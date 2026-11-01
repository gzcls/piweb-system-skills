# -*- coding: utf-8 -*-
"""跨平台文本读取、OOXML 安全字符清理与原子文件写入工具。

Word 的 XML 部件必须是 XML 1.0 文本。用户上传的 Markdown、JSON 或合同
批注可能来自 Windows 编辑器（GB18030/UTF-16）、带 BOM 的编辑器，甚至含
有 NUL/孤立代理项。这里集中处理这些边界，避免把乱码或半成品写入交付物。
"""
from __future__ import annotations

import os
import tempfile
from pathlib import Path


def _normalise_encoding(encoding: str) -> str:
    value = encoding.strip().lower().replace("_", "-")
    if value in {"utf8", "utf-8"}:
        # utf-8-sig 同时兼容有 BOM 与无 BOM 的 UTF-8。
        return "utf-8-sig"
    return encoding


def _looks_like_utf16(raw: bytes) -> tuple[str, ...]:
    """根据 NUL 字节分布返回无 BOM UTF-16 的候选顺序。"""
    sample = raw[:8192]
    if len(sample) < 4:
        return ()

    odd = sample[1::2]
    even = sample[0::2]
    odd_ratio = odd.count(0) / max(len(odd), 1)
    even_ratio = even.count(0) / max(len(even), 1)
    candidates: list[str] = []
    if odd_ratio >= 0.25:
        candidates.append("utf-16-le")
    if even_ratio >= 0.25:
        candidates.append("utf-16-be")

    # 纯中文 UTF-16 文本的每个码元都可能没有 NUL。尝试两种端序，
    # 仅接受“绝大多数可打印且含 CJK 字符”的候选，避免把 GB18030 误判成 UTF-16。
    if not candidates and len(raw) % 2 == 0:
        scored: list[tuple[int, float, str]] = []
        for enc in ("utf-16-le", "utf-16-be"):
            try:
                decoded = raw.decode(enc)
            except UnicodeDecodeError:
                continue
            if not decoded:
                continue
            printable = sum(
                1 for ch in decoded if ch in "\t\n\r" or ord(ch) >= 0x20
            ) / len(decoded)
            cjk_count = sum(1 for ch in decoded if "\u3400" <= ch <= "\u9fff")
            cjk_ratio = cjk_count / len(decoded)
            if printable >= 0.8 and cjk_count >= 2 and cjk_ratio >= 0.25:
                scored.append((cjk_count, printable, enc))
        candidates.extend(enc for _, _, enc in sorted(scored, reverse=True))
    return tuple(candidates)


def read_text_auto(path: str | Path, encoding: str | None = None) -> tuple[str, str]:
    """读取文本并返回 ``(内容, 实际编码)``。

    识别顺序为显式编码、UTF-8/UTF-8 BOM、UTF-16/UTF-32 BOM、无 BOM UTF-16，
    最后是 Windows 中文环境常见的 GB18030。所有尝试均为严格解码，避免用
    ``errors='replace'`` 静默制造不可见乱码。
    """
    p = Path(path)
    raw = p.read_bytes()

    if encoding:
        enc = _normalise_encoding(encoding)
        try:
            decoded = raw.decode(enc)
            if decoded.startswith("\ufeff"):
                decoded = decoded[1:]
            return decoded, enc
        except UnicodeDecodeError as exc:
            raise UnicodeError(f"无法按 {encoding} 解码文本文件: {p}") from exc

    bom_encodings = (
        (b"\xef\xbb\xbf", "utf-8-sig"),
        (b"\xff\xfe\x00\x00", "utf-32"),
        (b"\x00\x00\xfe\xff", "utf-32"),
        (b"\xff\xfe", "utf-16"),
        (b"\xfe\xff", "utf-16"),
    )
    for bom, enc in bom_encodings:
        if raw.startswith(bom):
            try:
                return raw.decode(enc), enc
            except UnicodeDecodeError as exc:
                raise UnicodeError(f"文本文件带 {enc} BOM，但内容损坏: {p}") from exc

    utf8_text: str | None = None
    try:
        utf8_text = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        pass

    # UTF-16 无 BOM 的中文在 UTF-8 视角下有大量 C0 控制字符/NUL；
    # 先尝试已识别的 UTF-16 候选，避免把它误当成“可解码但乱码”的 UTF-8。
    utf8_has_c0 = bool(
        utf8_text
        and any(
            0 <= ord(ch) < 0x20 and ch not in "\t\n\r" for ch in utf8_text
        )
    )
    for enc in _looks_like_utf16(raw) if utf8_text is None or utf8_has_c0 else ():
        try:
            return raw.decode(enc), enc
        except UnicodeDecodeError:
            continue

    if utf8_text is not None:
        return utf8_text, "utf-8"

    try:
        return raw.decode("gb18030"), "gb18030"
    except UnicodeDecodeError as exc:
        raise UnicodeError(
            f"无法识别文本文件编码: {p}；请使用 --encoding 指定 utf-8、utf-16 或 gb18030"
        ) from exc


def sanitize_xml_text(text: str, *, preserve_line_controls: bool = True) -> tuple[str, int]:
    """替换 XML 1.0 禁止字符，并返回 ``(清理后文本, 变更数量)``。

    ``\x0b``/``\x0c`` 常由复制粘贴产生，按换行处理以免相邻条款黏在一起；
    NUL、其他 C0 控制字符、U+FFFE/U+FFFF 与孤立 UTF-16 代理项改为空格，
    既保证 XML 合法，也避免清理后把两个词黏成一个词。
    """
    if not isinstance(text, str):
        text = str(text)
    if not text:
        return "", 0

    result: list[str] = []
    changed = 0
    for ch in str(text):
        code = ord(ch)
        if (
            code in (0x9, 0xA, 0xD)
            or 0x20 <= code <= 0xD7FF
            or 0xE000 <= code <= 0xFFFD
            or 0x10000 <= code <= 0x10FFFF
        ):
            result.append(ch)
            continue

        if preserve_line_controls and code in (0x0B, 0x0C):
            result.append("\n")
        else:
            result.append(" ")
        changed += 1

    return "".join(result), changed


def sanitize_xml_attribute(text: str) -> tuple[str, int]:
    """清理 XML 属性值；换行/制表符改为空格，避免作者等字段变成多行。"""
    clean, changed = sanitize_xml_text(text, preserve_line_controls=False)
    normalised = clean.replace("\r", " ").replace("\n", " ").replace("\t", " ")
    changed += sum(1 for old, new in zip(clean, normalised) if old != new)
    return normalised, changed


def atomic_write_bytes(path: str | Path, data: bytes) -> None:
    """在目标目录写临时文件并原子替换，避免留下半个交付文件。"""
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp_name = tempfile.mkstemp(
        prefix=f".{target.name}.", suffix=".tmp", dir=str(target.parent)
    )
    try:
        with os.fdopen(fd, "wb") as handle:
            handle.write(data)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(tmp_name, target)
    except Exception:
        try:
            os.unlink(tmp_name)
        except FileNotFoundError:
            pass
        raise
