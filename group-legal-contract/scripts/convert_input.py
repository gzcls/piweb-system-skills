#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""将旧版 Word/WPS/RTF 文件安全转换为可修订的 `.docx`。

LibreOffice 只负责格式转换；转换完成后仍会检查 ZIP CRC、OOXML 必需部件、
XML 可解析性和 python-docx 重开结果。输入文件只读，输出先写临时文件并在
校验通过后原子替换，避免把半成品交给后续 `docx_extract.py`。
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from lib.ooxml import sanitize_docx_file, validate_docx_file  # noqa: E402
from lib.textio import atomic_write_bytes  # noqa: E402


LEGACY_SUFFIXES = {".doc", ".wps", ".rtf"}


def _find_soffice() -> str | None:
    return shutil.which("soffice") or shutil.which("libreoffice")


def _same_file(left: Path, right: Path) -> bool:
    if not right.exists():
        return False
    try:
        return os.path.samefile(left, right)
    except (FileNotFoundError, OSError):
        return left.resolve() == right.resolve()


def _find_converted_docx(folder: Path, source_stem: str) -> Path:
    expected = folder / f"{source_stem}.docx"
    if expected.is_file():
        return expected
    candidates = sorted(folder.glob("*.docx"))
    if len(candidates) == 1:
        return candidates[0]
    if not candidates:
        raise RuntimeError("LibreOffice 未生成 .docx 文件")
    raise RuntimeError(
        "LibreOffice 生成了多个 .docx 文件，无法安全判断目标: "
        + ", ".join(p.name for p in candidates)
    )


def convert_to_docx(
    input_path: str | Path,
    out_docx: str | Path,
    timeout: int = 120,
) -> dict[str, Any]:
    """转换旧格式并返回转换与校验信息。"""
    source = Path(input_path)
    target = Path(out_docx)
    if not source.is_file():
        raise FileNotFoundError(f"输入文件不存在: {source}")
    if target.suffix.lower() != ".docx":
        raise ValueError(f"Word 输出文件必须使用 .docx 扩展名: {target}")
    if _same_file(source, target):
        raise ValueError("输出路径不能覆盖输入文件")
    target.parent.mkdir(parents=True, exist_ok=True)

    suffix = source.suffix.lower()
    if suffix == ".docx":
        atomic_write_bytes(target, source.read_bytes())
        validation = validate_docx_file(target, strict=True)
        if not validation["valid"]:
            raise ValueError(f"输入 .docx 校验失败: {'; '.join(validation['errors'])}")
        return {
            "input": str(source),
            "output": str(target),
            "source_format": ".docx",
            "converted": False,
            "validation": validation,
        }

    if suffix not in LEGACY_SUFFIXES:
        raise ValueError(
            f"不支持的输入格式: {source.suffix or '无扩展名'}；支持 .doc、.wps、.rtf、.docx"
        )

    soffice = _find_soffice()
    if not soffice:
        raise RuntimeError("未检测到 LibreOffice (soffice)，无法转换旧版 Word/WPS 文件")

    with tempfile.TemporaryDirectory(prefix="legacy_to_docx_") as tmp_root:
        tmp_root_p = Path(tmp_root)
        converted_dir = tmp_root_p / "converted"
        converted_dir.mkdir()
        profile_dir = tmp_root_p / "profile"
        profile_uri = profile_dir.resolve().as_uri()
        cmd = [
            soffice,
            "--headless",
            f"-env:UserInstallation={profile_uri}",
            "--convert-to",
            "docx:Office Open XML Text",
            "--outdir",
            str(converted_dir),
            str(source.resolve()),
        ]
        try:
            proc = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=timeout,
                check=False,
            )
        except subprocess.TimeoutExpired as exc:
            raise RuntimeError(f"LibreOffice 转换超时（>{timeout} 秒）") from exc

        if proc.returncode != 0:
            detail = (proc.stderr or proc.stdout or "").strip()
            raise RuntimeError(f"LibreOffice 转换失败（退出码 {proc.returncode}）: {detail}")

        converted = _find_converted_docx(converted_dir, source.stem)
        # 保留复杂文档的 AlternateContent，后续 redline 保存时再做兼容性规范化。
        sanitize_docx_file(converted, unwrap_wps_ac=False)
        validation = validate_docx_file(converted, strict=True)
        if not validation["valid"]:
            raise ValueError(f"转换后的 .docx 校验失败: {'; '.join(validation['errors'])}")
        atomic_write_bytes(target, converted.read_bytes())

    final_validation = validate_docx_file(target, strict=True)
    if not final_validation["valid"]:
        raise ValueError(f"输出 .docx 校验失败: {'; '.join(final_validation['errors'])}")
    return {
        "input": str(source),
        "output": str(target),
        "source_format": suffix,
        "converted": True,
        "validation": final_validation,
        "warning": "旧格式已转换为 .docx，请在 Word/WPS 中核对版式、字体和文本框。",
    }


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8")

    parser = argparse.ArgumentParser(description="将 .doc/.wps/.rtf 安全转换为 .docx")
    parser.add_argument("input", help="旧版 Word/WPS/RTF 或现有 .docx 文件")
    parser.add_argument("--out", required=True, help="输出 .docx 路径")
    parser.add_argument("--timeout", type=int, default=120, help="LibreOffice 超时时间（秒）")
    args = parser.parse_args(argv)

    try:
        result = convert_to_docx(args.input, args.out, timeout=max(1, args.timeout))
    except Exception as exc:
        print(json.dumps({"error": str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 1
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
