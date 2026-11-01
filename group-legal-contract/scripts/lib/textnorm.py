# -*- coding: utf-8 -*-
"""文本归一化、偏移映射、条款号提取与 revisions 校验工具。

匹配前归一化规则：
- 全角 ASCII（\\uFF01–\\uFF5E）转半角（\\u0021–\\u007E），涵盖全角括号、冒号、逗号、数字、字母等
- 中英文双引号（“ ” „ ‟ ＂）统一为 "
- 中英文单引号（‘ ’ ‚ ‛ ＇）统一为 '
- 全角空格（\\u3000）、不换行空格（\\u00A0）、换行（\\r/\\n）、制表符（\\t）及连续空白折叠为单个半角空格
- 维护归一化字符索引到原始字符串区间 (raw_start, raw_end) 的精确映射，以便写回时保持原字符不变。
"""
from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

from .textio import read_text_auto

_DOUBLE_QUOTES = frozenset({"“", "”", "„", "‟", "＂", "\u201c", "\u201d"})
_SINGLE_QUOTES = frozenset({"‘", "’", "‚", "‛", "＇", "\u2018", "\u2019"})
_WHITESPACE_CHARS = frozenset({" ", "\t", "\n", "\r", "\u3000", "\u00a0", "\v", "\f"})
_ZERO_WIDTH_CHARS = frozenset({"\u200b", "\u200c", "\u200d", "\ufeff"})


def _normalize_single_char(ch: str) -> str:
    """将单个非空白字符归一化。"""
    if ch in _DOUBLE_QUOTES:
        return '"'
    if ch in _SINGLE_QUOTES:
        return "'"
    code = ord(ch)
    # 全角 ASCII (\uFF01..\uFF5E) -> 半角 (\u0021..\u007E)
    if 0xFF01 <= code <= 0xFF5E:
        return chr(code - 0xFEE0)
    # 常见全角顿号/句号保持不变，仅转换常见破折号或波浪线变体（单字符对单字符）
    if ch == "～":
        return "~"
    return ch


def normalize_with_map(text: str) -> tuple[str, list[tuple[int, int]]]:
    """将字符串归一化，并返回 (归一化文本, 每个归一化字符对应的原字符串区间列表)。

    对于连续空白折叠成的单个空格，其区间为 (首个空白位置, 末尾空白位置+1)。
    对于零宽字符，直接并入相邻字符区间或跳过。
    """
    if not text:
        return "", []

    norm_chars: list[str] = []
    spans: list[tuple[int, int]] = []
    n = len(text)
    i = 0

    while i < n:
        ch = text[i]
        if ch in _ZERO_WIDTH_CHARS:
            i += 1
            continue
        if ch in _WHITESPACE_CHARS:
            ws_start = i
            i += 1
            while i < n and (text[i] in _WHITESPACE_CHARS or text[i] in _ZERO_WIDTH_CHARS):
                i += 1
            ws_end = i
            norm_chars.append(" ")
            spans.append((ws_start, ws_end))
        else:
            char_start = i
            norm_ch = _normalize_single_char(ch)
            i += 1
            # 若紧随零宽字符，将其吸收进当前字符的 span，避免切开原文本时遗留孤立零宽字符
            while i < n and text[i] in _ZERO_WIDTH_CHARS:
                i += 1
            norm_chars.append(norm_ch)
            spans.append((char_start, i))

    return "".join(norm_chars), spans


def normalize_text(text: str) -> str:
    """返回归一化后的文本字符串。"""
    norm, _ = normalize_with_map(text)
    return norm


def find_normalized_matches(raw_text: str, match_str: str) -> list[tuple[int, int]]:
    """在 raw_text 中按归一化规则查找 match_str 的所有不重叠出现位置，返回原字符串区间列表。"""
    if not raw_text or not match_str:
        return []

    norm_raw, raw_spans = normalize_with_map(raw_text)
    norm_match, _ = normalize_with_map(match_str)

    # 若 match_str 两端有空白但内部有非空白字符，允许精确匹配归一化后的串；
    # 若完全是空白，则退化为在原串直接查找
    if not norm_match:
        return []

    results: list[tuple[int, int]] = []
    m_len = len(norm_match)
    start_pos = 0
    while start_pos <= len(norm_raw) - m_len:
        idx = norm_raw.find(norm_match, start_pos)
        if idx == -1:
            break
        orig_start = raw_spans[idx][0]
        orig_end = raw_spans[idx + m_len - 1][1]
        results.append((orig_start, orig_end))
        start_pos = idx + m_len

    # 若带首尾空格未命中，尝试去除 match_str 首尾空白后再匹配一次
    if not results and norm_match.strip() and norm_match != norm_match.strip():
        norm_match_stripped = norm_match.strip()
        m_len = len(norm_match_stripped)
        start_pos = 0
        while start_pos <= len(norm_raw) - m_len:
            idx = norm_raw.find(norm_match_stripped, start_pos)
            if idx == -1:
                break
            orig_start = raw_spans[idx][0]
            orig_end = raw_spans[idx + m_len - 1][1]
            results.append((orig_start, orig_end))
            start_pos = idx + m_len

    return results


def locate_match(
    raw_text: str,
    match_str: str | None,
    occurrence: int | None = None,
) -> tuple[tuple[int, int] | None, str | None]:
    """在段落文本中定位目标片段。

    返回:
        ((raw_start, raw_end), None): 定位成功
        (None, 错误原因): 找不到、未提供 match 或多处匹配未指定 occurrence
    """
    if not match_str:
        return None, "未提供 anchor.match 匹配文本"

    matches = find_normalized_matches(raw_text, match_str)
    if not matches:
        return None, f"在段落中未找到匹配文本: “{match_str}”"

    if occurrence is not None:
        if occurrence < 1 or occurrence > len(matches):
            return (
                None,
                f"指定的 occurrence={occurrence} 超出实际匹配次数({len(matches)}): “{match_str}”",
            )
        return matches[occurrence - 1], None

    if len(matches) > 1:
        return (
            None,
            f"匹配文本“{match_str}”在同段落中出现 {len(matches)} 处且未指定 occurrence，拒绝猜测",
        )

    return matches[0], None


# 条款号识别正则（按优先级排列）
_CLAUSE_PATTERNS: list[re.Pattern[str]] = [
    # 第X编 / 第X章 / 第X节 / 第X条（支持中文数字或阿拉伯数字）
    re.compile(r"^\s*(第\s*[零〇○一二两三四五六七八九十百千万\d]+\s*[编章节条])"),
    # 类似 *[版本A...]：1.2 或直接 1.1 / 1.2.3
    re.compile(r"^\s*(?:[*•\-]\s*)?(?:【[^】]+】|\(\s*[^)]+\)|\[[^\]]+\])[：:\s]*(\d+(?:\.\d+)+)"),
    # 多级数字条款号：1.1、3.2.1（排除年份如 2026.10）
    re.compile(r"^\s*(\d{1,2}(?:\.\d{1,3})+)(?=\s|[^\d.]|$)"),
    # 中文序号：一、 二、 （一） (一)
    re.compile(r"^\s*([一二三四五六七八九十百]+)[、.．]"),
    re.compile(r"^\s*([（(][一二三四五六七八九十百]+[)）])"),
    # 括号阿拉伯数字：（1） (1)
    re.compile(r"^\s*([（(]\d{1,3}[)）])"),
    # 带点/顿号的阿拉伯数字：1. / 1、
    re.compile(r"^\s*(\d{1,3})[、.．](?!\d)"),
    # 带圈数字：① ② ...
    re.compile(r"^\s*([①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳])"),
    # 附件X
    re.compile(r"^\s*(附件[零〇○一二三四五六七八九十\d]+)"),
]


def extract_clause_ref(list_label: str | None, text: str | None) -> str:
    """从自动编号标签或段落首部文本提取归一化条款号。"""
    if list_label:
        cleaned = list_label.strip()
        # 去掉末尾多余标点（如 "1." -> "1", "一、" -> "一"，但保留括号与“第X条”）
        cleaned = re.sub(r"[、.．:：\s]+$", "", cleaned)
        if cleaned:
            return cleaned

    if not text:
        return ""

    # 去除首部 markdown 粗体符号干扰
    probe = re.sub(r"^\s*\*+\s*", "", text)
    for pat in _CLAUSE_PATTERNS:
        m = pat.match(probe)
        if m:
            ref = re.sub(r"\s+", "", m.group(1))
            # 统一全角括号或保持常见格式
            return ref

    return ""


_VALID_OPS = frozenset(
    {
        "replace",
        "delete",
        "insert_text",
        "insert_paragraph_after",
        "insert_paragraph_before",
        "delete_paragraph",
        "comment_only",
    }
)
_VALID_RISKS = frozenset({"高", "中", "低"})


def _builtin_validate_revisions(data: Any) -> list[str]:
    """当 jsonschema 未安装或额外检查时的内置轻量校验。"""
    errors: list[str] = []
    if not isinstance(data, dict):
        return ["根节点必须为 JSON 对象 (object)"]
    if "revisions" not in data or not isinstance(data.get("revisions"), list):
        return ["缺少必需的 revisions 数组字段"]
    if "author" in data and (not isinstance(data["author"], str) or not data["author"].strip()):
        errors.append("author 字段必须为非空字符串")

    for idx, item in enumerate(data["revisions"]):
        prefix = f"revisions[{idx}]"
        if not isinstance(item, dict):
            errors.append(f"{prefix}: 必须为对象")
            continue
        for req_key in ("id", "op", "anchor"):
            if req_key not in item:
                errors.append(f"{prefix}: 缺少必需字段 '{req_key}'")
        rev_id = item.get("id")
        if rev_id is not None and (not isinstance(rev_id, str) or not rev_id.strip()):
            errors.append(f"{prefix}.id: 必须为非空字符串")
        op = item.get("op")
        if op not in _VALID_OPS:
            errors.append(f"{prefix}.op: 非法操作类型 '{op}'")
        anchor = item.get("anchor")
        if isinstance(anchor, dict):
            block_id = anchor.get("block_id")
            if not isinstance(block_id, str) or not block_id.strip():
                errors.append(f"{prefix}.anchor.block_id: 必须为非空字符串")
            occ = anchor.get("occurrence")
            if occ is not None and (not isinstance(occ, int) or isinstance(occ, bool) or occ < 1):
                errors.append(f"{prefix}.anchor.occurrence: 必须为 >= 1 的整数")
            pos = anchor.get("position")
            if pos is not None and pos not in ("before", "after"):
                errors.append(f"{prefix}.anchor.position: 必须为 'before' 或 'after'")
        elif anchor is not None:
            errors.append(f"{prefix}.anchor: 必须为对象")

        comment = item.get("comment")
        if comment is not None:
            if not isinstance(comment, dict):
                errors.append(f"{prefix}.comment: 必须为对象")
            else:
                if comment.get("risk") not in _VALID_RISKS:
                    errors.append(f"{prefix}.comment.risk: 必须为 '高'、'中' 或 '低'")
                if not isinstance(comment.get("issue"), str) or not comment.get("issue", "").strip():
                    errors.append(f"{prefix}.comment.issue: 必须为非空字符串")
                if "basis" in comment and not isinstance(comment["basis"], list):
                    errors.append(f"{prefix}.comment.basis: 必须为字符串数组")

    return errors


def validate_revisions_data(data: Any, schema_path: Path | None = None) -> list[str]:
    """校验 revisions.json 数据，优先使用 jsonschema，缺失时降级为内置校验。"""
    errors: list[str] = []
    if schema_path is None:
        candidate = (
            Path(__file__).resolve().parent.parent.parent
            / "assets"
            / "schemas"
            / "revisions.schema.json"
        )
        if candidate.is_file():
            schema_path = candidate

    used_jsonschema = False
    if schema_path and schema_path.is_file():
        try:
            import jsonschema  # type: ignore

            schema_text, _ = read_text_auto(schema_path)
            schema = json.loads(schema_text)
            validator = jsonschema.Draft202012Validator(schema)
            for err in sorted(validator.iter_errors(data), key=lambda e: list(e.path)):
                loc = ".".join(str(p) for p in err.path) or "<root>"
                errors.append(f"{loc}: {err.message}")
            used_jsonschema = True
        except ImportError:
            used_jsonschema = False

    if not used_jsonschema:
        errors.extend(_builtin_validate_revisions(data))

    return errors
