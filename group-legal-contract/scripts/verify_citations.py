#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""扫描文本 / JSON / docx 中的法条引用并逐条核验（verify_citations.py）。

用法示例：
  python scripts/verify_citations.py work/revisions.json --out work/citation_report.json --md-out work/appendix_b.md
  python scripts/verify_citations.py 意见书.md --format markdown
  python scripts/verify_citations.py 审查修订版.docx
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from lib.kb import load_index
from lib.citations import extract_citations, verify_citation_matches, VerificationItem
from lib.textio import read_text_auto


_W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _extract_text_from_docx(docx_path: Path, exclude_sections: list[str] | None = None) -> str:
    """从 .docx 的正文与批注 XML 中提取纯文本（按段落分行）。"""
    lines: list[str] = []
    with zipfile.ZipFile(docx_path, "r") as zf:
        for part_name in ("word/document.xml", "word/comments.xml"):
            if part_name not in zf.namelist():
                continue
            xml_bytes = zf.read(part_name)
            root = ET.fromstring(xml_bytes)
            for p in root.iter(f"{_W_NS}p"):
                chunks: list[str] = []
                def append_visible(elem: ET.Element, deleted: bool = False) -> None:
                    # Track revision deletion containers explicitly.  A
                    # deleted citation must not count as a live citation in a
                    # signed/clean document; w:delText is therefore ignored as
                    # well as all descendants of w:del.
                    local = elem.tag.rsplit("}", 1)[-1]
                    now_deleted = deleted or local in {"del", "delText"}
                    if not now_deleted and local == "t" and elem.text:
                        chunks.append(elem.text)
                    for child in list(elem):
                        append_visible(child, now_deleted)

                append_visible(p)
                para_str = "".join(chunks).strip()
                if para_str:
                    lines.append(para_str)
    text = "\n".join(lines)
    if exclude_sections:
        wanted = [re.sub(r"\s+", "", x).lower() for x in exclude_sections if x.strip()]
        kept: list[str] = []
        for line in text.splitlines():
            norm_line = re.sub(r"\s+", "", line).lower()
            if any(key in norm_line for key in wanted):
                break
            kept.append(line)
        text = "\n".join(kept)
    return text


def _extract_text_from_json(json_path: Path) -> str:
    """将 JSON 中的所有字符串值按行拼接，同时保留行号可读性。"""
    raw, _ = read_text_auto(json_path)
    try:
        data = json.loads(raw)
    except Exception:
        return raw

    strings: list[str] = []

    def walk(obj):
        if isinstance(obj, str):
            strings.append(obj)
        elif isinstance(obj, list):
            for item in obj:
                walk(item)
        elif isinstance(obj, dict):
            for v in obj.values():
                walk(v)

    walk(data)
    return "\n".join(strings)


def _strip_markdown_sections(text: str, excluded_headings: list[str] | None = None) -> tuple[str, list[str]]:
    """移除指定 Markdown 标题及其子内容，避免把已生成的附录再次当作正文扫描。

    标题匹配采用包含关系（忽略大小写和空白），例如传入 ``附录 B`` 可同时
    排除 ``## 附录 B：法条引用自动核验表`` 和 ``### 附录 B：...``。返回值
    中保留了实际命中的标题，便于把排除范围写入核验报告。
    """
    wanted = [re.sub(r"\s+", "", x).lower() for x in (excluded_headings or []) if x.strip()]
    if not wanted:
        return text, []

    kept: list[str] = []
    excluded: list[str] = []
    active_level: int | None = None
    heading_re = re.compile(r"^(#{1,6})\s+(.+?)\s*$")

    for line in text.splitlines():
        hm = heading_re.match(line)
        if hm:
            level = len(hm.group(1))
            title = hm.group(2).strip()
            norm_title = re.sub(r"\s+", "", title).lower()
            if active_level is not None and level <= active_level:
                active_level = None
            if any(key in norm_title for key in wanted):
                active_level = level
                if title not in excluded:
                    excluded.append(title)
                continue
        if active_level is None:
            kept.append(line)

    return "\n".join(kept), excluded


def read_input_text(path: Path, exclude_sections: list[str] | None = None) -> str:
    suffix = path.suffix.lower()
    if suffix == ".docx":
        return _extract_text_from_docx(path, exclude_sections)
    if suffix == ".json":
        return _extract_text_from_json(path)
    text, _ = read_text_auto(path)
    if suffix in {".md", ".markdown"} and exclude_sections:
        text, _ = _strip_markdown_sections(text, exclude_sections)
    return text


def _dedupe_items(items: list[VerificationItem]) -> tuple[list[VerificationItem], int]:
    """按法源、条号和核验状态去重，返回 (去重后条目, 删除数量)。"""
    out: list[VerificationItem] = []
    seen: set[tuple[str, str | None, str]] = set()
    removed = 0
    for item in items:
        key = (item.law_id or item.law_query.strip(), item.article_label, item.status)
        if key in seen:
            removed += 1
            continue
        seen.add(key)
        out.append(item)
    return out, removed


def render_markdown_table(report: dict) -> str:
    """生成可直接嵌入《法律审查意见书》附录 B 的 Markdown 核验表。"""
    lines = [
        "### 附录 B：法条引用自动核验表",
        "",
        f"> **知识库快照**：`{report.get('kb_snapshot')}` ｜ **核验条数**：`{report['summary']['total']}` ｜ "
        f"**扫描出现次数**：`{report['summary'].get('occurrences_scanned', report['summary']['total'])}` ｜ "
        f"**已核验**：`{report['summary']['verified']}` ｜ **状态待核实**：`{report['summary']['status_unverified']}` ｜ "
        f"**失效/废止**：`{report['summary'].get('law_not_in_force', 0)}` ｜ "
        f"**库外法源**：`{report['summary']['out_of_kb']}` ｜ **条号不存在**：`{report['summary']['article_not_found']}`",
        "",
        "| 序号 | 引用原文 | 识别法源 | 条号 | 核验结论 | 本地库状态 / 施行日期 | 核验说明 |",
        "| :---: | :--- | :--- | :--- | :---: | :--- | :--- |",
    ]
    items = report.get("items", [])
    if not items:
        lines.append("| — | （未检测到法条引用） | — | — | — | — | — |")
        return "\n".join(lines)

    for idx_i, it in enumerate(items, start=1):
        raw_cit = (it.get("raw_citation") or "").replace("|", "｜")
        law_name = (it.get("law_title") or it.get("law_query") or "").replace("|", "｜")
        art_lbl = (it.get("article_label") or "—").replace("|", "｜")
        status_cn = it.get("status_cn") or it.get("status")
        law_meta = "—"
        if it.get("law_id"):
            eff = it.get("effective_date") or "见正文"
            law_meta = f"`{it['law_id']}` {it.get('law_status') or ''}（{eff}）"
        note = (it.get("note") or "通过").replace("|", "｜")
        lines.append(f"| {idx_i} | {raw_cit} | 《{law_name}》 | {art_lbl} | **{status_cn}** | {law_meta} | {note} |")
    return "\n".join(lines)


def verify_file(
    input_path: str | Path,
    allow_unmarked_out_of_kb: bool = False,
    index_path: str | None = None,
    exclude_sections: list[str] | None = None,
    dedupe: bool = False,
    as_of: str | None = None,
) -> tuple[int, dict]:
    input_path = Path(input_path)
    idx = load_index(index_path)
    if input_path.suffix.lower() in {".md", ".markdown"} and exclude_sections:
        raw_text, _ = read_text_auto(input_path)
        text, excluded_headings = _strip_markdown_sections(raw_text, exclude_sections)
    else:
        text = read_input_text(input_path, exclude_sections)
        excluded_headings = []
    matches = extract_citations(text, index_path=index_path)
    items: list[VerificationItem] = verify_citation_matches(matches, index_path=index_path, as_of=as_of)
    occurrence_count = len(items)
    duplicate_count = 0
    if dedupe:
        items, duplicate_count = _dedupe_items(items)

    summary = {
        "total": len(items),
        "occurrences_scanned": occurrence_count,
        "duplicates_removed": duplicate_count,
        "verified": sum(1 for x in items if x.status == "verified"),
        "status_unverified": sum(1 for x in items if x.status == "status_unverified"),
        "out_of_kb": sum(1 for x in items if x.status == "out_of_kb"),
        "unmarked_out_of_kb": sum(1 for x in items if x.status == "out_of_kb" and not x.acknowledged_out_of_kb),
        "article_not_found": sum(1 for x in items if x.status == "article_not_found"),
        "ambiguous_law": sum(1 for x in items if x.status == "ambiguous_law"),
        "not_yet_effective": sum(1 for x in items if x.status == "not_yet_effective"),
        "law_not_in_force": sum(1 for x in items if x.status == "law_not_in_force"),
    }

    has_fatal = (
        summary["article_not_found"] > 0
        or summary["ambiguous_law"] > 0
        or summary["not_yet_effective"] > 0
        or summary["law_not_in_force"] > 0
        or (not allow_unmarked_out_of_kb and summary["unmarked_out_of_kb"] > 0)
    )
    exit_code = 1 if has_fatal else 0

    report = {
        "input": str(input_path),
        "kb_snapshot": idx.get("kb_snapshot"),
        "as_of": as_of,
        "passed": not has_fatal,
        "excluded_sections_requested": exclude_sections or [],
        "excluded_sections": excluded_headings,
        "deduplicated": dedupe,
        "summary": summary,
        "items": [x.to_dict() for x in items],
    }
    return exit_code, report


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8")

    parser = argparse.ArgumentParser(description="扫描并核验文档（md/txt/json/docx）中的法条引用")
    parser.add_argument("input", help="待核验的文件路径（.md / .txt / .json / .docx）")
    parser.add_argument("--out", default=None, help="输出 JSON 报告文件路径")
    parser.add_argument("--md-out", default=None, help="输出 Markdown 核验表文件路径")
    parser.add_argument(
        "--allow-unmarked-out-of-kb",
        action="store_true",
        help="允许未标注【库外法源·待人工核实】的库外法源不触发非零退出码（用于扫描原始清单）",
    )
    parser.add_argument(
        "--exclude-section",
        action="append",
        default=[],
        help="Markdown 标题关键词；排除该标题及其子内容（可重复，用于跳过已生成的附录）",
    )
    parser.add_argument(
        "--dedupe",
        action="store_true",
        help="按法源/条号/状态去重；报告会保留原始扫描数和删除数量",
    )
    parser.add_argument("--index", default=None, help="指定 kb_index.json 路径")
    parser.add_argument("--as-of", default=None, help="按指定日期 YYYY-MM-DD 检查法规生效时点")
    parser.add_argument("--format", choices=["json", "markdown"], default="json", help="标准输出格式（默认 json）")
    args = parser.parse_args(argv)

    in_path = Path(args.input)
    if not in_path.exists():
        print(json.dumps({"error": f"输入文件不存在: {in_path}"}, ensure_ascii=False), file=sys.stderr)
        return 1

    try:
        code, report = verify_file(
            input_path=in_path,
            allow_unmarked_out_of_kb=args.allow_unmarked_out_of_kb,
            index_path=args.index,
            exclude_sections=args.exclude_section,
            dedupe=args.dedupe,
            as_of=args.as_of,
        )
    except Exception as exc:
        print(json.dumps({"error": str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 1

    if args.out:
        out_p = Path(args.out)
        out_p.parent.mkdir(parents=True, exist_ok=True)
        out_p.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    md_table = render_markdown_table(report)
    if args.md_out:
        md_p = Path(args.md_out)
        md_p.parent.mkdir(parents=True, exist_ok=True)
        md_p.write_text(md_table + "\n", encoding="utf-8")

    if args.format == "markdown":
        print(md_table)
    else:
        print(json.dumps(report, ensure_ascii=False, indent=2))

    return code


if __name__ == "__main__":
    raise SystemExit(main())
