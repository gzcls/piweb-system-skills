#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""将合同范本或审查意见书 Markdown 子集渲染为集团公文风格 .docx。

支持特性：
- # ~ #### 标题、普通段落、**粗体**、`行内代码`（去除反引号输出普通文字）
- 有序列表、无序列表、复选框列表 (- [ ] / - [x])、引用块 >、水平线 ---
- 表格（含单元格内 <br> 换行、列对齐、表头底纹；>=6 列宽表自动切换横向页面节）
- [●] 与 `[说明]` 占位符自动黄色高亮（--placeholder-highlight，默认开启）
- 范本批注块识别（首行匹配 `^>\\s*\\*?\\*?【[^】]*批注[^】]*】\\*?\\*?` 的连续 `>` 引用块）：
  - `--notes-as-comments`：将批注块转为挂载在前一段落上的 Word 批注
  - `--strip-notes`：剥离删除所有批注块（保留普通引用块如适用场景/使用说明）
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import re
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any

import docx
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from lib.ooxml import (
    DocxPackage,
    add_comment_entry,
    attach_comment_to_whole_paragraph,
    sanitize_docx_file,
    validate_docx_file,
    w14_tag,
    w_tag,
)
from lib.textio import read_text_auto, sanitize_xml_text


DEFAULT_REFERENCE = (
    SCRIPT_DIR.parent / "assets" / "docx" / "reference.docx"
)

# 批注块首行识别正则：以 `> **【…批注…】**` 或 `> 【…批注…】` 开头
NOTE_BLOCK_HEADER_RE = re.compile(r"^>\s*\*?\*?【[^】]*批注[^】]*】\*?\*?")
HEADING_RE = re.compile(r"^\s*(#{1,6})\s+(.+?)\s*$")
HR_RE = re.compile(r"^\s*(?:-\s*|\*\s*|_\s*){3,}$")
TABLE_SEP_CELL_RE = re.compile(r"^:?-{2,}:?$")
CHECKBOX_RE = re.compile(r"^\s*[-*+]\s+\[([ xX])\]\s+(.*)$")
UNORDERED_LIST_RE = re.compile(r"^\s*[-*+]\s+(.+)$")

# 签署版禁止出现的内部工作痕迹。该扫描器同时检查正文、表格、页眉页脚
# 和脚注，避免只读 Document.paragraphs 导致内部文本漏检。
SIGNING_INTERNAL_MARKER_RE = re.compile(
    r"(?:起草批注|谈判批注|风险批注|集团法务|内部(?:使用)?|TODO|FIXME|版本\s*[AB])",
    re.IGNORECASE,
)
SIGNING_PLACEHOLDER_RE = re.compile(
    r"(?:\[[^\]\n]*(?:●|待填写|填写|请输入|待确认|待核实)[^\]\n]*\]|"
    r"\{\{[^}\n]+\}\}|\bTODO\b|\bFIXME\b)",
    re.IGNORECASE,
)
_W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _set_run_cjk_font(
    run,
    east_asia: str = "宋体",
    ascii_font: str = "Times New Roman",
    size_pt: float | None = None,
) -> None:
    """为新建 run 显式设置东亚字体属性，确保中文渲染为宋体/黑体/仿宋。"""
    run.font.name = ascii_font
    if size_pt is not None:
        run.font.size = Pt(size_pt)
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), ascii_font)
    rFonts.set(qn("w:hAnsi"), ascii_font)
    rFonts.set(qn("w:eastAsia"), east_asia)
    rFonts.set(qn("w:cs"), east_asia)


def scan_signing_docx(path: str | Path) -> dict[str, Any]:
    """扫描签署版 DOCX 的内部内容、占位符和批注残留。

    返回结构可直接嵌入起草流水线报告。扫描只读取可见 ``w:t``，跳过
    ``w:del``/``w:delText``，并覆盖正文、表格、页眉页脚及脚注；存在有效
    ``comments.xml`` 或批注元素本身即视为签署版未清洁。
    """
    docx_path = Path(path)
    internal: list[dict[str, Any]] = []
    placeholders: list[dict[str, Any]] = []
    scanned_parts: list[str] = []
    try:
        with zipfile.ZipFile(docx_path, "r") as archive:
            names = [
                name for name in archive.namelist()
                if name.startswith("word/") and name.endswith(".xml")
                and any(key in name for key in ("document.xml", "header", "footer", "footnotes", "endnotes"))
            ]
            for part_name in names:
                try:
                    root = ET.fromstring(archive.read(part_name))
                except Exception:
                    continue
                scanned_parts.append(part_name)
                lines: list[str] = []

                def visit(elem: ET.Element, deleted: bool = False) -> None:
                    local = elem.tag.rsplit("}", 1)[-1]
                    now_deleted = deleted or local in {"del", "delText"}
                    if not now_deleted and local == "t" and elem.text:
                        lines.append(elem.text)
                    for child in list(elem):
                        visit(child, now_deleted)

                visit(root)
                text = "".join(lines)
                for line_no, line in enumerate(text.splitlines() or [text], 1):
                    if not line.strip():
                        continue
                    if SIGNING_INTERNAL_MARKER_RE.search(line):
                        internal.append({"part": part_name, "line": line_no, "text": line.strip()})
                    for match in SIGNING_PLACEHOLDER_RE.finditer(line):
                        placeholders.append({"part": part_name, "line": line_no, "text": match.group(0)})

            comments_xml = archive.read("word/comments.xml") if "word/comments.xml" in archive.namelist() else b""
            has_comments = False
            if comments_xml:
                try:
                    comments_root = ET.fromstring(comments_xml)
                    has_comments = any(el.tag.rsplit("}", 1)[-1] == "comment" for el in comments_root.iter())
                except ET.ParseError:
                    has_comments = True
    except (OSError, KeyError, zipfile.BadZipFile, ET.ParseError) as exc:
        return {
            "passed": False,
            "error": str(exc),
            "has_comments": False,
            "internal_markers": [],
            "placeholders": [],
            "placeholder_count": 0,
            "scanned_parts": scanned_parts,
        }

    return {
        "passed": not internal and not placeholders and not has_comments,
        "has_comments": has_comments,
        "internal_markers": internal,
        "placeholders": placeholders,
        "placeholder_count": len(placeholders),
        "scanned_parts": scanned_parts,
    }


def strip_docx_comments(path: str | Path) -> dict[str, Any]:
    """Remove comments inherited from a reference template for signing output.

    ``python-docx`` preserves a template's comments even when Markdown notes
    are stripped.  A signing copy must remove the comments part, its OPC
    relationship/content-type entries, and all range/reference markers while
    leaving visible document text untouched.
    """
    source = Path(path).resolve()
    temporary = source.with_name(f".{source.name}.{os.getpid()}.comments.tmp")
    removed_markers = 0

    def local(tag: str) -> str:
        return tag.rsplit("}", 1)[-1]

    def clean_xml(name: str, data: bytes) -> bytes:
        nonlocal removed_markers
        try:
            root = ET.fromstring(data)
        except ET.ParseError:
            return data
        if name.endswith("document.xml") or "/header" in name or "/footer" in name or name.endswith("footnotes.xml") or name.endswith("endnotes.xml"):
            for parent in root.iter():
                for child in list(parent):
                    if local(child.tag) in {"commentRangeStart", "commentRangeEnd", "commentReference"}:
                        parent.remove(child)
                        removed_markers += 1
        if name == "word/_rels/document.xml.rels":
            for child in list(root):
                if child.attrib.get("Type", "").endswith("/comments"):
                    root.remove(child)
                    removed_markers += 1
        if name == "[Content_Types].xml":
            for child in list(root):
                if child.attrib.get("PartName", "").lower() == "/word/comments.xml":
                    root.remove(child)
                    removed_markers += 1
        return ET.tostring(root, encoding="utf-8", xml_declaration=True)

    try:
        with zipfile.ZipFile(source, "r") as archive, zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as target:
            for info in archive.infolist():
                if info.filename == "word/comments.xml":
                    continue
                payload = archive.read(info.filename)
                if info.filename.endswith(".xml") or info.filename.endswith(".rels"):
                    payload = clean_xml(info.filename, payload)
                target.writestr(info, payload)
        os.replace(temporary, source)
        validation = validate_docx_file(source, strict=True)
        if not validation.get("valid"):
            raise ValueError("删除模板批注后 DOCX 结构校验失败")
        return {"removed": removed_markers, "path": str(source), "validated": True}
    finally:
        if temporary.exists():
            temporary.unlink()


def _set_cell_shading(cell, fill_hex: str = "EAEAEA") -> None:
    """设置表格单元格背景底纹颜色（严格遵循 ECMA-376 CT_TcPrBase 子元素顺序）。"""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        after_tags = {
            qn("w:noWrap"), qn("w:tcMar"), qn("w:textDirection"), qn("w:tcFitText"),
            qn("w:vAlign"), qn("w:hideMark"), qn("w:headers"), qn("w:cellIns"),
            qn("w:cellDel"), qn("w:cellMerge"), qn("w:tcPrChange"),
        }
        insert_idx = len(tcPr)
        for idx, ch in enumerate(tcPr):
            if ch.tag in after_tags:
                insert_idx = idx
                break
        tcPr.insert(insert_idx, shd)
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)


def _set_table_borders(table) -> None:
    """为表格添加标准细实线边框（严格遵循 ECMA-376 CT_TblPrBase 子元素顺序，置于 tblLook 之前）。"""
    tblPr = table._tbl.tblPr
    tblBorders = tblPr.find(qn("w:tblBorders"))
    if tblBorders is None:
        tblBorders = OxmlElement("w:tblBorders")
        after_tags = {
            qn("w:shd"), qn("w:tblLayout"), qn("w:tblCellMar"),
            qn("w:tblLook"), qn("w:tblCaption"), qn("w:tblDescription"), qn("w:tblPrChange"),
        }
        insert_idx = len(tblPr)
        for idx, ch in enumerate(tblPr):
            if ch.tag in after_tags:
                insert_idx = idx
                break
        tblPr.insert(insert_idx, tblBorders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        edge_el = tblBorders.find(qn(f"w:{edge}"))
        if edge_el is None:
            edge_el = OxmlElement(f"w:{edge}")
            tblBorders.append(edge_el)
        edge_el.set(qn("w:val"), "single")
        edge_el.set(qn("w:sz"), "4")
        edge_el.set(qn("w:space"), "0")
        edge_el.set(qn("w:color"), "666666")


def _is_placeholder_backtick(content: str) -> bool:
    """判断反引号包裹的内容是否属于待填占位符（如 `[●]`、`[填写公司全称]`、`[●]市[●]区`、`2026年[●]月[●]日`）。"""
    s = content.strip()
    if not s:
        return False
    # 排除引用说明中字面展示的 `> 【起草/谈判批注】` 或文件路径如 `legal_kb/...`
    if s.startswith(">") or "legal_kb/" in s:
        return False
    if "●" in s:
        return True
    if s.startswith("[") and s.endswith("]"):
        return True
    if re.search(r"\[[^\]]+\]", s):
        return True
    return False


def _append_plain_or_bare_placeholder(
    p,
    text: str,
    bold: bool,
    highlight_placeholders: bool,
    east_asia_font: str = "宋体",
    size_pt: float | None = None,
) -> None:
    """将无反引号的文本写入段落，若包含裸 `[●]` 等占位符则对其应用黄色高亮。"""
    if not text:
        return
    if not highlight_placeholders or "●" not in text:
        r = p.add_run(text)
        r.bold = bold
        _set_run_cjk_font(r, east_asia=east_asia_font, size_pt=size_pt)
        return

    # 切分含 ● 的方括号占位符（如 [●]）
    parts = re.split(r"(\[[^\]]*●[^\]]*\])", text)
    for part in parts:
        if not part:
            continue
        r = p.add_run(part)
        r.bold = bold
        _set_run_cjk_font(r, east_asia=east_asia_font, size_pt=size_pt)
        if part.startswith("[") and part.endswith("]") and "●" in part:
            r.font.highlight_color = WD_COLOR_INDEX.YELLOW


def _render_inline_runs(
    p,
    text: str,
    highlight_placeholders: bool = True,
    base_bold: bool = False,
    east_asia_font: str = "宋体",
    size_pt: float | None = None,
) -> None:
    """解析行内 **粗体**、`反引号代码/占位符` 以及裸 [●]，写入段落 p。"""
    if not text:
        return

    # 第一层：按 **...** 切分粗体区间
    bold_tokens = re.split(r"(\*\*.*?\*\*)", text)
    for b_tok in bold_tokens:
        if not b_tok:
            continue
        if b_tok.startswith("**") and b_tok.endswith("**") and len(b_tok) >= 4:
            is_bold = True
            inner_bold = b_tok[2:-2]
        else:
            is_bold = base_bold
            inner_bold = b_tok

        # 第二层：按 `...` 切分反引号区间
        code_tokens = re.split(r"(`[^`]+`)", inner_bold)
        for c_tok in code_tokens:
            if not c_tok:
                continue
            if c_tok.startswith("`") and c_tok.endswith("`") and len(c_tok) >= 2:
                code_inner = c_tok[1:-1]
                r = p.add_run(code_inner)
                r.bold = is_bold
                _set_run_cjk_font(r, east_asia=east_asia_font, size_pt=size_pt)
                if highlight_placeholders and _is_placeholder_backtick(code_inner):
                    r.font.highlight_color = WD_COLOR_INDEX.YELLOW
            else:
                _append_plain_or_bare_placeholder(
                    p,
                    c_tok,
                    bold=is_bold,
                    highlight_placeholders=highlight_placeholders,
                    east_asia_font=east_asia_font,
                    size_pt=size_pt,
                )


def _split_table_row(line: str) -> list[str]:
    """解析 Markdown 表格单行，返回单元格文本列表。

    ``str.split('|')`` 会把合同中的 ``\\|`` 和反引号代码片段误拆成列，
    进而改变列数、触发错误的横向分页。这里按 Markdown 的转义和代码
    区间逐字符扫描，并在拆分时还原常见的转义字符。
    """
    s = line.strip()
    cells: list[str] = []
    buf: list[str] = []
    escaped = False
    in_code = False
    for ch in s:
        if escaped:
            if ch in {"|", "\\", "`"}:
                buf.append(ch)
            else:
                buf.extend(("\\", ch))
            escaped = False
            continue
        if ch == "\\":
            escaped = True
            continue
        if ch == "`":
            in_code = not in_code
            buf.append(ch)
            continue
        if ch == "|" and not in_code:
            cells.append("".join(buf).strip())
            buf = []
            continue
        buf.append(ch)
    if escaped:
        buf.append("\\")
    cells.append("".join(buf).strip())
    # 外围 | 只是 Markdown 语法，不是空白单元格。
    if s.startswith("|") and cells:
        cells = cells[1:]
    if s.endswith("|") and cells:
        cells = cells[:-1]
    return cells or [""]


def _is_table_separator_line(line: str) -> bool:
    """判断是否为 Markdown 表格对齐分隔行（如 | :---: | :--- |）。"""
    s = line.strip()
    if "|" not in s:
        return False
    cells = _split_table_row(s)
    if not cells:
        return False
    return all(bool(TABLE_SEP_CELL_RE.match(c)) for c in cells)


def _parse_table_alignments(sep_line: str) -> list[Any]:
    cells = _split_table_row(sep_line)
    aligns: list[Any] = []
    for c in cells:
        if c.startswith(":") and c.endswith(":"):
            aligns.append(WD_ALIGN_PARAGRAPH.CENTER)
        elif c.endswith(":"):
            aligns.append(WD_ALIGN_PARAGRAPH.RIGHT)
        else:
            aligns.append(WD_ALIGN_PARAGRAPH.LEFT)
    return aligns


def _clean_note_block_text(quote_lines: list[str]) -> str:
    """将批注引用块去除行首 `>` 与多余 Markdown 标记，组装为干净的批注文本。"""
    cleaned_lines: list[str] = []
    for raw_line in quote_lines:
        line = re.sub(r"^\s*>\s?", "", raw_line).strip()
        # 去除粗体星号与反引号
        line = re.sub(r"\*\*(.*?)\*\*", r"\1", line)
        line = re.sub(r"`([^`]+)`", r"\1", line)
        if line:
            cleaned_lines.append(line)
    return "\n".join(cleaned_lines)


def _ensure_landscape_section_for_wide_table(doc: docx.Document) -> None:
    """当遇到 >= 6 列的宽表时，将当前节或新建节设为 A4 横向 (297mm x 210mm)。"""
    curr_sec = doc.sections[-1]
    if curr_sec.orientation == WD_ORIENT.LANDSCAPE:
        return

    # 若整个文档尚未有很多正文或就在当前节直接切为横向，或者新建横向节
    # 直接将当前节（或新增连续/分页节）设为横向，确保宽表在横向页面中完整展示
    if len(doc.paragraphs) > 0:
        new_sec = doc.add_section(WD_SECTION.NEW_PAGE)
        new_sec.orientation = WD_ORIENT.LANDSCAPE
        new_sec.page_width = Mm(297)
        new_sec.page_height = Mm(210)
        new_sec.top_margin = Mm(20)
        new_sec.bottom_margin = Mm(20)
        new_sec.left_margin = Mm(18)
        new_sec.right_margin = Mm(18)
    else:
        curr_sec.orientation = WD_ORIENT.LANDSCAPE
        curr_sec.page_width = Mm(297)
        curr_sec.page_height = Mm(210)


def _ensure_portrait_section_after_wide_table(doc: docx.Document) -> None:
    """在宽表后的正文恢复 A4 纵向节，避免后续条款继续横向排版。"""
    curr_sec = doc.sections[-1]
    if curr_sec.orientation != WD_ORIENT.LANDSCAPE:
        return
    new_sec = doc.add_section(WD_SECTION.NEW_PAGE)
    new_sec.orientation = WD_ORIENT.PORTRAIT
    new_sec.page_width = Mm(210)
    new_sec.page_height = Mm(297)
    new_sec.top_margin = Mm(20)
    new_sec.bottom_margin = Mm(20)
    new_sec.left_margin = Mm(25)
    new_sec.right_margin = Mm(25)


def convert_md_to_docx(
    input_md: str | Path,
    out_docx: str | Path,
    reference_docx: str | Path | None = None,
    placeholder_highlight: bool = True,
    notes_as_comments: bool = False,
    strip_notes: bool = False,
    comment_author: str = "集团法务起草批注",
    input_encoding: str | None = None,
) -> dict[str, Any]:
    """将 Markdown 文件渲染为 .docx 并返回转换统计信息。"""
    md_path = Path(input_md)
    out_path = Path(out_docx)
    if out_path.suffix.lower() != ".docx":
        raise ValueError(f"Word 输出文件必须使用 .docx 扩展名: {out_path}")
    ref_path = Path(reference_docx) if reference_docx else DEFAULT_REFERENCE

    if ref_path and ref_path.is_file():
        doc = docx.Document(str(ref_path))
        # 确保模板中可能残留的初始空段被清空
        body = doc._body._body
        for p in list(body.findall(qn("w:p"))):
            if not (p.text or "").strip() and len(p.findall(qn("w:r"))) == 0:
                body.remove(p)
    else:
        doc = docx.Document()

    md_text, detected_encoding = read_text_auto(md_path, encoding=input_encoding)
    md_text, invalid_xml_chars_removed = sanitize_xml_text(md_text)
    lines = md_text.splitlines()
    n = len(lines)
    i = 0

    last_p_elem = None
    pending_comments_for_para: list[tuple[Any, str]] = []
    para_marker_counter = 0
    wide_table_active = False

    stats = {
        "headings": 0,
        "paragraphs": 0,
        "tables": 0,
        "wide_tables": 0,
        "portrait_section_restores": 0,
        "note_blocks_detected": 0,
        "note_blocks_stripped": 0,
        "note_blocks_as_comments": 0,
        "regular_Format_quotes": 0,
        "input_encoding": detected_encoding,
        "invalid_xml_chars_removed": invalid_xml_chars_removed,
    }

    while i < n:
        raw_line = lines[i]
        stripped = raw_line.strip()

        # 宽表结束后，只有当下一段确实不是另一张表时才插入纵向节。
        # 这样连续宽表可共享横向页面，普通条款则恢复默认版式。
        if wide_table_active and stripped:
            next_is_table = "|" in stripped and (i + 1 < n) and _is_table_separator_line(lines[i + 1])
            if not next_is_table:
                _ensure_portrait_section_after_wide_table(doc)
                stats["portrait_section_restores"] += 1
                wide_table_active = False

        # 1. 空行
        if not stripped:
            i += 1
            continue

        # 2. 水平分隔线 ---
        if HR_RE.match(stripped):
            i += 1
            continue

        # 3. 标题 # ~ ######
        m_head = HEADING_RE.match(raw_line)
        if m_head:
            hashes, title_text = m_head.group(1), m_head.group(2).strip()
            level = min(len(hashes), 4)
            style_name = f"Heading {level}"
            p = doc.add_paragraph(style=style_name if style_name in doc.styles else None)
            font_size = {1: 16.0, 2: 14.0, 3: 12.0, 4: 12.0}.get(level, 12.0)
            _render_inline_runs(
                p,
                title_text,
                highlight_placeholders=placeholder_highlight,
                base_bold=True,
                east_asia_font="黑体",
                size_pt=font_size,
            )
            last_p_elem = p._p
            stats["headings"] += 1
            i += 1
            continue

        # 4. 引用块 > （区分范本批注块与普通引用块）
        if stripped.startswith(">"):
            quote_lines: list[str] = []
            while i < n and lines[i].strip().startswith(">"):
                quote_lines.append(lines[i])
                i += 1

            first_q_line = quote_lines[0].strip()
            is_note_block = bool(NOTE_BLOCK_HEADER_RE.match(first_q_line))

            if is_note_block:
                stats["note_blocks_detected"] += 1
                if strip_notes:
                    stats["note_blocks_stripped"] += 1
                    continue
                if notes_as_comments:
                    note_text = _clean_note_block_text(quote_lines)
                    if last_p_elem is None:
                        p_dummy = doc.add_paragraph()
                        last_p_elem = p_dummy._p
                    para_marker_counter += 1
                    # w14:paraId 是 ST_LongHexNumber，必须使用 8 位十六进制；
                    # 使用临时合法值，写入 comments.xml 后会从正文移除。
                    marker_id = f"{para_marker_counter:08X}"
                    last_p_elem.set(w14_tag("paraId"), marker_id)
                    pending_comments_for_para.append((marker_id, note_text))
                    stats["note_blocks_as_comments"] += 1
                    continue

            # 普通引用块，或未启用 strip/notes-as-comments 时的批注块
            stats["regular_Format_quotes"] += 1
            q_style = "Quote" if "Quote" in doc.styles else None
            for q_line in quote_lines:
                q_content = re.sub(r"^\s*>\s?", "", q_line).strip()
                if not q_content:
                    continue
                p = doc.add_paragraph(style=q_style)
                _render_inline_runs(
                    p,
                    q_content,
                    highlight_placeholders=placeholder_highlight,
                    base_bold=False,
                    east_asia_font="仿宋",
                    size_pt=10.5,
                )
                last_p_elem = p._p
            continue

        # 5. 表格（当前行含 | 且下一行为对齐分隔行 | --- |）
        if "|" in stripped and (i + 1 < n) and _is_table_separator_line(lines[i + 1]):
            header_cells = _split_table_row(lines[i])
            alignments = _parse_table_alignments(lines[i + 1])
            num_cols = max(len(header_cells), len(alignments))
            i += 2

            body_rows: list[list[str]] = []
            while i < n and "|" in lines[i].strip() and not HR_RE.match(lines[i].strip()):
                if _is_table_separator_line(lines[i]):
                    i += 1
                    continue
                row_cells = _split_table_row(lines[i])
                body_rows.append(row_cells)
                i += 1

            if num_cols >= 6:
                _ensure_landscape_section_for_wide_table(doc)
                stats["wide_tables"] += 1
                wide_table_active = True

            table = doc.add_table(rows=1 + len(body_rows), cols=num_cols)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            if "Table Grid" in doc.styles:
                table.style = "Table Grid"
            _set_table_borders(table)

            # 填充表头行
            hdr_row = table.rows[0]
            for c_idx in range(num_cols):
                cell = hdr_row.cells[c_idx]
                _set_cell_shading(cell, "EAEAEA")
                cell_txt = header_cells[c_idx] if c_idx < len(header_cells) else ""
                sub_lines = re.split(r"<br\s*/?>", cell_txt, flags=re.IGNORECASE)
                for s_idx, s_line in enumerate(sub_lines):
                    cp = cell.paragraphs[0] if s_idx == 0 else cell.add_paragraph()
                    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    _render_inline_runs(
                        cp,
                        s_line.strip(),
                        highlight_placeholders=placeholder_highlight,
                        base_bold=True,
                        east_asia_font="黑体",
                        size_pt=10.0,
                    )

            # 填充数据行
            for r_idx, r_cells in enumerate(body_rows, start=1):
                tr_row = table.rows[r_idx]
                for c_idx in range(num_cols):
                    cell = tr_row.cells[c_idx]
                    cell_txt = r_cells[c_idx] if c_idx < len(r_cells) else ""
                    col_align = (
                        alignments[c_idx] if c_idx < len(alignments) else WD_ALIGN_PARAGRAPH.LEFT
                    )
                    sub_lines = re.split(r"<br\s*/?>", cell_txt, flags=re.IGNORECASE)
                    for s_idx, s_line in enumerate(sub_lines):
                        cp = cell.paragraphs[0] if s_idx == 0 else cell.add_paragraph()
                        cp.alignment = col_align
                        _render_inline_runs(
                            cp,
                            s_line.strip(),
                            highlight_placeholders=placeholder_highlight,
                            base_bold=False,
                            east_asia_font="宋体",
                            size_pt=10.0,
                        )
                        last_p_elem = cp._p

            stats["tables"] += 1
            continue

        # 6. 复选框列表 - [ ] / - [x]
        m_chk = CHECKBOX_RE.match(raw_line)
        if m_chk:
            checked = m_chk.group(1).lower() == "x"
            item_text = m_chk.group(2).strip()
            box_symbol = "☑ " if checked else "☐ "
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Pt(14)
            _render_inline_runs(
                p,
                box_symbol + item_text,
                highlight_placeholders=placeholder_highlight,
                base_bold=False,
                east_asia_font="宋体",
                size_pt=12.0,
            )
            last_p_elem = p._p
            stats["paragraphs"] += 1
            i += 1
            continue

        # 7. 无序列表 * / -
        m_ul = UNORDERED_LIST_RE.match(raw_line)
        if m_ul:
            item_text = m_ul.group(1).strip()
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Pt(14)
            _render_inline_runs(
                p,
                "• " + item_text,
                highlight_placeholders=placeholder_highlight,
                base_bold=False,
                east_asia_font="宋体",
                size_pt=12.0,
            )
            last_p_elem = p._p
            stats["paragraphs"] += 1
            i += 1
            continue

        # 8. 普通段落或有序列表行（保留原文行首序号以便条款定位）
        p = doc.add_paragraph()
        _render_inline_runs(
            p,
            stripped,
            highlight_placeholders=placeholder_highlight,
            base_bold=False,
            east_asia_font="宋体",
            size_pt=12.0,
        )
        last_p_elem = p._p
        stats["paragraphs"] += 1
        i += 1

    out_path.parent.mkdir(parents=True, exist_ok=True)

    # 所有分支先写同目录临时文件、校验通过后再替换目标，避免中断后留下坏 docx。
    fd, tmp_output_name = tempfile.mkstemp(
        prefix=f".{out_path.name}.", suffix=".docx", dir=str(out_path.parent)
    )
    os.close(fd)
    tmp_output_path = Path(tmp_output_name)
    tmp_docx_path: Path | None = None
    try:
        if not pending_comments_for_para:
            doc.save(str(tmp_output_path))
            sanitize_docx_file(tmp_output_path, unwrap_wps_ac=False)
        else:
            with tempfile.NamedTemporaryFile(
                suffix=".docx", prefix=".md-to-docx-", delete=False, dir=str(out_path.parent)
            ) as tmp_f:
                tmp_docx_path = Path(tmp_f.name)
            doc.save(str(tmp_docx_path))
            pkg = DocxPackage(tmp_docx_path)
            id_alloc = pkg.create_id_allocator()
            doc_root = pkg.get_xml("word/document.xml")
            date_iso = dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

            if doc_root is not None:
                para_by_marker: dict[str, Any] = {}
                for p_el in doc_root.iter(w_tag("p")):
                    pid = p_el.get(w14_tag("paraId"))
                    if pid:
                        para_by_marker[pid] = p_el

                for marker_id, note_text in pending_comments_for_para:
                    target_p = para_by_marker.get(marker_id)
                    if target_p is not None:
                        cid = id_alloc.next_id()
                        add_comment_entry(pkg, cid, comment_author, date_iso, note_text)
                        attach_comment_to_whole_paragraph(target_p, cid)

                for p_el in para_by_marker.values():
                    if p_el.get(w14_tag("paraId")) in {
                        marker for marker, _ in pending_comments_for_para
                    }:
                        del p_el.attrib[w14_tag("paraId")]

                pkg.mark_modified("word/document.xml")

            pkg.save(tmp_output_path)

        validation = validate_docx_file(tmp_output_path, strict=True)
        if not validation["valid"]:
            raise ValueError(f"生成的 .docx 校验失败: {'; '.join(validation['errors'])}")
        os.replace(tmp_output_path, out_path)
        stats["docx_validation"] = validation
    finally:
        tmp_output_path.unlink(missing_ok=True)
        if tmp_docx_path is not None:
            tmp_docx_path.unlink(missing_ok=True)

    return {
        "input": str(md_path),
        "output": str(out_path),
        "reference": str(ref_path) if ref_path else None,
        "stats": stats,
    }


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8")

    parser = argparse.ArgumentParser(
        description="将合同范本或审查意见书 Markdown 渲染为集团公文格式 .docx"
    )
    parser.add_argument("input", help="输入 Markdown 文件路径")
    parser.add_argument("--out", required=True, help="输出 .docx 文件路径")
    parser.add_argument(
        "--reference",
        default=str(DEFAULT_REFERENCE),
        help="参考样式模板 .docx 路径（默认 assets/docx/reference.docx）",
    )
    parser.add_argument(
        "--encoding",
        default=None,
        help="输入 Markdown 编码（默认自动识别 UTF-8/UTF-16/GB18030）",
    )
    parser.add_argument(
        "--placeholder-highlight",
        dest="placeholder_highlight",
        action="store_true",
        default=True,
        help="将 [●] 及 `[说明]` 占位符设为黄色高亮（默认启用）",
    )
    parser.add_argument(
        "--no-placeholder-highlight",
        dest="placeholder_highlight",
        action="store_false",
        help="关闭占位符黄色高亮",
    )
    note_group = parser.add_mutually_exclusive_group()
    note_group.add_argument(
        "--notes-as-comments",
        action="store_true",
        help="将 `> **【…批注…】**` 引用块转换为挂载在前一段落上的 Word 批注",
    )
    note_group.add_argument(
        "--strip-notes",
        action="store_true",
        help="剥离删除所有 `> **【…批注…】**` 批注块",
    )
    args = parser.parse_args(argv)

    in_path = Path(args.input)
    if not in_path.is_file():
        print(
            json.dumps({"error": f"输入 Markdown 文件不存在: {in_path}"}, ensure_ascii=False),
            file=sys.stderr,
        )
        return 1

    try:
        res = convert_md_to_docx(
            input_md=in_path,
            out_docx=args.out,
            reference_docx=args.reference,
            placeholder_highlight=args.placeholder_highlight,
            notes_as_comments=args.notes_as_comments,
            strip_notes=args.strip_notes,
            input_encoding=args.encoding,
        )
    except Exception as exc:
        print(json.dumps({"error": f"渲染 docx 失败: {exc}"}, ensure_ascii=False), file=sys.stderr)
        return 1

    print(json.dumps(res, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
