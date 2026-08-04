# -*- coding: utf-8 -*-
"""
docx_build —— 用结构化数据(JSON spec)生成 Word 文档（独立脚本）。

python-docx 实现，中文用 宋体/黑体。spec 与 data-to-pdf 基本一致，
同一份 spec 可分别生成 PDF 或 Word。

用法：
    python3 docx_build.py /workspace/spec.json -o /workspace/report.docx
"""
from __future__ import annotations

import json
import os

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

BODY_CN, HEAD_CN, ASCII_F = "宋体", "黑体", "Times New Roman"
HEAD_FG = RGBColor(0x1F, 0x3B, 0x5C)
_ALIGN = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
          "right": WD_ALIGN_PARAGRAPH.RIGHT}
_LEVEL = {"重要": (0xC0, 0, 0), "高": (0xC0, 0, 0), "中": (0xB5, 0x6A, 0),
          "低": (0x1E, 0x6B, 0x2E), "提示": (0x55, 0x55, 0x55)}
_PAGE = {"A4": (Cm(21.0), Cm(29.7)), "A3": (Cm(29.7), Cm(42.0)),
         "LETTER": (Cm(21.59), Cm(27.94))}


def _run(run, size=10.5, bold=False, cn=BODY_CN, color=None):
    run.font.name = ASCII_F
    run.font.size = Pt(size)
    run.font.bold = bold
    run._element.rPr.rFonts.set(qn("w:eastAsia"), cn)
    if color:
        run.font.color.rgb = RGBColor(*color)


def _add_rich(p, text, size=10.5, color=None):
    """支持 <b>…</b> 加粗与 \n 换行（与 data-to-pdf 的 <b> 语法一致）。"""
    import re
    bold = False
    for seg in re.split(r"(<b>|</b>)", str(text)):
        if seg == "<b>":
            bold = True
            continue
        if seg == "</b>":
            bold = False
            continue
        for i, ln in enumerate(seg.split("\n")):
            if i > 0:
                p.add_run().add_break()
            if ln:
                _run(p.add_run(ln), size=size, bold=bold, color=color)
    return p


def _fixed_layout(table):
    """固定表格布局，使设置的列宽在 Word 中真正生效。"""
    tblPr = table._tbl.tblPr
    el = OxmlElement("w:tblLayout")
    el.set(qn("w:type"), "fixed")
    tblPr.append(el)


def _shade(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), hexcolor)
    tcPr.append(shd)


def _cell(cell, text, size=9, bold=False, cn=BODY_CN, align="left", color=None):
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    for i, line in enumerate(str(text).split("\n")):
        p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        p.alignment = _ALIGN.get(align, WD_ALIGN_PARAGRAPH.LEFT)
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        _run(p.add_run(line), size=size, bold=bold, cn=cn, color=color)


def _table(doc, block, avail_cm):
    headers = block.get("headers", [])
    rows = block.get("rows", [])
    ncol = len(headers) if headers else (len(rows[0]) if rows else 1)
    aligns = block.get("aligns") or ["left"] * ncol
    level_col = block.get("level_col")

    t = doc.add_table(rows=0, cols=ncol)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.style = "Table Grid"
    t.autofit = False
    t.allow_autofit = False
    if headers:
        cells = t.add_row().cells
        for i, h in enumerate(headers):
            _cell(cells[i], h, bold=True, cn=HEAD_CN, align="center", color=(0xFF, 0xFF, 0xFF))
            _shade(cells[i], "1F3B5C")
    for ri, row in enumerate(rows):
        cells = t.add_row().cells
        for ci, val in enumerate(row):
            col = None
            if level_col == ci:
                col = _LEVEL.get(str(val).strip())
            _cell(cells[ci], val, align=aligns[ci] if ci < len(aligns) else "left",
                  color=col, bold=bool(col))
            if headers and ri % 2 == 1:
                _shade(cells[ci], "F2F5F9")
    widths = block.get("widths_cm")
    if widths and len(widths) == ncol:
        scale = avail_cm / sum(widths)
        if scale < 1:
            widths = [w * scale for w in widths]
        _fixed_layout(t)
        for ci, w in enumerate(widths):
            for r in t.rows:
                r.cells[ci].width = Cm(w)
    return t


def _heading(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    _run(p.add_run(text), size=13, bold=True, cn=HEAD_CN, color=(0x1F, 0x3B, 0x5C))
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    b = OxmlElement("w:bottom")
    b.set(qn("w:val"), "single")
    b.set(qn("w:sz"), "6")
    b.set(qn("w:space"), "3")
    b.set(qn("w:color"), "1F3B5C")
    pbdr.append(b)
    pPr.append(pbdr)


def _footer_pagenum(section):
    p = section.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    def fld(instr):
        r = p.add_run()
        b = OxmlElement("w:fldChar"); b.set(qn("w:fldCharType"), "begin")
        it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = instr
        e = OxmlElement("w:fldChar"); e.set(qn("w:fldCharType"), "end")
        r._r.append(b); r._r.append(it); r._r.append(e)
    _run(p.add_run("第 "), size=9)
    fld("PAGE")
    _run(p.add_run(" 页 / 共 "), size=9)
    fld("NUMPAGES")
    _run(p.add_run(" 页"), size=9)


def build_docx(spec: dict, out_path: str) -> str:
    doc = Document()
    sec = doc.sections[0]
    w, h = _PAGE.get(str(spec.get("page_size", "A4")).upper(), _PAGE["A4"])
    sec.page_width, sec.page_height = w, h
    m = spec.get("margins_cm", [2.2, 2.0, 2.2, 2.0])
    sec.top_margin, sec.bottom_margin = Cm(m[0]), Cm(m[1])
    sec.left_margin, sec.right_margin = Cm(m[2]), Cm(m[3])
    avail_cm = (w - sec.left_margin - sec.right_margin) / Cm(1)

    normal = doc.styles["Normal"]
    normal.font.name = ASCII_F
    normal.font.size = Pt(10.5)
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), BODY_CN)

    if spec.get("title"):
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _run(p.add_run(spec["title"]), size=20, bold=True, cn=HEAD_CN, color=(0x1F, 0x3B, 0x5C))
    if spec.get("subtitle"):
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _run(p.add_run(spec["subtitle"]), size=10, cn=BODY_CN, color=(0x55, 0x55, 0x55))

    info = spec.get("info_table")
    if info:
        t = _table(doc, {"headers": info.get("headers", []), "rows": info.get("rows", []),
                         "widths_cm": info.get("widths_cm"), "aligns": info.get("aligns")}, avail_cm)
        for c in info.get("label_cols", []):
            start = 1 if info.get("headers") else 0
            for r in t.rows[start:]:
                _shade(r.cells[c], "E8EDF3")

    for s in spec.get("sections", []):
        if s.get("heading"):
            _heading(doc, s["heading"])
        for block in s.get("blocks", []):
            bt = block.get("type")
            if bt == "para":
                _add_rich(doc.add_paragraph(), block.get("text", ""), 10.5)
            elif bt == "subhead":
                p = doc.add_paragraph(); _run(p.add_run(block.get("text", "")), size=11, bold=True, cn=HEAD_CN)
            elif bt == "list":
                ordered = block.get("ordered")
                for i, x in enumerate(block.get("items", []), 1):
                    p = doc.add_paragraph()
                    p.paragraph_format.left_indent = Pt(16)
                    prefix = f"{i}. " if ordered else "• "
                    _run(p.add_run(prefix + str(x)), size=10.5)
            elif bt == "table":
                _table(doc, block, avail_cm)
            elif bt == "note":
                _add_rich(doc.add_paragraph(), block.get("text", ""), 9, color=(0x55, 0x55, 0x55))
            elif bt == "pagebreak":
                doc.add_page_break()

    if spec.get("footer_pagenum", True):
        _footer_pagenum(sec)

    out_path = os.path.abspath(out_path)
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    doc.save(out_path)
    return out_path


def _main(argv=None) -> int:
    import argparse
    ap = argparse.ArgumentParser(prog="docx_build", description="用 JSON spec 生成 Word 文档")
    ap.add_argument("spec", help="报告内容 JSON（建议放 /workspace）")
    ap.add_argument("-o", "--out", required=True, help="输出 .docx 路径（建议 /workspace）")
    args = ap.parse_args(argv)
    with open(args.spec, "r", encoding="utf-8") as f:
        spec = json.load(f)
    print(f"[OK] 已生成 {build_docx(spec, args.out)}", flush=True)
    return 0


def _cli():
    import sys
    try:
        return _main()
    except SystemExit:
        raise
    except Exception as e:
        print(f"[ERROR] {type(e).__name__}: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(_cli())
