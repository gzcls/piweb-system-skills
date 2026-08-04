# -*- coding: utf-8 -*-
"""
report2pdf —— 用结构化数据(JSON spec)生成可打印 PDF 报告（独立脚本）。

纯 Python（reportlab），中文用内置 CID 字体 STSong-Light，
*不依赖* MS Word / LibreOffice / 浏览器 / 系统字体，跨平台、可服务器部署。

用法：
    python3 report2pdf.py /workspace/spec.json -o /workspace/report.pdf

spec 字段说明见同 skill 的 references/spec.md。
"""
from __future__ import annotations

import json
import os
from typing import Any

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A3, A4, LETTER
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.platypus import (
    HRFlowable, ListFlowable, ListItem, Paragraph,
    SimpleDocTemplate, Spacer, Table, TableStyle,
)

_FONT = "STSong-Light"
_FONT_REGISTERED = False


def _ensure_font():
    global _FONT_REGISTERED
    if not _FONT_REGISTERED:
        pdfmetrics.registerFont(UnicodeCIDFont(_FONT))
        pdfmetrics.registerFontFamily(_FONT, normal=_FONT, bold=_FONT, italic=_FONT, boldItalic=_FONT)
        _FONT_REGISTERED = True


_PAGE = {"A4": A4, "A3": A3, "LETTER": LETTER}
_ALIGN = {"left": TA_LEFT, "center": TA_CENTER, "right": TA_RIGHT}
_HEADER_BG = colors.HexColor("#1F3B5C")
_LABEL_BG = colors.HexColor("#E8EDF3")
_ZEBRA_BG = colors.HexColor("#F2F5F9")
_HEAD_FG = colors.HexColor("#1F3B5C")
_GRID = colors.HexColor("#9AA7B4")
_LEVEL_COLOR = {
    "重要": colors.HexColor("#C00000"), "高": colors.HexColor("#C00000"),
    "中": colors.HexColor("#B56A00"), "低": colors.HexColor("#1E6B2E"),
    "提示": colors.HexColor("#555555"),
}


def _styles():
    _ensure_font()
    ss = getSampleStyleSheet()
    base = ss["Normal"]
    base.fontName = _FONT
    return {
        "title": ParagraphStyle("title", parent=base, fontName=_FONT, fontSize=20, leading=26,
                                alignment=TA_CENTER, textColor=_HEAD_FG, spaceAfter=2),
        "subtitle": ParagraphStyle("subtitle", parent=base, fontName=_FONT, fontSize=10, leading=14,
                                   alignment=TA_CENTER, textColor=colors.HexColor("#555555"), spaceAfter=10),
        "heading": ParagraphStyle("heading", parent=base, fontName=_FONT, fontSize=13, leading=18,
                                  textColor=_HEAD_FG, spaceBefore=12, spaceAfter=4),
        "subhead": ParagraphStyle("subhead", parent=base, fontName=_FONT, fontSize=11, leading=15,
                                  spaceBefore=8, spaceAfter=2),
        "body": ParagraphStyle("body", parent=base, fontName=_FONT, fontSize=10.5, leading=15.5, spaceAfter=4),
        "cell": ParagraphStyle("cell", parent=base, fontName=_FONT, fontSize=9, leading=12),
        "cell_h": ParagraphStyle("cell_h", parent=base, fontName=_FONT, fontSize=9, leading=12,
                                 textColor=colors.white, alignment=TA_CENTER),
    }


def _make_numbered_canvas(enable: bool):
    from reportlab.pdfgen import canvas

    class C(canvas.Canvas):
        def __init__(self, *a, **kw):
            super().__init__(*a, **kw)
            self._saved = []

        def showPage(self):
            self._saved.append(dict(self.__dict__))
            self._startPage()

        def save(self):
            total = len(self._saved)
            for st in self._saved:
                self.__dict__.update(st)
                if enable:
                    self.setFont(_FONT, 9)
                    self.setFillColor(colors.HexColor("#555555"))
                    self.drawCentredString(self._pagesize[0] / 2.0, 1.0 * cm,
                                           f"第 {self._pageNumber} 页 / 共 {total} 页")
                super().showPage()
            super().save()

    return C


def _mk(text, style):
    return Paragraph(str(text).replace("\n", "<br/>"), style)


def _build_table(block, styles, avail):
    headers = block.get("headers", [])
    rows = block.get("rows", [])
    ncol = len(headers) if headers else (len(rows[0]) if rows else 1)

    widths_cm = block.get("widths_cm")
    if widths_cm and len(widths_cm) == ncol:
        col_widths = [w * cm for w in widths_cm]
        scale = avail / sum(col_widths)
        if scale < 1:
            col_widths = [w * scale for w in col_widths]
    else:
        col_widths = [avail / ncol] * ncol

    aligns = block.get("aligns") or ["left"] * ncol
    data = []
    if headers:
        data.append([_mk(h, styles["cell_h"]) for h in headers])
    for row in rows:
        line = []
        for ci, val in enumerate(row):
            a = aligns[ci] if ci < len(aligns) else "left"
            st = ParagraphStyle(f"c{ci}", parent=styles["cell"], alignment=_ALIGN.get(a, TA_LEFT))
            line.append(_mk(val, st))
        data.append(line)

    t = Table(data, colWidths=col_widths, repeatRows=1 if headers else 0, hAlign="CENTER")
    ts = [
        ("GRID", (0, 0), (-1, -1), 0.5, _GRID),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ]
    body_start = 0
    if headers:
        ts.append(("BACKGROUND", (0, 0), (-1, 0), _HEADER_BG))
        body_start = 1
        for r in range(len(rows)):
            if r % 2 == 1:
                ts.append(("BACKGROUND", (0, r + 1), (-1, r + 1), _ZEBRA_BG))
    lvl_col = block.get("level_col")
    if lvl_col is not None:
        for r, row in enumerate(rows):
            if lvl_col < len(row):
                c = _LEVEL_COLOR.get(str(row[lvl_col]).strip())
                if c:
                    ts.append(("TEXTCOLOR", (lvl_col, r + body_start), (lvl_col, r + body_start), c))
    t.setStyle(TableStyle(ts))
    return t


def build_report(spec: dict, out_path: str) -> str:
    styles = _styles()
    page = _PAGE.get(str(spec.get("page_size", "A4")).upper(), A4)
    m = spec.get("margins_cm", [2.2, 2.0, 2.2, 2.0])
    top, bottom, left, right = [x * cm for x in m]

    out_path = os.path.abspath(out_path)
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)

    doc = SimpleDocTemplate(out_path, pagesize=page, topMargin=top, bottomMargin=bottom,
                            leftMargin=left, rightMargin=right,
                            title=spec.get("title", "报告"), author="data-to-pdf")
    avail = page[0] - left - right

    story: list[Any] = []
    if spec.get("title"):
        story.append(_mk(spec["title"], styles["title"]))
    if spec.get("subtitle"):
        story.append(_mk(spec["subtitle"], styles["subtitle"]))

    info = spec.get("info_table")
    if info:
        t = _build_table({"headers": info.get("headers", []), "rows": info.get("rows", []),
                          "widths_cm": info.get("widths_cm"), "aligns": info.get("aligns")}, styles, avail)
        label_cols = set(info.get("label_cols", []))
        if label_cols:
            start_row = 1 if info.get("headers") else 0
            t.setStyle(TableStyle([("BACKGROUND", (c, start_row), (c, -1), _LABEL_BG) for c in label_cols]))
        story.append(t)
        story.append(Spacer(1, 6))

    for sec in spec.get("sections", []):
        if sec.get("heading"):
            story.append(_mk(sec["heading"], styles["heading"]))
            story.append(HRFlowable(width="100%", thickness=0.8, color=_HEAD_FG, spaceBefore=1, spaceAfter=6))
        for block in sec.get("blocks", []):
            bt = block.get("type")
            if bt == "para":
                story.append(_mk(block.get("text", ""), styles["body"]))
            elif bt == "subhead":
                story.append(_mk(block.get("text", ""), styles["subhead"]))
            elif bt == "list":
                items = [ListItem(_mk(x, styles["body"]), leftIndent=12) for x in block.get("items", [])]
                story.append(ListFlowable(items, bulletType="1" if block.get("ordered") else "bullet",
                                          bulletFontName=_FONT, leftIndent=16))
                story.append(Spacer(1, 3))
            elif bt == "table":
                story.append(_build_table(block, styles, avail))
                story.append(Spacer(1, 4))
            elif bt == "spacer":
                story.append(Spacer(1, float(block.get("h", 6))))
            elif bt == "note":
                st = ParagraphStyle("note", parent=styles["body"], fontSize=9,
                                    textColor=colors.HexColor("#555555"))
                story.append(_mk(block.get("text", ""), st))

    doc.build(story, canvasmaker=_make_numbered_canvas(bool(spec.get("footer_pagenum", True))))
    return out_path


def _main(argv=None) -> int:
    import argparse

    ap = argparse.ArgumentParser(
        prog="report2pdf",
        description="用 JSON spec 生成可打印 PDF 报告（纯 Python，无 Office/浏览器依赖）")
    ap.add_argument("spec", help="报告内容 JSON 文件（建议放 /workspace）")
    ap.add_argument("-o", "--out", required=True, help="输出 PDF 路径（建议 /workspace/xxx.pdf）")
    args = ap.parse_args(argv)

    with open(args.spec, "r", encoding="utf-8") as f:
        spec = json.load(f)
    out = build_report(spec, args.out)
    print(f"[OK] 已生成 {out}", flush=True)
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
