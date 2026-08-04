# -*- coding: utf-8 -*-
"""
pdf_tool —— PDF 常用操作工具箱（独立脚本，子命令式）。

子命令：
  merge          合并多个 PDF
  split          按页/页码区间拆分
  rotate         旋转页面
  watermark      加文字水印（支持中文）
  encrypt        加密码
  extract-text   提取文字
  extract-tables 提取表格为 CSV

依赖：pypdf（合并/拆分/旋转/加密/水印底层）、reportlab（水印文字）、pdfplumber（提表格）。
均为沙箱镜像已预装。

示例：
  python3 pdf_tool.py merge a.pdf b.pdf -o /workspace/out.pdf
  python3 pdf_tool.py split in.pdf -o /workspace/parts --ranges 1-3,4-6
  python3 pdf_tool.py rotate in.pdf --angle 90 -o /workspace/r.pdf
  python3 pdf_tool.py watermark in.pdf --text 机密 -o /workspace/wm.pdf
  python3 pdf_tool.py encrypt in.pdf --password 1234 -o /workspace/enc.pdf
  python3 pdf_tool.py extract-text in.pdf -o /workspace/text.txt
  python3 pdf_tool.py extract-tables in.pdf -o /workspace/tables.csv
"""
from __future__ import annotations

import csv
import io
import os

from pypdf import PdfReader, PdfWriter


def _parse_ranges(spec: str, n: int) -> list[list[int]]:
    """'1-3,4-6,8' -> [[0,1,2],[3,4,5],[7]]（0 基）。"""
    groups = []
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            groups.append([i - 1 for i in range(int(a), int(b) + 1) if 1 <= i <= n])
        else:
            i = int(part)
            if 1 <= i <= n:
                groups.append([i - 1])
    return groups


def _parse_pages(spec: str | None, n: int) -> list[int]:
    if not spec:
        return list(range(n))
    out: list[int] = []
    for part in spec.split(","):
        part = part.strip()
        if "-" in part:
            a, b = part.split("-", 1)
            out += [i - 1 for i in range(int(a), int(b) + 1) if 1 <= i <= n]
        elif part:
            i = int(part)
            if 1 <= i <= n:
                out.append(i - 1)
    return out


def merge(inputs, out):
    w = PdfWriter()
    for f in inputs:
        for p in PdfReader(f).pages:
            w.add_page(p)
    _save(w, out)
    return [out]


def split(src, out_dir, ranges=None):
    os.makedirs(out_dir, exist_ok=True)
    reader = PdfReader(src)
    n = len(reader.pages)
    base = os.path.splitext(os.path.basename(src))[0]
    saved = []
    groups = _parse_ranges(ranges, n) if ranges else [[i] for i in range(n)]
    for gi, g in enumerate(groups, 1):
        w = PdfWriter()
        for idx in g:
            w.add_page(reader.pages[idx])
        out = os.path.join(out_dir, f"{base}_part{gi:02d}.pdf")
        _save(w, out)
        saved.append(out)
    return saved


def rotate(src, out, angle=90, pages=None):
    reader = PdfReader(src)
    target = set(_parse_pages(pages, len(reader.pages)))
    w = PdfWriter()
    for i, p in enumerate(reader.pages):
        if i in target:
            p.rotate(angle)
        w.add_page(p)
    _save(w, out)
    return [out]


def _make_watermark_overlay(text, width, height):
    from reportlab.lib import colors
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.pdfgen import canvas
    try:
        pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
    except Exception:
        pass
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(width, height))
    fs = max(18, int(width / 16))  # 字号随页面宽度缩放
    c.setFont("STSong-Light", fs)
    c.setFillColor(colors.grey, alpha=0.18)
    c.saveState()
    c.translate(width / 2, height / 2)
    c.rotate(45)
    step_x = max(160, int(fs * len(text) * 1.4) + 80)
    step_y = max(110, fs * 3)
    for x in range(-int(width), int(width), step_x):
        for y in range(-int(height), int(height), step_y):
            c.drawCentredString(x, y, text)
    c.restoreState()
    c.save()
    buf.seek(0)
    return PdfReader(buf).pages[0]


def watermark(src, out, text):
    reader = PdfReader(src)
    w = PdfWriter()
    for p in reader.pages:
        ov = _make_watermark_overlay(text, float(p.mediabox.width), float(p.mediabox.height))
        p.merge_page(ov)
        w.add_page(p)
    _save(w, out)
    return [out]


def encrypt(src, out, password):
    reader = PdfReader(src)
    w = PdfWriter()
    for p in reader.pages:
        w.add_page(p)
    w.encrypt(password)
    _save(w, out)
    return [out]


def extract_text(src, out, pages=None):
    reader = PdfReader(src)
    idxs = _parse_pages(pages, len(reader.pages))
    parts = []
    for i in idxs:
        parts.append(f"=== 第 {i + 1} 页 ===\n" + (reader.pages[i].extract_text() or ""))
    text = "\n\n".join(parts)
    os.makedirs(os.path.dirname(os.path.abspath(out)) or ".", exist_ok=True)
    with open(out, "w", encoding="utf-8") as f:
        f.write(text)
    return [out]


def extract_tables(src, out, pages=None):
    import pdfplumber
    os.makedirs(os.path.dirname(os.path.abspath(out)) or ".", exist_ok=True)
    rows_written = 0
    with pdfplumber.open(src) as pdf, open(out, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        idxs = _parse_pages(pages, len(pdf.pages))
        for i in idxs:
            for ti, table in enumerate(pdf.pages[i].extract_tables(), 1):
                writer.writerow([f"# 第{i + 1}页 表{ti}"])
                for row in table:
                    writer.writerow(["" if c is None else c for c in row])
                    rows_written += 1
                writer.writerow([])
    print(f"[info] 共写出 {rows_written} 行表格数据", flush=True)
    return [out]


def _save(writer, out):
    out = os.path.abspath(out)
    os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
    with open(out, "wb") as f:
        writer.write(f)


def _main(argv=None) -> int:
    import argparse
    ap = argparse.ArgumentParser(prog="pdf_tool", description="PDF 工具箱")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("merge"); p.add_argument("inputs", nargs="+"); p.add_argument("-o", "--out", required=True)
    p = sub.add_parser("split"); p.add_argument("src"); p.add_argument("-o", "--out", required=True); p.add_argument("--ranges", default=None)
    p = sub.add_parser("rotate"); p.add_argument("src"); p.add_argument("-o", "--out", required=True); p.add_argument("--angle", type=int, default=90); p.add_argument("--pages", default=None)
    p = sub.add_parser("watermark"); p.add_argument("src"); p.add_argument("-o", "--out", required=True); p.add_argument("--text", required=True)
    p = sub.add_parser("encrypt"); p.add_argument("src"); p.add_argument("-o", "--out", required=True); p.add_argument("--password", required=True)
    p = sub.add_parser("extract-text"); p.add_argument("src"); p.add_argument("-o", "--out", required=True); p.add_argument("--pages", default=None)
    p = sub.add_parser("extract-tables"); p.add_argument("src"); p.add_argument("-o", "--out", required=True); p.add_argument("--pages", default=None)

    a = ap.parse_args(argv)
    if a.cmd == "merge":
        files = merge(a.inputs, a.out)
    elif a.cmd == "split":
        files = split(a.src, a.out, a.ranges)
    elif a.cmd == "rotate":
        files = rotate(a.src, a.out, a.angle, a.pages)
    elif a.cmd == "watermark":
        files = watermark(a.src, a.out, a.text)
    elif a.cmd == "encrypt":
        files = encrypt(a.src, a.out, a.password)
    elif a.cmd == "extract-text":
        files = extract_text(a.src, a.out, a.pages)
    elif a.cmd == "extract-tables":
        files = extract_tables(a.src, a.out, a.pages)
    else:
        ap.error("未知子命令")
    for f in files:
        print(f)
    print(f"[OK] {a.cmd} 完成，共 {len(files)} 个输出", flush=True)
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
