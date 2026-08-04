# -*- coding: utf-8 -*-
"""
pdf2img —— 把 PDF 转为图片（独立脚本，可直接 python3 运行）。

两种能力：
  1) 默认           : 把每一页“渲染(rasterize)”成位图（最常用，所见即所得）
  2) --extract      : 抽取页面中“内嵌的原始图片对象”（保留原图，不含文字/矢量）

依赖：PyMuPDF（沙箱镜像已预装）。纯 Python，无需 Poppler/Ghostscript/ImageMagick。
"""
from __future__ import annotations

import os
from typing import Iterable, Sequence

try:
    import fitz  # PyMuPDF
except ImportError as e:  # pragma: no cover
    raise ImportError("需要 PyMuPDF：pip install pymupdf（沙箱镜像应已预装）") from e


def _parse_pages(pages, page_count: int) -> list[int]:
    if pages is None:
        return list(range(page_count))
    result: list[int] = []

    def add(one_based: int):
        idx = one_based - 1
        if 0 <= idx < page_count and idx not in result:
            result.append(idx)

    if isinstance(pages, str):
        for part in pages.split(","):
            part = part.strip()
            if not part:
                continue
            if "-" in part:
                a, b = part.split("-", 1)
                for n in range(int(a), int(b) + 1):
                    add(n)
            else:
                add(int(part))
    elif isinstance(pages, Iterable):
        for n in pages:
            add(int(n))
    else:
        raise TypeError("pages 应为 None / str / 可迭代的 1 基页码")
    return result


def pdf_to_images(pdf_path, out_dir=None, dpi=144, fmt="png", pages=None,
                  prefix="page_", zoom=None, alpha=False) -> list[str]:
    """把 PDF 各页渲染为图片，返回生成文件的绝对路径列表。"""
    if not os.path.isfile(pdf_path):
        raise FileNotFoundError(pdf_path)
    if out_dir is None:
        base = os.path.splitext(os.path.basename(pdf_path))[0]
        out_dir = os.path.join(os.path.dirname(os.path.abspath(pdf_path)), f"{base}_images")
    os.makedirs(out_dir, exist_ok=True)

    fmt = fmt.lower().lstrip(".")
    if fmt in ("jpg", "jpeg"):
        alpha = False  # JPEG 不支持透明通道，避免保存报错
    factor = zoom if zoom is not None else (dpi / 72.0)
    mat = fitz.Matrix(factor, factor)

    saved: list[str] = []
    with fitz.open(pdf_path) as doc:
        idxs = _parse_pages(pages, doc.page_count)
        width = max(3, len(str(doc.page_count)))
        for idx in idxs:
            pix = doc[idx].get_pixmap(matrix=mat, alpha=alpha)
            out = os.path.join(out_dir, f"{prefix}{idx + 1:0{width}d}.{fmt}")
            pix.save(out)
            saved.append(os.path.abspath(out))
    return saved


def extract_embedded_images(pdf_path, out_dir=None, pages=None, prefix="img_",
                            min_width=0, min_height=0) -> list[str]:
    """抽取 PDF 中“内嵌的原始图片对象”。返回已保存路径列表。"""
    if not os.path.isfile(pdf_path):
        raise FileNotFoundError(pdf_path)
    if out_dir is None:
        base = os.path.splitext(os.path.basename(pdf_path))[0]
        out_dir = os.path.join(os.path.dirname(os.path.abspath(pdf_path)), f"{base}_embedded")
    os.makedirs(out_dir, exist_ok=True)

    saved: list[str] = []
    seen: set[int] = set()
    with fitz.open(pdf_path) as doc:
        idxs = _parse_pages(pages, doc.page_count)
        pw = max(3, len(str(doc.page_count)))
        for idx in idxs:
            for i, info in enumerate(doc[idx].get_images(full=True)):
                xref = info[0]
                if xref in seen:
                    continue
                seen.add(xref)
                pix = fitz.Pixmap(doc, xref)
                if pix.width < min_width or pix.height < min_height:
                    continue
                if pix.n - pix.alpha >= 4:  # CMYK -> RGB
                    pix = fitz.Pixmap(fitz.csRGB, pix)
                out = os.path.join(out_dir, f"{prefix}p{idx + 1:0{pw}d}_{i + 1:02d}.png")
                pix.save(out)
                saved.append(os.path.abspath(out))
    return saved


def _main(argv: Sequence[str] | None = None) -> int:
    import argparse

    ap = argparse.ArgumentParser(
        prog="pdf2img",
        description="把 PDF 转为图片（默认整页渲染；--extract 抽取内嵌原图）")
    ap.add_argument("pdf", help="源 PDF 路径（建议放在 /workspace 下）")
    ap.add_argument("-o", "--out", default=None, help="输出目录（建议 /workspace/...；默认 <pdf名>_images）")
    ap.add_argument("--dpi", type=int, default=144, help="分辨率，默认 144（150~200 适合预览）")
    ap.add_argument("--zoom", type=float, default=None, help="缩放倍数（覆盖 --dpi）")
    ap.add_argument("--format", default="png", help="图片格式 png/jpg/webp...，默认 png")
    ap.add_argument("--pages", default=None, help='页码，如 "1-3,5,8-10"，默认全部')
    ap.add_argument("--prefix", default="page_", help="文件名前缀")
    ap.add_argument("--alpha", action="store_true", help="保留透明通道")
    ap.add_argument("--extract", action="store_true", help="改为抽取内嵌原始图片对象")
    ap.add_argument("--min-w", type=int, default=0, help="抽取模式：最小宽度像素")
    ap.add_argument("--min-h", type=int, default=0, help="抽取模式：最小高度像素")
    args = ap.parse_args(argv)

    if args.extract:
        files = extract_embedded_images(args.pdf, args.out, pages=args.pages,
                                        prefix="img_", min_width=args.min_w, min_height=args.min_h)
    else:
        files = pdf_to_images(args.pdf, args.out, dpi=args.dpi, fmt=args.format,
                              pages=args.pages, prefix=args.prefix, zoom=args.zoom, alpha=args.alpha)
    for f in files:
        print(f)
    print(f"[OK] 共 {len(files)} 个文件", flush=True)
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
