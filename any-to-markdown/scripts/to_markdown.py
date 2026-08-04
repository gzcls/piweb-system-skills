# -*- coding: utf-8 -*-
"""
to_markdown —— 把各种文档转成 Markdown（独立脚本）。

基于 markitdown，支持 Word/Excel/PPT/PDF/HTML/CSV/JSON/图片 等 → Markdown 文本，
便于把文档内容喂给大模型阅读、检索、再加工。

依赖：markitdown[all]（沙箱镜像已预装）。

用法：
    python3 to_markdown.py /workspace/in.docx -o /workspace/out.md
    python3 to_markdown.py /workspace/in.pdf            # 不给 -o 则打印到 stdout
"""
from __future__ import annotations

import os
import sys


def to_markdown(src: str) -> str:
    if not os.path.isfile(src):
        raise FileNotFoundError(src)
    try:
        from markitdown import MarkItDown
    except ImportError as e:  # pragma: no cover
        raise ImportError('需要 markitdown[all]：沙箱镜像应已预装；若仍缺失，请重建镜像') from e
    md = MarkItDown()
    result = md.convert(src)
    return result.text_content or ""


def _main(argv=None) -> int:
    import argparse
    ap = argparse.ArgumentParser(prog="to_markdown",
                                 description="把 Word/Excel/PPT/PDF/HTML/图片 等转成 Markdown")
    ap.add_argument("src", help="源文件（建议放 /workspace）")
    ap.add_argument("-o", "--out", default=None, help="输出 .md 路径（省略则打印到 stdout）")
    args = ap.parse_args(argv)

    text = to_markdown(args.src)
    if args.out:
        out = os.path.abspath(args.out)
        os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
        with open(out, "w", encoding="utf-8") as f:
            f.write(text)
        print(f"[OK] 已生成 {out}（{len(text)} 字符）", flush=True)
    else:
        sys.stdout.write(text)
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
