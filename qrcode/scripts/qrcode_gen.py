# -*- coding: utf-8 -*-
"""
qrcode_gen —— 生成二维码 PNG（独立脚本）。

依赖：qrcode[pil]（沙箱镜像已预装；若当前运行镜像较旧，需重建）。

用法：
    python3 qrcode_gen.py "https://example.com" -o /workspace/qr.png
    python3 qrcode_gen.py "文本内容" -o /workspace/qr.png --ecc H --size 12 --border 4 --fill black --bg white
"""
from __future__ import annotations

import os


def make_qr(data: str, out: str, ecc: str = "M", box_size: int = 10, border: int = 4,
            fill: str = "black", bg: str = "white") -> str:
    try:
        import qrcode
        from qrcode.constants import (ERROR_CORRECT_H, ERROR_CORRECT_L,
                                      ERROR_CORRECT_M, ERROR_CORRECT_Q)
    except ImportError as e:  # pragma: no cover
        raise ImportError("需要 qrcode[pil]：沙箱镜像应已预装；若仍缺失，请重建镜像") from e

    levels = {"L": ERROR_CORRECT_L, "M": ERROR_CORRECT_M, "Q": ERROR_CORRECT_Q, "H": ERROR_CORRECT_H}
    qr = qrcode.QRCode(version=None, error_correction=levels.get(ecc.upper(), ERROR_CORRECT_M),
                       box_size=box_size, border=border)
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color=fill, back_color=bg)
    out = os.path.abspath(out)
    os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
    img.save(out)
    return out


def _main(argv=None) -> int:
    import argparse
    ap = argparse.ArgumentParser(prog="qrcode_gen", description="生成二维码 PNG")
    ap.add_argument("data", help="二维码内容（网址或文本）")
    ap.add_argument("-o", "--out", required=True, help="输出 PNG 路径（建议 /workspace）")
    ap.add_argument("--ecc", default="M", help="纠错等级 L/M/Q/H（默认 M）")
    ap.add_argument("--size", type=int, default=10, help="每个点的像素(box_size)，默认 10")
    ap.add_argument("--border", type=int, default=4, help="边框格数，默认 4")
    ap.add_argument("--fill", default="black", help="前景色，默认 black")
    ap.add_argument("--bg", default="white", help="背景色，默认 white")
    args = ap.parse_args(argv)
    out = make_qr(args.data, args.out, args.ecc, args.size, args.border, args.fill, args.bg)
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
