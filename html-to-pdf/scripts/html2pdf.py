# -*- coding: utf-8 -*-
"""
html2pdf —— 用无头 Chromium / Chrome / Edge 把 HTML+CSS 渲染成 PDF（独立脚本）。

适合需要“浏览器级”视觉效果（渐变、卡片、阴影、Web 字体、复杂 CSS）的报告。
依赖：本机/容器存在 Chromium 或 Chrome 或 Edge（沙箱镜像已预装 chromium）。无需 pip 新依赖。

要点：
- 要让背景色/底纹打印出来，HTML 里需加：
      * { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
- 纸张/页边距用 CSS @page 控制：@page { size: A4; margin: 12mm; }
- emoji 需系统装有彩色 emoji 字体（如 fonts-noto-color-emoji），否则可能显示为方块。
"""
from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
from pathlib import Path

_CANDIDATES = [
    # Linux（沙箱）
    "/usr/bin/chromium", "/usr/bin/chromium-browser",
    "/usr/bin/google-chrome", "/usr/bin/microsoft-edge",
    # Windows - Edge / Chrome
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    # macOS
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
]
_PATH_NAMES = ["chromium", "chromium-browser", "google-chrome", "chrome", "msedge", "microsoft-edge"]


def find_browser(explicit: str | None = None) -> str:
    if explicit:
        if os.path.isfile(explicit):
            return explicit
        raise FileNotFoundError(f"指定的浏览器不存在：{explicit}")
    for name in _PATH_NAMES:
        p = shutil.which(name)
        if p:
            return p
    for p in _CANDIDATES:
        if os.path.isfile(p):
            return p
    raise FileNotFoundError("未找到 Chromium/Chrome/Edge。请安装其一或用 --browser 指定路径。")


def html_to_pdf(html, out_pdf, is_file=None, browser=None, timeout=120, virtual_time_ms=2000) -> str:
    """把 HTML 渲染成 PDF，返回输出路径。html 可为文件路径或源字符串（见 is_file）。"""
    exe = find_browser(browser)
    out_pdf = os.path.abspath(out_pdf)
    os.makedirs(os.path.dirname(out_pdf) or ".", exist_ok=True)

    if is_file is None:
        is_file = os.path.isfile(html)

    tmp_html = None
    try:
        if is_file:
            src_path = os.path.abspath(html)
        else:
            fd, tmp_html = tempfile.mkstemp(suffix=".html")
            with os.fdopen(fd, "w", encoding="utf-8") as f:
                f.write(html)
            src_path = tmp_html

        url = Path(src_path).as_uri()  # 正确处理空格/中文/反斜杠

        # 容器友好参数：--no-sandbox（非 root）、--disable-dev-shm-usage（/dev/shm 小）
        common = ["--disable-gpu", "--no-sandbox", "--disable-dev-shm-usage"]
        args = [exe, "--headless=new", *common, "--no-pdf-header-footer",
                f"--virtual-time-budget={virtual_time_ms}",
                f"--print-to-pdf={out_pdf}", url]
        proc = subprocess.run(args, capture_output=True, text=True, timeout=timeout)

        if not os.path.isfile(out_pdf):  # 旧版浏览器回退
            args_fb = [exe, "--headless", *common, "--print-to-pdf-no-header",
                       f"--print-to-pdf={out_pdf}", url]
            proc = subprocess.run(args_fb, capture_output=True, text=True, timeout=timeout)

        if not os.path.isfile(out_pdf):
            raise RuntimeError("PDF 生成失败。浏览器：%s\nstdout:%s\nstderr:%s"
                               % (exe, proc.stdout[-800:], proc.stderr[-800:]))
        return out_pdf
    finally:
        if tmp_html and os.path.isfile(tmp_html):
            os.remove(tmp_html)


def _main(argv=None) -> int:
    import argparse

    ap = argparse.ArgumentParser(
        prog="html2pdf", description="用无头 Chromium/Chrome/Edge 把 HTML+CSS 渲染成 PDF")
    ap.add_argument("html", help="HTML 文件路径（建议放 /workspace）")
    ap.add_argument("-o", "--out", required=True, help="输出 PDF 路径（建议 /workspace/xxx.pdf）")
    ap.add_argument("--browser", default=None, help="显式指定浏览器可执行文件")
    ap.add_argument("--wait", type=int, default=2000, help="渲染等待毫秒，默认 2000")
    args = ap.parse_args(argv)

    out = html_to_pdf(args.html, args.out, is_file=True, browser=args.browser, virtual_time_ms=args.wait)
    print(f"[OK] 已生成 {out}（浏览器：{find_browser(args.browser)}）", flush=True)
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
