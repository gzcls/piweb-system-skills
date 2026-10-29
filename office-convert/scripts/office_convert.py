# -*- coding: utf-8 -*-
"""
office_convert —— 用 LibreOffice 无头模式转换文档格式（独立脚本）。

能力（保留原排版）：
  - Office → PDF：docx/xlsx/pptx/doc/xls/ppt/odt → pdf
  - 格式互转：doc↔docx、xls↔xlsx、ppt↔pptx、*→odt 等
  - 文档 → 图片：可先转 PDF 再用 pdf-to-image 技能转图

依赖：LibreOffice（沙箱镜像已预装：impress/calc/writer）。无需 MS Office。

用法：
    python3 office_convert.py /workspace/in.docx --to pdf  -o /workspace
    python3 office_convert.py /workspace/in.xls  --to xlsx -o /workspace
"""
from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
from pathlib import Path

_CANDIDATES = [
    "/usr/bin/soffice", "/usr/bin/libreoffice", "/opt/libreoffice/program/soffice",
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
]
_NAMES = ["soffice", "libreoffice"]


def find_soffice(explicit: str | None = None) -> str:
    if explicit:
        if os.path.isfile(explicit):
            return explicit
        raise FileNotFoundError(f"指定的 LibreOffice 不存在：{explicit}")
    for n in _NAMES:
        p = shutil.which(n)
        if p:
            return p
    for p in _CANDIDATES:
        if os.path.isfile(p):
            return p
    raise FileNotFoundError("未找到 LibreOffice/soffice。沙箱镜像应已预装；本机需自行安装。")


def convert(src: str, to: str, out_dir: str | None = None, soffice: str | None = None,
            timeout: int = 180) -> str:
    """把 src 转成 to 指定格式，输出到 out_dir，返回输出文件路径。"""
    if not os.path.isfile(src):
        raise FileNotFoundError(src)
    exe = find_soffice(soffice)
    src = os.path.abspath(src)
    out_dir = os.path.abspath(out_dir or os.path.dirname(src))
    os.makedirs(out_dir, exist_ok=True)

    # 独立用户配置目录，避免并发/只读 HOME 问题
    profile = tempfile.mkdtemp(prefix="lo_profile_")
    try:
        # LibreOffice expects a URI here.  ``file://C:\\...`` is malformed on
        # Windows; Path.as_uri() emits ``file:///C:/...`` on all platforms.
        profile_uri = Path(profile).resolve().as_uri()
        args = [
            exe, "--headless", "--norestore", "--nolockcheck", "--nodefault",
            f"-env:UserInstallation={profile_uri}",
            "--convert-to", to, "--outdir", out_dir, src,
        ]
        # LibreOffice emits UTF-8 on Windows; use an explicit codec so a
        # non-UTF-8 system locale cannot break subprocess output collection.
        proc = subprocess.run(
            args, capture_output=True, text=True, encoding="utf-8",
            errors="replace", timeout=timeout,
        )
        base = os.path.splitext(os.path.basename(src))[0]
        ext = to.split(":")[0]  # 形如 "pdf" 或 "xlsx:Calc MS Excel 2007 XML"
        out = os.path.join(out_dir, f"{base}.{ext}")
        if not os.path.isfile(out):
            raise RuntimeError(
                "转换失败。soffice：%s\nstdout:%s\nstderr:%s"
                % (exe, proc.stdout[-800:], proc.stderr[-800:]))
        return out
    finally:
        shutil.rmtree(profile, ignore_errors=True)


def _main(argv=None) -> int:
    import argparse
    ap = argparse.ArgumentParser(prog="office_convert",
                                 description="LibreOffice 无头转换：Office↔PDF、格式互转")
    ap.add_argument("src", help="源文件（建议放 /workspace）")
    ap.add_argument("--to", required=True, help="目标格式，如 pdf / docx / xlsx / pptx / png / odt")
    ap.add_argument("-o", "--out", default=None, help="输出目录（建议 /workspace；默认源文件同目录）")
    ap.add_argument("--soffice", default=None, help="显式指定 soffice 可执行文件")
    args = ap.parse_args(argv)
    out = convert(args.src, args.to, args.out, args.soffice)
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
