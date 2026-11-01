# -*- coding: utf-8 -*-
"""运行环境自检脚本。

检查项：
- 必需项：Python >= 3.10、python-docx >= 1.1、lxml、jsonschema
- 可选项：LibreOffice (soffice)、PyMuPDF（fitz，实际 PDF 渲染/视觉核验依赖）、pdfplumber

输出 JSON 与人类可读结论；若缺少任何必需项，退出码为 1，否则为 0。
"""
from __future__ import annotations

import argparse
import importlib
import importlib.metadata
import json
import shutil
import subprocess
import sys
from typing import Any


def _parse_version(ver_str: str) -> tuple[int, ...]:
    """将版本号字符串解析为整数元组以便比较。"""
    parts: list[int] = []
    for token in ver_str.replace("-", ".").split("."):
        num = ""
        for ch in token:
            if ch.isdigit():
                num += ch
            else:
                break
        if num:
            parts.append(int(num))
    return tuple(parts)


def _get_pkg_version(dist_name: str, mod: Any) -> str:
    """优先使用 importlib.metadata 获取包版本，避免触发 __version__ 弃用警告。"""
    try:
        return importlib.metadata.version(dist_name)
    except Exception:
        return str(getattr(mod, "__version__", "unknown"))


def check_python() -> dict[str, Any]:
    """检查 Python 版本（要求 >= 3.10）。"""
    ver_info = sys.version_info
    ver_str = f"{ver_info.major}.{ver_info.minor}.{ver_info.micro}"
    ok = (ver_info.major, ver_info.minor) >= (3, 10)
    return {
        "name": "python",
        "required": True,
        "available": True,
        "ok": ok,
        "version": ver_str,
        "min_version": "3.10.0",
        "message": "满足要求" if ok else f"Python 版本过低（当前 {ver_str}，需 >= 3.10）",
    }


def check_python_package(
    import_name: str,
    dist_name: str,
    min_version: str | None = None,
    required: bool = True,
) -> dict[str, Any]:
    """检查指定 Python 包是否可导入且版本达标。"""
    try:
        mod = importlib.import_module(import_name)
        ver_str = _get_pkg_version(dist_name, mod)
        if min_version and ver_str != "unknown":
            ok = _parse_version(ver_str) >= _parse_version(min_version)
        else:
            ok = True
        msg = "满足要求" if ok else f"版本过低（当前 {ver_str}，需 >= {min_version}）"
        return {
            "name": dist_name,
            "required": required,
            "available": True,
            "ok": ok,
            "version": ver_str,
            "min_version": min_version,
            "message": msg,
        }
    except Exception as exc:
        return {
            "name": dist_name,
            "required": required,
            "available": False,
            "ok": False,
            "version": None,
            "min_version": min_version,
            "message": f"未安装或导入失败: {exc}",
        }


def check_soffice() -> dict[str, Any]:
    """检查可选外部工具 LibreOffice (soffice)。"""
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if not exe:
        return {
            "name": "soffice",
            "required": False,
            "available": False,
            "ok": False,
            "version": None,
            "path": None,
            "message": "未检测到 LibreOffice (soffice)，可选功能（旧格式转换/PDF渲染检查）不可用",
        }
    ver_str = "unknown"
    try:
        proc = subprocess.run(
            [exe, "--version"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=10,
        )
        out = (proc.stdout or proc.stderr or "").strip()
        if out:
            ver_str = out.splitlines()[0].strip()
    except Exception:
        pass
    return {
        "name": "soffice",
        "required": False,
        "available": True,
        "ok": True,
        "version": ver_str,
        "path": exe,
        "message": "可用",
    }


def run_checks() -> dict[str, Any]:
    """执行全部环境检查并汇总结果。"""
    checks = [
        check_python(),
        check_python_package("docx", "python-docx", min_version="1.1.0", required=True),
        check_python_package("lxml", "lxml", min_version="4.9.0", required=True),
        check_python_package("jsonschema", "jsonschema", min_version="4.0.0", required=True),
        check_soffice(),
        # docx_visual_verify.py 使用 fitz；此前只检查 pdfplumber 会把真正的
        # 视觉核验依赖误报为已满足，导致运行时才发现无法渲染 PDF。
        check_python_package("pymupdf", "PyMuPDF", min_version="1.20.0", required=False),
        check_python_package("pdfplumber", "pdfplumber", min_version=None, required=False),
    ]
    missing_required = [c["name"] for c in checks if c["required"] and not c["ok"]]
    missing_optional = [c["name"] for c in checks if not c["required"] and not c["ok"]]
    all_required_ok = len(missing_required) == 0

    if all_required_ok:
        if missing_optional:
            conclusion = (
                f"环境自检通过：所有必需组件均已就绪；"
                f"可选组件未就绪：{', '.join(missing_optional)}。"
            )
        else:
            conclusion = "环境自检通过：所有必需组件与可选组件均已就绪。"
    else:
        conclusion = f"环境自检失败：缺少或不满足必需组件：{', '.join(missing_required)}。"

    return {
        "ok": all_required_ok,
        "missing_required": missing_required,
        "missing_optional": missing_optional,
        "conclusion": conclusion,
        "checks": checks,
    }


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8")

    parser = argparse.ArgumentParser(description="检查集团法务合同 Skill 运行环境依赖")
    parser.add_argument(
        "--json-only",
        action="store_true",
        help="仅输出 JSON，不附加人类可读分割文本",
    )
    args = parser.parse_args(argv)

    report = run_checks()
    print(json.dumps(report, ensure_ascii=False, indent=2))
    if not args.json_only:
        print("\n=== 环境检查结论 ===")
        for item in report["checks"]:
            req_tag = "必需" if item["required"] else "可选"
            status_tag = "通过" if item["ok"] else ("缺失/不可用" if not item["available"] else "不满足")
            ver = item.get("version") or "-"
            print(f"- [{req_tag}] {item['name']} (版本: {ver}): {status_tag} — {item['message']}")
        print(report["conclusion"])

    return 0 if report["ok"] else 1


if __name__ == "__main__":
    sys.exit(main())
