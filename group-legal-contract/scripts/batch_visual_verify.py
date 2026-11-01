#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""并行执行多个最终 DOCX 的版面检查。

交付前通常只需检查审查修订版和法律审查意见书。该脚本复用
``docx_visual_verify.inspect_docx``，为每个文档使用独立临时目录，避免串行
启动 LibreOffice 造成无谓等待，也不对中间清洁版重复做完整渲染。
"""
from __future__ import annotations

import argparse
import json
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Any

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from docx_visual_verify import inspect_docx  # noqa: E402


def verify_many(
    paths: list[str | Path],
    max_workers: int | None = None,
    *,
    require_render: bool = False,
    fail_on_warnings: bool = False,
    cache_dir: str | Path | None = None,
) -> dict[str, Any]:
    requested = [str(Path(p).resolve()) for p in paths]
    # 同一交付文件可能同时被 docx_verify 和最终批量检查传入；批量阶段
    # 只渲染一次，保留去重数量便于审计耗时优化。
    normalized = list(dict.fromkeys(requested))
    requested_workers = max_workers or 4
    workers = min(max(1, requested_workers), 4, max(1, len(normalized)))
    results: dict[str, Any] = {}
    with ThreadPoolExecutor(max_workers=workers, thread_name_prefix="docx-visual") as pool:
        if require_render or fail_on_warnings:
            future_map = {
                pool.submit(inspect_docx, p, require_render=require_render, fail_on_warnings=fail_on_warnings, cache_dir=cache_dir): p
                for p in normalized
            }
        else:
            # Keep the default call signature compatible with lightweight test
            # doubles and downstream wrappers that predate the optional flags.
            if cache_dir is not None:
                future_map = {pool.submit(inspect_docx, p, cache_dir=cache_dir): p for p in normalized}
            else:
                future_map = {pool.submit(inspect_docx, p): p for p in normalized}
        for fut in as_completed(future_map):
            p = future_map[fut]
            try:
                results[p] = fut.result()
            except Exception as exc:
                results[p] = {
                    "status": "error",
                    "path": p,
                    "passed": False,
                    "findings": [{"severity": "error", "kind": "exception", "message": str(exc)}],
                }
    ordered = [results[p] for p in normalized]
    status_counts: dict[str, int] = {}
    for result in ordered:
        status = str(result.get("status", "error"))
        status_counts[status] = status_counts.get(status, 0) + 1
    return {
        "requested_count": len(requested),
        "count": len(ordered),
        "deduplicated_count": len(requested) - len(normalized),
        "workers": workers,
        "require_render": require_render,
        "fail_on_warnings": fail_on_warnings,
        "status_counts": status_counts,
        "warning_count": sum(int(r.get("warning_count", 0)) for r in ordered),
        "error_count": sum(int(r.get("error_count", 0)) for r in ordered),
        "passed": all(bool(r.get("passed")) for r in ordered),
        "results": ordered,
    }


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description="并行检查多个最终 DOCX 的渲染与版面")
    parser.add_argument("docx", nargs="+", help="待检查的 .docx 文件，可传多个")
    parser.add_argument("--workers", type=int, default=None, help="并行数（默认最多4）")
    parser.add_argument("--require-render", action="store_true", help="渲染依赖缺失或渲染失败时判定失败")
    parser.add_argument("--fail-on-warnings", action="store_true", help="存在版面警告时判定失败")
    parser.add_argument("--out", default=None, help="写入 JSON 报告路径")
    parser.add_argument("--cache-dir", default=None, help="视觉报告缓存目录")
    args = parser.parse_args(argv)
    try:
        report = verify_many(
            args.docx,
            max_workers=max(1, args.workers) if args.workers else None,
            require_render=args.require_render,
            fail_on_warnings=args.fail_on_warnings,
            cache_dir=args.cache_dir,
        )
    except Exception as exc:
        print(json.dumps({"error": str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 1
    if args.out:
        out = Path(args.out)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
