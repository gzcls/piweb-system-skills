#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""模型与脚本阶段耗时打点与汇总工具 (stage_mark.py)。

用法：
  python scripts/stage_mark.py <out_dir_or_work_dir> <stage_name> [--note "..."]
  python scripts/stage_mark.py <out_dir_or_work_dir> --summary
"""
from __future__ import annotations

import argparse
import json
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from lib.runtime_state import atomic_json, atomic_text, file_lock


def _resolve_work_dir(path: str | Path) -> Path:
    p = Path(path).resolve()
    if p.name == "work" or (p / "preflight.json").is_file() or (p / "stage-marks.json").is_file():
        return p
    return p / "work"


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="milliseconds").replace("+00:00", "Z")


def record_stage_mark(
    out_or_work_dir: str | Path,
    stage: str,
    *,
    note: str | None = None,
    extra: dict[str, Any] | None = None,
    run_id: str | None = None,
    actor: str | None = None,
) -> dict[str, Any]:
    """向 work/stage-marks.json 追加一条阶段打点记录。"""
    work = _resolve_work_dir(out_or_work_dir)
    work.mkdir(parents=True, exist_ok=True)
    marks_path = work / "stage-marks.json"

    with file_lock(work / ".stage-marks.lock"):
        data: dict[str, Any] = {"marks": []}
        if marks_path.is_file():
            try:
                loaded = json.loads(marks_path.read_text(encoding="utf-8"))
                if isinstance(loaded, dict) and isinstance(loaded.get("marks"), list):
                    data = loaded
            except Exception:
                data = {"marks": []}

        now_epoch = round(time.time(), 3)
        now_mono = time.perf_counter()
        marks: list[dict[str, Any]] = data["marks"]
        first_epoch = marks[0].get("epoch_s", now_epoch) if marks else now_epoch
        prev_epoch = marks[-1].get("epoch_s", now_epoch) if marks else now_epoch
        first_mono = marks[0].get("monotonic_s") if marks else None
        prev_mono = marks[-1].get("monotonic_s") if marks else None
        effective_run_id = run_id or (extra or {}).get("run_id")

        entry: dict[str, Any] = {
            "stage": stage,
            "timestamp_utc": _utc_now_iso(),
            "epoch_s": now_epoch,
            "delta_from_prev_s": round(max(0.0, now_epoch - float(prev_epoch)), 3),
            "elapsed_from_start_s": round(max(0.0, now_epoch - float(first_epoch)), 3),
            "monotonic_s": now_mono,
        }
        if prev_mono is not None:
            entry["monotonic_delta_s"] = round(max(0.0, now_mono - float(prev_mono)), 3)
        if first_mono is not None:
            entry["monotonic_elapsed_s"] = round(max(0.0, now_mono - float(first_mono)), 3)
        if effective_run_id:
            entry["run_id"] = str(effective_run_id)
        if actor:
            entry["actor"] = str(actor)
        if note:
            entry["note"] = note
        if extra:
            entry["extra"] = extra
        marks.append(entry)
        atomic_json(marks_path, data)
        return entry


def build_time_summary(out_or_work_dir: str | Path, *, run_id: str | None = None) -> dict[str, Any]:
    """结合 stage-marks.json、preflight.json 与 pipeline-report.json 生成耗时汇总。"""
    work = _resolve_work_dir(out_or_work_dir)
    marks_path = work / "stage-marks.json"
    preflight_path = work / "preflight.json"
    pipeline_path = work / "pipeline-report.json"

    marks: list[dict[str, Any]] = []
    if marks_path.is_file():
        try:
            marks = json.loads(marks_path.read_text(encoding="utf-8")).get("marks") or []
        except Exception:
            marks = []
    if run_id:
        marks = [m for m in marks if str(m.get("run_id") or "") == str(run_id)]

    preflight_timings: dict[str, Any] = {}
    if preflight_path.is_file():
        try:
            preflight_timings = json.loads(preflight_path.read_text(encoding="utf-8")).get("timings_ms") or {}
        except Exception:
            preflight_timings = {}

    pipeline_timings: dict[str, Any] = {}
    if pipeline_path.is_file():
        try:
            pipeline_timings = json.loads(pipeline_path.read_text(encoding="utf-8")).get("timings_ms") or {}
        except Exception:
            pipeline_timings = {}

    # 计算模型分析区间（preflight_done -> draft_built 或 pipeline_started）
    stage_epochs = {
        m["stage"]: float(m.get("monotonic_s", m.get("epoch_s", 0.0)))
        for m in marks
        if "stage" in m and ("epoch_s" in m or "monotonic_s" in m)
    }
    model_breakdown: list[dict[str, Any]] = []
    for idx in range(1, len(marks)):
        prev_m = marks[idx - 1]
        cur_m = marks[idx]
        model_breakdown.append({
            "from_stage": prev_m.get("stage"),
            "to_stage": cur_m.get("stage"),
            "seconds": cur_m.get("monotonic_delta_s", cur_m.get("delta_from_prev_s", 0.0)),
            "note": cur_m.get("note") or "",
            "actor": cur_m.get("actor") or "unclassified",
        })

    classified_totals = {"script": 0.0, "model": 0.0, "wait": 0.0, "unclassified": 0.0}
    for interval in model_breakdown:
        actor = str(interval.get("actor") or "unclassified")
        bucket = actor if actor in classified_totals else "unclassified"
        classified_totals[bucket] = round(classified_totals[bucket] + float(interval.get("seconds") or 0.0), 3)

    wall_clock_interval_s = None
    start_key = "preflight_done" if "preflight_done" in stage_epochs else (marks[0]["stage"] if marks else None)
    end_key = "draft_built" if "draft_built" in stage_epochs else (
        "pipeline_done" if "pipeline_done" in stage_epochs else None
    )
    if start_key and end_key and end_key in stage_epochs and start_key in stage_epochs:
        wall_clock_interval_s = round(max(0.0, stage_epochs[end_key] - stage_epochs[start_key]), 3)
    has_model_marks = any(interval.get("actor") == "model" for interval in model_breakdown)
    model_total_s = classified_totals["model"] if has_model_marks else None

    summary = {
        "work_dir": str(work),
        "generated_at_utc": _utc_now_iso(),
        "preflight_timings_ms": preflight_timings,
        "pipeline_timings_ms": pipeline_timings,
        "stage_marks": marks,
        "intervals": model_breakdown,
        "model_analysis_total_s": model_total_s,
        "wall_clock_interval_s": wall_clock_interval_s,
        "model_analysis_status": "measured" if has_model_marks else "not_recorded",
        "classified_totals_s": classified_totals,
        "model_wait_excluded_total_s": classified_totals["model"] if any(i.get("actor") == "model" for i in model_breakdown) else None,
    }

    atomic_json(work / "time-summary.json", summary)

    md_lines = [
        "# 审核耗时记录",
        "",
        "## 一、 阶段打点明细（`stage-marks.json`）",
        "",
        "| 序号 | 阶段标识 | UTC 时间 | 距上一打点耗时 | 累计耗时 | 备注 |",
        "| :---: | :--- | :--- | ---: | ---: | :--- |",
    ]
    for idx, m in enumerate(marks, 1):
        md_lines.append(
            f"| {idx} | `{m.get('stage')}` | {m.get('timestamp_utc')} | "
            f"{m.get('delta_from_prev_s', 0):.3f} 秒 | {m.get('elapsed_from_start_s', 0):.3f} 秒 | "
            f"{m.get('note') or '—'} |"
        )
    if not marks:
        md_lines.append("| — | 暂无打点 | — | — | — | — |")

    if wall_clock_interval_s is not None:
        md_lines.extend([
            "",
            f"> **阶段墙钟区间（`{start_key}` ➔ `{end_key}`）**：**{wall_clock_interval_s:.2f} 秒**；其中模型时间只有在打点 `actor=model` 时才计入 `model_analysis_total_s`。",
        ])

    md_lines.extend([
        "",
        "## 二、 确定性脚本阶段耗时（毫秒）",
        "",
        "| 脚本阶段 | 耗时 (ms) |",
        "| :--- | ---: |",
    ])
    for k, v in preflight_timings.items():
        md_lines.append(f"| `preflight.{k}` | {v} |")
    for k, v in pipeline_timings.items():
        md_lines.append(f"| `pipeline.{k}` | {v} |")

    atomic_text(work / "time-summary.md", "\n".join(md_lines) + "\n")
    return summary


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8")

    parser = argparse.ArgumentParser(description="记录审核阶段时间戳或生成耗时汇总")
    parser.add_argument("out_dir", help="本次审核输出目录（或 work/ 目录）")
    parser.add_argument("stage", nargs="?", default=None, help="阶段标识（如 context_read、citations_fetched、draft_written）")
    parser.add_argument("--note", default=None, help="附加说明")
    parser.add_argument("--actor", choices=["script", "model", "wait", "unclassified"], default=None, help="阶段耗时归属，用于区分脚本、模型和等待")
    parser.add_argument("--summary", action="store_true", help="生成 time-summary.json 与 time-summary.md")
    args = parser.parse_args(argv)

    if not args.stage and not args.summary:
        parser.error("必须提供 stage 名称或 --summary")

    result: dict[str, Any] = {}
    if args.stage:
        result["mark"] = record_stage_mark(args.out_dir, args.stage, note=args.note, actor=args.actor)
    if args.summary:
        result["summary"] = build_time_summary(args.out_dir)

    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
