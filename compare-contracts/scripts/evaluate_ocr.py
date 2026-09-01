#!/usr/bin/env python3
"""Calculate character error rate from extraction JSON and human ground truth."""

from __future__ import annotations

import argparse
import json
import math
import re
from pathlib import Path


CONTROL_FILES = {"extraction-summary.json", "review-queue.json"}


def normalize_text(text: str, mode: str) -> str:
    if mode == "exact":
        return text.replace("\r\n", "\n").replace("\r", "\n")
    if mode == "whitespace":
        return re.sub(r"\s+", "", text)
    raise ValueError(f"unsupported normalization mode: {mode}")


def levenshtein_distance(reference: str, hypothesis: str) -> int:
    if len(reference) < len(hypothesis):
        reference, hypothesis = hypothesis, reference
    previous = list(range(len(hypothesis) + 1))
    for row, reference_char in enumerate(reference, start=1):
        current = [row]
        for column, hypothesis_char in enumerate(hypothesis, start=1):
            current.append(
                min(
                    current[column - 1] + 1,
                    previous[column] + 1,
                    previous[column - 1] + (reference_char != hypothesis_char),
                )
            )
        previous = current
    return previous[-1]


def load_extractions(root: Path) -> tuple[dict[tuple[str, int | None], str], dict]:
    if not root.is_dir():
        raise ValueError(f"extraction directory does not exist: {root}")
    summary_path = root / "extraction-summary.json"
    files = []
    if summary_path.exists():
        summary = json.loads(summary_path.read_text(encoding="utf-8"))
        if not isinstance(summary, dict):
            raise ValueError("extraction summary must be an object")
        if summary.get("error_count") != 0 or summary.get("errors"):
            raise ValueError("cannot evaluate an extraction batch with file errors")
        records = summary.get("records")
        if not isinstance(records, list):
            raise ValueError("extraction summary records must be a list")
        files = [
            root / record["output"]
            for record in records
            if isinstance(record, dict)
            and isinstance(record.get("output"), str)
            and record["output"]
        ]
    else:
        files = sorted(path for path in root.glob("*.json") if path.name not in CONTROL_FILES)
    if not files:
        raise ValueError(f"no extraction JSON files found in: {root}")
    missing = [path for path in files if not path.is_file()]
    if missing:
        raise ValueError(
            "extraction summary references missing outputs: "
            + ", ".join(path.name for path in missing[:12])
        )
    records = {}
    aliases = {}
    for path in files:
        payload = json.loads(path.read_text(encoding="utf-8"))
        relative = payload.get("relative_path") or payload.get("source")
        source = payload.get("source")
        for page in payload.get("pages", []):
            page_number = page.get("page")
            text = page.get("text")
            if not isinstance(text, str):
                continue
            records[(relative, page_number)] = text
            if source:
                aliases[(source, page_number)] = (relative, page_number)
    return records, aliases


def load_ground_truth(path: Path) -> list[dict]:
    payload = json.loads(path.read_text(encoding="utf-8"))
    records = payload.get("records") if isinstance(payload, dict) else payload
    if not isinstance(records, list) or not records:
        raise ValueError("ground truth must be a non-empty list or an object with records")
    for index, record in enumerate(records, start=1):
        if not isinstance(record, dict):
            raise ValueError(f"ground-truth record {index} is not an object")
        if not record.get("source") or not isinstance(record.get("text"), str):
            raise ValueError(f"ground-truth record {index} requires source and text")
    return records


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("extraction", type=Path, help="Directory containing extraction JSON")
    parser.add_argument("ground_truth", type=Path, help="Human-corrected JSON records")
    parser.add_argument("--normalization", choices=("exact", "whitespace"), default="whitespace")
    parser.add_argument("--max-cer", type=float, help="Fail when aggregate CER exceeds this value")
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()

    if args.max_cer is not None and (
        not math.isfinite(args.max_cer) or args.max_cer < 0
    ):
        parser.error("--max-cer must be a non-negative finite number")
    if args.output and args.output.resolve().parent == args.extraction.resolve():
        parser.error("--output must be outside the extraction directory")

    try:
        extractions, aliases = load_extractions(args.extraction.resolve())
        ground_truth = load_ground_truth(args.ground_truth.resolve())
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        parser.error(str(exc))

    results = []
    missing = []
    total_edits = 0
    total_reference_chars = 0
    for record in ground_truth:
        key = (record["source"], record.get("page"))
        resolved_key = key if key in extractions else aliases.get(key)
        if resolved_key is None or resolved_key not in extractions:
            missing.append({"source": record["source"], "page": record.get("page")})
            continue
        reference = normalize_text(record["text"], args.normalization)
        hypothesis = normalize_text(extractions[resolved_key], args.normalization)
        edits = levenshtein_distance(reference, hypothesis)
        reference_chars = len(reference)
        cer = edits / reference_chars if reference_chars else (0.0 if not hypothesis else 1.0)
        total_edits += edits
        total_reference_chars += reference_chars
        results.append(
            {
                "source": resolved_key[0],
                "page": resolved_key[1],
                "reference_chars": reference_chars,
                "hypothesis_chars": len(hypothesis),
                "edit_distance": edits,
                "cer": round(cer, 6),
                "exact_match": reference == hypothesis,
            }
        )

    aggregate_cer = (
        total_edits / total_reference_chars
        if total_reference_chars
        else (0.0 if results else None)
    )
    report = {
        "extraction_directory": str(args.extraction.resolve()),
        "ground_truth": str(args.ground_truth.resolve()),
        "normalization": args.normalization,
        "evaluated_page_count": len(results),
        "missing_page_count": len(missing),
        "total_reference_chars": total_reference_chars,
        "total_edit_distance": total_edits,
        "aggregate_cer": round(aggregate_cer, 6) if aggregate_cer is not None else None,
        "missing": missing,
        "results": results,
    }
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))

    threshold_failed = (
        args.max_cer is not None
        and aggregate_cer is not None
        and aggregate_cer > args.max_cer
    )
    return 1 if missing or not results or threshold_failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
