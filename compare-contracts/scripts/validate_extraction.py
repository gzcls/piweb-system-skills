#!/usr/bin/env python3
"""Validate contract extraction artifacts, OCR metadata, and batch reconciliation."""

from __future__ import annotations

import argparse
import hashlib
import json
import math
from collections import Counter
from pathlib import Path

CONTROL_FILES = {"extraction-summary.json", "review-queue.json"}
REQUIRED_FIELDS = {
    "source",
    "relative_path",
    "extension",
    "content_sha256",
    "status",
    "pages",
}
VALID_STATUSES = {"partial", "truncated", "complete"}
VALID_REVIEW_ACTIONS = {
    "enhanced-ocr-then-verify",
    "manual-verify-original",
    "verify-material-fields",
}
REQUIRED_LEDGER_ACTIONS = {
    "enhanced-ocr-then-verify",
    "manual-verify-original",
}
VALID_LEDGER_STATUSES = {"pending", "resolved", "unclear"}


def config_fingerprint(config: dict) -> str:
    encoded = json.dumps(config, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(encoded.encode("utf-8")).hexdigest()


def finite_number(value) -> bool:
    return (
        isinstance(value, (int, float))
        and not isinstance(value, bool)
        and math.isfinite(float(value))
    )


def supports_review_actions(payload: dict) -> bool:
    config = payload.get("extraction_config")
    version = config.get("extractor_version") if isinstance(config, dict) else None
    return isinstance(version, int) and not isinstance(version, bool) and version >= 4


def load_batch(root: Path) -> tuple[list[Path], dict[str, dict], dict | None, list[str]]:
    summary_path = root / "extraction-summary.json"
    if not summary_path.exists():
        files = sorted(path for path in root.glob("*.json") if path.name not in CONTROL_FILES)
        return files, {}, None, []

    errors = []
    try:
        summary = json.loads(summary_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        return [], {}, None, [f"invalid extraction summary: {exc}"]
    if not isinstance(summary, dict):
        return [], {}, None, ["extraction summary must be an object"]

    records = summary.get("records")
    batch_errors = summary.get("errors")
    if not isinstance(records, list):
        errors.append("summary records must be a list")
        records = []
    if not isinstance(batch_errors, list):
        errors.append("summary errors must be a list")
        batch_errors = []

    counts = {}
    for field in ("file_count", "success_count", "error_count"):
        value = summary.get(field)
        if not isinstance(value, int) or isinstance(value, bool) or value < 0:
            errors.append(f"summary {field} must be a non-negative integer")
            counts[field] = None
        else:
            counts[field] = value
    if counts["success_count"] is not None and counts["success_count"] != len(records):
        errors.append("summary success_count does not match records")
    if counts["error_count"] is not None and counts["error_count"] != len(batch_errors):
        errors.append("summary error_count does not match errors")
    if all(counts[field] is not None for field in counts):
        if counts["file_count"] != counts["success_count"] + counts["error_count"]:
            errors.append("summary file_count does not equal successes plus errors")
    if batch_errors or (counts["error_count"] or 0) > 0:
        errors.append(f"extraction batch contains {len(batch_errors)} failed file(s)")

    record_map = {}
    files = []
    for index, record in enumerate(records, start=1):
        if not isinstance(record, dict):
            errors.append(f"summary record {index} is not an object")
            continue
        output = record.get("output")
        if not isinstance(output, str) or not output:
            errors.append(f"summary record {index} has no output")
            continue
        if Path(output).name != output:
            errors.append(f"summary record {index} output must be a filename")
            continue
        if output in record_map:
            errors.append(f"duplicate summary output: {output}")
            continue
        record_map[output] = record
        files.append(root / output)
    for path in files:
        if not path.is_file():
            errors.append(f"summary references missing output: {path.name}")
    files = [path for path in files if path.is_file()]
    return files, record_map, summary, errors


def validate_attempts(
    page: dict,
    locator: str,
    require_diagnostics: bool = False,
) -> tuple[list[str], list[str], dict]:
    errors = []
    warnings = []
    attempts = page.get("ocr_attempts")
    if not isinstance(attempts, list) or not attempts:
        warnings.append(f"{locator}: OCR page has no attempt metadata")
        return errors, warnings, {"attempt_count": 0, "adaptive": False, "high_resolution_selected": False}

    for attempt_index, attempt in enumerate(attempts, start=1):
        prefix = f"{locator}: attempt {attempt_index}"
        if not isinstance(attempt, dict):
            errors.append(f"{prefix} is not an object")
            continue
        text = attempt.get("text")
        if not isinstance(text, str):
            errors.append(f"{prefix} text must be a string")
        digest = attempt.get("text_sha256")
        if not isinstance(digest, str) or len(digest) != 64:
            errors.append(f"{prefix} has invalid text_sha256")
        elif isinstance(text, str):
            actual = hashlib.sha256(text.encode("utf-8")).hexdigest()
            if actual != digest:
                errors.append(f"{prefix} text hash mismatch")

        confidence = attempt.get("confidence")
        if confidence is not None and (
            not finite_number(confidence) or not 0.0 <= float(confidence) <= 1.0
        ):
            errors.append(f"{prefix} confidence must be null or between 0 and 1")
        for field in ("effective_scale", "quality_score"):
            value = attempt.get(field)
            if not finite_number(value) or float(value) < 0:
                errors.append(f"{prefix} {field} must be a non-negative finite number")
        text_chars = attempt.get("text_chars")
        if not isinstance(text_chars, int) or isinstance(text_chars, bool) or text_chars < 0:
            errors.append(f"{prefix} text_chars must be a non-negative integer")
        elif isinstance(text, str) and text_chars != len(text):
            errors.append(f"{prefix} text_chars does not match text length")
        lines = attempt.get("lines")
        if not isinstance(lines, list):
            errors.append(f"{prefix} lines must be a list")
        else:
            for line_index, line in enumerate(lines, start=1):
                if not isinstance(line, dict) or not isinstance(line.get("text"), str):
                    errors.append(f"{prefix} line {line_index} is invalid")
                    continue
                line_confidence = line.get("confidence")
                if line_confidence is not None and (
                    not finite_number(line_confidence)
                    or not 0.0 <= float(line_confidence) <= 1.0
                ):
                    errors.append(f"{prefix} line {line_index} confidence is invalid")
        if require_diagnostics:
            for field in ("numeric_token_count", "material_value_count"):
                value = attempt.get(field)
                if not isinstance(value, int) or isinstance(value, bool) or value < 0:
                    errors.append(f"{prefix} {field} must be a non-negative integer")
            short_line_ratio = attempt.get("short_line_ratio")
            if not finite_number(short_line_ratio) or not 0.0 <= float(short_line_ratio) <= 1.0:
                errors.append(f"{prefix} short_line_ratio must be between 0 and 1")

    selected = page.get("selected_attempt")
    selected_index = 0
    if not isinstance(selected, int) or isinstance(selected, bool) or not 1 <= selected <= len(attempts):
        errors.append(f"{locator}: invalid selected_attempt {selected!r}")
    else:
        selected_index = selected - 1
        chosen = attempts[selected_index]
        if isinstance(chosen, dict):
            if isinstance(chosen.get("text"), str) and page.get("text") != chosen["text"]:
                errors.append(f"{locator}: selected attempt text does not match page text")
            chosen_confidence = chosen.get("confidence")
            page_confidence = page.get("ocr_confidence")
            if finite_number(chosen_confidence) and finite_number(page_confidence):
                if not math.isclose(float(chosen_confidence), float(page_confidence), abs_tol=0.0001):
                    errors.append(f"{locator}: selected attempt confidence mismatch")

    return errors, warnings, {
        "attempt_count": len(attempts),
        "adaptive": len(attempts) > 1,
        "high_resolution_selected": len(attempts) > 1 and selected_index > 0,
    }


def validate_file(
    path: Path,
    require_complete: bool,
    min_confidence: float,
    require_config: bool,
) -> tuple[list[str], list[str], dict, dict | None]:
    errors = []
    warnings = []
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:  # noqa: BLE001
        return [f"invalid JSON: {exc}"], warnings, {}, None
    if not isinstance(payload, dict):
        return ["top-level extraction must be an object"], warnings, {}, None

    missing = sorted(REQUIRED_FIELDS - payload.keys())
    if missing:
        errors.append(f"missing fields: {', '.join(missing)}")

    config = payload.get("extraction_config")
    fingerprint = payload.get("extraction_config_fingerprint")
    if isinstance(config, dict) and isinstance(fingerprint, str):
        if config_fingerprint(config) != fingerprint:
            errors.append("extraction configuration fingerprint mismatch")
    elif require_config:
        errors.append("missing extraction configuration metadata")
    else:
        warnings.append("legacy extraction has no configuration fingerprint")
    require_review_actions = supports_review_actions(payload)

    status = payload.get("status")
    if not isinstance(status, str) or status not in VALID_STATUSES:
        errors.append(f"invalid status: {status!r}")
    if require_complete and status != "complete":
        errors.append(f"status is {status!r}, expected 'complete'")

    pages = payload.get("pages")
    if not isinstance(pages, list):
        errors.append("pages must be a list")
        pages = []
    units_completed = payload.get("units_completed")
    if not isinstance(units_completed, int) or isinstance(units_completed, bool) or units_completed < 0:
        errors.append("units_completed must be a non-negative integer")
    elif units_completed != len(pages):
        errors.append(f"units_completed mismatch: {units_completed} vs {len(pages)} pages")

    page_count = payload.get("page_count")
    if page_count is not None and (
        not isinstance(page_count, int) or isinstance(page_count, bool) or page_count < 0
    ):
        errors.append("page_count must be null or a non-negative integer")
        page_count = None
    if status == "complete" and page_count is not None and len(pages) != page_count:
        errors.append(f"page count mismatch: extracted={len(pages)} expected={page_count}")
    if status == "truncated" and page_count is not None and len(pages) >= page_count:
        errors.append("truncated extraction must contain fewer units than page_count")

    low_text_units = []
    low_confidence_units = []
    missing_confidence_units = []
    text_chars = 0
    review_units = 0
    manual_review_units = 0
    review_action_counts = Counter()
    adaptive_units = 0
    high_resolution_selected = 0
    confidences = []
    page_numbers = []

    for index, page in enumerate(pages, start=1):
        if not isinstance(page, dict):
            errors.append(f"unit {index} is not an object")
            continue
        raw_locator = page.get("locator")
        if not isinstance(raw_locator, str) or not raw_locator.strip():
            errors.append(f"unit {index}: locator must be a non-empty string")
            locator = f"unit {index}"
        else:
            locator = raw_locator
        text = page.get("text")
        if not isinstance(text, str):
            errors.append(f"{locator}: text must be a string")
            continue
        text_chars += len(text)
        if len(text.strip()) < 20:
            low_text_units.append(locator)

        page_number = page.get("page")
        if page_number is not None:
            if not isinstance(page_number, int) or isinstance(page_number, bool) or page_number < 1:
                errors.append(f"{locator}: page must be null or a positive integer")
            else:
                page_numbers.append(page_number)
                if page_number != index:
                    errors.append(f"{locator}: expected sequential page number {index}")

        mode = page.get("extraction_mode")
        if not isinstance(mode, str) or not mode:
            errors.append(f"{locator}: extraction_mode must be a non-empty string")
            mode = ""
        confidence = page.get("ocr_confidence")
        if mode.startswith("ocr"):
            if confidence is None:
                missing_confidence_units.append(locator)
            elif not finite_number(confidence) or not 0.0 <= float(confidence) <= 1.0:
                errors.append(f"{locator}: OCR confidence must be between 0 and 1")
            else:
                confidence = float(confidence)
                confidences.append(confidence)
                if confidence < min_confidence:
                    low_confidence_units.append(f"{locator} ({confidence:.2f})")
            attempt_errors, attempt_warnings, attempt_metrics = validate_attempts(
                page,
                locator,
                require_review_actions,
            )
            errors.extend(attempt_errors)
            warnings.extend(attempt_warnings)
            adaptive_units += int(attempt_metrics["adaptive"])
            high_resolution_selected += int(attempt_metrics["high_resolution_selected"])

        reasons = page.get("review_reasons", [])
        if not isinstance(reasons, list) or any(not isinstance(reason, str) for reason in reasons):
            errors.append(f"{locator}: review_reasons must be a list of strings")
            reasons = []
        elif reasons:
            review_units += 1

        if require_review_actions:
            manual_flags = page.get("manual_review_flags")
            if not isinstance(manual_flags, list) or any(
                not isinstance(flag, str) or not flag for flag in manual_flags
            ):
                errors.append(f"{locator}: manual_review_flags must be a list of strings")
                manual_flags = []
            if manual_flags:
                manual_review_units += 1
                missing_manual_reasons = [
                    flag for flag in manual_flags if f"manual:{flag}" not in reasons
                ]
                if missing_manual_reasons:
                    errors.append(f"{locator}: manual review flags are missing matching reasons")

            action = page.get("recommended_action")
            if action is not None and action not in VALID_REVIEW_ACTIONS:
                errors.append(f"{locator}: invalid recommended_action {action!r}")
            elif action is not None:
                review_action_counts[action] += 1
            if reasons and action is None:
                errors.append(f"{locator}: review reasons require a recommended_action")
            if not reasons and action is not None:
                errors.append(f"{locator}: recommended_action requires review reasons")
            if manual_flags and action != "manual-verify-original":
                errors.append(f"{locator}: manual review flags require original-page verification")

            layout_table = page.get("layout_table")
            if layout_table is not None:
                if not isinstance(layout_table, dict):
                    errors.append(f"{locator}: layout_table must be an object")
                else:
                    rows = layout_table.get("rows")
                    row_count = layout_table.get("row_count")
                    if not isinstance(rows, list):
                        errors.append(f"{locator}: layout_table rows must be a list")
                    elif row_count != len(rows):
                        errors.append(f"{locator}: layout_table row_count does not match rows")
                    elif any(
                        not isinstance(row, dict) or not isinstance(row.get("cells"), list)
                        for row in rows
                    ):
                        errors.append(f"{locator}: layout_table contains an invalid row")

    if len(page_numbers) != len(set(page_numbers)):
        errors.append("duplicate page numbers")
    if page_count is not None and page_numbers and page_numbers != list(range(1, len(pages) + 1)):
        errors.append("page sequence is not contiguous from 1")
    if low_text_units:
        warnings.append(f"low-text units: {', '.join(low_text_units[:12])}")
    if low_confidence_units:
        warnings.append(f"low OCR confidence: {', '.join(low_confidence_units[:12])}")
    if missing_confidence_units:
        warnings.append(f"missing OCR confidence: {', '.join(missing_confidence_units[:12])}")
    if not pages:
        errors.append("no extracted units")
    elif text_chars == 0:
        errors.append("extraction contains no text")

    metrics = {
        "source": payload.get("relative_path"),
        "status": status,
        "config_fingerprint": fingerprint,
        "units": len(pages),
        "page_count": page_count,
        "text_chars": text_chars,
        "ocr_units": len(confidences) + len(missing_confidence_units),
        "adaptive_units": adaptive_units,
        "high_resolution_selected": high_resolution_selected,
        "review_units": review_units,
        "manual_review_units": manual_review_units,
        "review_action_counts": dict(sorted(review_action_counts.items())),
        "average_confidence": round(sum(confidences) / len(confidences), 4) if confidences else None,
        "minimum_confidence": round(min(confidences), 4) if confidences else None,
    }
    return errors, warnings, metrics, payload


def reconcile_record(record: dict, payload: dict, summary: dict | None) -> list[str]:
    errors = []
    comparisons = {
        "source": payload.get("relative_path"),
        "status": payload.get("status"),
        "units_completed": payload.get("units_completed"),
        "page_count": payload.get("page_count"),
    }
    for field, actual in comparisons.items():
        if record.get(field) != actual:
            errors.append(f"summary {field} does not match extraction output")
    expected_review_count = sum(
        1 for page in payload.get("pages", []) if isinstance(page, dict) and page.get("review_reasons")
    )
    if record.get("review_queue_count") != expected_review_count:
        errors.append("summary review_queue_count does not match extraction output")
    if summary is not None:
        summary_fingerprint = summary.get("extraction_config_fingerprint")
        payload_fingerprint = payload.get("extraction_config_fingerprint")
        if summary_fingerprint != payload_fingerprint:
            errors.append("summary configuration fingerprint does not match extraction output")
    return errors


def expected_review_entries(payloads: list[dict]) -> list[dict]:
    entries = []
    for payload in payloads:
        for page in payload.get("pages", []):
            if not isinstance(page, dict):
                continue
            reasons = page.get("review_reasons") or []
            if not reasons:
                continue
            entry = {
                "source": payload.get("relative_path"),
                "page": page.get("page"),
                "locator": page.get("locator"),
                "priority": page.get("review_priority"),
                "reasons": reasons,
                "material_signals": page.get("material_signals", []),
                "ocr_confidence": page.get("ocr_confidence"),
                "ocr_scale": page.get("ocr_scale"),
            }
            if supports_review_actions(payload):
                entry["manual_review_flags"] = page.get("manual_review_flags", [])
                entry["recommended_action"] = page.get("recommended_action")
            entries.append(entry)
    return entries


def validate_review_queue(root: Path, summary: dict | None, payloads: list[dict]) -> list[str]:
    if summary is None or summary.get("schema_version") != 2:
        return []
    path = root / "review-queue.json"
    if not path.is_file():
        return ["missing review-queue.json for schema version 2 batch"]
    try:
        queue = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        return [f"invalid review queue: {exc}"]
    if not isinstance(queue, dict):
        return ["review queue must be an object"]

    errors = []
    entries = queue.get("entries")
    if not isinstance(entries, list):
        errors.append("review queue entries must be a list")
        entries = []
    for index, entry in enumerate(entries, start=1):
        if not isinstance(entry, dict):
            errors.append(f"review queue entry {index} is not an object")
        elif "text" in entry:
            errors.append(f"review queue entry {index} must not contain extracted text")

    if queue.get("schema_version") != summary.get("schema_version"):
        errors.append("review queue schema version does not match summary")
    if queue.get("extraction_config_fingerprint") != summary.get(
        "extraction_config_fingerprint"
    ):
        errors.append("review queue configuration fingerprint does not match summary")
    if queue.get("entry_count") != len(entries):
        errors.append("review queue entry_count does not match entries")
    high_count = sum(
        1 for entry in entries if isinstance(entry, dict) and entry.get("priority") == "high"
    )
    if queue.get("high_priority_count") != high_count:
        errors.append("review queue high_priority_count does not match entries")
    if summary.get("review_queue_count") != len(entries):
        errors.append("summary review_queue_count does not match review queue")

    if any(supports_review_actions(payload) for payload in payloads):
        action_counts = dict(
            sorted(
                Counter(
                    entry.get("recommended_action")
                    for entry in entries
                    if isinstance(entry, dict) and entry.get("recommended_action")
                ).items()
            )
        )
        if queue.get("action_counts") != action_counts:
            errors.append("review queue action_counts does not match entries")
        if summary.get("review_action_counts") != action_counts:
            errors.append("summary review_action_counts does not match review queue")

    expected = expected_review_entries(payloads)
    if entries != expected:
        errors.append("review queue entries do not match extraction outputs")
    return errors


def validate_review_ledger(
    extraction_root: Path,
    ledger_path: Path,
    require_resolved: bool,
) -> tuple[list[str], list[str], dict]:
    errors = []
    warnings = []
    metrics = {
        "required_action_count": 0,
        "ledger_entry_count": 0,
        "resolved_count": 0,
        "unclear_count": 0,
        "pending_count": 0,
        "missing_count": 0,
    }
    try:
        queue = json.loads((extraction_root / "review-queue.json").read_text(encoding="utf-8"))
        ledger = json.loads(ledger_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        return [f"invalid review ledger input: {exc}"], warnings, metrics
    if not isinstance(queue, dict) or not isinstance(ledger, dict):
        return ["review queue and review ledger must be objects"], warnings, metrics

    if ledger.get("schema_version") != 1:
        errors.append("review ledger schema_version must be 1")
    queue_fingerprint = queue.get("extraction_config_fingerprint")
    if ledger.get("extraction_config_fingerprint") != queue_fingerprint:
        errors.append("review ledger fingerprint does not match extraction")

    queue_entries = queue.get("entries")
    if not isinstance(queue_entries, list):
        return [*errors, "review queue entries must be a list"], warnings, metrics
    queue_index = {
        (entry.get("source"), entry.get("page"), entry.get("recommended_action")): entry
        for entry in queue_entries
        if isinstance(entry, dict)
    }
    required_keys = {
        key for key in queue_index if key[2] in REQUIRED_LEDGER_ACTIONS
    }
    metrics["required_action_count"] = len(required_keys)

    entries = ledger.get("entries")
    if not isinstance(entries, list):
        return [*errors, "review ledger entries must be a list"], warnings, metrics
    metrics["ledger_entry_count"] = len(entries)
    ledger_index = {}
    for index, entry in enumerate(entries, start=1):
        prefix = f"review ledger entry {index}"
        if not isinstance(entry, dict):
            errors.append(f"{prefix} must be an object")
            continue
        source = entry.get("source")
        page = entry.get("page")
        action = entry.get("action")
        status = entry.get("status")
        if not isinstance(source, str) or not source:
            errors.append(f"{prefix} requires source")
        if page is not None and (
            not isinstance(page, int) or isinstance(page, bool) or page < 1
        ):
            errors.append(f"{prefix} page must be null or a positive integer")
        if action not in VALID_REVIEW_ACTIONS:
            errors.append(f"{prefix} has invalid action {action!r}")
        if status not in VALID_LEDGER_STATUSES:
            errors.append(f"{prefix} has invalid status {status!r}")

        key = (source, page, action)
        if key in ledger_index:
            errors.append(f"{prefix} duplicates an earlier source/page/action")
        else:
            ledger_index[key] = entry
        if key not in queue_index:
            errors.append(f"{prefix} does not match the current review queue")

        if status in VALID_LEDGER_STATUSES:
            metrics[f"{status}_count"] += 1
        if status in {"resolved", "unclear"} and action in REQUIRED_LEDGER_ACTIONS:
            if entry.get("original_checked") is not True:
                errors.append(f"{prefix} must confirm original_checked")
        if status == "unclear" and not str(entry.get("notes", "")).strip():
            errors.append(f"{prefix} with unclear status requires notes")

    missing = sorted(
        required_keys - set(ledger_index),
        key=lambda value: (str(value[0]), str(value[1])),
    )
    pending = [
        key
        for key in required_keys.intersection(ledger_index)
        if ledger_index[key].get("status") == "pending"
    ]
    metrics["missing_count"] = len(missing)
    if missing:
        message = f"review ledger is missing {len(missing)} required action(s)"
        (errors if require_resolved else warnings).append(message)
    if pending:
        message = f"review ledger has {len(pending)} pending required action(s)"
        (errors if require_resolved else warnings).append(message)
    return errors, warnings, metrics


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path, help="Directory containing extraction JSON files")
    parser.add_argument("--require-complete", action="store_true")
    parser.add_argument("--require-config", action="store_true")
    parser.add_argument("--min-confidence", type=float, default=0.80)
    parser.add_argument("--strict", action="store_true", help="Treat warnings as failures")
    parser.add_argument("--fail-on-review", action="store_true", help="Fail when review pages remain")
    parser.add_argument("--review-ledger", type=Path, help="Human review resolution JSON")
    parser.add_argument(
        "--require-actions-resolved",
        action="store_true",
        help="Fail when required original-page actions are missing or pending",
    )
    parser.add_argument("--allow-empty", action="store_true", help="Allow a batch with no outputs")
    parser.add_argument("--output", type=Path, help="Write machine-readable validation JSON")
    args = parser.parse_args()

    if not math.isfinite(args.min_confidence) or not 0.0 <= args.min_confidence <= 1.0:
        parser.error("--min-confidence must be finite and between 0 and 1")
    if args.require_actions_resolved and args.review_ledger is None:
        parser.error("--require-actions-resolved requires --review-ledger")
    root = args.input.resolve()
    if not root.is_dir():
        parser.error(f"input directory does not exist: {root}")
    if args.output and args.output.resolve().parent == root:
        parser.error("--output must be outside the extraction directory")

    files, record_map, summary, batch_errors = load_batch(root)
    if summary is None and not files and not batch_errors:
        parser.error(f"no extraction JSON files found in: {root}")
    if summary is not None and not files and not batch_errors and not args.allow_empty:
        parser.error("extraction summary contains no successful outputs")

    results = []
    total_errors = len(batch_errors)
    total_warnings = 0
    total_review_units = 0
    payloads = []
    for path in files:
        errors, warnings, metrics, payload = validate_file(
            path,
            args.require_complete,
            args.min_confidence,
            args.require_config,
        )
        record = record_map.get(path.name)
        if record is not None and payload is not None:
            errors.extend(reconcile_record(record, payload, summary))
        if payload is not None:
            payloads.append(payload)
        total_errors += len(errors)
        total_warnings += len(warnings)
        total_review_units += metrics.get("review_units", 0)
        results.append(
            {
                "file": path.name,
                "errors": errors,
                "warnings": warnings,
                "metrics": metrics,
            }
        )
        status = "ERROR" if errors else "WARN" if warnings else "OK"
        print(f"[{status}] {path.name}")
        for message in errors:
            print(f"  error: {message}")
        for message in warnings:
            print(f"  warning: {message}")

    queue_errors = validate_review_queue(root, summary, payloads)
    batch_errors.extend(queue_errors)
    total_errors += len(queue_errors)
    ledger_metrics = None
    if args.review_ledger is not None:
        ledger_errors, ledger_warnings, ledger_metrics = validate_review_ledger(
            root,
            args.review_ledger.resolve(),
            args.require_actions_resolved,
        )
        batch_errors.extend(ledger_errors)
        total_errors += len(ledger_errors)
        total_warnings += len(ledger_warnings)
        for message in ledger_warnings:
            print(f"[WARN] review ledger: {message}")
    for message in batch_errors:
        print(f"[ERROR] batch: {message}")
    report = {
        "directory": str(root),
        "file_count": len(files),
        "error_count": total_errors,
        "warning_count": total_warnings,
        "review_unit_count": total_review_units,
        "review_ledger": ledger_metrics,
        "batch_errors": batch_errors,
        "results": results,
    }
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(
        f"Validated {len(files)} files: {total_errors} errors, {total_warnings} warnings, "
        f"{total_review_units} review units"
    )
    failed = total_errors or (args.strict and total_warnings) or (args.fail_on_review and total_review_units)
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
