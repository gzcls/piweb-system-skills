#!/usr/bin/env python3
"""Extract contract text with embedded-text-first and adaptive local OCR."""

from __future__ import annotations

import argparse
import difflib
import hashlib
import importlib.metadata
import importlib.util
import json
import math
import re
import statistics
import sys
import tempfile
import time
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree

SCHEMA_VERSION = 2
EXTRACTOR_VERSION = 4
SUPPORTED_EXTENSIONS = {
    ".pdf",
    ".docx",
    ".doc",
    ".txt",
    ".md",
    ".png",
    ".jpg",
    ".jpeg",
    ".tif",
    ".tiff",
}
IGNORED_DIRECTORIES = {
    ".contract-analysis",
    ".contract-tools",
    ".git",
    ".tools",
    "extracted_text",
    "node_modules",
}
IGNORED_FILENAMES = {
    "agents.md",
    "claude.md",
    "contributing.md",
    "readme.md",
}

QUALITY_PROFILES = {
    "fast": {
        "primary_scale": 1.25,
        "adaptive_ocr": False,
        "second_pass_scale": 2.0,
        "second_pass_confidence": 0.80,
        "second_pass_min_chars": 40,
        "second_pass_material": False,
        "review_confidence": 0.80,
        "max_render_megapixels": 8.0,
    },
    "hybrid": {
        "primary_scale": 1.25,
        "adaptive_ocr": True,
        "second_pass_scale": 2.5,
        "second_pass_confidence": 0.80,
        "second_pass_min_chars": 80,
        "second_pass_material": False,
        "review_confidence": 0.86,
        "max_render_megapixels": 12.0,
    },
    "balanced": {
        "primary_scale": 1.5,
        "adaptive_ocr": True,
        "second_pass_scale": 2.5,
        "second_pass_confidence": 0.88,
        "second_pass_min_chars": 80,
        "second_pass_material": False,
        "review_confidence": 0.86,
        "max_render_megapixels": 12.0,
    },
    "high": {
        "primary_scale": 1.5,
        "adaptive_ocr": True,
        "second_pass_scale": 3.0,
        "second_pass_confidence": 0.92,
        "second_pass_min_chars": 120,
        "second_pass_material": True,
        "review_confidence": 0.90,
        "max_render_megapixels": 16.0,
    },
}

LOW_CONFIDENCE_COVERAGE_THRESHOLD = 0.82
QUALITY_REVIEW_REASONS = frozenset(
    {
        "empty-text",
        "ocr-empty",
        "low-text",
        "missing-confidence",
        "low-confidence",
        "attempt-disagreement",
    }
)
NUMERIC_TOKEN_PATTERN = re.compile(r"\d[\d,.]*|[一二三四五六七八九十百千万]+")
MATERIAL_VALUE_PATTERN = re.compile(
    r"(?:[¥￥$]\s*\d[\d,.]*|\d[\d,.]*\s*(?:%|％|元|万元|天|日|个?月|年|小时))",
    re.IGNORECASE,
)
TABLE_KEYWORD_PATTERN = re.compile(
    r"GMV|GSV|NET|KPI|权重|考核|合计|销售目标|业绩目标|佣金|费率|"
    r"score|weight|target|commission|total",
    re.IGNORECASE,
)
CERTIFICATE_PATTERN = re.compile(
    r"营业执照|business license|certificate of incorporation|company registration",
    re.IGNORECASE,
)
CERTIFICATE_FIELD_PATTERN = re.compile(
    r"统一?社会信用代码|成立日期|登记机关|\(副本\)|（副本）|"
    r"registration number|incorporation date|registrar",
    re.IGNORECASE,
)
SIGNATURE_PATTERN = re.compile(
    r"(?:甲方|乙方|party\s+[ab])\s*[（(]?(?:盖章|签章|签字|signature)|"
    r"(?:盖章|签章|签字|signature)\s*[）)]?[:：]?\s*(?:甲方|乙方|party\s+[ab])",
    re.IGNORECASE,
)

MATERIAL_PATTERNS = (
    (
        "money",
        re.compile(
            r"人民币|金额|费用|价格|价款|税费|发票|付款|支付|结算|"
            r"[¥￥]|\d[\d,.]*\s*(?:元|万元|%|％)"
        ),
    ),
    (
        "date-term",
        re.compile(
            r"\d{4}\s*年\s*\d{1,2}\s*月|期限|有效期|续期|续约|通知期|"
            r"工作日|自然日|calendar days?|business days?",
            re.IGNORECASE,
        ),
    ),
    (
        "termination",
        re.compile(r"终止|解除|到期|撤销|terminate|termination|cancel", re.IGNORECASE),
    ),
    (
        "liability",
        re.compile(r"责任|赔偿|违约|补偿|免责|上限|indemn|liabil|damages", re.IGNORECASE),
    ),
    (
        "data-ip",
        re.compile(
            r"数据|个人信息|隐私|保密|知识产权|著作权|商标|专利|"
            r"confidential|privacy|intellectual property",
            re.IGNORECASE,
        ),
    ),
)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def safe_stem(value: str) -> str:
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value).strip(" .") or "contract"


def write_json(path: Path, payload: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            dir=path.parent,
            prefix=path.name + ".",
            suffix=".tmp",
            delete=False,
        ) as handle:
            json.dump(payload, handle, ensure_ascii=False, indent=2)
            temporary_path = Path(handle.name)
        temporary_path.replace(path)
    finally:
        if temporary_path is not None and temporary_path.exists():
            temporary_path.unlink()


def add_dependency_paths(input_root: Path, vendor_dirs: list[Path]) -> None:
    candidates = [
        input_root / ".contract-tools" / "runtime",
        input_root / ".contract-tools" / "pymupdf",
        input_root / ".contract-tools" / "ocr",
        input_root / ".tools" / "pymupdf",
        input_root / ".tools" / "ocr",
        *vendor_dirs,
    ]
    for path in candidates:
        resolved = path.resolve()
        if resolved.exists() and str(resolved) not in sys.path:
            sys.path.insert(0, str(resolved))


def package_version(distribution: str) -> str | None:
    try:
        return importlib.metadata.version(distribution)
    except importlib.metadata.PackageNotFoundError:
        return None


def process_memory_snapshot() -> dict:
    """Return best-effort process RSS metrics without adding a dependency."""
    if sys.platform == "win32":
        try:
            import ctypes
            from ctypes import wintypes

            size_type = ctypes.c_size_t

            class ProcessMemoryCounters(ctypes.Structure):
                _fields_ = [
                    ("cb", wintypes.DWORD),
                    ("PageFaultCount", wintypes.DWORD),
                    ("PeakWorkingSetSize", size_type),
                    ("WorkingSetSize", size_type),
                    ("QuotaPeakPagedPoolUsage", size_type),
                    ("QuotaPagedPoolUsage", size_type),
                    ("QuotaPeakNonPagedPoolUsage", size_type),
                    ("QuotaNonPagedPoolUsage", size_type),
                    ("PagefileUsage", size_type),
                    ("PeakPagefileUsage", size_type),
                ]

            counters = ProcessMemoryCounters()
            counters.cb = ctypes.sizeof(counters)
            kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
            psapi = ctypes.WinDLL("psapi", use_last_error=True)
            kernel32.GetCurrentProcess.restype = wintypes.HANDLE
            psapi.GetProcessMemoryInfo.argtypes = [
                wintypes.HANDLE,
                ctypes.POINTER(ProcessMemoryCounters),
                wintypes.DWORD,
            ]
            psapi.GetProcessMemoryInfo.restype = wintypes.BOOL
            if psapi.GetProcessMemoryInfo(
                kernel32.GetCurrentProcess(),
                ctypes.byref(counters),
                counters.cb,
            ):
                return {
                    "current_rss_bytes": int(counters.WorkingSetSize),
                    "peak_rss_bytes": int(counters.PeakWorkingSetSize),
                    "source": "windows-process-memory-counters",
                }
        except (AttributeError, OSError, ValueError):
            pass
    else:
        try:
            import resource

            peak = int(resource.getrusage(resource.RUSAGE_SELF).ru_maxrss)
            if sys.platform != "darwin":
                peak *= 1024
            return {
                "current_rss_bytes": None,
                "peak_rss_bytes": peak,
                "source": "getrusage",
            }
        except (ImportError, OSError, ValueError):
            pass
    return {
        "current_rss_bytes": None,
        "peak_rss_bytes": None,
        "source": "unavailable",
    }


def finite_float(value: str) -> float:
    parsed = float(value)
    if not math.isfinite(parsed):
        raise argparse.ArgumentTypeError("value must be finite")
    return parsed


def nonnegative_int(value: str) -> int:
    parsed = int(value)
    if parsed < 0:
        raise argparse.ArgumentTypeError("value must be zero or greater")
    return parsed


def positive_int(value: str) -> int:
    parsed = int(value)
    if parsed < 1:
        raise argparse.ArgumentTypeError("value must be positive")
    return parsed


def resolve_model_paths(args: argparse.Namespace) -> dict[str, Path]:
    paths = {}
    for kind in ("det", "rec", "cls"):
        supplied = getattr(args, f"{kind}_model_path")
        if supplied is None:
            continue
        resolved = supplied.resolve()
        if not resolved.is_file():
            raise ValueError(f"{kind} model does not exist: {resolved}")
        paths[kind] = resolved
    return paths


def default_model_paths() -> tuple[Path | None, dict[str, Path]]:
    spec = importlib.util.find_spec("rapidocr_onnxruntime")
    if spec is None or spec.origin is None:
        return None, {}
    package_root = Path(spec.origin).resolve().parent
    names = {
        "det": "ch_PP-OCRv3_det_infer.onnx",
        "rec": "ch_PP-OCRv3_rec_infer.onnx",
        "cls": "ch_ppocr_mobile_v2.0_cls_infer.onnx",
    }
    return package_root, {
        kind: package_root / "models" / filename
        for kind, filename in names.items()
        if (package_root / "models" / filename).is_file()
    }


def model_manifest(model_paths: dict[str, Path]) -> dict:
    package_root, defaults = default_model_paths()
    resolved_models = {**defaults, **model_paths}
    manifest = {
        "source": "custom-compatible-onnx" if model_paths else "rapidocr-default",
        "rapidocr_onnxruntime": package_version("rapidocr_onnxruntime"),
        "models": {
            kind: {
                "filename": path.name,
                "size_bytes": path.stat().st_size,
                "sha256": sha256_file(path),
                "custom": kind in model_paths,
            }
            for kind, path in sorted(resolved_models.items())
        },
    }
    if package_root is not None:
        config_path = package_root / "config.yaml"
        if config_path.is_file():
            manifest["engine_config"] = {
                "filename": config_path.name,
                "sha256": sha256_file(config_path),
            }
    return manifest


def resolve_quality_settings(args: argparse.Namespace) -> dict:
    settings = dict(QUALITY_PROFILES[args.quality_profile])
    overrides = {
        "primary_scale": args.scale,
        "adaptive_ocr": args.adaptive_ocr,
        "second_pass_scale": args.second_pass_scale,
        "second_pass_confidence": args.second_pass_confidence,
        "second_pass_min_chars": args.second_pass_min_chars,
        "second_pass_material": args.second_pass_material,
        "review_confidence": args.review_confidence,
        "max_render_megapixels": args.max_render_megapixels,
    }
    settings.update({key: value for key, value in overrides.items() if value is not None})
    settings["quality_profile"] = args.quality_profile

    for name in ("primary_scale", "second_pass_scale"):
        if not math.isfinite(settings[name]) or settings[name] < 0.5:
            raise ValueError(f"{name} must be at least 0.5")
    for name in ("second_pass_confidence", "review_confidence"):
        if not math.isfinite(settings[name]) or not 0.0 <= settings[name] <= 1.0:
            raise ValueError(f"{name} must be between 0 and 1")
    if settings["second_pass_min_chars"] < 1:
        raise ValueError("second_pass_min_chars must be positive")
    if not math.isfinite(settings["max_render_megapixels"]) or settings["max_render_megapixels"] < 1.0:
        raise ValueError("max_render_megapixels must be at least 1")
    return settings


def build_extraction_config(
    args: argparse.Namespace,
    settings: dict,
    models: dict,
) -> dict:
    return {
        "schema_version": SCHEMA_VERSION,
        "extractor_version": EXTRACTOR_VERSION,
        "ocr_mode": args.ocr,
        "page_limit": args.max_pages,
        "min_text_chars": max(1, args.min_text_chars),
        **settings,
        "dependencies": {
            "PyPDF2": package_version("PyPDF2"),
            "PyMuPDF": package_version("PyMuPDF"),
            "rapidocr_onnxruntime": package_version("rapidocr_onnxruntime"),
            "onnxruntime": package_version("onnxruntime"),
            "Pillow": package_version("Pillow"),
        },
        "model_bundle": models,
    }


def config_fingerprint(config: dict) -> str:
    encoded = json.dumps(config, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(encoded.encode("utf-8")).hexdigest()


def compact_text(text: str) -> str:
    return re.sub(r"\s+", "", text)


def detect_material_signals(text: str) -> list[str]:
    return [name for name, pattern in MATERIAL_PATTERNS if pattern.search(text)]


def line_texts(lines: list[dict] | list[str]) -> list[str]:
    values = []
    for line in lines:
        value = line.get("text", "") if isinstance(line, dict) else line
        compact = compact_text(str(value))
        if compact:
            values.append(compact)
    return values


def attempt_diagnostics(text: str, lines: list[dict] | list[str]) -> dict:
    compact = compact_text(text)
    values = line_texts(lines)
    short_lines = sum(len(value) <= 12 for value in values)
    return {
        "numeric_token_count": len(NUMERIC_TOKEN_PATTERN.findall(compact)),
        "material_value_count": len(MATERIAL_VALUE_PATTERN.findall(compact)),
        "short_line_ratio": round(short_lines / len(values), 4) if values else 0.0,
    }


def build_layout_table(lines: list[dict]) -> dict | None:
    cells = []
    for line in lines:
        if not isinstance(line, dict) or not compact_text(str(line.get("text", ""))):
            continue
        box = line.get("box")
        if not isinstance(box, list) or len(box) < 4:
            continue
        try:
            xs = [float(point[0]) for point in box]
            ys = [float(point[1]) for point in box]
        except (IndexError, TypeError, ValueError):
            continue
        height = max(ys) - min(ys)
        cells.append(
            {
                "text": line["text"],
                "confidence": line.get("confidence"),
                "box": box,
                "left": min(xs),
                "center_y": (min(ys) + max(ys)) / 2,
                "height": max(height, 1.0),
            }
        )
    if not cells:
        return None

    tolerance = max(2.0, statistics.median(cell["height"] for cell in cells) * 0.65)
    clustered_rows = []
    for cell in sorted(cells, key=lambda item: (item["center_y"], item["left"])):
        if not clustered_rows:
            clustered_rows.append([cell])
            continue
        current = clustered_rows[-1]
        row_center = statistics.fmean(item["center_y"] for item in current)
        if abs(cell["center_y"] - row_center) <= tolerance:
            current.append(cell)
        else:
            clustered_rows.append([cell])

    rows = []
    for index, row in enumerate(clustered_rows, start=1):
        ordered = sorted(row, key=lambda item: item["left"])
        rows.append(
            {
                "row": index,
                "text": " | ".join(item["text"] for item in ordered),
                "cells": [
                    {
                        "text": item["text"],
                        "confidence": item["confidence"],
                        "box": item["box"],
                    }
                    for item in ordered
                ],
            }
        )
    return {
        "method": "ocr-box-row-clustering-v1",
        "row_count": len(rows),
        "rows": rows,
    }


def detect_manual_review_flags(text: str, lines: list[dict] | list[str]) -> list[str]:
    compact = compact_text(text)
    diagnostics = attempt_diagnostics(text, lines)
    flags = []
    if CERTIFICATE_PATTERN.search(compact) and CERTIFICATE_FIELD_PATTERN.search(compact):
        flags.append("identity-document")

    has_structured_signing_block = (
        ("甲方" in compact and "乙方" in compact)
        and bool(
            re.search(
                r"法定代表人|授权签约人|代表人[:：]|authorized representative",
                compact,
                re.IGNORECASE,
            )
        )
        and bool(
            re.search(
                r"\d{4}.{0,2}年|\d{4}[-/.]\d{1,2}|date[:：]",
                compact,
                re.IGNORECASE,
            )
        )
    )
    if SIGNATURE_PATTERN.search(compact) or has_structured_signing_block:
        flags.append("signature-or-stamp")

    values = line_texts(lines)
    if (
        len(values) >= 20
        and diagnostics["short_line_ratio"] >= 0.45
        and diagnostics["numeric_token_count"] >= 6
        and TABLE_KEYWORD_PATTERN.search(compact)
    ):
        flags.append("complex-table")
    return flags


def text_similarity(first: str, second: str) -> float:
    left = compact_text(first)[:5000]
    right = compact_text(second)[:5000]
    if not left and not right:
        return 1.0
    return difflib.SequenceMatcher(None, left, right, autojunk=False).ratio()


def attempt_quality_score(text: str, confidence: float | None, line_count: int) -> float:
    confidence_component = (confidence or 0.0) * 0.90
    coverage_component = min(len(compact_text(text)) / 800.0, 1.0) * 0.08
    line_component = min(line_count / 20.0, 1.0) * 0.02
    return round(confidence_component + coverage_component + line_component, 6)


def use_low_confidence_coverage_selection(attempts: list[dict]) -> bool:
    confidences = [
        float(attempt["confidence"])
        for attempt in attempts
        if attempt.get("confidence") is not None
    ]
    return len(attempts) > 1 and (
        not confidences or max(confidences) < LOW_CONFIDENCE_COVERAGE_THRESHOLD
    )


def attempt_structure_value(attempt: dict, field: str) -> int:
    value = attempt.get(field)
    if isinstance(value, int) and not isinstance(value, bool) and value >= 0:
        return value
    diagnostics = attempt_diagnostics(attempt.get("text", ""), attempt.get("lines", []))
    return diagnostics[field]


def select_best_attempt(attempts: list[dict]) -> int:
    if use_low_confidence_coverage_selection(attempts):
        return max(
            range(len(attempts)),
            key=lambda index: (
                attempt_structure_value(attempts[index], "material_value_count"),
                attempt_structure_value(attempts[index], "numeric_token_count"),
                attempts[index]["text_chars"],
                attempts[index]["line_count"],
                attempts[index]["confidence"]
                if attempts[index]["confidence"] is not None
                else -1.0,
                -index,
            ),
        )
    return max(
        range(len(attempts)),
        key=lambda index: (
            attempts[index]["quality_score"],
            attempts[index]["confidence"] if attempts[index]["confidence"] is not None else -1.0,
            attempts[index]["text_chars"],
            -index,
        ),
    )


def second_pass_reasons(attempt: dict, settings: dict, material_signals: list[str]) -> list[str]:
    reasons = []
    if attempt["confidence"] is None:
        reasons.append("missing-confidence")
    elif attempt["confidence"] < settings["second_pass_confidence"]:
        reasons.append("low-confidence")
    if attempt["text_chars"] < settings["second_pass_min_chars"]:
        reasons.append("low-text")
    if settings["second_pass_material"] and material_signals:
        reasons.append("material-page")
    return reasons


def review_reasons(
    selected: dict,
    settings: dict,
    material_signals: list[str],
    attempt_similarity: float | None,
    manual_review_flags: list[str] | None = None,
) -> list[str]:
    reasons = []
    if selected["confidence"] is None:
        reasons.append("missing-confidence")
    elif selected["confidence"] < settings["review_confidence"]:
        reasons.append("low-confidence")
    if selected["text_chars"] < settings["second_pass_min_chars"]:
        reasons.append("low-text")
    if attempt_similarity is not None and attempt_similarity < 0.92:
        reasons.append("attempt-disagreement")
    reasons.extend(f"material:{signal}" for signal in material_signals)
    reasons.extend(f"manual:{flag}" for flag in (manual_review_flags or []))
    return list(dict.fromkeys(reasons))


def recommended_review_action(
    reasons: list[str],
    manual_review_flags: list[str],
    attempt_count: int,
    adaptive_enabled: bool = False,
) -> str | None:
    if manual_review_flags:
        return "manual-verify-original"
    if QUALITY_REVIEW_REASONS.intersection(reasons):
        if attempt_count <= 1 and not adaptive_enabled:
            return "enhanced-ocr-then-verify"
        return "manual-verify-original"
    if any(reason.startswith("material:") for reason in reasons):
        return "verify-material-fields"
    return None


def review_priority(reasons: list[str]) -> str | None:
    if not reasons:
        return None
    high_markers = {
        "empty-text",
        "ocr-empty",
        "low-text",
        "missing-confidence",
        "low-confidence",
        "attempt-disagreement",
        "material:money",
        "material:termination",
        "material:liability",
        "manual:identity-document",
        "manual:signature-or-stamp",
        "manual:complex-table",
    }
    return "high" if high_markers.intersection(reasons) else "medium"


class OCRRuntime:
    def __init__(self, model_paths: dict[str, Path], max_megapixels: float):
        self.model_paths = model_paths
        self.max_megapixels = max_megapixels
        self._engine = None
        self._pymupdf = None
        self._numpy = None
        self._cv2 = None

    def require(self) -> None:
        if self._engine is not None:
            return
        try:
            import cv2
            import numpy
            import pymupdf
            from rapidocr_onnxruntime import RapidOCR
        except ImportError as exc:
            raise RuntimeError(
                "OCR dependencies are missing. Install the pinned local runtime with:\n"
                "  python -m pip install --target .contract-tools/runtime "
                "-r <skill-dir>/scripts/requirements-ocr.txt"
            ) from exc
        kwargs = {
            f"{kind}_model_path": str(path)
            for kind, path in self.model_paths.items()
        }
        self._numpy = numpy
        self._pymupdf = pymupdf
        self._cv2 = cv2
        self._engine = RapidOCR(**kwargs)

    @staticmethod
    def _normalized_box(raw_box, coordinate_scale: float) -> list[list[float]] | None:
        if raw_box is None:
            return None
        try:
            return [
                [round(float(point[0]) / coordinate_scale, 2), round(float(point[1]) / coordinate_scale, 2)]
                for point in raw_box
            ]
        except (TypeError, ValueError, IndexError):
            return None

    @classmethod
    def normalize_result(cls, raw_result, coordinate_scale: float) -> tuple[list[dict], list[float]]:
        result = raw_result[0] if isinstance(raw_result, tuple) else raw_result
        elapsed = raw_result[1] if isinstance(raw_result, tuple) and len(raw_result) > 1 else None
        engine_times = []
        if isinstance(elapsed, (list, tuple)):
            for value in elapsed:
                try:
                    engine_times.append(float(value))
                except (TypeError, ValueError):
                    pass
        if result is None:
            return [], engine_times

        lines = []
        if hasattr(result, "txts"):
            texts = list(getattr(result, "txts", []))
            scores = list(getattr(result, "scores", []))
            boxes = list(getattr(result, "boxes", []))
            for index, value in enumerate(texts):
                text = str(value).strip()
                if not text:
                    continue
                score = None
                if index < len(scores):
                    try:
                        score = float(scores[index])
                    except (TypeError, ValueError):
                        pass
                raw_box = boxes[index] if index < len(boxes) else None
                lines.append(
                    {
                        "text": text,
                        "confidence": round(score, 6) if score is not None else None,
                        "box": cls._normalized_box(raw_box, coordinate_scale),
                    }
                )
            return lines, engine_times

        for item in result:
            if len(item) < 2:
                continue
            text = str(item[1]).strip()
            if not text:
                continue
            score = None
            if len(item) > 2:
                try:
                    score = float(item[2])
                except (TypeError, ValueError):
                    pass
            lines.append(
                {
                    "text": text,
                    "confidence": round(score, 6) if score is not None else None,
                    "box": cls._normalized_box(item[0], coordinate_scale),
                }
            )
        return lines, engine_times

    def _effective_scale(self, width: float, height: float, requested_scale: float) -> float:
        requested_pixels = width * height * requested_scale * requested_scale
        maximum_pixels = self.max_megapixels * 1_000_000
        if requested_pixels <= maximum_pixels:
            return requested_scale
        effective_scale = math.sqrt(maximum_pixels / max(width * height, 1.0))
        if effective_scale < 0.5:
            raise RuntimeError(
                "source page exceeds the render pixel cap even at scale 0.5; "
                "raise --max-render-megapixels only when sufficient memory is available"
            )
        return effective_scale

    @staticmethod
    def _attempt(
        lines: list[dict],
        requested_scale: float,
        effective_scale: float,
        width: int,
        height: int,
        engine_times: list[float],
    ) -> dict:
        text = "\n".join(line["text"] for line in lines)
        scores = [line["confidence"] for line in lines if line["confidence"] is not None]
        confidence = statistics.fmean(scores) if scores else None
        diagnostics = attempt_diagnostics(text, lines)
        return {
            "requested_scale": round(requested_scale, 4),
            "effective_scale": round(effective_scale, 4),
            "render_width": width,
            "render_height": height,
            "render_megapixels": round(width * height / 1_000_000, 4),
            "confidence": round(confidence, 4) if confidence is not None else None,
            "text_chars": len(text),
            "line_count": len(lines),
            "quality_score": attempt_quality_score(text, confidence, len(lines)),
            **diagnostics,
            "text_sha256": hashlib.sha256(text.encode("utf-8")).hexdigest(),
            "engine_seconds": round(sum(engine_times), 4) if engine_times else None,
            "text": text,
            "lines": lines,
        }

    def pdf_page(self, document, page_index: int, requested_scale: float) -> dict:
        self.require()
        page = document[page_index]
        effective_scale = self._effective_scale(page.rect.width, page.rect.height, requested_scale)
        pixmap = page.get_pixmap(
            matrix=self._pymupdf.Matrix(effective_scale, effective_scale),
            colorspace=self._pymupdf.csRGB,
            alpha=False,
        )
        image = self._numpy.frombuffer(pixmap.samples, dtype=self._numpy.uint8).reshape(
            pixmap.height, pixmap.width, pixmap.n
        )
        lines, engine_times = self.normalize_result(self._engine(image), effective_scale)
        return self._attempt(
            lines,
            requested_scale,
            effective_scale,
            pixmap.width,
            pixmap.height,
            engine_times,
        )

    def image_page_count(self, path: Path) -> int:
        try:
            from PIL import Image
        except ImportError as exc:
            raise RuntimeError("Pillow is required for image extraction") from exc
        with Image.open(path) as image:
            return int(getattr(image, "n_frames", 1))

    def image_frame(self, path: Path, frame_index: int, requested_scale: float) -> dict:
        self.require()
        try:
            from PIL import Image
        except ImportError as exc:
            raise RuntimeError("Pillow is required for image extraction") from exc
        with Image.open(path) as image:
            image.seek(frame_index)
            source = self._numpy.asarray(image.convert("RGB"))[:, :, ::-1].copy()
        source_height, source_width = source.shape[:2]
        effective_scale = self._effective_scale(source_width, source_height, requested_scale)
        if not math.isclose(effective_scale, 1.0):
            source = self._cv2.resize(
                source,
                None,
                fx=effective_scale,
                fy=effective_scale,
                interpolation=self._cv2.INTER_CUBIC,
            )
        height, width = source.shape[:2]
        lines, engine_times = self.normalize_result(self._engine(source), effective_scale)
        return self._attempt(
            lines,
            requested_scale,
            effective_scale,
            width,
            height,
            engine_times,
        )

    def open_pdf(self, path: Path):
        self.require()
        return self._pymupdf.open(path)


def build_ocr_page(
    page_number: int,
    locator: str,
    run_attempt,
    settings: dict,
) -> dict:
    attempts = [run_attempt(settings["primary_scale"])]
    initial_signals = detect_material_signals(attempts[0]["text"])
    rerun_reasons = second_pass_reasons(attempts[0], settings, initial_signals)
    if (
        settings["adaptive_ocr"]
        and rerun_reasons
        and settings["second_pass_scale"] > settings["primary_scale"]
    ):
        attempts.append(run_attempt(settings["second_pass_scale"]))

    selected_index = select_best_attempt(attempts)
    selected = attempts[selected_index]
    similarity = None
    if len(attempts) > 1:
        similarity = text_similarity(attempts[0]["text"], attempts[1]["text"])
    signals = detect_material_signals(selected["text"])
    manual_flags = detect_manual_review_flags(selected["text"], selected.get("lines", []))
    reasons = review_reasons(selected, settings, signals, similarity, manual_flags)
    action = recommended_review_action(
        reasons,
        manual_flags,
        len(attempts),
        settings["adaptive_ocr"],
    )
    selection_reason = "single-pass"
    if len(attempts) > 1:
        if use_low_confidence_coverage_selection(attempts):
            selection_reason = (
                "low-confidence-coverage-second-pass"
                if selected_index == 1
                else "low-confidence-coverage-primary-retained"
            )
        else:
            selection_reason = (
                "higher-quality-second-pass" if selected_index == 1 else "primary-retained"
            )

    page_payload = {
        "page": page_number,
        "locator": locator,
        "extraction_mode": "ocr-adaptive" if len(attempts) > 1 else "ocr",
        "ocr_confidence": selected["confidence"],
        "ocr_scale": selected["effective_scale"],
        "ocr_quality_score": selected["quality_score"],
        "selected_attempt": selected_index + 1,
        "selection_reason": selection_reason,
        "second_pass_reasons": rerun_reasons if len(attempts) > 1 else [],
        "attempt_similarity": round(similarity, 4) if similarity is not None else None,
        "material_signals": signals,
        "manual_review_flags": manual_flags,
        "review_reasons": reasons,
        "review_priority": review_priority(reasons),
        "recommended_action": action,
        "requires_original_review": bool(reasons),
        "ocr_attempts": attempts,
        "text": selected["text"],
    }
    if "complex-table" in manual_flags:
        layout_table = build_layout_table(selected.get("lines", []))
        if layout_table is not None:
            page_payload["layout_table"] = layout_table
    return page_payload


def should_select_ocr(ocr_page: dict, embedded_text: str, ocr_mode: str) -> tuple[bool, str]:
    ocr_text = ocr_page.get("text", "")
    if not compact_text(ocr_text):
        return False, "ocr-empty-fallback"
    if not compact_text(embedded_text):
        return True, "embedded-text-empty"
    if ocr_mode == "always":
        return True, "ocr-explicitly-requested"

    confidence = ocr_page.get("ocr_confidence")
    if confidence is None or confidence < 0.50:
        return False, "ocr-confidence-insufficient"
    ocr_chars = len(compact_text(ocr_text))
    embedded_chars = len(compact_text(embedded_text))
    if ocr_chars < max(20, math.ceil(embedded_chars * 1.10)):
        return False, "embedded-text-more-complete"
    return True, "ocr-more-complete"


def embedded_page(
    page_number: int,
    text: str,
    ocr_candidate: dict | None = None,
    selection_reason: str | None = None,
) -> dict:
    signals = detect_material_signals(text)
    text_lines = text.splitlines()
    manual_flags = detect_manual_review_flags(text, text_lines)
    reasons = []
    text_chars = len(compact_text(text))
    if text_chars == 0:
        reasons.append("empty-text")
    elif text_chars < 20:
        reasons.append("low-text")
    if selection_reason == "ocr-empty-fallback":
        reasons.append("ocr-empty")
    reasons.extend(f"material:{signal}" for signal in signals)
    reasons.extend(f"manual:{flag}" for flag in manual_flags)
    action = recommended_review_action(reasons, manual_flags, 0)
    payload = {
        "page": page_number,
        "locator": f"page {page_number}",
        "extraction_mode": "embedded-text",
        "ocr_confidence": None,
        "ocr_scale": None,
        "material_signals": signals,
        "manual_review_flags": manual_flags,
        "review_reasons": reasons,
        "review_priority": review_priority(reasons),
        "recommended_action": action,
        "requires_original_review": bool(reasons),
        "text": text,
    }
    if ocr_candidate is not None:
        payload["ocr_candidate"] = {
            key: value
            for key, value in ocr_candidate.items()
            if key not in {"page", "locator", "text", "review_reasons", "review_priority"}
        }
        payload["selection_reason"] = selection_reason or "embedded-text-selected"
    return payload


def summarize_quality(pages: list[dict]) -> dict:
    ocr_pages = [page for page in pages if page.get("extraction_mode", "").startswith("ocr")]
    confidences = [
        page["ocr_confidence"]
        for page in ocr_pages
        if page.get("ocr_confidence") is not None
    ]
    adaptive_pages = [page for page in ocr_pages if len(page.get("ocr_attempts", [])) > 1]
    high_resolution_selected = [
        page
        for page in adaptive_pages
        if page.get("selected_attempt") == 2
    ]
    review_pages = [page for page in pages if page.get("review_reasons")]
    manual_review_pages = [page for page in pages if page.get("manual_review_flags")]
    action_counts = Counter(
        page["recommended_action"]
        for page in pages
        if page.get("recommended_action")
    )
    return {
        "ocr_page_count": len(ocr_pages),
        "adaptive_page_count": len(adaptive_pages),
        "high_resolution_selected_count": len(high_resolution_selected),
        "review_page_count": len(review_pages),
        "manual_review_page_count": len(manual_review_pages),
        "review_action_counts": dict(sorted(action_counts.items())),
        "average_ocr_confidence": round(statistics.fmean(confidences), 4) if confidences else None,
        "minimum_ocr_confidence": round(min(confidences), 4) if confidences else None,
    }


def iter_contracts(root: Path, match: str | None):
    for path in sorted(root.rglob("*"), key=lambda item: str(item).lower()):
        if (
            not path.is_file()
            or path.suffix.lower() not in SUPPORTED_EXTENSIONS
            or path.name.lower() in IGNORED_FILENAMES
        ):
            continue
        relative_parts = path.relative_to(root).parts
        if any(part in IGNORED_DIRECTORIES for part in relative_parts[:-1]):
            continue
        if match and match.lower() not in path.name.lower():
            continue
        yield path


def base_payload(
    path: Path,
    root: Path,
    content_hash: str,
    extraction_config: dict,
    extraction_fingerprint: str,
) -> dict:
    relative = path.relative_to(root)
    group = relative.parts[0] if len(relative.parts) > 1 else "ungrouped"
    return {
        "schema_version": SCHEMA_VERSION,
        "source": path.name,
        "group": group,
        "relative_path": relative.as_posix(),
        "extension": path.suffix.lower(),
        "content_sha256": content_hash,
        "size_bytes": path.stat().st_size,
        "extraction_config": extraction_config,
        "extraction_config_fingerprint": extraction_fingerprint,
    }


def read_reusable_output(
    output_path: Path,
    content_hash: str,
    extraction_fingerprint: str,
) -> dict | None:
    if not output_path.exists():
        return None
    try:
        existing = json.loads(output_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    if (
        existing.get("content_sha256") == content_hash
        and existing.get("extraction_config_fingerprint") == extraction_fingerprint
    ):
        return existing
    return None


def reusable_page_prefix(payload: dict, target_pages: int) -> list[dict]:
    pages = payload.get("pages")
    if not isinstance(pages, list):
        return []
    prefix = pages[:target_pages]
    for index, page in enumerate(prefix, start=1):
        if (
            not isinstance(page, dict)
            or page.get("page") != index
            or not isinstance(page.get("text"), str)
        ):
            return []
    return prefix


def extract_pdf(
    path: Path,
    root: Path,
    output_path: Path,
    content_hash: str,
    ocr_mode: str,
    min_text_chars: int,
    max_pages: int,
    force: bool,
    runtime: OCRRuntime,
    settings: dict,
    extraction_config: dict,
    extraction_fingerprint: str,
) -> dict:
    try:
        from PyPDF2 import PdfReader
    except ImportError as exc:
        raise RuntimeError("PyPDF2 is required for PDF extraction") from exc

    reader = PdfReader(str(path))
    total_pages = len(reader.pages)
    target_pages = min(total_pages, max_pages) if max_pages else total_pages
    pages = []
    previous_elapsed = 0.0
    if not force:
        existing = read_reusable_output(output_path, content_hash, extraction_fingerprint)
        if existing is not None:
            pages = reusable_page_prefix(existing, target_pages)
            previous_elapsed = float(existing.get("elapsed_seconds", 0.0))
            if len(pages) == target_pages:
                desired_status = "complete" if target_pages == total_pages else "truncated"
                needs_finalization = (
                    existing.get("status") != desired_status
                    or existing.get("units_completed") != target_pages
                    or len(existing.get("pages", [])) != target_pages
                )
                existing["pages"] = pages
                existing["units_completed"] = len(pages)
                existing["status"] = desired_status
                existing["truncated"] = target_pages != total_pages
                existing["quality_summary"] = summarize_quality(pages)
                if needs_finalization:
                    write_json(output_path, existing)
                return existing

    document = None
    started = time.time()
    payload = {
        **base_payload(path, root, content_hash, extraction_config, extraction_fingerprint),
        "page_count": total_pages,
        "units_completed": len(pages),
        "status": "partial",
        "ocr_mode": ocr_mode,
        "ocr_scale": settings["primary_scale"],
        "elapsed_seconds": previous_elapsed,
        "pages": pages,
    }
    try:
        for page_index in range(len(pages), target_pages):
            embedded_text = (reader.pages[page_index].extract_text() or "").strip()
            use_ocr = ocr_mode == "always" or (
                ocr_mode == "auto" and len(embedded_text) < min_text_chars
            )
            if use_ocr:
                document = document or runtime.open_pdf(path)
                ocr_page = build_ocr_page(
                    page_index + 1,
                    f"page {page_index + 1}",
                    lambda scale, index=page_index, pdf=document: runtime.pdf_page(pdf, index, scale),
                    settings,
                )
                select_ocr, selection_reason = should_select_ocr(
                    ocr_page,
                    embedded_text,
                    ocr_mode,
                )
                if select_ocr:
                    page_payload = ocr_page
                    page_payload["source_selection_reason"] = selection_reason
                else:
                    page_payload = embedded_page(
                        page_index + 1,
                        embedded_text,
                        ocr_page,
                        selection_reason,
                    )
            else:
                page_payload = embedded_page(page_index + 1, embedded_text)

            pages.append(page_payload)
            payload["units_completed"] = len(pages)
            payload["elapsed_seconds"] = round(previous_elapsed + time.time() - started, 1)
            payload["quality_summary"] = summarize_quality(pages)
            write_json(output_path, payload)
            suffix = " adaptive" if len(page_payload.get("ocr_attempts", [])) > 1 else ""
            print(
                f"  page {page_index + 1}/{target_pages}: "
                f"{page_payload['extraction_mode']}{suffix}",
                flush=True,
            )
    finally:
        if document is not None:
            document.close()

    payload["status"] = "complete" if target_pages == total_pages else "truncated"
    payload["truncated"] = target_pages != total_pages
    payload["quality_summary"] = summarize_quality(pages)
    write_json(output_path, payload)
    return payload


def extract_docx(
    path: Path,
    root: Path,
    content_hash: str,
    extraction_config: dict,
    extraction_fingerprint: str,
) -> dict:
    with zipfile.ZipFile(path) as archive:
        xml = archive.read("word/document.xml")
    tree = ElementTree.fromstring(xml)
    namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    paragraphs = []
    for paragraph in tree.findall(".//w:p", namespace):
        text = "".join(node.text or "" for node in paragraph.findall(".//w:t", namespace))
        if text.strip():
            paragraphs.append(text.strip())
    text = "\n".join(paragraphs)
    signals = detect_material_signals(text)
    reasons = [f"material:{signal}" for signal in signals]
    manual_flags = detect_manual_review_flags(text, text.splitlines())
    reasons.extend(f"manual:{flag}" for flag in manual_flags)
    action = recommended_review_action(reasons, manual_flags, 0)
    return {
        **base_payload(path, root, content_hash, extraction_config, extraction_fingerprint),
        "page_count": None,
        "units_completed": 1,
        "status": "complete",
        "quality_summary": summarize_quality([]),
        "pages": [
            {
                "page": None,
                "locator": "document body",
                "extraction_mode": "docx-xml",
                "ocr_confidence": None,
                "material_signals": signals,
                "manual_review_flags": manual_flags,
                "review_reasons": reasons,
                "review_priority": review_priority(reasons),
                "recommended_action": action,
                "requires_original_review": bool(reasons),
                "text": text,
            }
        ],
    }


def extract_text_file(
    path: Path,
    root: Path,
    content_hash: str,
    extraction_config: dict,
    extraction_fingerprint: str,
) -> dict:
    text = path.read_text(encoding="utf-8-sig", errors="replace")
    signals = detect_material_signals(text)
    reasons = [f"material:{signal}" for signal in signals]
    manual_flags = detect_manual_review_flags(text, text.splitlines())
    reasons.extend(f"manual:{flag}" for flag in manual_flags)
    action = recommended_review_action(reasons, manual_flags, 0)
    return {
        **base_payload(path, root, content_hash, extraction_config, extraction_fingerprint),
        "page_count": None,
        "units_completed": 1,
        "status": "complete",
        "quality_summary": summarize_quality([]),
        "pages": [
            {
                "page": None,
                "locator": "document body",
                "extraction_mode": "direct-text",
                "ocr_confidence": None,
                "material_signals": signals,
                "manual_review_flags": manual_flags,
                "review_reasons": reasons,
                "review_priority": review_priority(reasons),
                "recommended_action": action,
                "requires_original_review": bool(reasons),
                "text": text,
            }
        ],
    }


def extract_image(
    path: Path,
    root: Path,
    output_path: Path,
    content_hash: str,
    ocr_mode: str,
    max_pages: int,
    force: bool,
    runtime: OCRRuntime,
    settings: dict,
    extraction_config: dict,
    extraction_fingerprint: str,
) -> dict:
    if ocr_mode == "never":
        raise RuntimeError("image inputs require OCR; remove --ocr never or exclude the image")
    total_pages = runtime.image_page_count(path)
    target_pages = min(total_pages, max_pages) if max_pages else total_pages
    pages = []
    previous_elapsed = 0.0
    if not force:
        existing = read_reusable_output(output_path, content_hash, extraction_fingerprint)
        if existing is not None:
            pages = reusable_page_prefix(existing, target_pages)
            previous_elapsed = float(existing.get("elapsed_seconds", 0.0))
            if len(pages) == target_pages:
                desired_status = "complete" if target_pages == total_pages else "truncated"
                needs_finalization = (
                    existing.get("status") != desired_status
                    or existing.get("units_completed") != target_pages
                    or len(existing.get("pages", [])) != target_pages
                )
                existing["pages"] = pages
                existing["units_completed"] = len(pages)
                existing["status"] = desired_status
                existing["truncated"] = target_pages != total_pages
                existing["quality_summary"] = summarize_quality(pages)
                if needs_finalization:
                    write_json(output_path, existing)
                return existing

    started = time.time()
    payload = {
        **base_payload(path, root, content_hash, extraction_config, extraction_fingerprint),
        "page_count": total_pages,
        "units_completed": len(pages),
        "status": "partial",
        "ocr_mode": ocr_mode,
        "ocr_scale": settings["primary_scale"],
        "elapsed_seconds": previous_elapsed,
        "pages": pages,
    }
    for frame_index in range(len(pages), target_pages):
        page = build_ocr_page(
            frame_index + 1,
            f"image frame {frame_index + 1}",
            lambda scale, index=frame_index: runtime.image_frame(path, index, scale),
            settings,
        )
        pages.append(page)
        payload["units_completed"] = len(pages)
        payload["elapsed_seconds"] = round(previous_elapsed + time.time() - started, 1)
        payload["quality_summary"] = summarize_quality(pages)
        write_json(output_path, payload)

    payload["status"] = "complete" if target_pages == total_pages else "truncated"
    payload["truncated"] = target_pages != total_pages
    payload["quality_summary"] = summarize_quality(pages)
    write_json(output_path, payload)
    return payload


def output_name(path: Path, root: Path, extraction_fingerprint: str) -> str:
    relative = path.relative_to(root)
    group = relative.parts[0] if len(relative.parts) > 1 else "ungrouped"
    source_id = hashlib.sha256(relative.as_posix().encode("utf-8")).hexdigest()[:10]
    return (
        f"{safe_stem(group)}__{safe_stem(path.stem)}__"
        f"{source_id}__{extraction_fingerprint[:8]}.json"
    )


def review_entries(payload: dict) -> list[dict]:
    entries = []
    for page in payload.get("pages", []):
        reasons = page.get("review_reasons") or []
        if not reasons:
            continue
        entries.append(
            {
                "source": payload.get("relative_path"),
                "page": page.get("page"),
                "locator": page.get("locator"),
                "priority": page.get("review_priority"),
                "reasons": reasons,
                "material_signals": page.get("material_signals", []),
                "manual_review_flags": page.get("manual_review_flags", []),
                "recommended_action": page.get("recommended_action"),
                "ocr_confidence": page.get("ocr_confidence"),
                "ocr_scale": page.get("ocr_scale"),
            }
        )
    return entries


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path, help="Directory containing contract files")
    parser.add_argument("--output-dir", type=Path)
    parser.add_argument("--ocr", choices=("auto", "always", "never"), default="auto")
    parser.add_argument(
        "--quality-profile",
        choices=tuple(QUALITY_PROFILES),
        default="hybrid",
        help="OCR quality/cost policy (default: hybrid)",
    )
    parser.add_argument("--scale", type=finite_float, help="Override primary render scale")
    parser.add_argument(
        "--adaptive-ocr",
        action=argparse.BooleanOptionalAction,
        default=None,
        help="Enable or disable a conditional second OCR pass",
    )
    parser.add_argument("--second-pass-scale", type=finite_float)
    parser.add_argument("--second-pass-confidence", type=finite_float)
    parser.add_argument("--second-pass-min-chars", type=positive_int)
    parser.add_argument(
        "--second-pass-material",
        action=argparse.BooleanOptionalAction,
        default=None,
        help="Second-pass pages containing material contract signals",
    )
    parser.add_argument("--review-confidence", type=finite_float)
    parser.add_argument("--max-render-megapixels", type=finite_float)
    parser.add_argument("--min-text-chars", type=positive_int, default=40)
    parser.add_argument("--max-pages", type=nonnegative_int, default=0, help="Limit each PDF; 0 means all")
    parser.add_argument("--match", help="Only process filenames containing this text")
    parser.add_argument("--det-model-path", type=Path)
    parser.add_argument("--rec-model-path", type=Path)
    parser.add_argument("--cls-model-path", type=Path)
    parser.add_argument("--force", action="store_true")
    parser.add_argument(
        "--allow-errors",
        action="store_true",
        help="Return success even when individual files fail",
    )
    parser.add_argument("--allow-empty", action="store_true", help="Allow an empty file selection")
    parser.add_argument(
        "--fail-on-error",
        action="store_true",
        help=argparse.SUPPRESS,
    )
    parser.add_argument("--vendor-dir", type=Path, action="append", default=[])
    args = parser.parse_args()

    batch_started = time.perf_counter()
    memory_start = process_memory_snapshot()

    root = args.input.resolve()
    if not root.is_dir():
        parser.error(f"input directory does not exist: {root}")
    output_dir = (args.output_dir or root / ".contract-analysis" / "extracted_text").resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    add_dependency_paths(root, args.vendor_dir)
    try:
        settings = resolve_quality_settings(args)
        model_paths = resolve_model_paths(args)
    except ValueError as exc:
        parser.error(str(exc))
    models = model_manifest(model_paths)
    extraction_config = build_extraction_config(args, settings, models)
    extraction_fingerprint = config_fingerprint(extraction_config)
    runtime = OCRRuntime(model_paths, settings["max_render_megapixels"])

    records = []
    errors = []
    queue = []
    contracts = list(iter_contracts(root, args.match))
    if not contracts and not args.allow_empty:
        parser.error("no supported contract files matched the input and filters")
    for index, path in enumerate(contracts, start=1):
        print(f"[{index}/{len(contracts)}] {path.relative_to(root)}", flush=True)
        try:
            content_hash = sha256_file(path)
            output_path = output_dir / output_name(path, root, extraction_fingerprint)
            suffix = path.suffix.lower()
            if suffix == ".pdf":
                payload = extract_pdf(
                    path,
                    root,
                    output_path,
                    content_hash,
                    args.ocr,
                    max(1, args.min_text_chars),
                    max(0, args.max_pages),
                    args.force,
                    runtime,
                    settings,
                    extraction_config,
                    extraction_fingerprint,
                )
            elif suffix == ".docx":
                payload = extract_docx(
                    path,
                    root,
                    content_hash,
                    extraction_config,
                    extraction_fingerprint,
                )
                write_json(output_path, payload)
            elif suffix in {".txt", ".md"}:
                payload = extract_text_file(
                    path,
                    root,
                    content_hash,
                    extraction_config,
                    extraction_fingerprint,
                )
                write_json(output_path, payload)
            elif suffix == ".doc":
                raise RuntimeError("legacy .doc must be converted to .docx or PDF before extraction")
            else:
                payload = extract_image(
                    path,
                    root,
                    output_path,
                    content_hash,
                    args.ocr,
                    args.max_pages,
                    args.force,
                    runtime,
                    settings,
                    extraction_config,
                    extraction_fingerprint,
                )
            file_queue = review_entries(payload)
            queue.extend(file_queue)
            records.append(
                {
                    "source": payload["relative_path"],
                    "output": output_path.name,
                    "status": payload["status"],
                    "units_completed": payload["units_completed"],
                    "page_count": payload.get("page_count"),
                    "quality_summary": payload.get("quality_summary"),
                    "review_queue_count": len(file_queue),
                }
            )
        except Exception as exc:  # noqa: BLE001
            errors.append({"source": path.relative_to(root).as_posix(), "error": str(exc)})
            print(f"  ERROR: {exc}", file=sys.stderr, flush=True)

    action_counts = Counter(
        entry["recommended_action"]
        for entry in queue
        if entry.get("recommended_action")
    )
    review_queue = {
        "schema_version": SCHEMA_VERSION,
        "extraction_config_fingerprint": extraction_fingerprint,
        "entry_count": len(queue),
        "high_priority_count": sum(entry["priority"] == "high" for entry in queue),
        "action_counts": dict(sorted(action_counts.items())),
        "entries": queue,
    }
    write_json(output_dir / "review-queue.json", review_queue)
    memory_end = process_memory_snapshot()
    resource_usage = {
        "elapsed_seconds": round(time.perf_counter() - batch_started, 3),
        "memory_start": memory_start,
        "memory_end": memory_end,
    }
    summary = {
        "schema_version": SCHEMA_VERSION,
        "input_root": str(root),
        "output_dir": str(output_dir),
        "extraction_config": extraction_config,
        "extraction_config_fingerprint": extraction_fingerprint,
        "file_count": len(contracts),
        "success_count": len(records),
        "error_count": len(errors),
        "review_queue_count": len(queue),
        "review_action_counts": dict(sorted(action_counts.items())),
        "resource_usage": resource_usage,
        "records": records,
        "errors": errors,
    }
    write_json(output_dir / "extraction-summary.json", summary)
    print(
        f"Completed: {len(records)} succeeded, {len(errors)} failed, "
        f"{len(queue)} review pages; elapsed={resource_usage['elapsed_seconds']:.1f}s; "
        f"output={output_dir}",
        flush=True,
    )
    return 1 if errors and (args.fail_on_error or not args.allow_errors) else 0


if __name__ == "__main__":
    raise SystemExit(main())
