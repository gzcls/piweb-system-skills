from __future__ import annotations

import hashlib
import io
import json
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from unittest.mock import patch

SCRIPT_DIR = Path(__file__).resolve().parents[1]
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from evaluate_ocr import levenshtein_distance, normalize_text  # noqa: E402
from evaluate_ocr import main as evaluate_main
from extract_contracts import (  # noqa: E402
    QUALITY_PROFILES,
    add_dependency_paths,
    build_layout_table,
    build_ocr_page,
    config_fingerprint,
    detect_manual_review_flags,
    detect_material_signals,
    embedded_page,
    extract_image,
    output_name,
    process_memory_snapshot,
    recommended_review_action,
    review_entries,
    second_pass_reasons,
    select_best_attempt,
    should_select_ocr,
    text_similarity,
)
from validate_extraction import (  # noqa: E402
    load_batch,
    validate_file,
    validate_review_ledger,
    validate_review_queue,
)
from validate_extraction import (
    main as validate_main,
)


class OCRQualityTests(unittest.TestCase):
    @staticmethod
    def fake_attempt(text: str) -> dict:
        return {
            "requested_scale": 1.25,
            "effective_scale": 1.25,
            "render_width": 100,
            "render_height": 100,
            "render_megapixels": 0.01,
            "confidence": 0.95,
            "text_chars": len(text),
            "line_count": 1,
            "quality_score": 0.95,
            "text_sha256": "0" * 64,
            "engine_seconds": 0.01,
            "text": text,
            "lines": [
                {
                    "text": text,
                    "confidence": 0.95,
                    "box": [[0.0, 0.0], [1.0, 0.0], [1.0, 1.0], [0.0, 1.0]],
                }
            ],
        }

    def test_material_signals_cover_contract_risk_families(self):
        signals = detect_material_signals("付款金额100万元，终止时承担赔偿责任并删除个人信息。")
        self.assertEqual(
            {"money", "termination", "liability", "data-ip"},
            set(signals),
        )

    def test_balanced_profile_reruns_low_confidence(self):
        attempt = {"confidence": 0.79, "text_chars": 500}
        reasons = second_pass_reasons(attempt, QUALITY_PROFILES["balanced"], [])
        self.assertIn("low-confidence", reasons)

    def test_hybrid_profile_only_reruns_clear_quality_risks(self):
        profile = QUALITY_PROFILES["hybrid"]
        self.assertEqual([], second_pass_reasons({"confidence": 0.81, "text_chars": 500}, profile, []))
        self.assertEqual(
            ["low-confidence"],
            second_pass_reasons({"confidence": 0.79, "text_chars": 500}, profile, []),
        )
        self.assertEqual(
            ["low-text"],
            second_pass_reasons({"confidence": 0.95, "text_chars": 79}, profile, []),
        )

    def test_high_profile_reruns_material_pages(self):
        attempt = {"confidence": 0.99, "text_chars": 500}
        reasons = second_pass_reasons(attempt, QUALITY_PROFILES["high"], ["money"])
        self.assertEqual(["material-page"], reasons)

    def test_attempt_selection_prefers_quality_then_confidence(self):
        attempts = [
            {"quality_score": 0.80, "confidence": 0.90, "text_chars": 500},
            {"quality_score": 0.85, "confidence": 0.88, "text_chars": 700},
        ]
        self.assertEqual(1, select_best_attempt(attempts))

    def test_low_confidence_selection_prefers_structured_coverage(self):
        attempts = [
            {
                "quality_score": 0.82,
                "confidence": 0.81,
                "text_chars": 500,
                "line_count": 20,
                "material_value_count": 5,
                "numeric_token_count": 10,
            },
            {
                "quality_score": 0.78,
                "confidence": 0.79,
                "text_chars": 540,
                "line_count": 24,
                "material_value_count": 6,
                "numeric_token_count": 12,
            },
        ]
        self.assertEqual(1, select_best_attempt(attempts))

    def test_manual_review_flags_cover_hard_contract_pages(self):
        certificate = detect_manual_review_flags(
            "营业执照 统一社会信用代码 123 登记机关",
            ["营业执照", "统一社会信用代码 123", "登记机关"],
        )
        self.assertIn("identity-document", certificate)

        signing = detect_manual_review_flags(
            "甲方：A 乙方：B 代表人：张三 2026年8月20日",
            ["甲方：A", "乙方：B", "代表人：张三", "2026年8月20日"],
        )
        self.assertIn("signature-or-stamp", signing)

        table_lines = [f"KPI {index} 20%" for index in range(20)]
        table = detect_manual_review_flags("\n".join(table_lines), table_lines)
        self.assertIn("complex-table", table)

    def test_layout_table_clusters_rows_and_orders_cells(self):
        lines = [
            {"text": "B", "confidence": 0.9, "box": [[20, 0], [30, 0], [30, 10], [20, 10]]},
            {"text": "A", "confidence": 0.9, "box": [[0, 0], [10, 0], [10, 10], [0, 10]]},
            {"text": "D", "confidence": 0.9, "box": [[20, 30], [30, 30], [30, 40], [20, 40]]},
            {"text": "C", "confidence": 0.9, "box": [[0, 30], [10, 30], [10, 40], [0, 40]]},
        ]
        table = build_layout_table(lines)
        self.assertIsNotNone(table)
        self.assertEqual(2, table["row_count"])
        self.assertEqual(["A | B", "C | D"], [row["text"] for row in table["rows"]])

    def test_manual_flags_force_original_page_verification(self):
        self.assertEqual(
            "manual-verify-original",
            recommended_review_action(
                ["material:money", "manual:complex-table"],
                ["complex-table"],
                2,
            ),
        )
        self.assertEqual(
            "enhanced-ocr-then-verify",
            recommended_review_action(["low-confidence"], [], 1),
        )
        self.assertEqual(
            "verify-material-fields",
            recommended_review_action(["material:money"], [], 1),
        )

    def test_build_ocr_page_records_low_confidence_selection_and_action(self):
        attempts = [
            {
                **self.fake_attempt("金额100万元"),
                "confidence": 0.79,
                "quality_score": 0.82,
                "material_value_count": 1,
                "numeric_token_count": 1,
            },
            {
                **self.fake_attempt("付款金额100万元，期限30日"),
                "confidence": 0.78,
                "quality_score": 0.80,
                "material_value_count": 2,
                "numeric_token_count": 2,
            },
        ]
        attempts[0]["text_chars"] = len(attempts[0]["text"])
        attempts[1]["text_chars"] = len(attempts[1]["text"])
        iterator = iter(attempts)
        page = build_ocr_page(1, "page 1", lambda _scale: next(iterator), QUALITY_PROFILES["hybrid"])
        self.assertEqual(2, page["selected_attempt"])
        self.assertEqual("low-confidence-coverage-second-pass", page["selection_reason"])
        self.assertEqual("manual-verify-original", page["recommended_action"])

    def test_text_similarity_ignores_whitespace(self):
        self.assertEqual(1.0, text_similarity("合同\n金额", "合 同 金 额"))

    def test_empty_ocr_never_replaces_embedded_text(self):
        selected, reason = should_select_ocr(
            {"text": "", "ocr_confidence": None},
            "有效合同文本",
            "always",
        )
        self.assertFalse(selected)
        self.assertEqual("ocr-empty-fallback", reason)

    def test_empty_ocr_fallback_is_high_priority_review(self):
        page = embedded_page(
            1,
            "",
            {"text": "", "ocr_confidence": None},
            "ocr-empty-fallback",
        )
        self.assertEqual("high", page["review_priority"])
        self.assertIn("empty-text", page["review_reasons"])
        self.assertIn("ocr-empty", page["review_reasons"])

    def test_config_fingerprint_changes_with_scale(self):
        first = config_fingerprint({"primary_scale": 1.5})
        second = config_fingerprint({"primary_scale": 2.5})
        self.assertNotEqual(first, second)

    def test_process_memory_snapshot_has_stable_shape(self):
        snapshot = process_memory_snapshot()
        self.assertEqual(
            {"current_rss_bytes", "peak_rss_bytes", "source"},
            set(snapshot),
        )
        self.assertIsInstance(snapshot["source"], str)

    def test_dependency_paths_discover_legacy_project_tools(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            ocr = root / ".tools" / "ocr"
            pymupdf = root / ".tools" / "pymupdf"
            ocr.mkdir(parents=True)
            pymupdf.mkdir(parents=True)
            original_path = list(sys.path)
            try:
                add_dependency_paths(root, [])
                self.assertIn(str(ocr.resolve()), sys.path)
                self.assertIn(str(pymupdf.resolve()), sys.path)
            finally:
                sys.path[:] = original_path

    def test_output_name_uses_relative_path_identity(self):
        root = Path("C:/contracts")
        first = output_name(root / "group" / "a" / "same.pdf", root, "a" * 64)
        second = output_name(root / "group" / "b" / "same.pdf", root, "a" * 64)
        self.assertNotEqual(first, second)

    def test_levenshtein_and_normalization(self):
        self.assertEqual(1, levenshtein_distance("合同金额", "合同金领"))
        self.assertEqual("合同金额", normalize_text("合同\n 金额", "whitespace"))

    def test_cer_evaluator_end_to_end_and_threshold(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            extraction = root / "extracted"
            extraction.mkdir()
            artifact = {
                "relative_path": "Supplier A/agreement.pdf",
                "source": str(root / "Supplier A" / "agreement.pdf"),
                "pages": [{"page": 1, "text": "合同金额 100 元"}],
            }
            (extraction / "contract.json").write_text(
                json.dumps(artifact, ensure_ascii=False), encoding="utf-8"
            )
            (extraction / "extraction-summary.json").write_text(
                json.dumps(
                    {
                        "error_count": 0,
                        "errors": [],
                        "records": [{"output": "contract.json"}],
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )
            truth = root / "truth.json"
            truth.write_text(
                json.dumps(
                    {
                        "records": [
                            {
                                "source": "Supplier A/agreement.pdf",
                                "page": 1,
                                "text": "合同金额100元",
                            }
                        ]
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )
            report = root / "evaluation.json"
            argv = [
                "evaluate_ocr.py",
                str(extraction),
                str(truth),
                "--max-cer",
                "0",
                "--output",
                str(report),
            ]
            with patch.object(sys, "argv", argv), redirect_stdout(io.StringIO()):
                self.assertEqual(0, evaluate_main())
            result = json.loads(report.read_text(encoding="utf-8"))
            self.assertEqual(0.0, result["aggregate_cer"])
            self.assertTrue(result["results"][0]["exact_match"])

    def test_batch_summary_reports_extraction_failures(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            summary = {
                "file_count": 1,
                "success_count": 0,
                "error_count": 1,
                "records": [],
                "errors": [{"source": "legacy.doc", "error": "conversion required"}],
            }
            (root / "extraction-summary.json").write_text(json.dumps(summary), encoding="utf-8")
            files, _, loaded, errors = load_batch(root)
            self.assertEqual([], files)
            self.assertIsNotNone(loaded)
            self.assertTrue(any("failed file" in message for message in errors))

    def test_validator_rejects_empty_summary_without_opt_in(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "extraction-summary.json").write_text(
                json.dumps(
                    {
                        "file_count": 0,
                        "success_count": 0,
                        "error_count": 0,
                        "records": [],
                        "errors": [],
                    }
                ),
                encoding="utf-8",
            )
            with patch.object(sys, "argv", ["validate_extraction.py", str(root)]):
                with self.assertRaises(SystemExit) as failure:
                    validate_main()
            self.assertEqual(2, failure.exception.code)

            with patch.object(
                sys,
                "argv",
                ["validate_extraction.py", str(root), "--allow-empty"],
            ), redirect_stdout(io.StringIO()):
                self.assertEqual(0, validate_main())

    def test_validator_handles_non_object_json(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "bad.json"
            path.write_text("[]", encoding="utf-8")
            errors, _, _, payload = validate_file(path, False, 0.8, False)
            self.assertIsNone(payload)
            self.assertIn("top-level extraction must be an object", errors)

    def test_validator_rejects_extraction_with_no_text(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "blank.json"
            path.write_text(
                json.dumps(
                    {
                        "source": "blank.pdf",
                        "relative_path": "blank.pdf",
                        "extension": ".pdf",
                        "content_sha256": "0" * 64,
                        "status": "complete",
                        "page_count": 1,
                        "units_completed": 1,
                        "pages": [
                            {
                                "page": 1,
                                "locator": "page 1",
                                "extraction_mode": "embedded-text",
                                "ocr_confidence": None,
                                "review_reasons": ["empty-text"],
                                "text": "",
                            }
                        ],
                    }
                ),
                encoding="utf-8",
            )
            errors, _, _, _ = validate_file(path, False, 0.8, False)
            self.assertIn("extraction contains no text", errors)

    def test_missing_review_queue_is_a_batch_error(self):
        summary = {
            "schema_version": 2,
            "extraction_config_fingerprint": "a" * 64,
            "review_queue_count": 0,
        }
        with tempfile.TemporaryDirectory() as directory:
            errors = validate_review_queue(Path(directory), summary, [])
        self.assertIn("missing review-queue.json for schema version 2 batch", errors)

    def test_v4_review_actions_validate_end_to_end(self):
        text = "付款金额100万元，" * 12
        attempt = {
            **self.fake_attempt(text),
            "numeric_token_count": 12,
            "material_value_count": 12,
            "short_line_ratio": 1.0,
        }
        attempt["text_sha256"] = hashlib.sha256(text.encode("utf-8")).hexdigest()
        page = build_ocr_page(
            1,
            "page 1",
            lambda _scale: attempt,
            QUALITY_PROFILES["hybrid"],
        )
        config = {"extractor_version": 4}
        fingerprint = config_fingerprint(config)
        payload = {
            "source": "contract.pdf",
            "relative_path": "contract.pdf",
            "extension": ".pdf",
            "content_sha256": "0" * 64,
            "status": "complete",
            "page_count": 1,
            "units_completed": 1,
            "extraction_config": config,
            "extraction_config_fingerprint": fingerprint,
            "pages": [page],
        }

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            artifact = root / "contract.json"
            artifact.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
            errors, _, _, loaded = validate_file(artifact, True, 0.8, True)
            self.assertEqual([], errors)

            entries = review_entries(payload)
            action_counts = {"verify-material-fields": 1}
            summary = {
                "schema_version": 2,
                "extraction_config_fingerprint": fingerprint,
                "review_queue_count": 1,
                "review_action_counts": action_counts,
            }
            (root / "review-queue.json").write_text(
                json.dumps(
                    {
                        "schema_version": 2,
                        "extraction_config_fingerprint": fingerprint,
                        "entry_count": 1,
                        "high_priority_count": 1,
                        "action_counts": action_counts,
                        "entries": entries,
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )
            self.assertEqual([], validate_review_queue(root, summary, [loaded]))

    def test_review_ledger_requires_original_check_for_manual_actions(self):
        fingerprint = "a" * 64
        queue = {
            "extraction_config_fingerprint": fingerprint,
            "entries": [
                {
                    "source": "contract.pdf",
                    "page": 3,
                    "recommended_action": "manual-verify-original",
                },
                {
                    "source": "contract.pdf",
                    "page": 1,
                    "recommended_action": "verify-material-fields",
                },
            ],
        }
        ledger = {
            "schema_version": 1,
            "extraction_config_fingerprint": fingerprint,
            "entries": [
                {
                    "source": "contract.pdf",
                    "page": 3,
                    "action": "manual-verify-original",
                    "status": "resolved",
                    "original_checked": True,
                }
            ],
        }
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "review-queue.json").write_text(json.dumps(queue), encoding="utf-8")
            ledger_path = root / "review-ledger.json"
            ledger_path.write_text(json.dumps(ledger), encoding="utf-8")
            errors, warnings, metrics = validate_review_ledger(root, ledger_path, True)
            self.assertEqual([], errors)
            self.assertEqual([], warnings)
            self.assertEqual(1, metrics["required_action_count"])
            self.assertEqual(1, metrics["resolved_count"])

            ledger["entries"][0]["original_checked"] = False
            ledger_path.write_text(json.dumps(ledger), encoding="utf-8")
            errors, _, _ = validate_review_ledger(root, ledger_path, True)
            self.assertTrue(any("original_checked" in error for error in errors))

            ledger["entries"] = []
            ledger_path.write_text(json.dumps(ledger), encoding="utf-8")
            errors, _, metrics = validate_review_ledger(root, ledger_path, True)
            self.assertTrue(any("missing 1 required action" in error for error in errors))
            self.assertEqual(1, metrics["missing_count"])

    def test_image_ocr_never_fails_before_runtime_use(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            with self.assertRaisesRegex(RuntimeError, "image inputs require OCR"):
                extract_image(
                    root / "scan.png",
                    root,
                    root / "scan.json",
                    "1" * 64,
                    "never",
                    0,
                    False,
                    None,
                    QUALITY_PROFILES["fast"],
                    {},
                    "2" * 64,
                )

    def test_multiframe_image_and_partial_terminal_recovery(self):
        class FakeRuntime:
            def __init__(self):
                self.frames = []

            def image_page_count(self, _path):
                return 2

            def image_frame(self, _path, frame_index, _scale):
                self.frames.append(frame_index)
                return OCRQualityTests.fake_attempt(f"frame {frame_index + 1}")

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "scan.tiff"
            source.write_bytes(b"test fixture")
            output = root / "scan.json"
            runtime = FakeRuntime()
            payload = extract_image(
                source,
                root,
                output,
                "3" * 64,
                "auto",
                0,
                True,
                runtime,
                dict(QUALITY_PROFILES["fast"]),
                {},
                "4" * 64,
            )
            self.assertEqual([0, 1], runtime.frames)
            self.assertEqual(2, payload["page_count"])
            self.assertEqual([1, 2], [page["page"] for page in payload["pages"]])

            payload["status"] = "partial"
            output.write_text(json.dumps(payload), encoding="utf-8")
            recovery_runtime = FakeRuntime()
            recovered = extract_image(
                source,
                root,
                output,
                "3" * 64,
                "auto",
                0,
                False,
                recovery_runtime,
                dict(QUALITY_PROFILES["fast"]),
                {},
                "4" * 64,
            )
            self.assertEqual([], recovery_runtime.frames)
            self.assertEqual("complete", recovered["status"])


if __name__ == "__main__":
    unittest.main()
