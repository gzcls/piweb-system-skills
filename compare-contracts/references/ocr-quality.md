# OCR Quality and Model Evaluation

Use this reference when scanned contracts need OCR, when extraction quality is questioned, or when changing OCR models or parameters.

- [Quality Profiles](#quality-profiles)
- [Adaptive Output](#adaptive-output)
- [Calibration Evidence](#calibration-evidence)
- [Model Provenance and Cache](#model-provenance-and-cache)
- [Ground-Truth Evaluation](#ground-truth-evaluation)

## Quality Profiles

| Profile | First pass | Conditional pass | Material-page rerun | Pixel cap | Use case |
|---|---:|---:|---|---:|---|
| `fast` | 1.25x | disabled | no | 8 MP | inventory, triage, quick search |
| `hybrid` | 1.25x | 2.5x below 0.80 confidence or 80 characters | no | 12 MP | default contract comparison |
| `balanced` | 1.5x | 2.5x below 0.88 confidence or 80 characters | no | 12 MP | measured fallback for difficult files |
| `high` | 1.5x | 3.0x below 0.92 confidence, 120 characters, or material signals | yes | 16 MP | decision-critical extraction |

The pixel cap is enforced per rendered page. If the source cannot fit under the cap even at 0.5x, extraction fails instead of risking uncontrolled memory use.

Use `hybrid` for the full corpus. It combines a fast first pass with page-level enhancement only for clear quality failures. Use `balanced` or `high` for selected files only when ground truth or the review queue justifies the extra work. A higher render scale is only a candidate: the extractor preserves both attempts and records how it selected one.

Do not rerun a page merely because it contains money, dates, termination, liability, data, or IP language. Material signals require verification when cited; they do not prove that higher-resolution OCR will help.

## Adaptive Output

Each OCR page records:

- selected text, average confidence, effective scale, and quality score;
- all OCR attempts with text, line confidence, normalized boxes, render size, and model time;
- similarity between attempts and why a second pass ran;
- material signals, manual-review flags, review reasons, priority, and a recommended action.

When all attempts remain below 0.82 confidence, small confidence differences are poorly calibrated. The selector therefore compares material-value count, numeric-token count, text coverage, and line coverage. Low-confidence and materially disagreeing attempts still require original-page verification.

The extractor detects three manual gates from text and OCR-line structure:

- `identity-document`: business licenses, incorporation certificates, and registration records;
- `signature-or-stamp`: signing blocks, representatives, dates, signatures, and stamps;
- `complex-table`: dense numeric KPI, GMV/GSV/NET, commission, target, or score tables.

These are routing signals, not proof that every flagged page contains handwriting or a bad table. They always produce `manual-verify-original` because ordinary OCR confidence cannot validate visual identity, handwriting, stamps, merged cells, or table reading order.

For `complex-table` pages, `layout_table` groups OCR boxes into left-to-right row candidates and retains each cell's box and confidence. Use it to recover scan order and locate values; verify merged cells, headers, units, and cited values against the original page.

`extraction-summary.json` records total wall-clock time and process memory snapshots. On Windows, `memory_end.peak_rss_bytes` is the extractor process's peak working set and includes Python, OCR models, ONNX Runtime, rendered pages, and other loaded libraries. It is useful for comparing like-for-like runs, but it is not a per-page allocation measurement.

`review-queue.json` contains only file/page locators and routing metadata, not extracted contract text. Process its actions as follows:

1. Resolve `enhanced-ocr-then-verify` with a measured stronger file run and page inspection.
2. Resolve `manual-verify-original` directly against the original page.
3. Resolve `verify-material-fields` for facts actually used in analysis or the report.

### Review Ledger

Create a local `review-ledger.json` for actions that require OCR enhancement or original-page inspection:

```json
{
  "schema_version": 1,
  "extraction_config_fingerprint": "<fingerprint from review-queue.json>",
  "entries": [
    {
      "source": "Supplier A/agreement.pdf",
      "page": 12,
      "action": "manual-verify-original",
      "status": "resolved",
      "original_checked": true,
      "notes": "Verified the amount and table header against the scan."
    }
  ]
}
```

Use `pending`, `resolved`, or `unclear` status. Both `resolved` and `unclear` must confirm `original_checked`; `unclear` also requires notes. The final extraction gate reconciles all `enhanced-ocr-then-verify` and `manual-verify-original` queue entries with this ledger. Keep cited material-field verification in `source-map.json`, because material-only queue entries need review only when used.

Confidence is a model score, not an accuracy percentage. Always inspect the original page for decision-critical values even when confidence is high.

## Calibration Evidence

On one 6-contract, 140-page scanned Chinese corpus, the former full-corpus `balanced` default took 56 minutes 21 seconds versus 31 minutes 21 seconds for the earlier lower-scale run. On the same 18-page risk-stratified truth set, pooled CER changed only from 13.8552% to 13.8478%. The larger run reduced the worst-page CER but did not materially improve aggregate CER or strict key-field matches.

Replaying the new 0.80 quality threshold against that run's first attempts selected 32 pages for a second pass instead of 66. Their recorded OCR engine time indicates an estimated 632-second reduction in second-pass work before accounting for the faster 1.25x first pass. This is an estimate, not a measured hybrid wall-clock benchmark.

Treat these figures as the reason for the default policy, not as a universal guarantee. Recalibrate thresholds on representative ground truth when the language, scan source, model bundle, or document layout changes.

## Model Provenance and Cache

The extraction configuration stores package versions, engine configuration hash, and SHA-256 for the default or custom ONNX models. The full configuration is fingerprinted. A model, scale, profile, OCR mode, or page-limit change creates a distinct output identity instead of reusing stale text.

Custom model flags are:

```powershell
--det-model-path <compatible-det.onnx>
--rec-model-path <compatible-rec.onnx>
--cls-model-path <compatible-cls.onnx>
```

RapidOCR 1.2.3 only accepts models compatible with its existing preprocessing and decoder:

- detection must expose a PaddleOCR DB-style probability-map output;
- recognition must expose CTC output and a compatible embedded character table;
- classification must match the existing `0/180` orientation labels.

A file path alone does not make a newer model compatible. Test every model bundle on ground truth before using it for material findings.

## Ground-Truth Evaluation

Create a local UTF-8 JSON file with human-corrected pages:

```json
{
  "records": [
    {
      "source": "Supplier A/agreement.pdf",
      "page": 12,
      "text": "Human-corrected page text"
    }
  ]
}
```

Run:

```powershell
python <skill-dir>/scripts/evaluate_ocr.py <extracted-dir> <ground-truth.json> --output <evaluation.json>
```

The default comparison ignores whitespace and reports page-level plus aggregate character error rate (`CER`). Use `--normalization exact` when line breaks and spacing matter. Use `--max-cer <value>` as a CI or acceptance gate only after the project defines an acceptable threshold.

For a model or parameter A/B test:

1. Use the same representative ground-truth pages for both runs.
2. Include body text, small print, numbers, tables, stamps, and poor scans.
3. Compare aggregate CER, worst-page CER, missing text, runtime, and memory.
4. Require exact agreement for sampled material numbers, dates, percentages, and caps.
5. Keep the current model if a candidate only raises confidence without lowering CER.
