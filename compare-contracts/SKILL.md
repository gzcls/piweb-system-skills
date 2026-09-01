---
name: compare-contracts
description: Extract, normalize, verify, and compare multiple contracts from PDF, scanned PDF, DOCX, image, TXT, or Markdown sources, then produce a source-cited HTML/CSS report. Use when Codex is asked to compare contract terms, suppliers, customer/vendor positions, old and new versions, fees, KPIs, payment, termination, liabilities, data/IP duties, or risks, or to assess and improve OCR quality for contract folders.
---

# Compare Contracts

## Overview

Create evidence-backed contract comparisons. Use deterministic scripts for inventory, extraction, and validation; use model judgment for term normalization, risk analysis, and recommendations.

Do not treat this workflow as legal advice. State the review perspective and identify terms that need qualified legal review.

## Workflow

### 1. Establish scope

1. Inspect the input directory and any user instructions before processing files.
2. Determine the comparison groups from folder structure, filenames, or the user's description.
3. Record the review perspective: buyer/customer, supplier/vendor, employer, employee, landlord, tenant, or neutral. Default to neutral when it cannot be inferred safely.
4. Keep source files unchanged. Put generated intermediates under `<input-dir>/.contract-analysis/` unless the user specifies another output location.
5. Keep contract text and extracted data local unless the user explicitly authorizes external services.

### 2. Inventory the corpus

Run:

```powershell
python <skill-dir>/scripts/inventory_contracts.py <input-dir> --output <input-dir>/.contract-analysis/inventory.json --hash
```

Review file counts, groups, page counts, hashes, duplicates, and extraction hints. The script classifies PDFs as embedded-text, scanned, mixed, or unknown from a page sample.

Legacy `.doc` files require conversion to `.docx` or PDF with a locally available office tool before extraction. Do not rename the extension without converting the file.

### 3. Extract text with provenance

Run the hybrid profile for the full corpus:

```powershell
python <skill-dir>/scripts/extract_contracts.py <input-dir> --ocr auto --quality-profile hybrid
```

The extractor tries embedded PDF text first and falls back to local OCR page by page. `hybrid` uses a 1.25x first pass for every scanned page, then runs 2.5x only below 0.80 confidence or 80 characters. It preserves both attempts and writes one JSON file per contract plus `extraction-summary.json` and `review-queue.json`.

Do not use confidence as an accuracy percentage. When both attempts remain below 0.82 confidence, the selector favors material-value, numeric, text, and line coverage over a small confidence difference. It still sends low-confidence or disagreeing attempts to original-page review.

Use each queue entry's `recommended_action`:

- `enhanced-ocr-then-verify`: rerun the affected file with a measured stronger profile, then check the page.
- `manual-verify-original`: inspect the original page; more OCR is not an acceptance test. This action is mandatory for detected certificates, signing/stamp layouts, and complex tables.
- `verify-material-fields`: check each money, date, termination, liability, data, or IP fact that will be cited; do not interpret the entry as an OCR failure by itself.

The summary also records wall-clock elapsed time and start/end process memory, including peak RSS when the operating system exposes it. Treat this as batch-level process usage, not per-page attribution.

If OCR dependencies are unavailable, install the verified versions into a project-local directory:

```powershell
python -m pip install --target <input-dir>/.contract-tools/runtime -r <skill-dir>/scripts/requirements-ocr.txt
```

The project-local runtime is discovered automatically. For backward compatibility, the extractor also discovers `.tools/ocr` and `.tools/pymupdf`; use repeated `--vendor-dir` only for other locations. Existing results are reused only when the source, model bundle, profile, parameters, and page limit match; use `--force` only when intentional.

Run only one extractor process per output directory at a time. Atomic writes prevent partial JSON files, but concurrent writers can still race when replacing the shared summary.

Use `balanced` or `high` only for a selected file when the hybrid review queue justifies the extra cost:

```powershell
python <skill-dir>/scripts/extract_contracts.py <input-dir> --ocr always --quality-profile balanced --match <filename-fragment>
```

`--match` filters files, not pages. Within a file, the adaptive policy still decides which pages receive a second pass. Do not rerun the full corpus at higher resolution merely because pages contain material terms.

Read `references/ocr-quality.md` before changing scales, confidence thresholds, pixel caps, or ONNX models. A larger render is not automatically better; compare retained attempts and the review queue.

Validate extraction before analysis:

```powershell
python <skill-dir>/scripts/validate_extraction.py <input-dir>/.contract-analysis/extracted_text --require-complete --require-config --output <input-dir>/.contract-analysis/extraction-validation.json
```

Resolve every batch error and every quality/manual action in `review-queue.json`. Verify all material values actually used in the comparison. A structurally valid extraction is not proof that every material number is correct. When changing a model or quality policy, evaluate human-corrected sample pages with `scripts/evaluate_ocr.py`; confidence alone is not an accuracy metric.

Record every `enhanced-ocr-then-verify` and `manual-verify-original` result in `review-ledger.json` using the schema in `references/ocr-quality.md`. After review, enforce the human gate:

```powershell
python <skill-dir>/scripts/validate_extraction.py <input-dir>/.contract-analysis/extracted_text --require-complete --require-config --review-ledger <input-dir>/.contract-analysis/review-ledger.json --require-actions-resolved --output <input-dir>/.contract-analysis/final-extraction-validation.json
```

Do not use `--fail-on-review` as the final gate: material-only queue entries remain intentionally open until analysis decides which fields will be cited. Track each cited `verify-material-fields` check in `source-map.json` instead.

### 4. Normalize comparable terms

Read these references before analyzing terms:

- `references/comparison-schema.md`
- `references/citation-rules.md`
- `references/risk-checklist.md`
- `references/ocr-quality.md` when OCR is used or evaluated

Create two structured artifacts in the working output directory:

- `comparison-data.json`: parties, dates, term, scope, fees, payment, service levels, termination, liability, indemnity, insurance, confidentiality, data, IP, compliance, dispute terms, and notable omissions.
- `source-map.json`: every material finding mapped to contract ID, file, page, clause/section, exact or close source text, extraction method, and verification state.

Align units and bases before comparing values. Examples include monthly versus annual fees, tax-inclusive versus tax-exclusive prices, calendar versus business days, one-time versus recurring charges, and capped versus uncapped liability.

For pages flagged `complex-table`, use the page's `layout_table` row/cell candidate as a reading aid, then verify cited cells against the original image. Never treat clustered OCR boxes as authoritative table structure, especially with merged cells.

Use `not stated` for an absent term, `not found` when the review could not locate it, and `unclear` when the source is ambiguous or extraction quality is inadequate. Do not collapse these states.

### 5. Verify material findings

Open the original page at high resolution for every decision-critical item, especially:

- prices, percentages, dates, durations, notice periods, thresholds, and caps;
- tables, handwritten changes, stamps, signatures, and annex references;
- termination, liability, indemnity, exclusivity, renewal, audit, data, and IP clauses;
- contradictory or unusually favorable language.

Do not silently repair OCR. Preserve the extracted text, record the corrected reading separately, and mark how it was verified.

Keep three kinds of statements distinct:

- Contract fact: directly supported by a source citation.
- Analyst inference: reasoned from one or more contract facts.
- Recommendation: an action or negotiation position, with its rationale.

### 6. Build the HTML/CSS report

Copy `assets/report-template/index.html` and `styles.css` into the report directory, then replace every placeholder. The report must work offline and remain readable on desktop, mobile, and print/PDF.

Include these sections unless the user requests a narrower report:

- Executive conclusion and decision drivers.
- Contract/entity inventory.
- Clause-by-clause comparison matrix.
- Commercial and operational comparison.
- Risk and missing-term analysis.
- Per-contract synopsis.
- Method, limitations, and legal-review notice.

Put a visible page-level source citation next to every material claim. Prefer accurate ranges or qualitative severity over invented precision. Do not produce an arbitrary weighted score unless the user provides or approves the scoring model.

### 7. Validate and inspect

Run:

```powershell
python <skill-dir>/scripts/validate_report.py <output-dir>/index.html
```

Fix broken local links, unresolved placeholders, missing required sections, and missing citations. Then inspect the report in a browser at desktop and mobile widths and check print preview. Verify tables scroll or reflow without clipping and that source citations remain legible.

## Completion Standard

A comparison is complete only when:

- all in-scope files appear in the inventory and extraction summary;
- extraction errors and incomplete files are either resolved or disclosed;
- every quality/manual review action is resolved, and every cited material field is checked against the original page;
- any model or OCR-policy change is accepted using ground-truth CER, not confidence alone;
- material facts have page-level citations and verification states;
- differences use normalized units and distinguish absence from uncertainty;
- risks state whose perspective they use and why the term matters;
- the HTML report validates, loads locally, and prints cleanly;
- limitations include OCR uncertainty and a legal-review disclaimer.
