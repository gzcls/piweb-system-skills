# Citation and Verification Rules

## Source Map

Give every material source an ID and store it in `source-map.json`:

```json
{
  "source_id": "S-C001-004",
  "contract_id": "C001",
  "file": "Supplier A/agreement.pdf",
  "page": 12,
  "section": "8.2 Termination for Convenience",
  "source_text": "Customer may terminate ... on 30 days' written notice.",
  "extraction_method": "ocr",
  "ocr_confidence": 0.94,
  "verification": "original_checked",
  "review_note": null
}
```

Use the PDF's physical page index for `page`. If the printed page number differs, record it in `review_note`; never silently substitute one for the other.

## Evidence Hierarchy

Prefer evidence in this order:

1. Original page checked visually.
2. Embedded text with a stable page location.
3. OCR text checked against the rendered page.
4. OCR text not yet checked.
5. Filename, metadata, or contextual inference.

Items at levels 4 and 5 cannot support a high-confidence material conclusion without an explicit limitation.

## Mandatory Original-Page Review

Visually verify the source page for:

- currency values, percentages, dates, durations, thresholds, and liability caps;
- cells in dense tables and text near page breaks;
- handwritten amendments, stamps, signatures, strikethroughs, or poor scans;
- provisions that drive a high-severity risk or final recommendation;
- any apparent conflict between clauses, schedules, or documents.

## Verification States

- `extracted`: captured but not visually checked.
- `original_checked`: checked against the original page.
- `cross_checked`: checked against another clause, schedule, or document.
- `uncertain`: unresolved ambiguity or unreadable source.

Never label a finding verified solely because an OCR confidence score is high.

## Report Citations

Each material claim in the HTML should show a concise citation such as `C001, p. 12, §8.2`. Link to a local source only when the link works offline and does not expose a sensitive absolute path.

Use citations for:

- all numbers and dates;
- stated rights, duties, exceptions, caps, and omissions;
- each row of the comparison matrix when the row contains a conclusion;
- factual premises behind a risk or recommendation.

Analyst inferences should cite all supporting facts and be labeled as analysis. Recommendations should name their rationale but should not be presented as contract language.

## Privacy

Do not embed complete contract text in the report. Quote only the minimum needed to support a finding. Avoid exposing personal identifiers, bank details, signatures, or absolute filesystem paths unless the user explicitly needs them.
