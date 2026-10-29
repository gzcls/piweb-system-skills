# Document design guide

## Choose the visual direction

User templates and explicit preferences take precedence. For targeted edits, preserve the existing system. Otherwise select a restrained starting point:

| Document | Direction |
| --- | --- |
| Contract/agreement | Black/gray, light table headers, clear clauses, generous signature area |
| Business report | One accent color, meaningful data emphasis, consistent chart/table styles |
| Internal notice | Compact metadata, obvious title and action/date hierarchy |
| Branded proposal | Supplied brand typography/colors; retain legible print contrast |

These are design defaults, not statutory formatting standards. Do not force a contract template onto reports or add a cover/TOC to a short document without a reason.

## Typography and styles

Define reusable styles for Title, Heading1/2, body, table, note, and signature. Use Heading1 for primary clauses; do not skip levels just to obtain a visual size. Prefer automatic numbering for clauses users will insert/delete, with hanging indents for continuation lines.

Suggested sizes: title 18–22 pt, heading 12–14 pt, body 11–12 pt, table 10–11 pt, notes/footer 9–10 pt. Use no more than two font families by default. Configure East Asian and Latin font slots explicitly and Chinese language properties for Chinese documents. Check installed fonts before choosing them. For a formal Chinese print document, an installed Song-style body font and a bold sans-serif heading font are suitable; a sans-serif body is also appropriate for a modern screen-oriented report. Recheck pagination after any substitution.

Use regular gray text rather than synthetic Chinese italics for secondary information. Keep Chinese closing punctuation off line starts; inspect mixed English/CJK text and currency symbols. Use consistent paragraph spacing (e.g. 4–6 pt after body, 8–12 pt before headings) and readable proportional line spacing. Avoid fixed line heights that clip glyphs. In docx-js font sizes use half-points; paragraph spacing and geometry use twips/DXA (1 pt = 20 twips, 1 inch = 1440 twips).

## Geometry and tables

Compute dimensions once per section; derive table/column widths from the usable region. For A4 11906 DXA wide with 1300 DXA left/right margins, available width is 9306 DXA. Nested content must also subtract cell margins or column gaps. Set table width, column grid, and cell widths consistently. For a merged cell use a grid span and the sum of the covered columns, not an arbitrary width.

Align major blocks to shared left/right edges unless a deliberate inset serves the content. Use restrained thin borders, light gray fills for print, and adequate cell padding. Left-align descriptions, right-align amounts with consistent decimals, and center short unit/count fields when useful. Avoid accidental header wraps. Repeat header rows on multi-page tables; keep short item rows intact, but allow exceptionally long rows to split rather than overflow.

## Pagination and signing

Keep headings with the next paragraph, enable widow/orphan control, and keep small logical blocks together. Avoid applying keep-next to every paragraph, which can create large gaps. Fix excessive spacing and inefficient metadata/table layouts before adjusting margins. Accept another page when needed; do not trade readable type or usable signing space for an arbitrary page target.

For signature blocks, reserve practical blank space (often 25–35 mm, adjusted to the use). Keep party labels, signature/date fields, and their space together. Use tab leaders or paragraph borders for filling lines instead of repeated underscores. A dedicated signature page is valid when intentional and identified with the contract title/number.

For multi-page business documents, use dynamic PAGE/NUMPAGES fields and a compact identifier in the header when useful. Respect different first-page headers when the design needs them. Do not use decorative tables as horizontal rules.

## Verification and delivery

Use the SKILL.md acceptance checklist. Inspect each page at whole-page scale for balance and enlarged scale for glyphs/table edges. Verify DOCX structure separately from visual rendering: correct XML does not prove good layout, and a good PDF does not prove editable numbering is correct.

Render into a fresh directory, verify exit status, PDF existence/page count/text, and inspect final page images after edits. Preserve the editable DOCX; offer a PDF preview when generated. State whether validation used LibreOffice, Word, or WPS. Never claim identical rendering across applications unless checked there.


## Reusable implementation

Use `scripts/style_presets.js` rather than copying colors and sizes into each generator. Call `getStylePreset(name, overrides)` for values and `createDocumentStyles(preset)` for the Document `styles` option. Use `scripts/signature_helpers.js` for editable signature lines built from paragraph borders and separate date fields.

Run `scripts/design_validate.py` after rendering. Pass `--expected-font` for each required family; the validator normalizes PDF subset names. It checks both horizontal and vertical bounds, declared DOCX fonts, empty pages, and headings stranded near a page bottom. These checks supplement visual inspection; they do not prove Word/WPS equivalence.
