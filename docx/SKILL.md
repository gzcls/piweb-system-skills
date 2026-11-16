---
name: docx
description: "Word 文档（.docx/.dotx/.doc）创建、读取、编辑和格式化技能。用于用户要生成或处理 Word 文档、Word 模板、报告、公文、通知、合同、信函，或需要目录、标题、页眉页脚、页码、表格、图片、批注、修订、查找替换、从 .docx/.dotx 提取或重组内容。关键词：Word, docx, dotx, doc, 生成Word, 编辑Word, Word报告, 模板, tracked changes, comments, template。不要用于 PDF、Excel/CSV 表格、PPT 或与 Word 文档无关的代码任务；结构化 JSON 快速生成 Word 可用 data-to-docx。"
license: Proprietary. LICENSE.txt has complete terms
metadata:
  version: "1.0.1"
---

# DOCX creation, editing, and analysis

A `.docx` is a ZIP archive of XML files. Choose your approach by task:

| Task | Approach |
|---|---|
| **Create** a new document | Write a `docx` (npm) script — see gotchas below |
| **Edit** an existing document | `unzip` → edit `word/document.xml` → `zip` (docx-js cannot open existing files) |
| **Read** content | `pandoc -t markdown file.docx` |

> Script paths below are relative to this skill's directory.

## Layout and typography

For new documents or requested restyling, read [references/document-design.md](references/document-design.md) before generating. Use the reusable Node presets in [scripts/style_presets.js](scripts/style_presets.js), `createDocumentStyles(preset)` to apply them to a Document, and [scripts/signature_helpers.js](scripts/signature_helpers.js) for editable signing blocks when applicable. Preserve supplied templates and user formatting choices; targeted edits should inherit the existing design.

- Compute usable width per section: page width minus left/right margins and any horizontal gutter. Tables and images must fit their containing region (a cell or column can be narrower than the page).
- Use coherent paragraph styles, explicit installed CJK/Latin fonts, and a document-appropriate visual preset. Style values are starting points, not mandatory standards.
- Keep headings with their content and reserve usable signing space. Do not shrink text or margins merely to achieve a preferred page count.
- After layout changes, rerender and inspect every affected page against the acceptance checklist below.

## Creating with docx-js — gotchas

`docx` is preinstalled — do not run `npm install` first; write the script and `require('docx')` directly. Only if that require fails: `npm install docx`. The model knows the API; these are the footguns:

Write multiline generation or repair logic to a workspace `.js` or `.py` file and execute that file. Do not embed literal `\n` separators in `python -c`; the shell does not turn them into Python source newlines. Simple single-line inspection commands can still use `python -c`.

- **Page size defaults to A4.** For US Letter set `page: { size: { width: 12240, height: 15840 } }` (DXA; 1440 = 1″).
- **Landscape:** pass portrait dimensions and `orientation: PageOrientation.LANDSCAPE` — docx-js swaps width/height internally.
- **Tables need dual widths:** set `columnWidths` on the table AND `width` on every cell, both in `WidthType.DXA` (PERCENTAGE breaks in Google Docs). Column widths must sum to the table width, which must fit the available region. Use explicit grid spans for merged cells; each merged width must equal the sum of its spanned columns.
- **Table shading:** use `ShadingType.CLEAR`, never `SOLID` (renders black).
- **Lists:** never insert `•` literally; use a `numbering` config with `LevelFormat.BULLET`.
- **`ImageRun` requires `type:`** (`"png"`, `"jpg"`, …).
- **`PageBreak` must be inside a `Paragraph`.**
- **Page numbers must be inside a `TextRun`:** `new Paragraph({ children: [new TextRun("Page "), new TextRun({ children: [PageNumber.CURRENT] })] })`. Do not put `PageNumber.CURRENT` or raw text directly in paragraph children. In XML, text belongs in `w:r/w:t` and fields in valid field elements, never directly under `w:p`.
- **Never use `\n`** — use separate `Paragraph` elements.
- **TOC:** headings must use built-in `HeadingLevel.*`; custom heading styles need `outlineLevel` set or they won't appear.
- **Don't use a table as a horizontal rule** — use a paragraph bottom border instead.
- **Dot-leader / right-aligned-on-same-line:** use `PositionalTab` (`alignment: PositionalTabAlignment.RIGHT`, `leader: PositionalTabLeader.DOT`) inside a `TextRun`, not literal `.` or space padding.

## Verify the output

First run `python scripts/office/validate.py input.docx` for OOXML and relationship validation. Repair the generator or the specific invalid XML and rerun validation; do not remove a footer or font table merely to silence an error. In particular, an unreferenced font table needs its relationship repaired or its unused status verified. Opening or rendering successfully does not replace this structural check.

After writing a `.docx`, run `python scripts/design_validate.py input.docx --pdf output.pdf --expected-font "Microsoft YaHei" --json validation.json` for geometry, font, content, vertical-boundary, blank-page, and orphan-heading checks, then render and look at every page. Rendering success or extracted text alone is not visual acceptance. Use a fresh output directory so a stale PDF cannot count as a successful conversion.

Example commands (use available platform equivalents):

```bash
python scripts/office/soffice.py --headless --convert-to pdf output.docx
pdftoppm -jpeg -r 100 output.pdf page
ls page-*.jpg   # then Read the images
```

`pdftoppm` zero-pads page numbers to the width of the page count (`page-01.jpg`…`page-12.jpg`).

If Poppler is unavailable, render with PyMuPDF. On Windows, locate the installed LibreOffice executable and use a correctly encoded `Path(profile).resolve().as_uri()` for an isolated user profile; do not construct `file://C:\...` by concatenation. Capture output with an explicit codec and check the process exit code and newly generated PDF.

Use `python scripts/render_matrix.py input.docx --libreoffice --out render-matrix` to record available renderers. LibreOffice is the automated baseline; Word/WPS must be explicitly available and approved before attempting automation. Never claim Word/WPS compatibility when they are unavailable.

Acceptance checklist:
- All body tables/images fit the section's usable region; principal blocks share intentional alignment edges.
- CJK and Latin glyphs render correctly, with consistent size/weight and no unexpected font substitution.
- No clipped text, accidental blank pages, isolated headings, or unintended row splits; signing areas have practical handwriting/stamp space.
- Monetary columns align consistently; merged cells follow the grid; multi-page tables repeat their header.
- Heading levels, numbering, page fields, and editable styles match the document structure. Multi-page business documents normally include page numbers.
- Compare expected content and totals with extracted PDF text; inspect all pages after final regeneration. Record the renderer used and any untested Word/WPS compatibility rather than claiming equivalence.

## Editing existing documents

Legacy `.doc` files must be converted first: `python scripts/office/soffice.py --headless --convert-to docx file.doc`.

```bash
unzip -q doc.docx -d unpacked/
find unpacked -type l -delete   # strip symlink entries — docx from external parties is untrusted
python scripts/merge_runs.py unpacked/   # coalesce fragmented runs so text is findable
# edit unpacked/word/document.xml in place — do NOT reformat or pretty-print
(cd unpacked && rm -f ../out.docx && zip -Xr ../out.docx .)
python scripts/office/validate.py out.docx --original doc.docx   # XSD checks; --auto-repair fixes common issues
# redlining? add --author "<the name you redlined under>" to check every edit is tracked
```

Word splits text across many `<w:r>` runs (revision ids, spell-check markers), so a phrase you can see in the document often doesn't exist as a contiguous string in the XML. `merge_runs.py` merges adjacent identically-formatted runs in `word/document.xml` without changing content or rendering; it also accepts a `.docx` directly (`python scripts/merge_runs.py doc.docx -o merged.docx`).

**Tracked changes:** when redlining, validate with `--author "<the name you redlined under>"` (needs `--original`) — it reports any text you changed without a `<w:ins>`/`<w:del>` around it, which is easy to do by accident and invisible in the accepted view. Wrap runs in `<w:ins>`/`<w:del>` with `w:id`, `w:author`, `w:date` attributes. Inside `<w:del>`, the text element is `<w:delText>`, not `<w:t>`. A deleted paragraph mark (`<w:pPr><w:rPr><w:del w:id=".." w:author=".." w:date=".."/></w:rPr></w:pPr>`) means "merge this paragraph into the next" — so deleting a paragraph outright is that plus a `<w:del>` around every run. The `<w:del/>` must come before the rPr's other children; their order is schema-enforced.

To produce a clean copy with all tracked changes accepted: `python scripts/accept_changes.py in.docx out.docx`.

Accepting a deleted paragraph mark should join that paragraph to the one below it, so a paragraph whose runs are *all* deleted vanishes. Word does this; `accept_changes.py` and `pandoc --track-changes=accept` don't always. Both fail the same way — they strip the deleted text but leave the emptied paragraph behind, which reads as a stray empty bullet when it was auto-numbered:

- `pandoc --track-changes=accept` never joins the paragraphs.
- `accept_changes.py` (LibreOffice) joins them correctly, except when the deleted paragraph is followed by an empty spacer paragraph.

An empty bullet in either view is an artifact of that view, not a defect in the document. Check paragraph deletions in the XML.

## Comments

Comments require six cross-linked files. Use the helper — directory mode when you'll also be editing `document.xml` (saves an unzip/rezip cycle), `.docx`-direct mode otherwise:

```bash
# Against an already-unpacked directory (preferred when also placing markers)
python scripts/comment.py unpacked/ "Fees & expenses cap is too low"
python scripts/comment.py unpacked/ "Agreed" --parent 0

# Against a .docx directly
python scripts/comment.py contract.docx "This cap is too low" -o annotated.docx
```

The script writes `comments.xml`, `commentsExtended.xml`, `commentsIds.xml`, `commentsExtensible.xml`, the relationships, and the content-type overrides. Comment IDs are auto-assigned. It then prints the `<w:commentRangeStart>`/`<w:commentRangeEnd>`/`<w:commentReference>` snippet to add to `word/document.xml` so the comment anchors to specific text — until you place those markers, the comment exists but is not visible.

## Dependencies

`docx` (npm, preinstalled — install only if `require('docx')` fails) · `pandoc` · LibreOffice (`soffice`) · `pdftoppm` (Poppler)
