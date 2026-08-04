---
name: pptx
description: "PPT/PowerPoint 演示文稿（.pptx）创建、读取、编辑和模板化技能。用于生成汇报PPT/路演/培训/总结/方案，或读取、解析、提取、修改、合并、拆分现有 .pptx，处理模板、版式、母版、图片、图表、备注、批注、speaker notes。关键词：PPT, PPTX, PowerPoint, 幻灯片, 演示文稿, slides, deck, presentation, pitch deck, template。只要 .pptx 作为输入或输出就用本技能；不要用于纯 PDF/Word/Excel 任务。"
license: Proprietary. LICENSE.txt has complete terms
---

# PPTX Skill

## 🔴 核心交互与智能默认准则 (CRITICAL INTERACTION & SMART DEFAULT RULE)

**任何时候只要用户发出生成 PPT 的任务指令（如“生成天气报告”、“制作总结汇报”等），执行分支如下：**

1. **智能默认策略（Smart Default）**：
   - ⚠️ **核心准则**：为了提供最畅快的单回合生成体验，并防止因为 AI 在生成中途提问打断而导致前端输入框被锁定，**无需在中途向用户提问**。
   - **默认行为**：除非用户在初始指令中**明确表示拒绝模板**（例如说“不要模板背景”、“制作普通白底PPT”），AI 代理**必须默认且直接套用‘华新集团模板风格’**进行生成。
   - **交付提示**：在完成生成并交付 PPT 文件的回复最后，附带以下友好提示：
     > *“💡 提示：本报告已默认采用‘华新集团模板风格’进行对齐渲染。如果您需要常规无背景图的版本，请随时对我发出‘使用常规标准流程重新生成’的指令。”*

2. **分支处理路径**：
   * **情况 A：套用华新集团模板（默认）**：
     - 模板背景图和设计规约已**预先提取**并存储于技能包 `assets/` 目录和 `huaxin_theme_config.json` 中。生成时**直接使用下方路径**加载即可。
     - ⚠️ **核心警告**：由于您的执行目录（Cwd）通常是项目临时目录而非技能包目录，使用普通的相对路径会导致文件找不到。因此，**必须使用相对于本 SKILL.md 所在目录（即技能包根目录）的相对路径，在代码中通过环境变量或系统路径拼接得出绝对路径**来加载模板背景与配置文件：
       * 封面页/封底页背景图：`assets/huaxin_bg_title.jpeg`
       * 其他所有页面背景图：`assets/huaxin_bg_content.jpeg`（目录页、过渡页、正文页、总结页等）
       * 设计规约配置文件：`scripts/office/huaxin_theme_config.json`
     - 强制根据上述配置文件渲染各个幻灯片的标题及正文坐标、字号、颜色（具体规范详见 [pptxgenjs.md](pptxgenjs.md)）。
     - ⚠️ **华新模板设计硬约束（必须严格遵守）**：
       1. **背景图分配总规则**：**仅封面页（title_slide）和封底页（end_slide）使用 `huaxin_bg_title.jpeg` 作为背景图**；**除此之外的所有其他页面**（包括但不限于：目录页、过渡页、正文页、总结页、数据页、对比页等任何类型）**一律使用 `huaxin_bg_content.jpeg` 作为背景图**。配置文件中的 `bg_rules` 字段记录了此规则，各 slide 配置的 `bg_image` 字段也已预先设置好正确的背景图路径，生成代码时应直接引用对应的 slide 配置即可。
       2. **首页汇报人及日期标签**：生成 PPT 报告时，必须在首页右下角添加一行文字：`汇报人：华新智能小助手   |   汇报日期：XXXX年XX月XX日`。如用户没有明确说明汇报人，默认统一使用 `华新智能小助手`。此标签必须靠右边对齐（文字右侧边距留空至少 5mm），且字体颜色必须使用深色（如 `#44546A` 深灰或 `#334155`）。
       3. **首页底部白底区域规范**：首页背景图 `huaxin_bg_title.jpeg` 底部约 1/4 的区域底色为纯白色。除了在该区域的右下角放置上述深色的“汇报标签”外，主标题、副标题、气泡框等其他所有幻灯片内容必须尽可能布局在上方 3/4 的深色渐变区域中（如气泡框 y 坐标设为 2.8 左右）。绝对禁止在底部白色底色区域内显示任何浅色文字。
       4. **正文页白底内容区域规范**：正文页背景图 `huaxin_bg_content.jpeg` 只有顶部切角长条是深色，中间和下方（即 `content_area` 内）全都是白色/浅色背景。因此，在正文内容区域呈现的任何文字（包括正文列表、卡片内部说明、图表文本标签等），必须强制使用深色字体（如 `#44546A` 深灰、`#0F172A` 近黑等），严禁使用白色、浅灰色等浅色文字，保证高对比度和易读性。
   * **情况 B：用户明确拒绝模板（或要求常规版本重做）**：
     - 走常规的 PPT 生成流程（如使用 `pptxgenjs` 从零生成无背景图/普通配色的 PPT，或执行普通 XML 模板解包流）。

---

## Quick Reference

| Task | Guide |
|------|-------|
| Read/analyze content | `python -m markitdown presentation.pptx` |
| Create a seven-day weather report | Use `scripts/build_weather_report.py` with structured JSON (workflow below) |
| Edit or create from template (XML Mode) | Read [editing.md](editing.md) |
| Create from template (Background + Layout Mode) | Read [pptxgenjs.md](pptxgenjs.md) |
| Create from scratch | Read [pptxgenjs.md](pptxgenjs.md) |

---

## Reading Content

```bash
# Text extraction
python -m markitdown presentation.pptx

# Visual overview
python scripts/thumbnail.py presentation.pptx

# Raw XML
python scripts/office/unpack.py presentation.pptx unpacked/
```

---

## Editing Workflow

**Read [editing.md](editing.md) for full details.**

1. Analyze template with `thumbnail.py`
2. Unpack → manipulate slides → edit content → clean → pack

---

## Creating from Scratch / Visual Template

**Read [pptxgenjs.md](pptxgenjs.md) for full details.**

* **从零生成**：当没有现有模板可用时，使用 pptxgenjs 代码直接绘制所有图形和背景。
* **基于华新集团模板生成（推荐 · 默认）**：背景图和设计规约已**预先提取**，存储在 `assets/` 和 `huaxin_theme_config.json` 中。封面/封底使用 `huaxin_bg_title.jpeg`，其他所有页面（目录、过渡、正文、总结等）使用 `huaxin_bg_content.jpeg`。直接在 pptxgenjs 中读取配置并加载背景图即可，无需额外提取步骤。

---

## Seven-day Weather Reports (Preferred Structured Workflow)

For a seven-day weather deck, do **not** author a temporary JavaScript renderer. The stable renderer already handles the Huaxin template, layout, calculations, citations, and QA artifacts.

### Source rules

1. Fetch one primary page that actually contains the seven-day forecast table. A stale news article or search-result snippet is not forecast data.
2. Use a secondary source only when it independently provides forecast facts for the same dates.
3. Put a value into `precipProbability`, `humidity`, `uv`, or `alerts` only when the cited source explicitly provides it. Omit unknown optional fields.
4. Every alert requires its own `sourceUrl`. Never infer typhoons, warnings, or causes from a general news item.
5. Preserve the retrieval time in ISO-8601 form. The renderer places sources and retrieval time in the deck.

### Input and build

Write one JSON file using [references/weather-report.schema.json](references/weather-report.schema.json). Use [references/weather-report.example.json](references/weather-report.example.json) only as a field example; replace every sample value and URL with retrieved facts.

Then run exactly one build command from the project directory:

```bash
python /skills/pptx/scripts/build_weather_report.py weather.json weather-report.pptx --qa-dir .pptx-qa
```

The command validates the JSON, creates a seven-slide deck, checks PPTX package integrity, extracts text to `.pptx-qa/content.txt`, and renders `.pptx-qa/overview.jpg`.

After it succeeds:

1. Read `.pptx-qa/overview.jpg` once when the active model supports image input.
2. Read `.pptx-qa/content.txt` only when the build reports a content problem or visual inspection suggests missing text.
3. Fix the JSON and rerun the same command only when QA identifies a concrete issue.
4. Delete the temporary JSON and `.pptx-qa` directory after verification; keep the requested `.pptx`.

This workflow should normally use 5-8 tool calls total: source search/fetch, one JSON write, one build command, one visual read, and at most one evidence-based correction.

---

## Execution Discipline

- Prefer relative workspace paths. `/workspace` is the only absolute workspace root shared by bash and file tools; `/skills` is the read-only skill root.
- Use `write` for a new source/JSON file and `edit` for a precise correction. Do not rebuild a large file through shell heredocs or repeated `sed` substitutions.
- Batch syntax checking, generation, conversion, and deterministic QA in one maintained pipeline command when a bundled script exists.
- After the same error class appears twice, stop retrying variants of the same command. Read the exact failing region, change one thing with `edit`/`write`, then rerun the pipeline once.
- Do not probe the same artifact through `/workspace`, `/tmp`, `/data`, and host paths. If a read fails, verify the one relative path from the current project directory.

---

## Design Ideas

**Don't create boring slides.** Plain bullets on a white background won't impress anyone. Consider ideas from this list for each slide.

### Before Starting

- **Pick a bold, content-informed color palette**: The palette should feel designed for THIS topic. If swapping your colors into a completely different presentation would still "work," you haven't made specific enough choices.
- **Dominance over equality**: One color should dominate (60-70% visual weight), with 1-2 supporting tones and one sharp accent. Never give all colors equal weight.
- **Dark/light contrast**: Dark backgrounds for title + conclusion slides, light for content ("sandwich" structure). Or commit to dark throughout for a premium feel.
- **Commit to a visual motif**: Pick ONE distinctive element and repeat it — rounded image frames, icons in colored circles, thick single-side borders. Carry it across every slide.

### Color Palettes

Choose colors that match your topic — don't default to generic blue. Use these palettes as inspiration:

| Theme | Primary | Secondary | Accent |
|-------|---------|-----------|--------|
| **Midnight Executive** | `1E2761` (navy) | `CADCFC` (ice blue) | `FFFFFF` (white) |
| **Forest & Moss** | `2C5F2D` (forest) | `97BC62` (moss) | `F5F5F5` (cream) |
| **Coral Energy** | `F96167` (coral) | `F9E795` (gold) | `2F3C7E` (navy) |
| **Warm Terracotta** | `B85042` (terracotta) | `E7E8D1` (sand) | `A7BEAE` (sage) |
| **Ocean Gradient** | `065A82` (deep blue) | `1C7293` (teal) | `21295C` (midnight) |
| **Charcoal Minimal** | `36454F` (charcoal) | `F2F2F2` (off-white) | `212121` (black) |
| **Teal Trust** | `028090` (teal) | `00A896` (seafoam) | `02C39A` (mint) |
| **Berry & Cream** | `6D2E46` (berry) | `A26769` (dusty rose) | `ECE2D0` (cream) |
| **Sage Calm** | `84B59F` (sage) | `69A297` (eucalyptus) | `50808E` (slate) |
| **Cherry Bold** | `990011` (cherry) | `FCF6F5` (off-white) | `2F3C7E` (navy) |

### For Each Slide

**Every slide needs a visual element** — image, chart, icon, or shape. Text-only slides are forgettable.

**Layout options:**
- Two-column (text left, illustration on right)
- Icon + text rows (icon in colored circle, bold header, description below)
- 2x2 or 2x3 grid (image on one side, grid of content blocks on other)
- Half-bleed image (full left or right side) with content overlay

**Data display:**
- Large stat callouts (big numbers 60-72pt with small labels below)
- Comparison columns (before/after, pros/cons, side-by-side options)
- Timeline or process flow (numbered steps, arrows)

**Visual polish:**
- Icons in small colored circles next to section headers
- Italic accent text for key stats or taglines

### Typography

**Choose an interesting font pairing** — don't default to Arial. Pick a header font with personality and pair it with a clean body font.

| Header Font | Body Font |
|-------------|-----------|
| Georgia | Calibri |
| Arial Black | Arial |
| Calibri | Calibri Light |
| Cambria | Calibri |
| Trebuchet MS | Calibri |
| Impact | Arial |
| Palatino | Garamond |
| Consolas | Calibri |

| Element | Size |
|---------|------|
| Slide title | 36-44pt bold |
| Section header | 20-24pt bold |
| Body text | 14-16pt |
| Captions | 10-12pt muted |

### Spacing

- 0.5" minimum margins
- 0.3-0.5" between content blocks
- Leave breathing room—don't fill every inch

### Avoid (Common Mistakes)

- **Don't repeat the same layout** — vary columns, cards, and callouts across slides
- **Don't center body text** — left-align paragraphs and lists; center only titles
- **Don't skimp on size contrast** — titles need 36pt+ to stand out from 14-16pt body
- **Don't default to blue** — pick colors that reflect the specific topic
- **Don't mix spacing randomly** — choose 0.3" or 0.5" gaps and use consistently
- **Don't style one slide and leave the rest plain** — commit fully or keep it simple throughout
- **Don't create text-only slides** — add images, icons, charts, or visual elements; avoid plain title + bullets
- **Don't forget text box padding** — when aligning lines or shapes with text edges, set `margin: 0` on the text box or offset the shape to account for padding
- **Don't use low-contrast elements** — icons AND text need strong contrast against the background; avoid light text on light backgrounds or dark text on dark backgrounds. Specifically, for the Huaxin template: never use light/white/gray text in the white body areas (such as content_area of content slides or the bottom 1/4 area of title/end slides), and always use high-contrast dark text (like #44546A, #0F172A).
- **NEVER use accent lines under titles** — these are a hallmark of AI-generated slides; use whitespace or background color instead.

---

## QA (Required)

Run one deterministic content check and one visual check. Treat QA as evidence gathering: report and fix concrete defects, but do not invent a defect merely to force another iteration.

### Content QA

```bash
python -m markitdown output.pptx
```

Check for missing content, typos, wrong order.

**When using templates, check for leftover placeholder text:**

```bash
python -m markitdown output.pptx | grep -iE "xxxx|lorem|ipsum|this.*(page|slide).*layout"
```

If grep returns results, fix them before declaring success.

### Visual QA

For weather reports, use the generated `.pptx-qa/overview.jpg`. For other decks, convert slides to images as described below. Inspect directly with the `read` tool when the active model supports images. A separate reviewer is optional for high-stakes or unusually complex decks when a delegation tool is actually available; do not add delegation as a mandatory step for every deck.

Convert slides to images (see [Converting to Images](#converting-to-images)), then use this prompt:

```
Visually inspect these slides. Assume there are issues — find them.

Look for:
- Overlapping elements (text through shapes, lines through words, stacked elements)
- Text overflow or cut off at edges/box boundaries
- Decorative lines positioned for single-line text but title wrapped to two lines
- Source citations or footers colliding with content above
- Elements too close (< 0.3" gaps) or cards/sections nearly touching
- Uneven gaps (large empty area in one place, cramped in another)
- Insufficient margin from slide edges (< 0.5")
- Columns or similar elements not aligned consistently
- Low-contrast text (e.g., light gray text on cream-colored background)
- Low-contrast icons (e.g., dark icons on dark backgrounds without a contrasting circle)
- Text boxes too narrow causing excessive wrapping
- Leftover placeholder content

For each slide, list actual issues or areas of concern. A clean slide may be reported as clean.

Read and analyze these images:
1. /path/to/slide-01.jpg (Expected: [brief description])
2. /path/to/slide-02.jpg (Expected: [brief description])

Report ALL issues found, including minor ones.
```

### Verification Loop

1. Generate slides → Convert to images → Inspect
2. List concrete issues found
3. If issues exist, fix them and re-verify affected slides
4. Stop after a clean full pass; a no-op fix cycle is neither required nor useful

When the active model cannot inspect images, complete package/text checks, state that visual inspection was unavailable, and do not loop through unreadable image paths.

---

## Converting to Images

Convert presentations to individual slide images for visual inspection:

```bash
python scripts/office/soffice.py --headless --convert-to pdf output.pptx
pdftoppm -jpeg -r 150 output.pdf slide
```

This creates `slide-01.jpg`, `slide-02.jpg`, etc.

To re-render specific slides after fixes:

```bash
pdftoppm -jpeg -r 150 -f N -l N output.pdf slide-fixed
```

---

## Dependencies

- `pip install "markitdown[pptx]"` - text extraction
- `pip install Pillow` - thumbnail grids
- `npm install -g pptxgenjs` - creating from scratch
- LibreOffice (`soffice`) - PDF conversion (auto-configured for sandboxed environments via `scripts/office/soffice.py`)
- Poppler (`pdftoppm`) - PDF to images
