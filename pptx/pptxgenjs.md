# PptxGenJS Tutorial

## Setup & Basic Structure

```javascript
const pptxgen = require("pptxgenjs");

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';  // or 'LAYOUT_16x10', 'LAYOUT_4x3', 'LAYOUT_WIDE'
pres.author = 'Your Name';
pres.title = 'Presentation Title';

let slide = pres.addSlide();
slide.addText("Hello World!", { x: 0.5, y: 0.5, fontSize: 36, color: "363636" });

pres.writeFile({ fileName: "Presentation.pptx" });
```

## Layout Dimensions

Slide dimensions (coordinates in inches):
- `LAYOUT_16x9`: 10" × 5.625" (default)
- `LAYOUT_16x10`: 10" × 6.25"
- `LAYOUT_4x3`: 10" × 7.5"
- `LAYOUT_WIDE`: 13.3" × 7.5"

---

## Text & Formatting

```javascript
// Basic text
slide.addText("Simple Text", {
  x: 1, y: 1, w: 8, h: 2, fontSize: 24, fontFace: "Arial",
  color: "363636", bold: true, align: "center", valign: "middle"
});

// Character spacing (use charSpacing, not letterSpacing which is silently ignored)
slide.addText("SPACED TEXT", { x: 1, y: 1, w: 8, h: 1, charSpacing: 6 });

// Rich text arrays
slide.addText([
  { text: "Bold ", options: { bold: true } },
  { text: "Italic ", options: { italic: true } }
], { x: 1, y: 3, w: 8, h: 1 });

// Multi-line text (requires breakLine: true)
slide.addText([
  { text: "Line 1", options: { breakLine: true } },
  { text: "Line 2", options: { breakLine: true } },
  { text: "Line 3" }  // Last item doesn't need breakLine
], { x: 0.5, y: 0.5, w: 8, h: 2 });

// Text box margin (internal padding)
slide.addText("Title", {
  x: 0.5, y: 0.3, w: 9, h: 0.6,
  margin: 0  // Use 0 when aligning text with other elements like shapes or icons
});
```

**Tip:** Text boxes have internal margin by default. Set `margin: 0` when you need text to align precisely with shapes, lines, or icons at the same x-position.

---

## Lists & Bullets

```javascript
// ✅ CORRECT: Multiple bullets
slide.addText([
  { text: "First item", options: { bullet: true, breakLine: true } },
  { text: "Second item", options: { bullet: true, breakLine: true } },
  { text: "Third item", options: { bullet: true } }
], { x: 0.5, y: 0.5, w: 8, h: 3 });

// ❌ WRONG: Never use unicode bullets
slide.addText("• First item", { ... });  // Creates double bullets

// Sub-items and numbered lists
{ text: "Sub-item", options: { bullet: true, indentLevel: 1 } }
{ text: "First", options: { bullet: { type: "number" }, breakLine: true } }
```

---

## Shapes

```javascript
slide.addShape(pres.shapes.RECTANGLE, {
  x: 0.5, y: 0.8, w: 1.5, h: 3.0,
  fill: { color: "FF0000" }, line: { color: "000000", width: 2 }
});

slide.addShape(pres.shapes.OVAL, { x: 4, y: 1, w: 2, h: 2, fill: { color: "0000FF" } });

slide.addShape(pres.shapes.LINE, {
  x: 1, y: 3, w: 5, h: 0, line: { color: "FF0000", width: 3, dashType: "dash" }
});

// With transparency
slide.addShape(pres.shapes.RECTANGLE, {
  x: 1, y: 1, w: 3, h: 2,
  fill: { color: "0088CC", transparency: 50 }
});

// Rounded rectangle (rectRadius only works with ROUNDED_RECTANGLE, not RECTANGLE)
// ⚠️ Don't pair with rectangular accent overlays — they won't cover rounded corners. Use RECTANGLE instead.
slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 1, y: 1, w: 3, h: 2,
  fill: { color: "FFFFFF" }, rectRadius: 0.1
});

// With shadow
slide.addShape(pres.shapes.RECTANGLE, {
  x: 1, y: 1, w: 3, h: 2,
  fill: { color: "FFFFFF" },
  shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.15 }
});
```

Shadow options:

| Property | Type | Range | Notes |
|----------|------|-------|-------|
| `type` | string | `"outer"`, `"inner"` | |
| `color` | string | 6-char hex (e.g. `"000000"`) | No `#` prefix, no 8-char hex — see Common Pitfalls |
| `blur` | number | 0-100 pt | |
| `offset` | number | 0-200 pt | **Must be non-negative** — negative values corrupt the file |
| `angle` | number | 0-359 degrees | Direction the shadow falls (135 = bottom-right, 270 = upward) |
| `opacity` | number | 0.0-1.0 | Use this for transparency, never encode in color string |

To cast a shadow upward (e.g. on a footer bar), use `angle: 270` with a positive offset — do **not** use a negative offset.

**Note**: Gradient fills are not natively supported. Use a gradient image as a background instead.

---

## Images

### Image Sources

```javascript
// From file path
slide.addImage({ path: "images/chart.png", x: 1, y: 1, w: 5, h: 3 });

// From URL
slide.addImage({ path: "https://example.com/image.jpg", x: 1, y: 1, w: 5, h: 3 });

// From base64 (faster, no file I/O)
slide.addImage({ data: "image/png;base64,iVBORw0KGgo...", x: 1, y: 1, w: 5, h: 3 });
```

### Image Options

```javascript
slide.addImage({
  path: "image.png",
  x: 1, y: 1, w: 5, h: 3,
  rotate: 45,              // 0-359 degrees
  rounding: true,          // Circular crop
  transparency: 50,        // 0-100
  flipH: true,             // Horizontal flip
  flipV: false,            // Vertical flip
  altText: "Description",  // Accessibility
  hyperlink: { url: "https://example.com" }
});
```

### Image Sizing Modes

```javascript
// Contain - fit inside, preserve ratio
{ sizing: { type: 'contain', w: 4, h: 3 } }

// Cover - fill area, preserve ratio (may crop)
{ sizing: { type: 'cover', w: 4, h: 3 } }

// Crop - cut specific portion
{ sizing: { type: 'crop', x: 0.5, y: 0.5, w: 2, h: 2 } }
```

### Calculate Dimensions (preserve aspect ratio)

```javascript
const origWidth = 1978, origHeight = 923, maxHeight = 3.0;
const calcWidth = maxHeight * (origWidth / origHeight);
const centerX = (10 - calcWidth) / 2;

slide.addImage({ path: "image.png", x: centerX, y: 1.2, w: calcWidth, h: maxHeight });
```

### Supported Formats

- **Standard**: PNG, JPG, GIF (animated GIFs work in Microsoft 365)
- **SVG**: Works in modern PowerPoint/Microsoft 365

---

## Icons

Use react-icons to generate SVG icons, then rasterize to PNG for universal compatibility.

### Setup

```javascript
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const { FaCheckCircle, FaChartLine } = require("react-icons/fa");

function renderIconSvg(IconComponent, color = "#000000", size = 256) {
  return ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComponent, { color, size: String(size) })
  );
}

async function iconToBase64Png(IconComponent, color, size = 256) {
  const svg = renderIconSvg(IconComponent, color, size);
  const pngBuffer = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + pngBuffer.toString("base64");
}
```

### Add Icon to Slide

```javascript
const iconData = await iconToBase64Png(FaCheckCircle, "#4472C4", 256);

slide.addImage({
  data: iconData,
  x: 1, y: 1, w: 0.5, h: 0.5  // Size in inches
});
```

**Note**: Use size 256 or higher for crisp icons. The size parameter controls the rasterization resolution, not the display size on the slide (which is set by `w` and `h` in inches).

### Icon Libraries

Install: `npm install -g react-icons react react-dom sharp`

Popular icon sets in react-icons:
- `react-icons/fa` - Font Awesome
- `react-icons/md` - Material Design
- `react-icons/hi` - Heroicons
- `react-icons/bi` - Bootstrap Icons

---

## Slide Backgrounds

```javascript
// Solid color
slide.background = { color: "F1F1F1" };

// Color with transparency
slide.background = { color: "FF3399", transparency: 50 };

// Image from URL
slide.background = { path: "https://example.com/bg.jpg" };

// Image from base64
slide.background = { data: "image/png;base64,iVBORw0KGgo..." };
```

---

## Tables

```javascript
slide.addTable([
  ["Header 1", "Header 2"],
  ["Cell 1", "Cell 2"]
], {
  x: 1, y: 1, w: 8, h: 2,
  border: { pt: 1, color: "999999" }, fill: { color: "F1F1F1" }
});

// Advanced with merged cells
let tableData = [
  [{ text: "Header", options: { fill: { color: "6699CC" }, color: "FFFFFF", bold: true } }, "Cell"],
  [{ text: "Merged", options: { colspan: 2 } }]
];
slide.addTable(tableData, { x: 1, y: 3.5, w: 8, colW: [4, 4] });
```

---

## Charts

```javascript
// Bar chart
slide.addChart(pres.charts.BAR, [{
  name: "Sales", labels: ["Q1", "Q2", "Q3", "Q4"], values: [4500, 5500, 6200, 7100]
}], {
  x: 0.5, y: 0.6, w: 6, h: 3, barDir: 'col',
  showTitle: true, title: 'Quarterly Sales'
});

// Line chart
slide.addChart(pres.charts.LINE, [{
  name: "Temp", labels: ["Jan", "Feb", "Mar"], values: [32, 35, 42]
}], { x: 0.5, y: 4, w: 6, h: 3, lineSize: 3, lineSmooth: true });

// Pie chart
slide.addChart(pres.charts.PIE, [{
  name: "Share", labels: ["A", "B", "Other"], values: [35, 45, 20]
}], { x: 7, y: 1, w: 5, h: 4, showPercent: true });
```

### Better-Looking Charts

Default charts look dated. Apply these options for a modern, clean appearance:

```javascript
slide.addChart(pres.charts.BAR, chartData, {
  x: 0.5, y: 1, w: 9, h: 4, barDir: "col",

  // Custom colors (match your presentation palette)
  chartColors: ["0D9488", "14B8A6", "5EEAD4"],

  // Clean background
  chartArea: { fill: { color: "FFFFFF" }, roundedCorners: true },

  // Muted axis labels
  catAxisLabelColor: "64748B",
  valAxisLabelColor: "64748B",

  // Subtle grid (value axis only)
  valGridLine: { color: "E2E8F0", size: 0.5 },
  catGridLine: { style: "none" },

  // Data labels on bars
  showValue: true,
  dataLabelPosition: "outEnd",
  dataLabelColor: "1E293B",

  // Hide legend for single series
  showLegend: false,
});
```

**Key styling options:**
- `chartColors: [...]` - hex colors for series/segments
- `chartArea: { fill, border, roundedCorners }` - chart background
- `catGridLine/valGridLine: { color, style, size }` - grid lines (`style: "none"` to hide)
- `lineSmooth: true` - curved lines (line charts)
- `legendPos: "r"` - legend position: "b", "t", "l", "r", "tr"

---

## Slide Masters

```javascript
pres.defineSlideMaster({
  title: 'TITLE_SLIDE', background: { color: '283A5E' },
  objects: [{
    placeholder: { options: { name: 'title', type: 'title', x: 1, y: 2, w: 8, h: 2 } }
  }]
});

let titleSlide = pres.addSlide({ masterName: "TITLE_SLIDE" });
titleSlide.addText("My Title", { placeholder: "title" });
```

---

## Common Pitfalls

⚠️ These issues cause file corruption, visual bugs, or broken output. Avoid them.

1. **NEVER use "#" with hex colors** - causes file corruption
   ```javascript
   color: "FF0000"      // ✅ CORRECT
   color: "#FF0000"     // ❌ WRONG
   ```

2. **NEVER encode opacity in hex color strings** - 8-char colors (e.g., `"00000020"`) corrupt the file. Use the `opacity` property instead.
   ```javascript
   shadow: { type: "outer", blur: 6, offset: 2, color: "00000020" }          // ❌ CORRUPTS FILE
   shadow: { type: "outer", blur: 6, offset: 2, color: "000000", opacity: 0.12 }  // ✅ CORRECT
   ```

3. **Use `bullet: true`** - NEVER unicode symbols like "•" (creates double bullets)

4. **Use `breakLine: true`** between array items or text runs together

5. **Avoid `lineSpacing` with bullets** - causes excessive gaps; use `paraSpaceAfter` instead

6. **Each presentation needs fresh instance** - don't reuse `pptxgen()` objects

7. **NEVER reuse option objects across calls** - PptxGenJS mutates objects in-place (e.g. converting shadow values to EMU). Sharing one object between multiple calls corrupts the second shape.
   ```javascript
   const shadow = { type: "outer", blur: 6, offset: 2, color: "000000", opacity: 0.15 };
   slide.addShape(pres.shapes.RECTANGLE, { shadow, ... });  // ❌ second call gets already-converted values
   slide.addShape(pres.shapes.RECTANGLE, { shadow, ... });

   const makeShadow = () => ({ type: "outer", blur: 6, offset: 2, color: "000000", opacity: 0.15 });
   slide.addShape(pres.shapes.RECTANGLE, { shadow: makeShadow(), ... });  // ✅ fresh object each time
   slide.addShape(pres.shapes.RECTANGLE, { shadow: makeShadow(), ... });
   ```

8. **Don't use `ROUNDED_RECTANGLE` with accent borders** - rectangular overlay bars won't cover rounded corners. Use `RECTANGLE` instead.
   ```javascript
   // ❌ WRONG: Accent bar doesn't cover rounded corners
   slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1, y: 1, w: 3, h: 1.5, fill: { color: "FFFFFF" } });
   slide.addShape(pres.shapes.RECTANGLE, { x: 1, y: 1, w: 0.08, h: 1.5, fill: { color: "0891B2" } });

   // ✅ CORRECT: Use RECTANGLE for clean alignment
   slide.addShape(pres.shapes.RECTANGLE, { x: 1, y: 1, w: 3, h: 1.5, fill: { color: "FFFFFF" } });
   slide.addShape(pres.shapes.RECTANGLE, { x: 1, y: 1, w: 0.08, h: 1.5, fill: { color: "0891B2" } });
   ```

---

## Quick Reference

- **Shapes**: RECTANGLE, OVAL, LINE, ROUNDED_RECTANGLE
- **Charts**: BAR, LINE, PIE, DOUGHNUT, SCATTER, BUBBLE, RADAR
- **Layouts**: LAYOUT_16x9 (10"×5.625"), LAYOUT_16x10, LAYOUT_4x3, LAYOUT_WIDE
- **Alignment**: "left", "center", "right"
- **Chart data labels**: "outEnd", "inEnd", "center"

---

## 使用 PPT 模板背景与样式约束（方案 B 最佳实践）

若需要根据华新集团 PPT 模板生成幻灯片，最稳定且灵活的方式是：**读取存储于技能目录下的设计规约配置 `huaxin_theme_config.json` 来进行精准排版**。

### 1. 工作流说明

模板背景图和设计规约已**预先提取**，直接使用以下虚拟路径资源（可在容器与宿主机中通用）：
   * 封面页(title_slide)/封底页(end_slide) 背景图：`assets/huaxin_bg_title.jpeg`
   * **其他所有页面** 背景图：`assets/huaxin_bg_content.jpeg`（包括目录页、过渡页、正文页、总结页等）
   * 设计规约配置文件虚拟路径：`/skills/pptx/scripts/office/huaxin_theme_config.json`

> ⚠️ **背景图分配总规则**：只有封面和封底两种页面使用 `huaxin_bg_title.jpeg`，其余任何类型的幻灯片一律使用 `huaxin_bg_content.jpeg`。配置文件的 `bg_rules` 字段和各 slide 类型的 `bg_image` 字段已预设好正确映射，生成代码时直接引用即可。

使用 `pptxgenjs` 读取配置参数，以技能包根目录 `/skills/pptx` 拼接上述相对图片路径为背景，并在配置限定的坐标及样式内绘制标题和正文。
   ⚠️ **注意**：编写生成代码时必须使用上述**虚拟绝对路径** `/skills/pptx/...`（即在配置的相对路径前拼接 `/skills/pptx`）。如果使用硬编码的 Windows 绝对路径，会导致沙箱运行失败；如果使用普通的相对路径，在项目执行目录不一致时也会导致读取失败。

### 2. 完整生成代码示例

```javascript
const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

// 技能包统一虚拟根目录
const SKILL_ROOT = "/skills/pptx";

// 1. 读取设计规约配置 (使用虚拟绝对路径)
const config = JSON.parse(fs.readFileSync(path.join(SKILL_ROOT, "scripts/office/huaxin_theme_config.json"), "utf8"));

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';

// ------------------------------------------------
// 2. 首页（Title Slide）
// ------------------------------------------------
let slide1 = pres.addSlide();
const titleSlideConf = config.slides.title_slide;

// 设置背景图 (拼接相对路径与虚拟根目录)
slide1.background = { path: path.join(SKILL_ROOT, titleSlideConf.bg_image) };

// 添加主标题 (完全匹配模板顶部或指定区域的位置、字体、字号和颜色)
slide1.addText("2026年企业核心战略规划报告", {
  x: titleSlideConf.title.x,
  y: titleSlideConf.title.y,
  w: titleSlideConf.title.w,
  h: titleSlideConf.title.h,
  fontFace: config.fonts.header,
  fontSize: titleSlideConf.title.fontSize,
  color: titleSlideConf.title.color,
  bold: titleSlideConf.title.bold,
  align: titleSlideConf.title.align,
  margin: 0
});
// 添加副标题
slide1.addText("2026年6月4日 · 广州市番禺区", {
  x: titleSlideConf.subtitle.x,
  y: titleSlideConf.subtitle.y,
  w: titleSlideConf.subtitle.w,
  h: titleSlideConf.subtitle.h,
  fontFace: config.fonts.body,
  fontSize: titleSlideConf.subtitle.fontSize,
  color: titleSlideConf.subtitle.color,
  bold: false,
  align: titleSlideConf.subtitle.align,
  margin: 0
});

// 添加汇报人与汇报日期标签 (强制在首页白底区域的右下角右对齐，使用深色字体以保持高对比度)
slide1.addText("汇报人：华新智能小助手   |   汇报日期：2026年6月4日", {
  x: 4.70,
  y: 5.00,
  w: 4.8,
  h: 0.22,
  fontFace: config.fonts.body,
  fontSize: 10.5,
  color: config.colors.primary, // 使用官方提取的主色深灰色，严禁使用白色/浅色
  align: "right",
  margin: 0
});

// ------------------------------------------------
// 3. 目录页（TOC Slide） —— 使用 content 背景
// ------------------------------------------------
let slideToc = pres.addSlide();
const tocSlideConf = config.slides.toc_slide;

// ⚠️ 目录页使用 huaxin_bg_content.jpeg（非 title 背景）
slideToc.background = { path: path.join(SKILL_ROOT, tocSlideConf.bg_image) };

slideToc.addText("目 录", {
  x: tocSlideConf.title.x,
  y: tocSlideConf.title.y,
  w: tocSlideConf.title.w,
  h: tocSlideConf.title.h,
  fontFace: config.fonts.header,
  fontSize: tocSlideConf.title.fontSize,
  color: tocSlideConf.title.color,
  bold: tocSlideConf.title.bold,
  align: tocSlideConf.title.align,
  margin: 0
});

slideToc.addText([
  { text: "一、核心业务增长概览", options: { bullet: true, breakLine: true } },
  { text: "二、下季度关键落地行动", options: { bullet: true } }
], {
  x: tocSlideConf.content_area.x,
  y: tocSlideConf.content_area.y,
  w: tocSlideConf.content_area.w,
  h: tocSlideConf.content_area.h,
  fontFace: config.fonts.body,
  fontSize: 18,
  color: config.colors.primary, // 深色字体，白底高对比
  paraSpaceAfter: 12
});

// ------------------------------------------------
// 4. 过渡页（Transition Slide） —— 使用 content 背景
// ------------------------------------------------
let slideTrans = pres.addSlide();
const transSlideConf = config.slides.transition_slide;

// ⚠️ 过渡页使用 huaxin_bg_content.jpeg（非 title 背景）
slideTrans.background = { path: path.join(SKILL_ROOT, transSlideConf.bg_image) };

slideTrans.addText("第一部分", {
  x: transSlideConf.title.x,
  y: transSlideConf.title.y,
  w: transSlideConf.title.w,
  h: transSlideConf.title.h,
  fontFace: config.fonts.header,
  fontSize: transSlideConf.title.fontSize,
  color: transSlideConf.title.color,
  bold: transSlideConf.title.bold,
  align: transSlideConf.title.align,
  margin: 0
});

slideTrans.addText("核心业务增长概览", {
  x: transSlideConf.content_area.x,
  y: 2.5,
  w: transSlideConf.content_area.w,
  h: 1.0,
  fontFace: config.fonts.header,
  fontSize: 28,
  color: config.colors.brand_blue, // 深色，白底区域高对比
  bold: true,
  align: "center",
  margin: 0
});

// ------------------------------------------------
// 5. 正文页（Content Slides） —— 使用 content 背景
// ------------------------------------------------
const contentData = [
  {
    title: "一、 核心业务增长概览",
    bullets: [
      { text: "业务收入同比增长达 35%，突破历史新高", options: { bullet: true, breakLine: true } },
      { text: "新兴市场占有率提升至 18%，成为主要增长引擎", options: { bullet: true, breakLine: true } },
      { text: "客户留存率保持在 92% 的行业领先水平", options: { bullet: true } }
    ]
  },
  {
    title: "二、 下季度关键落地行动",
    bullets: [
      { text: "加速数字化平台升级，优化用户交互体验", options: { bullet: true, breakLine: true } },
      { text: "拓展华东及华南线下服务网络，增加 50 个网点", options: { bullet: true, breakLine: true } },
      { text: "开展新一轮人才引进计划，聚焦研发与算法专家", options: { bullet: true } }
    ]
  }
];

const contentSlideConf = config.slides.content_slide;

contentData.forEach(item => {
  let slide = pres.addSlide();
  
  // 正文页使用 huaxin_bg_content.jpeg
  slide.background = { path: path.join(SKILL_ROOT, contentSlideConf.bg_image) };
  
  // 添加正文标题 (匹配模板顶部位置与字号颜色)
  slide.addText(item.title, {
    x: contentSlideConf.title.x,
    y: contentSlideConf.title.y,
    w: contentSlideConf.title.w,
    h: contentSlideConf.title.h,
    fontFace: config.fonts.header,
    fontSize: contentSlideConf.title.fontSize,
    color: contentSlideConf.title.color,
    bold: contentSlideConf.title.bold,
    align: contentSlideConf.title.align,
    margin: 0
  });
  
  // 在限定的正文区域 (content_area) 绘制内容，确保绝对不对齐冲突或溢出。
  // ⚠️ 警告：正文主要背景是白色底色，必须强制使用高可读性的深色字体(如 "44546A"、"0F172A"、"334155")，绝不允许使用白色/灰色等浅色文字！
  slide.addText(item.bullets, {
    x: contentSlideConf.content_area.x,
    y: contentSlideConf.content_area.y,
    w: contentSlideConf.content_area.w,
    h: contentSlideConf.content_area.h,
    fontFace: config.fonts.body,
    fontSize: 15,
    color: config.colors.primary, // 官方深蓝灰色主文字，高对比度
    paraSpaceAfter: 8
  });
});

// ------------------------------------------------
// 6. 总结页（Summary Slide） —— 使用 content 背景
// ------------------------------------------------
let slideSummary = pres.addSlide();
const summarySlideConf = config.slides.summary_slide;

// ⚠️ 总结页使用 huaxin_bg_content.jpeg（非 title 背景）
slideSummary.background = { path: path.join(SKILL_ROOT, summarySlideConf.bg_image) };

slideSummary.addText("总结与展望", {
  x: summarySlideConf.title.x,
  y: summarySlideConf.title.y,
  w: summarySlideConf.title.w,
  h: summarySlideConf.title.h,
  fontFace: config.fonts.header,
  fontSize: summarySlideConf.title.fontSize,
  color: summarySlideConf.title.color,
  bold: summarySlideConf.title.bold,
  align: summarySlideConf.title.align,
  margin: 0
});

slideSummary.addText([
  { text: "本季度整体表现超出预期目标", options: { bullet: true, breakLine: true } },
  { text: "下阶段将聚焦数字化转型和人才建设", options: { bullet: true } }
], {
  x: summarySlideConf.content_area.x,
  y: summarySlideConf.content_area.y,
  w: summarySlideConf.content_area.w,
  h: summarySlideConf.content_area.h,
  fontFace: config.fonts.body,
  fontSize: 15,
  color: config.colors.primary,
  paraSpaceAfter: 8
});

// ------------------------------------------------
// 7. 封底页（End Slide） —— 使用 title 背景
// ------------------------------------------------
let slideEnd = pres.addSlide();
const endSlideConf = config.slides.end_slide;

// ⚠️ 封底页使用 huaxin_bg_title.jpeg（与封面页相同）
slideEnd.background = { path: path.join(SKILL_ROOT, endSlideConf.bg_image) };

slideEnd.addText("谢谢观看\nThank You", {
  x: endSlideConf.title.x,
  y: endSlideConf.title.y,
  w: endSlideConf.title.w,
  h: endSlideConf.title.h,
  fontFace: config.fonts.header,
  fontSize: endSlideConf.title.fontSize,
  color: endSlideConf.title.color,
  bold: endSlideConf.title.bold,
  align: endSlideConf.title.align,
  margin: 0
});

// 8. 导出文件
pres.writeFile({ fileName: "Generated_Template_Presentation.pptx" });
```

