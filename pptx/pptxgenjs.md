# PptxGenJS 生成指南

使用 `pptxgenjs` 从零生成或套用华新背景时，先读取
`scripts/office/huaxin_theme_config.json`，所有画布、正文区、间距和组件 Token 以它为准。

## 基础对象

```js
const pptx = new pptxgen();
pptx.layout = "LAYOUT_16x9";
slide.addText("正文", { x: 0.5, y: 0.84, w: 9, h: 0.35, margin: 0 });
```

文本与形状需要精确贴合时显式设置 `margin: 0`。正文内容从 `y=0.84` 递推，不能恢复旧的 `y=1.31`。

## 圆角组件

```js
const radius = 0.125; // theme.components.corner_radius.default_in
slide.addShape(pptx.ShapeType.roundRect, {
  x: 0.5, y: 0.84, w: 2.1, h: 1.2,
  rectRadius: radius,
  fill: { color: "FFFFFF" },
  line: { color: "ADBACA", width: 0.75 }
});
```

- 卡片、指标块、面板和图片容器使用 `ROUNDED_RECTANGLE`，半径控制在 `0.104–0.167` 英寸。
- 同组组件使用同一半径；半径不得超过组件高度的一半。
- 状态标签可以使用满高胶囊：`Math.min(height / 2, 0.167)`。
- 标题栏、页脚、全宽色带、表格网格、背景和贴边强调条使用 `RECTANGLE`。
- `rectRadius` 只对 `ROUNDED_RECTANGLE` 生效；圆角容器不要叠加矩形强调条，否则角部会露出。

## 阴影和图像

阴影保持低透明度和短距离，例如 `opacity: 0.15`、`blur: 6`、`offset: 2`。不要复用已经传入过的选项对象，PptxGenJS 会原地转换部分数值。图片使用裁切填充，不拉伸；同组图片沿用相同圆角和边框语言。

## 生成后检查

```bash
python scripts/validate_skill.py --skill-dir . --pptx ../output/report.pptx
```

结构检查不能替代渲染和 PowerPoint/WPS 实测；交付前仍需逐页查看圆角、对齐、间距和文字可读性。
