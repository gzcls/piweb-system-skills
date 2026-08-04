---
name: data-to-pdf
description: "用结构化数据生成可打印的中文 PDF 报告（纯 Python，无需 Office/浏览器）。
  把标题、信息表、章节、表格、列表等内容写成一个 JSON，一键生成带表头配色、斑马纹、页码的规整 PDF。
  用于：生成报表/审查报告/台账/清单/汇总表/结算单并导出 PDF，把数据或表格做成正式 PDF 文档。
  操作：生成PDF、生成报告、导出PDF、制作报表、数据转PDF、JSON生成PDF、表格导出PDF、做台账/清单PDF。
  关键词：PDF报告, 生成PDF, 导出PDF, 报表, 报告, 台账, 清单, 汇总表, 中文PDF, 表格转PDF, reportlab, json to pdf, report, table, printable。"
required_tools:
  - bash
recommended_tools:
  - read
  - write
---
# data-to-pdf

用一个 **JSON spec** 描述报告内容（标题、信息表、章节、段落、列表、表格），生成排版规整、
可直接打印的 **中文 PDF**。引擎是 reportlab，中文用内置 CID 字体 `STSong-Light`，
**不依赖 Word / LibreOffice / 浏览器 / 系统字体**，适合在沙箱/服务器里稳定运行。

## 适用场景
- 用户说「生成一份 PDF 报告 / 把这些数据导出成 PDF / 做个报表/台账/清单 PDF」
- 已经有结构化数据（表格、清单、统计结果），需要落成正式 PDF 文档
- 需要**纯后端、无 Office、无浏览器**也能出 PDF 的场景
- 若需要「网页级精美样式（渐变/卡片/Web 字体）」，改用 `html-to-pdf` 技能

## 依赖
- Python 库 `reportlab`（沙箱镜像已预装）。中文字体内置，无需额外安装。

## 使用方法
1) 先把报告内容写成 JSON，存到 `/workspace/spec.json`（字段见下方「最小示例」，完整字段见 references/spec.md）。
2) 运行：
```bash
python3 /skills/data-to-pdf/scripts/report2pdf.py /workspace/spec.json -o /workspace/report.pdf
```

### 最小示例（spec.json）
```json
{
  "title": "示例报告",
  "subtitle": "副标题，可省略",
  "footer_pagenum": true,
  "info_table": {
    "headers": ["项目", "内容", "项目", "内容"],
    "rows": [["编号", "A-001", "金额", "16,100.07 元"]],
    "label_cols": [0, 2]
  },
  "sections": [
    {"heading": "一、摘要", "blocks": [
      {"type": "para", "text": "支持 <b>加粗</b>，\n 表示换行。"},
      {"type": "list", "ordered": true, "items": ["要点一", "要点二"]},
      {"type": "table",
       "headers": ["序号", "项目", "数量", "合价(元)"],
       "rows": [["1", "污水井", "4.00", "5,285.57"]],
       "aligns": ["center", "left", "right", "right"]},
      {"type": "table",
       "headers": ["编号", "级别", "问题"],
       "rows": [["P1", "重要", "口径不符"]],
       "level_col": 1}
    ]}
  ]
}
```
> 完整字段（页面尺寸/页边距、列宽 `widths_cm`、对齐 `aligns`、`subhead`/`note`/`spacer` 块、级别着色等）见
> `/skills/data-to-pdf/references/spec.md`，需要时再读取。同目录 `scripts/example_spec.json` 是可直接套用的模板。

## 输出说明
- 生成单个 PDF（路径由 `-o` 指定），A4、含表头深色反白、斑马纹、可选页脚「第 X 页 / 共 Y 页」。
- `level_col` 指定的列会按「重要(红)/中(橙)/低(绿)/提示(灰)」自动着色。

## 注意事项
- 输入 spec 与输出 PDF 都放 `/workspace`；`/skills` 只读，不能往里写。
- spec 必须是合法 JSON（注意中文引号、逗号）；字符串里换行写单反斜杠 `\n`（JSON 转义），加粗用 `<b>…</b>`。
- 本技能不渲染 HTML/CSS。要浏览器级视觉效果请用 `html-to-pdf`。
