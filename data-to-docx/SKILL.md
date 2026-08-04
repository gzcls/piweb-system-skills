---
name: data-to-docx
description: "用结构化数据生成 Word 文档（.docx，支持中文，无需 MS Office）。把标题、信息表、章节、段落、列表、表格写成一个 JSON，一键生成排版规整、带表头配色和页码的 Word 文档。
  用于：生成报告/通知/公文/合同/周报/纪要并导出 Word，把数据或表格做成正式 .docx 文档。
  操作：生成Word、生成docx、导出Word、做通知公文、写报告、表格导出Word、数据转Word。
  关键词：Word, docx, 生成Word, 导出Word, 公文, 通知, 合同, 报告, 周报, 纪要, 中文文档, 表格, python-docx, word document, report。"
required_tools:
  - bash
recommended_tools:
  - read
  - write
---
# data-to-docx

用一个 **JSON spec** 描述内容，生成可编辑的 **Word（.docx）** 文档。引擎是 `python-docx`，
中文用 宋体/黑体，**不需要 MS Office**。spec 与 `data-to-pdf` 基本一致——**同一份 spec 可分别出 PDF 或 Word**。

## 适用场景
- 用户说「生成一份 Word / 导出 docx / 做个通知、公文、合同、周报、纪要」
- 已有结构化数据，需要落成**可继续编辑**的 Word（要不可编辑的最终稿用 `data-to-pdf` 或 `office-convert`）

## 依赖
- Python 库 `python-docx`（沙箱镜像已预装）。

## 使用方法
```bash
python3 /skills/data-to-docx/scripts/docx_build.py /workspace/spec.json -o /workspace/report.docx
```

### spec 最小示例
```json
{
  "title": "项目周报",
  "subtitle": "2026 年第 25 周",
  "footer_pagenum": true,
  "info_table": {"headers": ["项目","内容","项目","内容"],
    "rows": [["部门","工程部","负责人","张三"]], "label_cols": [0,2]},
  "sections": [
    {"heading": "一、本周进展", "blocks": [
      {"type": "para", "text": "支持 <b>加粗</b> 与 \n 换行。"},
      {"type": "list", "ordered": true, "items": ["完成 A", "推进 B"]},
      {"type": "table", "headers": ["任务","状态"], "rows": [["A","完成"],["B","进行中"]],
       "aligns": ["left","center"]}
    ]}
  ]
}
```

### block 类型
`para` 段落 / `subhead` 小标题 / `list`(+`ordered`) 列表 / `table`(`headers`,`rows`,`widths_cm`,`aligns`,`level_col`) / `note` 灰字 / `pagebreak` 分页。
`level_col` 列按「重要/高(红)/中(橙)/低(绿)/提示(灰)」着色。
`info_table`、`table`、顶层字段（`page_size`/`margins_cm`/`footer_pagenum` 等）的完整说明同 `data-to-pdf` 的 references/spec.md；**差异**：docx 额外支持 `pagebreak`（分页），**不支持** `spacer`（留白）——pdf 用 `spacer`，docx 用 `pagebreak`。

## 输出说明
- 生成单个 `.docx`：A4、表头深色反白、斑马纹、可选页脚「第 X 页 / 共 Y 页」。
- 页码字段在 Word 打开后会自动计算显示。

## 注意事项
- 输入 spec、输出 docx 都放 `/workspace`；`/skills` 只读。
- spec 必须是合法 JSON；换行在 JSON 字符串里写单反斜杠 `\n`，加粗用 `<b>…</b>`（与 data-to-pdf 语法一致）。
- 要 PDF 终稿：先出 docx 再用 `office-convert` 转 PDF，或直接用 `data-to-pdf`。
