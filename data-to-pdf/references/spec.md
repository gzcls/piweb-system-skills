# data-to-pdf · spec 完整字段说明

报告内容用一个 JSON 对象描述，传给 `report2pdf.py`。除 `title` 外字段均可省略。

## 顶层字段
| 字段 | 类型 | 默认 | 说明 |
|------|------|------|------|
| `title` | string | — | 报告大标题（居中） |
| `subtitle` | string | — | 副标题（灰色居中） |
| `page_size` | string | `A4` | `A4` / `A3` / `LETTER` |
| `margins_cm` | number[4] | `[2.2,2.0,2.2,2.0]` | 页边距 `[上,下,左,右]`，单位 cm |
| `footer_pagenum` | bool | `true` | 页脚显示「第 X 页 / 共 Y 页」 |
| `info_table` | object | — | 顶部信息表，见下 |
| `sections` | object[] | `[]` | 章节数组，见下 |

## info_table（顶部信息表）
| 字段 | 说明 |
|------|------|
| `headers` | 表头数组，可省略 |
| `rows` | 二维数组，每行一条 |
| `widths_cm` | 各列宽度（cm），可省=均分；总宽超页面会自动等比缩放 |
| `aligns` | 各列对齐：`left`/`center`/`right` |
| `label_cols` | 指定哪些列作浅蓝底「标签列」，如 `[0,2]` |

## sections[]（章节）
| 字段 | 说明 |
|------|------|
| `heading` | 章节标题（深蓝、带下划线），可省 |
| `blocks` | 块数组，按顺序渲染 |

## block 类型
| `type` | 字段 | 说明 |
|--------|------|------|
| `para` | `text` | 段落。支持 `<b>加粗</b>`，`\n` 换行 |
| `subhead` | `text` | 小标题 |
| `list` | `items`, `ordered` | 列表；`ordered:true` 为有序(1.2.3)，否则圆点 |
| `table` | `headers`,`rows`,`widths_cm`,`aligns`,`level_col` | 表格，见下 |
| `note` | `text` | 灰色小字备注 |
| `spacer` | `h` | 纵向留白，`h` 为高度(pt) |

### table 块细节
- `headers`：表头数组（深色背景、白字、自动重复表头跨页）。
- `rows`：二维数组；行从 0 起，偶数行自动斑马纹底色。
- `widths_cm`：各列宽（cm）。省略则均分；总宽超出可用宽度时整体等比缩小。
- `aligns`：各列对齐 `left`/`center`/`right`（默认 left）。
- `level_col`：指定某列为「级别列」，该列单元格文字按值自动着色：
  - `重要`/`高` → 红，`中` → 橙，`低` → 绿，`提示` → 灰。

## 完整示例
见同 skill 的 `scripts/example_spec.json`，可直接复制修改后运行：
```bash
python3 /skills/data-to-pdf/scripts/report2pdf.py /workspace/spec.json -o /workspace/report.pdf
```

## 常见坑
- spec 必须是**合法 JSON**：键名用双引号、注意中文逗号/引号、末项不要多余逗号。
- 字符串里换行写 `\n`（JSON 转义），加粗写 `<b>…</b>`。
- 输入/输出都放 `/workspace`，不要写只读的 `/skills`。
