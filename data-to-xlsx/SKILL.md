---
name: data-to-xlsx
description: "生成 Excel 表格（.xlsx，支持中文，无需 MS Office）。把数据写成 JSON 或直接给 CSV，一键生成带表头样式、边框、冻结首行、自动筛选、数字格式、可选图表的 Excel；支持多工作表。
  用于：导出报表/台账/明细表/汇总表、CSV 转 Excel、把数据做成带图表的 xlsx。
  操作：生成Excel、导出xlsx、做报表台账、CSV转Excel、数据转Excel、表格加图表。
  关键词：Excel, xlsx, 生成Excel, 导出Excel, 报表, 台账, 明细表, 汇总表, CSV转Excel, 图表, openpyxl, spreadsheet。"
required_tools:
  - bash
recommended_tools:
  - read
  - write
---
# data-to-xlsx

生成 **Excel（.xlsx）**。两种输入：① JSON spec（多 sheet + 样式 + 图表）；② 直接给 **CSV** 快速转。
引擎 `openpyxl`，**不需要 MS Office**。

## 适用场景
- 用户说「导出 Excel / 做个报表、台账、明细表 / 把这些数据做成 xlsx / CSV 转 Excel」
- 需要带表头样式、冻结首行、筛选、千分位、图表的表格

## 依赖
- Python 库 `openpyxl`（沙箱镜像已预装；`pandas` 也已预装，可配合 data-analysis）。

## 使用方法
### A. CSV 一键转 Excel
```bash
python3 /skills/data-to-xlsx/scripts/xlsx_build.py --from-csv /workspace/data.csv -o /workspace/out.xlsx --sheet 数据
# 分号 CSV / TSV：用 --delimiter 指定分隔符（默认逗号）
python3 /skills/data-to-xlsx/scripts/xlsx_build.py --from-csv /workspace/data.csv -o /workspace/out.xlsx --delimiter ";"
```
### B. JSON spec（多 sheet / 图表）
```bash
python3 /skills/data-to-xlsx/scripts/xlsx_build.py /workspace/spec.json -o /workspace/out.xlsx
```

### spec 示例
```json
{
  "sheets": [
    {
      "name": "销售",
      "headers": ["月份", "销售额", "成本"],
      "rows": [["1月", 12000, 8000], ["2月", 15000, 9000], ["3月", 18000, 9500]],
      "widths": [12, 14, 14],
      "number_formats": {"1": "#,##0.00", "2": "#,##0.00"},
      "freeze_header": true,
      "autofilter": true,
      "chart": {"type": "bar", "title": "月度销售", "cats_col": 0, "data_cols": [1, 2], "anchor": "F2"}
    }
  ]
}
```

### 字段
- 顶层 `sheets[]`（也可直接传单个 sheet 对象）。
- sheet：`name`（超 31 字符按 Excel 限制自动截断）/ `headers` / `rows` / `widths`(省略=自动) / `number_formats`(键为列号字符串，0 基，如 `{"1":"#,##0.00"}`) / `freeze_header` / `autofilter` / `chart`。
- chart：`type`(bar/line/pie) / `title` / `cats_col`(分类列,0基) / `data_cols`(数据列,0基数组) / `anchor`(单元格)。**需该 sheet 同时有 `headers` 和 `rows` 才会生成图表**，否则静默跳过。

## 输出说明
- 单个 `.xlsx`：表头深蓝反白、细边框、首行冻结、自动筛选；列宽自动估算或按 `widths`。
- 有 `chart` 时在指定 `anchor` 处插入图表。

## 注意事项
- 输入/输出都放 `/workspace`；`/skills` 只读。
- CSV 模式会把纯数字转成数值（便于求和/画图），但**自动保留前导零/超长数字**（编号、手机号、身份证等按文本）；要完全不转换加 `--no-convert`。编码按 UTF-8（兼容 BOM）。非逗号分隔用 `--delimiter`。
- 复杂统计/分组聚合建议先用 `data-analysis`（pandas）算好，再用本技能落 Excel。
