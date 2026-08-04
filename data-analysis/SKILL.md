---
name: data-analysis
description: "分析 CSV/Excel 数据并出图。对表格数据做统计摘要、分组聚合，并用 matplotlib 生成柱状图/折线图/饼图/直方图/散点图（支持中文）。
  用于：分析表格/CSV/Excel、算统计指标、分组汇总求和平均、把数据画成图表、做数据可视化。
  操作：分析数据、统计摘要、分组汇总、求和平均、数据画图、生成图表、柱状图折线图饼图、可视化。
  关键词：数据分析, 统计, 分组聚合, 汇总, 求和, 平均, 图表, 可视化, 柱状图, 折线图, 饼图, 直方图, CSV, Excel, pandas, matplotlib, chart, analysis。"
required_tools:
  - bash
recommended_tools:
  - read
  - write
---
# data-analysis

对 CSV / Excel 做**统计摘要、分组聚合**，并用 matplotlib **出图**。引擎 `pandas` + `matplotlib`，中文已配字体。

## 适用场景
- 「分析下这个表格 / 算一下统计指标 / 看看缺失值」
- 「按地区汇总销售额 / 分组求和、求平均」
- 「把这组数据画成柱状图 / 折线图 / 饼图」

## 依赖
- `pandas`、`matplotlib`、`openpyxl`（读 xlsx）。沙箱镜像已预装。中文字体用镜像的 `fonts-noto-cjk`。

## 使用方法
```bash
# 1) 统计摘要（行列数、字段类型、缺失值、描述统计）→ Markdown
python3 /skills/data-analysis/scripts/analyze.py /workspace/data.csv --summary -o /workspace/summary.md

# 2) 分组聚合 → CSV
python3 /skills/data-analysis/scripts/analyze.py /workspace/data.xlsx --groupby 地区 --agg sum --y 销售额 -o /workspace/g.csv

# 3) 出图 → PNG
python3 /skills/data-analysis/scripts/analyze.py /workspace/data.csv --chart bar --x 月份 --y 销售额,成本 -o /workspace/chart.png
```

### 参数表
| 参数 | 说明 |
|------|------|
| `src` | CSV 或 Excel（按扩展名自动识别） |
| `-o, --out` | **必填**。输出路径：摘要用 `.md`、聚合用 `.csv`、图用 `.png` |
| `--summary` | 输出统计摘要 |
| `--groupby` / `--agg` | 分组列 / 聚合方式(sum/mean/count/max/min)；`count` 统计非空行数（含非数值列），其余聚合仅作用于数值列 |
| `--chart` | 图类型 bar/line/pie/hist/scatter |
| `--x` / `--y` | X(分类)列 / Y 列(逗号分隔) |
| `--title` | 图表标题 |

## 输出说明
- `--summary`：Markdown，含字段类型表 + 数值列描述统计。
- `--groupby`：聚合结果 CSV（UTF-8 BOM，Excel 可直接打开）。
- `--chart`：PNG 图（150 DPI）。

## 注意事项
- 输入/输出都放 `/workspace`；`/skills` 只读。
- `--summary` / `--groupby` / `--chart` 三选一，都不给会报错。
- 缺省行为：bar/line 不给 `--x` 用行索引作分类轴；pie 不给 `--y` 取第一个数值列。
- 出的图想进 Excel/PPT/报告，可配合 `data-to-xlsx` / `data-to-pptx` / `data-to-pdf`。
- 超大数据吃内存（512MB 上限）——必要时先筛列/抽样，或调高 `AP_SANDBOX_MEMORY`。
