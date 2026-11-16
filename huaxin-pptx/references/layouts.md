# 版式配方与 spec 参考

本文所有坐标单位均为英寸、字号单位均为 pt，均基于运行画布 10 × 5.625 英寸。所有数值，包括配方常量（`recipes` 段：结论条高度、胶囊尺寸、图表侧栏、表格行高、各配方最小高度等），都来自 `scripts/office/huaxin_theme_config.json`，修改时只改那个文件。
各处的字数上限都由“文本框宽度 × 72 ÷ 字号”**计算得出**，下表中的数字是按当前主题算出的结果。

## 1. 构建命令

```bash
python scripts/hx_pptx.py example-spec > ../work/spec.json      # 完整示例
python scripts/hx_pptx.py validate-spec --spec ../work/spec.json # 只读结构预检，VALID/退出0 或 SPEC_ERROR/退出2
python scripts/hx_pptx.py build --spec ../work/spec.json --output ../output/deck.pptx [--force]
python scripts/hx_pptx.py slots --slide 9                        # 查看模板第 9 页可替换的槽位
```

构建结果输出为 JSON：`status=BUILT` 表示已生成，`warnings` 中的每一项都要处理；`status=SPEC_ERROR` 表示违反硬规则，需按提示压缩、重组或拆页。

## 2. spec 结构

```json
{
  "title": "文档属性标题（可选）",
  "font": "可选，覆盖默认字体 Microsoft YaHei（作用于全部页面与配方）",
  "complete": true,
  "slides": [ { "type": "cover", "...": "..." } ]
}
```

`complete=false` 用于局部续改或只生成部分页面，此时不检查是否以首页开始、以答谢页结束。

## 3. 页面类型

| type | 来源 | 字段（* 为必填） | 规则 |
| --- | --- | --- | --- |
| `cover` | 模板第 1 页 | `title*`（≤2 行；每行 ≤21 字用 28.5 pt，≤30 字用 19.5 pt），`subtitle`（单行 ≤41 字），`presenter`，`date`，`notes` | 删除模板原有文字，右下角合并为一行 `汇报人｜汇报日期`；缺失的字段保留占位文字，不自动填写今天的日期 |
| `toc` | 模板第 2 页 | `items*`：2–6 项，每项为字符串或 `{title(≤16字), desc(≤38字，两行)}`；`label`（默认“目录”） | 不超过 3 项时自动居中为单栏；未用到的编号组会被删除 |
| `section` | 模板第 4 页 | `number`（如 `"02"`），`title*`（1–2 行，每行 ≤10 字，见 `recipes.section_title_chars`），`label`（默认 `PART TWO`） | 删除标题栏中的占位标题，只保留左上章节标签 |
| `content` | 正文版式（与模板第 5 页相同） | `title*`（≤22 字用 22.5 pt，≤25 字用 19.5 pt），`number`（章节标签），`conclusion`（≤34 字），`layout`（1 个配方，或由最多 2 个配方组成的数组），`source`，`notes` | 版式块自上而下叠放，每块之间间隔 0.18 |
| `template` | 模板任意页 | `source*`（页码），`title`，`number`，`texts`（`{"槽位号": "新文本"}`），`delete`（槽位号数组），`charts`，`tables`，`source_note`，`notes` | 先运行 `slots` 查看槽位；含图表的模板页只能复用一次；低于 11 pt 的替换文字会被提升到 11 pt |
| `closing` | 模板第 30 页 | `title`（默认“感谢聆听”），`subtitle` | 删除 “Thank you” |

`template` 页中，`charts` 的每一项为 `{index, categories, series:[{name, values}], chart_title, highlight, x_values}`（散点图槽位传 `x_values` 时保留散点图，仅传文本 `categories` 时自动转为同位置折线图），`tables` 的每一项为 `{index, rows:[[...]]}`，`rows` 中填 `null` 表示保留原单元格内容。

## 4. 正文配方（`content.layout`）

每个版式块必须以 `type` 选择配方，不能使用 `recipe`。图表的 `type` 固定为 `chart`，具体图种放在 `kind`。例如：

```json
{"type":"chart","kind":"doughnut","unit":"%","period":"示例期间","categories":["A","B"],"series":[{"name":"构成","values":[60,40]}]}
```

叠放时 `layout` 是对象数组，例如 `[{"type":"kpis","items":[{"value":"60","label":"A"},{"value":"40","label":"B"}]},{"type":"bullets","items":["示例结论"]}]`。`validate-spec` 检查这些输入结构并报告页码和字段路径；它不证明页面容量、模板槽位或视觉质量通过。

正文区：x 0.50，y 0.84，w 9.00，底边 5.02；如果设置了 `conclusion`，结论条高 0.50，其下的内容从 y 1.52 开始。

| 配方 | 字段 | 布局与样式 | 硬限制 |
| --- | --- | --- | --- |
| `cards` | `items`：`{title, body, accent}` | 2–4 列等宽圆角卡片（`HX_CARD_n`）；卡片标题 13.5 pt 蓝紫，正文 11.25 pt；强调卡为蓝紫底白字 | 2–4 张；强调卡最多 1 张；正文超过 45 字会给出警告 |
| `kpis` | `items`：`{value, unit, label, note, accent}` | 指标块（`HX_TILE_n`），数字 25.5 pt，单位 13.5 pt；强调指标用红色 | 2–4 个；强调指标最多 1 个 |
| `bullets` | `items`：字符串或 `{title, body, accent}` | 左侧编号胶囊（`HX_PILL_n`）+ 右侧文字，行与行之间用细分隔线 | 1–5 条 |
| `process` | `steps`：`{label, title, body, accent}` | 横向连线 + 节点胶囊 + 步骤面板（`HX_STEP_n`） | 3–6 步；强调节点最多 1 个 |
| `chart` | `kind`（column/bar/stacked_column/stacked_bar/line/pie/doughnut），`categories`，`series`，`unit`，`period`，`highlight`（单系列标红分类序号），`number_format`（可选，默认自动推断），`points`（≤2），`height` | 原生可编辑图表 `HX_CHART_1`，字号 10 pt；条形图默认自上而下排列分类；有 `points` 时左侧图表宽 5.85、右侧结论面板宽 2.97（`HX_SUPPORTING_CONCLUSION_n`）；单系列图直接显示数据标签 | 饼图、环形图只能 1 个系列且不超过 5 类；可用高度不足 1.6 时报错 |
| `table` | `columns`，`rows`，`col_widths`（比例），`emphasis_rows`，`height` | 原生表格 `HX_TABLE_1`：表头蓝紫底白字，隔行 `#F3F5F9`，首列左对齐，其余列居中 | 不超过 7 列；行高小于 0.30 时报错 |
| `two_column` | `left`、`right`：`{title, bullets, body}` | 两个面板（`HX_PANEL_1/2`），宽度各 4.41 | 两侧要点合计不超过 5 条 |
| `image_text` | `image`（本地路径），`image_side`（left/right），`ratio`（16:9 或 4:3），`title`，`bullets`，`body` | 图片裁切填充到圆角框（`HX_IMAGE_FRAME_1`），另一侧为面板 | 要点不超过 5 条；分辨率低于约 110 px/英寸时给出警告 |

叠放规则：最后一个块可以向下拉伸；`chart` 或 `table` 放在上方时，必须指定 `height`。常用组合有 `kpis + bullets`、`kpis + chart(height)`、`cards + table`。

### 栅格坐标速查

| 栏 | 列宽 | 各列 x |
| --- | --- | --- |
| 1 | 9.00 | 0.50 |
| 2（6/6） | 4.41 | 0.50 / 5.09 |
| 3（4/4/4） | 2.88 | 0.50 / 3.56 / 6.62 |
| 4（3/3/3/3） | 2.115 | 0.50 / 2.795 / 5.09 / 7.385 |
| 图表 + 结论 | 5.85 + 2.97 | 0.50 / 6.53 |

间距节奏为 0.06 / 0.12 / 0.18 / 0.24 / 0.36；卡片内边距 0.18；卡片间距 0.18。

### 字数与容量参考（Microsoft YaHei 11.25 pt，按字宽估算；最终以渲染结果为准）

| 位置 | 建议上限 |
| --- | --- |
| 正文标题 | 22 字（最多 25 字） |
| 结论条 | 34 字 |
| 四栏卡片正文 | 每行约 11 字，建议不超过 4 行 |
| 三栏卡片正文 | 每行约 16 字，建议不超过 3 行 |
| 双栏面板要点 | 每行约 24 字，每条不超过两行 |
| bullets 单条正文 | 每行约 53 字，不超过两行 |

## 5. 库函数（配方无法表达的页面）

```python
import sys; sys.path.insert(0, r"<技能目录>/scripts"); import hx_pptx as hx
prs, report = hx.build_deck(spec)                 # 先生成结构页和常规页
slide = hx.add_body_slide(prs, "观点标题", number="03")
cols = hx.columns(3)                              # [(x, w), ...]
hx.add_card(slide, cols[0][0], 0.84, cols[0][1], 1.6, "卡片标题", "正文", name="HX_CARD_1")
hx.add_pill(slide, 0.5, 2.6, 0.6, 0.28, "01", name="HX_PILL_1")
hx.add_text(slide, 0.5, 3.0, 9.0, 0.6, "说明文字", name="HX_TEXT_1")
hx.add_source(slide, "资料来源：……")
hx.save(prs, "../output/deck.pptx")
```

注意：`add_body_slide` 生成的页面会追加到末尾。如果需要插在中间，就在 spec 中按顺序放一个只有标题的 `content` 页，然后取 `prs.slides[i]` 在上面补画。
其余函数：`add_chart_panel(slide, x, y, w, h, block, index=1)` 在任意分栏或双图对比槽位绘制带单位行与完整主题样式的原生图表（互不重叠的左右半宽双子图视为 1 组对比图）；`add_rounded` 和 `set_corner_radius` 自动换算圆角；`add_image_fill` 裁切填充；`add_rect` 绘制结构矩形；`fill_text` 写入多段落文字；`set_notes` 写演讲备注；`check_title`、`capacity` 计算字数上限；`with hx.use_font("字体名"):` 临时更换补画内容的字体。

## 6. 对象命名规范（检查脚本依赖这些名称）

| 名称包含 | 含义 | 检查 |
| --- | --- | --- |
| `CARD` / `TILE` / `PANEL` | 主要卡片 | 每页 ≤4；必须是圆角；正文汉字 ≤45（第一段视为标题，不计入） |
| `PILL` / `TAG` / `BADGE` | 胶囊标签 | 半径 = min(h/2, 0.167)，不计入卡片数 |
| `STEP` | 流程步骤面板 | 不计入卡片数，检查圆角 |
| `SUPPORTING_CONCLUSION` | 辅助结论 | 每页 ≤2 |
| `SOURCE` / `FOOTNOTE` / `CAPTION` / `LABEL` / `UNIT` / `NOTE` / `METADATA` | 注释类文字 | 字号下限 8 pt（其余文字为 11 pt） |
| `HX_BODY_TITLE`、`HX_SECTION_TAB`、`HX_COVER_*` | 页面框架元素 | 不计入字号层级（每页 ≤4 级） |

没有语义名称、但带文字的圆角形状，会被视为“未命名卡片”并给出警告。
