# huaxin-pptx 2.1.0 验证记录

验证日期：2026-11-14。环境：Windows，Python 3.13，python-pptx 1.0.2，Pillow 11.1.0，PyYAML 6.0.3，jsonschema 4.26.0，PyMuPDF 1.28.2，LibreOffice（headless，`C:\Program Files\LibreOffice\program\soffice.com`）。
测试中生成的 PPTX、PDF、PNG 和证据 JSON 都放在技能目录之外，没有打入技能包。

## 确定性检查：通过

| # | 检查 | 结果 |
| --- | --- | --- |
| 1 | `validate_skill.py`：包结构、YAML、相对链接、资源 SHA-256、模板关系与 XML、严格主题 Schema（v3，`additionalProperties: false`）、交叉不变量（含 `spacing.scale` 节奏与表格行高上下限）、颜色与字体是否写进 style-guide | PASS，无 warning |
| 2 | `tests/validator-regression.py`（16 个变异用例）：未知键拦截、嵌套 Schema 错误、`spacing.scale` 节奏、正文区与 `safe_area.top` 绑定、坐标空间、圆角与表格行高上下限顺序、页脚安全线、`colors.light_gray` 必填、`qa.template_markers` 非空 | PASS |
| 3 | `tests/density-regression.py`：按页统计的密度上限、长稿不按全稿总数误报、实际圆角检测（拦截默认 0.25 英寸圆角）、胶囊半径、不含 `HX_` 前缀的语义命名对象（`STEP` / `TAG` 等）圆角检查、同页圆角一致、正文和注释字号下限、字号层级、13.333 英寸旧稿按 `canvas_scale` 缩放检查 | PASS |
| 4 | `tests/builder-regression.py`：示例 spec、`table`、`two_column`、`image_text`（PIL 临时测试图）、`kpis + line chart` 纵向叠放、`template` 第 9 页槽位与第 22 页原生图表/表格数据替换、尾页应用 `slides.closing.title` 宽度、`font` 覆盖与恢复、`create_starter.py` 回退构建，以及 9 类非法 spec 的报错 | PASS |
| 5 | `tests/visual-evidence-regression.py`（免 LibreOffice）：两步确认 `attest()`、缺少备注拦截、页码不全拦截、PPTX 改动后哈希不一致拦截、渲染阶段直接传 `--review-status pass` 拦截 | PASS |
| 6 | `tests/package-integrity.py`：技能名、目录名、版本（`2.1.0`）一致，PptxGenJS 遗留文件已删除，`scripts/office/pack.py` 将解包目录压回 `.pptx` 时过滤 `.git`、`__pycache__`、`node_modules`、`.pyc` 等开发临时文件 | PASS |
| 7 | AST 静态扫描与运行时配置追踪：全部 `.py` 文件无未使用的函数、常量或参数；主题配置 117 个叶子节点无死配置 | 通过 |
| 8 | 运行后技能目录中没有 `__pycache__` 或 `.pyc`；从其他工作目录用绝对路径运行脚本正常；构建时拒绝写入技能目录、未加 `--force` 时拒绝覆盖已有文件 | 通过 |
| 9 | `create_starter.py --component-fixture`（统一 10 × 5.625 英寸画布，复用 `hx_pptx` 辅助函数）`check_deck` 为 PASS；1.x 默认圆角（0.253 英寸）回归夹具仍能被拦截 | 通过 |

## 构建与视觉检查

| 文件 | 内容 | check_deck | 渲染与逐像素对比 |
| --- | --- | --- | --- |
| 示例稿（`hx_pptx.py example-spec`），8 页 | 首页、目录、章节页、指标 + 要点、卡片、柱图 + 辅助结论、流程、尾页 | PASS | LibreOffice 渲染为 PNG 人工逐页查看（`LIBREOFFICE_PASS`）；前 7 页与 2.0.0 逐像素 100% 一致，第 8 页（尾页）因应用 `slides.closing.title`（宽 8.3475 英寸居中）消除原模板 3.49 英寸窄框偏右问题，已目视复核居中正常 |
| 全配方测试稿，13 页 | 另含环形图、折线图、表格、双栏、5 步流程、四卡片，以及模板第 22 页（原生图表与表格）的复用 | PASS | LibreOffice 渲染为 PNG 人工逐页查看（`LIBREOFFICE_PASS`）；前 12 页与 2.0.0 逐像素 100% 一致，第 13 页（尾页）同上应用 `slides.closing.title` 居中，已目视复核 |
| 模板页复用测试，2 页 | 第 9 页按槽位替换文字；第 17 页故意不替换 | 第 2 页报告残留 `ADD TITLE` 和 `according to your` | 符合预期（`qa.template_markers` 单一来源残留检测有效） |

目视检查与重构回归中发现的问题及处理：
- 卡片阴影：LibreOffice 会按 `effectRef` 渲染主题阴影，已把 `effectRef` 设为 `idx=0`。
- 章节页大号编号换行：改为不自动换行，并保留模板原字体。
- 目录不超过 3 项时内容偏左：改为按 `recipes.toc.center_shift` 水平右移居中。
- 尾页标题框偏窄偏右：原模板第 30 页文本框仅宽 3.49 英寸且水平中心偏右（`x=5.16`），已统一应用 `slides.closing.title`（`x=0.825, w=8.3475` 居中）。
- 模板第 2 页隐藏的“02”文本框：构建时删除；模板图表中 9 pt 的文字提升到 10 pt。
- 表格单元格内边距：重构时逐像素对比捕获第 8 页偏移，确认 `python-pptx` 表格使用 `tcPr` 的 `cell.margin_*` 而非 `bodyPr` inset，已恢复为左右 0.08 英寸、上下 0.03 英寸并验证像素级一致。
- 视觉证据：在渲染命令中直接带 `--review-status pass` 会被拒绝；deck 改动后再执行 `--attest` 也会被拒绝（哈希不一致）。

## 边界
- 本记录只证明上述结构检查、LibreOffice 渲染、逐像素对比和人工逐页查看 PNG 的结果。**没有**在 Microsoft PowerPoint 或 WPS 中打开实测，也没有在所有宿主中验证导入。
- `tests/` 目录按决策保留在技能目录内用于开发回归；按宿主渐进披露机制，`SKILL.md` 默认运行时不会加载其中的测试脚本。
- 字数容量按字宽与字号动态估算，最终以渲染结果为准；已有 PPT 中的复杂动画、宏和嵌入对象没有验证。
- `acceptance-cases.md` 中的完整业务场景是后续的回归标准，并非都已完成智能体端到端测试。
