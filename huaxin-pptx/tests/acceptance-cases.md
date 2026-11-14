# 验收场景

这些是人工或智能体执行的回归场景，并不代表已经在所有宿主上执行过自动化测试。每次重大规则修改后，按实际任务选择运行，并记录结果。

| 场景 | 示例输入 | 应有行为 | 不应出现 |
| --- | --- | --- | --- |
| 主题新建 | 用华新模板做一份 6 页信息条线工作汇报，内容另附 | 6 页中包含首页和答谢尾页；先给页表，再用 `hx_pptx.py build` 生成；资料中没有的业务数字不编造 | 在 6 页正文之外再加首页和尾页；只回复大纲 |
| 文档转稿 | 根据通知制作会议讨论 PPT | 读取附件实际内容；把长文压缩成观点、流程或卡片 | 只凭文件名推测；整段照抄 |
| 已有 PPT 改版 | 统一为华新风格，保留数据与备注 | 先按 `editing.md` 选择路径并说明可能的损失；保留事实、数据含义和备注 | 静默丢失备注或改动数据 |
| 局部续改 | 只调整第 3 页流程图，其他不变 | 只改指定页面；使用 `complete=false` 或直接编辑该页 | 全稿重排；重复添加尾页 |
| 无汇报人/日期 | 没有提供这两项信息 | 首页保留 `汇报人｜汇报日期`；不单独追问 | 擅自填写当天日期或编造人员姓名 |
| 密度过高 | 两页内塞入大量原文 | 构建报出 SPEC_ERROR 后压缩、重组或拆页，必要时说明页数冲突 | 缩小字号硬塞；绕过构建检查 |
| 标题过长 | 正文标题超过 25 字 | 改写成 22 字以内的观点句 | 两行标题；标题溢出标题栏 |
| 复用信息图 | 想用模板第 9 页的四卡片 | 先运行 `slots --slide 9`，再按槽位替换并删除英文说明 | 残留 `ADD TITLE` 或 `according to your` |
| 复用原生图表 | 用模板第 22 页展示季度数据 | 用 `charts` 和 `tables` 替换数据，图表文字不低于 10 pt | 同一张图表页复用两次 |
| 仅需大纲 | 只要 6 页大纲，不要文件 | 只输出页纲 | 强行生成 PPTX |
| 明确其他模板 | 请用另一品牌模板，不用华新 | 不套用本技能的华新素材和规则 | 覆盖用户指定的模板 |
| 缺渲染器 | 环境只能写 PPTX | 报告“未执行视觉质检”；只说明 `check_deck` 的结果 | 用结构检查冒充目视检查 |
| 视觉证据 | 已安装 LibreOffice | 先 render（RENDERED），逐页看过后 `--attest`（PASS）；`check_deck --visual-evidence` 显示 `LIBREOFFICE_PASS` | 渲染时直接声明通过；deck 改动后仍沿用旧证据 |

## 确定性工具测试
1. `python -B scripts/validate_skill.py` 返回 0，且技能目录中不产生 `__pycache__` 或 `.pyc`。
2. `python -B tests/validator-regression.py`、`tests/density-regression.py`、`tests/builder-regression.py`、`tests/visual-evidence-regression.py`、`tests/package-integrity.py` 全部通过。
3. 从其他工作目录用绝对路径运行脚本，仍能正确定位资源。
4. 把包复制到临时目录后删除背景或改动资源字节，校验应失败，退出码为 1。不要在正式包上做这项测试。
5. `hx_pptx.py build` 和 `create_starter.py` 都拒绝写入技能目录，默认也拒绝覆盖已有文件。
6. `hx_pptx.py build` 的输出中没有重复的 ZIP 条目，并且 `check_deck.py` 的结果为 PASS。
7. `render_visual.py` 在 Windows 上能完成渲染；在 deck 或 PNG 被改动之后执行 `--attest` 必须失败。
