# pptx

华新风格的可编辑 PPTX 制作 Skill。版本 **1.3.0**。
依据用户提供的 `指令.md`、`template-style-guide.md` 和原始风格压缩包整合。

## 包含什么

```text
 pptx/
├── SKILL.md                         # 标准入口：YAML 元数据 + 制作流程
├── agents/openai.yaml              # 可选展示与触发配置
├── assets/
│   ├── huaxin-template-16x9.pptx    # 原始 30 页可编辑模板，字节不变
│   ├── cover-background.jpeg       # 首页 / 答谢尾页
│   └── body-background.jpeg        # 正文页
├── references/
│   ├── style-guide.md              # 唯一的完整视觉参数表
│   ├── quality-checklist.md        # 内容、视觉和可编辑性质检
│   ├── source-notes.md             # 来源、冲突处理和实施说明
│   ├── theme-config.schema.json    # 主题和组件 Token schema
│   └── assets-manifest.json        # 路径、尺寸、数量、SHA-256
├── scripts/
│   ├── create_starter.py           # 可选：生成未完成的可编辑起稿
│   ├── office/                     # Office 解包、打包、清理、结构校验和 LibreOffice 渲染
│   │   └── render_visual.py        # 生成 PDF/PNG 视觉证据清单
│   └── validate_skill.py           # 可选：检查包结构、素材和 PPTX
├── tests/
│   ├── acceptance-cases.md          # 回归场景与验收标准
│   ├── package-integrity.py          # 包版本、开发缓存和打包过滤回归
│   ├── validator-regression.py       # 嵌套 Schema 与页面边界回归
│   ├── density-regression.py         # 卡片、项目符号、中文字符、图表和字号层级回归
│   └── validation-report.md         # 本次验证记录
├── requirements.txt               # 可选脚本依赖
└── README.md
```

## 安装与使用

这是通用 Agent Skills 目录包。解压后应直接得到 `pptx/SKILL.md`，不要只复制入口文件而丢失资源。
目标宿主支持目录型技能时，把整个 `pptx` 文件夹放进其技能目录。
目标宿主支持压缩包导入时，按照其入口导入 ZIP；不承诺所有产品都接受同一后缀或入口。
`.skill` 备用包与 ZIP 内容一致，仅适用于明确支持该后缀的导入器。
本包不向用户账号或技能市场自动安装；平台若使用共享技能目录，应在文件更新后按部署流程刷新其持久化 catalog。

### Codex 本地示例

截至 2026-10-17 的官方文档给出用户级目录 `~/.agents/skills/` 和项目级目录 `.agents/skills/`。
例如用户级安装后的路径为：

```text
~/.agents/skills/pptx/SKILL.md
```

在支持的 Codex CLI / IDE 中可用 `$pptx` 明确调用。
这些示例不是 Web 模板库安装指令；不同宿主的实际入口以其文档为准。

```text
使用 $pptx，将我提供的通知制作成 6 页会议讨论 PPT，
保留审批节点和关键要求，使用华新原始背景，输出可编辑 PPTX。
```

```text
使用 pptx 改版这份演示文稿，保留数据和备注，
统一标题栏、字体与卡片样式，不修改其他已确认内容。
```

## 核心规则

默认 16:9；优先编辑原模板副本；首页和尾页复用首页背景，正文复用正文背景。
正文标题在顶部渐变栏内，白色粗体，避让已有品牌元素，不新增重复页码或 Logo。
首页右下角保留一行可编辑的汇报人和日期标签，未提供时保留字段，不擅自填入当前日期。
每页一个核心结论，使用原生文本、形状和图表；全稿必须完成内容核对与逐页视觉检查。
设计基准为 13.333 × 7.5 英寸，原模板运行画布为 10 × 5.625 英寸，运行对象统一按 0.75 缩放。
普通正文首排绑定主题 `safe_area.top`，当前默认从 `y=0.84` 英寸开始；正文运行字号不低于 11 pt，图表文字不低于 10 pt，脚注不低于 8 pt。
采用 12 栏栅格和 0.08/0.16/0.24/0.32/0.48 英寸设计间距节奏，内容拥挤时拆页或重组。
卡片、指标块、信息面板和图片容器使用 `ROUNDED_RECTANGLE`，统一圆角 `10–16 px`；
标题栏、页脚、全宽色带、表格网格和贴边强调条使用 `RECTANGLE`。

注意：原模板为 **10 × 5.625 英寸**，完整视觉规范以 **13.333 × 7.5 英寸** 为设计基准。
二者都是 16:9，但数值坐标及字号不能直接混用，详见制作入口与来源说明。
`scripts/office/huaxin_theme_config.json` 以运行画布为唯一几何来源；起稿脚本在
`--canvas design` 时按 `runtime_scale.x/y` 分轴反向换算坐标和字号，避免把两套单位混用。

## 可选脚本

脚本不负责调用模型，也不自动检索、发布、安装或上传文件。
只阅读技能或使用宿主自带 PPT 编辑能力时，无需安装这些依赖。
执行辅助脚本时需 Python 3.10+，并在有权限的 Python 环境中安装依赖：

```bash
python -m pip install -r requirements.txt
python -B scripts/validate_skill.py
```

`package.json` 仅声明可选的 PptxGenJS 依赖；`node_modules` 不随技能包发布。
需要使用 PptxGenJS 时，在技能目录外的工作环境按 `package-lock.json` 执行 `npm ci`。

从其他工作目录运行也可以；脚本通过自身位置定位模板，不依赖当前目录。

生成**未完成的起稿**（输出必须放在技能目录外）：

```bash
python scripts/create_starter.py --output ../output/工作汇报-起稿.pptx \
  --title "集团信息条线工作汇报" --subtitle "重点工作与下一步计划" \
  --body-title "围绕重点事项形成协同工作安排" \
  --body-title "明确责任分工与后续推进计划" \
  --presenter "集团信息部" --date "2026年9月"
```

需要验证主题卡片、页脚和正文首排时，可显式生成组件夹具；该夹具仍是测试稿，不是业务成品：

```bash
python -B scripts/create_starter.py --component-fixture \
  --output ../output/工作汇报-组件夹具.pptx \
  --title "集团信息条线工作汇报" --body-title "围绕重点事项形成协同工作安排"
```

省略 `--presenter` 或 `--date` 时保留相应字段。默认以设计基准尺寸新建；
`--canvas original` 使用原模板尺寸，但这仍是背景起稿，**不是原模板无损编辑**。
默认不会覆盖同名输出，确需替换普通输出文件时提供 `--force`。
若指定字体在宿主不可用，需选择已安装的字体并重新检查；本包不提供字体文件。

对生成稿执行基础结构检查：

```bash
python -B scripts/validate_skill.py --pptx ../output/工作汇报-起稿.pptx \
  --json ../output/结构检查.json
```

本机有 LibreOffice 时，用渲染脚本生成可复核的 PDF/PNG 证据。输出和证据 JSON 必须放在技能目录外；只有人工逐页检查 PNG 后才传入 `--review-status pass`：

```bash
python -B scripts/office/render_visual.py \
  --pptx ../output/工作汇报-组件夹具.pptx \
  --output-dir ../output/render \
  --json ../output/视觉证据.json \
  --review-status pass --reviewed-pages 1,2,3

python -B scripts/validate_skill.py \
  --pptx ../output/工作汇报-组件夹具.pptx \
  --visual-evidence ../output/视觉证据.json \
  --json ../output/结构与视觉检查.json
```

`validate_skill.py` 依赖 PyYAML 和 jsonschema；起稿脚本还需要 python-pptx、Pillow。
其中 `validator-regression.py` 用于确认嵌套 Schema 错误、正文区与安全区绑定、页面矩形越界和页脚顶部安全线不会被静默放过；`density-regression.py` 用小型 PPTX fixture 验证密度上限、辅助结论、圆角几何、未命名圆角组件和字号层级；`render_visual.py` 负责 LibreOffice PDF/PNG 渲染与带 SHA-256 的视觉证据清单。
验证器与回归测试会主动关闭 Python 字节码写入；建议保留 `-B`，确保只读检查不会在技能目录生成 `__pycache__` 或 `.pyc` 文件。
它会检查元数据、相对引用、必要文件、资产校验值、图片尺寸、PPTX 基本结构以及显著离页对象。
没有提供 `--visual-evidence` 时，它不判断文案是否真实、文字是否被裁切、配色是否美观；有视觉证据时，它只验证渲染文件、哈希、页数和人工复核标记，不能替代 PowerPoint/WPS 实测。

编辑已有 Office 包时，`scripts/office/pack.py <unpacked-dir> <output.pptx>` 默认会在发布前执行 ZIP/XML 校验；`--no-validate` 仅用于故意不完整的测试夹具。需要单独复核时运行 `python scripts/office/validate.py <output.pptx>`。

## 本次验证

当前版本的确定性检查与两种画布的结构验证结果见 `tests/validation-report.md`。
详见 `tests/validation-report.md`。没有宣称已在所有宿主安装或通过 PowerPoint/WPS 实测。

## 版本与来源

`references/source-notes.md` 说明所有来源、字体换算与有实质影响的合并规则。
原始模板与两张背景保持字节不变；资源清单可用于复核。
未给原素材重新授予开源许可；再分发与品牌使用仍需符合原有授权。

格式参考（仅用于包封装）：
- Agent Skills Specification：`https://agentskills.io/specification`
- OpenAI Build skills：`https://learn.chatgpt.com/docs/build-skills`
