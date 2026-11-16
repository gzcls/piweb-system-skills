# huaxin-pptx

用于制作华新品牌风格可编辑 PPTX 的 Agent Skill，当前版本 **2.1.1**（1.x 版本名为 `pptx`，变更说明见 `references/source-notes.md`）。

## 目录结构

```text
huaxin-pptx/
├── SKILL.md                         # 入口：边界、唯一坐标系、流程、硬约束、诚实性
├── editing.md                       # 改版和续改已有 PPT（重建、局部修改、XML 最小修改）
├── agents/openai.yaml               # 展示与触发配置（$huaxin-pptx）
├── assets/                          # 原始 30 页模板、首页背景、正文背景（按字节原样保留，只读）
├── references/
│   ├── layouts.md                   # spec 格式、版式配方、栅格坐标、字数上限、命名规范、库函数
│   ├── template-map.md              # 原模板逐页索引与复用建议
│   ├── style-guide.md               # 视觉细则（只用运行画布数值）
│   ├── quality-checklist.md         # 交付前质检清单
│   ├── source-notes.md              # 来源、合并处理、各版本变更原因
│   ├── theme-config.schema.json     # 主题配置 Schema
│   └── assets-manifest.json         # 资源路径、尺寸、SHA-256
├── scripts/
│   ├── hx_pptx.py                   # 首选：克隆模板页 + 配方生成（CLI 与库）
│   ├── check_deck.py                # 成品检查：密度、字号下限、圆角、模板残留
│   ├── create_starter.py            # 回退：用背景图生成未完成起稿 / 组件夹具（复用 hx_pptx 辅助函数）
│   ├── validate_skill.py            # 维护用：完整包校验（含成品检查）
│   └── office/                      # 主题配置、LibreOffice 渲染与 attest、解包、打包、校验
├── tests/                           # 回归测试、验收场景与验证记录（不在运行时加载）
└── requirements.txt
```

## 安装

把整个 `huaxin-pptx` 文件夹放进宿主的技能目录，只复制 `SKILL.md` 会丢失资源。
- Codex 本地示例：`~/.agents/skills/huaxin-pptx/SKILL.md`，用 `$huaxin-pptx` 显式调用。
- PIWeb：通过受管的 File Skill 安装接口导入。更名后会产生新的技能标识，旧的 `pptx` 技能需要按部署流程下线或替换。

依赖：Python 3.10+ 和 `python -m pip install -r requirements.txt`；需要视觉证据时，还要安装 LibreOffice。

## 快速开始

```bash
python scripts/hx_pptx.py example-spec > ../work/spec.json      # 修改为实际内容
python scripts/hx_pptx.py validate-spec --spec ../work/spec.json # 结构预检，不写输出文件
python scripts/hx_pptx.py build --spec ../work/spec.json --output ../output/汇报.pptx
python scripts/check_deck.py ../output/汇报.pptx
python scripts/office/render_visual.py --pptx ../output/汇报.pptx --output-dir ../output/render --json ../output/evidence.json
# 逐页查看 ../output/render/汇报/png/*.png，修复问题后：
python scripts/office/render_visual.py --attest ../output/evidence.json --reviewed-pages all --review-notes "逐页检查要点"
python scripts/check_deck.py ../output/汇报.pptx --visual-evidence ../output/evidence.json
```

示例调用：

```text
使用 $huaxin-pptx，将我提供的通知制作成 6 页会议讨论 PPT，保留审批节点和关键要求，输出可编辑 PPTX。
```

## 维护

```bash
python -B scripts/validate_skill.py
python -B tests/validator-regression.py
python -B tests/density-regression.py
python -B tests/builder-regression.py
python -B tests/spec-regression.py
python -B tests/visual-evidence-regression.py
python -B tests/package-integrity.py
```

- 修改颜色、字号、坐标或配方常量时，只改 `scripts/office/huaxin_theme_config.json`（Schema 禁止未知键，校验器检查各字段之间是否一致）；新增或修改颜色时同步更新 `style-guide.md` 的色板，校验器会检查。
- 脚本不会调用模型，也不会自动检索、上传或发布文件；所有输出都必须放在技能目录之外。
- 原始模板和背景的授权保持不变；本包没有为这些素材授予开源许可。

验证结果见 `tests/validation-report.md`。该记录**不代表**已在 PowerPoint 或 WPS 中实测，也不代表已在所有宿主中导入成功。
