# 改版与续改已有 PPT

## 1. 选择路径

| 情况 | 做法 |
| --- | --- |
| 用户内容可以重新组织成华新版式 | 先提取内容（文字、数据、备注），写成 spec，用 `hx_pptx.py build` **重建**；这是最稳妥的统一风格方式 |
| 只需在华新稿上局部修改 | 用 python-pptx 打开副本，只修改目标页；新增对象用 `hx_pptx` 的库函数（见 `references/layouts.md` 第 5 节），对象名使用 `HX_*` |
| 必须保留原文件中的动画、复杂图表或嵌入对象 | 解包后对 XML 做最小修改（见第 3 节），不要整页重建 |

开始之前先说明所选路径可能丢失什么（例如重建会丢失原有动画），并征得用户同意或换用更稳妥的路径。

## 2. 改版检查要点
- 先找出重复、过密、层级混乱或质量低的页面，再按页表重构。
- 保留用户要求保留的事实、数据含义、图表数据、备注和品牌元素；局部修改不得改动其他页面。
- 读取原稿内容：`python -c "from pptx import Presentation; ..."` 逐页导出文本、表格和备注，或使用宿主自带的文档解析能力。不要只看缩略图就推测内容。

## 3. XML 级最小修改
1. 使用宿主的渲染能力或 `scripts/office/render_visual.py` 查看全稿；没有渲染器时，要在报告中写明“未执行视觉质检”。
2. `python scripts/office/unpack.py <input.pptx> <技能目录外的临时目录>`。
3. 只修改目标幻灯片 XML 及必要的关系文件。删除对象时，要同时删除它在 `.rels` 中不再被引用的关系。
4. `python scripts/office/clean.py <临时目录> --dry-run` 预览，然后去掉 `--dry-run` 执行。注意：该脚本**只删除 `.tmp` 和 `.bak` 临时文件**，不会清理未引用的部件。
5. `python scripts/office/pack.py <临时目录> <output.pptx>` 打包。打包前默认执行 ZIP/XML 校验；`--no-validate` 只用于故意不完整的测试夹具。也可以单独运行 `python scripts/office/validate.py <output.pptx>`。
6. 运行 `python scripts/check_deck.py <output.pptx>`，重新渲染并逐页检查。

## 4. 边界
- 不要把整页正文栅格化为图片；不要删除用户要求保留的备注、数据或对象。
- XML 能解析、python-pptx 能打开，都不等于已在 PowerPoint 或 WPS 中实测通过。
