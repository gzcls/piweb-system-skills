# PPTX 编辑指南

优先复制原始模板再编辑，保留母版、背景、页码、备注和可编辑对象。只有结构编辑能力不足时，才解包后针对 XML 做最小修改。

## 流程

1. 使用宿主提供的 PowerPoint/WPS/LibreOffice 渲染能力查看全稿；本包不提供 `thumbnail.py`，没有渲染器时须在报告中标记“未执行视觉质检”。
2. 使用 `scripts/office/unpack.py` 解包到技能目录外的临时目录。
3. 只修改目标页和必要关系，清理未引用对象。
4. 使用 `scripts/office/clean.py --dry-run` 预览清理范围。
5. 使用 `scripts/office/pack.py <unpacked-dir> <output.pptx>` 打包；脚本默认在发布输出前执行 ZIP/XML 结构校验。需要处理故意不完整的测试夹具时才使用 `--no-validate`，不要把 `--original` 当作无损编辑或校验选项。
   也可以单独运行 `python scripts/office/validate.py <output.pptx>` 复核已打包文件。
6. 重新渲染并逐页检查。

## 组件形状

卡片、指标块、信息面板和图片容器遵循主题配置的圆角 Token：`0.104–0.167` 英寸，默认 `0.125` 英寸；同组保持一致。标题栏、页脚、全宽色带、表格网格和贴边强调条使用矩形，避免用矩形覆盖圆角组件。

## 边界

不要把整页正文栅格化为图片；不要删除用户要求保留的备注、数据或对象；不要把 XML 解析通过描述为 PowerPoint/WPS 实测通过。
