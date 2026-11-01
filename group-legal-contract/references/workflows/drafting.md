# 工作流 B：合同起草与标准模板生成（`workflows/drafting.md`）

本工作流用于根据用户提供的业务场景、立场和交易要素，基于集团标准合同范本或板块专属核心条款，起草生成符合集团公文规范的 **《合同签署草稿版.docx》** 与 **《合同起草说明版（带批注）.docx》**。

---

## 第一步：收集交易要素与匹配标准底稿

1. **要素识别**：对照 [../../assets/schemas/draft-request.schema.json](../../assets/schemas/draft-request.schema.json) 梳理以下信息：
   - 合同类型与适用业务板块；
   - 我方合同地位（如：买方/卖方、委托方/受托方、出租方/承租方、披露方/接收方）；
   - 合同双方主体信息（全称、统一社会信用代码、地址、联系人）；
   - 核心商务条件（标的物/服务范围、金额与税率、付款节点、履约/租赁期限、交付与验收方式、质保期、违约金比例、争议解决机构）。
   - **未知项留白原则**：用户未提供的具体商务参数，一律使用 `[●]` 或 `` `[填写说明]` `` 占位，**严禁凭空编造主体名称、金额、账号或日期**。
2. **底稿路由**：查阅 [../routing/contract-type-map.md](../routing/contract-type-map.md)：
   - **情形 1（命中 `assets/templates/` 标准商事合同范本）**：读取对应范本作为基础底稿（如双向 NDA、技术开发、货物买卖与设备采购、写字楼园区租赁及物业、股权转让、有限公司章程及决议、DPA 数据委托处理协议、劳动用工与保密竞业全套范本）。
   - **情形 2（命中 `references/kb/sectors/` 板块专属核心条款范本）**：读取对应板块文档中的核心条款范本（如建工施工总包/分包、餐饮特许经营加盟、国际货物买卖中英双语条款、增资扩股与股东协议、OEM/ODM 代工、工厂 EHS 安全协议、医院 GCP/医疗和解、政策性储备粮承储轮换），结合通用商事合同骨架组装完整合同。
   - **情形 3（未命中现成范本）**：依据 [../policies/group-redlines.md](../policies/group-redlines.md) 及对应专类审查清单中的“必备条款与我方立场设计要点”组合起草，并在文首显式标注：`> 【重要提示】本草稿系根据集团审查规则组合起草（非集团固化标准范本），签署前须经法务共享中心重点复核。`

---

## 第二步：立场适配与集团强制红线嵌入

1. **立场版本切换**：
   - 标准范本中带有 `[版本A（偏重甲方/买方/出租方/披露方）]` 与 `[版本B（偏重乙方/卖方/承租方/接收方）]` 分支的，根据我方实际地位保留最有利版本，或保留双版本供业务谈判比选（按用户要求）。
2. **强制红线自检**：对照 [../policies/group-redlines.md](../policies/group-redlines.md) 检查草稿是否已完整嵌入：
   - 明确的质量与终验标准（采购类区分外观签收与质量终验）；
   - “先票后款”及增值税专票虚开/异常连带赔偿条款；
   - 大型企业对中小企业采购时的账期合规（不设置违规“背靠背”条款）；
   - 知识产权归属反转条款（委托开发默认归受托方，我方作为委托方必须显式约定归我方）；
   - 反商业贿赂与阳光廉洁条款（适配 2025 修订《反不正当竞争法》第 8、24 条）；
   - 违约金、维权费用承担（含律师费、保全费、保全保险费）及法定有效司法送达地址确认条款；
   - 唯一且明确的争议解决条款（严禁“或裁或诉”，不动产/建工遵守专属管辖）。
3. **法条核实**：草稿正文或起草批注中引用的所有法律法规条号，必须通过 `cite_law.py` 核实，并使用 `verify_citations.py` 校验通过。

---

## 第三步：生成 Markdown 底稿并渲染双版本 `.docx`

推荐通过统一流水线执行本节，避免人工漏掉请求字段、引用和签署版清洁检查：

自动模板路由读取随 Skill 发布的只读 `assets/template_catalog.json`，并按模板 SHA 校验文件完整性。若模板目录不可用，流水线会明确生成“无标准模板”的通用工作稿并标记人工复核；正式交付前必须使用完整发布包。

```bash
python3 <SKILL_DIR>/scripts/drafting_pipeline.py \
  --request "<输出目录>/work/draft_request.json" \
  --draft-md "<输出目录>/work/draft_with_notes.md" \
  --out-dir "<输出目录>" \
  --require-complete --strict
```

流水线会写入 `work/manifest.json`、`work/lint_report.json` 和
`work/citation_report.json`，并在 `work/` 下保存请求、底稿、schema、路由表和
最终 DOCX 的 SHA-256。`--no-visual` 仅适用于开发调试；签署交付不得跳过视觉门禁。

1. 在任务输出目录 `output/<YYYYMMDD-HHmmss>_<合同简称>/work/` 下保存填充后的 Markdown 底稿 `draft_with_notes.md`。
   - 确保范本中的起草/谈判提示保留标准引用块格式：`> **【起草/谈判批注】** ...` 或 `> **【风险批注】** ...`。
2. **执行法条引用核验**：
   ```bash
   python3 <SKILL_DIR>/scripts/verify_citations.py "<输出目录>/work/draft_with_notes.md" \
     --out "<输出目录>/work/citation_report.json"
   ```
   确认退出码为 0。
3. **渲染版本 1：《签署草稿版.docx》**（剥离所有内部起草批注，`[●]` 黄色高亮，供对外发送或填写签署）：
   ```bash
   python3 <SKILL_DIR>/scripts/md_to_docx.py "<输出目录>/work/draft_with_notes.md" \
     --out "<输出目录>/《合同名》_签署草稿版.docx" \
     --reference "<SKILL_DIR>/assets/docx/reference.docx" \
     --strip-notes \
     --placeholder-highlight
   ```
4. **渲染版本 2：《起草说明版.docx》**（将 `> **【…批注…】**` 转换为挂载在对应条款段落上的 Word 批注，供内部法务与业务谈判参考）：
   ```bash
   python3 <SKILL_DIR>/scripts/md_to_docx.py "<输出目录>/work/draft_with_notes.md" \
     --out "<输出目录>/《合同名》_起草说明版.docx" \
     --reference "<SKILL_DIR>/assets/docx/reference.docx" \
     --notes-as-comments \
     --placeholder-highlight
   ```

---

## 第四步：输出《待业务补充确认事项清单》

在向用户交付上述两份 `.docx` 文件的同时，在回复中结构化列出：
1. **待填写的 `[●]` 商务占位符清单**（按条款顺序列出需要业务部门补充的具体数值、日期、主体信息及附件名称）；
2. **核心立场保护说明**（简述本草稿已为我方重点设防的 3–5 项核心条款及谈判退让底线）；
3. **知识库版本快照声明**。
