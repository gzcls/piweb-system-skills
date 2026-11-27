# 工作流 A：合同合规审核与 Word 修订交付（`workflows/review.md`）

本工作流用于接收待审合同（`.docx` 首选，也支持先转换的 `.doc` / `.wps` / `.rtf` 以及 `.md` / `.txt`），依据集团规则与本地法律法规库完成审查，并交付**带修订痕迹与批注的 `.docx`** 及 **《合同法律审查意见书》`.docx`**。

> **高效审查准则（防上下文膨胀）**：
> 1. 优先运行 `preflight_contract.py` 自动生成**审查上下文包 `work/review_context.md`** 与**压缩合同视图 `work/contract_compact.txt`**；
> 2. 优先读取单页速查表 [../policies/review-quickref.md](../policies/review-quickref.md)，免去逐份通读 5 份策略文件、路由大表与超长功能参数表；
> 3. 补充法条统一使用 `cite_law.py --batch` **一次性批量核取**，严禁逐条多次往返调用；
> 4. 模型只编写一份单一来源台账 **`work/review-draft.json`**（结构见 [../../assets/schemas/review-draft.schema.json](../../assets/schemas/review-draft.schema.json)），由脚本确定性派生 `revisions.json`、`review-result.json` 与 `review_report.md`。

---

## 第一步：建立任务输出目录与生成审查上下文包

1. 在工作区创建本次审查输出目录：
   ```text
   output/<YYYYMMDD-HHmmss>_<合同简称>/
   └── work/
   ```
2. 调用 `preflight_contract.py` 一次性完成转换、结构提取、确定性预检、自动路由与审查上下文包生成（自动打点 `preflight_done`）：
   ```bash
   python3 <SKILL_DIR>/scripts/preflight_contract.py "<输入合同.doc>" \
     --out-dir "<输出目录>"
   ```
   生成的 `work/` 产物包括：
   - `work/base.docx` 与 `work/blocks.json`：含每段稳定锚点 `block_id`（如 `p-0015` 或 `t0-r2-c1-p0`）、自动编号 `list_label`、纯文本 `text` 及既有修订标记；
   - `work/contract_compact.txt`：**压缩合同视图**（100% 保留正文段落、签署/价格/付款/验收表及含金额/期限/权责关键词的单元格，自动折叠纯产品功能描述大表）；
   - `work/review_context.md`：**一站式审查上下文包**（含主体与金额算术核对、自动路由结论、必查清单覆盖、预检高优风险线索 `candidate_findings`、命中专类清单核心要点、以及已定位的常用法条原文节选卡片）；
   - `work/prefetched_laws.json`：预取法条载荷（节选仅用于定位，正式引用仍须核取完整原文）；`work/fold_manifest.json`：压缩视图的逐表保留/折叠记录。
3. 若输入为纯文本或 Markdown（非 `.docx`/`.doc`），先使用 `md_to_docx.py` 渲染为 `work/base_contract.docx`，再对该 `.docx` 执行 `preflight_contract.py`。

---

## 第二步：快速读取上下文与按需展开

1. **核心读取（仅需 3 份文件，可在同一轮并行调用 `read_file`）**：
   - 读取全规则速查表：[../policies/review-quickref.md](../policies/review-quickref.md)（已完整包含 GR-01~GR-35 及 `C0101`/`C0102` 通用双视角策略）；
   - 读取审查上下文包：`<输出目录>/work/review_context.md`（已内联命中专类/板块清单的核查列、甲乙方修改策略、核心条款范本及预取法条定位节选；若因单次 46KB 字节上限提示截断，可按行号分段读取）；
   - 读取压缩合同视图：`<输出目录>/work/contract_compact.txt`。
2. **确认核心五要素**：我方合同地位（甲/乙方及交易角色）、合同类型、所属业务板块、合同金额/期限、特殊关注点（若未提供且可推断，按审慎保护我方立场设定默认假设并在 `assumptions` 中声明）。
3. **按需展开（仅在必要时执行）**：
   - 若压缩视图中某张被折叠的表格疑似包含特殊商务约定，按需展开该表：
     ```bash
     python3 <SKILL_DIR>/scripts/docx_extract.py "<输出目录>/work/blocks.json" --table t33
     ```
   - 若属于复杂跨板块合同且 `review_context.md` 的清单摘要未覆盖特定行业细节，再按需查阅 [../routing/contract-type-map.md](../routing/contract-type-map.md)、[../routing/sector-map.md](../routing/sector-map.md) 或对应 `references/kb/sectors/Sxxxx.md`、[../policies/group-redlines.md](../policies/group-redlines.md)。
4. **阶段打点（可选）**：
   ```bash
   python3 <SKILL_DIR>/scripts/stage_mark.py "<输出目录>" context_read --note "上下文与合同读取完成"
   ```

---

## 第三步：穿透式排查与批量核取补充法条

1. 按照 [../policies/rule-priority.md](../policies/rule-priority.md) 与 [../policies/review-quickref.md](../policies/review-quickref.md) 的优先级，结合 `review_context.md` 中的预检风险线索（`STOP_SERVICE`、`LIABILITY_CAP`、`DATA_TERMS`、`AMOUNT_VARIANTS`、`BLANK_FIELD`、`STALE_DATE`、`OLD_CONTRACT_LAW` 等），对待审合同进行穿透式排查：
   - 主体资格、出资认缴/实缴与对外担保决议核验；
   - 标的、交付、验收（区分外观签收与终验）、质量异议期；
   - 付款节点、发票合规（先票后款、虚开连带赔偿）、中小企业账期合规（禁止违规“背靠背”）；
   - 知识产权归属（委托开发须反转默认规则归我方）、数据安全（790号令双清单、24小时漏洞报告、三年记录与禁停服锁数据）、商业秘密保护；
   - 违约金与赔偿范围（将律师费、保全费、保全保险费纳入违约方承担范围，将人身损害、故意/重大过失、数据/IP/保密违约排除在责任上限之外）；
   - 解除权、争议解决（严禁“或裁或诉”，建工/租赁遵守不动产专属管辖）与司法送达确认条款；
   - 反商业贿赂双罚制廉洁条款。
2. **检查预取法条并一次性批量核取补充法条（反锚定与精准引用铁律）**：
   - 先查看 `work/review_context.md` 第四部分已预取核验的完整法条卡片；已列出的条文可直接引用。
   - **严禁惰性引用“万金油条款”**：禁止仅用《民法典》第470条、第509条等概括性规定替代合同分则专条或专项法规。凡审查点涉及预取卡片未覆盖的具体制度（如替代履行、抵销、特殊行业规范等），**必须**通过 `cite_law.py --batch` 补充核取。
   - **支持条号与 `?关键词` 混合批量核取**：若对某条文的准确条号不确定，可直接在 `--batch` 中写入 `?关键词`（脚本会自动在该法规内匹配前 3 条相关条文并返回完整原文；若数字条号因修法偏移未命中，也会自动返回前后相邻 2 条的 `nearby_articles_hint`）：
     ```bash
     python3 <SKILL_DIR>/scripts/cite_law.py \
       --batch "民法典:581,?不安抗辩;反不正当竞争法:8,15,?大型企业;民事诉讼法:34" \
       --out "<输出目录>/work/extra_laws.json"
     ```
   - 若某法源返回 `out_of_kb`，按 [../policies/citation-policy.md](../policies/citation-policy.md) 规则，仅引用法名与条号并标注 `【库外法源·待人工核实】`，严禁编造条文原文。
3. **阶段打点（可选）**：
   ```bash
   python3 <SKILL_DIR>/scripts/stage_mark.py "<输出目录>" citations_fetched --note "补充法条批量核取完成"
   ```

---

## 第四步：编写单一来源台账 `work/review-draft.json`

遵循 [../policies/revision-style.md](../policies/revision-style.md) 与 [../policies/risk-rating.md](../policies/risk-rating.md)，直接编写一份 `work/review-draft.json`（严格符合 [../../assets/schemas/review-draft.schema.json](../../assets/schemas/review-draft.schema.json)）：

1. **全局字段**：
   - `contract_name`、`our_position`、`contract_type`、`sector`、`subject_and_amount`、`assumptions`；
   - `executive_summary`：由模型撰写 **150–400 字** 的总体风险评价与核心审查结论综述（将直接嵌入《法律审查意见书》第一部分）；
   - `top_risks`：可直接传高风险修订编号数组（如 `["R-05", "R-12", "R-10", "R-03"]`）或 `{id, title, summary}` 对象数组；
   - `pre_signing_checks`：签约前商务与合规核查事项字符串数组；
   - `checklist_coverage`：可选（缺省时自动从 `preflight.json` 路由清单生成）。
2. **`revisions` 数组（唯一问题源）**：
   - 每条修订分配唯一编号 `R-01`, `R-02`, ...；
   - **最小改动原则**：优先使用 `replace`、`insert_text`、`delete` 对具体词句进行精确修改，`anchor.match` 必须直接复制对应 `block_id` 的原文片段（6–40 字、段内唯一）；
   - **缺失条款补齐**：原合同缺失的必备保护性条款，使用 `insert_paragraph_after` 锚定在相关章节末尾段落后插入完整条文（脚本会自动将其同时归入意见书第三部分“建议新增的保护性条款”）；
   - **商务待确认或含既有修订段落**：使用 `comment_only` 并填写 `comment.suggestion`；
   - 可选辅助字段：`title`（简短问题标题，供 Top 风险和新增条款标题使用）、`clause_ref`（缺省时脚本自动从 `blocks.json` 提取）、`original_excerpt`（缺省时脚本自动取 `anchor.match` 或段落前 140 字）。

---

## 第五步：一键派生与确定性收尾流水线（强制闭环）

编写完 `work/review-draft.json` 后，直接调用统一流水线入口：

```bash
python3 <SKILL_DIR>/scripts/contract_review_pipeline.py "<输入合同.doc>" \
  --out-dir "<输出目录>" \
  --draft "<输出目录>/work/review-draft.json"
```

流水线依次自动完成：
1. **派生三份文件（`build_review_outputs.py`）**：从 `review-draft.json` 确定性生成 `work/revisions.json`（符合 [../../assets/schemas/revisions.schema.json](../../assets/schemas/revisions.schema.json)）、`work/review-result.json`（符合 [../../assets/schemas/review-result.schema.json](../../assets/schemas/review-result.schema.json)）以及包含完整四部分正文与附录 A 的 `work/review_report.md`（参照 [../../assets/report/redline-report.md](../../assets/report/redline-report.md)）；
2. **台账与锚点闭环校验（`validate_review_bundle.py`）**：校验 Schema、编号对齐、`block_id` 存在性及行内 `match` 唯一性；
3. **双重法条引用核验（`verify_citations.py`）**：核验 `revisions.json` 与 `review_report.md` 全部法条引用，生成并自动嵌入**附录 B：法条引用自动核验表**；
4. **写入 Word 修订与批注（`docx_redline.py`）**：生成 `《合同名》_审查修订版.docx`；
5. **双向完整性校验与清洁版导出（`docx_verify.py`）**：验证“全部拒绝 = 原文”、“全部接受 = 预期文本”及 OOXML 结构合法，导出 `《合同名》_清洁版.docx`；
6. **意见书渲染与并行版面检查（`md_to_docx.py` + `batch_visual_verify.py`）**：按参考模板渲染 `《合同名》_法律审查意见书.docx`，并对修订版、清洁版和意见书执行 PDF 版面与空白页检查；缺少渲染器时默认阻断；
7. **耗时汇总（`stage_mark.py`）**：自动记录 `pipeline_done` 并生成 `work/time-summary.json` 与 `work/time-summary.md`。

---

## 第六步：单步调试命令与结果交付

如需在不启动完整流水线的情况下单独调试某一环节，可使用以下分步命令：

1. **单独从 `review-draft.json` 派生三份文件**：
   ```bash
   python3 <SKILL_DIR>/scripts/build_review_outputs.py "<输出目录>/work/review-draft.json"
   ```
2. **单独执行台账与锚点校验**：
   ```bash
   python3 <SKILL_DIR>/scripts/validate_review_bundle.py \
     --revisions "<输出目录>/work/revisions.json" \
     --review-result "<输出目录>/work/review-result.json" \
     --blocks "<输出目录>/work/blocks.json" \
     --out "<输出目录>/work/bundle-validation.json"
   ```
3. **单独执行法条引用核验**：
   ```bash
   python3 <SKILL_DIR>/scripts/verify_citations.py "<输出目录>/work/revisions.json" \
     --out "<输出目录>/work/citation_report.json" \
     --md-out "<输出目录>/work/appendix_b.md"
   ```
4. **单独写入 Word 修订并做双向校验**：
   ```bash
   python3 <SKILL_DIR>/scripts/docx_redline.py "<输出目录>/work/base.docx" "<输出目录>/work/revisions.json" \
     --out "<输出目录>/《合同名》_审查修订版.docx" \
     --log "<输出目录>/work/apply_log.json" --track-on

   python3 <SKILL_DIR>/scripts/docx_verify.py "<输出目录>/work/base.docx" \
     "<输出目录>/《合同名》_审查修订版.docx" \
     "<输出目录>/work/revisions.json" \
     --log "<输出目录>/work/apply_log.json" \
     --report "<输出目录>/work/verify_report.json" \
     --clean-out "<输出目录>/《合同名》_清洁版.docx"
   ```
   > **排障注意**：`docx_verify.py` 的 `stdout` 仅输出精简摘要（不含 `checks` 字段），查看 `checks.structure` / `checks.accept_all` 等明细必须读取 `--report` 写入的 `work/verify_report.json`，不要通过 `2>&1 | python3 -c ...` 解析 `stdout`。如需运行多行 Python 检查脚本，请写入工作区 `.py` 文件执行，不要拼接多行 `python3 -c "..."` 字符串。
5. **交付展示**：向用户展示核心审查结论摘要（综合评级、Top 红线风险、修订应用统计与 `work/time-summary.md` 耗时统计），并提供输出目录中的三份 `.docx` 交付文件路径。
