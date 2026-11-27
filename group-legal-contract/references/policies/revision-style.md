# Word 修订痕迹与批注编写规范（`policies/revision-style.md`）

本规范定义模型如何构造 `revisions.json`（参见 [../../assets/schemas/revisions.schema.json](../../assets/schemas/revisions.schema.json)），以保证 `docx_redline.py` 生成的 `.docx` 符合执业律师与集团法务的审阅习惯。

---

## 一、 修订粒度：最小改动原则（Minimal Diff）

1. **禁止整段删改**：若仅修改某一条款中的期限、比例、主体或增补半句话，**严禁**将整个段落作为 `match` 替换。应仅选取需要改动的目标短语及其紧邻的 3–10 个定位字符作为 `anchor.match`。
   - ✅ **正确示例**：`match: "验收合格后九十日内付款"` ➔ `new_text: "收到乙方开具的合规增值税专用发票且终验合格后三十（30）日内付款"`。
   - ❌ **错误示例**：把整条 300 字的付款条款全部放入 `match` 重写。
2. **`anchor.match` 精确定位要求**：
   - `match` 字符串必须**直接取自** `docx_extract.py` 生成的 `blocks.json` 中对应 `block_id` 的 `text` 字段；
   - 不要把 `contract.txt` 行首的 `[p-0015|第八条]` 标签或 `list_label`（自动编号文本）包含进 `match` 中，因为自动编号并不在段落的 `w:t` 文本流内；
   - 若目标短语在同一段落中出现多次，要么扩大左右上下文使 `match` 在该段内唯一，要么显式指定 `"occurrence": 1`（或 `2`）。

---

## 二、 六种操作类型（`op`）的选用准则

| 操作类型 (`op`) | 适用场景 | 必填字段 |
| :--- | :--- | :--- |
| `replace` | 将原句中的部分词句替换为新措辞（最常用） | `anchor.block_id`, `anchor.match`, `new_text`, `comment` |
| `insert_text` | 在原句某锚点词句之前（`before`）或之后（`after`）补充文字，不删除原文 | `anchor.block_id`, `anchor.match`, `anchor.position` (`after`/`before`), `new_text`, `comment` |
| `delete` | 删除原句中不合理的免责、兜底或冗余短语 | `anchor.block_id`, `anchor.match`, `comment` |
| `insert_paragraph_after` / `insert_paragraph_before` | 原合同缺失必备保护性条款（如缺失反商业贿赂、知识产权归属、维权费用与送达条款），在指定段落前/后插入完整新段落 | `anchor.block_id`, `new_paragraphs`（字符串数组，非空）, `comment`（可选 `style_from`） |
| `delete_paragraph` | 整段属于违法无效或严重损害我方利益且无法通过局部修改挽救的条款 | `anchor.block_id`, `comment` |
| `comment_only` | 需业务部门核实确认商业参数（如价格、工期、技术指标）、需提示背景风险但不宜直接代改文字，或原段落存在复杂域/已有修订 | `anchor.block_id`（可选 `anchor.match`），`comment` |

### 确定性修订引擎行为保证（无需测试或查阅脚本源码）

1. **同段多处修订自动无损合并**：同一 `block_id` 段落内可提交多条不重叠的 `replace` / `delete` / `insert_text` 指令，`docx_redline.py` 内部会先在原段定位区间再按字符偏移**从后往前**应用，且 `replace` 自动计算字符级最小 diff（保留公共前后缀，仅包裹差异字词）。若同段两条指令区间重叠，后一条会自动安全降级为批注。
2. **多段插入严格正序**：`insert_paragraph_after` / `insert_paragraph_before` 的 `new_paragraphs` 数组严格按数组正序依次插入；同一锚点段落若有多条 `insert_paragraph_after`，亦按指令出现顺序依次排列在锚点段之后（默认继承锚点段或 `style_from` 段落样式与缩进）。
3. **禁止手造实验样本与错误校验方式**：
   - `revisions.json` 必须由 `build_review_outputs.py` 从 `review-draft.json` 确定性派生，**严禁**手工构造 `revisions.json` 或创建 `exp` / `test` 临时 `.docx` 试跑探测；
   - `python-docx` 的 `Paragraph.text` 只读取直接子级 `<w:r>` 而忽略 `<w:ins>` / `<w:del>` 修订文本，**严禁**直接用 `Document(revised.docx)` 读取 `p.text` 判断修订结果；如需抽检接受修订后的正文，只能读取 `docx_verify.py --clean-out` 导出的清洁版。

---

## 三、 批注（`comment`）结构与写作标准

每条修订指令必须附带结构化 `comment` 对象，由 `docx_redline.py` 自动格式化为：
`【R-07｜高风险】问题：…／依据：…／底线：…／建议：…`

1. **`id` 编号对齐**：修订指令 `id` 统一采用 `R-01`, `R-02`, `R-03`...，并与《合同法律审查意见书》第二部分“逐条审查修改对照表”及第三部分“新增条款”的编号保持一致。
2. **`risk`（风险等级）**：严格填 `"高"`、`"中"` 或 `"低"`（口径见 [risk-rating.md](risk-rating.md)）。
3. **`issue`（核心风险点分析）**：简明扼要指出原文存在什么法律漏洞、违反哪条红线或对我方造成何种不利后果（控制在 30–120 字）。
4. **`basis`（法规与规则依据数组）**：
   - 填入经 `cite_law.py` 核实的精确法条名称与条号（如 `["《中华人民共和国民法典》第五百八十五条", "《最高人民法院关于适用〈民法典〉合同编通则若干问题的解释》第六十五条"]`）；
   - 库外法源必须带标注（如 `"《中华人民共和国民事诉讼法》第三十四条【库外法源·待人工核实】"`）。
5. **`fallback`（谈判退让底线）**：给出对方拒绝首选修改方案时，我方业务与法务可接受的次优折中方案（如：“若对方不同意30日账期，最长不超过45日，但必须保留‘先票后款’前置条件”）。
6. **`suggestion`（补充说明，可选）**：当 `op` 为 `comment_only` 或发生自动降级时，提供推荐替换措辞或待核实动作。

---

## 四、 合同措辞风格规范

- **数字与期间**：重要金额、期限、比例建议采用“中文大写/汉字 + 阿拉伯数字括注”或严谨表述（如“三十（30）个自然日”、“百分之二十（20%）”、“人民币伍拾万元整（¥500,000.00）”），明确区分“自然日”与“工作日”。
- **主体称谓一致**：严格使用合同首部定义的简称（如“甲方”、“乙方”或“买方”、“卖方”），切勿在同一份合同中混用“甲方”与“发包人/委托方”（除非首部已如此定义）。
- **杜绝模糊词**：将“及时通知”、“合理期限”、“相应责任”改为具体天数与明确违约后果。
