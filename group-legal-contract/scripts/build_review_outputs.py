#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""从单一来源台账 (review-draft.json) 确定性派生三份审核产物。

输入：
  - work/review-draft.json（模型一次性编写的单一来源台账）
  - work/blocks.json（可选，自动补齐 clause_ref 与 original_excerpt）
  - work/preflight.json（可选，自动补齐合同名、路由清单覆盖与标的摘要）

输出：
  1. work/revisions.json（符合 assets/schemas/revisions.schema.json）
  2. work/review-result.json（符合 assets/schemas/review-result.schema.json）
  3. work/review_report.md（符合 assets/report/redline-report.md 四部分 + 附录 A + 附录 B 占位符）
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import time
from datetime import date
from pathlib import Path
from typing import Any

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from lib.kb import entry_by_id, load_index, resolve_law  # noqa: E402
from lib.textio import read_text_auto  # noqa: E402


_ALLOWED_REVISION_KEYS = {
    "id",
    "op",
    "anchor",
    "new_text",
    "new_paragraphs",
    "style_from",
    "comment",
}
# These fields are intentionally retained in the draft but are not part of the
# lower-level revisions.json schema.  Previously the builder silently dropped
# them, making an apparently valid model response differ from the emitted
# review result.  Validate them before deriving any output instead.
_ALLOWED_DRAFT_REVISION_KEYS = _ALLOWED_REVISION_KEYS | {
    "clause_ref",
    "original_excerpt",
    "title",
    "top_risk_summary",
    "missing_clause_title",
    "include_in_missing_clauses",
    "evidence",
}
_ALLOWED_ANCHOR_KEYS = {"block_id", "match", "occurrence", "position"}
_ALLOWED_COMMENT_KEYS = {"risk", "issue", "basis", "fallback", "suggestion"}
_VALID_OPS = {
    "replace",
    "delete",
    "insert_text",
    "insert_paragraph_after",
    "insert_paragraph_before",
    "delete_paragraph",
    "comment_only",
}
_SCHEMA_PATH = SCRIPT_DIR.parent / "assets" / "schemas" / "review-draft.schema.json"

_RISK_BADGE = {"高": "🔴 高", "中": "🟡 中", "低": "🟢 低"}
_OVERALL_DISPLAY = {
    "高": "🔴 重大风险（不建议直接签署，须重大修改）",
    "中": "🟡 中等风险（修改核心条款后可签署）",
    "低": "🟢 低风险（微调文字后即可签署）",
}
_SIGNING_MAP = {
    "高": "不建议直接签署（须重大修改）",
    "中": "修改核心条款后可签署",
    "低": "微调文字后即可签署",
}


def _load_json(path: str | Path) -> dict[str, Any]:
    p = Path(path).resolve()
    text, _ = read_text_auto(p)
    data = json.loads(text)
    if not isinstance(data, dict):
        raise ValueError(f"JSON 根节点必须是对象: {p}")
    return data


def _draft_schema_errors(draft: dict[str, Any]) -> list[str]:
    """Return deterministic JSON-schema errors for the model-authored draft."""
    try:
        import jsonschema  # type: ignore
    except ImportError as exc:  # pragma: no cover - dependency is checked by preflight
        return [f"review-draft.json: 缺少 jsonschema 依赖: {exc}"]
    schema = json.loads(_SCHEMA_PATH.read_text(encoding="utf-8"))
    validator = jsonschema.Draft202012Validator(schema)
    errors: list[str] = []
    for error in sorted(validator.iter_errors(draft), key=lambda item: list(item.path)):
        loc = ".".join(str(part) for part in error.path) or "<root>"
        errors.append(f"review-draft.json.{loc}: {error.message}")
    return errors


def _validate_draft_relations(
    draft: dict[str, Any],
    blocks_map: dict[str, dict[str, Any]],
    preflight: dict[str, Any] | None = None,
) -> None:
    """Validate conditions JSON Schema cannot express safely.

    The model writes one source ledger and all downstream artefacts are derived
    from it.  Any dropped row, dangling Top-risk ID, risk mismatch, or operation
    lacking its payload would otherwise produce a polished but incomplete report.
    Fail before writing *any* derived output.
    """
    errors = _draft_schema_errors(draft)
    routing = (preflight or {}).get("routing") or {}
    if routing.get("routing_confirmation_required"):
        confirmation = draft.get("routing_confirmation") or {}
        if confirmation.get("confirmed") is not True:
            errors.append(
                "自动路由置信度不足：必须在 routing_confirmation.confirmed=true 中记录确认理由"
            )
    try:
        schema_properties = set(
            (json.loads(_SCHEMA_PATH.read_text(encoding="utf-8")).get("properties") or {}).keys()
        )
        unknown_root = sorted(set(draft) - schema_properties)
        if unknown_root:
            errors.append(
                "review-draft.json 存在未支持顶层字段（将被忽略）: " + ", ".join(unknown_root)
            )
    except Exception as exc:
        errors.append(f"review-draft.schema.json 无法读取: {exc}")
    raw_revisions = draft.get("revisions")
    if not isinstance(raw_revisions, list):
        errors.append("review-draft.json.revisions 必须是数组")
        raw_revisions = []

    revision_ids: list[str] = []
    for idx, item in enumerate(raw_revisions):
        prefix = f"revisions[{idx}]"
        if not isinstance(item, dict):
            errors.append(f"{prefix} 必须是对象")
            continue
        unknown = sorted(set(item) - _ALLOWED_DRAFT_REVISION_KEYS)
        if unknown:
            errors.append(f"{prefix} 存在未支持字段（将被丢弃）: {', '.join(unknown)}")
        rid = str(item.get("id") or "").strip()
        if not rid:
            errors.append(f"{prefix}.id 不能为空")
        elif rid in revision_ids:
            errors.append(f"{prefix}.id 重复: {rid}")
        revision_ids.append(rid)
        op = str(item.get("op") or "").strip()
        if op not in _VALID_OPS:
            errors.append(f"{prefix}.op 不支持: {op or '<empty>'}")
        anchor = item.get("anchor")
        if not isinstance(anchor, dict):
            errors.append(f"{prefix}.anchor 必须是对象")
            anchor = {}
        block_id = str(anchor.get("block_id") or "").strip()
        if not block_id:
            errors.append(f"{prefix}.anchor.block_id 不能为空")
        elif blocks_map and block_id not in blocks_map:
            errors.append(f"{prefix}.anchor.block_id 不存在: {block_id}")
        if op in {"replace", "insert_text"}:
            if not isinstance(item.get("new_text"), str) or not item.get("new_text", "").strip():
                errors.append(f"{prefix}: {op} 必须提供非空 new_text")
        elif op in {"insert_paragraph_after", "insert_paragraph_before"}:
            paras = item.get("new_paragraphs")
            if not isinstance(paras, list) or not any(str(p).strip() for p in paras):
                errors.append(f"{prefix}: {op} 必须提供非空 new_paragraphs")
        if op in {"replace", "delete", "insert_text"}:
            match = anchor.get("match")
            if not isinstance(match, str) or not match.strip():
                errors.append(f"{prefix}: {op} 必须提供非空 anchor.match")
        excerpt = item.get("original_excerpt")
        if excerpt is not None:
            excerpt_text = str(excerpt).strip()
            block_text = str((blocks_map.get(block_id) or {}).get("text") or "")
            # A model may abbreviate a long clause with an ellipsis, but the
            # retained core must still be a literal span of the extracted
            # contract paragraph. This prevents fabricated excerpts from
            # becoming the evidence shown in the report.
            core = re.sub(r"^(?:\.{3}|…)+|(?:\.{3}|…)+$", "", excerpt_text).strip()
            if core and block_text and core not in block_text:
                errors.append(f"{prefix}.original_excerpt 不存在于 anchor.block_id 对应原文")
            elif core and not block_text and block_id:
                errors.append(f"{prefix}.original_excerpt 无法在缺失的原文块中核验")
        evidence = item.get("evidence")
        if evidence is not None:
            if not isinstance(evidence, dict):
                errors.append(f"{prefix}.evidence 必须是对象")
            else:
                source_id = str(evidence.get("source_block_id") or block_id).strip()
                quote = str(evidence.get("quote") or "")
                if source_id != block_id:
                    errors.append(f"{prefix}.evidence.source_block_id 必须等于 anchor.block_id")
                source_text = str((blocks_map.get(source_id) or {}).get("text") or "")
                if quote and source_text and quote not in source_text:
                    errors.append(f"{prefix}.evidence.quote 不存在于 source_block_id 原文")
                if quote and not source_text:
                    errors.append(f"{prefix}.evidence.quote 无法在 source_block_id 原文中核验")
                start, end = evidence.get("span_start"), evidence.get("span_end")
                if start is not None or end is not None:
                    if not isinstance(start, int) or not isinstance(end, int) or start < 0 or end < start or end > len(source_text):
                        errors.append(f"{prefix}.evidence span_start/span_end 越界或类型错误")
                    elif quote and source_text[start:end] != quote:
                        errors.append(f"{prefix}.evidence span 与 quote 不一致")
                quoted_hash = str(evidence.get("quote_sha256") or "").strip()
                if quoted_hash and quoted_hash != hashlib.sha256(quote.encode("utf-8")).hexdigest():
                    errors.append(f"{prefix}.evidence.quote_sha256 与 quote 不一致")
        comment = item.get("comment")
        if not isinstance(comment, dict):
            errors.append(f"{prefix}.comment 必须是对象")
        else:
            unknown_comment = sorted(set(comment) - _ALLOWED_COMMENT_KEYS)
            if unknown_comment:
                errors.append(f"{prefix}.comment 存在未支持字段: {', '.join(unknown_comment)}")
            if not str(comment.get("issue") or "").strip():
                errors.append(f"{prefix}.comment.issue 不能为空")
            if comment.get("risk") == "高" and not any(
                str(value).strip() for value in (comment.get("basis") or [])
            ):
                errors.append(f"{prefix}.comment.basis：高风险修订必须提供至少一项证据/法规依据")

    if len(revision_ids) != len(set(revision_ids)):
        # Keep this explicit even though the per-row message above points at
        # the duplicate; it makes the gate easy to find in CI output.
        errors.append("review-draft.json.revisions 存在重复修订编号")
    revision_id_set = {rid for rid in revision_ids if rid}

    raw_top = draft.get("top_risks")
    if raw_top is not None:
        if not isinstance(raw_top, list):
            errors.append("top_risks 必须是数组")
        else:
            for idx, entry in enumerate(raw_top):
                rid = entry.get("id") if isinstance(entry, dict) else entry
                if isinstance(rid, str) and rid and rid not in revision_id_set:
                    errors.append(f"top_risks[{idx}] 引用了不存在的修订编号: {rid}")
                if isinstance(entry, dict):
                    for field in ("id", "title", "summary"):
                        if not str(entry.get(field) or "").strip():
                            errors.append(f"top_risks[{idx}].{field} 不能为空")
                elif not isinstance(entry, str) or not entry.strip():
                    errors.append(f"top_risks[{idx}] 必须是非空修订编号或对象")

    raw_missing = draft.get("missing_clauses")
    if raw_missing is not None:
        if not isinstance(raw_missing, list):
            errors.append("missing_clauses 必须是数组")
        else:
            seen_missing: set[str] = set()
            for idx, entry in enumerate(raw_missing):
                if not isinstance(entry, dict):
                    errors.append(f"missing_clauses[{idx}] 必须是对象")
                    continue
                rid = str(entry.get("id") or "").strip()
                if rid and rid not in revision_id_set:
                    errors.append(f"missing_clauses[{idx}] 引用了不存在的修订编号: {rid}")
                if rid in seen_missing:
                    errors.append(f"missing_clauses[{idx}] 重复编号: {rid}")
                seen_missing.add(rid)
                for field in ("id", "title", "reason", "suggested_clause_text"):
                    if not str(entry.get(field) or "").strip():
                        errors.append(f"missing_clauses[{idx}].{field} 不能为空")

    computed_risk = _compute_overall_risk(
        [item for item in raw_revisions if isinstance(item, dict)]
    )
    stated_risk = str(draft.get("overall_risk") or "").strip()
    if stated_risk and stated_risk != computed_risk:
        errors.append(f"overall_risk={stated_risk} 与修订风险推导值 {computed_risk} 不一致")
    stated_signing = str(draft.get("signing_recommendation") or "").strip()
    if stated_signing and stated_signing != _SIGNING_MAP.get(stated_risk or computed_risk):
        errors.append("signing_recommendation 与 overall_risk 不一致")

    if errors:
        raise ValueError("review-draft.json 预校验失败:\n- " + "\n- ".join(errors))


_LAW_BASIS_RE = re.compile(r"《(?P<law>[^》]+)》(?P<tail>[^；;。\n]*)")
_ARTICLE_LABEL_RE = re.compile(r"第[一二三四五六七八九十百千万零〇两\d]+条(?:之[一二三四五六七八九十百千万零〇两\d]+)?(?:第[一二三四五六七八九十百千万零〇两\d]+款)?")


def _build_law_basis_evidence(basis: list[str]) -> list[dict[str, Any]]:
    """Turn free-form legal basis strings into source-bound audit records."""
    evidence: list[dict[str, Any]] = []
    for raw in basis:
        raw_text = str(raw).strip()
        match = _LAW_BASIS_RE.search(raw_text)
        if not match:
            evidence.append({"citation": raw_text, "evidence_type": "law_citation", "source_status": "manual_confirmation"})
            continue
        law_query = match.group("law").strip()
        labels = _ARTICLE_LABEL_RE.findall(match.group("tail"))
        status, targets = resolve_law(law_query, "")
        if status == "ok" and targets:
            entry = entry_by_id(targets[0].entry_id, "") or {}
            evidence.append({
                "citation": raw_text,
                "law_id": entry.get("id"),
                "law_title": entry.get("title"),
                "article_labels": labels,
                "source_path": entry.get("path"),
                "source_sha256": entry.get("sha256"),
                "evidence_type": "law_citation",
                "source_status": "located",
            })
        else:
            evidence.append({
                "citation": raw_text,
                "law_query": law_query,
                "article_labels": labels,
                "evidence_type": "law_citation",
                "source_status": "manual_confirmation" if status != "ambiguous" else "ambiguous",
            })
    return evidence


def _write_json(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(f".{path.name}.{time.time_ns()}.tmp")
    try:
        tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        tmp.replace(path)
    finally:
        if tmp.exists():
            tmp.unlink()


def _blocks_lookup(blocks_path: Path | None) -> dict[str, dict[str, Any]]:
    if not blocks_path or not blocks_path.is_file():
        return {}
    data = _load_json(blocks_path)
    out: dict[str, dict[str, Any]] = {}
    for b in data.get("blocks") or []:
        if not isinstance(b, dict):
            continue
        bid = str(b.get("block_id") or "")
        if bid:
            out[bid] = b
        seq = str(b.get("seq_id") or "")
        if seq:
            out.setdefault(seq, b)
    return out


def _default_kb_snapshot() -> str:
    try:
        idx = load_index()
        snap = idx.get("kb_snapshot") or str(date.today())
        commit = idx.get("source_commit") or "unknown"
        return f"快照日期 {snap}（提交号 {commit}）"
    except Exception:
        return f"快照日期 {date.today().isoformat()}"


def _infer_title(item: dict[str, Any]) -> str:
    for key in ("title", "missing_clause_title"):
        val = str(item.get(key) or "").strip()
        if val:
            return val
    # 尝试从新增段落开头的【标题】提取
    candidates: list[str] = []
    if item.get("new_paragraphs"):
        candidates.extend(str(p) for p in item["new_paragraphs"] if p)
    if item.get("new_text"):
        candidates.append(str(item["new_text"]))
    comment = item.get("comment") or {}
    if comment.get("suggestion"):
        candidates.append(str(comment["suggestion"]))
    for cand in candidates:
        m = re.search(r"【([^】]{2,28})】", cand)
        if m:
            return m.group(1).strip()
    issue = str(comment.get("issue") or "").strip()
    if issue:
        first_clause = re.split(r"[，。；;、：:]", issue)[0].strip()
        return first_clause[:24] if len(first_clause) > 24 else first_clause
    return f"条款修订项 {item.get('id', '')}".strip()


def _clean_revision_for_schema(item: dict[str, Any]) -> dict[str, Any]:
    cleaned: dict[str, Any] = {}
    for k in _ALLOWED_REVISION_KEYS:
        if k not in item or item[k] is None:
            continue
        if k == "anchor" and isinstance(item[k], dict):
            cleaned["anchor"] = {
                ak: av for ak, av in item[k].items() if ak in _ALLOWED_ANCHOR_KEYS and av is not None
            }
        elif k == "comment" and isinstance(item[k], dict):
            cleaned["comment"] = {
                ck: cv for ck, cv in item[k].items() if ck in _ALLOWED_COMMENT_KEYS and cv is not None
            }
        else:
            cleaned[k] = item[k]
    return cleaned


def _compute_overall_risk(revisions: list[dict[str, Any]]) -> str:
    high_cnt = sum(1 for r in revisions if (r.get("comment") or {}).get("risk") == "高")
    med_cnt = sum(1 for r in revisions if (r.get("comment") or {}).get("risk") == "中")
    if high_cnt >= 1 or med_cnt >= 5:
        return "高"
    if med_cnt >= 1:
        return "中"
    return "低"


def _escape_md_cell(text: str) -> str:
    s = str(text or "").strip()
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    s = s.replace("|", "｜")
    s = re.sub(r"\n+", "<br>", s)
    return s


def _render_report_md(
    draft: dict[str, Any],
    review_result: dict[str, Any],
    preflight: dict[str, Any] | None,
) -> str:
    contract_name = review_result["contract_name"]
    dept = (
        str(draft.get("department_or_sector") or "").strip()
        or str(review_result.get("sector") or "").strip()
        or "未提供（按通用商事合同口径审查）"
    )
    our_pos = review_result["our_position"]
    subject_amount = str(draft.get("subject_and_amount") or "").strip()
    if not subject_amount:
        explicit_total = ((preflight or {}).get("summary") or {}).get("amounts", {}).get("explicit_total")
        if explicit_total is not None:
            subject_amount = f"{review_result['contract_type']}；合同标示总金额约 {explicit_total:g} 元（以最终确认价格表为准）"
        else:
            subject_amount = f"{review_result['contract_type']}（详见合同正文及价格附件）"

    # 汇总引用过的法源名称（不带条号，避免首部概览触发多余的条号核验计数）
    law_names: list[str] = []
    for issue in review_result.get("issues") or []:
        for b in issue.get("legal_basis") or []:
            for m in re.finditer(r"《([^》]+)》", str(b)):
                name = f"《{m.group(1)}》"
                if name not in law_names:
                    law_names.append(name)
    laws_overview = "、".join(law_names[:8]) if law_names else "《中华人民共和国民法典》及相关专项法规"

    chk_titles = [
        f"《{c['checklist_title']}》"
        for c in (review_result.get("checklist_coverage") or [])
        if c.get("checklist_title")
    ]
    chks_overview = "、".join(chk_titles[:6]) if chk_titles else "《集团法务强制红线速查表》《集团法务合同审核要点》《通用商事合同审查总清单》"

    kb_snapshot = review_result["kb_snapshot"]
    review_date = str(draft.get("review_date") or "").strip()
    if not review_date:
        today = date.today()
        review_date = f"{today.year}年{today.month:02d}月{today.day:02d}日"

    assumptions = review_result.get("assumptions") or []
    has_existing_rev = bool(((preflight or {}).get("extraction") or {}).get("has_existing_revisions"))
    if has_existing_rev and not any("修订痕迹" in a for a in assumptions):
        assumptions = list(assumptions) + [
            "预检检测到原文档带有未接受的修订痕迹，已按“当前接受视图”提取审查；涉及原有修订的段落请在定稿前由人工在 Word 中确认接受或拒绝。"
        ]
    assumption_str = "；".join(str(a).rstrip("。；;") for a in assumptions if str(a).strip())
    if assumption_str:
        assumption_str += "。"
    else:
        assumption_str = f"按 {our_pos} 最审慎保护原则执行审查。"

    overall_risk = review_result["overall_risk"]
    overall_label = _OVERALL_DISPLAY.get(overall_risk, overall_risk)
    signing_rec = review_result["signing_recommendation"]
    exec_summary = str(draft.get("executive_summary") or "").strip()

    lines: list[str] = [
        f"# 《{contract_name}》法律合规审查意见书",
        "",
        f"* **委托部门 / 所属板块**：{dept}",
        f"* **我方合同地位**：{our_pos}",
        f"* **合同交易标的与金额**：{subject_amount}",
        "* **审查依据（本地知识库已加载）**：",
        f"  1. 法律法规库：{laws_overview}；",
        f"  2. 审查清单与板块库：{chks_overview}。",
        f"* **知识库快照与版本**：{kb_snapshot}",
        f"* **审查日期**：{review_date}",
        "",
        f"> **审查声明与默认假设**：本意见书由集团法务 AI 审查助手生成，仅供内部法务初审与商务谈判参考，最终定稿须经授权法务复核签批。{assumption_str}",
        "",
        "---",
        "",
        "## 一、 总体风险评估与审查结论（Executive Summary）",
        "",
        f"* **综合风险评级**：{overall_label}",
        f"* **签署结论建议**：{signing_rec}",
    ]
    if exec_summary:
        lines.extend(["", exec_summary])

    top_risks = review_result.get("top_risks") or []
    lines.extend([
        "",
        f"* **核心商业与法律隐患摘要（Top {len(top_risks)} 红线问题）**：",
    ])
    for idx, tr in enumerate(top_risks, 1):
        lines.append(f"  {idx}. **【{tr['title']}（对应 {tr['id']}）】**：{tr['summary']}")

    lines.extend([
        "",
        "---",
        "",
        "## 二、 逐条审查修改对照表（Clause-by-Clause Redline Table）",
        "",
        "| 编号 | 原合同条款 | 风险等级 | 原文摘录 | 风险分析与法规依据 | 建议修改后措辞（Redline） | 谈判退让底线（Fallback） |",
        "| :---: | :---: | :---: | :--- | :--- | :--- | :--- |",
    ])

    for issue in review_result.get("issues") or []:
        rid = _escape_md_cell(issue["id"])
        cref = _escape_md_cell(issue.get("clause_ref") or issue.get("block_id") or "—")
        rbadge = _RISK_BADGE.get(issue["risk"], issue["risk"])
        excerpt = _escape_md_cell(issue.get("original_excerpt") or "—")
        if excerpt != "—" and not (excerpt.startswith("“") and excerpt.endswith("”")):
            excerpt = f"“{excerpt}”"
        basis_list = [str(x).strip() for x in (issue.get("legal_basis") or []) if str(x).strip()]
        analysis_cell = f"**风险点**：{_escape_md_cell(issue['issue_analysis'])}"
        if basis_list:
            analysis_cell += f"<br>**法规依据**：{_escape_md_cell('；'.join(basis_list))}"
        suggested = _escape_md_cell(issue.get("suggested_text") or "")
        fallback = _escape_md_cell(issue.get("fallback") or "可视情况协商")
        lines.append(
            f"| {rid} | {cref} | {rbadge} | {excerpt} | {analysis_cell} | {suggested} | {fallback} |"
        )

    lines.extend([
        "",
        "---",
        "",
        "## 三、 原合同缺失但必须新增的保护性条款（Missing Clauses）",
        "",
    ])
    missing = review_result.get("missing_clauses") or []
    if missing:
        lines.append(
            f"针对原合同未涵盖但我方处于 {our_pos} 地位必须防范的法律漏洞，已在《审查修订版.docx》中新增或建议补充以下保护性条款："
        )
        lines.append("")
        for idx, mc in enumerate(missing, 1):
            basis_tail = ""
            if mc.get("legal_basis"):
                basis_tail = f"（依据：{'；'.join(mc['legal_basis'])}）"
            lines.extend([
                f"### {idx}. 建议新增条款（对应 {mc['id']}）：【{mc['title']}】",
                f"* **增补理由**：{mc['reason']}{basis_tail}",
                "> **建议直接插入以下条文**：  ",
                f"> “{mc['suggested_clause_text']}”",
                "",
            ])
    else:
        lines.append("原合同结构基本覆盖核心模块，相关完善内容均已通过上方逐条修订直接写入。")
        lines.append("")

    lines.extend([
        "---",
        "",
        "## 四、 签约前商务与合规核查事项（Pre-Signing Checklist）",
        "",
        "在正式加盖公章前，请业务部门配合核实以下事项：",
    ])
    for chk in review_result.get("pre_signing_checks") or []:
        lines.append(f"- [ ] {chk}")

    lines.extend([
        "",
        "---",
        "",
        "## 附录 A：适用审查清单覆盖表",
        "",
        "| 清单编号 | 清单名称 | 适用状态 | 核查结果摘要 |",
        "| :---: | :--- | :---: | :--- |",
    ])
    for cov in review_result.get("checklist_coverage") or []:
        cid = _escape_md_cell(cov.get("checklist_id") or "—")
        ctitle = _escape_md_cell(cov.get("checklist_title") or "")
        cstat = _escape_md_cell(cov.get("status") or "已全项核查")
        cnotes = _escape_md_cell(cov.get("notes") or "已覆盖")
        lines.append(f"| {cid} | `{ctitle}` | {cstat} | {cnotes} |")

    lines.extend([
        "",
        "---",
        "",
        "## 附录 B：法条引用自动核验表",
        "",
        "<!-- APPENDIX_B_PLACEHOLDER -->",
        "",
    ])
    return "\n".join(lines)


def build_outputs_from_draft(
    draft_path: str | Path,
    *,
    blocks_path: str | Path | None = None,
    preflight_path: str | Path | None = None,
    revisions_out: str | Path | None = None,
    review_result_out: str | Path | None = None,
    report_md_out: str | Path | None = None,
) -> dict[str, Any]:
    """读取 review-draft.json 并生成 revisions.json、review-result.json、review_report.md。"""
    draft_p = Path(draft_path).resolve()
    work_dir = draft_p.parent
    draft = _load_json(draft_p)

    blocks_p = Path(blocks_path).resolve() if blocks_path else (work_dir / "blocks.json")
    preflight_p = Path(preflight_path).resolve() if preflight_path else (work_dir / "preflight.json")
    revisions_p = Path(revisions_out).resolve() if revisions_out else (work_dir / "revisions.json")
    result_p = Path(review_result_out).resolve() if review_result_out else (work_dir / "review-result.json")
    report_p = Path(report_md_out).resolve() if report_md_out else (work_dir / "review_report.md")

    blocks_map = _blocks_lookup(blocks_p)
    preflight: dict[str, Any] | None = None
    if preflight_p.is_file():
        try:
            preflight = _load_json(preflight_p)
        except Exception:
            preflight = None

    # Gate the single source of truth before deriving any files.  This keeps a
    # failed model response from leaving a new revisions.json beside an older
    # review-result.json and accidentally mixing generations.
    _validate_draft_relations(draft, blocks_map, preflight)
    raw_revisions = draft.get("revisions") or []

    clean_revisions = [_clean_revision_for_schema(item) for item in raw_revisions]
    revisions_obj = {
        "author": str(draft.get("author") or "集团法务AI审查"),
        "revisions": clean_revisions,
    }

    # 构建 issues 与候选 missing_clauses
    issues: list[dict[str, Any]] = []
    auto_missing: list[dict[str, Any]] = []
    rev_by_id: dict[str, dict[str, Any]] = {}

    for item in raw_revisions:
        if not isinstance(item, dict):
            continue
        rid = str(item.get("id") or "")
        rev_by_id[rid] = item
        anchor = item.get("anchor") or {}
        bid = str(anchor.get("block_id") or "")
        block = blocks_map.get(bid) or {}
        comment = item.get("comment") or {}

        clause_ref = (
            str(item.get("clause_ref") or "").strip()
            or str(block.get("clause_ref") or "").strip()
            or str(block.get("list_label") or "").strip()
            or bid
        )
        excerpt = str(item.get("original_excerpt") or "").strip()
        if not excerpt:
            if anchor.get("match"):
                excerpt = str(anchor["match"]).strip()
            elif block.get("text"):
                raw_txt = str(block["text"]).strip()
                excerpt = raw_txt[:140] + ("…" if len(raw_txt) > 140 else "")
            else:
                excerpt = "（新增条款或整段批注）"

        op = str(item.get("op") or "")
        new_paras = [str(p).strip() for p in (item.get("new_paragraphs") or []) if str(p).strip()]
        if item.get("new_text"):
            suggested = str(item["new_text"]).strip()
        elif new_paras:
            suggested = "\n".join(new_paras)
        elif comment.get("suggestion"):
            suggested = str(comment["suggestion"]).strip()
        elif op == "delete":
            suggested = "【建议删除该条款/表述】（删除原文失衡或违规内容）"
        else:
            suggested = str(comment.get("fallback") or "按批注意见修改").strip()

        basis = [str(b).strip() for b in (comment.get("basis") or []) if str(b).strip()]
        fallback = str(comment.get("fallback") or "请业务与法务结合谈判地位评估").strip()
        law_basis_evidence = _build_law_basis_evidence(basis)

        evidence = item.get("evidence") if isinstance(item.get("evidence"), dict) else None
        source_text = str(block.get("text") or "")
        quote = str((evidence or {}).get("quote") or "").strip()
        if not quote and anchor.get("match") and str(anchor.get("match")) in source_text:
            quote = str(anchor.get("match")).strip()
        if not quote and source_text and op not in {"insert_paragraph_after", "insert_paragraph_before"}:
            quote = source_text[: min(len(source_text), 240)]
        span_start = source_text.find(quote) if quote else -1
        evidence_obj = {
            "source_block_id": str((evidence or {}).get("source_block_id") or bid),
            "quote": quote,
            "span_start": span_start if span_start >= 0 else None,
            "span_end": (span_start + len(quote)) if span_start >= 0 else None,
            "quote_sha256": hashlib.sha256(quote.encode("utf-8")).hexdigest() if quote else None,
            "evidence_type": str((evidence or {}).get("evidence_type") or ("contract_text" if quote else "manual")),
            "source_status": "located" if span_start >= 0 else "manual_confirmation",
        }

        issues.append({
            "id": rid,
            "block_id": bid,
            "clause_ref": clause_ref,
            "risk": str(comment.get("risk") or "中"),
            "original_excerpt": excerpt,
            "issue_analysis": str(comment.get("issue") or "").strip(),
            "suggested_text": suggested,
            "legal_basis": basis,
            "legal_basis_evidence": law_basis_evidence,
            "fallback": fallback,
            "evidence": evidence_obj,
        })

        inc_missing = item.get("include_in_missing_clauses")
        if inc_missing is True or (
            inc_missing is None and op in ("insert_paragraph_after", "insert_paragraph_before")
        ):
            clause_text = "\n".join(new_paras) if new_paras else suggested
            auto_missing.append({
                "id": rid,
                "title": _infer_title(item),
                "reason": str(comment.get("issue") or "").strip(),
                "suggested_clause_text": clause_text,
                "legal_basis": basis,
            })

    overall_risk = str(draft.get("overall_risk") or "").strip() or _compute_overall_risk(raw_revisions)
    signing_rec = (
        str(draft.get("signing_recommendation") or "").strip()
        or _SIGNING_MAP.get(overall_risk, "修改核心条款后可签署")
    )

    # 构建 top_risks（1..5 项）
    raw_top = draft.get("top_risks") or []
    top_risks: list[dict[str, str]] = []
    if isinstance(raw_top, list) and raw_top:
        for entry in raw_top[:5]:
            if isinstance(entry, dict) and entry.get("id") and entry.get("title") and entry.get("summary"):
                top_risks.append({
                    "id": str(entry["id"]),
                    "title": str(entry["title"]),
                    "summary": str(entry["summary"]),
                })
            elif isinstance(entry, str) and entry in rev_by_id:
                r_item = rev_by_id[entry]
                cm = r_item.get("comment") or {}
                top_risks.append({
                    "id": entry,
                    "title": _infer_title(r_item),
                    "summary": str(r_item.get("top_risk_summary") or cm.get("issue") or "").strip(),
                })
    if not top_risks:
        rank_map = {"高": 0, "中": 1, "低": 2}
        sorted_revs = sorted(
            (r for r in raw_revisions if isinstance(r, dict)),
            key=lambda r: rank_map.get(str((r.get("comment") or {}).get("risk") or "中"), 9),
        )
        for r_item in sorted_revs[:4]:
            cm = r_item.get("comment") or {}
            top_risks.append({
                "id": str(r_item.get("id") or "R-01"),
                "title": _infer_title(r_item),
                "summary": str(r_item.get("top_risk_summary") or cm.get("issue") or "").strip(),
            })
    if not top_risks:
        top_risks = [{"id": "R-01", "title": "常规条款复核", "summary": "未发现重大红线问题，请按常规流程复核。"}]

    explicit_missing = draft.get("missing_clauses")
    if isinstance(explicit_missing, list):
        missing_clauses = explicit_missing
    else:
        missing_clauses = auto_missing

    pre_checks = draft.get("pre_signing_checks")
    if not isinstance(pre_checks, list) or not pre_checks:
        pre_checks = [
            "主体与资信穿透核查：通过国家企业信用信息公示系统及中国执行信息公开网核实对方非失信被执行人，且注册资本认缴/实缴与交易规模匹配。",
            "授权与决议核查：如对方签字人非法定代表人本人，获取加盖公章的《授权委托书》原件；涉及对外担保的，调取有效股东会/董事会决议。",
            "附件完整性与骑缝章：确认技术协议、价格清单、验收标准及廉洁条款均已定稿，并与正文一并加盖骑缝章。",
        ]

    coverage = draft.get("checklist_coverage")
    if not isinstance(coverage, list) or not coverage:
        routing_chks = ((preflight or {}).get("routing") or {}).get("checklists") or []
        if routing_chks:
            coverage = [
                {
                    "checklist_id": str(c.get("id") or "—"),
                    "checklist_title": str(c.get("title") or ""),
                    "status": "待人工确认",
                    "notes": str(c.get("reason") or "由自动路由生成，需模型逐项确认后方可标记为已全项核查"),
                }
                for c in routing_chks
            ]
        else:
            coverage = [
                {
                    "checklist_id": "group-redlines",
                    "checklist_title": "集团法务强制红线速查表（group-redlines.md）",
                    "status": "待人工确认",
                    "notes": f"自动底座清单，共形成 {len(issues)} 项修订与批注，需模型逐项确认",
                },
                {
                    "checklist_id": "C0102",
                    "checklist_title": "集团法务合同审核要点",
                    "status": "待人工确认",
                    "notes": "通用及专项章节已注入上下文，需模型逐项确认",
                },
                {
                    "checklist_id": "C0101",
                    "checklist_title": "通用商事合同审查总清单（主审与复核双视角）",
                    "status": "待人工确认",
                    "notes": "通用清单已注入上下文，需模型逐项确认",
                },
            ]

    source_name = ((preflight or {}).get("source") or {}).get("name") or ""
    default_contract_name = Path(source_name).stem if source_name else "待审合同"
    contract_name = str(draft.get("contract_name") or "").strip() or default_contract_name

    review_result_obj: dict[str, Any] = {
        "contract_name": contract_name,
        "our_position": str(draft["our_position"]),
        "contract_type": str(draft["contract_type"]),
        "sector": draft.get("sector"),
        "assumptions": [str(a) for a in (draft.get("assumptions") or [])],
        "kb_snapshot": str(draft.get("kb_snapshot") or "").strip() or _default_kb_snapshot(),
        "overall_risk": overall_risk,
        "signing_recommendation": signing_rec,
        "top_risks": top_risks,
        "issues": issues,
        "missing_clauses": missing_clauses,
        "pre_signing_checks": [str(x) for x in pre_checks],
        "checklist_coverage": coverage,
    }
    report_md_text = _render_report_md(draft, review_result_obj, preflight)
    report_p.parent.mkdir(parents=True, exist_ok=True)
    report_tmp = report_p.with_name(f".{report_p.name}.{time.time_ns()}.tmp")
    try:
        report_tmp.write_text(report_md_text, encoding="utf-8")
        report_tmp.replace(report_p)
    finally:
        if report_tmp.exists():
            report_tmp.unlink()

    # 所有确定性内容均在内存中完成并通过关系校验后才一次性提交三个
    # 产物；中途异常不会留下“新 revisions + 旧 review-result”的混合代际。
    _write_json(revisions_p, revisions_obj)
    _write_json(result_p, review_result_obj)

    return {
        "draft": str(draft_p),
        "revisions_json": str(revisions_p),
        "review_result_json": str(result_p),
        "review_report_md": str(report_p),
        "counts": {
            "revisions": len(clean_revisions),
            "issues": len(issues),
            "top_risks": len(top_risks),
            "missing_clauses": len(missing_clauses),
            "report_chars": len(report_md_text),
        },
    }


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8")

    parser = argparse.ArgumentParser(
        description="从 review-draft.json 确定性派生 revisions.json、review-result.json 与 review_report.md"
    )
    parser.add_argument("draft", help="review-draft.json 路径")
    parser.add_argument("--blocks", default=None, help="可选：blocks.json 路径（默认同目录 work/blocks.json）")
    parser.add_argument("--preflight", default=None, help="可选：preflight.json 路径（默认同目录 work/preflight.json）")
    parser.add_argument("--revisions-out", default=None, help="输出 revisions.json 路径")
    parser.add_argument("--review-result-out", default=None, help="输出 review-result.json 路径")
    parser.add_argument("--report-md-out", default=None, help="输出 review_report.md 路径")
    args = parser.parse_args(argv)

    try:
        summary = build_outputs_from_draft(
            args.draft,
            blocks_path=args.blocks,
            preflight_path=args.preflight,
            revisions_out=args.revisions_out,
            review_result_out=args.review_result_out,
            report_md_out=args.report_md_out,
        )
    except Exception as exc:
        print(json.dumps({"error": str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 1

    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
