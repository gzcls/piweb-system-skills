#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""按 revisions.json 向 .docx 写入 OOXML 修订痕迹（w:ins / w:del）与批注（word/comments.xml）。

用法:
  python scripts/docx_redline.py <input.docx> <revisions.json> --out <修订版.docx> \\
      [--log apply_log.json] [--author 集团法务AI审查] [--track-on] [--date 2026-10-30T10:00:00Z]

退出码:
  0: 全部成功或仅存在自动降级为批注 (degraded_to_comment)
  2: 存在失败项 (failed，如 block_id 不存在)
  1: 输入文件缺失、参数非法或发生未预期异常
"""
from __future__ import annotations

import argparse
import copy
import datetime as dt
import json
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from lxml import etree

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from lib.ooxml import (
    BlockRecord,
    CharAtom,
    DocxPackage,
    IdAllocator,
    add_comment_entry,
    attach_comment_around_elements,
    attach_comment_to_whole_paragraph,
    build_char_atoms,
    build_runs_preserving_rPr,
    clone_rPr,
    compute_minimal_diff_opcodes,
    enable_track_revisions,
    format_comment_text,
    iter_document_blocks,
    make_del_element,
    make_ins_element,
    make_text_run,
    paragraph_has_existing_revisions,
    set_xml_space_preserve,
    split_run_at_char_index,
    validate_docx_file,
    w14_tag,
    w_tag,
)
from lib.textnorm import locate_match, validate_revisions_data
from lib.textio import read_text_auto, sanitize_xml_attribute


DEFAULT_AUTHOR = "集团法务AI审查"


@dataclass
class _InlineTask:
    """同一段落内待按偏移从后往前执行的修订/批注任务。"""

    order_idx: int
    rev: dict[str, Any]
    raw_start: int
    raw_end: int


def _fallback_suggestion_for_op(rev: dict[str, Any]) -> str:
    """当修改类操作降级为批注且原 comment 未含 suggestion 时，自动生成修改建议说明。"""
    op = rev.get("op", "")
    anchor = rev.get("anchor") or {}
    match_str = anchor.get("match") or ""
    new_text = rev.get("new_text") or ""
    new_paras = rev.get("new_paragraphs") or []

    if op == "replace":
        return f"将“{match_str}”修改为“{new_text}”"
    if op == "delete":
        return f"删除“{match_str}”"
    if op == "insert_text":
        pos = "之前" if anchor.get("position") == "before" else "之后"
        return f"在“{match_str}”{pos}插入“{new_text}”"
    if op == "delete_paragraph":
        return "删除本整段条款"
    if op in ("insert_paragraph_after", "insert_paragraph_before"):
        pos = "前" if op == "insert_paragraph_before" else "后"
        joined = "；".join(new_paras) if new_paras else new_text
        return f"在本段{pos}新增条款：{joined}"
    return ""


def _attach_degraded_comment(
    pkg: DocxPackage,
    id_alloc: IdAllocator,
    p_elem: etree._Element,
    rev: dict[str, Any],
    author: str,
    date_iso: str,
    degrade_reason: str,
) -> None:
    """将降级后的修订写入目标段落的整段级批注。"""
    cid = id_alloc.next_id()
    fallback_sug = _fallback_suggestion_for_op(rev)
    c_text = format_comment_text(rev["id"], rev.get("comment"), fallback_suggestion=fallback_sug)
    add_comment_entry(pkg, cid, author, date_iso, c_text)
    attach_comment_to_whole_paragraph(p_elem, cid)


def _extract_first_rPr_from_paragraph(p_elem: etree._Element) -> etree._Element | None:
    """从段落首个含 w:rPr 的 w:r（或 pPr/rPr）中提取字体样式模板。"""
    for r_el in p_elem.findall(w_tag("r")):
        cloned = clone_rPr(r_el)
        if cloned is not None:
            return cloned
    pPr = p_elem.find(w_tag("pPr"))
    if pPr is not None:
        rPr = pPr.find(w_tag("rPr"))
        if rPr is not None:
            cloned = copy.deepcopy(rPr)
            for tag_name in ("ins", "del", "rPrChange"):
                for el in cloned.findall(w_tag(tag_name)):
                    cloned.remove(el)
            return cloned
    return None


def _build_inserted_paragraph(
    text: str,
    style_source_p: etree._Element,
    id_alloc: IdAllocator,
    author: str,
    date_iso: str,
) -> etree._Element:
    """构造带修订标记的新增段落 <w:p>：
    - 复制参照段落的 w:pPr（保留缩进、样式、自动编号）；
    - 在 w:pPr/w:rPr 中写入 <w:ins/> 标记段落标记本身为新增；
    - 段落正文包裹在 <w:ins><w:r>...</w:r></w:ins> 中。
    """
    new_p = etree.Element(w_tag("p"))
    src_pPr = style_source_p.find(w_tag("pPr"))
    if src_pPr is not None:
        new_pPr = copy.deepcopy(src_pPr)
        # 移除旧修订记录
        for chg_tag in ("pPrChange", "sectPr"):
            for el in new_pPr.findall(w_tag(chg_tag)):
                new_pPr.remove(el)
    else:
        new_pPr = etree.Element(w_tag("pPr"))

    p_rPr = new_pPr.find(w_tag("rPr"))
    if p_rPr is None:
        p_rPr = etree.SubElement(new_pPr, w_tag("rPr"))
    else:
        for old_tag in ("ins", "del", "rPrChange"):
            for el in p_rPr.findall(w_tag(old_tag)):
                p_rPr.remove(el)

    # 段落标记本身标记为 w:ins
    p_ins = etree.Element(
        w_tag("ins"),
        {
            w_tag("id"): id_alloc.next_id(),
            w_tag("author"): sanitize_xml_attribute(str(author))[0],
            w_tag("date"): sanitize_xml_attribute(str(date_iso))[0],
        },
    )
    p_rPr.insert(0, p_ins)
    new_p.append(new_pPr)

    rPr_tpl = _extract_first_rPr_from_paragraph(style_source_p)
    r_el = make_text_run(text, rPr_tpl, is_del=False)
    ins_el = make_ins_element(id_alloc.next_id(), author, date_iso, [r_el])
    new_p.append(ins_el)
    return new_p


def _apply_delete_paragraph(
    p_elem: etree._Element,
    id_alloc: IdAllocator,
    author: str,
    date_iso: str,
) -> tuple[etree._Element | None, etree._Element | None]:
    """将整段标记为删除：
    - 段内所有 w:r 包裹进 <w:del>，w:t 转为 w:delText；
    - 段落标记 w:pPr/w:rPr 添加 <w:del/>。
    """
    normalize_atoms_res = build_char_atoms(p_elem)
    _, atoms = normalize_atoms_res

    pPr = p_elem.find(w_tag("pPr"))
    if pPr is None:
        pPr = etree.Element(w_tag("pPr"))
        p_elem.insert(0, pPr)
    p_rPr = pPr.find(w_tag("rPr"))
    if p_rPr is None:
        p_rPr = etree.SubElement(pPr, w_tag("rPr"))
    p_del = etree.Element(
        w_tag("del"),
        {
            w_tag("id"): id_alloc.next_id(),
            w_tag("author"): sanitize_xml_attribute(str(author))[0],
            w_tag("date"): sanitize_xml_attribute(str(date_iso))[0],
        },
    )
    p_rPr.insert(0, p_del)

    # 将段内所有直接 w:r 转为 w:del
    direct_runs = [ch for ch in p_elem if ch.tag == w_tag("r")]
    if not direct_runs:
        return None, None

    first_run_idx = list(p_elem).index(direct_runs[0])
    del_runs: list[etree._Element] = []
    for r_el in direct_runs:
        new_r = copy.deepcopy(r_el)
        for t_el in new_r.findall(w_tag("t")):
            t_el.tag = w_tag("delText")
            set_xml_space_preserve(t_el)
        del_runs.append(new_r)
        p_elem.remove(r_el)

    del_wrapper = make_del_element(id_alloc.next_id(), author, date_iso, del_runs)
    p_elem.insert(first_run_idx, del_wrapper)
    return del_wrapper, del_wrapper


def _execute_inline_task(
    pkg: DocxPackage,
    id_alloc: IdAllocator,
    p_elem: etree._Element,
    atoms: list[CharAtom],
    task: _InlineTask,
    author: str,
    date_iso: str,
) -> list[CharAtom]:
    """在段落 p_elem 的 [raw_start, raw_end) 区间上执行单条行内操作（从后往前调用）。"""
    rev = task.rev
    op = rev["op"]
    raw_start, raw_end = task.raw_start, task.raw_end

    # 先切分右侧边界 raw_end，再切分左侧边界 raw_start，保证左侧偏移不变
    atoms = split_run_at_char_index(p_elem, atoms, raw_end)
    atoms = split_run_at_char_index(p_elem, atoms, raw_start)

    slice_atoms = atoms[raw_start:raw_end]
    target_runs: list[etree._Element] = []
    for a in slice_atoms:
        if a.run_elem is not None and (not target_runs or target_runs[-1] is not a.run_elem):
            target_runs.append(a.run_elem)

    if not target_runs:
        return atoms

    children_list = list(p_elem)
    insert_pos = children_list.index(target_runs[0])

    if op == "comment_only":
        cid = id_alloc.next_id()
        c_text = format_comment_text(rev["id"], rev.get("comment"))
        add_comment_entry(pkg, cid, author, date_iso, c_text)
        attach_comment_around_elements(p_elem, target_runs[0], target_runs[-1], cid)
        return atoms

    if op == "delete":
        del_runs = build_runs_preserving_rPr(slice_atoms, is_del=True)
        del_el = make_del_element(id_alloc.next_id(), author, date_iso, del_runs)
        for r_el in target_runs:
            p_elem.remove(r_el)
        p_elem.insert(insert_pos, del_el)
        if rev.get("comment"):
            cid = id_alloc.next_id()
            c_text = format_comment_text(rev["id"], rev["comment"])
            add_comment_entry(pkg, cid, author, date_iso, c_text)
            attach_comment_around_elements(p_elem, del_el, del_el, cid)
        return atoms

    if op == "insert_text":
        position = (rev.get("anchor") or {}).get("position") or "after"
        new_text = rev.get("new_text") or ""
        if position == "before":
            ref_rPr = clone_rPr(slice_atoms[0].run_elem)
            ins_run = make_text_run(new_text, ref_rPr, is_del=False)
            ins_el = make_ins_element(id_alloc.next_id(), author, date_iso, [ins_run])
            p_elem.insert(insert_pos, ins_el)
            first_el, last_el = ins_el, target_runs[-1]
        else:
            ref_rPr = clone_rPr(slice_atoms[-1].run_elem)
            ins_run = make_text_run(new_text, ref_rPr, is_del=False)
            ins_el = make_ins_element(id_alloc.next_id(), author, date_iso, [ins_run])
            after_pos = list(p_elem).index(target_runs[-1]) + 1
            p_elem.insert(after_pos, ins_el)
            first_el, last_el = target_runs[0], ins_el

        if rev.get("comment"):
            cid = id_alloc.next_id()
            c_text = format_comment_text(rev["id"], rev["comment"])
            add_comment_entry(pkg, cid, author, date_iso, c_text)
            attach_comment_around_elements(p_elem, first_el, last_el, cid)
        return atoms

    if op == "replace":
        orig_slice = "".join(a.char for a in slice_atoms)
        new_text = rev.get("new_text") if rev.get("new_text") is not None else ""
        opcodes = compute_minimal_diff_opcodes(orig_slice, new_text, min_equal_len=2)

        emitted: list[etree._Element] = []
        # 若原文片段与新文本完全一致，做一次显式 del+ins 以留痕
        if len(opcodes) == 1 and opcodes[0][0] == "equal":
            del_runs = build_runs_preserving_rPr(slice_atoms, is_del=True)
            emitted.append(make_del_element(id_alloc.next_id(), author, date_iso, del_runs))
            if new_text:
                ins_run = make_text_run(new_text, clone_rPr(slice_atoms[0].run_elem), is_del=False)
                emitted.append(make_ins_element(id_alloc.next_id(), author, date_iso, [ins_run]))
        else:
            for tag, i1, i2, j1, j2 in opcodes:
                sub_atoms = slice_atoms[i1:i2]
                if tag == "equal":
                    emitted.extend(build_runs_preserving_rPr(sub_atoms, is_del=False))
                else:
                    if i2 > i1:
                        del_runs = build_runs_preserving_rPr(sub_atoms, is_del=True)
                        emitted.append(
                            make_del_element(id_alloc.next_id(), author, date_iso, del_runs)
                        )
                    if j2 > j1:
                        ref_atom = (
                            sub_atoms[0]
                            if sub_atoms
                            else (slice_atoms[i1] if i1 < len(slice_atoms) else slice_atoms[-1])
                        )
                        ins_run = make_text_run(
                            new_text[j1:j2],
                            clone_rPr(ref_atom.run_elem),
                            is_del=False,
                        )
                        emitted.append(
                            make_ins_element(id_alloc.next_id(), author, date_iso, [ins_run])
                        )

        for r_el in target_runs:
            p_elem.remove(r_el)
        for offset, el in enumerate(emitted):
            p_elem.insert(insert_pos + offset, el)

        if rev.get("comment") and emitted:
            cid = id_alloc.next_id()
            c_text = format_comment_text(rev["id"], rev["comment"])
            add_comment_entry(pkg, cid, author, date_iso, c_text)
            attach_comment_around_elements(p_elem, emitted[0], emitted[-1], cid)

    return atoms


def apply_revisions(
    input_docx: str | Path,
    revisions_data: dict[str, Any],
    out_docx: str | Path,
    log_path: str | Path | None = None,
    author_override: str | None = None,
    track_on: bool = False,
    date_iso: str | None = None,
) -> dict[str, Any]:
    """执行全部修订指令并写入输出 .docx 与日志。"""
    pkg = DocxPackage(input_docx)
    id_alloc = pkg.create_id_allocator()

    author = (author_override or revisions_data.get("author") or DEFAULT_AUTHOR).strip()
    if not date_iso:
        date_iso = dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    blocks = iter_document_blocks(pkg, include_headers_footers=True)
    block_map: dict[str, BlockRecord] = {}
    for b in blocks:
        block_map[b.block_id] = b
        if b.seq_id not in block_map:
            block_map[b.seq_id] = b

    any_existing_rev = any(b.has_existing_revisions for b in blocks)
    if any_existing_rev:
        print(
            "【警告】输入文档包含已有修订痕迹（w:ins/w:del）。对含已有修订段落的修改类操作将自动降级为批注。",
            file=sys.stdout,
        )

    revisions_list: list[dict[str, Any]] = revisions_data.get("revisions") or []
    log_entries: list[dict[str, Any] | None] = [None] * len(revisions_list)

    # 按目标段落分组处理，以支持同段落多处修订按偏移从后往前执行
    para_groups: dict[str, list[tuple[int, dict[str, Any]]]] = {}
    for idx, rev in enumerate(revisions_list):
        rev_id = str(rev.get("id") or f"REV-{idx + 1}")
        op = str(rev.get("op") or "")
        anchor = rev.get("anchor") or {}
        block_id = str(anchor.get("block_id") or "")

        if not block_id or block_id not in block_map:
            log_entries[idx] = {
                "id": rev_id,
                "op": op,
                "status": "failed",
                "reason": f"未找到锚点段落 block_id='{block_id}'",
                "block_id": block_id,
            }
            continue

        target_block = block_map[block_id]
        para_groups.setdefault(target_block.block_id, []).append((idx, rev))

    # 逐个目标段落执行
    for canonical_block_id, indexed_revs in para_groups.items():
        block = block_map[canonical_block_id]
        p_elem = block.p_elem
        has_existing_rev = paragraph_has_existing_revisions(p_elem)

        # 若目标段落在 header/footer/footnote，首期仅提取不直接修订正文，降级到首段批注或提示
        if block.location not in ("body", "table"):
            fallback_p = next((b.p_elem for b in blocks if b.location in ("body", "table")), None)
            for idx, rev in indexed_revs:
                req_id = rev["anchor"]["block_id"]
                if fallback_p is not None:
                    _attach_degraded_comment(
                        pkg,
                        id_alloc,
                        fallback_p,
                        rev,
                        author,
                        date_iso,
                        f"首期不支持直接修订 {block.location} 区域，降级为批注",
                    )
                    log_entries[idx] = {
                        "id": rev["id"],
                        "op": rev["op"],
                        "status": "degraded_to_comment",
                        "reason": f"目标位于 {block.location} 区域，首期降级为批注",
                        "block_id": req_id,
                    }
                else:
                    log_entries[idx] = {
                        "id": rev["id"],
                        "op": rev["op"],
                        "status": "failed",
                        "reason": f"目标位于 {block.location} 区域且无可用正文段落挂载批注",
                        "block_id": req_id,
                    }
            continue

        raw_text, atoms = build_char_atoms(p_elem)
        accepted_inline_tasks: list[_InlineTask] = []
        occupied_spans: list[tuple[int, int]] = []
        paragraph_deleted = False

        # 第一遍：在原始段落文本上定位所有行内操作，处理段落级操作与降级
        deferred_para_inserts_before: list[tuple[int, dict[str, Any]]] = []
        deferred_para_inserts_after: list[tuple[int, dict[str, Any]]] = []
        deferred_para_delete: list[tuple[int, dict[str, Any]]] = []
        deferred_whole_para_comments: list[tuple[int, dict[str, Any], str, str | None]] = []

        for idx, rev in indexed_revs:
            rev_id = rev["id"]
            op = rev["op"]
            anchor = rev.get("anchor") or {}
            req_block_id = anchor.get("block_id", canonical_block_id)
            match_str = anchor.get("match")
            occurrence = anchor.get("occurrence")

            if op in ("insert_paragraph_before", "insert_paragraph_after"):
                new_paras = rev.get("new_paragraphs") or (
                    [rev["new_text"]] if rev.get("new_text") else []
                )
                if not new_paras or not any(str(p).strip() for p in new_paras):
                    deferred_whole_para_comments.append(
                        (
                            idx,
                            rev,
                            "degraded_to_comment",
                            f"{op} 未提供有效的 new_paragraphs 内容，降级为段落批注",
                        )
                    )
                elif op == "insert_paragraph_before":
                    deferred_para_inserts_before.append((idx, rev))
                else:
                    deferred_para_inserts_after.append((idx, rev))
                continue

            if op == "delete_paragraph":
                if has_existing_rev:
                    print(
                        f"【警告】段落 {req_block_id} 已含修订痕迹，修订 {rev_id} ({op}) 降级为批注。",
                        file=sys.stdout,
                    )
                    deferred_whole_para_comments.append(
                        (
                            idx,
                            rev,
                            "degraded_to_comment",
                            "目标段落已存在修订痕迹(w:ins/w:del)，整段删除降级为批注",
                        )
                    )
                    continue
                prot_reasons = [
                    a.protect_reason
                    for a in atoms
                    if a.is_protected
                    and a.protect_reason
                    and ("超链接" in a.protect_reason or "域" in a.protect_reason or "内容控件" in a.protect_reason)
                ]
                if prot_reasons:
                    deferred_whole_para_comments.append(
                        (
                            idx,
                            rev,
                            "degraded_to_comment",
                            f"目标段落包含不可安全删除的结构({prot_reasons[0]})，降级为批注",
                        )
                    )
                    continue
                deferred_para_delete.append((idx, rev))
                paragraph_deleted = True
                continue

            if op == "comment_only":
                if match_str:
                    span, _ = locate_match(raw_text, match_str, occurrence)
                    if span is not None:
                        s, e = span
                        has_prot = any(a.is_protected for a in atoms[s:e])
                        overlaps = any(max(s, os) < min(e, oe) for os, oe in occupied_spans)
                        if not has_prot and not overlaps and not paragraph_deleted:
                            occupied_spans.append((s, e))
                            accepted_inline_tasks.append(_InlineTask(idx, rev, s, e))
                            log_entries[idx] = {
                                "id": rev_id,
                                "op": op,
                                "status": "applied",
                                "reason": None,
                                "block_id": req_block_id,
                            }
                            continue
                deferred_whole_para_comments.append((idx, rev, "applied", None))
                continue

            # 修改类行内操作: replace / delete / insert_text
            if has_existing_rev:
                print(
                    f"【警告】段落 {req_block_id} 已含修订痕迹，修订 {rev_id} ({op}) 降级为批注。",
                    file=sys.stdout,
                )
                deferred_whole_para_comments.append(
                    (
                        idx,
                        rev,
                        "degraded_to_comment",
                        "目标段落已存在修订痕迹(w:ins/w:del)，修改类操作降级为批注",
                    )
                )
                continue

            if paragraph_deleted:
                deferred_whole_para_comments.append(
                    (
                        idx,
                        rev,
                        "degraded_to_comment",
                        "同段落已被标记为整段删除，该局部修改降级为批注",
                    )
                )
                continue

            if op == "replace" and rev.get("new_text") is None:
                deferred_whole_para_comments.append(
                    (
                        idx,
                        rev,
                        "degraded_to_comment",
                        "replace 操作缺少 new_text 字段，降级为批注",
                    )
                )
                continue

            if op == "insert_text" and not rev.get("new_text"):
                deferred_whole_para_comments.append(
                    (
                        idx,
                        rev,
                        "degraded_to_comment",
                        "insert_text 操作缺少非空 new_text 字段，降级为批注",
                    )
                )
                continue

            span, loc_err = locate_match(raw_text, match_str, occurrence)
            if span is None:
                deferred_whole_para_comments.append(
                    (idx, rev, "degraded_to_comment", loc_err or "未找到匹配文本")
                )
                continue

            s, e = span
            prot_atoms = [a for a in atoms[s:e] if a.is_protected]
            if prot_atoms:
                reason_str = prot_atoms[0].protect_reason or "跨越受保护结构"
                deferred_whole_para_comments.append(
                    (
                        idx,
                        rev,
                        "degraded_to_comment",
                        f"匹配范围{reason_str}，不可安全拆分修改，降级为段落批注",
                    )
                )
                continue

            if any(max(s, os) < min(e, oe) for os, oe in occupied_spans):
                deferred_whole_para_comments.append(
                    (
                        idx,
                        rev,
                        "degraded_to_comment",
                        "与同段落另一条修订的匹配区间重叠，降级为段落批注",
                    )
                )
                continue

            occupied_spans.append((s, e))
            accepted_inline_tasks.append(_InlineTask(idx, rev, s, e))
            log_entries[idx] = {
                "id": rev_id,
                "op": op,
                "status": "applied",
                "reason": None,
                "block_id": req_block_id,
            }

        # 第二遍：按 raw_start 从大到小（从后往前）执行所有行内任务
        accepted_inline_tasks.sort(key=lambda t: (t.raw_start, t.raw_end), reverse=True)
        for task in accepted_inline_tasks:
            atoms = _execute_inline_task(pkg, id_alloc, p_elem, atoms, task, author, date_iso)

        # 第三遍：执行整段删除（若有）
        for idx, rev in deferred_para_delete:
            req_block_id = rev["anchor"].get("block_id", canonical_block_id)
            first_del, last_del = _apply_delete_paragraph(p_elem, id_alloc, author, date_iso)
            if rev.get("comment"):
                cid = id_alloc.next_id()
                c_text = format_comment_text(rev["id"], rev["comment"])
                add_comment_entry(pkg, cid, author, date_iso, c_text)
                if first_del is not None and last_del is not None:
                    attach_comment_around_elements(p_elem, first_del, last_del, cid)
                else:
                    attach_comment_to_whole_paragraph(p_elem, cid)
            log_entries[idx] = {
                "id": rev["id"],
                "op": rev["op"],
                "status": "applied",
                "reason": None,
                "block_id": req_block_id,
            }

        # 第四遍：挂载整段级批注（含 comment_only 及降级项）
        for idx, rev, status_val, reason_val in deferred_whole_para_comments:
            req_block_id = rev["anchor"].get("block_id", canonical_block_id)
            if status_val == "degraded_to_comment":
                _attach_degraded_comment(
                    pkg, id_alloc, p_elem, rev, author, date_iso, reason_val or "降级为批注"
                )
            else:
                cid = id_alloc.next_id()
                c_text = format_comment_text(rev["id"], rev.get("comment"))
                add_comment_entry(pkg, cid, author, date_iso, c_text)
                attach_comment_to_whole_paragraph(p_elem, cid)

            log_entries[idx] = {
                "id": rev["id"],
                "op": rev["op"],
                "status": status_val,
                "reason": reason_val,
                "block_id": req_block_id,
            }

        # 第五遍：执行段落前/后插入（insert_paragraph_before / insert_paragraph_after）
        parent_container = p_elem.getparent()
        if parent_container is not None:
            for idx, rev in deferred_para_inserts_before:
                req_block_id = rev["anchor"].get("block_id", canonical_block_id)
                style_from_id = rev.get("style_from") or req_block_id
                style_src_p = (
                    block_map[style_from_id].p_elem if style_from_id in block_map else p_elem
                )
                new_paras = rev.get("new_paragraphs") or [rev["new_text"]]
                created_ps: list[etree._Element] = []
                anchor_pos = list(parent_container).index(p_elem)
                for offset, p_text in enumerate(new_paras):
                    new_p = _build_inserted_paragraph(
                        str(p_text), style_src_p, id_alloc, author, date_iso
                    )
                    parent_container.insert(anchor_pos + offset, new_p)
                    created_ps.append(new_p)
                if rev.get("comment") and created_ps:
                    cid = id_alloc.next_id()
                    c_text = format_comment_text(rev["id"], rev["comment"])
                    add_comment_entry(pkg, cid, author, date_iso, c_text)
                    attach_comment_to_whole_paragraph(created_ps[0], cid)
                log_entries[idx] = {
                    "id": rev["id"],
                    "op": rev["op"],
                    "status": "applied",
                    "reason": None,
                    "block_id": req_block_id,
                }

            # 若同一个锚点段有多条 insert_paragraph_after，按指令出现顺序依次排在锚点段之后
            last_inserted_after_p = p_elem
            for idx, rev in deferred_para_inserts_after:
                req_block_id = rev["anchor"].get("block_id", canonical_block_id)
                style_from_id = rev.get("style_from") or req_block_id
                style_src_p = (
                    block_map[style_from_id].p_elem if style_from_id in block_map else p_elem
                )
                new_paras = rev.get("new_paragraphs") or [rev["new_text"]]
                created_ps = []
                base_pos = list(parent_container).index(last_inserted_after_p) + 1
                for offset, p_text in enumerate(new_paras):
                    new_p = _build_inserted_paragraph(
                        str(p_text), style_src_p, id_alloc, author, date_iso
                    )
                    parent_container.insert(base_pos + offset, new_p)
                    created_ps.append(new_p)
                    last_inserted_after_p = new_p
                if rev.get("comment") and created_ps:
                    cid = id_alloc.next_id()
                    c_text = format_comment_text(rev["id"], rev["comment"])
                    add_comment_entry(pkg, cid, author, date_iso, c_text)
                    attach_comment_to_whole_paragraph(created_ps[0], cid)
                log_entries[idx] = {
                    "id": rev["id"],
                    "op": rev["op"],
                    "status": "applied",
                    "reason": None,
                    "block_id": req_block_id,
                }

    pkg.mark_modified("word/document.xml")
    if track_on:
        enable_track_revisions(pkg)

    pkg.save(out_docx)
    validation = validate_docx_file(out_docx, strict=True)
    if not validation["valid"]:
        raise ValueError(f"生成的修订版 .docx 校验失败: {'; '.join(validation['errors'])}")

    final_items = [e for e in log_entries if e is not None]
    counts = {
        "total": len(final_items),
        "applied": sum(1 for e in final_items if e["status"] == "applied"),
        "degraded_to_comment": sum(1 for e in final_items if e["status"] == "degraded_to_comment"),
        "failed": sum(1 for e in final_items if e["status"] == "failed"),
    }
    apply_log: dict[str, Any] = {
        "input": str(input_docx),
        "output": str(out_docx),
        "author": author,
        "date": date_iso,
        "track_revisions": track_on,
        "docx_validation": validation,
        "summary": counts,
        "items": final_items,
    }

    if log_path is not None:
        lp = Path(log_path)
        lp.parent.mkdir(parents=True, exist_ok=True)
        lp.write_text(json.dumps(apply_log, ensure_ascii=False, indent=2), encoding="utf-8")

    return apply_log


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8")

    parser = argparse.ArgumentParser(
        description="按 revisions.json 向 .docx 写入 OOXML 修订痕迹与批注"
    )
    parser.add_argument("input", help="原始合同 .docx 文件路径（只读）")
    parser.add_argument("revisions", help="修订指令 revisions.json 文件路径")
    parser.add_argument("--out", required=True, help="输出带修订痕迹的 .docx 路径")
    parser.add_argument("--log", default=None, help="输出 apply_log.json 路径")
    parser.add_argument("--author", default=None, help="覆盖修订与批注作者署名（默认：集团法务AI审查）")
    parser.add_argument("--track-on", action="store_true", help="在 settings.xml 中开启 w:trackRevisions")
    parser.add_argument("--date", default=None, help="指定修订时间戳（ISO-8601，如 2026-10-30T10:00:00Z）")
    args = parser.parse_args(argv)

    in_path = Path(args.input)
    rev_path = Path(args.revisions)
    if not in_path.is_file():
        print(json.dumps({"error": f"输入 docx 不存在: {in_path}"}, ensure_ascii=False), file=sys.stderr)
        return 1
    if not rev_path.is_file():
        print(
            json.dumps({"error": f"revisions.json 不存在: {rev_path}"}, ensure_ascii=False),
            file=sys.stderr,
        )
        return 1

    try:
        revisions_text, _ = read_text_auto(rev_path)
        revisions_data = json.loads(revisions_text)
    except Exception as exc:
        print(
            json.dumps({"error": f"解析 revisions.json 失败: {exc}"}, ensure_ascii=False),
            file=sys.stderr,
        )
        return 1

    schema_errors = validate_revisions_data(revisions_data)
    if schema_errors:
        print(
            json.dumps(
                {"error": "revisions.json Schema 校验未通过", "details": schema_errors},
                ensure_ascii=False,
                indent=2,
            ),
            file=sys.stderr,
        )
        return 1

    try:
        apply_log = apply_revisions(
            input_docx=in_path,
            revisions_data=revisions_data,
            out_docx=args.out,
            log_path=args.log,
            author_override=args.author,
            track_on=args.track_on,
            date_iso=args.date,
        )
    except Exception as exc:
        print(json.dumps({"error": f"写入修订失败: {exc}"}, ensure_ascii=False), file=sys.stderr)
        return 1

    print(json.dumps(apply_log["summary"], ensure_ascii=False, indent=2))
    return 2 if apply_log["summary"]["failed"] > 0 else 0


if __name__ == "__main__":
    raise SystemExit(main())
