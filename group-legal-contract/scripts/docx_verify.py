#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""校验修订版 .docx 的双向文本一致性与 OOXML 结构完整性，并可导出“全部接受”后的清洁版。

校验内容：
1. 全部拒绝 = 原文：在 XML 层移除本次新增 w:ins、还原 w:del、忽略新增批注，逐段比对与原文档文本 100% 一致；
2. 全部接受 = 预期：在 XML 层接受本次新增 w:ins、移除 w:del，验证每条 applied 修订均精确生效；
3. 结构合法性：所有 XML 可解析、python-docx 可重新打开、修订/批注 w:id 唯一、comment 三元组（Start/End/Reference）成对且与 comments.xml 一致，并通过 OPC relationships / Content_Types 严格包检查；
4. 可选渲染检查 (--render-check)：调用 LibreOffice (soffice) 转换为 PDF 验证版面渲染不报错。

任一硬性检查失败退出码为 1，全部通过退出码为 0。
"""
from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Any

import docx
from lxml import etree

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from lib.ooxml import (
    COMMENTS_CONTENT_TYPE,
    COMMENTS_REL_TYPE,
    CT_NS,
    PKG_REL_NS,
    PPR_ORDER,
    RPR_ORDER,
    SECTPR_ORDER,
    SETTINGS_ORDER,
    STYLE_ORDER,
    TBLPR_ORDER,
    TCPR_ORDER,
    W_NS,
    DocxPackage,
    build_char_atoms,
    extract_paragraph_accepted_text,
    iter_document_blocks,
    paragraph_has_existing_revisions,
    set_xml_space_preserve,
    validate_docx_file,
    w_tag,
)
from lib.textnorm import locate_match
from lib.textio import read_text_auto, sanitize_xml_text


def _collect_ids_from_pkg(pkg: DocxPackage) -> tuple[set[str], set[str]]:
    """收集文档包中已有的 (修订 w:id 集合, 批注 w:id 集合)。"""
    rev_ids: set[str] = set()
    com_ids: set[str] = set()
    id_attr = w_tag("id")

    doc_root = pkg.get_xml("word/document.xml")
    if doc_root is not None:
        for el in doc_root.iter():
            val = el.get(id_attr)
            if not val:
                continue
            if el.tag in (w_tag("ins"), w_tag("del"), w_tag("moveFrom"), w_tag("moveTo")):
                rev_ids.add(val)
            elif el.tag in (
                w_tag("commentRangeStart"),
                w_tag("commentRangeEnd"),
                w_tag("commentReference"),
            ):
                com_ids.add(val)

    com_root = pkg.get_xml("word/comments.xml")
    if com_root is not None:
        for el in com_root.findall(w_tag("comment")):
            val = el.get(id_attr)
            if val:
                com_ids.add(val)

    return rev_ids, com_ids


def _remove_new_comments_in_tree(
    doc_root: etree._Element,
    orig_comment_ids: set[str],
) -> None:
    """从 document.xml 树中移除本次新增的 commentRangeStart/End 与 commentReference run。"""
    id_attr = w_tag("id")
    for c_tag in (w_tag("commentRangeStart"), w_tag("commentRangeEnd")):
        for el in list(doc_root.iter(c_tag)):
            cid = el.get(id_attr, "")
            if cid not in orig_comment_ids:
                parent = el.getparent()
                if parent is not None:
                    parent.remove(el)

    for ref_el in list(doc_root.iter(w_tag("commentReference"))):
        cid = ref_el.get(id_attr, "")
        if cid not in orig_comment_ids:
            r_parent = ref_el.getparent()
            if r_parent is not None and r_parent.tag == w_tag("r"):
                other_children = [
                    c for c in r_parent if c.tag not in (w_tag("rPr"), w_tag("commentReference"))
                ]
                if not other_children:
                    p_parent = r_parent.getparent()
                    if p_parent is not None:
                        p_parent.remove(r_parent)
                    continue
            if r_parent is not None:
                r_parent.remove(ref_el)


def simulate_reject_all(
    revised_docx: str | Path,
    orig_rev_ids: set[str],
    orig_comment_ids: set[str],
) -> DocxPackage:
    """在内存中对 revised_docx 模拟“全部拒绝本次修订”（移除新增 w:ins、还原新增 w:del、移除新增批注）。"""
    pkg = DocxPackage(revised_docx)
    doc_root = pkg.get_xml("word/document.xml")
    if doc_root is None:
        return pkg

    id_attr = w_tag("id")

    # 1. 处理段落级新增或删除标记（位于 w:p/w:pPr/w:rPr 下的 w:ins / w:del）
    for p_el in list(doc_root.iter(w_tag("p"))):
        pPr = p_el.find(w_tag("pPr"))
        if pPr is None:
            continue
        p_rPr = pPr.find(w_tag("rPr"))
        if p_rPr is None:
            continue

        # 若段落标记含本次新增的 w:ins，且段内无非新增文本，拒绝修订即移除整段
        p_ins_list = [
            el for el in p_rPr.findall(w_tag("ins")) if el.get(id_attr, "") not in orig_rev_ids
        ]
        if p_ins_list:
            parent = p_el.getparent()
            if parent is not None:
                parent.remove(p_el)
                continue

        # 若段落标记含本次新增的 w:del，拒绝修订即移除该段落级 w:del 标记以保留段落
        for d_el in list(p_rPr.findall(w_tag("del"))):
            if d_el.get(id_attr, "") not in orig_rev_ids:
                p_rPr.remove(d_el)

    # 2. 移除正文内所有本次新增的 <w:ins>（非 pPr 内）
    for ins_el in list(doc_root.iter(w_tag("ins"))):
        if ins_el.get(id_attr, "") in orig_rev_ids:
            continue
        parent = ins_el.getparent()
        if parent is None:
            continue
        parent.remove(ins_el)

    # 3. 还原正文内所有本次新增的 <w:del>（将 w:delText 转回 w:t 并解包）
    for del_el in list(doc_root.iter(w_tag("del"))):
        if del_el.get(id_attr, "") in orig_rev_ids:
            continue
        parent = del_el.getparent()
        if parent is None:
            continue
        idx = list(parent).index(del_el)
        for child in list(del_el):
            for dt_el in child.iter(w_tag("delText")):
                dt_el.tag = w_tag("t")
                set_xml_space_preserve(dt_el)
            parent.insert(idx, child)
            idx += 1
        parent.remove(del_el)

    # 4. 移除本次新增的批注标记
    _remove_new_comments_in_tree(doc_root, orig_comment_ids)
    pkg.mark_modified("word/document.xml")
    return pkg


def simulate_accept_all(
    revised_docx: str | Path,
    orig_rev_ids: set[str],
    orig_comment_ids: set[str],
    orig_had_comments_part: bool = False,
    orig_had_track_revisions: bool = False,
) -> DocxPackage:
    """在内存中对 revised_docx 模拟“全部接受本次修订”并移除本次新增批注（用于校验与导出清洁版）。"""
    pkg = DocxPackage(revised_docx)
    doc_root = pkg.get_xml("word/document.xml")
    if doc_root is None:
        return pkg

    id_attr = w_tag("id")

    # 1. 处理段落级删除与新增标记
    for p_el in list(doc_root.iter(w_tag("p"))):
        pPr = p_el.find(w_tag("pPr"))
        if pPr is None:
            continue
        p_rPr = pPr.find(w_tag("rPr"))
        if p_rPr is None:
            continue

        # 整段删除：接受修订时移除该段落（若在表格单元格 w:tc 中且是唯一段落，保留空段落以满足 OOXML 约束）
        p_del_list = [
            el for el in p_rPr.findall(w_tag("del")) if el.get(id_attr, "") not in orig_rev_ids
        ]
        if p_del_list:
            parent = p_el.getparent()
            if parent is not None:
                if parent.tag == w_tag("tc") and len(parent.findall(w_tag("p"))) <= 1:
                    # 清空该单元格内唯一段落的删除内容与标记，保留空 <w:p>
                    for d_el in p_del_list:
                        p_rPr.remove(d_el)
                    for ch in list(p_el):
                        if ch.tag != w_tag("pPr"):
                            p_el.remove(ch)
                else:
                    parent.remove(p_el)
                continue

        # 新增段落：接受修订时移除段落标记上的 w:ins
        for i_el in list(p_rPr.findall(w_tag("ins"))):
            if i_el.get(id_attr, "") not in orig_rev_ids:
                p_rPr.remove(i_el)

    # 2. 移除所有本次新增的 <w:del>
    for del_el in list(doc_root.iter(w_tag("del"))):
        if del_el.get(id_attr, "") in orig_rev_ids:
            continue
        parent = del_el.getparent()
        if parent is not None:
            parent.remove(del_el)

    # 3. 解包所有本次新增的 <w:ins>（保留内部 <w:r>）
    for ins_el in list(doc_root.iter(w_tag("ins"))):
        if ins_el.get(id_attr, "") in orig_rev_ids:
            continue
        parent = ins_el.getparent()
        if parent is None:
            continue
        idx = list(parent).index(ins_el)
        for child in list(ins_el):
            parent.insert(idx, child)
            idx += 1
        parent.remove(ins_el)

    # 4. 清理本次新增的批注标记与 comments.xml 中的对应批注
    _remove_new_comments_in_tree(doc_root, orig_comment_ids)
    pkg.mark_modified("word/document.xml")

    com_root = pkg.get_xml("word/comments.xml")
    if com_root is not None:
        for com_el in list(com_root.findall(w_tag("comment"))):
            cid = com_el.get(id_attr, "")
            if cid not in orig_comment_ids:
                com_root.remove(com_el)
        if not orig_had_comments_part and len(com_root.findall(w_tag("comment"))) == 0:
            pkg.remove_part("word/comments.xml")
            rels_root = pkg.get_xml("word/_rels/document.xml.rels")
            if rels_root is not None:
                for rel in list(rels_root):
                    if rel.get("Type") == COMMENTS_REL_TYPE:
                        rels_root.remove(rel)
                pkg.mark_modified("word/_rels/document.xml.rels")
            ct_root = pkg.get_xml("[Content_Types].xml")
            if ct_root is not None:
                for ovr in list(ct_root):
                    if ovr.get("PartName") in ("/word/comments.xml", "word/comments.xml"):
                        ct_root.remove(ovr)
                pkg.mark_modified("[Content_Types].xml")
        else:
            pkg.mark_modified("word/comments.xml")

    if not orig_had_track_revisions:
        settings_root = pkg.get_xml("word/settings.xml")
        if settings_root is not None:
            tr_el = settings_root.find(w_tag("trackRevisions"))
            if tr_el is not None:
                settings_root.remove(tr_el)
                pkg.mark_modified("word/settings.xml")

    return pkg


def check_structure_integrity(revised_docx: str | Path) -> dict[str, Any]:
    """检查修订版 .docx 的 XML 结构合法性、ID 唯一性、批注三元组配对及 python-docx 可重开性。"""
    errors: list[str] = []
    rev_path = Path(revised_docx)

    # 包级检查覆盖 OPC relationships、Content_Types 和部件路径等常见损坏点。
    # 这里使用 strict=True 作为修订校验的硬性门槛；普通生成流程也会在落盘前执行同一检查。
    strict_package_res = validate_docx_file(rev_path, reopen=False, strict=True)
    if not strict_package_res["valid"]:
        errors.extend(f"严格包校验: {err}" for err in strict_package_res["errors"])

    # 1. 检查所有 XML / rels 文件是否可被 lxml 解析
    xml_parse_ok = True
    try:
        with zipfile.ZipFile(rev_path, "r") as zf:
            names = zf.namelist()
            if len(names) != len(set(names)):
                errors.append("ZIP 包包含重复部件名称")
            bad_name = zf.testzip()
            if bad_name:
                errors.append(f"ZIP CRC 校验失败: {bad_name}")
            required_parts = {"[Content_Types].xml", "_rels/.rels", "word/document.xml"}
            missing_parts = sorted(required_parts - set(names))
            if missing_parts:
                errors.append(f"缺少必需 OOXML 部件: {', '.join(missing_parts)}")
            parser = etree.XMLParser(remove_blank_text=False, resolve_entities=False)
            for name in names:
                if name.endswith(".xml") or name.endswith(".rels"):
                    try:
                        etree.fromstring(zf.read(name), parser=parser)
                    except Exception as exc:
                        xml_parse_ok = False
                        errors.append(f"XML 解析失败 ({name}): {exc}")
    except Exception as exc:
        xml_parse_ok = False
        errors.append(f"ZIP 读取失败: {exc}")

    # 2. 检查 python-docx 是否可正常打开并遍历段落与表格
    docx_reopen_ok = True
    try:
        d = docx.Document(str(rev_path))
        _ = len(d.paragraphs)
        _ = len(d.tables)
    except Exception as exc:
        docx_reopen_ok = False
        errors.append(f"python-docx 重新打开失败: {exc}")

    try:
        pkg = DocxPackage(rev_path)
    except Exception as exc:
        errors.append(f"DOCX 包加载失败: {exc}")
        return {
            "valid": False,
            "xml_well_formed": xml_parse_ok,
            "python_docx_reopen": docx_reopen_ok,
            "strict_package_valid": strict_package_res["valid"],
            "ecma376_compliant": False,
            "revision_tag_count": 0,
            "comment_count": 0,
            "errors": errors,
        }
    doc_root = pkg.get_xml("word/document.xml")
    id_attr = w_tag("id")

    # 3. 检查修订 ID (w:ins / w:del) 唯一性
    rev_id_list: list[str] = []
    duplicate_rev_ids: list[str] = []
    seen_rev_ids: set[str] = set()
    if doc_root is not None:
        for el in doc_root.iter():
            if el.tag in (w_tag("ins"), w_tag("del")):
                rid = el.get(id_attr)
                if not rid:
                    errors.append(f"存在缺少 w:id 属性的修订节点 <{etree.QName(el.tag).localname}>")
                    continue
                rev_id_list.append(rid)
                if rid in seen_rev_ids:
                    duplicate_rev_ids.append(rid)
                seen_rev_ids.add(rid)
    if duplicate_rev_ids:
        errors.append(f"修订节点 w:id 存在重复: {sorted(set(duplicate_rev_ids))}")

    # 4. 检查批注三元组配对、顺序与 comments.xml / rels / Content_Types 完整性
    start_ids: list[str] = []
    end_ids: list[str] = []
    ref_ids: list[str] = []
    order_events: dict[str, list[str]] = {}

    if doc_root is not None:
        for el in doc_root.iter():
            if el.tag == w_tag("commentRangeStart"):
                cid = el.get(id_attr, "")
                start_ids.append(cid)
                order_events.setdefault(cid, []).append("start")
            elif el.tag == w_tag("commentRangeEnd"):
                cid = el.get(id_attr, "")
                end_ids.append(cid)
                order_events.setdefault(cid, []).append("end")
            elif el.tag == w_tag("commentReference"):
                cid = el.get(id_attr, "")
                ref_ids.append(cid)
                order_events.setdefault(cid, []).append("ref")

    if len(start_ids) != len(set(start_ids)):
        errors.append("commentRangeStart 存在重复 w:id")
    if len(end_ids) != len(set(end_ids)):
        errors.append("commentRangeEnd 存在重复 w:id")
    if len(ref_ids) != len(set(ref_ids)):
        errors.append("commentReference 存在重复 w:id")

    start_set, end_set, ref_set = set(start_ids), set(end_ids), set(ref_ids)
    if not (start_set == end_set == ref_set):
        errors.append(
            f"批注三元素 ID 集合不匹配: start={sorted(start_set)}, end={sorted(end_set)}, ref={sorted(ref_set)}"
        )

    for cid, evs in order_events.items():
        if "start" in evs and "end" in evs:
            if evs.index("start") > evs.index("end"):
                errors.append(f"批注 id={cid} 的 commentRangeEnd 出现在 commentRangeStart 之前")

    comment_xml_ids: list[str] = []
    com_root = pkg.get_xml("word/comments.xml")
    if com_root is not None:
        for c_el in com_root.findall(w_tag("comment")):
            cid = c_el.get(id_attr, "")
            if not cid:
                errors.append("word/comments.xml 中存在无 w:id 的 <w:comment>")
            comment_xml_ids.append(cid)
        if len(comment_xml_ids) != len(set(comment_xml_ids)):
            errors.append("word/comments.xml 中 <w:comment> 的 w:id 存在重复")

    if start_set:
        if com_root is None:
            errors.append("正文包含批注引用，但缺少 word/comments.xml 部件")
        else:
            missing_in_xml = start_set - set(comment_xml_ids)
            if missing_in_xml:
                errors.append(f"正文中引用的批注 ID 在 word/comments.xml 中不存在: {sorted(missing_in_xml)}")

        # 检查 document.xml.rels
        rels_root = pkg.get_xml("word/_rels/document.xml.rels")
        has_rel = False
        if rels_root is not None:
            for rel in rels_root:
                if rel.get("Type") == COMMENTS_REL_TYPE:
                    has_rel = True
                    break
        if not has_rel:
            errors.append("word/_rels/document.xml.rels 缺少 comments 关系声明")

        # 检查 [Content_Types].xml
        ct_root = pkg.get_xml("[Content_Types].xml")
        has_ct = False
        if ct_root is not None:
            for child in ct_root:
                if (
                    child.get("PartName") in ("/word/comments.xml", "word/comments.xml")
                    and child.get("ContentType") == COMMENTS_CONTENT_TYPE
                ):
                    has_ct = True
                    break
        if not has_ct:
            errors.append("[Content_Types].xml 缺少 /word/comments.xml 的 ContentType 声明")

    # 5. 检查 ECMA-376 <xsd:sequence> 元素顺序与单次出现约束（防止桌面版 Word 报“文件已损坏”）
    ecma_errors = _check_ecma376_sequence_and_duplicates(pkg)
    errors.extend(ecma_errors)

    return {
        "valid": len(errors) == 0,
        "xml_well_formed": xml_parse_ok,
        "python_docx_reopen": docx_reopen_ok,
        "strict_package_valid": strict_package_res["valid"],
        "ecma376_compliant": len(ecma_errors) == 0,
        "revision_tag_count": len(rev_id_list),
        "comment_count": len(comment_xml_ids),
        "errors": errors,
    }


def _check_ecma376_sequence_and_duplicates(pkg: DocxPackage) -> list[str]:
    """检查文档包内各 XML 部件是否满足 ECMA-376 <xsd:sequence> 顺序及 maxOccurs=1 约束。"""
    errs: list[str] = []
    checks = [
        (w_tag("pPr"), PPR_ORDER, set()),
        (w_tag("rPr"), RPR_ORDER, {w_tag("ins"), w_tag("del"), w_tag("moveFrom"), w_tag("moveTo")}),
        (w_tag("tblPr"), TBLPR_ORDER, set()),
        (w_tag("tcPr"), TCPR_ORDER, {w_tag("headers")}),
        (w_tag("sectPr"), SECTPR_ORDER, {w_tag("headerReference"), w_tag("footerReference")}),
    ]
    part_names = [
        name
        for name in list(pkg.raw_parts.keys()) + list(pkg.xml_parts.keys())
        if name.startswith("word/") and name.endswith(".xml")
    ]
    for name in sorted(set(part_names)):
        root = pkg.get_xml(name)
        if root is None:
            continue
        if root.tag == w_tag("settings"):
            order_map = {w_tag(k): i for i, k in enumerate(SETTINGS_ORDER)}
            last_idx = -1
            last_local = ""
            for ch in root:
                if not isinstance(ch.tag, str) or not ch.tag.startswith(f"{{{W_NS}}}"):
                    continue
                idx = order_map.get(ch.tag, 9999)
                local = etree.QName(ch.tag).localname
                if idx < last_idx:
                    errs.append(f"{name}: <w:settings> 子元素顺序违反 ECMA-376 (<w:{local}> 出现在 <w:{last_local}> 之后)")
                    break
                last_idx, last_local = idx, local

        for container_tag, order_list, multi_tags in checks:
            order_map = {w_tag(k): i for i, k in enumerate(order_list)}
            c_local = etree.QName(container_tag).localname
            for el in root.iter(container_tag):
                seen_tags: set[str] = set()
                last_idx = -1
                last_local = ""
                for ch in el:
                    if not isinstance(ch.tag, str) or not ch.tag.startswith(f"{{{W_NS}}}"):
                        continue
                    local = etree.QName(ch.tag).localname
                    if ch.tag in seen_tags and ch.tag not in multi_tags:
                        errs.append(f"{name}: <w:{c_local}> 内存在重复子元素 <w:{local}>")
                        break
                    seen_tags.add(ch.tag)
                    idx = order_map.get(ch.tag, 9999)
                    if idx < last_idx:
                        errs.append(
                            f"{name}: <w:{c_local}> 子元素顺序违反 ECMA-376 (<w:{local}> 出现在 <w:{last_local}> 之后)"
                        )
                        break
                    last_idx, last_local = idx, local
                if len(errs) >= 10:
                    return errs
    return errs


def _infer_statuses_without_log(
    orig_pkg: DocxPackage,
    revisions_list: list[dict[str, Any]],
) -> dict[str, str]:
    """当未传入 apply_log.json 时，根据原文档结构推导各条修订的预期状态。"""
    blocks = iter_document_blocks(orig_pkg, include_headers_footers=True)
    block_map = {}
    for b in blocks:
        block_map[b.block_id] = b
        if b.seq_id not in block_map:
            block_map[b.seq_id] = b

    status_map: dict[str, str] = {}
    para_groups: dict[str, list[dict[str, Any]]] = {}
    for rev in revisions_list:
        rid = rev["id"]
        bid = (rev.get("anchor") or {}).get("block_id", "")
        if not bid or bid not in block_map:
            status_map[rid] = "failed"
            continue
        tb = block_map[bid]
        if tb.location not in ("body", "table"):
            status_map[rid] = "degraded_to_comment"
            continue
        para_groups.setdefault(tb.block_id, []).append(rev)

    for canonical_id, revs in para_groups.items():
        b = block_map[canonical_id]
        raw_text, atoms = build_char_atoms(b.p_elem)
        has_rev = paragraph_has_existing_revisions(b.p_elem)
        occupied: list[tuple[int, int]] = []
        para_deleted = False

        for rev in revs:
            rid = rev["id"]
            op = rev["op"]
            anchor = rev.get("anchor") or {}
            if op in ("insert_paragraph_before", "insert_paragraph_after"):
                new_ps = rev.get("new_paragraphs") or ([rev["new_text"]] if rev.get("new_text") else [])
                status_map[rid] = "applied" if new_ps else "degraded_to_comment"
            elif op == "delete_paragraph":
                if has_rev:
                    status_map[rid] = "degraded_to_comment"
                else:
                    status_map[rid] = "applied"
                    para_deleted = True
            elif op == "comment_only":
                status_map[rid] = "applied"
                if anchor.get("match"):
                    sp, _ = locate_match(raw_text, anchor.get("match"), anchor.get("occurrence"))
                    if sp is not None and not any(a.is_protected for a in atoms[sp[0] : sp[1]]):
                        if not any(max(sp[0], os) < min(sp[1], oe) for os, oe in occupied):
                            occupied.append(sp)
            else:
                if has_rev or para_deleted:
                    status_map[rid] = "degraded_to_comment"
                    continue
                sp, _ = locate_match(raw_text, anchor.get("match"), anchor.get("occurrence"))
                if sp is None or any(a.is_protected for a in atoms[sp[0] : sp[1]]):
                    status_map[rid] = "degraded_to_comment"
                    continue
                if any(max(sp[0], os) < min(sp[1], oe) for os, oe in occupied):
                    status_map[rid] = "degraded_to_comment"
                    continue
                occupied.append(sp)
                status_map[rid] = "applied"

    return status_map


def _detect_existing_revision_conflicts(
    original_blocks: list[Any],
    revisions_data: dict[str, Any],
) -> list[dict[str, Any]]:
    """识别无法由结构校验判断的“既有修订+新条款”语义合并风险。

    原文含有未接受的 w:ins/w:del 时，插入段落不会自动删除旧条款；
    该情况不应使 OOXML 校验失败，但必须进入交付报告供人工确认。
    """
    by_id: dict[str, Any] = {}
    for block in original_blocks:
        by_id[str(getattr(block, "block_id", ""))] = block
        seq_id = str(getattr(block, "seq_id", ""))
        if seq_id:
            by_id.setdefault(seq_id, block)
    warnings: list[dict[str, Any]] = []
    risky_ops = {
        "replace",
        "delete",
        "insert_text",
        "delete_paragraph",
        "insert_paragraph_before",
        "insert_paragraph_after",
    }
    for revision in revisions_data.get("revisions") or []:
        if not isinstance(revision, dict) or revision.get("op") not in risky_ops:
            continue
        anchor = revision.get("anchor") or {}
        block_id = str(anchor.get("block_id") or "")
        block = by_id.get(block_id)
        if not block or not getattr(block, "has_existing_revisions", False):
            continue
        op = str(revision.get("op") or "")
        category = (
            "inserted_replacement_may_coexist"
            if op in {"insert_paragraph_before", "insert_paragraph_after"}
            else "inline_edit_on_existing_revision"
        )
        warnings.append(
            {
                "id": str(revision.get("id") or ""),
                "block_id": block_id,
                "op": op,
                "category": category,
                "message": (
                    "锚点段落含原有未接受修订；新增段落不会自动删除旧条款，"
                    "签署前需在 Word 中接受/拒绝原修订并确认新旧条款是否冲突。"
                    if category == "inserted_replacement_may_coexist"
                    else "锚点段落含原有未接受修订；该修改可能已降级为批注，需人工确认最终文本。"
                ),
            }
        )
    return warnings


def _verify_accept_all_expectations(
    orig_pkg: DocxPackage,
    revised_pkg: DocxPackage,
    accepted_pkg: DocxPackage,
    revisions_data: dict[str, Any],
    status_map: dict[str, str],
) -> dict[str, Any]:
    """验证模拟“全部接受”后的文档是否符合 revisions.json 中所有 applied 与 degraded 修订的预期。"""
    errors: list[str] = []
    verified_details: list[dict[str, Any]] = []

    orig_blocks = iter_document_blocks(orig_pkg, include_headers_footers=True)
    orig_map = {}
    for b in orig_blocks:
        orig_map[b.block_id] = b
        if b.seq_id not in orig_map:
            orig_map[b.seq_id] = b

    acc_blocks = iter_document_blocks(accepted_pkg, include_headers_footers=True)
    acc_all_texts = [b.text for b in acc_blocks]

    # 收集 revised_pkg 中所有批注文本以便校验 comment_only 与 degraded_to_comment
    com_root = revised_pkg.get_xml("word/comments.xml")
    all_comment_texts: list[str] = []
    if com_root is not None:
        for c_el in com_root.findall(w_tag("comment")):
            t_chunks = [t.text for t in c_el.iter(w_tag("t")) if t.text]
            all_comment_texts.append("".join(t_chunks))
    joined_comments = "\n".join(all_comment_texts)

    revisions_list: list[dict[str, Any]] = revisions_data.get("revisions") or []

    def safe_text(value: Any) -> str:
        return sanitize_xml_text(str(value))[0]

    # 按原始锚点段落聚合 applied 的修订，计算每个锚点段在全部接受后的预期文本
    para_applied_inline: dict[str, list[tuple[int, int, dict[str, Any]]]] = {}
    para_deleted_ids: set[str] = set()
    inserted_para_expectations: list[tuple[str, str, list[str]]] = []

    for rev in revisions_list:
        rid = rev["id"]
        op = rev["op"]
        st = status_map.get(rid, "applied")
        anchor = rev.get("anchor") or {}
        bid = anchor.get("block_id", "")

        if st == "failed":
            continue

        if st == "degraded_to_comment":
            # 验证批注中已写入该修订 ID 的批注
            if f"【{rid}｜" not in joined_comments:
                errors.append(f"降级修订 {rid} 未在 word/comments.xml 中找到对应批注标记 【{rid}｜")
                verified_details.append({"id": rid, "status": st, "ok": False})
            else:
                verified_details.append({"id": rid, "status": st, "ok": True})
            continue

        if bid not in orig_map:
            errors.append(f"修订 {rid} 标记为 applied，但原文档中不存在 block_id={bid}")
            verified_details.append({"id": rid, "status": st, "ok": False})
            continue

        orig_b = orig_map[bid]
        canon_id = orig_b.block_id

        if op == "comment_only":
            if f"【{rid}｜" not in joined_comments:
                errors.append(f"纯批注修订 {rid} 未在 word/comments.xml 中找到对应批注 【{rid}｜")
                verified_details.append({"id": rid, "status": st, "ok": False})
            else:
                verified_details.append({"id": rid, "status": st, "ok": True})
            continue

        if op == "delete_paragraph":
            para_deleted_ids.add(canon_id)
            verified_details.append({"id": rid, "status": st, "ok": True})
            continue

        if op in ("insert_paragraph_after", "insert_paragraph_before"):
            new_ps = rev.get("new_paragraphs") or ([rev["new_text"]] if rev.get("new_text") else [])
            inserted_para_expectations.append((rid, op, [safe_text(x) for x in new_ps]))
            continue

        # replace / delete / insert_text
        span, loc_err = locate_match(orig_b.text, anchor.get("match"), anchor.get("occurrence"))
        if span is None:
            errors.append(f"修订 {rid} ({op}) 无法在原段落 {bid} 中重现定位: {loc_err}")
            verified_details.append({"id": rid, "status": st, "ok": False})
            continue

        para_applied_inline.setdefault(canon_id, []).append((span[0], span[1], rev))

    # 建立原文档未删除且未新增的段落与 accepted_pkg 段落的对应关系
    # 通过遍历原文档块顺序，结合 insert_paragraph_before/after 和 delete_paragraph 推导预期段落序列
    before_inserts_map: dict[str, list[str]] = {}
    after_inserts_map: dict[str, list[str]] = {}
    for rev in revisions_list:
        rid = rev["id"]
        if status_map.get(rid) != "applied":
            continue
        op = rev["op"]
        bid = (rev.get("anchor") or {}).get("block_id", "")
        if bid not in orig_map:
            continue
        canon_id = orig_map[bid].block_id
        if op == "insert_paragraph_before":
            new_ps = rev.get("new_paragraphs") or ([rev["new_text"]] if rev.get("new_text") else [])
            before_inserts_map.setdefault(canon_id, []).extend(safe_text(x) for x in new_ps)
        elif op == "insert_paragraph_after":
            new_ps = rev.get("new_paragraphs") or ([rev["new_text"]] if rev.get("new_text") else [])
            after_inserts_map.setdefault(canon_id, []).extend(safe_text(x) for x in new_ps)

    expected_doc_sequence: list[tuple[str, str]] = []
    expected_by_canon_id: dict[str, str] = {}

    for b in orig_blocks:
        cid = b.block_id
        for ins_txt in before_inserts_map.get(cid, []):
            expected_doc_sequence.append((f"ins_before:{cid}", ins_txt))

        if cid in para_deleted_ids:
            # 若是表格单元格唯一段落，simulate_accept_all 会保留空段落以符合 OOXML
            parent = b.p_elem.getparent()
            if parent is not None and parent.tag == w_tag("tc") and len(parent.findall(w_tag("p"))) <= 1:
                expected_doc_sequence.append((cid, ""))
                expected_by_canon_id[cid] = ""
        else:
            exp_text = b.text
            inline_items = para_applied_inline.get(cid, [])
            inline_items_sorted = sorted(inline_items, key=lambda x: (x[0], x[1]), reverse=True)
            for s, e, rev in inline_items_sorted:
                op = rev["op"]
                if op == "replace":
                    exp_text = exp_text[:s] + safe_text(rev.get("new_text") or "") + exp_text[e:]
                elif op == "delete":
                    exp_text = exp_text[:s] + exp_text[e:]
                elif op == "insert_text":
                    pos = (rev.get("anchor") or {}).get("position") or "after"
                    nt = safe_text(rev.get("new_text") or "")
                    if pos == "before":
                        exp_text = exp_text[:s] + nt + exp_text[s:]
                    else:
                        exp_text = exp_text[:e] + nt + exp_text[e:]
            expected_by_canon_id[cid] = exp_text
            expected_doc_sequence.append((cid, exp_text))

        for ins_txt in after_inserts_map.get(cid, []):
            expected_doc_sequence.append((f"ins_after:{cid}", ins_txt))

    # 1. 全局序列比对：接受全部修订后的段落序列应与推导出的预期序列完全一致
    if len(acc_blocks) != len(expected_doc_sequence):
        errors.append(
            f"全部接受后段落总数 ({len(acc_blocks)}) 与预期段落总数 ({len(expected_doc_sequence)}) 不一致"
        )
    else:
        for idx_p, (acc_b, (exp_tag, exp_txt)) in enumerate(zip(acc_blocks, expected_doc_sequence)):
            if acc_b.text != exp_txt:
                errors.append(
                    f"第 {idx_p + 1} 段 ({exp_tag}) 全部接受后文本不符: 期望={exp_txt!r}, 实际={acc_b.text!r}"
                )

    # 2. 逐条验证行内修订与新增段落修订的语义条件
    for cid, items in para_applied_inline.items():
        orig_text = orig_map[cid].text
        exp_text = expected_by_canon_id.get(cid, "")
        for s, e, rev in items:
            rid = rev["id"]
            op = rev["op"]
            matched_orig = orig_text[s:e]
            nt = safe_text(rev.get("new_text") or "")
            item_ok = True
            if op in ("replace", "insert_text") and nt and nt not in exp_text:
                errors.append(f"修订 {rid} ({op}) 接受后段落未包含 new_text: {nt!r}")
                item_ok = False
            if op in ("replace", "delete"):
                # 若被删原片段在原段仅出现 1 次且不是 new_text 的子串，则接受后不应再出现
                if orig_text.count(matched_orig) == 1 and (op == "delete" or matched_orig not in nt):
                    if matched_orig in exp_text:
                        errors.append(
                            f"修订 {rid} ({op}) 接受后仍残留被删内容: {matched_orig!r}"
                        )
                        item_ok = False
            verified_details.append({"id": rid, "status": "applied", "ok": item_ok})

    for rid, op, new_ps in inserted_para_expectations:
        item_ok = True
        for np_txt in new_ps:
            if np_txt not in acc_all_texts:
                errors.append(f"修订 {rid} ({op}) 新增段落未出现在全部接受后的文档中: {np_txt!r}")
                item_ok = False
        verified_details.append({"id": rid, "status": "applied", "ok": item_ok})

    return {
        "matches_expected": len(errors) == 0,
        "verified_items": verified_details,
        "errors": errors,
    }


def run_render_check(revised_docx: str | Path, out_dir: str | Path) -> dict[str, Any]:
    """可选调用 LibreOffice (soffice) 将修订版 .docx 转换为 PDF，检查渲染是否报错。"""
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if not exe:
        return {
            "executed": False,
            "ok": True,
            "pdf_path": None,
            "message": "未检测到 soffice，跳过可选 PDF 渲染检查",
        }

    rev_path = Path(revised_docx).resolve()
    out_p = Path(out_dir).resolve()
    out_p.mkdir(parents=True, exist_ok=True)

    # 使用独立临时 UserInstallation 避免并发或宿主 LibreOffice 锁冲突
    with tempfile.TemporaryDirectory(prefix="lo_profile_") as tmp_profile:
        profile_uri = Path(tmp_profile).resolve().as_uri()
        cmd = [
            exe,
            f"-env:UserInstallation={profile_uri}",
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(out_p),
            str(rev_path),
        ]
        try:
            proc = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=60,
            )
        except Exception as exc:
            return {
                "executed": True,
                "ok": False,
                "pdf_path": None,
                "message": f"调用 soffice 异常: {exc}",
            }

    expected_pdf = out_p / f"{rev_path.stem}.pdf"
    ok = proc.returncode == 0 and expected_pdf.is_file() and expected_pdf.stat().st_size > 0
    return {
        "executed": True,
        "ok": ok,
        "returncode": proc.returncode,
        "pdf_path": str(expected_pdf) if expected_pdf.is_file() else None,
        "message": "LibreOffice PDF 渲染转换成功" if ok else f"转换失败: {proc.stderr or proc.stdout}",
    }


def verify_docx(
    original_docx: str | Path,
    revised_docx: str | Path,
    revisions_json: str | Path,
    log_path: str | Path | None = None,
    report_path: str | Path | None = None,
    clean_out: str | Path | None = None,
    render_check_dir: str | Path | None = None,
    fail_on_semantic_warning: bool = False,
) -> dict[str, Any]:
    """执行完整校验流程并返回报告字典。"""
    orig_pkg = DocxPackage(original_docx)
    rev_pkg = DocxPackage(revised_docx)
    revisions_text, _ = read_text_auto(revisions_json)
    revisions_data = json.loads(revisions_text)

    orig_rev_ids, orig_com_ids = _collect_ids_from_pkg(orig_pkg)
    orig_had_comments_part = "word/comments.xml" in orig_pkg.raw_parts
    orig_settings = orig_pkg.get_xml("word/settings.xml")
    orig_had_track = (
        orig_settings is not None and orig_settings.find(w_tag("trackRevisions")) is not None
    )

    # 1. 模拟“全部拒绝”并与原文档逐段比对
    rejected_pkg = simulate_reject_all(revised_docx, orig_rev_ids, orig_com_ids)
    orig_blocks = iter_document_blocks(orig_pkg, include_headers_footers=True)
    rej_blocks = iter_document_blocks(rejected_pkg, include_headers_footers=True)
    semantic_warnings = _detect_existing_revision_conflicts(orig_blocks, revisions_data)

    reject_mismatches: list[dict[str, Any]] = []
    if len(orig_blocks) != len(rej_blocks):
        reject_mismatches.append(
            {
                "error": f"全部拒绝后段落块数量 ({len(rej_blocks)}) 与原文档 ({len(orig_blocks)}) 不一致"
            }
        )
    else:
        for idx, (ob, rb) in enumerate(zip(orig_blocks, rej_blocks)):
            if ob.text != rb.text or ob.location != rb.location:
                reject_mismatches.append(
                    {
                        "index": idx,
                        "block_id": ob.block_id,
                        "orig_text": ob.text,
                        "rejected_text": rb.text,
                    }
                )
    reject_all_ok = len(reject_mismatches) == 0

    # 2. 模拟“全部接受”并导出清洁版（若请求）
    accepted_pkg = simulate_accept_all(
        revised_docx,
        orig_rev_ids,
        orig_com_ids,
        orig_had_comments_part=orig_had_comments_part,
        orig_had_track_revisions=orig_had_track,
    )

    clean_out_ok = True
    clean_out_err: str | None = None
    if clean_out is not None:
        try:
            accepted_pkg.save(clean_out)
            clean_validation = validate_docx_file(clean_out, strict=True)
            if not clean_validation["valid"]:
                raise ValueError("；".join(clean_validation["errors"]))
        except Exception as exc:
            clean_out_ok = False
            clean_out_err = f"清洁版导出或打开失败: {exc}"

    # 读取或推导各修订状态
    status_map: dict[str, str] = {}
    failed_in_log = 0
    if log_path is not None and Path(log_path).is_file():
        log_text, _ = read_text_auto(log_path)
        log_data = json.loads(log_text)
        for item in log_data.get("items") or []:
            status_map[item["id"]] = item["status"]
            if item["status"] == "failed":
                failed_in_log += 1
    else:
        status_map = _infer_statuses_without_log(
            orig_pkg, revisions_data.get("revisions") or []
        )
        failed_in_log = sum(1 for v in status_map.values() if v == "failed")

    accept_res = _verify_accept_all_expectations(
        orig_pkg, rev_pkg, accepted_pkg, revisions_data, status_map
    )
    accept_all_ok = accept_res["matches_expected"] and (failed_in_log == 0)
    if failed_in_log > 0:
        accept_res["errors"].append(f"存在 {failed_in_log} 条状态为 failed 的修订未生效")

    # 3. 结构完整性检查
    struct_res = check_structure_integrity(revised_docx)
    if not clean_out_ok and clean_out_err:
        struct_res["valid"] = False
        struct_res["errors"].append(clean_out_err)

    # 4. 可选渲染检查
    render_res: dict[str, Any] | None = None
    render_ok = True
    if render_check_dir is not None:
        render_res = run_render_check(revised_docx, render_check_dir)
        render_ok = render_res["ok"]

    semantic_ok = not semantic_warnings or not fail_on_semantic_warning
    passed = bool(reject_all_ok and accept_all_ok and struct_res["valid"] and render_ok and semantic_ok)

    report: dict[str, Any] = {
        "passed": passed,
        "original_docx": str(original_docx),
        "revised_docx": str(revised_docx),
        "revisions_json": str(revisions_json),
        "clean_out": str(clean_out) if clean_out else None,
        "reject_all_matches_original": reject_all_ok,
        "accept_all_matches_expected": accept_all_ok,
        "structure_valid": struct_res["valid"],
        "semantic_review_required": bool(semantic_warnings),
        "semantic_warnings": semantic_warnings,
        "checks": {
            "reject_all": {
                "passed": reject_all_ok,
                "original_block_count": len(orig_blocks),
                "rejected_block_count": len(rej_blocks),
                "mismatches": reject_mismatches,
            },
            "accept_all": {
                "passed": accept_all_ok,
                "failed_revisions_count": failed_in_log,
                "verified_items": accept_res["verified_items"],
                "errors": accept_res["errors"],
            },
            "structure": struct_res,
            "semantic": {
                "passed": semantic_ok,
                "fail_on_warning": fail_on_semantic_warning,
                "review_required": bool(semantic_warnings),
                "warnings": semantic_warnings,
            },
            "render": render_res,
        },
    }

    if report_path is not None:
        rp = Path(report_path)
        rp.parent.mkdir(parents=True, exist_ok=True)
        rp.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    return report


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8")

    parser = argparse.ArgumentParser(
        description="校验修订版 .docx（全部拒绝=原文、全部接受=预期、OOXML结构完整性）并可输出清洁版"
    )
    parser.add_argument("original", help="原始合同 .docx 路径")
    parser.add_argument("revised", help="修订版合同 .docx 路径")
    parser.add_argument("revisions", help="revisions.json 路径")
    parser.add_argument("--log", default=None, help="docx_redline.py 输出的 apply_log.json 路径")
    parser.add_argument("--report", default=None, help="输出校验报告 verify_report.json 路径")
    parser.add_argument("--clean-out", default=None, help="输出全部接受修订后的清洁版 .docx 路径")
    parser.add_argument(
        "--render-check",
        default=None,
        help="可选：使用 LibreOffice (soffice) 转换 PDF 的输出目录以验证版面渲染",
    )
    parser.add_argument(
        "--fail-on-semantic-warning",
        action="store_true",
        help="检测到原文既有修订与新修订可能冲突时，以失败退出；默认仅写入报告告警",
    )
    args = parser.parse_args(argv)

    for label, p_str in [
        ("原始 docx", args.original),
        ("修订版 docx", args.revised),
        ("revisions.json", args.revisions),
    ]:
        if not Path(p_str).is_file():
            print(
                json.dumps({"error": f"{label}文件不存在: {p_str}"}, ensure_ascii=False),
                file=sys.stderr,
            )
            return 1

    try:
        report = verify_docx(
            original_docx=args.original,
            revised_docx=args.revised,
            revisions_json=args.revisions,
            log_path=args.log,
            report_path=args.report,
            clean_out=args.clean_out,
            render_check_dir=args.render_check,
            fail_on_semantic_warning=args.fail_on_semantic_warning,
        )
    except Exception as exc:
        print(json.dumps({"error": f"校验执行异常: {exc}"}, ensure_ascii=False), file=sys.stderr)
        return 1

    summary = {
        "passed": report["passed"],
        "reject_all_matches_original": report["reject_all_matches_original"],
        "accept_all_matches_expected": report["accept_all_matches_expected"],
        "structure_valid": report["structure_valid"],
        "clean_out": report["clean_out"],
    }
    if report["checks"]["render"] is not None:
        summary["render_ok"] = report["checks"]["render"]["ok"]

    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0 if report["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
