#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""按“法名/文号 + 条号”精确提取本地知识库中的法条原文（cite_law.py）。

用法示例：
  python scripts/cite_law.py --law 民法典 --article 585
  python scripts/cite_law.py --law "法释〔2023〕13号" --article "第64-66条"
  python scripts/cite_law.py --law 反不正当竞争法 --article 22 --paragraph 3
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from lib.cn_num import CN_NUM_RE, parse_num, int_to_cn
from lib.kb import (
    load_index,
    entry_by_id,
    resolve_law,
    find_articles,
    law_articles,
    split_paragraphs,
    LawTarget,
)
from lib.citations import parse_article_chain, ArticleRef


def _split_article_and_keyword_args(arg: str) -> tuple[str, list[str]]:
    """将条号表达式中的 `?关键词` 拆出，支持如 `584,585,?替代履行`。"""
    num_parts: list[str] = []
    kw_parts: list[str] = []
    for token in re.split(r"[,，、；;\s]+", arg.strip()):
        t = token.strip()
        if not t:
            continue
        if t.startswith(("?", "？")) and len(t) > 1:
            kw_parts.append(t[1:].strip())
        else:
            num_parts.append(t)
    return "、".join(num_parts), kw_parts


def _parse_article_arg(arg: str, default_para: int | None = None) -> list[ArticleRef]:
    """支持 '585', '第五百八十五条', '584-586', '143,522,523', '第22条第3款' 等写法。"""
    s = arg.strip()
    if not s:
        return []
    # 先尝试标准条文串解析
    normalized = s if "条" in s else "、".join(
        f"第{p.strip()}条" if not p.strip().startswith("第") else p.strip()
        for p in re.split(r"[,，、；;\s]+", s) if p.strip()
    )
    # 处理純数字区间如 584-586
    m_rng = re.fullmatch(rf"第?({CN_NUM_RE}|\d+)条?\s*[\-—~～至到]\s*第?({CN_NUM_RE}|\d+)条?", s)
    if m_rng:
        start_n = parse_num(m_rng.group(1))
        end_n = parse_num(m_rng.group(2))
        if start_n and end_n and start_n <= end_n and (end_n - start_n) <= 50:
            return [ArticleRef(num=n, paragraph=default_para, raw=s) for n in range(start_n, end_n + 1)]

    refs = parse_article_chain(normalized)
    if not refs:
        # 兜底：匹配“23之2”或单数字
        m_single = re.fullmatch(rf"第?(?P<num>{CN_NUM_RE}|\d+)条?(?:之(?P<suf>{CN_NUM_RE}|\d+))?", s)
        if m_single:
            num = parse_num(m_single.group("num"))
            suf = parse_num(m_single.group("suf")) if m_single.group("suf") else None
            if num:
                refs = [ArticleRef(num=num, suffix=suf, paragraph=default_para, raw=s)]
    if default_para is not None:
        for r in refs:
            if r.paragraph is None:
                r.paragraph = default_para
    return refs


def cite_law(
    law_query: str,
    article_str: str | None = None,
    paragraph: int | None = None,
    index_path: str | None = None,
) -> tuple[int, dict]:
    idx = load_index(index_path)

    # 支持直接传条目 ID（如 L0102）
    direct_entry = entry_by_id(law_query.strip(), str(index_path or ""))
    if direct_entry and direct_entry.get("kind") == "law":
        status = "ok"
        targets = [LawTarget(direct_entry["id"], None, "id", direct_entry.get("doc_kind") == "compilation")]
    else:
        status, targets = resolve_law(law_query, str(index_path or ""))

    if status == "not_in_kb":
        return 2, {
            "status": "out_of_kb",
            "status_cn": "库外法源·待人工核实",
            "query": {"law": law_query, "article": article_str, "paragraph": paragraph},
            "kb_snapshot": idx.get("kb_snapshot"),
            "message": f"本地知识库未收录《{law_query}》。引用时必须标注【库外法源·待人工核实】，严禁编造条文原文。",
            "articles": [],
        }

    if status == "ambiguous":
        cands = []
        for t in targets:
            e = entry_by_id(t.entry_id, str(index_path or ""))
            if e:
                cands.append({"id": e["id"], "title": e["title"], "component": t.component})
        return 3, {
            "status": "ambiguous_law",
            "status_cn": "法源名称存在歧义",
            "query": {"law": law_query, "article": article_str, "paragraph": paragraph},
            "kb_snapshot": idx.get("kb_snapshot"),
            "candidates": cands,
            "message": "法源简称命中多部法规，请使用更具体的名称、文号或条目 ID。",
            "articles": [],
        }

    primary_target = targets[0]
    primary_entry = entry_by_id(primary_target.entry_id, str(index_path or ""))

    if not article_str:
        return 0, {
            "status": "ok",
            "query": {"law": law_query},
            "kb_snapshot": idx.get("kb_snapshot"),
            "law": {
                "id": primary_entry["id"],
                "title": primary_entry["title"],
                "path": primary_entry["path"],
                "doc_type": primary_entry.get("doc_type"),
                "doc_kind": primary_entry.get("doc_kind"),
                "doc_nos": primary_entry.get("doc_nos", []),
                "status": primary_entry.get("status"),
                "effective_date": primary_entry.get("effective_date"),
                "article_count": primary_entry.get("article_count"),
            },
            "articles": [],
        }

    num_arg, kw_queries = _split_article_and_keyword_args(article_str)
    art_refs = _parse_article_arg(num_arg, default_para=paragraph) if num_arg else []
    if not art_refs and not kw_queries:
        return 4, {
            "status": "invalid_article",
            "status_cn": "无法解析条号",
            "query": {"law": law_query, "article": article_str},
            "message": f"无法解析条号参数：{article_str}",
            "articles": [],
        }

    out_articles: list[dict] = []
    missing_labels: list[str] = []
    nearby_hints: list[dict] = []

    for ref in art_refs:
        found = []
        used_target = primary_target
        for cand in targets:
            hits = find_articles(cand, ref.num, ref.suffix, index_path=index_path)
            if hits:
                found = hits
                used_target = cand
                break

        if not found:
            missing_labels.append(ref.label)
            # 自动探测前后相邻 2 条（应对修法条号偏移或邻近条号记忆偏差）
            for delta in (-2, -1, 1, 2):
                cand_num = ref.num + delta
                if cand_num <= 0:
                    continue
                for cand in targets:
                    near_hits = find_articles(cand, cand_num, None, index_path=index_path)
                    for nh in near_hits:
                        nearby_hints.append({
                            "for_missing": ref.label,
                            "label": nh.label,
                            "num": nh.num,
                            "preview": re.sub(r"\s+", " ", nh.text)[:160],
                        })
            continue

        used_entry = entry_by_id(used_target.entry_id, str(index_path or "")) or primary_entry
        for art in found:
            paras = split_paragraphs(art.text)
            selected_para_text = None
            para_warning = None
            if ref.paragraph is not None:
                if 1 <= ref.paragraph <= len(paras):
                    selected_para_text = paras[ref.paragraph - 1]
                else:
                    para_warning = f"该条共切分出 {len(paras)} 款，未找到第 {ref.paragraph} 款，已返回整条原文"

            out_articles.append({
                "law_id": used_entry["id"],
                "law_title": used_entry["title"],
                "path": used_entry["path"],
                "doc_nos": used_entry.get("doc_nos", []),
                "law_status": used_entry.get("status"),
                "effective_date": used_entry.get("effective_date"),
                "excerpt": used_target.excerpt,
                "component": art.component,
                "num": art.num,
                "suffix": art.suffix,
                "label": art.label,
                "title": art.title,
                "line": art.line,
                "paragraph_count": len(paras),
                "requested_paragraph": ref.paragraph,
                "paragraph_text": selected_para_text,
                "paragraph_warning": para_warning,
                "text": art.text,
            })

    # 支持在条号参数中直接用 `?关键词` 在该法规内按关键词检索匹配条文（最多返回前 3 条）
    if kw_queries:
        seen_art_nums = {(a["law_id"], a["label"]) for a in out_articles}
        for kw in kw_queries:
            matched_kw = 0
            for cand in targets:
                cand_entry = entry_by_id(cand.entry_id, str(index_path or "")) or primary_entry
                all_arts, _ = law_articles(cand.entry_id, index_path=index_path)
                for art in all_arts:
                    if cand.component and art.component and not (
                        art.component == cand.component or art.component.startswith(cand.component + ".")
                    ):
                        continue
                    if kw in art.text or (art.title and kw in art.title):
                        key = (cand_entry["id"], art.label)
                        if key not in seen_art_nums:
                            seen_art_nums.add(key)
                            paras = split_paragraphs(art.text)
                            out_articles.append({
                                "law_id": cand_entry["id"],
                                "law_title": cand_entry["title"],
                                "path": cand_entry["path"],
                                "doc_nos": cand_entry.get("doc_nos", []),
                                "law_status": cand_entry.get("status"),
                                "effective_date": cand_entry.get("effective_date"),
                                "excerpt": cand.excerpt,
                                "component": art.component,
                                "num": art.num,
                                "suffix": art.suffix,
                                "label": art.label,
                                "title": art.title,
                                "line": art.line,
                                "paragraph_count": len(paras),
                                "requested_paragraph": None,
                                "paragraph_text": None,
                                "paragraph_warning": None,
                                "matched_keyword": kw,
                                "text": art.text,
                            })
                        matched_kw += 1
                        if matched_kw >= 3:
                            break
                if matched_kw > 0:
                    break
            if matched_kw == 0:
                missing_labels.append(f"?{kw}")

    exit_code = 0 if (out_articles and not missing_labels) else 4
    status_str = "ok" if exit_code == 0 else ("partial" if out_articles else "article_not_found")
    warnings: list[str] = []
    if primary_target.excerpt:
        warnings.append("当前法源来自专题汇编节选文件，非单行法全文。")
    if any(k in (primary_entry.get("status") or "") for k in ("待核实", "失效", "废止")):
        warnings.append(f"当前法源台账状态为【{primary_entry.get('status')}】，引用时请显式提示。")
    if missing_labels:
        if primary_target.excerpt:
            warnings.append(f"汇编节选未收录 {', '.join(missing_labels)}，引用时须标注【库外法源·待人工核实】。")
        else:
            warnings.append(
                f"在《{primary_entry['title']}》中未找到 {', '.join(missing_labels)}；"
                "严禁编造条号，请检查是否为修法前旧条号或参考 nearby_articles_hint。"
            )

    payload_out: dict = {
        "status": status_str,
        "kb_snapshot": idx.get("kb_snapshot"),
        "query": {"law": law_query, "article": article_str, "paragraph": paragraph},
        "law": {
            "id": primary_entry["id"],
            "title": primary_entry["title"],
            "path": primary_entry["path"],
            "doc_type": primary_entry.get("doc_type"),
            "doc_kind": primary_entry.get("doc_kind"),
            "doc_nos": primary_entry.get("doc_nos", []),
            "status": primary_entry.get("status"),
            "effective_date": primary_entry.get("effective_date"),
            "article_count": primary_entry.get("article_count"),
        },
        "missing_articles": missing_labels,
        "warnings": warnings,
        "articles": out_articles,
    }
    if nearby_hints:
        payload_out["nearby_articles_hint"] = nearby_hints[:8]
    return exit_code, payload_out


def _parse_batch_spec(raw_specs: list[str]) -> list[dict]:
    """解析批量法条查询参数。

    支持：
    1. 内联表达式：`民法典:470,509,584-586;反不正当竞争法:8,15`
    2. JSON 文件或 `-`（标准输入）：
       - `[{"law": "民法典", "article": "584,585"}, ...]`
       - `{"民法典": "470,509", "反不正当竞争法": ["8", "15"]}`
    3. 文本文件：每行或分号分隔的 `法名:条号列表`
    """
    chunks: list[str] = []
    for spec in raw_specs:
        s = spec.strip()
        if not s:
            continue
        if s == "-":
            chunks.append(sys.stdin.read())
            continue
        cand_path = Path(s)
        if cand_path.is_file():
            chunks.append(cand_path.read_text(encoding="utf-8-sig"))
        else:
            chunks.append(s)

    items: list[dict] = []
    for chunk in chunks:
        text = chunk.strip()
        if not text:
            continue
        if text[0] in "[{":
            data = json.loads(text)
            if isinstance(data, dict):
                if "items" in data and isinstance(data["items"], list):
                    data = data["items"]
                else:
                    for law_k, art_v in data.items():
                        if isinstance(art_v, list):
                            art_str = ",".join(str(x) for x in art_v)
                        else:
                            art_str = str(art_v) if art_v is not None else None
                        items.append({"law": str(law_k), "article": art_str, "paragraph": None})
                    continue
            if isinstance(data, list):
                for entry in data:
                    if isinstance(entry, dict) and entry.get("law"):
                        art_v = entry.get("article") or entry.get("articles")
                        if isinstance(art_v, list):
                            art_str = ",".join(str(x) for x in art_v)
                        else:
                            art_str = str(art_v) if art_v is not None else None
                        para_raw = entry.get("paragraph")
                        para_num = (
                            parse_num(re.sub(r"[第款\s]", "", str(para_raw)))
                            if para_raw is not None
                            else None
                        )
                        items.append({
                            "law": str(entry["law"]),
                            "article": art_str,
                            "paragraph": para_num,
                        })
                continue

        for part in re.split(r"[;；\r\n]+", text):
            part = part.strip()
            if not part or part.startswith("#"):
                continue
            m = re.match(r"^([^:：]+?)\s*[:：]\s*(.+)$", part)
            if m:
                items.append({
                    "law": m.group(1).strip(),
                    "article": m.group(2).strip(),
                    "paragraph": None,
                })
            else:
                items.append({"law": part, "article": None, "paragraph": None})
    return items


def cite_law_batch(
    specs: list[str] | list[dict],
    index_path: str | None = None,
) -> tuple[int, dict]:
    """一次核取多部法规的多个条号，减少工具往返次数。"""
    idx = load_index(index_path)
    if specs and isinstance(specs[0], dict):
        parsed_items = specs  # type: ignore[assignment]
    else:
        parsed_items = _parse_batch_spec([str(x) for x in specs])

    if not parsed_items:
        return 4, {
            "status": "invalid_batch",
            "kb_snapshot": idx.get("kb_snapshot"),
            "message": "批量查询参数为空",
            "results": [],
        }

    results: list[dict] = []
    codes: list[int] = []
    total_articles = 0
    ok_count = 0
    out_of_kb_count = 0
    ambiguous_count = 0
    missing_count = 0

    for item in parsed_items:
        code, payload = cite_law(
            law_query=item["law"],
            article_str=item.get("article"),
            paragraph=item.get("paragraph"),
            index_path=index_path,
        )
        codes.append(code)
        results.append(payload)
        arts = payload.get("articles") or []
        total_articles += len(arts)
        st = payload.get("status")
        if st == "ok":
            ok_count += 1
        elif st == "out_of_kb":
            out_of_kb_count += 1
        elif st == "ambiguous_law":
            ambiguous_count += 1
        else:
            missing_count += 1

    hard_errors = [c for c in codes if c not in (0, 2)]
    if hard_errors:
        exit_code = hard_errors[0]
        status_str = "partial" if (ok_count > 0 or total_articles > 0) else "failed"
    elif out_of_kb_count > 0:
        exit_code = 0
        status_str = "ok_with_out_of_kb"
    else:
        exit_code = 0
        status_str = "ok"

    return exit_code, {
        "status": status_str,
        "kb_snapshot": idx.get("kb_snapshot"),
        "summary": {
            "requested_items": len(parsed_items),
            "ok_items": ok_count,
            "out_of_kb_items": out_of_kb_count,
            "ambiguous_items": ambiguous_count,
            "missing_or_partial_items": missing_count,
            "total_articles": total_articles,
        },
        "results": results,
    }


def _print_single_text(payload: dict) -> None:
    law_info = payload.get("law")
    if law_info:
        print(
            f"【{law_info['id']}】《{law_info['title']}》 "
            f"（状态：{law_info.get('status')}｜施行日期：{law_info.get('effective_date') or '见正文'}）"
        )
    for w in payload.get("warnings", []):
        print(f"[提示] {w}")
    if payload.get("message"):
        print(payload["message"])
    for art in payload.get("articles", []):
        print(f"\n--- {art['label']} ({art['path']}:{art['line']}) ---")
        if art.get("paragraph_text"):
            print(f"[第{art['requested_paragraph']}款] {art['paragraph_text']}")
        else:
            print(art["text"])
        if art.get("paragraph_warning"):
            print(f"[款号提示] {art['paragraph_warning']}")


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8")

    parser = argparse.ArgumentParser(description="按法名/文号与条号提取本地知识库法条原文（支持单条与批量）")
    parser.add_argument("--law", default=None, help="法规名称、简称、文号（如 法释〔2023〕13号）或 ID（如 L0102）")
    parser.add_argument("--article", default=None, help="条号（如 585、第五百八十五条、584-586、143,522,523）")
    parser.add_argument("--paragraph", default=None, help="款号（如 2 或 第二款）")
    parser.add_argument(
        "--batch",
        action="append",
        default=[],
        help="批量查询表达式或文件路径（如 '民法典:470,509,585;反不正当竞争法:8,15' 或 JSON/文本文件）",
    )
    parser.add_argument("--out", default=None, help="可选：将 JSON 结果写入指定文件")
    parser.add_argument("--index", default=None, help="指定 kb_index.json 路径")
    parser.add_argument("--format", choices=["json", "text"], default="json", help="输出格式（默认 json）")
    args = parser.parse_args(argv)

    if not args.law and not args.batch:
        parser.error("必须提供 --law 或 --batch 参数")

    try:
        if args.batch:
            code, payload = cite_law_batch(args.batch, index_path=args.index)
        else:
            para_num = parse_num(re.sub(r"[第款\s]", "", args.paragraph)) if args.paragraph else None
            code, payload = cite_law(
                law_query=args.law,
                article_str=args.article,
                paragraph=para_num,
                index_path=args.index,
            )
    except Exception as exc:
        print(json.dumps({"error": str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 1

    if args.out:
        out_p = Path(args.out).resolve()
        out_p.parent.mkdir(parents=True, exist_ok=True)
        out_p.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    if args.format == "json":
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    else:
        if args.batch:
            s = payload.get("summary") or {}
            print(
                f"=== 批量法条核取（请求 {s.get('requested_items', 0)} 项｜"
                f"命中条文 {s.get('total_articles', 0)} 条｜库外 {s.get('out_of_kb_items', 0)} 项） ==="
            )
            for idx_item, item_payload in enumerate(payload.get("results", []), 1):
                print(f"\n[{idx_item}] 查询：{item_payload.get('query')}")
                _print_single_text(item_payload)
        else:
            _print_single_text(payload)
    return code


if __name__ == "__main__":
    raise SystemExit(main())
