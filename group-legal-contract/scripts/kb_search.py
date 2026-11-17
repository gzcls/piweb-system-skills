#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""知识库关键词与标签检索工具（kb_search.py）。

用法示例：
  python scripts/kb_search.py --q "违约金 过高" --kind law --top 5
  python scripts/kb_search.py --q "预付式消费" --sector 03 --top 10
  python scripts/kb_search.py --contract-type lease --kind checklist
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

# 确保可以导入同目录下的 lib 包
SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from lib.kb import load_index, read_entry_text


def _split_keywords(queries: list[str]) -> list[str]:
    kws: list[str] = []
    for q in queries:
        for part in re.split(r"[\s,，、；;]+", q.strip()):
            if part and part not in kws:
                kws.append(part)
    return kws


def _match_filters(entry: dict, kind: str | None, sector: str | None, contract_type: str | None) -> bool:
    if kind and kind != "all" and entry.get("kind") != kind:
        return False
    if sector:
        sec_val = entry.get("sector") or entry.get("category") or ""
        if sector.lower() not in sec_val.lower():
            return False
    if contract_type:
        cts = entry.get("contract_types") or []
        if contract_type not in cts:
            return False
    return True


def search_kb(
    queries: list[str],
    kind: str | None = None,
    sector: str | None = None,
    contract_type: str | None = None,
    top: int = 10,
    max_chars: int = 280,
    max_per_file: int = 3,
    index_path: str | None = None,
) -> dict:
    idx = load_index(index_path)
    keywords = _split_keywords(queries)

    candidates = [
        e for e in idx.get("entries", [])
        if _match_filters(e, kind, sector, contract_type)
    ]

    # 无关键词时，仅返回条目元数据列表
    if not keywords:
        hits = []
        for e in candidates[:top]:
            hits.append({
                "id": e["id"],
                "kind": e["kind"],
                "title": e["title"],
                "path": e["path"],
                "sector": e.get("sector") or e.get("category"),
                "status": e.get("status"),
                "est_tokens": e.get("est_tokens"),
                "line": 1,
                "score": 1.0,
                "snippet": f"【{e['kind']}】{e['title']}",
            })
        return {
            "kb_snapshot": idx.get("kb_snapshot"),
            "keywords": [],
            "filters": {"kind": kind, "sector": sector, "contract_type": contract_type},
            "total_hits": len(hits),
            "hits": hits,
        }

    all_hits: list[dict] = []
    kw_Lowers = [k.lower() for k in keywords]

    for e in candidates:
        title_text = (e.get("title", "") + " " + " ".join(e.get("aliases", []))).lower()
        title_bonus = sum(2.0 for k in kw_Lowers if k in title_text)

        text = read_entry_text(e, index_path)
        lines = text.splitlines()

        file_hits: list[dict] = []
        last_taken_line = -10

        for idx_0, raw_line in enumerate(lines):
            line_low = raw_line.lower()
            matched_kws = [k for k in kw_Lowers if k in line_low]
            if not matched_kws:
                continue
            # 合并相邻行：预留前后文窗口
            if idx_0 - last_taken_line <= 2 and file_hits:
                continue
            start_l = max(0, idx_0 - 1)
            end_l = min(len(lines), idx_0 + 3)
            window_text = "\n".join(l.strip() for l in lines[start_l:end_l] if l.strip())
            window_low = window_text.lower()
            window_kws = [k for k in kw_Lowers if k in window_low]
            score = len(window_kws) * 3.0 + len(matched_kws) * 1.5 + title_bonus

            if len(window_text) > max_chars:
                window_text = window_text[:max_chars].rstrip() + "…"

            file_hits.append({
                "id": e["id"],
                "kind": e["kind"],
                "title": e["title"],
                "path": e["path"],
                "sector": e.get("sector") or e.get("category"),
                "status": e.get("status"),
                "line": idx_0 + 1,
                "score": round(score, 2),
                "matched_keywords": window_kws,
                "snippet": window_text,
            })
            last_taken_line = idx_0

        # 若仅标题命中而正文未逐行命中，也保留一条标题命中记录
        if not file_hits and title_bonus > 0:
            preview = "\n".join(l.strip() for l in lines[:5] if l.strip())
            if len(preview) > max_chars:
                preview = preview[:max_chars].rstrip() + "…"
            file_hits.append({
                "id": e["id"],
                "kind": e["kind"],
                "title": e["title"],
                "path": e["path"],
                "sector": e.get("sector") or e.get("category"),
                "status": e.get("status"),
                "line": 1,
                "score": round(title_bonus, 2),
                "matched_keywords": [k for k in kw_Lowers if k in title_text],
                "snippet": preview,
            })

        file_hits.sort(key=lambda x: (-x["score"], x["line"]))
        all_hits.extend(file_hits[:max_per_file])

    all_hits.sort(key=lambda x: (-x["score"], x["id"], x["line"]))
    top_hits = all_hits[:top]

    return {
        "kb_snapshot": idx.get("kb_snapshot"),
        "keywords": keywords,
        "filters": {"kind": kind, "sector": sector, "contract_type": contract_type},
        "total_hits": len(all_hits),
        "hits": top_hits,
    }


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8")

    parser = argparse.ArgumentParser(description="检索集团法务知识库（法规/审查清单/范本/板块专属库）")
    parser.add_argument("--q", "--query", dest="queries", action="append", default=[], help="检索关键词（可空格分隔或多次传入）")
    parser.add_argument("--kind", choices=["all", "law", "checklist", "template", "sector", "template-note"], default="all", help="限定文档类别")
    parser.add_argument("--sector", default=None, help="按板块或分类过滤（如 03、餐饮、01_民商事）")
    parser.add_argument("--contract-type", dest="contract_type", default=None, help="按合同类型代码过滤（如 sale、lease、tech）")
    parser.add_argument("--top", type=int, default=10, help="最多返回条数（默认 10）")
    parser.add_argument("--max-chars", type=int, default=280, help="每条片段最大字符数（默认 280）")
    parser.add_argument("--max-per-file", type=int, default=3, help="单文件最多返回片段数（默认 3）")
    parser.add_argument("--index", default=None, help="指定 kb_index.json 路径")
    parser.add_argument("--format", choices=["json", "text"], default="json", help="输出格式（默认 json）")
    args = parser.parse_args(argv)

    try:
        res = search_kb(
            queries=args.queries,
            kind=None if args.kind == "all" else args.kind,
            sector=args.sector,
            contract_type=args.contract_type,
            top=args.top,
            max_chars=args.max_chars,
            max_per_file=args.max_per_file,
            index_path=args.index,
        )
    except Exception as exc:
        print(json.dumps({"error": str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 1

    if args.format == "json":
        print(json.dumps(res, ensure_ascii=False, indent=2))
    else:
        print(f"知识库快照: {res['kb_snapshot']} | 关键词: {res['keywords']} | 命中总数: {res['total_hits']}")
        for idx_h, h in enumerate(res["hits"], start=1):
            print(f"\n[{idx_h}] {h['id']} ({h['kind']}) 《{h['title']}》 {h['path']}:{h['line']} [score={h['score']}]")
            print(f"    {h['snippet']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
