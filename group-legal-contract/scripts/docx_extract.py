#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""从 .docx 合同中提取带稳定锚点的结构化段落块 (blocks.json) 与紧凑可读文本 (contract.txt)。

用法:
  python scripts/docx_extract.py <input.docx> --out-dir <dir>
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from lib.ooxml import DocxPackage, iter_document_blocks


def extract_docx(
    input_docx: str | Path,
    out_dir: str | Path | None = None,
) -> dict[str, Any]:
    """解析 .docx 并生成结构化块数据与紧凑文本。"""
    in_path = Path(input_docx)
    pkg = DocxPackage(in_path)
    records = iter_document_blocks(pkg, include_headers_footers=True)

    block_dicts = [r.to_dict() for r in records]
    doc_has_rev = any(b["has_existing_revisions"] for b in block_dicts)
    doc_has_com = any(b["has_comments"] for b in block_dicts)

    # 生成 contract.txt 行：每行格式 [p-0153|第八条] 文本
    # 空段落（无文本且无编号）在 contract.txt 中跳过以保持紧凑，blocks.json 完整保留
    txt_lines: list[str] = []
    for b in block_dicts:
        text_val = b["text"]
        label_val = b["list_label"]
        if not text_val.strip() and not label_val.strip():
            continue
        # 若自动编号不在段落文本开头，可在展示文本中体现或放在 | 右侧的 clause_ref
        clause_tag = b["clause_ref"] or label_val or ""
        # 将段落内换行压平为空格或 \\n 符号以保证 contract.txt 单段单行格式
        single_line_text = text_val.replace("\r\n", "↵").replace("\n", "↵").replace("\r", "↵")
        txt_lines.append(f"[{b['block_id']}|{clause_tag}] {single_line_text}")

    contract_txt = "\n".join(txt_lines) + ("\n" if txt_lines else "")

    result: dict[str, Any] = {
        "source": str(in_path),
        "block_count": len(block_dicts),
        "has_existing_revisions": doc_has_rev,
        "has_comments": doc_has_com,
        "blocks": block_dicts,
    }

    if out_dir is not None:
        out_p = Path(out_dir)
        out_p.mkdir(parents=True, exist_ok=True)
        blocks_json_path = out_p / "blocks.json"
        contract_txt_path = out_p / "contract.txt"
        blocks_json_path.write_text(
            json.dumps(result, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        contract_txt_path.write_text(contract_txt, encoding="utf-8")
        result["blocks_json_path"] = str(blocks_json_path)
        result["contract_txt_path"] = str(contract_txt_path)

    return result


import re


def format_table_blocks(
    blocks: list[dict[str, Any]],
    table_ids: list[str],
) -> list[str]:
    """按表编号（如 ['t33'] 或 ['33']）展开表格的所有非空单元格行。"""
    normalized: set[str] = set()
    for raw in table_ids:
        for token in re.split(r"[,，\s]+", str(raw).strip()):
            if not token:
                continue
            t = token.lower()
            if not t.startswith("t"):
                t = f"t{t}"
            normalized.add(t)

    lines: list[str] = []
    for b in blocks:
        bid = str(b.get("block_id") or "")
        m = re.match(r"^(t\d+)-", bid)
        if not m or m.group(1) not in normalized:
            continue
        text_val = str(b.get("text") or "")
        label_val = str(b.get("list_label") or "")
        if not text_val.strip() and not label_val.strip():
            continue
        clause_tag = b.get("clause_ref") or label_val or ""
        single_line_text = text_val.replace("\r\n", "↵").replace("\n", "↵").replace("\r", "↵")
        lines.append(f"[{bid}|{clause_tag}] {single_line_text}")
    return lines


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8")

    parser = argparse.ArgumentParser(
        description="提取 .docx 合同段落锚点结构 (blocks.json) 与紧凑文本 (contract.txt)"
    )
    parser.add_argument("input", help="待提取的输入 .docx 文件路径（使用 --table 时也可直接传 blocks.json）")
    parser.add_argument("--out-dir", default=None, help="输出目录（写入 blocks.json 和 contract.txt）")
    parser.add_argument(
        "--table",
        action="append",
        default=[],
        help="按需展开指定表格（如 --table t33 或 --table t3,t6），适用于阅读压缩视图后展开特定表格",
    )
    args = parser.parse_args(argv)

    in_path = Path(args.input)
    if not in_path.is_file():
        print(
            json.dumps({"error": f"输入文件不存在: {in_path}"}, ensure_ascii=False),
            file=sys.stderr,
        )
        return 1

    if args.table:
        try:
            if in_path.suffix.lower() == ".json":
                data = json.loads(in_path.read_text(encoding="utf-8"))
                blocks = data.get("blocks") or []
            else:
                res = extract_docx(in_path, args.out_dir)
                blocks = res.get("blocks") or []
            lines = format_table_blocks(blocks, args.table)
        except Exception as exc:
            print(json.dumps({"error": f"展开表格失败: {exc}"}, ensure_ascii=False), file=sys.stderr)
            return 1
        print("\n".join(lines))
        return 0

    if not args.out_dir:
        parser.error("未指定 --table 时必须提供 --out-dir")

    try:
        res = extract_docx(in_path, args.out_dir)
    except Exception as exc:
        print(
            json.dumps({"error": f"提取失败: {exc}"}, ensure_ascii=False),
            file=sys.stderr,
        )
        return 1

    if res.get("has_existing_revisions"):
        print(
            "【警告】检测到输入文档已包含未接受的修订痕迹（w:ins/w:del），已按“当前接受视图”提取文本。",
            file=sys.stdout,
        )

    summary = {
        "source": res["source"],
        "block_count": res["block_count"],
        "has_existing_revisions": res["has_existing_revisions"],
        "has_comments": res["has_comments"],
        "blocks_json": res.get("blocks_json_path"),
        "contract_txt": res.get("contract_txt_path"),
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
