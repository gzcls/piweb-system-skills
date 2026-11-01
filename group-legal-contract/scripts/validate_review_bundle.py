#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""校验合同审核的结构化输出闭环。

该脚本把 ``review-result.json`` 作为审核问题台账，校验它与
``revisions.json``、``blocks.json`` 的一致性。它不判断法律结论，只阻止
结构不完整、修订编号漂移、锚点不存在或行内修订缺少原文匹配的文件进入
引用核验和 Word 写入阶段。
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

from lib.textio import read_text_auto  # noqa: E402
from lib.textnorm import locate_match, validate_revisions_data  # noqa: E402


SCHEMA_DIR = SCRIPT_DIR.parent / "assets" / "schemas"


def _load_json(path: str | Path) -> Any:
    p = Path(path).resolve()
    if not p.is_file():
        raise FileNotFoundError(f"JSON 文件不存在: {p}")
    text, _ = read_text_auto(p)
    try:
        return json.loads(text)
    except json.JSONDecodeError as exc:
        raise ValueError(f"JSON 解析失败 {p}: {exc}") from exc


def _schema_errors(data: Any, schema_path: Path, label: str) -> list[str]:
    try:
        import jsonschema  # type: ignore
    except ImportError as exc:  # pragma: no cover - check_env 将其列为必需依赖
        return [f"{label}: 缺少 jsonschema 依赖: {exc}"]
    schema = _load_json(schema_path)
    validator = jsonschema.Draft202012Validator(schema)
    errors: list[str] = []
    for err in sorted(validator.iter_errors(data), key=lambda e: list(e.path)):
        loc = ".".join(str(p) for p in err.path) or "<root>"
        errors.append(f"{label}.{loc}: {err.message}")
    return errors


def _blocks_map(blocks_path: str | Path | None) -> dict[str, dict[str, Any]]:
    if not blocks_path:
        return {}
    data = _load_json(blocks_path)
    blocks = data.get("blocks") if isinstance(data, dict) else None
    if not isinstance(blocks, list):
        raise ValueError("blocks.json 缺少 blocks 数组")
    result: dict[str, dict[str, Any]] = {}
    for block in blocks:
        if not isinstance(block, dict):
            continue
        block_id = str(block.get("block_id") or "")
        if block_id:
            result[block_id] = block
        # docx_redline 支持 seq_id 作为兼容锚点，校验阶段也同步支持。
        seq_id = str(block.get("seq_id") or "")
        if seq_id:
            result.setdefault(seq_id, block)
    return result


def validate_bundle(
    revisions_path: str | Path,
    review_result_path: str | Path | None = None,
    blocks_path: str | Path | None = None,
    *,
    require_matching_ids: bool = True,
    fail_on_warnings: bool = False,
) -> dict[str, Any]:
    """返回结构化校验结果；不通过时 ``valid`` 为 False。"""
    errors: list[str] = []
    warnings: list[str] = []
    revisions = _load_json(revisions_path)
    if not isinstance(revisions, dict):
        errors.append("revisions.json 必须是 JSON 对象")
        revisions = {}
    errors.extend(validate_revisions_data(revisions))

    review_result: dict[str, Any] | None = None
    if review_result_path:
        review_result = _load_json(review_result_path)
        if not isinstance(review_result, dict):
            errors.append("review-result.json 必须是 JSON 对象")
            review_result = {}
        else:
            errors.extend(
                _schema_errors(
                    review_result,
                    SCHEMA_DIR / "review-result.schema.json",
                    "review-result.json",
                )
            )

    revisions_list = revisions.get("revisions") or []
    revision_ids = [str(item.get("id") or "") for item in revisions_list if isinstance(item, dict)]
    if len(revision_ids) != len(set(revision_ids)):
        errors.append("revisions.json 存在重复修订编号")
    if any(not item_id for item_id in revision_ids):
        errors.append("revisions.json 存在空修订编号")

    issue_ids: list[str] = []
    issue_map: dict[str, dict[str, Any]] = {}
    if review_result is not None:
        issues = review_result.get("issues") or []
        for issue in issues:
            if not isinstance(issue, dict):
                continue
            issue_id = str(issue.get("id") or "")
            issue_ids.append(issue_id)
            if issue_id:
                issue_map[issue_id] = issue
        if len(issue_ids) != len(set(issue_ids)):
            errors.append("review-result.json 存在重复问题编号")
        if require_matching_ids and set(revision_ids) != set(issue_ids):
            missing_in_result = sorted(set(revision_ids) - set(issue_ids))
            missing_in_revisions = sorted(set(issue_ids) - set(revision_ids))
            if missing_in_result:
                errors.append(f"review-result.json 缺少修订编号: {', '.join(missing_in_result)}")
            if missing_in_revisions:
                errors.append(f"revisions.json 缺少问题编号: {', '.join(missing_in_revisions)}")

    blocks = _blocks_map(blocks_path)
    inline_ops = {"replace", "delete", "insert_text"}
    for item in revisions_list:
        if not isinstance(item, dict):
            continue
        item_id = str(item.get("id") or "")
        anchor = item.get("anchor") or {}
        block_id = str(anchor.get("block_id") or "")
        op = str(item.get("op") or "")
        if blocks and block_id not in blocks:
            errors.append(f"{item_id}: 找不到锚点 block_id={block_id}")
            continue
        block = blocks.get(block_id)
        match = anchor.get("match")
        if op in inline_ops and not isinstance(match, str):
            errors.append(f"{item_id}: {op} 操作必须提供 anchor.match")
        if op in {"replace", "insert_text"} and not str(item.get("new_text") or "").strip():
            errors.append(f"{item_id}: {op} 操作必须提供非空 new_text")
        if op in {"insert_paragraph_before", "insert_paragraph_after"}:
            paragraphs = item.get("new_paragraphs")
            if not isinstance(paragraphs, list) or not any(str(value).strip() for value in paragraphs):
                errors.append(f"{item_id}: {op} 操作必须提供非空 new_paragraphs")
        if op == "delete" and not isinstance(match, str):
            errors.append(f"{item_id}: delete 操作必须提供 anchor.match，以证明删除目标")
        if op in inline_ops and block and isinstance(match, str):
            span, reason = locate_match(str(block.get("text") or ""), match, anchor.get("occurrence"))
            if span is None:
                errors.append(f"{item_id}: anchor.match 无法在 {block_id} 中定位（{reason or '未找到'}）")

        issue = issue_map.get(item_id)
        if issue:
            issue_block = str(issue.get("block_id") or "")
            if issue_block and issue_block != block_id:
                errors.append(f"{item_id}: review-result.block_id 与 revisions.anchor.block_id 不一致")
            comment = item.get("comment") or {}
            if comment.get("risk") and issue.get("risk") and comment.get("risk") != issue.get("risk"):
                errors.append(f"{item_id}: comment.risk 与 review-result.issue.risk 不一致")

        # 既有修订段落上的插入不会自动删除原条款，交由 docx_verify 生成语义告警。
        if block and block.get("has_existing_revisions") and op in {
            "replace",
            "delete",
            "insert_text",
            "delete_paragraph",
            "insert_paragraph_before",
            "insert_paragraph_after",
        }:
            warnings.append(f"{item_id}: 锚点 {block_id} 含既有修订，需人工确认新旧条款是否冲突")

    if fail_on_warnings and warnings:
        errors.extend([f"严格门禁：{warning}" for warning in warnings])

    return {
        "valid": not errors,
        "errors": errors,
        "warnings": warnings,
        "blocked_by_warnings": bool(fail_on_warnings and warnings),
        "summary": {
            "revision_count": len(revision_ids),
            "issue_count": len(issue_ids),
            "anchor_count": len(blocks),
            "existing_revision_warnings": sum("含既有修订" in item for item in warnings),
            "warning_count": len(warnings),
        },
        "inputs": {
            "revisions": str(Path(revisions_path).resolve()),
            "review_result": str(Path(review_result_path).resolve()) if review_result_path else None,
            "blocks": str(Path(blocks_path).resolve()) if blocks_path else None,
        },
    }


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description="校验合同审核问题台账、修订指令和原文锚点")
    parser.add_argument("--revisions", required=True, help="revisions.json")
    parser.add_argument("--review-result", default=None, help="review-result.json")
    parser.add_argument("--blocks", default=None, help="docx_extract.py 生成的 blocks.json")
    parser.add_argument("--out", default=None, help="输出校验 JSON")
    parser.add_argument(
        "--fail-on-warning",
        action="store_true",
        help="将既有修订等警告提升为阻断错误（发布流水线默认启用）",
    )
    args = parser.parse_args(argv)
    try:
        report = validate_bundle(
            args.revisions,
            args.review_result,
            args.blocks,
            fail_on_warnings=args.fail_on_warning,
        )
    except Exception as exc:
        report = {"valid": False, "errors": [str(exc)], "warnings": [], "summary": {}}
    if args.out:
        out = Path(args.out).resolve()
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report["valid"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
