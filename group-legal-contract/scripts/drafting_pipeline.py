#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""合同起草/模板交付流水线。

该模块把「起草请求校验、条件必填、模板路由、Markdown lint、法条核验、
双版本 DOCX 渲染以及结构/视觉门禁」收敛成一个可重复执行的入口。它只
负责起草产物，不改变审核流水线的输入或输出。

默认允许带待确认占位符的工作稿生成，但会将结果标记为 ``manual_review``；
使用 ``--require-complete`` 或 ``--strict`` 时，签署版仍有占位符会阻断发布。
视觉检查依赖可选的 LibreOffice/PyMuPDF。依赖缺失会保留 ``skipped`` 信号，
不会伪造视觉通过。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import sys
import tempfile
import uuid
import zipfile
from pathlib import Path
from typing import Any, Iterable

SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_ROOT = SCRIPT_DIR.parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from docx_visual_verify import inspect_docx
from contract_review_pipeline import _publish_staged_tree, _recover_publish_journals
from lib.ooxml import validate_docx_file
from md_to_docx import convert_md_to_docx, scan_signing_docx, strip_docx_comments
from verify_citations import verify_file

try:
    import jsonschema  # type: ignore
except ImportError:  # pragma: no cover - minimal installations still get useful lint
    jsonschema = None


REQUEST_SCHEMA = SKILL_ROOT / "assets" / "schemas" / "draft-request.schema.json"
ROUTE_MAP = SKILL_ROOT / "references" / "routing" / "contract-type-map.md"
DEFAULT_REFERENCE = SKILL_ROOT / "assets" / "docx" / "reference.docx"
TEMPLATE_CATALOG = SKILL_ROOT / "assets" / "template_catalog.json"

PLACEHOLDER_RE = re.compile(
    r"\[(?:●|填写说明|待填写|待确认|待核实|待集团填写|待补充|需补充|请填写|请确认|TODO|FIXME)(?:[^\]]*)\]",
    re.IGNORECASE,
)
UNRESOLVED_TEMPLATE_RE = re.compile(r"\{\{\s*[^{}]+\s*\}\}|\$\{[^{}]+\}")
INTERNAL_MARKER_RE = re.compile(
    r"(?:起草批注|谈判批注|风险批注|集团法务|内部(?:使用)?|TODO|FIXME|待确认|待核实|版本\s*[AB]|签署前)",
    re.IGNORECASE,
)
# 签署版扫描只拦截确定属于内部批注/版本控制的文本。普通合同中出现
# “签署前”“待确认”等业务语句时，仍由占位符 lint 给出人工复核提示，
# 避免把合法的履行条件误判为内部内容。
SIGNING_INTERNAL_RE = re.compile(
    r"(?:起草批注|谈判批注|风险批注|集团法务|内部(?:使用)?|\bTODO\b|\bFIXME\b|版本\s*[AB])",
    re.IGNORECASE,
)
HEADING_RE = re.compile(r"^\s*(#{1,6})\s+(.+?)\s*$")
CLAUSE_HEADING_RE = re.compile(r"^\s*(?:#{1,6}\s+)?([一二三四五六七八九十百千万\d]+)[、.．]\s*(.*)$")

# 商业字段只在确实存在交易履行内容的合同路由中强制要求。
REQUIRED_BY_ROUTE: dict[str, tuple[str, ...]] = {
    "sale": ("commercial_terms.total_amount", "commercial_terms.payment_schedule", "commercial_terms.delivery_and_acceptance"),
    "medical-procurement": ("commercial_terms.total_amount", "commercial_terms.payment_schedule", "commercial_terms.delivery_and_acceptance"),
    "food-procurement": ("commercial_terms.total_amount", "commercial_terms.payment_schedule", "commercial_terms.delivery_and_acceptance"),
    "commodity-trade": ("commercial_terms.total_amount", "commercial_terms.payment_schedule"),
    "intl-trade": ("commercial_terms.total_amount", "commercial_terms.payment_schedule", "commercial_terms.delivery_and_acceptance"),
    "tech": ("commercial_terms.term", "commercial_terms.ip_ownership"),
    "lease": ("commercial_terms.term", "commercial_terms.total_amount", "commercial_terms.payment_schedule"),
    "property-mgmt": ("commercial_terms.term", "commercial_terms.total_amount", "commercial_terms.payment_schedule"),
    "construction": ("commercial_terms.total_amount", "commercial_terms.term", "commercial_terms.delivery_and_acceptance"),
    "epc": ("commercial_terms.total_amount", "commercial_terms.term", "commercial_terms.delivery_and_acceptance"),
    "data": ("commercial_terms.term", "commercial_terms.ip_ownership"),
    "oem": ("commercial_terms.total_amount", "commercial_terms.delivery_and_acceptance"),
    "outsourcing": ("commercial_terms.term", "commercial_terms.delivery_and_acceptance"),
}


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _write_json_atomic(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(f".{path.name}.{uuid.uuid4().hex}.tmp")
    tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    os.replace(tmp, path)


def _get_path(data: dict[str, Any], dotted: str) -> Any:
    value: Any = data
    for part in dotted.split("."):
        if not isinstance(value, dict):
            return None
        value = value.get(part)
    return value


def _nonempty(value: Any) -> bool:
    if value is None or isinstance(value, bool):
        return False
    if isinstance(value, str):
        return bool(value.strip())
    if isinstance(value, (list, dict)):
        return bool(value)
    return True


def load_routes(path: Path = ROUTE_MAP) -> dict[str, dict[str, Any]]:
    """从生成的路由 Markdown 中提取类型名称、范本和核心法规 ID。"""
    routes: dict[str, dict[str, Any]] = {}
    if not path.is_file():
        return routes
    row_re = re.compile(r"^\|\s*`([^`]+)`\s*\|\s*\*\*(.*?)\*\*\s*\|(.+)$")
    for line in path.read_text(encoding="utf-8", errors="replace").splitlines():
        match = row_re.match(line.strip())
        if not match:
            continue
        route_id, title, remainder = match.groups()
        template_ids = sorted(set(re.findall(r"\bT\d{4}\b", remainder)))
        law_ids = sorted(set(re.findall(r"\bL\d{4}\b", remainder)))
        routes[route_id] = {"title": title.strip(), "template_ids": template_ids, "law_ids": law_ids}
    return routes


def _asset_root() -> Path:
    """Prefer source assets, then the last validated dist package."""
    if (SKILL_ROOT / "assets" / "template_catalog.json").is_file():
        return SKILL_ROOT
    dist_root = SKILL_ROOT.parent.parent / "dist" / SKILL_ROOT.name
    return dist_root if (dist_root / "assets" / "template_catalog.json").is_file() else SKILL_ROOT


def load_template_catalog(path: Path = TEMPLATE_CATALOG) -> dict[str, Any]:
    root = _asset_root()
    candidate = path if path.is_file() else root / "assets" / "template_catalog.json"
    if not candidate.is_file():
        return {"schema_version": 1, "templates": [], "path": str(candidate)}
    data = json.loads(candidate.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError(f"模板目录根节点必须是对象: {candidate}")
    templates = data.get("templates")
    if not isinstance(templates, list):
        raise ValueError(f"模板目录 templates 必须是数组: {candidate}")
    for entry in templates:
        if not isinstance(entry, dict) or not entry.get("id") or not entry.get("path"):
            raise ValueError(f"模板目录存在非法条目: {entry!r}")
        template_path = (candidate.parent.parent / str(entry["path"])).resolve()
        if not template_path.is_file():
            raise ValueError(f"模板目录指向不存在的文件: {entry.get('id')}")
        expected = str(entry.get("sha256") or "")
        if expected and _sha256(template_path) != expected:
            raise ValueError(f"模板文件 SHA256 不匹配: {entry.get('id')}")
    data["path"] = str(candidate.resolve())
    data["sha256"] = _sha256(candidate)
    return data


def select_template(request: dict[str, Any], route: dict[str, Any] | None) -> tuple[Path | None, dict[str, Any] | None]:
    catalog = load_template_catalog()
    entries = {str(item.get("id")): item for item in catalog.get("templates", []) if isinstance(item, dict)}
    requested = str(request.get("template_id") or "").strip()
    candidates = [requested] if requested else list((route or {}).get("template_ids") or [])
    for template_id in candidates:
        item = entries.get(template_id)
        if not item:
            continue
        path = Path(catalog["path"]).parent.parent / str(item["path"])
        return path.resolve(), item
    return None, None


def _render_route_skeleton(request: dict[str, Any]) -> str:
    """Create a transparent fallback when a route has no standard template."""
    terms = request.get("commercial_terms") or {}
    party_a = (request.get("party_a") or {}).get("name") or "[●]"
    party_b = (request.get("party_b") or {}).get("name") or "[●]"
    rows = [
        ("合同标的", terms.get("subject_matter") or "[●]"),
        ("合同金额", terms.get("total_amount") or "[●]"),
        ("付款安排", terms.get("payment_schedule") or "[●]"),
        ("期限", terms.get("term") or "[●]"),
        ("交付与验收", terms.get("delivery_and_acceptance") or "[●]"),
        ("知识产权", terms.get("ip_ownership") or "[●]"),
        ("争议解决", terms.get("dispute_resolution") or "[●]"),
    ]
    lines = [
        f"# {request.get('contract_type') or '合同'}（组合起草工作稿）", "",
        f"甲方：{party_a}", f"乙方：{party_b}", "",
        "> 本工作稿由通用结构自动生成，当前路由没有集团标准范本；请法务逐条确认并补充业务字段。", "",
        "## 一、交易要素", "", "| 要素 | 内容 |", "| --- | --- |",
    ]
    lines.extend(f"| {label} | {value} |" for label, value in rows)
    lines.extend([
        "", "## 二、通用履约与违约条款", "",
        "双方应按已确认的交易要素履行义务；交付、验收、付款、保密、知识产权、数据合规、违约责任及解除条件应以书面附件进一步明确。", "",
        "## 三、签署信息", "", "甲方（盖章）：[●]    乙方（盖章）：[●]", "签署日期：[●]年[●]月[●]日", "",
    ])
    return "\n".join(lines)


def validate_request(
    request: dict[str, Any],
    schema_path: Path = REQUEST_SCHEMA,
    *,
    require_complete: bool = False,
) -> dict[str, Any]:
    errors: list[str] = []
    warnings: list[str] = []
    schema_errors: list[dict[str, Any]] = []
    if not isinstance(request, dict):
        return {"valid": False, "errors": ["draft_request 必须是 JSON 对象"], "warnings": [], "missing_fields": [], "schema_errors": []}
    if jsonschema is None:
        warnings.append("未安装 jsonschema，已跳过 Draft 2020-12 schema 校验")
    elif schema_path.is_file():
        try:
            schema = json.loads(schema_path.read_text(encoding="utf-8"))
            validator = jsonschema.Draft202012Validator(schema)
            for error in sorted(validator.iter_errors(request), key=lambda e: list(e.path)):
                location = ".".join(str(x) for x in error.path) or "$"
                schema_errors.append({"path": location, "message": error.message})
                errors.append(f"{location}: {error.message}")
        except Exception as exc:
            errors.append(f"读取/执行请求 schema 失败: {exc}")
    else:
        errors.append(f"请求 schema 不存在: {schema_path}")

    contract_type = str(request.get("contract_type") or "").strip()
    our_position = str(request.get("our_position") or "").strip()
    if not contract_type:
        errors.append("contract_type 不能为空")
    if not our_position:
        errors.append("our_position 不能为空")

    # 所有合同起草都需要可识别的双方、标的和争议解决条款，
    # 以免渲染出无法签署的空模板。
    universal = ["party_a.name", "party_b.name", "commercial_terms.subject_matter", "commercial_terms.dispute_resolution"]
    required = list(universal) + list(REQUIRED_BY_ROUTE.get(contract_type, ()))
    for field in required:
        if not _nonempty(_get_path(request, field)):
            message = f"条件必填字段缺失: {field}"
            (errors if require_complete else warnings).append(message)

    # JSON Schema 的原始版本兼容历史请求，未对可选字符串统一设置
    # minLength，也未关闭额外字段。这里把空字符串和未知字段显式报告，
    # 供 --require-complete/--strict 决定是否阻断，而不破坏旧请求的读取。
    known_top_level = {
        "contract_type", "sector", "our_position", "party_a", "party_b",
        "commercial_terms", "special_requirements", "template_id",
    }
    for key in sorted(set(request) - known_top_level):
        warnings.append(f"请求包含 schema 未声明字段: {key}")
    for location, value in _walk_strings(request):
        if not value.strip():
            warnings.append(f"请求字段为空字符串: {location}")

    routes = load_routes()
    route = routes.get(contract_type)
    if not route:
        warnings.append(f"contract_type 未在路由表中命中: {contract_type or '（空）'}")
    template_catalog: dict[str, Any] = {"templates": []}
    try:
        template_catalog = load_template_catalog()
        available = {str(item.get("id")) for item in template_catalog.get("templates", [])}
        missing_templates = sorted(set((route or {}).get("template_ids", [])) - available)
        if missing_templates:
            warnings.append("路由引用的模板不在 template_catalog.json: " + ", ".join(missing_templates))
    except Exception as exc:
        warnings.append(f"模板目录校验失败: {exc}")
    return {
        "valid": not errors,
        "errors": errors,
        "warnings": warnings,
        "missing_fields": [
            item.split(": ", 1)[1]
            for item in [*errors, *warnings]
            if item.startswith("条件必填字段缺失: ")
        ],
        "schema_errors": schema_errors,
        "route": route,
        "template_catalog": {"path": template_catalog.get("path"), "sha256": template_catalog.get("sha256"), "template_count": len(template_catalog.get("templates", []))},
    }


def _walk_strings(value: Any, prefix: str = "$") -> Iterable[tuple[str, str]]:
    """遍历请求中的字符串，以便发现 schema 未约束的空字符串。"""
    if isinstance(value, str):
        yield prefix, value
    elif isinstance(value, dict):
        for key, child in value.items():
            yield from _walk_strings(child, f"{prefix}.{key}")
    elif isinstance(value, list):
        for idx, child in enumerate(value):
            yield from _walk_strings(child, f"{prefix}[{idx}]")


def _iter_table_rows(lines: list[str]) -> Iterable[tuple[int, list[str]]]:
    for idx, line in enumerate(lines):
        if "|" not in line or idx + 1 >= len(lines) or "|" not in lines[idx + 1]:
            continue
        sep = [cell.strip() for cell in lines[idx + 1].strip().strip("|").split("|")]
        if not sep or not all(re.fullmatch(r":?-{2,}:?", cell) for cell in sep):
            continue
        expected = len([cell for cell in line.strip().strip("|").split("|")])
        row_idx = idx + 2
        while row_idx < len(lines) and "|" in lines[row_idx] and lines[row_idx].strip():
            row = [cell.strip() for cell in lines[row_idx].strip().strip("|").split("|")]
            yield row_idx + 1, row
            row_idx += 1


def lint_markdown(path: Path) -> dict[str, Any]:
    text = path.read_text(encoding="utf-8", errors="replace")
    lines = text.splitlines()
    findings: list[dict[str, Any]] = []
    placeholders = [{"line": idx, "text": line.strip()} for idx, line in enumerate(lines, 1) if PLACEHOLDER_RE.search(line) or UNRESOLVED_TEMPLATE_RE.search(line)]
    for item in placeholders:
        findings.append({"severity": "warning", "kind": "placeholder", **item})
    for idx, line in enumerate(lines, 1):
        if INTERNAL_MARKER_RE.search(line) and ("批注" in line or "TODO" in line or "FIXME" in line or "内部" in line or "待" in line):
            findings.append({"severity": "warning", "kind": "internal_marker", "line": idx, "text": line.strip()})

    note_blocks: list[dict[str, Any]] = []
    for idx, line in enumerate(lines, 1):
        if re.match(r"^\s*>\s*\*?\*?【[^】]*批注[^】]*】", line):
            note_blocks.append({"line": idx, "text": line.strip()})
            findings.append({"severity": "info", "kind": "note_block", "line": idx, "text": line.strip()})

    heading_counts: dict[str, list[int]] = {}
    clause_counts: dict[str, list[int]] = {}
    for idx, line in enumerate(lines, 1):
        match = HEADING_RE.match(line)
        if match:
            title = re.sub(r"\s+", "", match.group(2)).lower()
            heading_counts.setdefault(title, []).append(idx)
        clause = CLAUSE_HEADING_RE.match(line)
        if clause:
            clause_counts.setdefault(clause.group(1), []).append(idx)
    for title, locations in heading_counts.items():
        if title and len(locations) > 1:
            findings.append({"severity": "warning", "kind": "duplicate_heading", "heading": title, "lines": locations})
    for clause, locations in clause_counts.items():
        if len(locations) > 1:
            findings.append({"severity": "warning", "kind": "duplicate_clause_number", "number": clause, "lines": locations})

    malformed_rows: list[dict[str, Any]] = []
    for line_no, row in _iter_table_rows(lines):
        # 找到上一个表头/分隔行，以校验列数一致性。
        sep_index = line_no - 2
        header_line = lines[sep_index - 1] if sep_index - 1 >= 0 else ""
        expected = len(header_line.strip().strip("|").split("|")) if "|" in header_line else len(row)
        if len(row) != expected:
            malformed_rows.append({"line": line_no, "expected_columns": expected, "actual_columns": len(row)})
            findings.append({"severity": "error", "kind": "malformed_table", **malformed_rows[-1]})

    errors = [item for item in findings if item["severity"] == "error"]
    return {
        "passed": not errors,
        "findings": findings,
        "placeholder_count": len(placeholders),
        "note_block_count": len(note_blocks),
        "malformed_table_count": len(malformed_rows),
        "internal_marker_count": sum(1 for item in findings if item["kind"] == "internal_marker"),
    }


def _docx_text_and_parts(path: Path) -> tuple[str, bool]:
    chunks: list[str] = []
    has_comments = False
    with zipfile.ZipFile(path) as zf:
        names = set(zf.namelist())
        has_comments = "word/comments.xml" in names
        import xml.etree.ElementTree as ET

        ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        for part in ("word/document.xml", "word/comments.xml"):
            if part not in names:
                continue
            root = ET.fromstring(zf.read(part))
            for node in root.iter(f"{ns}t"):
                if node.text:
                    chunks.append(node.text)
    return "\n".join(chunks), has_comments


def _scan_signing_docx(path: Path) -> dict[str, Any]:
    # 统一使用 md_to_docx 的 OOXML 扫描器，覆盖正文、表格、页眉页脚和
    # 脚注，并跳过已删除修订文本，避免签署版漏检内部内容。
    result = scan_signing_docx(path)
    # 占位符属于待业务补充项，由 require_complete/strict 门禁决定是否阻断；
    # “签署版内部内容扫描”本身只对批注、内部标记和 comments.xml 失败。
    result["clean_internal"] = not result.get("has_comments") and not result.get("internal_markers")
    result["passed"] = bool(result["clean_internal"])
    return result


def _visual_gate(path: Path, work_dir: Path, disabled: bool) -> dict[str, Any]:
    if disabled:
        return {"status": "skipped", "reason": "--no-visual specified", "passed": False}
    try:
        result = inspect_docx(path, work_dir=work_dir, cache_dir=Path(work_dir).parent / "cache")
        # inspect_docx 的 passed 只代表没有发现几何硬错误；status=skipped 仍需人工复核。
        result["gate_passed"] = result.get("status") == "ok" and bool(result.get("passed"))
        return result
    except Exception as exc:
        return {"status": "error", "passed": False, "gate_passed": False, "error": str(exc)}


def run_drafting_pipeline(
    request_path: str | Path,
    draft_md: str | Path | None,
    out_dir: str | Path,
    *,
    reference_docx: str | Path | None = None,
    require_complete: bool = False,
    strict: bool = False,
    no_visual: bool = False,
    allow_unverified: bool = False,
    schema_path: str | Path = REQUEST_SCHEMA,
) -> dict[str, Any]:
    request_path = Path(request_path).resolve()
    draft_md = Path(draft_md).resolve() if draft_md else None
    out_dir = Path(out_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    reference = Path(reference_docx).resolve() if reference_docx else DEFAULT_REFERENCE
    run_id = f"draft-{uuid.uuid4().hex}"
    # Every generated file is built in an isolated run tree.  The live output
    # directory is touched only after all deterministic gates pass and the
    # publish journal has been written by _publish_staged_tree().
    staged_root = out_dir / ".pipeline-runs" / run_id
    work_dir = staged_root / "work"
    work_dir.mkdir(parents=True, exist_ok=True)
    recovery = _recover_publish_journals(out_dir)
    if recovery:
        _write_json_atomic(work_dir / "recovery-report.json", {"recovered": recovery, "count": len(recovery)})
    errors: list[str] = []
    warnings: list[str] = []
    outputs: dict[str, str] = {}
    request: dict[str, Any] = {}

    try:
        request = json.loads(request_path.read_text(encoding="utf-8"))
    except Exception as exc:
        errors.append(f"读取 draft_request 失败: {exc}")
    request_report = validate_request(request, Path(schema_path), require_complete=require_complete) if not errors else {"valid": False, "errors": errors, "warnings": [], "missing_fields": [], "schema_errors": []}
    errors.extend(x for x in request_report.get("errors", []) if x not in errors)
    warnings.extend(request_report.get("warnings", []))
    selected_template: dict[str, Any] | None = None
    if draft_md is None:
        try:
            draft_md, selected_template = select_template(request, request_report.get("route"))
        except Exception as exc:
            errors.append(f"模板目录/路由加载失败: {exc}")
        if draft_md is None:
            draft_md = work_dir / "generated_route_skeleton.md"
            draft_md.write_text(_render_route_skeleton(request), encoding="utf-8")
            selected_template = {"id": None, "source": "generated_route_skeleton", "route": request.get("contract_type")}
            warnings.append("当前路由没有标准模板，已生成通用组合起草工作稿，必须人工逐条确认")
    if draft_md is not None and not draft_md.is_file():
        errors.append(f"draft Markdown 不存在: {draft_md}")

    lint_report = lint_markdown(draft_md) if draft_md is not None and draft_md.is_file() else {"passed": False, "findings": [], "placeholder_count": 0, "note_block_count": 0}
    if any(item.get("severity") == "error" for item in lint_report.get("findings", [])):
        errors.append("Markdown lint 存在结构性错误")
    if lint_report.get("placeholder_count", 0):
        msg = f"Markdown 含 {lint_report['placeholder_count']} 个未决占位符"
        (errors if require_complete else warnings).append(msg)

    citation_report: dict[str, Any] = {"passed": True, "summary": {"total": 0}}
    citation_code = 0
    if draft_md is not None and draft_md.is_file():
        try:
            citation_code, citation_report = verify_file(
                draft_md,
                allow_unmarked_out_of_kb=allow_unverified,
                dedupe=True,
                as_of=str(request.get("review_date") or "")[:10] or None,
            )
            if citation_code:
                errors.append("法条引用核验未通过")
            elif citation_report.get("summary", {}).get("status_unverified", 0):
                warnings.append("存在知识库状态待核实的法条引用")
        except Exception as exc:
            citation_code = 1
            errors.append(f"法条引用核验失败: {exc}")
    _write_json_atomic(work_dir / "citation_report.json", citation_report)
    _write_json_atomic(work_dir / "lint_report.json", lint_report)

    docx_reports: dict[str, Any] = {}
    signing_scan: dict[str, Any] = {}
    if not errors:
        signing = staged_root / f"{draft_md.stem}_签署草稿版.docx"
        notes = staged_root / f"{draft_md.stem}_起草说明版.docx"
        try:
            signing_source = draft_md
            template_sanitization: dict[str, Any] = {}
            if selected_template:
                # Generated templates may carry an editorial title such as
                # “附风险批注”. Remove only explicit internal labels from the
                # signing copy; placeholders remain visible and are governed
                # by require_complete/strict.
                raw_template = draft_md.read_text(encoding="utf-8", errors="replace")
                clean_template = re.sub(r"[（(]?附?(?:风险批注|起草批注|谈判批注)[）)]?", "", raw_template)
                signing_source = work_dir / "generated_signing_source.md"
                signing_source.write_text(clean_template, encoding="utf-8")
                template_sanitization = {"applied": True, "source": str(signing_source), "removed_internal_labels": raw_template.count("风险批注") + raw_template.count("起草批注") + raw_template.count("谈判批注")}
            convert_md_to_docx(signing_source, signing, reference_docx=reference, placeholder_highlight=True, strip_notes=True)
            convert_md_to_docx(draft_md, notes, reference_docx=reference, placeholder_highlight=True, notes_as_comments=True)
            signing_clean = strip_docx_comments(signing)
            outputs.update({"signing_docx": str(signing), "notes_docx": str(notes)})
            for label, path in (("signing", signing), ("notes", notes)):
                docx_reports[label] = validate_docx_file(path, strict=True)
                if not docx_reports[label].get("valid"):
                    errors.append(f"{label} DOCX 结构校验失败")
            signing_scan = _scan_signing_docx(signing)
            if not signing_scan.get("passed"):
                errors.append("签署版仍含内部批注或 comments.xml")
            if signing_scan.get("placeholder_count"):
                msg = f"签署版含 {signing_scan['placeholder_count']} 个未决占位符"
                (errors if require_complete else warnings).append(msg)
            visual_work = work_dir / "visual"
            docx_reports["signing_visual"] = _visual_gate(signing, visual_work / "signing", no_visual)
            docx_reports["notes_visual"] = _visual_gate(notes, visual_work / "notes", no_visual)
            for key in ("signing_visual", "notes_visual"):
                visual = docx_reports[key]
                if visual.get("status") == "error":
                    errors.append(f"{key} 视觉检查失败")
                elif visual.get("status") == "skipped":
                    warnings.append(f"{key} 视觉检查 skipped，需人工复核")
                elif not visual.get("gate_passed", visual.get("passed", False)):
                    errors.append(f"{key} 视觉检查发现硬错误")
        except Exception as exc:
            errors.append(f"双版本 DOCX 渲染失败: {exc}")

    if errors:
        status = "blocked"
    elif warnings:
        status = "manual_review"
    else:
        status = "releasable"
    if strict and status != "releasable":
        # 将 warning 原样提升为阻断原因，manifest 中保留具体的门禁信息，
        # 便于调用方定位是占位符、条件字段还是视觉依赖缺失。
        errors.extend(f"严格门禁: {warning}" for warning in warnings)
        errors.append("--strict 要求所有质量与视觉门禁通过")
        status = "blocked"
    manifest = {
        "run_id": run_id,
        "status": status,
        "output_dir": str(out_dir),
        "staging_dir": str(staged_root),
        "errors": errors,
        "warnings": warnings,
        "request": {"path": str(request_path), "sha256": _sha256(request_path) if request_path.is_file() else None},
        "draft_markdown": {"path": str(draft_md) if draft_md else None, "sha256": _sha256(draft_md) if draft_md and draft_md.is_file() else None},
        "schema": {"path": str(Path(schema_path).resolve()), "sha256": _sha256(Path(schema_path)) if Path(schema_path).is_file() else None},
        "route_map": {"path": str(ROUTE_MAP), "sha256": _sha256(ROUTE_MAP) if ROUTE_MAP.is_file() else None},
        "template_catalog": {
            "path": request_report.get("template_catalog", {}).get("path"),
            "sha256": request_report.get("template_catalog", {}).get("sha256"),
            "selected": selected_template,
        },
        "reference_docx": {"path": str(reference), "sha256": _sha256(reference) if reference.is_file() else None},
        "request_validation": request_report,
        "lint": lint_report,
        "citation": citation_report,
        "docx": docx_reports,
        "signing_scan": signing_scan,
        "signing_comment_cleanup": locals().get("signing_clean", {}),
        "template_sanitization": locals().get("template_sanitization", {}),
        "outputs": {key: {"path": value, "sha256": _sha256(Path(value)) if Path(value).is_file() else None} for key, value in outputs.items()},
        "options": {"require_complete": require_complete, "strict": strict, "no_visual": no_visual, "allow_unverified": allow_unverified},
    }
    _write_json_atomic(work_dir / "manifest.json", manifest)
    if status != "blocked":
        # Publication rewrites staged absolute paths in JSON audit files and
        # keeps the old public output recoverable if a move is interrupted.
        _publish_staged_tree(staged_root, out_dir, run_id)
        return json.loads((out_dir / "work" / "manifest.json").read_text(encoding="utf-8"))

    # Keep blocked runs isolated for forensic review, while indexing a small
    # failure manifest in the public work directory without replacing a prior
    # successful delivery.
    failure_path = out_dir / "work" / "pipeline-failures" / f"{run_id}.json"
    failure_path.parent.mkdir(parents=True, exist_ok=True)
    failure_manifest = dict(manifest)
    failure_manifest["failure"] = {
        "staging_dir": str(staged_root),
        "retained_for_forensics": True,
        "reason": "质量门禁阻断，未发布到正式输出目录",
    }
    _write_json_atomic(failure_path, failure_manifest)
    return manifest


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="合同起草请求校验、双版本渲染及结构/视觉门禁")
    parser.add_argument("--request", required=True, help="draft_request JSON")
    parser.add_argument("--draft-md", required=False, default=None, help="带起草批注的 Markdown；缺省按 template_catalog 自动选择")
    parser.add_argument("--out-dir", required=True, help="隔离输出目录")
    parser.add_argument("--reference-docx", default=None)
    parser.add_argument("--schema", default=str(REQUEST_SCHEMA))
    parser.add_argument("--require-complete", action="store_true", help="占位符和条件必填缺失时阻断")
    parser.add_argument("--strict", action="store_true", help="要求无 warning 且视觉门禁通过")
    parser.add_argument("--no-visual", action="store_true", help="跳过渲染器；manifest 记录 skipped，严格模式会阻断")
    parser.add_argument("--allow-unverified", action="store_true", help="允许未标注库外法源的法条引用")
    args = parser.parse_args(argv)
    report = run_drafting_pipeline(
        args.request,
        args.draft_md,
        args.out_dir,
        reference_docx=args.reference_docx,
        require_complete=args.require_complete,
        strict=args.strict,
        no_visual=args.no_visual,
        allow_unverified=args.allow_unverified,
        schema_path=args.schema,
    )
    print(json.dumps(report, ensure_ascii=False, indent=2))
    if report["status"] == "blocked":
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
