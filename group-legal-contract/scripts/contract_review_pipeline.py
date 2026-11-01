#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""合同审核确定性收尾流水线。

该入口不代替模型/法务编写风险意见，而是把审核完成后的确定性步骤串联起来：

1. 预检并复用源文件缓存；
2. 校验问题台账、修订指令和原文锚点；
3. 核验修订指令中的法条引用；
4. 写入 Word 修订并执行双向结构校验（默认不重复渲染 PDF）；
5. 核验意见书引用、转 Word；
6. 最后只做一次批量版面渲染。

所有路径在入口处解析为绝对路径，避免从 output/work 目录调用脚本时发生相对
技能路径错误。原始输入只读，所有产物写入指定输出目录。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import sys
import time
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from batch_visual_verify import verify_many  # noqa: E402
from build_review_outputs import build_outputs_from_draft  # noqa: E402
from docx_redline import apply_revisions  # noqa: E402
from docx_verify import verify_docx  # noqa: E402
from lib.kb import INDEX_PATH, kb_fingerprint, kb_source_consistency  # noqa: E402
from lib.runtime_state import atomic_json, file_lock  # noqa: E402
from lib.textio import read_text_auto  # noqa: E402
from md_to_docx import convert_md_to_docx  # noqa: E402
from preflight_contract import preflight_contract  # noqa: E402
from stage_mark import build_time_summary, record_stage_mark  # noqa: E402
from validate_review_bundle import validate_bundle  # noqa: E402
from verify_citations import verify_file  # noqa: E402


SKILL_DIR = SCRIPT_DIR.parent
DEFAULT_REFERENCE = SKILL_DIR / "assets" / "docx" / "report-reference.docx"


def _write_json(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(f".{path.name}.{uuid.uuid4().hex}.tmp")
    try:
        tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        tmp.replace(path)
    finally:
        if tmp.exists():
            tmp.unlink()


def _replace_path_values(value: Any, old: str, new: str) -> Any:
    """Rewrite staged absolute paths after a successful publish."""
    if isinstance(value, str):
        return value.replace(old, new)
    if isinstance(value, list):
        return [_replace_path_values(item, old, new) for item in value]
    if isinstance(value, dict):
        return {key: _replace_path_values(item, old, new) for key, item in value.items()}
    return value


def _refresh_artifact_hashes(value: Any) -> Any:
    """Refresh hashes for published JSON artifacts after path rewriting.

    ``blocks.json`` and similar audit files can contain an absolute staged
    path. Rewriting that path changes the file bytes; the preflight manifest
    must therefore refresh its ``*_sha256`` fields before it is published or
    the next isolated run will correctly reject the cache as corrupt.
    """
    if not isinstance(value, dict):
        return value
    artifacts = value.get("artifacts")
    if not isinstance(artifacts, dict):
        return value
    for key, raw_path in list(artifacts.items()):
        if key.endswith("_sha256") or not raw_path:
            continue
        hash_key = f"{key}_sha256"
        if hash_key not in artifacts:
            continue
        path = Path(str(raw_path))
        if path.is_file():
            artifacts[hash_key] = _sha256_file(path)
    return value


def _publish_staged_tree(staged: Path, out: Path, run_id: str) -> dict[str, Any]:
    """Publish a completed run with a rollback journal.

    Each file is moved under the output lock. Existing files are first moved to
    a backup inside the run directory, so an injected or filesystem failure can
    restore the previous output set. The journal is intentionally retained in
    ``work`` after commit for post-run inspection.
    """
    staged = staged.resolve()
    out = out.resolve()
    journal = staged / "publish-journal.json"
    entries: list[dict[str, Any]] = []
    committed: list[dict[str, Any]] = []

    files = sorted(
        p for p in staged.rglob("*")
        if p.is_file() and p.name != "publish-journal.json" and ".publish-backup" not in p.parts
    )
    for source in files:
        rel = source.relative_to(staged)
        target = out / rel
        backup = staged / ".publish-backup" / rel
        entries.append({"relative": rel.as_posix(), "target": str(target), "source": str(source), "backup": str(backup)})

    journal_data = {"run_id": run_id, "status": "prepared", "entries": entries}
    atomic_json(journal, journal_data)
    try:
        with file_lock(out / ".pipeline-publish.lock"):
            for item in entries:
                source = Path(item["source"])
                target = Path(item["target"])
                backup = Path(item["backup"])
                target.parent.mkdir(parents=True, exist_ok=True)
                if target.exists():
                    backup.parent.mkdir(parents=True, exist_ok=True)
                    os.replace(target, backup)
                    item["had_previous"] = True
                else:
                    item["had_previous"] = False
                # Persist the backup decision before moving the staged source;
                # this is the only durable fact needed to recover a crash in
                # the narrow window between the two os.replace calls.
                journal_data["status"] = "publishing"
                journal_data["in_flight"] = item["relative"]
                atomic_json(journal, journal_data)
                os.replace(source, target)
                committed.append(item)
                journal_data["status"] = "publishing"
                journal_data["committed"] = len(committed)
                journal_data.pop("in_flight", None)
                atomic_json(journal, journal_data)
            journal_data["status"] = "committed"
            journal_data["committed"] = len(committed)
            # Write the journal into the live output only after every target is
            # committed; this file is diagnostic and never an input artifact.
            atomic_json(out / "work" / "publish-journal.json", journal_data)
    except Exception as exc:
        # Include the in-flight item: a failure between moving an old target
        # to backup and moving the new source would otherwise leave the live
        # output path missing even though the item was not yet committed.
        for item in reversed(entries):
            target = Path(item["target"])
            backup = Path(item["backup"])
            try:
                if target.exists() and (item in committed or item.get("had_previous") is False):
                    target.unlink()
                if item.get("had_previous") and backup.exists():
                    backup.parent.mkdir(parents=True, exist_ok=True)
                    os.replace(backup, target)
            except Exception:
                # Keep the journal and original exception; recovery can use the
                # remaining backup files without masking the first failure.
                pass
        journal_data.update({"status": "rolled_back", "error": str(exc), "committed": len(committed)})
        atomic_json(journal, journal_data)
        raise

    # A manifest is the public return value. Its staged paths must not leak
    # into downstream reports after the transaction has completed.
    report_path = out / "work" / "pipeline-report.json"
    report: dict[str, Any] = {}
    # Rewrite every JSON audit file that embeds the staged absolute path, not
    # only the public manifest (time-summary and preflight artifacts also carry
    # work_dir/source paths used by downstream audit tooling).
    for json_path in (out / "work").rglob("*.json"):
        if json_path.name == "publish-journal.json":
            continue
        try:
            raw_json = json_path.read_text(encoding="utf-8")
            # Re-serializing a JSON artifact changes its byte hash even when
            # no staged path occurs inside it. Skip those files so the
            # preflight artifact hashes remain valid across publication.
            staged_json = str(staged).replace("\\", "\\\\")
            if str(staged) not in raw_json and staged_json not in raw_json:
                continue
            parsed = json.loads(raw_json)
            parsed = _replace_path_values(parsed, str(staged), str(out))
            parsed = _refresh_artifact_hashes(parsed)
            if json_path == report_path and isinstance(parsed, dict):
                parsed["output_dir"] = str(out)
                report = parsed
            atomic_json(json_path, parsed)
        except Exception:
            continue
    shutil.rmtree(staged, ignore_errors=True)
    return report


def _seed_cross_run_caches(source: Path, live_work: Path, staged_work: Path) -> dict[str, Any]:
    """Seed only validated preflight/visual caches into an isolated run.

    The review outputs themselves are deliberately excluded.  The preflight
    manifest is copied only when it is bound to the same source SHA and its
    declared files are copied under the new run's work directory with paths
    rewritten, allowing ``preflight_contract`` to perform its normal cache-key
    and artifact-hash checks.  This keeps transactional isolation without
    sacrificing cross-run cache hits.
    """
    seeded: dict[str, Any] = {"preflight": False, "visual_files": 0}
    manifest_path = live_work / "preflight.json"
    if manifest_path.is_file():
        try:
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            if manifest.get("source", {}).get("sha256") == _sha256_file(source):
                artifact_paths = manifest.get("artifacts") or {}
                copied: set[Path] = set()
                for key, raw in artifact_paths.items():
                    if not key.endswith("_sha256") and raw:
                        candidate = Path(str(raw)).resolve()
                        try:
                            rel = candidate.relative_to(live_work.resolve())
                        except ValueError:
                            continue
                        if not candidate.is_file():
                            continue
                        destination = staged_work / rel
                        destination.parent.mkdir(parents=True, exist_ok=True)
                        shutil.copy2(candidate, destination)
                        copied.add(destination)
                required = ("base_docx", "blocks_json", "contract_txt")
                if all((staged_work / Path(str(artifact_paths.get(name))).name).is_file() for name in required):
                    rewritten = _replace_path_values(manifest, str(live_work.resolve()), str(staged_work.resolve()))
                    _write_json(staged_work / "preflight.json", rewritten)
                    seeded["preflight"] = True
        except Exception:
            # A stale/corrupt cache is simply ignored; preflight will rebuild
            # it and record the miss in its manifest.
            pass

    visual_cache = live_work / "visual-cache"
    if visual_cache.is_dir():
        destination = staged_work / "visual-cache"
        shutil.copytree(visual_cache, destination, dirs_exist_ok=True)
        seeded["visual_files"] = sum(1 for item in destination.rglob("*") if item.is_file())
    return seeded


def _recover_publish_journals(out: Path) -> list[dict[str, Any]]:
    """Rollback interrupted publish transactions before starting a new run.

    A process can terminate after a target has been moved to backup but before
    the next journal checkpoint.  On the next invocation, inspect every
    surviving run journal under ``.pipeline-runs`` and restore the last known
    committed output set. Recovery is conservative: it rolls back an
    unfinished transaction instead of exposing a partially published result.
    """
    root = out / ".pipeline-runs"
    if not root.is_dir():
        return []
    recovered: list[dict[str, Any]] = []
    with file_lock(out / ".pipeline-publish.lock"):
        for journal in sorted(root.glob("*/publish-journal.json")):
            try:
                data = json.loads(journal.read_text(encoding="utf-8"))
            except Exception:
                continue
            if data.get("status") in {"committed", "rolled_back"}:
                continue
            entries = data.get("entries") or []
            committed = int(data.get("committed") or 0)
            restored = 0
            for index, item in reversed(list(enumerate(entries))):
                target = Path(str(item.get("target") or ""))
                backup = Path(str(item.get("backup") or ""))
                source = Path(str(item.get("source") or ""))
                try:
                    # Entries before the last checkpoint were published (or
                    # may have been published just before a crash); entries at
                    # the checkpoint can be inferred from source existence.
                    published = index < committed or (not source.exists() and target.exists())
                    if published and target.exists():
                        target.unlink()
                    if item.get("had_previous") and backup.exists():
                        target.parent.mkdir(parents=True, exist_ok=True)
                        os.replace(backup, target)
                        restored += 1
                except Exception:
                    continue
            data.update({"status": "recovered_rollback", "recovered_at_utc": _utc_now_iso(), "restored": restored})
            atomic_json(journal, data)
            recovered.append({"run_id": data.get("run_id"), "journal": str(journal), "restored": restored})
    return recovered


def _sha256_file(path: Path) -> str | None:
    """Return a content hash for manifest binding; missing files stay explicit."""
    if not path.is_file():
        return None
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _assert_output_path(path: Path, out_dir: Path) -> None:
    """Prevent a run from silently overwriting an unrelated directory."""
    try:
        path.relative_to(out_dir)
    except ValueError as exc:
        raise ValueError(f"流水线产物必须位于 --out-dir 内: {path}") from exc


def _assert_reusable_outputs(work: Path, source: Path) -> None:
    """Bind explicitly reused artefacts to the exact source/run manifest."""
    manifest_path = work / "pipeline-report.json"
    if not manifest_path.is_file():
        raise ValueError("复用已有审核台账需要 work/pipeline-report.json 作为来源绑定证明")
    try:
        manifest = _load_json(manifest_path)
    except Exception as exc:
        raise ValueError(f"旧流水线 manifest 无法解析，拒绝复用: {manifest_path}") from exc
    if manifest.get("status") != "succeeded":
        raise ValueError(f"旧流水线状态为 {manifest.get('status')!r}，仅允许复用 succeeded 产物")
    current_hash = _sha256_file(source)
    if manifest.get("source") != str(source) or manifest.get("source_sha256") != current_hash:
        raise ValueError("旧审核产物与当前输入合同路径或 SHA256 不一致，拒绝复用")
    artifacts = manifest.get("artifacts") or {}
    for item in artifacts.values():
        if not isinstance(item, dict):
            continue
        path = Path(str(item.get("path") or "")).resolve()
        expected = item.get("sha256")
        if expected and _sha256_file(path) != expected:
            raise ValueError(f"旧审核产物已被修改，拒绝复用: {path}")


def _write_failure_manifest(
    out_dir: Path,
    source: Path,
    exc: BaseException,
    *,
    run_id: str | None = None,
) -> None:
    work = out_dir / "work"
    manifest: dict[str, Any] = {}
    existing = work / "pipeline-report.json"
    if existing.is_file():
        try:
            parsed = json.loads(existing.read_text(encoding="utf-8"))
            if isinstance(parsed, dict):
                manifest.update(parsed)
        except Exception:
            pass
    manifest.update({
        "pipeline_version": "review-pipeline-1.3",
        "run_id": run_id or manifest.get("run_id") or f"run-{uuid.uuid4().hex}",
        "status": "failed",
        "source": str(source),
        "source_sha256": _sha256_file(source),
        "output_dir": str(out_dir),
        "error": {"type": type(exc).__name__, "message": str(exc)},
        "finished_at_utc": _utc_now_iso(),
    })
    try:
        _write_json(work / "pipeline-report.json", manifest)
    except Exception:
        # Preserve the original pipeline exception even if the failure report
        # itself cannot be written (for example, a read-only output mount).
        pass


def _load_json(path: Path) -> dict[str, Any]:
    text, _ = read_text_auto(path)
    data = json.loads(text)
    if not isinstance(data, dict):
        raise ValueError(f"JSON 根节点必须是对象: {path}")
    return data


def _stage(timings: dict[str, float], name: str, started: float) -> None:
    timings[name] = round((time.perf_counter() - started) * 1000, 1)


def _utc_now_iso() -> str:
    """返回可审计的 UTC 时间戳，避免本地时区歧义。"""
    return datetime.now(timezone.utc).isoformat(timespec="milliseconds").replace("+00:00", "Z")


def _embed_appendix(report_text: str, appendix_text: str) -> tuple[str, bool]:
    """将法条核验表嵌入意见书 Markdown。

    优先替换显式 ``<!-- APPENDIX_B_PLACEHOLDER -->``。若占位符前已经有
    “附录 B”标题，则去掉核验片段自带的首个标题，避免重复标题。没有占位符
    时追加到文末，保证最终 docx 一定包含核验表，同时不改写原始 Markdown。
    """
    marker = "<!-- APPENDIX_B_PLACEHOLDER -->"
    appendix = appendix_text.strip()
    if not appendix:
        return report_text, False
    if marker not in report_text:
        # 支持重复运行：若意见书已经包含本流水线生成的附录表格，则直接复用。
        if "法条引用自动核验表" in report_text and "| 序号 | 引用原文 |" in report_text:
            return report_text, False
        suffix = "\n\n" + appendix + "\n"
        return report_text.rstrip() + suffix, True

    before, after = report_text.split(marker, 1)
    appendix_lines = appendix.splitlines()
    recent = "\n".join(before.rstrip().splitlines()[-3:])
    if appendix_lines and "附录 B" in recent and appendix_lines[0].startswith("#"):
        appendix = "\n".join(appendix_lines[2:]).lstrip()
    rendered = before.rstrip() + "\n\n" + appendix + "\n" + after.lstrip("\r\n")
    return rendered, True


def _run_pipeline(
    input_path: str | Path,
    out_dir: str | Path,
    *,
    draft_path: str | Path | None = None,
    revisions_path: str | Path | None = None,
    review_result_path: str | Path | None = None,
    report_md_path: str | Path | None = None,
    reference_docx: str | Path | None = None,
    revised_out: str | Path | None = None,
    clean_out: str | Path | None = None,
    report_out: str | Path | None = None,
    force_preflight: bool = False,
    visual: bool = True,
    author: str | None = None,
    date_iso: str | None = None,
    allow_existing_revisions: bool = False,
    allow_semantic_warnings: bool = False,
    reuse_existing: bool = False,
    run_id: str | None = None,
) -> dict[str, Any]:
    source = Path(input_path).resolve()
    out = Path(out_dir).resolve()
    work = out / "work"
    out.mkdir(parents=True, exist_ok=True)
    work.mkdir(parents=True, exist_ok=True)
    if not source.is_file():
        raise FileNotFoundError(f"输入合同不存在: {source}")

    record_stage_mark(work, "pipeline_started", note="确定性收尾流水线启动", run_id=run_id, actor="script")

    stem = source.stem
    draft = Path(draft_path).resolve() if draft_path else (work / "review-draft.json")
    if not draft.is_file():
        draft = None

    revisions = Path(revisions_path).resolve() if revisions_path else work / "revisions.json"
    review_result = Path(review_result_path).resolve() if review_result_path else work / "review-result.json"
    report_md = Path(report_md_path).resolve() if report_md_path else None
    revised = Path(revised_out).resolve() if revised_out else out / f"{stem}_审查修订版.docx"
    clean = Path(clean_out).resolve() if clean_out else out / f"{stem}_清洁版.docx"
    report_docx = Path(report_out).resolve() if report_out else out / f"{stem}_法律审查意见书.docx"

    for artifact_path in (revisions, review_result, report_md, revised, clean, report_docx):
        if artifact_path is not None:
            _assert_output_path(Path(artifact_path), out)

    # A run without a draft must opt in to consuming existing ledgers.  This
    # prevents a newly preflighted source from being paired with yesterday's
    # revisions/review-result files merely because they share an output folder.
    if draft is None and not reuse_existing:
        stale = [p for p in (revisions, review_result, report_md) if p is not None and p.is_file()]
        if stale:
            raise ValueError(
                "检测到无 review-draft.json 但已有旧审核产物；为防止代际混用，请提供 draft 或显式 reuse_existing=True"
            )
    elif draft is None and reuse_existing:
        stale = [p for p in (revisions, review_result, report_md) if p is not None and p.is_file()]
        if stale:
            _assert_reusable_outputs(work, source)

    timings: dict[str, float] = {}
    pipeline_started = time.perf_counter()
    pipeline_started_at = _utc_now_iso()
    manifest: dict[str, Any] = {
        "pipeline_version": "review-pipeline-1.3",
        "run_id": run_id or f"run-{uuid.uuid4().hex}",
        "status": "running",
        "source": str(source),
        "source_sha256": _sha256_file(source),
        "output_dir": str(out),
        "inputs": {
            "draft": {"path": str(draft) if draft else None, "sha256": _sha256_file(draft) if draft else None},
            "revisions": {"path": str(revisions), "sha256": _sha256_file(revisions)},
            "review_result": {"path": str(review_result), "sha256": _sha256_file(review_result)},
            "report_markdown": {"path": str(report_md) if report_md else None, "sha256": _sha256_file(report_md) if report_md else None},
            "reference_docx": {
                "path": str(Path(reference_docx).resolve()) if reference_docx else str(DEFAULT_REFERENCE),
                "sha256": _sha256_file(Path(reference_docx).resolve() if reference_docx else DEFAULT_REFERENCE),
            },
        },
        "toolchain": {
            "pipeline_sha256": _sha256_file(Path(__file__).resolve()),
            "review_draft_schema_sha256": _sha256_file(SKILL_DIR / "assets" / "schemas" / "review-draft.schema.json"),
            "revisions_schema_sha256": _sha256_file(SKILL_DIR / "assets" / "schemas" / "revisions.schema.json"),
            "kb_index": {
                "path": str(INDEX_PATH),
                "sha256": _sha256_file(INDEX_PATH),
                "fingerprint": kb_fingerprint(INDEX_PATH),
                "source_consistency": kb_source_consistency(INDEX_PATH),
            },
        },
        "policy": {
            "allow_existing_revisions": allow_existing_revisions,
            "allow_semantic_warnings": allow_semantic_warnings,
            "reuse_existing": reuse_existing,
        },
        "outputs": {},
        "stages": {},
        "timing": {
            "started_at_utc": pipeline_started_at,
            "clock": "time.perf_counter + UTC wall clock",
        },
    }
    _write_json(work / "pipeline-report.json", manifest)

    started = time.perf_counter()
    preflight = preflight_contract(
        source,
        out,
        force=force_preflight,
        existing_revisions="accept-view" if allow_existing_revisions else "fail",
        run_id=run_id,
    )
    _stage(timings, "preflight_ms", started)
    manifest["preflight"] = {
        "cache_hit": preflight.get("cache", {}).get("hit", False),
        "base_docx": preflight["artifacts"]["base_docx"],
        "block_count": preflight["extraction"].get("block_count"),
        "high_findings": preflight["checks"].get("high_findings", []),
        "review_findings": preflight["checks"].get("review_findings", []),
    }
    base_docx = Path(preflight["artifacts"]["base_docx"]).resolve()
    blocks_path = Path(preflight["artifacts"]["blocks_json"]).resolve()
    manifest["preflight"]["base_docx_sha256"] = _sha256_file(base_docx)
    manifest["preflight"]["blocks_sha256"] = _sha256_file(blocks_path)

    if draft is not None:
        started = time.perf_counter()
        derived = build_outputs_from_draft(
            draft,
            blocks_path=blocks_path,
            preflight_path=work / "preflight.json",
            revisions_out=revisions,
            review_result_out=review_result,
            report_md_out=report_md or (work / "review_report.md"),
        )
        _stage(timings, "draft_derive_ms", started)
        manifest["draft_derived"] = derived
        if report_md is None:
            report_md = Path(derived["review_report_md"]).resolve()
        record_stage_mark(work, "draft_built", note=f"从 review-draft.json 派生三份台账与意见书（{timings['draft_derive_ms']}ms）", run_id=run_id, actor="script")

    revisions_exists = revisions.is_file()
    review_result_exists = review_result.is_file()
    if not revisions_exists:
        revisions = None
    if not review_result_exists:
        review_result = None
    if report_md is None and (work / "review_report.md").is_file() and draft is not None:
        report_md = (work / "review_report.md").resolve()

    if revisions is not None:
        started = time.perf_counter()
        bundle_report = validate_bundle(
            revisions,
            review_result if review_result is not None else None,
            blocks_path,
            fail_on_warnings=not allow_existing_revisions,
        )
        _write_json(work / "bundle-validation.json", bundle_report)
        _stage(timings, "bundle_validation_ms", started)
        manifest["bundle_validation"] = bundle_report
        if not bundle_report["valid"]:
            raise ValueError("审核问题台账/修订指令校验失败，详见 work/bundle-validation.json")

        started = time.perf_counter()
        revision_citation_code, revision_citation = verify_file(
            revisions,
            exclude_sections=[],
            dedupe=False,
            as_of=date_iso[:10] if date_iso else None,
        )
        _write_json(work / "citation_report_revisions.json", revision_citation)
        (work / "appendix_b_revisions.md").write_text(
            _render_citation_markdown(revision_citation), encoding="utf-8"
        )
        _stage(timings, "revision_citation_ms", started)
        manifest["revision_citations"] = revision_citation.get("summary", {})
        if revision_citation_code != 0:
            raise ValueError("revisions.json 法条引用核验失败")

        started = time.perf_counter()
        revisions_data = _load_json(revisions)
        apply_log = apply_revisions(
            input_docx=base_docx,
            revisions_data=revisions_data,
            out_docx=revised,
            log_path=work / "apply_log.json",
            author_override=author,
            track_on=True,
            date_iso=date_iso,
        )
        _stage(timings, "redline_ms", started)
        manifest["apply"] = apply_log.get("summary", {})

        started = time.perf_counter()
        verify_report = verify_docx(
            original_docx=base_docx,
            revised_docx=revised,
            revisions_json=revisions,
            log_path=work / "apply_log.json",
            report_path=work / "verify_report.json",
            clean_out=clean,
            render_check_dir=None,
        )
        _stage(timings, "docx_verify_ms", started)
        manifest["verification"] = {
            "passed": verify_report.get("passed"),
            "semantic_review_required": verify_report.get("semantic_review_required", False),
            "semantic_warning_count": len(verify_report.get("semantic_warnings") or []),
        }
        if not verify_report.get("passed"):
            raise ValueError("修订版双向校验失败，详见 work/verify_report.json")
        semantic_warnings = verify_report.get("semantic_warnings") or []
        if semantic_warnings and not allow_semantic_warnings:
            raise ValueError(
                "检测到语义校验警告，默认 fail-closed；如已完成人工复核，请显式设置 allow_semantic_warnings=True"
            )
        if semantic_warnings:
            manifest["verification"]["semantic_override"] = {
                "applied": True,
                "reason": "allow_semantic_warnings",
                "count": len(semantic_warnings),
            }
        manifest["outputs"]["revised"] = str(revised)
        manifest["outputs"]["clean"] = str(clean)

    if report_md is not None:
        if not report_md.is_file():
            raise FileNotFoundError(f"意见书 Markdown 不存在: {report_md}")
        started = time.perf_counter()
        report_citation_code, report_citation = verify_file(
            report_md,
            exclude_sections=["附录 B"],
            dedupe=True,
            as_of=date_iso[:10] if date_iso else None,
        )
        _write_json(work / "citation_report_review.json", report_citation)
        appendix_text = _render_citation_markdown(report_citation)
        appendix_path = work / "appendix_b_review.md"
        appendix_path.write_text(appendix_text, encoding="utf-8")
        _stage(timings, "report_citation_ms", started)
        manifest["report_citations"] = report_citation.get("summary", {})
        if report_citation_code != 0:
            raise ValueError("意见书法条引用核验失败")

        started = time.perf_counter()
        report_render_input = work / "report_with_appendix.md"
        report_render_text, appendix_embedded = _embed_appendix(
            read_text_auto(report_md)[0], appendix_text
        )
        report_render_input.write_text(report_render_text, encoding="utf-8")
        convert_md_to_docx(
            input_md=report_render_input,
            out_docx=report_docx,
            reference_docx=Path(reference_docx).resolve() if reference_docx else DEFAULT_REFERENCE,
            placeholder_highlight=True,
        )
        _stage(timings, "report_docx_ms", started)
        manifest["outputs"]["report"] = str(report_docx)
        manifest["outputs"]["report_render_input"] = str(report_render_input)
        manifest["report"] = {
            "source_markdown": str(report_md),
            "render_markdown": str(report_render_input),
            "appendix_path": str(appendix_path),
            "appendix_embedded": appendix_embedded,
        }

    # 修订版、清洁版和意见书必须共享同一版面门禁；只检查带修订版会漏掉
    # 接受修订后的布局变化。开启 visual 时要求实际渲染成功，缺少渲染器
    # 只能通过 --no-visual 明确跳过并在审计中留下记录。
    visual_paths = [Path(p) for p in (revised, clean, report_docx) if p.is_file()]
    if visual and visual_paths:
        started = time.perf_counter()
        visual_report = verify_many(visual_paths, require_render=True, cache_dir=work / "visual-cache")
        _write_json(work / "visual_batch.json", visual_report)
        _stage(timings, "visual_ms", started)
        manifest["visual"] = {
            "passed": visual_report.get("passed"),
            "count": visual_report.get("count"),
            "deduplicated_count": visual_report.get("deduplicated_count", 0),
        }
        if not visual_report.get("passed"):
            raise ValueError("最终版面检查失败，详见 work/visual_batch.json")

    if not manifest["outputs"]:
        raise ValueError("流水线未生成任何可发布审核产物；请提供 review-draft/revisions 或意见书")

    for name, path in (
        ("revisions", revisions),
        ("review_result", review_result),
        ("report_markdown", report_md),
    ):
        if path is not None and Path(path).is_file():
            manifest["outputs"][name] = str(Path(path))

    timings["total_ms"] = round((time.perf_counter() - pipeline_started) * 1000, 1)
    manifest["timing"].update(
        {
            "finished_at_utc": _utc_now_iso(),
            "elapsed_ms": timings["total_ms"],
            "phase_ms": dict(timings),
        }
    )
    manifest["timings_ms"] = timings
    # 保留一个便于看板读取的同名阶段映射；timings_ms 是审计原始数据。
    manifest["stages"] = dict(timings)
    manifest["status"] = "succeeded"
    manifest["artifacts"] = {
        name: {"path": path, "sha256": _sha256_file(Path(path))}
        for name, path in manifest["outputs"].items()
    }
    _write_json(work / "pipeline-report.json", manifest)
    record_stage_mark(work, "pipeline_done", note=f"确定性收尾流水线完成（{timings['total_ms']}ms）", run_id=run_id, actor="script")
    build_time_summary(work, run_id=run_id)
    return manifest


def run_pipeline(
    input_path: str | Path,
    out_dir: str | Path,
    *,
    draft_path: str | Path | None = None,
    revisions_path: str | Path | None = None,
    review_result_path: str | Path | None = None,
    report_md_path: str | Path | None = None,
    reference_docx: str | Path | None = None,
    revised_out: str | Path | None = None,
    clean_out: str | Path | None = None,
    report_out: str | Path | None = None,
    force_preflight: bool = False,
    visual: bool = True,
    author: str | None = None,
    date_iso: str | None = None,
    allow_existing_revisions: bool = False,
    allow_semantic_warnings: bool = False,
    reuse_existing: bool = False,
) -> dict[str, Any]:
    """Run the deterministic pipeline as an isolated, publishable transaction."""
    source = Path(input_path).resolve()
    out = Path(out_dir).resolve()
    run_id = f"run-{uuid.uuid4().hex}"
    staged = out / ".pipeline-runs" / run_id
    out.mkdir(parents=True, exist_ok=True)
    recovery = _recover_publish_journals(out)

    def stage_output_path(raw: str | Path | None) -> Path | None:
        if raw is None:
            return None
        resolved = Path(raw).resolve()
        try:
            rel = resolved.relative_to(out)
        except ValueError as exc:
            raise ValueError(f"流水线产物必须位于 --out-dir 内: {resolved}") from exc
        return staged / rel

    # Preserve the historical auto-detection semantics while ensuring that an
    # old ledger is copied into this run only when the caller explicitly opts
    # into reuse. A normal run starts with an empty work directory.
    original_work = out / "work"
    original_draft = Path(draft_path).resolve() if draft_path else original_work / "review-draft.json"
    if draft_path is None and original_draft.is_file():
        staged_draft = staged / "work" / "review-draft.json"
        staged_draft.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(original_draft, staged_draft)
        draft_arg: str | Path | None = staged_draft
    elif draft_path is not None:
        draft_arg = draft_path
    else:
        draft_arg = None

    staged_work = staged / "work"
    staged_work.mkdir(parents=True, exist_ok=True)
    if recovery:
        _write_json(staged_work / "recovery-report.json", {"recovered": recovery, "count": len(recovery)})
    cache_seed = _seed_cross_run_caches(source, original_work, staged_work)

    staged_revisions = stage_output_path(revisions_path)
    staged_review_result = stage_output_path(review_result_path)
    staged_report_md = stage_output_path(report_md_path)
    staged_revised = stage_output_path(revised_out)
    staged_clean = stage_output_path(clean_out)
    staged_report = stage_output_path(report_out)

    if reuse_existing and draft_arg is None:
        # Reuse requires a complete, source-bound manifest. Copy only the
        # declared input ledgers; generated artifacts will be republished by
        # the isolated run and are never read in place from the live folder.
        for name in ("revisions.json", "review-result.json", "review_report.md", "pipeline-report.json"):
            src = original_work / name
            if src.is_file():
                dst = staged / "work" / name
                dst.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(src, dst)
        if staged_revisions is None:
            staged_revisions = staged / "work" / "revisions.json"
        if staged_review_result is None:
            staged_review_result = staged / "work" / "review-result.json"
        if staged_report_md is None and (staged / "work" / "review_report.md").is_file():
            staged_report_md = staged / "work" / "review_report.md"
    elif draft_arg is None and not reuse_existing:
        original_revisions = Path(revisions_path).resolve() if revisions_path else original_work / "revisions.json"
        original_review_result = Path(review_result_path).resolve() if review_result_path else original_work / "review-result.json"
        original_report_md = Path(report_md_path).resolve() if report_md_path else None
        stale = [path for path in (original_revisions, original_review_result, original_report_md) if path is not None and path.is_file()]
        if stale:
            raise ValueError(
                "检测到无 review-draft.json 但已有旧审核产物；为防止代际混用，请提供 draft 或显式 reuse_existing=True"
            )

    try:
        # Serialise publication and cache writes for a shared output folder.
        # The lock is held for the whole run so two invocations cannot publish
        # competing manifests or consume each other's preflight cache.
        with file_lock(out / ".pipeline.lock"):
            result = _run_pipeline(
            source,
            staged,
            draft_path=draft_arg,
            revisions_path=staged_revisions,
            review_result_path=staged_review_result,
            report_md_path=staged_report_md,
            reference_docx=reference_docx,
            revised_out=staged_revised,
            clean_out=staged_clean,
            report_out=staged_report,
            force_preflight=force_preflight,
            visual=visual,
            author=author,
            date_iso=date_iso,
            allow_existing_revisions=allow_existing_revisions,
            allow_semantic_warnings=allow_semantic_warnings,
            reuse_existing=reuse_existing,
            run_id=run_id,
            )
            return _publish_staged_tree(staged, out, run_id)
    except Exception as exc:
        # Keep a source-bound failure record without replacing the last
        # successful public manifest. The isolated run directory is retained
        # for forensic inspection until the next cleanup operation.
        _write_failure_manifest(staged, source, exc, run_id=run_id)
        failure_path = out / "work" / "pipeline-failures" / f"{run_id}.json"
        try:
            failure_path.parent.mkdir(parents=True, exist_ok=True)
            staged_failure = staged / "work" / "pipeline-report.json"
            if staged_failure.is_file():
                shutil.copy2(staged_failure, failure_path)
            else:
                temp_failure_root = out / "work" / "pipeline-failures" / f".{run_id}-tmp"
                _write_failure_manifest(temp_failure_root, source, exc, run_id=run_id)
                generated = temp_failure_root / "work" / "pipeline-report.json"
                if generated.is_file():
                    os.replace(generated, failure_path)
                shutil.rmtree(temp_failure_root, ignore_errors=True)
        except Exception:
            pass
        raise


def _render_citation_markdown(report: dict[str, Any]) -> str:
    """轻量输出核验表，避免流水线再次启动 CLI 解析同一文件。"""
    summary = report.get("summary") or {}
    lines = [
        "### 法条引用自动核验表",
        "",
        f"> 核验截止日：`{report.get('as_of') or '未指定（仅按知识库当前状态）'}`",
        "",
        f"> 核验条数：`{summary.get('total', 0)}` ｜ 扫描出现次数：`{summary.get('occurrences_scanned', 0)}` ｜ 已核验：`{summary.get('verified', 0)}` ｜ 库外法源：`{summary.get('out_of_kb', 0)}` ｜ 条号不存在：`{summary.get('article_not_found', 0)}` ｜ 截止日未生效：`{summary.get('not_yet_effective', 0)}`",
        "",
        "| 序号 | 引用原文 | 条号 | 结论 | 说明 |",
        "| :---: | :--- | :--- | :---: | :--- |",
    ]
    for idx, item in enumerate(report.get("items") or [], 1):
        raw = str(item.get("raw_citation") or "").replace("|", "｜")
        article = str(item.get("article_label") or "").replace("|", "｜")
        status = str(item.get("status_cn") or item.get("status") or "").replace("|", "｜")
        note = str(item.get("note") or "通过").replace("|", "｜")
        lines.append(f"| {idx} | {raw} | {article} | {status} | {note} |")
    return "\n".join(lines) + "\n"


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description="串联合同审核预检、单一台账派生、校验、修订和一次性版面检查")
    parser.add_argument("input", help="原始合同文件（.doc/.docx/.wps/.rtf）")
    parser.add_argument("--out-dir", required=True, help="本次审核输出目录")
    parser.add_argument("--draft", default=None, help="单一来源台账 JSON（默认自动检测 out/work/review-draft.json）")
    parser.add_argument("--revisions", default=None, help="修订指令 JSON，默认 out/work/revisions.json")
    parser.add_argument("--review-result", default=None, help="问题台账 JSON，默认 out/work/review-result.json")
    parser.add_argument("--report-md", default=None, help="意见书 Markdown；提供后自动核验并转为 docx")
    parser.add_argument("--reference", default=None, help="意见书参考模板 docx")
    parser.add_argument("--revised-out", default=None, help="修订版 docx 输出路径")
    parser.add_argument("--clean-out", default=None, help="清洁版 docx 输出路径")
    parser.add_argument("--report-out", default=None, help="意见书 docx 输出路径")
    parser.add_argument("--force-preflight", action="store_true", help="忽略预检缓存")
    parser.add_argument("--no-visual", dest="visual", action="store_false", help="跳过最终批量版面渲染")
    parser.add_argument(
        "--allow-existing-revisions",
        action="store_true",
        help="人工确认原文已有修订痕迹后继续；默认发现冲突即阻断",
    )
    parser.add_argument(
        "--allow-semantic-warnings",
        action="store_true",
        help="人工复核语义校验警告后继续；默认 fail-closed",
    )
    parser.add_argument(
        "--reuse-existing",
        action="store_true",
        help="显式允许无 review-draft.json 时复用已有审核台账",
    )
    parser.add_argument("--author", default=None, help="修订作者")
    parser.add_argument("--date", default=None, help="修订时间戳")
    args = parser.parse_args(argv)
    try:
        report = run_pipeline(
            args.input,
            args.out_dir,
            draft_path=args.draft,
            revisions_path=args.revisions,
            review_result_path=args.review_result,
            report_md_path=args.report_md,
            reference_docx=args.reference,
            revised_out=args.revised_out,
            clean_out=args.clean_out,
            report_out=args.report_out,
            force_preflight=args.force_preflight,
            visual=args.visual,
            author=args.author,
            date_iso=args.date,
            allow_existing_revisions=args.allow_existing_revisions,
            allow_semantic_warnings=args.allow_semantic_warnings,
            reuse_existing=args.reuse_existing,
        )
    except Exception as exc:
        print(json.dumps({"error": str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 1
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
