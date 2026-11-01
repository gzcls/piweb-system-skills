#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""可选的 DOCX 渲染与版面风险检查（clean-room 实现）。

该模块只把渲染结果作为诊断信号：OOXML/ZIP 完整性仍由 docx_verify/lib.ooxml 负责。
LibreOffice 和 PyMuPDF 均为可选依赖；缺少时返回 skipped，而不是伪造通过。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path
from typing import Any
from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def _find_soffice() -> str | None:
    for name in ("soffice.com", "soffice.exe", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    candidates = (
        Path(os.environ.get("PROGRAMFILES", "")) / "LibreOffice/program/soffice.com",
        Path(os.environ.get("PROGRAMFILES(X86)", "")) / "LibreOffice/program/soffice.com",
    )
    return next((str(p) for p in candidates if p.exists()), None)


def _render_pdf(docx_path: Path, work: Path) -> tuple[Path | None, dict[str, Any]]:
    soffice = _find_soffice()
    if not soffice:
        return None, {"status": "skipped", "reason": "LibreOffice/soffice not found"}
    profile = work / "lo-profile"
    profile.mkdir(parents=True, exist_ok=True)
    out = work / "pdf"
    out.mkdir(parents=True, exist_ok=True)
    cmd = [soffice, "--headless", "--norestore", "--nolockcheck", f"-env:UserInstallation={profile.as_uri()}", "--convert-to", "pdf", "--outdir", str(out), str(docx_path)]
    try:
        proc = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=180,
        )
    except subprocess.TimeoutExpired as exc:
        return None, {
            "status": "error",
            "kind": "render_timeout",
            "timeout_seconds": 180,
            "message": f"LibreOffice 渲染超时: {exc}",
        }
    except OSError as exc:
        return None, {"status": "error", "kind": "render_exec", "message": str(exc)}
    pdf = out / (docx_path.stem + ".pdf")
    if proc.returncode != 0 or not pdf.exists():
        return None, {"status": "error", "returncode": proc.returncode, "stderr": proc.stderr[-2000:], "stdout": proc.stdout[-2000:]}
    return pdf, {"status": "ok", "renderer": "LibreOffice", "command": cmd}


def _xml_layout_findings(docx_path: Path) -> list[dict[str, Any]]:
    findings: list[dict[str, Any]] = []
    try:
        with zipfile.ZipFile(docx_path) as zf:
            document = etree.fromstring(zf.read("word/document.xml"))
    except Exception as exc:
        return [{"severity": "error", "kind": "xml_read", "message": str(exc)}]
    def usable_width(sect: etree._Element | None) -> int | None:
        if sect is None:
            return None
        pg = sect.find("w:pgSz", NS)
        mar = sect.find("w:pgMar", NS)
        if pg is None or mar is None:
            return None
        try:
            return (
                int(pg.get(f"{{{W_NS}}}w", "12240"))
                - int(mar.get(f"{{{W_NS}}}left", "1440"))
                - int(mar.get(f"{{{W_NS}}}right", "1440"))
            )
        except (TypeError, ValueError):
            return None

    body = document.find("w:body", NS)
    if body is None:
        return [{"severity": "error", "kind": "xml_body", "message": "document.xml 缺少 w:body"}]
    # A section break is attached to the paragraph that ends the preceding
    # section. Therefore a table before a paragraph-level sectPr belongs to
    # that sectPr's width; content after the break belongs to the body-level
    # sectPr (or the next break). Apply widths to each body region rather than
    # reusing the first section for every table.
    final_sect = body.find("w:sectPr", NS)
    final_width = usable_width(final_sect)
    table_widths: list[tuple[int | None, int]] = []
    pending_tables: list[etree._Element] = []
    section_no = 1

    def flush_region(width: int | None, number: int) -> None:
        for table in pending_tables:
            table_widths.extend((width, number) for _ in table.iter(f"{{{W_NS}}}tbl"))
        pending_tables.clear()

    for child in body:
        if child.tag == f"{{{W_NS}}}tbl":
            pending_tables.append(child)
            continue
        if child.tag == f"{{{W_NS}}}p":
            p_sect = child.find("w:pPr/w:sectPr", NS)
            if p_sect is not None:
                flush_region(usable_width(p_sect), section_no)
                section_no += 1
    flush_region(final_width, section_no)

    for idx, tbl in enumerate(document.findall(".//w:tbl", NS), 1):
        usable, table_section = table_widths[idx - 1] if idx <= len(table_widths) else (final_width, section_no)
        grid = tbl.find("w:tblGrid", NS)
        width = (
            sum(int(c.get(f"{{{W_NS}}}w", "0")) for c in grid.findall("w:gridCol", NS))
            if grid is not None
            else 0
        )
        layout = tbl.find("w:tblPr/w:tblLayout", NS)
        autofit = layout is None or layout.get(f"{{{W_NS}}}type", "autofit") == "autofit"
        if usable and width > usable:
            findings.append({"severity": "warning", "kind": "table_width", "table": idx, "section": table_section, "grid_width_twips": width, "usable_twips": usable, "auto_fit": autofit, "message": "OOXML 原始表格网格宽度超过所在分节可用宽度；需以 PDF 几何检查确认"})
    return findings


def _pdf_findings(pdf_path: Path) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    try:
        try:
            import pymupdf as fitz  # type: ignore
        except ImportError:
            import fitz  # type: ignore
    except ImportError:
        return [{"severity": "warning", "kind": "dependency", "message": "PyMuPDF 未安装，跳过 PDF 几何检查"}], {"status": "skipped", "reason": "PyMuPDF not installed"}
    findings: list[dict[str, Any]] = []
    font_counts: dict[str, int] = {}
    with fitz.open(pdf_path) as doc:
        pages = len(doc)
        for pno, page in enumerate(doc, 1):
            blocks = page.get_text("blocks")
            drawings = page.get_drawings()
            images = page.get_images(full=True)
            if not blocks and not drawings and not images:
                findings.append({"severity": "error", "kind": "blank_page", "page": pno, "message": "页面无可见文本、图形或图片"})
            rect = page.rect
            for block in blocks:
                x0, y0, x1, y1 = block[:4]
                tol = 3
                if x0 < -tol or y0 < -tol or x1 > rect.width + tol or y1 > rect.height + tol:
                    findings.append({"severity": "error", "kind": "out_of_bounds", "page": pno, "bbox": [x0, y0, x1, y1], "page_size": [rect.width, rect.height]})
                text = str(block[4]).strip()
                if text and len(text) < 80 and y1 > rect.height * 0.88 and (text.startswith(("一、", "二、", "三、", "四、", "五、", "1.", "2.")) or text.endswith(("审查意见", "合同", "条款"))):
                    findings.append({"severity": "warning", "kind": "orphan_heading", "page": pno, "text": text[:80], "message": "疑似标题落在页面底部，建议人工复核"})
            try:
                for block in blocks:
                    raw = page.get_text("dict", clip=fitz.Rect(block[:4]))
                    for line in raw.get("blocks", []):
                        for span in line.get("lines", []):
                            for item in span.get("spans", []):
                                name = item.get("font")
                                if name:
                                    font_counts[name] = font_counts.get(name, 0) + 1
            except Exception:
                pass
    return findings, {"status": "ok", "pages": pages, "fonts": font_counts}


def _visual_cache_key(path: Path, require_render: bool, fail_on_warnings: bool) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    soffice = _find_soffice()
    renderer = "missing"
    if soffice:
        try:
            stat = Path(soffice).stat()
            renderer = f"{soffice}:{stat.st_mtime_ns}:{stat.st_size}"
        except OSError:
            renderer = soffice
    try:
        import fitz  # type: ignore
        pymupdf = getattr(fitz, "__doc__", "")[:40]
    except Exception:
        pymupdf = "missing"
    font_fingerprint = _font_environment_fingerprint()
    digest.update(json.dumps({"renderer": renderer, "pymupdf": pymupdf, "fonts": font_fingerprint, "require_render": require_render, "fail_on_warnings": fail_on_warnings}, sort_keys=True).encode())
    return digest.hexdigest()


def _font_environment_fingerprint() -> str:
    """Fingerprint installed font files that can change LibreOffice layout."""
    roots: list[Path] = []
    if os.name == "nt":
        windir = os.environ.get("WINDIR")
        if windir:
            roots.append(Path(windir) / "Fonts")
    roots.extend([Path("/usr/share/fonts"), Path("/usr/local/share/fonts"), Path.home() / ".fonts"])
    digest = hashlib.sha256()
    for root in roots:
        if not root.is_dir():
            continue
        digest.update(str(root).encode("utf-8", "ignore"))
        try:
            files = sorted(p for p in root.rglob("*") if p.is_file() and p.suffix.lower() in {".ttf", ".otf", ".ttc", ".fon"})
            for font in files:
                try:
                    stat = font.stat()
                    digest.update(f"{font.relative_to(root)}:{stat.st_size}:{stat.st_mtime_ns}".encode("utf-8", "ignore"))
                except OSError:
                    continue
        except OSError:
            continue
    return digest.hexdigest()


def _file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def inspect_docx(
    docx_path: str | Path,
    *,
    work_dir: str | Path | None = None,
    require_render: bool = False,
    fail_on_warnings: bool = False,
    cache_dir: str | Path | None = None,
) -> dict[str, Any]:
    path = Path(docx_path).resolve()
    if path.suffix.lower() != ".docx":
        return {"status": "error", "path": str(path), "findings": [{"severity": "error", "kind": "extension", "message": "输入必须是 .docx"}], "passed": False}
    cache_path: Path | None = None
    cache_key = None
    if cache_dir is not None:
        cache_dir = Path(cache_dir)
        cache_dir.mkdir(parents=True, exist_ok=True)
        cache_key = _visual_cache_key(path, require_render, fail_on_warnings)
        cache_path = cache_dir / f"{cache_key}.json"
        if cache_path.is_file():
            try:
                cached = json.loads(cache_path.read_text(encoding="utf-8"))
                cached_pdf = cached.get("pdf_cache_path")
                pdf_ready = cached.get("status") != "ok" or (cached_pdf and Path(str(cached_pdf)).is_file())
                if cached.get("cache_key") == cache_key and cached.get("source_sha256") == _file_sha256(path) and pdf_ready:
                    cached["path"] = str(path)
                    cached["cache_hit"] = True
                    return cached
            except Exception:
                pass
    own = work_dir is None
    work = Path(work_dir) if work_dir else Path(tempfile.mkdtemp(prefix="docx_visual_"))
    try:
        findings = _xml_layout_findings(path)
        pdf, render = _render_pdf(path, work)
        pdf_info: dict[str, Any] = {}
        if pdf:
            pdf_findings, pdf_info = _pdf_findings(pdf)
            findings.extend(pdf_findings)
        elif render.get("status") == "skipped":
            findings.append({
                "severity": "warning",
                "kind": "render_skipped",
                "message": render.get("reason", "渲染依赖不可用"),
            })
        elif render.get("status") == "error":
            findings.append({
                "severity": "error",
                "kind": render.get("kind", "render_error"),
                "message": render.get("message") or render.get("stderr") or "DOCX 渲染失败",
            })
        hard_errors = [f for f in findings if f.get("severity") == "error"]
        warnings = [f for f in findings if f.get("severity") == "warning"]
        render_status = render.get("status", "error")
        render_required_failure = require_render and render_status != "ok"
        result = {
            "status": render_status,
            "path": str(path),
            "render": render,
            "pdf": pdf_info,
            "findings": findings,
            "warning_count": len(warnings),
            "error_count": len(hard_errors),
            "render_required": require_render,
            "passed": not hard_errors and not render_required_failure and not (fail_on_warnings and warnings),
        }
        if cache_path is not None and cache_key is not None:
            result["cache_hit"] = False
            result["cache_key"] = cache_key
            result["source_sha256"] = _file_sha256(path)
            if pdf is not None:
                cached_pdf = cache_path.with_suffix(".pdf")
                try:
                    shutil.copy2(pdf, cached_pdf)
                    result["pdf_cache_path"] = str(cached_pdf)
                except OSError:
                    result["pdf_cache_path"] = None
            temporary = cache_path.with_name(f".{cache_path.name}.{os.getpid()}.tmp")
            try:
                temporary.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
                os.replace(temporary, cache_path)
            finally:
                temporary.unlink(missing_ok=True)
        return result
    finally:
        if own:
            shutil.rmtree(work, ignore_errors=True)


def main() -> int:
    parser = argparse.ArgumentParser(description="DOCX 可选版面/渲染检查")
    parser.add_argument("docx")
    parser.add_argument("--work-dir")
    parser.add_argument("--require-render", action="store_true", help="渲染依赖缺失或渲染失败时判定失败")
    parser.add_argument("--fail-on-warnings", action="store_true", help="存在版面警告时判定失败")
    parser.add_argument("--json", action="store_true", help="输出 JSON")
    args = parser.parse_args()
    result = inspect_docx(args.docx, work_dir=args.work_dir, require_render=args.require_render, fail_on_warnings=args.fail_on_warnings)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result.get("passed") else 1


if __name__ == "__main__":
    raise SystemExit(main())
