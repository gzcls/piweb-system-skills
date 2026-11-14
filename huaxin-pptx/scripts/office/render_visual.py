#!/usr/bin/env python3
"""Render PPTX files with LibreOffice and write reproducible visual evidence.

The renderer only proves that LibreOffice produced a PDF and rasterized pages
(status=RENDERED).  It does not infer visual quality.  After every PNG has been
inspected, run ``--attest evidence.json --reviewed-pages all --review-notes ...``;
only then does the evidence carry manual_review.status=PASS.
All outputs belong outside the distributable skill directory.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

sys.dont_write_bytecode = True


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def resolve_soffice(explicit: Path | None) -> Path:
    candidates: list[Path] = []
    if explicit is not None:
        candidates.append(explicit.expanduser())
    found = shutil.which("soffice")
    if found:
        candidates.append(Path(found))
    program_files = os.environ.get("PROGRAMFILES")
    program_files_x86 = os.environ.get("PROGRAMFILES(X86)")
    for base in (program_files, program_files_x86):
        if base:
            candidates.append(Path(base) / "LibreOffice" / "program" / "soffice.exe")
    candidates.extend(
        [
            Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
            Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
        ]
    )
    for candidate in candidates:
        try:
            resolved = candidate.resolve()
        except OSError:
            continue
        if resolved.is_file():
            return resolved
    raise FileNotFoundError(
        "LibreOffice soffice executable was not found; pass --soffice or install LibreOffice."
    )


def import_pdf_renderer() -> Any:
    try:
        import pymupdf  # type: ignore

        return pymupdf
    except ImportError:
        try:
            import fitz  # type: ignore

            return fitz
        except ImportError as exc:
            raise RuntimeError(
                "PyMuPDF is required to rasterize LibreOffice PDF pages; "
                "install requirements.txt (PyMuPDF)."
            ) from exc


def render_deck(
    soffice: Path,
    pptx: Path,
    output_root: Path,
    scale: float,
) -> dict[str, Any]:
    if not pptx.is_file():
        raise FileNotFoundError(f"PPTX does not exist: {pptx}")
    if pptx.suffix.lower() != ".pptx":
        raise ValueError(f"Input must have a .pptx extension: {pptx}")
    deck_dir = output_root / pptx.stem
    deck_dir.mkdir(parents=True, exist_ok=True)
    pdf_path = deck_dir / f"{pptx.stem}.pdf"
    if pdf_path.exists():
        pdf_path.unlink()
    # LibreOffice may keep files in its profile briefly after exiting on Windows; use a
    # profile outside the output tree and tolerate cleanup errors instead of failing the render.
    profile = Path(tempfile.mkdtemp(prefix="lo-profile-"))
    try:
        profile_uri = profile.resolve().as_uri()
        command = [
            str(soffice),
            "--headless",
            f"-env:UserInstallation={profile_uri}",
            "--convert-to",
            "pdf",
            "--outdir",
            str(deck_dir),
            str(pptx),
        ]
        completed = subprocess.run(
            command,
            check=False,
            capture_output=True,
            text=True,
            timeout=180,
        )
    finally:
        shutil.rmtree(profile, ignore_errors=True)
    if completed.returncode != 0:
        raise RuntimeError(
            f"LibreOffice conversion failed ({completed.returncode}): "
            f"{completed.stderr.strip() or completed.stdout.strip()}"
        )
    if not pdf_path.is_file():
        raise RuntimeError(
            f"LibreOffice returned success but did not create {pdf_path}; "
            f"stdout={completed.stdout.strip()!r} stderr={completed.stderr.strip()!r}"
        )

    pdf_renderer = import_pdf_renderer()
    matrix = pdf_renderer.Matrix(scale, scale)
    png_dir = deck_dir / "png"
    png_dir.mkdir(exist_ok=True)
    for stale in png_dir.glob("page-*.png"):
        stale.unlink()
    pages: list[dict[str, Any]] = []
    document = pdf_renderer.open(pdf_path)
    try:
        for index, page in enumerate(document, start=1):
            pixmap = page.get_pixmap(matrix=matrix, alpha=False)
            png_path = png_dir / f"page-{index}.png"
            pixmap.save(str(png_path))
            pages.append(
                {
                    "page": index,
                    "width_pt": round(float(page.rect.width), 3),
                    "height_pt": round(float(page.rect.height), 3),
                    "raster_width_px": pixmap.width,
                    "raster_height_px": pixmap.height,
                    "png": str(png_path.resolve()),
                    "png_sha256": sha256(png_path),
                }
            )
    finally:
        document.close()
    return {
        "pptx": str(pptx.resolve()),
        "pptx_sha256": sha256(pptx),
        "pdf": str(pdf_path.resolve()),
        "pdf_sha256": sha256(pdf_path),
        "page_count": len(pages),
        "pages": pages,
        "conversion": {
            "returncode": completed.returncode,
            "stdout": completed.stdout.strip(),
            "stderr": completed.stderr.strip(),
        },
    }


def parse_review_pages(value: str | None, page_count: int) -> list[int]:
    """Parse 'all', '1-12' or '1,2,5-7' into a sorted page list."""
    if not value:
        return []
    if value.strip().lower() == "all":
        return list(range(1, page_count + 1))
    pages: set[int] = set()
    for item in value.split(","):
        item = item.strip()
        if not item:
            continue
        if "-" in item:
            start, end = (int(part) for part in item.split("-", 1))
            pages.update(range(start, end + 1))
        else:
            pages.add(int(item))
    if any(page < 1 or page > page_count for page in pages):
        raise ValueError(f"reviewed pages must be within 1-{page_count}")
    return sorted(pages)


def attest(evidence_path: Path, reviewed: str | None, notes: str, json_output: Path | None) -> dict[str, Any]:
    """Step 2: after a person/agent has looked at every PNG, record the review.

    The evidence must come from an earlier render run and every PPTX/PNG hash must still
    match, so a review cannot be attached to a deck that changed after rendering.
    """
    evidence = json.loads(evidence_path.read_text(encoding="utf-8"))
    if evidence.get("render_status") != "PASS" or not evidence.get("decks"):
        raise ValueError("evidence is not a successful render manifest")
    if len(notes.strip()) < 4:
        raise ValueError("--review-notes is required: summarise what was checked per page and any fixes")
    counts = {deck["page_count"] for deck in evidence["decks"]}
    if len(counts) != 1:
        raise ValueError("attest one deck per evidence file (decks have different page counts)")
    page_count = counts.pop()
    pages = parse_review_pages(reviewed, page_count)
    if pages != list(range(1, page_count + 1)):
        raise ValueError(f"every page must be reviewed before PASS (expected 1-{page_count}, got {pages})")
    for deck in evidence["decks"]:
        if sha256(Path(deck["pptx"])) != deck["pptx_sha256"]:
            raise ValueError(f"{deck['pptx']} changed after rendering; render again before attesting")
        for page in deck["pages"]:
            if sha256(Path(page["png"])) != page["png_sha256"]:
                raise ValueError(f"{page['png']} changed after rendering; render again before attesting")
    evidence["status"] = "PASS"
    evidence["manual_review"] = {
        "status": "PASS",
        "pages": pages,
        "notes": notes.strip(),
        "reviewed_at_utc": datetime.now(timezone.utc).isoformat(),
    }
    target = json_output or evidence_path
    target.write_text(json.dumps(evidence, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return evidence


def main() -> int:
    parser = argparse.ArgumentParser(
        description=__doc__,
        epilog=(
            "Two steps: (1) render: --pptx deck.pptx --output-dir DIR --json evidence.json; "
            "(2) inspect every PNG, then: --attest evidence.json --reviewed-pages all "
            "--review-notes \"what was checked / fixed\"."
        ),
    )
    parser.add_argument("--pptx", action="append", type=Path, help="PPTX to render; repeat for more decks.")
    parser.add_argument("--output-dir", type=Path, help="Output directory outside the skill package.")
    parser.add_argument("--json", dest="json_output", type=Path, help="Write evidence JSON here.")
    parser.add_argument("--soffice", type=Path, help="LibreOffice soffice executable.")
    parser.add_argument("--scale", type=float, default=1.5,
                        help="PDF raster scale (72 points/inch × scale; default: 1.5).")
    parser.add_argument("--attest", type=Path,
                        help="Step 2: mark an existing evidence JSON as reviewed after inspecting every PNG.")
    parser.add_argument("--reviewed-pages", help="'all', or pages/ranges such as 1-12 (attest mode).")
    parser.add_argument("--review-notes", default="", help="Required in attest mode: what was checked / fixed.")
    parser.add_argument("--review-status", choices=("pending", "pass"), help=argparse.SUPPRESS)
    args = parser.parse_args()

    package_root = Path(__file__).resolve().parents[2]

    def outside_package(path: Path, label: str) -> Path:
        resolved = path.expanduser().resolve()
        if resolved == package_root or package_root in resolved.parents:
            parser.error(f"{label} must be outside the distributable skill directory")
        return resolved

    json_output = outside_package(args.json_output, "--json") if args.json_output else None
    try:
        if args.attest:
            evidence = attest(args.attest.expanduser().resolve(), args.reviewed_pages, args.review_notes,
                              json_output)
            print(json.dumps({"status": evidence["status"], "manual_review": evidence["manual_review"]},
                             ensure_ascii=False, indent=2))
            return 0
        if args.review_status == "pass" or args.reviewed_pages:
            parser.error("render first, inspect every PNG, then run --attest EVIDENCE.json --reviewed-pages all")
        if not args.pptx or not args.output_dir:
            parser.error("render mode requires --pptx and --output-dir")
        if args.scale <= 0:
            parser.error("--scale must be positive")
        output_root = outside_package(args.output_dir, "--output-dir")
        output_root.mkdir(parents=True, exist_ok=True)
        soffice = resolve_soffice(args.soffice)
        decks = [render_deck(soffice, path.expanduser().resolve(), output_root, args.scale) for path in args.pptx]
        evidence: dict[str, Any] = {
            "schema_version": 1,
            "status": "RENDERED",
            "render_status": "PASS",
            "renderer": {
                "name": "LibreOffice",
                "executable": str(soffice),
                "version": "unspecified; inspect executable metadata",
                "mode": "headless PDF conversion + PyMuPDF rasterization",
            },
            "raster": {"scale": args.scale, "pixels_per_inch": round(72 * args.scale, 3)},
            "manual_review": {"status": "PENDING", "pages": [], "notes": ""},
            "generated_at_utc": datetime.now(timezone.utc).isoformat(),
            "decks": decks,
        }
        serialized = json.dumps(evidence, ensure_ascii=False, indent=2) + "\n"
        if json_output is not None:
            json_output.parent.mkdir(parents=True, exist_ok=True)
            json_output.write_text(serialized, encoding="utf-8")
        print(serialized, end="")
        return 0
    except (FileNotFoundError, OSError, RuntimeError, ValueError, KeyError, subprocess.TimeoutExpired) as exc:
        print(f"visual render failed: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
