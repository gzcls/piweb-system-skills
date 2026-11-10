#!/usr/bin/env python3
"""Render PPTX files with LibreOffice and write reproducible visual evidence.

The renderer only proves that LibreOffice produced a PDF and rasterized pages.
It does not infer visual quality.  A human must inspect the PNGs and rerun with
``--review-status pass`` before a validator report can claim visual evidence.
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


def parse_page_list(value: str | None) -> list[int]:
    if not value:
        return []
    pages: list[int] = []
    for item in value.split(","):
        item = item.strip()
        if not item:
            continue
        page = int(item)
        if page < 1:
            raise ValueError("reviewed page numbers are 1-based")
        if page not in pages:
            pages.append(page)
    return pages


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
    with tempfile.TemporaryDirectory(prefix="lo-profile-", dir=deck_dir) as profile:
        profile_uri = Path(profile).resolve().as_uri()
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


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--pptx",
        action="append",
        required=True,
        type=Path,
        help="PPTX to render; repeat for more than one deck.",
    )
    parser.add_argument(
        "--output-dir",
        required=True,
        type=Path,
        help="Output directory outside the skill package.",
    )
    parser.add_argument("--json", dest="json_output", type=Path, help="Write evidence JSON here.")
    parser.add_argument("--soffice", type=Path, help="LibreOffice soffice executable.")
    parser.add_argument(
        "--scale",
        type=float,
        default=1.5,
        help="PDF raster scale (72 points/inch × scale; default: 1.5).",
    )
    parser.add_argument(
        "--review-status",
        choices=("pending", "pass"),
        default="pending",
        help="Human PNG review status; pass only after inspecting every requested page.",
    )
    parser.add_argument(
        "--reviewed-pages",
        help="Comma-separated 1-based page numbers inspected for every deck.",
    )
    parser.add_argument("--review-notes", default="", help="Short human review note stored in evidence.")
    args = parser.parse_args()
    if args.scale <= 0:
        parser.error("--scale must be positive")

    output_root = args.output_dir.expanduser().resolve()
    output_root.mkdir(parents=True, exist_ok=True)
    package_root = Path(__file__).resolve().parents[2]
    if output_root == package_root or package_root in output_root.parents:
        parser.error("--output-dir must be outside the distributable skill directory")
    json_output = args.json_output.expanduser().resolve() if args.json_output else None
    if json_output is not None and (json_output == package_root or package_root in json_output.parents):
        parser.error("--json must be outside the distributable skill directory")
    try:
        reviewed_pages = parse_page_list(args.reviewed_pages)
        soffice = resolve_soffice(args.soffice)
        decks = [
            render_deck(soffice, path.expanduser().resolve(), output_root, args.scale)
            for path in args.pptx
        ]
        for deck in decks:
            page_count = deck["page_count"]
            if reviewed_pages and any(page > page_count for page in reviewed_pages):
                raise ValueError(
                    f"reviewed page exceeds {deck['pptx']} page count ({page_count})"
                )
            if args.review_status == "pass":
                expected_pages = set(range(1, page_count + 1))
                if set(reviewed_pages) != expected_pages:
                    raise ValueError(
                        f"--review-status pass requires every page of {deck['pptx']} "
                        f"to be listed in --reviewed-pages (expected 1-{page_count})"
                    )
        evidence: dict[str, Any] = {
            "schema_version": 1,
            "status": "PASS" if args.review_status == "pass" else "RENDERED",
            "render_status": "PASS",
            "renderer": {
                "name": "LibreOffice",
                "executable": str(soffice),
                "version": "unspecified; inspect executable metadata",
                "mode": "headless PDF conversion + PyMuPDF rasterization",
            },
            "raster": {"scale": args.scale, "pixels_per_inch": round(72 * args.scale, 3)},
            "manual_review": {
                "status": "PASS" if args.review_status == "pass" else "PENDING",
                "pages": reviewed_pages,
                "notes": args.review_notes,
            },
            "generated_at_utc": datetime.now(timezone.utc).isoformat(),
            "decks": decks,
        }
        serialized = json.dumps(evidence, ensure_ascii=False, indent=2) + "\n"
        if json_output is not None:
            json_output.parent.mkdir(parents=True, exist_ok=True)
            json_output.write_text(serialized, encoding="utf-8")
        print(serialized, end="")
        return 0
    except (FileNotFoundError, OSError, RuntimeError, ValueError, subprocess.TimeoutExpired) as exc:
        print(f"visual render failed: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
