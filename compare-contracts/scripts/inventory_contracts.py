#!/usr/bin/env python3
"""Inventory contract files and identify PDFs that likely need OCR."""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from pathlib import Path


SUPPORTED_EXTENSIONS = {
    ".pdf",
    ".docx",
    ".doc",
    ".txt",
    ".md",
    ".png",
    ".jpg",
    ".jpeg",
    ".tif",
    ".tiff",
}
IGNORED_DIRECTORIES = {
    ".contract-analysis",
    ".contract-tools",
    ".git",
    ".tools",
    "extracted_text",
    "node_modules",
}
IGNORED_FILENAMES = {
    "agents.md",
    "claude.md",
    "contributing.md",
    "readme.md",
}


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def pdf_profile(path: Path, sample_pages: int) -> dict:
    try:
        from PyPDF2 import PdfReader
    except ImportError:
        return {
            "page_count": None,
            "embedded_text_sample_chars": None,
            "extraction_hint": "unknown",
            "note": "PyPDF2 is not installed",
        }

    try:
        reader = PdfReader(str(path))
        total_pages = len(reader.pages)
        indexes = list(range(min(sample_pages, total_pages)))
        if total_pages > sample_pages and total_pages - 1 not in indexes:
            indexes.append(total_pages - 1)
        counts = [len((reader.pages[index].extract_text() or "").strip()) for index in indexes]
        nonempty = sum(count >= 40 for count in counts)
        if not counts or nonempty == 0:
            hint = "scanned"
        elif nonempty == len(counts):
            hint = "embedded-text"
        else:
            hint = "mixed"
        return {
            "page_count": total_pages,
            "sampled_pages": [index + 1 for index in indexes],
            "embedded_text_sample_chars": sum(counts),
            "extraction_hint": hint,
        }
    except Exception as exc:  # noqa: BLE001
        return {
            "page_count": None,
            "embedded_text_sample_chars": None,
            "extraction_hint": "error",
            "note": str(exc),
        }


def iter_contracts(root: Path):
    for path in sorted(root.rglob("*"), key=lambda item: str(item).lower()):
        if (
            not path.is_file()
            or path.suffix.lower() not in SUPPORTED_EXTENSIONS
            or path.name.lower() in IGNORED_FILENAMES
        ):
            continue
        if any(part in IGNORED_DIRECTORIES for part in path.relative_to(root).parts[:-1]):
            continue
        yield path


def build_inventory(root: Path, sample_pages: int, include_hash: bool) -> dict:
    files = []
    for index, path in enumerate(iter_contracts(root), start=1):
        relative = path.relative_to(root)
        first_part = relative.parts[0] if len(relative.parts) > 1 else "ungrouped"
        entry = {
            "id": f"C{index:03d}",
            "group": first_part,
            "name": path.name,
            "relative_path": relative.as_posix(),
            "extension": path.suffix.lower(),
            "size_bytes": path.stat().st_size,
        }
        if include_hash:
            entry["sha256"] = sha256_file(path)
        if path.suffix.lower() == ".pdf":
            entry.update(pdf_profile(path, sample_pages))
        elif path.suffix.lower() == ".doc":
            entry["extraction_hint"] = "legacy-doc-conversion-required"
        elif path.suffix.lower() in {".png", ".jpg", ".jpeg", ".tif", ".tiff"}:
            entry["extraction_hint"] = "ocr-required"
        else:
            entry["extraction_hint"] = "direct-text"
        files.append(entry)

    groups = {}
    for entry in files:
        groups[entry["group"]] = groups.get(entry["group"], 0) + 1
    return {
        "root": str(root.resolve()),
        "supported_extensions": sorted(SUPPORTED_EXTENSIONS),
        "file_count": len(files),
        "groups": groups,
        "files": files,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path, help="Directory containing contract files")
    parser.add_argument("--output", type=Path, help="Write JSON inventory to this file")
    parser.add_argument("--sample-pages", type=int, default=3)
    parser.add_argument("--hash", action="store_true", help="Calculate SHA-256 for every file")
    args = parser.parse_args()

    root = args.input.resolve()
    if not root.is_dir():
        parser.error(f"input directory does not exist: {root}")

    inventory = build_inventory(root, max(1, args.sample_pages), args.hash)
    payload = json.dumps(inventory, ensure_ascii=False, indent=2)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(payload, encoding="utf-8")
        print(f"Wrote {inventory['file_count']} files to {args.output}")
    else:
        sys.stdout.write(payload + "\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
