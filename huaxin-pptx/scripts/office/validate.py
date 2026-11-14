#!/usr/bin/env python3
"""Validate the basic ZIP/XML integrity of an Office package."""
from __future__ import annotations

import argparse
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile


def validate_package(package: Path, package_suffix: str | None = None) -> None:
    """Validate one packed Office package.

    This helper is intentionally side-effect free so ``pack.py`` can validate
    its temporary archive before publishing it to the requested output path.
    """
    if package.is_dir():
        raise ValueError("Use a packed .pptx/.docx/.xlsx file")
    with ZipFile(package) as archive:
        damaged = archive.testzip()
        if damaged:
            raise ValueError(f"ZIP CRC error: {damaged}")
        names = set(archive.namelist())
        required = {"[Content_Types].xml", "_rels/.rels"}
        suffix = (package_suffix or package.suffix).lower()
        if suffix == ".pptx": required.add("ppt/presentation.xml")
        missing = required - names
        if missing: raise ValueError(f"Missing package parts: {sorted(missing)}")
        for name in names:
            if name.endswith((".xml", ".rels")): ET.fromstring(archive.read(name))


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("path", type=Path)
    parser.add_argument("-v", "--verbose", action="store_true", help="Print the validated path (default).")
    args = parser.parse_args()
    package = args.path.resolve()
    try:
        validate_package(package)
    except (OSError, ValueError) as exc:
        raise SystemExit(str(exc)) from exc
    print(f"All validations PASSED: {package}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
