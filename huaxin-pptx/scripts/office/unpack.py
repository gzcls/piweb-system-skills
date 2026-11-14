#!/usr/bin/env python3
"""Safely unpack a DOCX/PPTX/XLSX ZIP package."""
from __future__ import annotations

import argparse
import shutil
from pathlib import Path, PurePosixPath
from zipfile import ZipFile


def safe_name(name: str) -> Path:
    path = PurePosixPath(name)
    if path.is_absolute() or ".." in path.parts or "\\" in name:
        raise ValueError(f"Unsafe archive path: {name}")
    return Path(*path.parts)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input_file", type=Path)
    parser.add_argument("output_directory", type=Path)
    args = parser.parse_args()
    output = args.output_directory.resolve()
    if output.exists():
        parser.error(f"Output directory already exists: {output}")
    output.mkdir(parents=True)
    try:
        with ZipFile(args.input_file) as archive:
            names = archive.namelist()
            if len(names) != len(set(names)):
                raise ValueError("Duplicate ZIP entries are not allowed")
            for info in archive.infolist():
                relative = safe_name(info.filename)
                target = (output / relative).resolve()
                if output not in target.parents and target != output:
                    raise ValueError(f"Archive path escaped output: {info.filename}")
                if info.is_dir():
                    target.mkdir(parents=True, exist_ok=True)
                    continue
                target.parent.mkdir(parents=True, exist_ok=True)
                with archive.open(info) as source, target.open("wb") as destination:
                    shutil.copyfileobj(source, destination)
    except Exception:
        shutil.rmtree(output, ignore_errors=True)
        raise
    print(f"Unpacked {args.input_file} to {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
