#!/usr/bin/env python3
"""Pack an unpacked Office directory into a ZIP-based Office file."""
from __future__ import annotations

import argparse
import os
import subprocess
import sys
import tempfile
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

EXCLUDED_DIRECTORY_NAMES = {
    ".git",
    ".pytest_cache",
    "__pycache__",
    "node_modules",
}
EXCLUDED_SUFFIXES = {".pyc", ".pyo"}


def should_include(path: Path) -> bool:
    """Return whether *path* belongs in a distributable Office package.

    Office package trees are generated artifacts. Development dependency trees
    and Python bytecode can be present when a package is prepared from a
    working directory, but are never valid Office parts and make the archive
    non-reproducible. Filter them at the pack boundary so callers do not need
    to remember a cleanup step.
    """
    if any(part in EXCLUDED_DIRECTORY_NAMES for part in path.parts):
        return False
    return path.suffix.lower() not in EXCLUDED_SUFFIXES


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input_directory", type=Path)
    parser.add_argument("output_file", type=Path)
    validation = parser.add_mutually_exclusive_group()
    validation.add_argument(
        "--validate",
        dest="validate",
        action="store_true",
        default=True,
        help="Validate the temporary archive before publishing it (default).",
    )
    validation.add_argument(
        "--no-validate",
        dest="validate",
        action="store_false",
        help="Skip ZIP/XML validation; use only for a deliberately partial test fixture.",
    )
    parser.add_argument("--force", action="store_true")
    args = parser.parse_args()
    source = args.input_directory.resolve()
    output = args.output_file.resolve()
    if not source.is_dir(): parser.error(f"Input directory does not exist: {source}")
    if output.exists() and not args.force: parser.error(f"Output exists; use --force: {output}")
    output.parent.mkdir(parents=True, exist_ok=True)
    # Snapshot the source before creating our temporary archive. If the output
    # lives inside the source tree, this prevents the temporary file (or an
    # existing output from a previous run) from being packed into itself.
    entries = [
        path for path in sorted(source.rglob("*"))
        if path.is_file()
        and path.resolve() != output
        and should_include(path.relative_to(source))
    ]
    temporary_suffix = output.suffix or ".tmp"
    fd, temporary = tempfile.mkstemp(prefix=f".{output.name}.", suffix=temporary_suffix, dir=output.parent)
    os.close(fd)
    try:
        with ZipFile(temporary, "w", ZIP_DEFLATED) as archive:
            for path in entries:
                archive.write(path, path.relative_to(source).as_posix())
        if args.validate:
            validator = Path(__file__).with_name("validate.py")
            result = subprocess.run(
                [sys.executable, str(validator), temporary],
                check=False,
                capture_output=True,
                text=True,
            )
            if result.returncode:
                detail = (result.stderr or result.stdout).strip()
                raise RuntimeError(f"Archive validation failed: {detail}")
        os.replace(temporary, output)
    finally:
        if os.path.exists(temporary): os.unlink(temporary)
    print(f"Successfully packed {source} to {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
