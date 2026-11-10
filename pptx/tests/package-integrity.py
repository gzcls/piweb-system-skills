#!/usr/bin/env python3
"""Deterministic checks for the distributable PPTX skill package.

This test intentionally uses only the Python standard library. It verifies
that package metadata stays aligned and that the Office pack helper filters
development-only files before they enter an archive.
"""
from __future__ import annotations

import argparse
import json
import subprocess
import sys
import tempfile
from pathlib import Path
from zipfile import ZipFile


EXCLUDED_DIRECTORY_NAMES = {".git", ".pytest_cache", "__pycache__", "node_modules"}
EXCLUDED_SUFFIXES = {".pyc", ".pyo"}


def check_metadata(root: Path) -> None:
    package = json.loads((root / "package.json").read_text(encoding="utf-8"))
    lock = json.loads((root / "package-lock.json").read_text(encoding="utf-8"))
    manifest = json.loads((root / "references/assets-manifest.json").read_text(encoding="utf-8"))

    version = package.get("version")
    if not isinstance(version, str) or not version:
        raise AssertionError("package.json must declare a non-empty version")
    if lock.get("version") != version or lock.get("packages", {}).get("", {}).get("version") != version:
        raise AssertionError("package-lock.json root version is out of sync with package.json")
    if manifest.get("skill_name") != "pptx" or manifest.get("version") != version:
        raise AssertionError("assets-manifest.json skill identity/version is out of sync")
    package_dependencies = package.get("dependencies", {})
    lock_dependencies = lock.get("packages", {}).get("", {}).get("dependencies", {})
    if package_dependencies != lock_dependencies:
        raise AssertionError("package-lock.json root dependencies are out of sync with package.json")


def check_no_development_artifacts(root: Path) -> None:
    offenders: list[str] = []
    for path in root.rglob("*"):
        relative = path.relative_to(root)
        if any(part in EXCLUDED_DIRECTORY_NAMES for part in relative.parts):
            offenders.append(relative.as_posix())
        elif path.is_file() and path.suffix.lower() in EXCLUDED_SUFFIXES:
            offenders.append(relative.as_posix())
    if offenders:
        raise AssertionError("development artifacts found: " + ", ".join(sorted(offenders)))


def check_pack_filter(root: Path) -> None:
    packer = root / "scripts/office/pack.py"
    with tempfile.TemporaryDirectory(prefix="pptx-skill-pack-test-") as temporary:
        workspace = Path(temporary)
        source = workspace / "office-tree"
        # Keep the output inside the source tree to catch temporary-file or
        # stale-output self-inclusion during packing.
        output = source / "output.pptx"
        (source / "ppt").mkdir(parents=True)
        (source / "node_modules/pkg").mkdir(parents=True)
        (source / "__pycache__").mkdir(parents=True)
        (source / "[Content_Types].xml").write_text("<Types/>", encoding="utf-8")
        (source / "_rels/.rels").parent.mkdir(parents=True, exist_ok=True)
        (source / "_rels/.rels").write_text("<Relationships/>", encoding="utf-8")
        (source / "ppt/presentation.xml").write_text(
            '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>',
            encoding="utf-8",
        )
        (source / "node_modules/pkg/index.js").write_text("development dependency", encoding="utf-8")
        (source / "__pycache__/helper.pyc").write_bytes(b"bytecode")
        (source / "scratch.pyc").write_bytes(b"bytecode")

        subprocess.run(
            [sys.executable, str(packer), str(source), str(output)],
            check=True,
            capture_output=True,
            text=True,
        )
        with ZipFile(output) as archive:
            names = set(archive.namelist())
        if "ppt/presentation.xml" not in names:
            raise AssertionError("pack helper omitted a regular Office part")
        if "output.pptx" in names or any(name.startswith(".output.pptx.") for name in names):
            raise AssertionError("pack helper included its output or temporary file")
        forbidden = {"node_modules/pkg/index.js", "__pycache__/helper.pyc", "scratch.pyc"}
        leaked = forbidden & names
        if leaked:
            raise AssertionError("pack helper leaked development files: " + ", ".join(sorted(leaked)))


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    args = parser.parse_args()
    root = args.root.expanduser().resolve()
    check_metadata(root)
    check_no_development_artifacts(root)
    check_pack_filter(root)
    print(f"Package integrity checks PASSED: {root}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
