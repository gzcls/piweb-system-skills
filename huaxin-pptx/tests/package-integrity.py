#!/usr/bin/env python3
"""Deterministic checks for the distributable huaxin-pptx skill package.

This test intentionally uses only the Python standard library. It verifies
that package identity/version stay aligned and that the Office pack helper
filters development-only files before they enter an archive.
"""
from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
import tempfile
from pathlib import Path
from zipfile import ZipFile


SKILL_NAME = "huaxin-pptx"
EXCLUDED_DIRECTORY_NAMES = {".git", ".pytest_cache", "__pycache__", "node_modules"}
EXCLUDED_SUFFIXES = {".pyc", ".pyo"}


def check_metadata(root: Path) -> None:
    skill = (root / "SKILL.md").read_text(encoding="utf-8")
    name = re.search(r"^name:\s*(\S+)", skill, re.M)
    version = re.search(r'^\s+version:\s*"([^"]+)"', skill, re.M)
    if not name or name.group(1) != SKILL_NAME or root.name != SKILL_NAME:
        raise AssertionError("SKILL.md name and directory must both be huaxin-pptx")
    if not version:
        raise AssertionError("SKILL.md must declare metadata.version")
    manifest = json.loads((root / "references/assets-manifest.json").read_text(encoding="utf-8"))
    if manifest.get("skill_name") != SKILL_NAME or manifest.get("version") != version.group(1):
        raise AssertionError("assets-manifest.json skill identity/version is out of sync with SKILL.md")
    agent = (root / "agents/openai.yaml").read_text(encoding="utf-8")
    if f"${SKILL_NAME}" not in agent:
        raise AssertionError("agents/openai.yaml default prompt must reference $huaxin-pptx")
    for retired in ("package.json", "package-lock.json", "pptxgenjs.md"):
        if (root / retired).exists():
            raise AssertionError(f"retired file still present: {retired}")


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
