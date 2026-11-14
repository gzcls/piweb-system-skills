#!/usr/bin/env python3
"""Delete leftover temporary files (*.tmp, *.bak) from an unpacked Office package.

It does NOT remove unreferenced parts; drop unused relationships/parts manually.
"""
from __future__ import annotations

import argparse
from pathlib import Path


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("unpacked_dir", type=Path)
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()
    root = args.unpacked_dir.resolve()
    if not root.is_dir() or root.name in {"", ".", ".."}: parser.error("Invalid unpacked directory")
    candidates = [p for p in root.rglob("*") if p.is_file() and p.suffix.lower() in {".tmp", ".bak"}]
    for path in candidates:
        print(("would remove " if args.dry_run else "remove ") + str(path))
        if not args.dry_run: path.unlink()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
