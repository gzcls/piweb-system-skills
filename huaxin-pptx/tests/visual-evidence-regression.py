#!/usr/bin/env python3
"""Regression checks for the two-step visual evidence flow (render -> attest).

Does not need LibreOffice: it fabricates a rendered-evidence manifest over real
temporary files, then exercises ``render_visual.attest`` and the validator.
"""
from __future__ import annotations

import hashlib
import importlib.util
import json
import subprocess
import sys
import tempfile
from pathlib import Path

sys.dont_write_bytecode = True
ROOT = Path(__file__).resolve().parents[1]


def load(name: str, relative: str):
    spec = importlib.util.spec_from_file_location(name, ROOT / relative)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def expect_value_error(function, fragment: str) -> None:
    try:
        function()
    except ValueError as exc:
        assert fragment in str(exc), (fragment, str(exc))
        return
    raise AssertionError(f"expected ValueError containing {fragment!r}")


def main() -> int:
    render = load("hx_render_visual", "scripts/office/render_visual.py")
    validator = load("hx_validator_visual", "scripts/validate_skill.py")
    with tempfile.TemporaryDirectory(prefix="hx-visual-regression-") as temporary:
        tmp = Path(temporary)
        deck, pdf = tmp / "deck.pptx", tmp / "deck.pdf"
        deck.write_bytes(b"pptx-bytes")
        pdf.write_bytes(b"pdf-bytes")
        pages = []
        for number in (1, 2):
            png = tmp / f"page-{number}.png"
            png.write_bytes(f"png-{number}".encode())
            pages.append({"page": number, "width_pt": 720, "height_pt": 405, "raster_width_px": 1080,
                          "raster_height_px": 608, "png": str(png), "png_sha256": sha(png)})
        evidence_path = tmp / "evidence.json"
        rendered = {"schema_version": 1, "status": "RENDERED", "render_status": "PASS",
                    "renderer": {"name": "LibreOffice"}, "manual_review": {"status": "PENDING", "pages": []},
                    "decks": [{"pptx": str(deck), "pptx_sha256": sha(deck), "pdf": str(pdf), "pdf_sha256": sha(pdf),
                               "page_count": 2, "pages": pages}]}
        evidence_path.write_text(json.dumps(rendered), encoding="utf-8")

        # A rendered-but-unreviewed manifest is not visual approval.
        _, errors, _ = validator.validate_visual_evidence(evidence_path, deck)
        assert any("manual_review.status must be PASS" in e for e in errors), errors
        # attest refuses missing notes, partial review and changed files.
        expect_value_error(lambda: render.attest(evidence_path, "all", "", None), "--review-notes is required")
        expect_value_error(lambda: render.attest(evidence_path, "1", "看过第一页", None), "every page must be reviewed")
        (tmp / "page-2.png").write_bytes(b"edited")
        expect_value_error(lambda: render.attest(evidence_path, "all", "逐页检查", None), "changed after rendering")
        (tmp / "page-2.png").write_bytes(b"png-2")
        # Success path: PASS evidence that the validator accepts for this deck.
        render.attest(evidence_path, "1-2", "两页均无溢出与遮挡", None)
        attested = json.loads(evidence_path.read_text(encoding="utf-8"))
        assert attested["status"] == "PASS" and attested["manual_review"]["pages"] == [1, 2], attested
        _, errors, matched = validator.validate_visual_evidence(evidence_path, deck)
        assert not errors and matched is not None, errors
        # Changing the deck afterwards invalidates the evidence.
        deck.write_bytes(b"pptx-bytes-v2")
        _, errors, _ = validator.validate_visual_evidence(evidence_path, deck)
        assert any("SHA-256 does not match" in e for e in errors), errors

        # Claiming a review in the same command as rendering is rejected by the CLI.
        cli = subprocess.run([sys.executable, "-B", str(ROOT / "scripts/office/render_visual.py"),
                              "--pptx", str(deck), "--output-dir", str(tmp / "out"), "--review-status", "pass"],
                             capture_output=True, text=True)
        assert cli.returncode != 0 and "--attest" in cli.stderr, cli.stderr

    print("Visual evidence regression checks PASSED")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
