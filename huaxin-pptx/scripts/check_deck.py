#!/usr/bin/env python3
"""Check a finished Huaxin deck (structure, density, typography floors, corner radii,
leftover template text). Lightweight: does not re-validate the skill package itself.

  python scripts/check_deck.py ../output/deck.pptx [--visual-evidence evidence.json] [--json report.json]

Exit code 0 = no errors (warnings may still need fixing), 1 = errors, 2 = missing dependency.
This is NOT a visual check; render and inspect every page as SKILL.md requires.
"""
from __future__ import annotations

import argparse
import importlib.util
import json
import sys
from pathlib import Path

sys.dont_write_bytecode = True
ROOT = Path(__file__).resolve().parents[1]


def load_validator():
    spec = importlib.util.spec_from_file_location("hx_validate_skill", ROOT / "scripts" / "validate_skill.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("pptx", type=Path)
    parser.add_argument("--visual-evidence", type=Path, help="Evidence JSON from render_visual.py (after --attest)")
    parser.add_argument("--json", dest="json_output", type=Path, help="Also write the report here")
    args = parser.parse_args()
    validator = load_validator()
    target = args.pptx.expanduser().resolve()
    theme = json.loads((ROOT / "scripts/office/huaxin_theme_config.json").read_text(encoding="utf-8"))
    report: dict = {"pptx": str(target), "status": "FAIL", "errors": [], "warnings": []}
    try:
        stats, warnings, errors = validator.inspect_deck(target, theme)
    except Exception as exc:  # noqa: BLE001 - report any parse failure as an error
        report["errors"].append(f"Could not inspect deck: {exc}")
        print(json.dumps(report, ensure_ascii=False, indent=2))
        return 1
    if args.visual_evidence:
        _, visual_errors, matched = validator.validate_visual_evidence(args.visual_evidence.expanduser().resolve(),
                                                                       target)
        errors.extend(visual_errors)
        if matched is not None and not visual_errors:
            stats["visual_check"] = "LIBREOFFICE_PASS"
    keys = ("slides", "width_in", "height_in", "visual_check", "named_cards_per_slide", "bullet_paragraphs_per_slide",
            "charts_per_slide", "typography_levels_per_slide", "min_explicit_font_size_pt", "min_chart_font_size_pt")
    report["summary"] = {key: stats.get(key) for key in keys}
    report["errors"] = errors
    report["warnings"] = sorted(set(warnings))
    report["status"] = "FAIL" if errors else ("PASS_WITH_WARNINGS" if report["warnings"] else "PASS")
    report["note"] = "Structural check only; it does not prove text fit, visual quality or factual accuracy."
    output = json.dumps(report, ensure_ascii=False, indent=2)
    if args.json_output:
        destination = args.json_output.expanduser().resolve()
        if destination == ROOT or ROOT in destination.parents:
            parser.error("--json must be outside the skill directory")
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_text(output + "\n", encoding="utf-8")
    print(output)
    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
