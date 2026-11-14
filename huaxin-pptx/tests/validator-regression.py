#!/usr/bin/env python3
"""Regression checks for theme validation: JSON Schema plus cross-field consistency."""
from __future__ import annotations

import copy
import importlib.util
import json
import shutil
import sys
import tempfile
from pathlib import Path

sys.dont_write_bytecode = True
ROOT = Path(__file__).resolve().parents[1]


def load_validator():
    path = ROOT / "scripts" / "validate_skill.py"
    spec = importlib.util.spec_from_file_location("pptx_validate_skill", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def main() -> int:
    validator = load_validator()
    with tempfile.TemporaryDirectory(prefix="pptx-validator-regression-") as temporary:
        fixture = Path(temporary)
        (fixture / "scripts" / "office").mkdir(parents=True)
        (fixture / "references").mkdir()
        shutil.copytree(ROOT / "assets", fixture / "assets")
        shutil.copy2(ROOT / "references/theme-config.schema.json", fixture / "references/theme-config.schema.json")
        config_path = fixture / "scripts/office/huaxin_theme_config.json"
        original = json.loads((ROOT / "scripts/office/huaxin_theme_config.json").read_text(encoding="utf-8"))

        def errors_for(mutate) -> list[str]:
            config = copy.deepcopy(original)
            mutate(config)
            config_path.write_text(json.dumps(config, ensure_ascii=False), encoding="utf-8")
            return validator.validate_theme_config(fixture)[1]

        cases = [
            # Nested type errors are caught by the recursive JSON Schema.
            (lambda c: c["grid"].update(gap="wide"), "JSON Schema: $.grid.gap"),
            # Spacing tokens must come from the rhythm scale.
            (lambda c: c["grid"].update(gap=0.2), "grid.gap (0.2) must be one of spacing.scale"),
            # Unknown keys are rejected (no silent dead configuration).
            (lambda c: c["components"].update(shadow={"opacity": 0.1}), "Additional properties"),
            # Geometry outside the canvas.
            (lambda c: c["slides"]["body"]["content_area"].update(x=2.0), "content_area must bind"),
            (lambda c: c["slides"]["cover"]["metadata"].update(x=7.0), "slides.cover.metadata exceeds the canvas"),
            # Content area is bound to the named safe-area token, not a duplicated literal.
            (lambda c: c["safe_area"].update(top=0.95), "content_area.y must equal safe_area.top"),
            (lambda c: (c["safe_area"].update(top=0.6), c["slides"]["body"]["content_area"].update(y=0.6)),
             "title_to_content"),
            # Footer line follows the clearance token; content must keep its distance from the footer.
            (lambda c: c["safe_area"].update(footer_top_clearance=0.24), "footer.y must equal"),
            (lambda c: c["slides"]["body"]["content_area"].update(h=4.4), "collides with the footer"),
            # Sizes must respect floors; radius and table row-height tokens must be ordered.
            (lambda c: c["typography"]["sizes_pt"].update(body=10.5), "sizes_pt.body (10.5) is below floor_pt.body"),
            (lambda c: c["components"]["corner_radius"].update(default_in=0.2), "min_in <= default_in <= max_in"),
            (lambda c: c["recipes"]["table"].update(row_h_min=0.5), "row_h_min must not exceed row_h_max"),
            # Required palette tokens and non-empty template markers are enforced by Schema.
            (lambda c: c["colors"].pop("light_gray"), "'light_gray' is a required property"),
            (lambda c: c["qa"].update(template_markers=[]), "$.qa.template_markers"),
            # Colour references must point at defined tokens.
            (lambda c: c["charts"]["series_colors"].append("orange"), "unknown colour token 'orange'"),
            # Backgrounds must exist and use the canonical names.
            (lambda c: c["slides"]["closing"].update(bg_image="assets/body-background.jpeg"), "canonical"),
        ]
        for mutate, expected in cases:
            errors = errors_for(mutate)
            assert any(expected in error for error in errors), (expected, errors)
        assert not errors_for(lambda c: None), errors_for(lambda c: None)

    print("Validator regression checks PASSED")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
