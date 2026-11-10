#!/usr/bin/env python3
"""Regression checks for recursive theme validation and geometry guards."""
from __future__ import annotations

import importlib.util
import json
import shutil
import sys
import tempfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def load_validator():
    # This test imports the validator as a module. Disable bytecode writes
    # before loading it so the canonical skill remains unchanged afterward.
    sys.dont_write_bytecode = True
    path = ROOT / "scripts" / "validate_skill.py"
    spec = importlib.util.spec_from_file_location("pptx_validate_skill", path)
    if spec is None or spec.loader is None:
        raise AssertionError(f"Could not load validator: {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def copy_theme_fixture(destination: Path) -> None:
    (destination / "scripts" / "office").mkdir(parents=True)
    (destination / "references").mkdir()
    (destination / "assets").mkdir()
    shutil.copy2(ROOT / "scripts/office/huaxin_theme_config.json", destination / "scripts/office/huaxin_theme_config.json")
    shutil.copy2(ROOT / "references/theme-config.schema.json", destination / "references/theme-config.schema.json")
    for filename in ("cover-background.jpeg", "body-background.jpeg"):
        shutil.copy2(ROOT / "assets" / filename, destination / "assets" / filename)


def main() -> int:
    validator = load_validator()
    with tempfile.TemporaryDirectory(prefix="pptx-validator-regression-") as temporary:
        fixture = Path(temporary)
        copy_theme_fixture(fixture)
        config_path = fixture / "scripts/office/huaxin_theme_config.json"
        config = json.loads(config_path.read_text(encoding="utf-8"))

        # A nested type error must be caught by JSON Schema, rather than only
        # by the validator's hand-written top-level checks.
        config["grid"]["gap"] = "not-a-number"
        config_path.write_text(json.dumps(config), encoding="utf-8")
        _, errors = validator.validate_theme_config(fixture)
        assert any("JSON Schema" in error and "grid.gap" in error for error in errors), errors

        # Restore valid config, then move a content rectangle off canvas. This
        # is intentionally a cross-field check outside JSON Schema's scope.
        config["grid"]["gap"] = 0.22
        config["slides"]["content_slide"]["content_area"]["x"] = 9.5
        config_path.write_text(json.dumps(config), encoding="utf-8")
        _, errors = validator.validate_theme_config(fixture)
        assert any("content_slide.content_area exceeds" in error for error in errors), errors

        # Restore the horizontal position before checking cross-field token
        # bindings below.
        config["slides"]["content_slide"]["content_area"]["x"] = 0.5
        # The content rectangle is bound to the named top safe-area token,
        # rather than a duplicated literal.  Changing the token alone must
        # fail until the rectangle is moved with it.
        config["safe_area"]["top"] = 0.95
        config_path.write_text(json.dumps(config), encoding="utf-8")
        _, errors = validator.validate_theme_config(fixture)
        assert any("content_area.y must equal safe_area.top" in error for error in errors), errors
        config["safe_area"]["top"] = 0.84

        # All geometry tokens in this runtime theme must remain in runtime
        # coordinates; design coordinates belong only to the design canvas.
        config["spacing"]["coordinate_space"] = "design"
        config_path.write_text(json.dumps(config), encoding="utf-8")
        _, errors = validator.validate_theme_config(fixture)
        assert any("spacing.coordinate_space must be 'runtime'" in error for error in errors), errors
        config["spacing"]["coordinate_space"] = "runtime"

        # The card token must use the shared default radius, avoiding a second
        # per-component radius that silently diverges from the theme.
        config["components"]["card"]["radius_in"] = 0.104
        config_path.write_text(json.dumps(config), encoding="utf-8")
        _, errors = validator.validate_theme_config(fixture)
        assert any("card radius_in must equal components.corner_radius.default_in" in error for error in errors), errors
        config["components"]["card"]["radius_in"] = config["components"]["corner_radius"]["default_in"]

        # Footer clearance is a named semantic token.  The compatibility
        # ``bottom`` value must stay equal to it and footer.y must sit on the
        # corresponding top clearance line.
        config["safe_area"]["footer_top_clearance"] = 0.24
        config_path.write_text(json.dumps(config), encoding="utf-8")
        _, errors = validator.validate_theme_config(fixture)
        assert any("bottom must equal safe_area.footer_top_clearance" in error for error in errors), errors
        assert any("footer.y must equal canvas.height" in error for error in errors), errors

        # Restore valid semantics before leaving the fixture.
        config["safe_area"]["footer_top_clearance"] = config["safe_area"]["bottom"]
        config_path.write_text(json.dumps(config), encoding="utf-8")
        _, errors = validator.validate_theme_config(fixture)
        assert not errors, errors

    print("Validator regression checks PASSED")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
