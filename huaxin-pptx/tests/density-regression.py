#!/usr/bin/env python3
"""Regression fixture for per-slide density, typography floors and corner-radius checks."""
from __future__ import annotations

import importlib.util
import sys
import tempfile
from pathlib import Path

sys.dont_write_bytecode = True


ROOT = Path(__file__).resolve().parents[1]
RADIUS = {"min_in": 0.104, "default_in": 0.125, "max_in": 0.167}
DENSITY = {
    "max_top_level_bullets": 5,
    "max_major_cards": 4,
    "max_cjk_chars_per_card": 45,
    "max_major_charts": 1,
    "max_supporting_conclusions": 2,
    "max_typography_levels": 4,
    "body_min_runtime_pt": 11,
    "chart_min_runtime_pt": 10,
    "caption_min_runtime_pt": 8,
    "max_donut_categories": 5,
}


def load_validator():
    path = ROOT / "scripts" / "validate_skill.py"
    spec = importlib.util.spec_from_file_location("pptx_validate_density", path)
    if spec is None or spec.loader is None:
        raise AssertionError(f"Could not load validator: {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _prs():
    from pptx import Presentation
    from pptx.util import Inches

    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)
    return prs


def _card(slide, index, x, y, text, radius=0.125, w=2.0, h=0.9):
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.util import Inches, Pt

    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    card.name = f"HX_CARD_{index}"
    card.adjustments[0] = radius / min(w, h)
    card.text_frame.text = "卡片标题"
    paragraph = card.text_frame.add_paragraph()
    paragraph.text = text
    for p in card.text_frame.paragraphs:
        for run in p.runs:
            run.font.size = Pt(11.25)
    return card


def make_fixture(path: Path, over_limit: bool) -> None:
    from pptx.chart.data import ChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.oxml.xmlchemy import OxmlElement
    from pptx.util import Inches, Pt

    prs = _prs()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    def add_text(x, y, w, h, value, size, name=None):
        textbox = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        if name:
            textbox.name = name
        run = textbox.text_frame.paragraphs[0].add_run()
        run.text = value
        run.font.size = Pt(size)

    add_text(0.5, 0.2, 9, 0.4, "正文标题", 22.5, name="HX_BODY_TITLE")
    add_text(0.5, 0.62, 9, 0.25, "结论标签", 13.5)
    add_text(0.5, 1.0, 9, 0.35, "关键数字", 18)
    card_count = 5 if over_limit else 4
    body = "卡片正文" * (12 if over_limit else 4)
    for index in range(card_count):
        _card(slide, index + 1, 0.5 + (index % 4) * 2.3, 1.5 + (index // 4) * 1.25, body)

    bullet_box = slide.shapes.add_textbox(Inches(0.5), Inches(4.1), Inches(4), Inches(1))
    bullet_box.name = "HX_BULLET_LIST"
    bullet_count = 6 if over_limit else 5
    for index in range(bullet_count):
        paragraph = bullet_box.text_frame.paragraphs[0] if index == 0 else bullet_box.text_frame.add_paragraph()
        paragraph.text = f"要点 {index + 1}"
        paragraph._p.get_or_add_pPr().append(OxmlElement("a:buChar"))
        paragraph._p.pPr[-1].set("char", "•")
        for run in paragraph.runs:
            run.font.size = Pt(11.25)

    chart_data = ChartData()
    chart_data.categories = ["A", "B", "C", "D", "E"]
    chart_data.add_series("占比", (20, 20, 20, 20, 20))
    slide.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, Inches(5.0), Inches(3.2), Inches(4.4), Inches(2.0), chart_data)
    if over_limit:
        add_text(5.0, 1.0, 4, 0.3, "额外层级", 12.75)
        add_text(5.0, 1.4, 4, 0.3, "过小正文", 9)
    for index in range(3 if over_limit else 2):
        add_text(5.0, 1.0 + index * 0.4, 4.0, 0.3, f"辅助结论 {index + 1}", 18,
                 name=f"HX_SUPPORTING_CONCLUSION_{index + 1}")
    add_text(0.5, 5.34, 9, 0.18, "来源：示例", 8, name="HX_SOURCE")
    prs.save(path)


def make_multi_slide_deck(path: Path) -> None:
    """Three slides with 3 cards each: deck totals exceed per-slide limits but no slide does."""
    prs = _prs()
    for _ in range(3):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        for index in range(3):
            _card(slide, index + 1, 0.5 + index * 3.0, 1.0, "正文")
    prs.save(path)


def make_component_contract_fixture(path: Path) -> None:
    """Deliberately ambiguous components and bad radii for object-level checks."""
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.util import Inches

    prs = _prs()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rounded = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(0.84), Inches(4), Inches(1.2))
    rounded.text_frame.text = "未命名卡片"  # default python-pptx name and default radius
    rectangle = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(5), Inches(0.84), Inches(4), Inches(1.2))
    rectangle.name = "HX_CARD_BAD_GEOMETRY"
    rectangle.text_frame.text = "错误卡片"
    _card(slide, 7, 0.5, 2.5, "默认圆角", radius=0.25, w=4, h=1.5)   # python-pptx default ≈ 0.25 in
    _card(slide, 8, 5.0, 2.5, "合规圆角", radius=0.125, w=4, h=1.5)
    pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(4.5), Inches(0.6), Inches(0.28))
    pill.name = "HX_PILL_1"
    pill.adjustments[0] = 0.5
    pill.text_frame.text = "01"
    step_bad = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.5), Inches(4.3), Inches(2.0), Inches(1.2))
    step_bad.name = "FLOW_STEP_BAD"  # semantic STEP without HX_ prefix (default ~0.20 in radius)
    step_bad.text_frame.text = "步骤"
    tag_bad = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.0), Inches(4.5), Inches(0.8), Inches(0.28))
    tag_bad.name = "STATUS_TAG_BAD"  # semantic TAG without HX_ prefix (wrong pill radius)
    tag_bad.adjustments[0] = 0.1
    tag_bad.text_frame.text = "待定"
    prs.save(path)


def make_legacy_13_canvas_fixture(path: Path, bad: bool) -> None:
    """13.333 × 7.5 in legacy deck (canvas_scale ≈ 1.3333): radii and floors scale proportionally."""
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.util import Inches, Pt

    prs = Presentation()
    prs.slide_width = Inches(13.333333)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    scale = 4 / 3
    radius = 0.35 if bad else 0.125 * scale
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.67), Inches(1.12), Inches(4.0), Inches(2.4))
    card.name = "HX_CARD_1"
    card.adjustments[0] = radius / 2.4
    run = card.text_frame.paragraphs[0].add_run()
    run.text = "旧画布正文"
    run.font.size = Pt(12 if bad else 15)  # scaled 11 pt floor = 14.67 pt
    prs.save(path)


def main() -> int:
    validator = load_validator()
    with tempfile.TemporaryDirectory(prefix="pptx-density-regression-") as temporary:
        valid_path = Path(temporary) / "valid.pptx"
        make_fixture(valid_path, over_limit=False)
        valid_stats, _ = validator.inspect_pptx(valid_path, check_bounds=True, radius_limits=RADIUS)
        assert valid_stats["bullet_paragraphs_per_slide"] == [5], valid_stats
        assert valid_stats["named_cards_per_slide"] == [4], valid_stats
        assert valid_stats["max_cjk_chars_per_named_card"] <= 45, valid_stats
        assert valid_stats["charts_per_slide"] == [1], valid_stats
        assert valid_stats["max_donut_categories"] == 5, valid_stats
        assert valid_stats["max_typography_levels_per_slide"] <= 4, valid_stats
        assert valid_stats["max_supporting_conclusions_per_slide"] == 2, valid_stats
        assert not valid_stats["radius_violations"], valid_stats
        diagnostics = validator.validate_pptx_density(valid_stats, DENSITY)
        assert not diagnostics, diagnostics

        invalid_path = Path(temporary) / "invalid.pptx"
        make_fixture(invalid_path, over_limit=True)
        invalid_stats, _ = validator.inspect_pptx(invalid_path, check_bounds=True, density_limits=DENSITY,
                                                  radius_limits=RADIUS)
        diagnostics = validator.validate_pptx_density(invalid_stats, DENSITY)
        for expected in ("bullet paragraphs", "named card/panel", "CJK characters", "typography levels",
                         "supporting conclusion", "below the body floor"):
            assert any(expected in item for item in diagnostics), (expected, diagnostics)

        multi_path = Path(temporary) / "multi.pptx"
        make_multi_slide_deck(multi_path)
        multi_stats, _ = validator.inspect_pptx(multi_path, check_bounds=True, radius_limits=RADIUS)
        assert multi_stats["named_card_count"] == 9 and multi_stats["named_cards_per_slide"] == [3, 3, 3]
        assert not validator.validate_pptx_density(multi_stats, DENSITY), "per-slide limits must not use deck totals"

        component_path = Path(temporary) / "component-contract.pptx"
        make_component_contract_fixture(component_path)
        component_stats, _ = validator.inspect_pptx(component_path, check_bounds=True, radius_limits=RADIUS)
        assert component_stats["unnamed_text_rounded_shape_count"] == 1, component_stats
        assert "HX_CARD_BAD_GEOMETRY" in component_stats["card_shape_violations"], component_stats
        assert any("HX_CARD_7" in item for item in component_stats["radius_violations"]), component_stats
        assert not any("HX_CARD_8" in item for item in component_stats["radius_violations"]), component_stats
        assert not any("HX_PILL_1" in item for item in component_stats["radius_violations"]), component_stats
        assert any("FLOW_STEP_BAD" in item for item in component_stats["radius_violations"]), component_stats
        assert any("STATUS_TAG_BAD" in item for item in component_stats["radius_violations"]), component_stats
        assert component_stats["radius_inconsistencies"], component_stats
        component_diagnostics = validator.validate_pptx_density(component_stats, DENSITY)
        for expected in ("text-bearing rounded component", "rounded rectangle geometry", "Corner radius violation"):
            assert any(expected in item for item in component_diagnostics), (expected, component_diagnostics)

        legacy_ok = Path(temporary) / "legacy-ok.pptx"
        make_legacy_13_canvas_fixture(legacy_ok, bad=False)
        legacy_stats, _ = validator.inspect_pptx(legacy_ok, check_bounds=True, density_limits=DENSITY,
                                                 radius_limits=RADIUS)
        assert abs(legacy_stats["canvas_scale"] - 4 / 3) < 0.001, legacy_stats
        assert not validator.validate_pptx_density(legacy_stats, DENSITY), legacy_stats

        legacy_bad = Path(temporary) / "legacy-bad.pptx"
        make_legacy_13_canvas_fixture(legacy_bad, bad=True)
        bad_stats, _ = validator.inspect_pptx(legacy_bad, check_bounds=True, density_limits=DENSITY,
                                              radius_limits=RADIUS)
        bad_diag = validator.validate_pptx_density(bad_stats, DENSITY)
        assert any("HX_CARD_1" in item for item in bad_stats["radius_violations"]), bad_stats
        assert any("below the body floor" in item for item in bad_diag), bad_diag

    print("Density regression checks PASSED")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
