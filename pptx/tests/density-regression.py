#!/usr/bin/env python3
"""Regression fixture for semantic PPTX density and typography limits."""
from __future__ import annotations

import importlib.util
import sys
import tempfile
from pathlib import Path

sys.dont_write_bytecode = True


ROOT = Path(__file__).resolve().parents[1]


def load_validator():
    path = ROOT / "scripts" / "validate_skill.py"
    spec = importlib.util.spec_from_file_location("pptx_validate_density", path)
    if spec is None or spec.loader is None:
        raise AssertionError(f"Could not load validator: {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def make_fixture(path: Path, over_limit: bool) -> None:
    from pptx import Presentation
    from pptx.chart.data import ChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.util import Inches, Pt
    from pptx.oxml.xmlchemy import OxmlElement

    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    def add_text(x: float, y: float, w: float, h: float, value: str, size: float,
                 name: str | None = None) -> None:
        textbox = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        if name:
            textbox.name = name
        textbox.text_frame.clear()
        paragraph = textbox.text_frame.paragraphs[0]
        run = paragraph.add_run()
        run.text = value
        run.font.size = Pt(size)

    add_text(0.5, 0.2, 9, 0.4, "正文标题", 22.5)
    add_text(0.5, 0.62, 9, 0.25, "结论标签", 13.5)
    add_text(0.5, 1.0, 9, 0.35, "关键数字", 18)
    card_count = 5 if over_limit else 4
    card_text = "卡片正文" * (12 if over_limit else 4)
    for index in range(card_count):
        card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(0.5 + (index % 4) * 2.3),
            Inches(1.5 + (index // 4) * 1.25),
            Inches(2.0),
            Inches(0.9),
        )
        card.name = f"HX_CARD_{index + 1}"
        card.text_frame.text = card_text
        for paragraph in card.text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(11.25)

    bullet_box = slide.shapes.add_textbox(Inches(0.5), Inches(4.1), Inches(4), Inches(1))
    bullet_box.name = "HX_BULLET_LIST"
    bullet_box.text_frame.clear()
    bullet_count = 6 if over_limit else 5
    for index in range(bullet_count):
        paragraph = bullet_box.text_frame.paragraphs[0] if index == 0 else bullet_box.text_frame.add_paragraph()
        paragraph.text = f"要点 {index + 1}"
        ppr = paragraph._p.get_or_add_pPr()
        bullet = OxmlElement("a:buChar")
        bullet.set("char", "•")
        ppr.append(bullet)
        for run in paragraph.runs:
            run.font.size = Pt(11.25)

    chart_data = ChartData()
    chart_data.categories = ["A", "B", "C", "D", "E"]
    chart_data.add_series("占比", (20, 20, 20, 20, 20))
    slide.shapes.add_chart(
        XL_CHART_TYPE.DOUGHNUT,
        Inches(5.0), Inches(3.2), Inches(4.4), Inches(2.0), chart_data
    )
    if over_limit:
        add_text(5.0, 1.0, 4, 0.3, "额外层级", 12.75)
    supporting_count = 3 if over_limit else 2
    for index in range(supporting_count):
        add_text(
            5.0, 1.0 + index * 0.4, 4.0, 0.3,
            f"辅助结论 {index + 1}", 18,
            name=f"HX_SUPPORTING_CONCLUSION_{index + 1}",
        )
    prs.save(path)


def make_component_contract_fixture(path: Path) -> None:
    """Create deliberately ambiguous components for object-level checks."""
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.util import Inches

    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rounded = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(0.84), Inches(4), Inches(1.2)
    )
    # Leave the generated name untouched: it must still be detected as an
    # unnamed text-bearing rounded component.
    rounded.text_frame.text = "未命名卡片"
    rectangle = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(5), Inches(0.84), Inches(4), Inches(1.2)
    )
    rectangle.name = "HX_CARD_BAD_GEOMETRY"
    rectangle.text_frame.text = "错误卡片"
    prs.save(path)


def main() -> int:
    validator = load_validator()
    density = {
        "max_top_level_bullets": 5,
        "max_major_cards": 4,
        "max_cjk_chars_per_card": 45,
        "max_major_charts": 1,
        "max_supporting_conclusions": 2,
        "max_typography_levels": 4,
        "body_min_runtime_pt": 11,
        "chart_min_runtime_pt": 10,
        "max_donut_categories": 5,
    }
    with tempfile.TemporaryDirectory(prefix="pptx-density-regression-") as temporary:
        valid_path = Path(temporary) / "valid.pptx"
        make_fixture(valid_path, over_limit=False)
        valid_stats, valid_warnings = validator.inspect_pptx(valid_path, check_bounds=True)
        assert valid_stats["bullet_paragraphs"] == 5, valid_stats
        assert valid_stats["named_card_count"] == 4, valid_stats
        assert valid_stats["max_cjk_chars_per_named_card"] <= 45, valid_stats
        assert valid_stats["charts"] == 1, valid_stats
        assert valid_stats["max_donut_categories"] == 5, valid_stats
        assert valid_stats["max_typography_levels_per_slide"] <= 4, valid_stats
        assert valid_stats["max_supporting_conclusions_per_slide"] == 2, valid_stats
        assert not validator.validate_pptx_density(valid_stats, density), valid_warnings

        invalid_path = Path(temporary) / "invalid.pptx"
        make_fixture(invalid_path, over_limit=True)
        invalid_stats, _ = validator.inspect_pptx(invalid_path, check_bounds=True)
        diagnostics = validator.validate_pptx_density(invalid_stats, density)
        assert any("bullet paragraphs" in item for item in diagnostics), diagnostics
        assert any("named card/panel" in item for item in diagnostics), diagnostics
        assert any("CJK characters" in item for item in diagnostics), diagnostics
        assert any("typography levels" in item for item in diagnostics), diagnostics

        component_path = Path(temporary) / "component-contract.pptx"
        make_component_contract_fixture(component_path)
        component_stats, _ = validator.inspect_pptx(component_path, check_bounds=True)
        assert component_stats["rounded_shape_count"] == 1, component_stats
        assert component_stats["unnamed_text_rounded_shape_count"] == 1, component_stats
        assert "HX_CARD_BAD_GEOMETRY" in component_stats["card_shape_violations"], component_stats
        component_diagnostics = validator.validate_pptx_density(component_stats, density)
        assert any("text-bearing rounded component" in item for item in component_diagnostics), component_diagnostics
        assert any("rounded rectangle geometry" in item for item in component_diagnostics), component_diagnostics
        assert any("supporting conclusion" in item for item in diagnostics), diagnostics

    print("Density regression checks PASSED")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
