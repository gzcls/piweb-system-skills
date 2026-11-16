#!/usr/bin/env python3
"""Regression checks for scripts/hx_pptx.py (template cloning + every recipe) and create_starter.py.

Builds decks in a temporary directory; never writes inside the skill package.
"""
from __future__ import annotations

import copy
import importlib.util
import subprocess
import sys
import tempfile
from pathlib import Path
from zipfile import ZipFile

sys.dont_write_bytecode = True
ROOT = Path(__file__).resolve().parents[1]


def load(name: str, relative: str):
    spec = importlib.util.spec_from_file_location(name, ROOT / relative)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module  # dataclasses resolve annotations through sys.modules
    spec.loader.exec_module(module)
    return module


def expect_spec_error(hx, spec: dict, fragment: str) -> None:
    try:
        hx.build_deck(spec)
    except hx.SpecError as exc:
        assert fragment in str(exc), (fragment, str(exc))
        return
    raise AssertionError(f"expected SpecError containing {fragment!r}")


def check_clean(validator, theme, path: Path, expected_slides: int) -> dict:
    with ZipFile(path) as archive:
        names = archive.namelist()
    assert len(names) == len(set(names)), "duplicate ZIP entries (foreign parts were cloned)"
    stats, warnings, errors = validator.inspect_deck(path, theme)
    assert not errors, errors
    assert not warnings, warnings
    assert stats["slides"] == expected_slides, stats["slides"]
    assert stats["width_in"] == 10.0 and stats["height_in"] == 5.625, stats
    assert not stats["radius_violations"] and not stats["radius_inconsistencies"], stats
    return stats


def recipe_spec(image: Path) -> dict:
    """Every recipe and both template-data paths that the example spec does not cover."""
    return {"complete": False, "font": "SimHei", "slides": [
        {"type": "content", "title": "重点项目进度一览", "layout": {
            "type": "table", "columns": ["项目", "部门", "进度"], "col_widths": [3, 2, 1], "emphasis_rows": [1],
            "rows": [["数据中台", "信息部", "85%"], ["财务共享", "财务部", "100%"]]}, "source": "示例"},
        {"type": "content", "title": "现状与目标对比", "layout": {
            "type": "two_column", "left": {"title": "现状", "bullets": ["系统分散", "人工巡检"]},
            "right": {"title": "目标", "bullets": ["统一平台", "智能运维"]}}},
        {"type": "content", "title": "试点基地现场情况", "layout": {
            "type": "image_text", "image": str(image), "title": "试点成效", "bullets": ["产能提升", "能耗下降"]}},
        {"type": "content", "title": "构成与趋势组合页", "layout": [
            {"type": "kpis", "items": [{"value": "52", "unit": "%", "label": "制造占比"},
                                       {"value": "18", "unit": "%", "label": "物流占比"}]},
            {"type": "chart", "kind": "line", "unit": "万吨", "categories": ["1月", "2月", "3月"],
             "series": [{"name": "一号线", "values": [1, 2, 3]}, {"name": "二号线", "values": [2, 2, 4]}]}]},
        {"type": "template", "source": 9, "title": "四类能力支撑业务转型", "number": "02", "delete": [12],
         "texts": {"13": "数据治理", "14": "统一口径", "15": "流程再造", "16": "流程线上化",
                   "17": "平台建设", "18": "共享平台", "19": "人才培养", "20": "复合人才"}},
        {"type": "template", "source": 22, "title": "收入与利润按季度增长", "source_note": "单位：万元",
         "charts": [{"index": 0, "chart_title": None, "categories": ["Q1", "Q2", "Q3", "Q4"],
                     "series": [{"name": "净利润", "values": [1, 4, 18, 22]},
                                {"name": "收入", "values": [9, 30, 90, 112]}]}],
         "tables": [{"index": 0, "rows": [[None, "收入", "净利润", "利润率"], ["一季度", "9", "1", "11%"],
                                          ["二季度", "30", "4", "13%"], ["三季度", "90", "18", "20%"],
                                          ["四季度", "112", "22", "20%"]]}]},
    ]}


def main() -> int:
    hx = load("hx_pptx_under_test", "scripts/hx_pptx.py")
    validator = load("hx_validator_under_test", "scripts/validate_skill.py")
    theme = hx.THEME
    from PIL import Image
    from pptx import Presentation

    with tempfile.TemporaryDirectory(prefix="hx-builder-regression-") as temporary:
        tmp = Path(temporary)
        # 1. Example spec: structural pages + kpis/bullets/cards/chart/process.
        output = tmp / "example.pptx"
        prs, report = hx.build_deck(copy.deepcopy(hx.EXAMPLE_SPEC))
        hx.save(prs, output)
        assert not report["warnings"], report["warnings"]
        check_clean(validator, theme, output, len(hx.EXAMPLE_SPEC["slides"]))
        deck = Presentation(str(output))
        text = lambda slide: " ".join(s.text_frame.text for s in slide.shapes if s.has_text_frame)  # noqa: E731
        assert "汇报人｜汇报日期" in text(deck.slides[0]) and "演讲人" not in text(deck.slides[0])
        assert "感谢聆听" in text(deck.slides[-1]) and "Thank you" not in text(deck.slides[-1])
        closing_box = next(s for s in deck.slides[-1].shapes if s.name == "HX_CLOSING_TITLE")
        assert abs(closing_box.width / 914400 - theme["slides"]["closing"]["title"]["w"]) < 0.01
        assert "单击此处" not in text(deck.slides[2]), "section page keeps the template bar title"
        assert not hx.audit_markers(deck), hx.audit_markers(deck)
        for slide in list(deck.slides)[3:-1]:
            title = next(s for s in slide.shapes if s.name == "HX_BODY_TITLE")
            assert abs(title.left / 914400 - theme["slides"]["body"]["title"]["x"]) < 0.01

        # 2. Remaining recipes, stacked blocks, template text/chart/table replacement and the font override.
        image = tmp / "photo.png"
        Image.new("RGB", (1600, 900), (55, 75, 155)).save(image)
        recipes = tmp / "recipes.pptx"
        prs, report = hx.build_deck(recipe_spec(image))
        hx.save(prs, recipes)
        assert not report["warnings"], report["warnings"]
        stats = check_clean(validator, theme, recipes, 6)
        assert stats["charts_per_slide"] == [0, 0, 0, 1, 0, 1], stats["charts_per_slide"]
        deck = Presentation(str(recipes))
        names = [{s.name for s in slide.shapes} for slide in deck.slides]
        assert "HX_TABLE_1" in names[0] and {"HX_PANEL_1", "HX_PANEL_2"} <= names[1]
        assert "HX_IMAGE_FRAME_1" in names[2] and {"HX_TILE_1", "HX_CHART_1"} <= names[3]
        assert "数据治理" in " ".join(s.text_frame.text for s in hx.iter_shapes(deck.slides[4].shapes) if s.has_text_frame)
        table = next(s for s in deck.slides[5].shapes if s.has_table).table
        assert table.cell(4, 1).text_frame.text == "112"
        chart = next(s for s in deck.slides[5].shapes if getattr(s, "has_chart", False) and s.has_chart).chart
        assert list(chart.plots[0].categories) == ["Q1", "Q2", "Q3", "Q4"] and not chart.has_title
        fonts = {e.get("typeface") for e in deck.slides[0].shapes[0]._element.iter() if e.tag.endswith("}ea")}
        assert fonts == {"SimHei"}, fonts
        assert hx._ACTIVE_FONT == theme["typography"]["font"], "use_font must restore the theme font"

        # 3. Fallback starter shares the helpers and passes the same deck checks.
        starter = tmp / "starter.pptx"
        subprocess.run([sys.executable, "-B", str(ROOT / "scripts/create_starter.py"), "--component-fixture",
                        "--output", str(starter), "--title", "测试稿", "--body-title", "测试标题"],
                       check=True, capture_output=True)
        check_clean(validator, theme, starter, 3)
        rejected = subprocess.run([sys.executable, "-B", str(ROOT / "scripts/create_starter.py"),
                                   "--output", str(ROOT / "x.pptx"), "--title", "测试"], capture_output=True)
        assert rejected.returncode != 0 and not (ROOT / "x.pptx").exists()

        # 5. Chart formatting, dual-chart helper & template 22/23 regressions.
        from lxml import etree
        from pptx.oxml.ns import qn

        chart_reg = tmp / "chart-regressions.pptx"
        prs_cr, rep_cr = hx.build_deck({
            "complete": False,
            "slides": [
                {
                    "type": "content",
                    "title": "负值柱状与单系列堆积边界",
                    "layout": [
                        {
                            "type": "chart",
                            "kind": "bar",
                            "height": 2.0,
                            "unit": "%",
                            "categories": ["项目A", "项目B", "项目C"],
                            "series": [{"name": "增速", "values": [28.6, -9.4, 14.25]}],
                            "highlight": 0,
                        },
                        {
                            "type": "cards",
                            "items": [{"title": "单系列堆积", "body": "验证居中白字与横向反转"},
                                      {"title": "负值不反相", "body": "验证负值不错误变白"}],
                        },
                    ],
                },
                {
                    "type": "content",
                    "title": "单系列堆积柱图居中标签",
                    "layout": {
                        "type": "chart",
                        "kind": "stacked_column",
                        "unit": "项",
                        "categories": ["Q1", "Q2", "Q3"],
                        "series": [{"name": "结项", "values": [6, 9, 14]}],
                        "highlight": 2,
                    },
                },
                {
                    "type": "content",
                    "title": "饼图小扇区外置与浅色深字对比",
                    "layout": {
                        "type": "chart",
                        "kind": "pie",
                        "unit": "百万元",
                        "categories": ["A", "B", "C", "D", "E"],
                        "series": [{"name": "占比", "values": [47.5, 24.8, 23.6, 2.3, 1.8]}],
                        "highlight": 1,
                    },
                },
                {
                    "type": "content",
                    "title": "单系列折线标签与高亮标记",
                    "layout": {
                        "type": "chart",
                        "kind": "line",
                        "unit": "万元",
                        "categories": ["Q1", "Q2", "Q3", "Q4"],
                        "series": [{"name": "均价", "values": [18.5, 21.2, 24.8, 29.6]}],
                        "highlight": 3,
                    },
                },
                {
                    "type": "template",
                    "source": 22,
                    "title": "模板二十二页标题与表头补齐",
                    "source_note": "单位：百万元｜合并口径",
                    "charts": [{
                        "index": 0,
                        "chart_title": "季度收益趋势",
                        "categories": ["Q1", "Q2", "Q3", "Q4"],
                        "series": [{"name": "利润", "values": [12, 19, 28, 36]},
                                   {"name": "收入", "values": [95, 130, 175, 210]}],
                    }],
                    "tables": [{
                        "index": 0,
                        "rows": [["季度", "收入", "利润", "利润率"],
                                 ["一季度", "95", "12", "12.6%"],
                                 ["二季度", "130", "19", "14.6%"],
                                 ["三季度", "175", "28", "16.0%"],
                                 ["四季度", "210", "36", "17.1%"]],
                    }],
                },
                {
                    "type": "template",
                    "source": 23,
                    "title": "模板二十三页双图全量替换",
                    "source_note": "说明：测试数据",
                    "charts": [
                        {
                            "index": 0,
                            "chart_title": "区域构成",
                            "categories": ["华东", "华南", "华北", "海外"],
                            "series": [{"name": "销售", "values": [12.6, 6.8, 4.2, 2.4]}],
                        },
                        {
                            "index": 1,
                            "chart_title": "客单均价",
                            "categories": ["Q1", "Q2", "Q3", "Q4"],
                            "series": [{"name": "均价", "values": [18.5, 21.2, 24.8, 29.6]}],
                            "highlight": 3,
                        },
                    ],
                },
            ],
        })
        dual_s = hx.add_body_slide(prs_cr, "双图并排对比辅助函数验证")
        cols2 = hx.columns(2)
        hx.add_chart_panel(dual_s, cols2[0][0], hx.CONTENT["y"], cols2[0][1], hx.CONTENT["h"],
                           {"kind": "doughnut", "unit": "%", "categories": ["A", "B", "C", "D", "E"],
                            "series": [{"name": "构成", "values": [34, 26, 18, 13, 9]}]}, index=1)
        hx.add_chart_panel(dual_s, cols2[1][0], hx.CONTENT["y"], cols2[1][1], hx.CONTENT["h"],
                           {"kind": "column", "unit": "%", "categories": ["战略", "重点", "长尾"],
                            "series": [{"name": "去年", "values": [78, 66, 36]},
                                       {"name": "今年", "values": [98, 86, 49]}]}, index=2)
        hx.save(prs_cr, chart_reg)
        assert not rep_cr["warnings"], rep_cr["warnings"]
        cr_stats = check_clean(validator, theme, chart_reg, 7)
        assert cr_stats["charts_per_slide"] == [1, 1, 1, 1, 1, 2, 2], cr_stats
        assert cr_stats["major_charts_per_slide"] == [1, 1, 1, 1, 1, 1, 1], cr_stats

        deck_cr = Presentation(str(chart_reg))
        assert not hx.audit_markers(deck_cr), hx.audit_markers(deck_cr)
        bar_xml = etree.tostring(next(s.chart for s in deck_cr.slides[0].shapes if getattr(s, "has_chart", False) and s.has_chart)._chartSpace, encoding="unicode")
        assert '<c:invertIfNegative val="0"/>' in bar_xml and '<c:tickLblPos val="low"/>' in bar_xml
        assert '<c:orientation val="maxMin"/>' in bar_xml
        sc_xml = etree.tostring(next(s.chart for s in deck_cr.slides[1].shapes if getattr(s, "has_chart", False) and s.has_chart)._chartSpace, encoding="unicode")
        assert '<c:dLblPos val="ctr"/>' in sc_xml and "outEnd" not in sc_xml
        pie_xml = etree.tostring(next(s.chart for s in deck_cr.slides[2].shapes if getattr(s, "has_chart", False) and s.has_chart)._chartSpace, encoding="unicode")
        assert 'val="outEnd"' in pie_xml and 'val="3E3A39"' in pie_xml and "firstSliceAng" in pie_xml
        line_xml = etree.tostring(next(s.chart for s in deck_cr.slides[3].shapes if getattr(s, "has_chart", False) and s.has_chart)._chartSpace, encoding="unicode")
        assert 'val="E63237"' in line_xml and '<c:max val="35.0"/>' in line_xml
        t22_slide = deck_cr.slides[4]
        t22_table_frame = next(s for s in t22_slide.shapes if s.has_table)
        assert t22_table_frame.table.cell(0, 0).text_frame.text == "季度"
        assert (t22_table_frame.left + t22_table_frame.width) / 914400 <= 9.52
        t22_src = next(s for s in t22_slide.shapes if s.name == "HX_SOURCE").text_frame.text
        assert t22_src.startswith("单位：") and "来源：单位" not in t22_src
        t23_charts = [s for s in hx.iter_shapes(deck_cr.slides[5].shapes) if hx._shape_kind(s) == "chart"]
        assert len(t23_charts) == 2 and t23_charts[0].left >= 0
        t23_pie_xml = etree.tostring(t23_charts[0].chart._chartSpace, encoding="unicode")
        assert "FABE00" not in t23_pie_xml.upper()

        # Unreplaced template 23 slot 1 ("Y 值") must be caught by audit_markers, and numeric x_values keeps XY_SCATTER.
        prs_unrep, _ = hx.build_deck({"complete": False, "slides": [{"type": "template", "source": 23, "title": "未替换右图"}]})
        assert any("Y 值" in m for m in hx.audit_markers(prs_unrep))
        prs_xy, _ = hx.build_deck({"complete": False, "slides": [{"type": "template", "source": 23, "title": "保留散点图",
                                   "charts": [{"index": 1, "x_values": [1.0, 2.5, 4.0],
                                               "series": [{"name": "实测", "values": [10, 20, 35]}]}]}]})
        assert not hx.audit_markers(prs_xy)
        expect_spec_error(hx, {"slides": [{"type": "content", "title": "越界高亮", "layout": {
            "type": "chart", "kind": "column", "unit": "%", "categories": ["A", "B"],
            "series": [{"name": "s", "values": [1, 2]}], "highlight": 5}}]}, "超出分类范围")
        _, hl_rep = hx.build_deck({"complete": False, "slides": [{"type": "content", "title": "多系列忽略单点高亮", "layout": {
            "type": "chart", "kind": "column", "unit": "%", "categories": ["A", "B"],
            "series": [{"name": "s1", "values": [1, 2]}, {"name": "s2", "values": [3, 4]}], "highlight": 1}}]})
        assert any("highlight" in w for w in hl_rep["warnings"]), hl_rep

    # 4. Hard rules raise SpecError instead of producing a crowded slide.
    content = lambda layout: {"slides": [{"type": "content", "title": "测试页", "layout": layout}]}  # noqa: E731
    long_title = "这是一个明显超过二十五个字的正文标题因此必须被拒绝并要求压缩"
    expect_spec_error(hx, {"slides": [{"type": "content", "title": long_title}]}, "放不下")
    expect_spec_error(hx, content({"type": "cards", "items": [{"title": f"卡{i}"} for i in range(5)]}), "2–4 项")
    expect_spec_error(hx, content({"type": "chart", "kind": "pie", "categories": list("ABCDEF"),
                                   "series": [{"name": "s", "values": [1] * 6}]}), "不超过 5 类")
    expect_spec_error(hx, content({"type": "table", "columns": ["a", "b"], "rows": [["1", "2"]] * 20}), "放不下")
    expect_spec_error(hx, content({"type": "bullets", "items": ["要点"] * 6}), "1–5 项")
    expect_spec_error(hx, content({"type": "two_column", "left": {"bullets": ["a"] * 3},
                                   "right": {"bullets": ["b"] * 3}}), "单页不超过")
    expect_spec_error(hx, content([{"type": "chart", "categories": ["a"], "series": [{"values": [1]}]},
                                   {"type": "bullets", "items": ["x"]}]), "必须指定 height")
    expect_spec_error(hx, {"slides": [{"type": "template", "source": 22}, {"type": "template", "source": 22}]},
                      "只能复用一次")
    expect_spec_error(hx, {"slides": [{"type": "toc", "items": ["只有一项"]}]}, "2–6 项")
    _, report = hx.build_deck({"slides": [{"type": "cover", "title": "测试"}]})
    assert any("closing" in item for item in report["warnings"]), report

    print("Builder regression checks PASSED")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
