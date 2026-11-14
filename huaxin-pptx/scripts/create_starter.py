#!/usr/bin/env python3
"""Create an UNFINISHED, editable Huaxin PPTX starter from the bundled backgrounds.

FALLBACK ONLY. The preferred path is ``scripts/hx_pptx.py build`` (clones the original
template). Use this starter only when template cloning cannot be used, or for QA fixtures.
It shares every geometry, size and helper with ``hx_pptx`` (10 × 5.625 in canvas).
Dependencies: python-pptx and Pillow. No network calls or font files are required.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parent))

import hx_pptx as hx  # noqa: E402  (path set above)
from pptx import Presentation  # noqa: E402
from pptx.util import Inches  # noqa: E402


def build_starter(args) -> "Presentation":
    prs = Presentation()
    prs.slide_width = Inches(hx.THEME["canvas"]["width"])
    prs.slide_height = Inches(hx.THEME["canvas"]["height"])
    prs.core_properties.title = args.title.replace("\n", " ")
    prs.core_properties.subject = "未完成的华新可编辑起稿：正文内容需另行制作并质检" + (
        "；包含主题组件验收夹具" if args.component_fixture else "")
    prs.core_properties.author = args.presenter.strip()
    prs.core_properties.keywords = "huaxin-pptx; starter; draft"
    width, height = hx.THEME["canvas"]["width"], hx.THEME["canvas"]["height"]
    content = hx.CONTENT

    def make_slide(page: dict):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        hx.add_image_fill(slide, hx.ROOT / page["bg_image"], 0, 0, width, height,
                          name="HX_BACKGROUND_COVER" if page is not hx.BODY else "HX_BACKGROUND_BODY", rounded=False)
        hx.set_notes(slide, "此页由 huaxin-pptx 起稿脚本（回退路径）创建，尚未完成正文制作和最终视觉质检。"
                            f"正文内容区默认从 y={content['y']:.2f} 英寸开始。"
                     + ("正文页包含 --component-fixture 主题组件验收夹具；交付前请替换示例文案。"
                        if args.component_fixture else ""))
        return slide

    slide = make_slide(hx.COVER)
    hx.add_cover_text(slide, args.title, args.subtitle, hx.cover_label(args.presenter, args.date))
    for title in args.body_title or ["请将本页标题改为核心结论"]:
        slide = make_slide(hx.BODY)
        hx.set_body_title(slide, title)
        if args.component_fixture:
            for index, ((x, w), card_title) in enumerate(zip(hx.columns(2), ("关键结论", "行动建议")), 1):
                hx.add_card(slide, x, content["y"], w, min(1.52, content["h"] - 0.62), card_title,
                            "用主题间距与字号承载一条可编辑正文。", name=f"HX_CARD_{index}")
            hx.add_source(slide, "来源：主题组件夹具（替换为真实来源）")
    slide = make_slide(hx.CLOSING)
    hx.add_text(slide, *hx.CLOSING["title"].values(), hx.closing_paragraphs(args.closing),
                name="HX_CLOSING_TITLE", anchor="middle")
    return prs


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--output", required=True, type=Path, help="Output .pptx outside the skill folder")
    parser.add_argument("--title", required=True, help="Cover title; at most two lines")
    parser.add_argument("--subtitle", default="", help="Optional one-line cover subtitle")
    parser.add_argument("--presenter", default="", help="Omit to retain 汇报人")
    parser.add_argument("--date", default="", help="Omit to retain 汇报日期; never inferred from the current date")
    parser.add_argument("--body-title", action="append", help="Repeat for each blank body slide")
    parser.add_argument("--component-fixture", action="store_true",
                        help="Add two theme cards and a source line to body slides for visual/structural QA")
    parser.add_argument("--closing", default="感谢聆听", help="Closing title; at most two lines")
    parser.add_argument("--font", default=None, help="Installed CJK-capable sans-serif font (default: theme font)")
    parser.add_argument("--force", action="store_true", help="Overwrite an ordinary output file, never skill assets")
    args = parser.parse_args()
    try:
        with hx.use_font(args.font):
            prs = build_starter(args)
        output = hx.save(prs, args.output, args.force)
    except hx.SpecError as exc:
        parser.error(str(exc))
    print(json.dumps({
        "output": str(output), "slides": len(prs.slides), "status": "UNFINISHED_STARTER",
        "canvas_inches": [hx.THEME["canvas"]["width"], hx.THEME["canvas"]["height"]],
        "presenter_date_label": hx.cover_label(args.presenter, args.date),
        "notes": [
            "正文页仅有标题，内容仍需制作；不要把起稿直接当作完成稿交付。",
            "此为原背景新建方案（回退路径），不是原模板克隆；首选 scripts/hx_pptx.py build。",
            "已重新解析文件；尚未执行外部渲染、字体可用性和目视质检。",
        ],
    }, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
