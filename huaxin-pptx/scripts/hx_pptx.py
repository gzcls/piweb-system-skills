#!/usr/bin/env python3
"""Huaxin PPTX builder: clone pages from the bundled original template and fill
them with editable, theme-compliant recipe layouts.

This module is both a CLI and an importable library. All geometry is in inches
on the original 10 x 5.625 in template canvas, and every size, colour and
recipe constant comes from ``scripts/office/huaxin_theme_config.json``.

CLI
---
  python scripts/hx_pptx.py build --spec deck.json --output ../output/deck.pptx
  python scripts/hx_pptx.py slots --slide 22          # list replaceable slots of a template page
  python scripts/hx_pptx.py example-spec              # print a complete example spec

Library
-------
  import sys; sys.path.insert(0, r"<skill>/scripts"); import hx_pptx as hx
  prs, report = hx.build_deck(spec)            # template pages + recipe layouts
  slide = hx.add_body_slide(prs, "观点标题")     # extra blank body page
  x, w = hx.columns(3)[0]; hx.add_card(slide, x, 0.84, w, 1.6, "卡片标题", "正文", name="HX_CARD_1")
  hx.save(prs, "../output/deck.pptx")

Dependencies: python-pptx, Pillow (image layouts only). No network access.
"""
from __future__ import annotations

import argparse
import contextlib
import copy
import json
import math
import os
import re
import sys
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable

sys.dont_write_bytecode = True

try:
    from pptx import Presentation
    from pptx.chart.data import CategoryChartData
    from pptx.dml.color import RGBColor
    from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
    from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
    from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
    from pptx.oxml.ns import qn
    from pptx.util import Emu, Inches, Pt
except ImportError as exc:  # pragma: no cover - environment guard
    raise SystemExit(f"Missing dependency: {exc}. Install requirements.txt (python-pptx, Pillow).")

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE_PATH = ROOT / "assets" / "huaxin-template-16x9.pptx"
THEME_PATH = ROOT / "scripts" / "office" / "huaxin_theme_config.json"
EMU_PER_IN = 914400
BODY_LAYOUT_NAME = "标题和内容"

# Template page numbers (1-based) used by the structural slide types.
PAGE_COVER, PAGE_TOC, PAGE_SECTION, PAGE_CLOSING = 1, 2, 4, 30
ENGLISH_ORDINALS = ["ONE", "TWO", "THREE", "FOUR", "FIVE", "SIX", "SEVEN", "EIGHT", "NINE", "TEN"]


class SpecError(ValueError):
    """Raised when a spec violates a hard rule (density, length, missing data)."""


# --------------------------------------------------------------------------- theme
THEME = json.loads(THEME_PATH.read_text(encoding="utf-8"))
COLORS = THEME["colors"]
TYPE = THEME["typography"]["sizes_pt"]
FLOOR = THEME["typography"]["floor_pt"]
LINE_SPACING = float(THEME["typography"]["line_spacing"])
DENSITY = THEME["density"]
RADIUS = THEME["components"]["corner_radius"]
BORDER_PT = float(THEME["components"]["border_pt"])
COVER = THEME["slides"]["cover"]
BODY = THEME["slides"]["body"]
CLOSING = THEME["slides"]["closing"]
CONTENT = BODY["content_area"]
CONTENT_BOTTOM = CONTENT["y"] + CONTENT["h"]
FOOTER = THEME["footer"]
CHARTS = THEME["charts"]
R = THEME["recipes"]
GAP = float(THEME["grid"]["gap"])
PAD = float(THEME["spacing"]["card_padding"])
TEMPLATE_MARKERS = tuple(THEME["qa"]["template_markers"])

ALIGN = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}
ANCHOR = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE, "bottom": MSO_ANCHOR.BOTTOM}
_ACTIVE_FONT = THEME["typography"]["font"]


@contextlib.contextmanager
def use_font(name: str | None):
    """Temporarily override the font used by every text helper (spec ``font`` / starter ``--font``)."""
    global _ACTIVE_FONT
    previous = _ACTIVE_FONT
    _ACTIVE_FONT = name or previous
    try:
        yield
    finally:
        _ACTIVE_FONT = previous


def rgb(name_or_hex: str):
    value = COLORS.get(name_or_hex, name_or_hex)
    return RGBColor.from_string(value.lstrip("#").upper())


# --------------------------------------------------------------------------- geometry helpers
def columns(n: int, x: float | None = None, w: float | None = None, gap: float = GAP) -> list[tuple[float, float]]:
    """Return ``n`` equal (x, width) columns inside the content area."""
    x0 = CONTENT["x"] if x is None else x
    total = CONTENT["w"] if w is None else w
    width = (total - gap * (n - 1)) / n
    return [(round(x0 + i * (width + gap), 4), round(width, 4)) for i in range(n)]


def _units(text: str) -> float:
    return sum(1.0 if ord(ch) > 0x2E80 else 0.55 for ch in text)


def capacity(width_in: float, size_pt: float) -> int:
    """How many CJK characters of ``size_pt`` fit on one line of ``width_in``."""
    return int(width_in * 72 // size_pt)


def estimate_lines(text: str, width_in: float, size_pt: float) -> int:
    per_line = max(1.0, width_in / (size_pt / 72.0))
    return sum(max(1, math.ceil(_units(line) / per_line)) for line in str(text).split("\n"))


def line_height(size_pt: float) -> float:
    return size_pt / 72.0 * 1.2 * LINE_SPACING


def text_height(text: str, width_in: float, size_pt: float) -> float:
    return estimate_lines(text, width_in, size_pt) * line_height(size_pt) if text else 0.0


def cjk_count(text: str) -> int:
    return len(re.findall(r"[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]", text or ""))


def _rect(spec: dict) -> tuple:
    return Inches(spec["x"]), Inches(spec["y"]), Inches(spec["w"]), Inches(spec["h"])


# --------------------------------------------------------------------------- text helpers
def set_font_name(run, name: str) -> None:
    rpr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        element = rpr.find(qn(tag))
        if element is None:
            element = rpr.makeelement(qn(tag), {})
            rpr.append(element)
        element.set("typeface", name)


def _style_run(run, size=None, color=None, bold=None, set_font: bool = True) -> None:
    if size is not None:
        run.font.size = Pt(size)
    if color is not None:
        run.font.color.rgb = rgb(color)
    if bold is not None:
        run.font.bold = bold
    if set_font:
        set_font_name(run, _ACTIVE_FONT)


def _add_bullet(paragraph, char: str = "•") -> None:
    ppr = paragraph._p.get_or_add_pPr()
    ppr.set("marL", str(int(Inches(0.16))))
    ppr.set("indent", str(int(-Inches(0.16))))
    for tag in ("a:buNone", "a:buChar", "a:buAutoNum"):
        for old in ppr.findall(qn(tag)):
            ppr.remove(old)
    ppr.append(ppr.makeelement(qn("a:buChar"), {"char": char}))


def fill_text(text_frame, paragraphs: list[dict], *, anchor: str = "top", margins: tuple = (0, 0, 0, 0),
              wrap: bool = True) -> None:
    """Replace a text frame's content.

    ``paragraphs`` items: {"runs": [{"text", "size", "color", "bold"}], "align", "bullet", "space_after"}.
    A shorthand item {"text", "size", "color", "bold", ...} becomes a single run.
    """
    text_frame.clear()
    text_frame.word_wrap = wrap
    text_frame.auto_size = MSO_AUTO_SIZE.NONE
    text_frame.vertical_anchor = ANCHOR[anchor]
    left, top, right, bottom = margins
    text_frame.margin_left, text_frame.margin_top = Inches(left), Inches(top)
    text_frame.margin_right, text_frame.margin_bottom = Inches(right), Inches(bottom)
    for index, spec in enumerate(paragraphs):
        paragraph = text_frame.paragraphs[0] if index == 0 else text_frame.add_paragraph()
        paragraph.alignment = ALIGN[spec.get("align", "left")]
        paragraph.line_spacing = LINE_SPACING
        paragraph.space_before = Pt(0)
        paragraph.space_after = Pt(spec.get("space_after", 0))
        if spec.get("bullet"):
            _add_bullet(paragraph)
        runs = spec.get("runs") or [{k: spec[k] for k in ("text", "size", "color", "bold") if k in spec}]
        for run_spec in runs:
            run = paragraph.add_run()
            run.text = str(run_spec.get("text", ""))
            _style_run(run, run_spec.get("size", TYPE["body"]), run_spec.get("color", "text_dark"),
                       run_spec.get("bold", False))


def replace_text_keep_format(shape, text: str, *, size=None, color=None, bold=None, keep_font: bool = False) -> None:
    """Replace text while keeping the first run's formatting from the template.

    The theme font is applied unless ``keep_font`` is set (decorative template numerals).
    """
    text_frame = shape.text_frame
    first_rpr = first_ppr = None
    for paragraph in text_frame.paragraphs:
        if first_ppr is None and paragraph._p.pPr is not None:
            first_ppr = copy.deepcopy(paragraph._p.pPr)
        first_rpr = next((copy.deepcopy(r._r.rPr) for r in paragraph.runs if r.text.strip() and r._r.rPr is not None),
                         None)
        if first_rpr is not None:
            break
    text_frame.clear()
    for index, line in enumerate(str(text).split("\n")):
        paragraph = text_frame.paragraphs[0] if index == 0 else text_frame.add_paragraph()
        if index > 0 and first_ppr is not None:
            if paragraph._p.pPr is not None:
                paragraph._p.remove(paragraph._p.pPr)
            paragraph._p.insert(0, copy.deepcopy(first_ppr))
        run = paragraph.add_run()
        run.text = line
        if first_rpr is not None:
            if run._r.rPr is not None:
                run._r.remove(run._r.rPr)
            run._r.insert(0, copy.deepcopy(first_rpr))
        _style_run(run, size, color, bold, set_font=not keep_font)


# --------------------------------------------------------------------------- shape helpers
def _no_shadow(shape) -> None:
    """Remove theme effects. ``effectLst`` covers PowerPoint; effectRef idx=0 covers LibreOffice/WPS."""
    shape.shadow.inherit = False
    style = shape._element.find(qn("p:style"))
    effect_ref = style.find(qn("a:effectRef")) if style is not None else None
    if effect_ref is not None:
        effect_ref.set("idx", "0")


def set_corner_radius(shape, radius_in: float | None = None) -> float:
    """Set a rounded rectangle radius in inches (python-pptx needs a ratio of min(w, h))."""
    radius = float(RADIUS["default_in"]) if radius_in is None else float(radius_in)
    short_side = min(shape.width, shape.height) / EMU_PER_IN
    radius = min(radius, short_side / 2)
    shape.adjustments[0] = radius / short_side if short_side else 0
    return radius


def _apply_fill_line(shape, fill, line) -> None:
    _no_shadow(shape)
    if fill is None:
        shape.fill.background()
    else:
        shape.fill.solid()
        shape.fill.fore_color.rgb = rgb(fill)
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = rgb(line)
        shape.line.width = Pt(BORDER_PT)


def add_rect(slide, x, y, w, h, *, fill="light_blue_gray", line=None, name="HX_RECT"):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.name = name
    _apply_fill_line(shape, fill, line)
    return shape


def add_rounded(slide, x, y, w, h, *, fill="white", line="light_blue_gray", name="HX_CARD",
                radius: float | None = None):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.name = name
    set_corner_radius(shape, radius)
    _apply_fill_line(shape, fill, line)
    return shape


def add_pill(slide, x, y, w, h, text, *, fill="brand_blue", color="white", size=None, name="HX_PILL"):
    shape = add_rounded(slide, x, y, w, h, fill=fill, line=None, name=name,
                        radius=min(h / 2, float(RADIUS["max_in"])))
    fill_text(shape.text_frame, [{"text": text, "size": size or TYPE["body"], "color": color, "bold": True,
                                  "align": "center"}], anchor="middle", wrap=False)
    return shape


def add_text(slide, x, y, w, h, paragraphs: list[dict] | str, *, name="HX_TEXT", anchor="top",
             size=None, color="text_dark", bold=False, align="left", wrap=True):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    box.name = name
    if isinstance(paragraphs, str):
        paragraphs = [{"text": line, "size": size or TYPE["body"], "color": color, "bold": bold, "align": align}
                      for line in paragraphs.split("\n")]
    fill_text(box.text_frame, paragraphs, anchor=anchor, wrap=wrap)
    return box


def add_card(slide, x, y, w, h, title: str, body: str = "", *, name="HX_CARD", accent=False):
    """Rounded card with title paragraph + body paragraph inside the shape (editable)."""
    card = add_rounded(slide, x, y, w, h, fill="brand_blue" if accent else "white",
                       line=None if accent else "light_blue_gray", name=name)
    paragraphs = [{"text": title, "size": TYPE["card_title"], "color": "white" if accent else "brand_blue",
                   "bold": True, "space_after": 4}]
    paragraphs += [{"text": line, "size": TYPE["body"], "color": "white" if accent else "text_dark"}
                   for line in str(body).split("\n") if body]
    fill_text(card.text_frame, paragraphs, margins=(PAD, PAD, PAD, PAD))
    return card


def add_source(slide, text: str):
    if not text:
        return None
    label = text if re.match(r"^(来源|资料来源|数据来源|注|Source)", text) else f"来源：{text}"
    return add_text(slide, FOOTER["x"], FOOTER["y"], FOOTER["w"], FOOTER["h"],
                    [{"text": label, "size": TYPE["caption"], "color": FOOTER["color"]}],
                    name="HX_SOURCE", anchor="middle")


def add_image_fill(slide, path: str | Path, x, y, w, h, *, name="HX_IMAGE_FRAME_1", rounded=True):
    """Insert an image cropped to fill the frame. Returns (picture, low-resolution warning or None)."""
    from PIL import Image

    path = Path(path).expanduser().resolve()
    if not path.is_file():
        raise SpecError(f"图片不存在：{path}")
    with Image.open(path) as image:
        iw, ih = image.size
    picture = slide.shapes.add_picture(str(path), Inches(x), Inches(y), Inches(w), Inches(h))
    picture.name = name
    image_ratio, frame_ratio = iw / ih, w / h
    if image_ratio > frame_ratio:
        crop = (1 - frame_ratio / image_ratio) / 2
        picture.crop_left = picture.crop_right = crop
        visible_px = iw * (1 - 2 * crop)
    else:
        crop = (1 - image_ratio / frame_ratio) / 2
        picture.crop_top = picture.crop_bottom = crop
        visible_px = iw
    warning = None
    if visible_px / w < R["image_min_ppi"]:
        warning = f"图片 {path.name} 分辨率偏低（约 {visible_px / w:.0f} px/in），可能模糊"
    if rounded:
        geometry = picture._element.spPr.find(qn("a:prstGeom"))
        geometry.set("prst", "roundRect")
        av = geometry.find(qn("a:avLst"))
        if av is None:
            av = geometry.makeelement(qn("a:avLst"), {})
            geometry.append(av)
        for old in list(av):
            av.remove(old)
        ratio = min(0.5, float(RADIUS["default_in"]) / min(w, h))
        av.append(av.makeelement(qn("a:gd"), {"name": "adj", "fmla": f"val {int(round(ratio * 100000))}"}))
    return picture, warning


def set_notes(slide, notes: str | None) -> None:
    if notes:
        slide.notes_slide.notes_text_frame.text = str(notes)


# --------------------------------------------------------------------------- template cloning
def open_template():
    return Presentation(str(TEMPLATE_PATH))


def iter_shapes(shapes) -> Iterable:
    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_shapes(shape.shapes)


def _remove(shape) -> None:
    element = shape._element
    if element.getparent() is not None:
        element.getparent().remove(element)


def clone_slide(prs, source_slide):
    """Append a deep copy of ``source_slide`` from the SAME package (media and charts shared by reference)."""
    new = prs.slides.add_slide(source_slide.slide_layout)
    for shape in list(new.shapes):
        _remove(shape)
    rid_map: dict[str, str] = {}
    for rid, rel in source_slide.part.rels.items():
        if rel.reltype.endswith(("/slideLayout", "/notesSlide")):
            continue
        rid_map[rid] = (new.part.relate_to(rel.target_ref, rel.reltype, is_external=True) if rel.is_external
                        else new.part.relate_to(rel.target_part, rel.reltype))
    tree = new.shapes._spTree
    for child in source_slide.shapes._spTree.iterchildren():
        if child.tag not in (qn("p:nvGrpSpPr"), qn("p:grpSpPr")):
            tree.append(copy.deepcopy(child))
    r_ns = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    for element in tree.iter():
        for attr, value in list(element.attrib.items()):
            if attr.startswith(r_ns) and value in rid_map:
                element.set(attr, rid_map[value])
    src_bg = source_slide._element.cSld.find(qn("p:bg"))
    if src_bg is not None:
        new._element.cSld.insert(0, copy.deepcopy(src_bg))
    return new


def _has_chart(slide) -> bool:
    return any(rel.reltype.endswith("/chart") for rel in slide.part.rels.values())


def _drop_original_slides(prs, count: int) -> None:
    id_list = prs.slides._sldIdLst
    for sld_id in list(id_list)[:count]:
        prs.part.drop_rel(sld_id.rId)
        id_list.remove(sld_id)


def new_body_slide(prs):
    """Blank body page = template page 5 (layout 标题和内容, background from the master, no shapes)."""
    layout = next((l for l in prs.slide_layouts if l.name == BODY_LAYOUT_NAME), prs.slide_layouts[1])
    slide = prs.slides.add_slide(layout)
    for shape in list(slide.shapes):
        _remove(shape)
    return slide


# --------------------------------------------------------------------------- page chrome
def _find_text_shape(slide, *, max_top: float, left_range: tuple[float, float], max_w: float = 99,
                     max_h: float = 99):
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        top, left = shape.top / EMU_PER_IN, shape.left / EMU_PER_IN
        if (top < max_top and left_range[0] <= left < left_range[1]
                and shape.width / EMU_PER_IN < max_w and shape.height / EMU_PER_IN < max_h):
            return shape
    return None


def _bar_title_shape(slide):
    return _find_text_shape(slide, max_top=0.6, left_range=(1.3, 3.0), max_h=0.6)


def _tab_shape(slide):
    return _find_text_shape(slide, max_top=0.6, left_range=(0.0, 1.3), max_w=1.0)


def check_title(text: str, kind: str = "body") -> float:
    """Return the font size for a title, or raise when it cannot fit (capacity = box width / font size)."""
    lines = str(text).split("\n")
    if not text or not str(text).strip():
        raise SpecError("标题不能为空")
    longest = max(_units(line) for line in lines)
    box = BODY["title"] if kind == "body" else (CLOSING["title"] if kind == "closing" else COVER["title"])
    first = TYPE["slide_title"] if kind == "body" else TYPE["cover_title"]
    max_lines = 1 if kind == "body" else 2
    if len(lines) > max_lines:
        raise SpecError(f"{'正文标题必须单行' if kind == 'body' else '首页/尾页标题最多两行'}：{text!r}")
    for size in (first, TYPE["long_title"]):
        if longest <= capacity(box["w"], size):
            return size
    raise SpecError(f"标题每行超过 {capacity(box['w'], TYPE['long_title'])} 字，放不下，请压缩为观点句：{text!r}")


def check_length(text: str, limit: int, label: str) -> None:
    if _units(text) > limit:
        raise SpecError(f"{label} 不超过 {limit} 字：{text!r}")


def set_body_title(slide, text: str):
    size = check_title(text, "body")
    shape = _bar_title_shape(slide) or slide.shapes.add_textbox(*_rect(BODY["title"]))
    shape.name = "HX_BODY_TITLE"
    shape.left, shape.top, shape.width, shape.height = _rect(BODY["title"])
    fill_text(shape.text_frame, [{"text": text, "size": size, "color": "white", "bold": True}],
              anchor="middle", wrap=False)
    return shape


def set_section_tab(slide, number: str | None):
    shape = _tab_shape(slide)
    if not number:
        return shape
    shape = shape or slide.shapes.add_textbox(*_rect(BODY["section_tab"]))
    shape.name = "HX_SECTION_TAB"
    shape.left, shape.top, shape.width, shape.height = _rect(BODY["section_tab"])
    fill_text(shape.text_frame, [{"text": str(number), "size": TYPE["section_tab"], "color": "white", "bold": True,
                                  "align": "center"}], anchor="middle", wrap=False)
    return shape


def add_body_slide(prs, title: str, *, number: str | None = None):
    """Append a blank body page with a bar title. ``prs`` must be opened from the bundled template."""
    slide = new_body_slide(prs)
    set_body_title(slide, title)
    set_section_tab(slide, number)
    return slide


def cover_label(presenter: str | None, date: str | None) -> str:
    return f"{(presenter or '').strip() or '汇报人'}｜{(date or '').strip() or '汇报日期'}"


def add_cover_text(slide, title: str, subtitle: str = "", label: str = "") -> None:
    """Cover title / subtitle / bottom-right label at theme positions (shared with create_starter)."""
    size = check_title(title, "cover")
    add_text(slide, *COVER["title"].values(),
             [{"text": line, "size": size, "color": "white", "bold": True, "align": "center"}
              for line in title.split("\n")], name="HX_COVER_TITLE", anchor="middle")
    if subtitle:
        if "\n" in subtitle:
            raise SpecError("cover.subtitle 必须单行")
        check_length(subtitle, capacity(COVER["subtitle"]["w"], TYPE["cover_subtitle"]), "cover.subtitle")
        add_text(slide, *COVER["subtitle"].values(),
                 [{"text": subtitle, "size": TYPE["cover_subtitle"], "color": "white", "align": "center"}],
                 name="HX_COVER_SUBTITLE", anchor="middle")
    if label:
        add_text(slide, *COVER["metadata"].values(),
                 [{"text": label, "size": TYPE["label"], "color": "primary", "align": "right"}],
                 name="HX_COVER_METADATA", anchor="middle", wrap=False)


def closing_paragraphs(title: str, subtitle: str = "") -> list[dict]:
    size = check_title(title, "closing")
    paragraphs = [{"text": line, "size": size, "color": "white", "bold": True, "align": "center"}
                  for line in title.split("\n")]
    if subtitle:
        paragraphs.append({"text": subtitle, "size": TYPE["cover_subtitle"], "color": "white", "align": "center"})
    return paragraphs


# --------------------------------------------------------------------------- build context
@dataclass
class BuildContext:
    prs: Any
    warnings: list[str] = field(default_factory=list)
    used_chart_sources: set[int] = field(default_factory=set)

    def template_slide(self, page: int):
        """Original template page; valid during build because the originals stay at the front."""
        return self.prs.slides[page - 1]


@dataclass
class Block:
    """Where a recipe draws: slide, top y, warning sink, and whether it may stretch downwards."""
    slide: Any
    y: float
    warnings: list[str]
    stretch: bool = True

    @property
    def available(self) -> float:
        return CONTENT_BOTTOM - self.y


# --------------------------------------------------------------------------- structural pages
def build_cover(ctx: BuildContext, spec: dict):
    slide = clone_slide(ctx.prs, ctx.template_slide(PAGE_COVER))
    for shape in list(slide.shapes):
        # Template title + "演讲人：" / "时间：" boxes are replaced by theme-positioned text.
        if shape.has_text_frame:
            _remove(shape)
    add_cover_text(slide, spec.get("title") or "", spec.get("subtitle") or "",
                   cover_label(spec.get("presenter"), spec.get("date")))
    set_notes(slide, spec.get("notes"))
    return slide


def build_toc(ctx: BuildContext, spec: dict):
    items = [item if isinstance(item, dict) else {"title": item} for item in spec.get("items") or []]
    if not 2 <= len(items) <= 6:
        raise SpecError("toc.items 需要 2–6 项")
    slide = clone_slide(ctx.prs, ctx.template_slide(PAGE_TOC))
    item_w = R["toc"]["item_w"]
    texts = [s for s in slide.shapes if s.has_text_frame]
    titles = [s for s in texts if s.text_frame.text.strip() == "在这里添加目录标题"]
    subs = [s for s in texts if s.text_frame.text.strip().startswith("according")]
    for shape in texts:  # template page 2 carries a hidden stray "02" box outside the groups
        if re.fullmatch(r"0\d", shape.text_frame.text.strip()) and shape.left / EMU_PER_IN > 1.0:
            _remove(shape)
    groups = []
    for shape in slide.shapes:
        if shape.shape_type != MSO_SHAPE_TYPE.GROUP:
            continue
        numbers = [s for s in iter_shapes(shape.shapes)
                   if s.has_text_frame and re.fullmatch(r"0\d", s.text_frame.text.strip())]
        if numbers:
            groups.append((numbers[0].text_frame.text.strip(), shape))
            numbers[0].text_frame.word_wrap = False
            replace_text_keep_format(numbers[0], numbers[0].text_frame.text.strip(), size=TYPE["card_title"],
                                     keep_font=True)
    order = lambda s: (0 if s.left / EMU_PER_IN < 4.5 else 1, s.top)
    titles.sort(key=order)
    subs.sort(key=order)
    groups.sort(key=lambda item: item[0])
    for index in range(6):
        if index >= len(items):
            for shape in (titles[index], subs[index], groups[index][1]):
                _remove(shape)
            continue
        item = items[index]
        check_length(item["title"], capacity(item_w, TYPE["card_title"]), "目录标题")
        box = titles[index]
        box.width, box.height = Inches(item_w), Inches(0.3)
        fill_text(box.text_frame, [{"text": item["title"], "size": TYPE["card_title"], "color": "text_dark",
                                    "bold": True}], anchor="middle", wrap=False)
        desc = item.get("desc") or ""
        if not desc:
            _remove(subs[index])
            continue
        check_length(desc, capacity(item_w, TYPE["body"]) * 2, "目录说明")
        sub = subs[index]
        sub.top, sub.width, sub.height = box.top + Inches(0.3), Inches(item_w), Inches(0.42)
        fill_text(sub.text_frame, [{"text": desc, "size": TYPE["body"], "color": "primary"}])
    if len(items) <= 3:  # single column: centre it instead of leaving it on the left half
        shift = Inches(R["toc"]["center_shift"])
        for shape in titles[:len(items)] + subs[:len(items)] + [g for _, g in groups[:len(items)]]:
            if shape._element.getparent() is not None:
                shape.left = shape.left + shift
    if _tab_shape(slide) is not None:
        set_section_tab(slide, spec.get("label") or "目录")
    if spec.get("title"):
        set_body_title(slide, spec["title"])
    set_notes(slide, spec.get("notes"))
    return slide


def build_section(ctx: BuildContext, spec: dict):
    number = str(spec.get("number") or "01")
    title = spec.get("title") or ""
    lines = title.split("\n")
    limit = R["section_title_chars"]
    if not title or len(lines) > 2 or any(_units(line) > limit for line in lines):
        raise SpecError(f"section.title 需 1–2 行、每行不超过 {limit} 字：{title!r}")
    slide = clone_slide(ctx.prs, ctx.template_slide(PAGE_SECTION))
    bar_title = _bar_title_shape(slide)
    if bar_title is not None:
        _remove(bar_title)
    set_section_tab(slide, number)
    index = int(number) - 1 if number.isdigit() else -1
    part = f"PART {ENGLISH_ORDINALS[index]}" if 0 <= index < len(ENGLISH_ORDINALS) else f"PART {number}"
    for shape in iter_shapes(slide.shapes):
        if not shape.has_text_frame:
            continue
        text = shape.text_frame.text.strip()
        if text in ("01", "PART ONE"):
            shape.text_frame.word_wrap = False
            replace_text_keep_format(shape, number if text == "01" else (spec.get("label") or part), keep_font=True)
        elif text.startswith("ADD YOUR"):
            replace_text_keep_format(shape, title, size=TYPE["cover_title"], color="primary", bold=True)
    set_notes(slide, spec.get("notes"))
    return slide


def build_closing(ctx: BuildContext, spec: dict):
    slide = clone_slide(ctx.prs, ctx.template_slide(PAGE_CLOSING))
    box = next(s for s in slide.shapes if s.has_text_frame and "谢谢观看" in s.text_frame.text)
    box.name = "HX_CLOSING_TITLE"
    box.left, box.top, box.width, box.height = _rect(CLOSING["title"])
    fill_text(box.text_frame, closing_paragraphs(spec.get("title") or "感谢聆听", spec.get("subtitle") or ""),
              anchor="middle")
    set_notes(slide, spec.get("notes"))
    return slide


# --------------------------------------------------------------------------- recipes
def _require_count(items, low, high, label):
    if not isinstance(items, list) or not low <= len(items) <= high:
        raise SpecError(f"{label} 需要 {low}–{high} 项，当前 {len(items) if isinstance(items, list) else 0} 项")


def _require_single_accent(items, label):
    if sum(1 for item in items if item.get("accent")) > 1:
        raise SpecError(f"{label}最多 1 个强调对象")


def _fit(height: float, available: float, label: str) -> None:
    if height > available + 1e-6:
        raise SpecError(f"{label} 内容超出正文区（需要 {height:.2f} in，可用 {available:.2f} in）；请压缩文字或拆页")


def _row_height(needed: float, available: float, stretch: bool, kind: str) -> float:
    """Height for a row of cards/tiles/panels: content height, or a share of the free space for the last block."""
    target = max(needed + 0.1, R["min_height"][kind])
    if stretch:
        target = max(target, available * R["stretch_ratio"].get(kind, R["stretch_ratio"]["default"]))
    return min(target, available)


def _item(entry) -> dict:
    return entry if isinstance(entry, dict) else {"title": "", "body": str(entry)}


def _conclusion(slide, text: str, y: float) -> float:
    check_length(text, capacity(CONTENT["w"] - 2 * PAD, TYPE["conclusion"]), "conclusion（一句话结论）")
    height = R["conclusion_h"]
    add_rounded(slide, CONTENT["x"], y, CONTENT["w"], height, fill="panel_fill", line=None, name="HX_CONCLUSION_BG")
    add_text(slide, CONTENT["x"] + PAD, y, CONTENT["w"] - 2 * PAD, height,
             [{"text": text, "size": TYPE["conclusion"], "color": "brand_blue", "bold": True}],
             name="HX_CONCLUSION_TEXT", anchor="middle")
    return y + height + GAP


def recipe_cards(b: Block, block: dict) -> float:
    items = block.get("items")
    _require_count(items, 2, DENSITY["max_major_cards"], "cards.items")
    _require_single_accent(items, "每组卡片")
    cols = columns(len(items))
    inner = cols[0][1] - 2 * PAD
    limit = DENSITY["max_cjk_chars_per_card"]
    for index, item in enumerate(items, 1):
        if cjk_count(item.get("body", "")) > limit:
            b.warnings.append(f"卡片 {index} 正文 {cjk_count(item['body'])} 字，超过 {limit} 字建议上限")
    needed = max(2 * PAD + text_height(item["title"], inner, TYPE["card_title"]) + 0.06
                 + text_height(item.get("body", ""), inner, TYPE["body"]) for item in items)
    _fit(needed, b.available, "cards")
    height = _row_height(needed, b.available, b.stretch, "cards")
    for index, ((x, w), item) in enumerate(zip(cols, items), 1):
        add_card(b.slide, x, b.y, w, height, item["title"], item.get("body", ""), name=f"HX_CARD_{index}",
                 accent=bool(item.get("accent")))
    return b.y + height


def recipe_kpis(b: Block, block: dict) -> float:
    items = block.get("items")
    _require_count(items, 2, DENSITY["max_major_cards"], "kpis.items")
    _require_single_accent(items, "每组指标")
    cols = columns(len(items))
    inner = cols[0][1] - 2 * PAD
    needed = max(2 * PAD + line_height(TYPE["key_number"]) + 0.04 + text_height(item.get("label", ""), inner, TYPE["body"])
                 + text_height(item.get("note", ""), inner, TYPE["body"]) for item in items)
    _fit(needed, b.available, "kpis")
    height = _row_height(needed, b.available, b.stretch, "kpis")
    for index, ((x, w), item) in enumerate(zip(cols, items), 1):
        tile = add_rounded(b.slide, x, b.y, w, height, name=f"HX_TILE_{index}")
        value_color = "brand_red" if item.get("accent") else "brand_blue"
        runs = [{"text": str(item["value"]), "size": TYPE["key_number"], "color": value_color, "bold": True}]
        if item.get("unit"):
            runs.append({"text": f" {item['unit']}", "size": TYPE["card_title"], "color": value_color, "bold": True})
        paragraphs = [{"runs": runs, "align": "center", "space_after": 2},
                      {"text": item.get("label", ""), "size": TYPE["body"], "bold": True, "align": "center"}]
        if item.get("note"):
            paragraphs.append({"text": item["note"], "size": TYPE["body"], "color": "primary", "align": "center"})
        fill_text(tile.text_frame, paragraphs, anchor="middle", margins=(PAD, PAD, PAD, PAD))
    return b.y + height


def recipe_bullets(b: Block, block: dict) -> float:
    items = [_item(entry) for entry in block.get("items") or []]
    _require_count(items, 1, DENSITY["max_top_level_bullets"], "bullets.items")
    pill_w, pill_h = R["pill"]["w"], R["pill"]["h"]
    row_gap = R["bullet_row_gap"]
    text_x = CONTENT["x"] + pill_w + GAP
    text_w = CONTENT["w"] - pill_w - GAP
    heights = [max(text_height(item.get("title", ""), text_w, TYPE["card_title"])
                   + text_height(item.get("body", ""), text_w, TYPE["body"]), pill_h) for item in items]
    needed = sum(heights) + row_gap * (len(items) - 1)
    _fit(needed, b.available, "bullets")
    spare = min((b.available - needed) / len(items), 0.25) if b.stretch else 0.0
    cursor = b.y
    for index, (item, h) in enumerate(zip(items, heights), 1):
        h += spare
        add_pill(b.slide, CONTENT["x"], cursor, pill_w, pill_h, f"{index:02d}",
                 fill="brand_red" if item.get("accent") else "brand_blue", name=f"HX_PILL_{index}")
        paragraphs = []
        if item.get("title"):
            paragraphs.append({"text": item["title"], "size": TYPE["card_title"], "color": "brand_blue", "bold": True})
        if item.get("body"):
            paragraphs.append({"text": item["body"], "size": TYPE["body"]})
        add_text(b.slide, text_x, cursor - 0.01, text_w, h, paragraphs, name=f"HX_POINT_{index}")
        cursor += h
        if index < len(items):
            add_rect(b.slide, text_x, cursor + row_gap / 2 - 0.005, text_w, 0.01, fill="light_gray",
                     name=f"HX_DIVIDER_{index}")
            cursor += row_gap
    return cursor


def recipe_process(b: Block, block: dict) -> float:
    steps = block.get("steps")
    _require_count(steps, 3, R["process"]["max_steps"], "process.steps")
    _require_single_accent(steps, "流程")
    cols = columns(len(steps))
    pill_h = R["process"]["pill_h"]
    first_center = cols[0][0] + cols[0][1] / 2
    last_center = cols[-1][0] + cols[-1][1] / 2
    add_rect(b.slide, first_center, b.y + pill_h / 2 - 0.015, last_center - first_center, 0.03,
             name="HX_PROCESS_LINE")
    panel_y = b.y + pill_h + 0.14
    available = CONTENT_BOTTOM - panel_y
    inner = cols[0][1] - 2 * PAD
    needed = max(2 * PAD + text_height(step["title"], inner, TYPE["card_title"]) + 0.05
                 + text_height(step.get("body", ""), inner, TYPE["body"]) for step in steps)
    _fit(needed, available, "process")
    height = _row_height(needed, available, b.stretch, "process")
    for index, ((x, w), step) in enumerate(zip(cols, steps), 1):
        label = step.get("label") or f"{index:02d}"
        pill_w = min(max(0.7, _units(label) * TYPE["body"] / 72 + 0.3), w - 0.1)
        accent = bool(step.get("accent"))
        add_pill(b.slide, x + (w - pill_w) / 2, b.y, pill_w, pill_h, label,
                 fill="brand_red" if accent else "brand_blue", name=f"HX_PILL_{index}")
        panel = add_rounded(b.slide, x, panel_y, w, height, fill="panel_fill", line="brand_red" if accent else None,
                            name=f"HX_STEP_{index}")
        paragraphs = [{"text": step["title"], "size": TYPE["card_title"], "color": "brand_blue", "bold": True,
                       "align": "center", "space_after": 3}]
        if step.get("body"):
            paragraphs.append({"text": step["body"], "size": TYPE["body"], "align": "center"})
        fill_text(panel.text_frame, paragraphs, margins=(PAD, PAD, PAD, PAD))
    return panel_y + height


CHART_TYPES = {"column": XL_CHART_TYPE.COLUMN_CLUSTERED, "bar": XL_CHART_TYPE.BAR_CLUSTERED,
               "stacked_column": XL_CHART_TYPE.COLUMN_STACKED, "stacked_bar": XL_CHART_TYPE.BAR_STACKED,
               "line": XL_CHART_TYPE.LINE_MARKERS, "pie": XL_CHART_TYPE.PIE, "doughnut": XL_CHART_TYPE.DOUGHNUT}


def style_chart(chart, kind: str, single_series: bool, highlight: int | None = None) -> None:
    chart.font.size = Pt(TYPE["chart"])
    chart.font.name = _ACTIVE_FONT
    chart.font.color.rgb = rgb("text_dark")
    chart.has_title = False
    plot = chart.plots[0]
    if kind in ("pie", "doughnut"):
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.RIGHT
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(TYPE["chart"])
        palette = CHARTS["pie_colors"]
        for index, point in enumerate(plot.series[0].points):
            color = "brand_red" if highlight == index else palette[index % len(palette)]
            if highlight is not None and highlight != index and color == "brand_red":
                color = "text_dark"
            point.format.fill.solid()
            point.format.fill.fore_color.rgb = rgb(color)
            point.format.line.color.rgb = rgb("white")
        plot.has_data_labels = True
        labels = plot.data_labels
        labels.number_format, labels.number_format_is_linked = "0%", False
        labels.show_percentage, labels.show_value = True, False
        labels.font.size = Pt(TYPE["chart"])
        labels.font.color.rgb = rgb("white")
        if kind == "doughnut":
            hole = plot._element.find(qn("c:holeSize"))
            if hole is None:
                hole = plot._element.makeelement(qn("c:holeSize"), {})
                plot._element.append(hole)
            hole.set("val", "60")
        return
    chart.has_legend = not single_series
    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.TOP
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(TYPE["chart"])
    palette = CHARTS["series_colors"]
    for index, series in enumerate(plot.series):
        color = palette[index % len(palette)]
        if kind == "line":
            series.format.line.color.rgb = rgb(color)
            series.format.line.width = Pt(2)
            series.smooth = False
            series.marker.format.fill.solid()
            series.marker.format.fill.fore_color.rgb = rgb(color)
            series.marker.format.line.color.rgb = rgb(color)
            continue
        series.format.fill.solid()
        series.format.fill.fore_color.rgb = rgb(color)
        if single_series and highlight is not None:
            point = series.points[highlight]
            point.format.fill.solid()
            point.format.fill.fore_color.rgb = rgb("brand_red")
    if kind != "line":
        plot.gap_width = 80
        if kind.startswith("stacked"):
            plot.overlap = 100
    value_axis, category_axis = chart.value_axis, chart.category_axis
    category_axis.tick_labels.font.size = Pt(TYPE["chart"])
    category_axis.format.line.color.rgb = rgb("light_gray")
    category_axis.has_major_gridlines = False
    value_axis.tick_labels.font.size = Pt(TYPE["chart"])
    if single_series and kind != "line":
        plot.has_data_labels = True
        plot.data_labels.font.size = Pt(TYPE["chart"])
        plot.data_labels.font.color.rgb = rgb("text_dark")
        plot.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
        value_axis.visible = False
        value_axis.has_major_gridlines = False
    else:
        value_axis.has_major_gridlines = True
        value_axis.major_gridlines.format.line.color.rgb = rgb("light_gray")
        value_axis.major_gridlines.format.line.width = Pt(0.5)
        value_axis.format.line.fill.background()


def chart_data_from(block: dict):
    categories = block.get("categories") or []
    series = block.get("series") or []
    if not categories or not series:
        raise SpecError("chart 需要 categories 与 series")
    data = CategoryChartData()
    data.categories = categories
    for item in series:
        values = item.get("values") or []
        if len(values) != len(categories):
            raise SpecError(f"系列 {item.get('name')!r} 的数值个数与分类数不一致")
        data.add_series(item.get("name") or "系列", values)
    return data


def recipe_chart(b: Block, block: dict) -> float:
    kind = block.get("kind", "column")
    if kind not in CHART_TYPES:
        raise SpecError(f"chart.kind 仅支持 {sorted(CHART_TYPES)}")
    if kind in ("pie", "doughnut"):
        if len(block.get("categories") or []) > CHARTS["max_donut_categories"]:
            raise SpecError(f"饼图/环形图不超过 {CHARTS['max_donut_categories']} 类；更多分类请改用排序条形图（kind=bar）")
        if len(block.get("series") or []) != 1:
            raise SpecError("饼图/环形图只能有 1 个系列")
    elif not block.get("unit"):
        b.warnings.append("图表未提供 unit（单位），请补充")
    points = [_item(entry) for entry in block.get("points") or []]
    if len(points) > DENSITY["max_supporting_conclusions"]:
        raise SpecError(f"图表页最多 {DENSITY['max_supporting_conclusions']} 个辅助结论（points）")
    c = R["chart"]
    available = min(b.available, float(block.get("height", b.available)))
    if available < c["min_h"]:
        raise SpecError(f"图表可用高度 {available:.2f} in 太小（至少 {c['min_h']} in）；请减少上方内容或拆页")
    chart_x, chart_w = CONTENT["x"], CONTENT["w"]
    if points:
        chart_w = CONTENT["w"] - c["side_w"] - GAP
        side_x = CONTENT["x"] + chart_w + GAP
        panel_h = min((available - GAP * (len(points) - 1)) / len(points), c["panel_max_h"])
        for index, point in enumerate(points, 1):
            panel = add_rounded(b.slide, side_x, b.y + (index - 1) * (panel_h + GAP), c["side_w"], panel_h,
                                fill="panel_fill", line=None, name=f"HX_SUPPORTING_CONCLUSION_{index}")
            paragraphs = []
            if point.get("title"):
                paragraphs.append({"text": point["title"], "size": TYPE["card_title"], "color": "brand_blue",
                                   "bold": True, "space_after": 3})
            if point.get("body"):
                paragraphs.append({"text": point["body"], "size": TYPE["body"]})
            fill_text(panel.text_frame, paragraphs, anchor="middle", margins=(PAD, PAD, PAD, PAD))
    caption = " ｜ ".join(part for part in (f"单位：{block['unit']}" if block.get("unit") else "",
                                           block.get("period") or "") if part)
    caption_h = c["caption_h"] if caption else 0.0
    if caption:
        add_text(b.slide, chart_x, b.y, chart_w, caption_h, [{"text": caption, "size": TYPE["chart"], "color": "primary"}],
                 name="HX_CHART_UNIT", anchor="middle")
    frame = b.slide.shapes.add_chart(CHART_TYPES[kind], Inches(chart_x), Inches(b.y + caption_h), Inches(chart_w),
                                     Inches(available - caption_h), chart_data_from(block))
    frame.name = "HX_CHART_1"
    style_chart(frame.chart, kind, len(block["series"]) == 1, block.get("highlight"))
    return b.y + available


def recipe_table(b: Block, block: dict) -> float:
    header = block.get("columns") or []
    rows = block.get("rows") or []
    t = R["table"]
    if not header or not rows:
        raise SpecError("table 需要 columns 与 rows")
    if any(len(row) != len(header) for row in rows):
        raise SpecError("table 每行单元格数必须与 columns 一致")
    if len(header) > t["max_cols"]:
        raise SpecError(f"表格不超过 {t['max_cols']} 列")
    available = min(b.available, float(block.get("height", b.available)))
    row_h = min(t["row_h_max"], (available - t["header_h"]) / len(rows))
    if row_h < t["row_h_min"]:
        raise SpecError(f"表格 {len(rows)} 行放不下（行高 {row_h:.2f} in < {t['row_h_min']:.2f} in）；请拆页或精简")
    total_h = t["header_h"] + row_h * len(rows)
    frame = b.slide.shapes.add_table(len(rows) + 1, len(header), Inches(CONTENT["x"]), Inches(b.y),
                                     Inches(CONTENT["w"]), Inches(total_h))
    frame.name = "HX_TABLE_1"
    table = frame.table
    table.first_row, table.horz_banding = True, False
    ratios = block.get("col_widths") or [1] * len(header)
    for index, ratio in enumerate(ratios):
        table.columns[index].width = Emu(int(Inches(CONTENT["w"] * ratio / float(sum(ratios)))))
    table.rows[0].height = Emu(int(Inches(t["header_h"])))
    for index in range(len(rows)):
        table.rows[index + 1].height = Emu(int(Inches(row_h)))
    emphasis = set(block.get("emphasis_rows") or [])

    def write(cell, text, *, header_cell=False, row_index=0, col_index=0):
        cell.fill.solid()
        cell.fill.fore_color.rgb = rgb("brand_blue" if header_cell else ("panel_fill" if row_index % 2 else "white"))
        cell.vertical_anchor = ANCHOR["middle"]
        # Table cells use tcPr margins (bodyPr insets are ignored inside tables).
        cell.margin_left = cell.margin_right = Inches(0.08)
        cell.margin_top = cell.margin_bottom = Inches(0.03)
        emphasized = not header_cell and row_index in emphasis
        fill_text(cell.text_frame, [{
            "text": str(text), "size": TYPE["body"], "bold": header_cell or emphasized,
            "color": "white" if header_cell else ("brand_red" if emphasized else "text_dark"),
            "align": "left" if col_index == 0 and not header_cell else "center"}],
            anchor="middle", margins=(0.08, 0.03, 0.08, 0.03))

    for col, text in enumerate(header):
        write(table.cell(0, col), text, header_cell=True, col_index=col)
    for r, row in enumerate(rows):
        for col, text in enumerate(row):
            write(table.cell(r + 1, col), text, row_index=r, col_index=col)
    return b.y + total_h


def _panel(slide, x, y, w, h, data: dict, name: str):
    panel = add_rounded(slide, x, y, w, h, name=name)
    paragraphs = []
    if data.get("title"):
        paragraphs.append({"text": data["title"], "size": TYPE["card_title"], "color": "brand_blue", "bold": True,
                           "space_after": 4})
    paragraphs += [{"text": bullet, "size": TYPE["body"], "bullet": True, "space_after": 3}
                   for bullet in data.get("bullets") or []]
    if data.get("body"):
        paragraphs.append({"text": data["body"], "size": TYPE["body"]})
    fill_text(panel.text_frame, paragraphs, margins=(PAD, PAD, PAD, PAD))
    return panel


def _panel_height(data: dict, w: float) -> float:
    inner = w - 2 * PAD - 0.16
    h = 2 * PAD + (text_height(data.get("title", ""), inner, TYPE["card_title"]) + 0.06 if data.get("title") else 0)
    h += sum(text_height(bullet, inner, TYPE["body"]) + 0.04 for bullet in data.get("bullets") or [])
    return h + text_height(data.get("body", ""), inner, TYPE["body"])


def recipe_two_column(b: Block, block: dict) -> float:
    left, right = block.get("left") or {}, block.get("right") or {}
    bullets = len(left.get("bullets") or []) + len(right.get("bullets") or [])
    if bullets > DENSITY["max_top_level_bullets"]:
        raise SpecError(f"双栏共 {bullets} 个要点，单页不超过 {DENSITY['max_top_level_bullets']} 个；请精简或拆页")
    cols = columns(2)
    needed = max(_panel_height(left, cols[0][1]), _panel_height(right, cols[1][1]))
    _fit(needed, b.available, "two_column")
    height = _row_height(needed, b.available, b.stretch, "two_column")
    _panel(b.slide, cols[0][0], b.y, cols[0][1], height, left, "HX_PANEL_1")
    _panel(b.slide, cols[1][0], b.y, cols[1][1], height, right, "HX_PANEL_2")
    return b.y + height


def recipe_image_text(b: Block, block: dict) -> float:
    cols = columns(2)
    image_col, text_col = (cols[0], cols[1]) if block.get("image_side", "left") == "left" else (cols[1], cols[0])
    image_h = min(b.available, image_col[1] * (9 / 16 if block.get("ratio", "16:9") == "16:9" else 3 / 4))
    _, warning = add_image_fill(b.slide, block["image"], image_col[0], b.y, image_col[1], image_h)
    if warning:
        b.warnings.append(warning)
    text = {k: block.get(k) for k in ("title", "bullets", "body")}
    if len(text.get("bullets") or []) > DENSITY["max_top_level_bullets"]:
        raise SpecError(f"图文页要点不超过 {DENSITY['max_top_level_bullets']} 个")
    needed = _panel_height(text, text_col[1])
    _fit(needed, b.available, "image_text")
    panel_h = max(image_h, min(needed + 0.1, b.available))
    _panel(b.slide, text_col[0], b.y, text_col[1], panel_h, text, "HX_PANEL_1")
    return b.y + panel_h


RECIPES = {"cards": recipe_cards, "kpis": recipe_kpis, "bullets": recipe_bullets, "process": recipe_process,
           "chart": recipe_chart, "table": recipe_table, "two_column": recipe_two_column,
           "image_text": recipe_image_text}


def build_content(ctx: BuildContext, spec: dict):
    """Body page: bar title + optional conclusion band + stacked recipe blocks + source line."""
    slide = new_body_slide(ctx.prs)
    set_body_title(slide, spec.get("title") or "")
    set_section_tab(slide, spec.get("number"))
    y = CONTENT["y"]
    if spec.get("conclusion"):
        y = _conclusion(slide, spec["conclusion"], y)
    layout = spec.get("layout")
    blocks = layout if isinstance(layout, list) else ([layout] if layout else [])
    if len(blocks) > DENSITY["max_blocks_per_slide"]:
        raise SpecError(f"单页最多叠放 {DENSITY['max_blocks_per_slide']} 个版式块（如 kpis + bullets）；内容更多请拆页")
    for index, block in enumerate(blocks):
        kind = block.get("type")
        if kind not in RECIPES:
            raise SpecError(f"layout.type 仅支持 {sorted(RECIPES)}")
        last = index == len(blocks) - 1
        if kind in ("chart", "table") and not last and "height" not in block:
            raise SpecError(f"{kind} 放在上方时必须指定 height（英寸）")
        y = RECIPES[kind](Block(slide, y, ctx.warnings, stretch=last), block) + GAP
    add_source(slide, spec.get("source") or "")
    set_notes(slide, spec.get("notes"))
    return slide


# --------------------------------------------------------------------------- reuse any template page
def _shape_kind(shape) -> str:
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        return "group"
    if getattr(shape, "has_chart", False) and shape.has_chart:
        return "chart"
    if getattr(shape, "has_table", False) and shape.has_table:
        return "table"
    return "picture" if shape.shape_type == MSO_SHAPE_TYPE.PICTURE else "shape"


def list_slots(slide) -> list[dict]:
    return [{"index": index, "kind": _shape_kind(shape), "name": shape.name,
             "text": shape.text_frame.text if shape.has_text_frame else ""}
            for index, shape in enumerate(iter_shapes(slide.shapes))]


def _enforce_chart_floor(chart) -> None:
    """Raise every explicit chart text size below the chart floor (template charts use 9 pt)."""
    floor = int(round(FLOOR["chart"] * 100))
    for element in chart._chartSpace.iter():
        size = element.get("sz") if element.tag.rsplit("}", 1)[-1] in ("rPr", "defRPr", "endParaRPr") else None
        if size is not None and size.isdigit() and int(size) < floor:
            element.set("sz", str(floor))


def build_template_page(ctx: BuildContext, spec: dict):
    source = int(spec.get("source", 0))
    if not 1 <= source <= PAGE_CLOSING:
        raise SpecError(f"template.source 需在 1–{PAGE_CLOSING}")
    original = ctx.template_slide(source)
    if _has_chart(original):
        if source in ctx.used_chart_sources:
            raise SpecError(f"含图表的模板页 {source} 只能复用一次；其余请用 content + chart 生成")
        ctx.used_chart_sources.add(source)
    slide = clone_slide(ctx.prs, original)
    shapes = list(iter_shapes(slide.shapes))
    for key, text in (spec.get("texts") or {}).items():
        index = int(key)
        if not 0 <= index < len(shapes) or not shapes[index].has_text_frame:
            raise SpecError(f"template.texts 的槽位 {index} 不是文本槽；先运行 slots --slide {source}")
        replace_text_keep_format(shapes[index], text)
        for paragraph in shapes[index].text_frame.paragraphs:  # template samples use 10 pt; the floor is 11 pt
            for run in paragraph.runs:
                if run.font.size is not None and run.font.size.pt < FLOOR["body"]:
                    run.font.size = Pt(FLOOR["body"])
    charts = [s for s in shapes if _shape_kind(s) == "chart"]
    for frame in charts:
        _enforce_chart_floor(frame.chart)
    for item in spec.get("charts") or []:
        chart = charts[int(item.get("index", 0))].chart
        chart.replace_data(chart_data_from(item))
        if "chart_title" in item:
            chart.has_title = bool(item["chart_title"])
            if item["chart_title"]:
                chart.chart_title.text_frame.text = item["chart_title"]
    tables = [s for s in shapes if _shape_kind(s) == "table"]
    for item in spec.get("tables") or []:
        table = tables[int(item.get("index", 0))].table
        for r, row in enumerate(item.get("rows") or []):
            for c, value in enumerate(row):
                if value is None:
                    continue
                cell = table.cell(r, c)
                if cell.text_frame.text.strip():
                    replace_text_keep_format(cell, str(value))
                else:
                    fill_text(cell.text_frame, [{"text": str(value)}], anchor="middle")
    for index in sorted({int(i) for i in spec.get("delete") or []}, reverse=True):
        if 0 <= index < len(shapes):
            _remove(shapes[index])
    if spec.get("title"):
        set_body_title(slide, spec["title"])
    set_section_tab(slide, spec.get("number"))
    add_source(slide, spec.get("source_note") or "")
    set_notes(slide, spec.get("notes"))
    return slide


# --------------------------------------------------------------------------- deck
PAGE_BUILDERS = {"cover": build_cover, "toc": build_toc, "section": build_section, "content": build_content,
                 "template": build_template_page, "closing": build_closing}


def audit_markers(prs) -> list[str]:
    findings = []
    for number, slide in enumerate(prs.slides, 1):
        parts = []
        for shape in iter_shapes(slide.shapes):
            if shape.has_text_frame:
                parts.append(shape.text_frame.text)
            elif _shape_kind(shape) == "table":
                parts += [cell.text_frame.text for row in shape.table.rows for cell in row.cells]
        text = " ".join(parts).casefold()
        findings += [f"第 {number} 页残留模板文字：{m}" for m in TEMPLATE_MARKERS if m.casefold() in text]
    return findings


def build_deck(spec: dict):
    """Build a Presentation from a spec. Returns (prs, report)."""
    pages = spec.get("slides") or []
    if not pages:
        raise SpecError("spec.slides 不能为空")
    prs = open_template()
    original_count = len(prs.slides)
    ctx = BuildContext(prs)
    with use_font(spec.get("font")):
        for number, page in enumerate(pages, 1):
            kind = page.get("type", "content")
            if kind not in PAGE_BUILDERS:
                raise SpecError(f"第 {number} 页：未知页面类型 {kind!r}")
            try:
                PAGE_BUILDERS[kind](ctx, page)
            except SpecError as exc:
                raise SpecError(f"第 {number} 页（{kind}）：{exc}") from None
    _drop_original_slides(prs, original_count)
    if spec.get("title"):
        prs.core_properties.title = spec["title"]
    prs.core_properties.keywords = "huaxin-pptx"
    types = [page.get("type", "content") for page in pages]
    if spec.get("complete", True):
        if types[0] != "cover":
            ctx.warnings.append("完整稿应以首页（cover）开始")
        if types[-1] != "closing":
            ctx.warnings.append("完整稿应以答谢尾页（closing）结束")
    ctx.warnings.extend(audit_markers(prs))
    report = {"slides": len(prs.slides), "types": types, "warnings": ctx.warnings, "status": "BUILT",
              "next_steps": ["python scripts/check_deck.py <output.pptx>",
                             "python scripts/office/render_visual.py --pptx <output.pptx> --output-dir <dir> --json <evidence.json>",
                             "逐页查看 PNG 后执行 render_visual.py --attest <evidence.json> --reviewed-pages all --review-notes ..."]}
    return prs, report


def save(prs, output: str | Path, force: bool = False) -> Path:
    output = Path(output).expanduser().resolve()
    if output == ROOT or ROOT in output.parents:
        raise SpecError("输出必须位于技能目录之外")
    if output.suffix.lower() != ".pptx":
        raise SpecError("输出文件必须以 .pptx 结尾")
    if output.exists() and not force:
        raise SpecError(f"输出已存在：{output}；确需覆盖请加 --force")
    output.parent.mkdir(parents=True, exist_ok=True)
    fd, temp_name = tempfile.mkstemp(prefix=".hx-", suffix=".pptx", dir=output.parent)
    os.close(fd)
    try:
        prs.save(temp_name)
        Presentation(temp_name)  # reopen check
        os.replace(temp_name, output)
    finally:
        if os.path.exists(temp_name):
            os.unlink(temp_name)
    return output


EXAMPLE_SPEC = {
    "title": "集团信息条线 2026 年工作汇报",
    "slides": [
        {"type": "cover", "title": "集团信息条线\n2026 年工作汇报", "subtitle": "重点进展与下一步安排",
         "presenter": "", "date": ""},
        {"type": "toc", "items": [{"title": "年度工作概览", "desc": "目标完成情况与关键指标"},
                                  {"title": "重点项目进展"}, {"title": "问题与下一步"}]},
        {"type": "section", "number": "01", "title": "年度工作概览"},
        {"type": "content", "number": "01", "title": "四项核心指标全部达成年度目标",
         "conclusion": "系统可用率与交付效率同步提升，成本保持受控",
         "layout": [
             {"type": "kpis", "items": [
                 {"value": "99.95", "unit": "%", "label": "核心系统可用率", "note": "目标 99.9%"},
                 {"value": "36", "unit": "项", "label": "上线项目数", "note": "同比 +9 项"},
                 {"value": "-12", "unit": "%", "label": "单位运维成本", "accent": True, "note": "自动化巡检贡献最大"},
                 {"value": "4.6", "unit": "分", "label": "用户满意度", "note": "满分 5 分"}]},
             {"type": "bullets", "items": [
                 {"title": "稳定性", "body": "全年无一级故障，核心系统平均恢复时间缩短至 18 分钟。"},
                 {"title": "效率", "body": "敏捷交付覆盖全部核心系统，需求平均交付周期缩短 30%。"}]}],
         "source": "集团信息部年度统计（示例数据）",
         "notes": "口播：先讲可用率，再讲成本下降的原因。"},
        {"type": "content", "number": "01", "title": "三类重点工作形成协同推进格局",
         "conclusion": "底座、系统、安全三线并进，支撑业务全面上云",
         "layout": {"type": "cards", "items": [
             {"title": "数字化底座", "body": "完成数据中台二期建设，统一主数据与指标口径，数据服务调用量同比翻番。"},
             {"title": "业务系统", "body": "供应链与财务共享系统上线，覆盖全部生产基地，单据处理效率提升 40%。",
              "accent": True},
             {"title": "安全合规", "body": "通过等保三级复评，建立常态化攻防演练机制，全年未发生重大安全事件。"}]}},
        {"type": "content", "number": "02", "title": "季度收入持续增长，四季度增幅最大",
         "layout": {"type": "chart", "kind": "column", "unit": "万元", "period": "2026 年 Q1–Q4",
                    "categories": ["一季度", "二季度", "三季度", "四季度"],
                    "series": [{"name": "收入", "values": [9, 30, 90, 112]}], "highlight": 3,
                    "points": [{"title": "增长主因", "body": "新业务线于三季度集中上线。"},
                               {"title": "风险提示", "body": "四季度含一次性项目收入。"}]},
         "source": "财务月报（示例数据）"},
        {"type": "content", "number": "03", "title": "下一步按三个阶段推进落地",
         "conclusion": "上半年完成试点验证，年中实现全集团推广",
         "layout": {"type": "process", "steps": [
             {"label": "1 月", "title": "方案确认", "body": "完成需求评审与预算批复，明确各基地责任人"},
             {"label": "3 月", "title": "试点上线", "body": "选取两家基地试运行，沉淀推广手册与培训材料"},
             {"label": "6 月", "title": "全面推广", "body": "覆盖全部基地并完成验收，转入常态化运营",
              "accent": True}]}},
        {"type": "closing", "title": "感谢聆听", "subtitle": "敬请批评指正"},
    ],
}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest="command", required=True)
    build = sub.add_parser("build", help="Build a deck from a JSON spec")
    build.add_argument("--spec", required=True, type=Path)
    build.add_argument("--output", required=True, type=Path)
    build.add_argument("--force", action="store_true")
    slots = sub.add_parser("slots", help="List replaceable slots of a template page")
    slots.add_argument("--slide", required=True, type=int)
    sub.add_parser("example-spec", help="Print a complete example spec")
    args = parser.parse_args()
    if args.command == "example-spec":
        result: dict = EXAMPLE_SPEC
    elif args.command == "slots":
        slide = open_template().slides[args.slide - 1]
        result = {"slide": args.slide, "layout": slide.slide_layout.name, "has_chart": _has_chart(slide),
                  "slots": [s for s in list_slots(slide) if s["kind"] != "shape" or s["text"].strip()]}
    else:
        try:
            prs, result = build_deck(json.loads(args.spec.read_text(encoding="utf-8")))
            result["output"] = str(save(prs, args.output, args.force))
        except SpecError as exc:
            print(json.dumps({"status": "SPEC_ERROR", "error": str(exc)}, ensure_ascii=False, indent=2))
            return 2
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
