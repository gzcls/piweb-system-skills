#!/usr/bin/env python3
"""Validate this skill's metadata, files, asset hashes and optional PPTX structure.

Requires PyYAML and ``jsonschema``. The latter is used for the complete
theme-config schema validation; the additional hand-written checks below cover
cross-field invariants that JSON Schema cannot express (for example, object
coordinates fitting inside the selected canvas).
This is NOT a visual-quality checker or PowerPoint/WPS compatibility certification.
When ``--visual-evidence`` is supplied, this script also verifies a separately
generated LibreOffice/PyMuPDF evidence manifest.  The evidence manifest must
explicitly record a human page review; a structural report alone never implies
that a slide was visually inspected.
"""
from __future__ import annotations

import argparse
import ast
import hashlib
import json
import posixpath
import re
import sys
import xml.etree.ElementTree as ET
from pathlib import Path, PurePosixPath
from urllib.parse import unquote
from zipfile import BadZipFile, ZipFile

# Validation is a read-only operation. Keep optional/imported helpers from
# leaving ``__pycache__`` or ``.pyc`` files inside the distributable skill
# directory when this script is run directly.
sys.dont_write_bytecode = True

ROOT = Path(__file__).resolve().parents[1]
NS = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'r': 'http://schemas.openxmlformats.org/package/2006/relationships',
      'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'}
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
FONT_SUFFIXES = {'.ttf', '.ttc', '.otf', '.woff', '.woff2', '.eot', '.fntdata'}


def jpeg_size(data: bytes) -> tuple[int, int]:
    if data[:2] != b'\xff\xd8':
        raise ValueError('Not a JPEG file.')
    index = 2
    sof = {0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf}
    while index < len(data):
        if data[index] != 0xff:
            raise ValueError('Invalid JPEG marker.')
        while index < len(data) and data[index] == 0xff:
            index += 1
        if index >= len(data):
            break
        marker = data[index]
        index += 1
        if marker in {0xd8, 0x01} or 0xd0 <= marker <= 0xd7:
            continue
        if marker in {0xd9, 0xda} or index + 2 > len(data):
            break
        length = int.from_bytes(data[index:index + 2], 'big')
        if length < 2 or index + length > len(data):
            raise ValueError('Truncated JPEG segment.')
        if marker in sof:
            if length < 8:
                raise ValueError('Truncated JPEG dimensions.')
            height = int.from_bytes(data[index + 3:index + 5], 'big')
            width = int.from_bytes(data[index + 5:index + 7], 'big')
            return width, height
        index += length
    raise ValueError('JPEG dimensions not found.')


def validate_visual_evidence(path: Path, target: Path | None = None) -> tuple[dict, list[str], dict | None]:
    """Validate a render_visual.py evidence manifest and optionally match a deck.

    The manifest is intentionally separate from this package: PDF/PNG output is
    generated in a caller-owned directory.  Hashes and page dimensions make the
    recorded result reproducible, while ``manual_review.status`` prevents an
    unattended render from being reported as a visual approval.
    """
    errors: list[str] = []
    try:
        evidence = json.loads(path.read_text(encoding='utf-8'))
    except (OSError, ValueError) as exc:
        return {}, [f'Visual evidence could not be read: {exc}'], None
    if not isinstance(evidence, dict):
        return {}, ['Visual evidence must be a JSON object.'], None
    if evidence.get('schema_version') != 1:
        errors.append('Visual evidence schema_version must be 1.')
    if evidence.get('status') != 'PASS' or evidence.get('render_status') != 'PASS':
        errors.append('Visual evidence must have status=PASS and render_status=PASS.')
    renderer = evidence.get('renderer')
    if not isinstance(renderer, dict) or renderer.get('name') != 'LibreOffice':
        errors.append('Visual evidence renderer.name must be LibreOffice.')
    manual = evidence.get('manual_review')
    if not isinstance(manual, dict) or manual.get('status') != 'PASS':
        errors.append('Visual evidence manual_review.status must be PASS after PNG inspection.')
    reviewed_pages = manual.get('pages') if isinstance(manual, dict) else None
    if not isinstance(reviewed_pages, list) or any(
        not isinstance(page, int) or isinstance(page, bool) or page < 1
        for page in reviewed_pages
    ):
        errors.append('Visual evidence manual_review.pages must be a list of positive page numbers.')
        reviewed_page_set: set[int] = set()
    else:
        reviewed_page_set = set(reviewed_pages)
    decks = evidence.get('decks')
    if not isinstance(decks, list) or not decks:
        errors.append('Visual evidence must contain at least one deck.')
        decks = []

    def check_hash(file_path: Path, expected: object, label: str) -> None:
        if not file_path.is_file():
            errors.append(f'Visual evidence {label} is missing: {file_path}')
            return
        if not isinstance(expected, str) or not re.fullmatch(r'[0-9a-f]{64}', expected):
            errors.append(f'Visual evidence {label} has an invalid SHA-256.')
            return
        try:
            actual = hashlib.sha256(file_path.read_bytes()).hexdigest()
        except OSError as exc:
            errors.append(f'Visual evidence {label} could not be hashed: {exc}')
            return
        if actual != expected:
            errors.append(f'Visual evidence {label} SHA-256 does not match.')

    matched: dict | None = None
    target_resolved = target.resolve() if target is not None else None
    for index, deck in enumerate(decks):
        if not isinstance(deck, dict):
            errors.append(f'Visual evidence decks[{index}] must be an object.')
            continue
        pptx_value, pdf_value = deck.get('pptx'), deck.get('pdf')
        if not isinstance(pptx_value, str) or not isinstance(pdf_value, str):
            errors.append(f'Visual evidence decks[{index}] must contain pptx and pdf paths.')
            continue
        pptx_path, pdf_path = Path(pptx_value), Path(pdf_value)
        check_hash(pptx_path, deck.get('pptx_sha256'), f'decks[{index}].pptx')
        check_hash(pdf_path, deck.get('pdf_sha256'), f'decks[{index}].pdf')
        page_count = deck.get('page_count')
        pages = deck.get('pages')
        if not isinstance(page_count, int) or page_count <= 0 or not isinstance(pages, list) or len(pages) != page_count:
            errors.append(f'Visual evidence decks[{index}] page_count/pages are inconsistent.')
        else:
            if isinstance(manual, dict) and manual.get('status') == 'PASS':
                expected_pages = set(range(1, page_count + 1))
                if reviewed_page_set != expected_pages:
                    errors.append(
                        f'Visual evidence decks[{index}] is marked reviewed but manual_review.pages '
                        f'does not cover every page (expected 1-{page_count}).'
                    )
            for page_index, page in enumerate(pages, start=1):
                if not isinstance(page, dict) or page.get('page') != page_index:
                    errors.append(f'Visual evidence decks[{index}] page numbering is invalid.')
                    continue
                png_value = page.get('png')
                if not isinstance(png_value, str):
                    errors.append(f'Visual evidence decks[{index}].pages[{page_index}] lacks a PNG path.')
                    continue
                check_hash(Path(png_value), page.get('png_sha256'), f'decks[{index}].pages[{page_index}].png')
                for dimension in ('width_pt', 'height_pt', 'raster_width_px', 'raster_height_px'):
                    value = page.get(dimension)
                    if not isinstance(value, (int, float)) or isinstance(value, bool) or value <= 0:
                        errors.append(f'Visual evidence decks[{index}].pages[{page_index}] has invalid {dimension}.')
        if target_resolved is not None:
            try:
                same_path = pptx_path.resolve() == target_resolved
            except OSError:
                same_path = False
            same_hash = False
            if target_resolved.is_file() and isinstance(deck.get('pptx_sha256'), str):
                same_hash = deck.get('pptx_sha256') == hashlib.sha256(target_resolved.read_bytes()).hexdigest()
            if same_path or same_hash:
                matched = deck
    if target_resolved is not None and matched is None:
        errors.append(f'Visual evidence does not contain the inspected PPTX: {target_resolved}')
    return evidence, errors, matched


_CJK_RE = re.compile(r'[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]')
_SUPPORTING_CONCLUSION_RE = re.compile(r'(^|[_ -])SUPPORTING[_ -]CONCLUSION([_ -]|$)')
_CARD_NAME_RE = re.compile(r'(^|[_ -])(CARD|TILE|PANEL)([_ -]|$)')
_PILL_NAME_RE = re.compile(r'(^|[_ -])(PILL|TAG|BADGE)([_ -]|$)')
_STEP_NAME_RE = re.compile(r'(^|[_ -])STEP([_ -]|$)')
_CAPTION_NAME_RE = re.compile(r'(SOURCE|FOOTNOTE|CAPTION|METADATA|UNIT|LABEL|NOTE)')
_CHROME_NAME_RE = re.compile(r'^HX_(BODY_TITLE|SECTION_TAB|COVER_|CLOSING_)')
_ROUNDED_GEOMETRIES = {
    'roundRect', 'round1Rect', 'round2SameRect', 'round2DiagRect', 'snipRoundRect',
}


def _font_sizes_from_xml(root: ET.Element) -> list[float]:
    """Read explicit and inherited text-run sizes from an Office XML part."""
    sizes: list[float] = []
    for element in root.iter():
        if element.tag.rsplit('}', 1)[-1] not in {'rPr', 'defRPr', 'endParaRPr'}:
            continue
        raw_size = element.get('sz')
        if raw_size is None:
            continue
        try:
            sizes.append(int(raw_size) / 100)
        except (TypeError, ValueError):
            continue
    return sizes


def _run_sizes(shape: ET.Element) -> list[float]:
    """Explicit sizes of runs that actually carry text (inherited sizes are unknown here)."""
    sizes: list[float] = []
    for run in shape.iter(f'{{{NS["a"]}}}r'):
        text = ''.join(t.text or '' for t in run.findall('a:t', NS)).strip()
        properties = run.find('a:rPr', NS)
        if not text or properties is None or properties.get('sz') is None:
            continue
        try:
            sizes.append(int(properties.get('sz')) / 100)
        except ValueError:
            continue
    return sizes


def _shape_name(shape: ET.Element) -> str:
    for path in ('p:nvSpPr/p:cNvPr', 'p:nvPicPr/p:cNvPr', 'p:nvGraphicFramePr/p:cNvPr', 'p:nvGrpSpPr/p:cNvPr'):
        name_element = shape.find(path, NS)
        if name_element is not None:
            return name_element.get('name', '').strip()
    return ''


def _shape_geometry(shape: ET.Element) -> str | None:
    geometry = shape.find('p:spPr/a:prstGeom', NS)
    return geometry.get('prst') if geometry is not None else None


def _shape_text(shape: ET.Element) -> str:
    return ''.join(text.text or '' for text in shape.findall('.//a:t', NS))


def _body_text(shape: ET.Element) -> str:
    """Text excluding the first paragraph when a shape has several (first = card title)."""
    paragraphs = shape.findall('p:txBody/a:p', NS)
    texts = [''.join(t.text or '' for t in p.findall('.//a:t', NS)) for p in paragraphs]
    texts = [t for t in texts if t.strip()]
    return ''.join(texts[1:]) if len(texts) > 1 else ''.join(texts)


def _shape_has_font_size(shape: ET.Element) -> bool:
    """Return whether text in a shape has an explicit or inherited size.

    ``defRPr`` and level defaults are included because Office commonly stores
    the inherited size there instead of on every run.  A text-bearing shape
    with no size anywhere is still ambiguous after a theme substitution, so it
    is surfaced as a diagnostic rather than silently accepted.
    """
    return any(
        element.get('sz') is not None
        for element in shape.iter()
        if element.tag.rsplit('}', 1)[-1] in {'rPr', 'defRPr', 'endParaRPr'}
    )


def _corner_radius_in(shape: ET.Element) -> tuple[float, float, float] | None:
    """Return (radius, width, height) in inches for a roundRect shape or picture."""
    geometry = shape.find('p:spPr/a:prstGeom', NS)
    transform = shape.find('p:spPr/a:xfrm', NS)
    if geometry is None or geometry.get('prst') != 'roundRect' or transform is None:
        return None
    extent = transform.find('a:ext', NS)
    if extent is None:
        return None
    width, height = int(extent.get('cx', '0')) / 914400, int(extent.get('cy', '0')) / 914400
    adjust = 16667
    guide = geometry.find('a:avLst/a:gd', NS)
    if guide is not None:
        match = re.search(r'val\s+(-?\d+)', guide.get('fmla', ''))
        if match:
            adjust = int(match.group(1))
    return min(width, height) * min(max(adjust, 0), 50000) / 100000, width, height


def validate_pptx_density(stats: dict, density: dict) -> list[str]:
    """Return deterministic per-slide density, typography and component diagnostics.

    Limits are PER SLIDE (the rules say "每页"), so a long deck is never flagged
    merely for having many slides.  Card limits apply to shapes named ``CARD``,
    ``TILE`` or ``PANEL`` plus unnamed text-bearing rounded rectangles; pills
    (``PILL/TAG/BADGE``), process steps (``STEP``) and supporting conclusions
    are counted separately.
    """
    diagnostics: list[str] = []

    def positive_int(key: str) -> int | None:
        value = density.get(key)
        return value if isinstance(value, int) and not isinstance(value, bool) and value >= 0 else None

    def per_slide(stat_key: str, limit_key: str, label: str) -> None:
        limit = positive_int(limit_key)
        values = stats.get(stat_key)
        if limit is None or not isinstance(values, list):
            return
        for index, value in enumerate(values, start=1):
            if isinstance(value, int) and value > limit:
                diagnostics.append(f'Slide {index} has {value} {label}; density.{limit_key} is {limit}.')

    per_slide('bullet_paragraphs_per_slide', 'max_top_level_bullets', 'bullet paragraphs')
    per_slide('named_cards_per_slide', 'max_major_cards', 'named card/panel shapes')
    chart_limit = positive_int('max_major_charts')
    chart_values = stats.get('major_charts_per_slide')
    raw_chart_values = stats.get('charts_per_slide')
    if not isinstance(chart_values, list):
        chart_values = raw_chart_values
    if chart_limit is not None and isinstance(chart_values, list):
        for index, value in enumerate(chart_values, start=1):
            if isinstance(value, int) and value > chart_limit:
                reported = (raw_chart_values[index - 1]
                            if isinstance(raw_chart_values, list) and index - 1 < len(raw_chart_values)
                            else value)
                diagnostics.append(f'Slide {index} has {reported} charts; density.max_major_charts is {chart_limit}.')
    per_slide('supporting_conclusion_counts', 'max_supporting_conclusions', 'named supporting conclusion shapes')
    per_slide('typography_levels_per_slide', 'max_typography_levels', 'explicit typography levels')
    cjk_limit = positive_int('max_cjk_chars_per_card')
    if cjk_limit is not None and isinstance(stats.get('max_cjk_chars_per_named_card'), int):
        if stats['max_cjk_chars_per_named_card'] > cjk_limit:
            diagnostics.append(
                f"A named card contains {stats['max_cjk_chars_per_named_card']} CJK characters in its body; "
                f"density.max_cjk_chars_per_card is {cjk_limit}."
            )
    unnamed_rounded = stats.get('unnamed_text_rounded_shape_count')
    if isinstance(unnamed_rounded, int) and unnamed_rounded:
        diagnostics.append(
            f"PPTX has {unnamed_rounded} text-bearing rounded component(s) without a "
            "semantic CARD/TILE/PANEL/PILL/STEP name; name them so density and radius checks apply."
        )
    invalid_card_shapes = stats.get('card_shape_violations')
    if isinstance(invalid_card_shapes, list) and invalid_card_shapes:
        diagnostics.append(
            "Semantic card component(s) do not use a rounded rectangle geometry: "
            + ', '.join(str(item) for item in invalid_card_shapes[:8])
            + (" …" if len(invalid_card_shapes) > 8 else '')
        )
    for item in (stats.get('radius_violations') or [])[:12]:
        diagnostics.append(f'Corner radius violation: {item}')
    for item in (stats.get('radius_inconsistencies') or [])[:12]:
        diagnostics.append(f'Inconsistent card radii on one slide: {item}')
    inheritance_gaps = stats.get('text_shapes_without_font_size')
    if isinstance(inheritance_gaps, int) and inheritance_gaps:
        diagnostics.append(
            f"PPTX has {inheritance_gaps} text-bearing shape(s) without explicit or "
            "inherited font size; typography may change with the active theme."
        )
    donut_limit = density.get('max_donut_categories')
    if isinstance(donut_limit, int) and not isinstance(donut_limit, bool):
        max_donut = stats.get('max_donut_categories')
        if isinstance(max_donut, int) and max_donut > donut_limit:
            diagnostics.append(
                f"A donut chart has {max_donut} categories; charts.max_donut_categories is {donut_limit}."
            )
    for item in (stats.get('body_font_floor_violations') or [])[:12]:
        diagnostics.append(f'Text below the body floor: {item}')
    chart_floor = density.get('chart_min_runtime_pt')
    chart_min = stats.get('min_chart_font_size_pt')
    if isinstance(chart_floor, (int, float)) and not isinstance(chart_floor, bool):
        if isinstance(chart_min, (int, float)) and chart_min < chart_floor * stats.get('canvas_scale', 1.0) - 0.01:
            diagnostics.append(
                f"PPTX chart text reaches {chart_min:g} pt below the {chart_floor:g} pt chart floor."
            )
    return diagnostics


def _inspect_slide(slide_root: ET.Element, label: str, scale: float, radius: dict,
                   body_floor: float | None, caption_floor: float) -> dict:
    """Per-slide statistics: density counts, typography, floor and corner-radius violations."""
    chart_frames = [gf for gf in slide_root.findall('.//p:graphicFrame', NS) if gf.find('.//c:chart', NS) is not None]
    raw_charts = len(slide_root.findall('.//c:chart', NS))
    major_charts = raw_charts
    if raw_charts == 2 and len(chart_frames) == 2:
        boxes = []
        for gf in chart_frames:
            xfrm = gf.find('p:xfrm', NS)
            off = xfrm.find('a:off', NS) if xfrm is not None else None
            ext = xfrm.find('a:ext', NS) if xfrm is not None else None
            if off is not None and ext is not None:
                x_in = int(off.get('x', '0')) / 914400
                w_in = int(ext.get('cx', '0')) / 914400
                boxes.append((x_in, x_in + w_in, w_in))
        if len(boxes) == 2:
            boxes.sort(key=lambda item: item[0])
            if all(0 < item[2] <= 4.8 * scale + 0.02 for item in boxes) and boxes[0][1] <= boxes[1][0] + 0.02 * scale:
                major_charts = 1
    result = {'bullets': 0, 'charts': raw_charts, 'major_charts': major_charts, 'cards': 0, 'supporting': 0,
              'levels': set(), 'sizes': [], 'rounded': 0, 'text_rounded': 0, 'unnamed_rounded': 0,
              'no_font_size': 0, 'max_cjk': 0, 'floor_violations': [], 'radius_violations': [],
              'card_shape_violations': [], 'radius_inconsistency': None}
    bullet_tags = {f'{{{NS["a"]}}}buChar', f'{{{NS["a"]}}}buAutoNum'}
    for paragraph in slide_root.findall('.//a:p', NS):
        properties = paragraph.find('a:pPr', NS)
        if (properties is not None and any(child.tag in bullet_tags for child in properties)
                and ''.join(t.text or '' for t in paragraph.findall('.//a:t', NS)).strip()):
            result['bullets'] += 1
    r_min, r_max = radius.get('min_in'), radius.get('max_in')
    card_radii: list[tuple[str, float]] = []
    for shape in slide_root.iter():
        if shape.tag not in (f'{{{NS["p"]}}}sp', f'{{{NS["p"]}}}pic'):
            continue
        raw_name = _shape_name(shape)
        name = raw_name.upper()
        is_rounded = _shape_geometry(shape) in _ROUNDED_GEOMETRIES
        has_text = bool(_shape_text(shape).strip()) if shape.tag.endswith('}sp') else False
        card_name = bool(_CARD_NAME_RE.search(name))
        exempt = bool(_PILL_NAME_RE.search(name) or _STEP_NAME_RE.search(name)
                      or _SUPPORTING_CONCLUSION_RE.search(name))
        result['supporting'] += bool(_SUPPORTING_CONCLUSION_RE.search(name))
        if has_text:
            result['no_font_size'] += not _shape_has_font_size(shape)
            sizes = _run_sizes(shape)
            caption = bool(_CAPTION_NAME_RE.search(name))
            if sizes:
                result['sizes'] += sizes
                if not caption and not _CHROME_NAME_RE.search(name):
                    result['levels'].update(sizes)
                floor = caption_floor if caption else body_floor
                if isinstance(floor, (int, float)) and min(sizes) < floor * scale - 0.01:
                    result['floor_violations'].append(
                        f'{label} "{raw_name or "<unnamed>"}" uses {min(sizes):g} pt '
                        f'(< {floor * scale:g} pt {"caption" if caption else "body"} floor)')
        if is_rounded:
            result['rounded'] += 1
            result['text_rounded'] += has_text
            result['unnamed_rounded'] += has_text and not (card_name or exempt)
        elif card_name:
            result['card_shape_violations'].append(raw_name or '<unnamed>')
        is_card = card_name or (is_rounded and has_text and not exempt)
        if is_card:
            result['cards'] += 1
            result['max_cjk'] = max(result['max_cjk'], len(_CJK_RE.findall(_body_text(shape))))
        measured = _corner_radius_in(shape) if (name.startswith('HX_') or card_name or exempt) and number_pair(r_min, r_max) else None
        if measured is None:
            continue
        value, _, shape_h = measured
        tolerance = 0.006 * scale
        if _PILL_NAME_RE.search(name):
            expected = min(shape_h / 2, r_max * scale)
            if abs(value - expected) > tolerance:
                result['radius_violations'].append(
                    f'{label} "{raw_name}" pill radius {value:.3f} in, expected {expected:.3f} in')
            continue
        low, high = r_min * scale, r_max * scale
        if not (low - tolerance <= value <= high + tolerance) and value < shape_h / 2 - tolerance:
            result['radius_violations'].append(f'{label} "{raw_name}" radius {value:.3f} in outside {low:.3f}–{high:.3f} in')
        if is_card:
            card_radii.append((raw_name, value))
    if card_radii and max(v for _, v in card_radii) - min(v for _, v in card_radii) > 0.01 * scale:
        result['radius_inconsistency'] = f'{label}: ' + ', '.join(f'{n}={v:.3f}' for n, v in card_radii[:6])
    return result


def _off_slide_objects(slide_root: ET.Element, width: int, height: int) -> bool:
    tree = slide_root.find('p:cSld/p:spTree', NS)
    tolerance = round(914400 * 0.02)
    for shape in tree if tree is not None else []:
        if shape.tag not in {f'{{{NS["p"]}}}{tag}' for tag in ('sp', 'pic', 'graphicFrame', 'cxnSp')}:
            continue
        transform = shape.find('p:spPr/a:xfrm', NS)
        transform = transform if transform is not None else shape.find('p:xfrm', NS)
        if transform is None or transform.get('rot') not in (None, '0'):
            continue
        off, ext = transform.find('a:off', NS), transform.find('a:ext', NS)
        if off is None or ext is None:
            continue
        x, y = int(off.get('x', '0')), int(off.get('y', '0'))
        w, h = int(ext.get('cx', '0')), int(ext.get('cy', '0'))
        if x < -tolerance or y < -tolerance or x + w > width + tolerance or y + h > height + tolerance:
            return True
    return False


def inspect_pptx(path: Path, check_bounds: bool = False, density_limits: dict | None = None,
                 radius_limits: dict | None = None, markers: tuple[str, ...] = (),
                 runtime_width_in: float = 10.0) -> tuple[dict, list[str]]:
    """Structural inspection of a PPTX package. Limits come from the theme via ``deck_limits``."""
    warnings: list[str] = []
    with ZipFile(path) as archive:
        infos = archive.infolist()
        names = archive.namelist()
        entries = set(names)
        if len(entries) != len(names):
            raise ValueError('PPTX has duplicate ZIP entries.')
        if sum(i.file_size for i in infos) > 300_000_000:
            raise ValueError('PPTX expands above the 300 MB diagnostic limit.')
        for name in names:
            parts = PurePosixPath(name)
            if parts.is_absolute() or '..' in parts.parts or '\\' in name:
                raise ValueError(f'Unsafe ZIP entry: {name}')
            if Path(name).suffix.lower() in FONT_SUFFIXES or name.startswith('ppt/fonts/'):
                raise ValueError(f'Embedded font detected: {name}')
        damaged = archive.testzip()
        if damaged:
            raise ValueError(f'PPTX CRC error: {damaged}')
        for required in ('[Content_Types].xml', '_rels/.rels', 'ppt/presentation.xml'):
            if required not in entries:
                raise ValueError(f'Missing PPTX part: {required}')
        xml = {name: ET.fromstring(archive.read(name)) for name in names if name.endswith(('.xml', '.rels'))}
    external_relationships = 0
    for name, root in xml.items():
        if not name.endswith('.rels'):
            continue
        base = '' if name == '_rels/.rels' else posixpath.dirname(posixpath.dirname(name))
        for relationship in root.findall('r:Relationship', NS):
            target = relationship.get('Target', '')
            if relationship.get('TargetMode') == 'External':
                external_relationships += 1
                continue
            target = unquote(target.split('#', 1)[0])
            resolved = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(base, target))
            if resolved not in entries:
                raise ValueError(f'Broken relationship in {name}: {target}')
    size = xml['ppt/presentation.xml'].find('p:sldSz', NS)
    if size is None:
        raise ValueError('Missing presentation dimensions.')
    width, height = int(size.get('cx', '0')), int(size.get('cy', '0'))
    if width <= 0 or height <= 0:
        raise ValueError('Invalid presentation dimensions.')
    scale = (width / 914400) / runtime_width_in
    # Slide order follows the presentation slide list, not part names.
    presentation_rels = xml.get('ppt/_rels/presentation.xml.rels')
    rel_targets = {rel.get('Id'): posixpath.normpath(posixpath.join('ppt', rel.get('Target', '')))
                   for rel in (presentation_rels.findall('r:Relationship', NS) if presentation_rels is not None else [])}
    references = xml['ppt/presentation.xml'].findall('p:sldIdLst/p:sldId', NS)
    slide_names = [rel_targets.get(ref.get(f'{{{NS_R}}}id')) for ref in references]
    slide_parts = [name for name in names if re.fullmatch(r'ppt/slides/slide\d+\.xml', name)]
    if len(references) != len(slide_parts) or any(name not in entries for name in slide_names):
        raise ValueError('Slide-part count does not match the presentation slide list.')

    limits = density_limits or {}
    slides = [_inspect_slide(xml[name], f'slide {number}', scale, radius_limits or {},
                             limits.get('body_min_runtime_pt'), limits.get('caption_min_runtime_pt', 8))
              for number, name in enumerate(slide_names, start=1)]
    all_sizes = [size for slide in slides for size in slide['sizes']]
    chart_parts = [root for name, root in xml.items() if re.fullmatch(r'ppt/charts/chart\d+\.xml', name)]
    chart_sizes = [size for root in chart_parts for size in _font_sizes_from_xml(root)]
    donut_categories = [len(root.findall('.//c:cat//c:pt', NS)) for root in chart_parts
                        if root.find('.//c:doughnutChart', NS) is not None]
    per = lambda key: [slide[key] for slide in slides]  # noqa: E731
    stats = {
        'slides': len(slide_names), 'width_emu': width, 'height_emu': height,
        'width_in': round(width / 914400, 6), 'height_in': round(height / 914400, 6),
        'canvas_scale': round(scale, 6), 'is_16_9': abs(width / height - 16 / 9) < 0.002,
        'layouts': sum(bool(re.fullmatch(r'ppt/slideLayouts/slideLayout\d+\.xml', n)) for n in names),
        'media_files': sum(n.startswith('ppt/media/') and not n.endswith('/') for n in names),
        'charts': len(chart_parts),
        'min_explicit_font_size_pt': min(all_sizes) if all_sizes else None,
        'typography_sizes_pt': sorted(set(all_sizes)),
        'typography_levels_per_slide': [len(levels) for levels in per('levels')],
        'max_typography_levels_per_slide': max((len(levels) for levels in per('levels')), default=0),
        'min_chart_font_size_pt': min(chart_sizes) if chart_sizes else None,
        'bullet_paragraphs_per_slide': per('bullets'),
        'named_cards_per_slide': per('cards'),
        'named_card_count': sum(per('cards')),
        'charts_per_slide': per('charts'),
        'major_charts_per_slide': per('major_charts'),
        'supporting_conclusion_counts': per('supporting'),
        'max_supporting_conclusions_per_slide': max(per('supporting'), default=0),
        'max_cjk_chars_per_named_card': max(per('max_cjk'), default=0),
        'rounded_shape_count': sum(per('rounded')),
        'text_bearing_rounded_shape_count': sum(per('text_rounded')),
        'unnamed_text_rounded_shape_count': sum(per('unnamed_rounded')),
        'text_shapes_without_font_size': sum(per('no_font_size')),
        'card_shape_violations': [item for slide in slides for item in slide['card_shape_violations']],
        'radius_violations': [item for slide in slides for item in slide['radius_violations']],
        'radius_inconsistencies': [item for item in per('radius_inconsistency') if item],
        'body_font_floor_violations': [item for slide in slides for item in slide['floor_violations']],
        'max_donut_categories': max(donut_categories, default=0),
        'external_relationships': external_relationships,
        'xml_parts_checked': len(xml), 'visual_check': 'NOT_PERFORMED',
    }
    if density_limits:
        warnings.extend(validate_pptx_density(stats, density_limits))
    if not stats['is_16_9']:
        warnings.append('Canvas is not 16:9. Confirm the user explicitly requested another ratio.')
    if check_bounds:
        for number, name in enumerate(slide_names, start=1):
            if _off_slide_objects(xml[name], width, height):
                warnings.append(f'slide {number}: a top-level object extends off slide; inspect its intent visually.')
            slide_texts = [t.text or '' for t in xml[name].findall('.//a:t', NS)]
            slide_rels_name = posixpath.join(posixpath.dirname(name), '_rels', f'{posixpath.basename(name)}.rels')
            slide_rels = xml.get(slide_rels_name)
            if slide_rels is not None:
                base_dir = posixpath.dirname(name)
                for rel in slide_rels.findall('r:Relationship', NS):
                    if rel.get('TargetMode') == 'External':
                        continue
                    rel_target = unquote(rel.get('Target', '').split('#', 1)[0])
                    resolved_part = (rel_target.lstrip('/') if rel_target.startswith('/')
                                     else posixpath.normpath(posixpath.join(base_dir, rel_target)))
                    if resolved_part.startswith('ppt/charts/chart') and resolved_part in xml:
                        chart_xml = xml[resolved_part]
                        slide_texts += [t.text or '' for t in chart_xml.findall('.//a:t', NS)]
                        slide_texts += [v.text or '' for v in chart_xml.findall('.//c:tx//c:v', NS)]
            content = ' '.join(slide_texts).casefold()
            warnings += [f'slide {number}: possible unfinished template text: {m}' for m in markers
                         if m.casefold() in content]
    return stats, warnings


def number_pair(first: object, second: object) -> bool:
    return all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in (first, second))


def deck_limits(theme_config: dict | None) -> dict:
    """Keyword arguments for ``inspect_pptx`` derived from the theme contract (the only source of limits)."""
    if not isinstance(theme_config, dict):
        return {}
    floors = theme_config['typography']['floor_pt']
    density = dict(theme_config['density'],
                   max_donut_categories=theme_config['charts']['max_donut_categories'],
                   body_min_runtime_pt=floors['body'], chart_min_runtime_pt=floors['chart'],
                   caption_min_runtime_pt=floors['caption'])
    return {'density_limits': density, 'radius_limits': theme_config['components']['corner_radius'],
            'markers': tuple(theme_config['qa']['template_markers']),
            'runtime_width_in': theme_config['canvas']['width']}


def inspect_deck(target: Path, theme_config: dict | None) -> tuple[dict, list[str], list[str]]:
    """Inspect a finished deck. Returns (stats, warnings, errors)."""
    limits = deck_limits(theme_config)
    stats, warnings = inspect_pptx(target, check_bounds=True, **limits)
    errors: list[str] = []
    caption_floor = limits.get('density_limits', {}).get('caption_min_runtime_pt', 8)
    minimum_font = stats.get('min_explicit_font_size_pt')
    if minimum_font is not None and minimum_font < caption_floor * stats.get('canvas_scale', 1.0) - 0.01:
        errors.append(f'Generated PPTX contains explicit text below the {caption_floor:g} pt absolute minimum.')
    return stats, warnings, errors


def validate_theme_config(root: Path) -> tuple[dict, list[str]]:
    """Validate the theme with JSON Schema, then check cross-field consistency.

    Values themselves (limits, sizes, colours) live only in the theme file. This
    function checks that they are coherent with each other — it deliberately
    does not re-hardcode them, which would create a second source of truth.
    """
    errors: list[str] = []
    path = root / 'scripts/office/huaxin_theme_config.json'
    schema_path = root / 'references/theme-config.schema.json'
    try:
        config = json.loads(path.read_text(encoding='utf-8'))
        schema = json.loads(schema_path.read_text(encoding='utf-8'))
    except (OSError, ValueError) as exc:
        return {}, [f'Theme config or schema could not be read: {exc}']
    try:
        from jsonschema import Draft202012Validator, SchemaError
    except ImportError:
        return config, ['Theme validation requires jsonschema. Install requirements.txt (jsonschema>=4.23,<5).']
    try:
        Draft202012Validator.check_schema(schema)
    except SchemaError as exc:
        return config, [f'Theme schema is invalid JSON Schema: {exc.message}']
    schema_errors = sorted(Draft202012Validator(schema).iter_errors(config), key=lambda item: list(map(str, item.path)))
    for item in schema_errors[:50]:
        location = ''.join(f'[{p}]' if isinstance(p, int) else f'.{p}' for p in item.absolute_path)
        errors.append(f'JSON Schema: ${location} {item.message}')
    if schema_errors:
        return config, errors  # cross-field checks assume a schema-valid document

    tolerance = 0.001
    canvas = config['canvas']
    width, height = canvas['width'], canvas['height']
    if abs(width / height - 16 / 9) > 0.002:
        errors.append('Theme canvas must be 16:9.')
    radius = config['components']['corner_radius']
    if not radius['min_in'] <= radius['default_in'] <= radius['max_in']:
        errors.append('Theme corner radius must satisfy min_in <= default_in <= max_in.')
    table = config['recipes']['table']
    if table['row_h_min'] > table['row_h_max']:
        errors.append('Theme recipes.table.row_h_min must not exceed row_h_max.')
    scale = config['spacing']['scale']
    for label, value in (('grid.gap', config['grid']['gap']), ('spacing.card_padding', config['spacing']['card_padding'])):
        if not any(abs(value - step) < tolerance for step in scale):
            errors.append(f'Theme {label} ({value}) must be one of spacing.scale {scale}.')

    rects = {f'slides.{page}.{key}': value for page, spec in config['slides'].items()
             for key, value in spec.items() if isinstance(value, dict)}
    rects['footer'] = config['footer']
    for label, rect in rects.items():
        if rect['x'] + rect['w'] > width + tolerance or rect['y'] + rect['h'] > height + tolerance:
            errors.append(f'Theme {label} exceeds the canvas ({width}×{height}).')

    safe, body, footer = config['safe_area'], config['slides']['body'], config['footer']
    content, title = body['content_area'], body['title']
    if abs(content['x'] - safe['left']) > tolerance or abs(content['x'] + content['w'] - (width - safe['right'])) > tolerance:
        errors.append('Theme slides.body.content_area must bind to the horizontal safe_area.')
    if abs(content['y'] - safe['top']) > tolerance:
        errors.append(f"Theme slides.body.content_area.y must equal safe_area.top ({safe['top']:.3f}).")
    if content['y'] + tolerance < title['y'] + title['h'] + config['spacing']['title_to_content']:
        errors.append('Theme content starts before title_to_content spacing is satisfied.')
    if abs(footer['y'] - (height - safe['footer_top_clearance'])) > tolerance:
        errors.append(f"Theme footer.y must equal canvas.height - safe_area.footer_top_clearance "
                      f"({height - safe['footer_top_clearance']:.3f}).")
    if content['y'] + content['h'] > footer['y'] - config['spacing']['content_to_footer'] + tolerance:
        errors.append('Theme content area collides with the footer spacing.')

    sizes, floors = config['typography']['sizes_pt'], config['typography']['floor_pt']
    for key, floor in (('body', 'body'), ('card_title', 'body'), ('chart', 'chart'), ('caption', 'caption'),
                       ('label', 'caption')):
        if sizes[key] < floors[floor]:
            errors.append(f'Theme typography.sizes_pt.{key} ({sizes[key]}) is below floor_pt.{floor} ({floors[floor]}).')
    if not sizes['long_title'] < sizes['slide_title']:
        errors.append('Theme long_title must be smaller than slide_title.')

    colors = config['colors']
    for label, value in [('footer.color', footer['color'])] + [
            (f'charts.{key}', ref) for key in ('series_colors', 'pie_colors') for ref in config['charts'][key]]:
        if not re.fullmatch(r'[0-9A-Fa-f]{6}', value) and value not in colors:
            errors.append(f'Theme {label} references unknown colour token {value!r}.')

    for page, spec in config['slides'].items():
        if not (root / spec['bg_image']).is_file():
            errors.append(f"Theme slides.{page}.bg_image does not exist: {spec['bg_image']}.")
    if config['slides']['body']['bg_image'] != 'assets/body-background.jpeg' or any(
            config['slides'][page]['bg_image'] != 'assets/cover-background.jpeg' for page in ('cover', 'closing')):
        errors.append('Theme backgrounds must use the canonical cover/body asset names.')
    return config, errors


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--skill-dir', type=Path, default=ROOT, help='Skill directory to inspect')
    parser.add_argument('--pptx', type=Path, help='Optional generated PPTX to inspect')
    parser.add_argument('--visual-evidence', type=Path,
                        help='Optional render_visual.py JSON evidence manifest')
    parser.add_argument('--json', dest='json_output', type=Path, help='Save a JSON report outside the inspected skill')
    args = parser.parse_args()
    root = args.skill_dir.expanduser().resolve()
    try:
        import yaml
    except ImportError:
        print('Missing PyYAML. Install requirements.txt or run: python -m pip install PyYAML', file=sys.stderr)
        return 2

    # Reject duplicate keys rather than silently accepting a later YAML value.
    class UniqueLoader(yaml.SafeLoader):
        pass

    def mapping(loader, node, deep=False):
        result = {}
        for key_node, value_node in node.value:
            key = loader.construct_object(key_node, deep=deep)
            if key in result:
                raise ValueError(f'Duplicate YAML key: {key}')
            result[key] = loader.construct_object(value_node, deep=deep)
        return result
    UniqueLoader.add_constructor(yaml.resolver.BaseResolver.DEFAULT_MAPPING_TAG, mapping)

    report: dict = {'skill_directory': str(root), 'status': 'FAIL', 'errors': [], 'warnings': [],
                   'checks': {}, 'limitations': [
                       'Structural and asset checks only; no visual inspection, factual validation or text-fit proof.',
                       'Density limits are checked for theme tokens and, when a PPTX is supplied, semantic bullet/card/CJK/chart/typography statistics; text fit, unnamed components and visual hierarchy still require generation-time or visual review.',
                       'No claim of an actual PowerPoint/WPS test or successful host installation.'
                   ]}
    errors, warnings, checks = report['errors'], report['warnings'], report['checks']
    visual_evidence: dict | None = None
    matched_visual_deck: dict | None = None
    visual_errors: list[str] = []
    skill_version: str | None = None

    def require(condition: bool, message: str) -> None:
        if not condition:
            errors.append(message)

    required = ['SKILL.md', 'README.md', 'LICENSE.txt', 'agents/openai.yaml', 'requirements.txt', 'editing.md',
                'references/style-guide.md', 'references/quality-checklist.md', 'references/layouts.md',
                'references/template-map.md', 'references/source-notes.md', 'references/assets-manifest.json',
                'assets/cover-background.jpeg', 'assets/body-background.jpeg',
                'assets/huaxin-template-16x9.pptx', 'scripts/create_starter.py', 'scripts/hx_pptx.py',
                'scripts/check_deck.py', 'scripts/validate_skill.py', 'tests/acceptance-cases.md',
                'tests/package-integrity.py', 'tests/validator-regression.py',
                'tests/density-regression.py', 'tests/builder-regression.py',
                'tests/visual-evidence-regression.py', 'tests/validation-report.md',
                'references/theme-config.schema.json', 'scripts/office/huaxin_theme_config.json',
                'scripts/office/render_visual.py']
    for name in required:
        require((root / name).is_file(), f'Missing required file: {name}')
    theme_config, theme_errors = validate_theme_config(root)
    errors.extend(theme_errors)
    if theme_config and not theme_errors:
        # Token drift guard: every theme colour and the font must be documented in the style guide.
        guide_path = root / 'references/style-guide.md'
        guide_text = guide_path.read_text(encoding='utf-8').upper() if guide_path.is_file() else ''
        for key, value in theme_config['colors'].items():
            require(f'#{value.upper()}' in guide_text,
                    f'Theme color {key}=#{value} is not documented in references/style-guide.md (token drift).')
        require(theme_config['typography']['font'].upper() in guide_text,
                'Theme typography.font is not documented in references/style-guide.md.')
        checks['theme'] = {
            'path': str(root / 'scripts/office/huaxin_theme_config.json'),
            'schema_version': theme_config['schema_version'],
            'canvas': theme_config['canvas'],
            'typography': theme_config['typography'],
            'density': theme_config['density'],
            'corner_radius': theme_config['components']['corner_radius'],
        }
    if args.visual_evidence:
        visual_path = args.visual_evidence.expanduser().resolve()
        visual_evidence, visual_errors, matched_visual_deck = validate_visual_evidence(
            visual_path,
            args.pptx.expanduser().resolve() if args.pptx else None,
        )
        errors.extend(visual_errors)
        if visual_evidence:
            checks['visual_evidence'] = {
                'path': str(visual_path),
                'schema_version': visual_evidence.get('schema_version'),
                'status': visual_evidence.get('status'),
                'render_status': visual_evidence.get('render_status'),
                'renderer': visual_evidence.get('renderer'),
                'raster': visual_evidence.get('raster'),
                'manual_review': visual_evidence.get('manual_review'),
                'decks': visual_evidence.get('decks'),
            }
            if not visual_errors:
                report['limitations'][0] = (
                    'LibreOffice visual evidence is attached for the recorded decks; '
                    'visual quality remains bounded by the explicitly recorded human page review.'
                )
    try:
        skill_text = (root / 'SKILL.md').read_text(encoding='utf-8')
        matched = re.match(r'\A---\s*\n(.*?)\n---\s*\n', skill_text, re.S)
        if matched is None:
            raise ValueError('SKILL.md must start with YAML frontmatter.')
        metadata = yaml.load(matched.group(1), Loader=UniqueLoader)
        if not isinstance(metadata, dict):
            raise ValueError('Frontmatter must be a YAML mapping.')
        name, description = metadata.get('name'), metadata.get('description')
        require(isinstance(name, str) and len(name) <= 64 and bool(re.fullmatch(r'[a-z0-9]+(?:-[a-z0-9]+)*', name)), 'Invalid skill name.')
        require(name == root.name, 'Skill name must match its directory name.')
        require(isinstance(description, str) and 0 < len(description) <= 1024, 'description must be a non-empty string of at most 1024 characters.')
        allowed = {'name', 'description', 'license', 'compatibility', 'metadata', 'allowed-tools'}
        require(not set(metadata) - allowed, 'Unknown frontmatter field(s).')
        if 'compatibility' in metadata:
            value = metadata['compatibility']
            require(isinstance(value, str) and 0 < len(value) <= 500, 'compatibility must be 1–500 characters.')
        if 'metadata' in metadata:
            value = metadata['metadata']
            require(isinstance(value, dict) and all(isinstance(k, str) and isinstance(v, str) for k, v in value.items()), 'metadata must map string keys to string values.')
            if isinstance(value, dict) and isinstance(value.get('version'), str):
                skill_version = value['version']
        require(len(skill_text.splitlines()) < 500, 'SKILL.md exceeds the recommended 500-line limit.')
        checks['frontmatter'] = {'name': name, 'description_chars': len(description) if isinstance(description, str) else 0,
                                 'skill_lines': len(skill_text.splitlines())}
        links = re.findall(r'\]\(([^)]+)\)', skill_text)
        for link in links:
            if '://' in link or link.startswith('#'):
                continue
            relative = link.split('#', 1)[0]
            resolved = (root / relative).resolve()
            require(root in resolved.parents, f'Resource link escapes the skill directory: {relative}')
            require(resolved.is_file(), f'Broken resource link: {relative}')
        checks['relative_links'] = len(links)
        agent = yaml.load((root / 'agents/openai.yaml').read_text(encoding='utf-8'), Loader=UniqueLoader)
        require(isinstance(agent, dict), 'Agent configuration must be a mapping.')
        interface = agent.get('interface', {})
        for field in ('display_name', 'short_description', 'default_prompt'):
            require(isinstance(interface.get(field), str) and bool(interface.get(field)), f'Missing agent interface field: {field}')
        require('$huaxin-pptx' in interface.get('default_prompt', ''), 'Agent default prompt references the wrong skill.')
        checks['agent_metadata'] = True
    except (OSError, ValueError, TypeError, AttributeError, yaml.YAMLError) as exc:
        errors.append(f'Metadata/reference validation failed: {exc}')

    for path in sorted(root.rglob('*')):
        if path.is_symlink():
            errors.append(f'Symlink is not expected in this distributable package: {path.relative_to(root)}')
        if path.is_file() and path.suffix.lower() in FONT_SUFFIXES:
            errors.append(f'Font file is not allowed in this package: {path.relative_to(root)}')
        if path.is_file() and path.suffix == '.py':
            try:
                ast.parse(path.read_text(encoding='utf-8'), filename=str(path))
            except (OSError, SyntaxError) as exc:
                errors.append(f'Invalid Python script {path.name}: {exc}')

    try:
        manifest = json.loads((root / 'references/assets-manifest.json').read_text(encoding='utf-8'))
        require(manifest['skill_name'] == 'huaxin-pptx', 'Asset manifest references the wrong skill.')
        if skill_version is not None:
            require(manifest.get('version') == skill_version,
                    'Asset manifest version must match SKILL.md metadata.version.')
        checks['assets'] = []
        for asset in manifest['assets']:
            path = (root / asset['path']).resolve()
            if root not in path.parents:
                raise ValueError('Asset path escapes skill directory.')
            data = path.read_bytes()
            require(len(data) == asset['bytes'], f'Asset size mismatch: {asset["path"]}')
            require(hashlib.sha256(data).hexdigest() == asset['sha256'], f'Asset hash mismatch: {asset["path"]}')
            if path.suffix.lower() == '.jpeg':
                dimensions = jpeg_size(data)
                require(dimensions == (asset['width_px'], asset['height_px']), f'JPEG dimensions mismatch: {asset["path"]}')
            checks['assets'].append(asset['path'])
        template = root / manifest['template']['path']
        stats, template_warnings = inspect_pptx(template)
        for field in ('slides', 'layouts', 'media_files', 'width_emu', 'height_emu'):
            require(stats[field] == manifest['template'][field], f'Template {field} mismatch.')
        require(stats['is_16_9'], 'Original template is expected to be 16:9.')
        with ZipFile(template) as archive:
            for asset in manifest['assets']:
                if 'source_part' in asset:
                    require(archive.read(asset['source_part']) == (root / asset['path']).read_bytes(), f'Background differs from original template: {asset["path"]}')
        if visual_evidence and not visual_errors:
            template_path = template.resolve()
            for deck in visual_evidence.get('decks', []):
                if isinstance(deck, dict) and isinstance(deck.get('pptx'), str):
                    try:
                        if Path(deck['pptx']).resolve() == template_path:
                            stats['visual_check'] = 'LIBREOFFICE_PASS'
                    except OSError:
                        pass
        checks['original_template'] = stats
        warnings.extend(template_warnings)
    except (OSError, KeyError, TypeError, ValueError, BadZipFile, ET.ParseError) as exc:
        errors.append(f'Asset/template validation failed: {exc}')

    if args.pptx:
        try:
            target = args.pptx.expanduser().resolve()
            stats, deck_warnings, deck_errors = inspect_deck(target, theme_config)
            if matched_visual_deck is not None and not visual_errors:
                stats['visual_check'] = 'LIBREOFFICE_PASS'
            checks['generated_pptx'] = {'path': str(target), **stats}
            errors.extend(deck_errors)
            warnings.extend(deck_warnings)
        except (OSError, ValueError, BadZipFile, ET.ParseError) as exc:
            errors.append(f'Generated PPTX validation failed: {exc}')

    report['status'] = 'PASS' if not errors else 'FAIL'
    report['warnings'] = sorted(set(warnings))
    if args.json_output:
        try:
            destination = args.json_output.expanduser().resolve()
            if destination == root or root in destination.parents or destination == ROOT or ROOT in destination.parents:
                raise ValueError('Write reports outside skill directories.')
            if destination.suffix.lower() != '.json':
                raise ValueError('JSON report path must end in .json.')
            if args.pptx and destination == args.pptx.expanduser().resolve():
                raise ValueError('A report cannot overwrite the inspected PPTX.')
            destination.parent.mkdir(parents=True, exist_ok=True)
            destination.write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
        except (OSError, ValueError) as exc:
            errors.append(f'Could not write JSON report: {exc}')
            report['status'] = 'FAIL'
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report['status'] == 'PASS' else 1


if __name__ == '__main__':
    raise SystemExit(main())
