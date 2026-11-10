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


def _shape_name(shape: ET.Element) -> str:
    name_element = shape.find('p:nvSpPr/p:cNvPr', NS)
    return (name_element.get('name', '') if name_element is not None else '').strip()


def _shape_geometry(shape: ET.Element) -> str | None:
    geometry = shape.find('p:spPr/a:prstGeom', NS)
    return geometry.get('prst') if geometry is not None else None


def _shape_text(shape: ET.Element) -> str:
    return ''.join(text.text or '' for text in shape.findall('.//a:t', NS))


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


def validate_pptx_density(stats: dict, density: dict) -> list[str]:
    """Return deterministic density diagnostics for semantic PPTX objects.

    Card limits apply to shapes named ``CARD``, ``TILE`` or ``PANEL`` plus
    text-bearing rounded rectangles.  This catches unnamed card components
    while leaving text-free decorative shapes outside the limit. Supporting conclusions use the explicit semantic
    name ``SUPPORTING_CONCLUSION`` and are capped per slide. Bullet, chart and
    typography limits are counted from the Office XML and therefore remain
    applicable to generated decks.
    """
    diagnostics: list[str] = []

    def positive_int(key: str) -> int | None:
        value = density.get(key)
        return value if isinstance(value, int) and not isinstance(value, bool) and value >= 0 else None

    bullets_limit = positive_int('max_top_level_bullets')
    if bullets_limit is not None and isinstance(stats.get('bullet_paragraphs'), int):
        if stats['bullet_paragraphs'] > bullets_limit:
            diagnostics.append(
                f"PPTX has {stats['bullet_paragraphs']} bullet paragraphs; "
                f"density.max_top_level_bullets is {bullets_limit}."
            )
    cards_limit = positive_int('max_major_cards')
    if cards_limit is not None and isinstance(stats.get('named_card_count'), int):
        if stats['named_card_count'] > cards_limit:
            diagnostics.append(
                f"PPTX has {stats['named_card_count']} named card/panel shapes; "
                f"density.max_major_cards is {cards_limit}."
            )
    cjk_limit = positive_int('max_cjk_chars_per_card')
    if cjk_limit is not None and isinstance(stats.get('max_cjk_chars_per_named_card'), int):
        if stats['max_cjk_chars_per_named_card'] > cjk_limit:
            diagnostics.append(
                f"A named card contains {stats['max_cjk_chars_per_named_card']} CJK characters; "
                f"density.max_cjk_chars_per_card is {cjk_limit}."
            )
    charts_limit = positive_int('max_major_charts')
    if charts_limit is not None and isinstance(stats.get('charts'), int) and stats['charts'] > charts_limit:
        diagnostics.append(
            f"PPTX has {stats['charts']} charts; density.max_major_charts is {charts_limit}."
        )
    unnamed_rounded = stats.get('unnamed_text_rounded_shape_count')
    if isinstance(unnamed_rounded, int) and unnamed_rounded:
        diagnostics.append(
            f"PPTX has {unnamed_rounded} text-bearing rounded component(s) without a "
            "semantic CARD/TILE/PANEL name; name them so density and radius checks apply."
        )
    invalid_card_shapes = stats.get('card_shape_violations')
    if isinstance(invalid_card_shapes, list) and invalid_card_shapes:
        diagnostics.append(
            "Semantic card component(s) do not use a rounded rectangle geometry: "
            + ', '.join(str(item) for item in invalid_card_shapes[:8])
            + (" …" if len(invalid_card_shapes) > 8 else '')
        )
    inheritance_gaps = stats.get('text_shapes_without_font_size')
    if isinstance(inheritance_gaps, int) and inheritance_gaps:
        diagnostics.append(
            f"PPTX has {inheritance_gaps} text-bearing shape(s) without explicit or "
            "inherited font size; typography may change with the active theme."
        )
    supporting_limit = positive_int('max_supporting_conclusions')
    supporting_max = stats.get('max_supporting_conclusions_per_slide')
    if (supporting_limit is not None and isinstance(supporting_max, int)
            and supporting_max > supporting_limit):
        diagnostics.append(
            f"A slide has {supporting_max} named supporting conclusion shapes; "
            f"density.max_supporting_conclusions is {supporting_limit}."
        )
    donut_limit = density.get('max_donut_categories')
    if isinstance(donut_limit, int) and not isinstance(donut_limit, bool):
        max_donut = stats.get('max_donut_categories')
        if isinstance(max_donut, int) and max_donut > donut_limit:
            diagnostics.append(
                f"A donut chart has {max_donut} categories; charts.max_donut_categories is {donut_limit}."
            )
    levels_limit = positive_int('max_typography_levels')
    slide_levels = stats.get('max_typography_levels_per_slide')
    if levels_limit is not None and isinstance(slide_levels, int):
        if slide_levels > levels_limit:
            diagnostics.append(
                f"A slide uses {slide_levels} explicit typography levels; "
                f"density.max_typography_levels is {levels_limit}."
            )
    body_floor = density.get('body_min_runtime_pt')
    if isinstance(body_floor, (int, float)) and not isinstance(body_floor, bool):
        body_min = stats.get('min_explicit_font_size_pt')
        if isinstance(body_min, (int, float)) and body_min < body_floor:
            diagnostics.append(
                f"PPTX has explicit text at {body_min:g} pt below the {body_floor:g} pt "
                'body floor; confirm it is caption/footer text.'
            )
    chart_floor = density.get('chart_min_runtime_pt')
    chart_min = stats.get('min_chart_font_size_pt')
    if isinstance(chart_floor, (int, float)) and not isinstance(chart_floor, bool):
        if isinstance(chart_min, (int, float)) and chart_min < chart_floor:
            diagnostics.append(
                f"PPTX chart text reaches {chart_min:g} pt below the {chart_floor:g} pt chart floor."
            )
    return diagnostics


def inspect_pptx(path: Path, check_bounds: bool = False,
                 density_limits: dict | None = None) -> tuple[dict, list[str]]:
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
        xml = {name: ET.fromstring(archive.read(name)) for name in names
               if name.endswith(('.xml', '.rels'))}
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
        slide_names = sorted(name for name in names if re.fullmatch(r'ppt/slides/slide\d+\.xml', name))
        references = xml['ppt/presentation.xml'].findall('p:sldIdLst/p:sldId', NS)
        if len(references) != len(slide_names):
            raise ValueError('Slide-part count does not match the presentation slide list.')
        explicit_font_sizes_pt: list[float] = []
        all_typography_sizes_pt: list[float] = []
        bullet_paragraphs = 0
        named_card_count = 0
        max_cjk_chars_per_named_card = 0
        supporting_conclusion_counts: list[int] = []
        typography_levels_per_slide: list[int] = []
        rounded_shape_count = 0
        text_bearing_rounded_shape_count = 0
        unnamed_text_rounded_shape_count = 0
        card_shape_violations: list[str] = []
        text_shapes_without_font_size = 0
        for name in slide_names:
            slide_root = xml[name]
            slide_font_sizes = _font_sizes_from_xml(slide_root)
            explicit_font_sizes_pt.extend(slide_font_sizes)
            all_typography_sizes_pt.extend(slide_font_sizes)
            typography_levels_per_slide.append(len(set(slide_font_sizes)))
            supporting_conclusion_count = 0
            for paragraph in slide_root.findall('.//a:p', NS):
                paragraph_properties = paragraph.find('a:pPr', NS)
                if paragraph_properties is not None and any(
                    child.tag in {f'{{{NS["a"]}}}buChar', f'{{{NS["a"]}}}buAutoNum'}
                    for child in list(paragraph_properties)
                ):
                    bullet_paragraphs += 1
            for shape in slide_root.findall('.//p:sp', NS):
                shape_name_raw = _shape_name(shape)
                shape_name = shape_name_raw.upper()
                geometry = _shape_geometry(shape)
                card_name = bool(_CARD_NAME_RE.search(shape_name))
                shape_text = _shape_text(shape)
                if shape_text and not _shape_has_font_size(shape):
                    text_shapes_without_font_size += 1
                is_rounded = geometry in _ROUNDED_GEOMETRIES
                if is_rounded:
                    rounded_shape_count += 1
                    if shape_text:
                        text_bearing_rounded_shape_count += 1
                        if not card_name:
                            unnamed_text_rounded_shape_count += 1
                if card_name and geometry not in _ROUNDED_GEOMETRIES:
                    card_shape_violations.append(shape_name_raw or '<unnamed>')
                if _SUPPORTING_CONCLUSION_RE.search(shape_name):
                    supporting_conclusion_count += 1
                # A text-bearing rounded shape is an implicit card even when a
                # producer forgot the semantic name.  Counting it closes the
                # old loophole while decorative, text-free rounded shapes stay
                # outside density limits.
                if not card_name and not (is_rounded and shape_text):
                    continue
                named_card_count += 1
                max_cjk_chars_per_named_card = max(
                    max_cjk_chars_per_named_card, len(_CJK_RE.findall(shape_text))
                )
            supporting_conclusion_counts.append(supporting_conclusion_count)
        chart_category_counts: list[int] = []
        donut_category_counts: list[int] = []
        chart_font_sizes_pt: list[float] = []
        for name, chart_root in xml.items():
            if not re.fullmatch(r'ppt/charts/chart\d+\.xml', name):
                continue
            chart_font_sizes_pt.extend(_font_sizes_from_xml(chart_root))
            categories = chart_root.findall('.//c:cat//c:pt', NS)
            category_count = len(categories)
            chart_category_counts.append(category_count)
            if chart_root.find('.//c:doughnutChart', NS) is not None:
                donut_category_counts.append(category_count)
        stats = {
            'slides': len(slide_names), 'width_emu': width, 'height_emu': height,
            'width_in': round(width / 914400, 6), 'height_in': round(height / 914400, 6),
            'is_16_9': abs(width / height - 16 / 9) < 0.002,
            'layouts': sum(bool(re.fullmatch(r'ppt/slideLayouts/slideLayout\d+\.xml', n)) for n in names),
            'media_files': sum(n.startswith('ppt/media/') and not n.endswith('/') for n in names),
            'charts': sum(bool(re.fullmatch(r'ppt/charts/chart\d+\.xml', n)) for n in names),
            'text_runs': sum(len(xml[n].findall('.//a:t', NS)) for n in slide_names),
            'explicit_font_sizes_pt': sorted(set(explicit_font_sizes_pt)),
            'min_explicit_font_size_pt': min(explicit_font_sizes_pt) if explicit_font_sizes_pt else None,
            'typography_levels': len(set(all_typography_sizes_pt)),
            'typography_levels_per_slide': typography_levels_per_slide,
            'max_typography_levels_per_slide': max(typography_levels_per_slide, default=0),
            'typography_sizes_pt': sorted(set(all_typography_sizes_pt)),
            'chart_font_sizes_pt': sorted(set(chart_font_sizes_pt)),
            'min_chart_font_size_pt': min(chart_font_sizes_pt) if chart_font_sizes_pt else None,
            'bullet_paragraphs': bullet_paragraphs,
            'named_card_count': named_card_count,
            'semantic_card_count': named_card_count - unnamed_text_rounded_shape_count,
            'rounded_shape_count': rounded_shape_count,
            'text_bearing_rounded_shape_count': text_bearing_rounded_shape_count,
            'unnamed_text_rounded_shape_count': unnamed_text_rounded_shape_count,
            'card_shape_violations': card_shape_violations,
            'text_shapes_without_font_size': text_shapes_without_font_size,
            'max_cjk_chars_per_named_card': max_cjk_chars_per_named_card,
            'supporting_conclusion_counts': supporting_conclusion_counts,
            'supporting_conclusion_count': sum(supporting_conclusion_counts),
            'max_supporting_conclusions_per_slide': max(supporting_conclusion_counts, default=0),
            'chart_category_counts': chart_category_counts,
            'max_donut_categories': max(donut_category_counts, default=0),
            'external_relationships': external_relationships,
            'xml_parts_checked': len(xml), 'visual_check': 'NOT_PERFORMED'
        }
        if density_limits:
            warnings.extend(validate_pptx_density(stats, density_limits))
        if not stats['is_16_9']:
            warnings.append('Canvas is not 16:9. Confirm the user explicitly requested another ratio.')
        if check_bounds:
            if stats['min_explicit_font_size_pt'] is not None and stats['min_explicit_font_size_pt'] < 8:
                warnings.append('Generated PPTX contains explicit text smaller than 8 pt.')
            tolerance = round(914400 * 0.02)
            for name in slide_names:
                tree = xml[name].find('p:cSld/p:spTree', NS)
                if tree is None:
                    continue
                for shape in tree:
                    if shape.tag not in {f'{{{NS["p"]}}}{tag}' for tag in ('sp', 'pic', 'graphicFrame', 'cxnSp')}:
                        continue
                    transform = shape.find('p:spPr/a:xfrm', NS)
                    if transform is None:
                        transform = shape.find('p:xfrm', NS)
                    if transform is None or transform.get('rot') not in (None, '0'):
                        continue
                    off, ext = transform.find('a:off', NS), transform.find('a:ext', NS)
                    if off is None or ext is None:
                        continue
                    x, y = int(off.get('x', '0')), int(off.get('y', '0'))
                    w, h = int(ext.get('cx', '0')), int(ext.get('cy', '0'))
                    if x < -tolerance or y < -tolerance or x + w > width + tolerance or y + h > height + tolerance:
                        warnings.append(f'{name}: a top-level object extends off slide; inspect its intent visually.')
                content = ''.join(t.text or '' for t in xml[name].findall('.//a:t', NS))
                for placeholder in ('此处填写', '在这里添加', 'according to your need', 'Lorem ipsum', '请将本页标题改为核心结论'):
                    if placeholder.casefold() in content.casefold():
                        warnings.append(f'{name}: possible unfinished template text: {placeholder}')
        return stats, warnings


def validate_theme_config(root: Path) -> tuple[dict, list[str]]:
    """Validate theme schema, scale invariants, layout tokens and asset references."""
    errors: list[str] = []
    path = root / 'scripts/office/huaxin_theme_config.json'
    schema_path = root / 'references/theme-config.schema.json'
    try:
        config = json.loads(path.read_text(encoding='utf-8'))
    except (OSError, ValueError) as exc:
        return {}, [f'Theme config could not be read: {exc}']
    if not isinstance(config, dict):
        return {}, ['Theme config must be a JSON object.']
    try:
        schema = json.loads(schema_path.read_text(encoding='utf-8'))
    except (OSError, ValueError) as exc:
        errors.append(f'Theme schema could not be read: {exc}')
        schema = {}
    if not isinstance(schema, dict):
        errors.append('Theme schema must be a JSON object.')
    else:
        # Run the actual recursive Draft 2020-12 validation before the
        # cross-field checks below.  The previous implementation only checked
        # top-level keys, which let malformed nested values pass silently.
        try:
            from jsonschema import Draft202012Validator, SchemaError
        except ImportError:
            errors.append(
                'Theme config requires the jsonschema package for recursive '
                'validation. Install requirements.txt (jsonschema>=4.23,<5).'
            )
        else:
            try:
                Draft202012Validator.check_schema(schema)
                validator = Draft202012Validator(schema)
                schema_errors = sorted(
                    validator.iter_errors(config),
                    key=lambda item: [str(part) for part in item.path],
                )

                def error_path(item: object) -> str:
                    path_items = list(getattr(item, 'absolute_path', ()))
                    if not path_items:
                        return '$'
                    rendered = '$'
                    for part in path_items:
                        rendered += f'[{part}]' if isinstance(part, int) else f'.{part}'
                    return rendered

                # Keep diagnostics readable for hand-edited configs while
                # still making it explicit that more errors may exist.
                for item in schema_errors[:50]:
                    errors.append(f'JSON Schema: {error_path(item)} {item.message}')
                if len(schema_errors) > 50:
                    errors.append(f'JSON Schema: {len(schema_errors) - 50} additional error(s) omitted.')
            except SchemaError as exc:
                errors.append(f'Theme schema is invalid JSON Schema: {exc.message}')
        schema_required = set(schema.get('required', []))
        schema_properties = set(schema.get('properties', {}))
        missing = sorted(schema_required - set(config))
        unknown = sorted(set(config) - schema_properties)
        if missing:
            errors.append(f'Theme config is missing schema fields: {", ".join(missing)}.')
        if unknown:
            errors.append(f'Theme config has fields absent from schema: {", ".join(unknown)}.')
        if schema.get('$schema') != 'https://json-schema.org/draft/2020-12/schema':
            errors.append('Theme schema must use JSON Schema draft 2020-12.')
        if '$defs' not in schema or not isinstance(schema.get('$defs'), dict):
            errors.append('Theme schema must define reusable $defs.')

    def number(value: object) -> bool:
        return isinstance(value, (int, float)) and not isinstance(value, bool)

    for key in ('canvas', 'design_canvas', 'safe_area', 'grid', 'spacing', 'density', 'components', 'typography', 'charts', 'images', 'slides'):
        if not isinstance(config.get(key), dict):
            errors.append(f'Theme {key} must be an object.')
    canvas = config.get('canvas', {})
    design_canvas = config.get('design_canvas', {})
    scale = config.get('runtime_scale', {})
    if isinstance(canvas, dict):
        if canvas.get('units') != 'in':
            errors.append("Theme canvas.units must be 'in'.")
        if not (number(canvas.get('width')) and number(canvas.get('height')) and canvas['width'] > 0 and canvas['height'] > 0):
            errors.append('Theme canvas dimensions must be numeric.')
        elif canvas.get('aspect_ratio') != '16:9' or abs(canvas['width'] / canvas['height'] - 16 / 9) > 0.002:
            errors.append('Theme runtime canvas must be 16:9.')
    if isinstance(design_canvas, dict):
        if design_canvas.get('units') != 'in':
            errors.append("Theme design_canvas.units must be 'in'.")
        if not (number(design_canvas.get('width')) and number(design_canvas.get('height'))):
            errors.append('Theme design canvas dimensions must be numeric.')
    if isinstance(scale, dict) and isinstance(canvas, dict) and isinstance(design_canvas, dict):
        sx, sy = scale.get('x'), scale.get('y')
        if not (number(sx) and number(sy) and sx > 0 and sy > 0):
            errors.append('Theme runtime_scale.x/y must be positive numbers.')
        else:
            design_width, design_height = design_canvas.get('width'), design_canvas.get('height')
            if not (number(design_width) and design_width > 0 and number(design_height) and design_height > 0):
                errors.append('Theme design canvas dimensions must be positive for scale validation.')
            elif abs(canvas.get('width', 0) / design_width - sx) > 0.002:
                errors.append('Theme runtime_scale.x does not match canvas/design_canvas width.')
            if number(design_height) and design_height > 0 and abs(canvas.get('height', 0) / design_height - sy) > 0.002:
                errors.append('Theme runtime_scale.y does not match canvas/design_canvas height.')

    grid = config.get('grid', {})
    if isinstance(grid, dict):
        if grid.get('columns') != 12:
            errors.append('Theme grid.columns must be 12.')
        gap, gap_range = grid.get('gap'), grid.get('gap_range', {})
        if not number(gap) or not isinstance(gap_range, dict) or not number(gap_range.get('min')) or not number(gap_range.get('max')):
            errors.append('Theme grid gap and gap_range must be numeric.')
        elif not gap_range['min'] <= gap <= gap_range['max']:
            errors.append('Theme grid.gap must be inside grid.gap_range.')
        if grid.get('coordinate_space') != 'runtime':
            errors.append("Theme grid.coordinate_space must be 'runtime'.")

    spacing = config.get('spacing', {})
    if isinstance(spacing, dict):
        if spacing.get('coordinate_space') != 'runtime':
            errors.append("Theme spacing.coordinate_space must be 'runtime'.")
        scale_values = spacing.get('scale', [])
        if not isinstance(scale_values, list) or not scale_values or any(not number(v) or v <= 0 for v in scale_values):
            errors.append('Theme spacing.scale must contain positive numbers.')
        elif sorted(scale_values) != scale_values or abs(spacing.get('base', 0) - scale_values[0]) > 0.001:
            errors.append('Theme spacing.scale must be sorted and start at spacing.base.')
        padding_range = spacing.get('card_padding_range', {})
        if not isinstance(padding_range, dict) or not number(padding_range.get('min')) or not number(padding_range.get('max')) or not number(spacing.get('card_padding')) or not padding_range['min'] <= spacing['card_padding'] <= padding_range['max']:
            errors.append('Theme card padding must be inside card_padding_range.')
        if number(spacing.get('card_gap')) and spacing['card_gap'] < 0.15:
            errors.append('Theme spacing.card_gap must be at least 0.15in in runtime coordinates.')

    density = config.get('density', {})
    if isinstance(density, dict):
        if density.get('max_top_level_bullets') != 5 or density.get('max_major_cards') != 4:
            errors.append('Theme density must cap bullets at 5 and major cards at 4.')
        if number(density.get('max_content_area_ratio')) and not 0 < density['max_content_area_ratio'] <= 0.75:
            errors.append('Theme density.max_content_area_ratio must be <= 0.75.')
        if density.get('body_min_runtime_pt') != 11 or density.get('chart_min_runtime_pt') != 10:
            errors.append('Theme density body/chart runtime floors must be 11pt/10pt.')

    components = config.get('components', {})
    radius = components.get('corner_radius', {}) if isinstance(components, dict) else {}
    minimum, default, maximum = radius.get('min_in'), radius.get('default_in'), radius.get('max_in')
    if radius.get('unit') != 'in':
        errors.append("Theme corner radius unit must be 'in'.")
    if not all(number(value) for value in (minimum, default, maximum)):
        errors.append('Theme corner radius bounds must be numeric.')
    elif not (0.104 <= minimum <= default <= maximum <= 0.167):
        errors.append('Theme corner radius must stay ordered within 0.104–0.167 inches.')
    card = components.get('card', {}) if isinstance(components, dict) else {}
    if card.get('shape') != 'ROUNDED_RECTANGLE':
        errors.append('Theme card shape must be ROUNDED_RECTANGLE.')
    if not number(card.get('radius_in')) or not 0.104 <= card['radius_in'] <= 0.167:
        errors.append('Theme card radius_in must be within 0.104–0.167 inches.')
    elif number(default) and abs(card['radius_in'] - default) > 0.001:
        errors.append('Theme card radius_in must equal components.corner_radius.default_in.')
    pill = components.get('pill', {}) if isinstance(components, dict) else {}
    if pill.get('shape') != 'ROUNDED_RECTANGLE':
        errors.append('Theme pill shape must be ROUNDED_RECTANGLE.')
    structural = components.get('structural', {}) if isinstance(components, dict) else {}
    if structural.get('shape') != 'RECTANGLE':
        errors.append('Theme structural shape must be RECTANGLE.')

    typography = config.get('typography', {})
    semantic_keys = ('cover_title_pt', 'cover_subtitle_pt', 'slide_title_pt', 'long_title_pt', 'conclusion_pt', 'card_title_pt', 'body_pt', 'caption_pt', 'label_pt', 'key_number_pt', 'chart_pt')
    if isinstance(typography, dict):
        design_values, runtime_values = typography.get('design', {}), typography.get('runtime', {})
        ranges = typography.get('ranges', {})
        runtime_floor = typography.get('runtime_floor_pt', {})
        if not isinstance(design_values, dict) or not isinstance(runtime_values, dict) or not isinstance(ranges, dict):
            errors.append('Theme typography.design/runtime/ranges must be objects.')
        else:
            sy = scale.get('y') if isinstance(scale, dict) else None
            for key in semantic_keys:
                design_value, runtime_value, value_range = design_values.get(key), runtime_values.get(key), ranges.get(key, {})
                if not (number(design_value) and number(runtime_value) and isinstance(value_range, dict)):
                    errors.append(f'Theme typography.{key} must define design/runtime values and ranges.')
                    continue
                if number(sy) and abs(runtime_value - design_value * sy) > 0.02:
                    errors.append(f'Theme typography.{key} runtime value must equal design value × runtime_scale.y.')
                for space in ('design', 'runtime'):
                    current = value_range.get(space, {})
                    if not isinstance(current, dict) or not number(current.get('min')) or not number(current.get('max')) or not current['min'] <= current['max']:
                        errors.append(f'Theme typography.{key}.{space} range is invalid.')
                    elif not current['min'] <= (design_value if space == 'design' else runtime_value) <= current['max']:
                        errors.append(f'Theme typography.{key} value is outside its {space} range.')
            if (number(runtime_floor.get('body')) and number(runtime_values.get('body_pt')) and
                    runtime_values['body_pt'] < runtime_floor['body']):
                errors.append('Theme typography.body_pt is below the body runtime floor.')
            if (number(runtime_floor.get('chart')) and number(runtime_values.get('chart_pt')) and
                    runtime_values['chart_pt'] < runtime_floor['chart']):
                errors.append('Theme typography.chart_pt is below the chart runtime floor.')
            aliases = {
                'title_pt': ('slide_title_pt', runtime_values.get('slide_title_pt')),
                'body_pt': ('body_pt', runtime_values.get('body_pt')),
                'caption_pt': ('caption_pt', runtime_values.get('caption_pt')),
            }
            for alias, (semantic, expected) in aliases.items():
                if number(expected) and number(typography.get(alias)) and abs(typography[alias] - expected) > 0.02:
                    errors.append(f'Theme typography.{alias} must mirror typography.runtime.{semantic}.')
    charts = config.get('charts', {})
    if isinstance(charts, dict):
        if charts.get('allow_3d') is not False:
            errors.append('Theme charts.allow_3d must be false.')
        if (not number(charts.get('min_font_size_pt')) or charts['min_font_size_pt'] < 10 or
                not isinstance(charts.get('max_donut_categories'), int) or charts['max_donut_categories'] > 5):
            errors.append('Theme chart text/category limits are too permissive.')

    def validate_asset_reference(reference: object, label: str) -> None:
        if not isinstance(reference, str) or not reference.startswith('assets/'):
            errors.append(f'Theme {label} must reference an assets/ path.')
            return
        if 'huaxin_bg_' in reference or 'huxin-pptx' in reference:
            errors.append(f'Theme {label} uses a retired asset/skill identifier: {reference}.')
        asset_path = root / reference
        if not asset_path.is_file():
            errors.append(f'Theme {label} does not exist: {reference}.')
    bg_rules = config.get('bg_rules', {})
    if isinstance(bg_rules, dict):
        validate_asset_reference(bg_rules.get('title_bg'), 'bg_rules.title_bg')
        validate_asset_reference(bg_rules.get('content_bg'), 'bg_rules.content_bg')
        if bg_rules.get('title_bg') != 'assets/cover-background.jpeg' or bg_rules.get('content_bg') != 'assets/body-background.jpeg':
            errors.append('Theme background paths must use the canonical cover/body asset names.')
    slides = config.get('slides', {})
    safe = config.get('safe_area', {})
    footer = config.get('footer', {})
    spacing_tokens = config.get('spacing', {})
    density_tokens = config.get('density', {})
    if isinstance(safe, dict):
        bottom = safe.get('bottom')
        footer_top_clearance = safe.get('footer_top_clearance')
        if safe.get('bottom_semantics') != 'footer_top_clearance':
            errors.append("Theme safe_area.bottom_semantics must be 'footer_top_clearance'.")
        if number(bottom) and number(footer_top_clearance) and abs(bottom - footer_top_clearance) > 0.001:
            errors.append('Theme safe_area.bottom must equal safe_area.footer_top_clearance.')
    if isinstance(slides, dict):
        def validate_rect_bounds(label: str, rect: object, width: object, height: object) -> None:
            if not isinstance(rect, dict):
                return
            if not number(width) or not number(height):
                return  # Recursive JSON Schema reports invalid canvas dimensions.
            values = [rect.get(key) for key in ('x', 'y', 'w', 'h')]
            if not all(number(value) for value in values):
                return  # Recursive JSON Schema reports the type details.
            x, y, w, h = values
            if x < 0 or y < 0 or x + w > width + 0.001 or y + h > height + 0.001:
                errors.append(
                    f'Theme {label} exceeds the runtime canvas bounds '
                    f'(x={x}, y={y}, w={w}, h={h}; canvas={width}×{height}).'
                )

        runtime_width = canvas.get('width') if isinstance(canvas, dict) else None
        runtime_height = canvas.get('height') if isinstance(canvas, dict) else None
        type_ranges = typography.get('ranges', {}) if isinstance(typography, dict) else {}
        runtime_ranges = {key: value.get('runtime', {}) for key, value in type_ranges.items() if isinstance(value, dict)}
        for name, slide in slides.items():
            if not isinstance(slide, dict):
                errors.append(f'Theme slides.{name} must be an object.')
                continue
            validate_asset_reference(slide.get('bg_image'), f'slides.{name}.bg_image')
            title = slide.get('title')
            validate_rect_bounds(f'slides.{name}.title', title, runtime_width, runtime_height)
            if isinstance(title, dict) and number(title.get('fontSize')):
                range_key = 'cover_title_pt' if name in {'title_slide', 'end_slide'} else 'slide_title_pt'
                allowed = runtime_ranges.get(range_key, {})
                if isinstance(allowed, dict) and number(allowed.get('min')) and number(allowed.get('max')):
                    if not allowed['min'] <= title['fontSize'] <= allowed['max']:
                        errors.append(f'Theme slides.{name}.title.fontSize must be inside typography.ranges.{range_key}.runtime.')
            elif isinstance(title, dict):
                errors.append(f'Theme slides.{name}.title.fontSize must be numeric.')
            subtitle = slide.get('subtitle')
            validate_rect_bounds(f'slides.{name}.subtitle', subtitle, runtime_width, runtime_height)
            if isinstance(subtitle, dict) and name == 'title_slide' and number(subtitle.get('fontSize')):
                allowed = runtime_ranges.get('cover_subtitle_pt', {})
                if isinstance(allowed, dict) and number(allowed.get('min')) and number(allowed.get('max')):
                    if not allowed['min'] <= subtitle['fontSize'] <= allowed['max']:
                        errors.append('Theme title_slide.subtitle.fontSize must be inside typography.ranges.cover_subtitle_pt.runtime.')
            content = slide.get('content_area')
            if isinstance(content, dict):
                validate_rect_bounds(f'slides.{name}.content_area', content, runtime_width, runtime_height)
                if isinstance(safe, dict) and number(safe.get('top')) and number(content.get('y')):
                    if abs(content['y'] - safe['top']) > 0.001:
                        errors.append(
                            f'Theme slides.{name}.content_area.y must equal safe_area.top '
                            f'({safe["top"]:.3f}).'
                        )
                elif isinstance(safe, dict) and not number(safe.get('top')):
                    errors.append('Theme safe_area.top must be numeric before content_area can be bound.')
                if (isinstance(safe, dict) and number(safe.get('left')) and number(safe.get('right'))
                        and number(runtime_width) and number(content.get('x')) and number(content.get('w'))):
                    expected_content_x = safe['left']
                    expected_content_w = runtime_width - safe['left'] - safe['right']
                    if abs(content['x'] - expected_content_x) > 0.001 or abs(content['w'] - expected_content_w) > 0.001:
                        errors.append(
                            f'Theme slides.{name}.content_area must bind to horizontal safe_area '
                            f'(x={expected_content_x:.3f}, w={expected_content_w:.3f}).'
                        )
                if (number(runtime_width) and number(runtime_height) and
                        all(number(content.get(key)) for key in ('w', 'h'))):
                    ratio = content['w'] * content['h'] / (runtime_width * runtime_height)
                    max_ratio = density_tokens.get('max_content_area_ratio') if isinstance(density_tokens, dict) else None
                    if number(max_ratio) and ratio > max_ratio + 0.001:
                        errors.append(
                            f'Theme slides.{name}.content_area ratio {ratio:.3f} '
                            f'exceeds density.max_content_area_ratio {max_ratio:.3f}.'
                        )
                # The title-to-content and content-to-footer rhythm is a
                # cross-field invariant.  It also catches regressions that
                # recreate the excessive top/bottom whitespace this theme was
                # designed to avoid.
                if isinstance(title, dict) and number(title.get('y')) and number(title.get('h')) and number(content.get('y')):
                    title_gap = spacing_tokens.get('title_to_content') if isinstance(spacing_tokens, dict) else None
                    if number(title_gap) and content['y'] + 0.001 < title['y'] + title['h'] + title_gap:
                        errors.append(f'Theme slides.{name} content starts before title_to_content spacing is satisfied.')
                if number(content.get('y')) and number(content.get('h')) and number(footer.get('y')):
                    block_gap = spacing_tokens.get('block_gap') if isinstance(spacing_tokens, dict) else None
                    if number(block_gap) and content['y'] + content['h'] > footer['y'] - block_gap + 0.001:
                        errors.append(f'Theme slides.{name} content collides with the footer rhythm.')
                if isinstance(safe, dict) and number(safe.get('left')) and number(safe.get('right')):
                    if number(content.get('x')) and number(content.get('w')) and number(runtime_width):
                        if content['x'] < safe['left'] - 0.001 or content['x'] + content['w'] > runtime_width - safe['right'] + 0.001:
                            errors.append(f'Theme slides.{name} content_area violates horizontal safe_area bounds.')
    if isinstance(footer, dict) and isinstance(canvas, dict):
        canvas_width, canvas_height = canvas.get('width'), canvas.get('height')
        if all(number(footer.get(key)) for key in ('y', 'height')) and number(canvas_height) and footer['y'] + footer['height'] > canvas_height + 0.001:
            errors.append('Theme footer y+height exceeds the runtime canvas height.')
        if number(canvas_height) and isinstance(safe, dict) and number(safe.get('footer_top_clearance')) and number(footer.get('y')):
            expected_footer_y = canvas_height - safe['footer_top_clearance']
            if abs(footer['y'] - expected_footer_y) > 0.001:
                errors.append(
                    'Theme footer.y must equal canvas.height - safe_area.footer_top_clearance '
                    f'({expected_footer_y:.3f}).'
                )
        if all(number(footer.get(key)) for key in ('page_x', 'page_width')) and number(canvas_width) and footer['page_x'] + footer['page_width'] > canvas_width + 0.001:
            errors.append('Theme footer page_x+page_width exceeds the runtime canvas width.')
        if all(number(footer.get(key)) for key in ('source_x', 'source_width')) and number(canvas_width) and footer['source_x'] + footer['source_width'] > canvas_width + 0.001:
            errors.append('Theme footer source_x+source_width exceeds the runtime canvas width.')
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

    def require(condition: bool, message: str) -> None:
        if not condition:
            errors.append(message)

    def number(value: object) -> bool:
        return isinstance(value, (int, float)) and not isinstance(value, bool)

    required = ['SKILL.md', 'README.md', 'LICENSE.txt', 'agents/openai.yaml', 'requirements.txt',
                'references/style-guide.md', 'references/quality-checklist.md',
                'references/source-notes.md', 'references/assets-manifest.json',
                'assets/cover-background.jpeg', 'assets/body-background.jpeg',
                'assets/huaxin-template-16x9.pptx', 'scripts/create_starter.py',
                 'scripts/validate_skill.py', 'tests/acceptance-cases.md',
                 'tests/package-integrity.py', 'tests/validator-regression.py',
                 'tests/density-regression.py',
                 'tests/validation-report.md',
                'references/theme-config.schema.json', 'scripts/office/huaxin_theme_config.json',
                'scripts/office/render_visual.py']
    for name in required:
        require((root / name).is_file(), f'Missing required file: {name}')
    theme_config, theme_errors = validate_theme_config(root)
    errors.extend(theme_errors)
    if theme_config:
        canvas = theme_config.get('canvas', {})
        canvas_width = canvas.get('width') if isinstance(canvas, dict) else None
        canvas_height = canvas.get('height') if isinstance(canvas, dict) else None
        content_area_ratios: dict[str, float] = {}
        if number(canvas_width) and number(canvas_height) and canvas_width > 0 and canvas_height > 0:
            for slide_name, slide in theme_config.get('slides', {}).items():
                area = slide.get('content_area') if isinstance(slide, dict) else None
                if isinstance(area, dict) and all(number(area.get(key)) for key in ('w', 'h')):
                    content_area_ratios[slide_name] = round(area['w'] * area['h'] / (canvas_width * canvas_height), 6)
        checks['theme'] = {
            'path': str(root / 'scripts/office/huaxin_theme_config.json'),
            'schema_validation': 'JSON Schema Draft 2020-12 (recursive)',
            'schema_version': theme_config.get('schema_version'),
            'grid': theme_config.get('grid'),
            'typography_runtime': theme_config.get('typography', {}).get('runtime'),
            'density': theme_config.get('density'),
            'corner_radius': theme_config.get('components', {}).get('corner_radius'),
            'card': theme_config.get('components', {}).get('card'),
            'structural': theme_config.get('components', {}).get('structural'),
            'content_area_ratios': content_area_ratios,
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
        require('$pptx' in interface.get('default_prompt', ''), 'Agent default prompt references the wrong skill.')
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
        require(manifest['skill_name'] == 'pptx', 'Asset manifest references the wrong skill.')
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
            density_limits = theme_config.get('density', {}) if isinstance(theme_config, dict) else None
            if isinstance(theme_config, dict):
                charts_config = theme_config.get('charts', {})
                if isinstance(charts_config, dict):
                    density_limits = dict(density_limits or {})
                    density_limits['max_donut_categories'] = charts_config.get('max_donut_categories')
            stats, deck_warnings = inspect_pptx(target, check_bounds=True, density_limits=density_limits)
            if matched_visual_deck is not None and not visual_errors:
                stats['visual_check'] = 'LIBREOFFICE_PASS'
            checks['generated_pptx'] = {'path': str(target), **stats}
            if stats.get('min_explicit_font_size_pt') is not None:
                minimum_font = stats['min_explicit_font_size_pt']
                require(minimum_font >= 8, 'Generated PPTX contains explicit text below the 8 pt absolute minimum.')
                # When density_limits is present, validate_pptx_density emits
                # the configured body-floor diagnostic with the same context.
                if minimum_font < 11 and not density_limits:
                    warnings.append('Generated PPTX contains explicit text below the 11 pt body readability floor; confirm it is caption/footer text.')
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
