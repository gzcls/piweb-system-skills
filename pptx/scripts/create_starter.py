#!/usr/bin/env python3
"""Create an UNFINISHED, editable Huaxin PPTX starter using bundled backgrounds.

This is a fallback scaffold, not an importer or a complete presentation generator.
Dependencies: python-pptx and Pillow. No network calls or font files are required.
"""
from __future__ import annotations

import argparse
import json
import os
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

THEME_CONFIG_PATH = ROOT / 'scripts/office/huaxin_theme_config.json'

@dataclass(frozen=True)
class ThemeTokens:
    """Runtime-canonical tokens loaded from the bundled theme contract.

    All slide geometry in the JSON contract is expressed in runtime inches.
    Design-mode output is derived by applying ``design/runtime`` separately on
    each axis.  Keeping this conversion here prevents callers from mixing the
    10 × 5.625 source canvas with the 13.333 × 7.5 design reference.
    """

    config: dict
    runtime_width: float
    runtime_height: float
    design_width: float
    design_height: float
    runtime_scale_x: float
    runtime_scale_y: float
    runtime_font_sizes: dict[str, float]
    design_font_sizes: dict[str, float]
    font_family: str
    content_start_y: float


def load_theme_tokens() -> ThemeTokens:
    """Load and sanity-check the machine-readable theme contract.

    The starter must fail fast when the contract is malformed.  Falling back
    to an embedded set of geometry values silently produces slides that no
    longer match the active theme and was the source of divergent layouts.
    """
    try:
        config = json.loads(THEME_CONFIG_PATH.read_text(encoding='utf-8'))
    except (OSError, json.JSONDecodeError) as exc:
        raise RuntimeError(f'Could not read theme config: {THEME_CONFIG_PATH}: {exc}') from exc
    try:
        canvas = config['canvas']
        design_canvas = config['design_canvas']
        scale = config['runtime_scale']
        typography = config['typography']
        fonts = config['fonts']
        safe_area = config['safe_area']
        runtime_width = float(canvas['width'])
        runtime_height = float(canvas['height'])
        design_width = float(design_canvas['width'])
        design_height = float(design_canvas['height'])
        runtime_scale_x = float(scale['x'])
        runtime_scale_y = float(scale['y'])
        if min(runtime_width, runtime_height, design_width, design_height,
               runtime_scale_x, runtime_scale_y) <= 0:
            raise ValueError('canvas dimensions and runtime scales must be positive')
        expected_x = runtime_width / design_width
        expected_y = runtime_height / design_height
        if abs(expected_x - runtime_scale_x) > 0.002 or abs(expected_y - runtime_scale_y) > 0.002:
            raise ValueError('runtime_scale does not match canvas/design_canvas dimensions')
        design_sizes = {
            key: float(value) for key, value in typography['design'].items()
        }
        runtime_sizes = {
            key: float(value) for key, value in typography['runtime'].items()
        }
        if set(design_sizes) != set(runtime_sizes):
            raise ValueError('typography.design and typography.runtime keys differ')
        for key, value in design_sizes.items():
            if abs(runtime_sizes[key] - value * runtime_scale_y) > 0.02:
                raise ValueError(f'typography.runtime.{key} is not design × runtime_scale.y')
        font_family = str(fonts['body']).strip()
        if not font_family:
            raise ValueError('fonts.body cannot be blank')
        content_start_y = float(safe_area['top'])
        if content_start_y < 0:
            raise ValueError('safe_area.top cannot be negative')
    except (KeyError, TypeError, ValueError) as exc:
        raise RuntimeError(f'Invalid theme config {THEME_CONFIG_PATH}: {exc}') from exc
    return ThemeTokens(
        config=config,
        runtime_width=runtime_width,
        runtime_height=runtime_height,
        design_width=design_width,
        design_height=design_height,
        runtime_scale_x=runtime_scale_x,
        runtime_scale_y=runtime_scale_y,
        runtime_font_sizes=runtime_sizes,
        design_font_sizes=design_sizes,
        font_family=font_family,
        content_start_y=content_start_y,
    )


THEME = load_theme_tokens()
DEFAULT_FONT_FAMILY = THEME.font_family


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', required=True, type=Path, help='Output .pptx outside the skill folder')
    parser.add_argument('--title', required=True, help='Cover title; at most two lines, 24 characters per line')
    parser.add_argument('--subtitle', default='', help='Optional cover subtitle; at most 36 characters')
    parser.add_argument('--presenter', default='', help='Omit to retain 汇报人')
    parser.add_argument('--date', default='', help='Omit to retain 汇报日期; never inferred from the current date')
    parser.add_argument('--body-title', action='append', help='Repeat for each blank body slide; up to 24 characters')
    parser.add_argument('--component-fixture', action='store_true',
                        help='Add a small editable card/footer fixture to body slides for theme and visual QA')
    parser.add_argument('--closing', default='感谢聆听', help='Closing title; at most two lines, 24 characters per line')
    parser.add_argument('--font', default=DEFAULT_FONT_FAMILY, help='Name of an installed CJK-capable sans-serif font')
    parser.add_argument('--canvas', choices=('design', 'original'), default='design')
    parser.add_argument('--force', action='store_true', help='Overwrite an ordinary output file, never skill assets')
    args = parser.parse_args()

    try:
        from PIL import Image
        from pptx import Presentation
        from pptx.dml.color import RGBColor
        from pptx.enum.shapes import MSO_SHAPE
        from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
        from pptx.util import Inches, Pt
    except ImportError as exc:
        print(f'Missing dependency: {exc}. Install requirements.txt in your Python environment.', file=sys.stderr)
        return 2

    output = args.output.expanduser().resolve()
    if output == ROOT or ROOT in output.parents:
        parser.error('Output must be outside the skill directory. Bundled resources are read-only.')
    if output.suffix.lower() != '.pptx':
        parser.error('--output must end in .pptx')
    if output.exists() and not args.force:
        parser.error('Output already exists. Choose a different path or explicitly pass --force.')
    if output.exists() and not output.is_file():
        parser.error('Output is not an ordinary file.')
    for label, text in [('title', args.title), ('closing', args.closing)]:
        lines = text.splitlines()
        if not text.strip() or len(lines) > 2 or any(len(line) > 24 for line in lines):
            parser.error(f'--{label} requires 1–2 lines with at most 24 characters per line.')
    if '\n' in args.subtitle or len(args.subtitle) > 36:
        parser.error('--subtitle must be one line with at most 36 characters.')
    body_titles = args.body_title or ['请将本页标题改为核心结论']
    for title in body_titles:
        if not title.strip() or '\n' in title or len(title) > 24:
            parser.error('Each --body-title must be one non-empty line with at most 24 characters.')
    label = f'{args.presenter.strip() or "汇报人"}｜{args.date.strip() or "汇报日期"}'
    if '\n' in label or len(label) > 40:
        parser.error('The combined presenter/date label must be one line with at most 40 characters.')
    if not args.font.strip():
        parser.error('--font cannot be blank.')

    slide_tokens = THEME.config['slides']
    title_slide_token = slide_tokens['title_slide']
    body_slide_token = slide_tokens['content_slide']
    end_slide_token = slide_tokens['end_slide']
    cover = ROOT / title_slide_token['bg_image']
    body = ROOT / body_slide_token['bg_image']
    for asset in (cover, body):
        if not asset.is_file():
            print(f'Missing asset: {asset}', file=sys.stderr)
            return 1

    # The JSON contract stores runtime geometry.  Select the target canvas and
    # derive x/y scales independently; this remains correct if a future theme
    # uses a non-uniform runtime scale instead of silently assuming one value.
    if args.canvas == 'design':
        target_width, target_height = THEME.design_width, THEME.design_height
        coord_scale_x = 1.0 / THEME.runtime_scale_x
        coord_scale_y = 1.0 / THEME.runtime_scale_y
        font_sizes = THEME.design_font_sizes
        output_space = 'design'
    else:
        target_width, target_height = THEME.runtime_width, THEME.runtime_height
        coord_scale_x = coord_scale_y = 1.0
        font_sizes = THEME.runtime_font_sizes
        output_space = 'runtime'
    prs = Presentation()
    prs.slide_width = Inches(target_width)
    prs.slide_height = Inches(target_height)
    prs.core_properties.title = args.title.replace('\n', ' ')
    prs.core_properties.subject = ('未完成的华新可编辑起稿：正文内容需另行制作并质检'
                                   + ('；包含主题组件验收夹具' if args.component_fixture else ''))
    prs.core_properties.author = args.presenter.strip()
    prs.core_properties.keywords = 'pptx; huaxin; starter; draft'

    def background(slide, path: Path) -> None:
        """Center-crop to fill without changing the source image or stretching it."""
        with Image.open(path) as image:
            iw, ih = image.size
        image_ratio = iw / ih
        frame_ratio = prs.slide_width / prs.slide_height
        picture = slide.shapes.add_picture(str(path), 0, 0, prs.slide_width, prs.slide_height)
        picture.name = 'HX_BACKGROUND_COVER' if path == cover else 'HX_BACKGROUND_BODY'
        if image_ratio > frame_ratio:
            crop = (1.0 - frame_ratio / image_ratio) / 2.0
            picture.crop_left = crop
            picture.crop_right = crop
        else:
            crop = (1.0 - image_ratio / frame_ratio) / 2.0
            picture.crop_top = crop
            picture.crop_bottom = crop

    def text(slide, value: str, name: str, xywh: tuple[float, float, float, float],
             size: float, color: str = 'FFFFFF', bold: bool = False,
             align=PP_ALIGN.CENTER) -> None:
        x, y, width, height = xywh
        shape = slide.shapes.add_textbox(
            Inches(x * coord_scale_x), Inches(y * coord_scale_y),
            Inches(width * coord_scale_x), Inches(height * coord_scale_y),
        )
        shape.name = name
        frame = shape.text_frame
        frame.clear()
        frame.word_wrap = True
        frame.auto_size = MSO_AUTO_SIZE.NONE
        frame.margin_left = frame.margin_right = 0
        frame.margin_top = frame.margin_bottom = 0
        frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        for index, line in enumerate(value.split('\n')):
            paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
            paragraph.alignment = align
            paragraph.space_before = Pt(0)
            paragraph.space_after = Pt(0)
            paragraph.line_spacing = 1.12
            run = paragraph.add_run()
            run.text = line
            run.font.name = args.font
            run.font.size = Pt(size)
            run.font.bold = bold
            run.font.color.rgb = RGBColor.from_string(color)

    def component_fixture(slide) -> None:
        """Place representative editable components using only theme tokens.

        This opt-in fixture gives visual and structural QA a real body page to
        inspect while keeping the default starter intentionally blank.  All
        geometry is read from the runtime content/footer tokens and converted
        with the same axis scales as the rest of the starter.
        """
        body_spec = body_slide_token['content_area']
        spacing = THEME.config['spacing']
        components = THEME.config['components']
        card_token = components['card']
        gap = float(spacing['card_gap'])
        card_w = (float(body_spec['w']) - gap) / 2
        card_y = float(body_spec['y'])
        card_h = min(1.52, float(body_spec['h']) - 0.62)
        fill = RGBColor.from_string(THEME.config['colors']['white'])
        ink = RGBColor.from_string(THEME.config['colors']['text_dark'])
        accent = RGBColor.from_string(THEME.config['colors']['brand_blue'])

        for index, card_title in enumerate(('关键结论', '行动建议'), start=1):
            x = float(body_spec['x']) + (index - 1) * (card_w + gap)
            card = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(x * coord_scale_x), Inches(card_y * coord_scale_y),
                Inches(card_w * coord_scale_x), Inches(card_h * coord_scale_y),
            )
            card.name = f'HX_CARD_{index}'
            card.fill.solid()
            card.fill.fore_color.rgb = fill
            card.line.color.rgb = accent
            card.line.width = Pt(float(components['border']['default_pt']))
            card.text_frame.clear()
            card.text_frame.margin_left = Inches(float(spacing['card_padding']) * coord_scale_x)
            card.text_frame.margin_right = Inches(float(spacing['card_padding']) * coord_scale_x)
            card.text_frame.margin_top = Inches(float(spacing['card_padding']) * coord_scale_y)
            card.text_frame.margin_bottom = Inches(float(spacing['card_padding']) * coord_scale_y)
            paragraph = card.text_frame.paragraphs[0]
            paragraph.alignment = PP_ALIGN.LEFT
            paragraph.space_after = Pt(3)
            run = paragraph.add_run()
            run.text = card_title
            run.font.name = args.font
            run.font.size = Pt(font_sizes['card_title_pt'])
            run.font.bold = True
            run.font.color.rgb = accent
            paragraph = card.text_frame.add_paragraph()
            paragraph.alignment = PP_ALIGN.LEFT
            paragraph.line_spacing = 1.12
            run = paragraph.add_run()
            run.text = '用主题间距与字号承载一条可编辑正文。'
            run.font.name = args.font
            run.font.size = Pt(font_sizes['body_pt'])
            run.font.color.rgb = ink

        footer = THEME.config['footer']
        footer_shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(float(footer['source_x']) * coord_scale_x), Inches(float(footer['y']) * coord_scale_y),
            Inches(float(footer['source_width']) * coord_scale_x), Inches(float(footer['height']) * coord_scale_y),
        )
        footer_shape.name = 'HX_FOOTER'
        footer_shape.fill.solid()
        footer_shape.fill.fore_color.rgb = accent
        footer_shape.line.color.rgb = accent
        footer_shape.text_frame.clear()
        footer_shape.text_frame.margin_left = footer_shape.text_frame.margin_right = 0
        footer_shape.text_frame.margin_top = footer_shape.text_frame.margin_bottom = 0
        footer_paragraph = footer_shape.text_frame.paragraphs[0]
        footer_paragraph.alignment = PP_ALIGN.LEFT
        footer_run = footer_paragraph.add_run()
        footer_run.text = '主题组件夹具｜可替换为业务页脚'
        footer_run.font.name = args.font
        footer_run.font.size = Pt(font_sizes['caption_pt'])
        footer_run.font.color.rgb = fill

    def make_slide(path: Path):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        background(slide, path)
        slide.notes_slide.notes_text_frame.text = (
            '此页由 pptx 起稿脚本创建，尚未完成正文制作和最终视觉质检。'
            '背景来源为随包原始华新模板；主要文本框可编辑。'
            f'正文内容区默认从 y={THEME.content_start_y:.2f} 英寸开始；'
            f'本页采用 {output_space} 坐标，几何按 x={coord_scale_x:.4f}/y={coord_scale_y:.4f} 换算。'
            + ('正文页包含 --component-fixture 主题组件验收夹具；交付前请替换示例文案。'
               if args.component_fixture else '')
        )
        return slide

    slide = make_slide(cover)
    # Longer explicit lines receive a smaller, still in-range cover title.
    cover_spec = title_slide_token['title']
    cover_size = (font_sizes['long_title_pt']
                  if max(map(len, args.title.splitlines())) > 19
                  else font_sizes['cover_title_pt'])
    text(slide, args.title, 'HX_COVER_TITLE', tuple(cover_spec[key] for key in ('x', 'y', 'w', 'h')),
         cover_size, color=cover_spec['color'], bold=cover_spec['bold'],
         align={'left': PP_ALIGN.LEFT, 'center': PP_ALIGN.CENTER, 'right': PP_ALIGN.RIGHT}[cover_spec['align']])
    if args.subtitle:
        subtitle_spec = title_slide_token['subtitle']
        text(slide, args.subtitle, 'HX_COVER_SUBTITLE',
             tuple(subtitle_spec[key] for key in ('x', 'y', 'w', 'h')),
             font_sizes['cover_subtitle_pt'], color=subtitle_spec['color'],
             align={'left': PP_ALIGN.LEFT, 'center': PP_ALIGN.CENTER, 'right': PP_ALIGN.RIGHT}[subtitle_spec['align']])
    # Actual bottom-right region is white: dark blue-gray is required for contrast.
    metadata_spec = title_slide_token['metadata']
    text(slide, label, 'HX_COVER_METADATA',
         tuple(metadata_spec[key] for key in ('x', 'y', 'w', 'h')),
         metadata_spec['fontSize'] if output_space == 'runtime'
         else metadata_spec['fontSize'] / THEME.runtime_scale_y,
         color=metadata_spec['color'],
         align={'left': PP_ALIGN.LEFT, 'center': PP_ALIGN.CENTER, 'right': PP_ALIGN.RIGHT}[metadata_spec['align']])

    for title in body_titles:
        slide = make_slide(body)
        title_spec = body_slide_token['title']
        title_size = (font_sizes['long_title_pt']
                      if len(title) > 20
                      else font_sizes['slide_title_pt'])
        text(slide, title, 'HX_BODY_TITLE', tuple(title_spec[key] for key in ('x', 'y', 'w', 'h')),
             title_size, color=title_spec['color'], bold=title_spec['bold'], align=PP_ALIGN.LEFT)
        if args.component_fixture:
            component_fixture(slide)
        # No duplicated page numbers, logo, decorative bar or unfinished body copy.

    slide = make_slide(cover)
    closing_spec = end_slide_token['title']
    closing_size = (font_sizes['long_title_pt']
                    if max(map(len, args.closing.splitlines())) > 19
                    else font_sizes['cover_title_pt'])
    text(slide, args.closing, 'HX_CLOSING_TITLE', tuple(closing_spec[key] for key in ('x', 'y', 'w', 'h')),
         closing_size, color=closing_spec['color'], bold=closing_spec['bold'],
         align={'left': PP_ALIGN.LEFT, 'center': PP_ALIGN.CENTER, 'right': PP_ALIGN.RIGHT}[closing_spec['align']])

    temp_path = None
    try:
        output.parent.mkdir(parents=True, exist_ok=True)
        fd, temp_name = tempfile.mkstemp(prefix='.pptx-', suffix='.pptx', dir=output.parent)
        os.close(fd)
        temp_path = Path(temp_name)
        prs.save(str(temp_path))
        loaded = Presentation(str(temp_path))
        if len(loaded.slides) != len(body_titles) + 2:
            raise ValueError('Unexpected slide count after reopening the starter.')
        if output.exists() and not args.force:
            raise FileExistsError('Output appeared while generating; refusing to overwrite it.')
        os.replace(temp_path, output)
        temp_path = None
    except (OSError, ValueError) as exc:
        print(f'Could not create starter: {exc}', file=sys.stderr)
        return 1
    finally:
        if temp_path is not None:
            temp_path.unlink(missing_ok=True)

    print(json.dumps({
        'output': str(output), 'slides': len(prs.slides), 'status': 'UNFINISHED_STARTER',
        'canvas_inches': [round(prs.slide_width / 914400, 6), round(prs.slide_height / 914400, 6)],
        'presenter_date_label': label, 'font_requested': args.font,
        'notes': [
            '正文页仅有标题，内容仍需制作；不要把起稿直接当作完成稿交付。',
            '如使用 --component-fixture，正文页会包含两张主题圆角卡片和主题页脚，用于 LibreOffice 视觉/结构验收。',
            '此为原背景新建方案，不是原模板导入或母版的无损编辑。',
            '已重新解析文件；尚未执行外部渲染、字体可用性和目视质检。'
        ]
    }, ensure_ascii=False, indent=2))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
