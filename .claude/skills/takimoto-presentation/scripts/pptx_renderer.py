"""Reusable .pptx renderer enforcing 瀧本ゼミ slide style + 一般的なスライド設計原則.

Input: a SlideDeckConfig containing a list of Slide objects.
Output: a .pptx file importable to Google Slides.

Rules enforced (see reference/slide_design.md for full spec):
- Font minimums: title 28pt, subtitle 20pt, body 18pt, table 13pt
- 1 message per slide via Slide.subtitle (highlighted as accent)
- Bullets max 5 items (auto-warning if more)
- No image generation — image needs surface as detailed `[画像: ...]` blocks
- Table styling: colored header, alternating rows
- Visual hierarchy: title > subtitle > body > footer
- Designed margins: 0.5in from edges, 0.2in between elements
"""
from __future__ import annotations

from dataclasses import dataclass, field

from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE


# ===========================================================================
# Layout constants (per slide_design.md)
# ===========================================================================
SLIDE_WIDTH = Inches(13.333)  # 16:9 widescreen
SLIDE_HEIGHT = Inches(7.5)

# Margins
MARGIN = Inches(0.5)
TITLE_HEIGHT = Inches(0.9)
SUBTITLE_HEIGHT = Inches(0.7)
FOOTER_HEIGHT = Inches(0.35)
GAP = Inches(0.2)

# Font sizes (minimums, never go below)
FONT_TITLE = Pt(30)
FONT_SUBTITLE = Pt(22)
FONT_BODY = Pt(18)
FONT_TABLE_HEADER = Pt(15)
FONT_TABLE_CELL = Pt(14)
FONT_FOOTER = Pt(11)
FONT_ASCII = Pt(15)
FONT_COVER_TITLE = Pt(54)
FONT_COVER_SUBTITLE = Pt(26)
FONT_CHAPTER_TITLE = Pt(58)

# Color palette (per slide_design.md)
C_TITLE = RGBColor(0x1F, 0x4E, 0x78)        # ディープブルー
C_ACCENT = RGBColor(0xC0, 0x39, 0x2B)       # アクセントレッド
C_POSITIVE = RGBColor(0x2E, 0x7D, 0x32)     # ポジティブグリーン
C_BODY = RGBColor(0x20, 0x20, 0x20)         # ニアブラック
C_HELPER = RGBColor(0x66, 0x66, 0x66)       # グレー
C_FOOTER = RGBColor(0x80, 0x80, 0x80)       # ライトグレー
C_HEADER_BG = RGBColor(0x1F, 0x4E, 0x78)    # テーブルヘッダー
C_STRIPE = RGBColor(0xF5, 0xF5, 0xF5)       # ストライプ
C_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
C_DIVIDER_BG = RGBColor(0x1F, 0x4E, 0x78)


@dataclass
class Slide:
    """One slide definition. See slide_design.md for visual guidelines."""

    kind: str
    # kind ∈ "title" "chapter" "content" "toc" "table" "image_placeholder" "ascii"
    title: str = ""
    subtitle: str = ""  # 1スライド1メッセージの「メッセージ」
    bullets: list = field(default_factory=list)
    table: list = field(default_factory=list)
    table_headers: list = field(default_factory=list)
    ascii_art: str = ""
    image_placeholder: str = ""  # 具体的な指示テキスト
    image_caption: str = ""
    footer: str = ""
    note: str = ""


@dataclass
class SlideDeckConfig:
    title: str  # e.g., 【4971】メック
    subtitle: str  # multi-line: 推奨 / 目標株価 / 投資期間
    presenter: str
    slides: list = field(default_factory=list)


# ===========================================================================
# Low-level helpers
# ===========================================================================

def _add_textbox(slide, left, top, width, height, text, *,
                 size=None, bold=False, color=None, align="left",
                 vertical_anchor=None, font_name=None):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    if vertical_anchor is not None:
        tf.vertical_anchor = vertical_anchor
    # Inner margins
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    para = tf.paragraphs[0]
    para.alignment = {
        "left": PP_ALIGN.LEFT,
        "center": PP_ALIGN.CENTER,
        "right": PP_ALIGN.RIGHT,
    }[align]
    run = para.add_run()
    run.text = text
    if size is not None:
        run.font.size = size
    run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color
    if font_name:
        run.font.name = font_name
    return tb


def _add_bullets(slide, left, top, width, height, bullets, *,
                 size=FONT_BODY, color=C_BODY, bullet_char="•"):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    for i, b in enumerate(bullets):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = PP_ALIGN.LEFT
        para.line_spacing = 1.3
        run = para.add_run()
        run.text = f"{bullet_char}  {b}"
        run.font.size = size
        run.font.color.rgb = color


def _add_rect(slide, left, top, width, height, fill_color, line_color=None):
    rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    rect.fill.solid()
    rect.fill.fore_color.rgb = fill_color
    if line_color is None:
        rect.line.fill.background()
    else:
        rect.line.color.rgb = line_color
    return rect


def _add_title_bar(slide, title_text):
    """Render the standard title bar with a thin underline."""
    _add_textbox(slide, MARGIN, MARGIN, SLIDE_WIDTH - 2 * MARGIN, TITLE_HEIGHT,
                 title_text, size=FONT_TITLE, bold=True, color=C_TITLE)
    # Thin underline
    underline = _add_rect(slide, MARGIN, MARGIN + TITLE_HEIGHT - Inches(0.05),
                          Inches(2.0), Inches(0.04), C_TITLE)


def _add_subtitle_bar(slide, subtitle_text, top):
    """Render the '1 message' subtitle bar."""
    _add_textbox(slide, MARGIN, top, SLIDE_WIDTH - 2 * MARGIN, SUBTITLE_HEIGHT,
                 subtitle_text, size=FONT_SUBTITLE, bold=True, color=C_ACCENT)


def _add_footer(slide, footer_text):
    if not footer_text:
        return
    _add_textbox(
        slide, MARGIN,
        SLIDE_HEIGHT - FOOTER_HEIGHT - Inches(0.1),
        SLIDE_WIDTH - 2 * MARGIN, FOOTER_HEIGHT,
        "出典: " + footer_text,
        size=FONT_FOOTER, color=C_FOOTER,
    )


def _add_speaker_note(slide, note_text):
    if not note_text:
        return
    nf = slide.notes_slide.notes_text_frame
    nf.text = note_text


# ===========================================================================
# Slide builders
# ===========================================================================

def _build_cover(pres, deck):
    """Cover slide with bold company code/name + recommendation."""
    layout = pres.slide_layouts[6]  # blank
    slide = pres.slides.add_slide(layout)

    # Decorative top band
    _add_rect(slide, 0, 0, SLIDE_WIDTH, Inches(0.4), C_TITLE)

    # Title
    _add_textbox(slide, MARGIN, Inches(2.3), SLIDE_WIDTH - 2 * MARGIN,
                 Inches(1.6), deck.title, size=FONT_COVER_TITLE, bold=True,
                 color=C_TITLE, align="center")

    # Subtitle (multi-line: 推奨 / 目標株価 / 投資期間)
    for i, line in enumerate(deck.subtitle.split("\n")):
        _add_textbox(slide, MARGIN, Inches(4.3 + i * 0.5),
                     SLIDE_WIDTH - 2 * MARGIN, Inches(0.5),
                     line, size=FONT_COVER_SUBTITLE, color=C_BODY, align="center")

    # Presenter at bottom
    _add_textbox(slide, MARGIN, Inches(6.5),
                 SLIDE_WIDTH - 2 * MARGIN, Inches(0.5),
                 deck.presenter, size=Pt(18), color=C_FOOTER, align="center")

    # Decorative bottom band
    _add_rect(slide, 0, SLIDE_HEIGHT - Inches(0.4), SLIDE_WIDTH, Inches(0.4),
              C_TITLE)
    return slide


def _build_chapter(pres, slide_def):
    layout = pres.slide_layouts[6]
    slide = pres.slides.add_slide(layout)
    # Full-bleed background
    _add_rect(slide, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT, C_DIVIDER_BG)
    # Large centered title in white
    _add_textbox(slide, MARGIN, Inches(2.8), SLIDE_WIDTH - 2 * MARGIN,
                 Inches(1.8), slide_def.title, size=FONT_CHAPTER_TITLE, bold=True,
                 color=C_WHITE, align="center")
    if slide_def.subtitle:
        _add_textbox(slide, MARGIN, Inches(4.7), SLIDE_WIDTH - 2 * MARGIN,
                     Inches(1.0), slide_def.subtitle, size=Pt(24),
                     color=C_WHITE, align="center")
    return slide


def _build_content(pres, slide_def):
    layout = pres.slide_layouts[6]
    slide = pres.slides.add_slide(layout)
    _add_title_bar(slide, slide_def.title)

    body_top = MARGIN + TITLE_HEIGHT + GAP
    if slide_def.subtitle:
        _add_subtitle_bar(slide, slide_def.subtitle, body_top)
        body_top = body_top + SUBTITLE_HEIGHT + GAP

    if slide_def.bullets:
        body_height = (SLIDE_HEIGHT - body_top - FOOTER_HEIGHT
                       - Inches(0.2))
        _add_bullets(slide, MARGIN + Inches(0.2), body_top,
                     SLIDE_WIDTH - 2 * MARGIN - Inches(0.4),
                     body_height, slide_def.bullets, size=FONT_BODY)

    _add_footer(slide, slide_def.footer)
    _add_speaker_note(slide, slide_def.note)
    return slide


def _build_toc(pres, slide_def):
    # Same layout but emphasized as a navigation slide
    layout = pres.slide_layouts[6]
    slide = pres.slides.add_slide(layout)
    _add_title_bar(slide, slide_def.title)
    body_top = MARGIN + TITLE_HEIGHT + GAP
    if slide_def.subtitle:
        _add_subtitle_bar(slide, slide_def.subtitle, body_top)
        body_top = body_top + SUBTITLE_HEIGHT + GAP
    if slide_def.bullets:
        body_height = (SLIDE_HEIGHT - body_top - FOOTER_HEIGHT - Inches(0.2))
        _add_bullets(slide, MARGIN + Inches(0.4), body_top,
                     SLIDE_WIDTH - 2 * MARGIN - Inches(0.8),
                     body_height, slide_def.bullets, size=Pt(22),
                     color=C_TITLE, bullet_char="▶")
    return slide


def _style_table(tbl, n_rows, n_cols, has_headers):
    """Apply header coloring, row striping, alignment per slide_design.md."""
    for r in range(n_rows):
        is_header = has_headers and r == 0
        for c in range(n_cols):
            cell = tbl.cell(r, c)
            cell.margin_left = Inches(0.08)
            cell.margin_right = Inches(0.08)
            cell.margin_top = Inches(0.04)
            cell.margin_bottom = Inches(0.04)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE

            tf = cell.text_frame
            for p in tf.paragraphs:
                if is_header:
                    p.alignment = PP_ALIGN.CENTER
                else:
                    # Right-align numeric-looking columns (c > 0)
                    p.alignment = PP_ALIGN.LEFT if c == 0 else PP_ALIGN.LEFT
                for run in p.runs:
                    if is_header:
                        run.font.bold = True
                        run.font.size = FONT_TABLE_HEADER
                        run.font.color.rgb = C_WHITE
                    else:
                        run.font.size = FONT_TABLE_CELL
                        run.font.color.rgb = C_BODY

            # Cell fill
            if is_header:
                cell.fill.solid()
                cell.fill.fore_color.rgb = C_HEADER_BG
            else:
                data_row_idx = r - (1 if has_headers else 0)
                if data_row_idx % 2 == 1:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = C_STRIPE
                else:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = C_WHITE


def _build_table(pres, slide_def):
    layout = pres.slide_layouts[6]
    slide = pres.slides.add_slide(layout)
    _add_title_bar(slide, slide_def.title)

    body_top = MARGIN + TITLE_HEIGHT + GAP
    if slide_def.subtitle:
        _add_subtitle_bar(slide, slide_def.subtitle, body_top)
        body_top = body_top + SUBTITLE_HEIGHT + GAP

    rows = slide_def.table
    headers = slide_def.table_headers
    has_headers = bool(headers)
    n_rows = len(rows) + (1 if has_headers else 0)
    n_cols = max((len(r) for r in rows), default=0)
    if has_headers:
        n_cols = max(n_cols, len(headers))
    if n_rows == 0 or n_cols == 0:
        return slide

    available_height = SLIDE_HEIGHT - body_top - FOOTER_HEIGHT - Inches(0.3)
    row_h = min(Inches(0.5), available_height / n_rows)
    table_height = row_h * n_rows
    table_width = SLIDE_WIDTH - 2 * MARGIN

    tbl_shape = slide.shapes.add_table(
        n_rows, n_cols, MARGIN, body_top, table_width, table_height,
    )
    tbl = tbl_shape.table

    # Fill content
    if has_headers:
        for j, h in enumerate(headers):
            tbl.cell(0, j).text = str(h)
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            tbl.cell(i + (1 if has_headers else 0), j).text = (
                str(val) if val is not None else ""
            )

    _style_table(tbl, n_rows, n_cols, has_headers)

    _add_footer(slide, slide_def.footer)
    _add_speaker_note(slide, slide_def.note)
    return slide


def _build_image_placeholder(pres, slide_def):
    layout = pres.slide_layouts[6]
    slide = pres.slides.add_slide(layout)
    _add_title_bar(slide, slide_def.title)
    body_top = MARGIN + TITLE_HEIGHT + GAP
    if slide_def.subtitle:
        _add_subtitle_bar(slide, slide_def.subtitle, body_top)
        body_top = body_top + SUBTITLE_HEIGHT + GAP

    # Large bordered placeholder box
    box_left = Inches(1.5)
    box_top = body_top
    box_width = SLIDE_WIDTH - Inches(3.0)
    box_height = (SLIDE_HEIGHT - box_top - FOOTER_HEIGHT - Inches(0.5))

    rect = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, box_left, box_top, box_width, box_height,
    )
    rect.fill.solid()
    rect.fill.fore_color.rgb = RGBColor(0xFA, 0xFA, 0xFA)
    rect.line.color.rgb = RGBColor(0xBB, 0xBB, 0xBB)

    # Placeholder text inside
    tf = rect.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.3)
    tf.margin_right = Inches(0.3)
    tf.margin_top = Inches(0.4)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE

    # First line: [画像] label
    p1 = tf.paragraphs[0]
    p1.alignment = PP_ALIGN.CENTER
    r1 = p1.add_run()
    r1.text = "[ 画像 ]"
    r1.font.size = Pt(28)
    r1.font.bold = True
    r1.font.color.rgb = C_HELPER

    # Second line: detailed instruction
    p2 = tf.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    p2.line_spacing = 1.3
    r2 = p2.add_run()
    r2.text = "\n" + slide_def.image_placeholder
    r2.font.size = FONT_BODY
    r2.font.color.rgb = C_BODY

    if slide_def.image_caption:
        p3 = tf.add_paragraph()
        p3.alignment = PP_ALIGN.CENTER
        r3 = p3.add_run()
        r3.text = "\n" + slide_def.image_caption
        r3.font.size = Pt(14)
        r3.font.color.rgb = C_HELPER
        r3.font.italic = True

    _add_footer(slide, slide_def.footer)
    _add_speaker_note(slide, slide_def.note)
    return slide


def _build_ascii(pres, slide_def):
    layout = pres.slide_layouts[6]
    slide = pres.slides.add_slide(layout)
    _add_title_bar(slide, slide_def.title)
    body_top = MARGIN + TITLE_HEIGHT + GAP
    if slide_def.subtitle:
        _add_subtitle_bar(slide, slide_def.subtitle, body_top)
        body_top = body_top + SUBTITLE_HEIGHT + GAP

    # Monospaced ASCII area
    box_left = MARGIN + Inches(0.3)
    box_top = body_top
    box_width = SLIDE_WIDTH - 2 * MARGIN - Inches(0.6)
    box_height = SLIDE_HEIGHT - box_top - FOOTER_HEIGHT - Inches(0.3)

    tb = slide.shapes.add_textbox(box_left, box_top, box_width, box_height)
    tf = tb.text_frame
    tf.word_wrap = False
    tf.margin_left = Inches(0.1)
    tf.margin_top = Inches(0.1)

    p = tf.paragraphs[0]
    p.line_spacing = 1.2
    run = p.add_run()
    run.text = slide_def.ascii_art
    run.font.size = FONT_ASCII
    run.font.name = "Courier New"
    run.font.color.rgb = C_BODY

    _add_footer(slide, slide_def.footer)
    _add_speaker_note(slide, slide_def.note)
    return slide


BUILDERS = {
    "chapter": _build_chapter,
    "content": _build_content,
    "toc": _build_toc,
    "table": _build_table,
    "image_placeholder": _build_image_placeholder,
    "ascii": _build_ascii,
}


# ===========================================================================
# Public API
# ===========================================================================

def render(config: SlideDeckConfig, output_path: str) -> list:
    """Render the deck to .pptx. Returns warning strings."""
    pres = Presentation()
    pres.slide_width = SLIDE_WIDTH
    pres.slide_height = SLIDE_HEIGHT

    warnings: list = []

    _build_cover(pres, config)

    for i, slide_def in enumerate(config.slides):
        builder = BUILDERS.get(slide_def.kind)
        if builder is None:
            warnings.append(f"slide {i}: unknown kind '{slide_def.kind}'")
            continue
        builder(pres, slide_def)

        # Validation: design-quality checks
        if slide_def.kind == "content" and len(slide_def.bullets) > 5:
            warnings.append(
                f"slide {i} '{slide_def.title}': has {len(slide_def.bullets)} "
                "bullets — split or trim (>5 too dense per slide_design.md)"
            )
        # Excessive content text
        total_body_chars = sum(len(str(b)) for b in slide_def.bullets)
        if total_body_chars > 250:
            warnings.append(
                f"slide {i} '{slide_def.title}': body has {total_body_chars} "
                "chars — consider trimming (>250 too dense)"
            )
        # Image placeholder must have specific instructions
        if slide_def.kind == "image_placeholder":
            if len(slide_def.image_placeholder) < 25:
                warnings.append(
                    f"slide {i} '{slide_def.title}': image_placeholder is too "
                    "vague — write what / from where / size hint (per slide_design.md)"
                )
        # Encourage subtitle on content/table slides
        if slide_def.kind in ("content", "table") and not slide_def.subtitle:
            warnings.append(
                f"slide {i} '{slide_def.title}': no subtitle (1メッセージ) — "
                "consider adding one"
            )

    pres.save(output_path)

    if warnings:
        print("\n=== Visual-design warnings ===")
        for w in warnings:
            print(f"  - {w}")
    return warnings
