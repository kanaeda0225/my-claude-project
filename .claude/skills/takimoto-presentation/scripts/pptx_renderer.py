"""Reusable .pptx renderer enforcing 瀧本ゼミ slide style + 一般的なスライド設計原則.

Input: a SlideDeckConfig containing a list of Slide objects.
Output: a .pptx file importable to Google Slides.

Key design (slide-master-like uniformity):
- Every content/table/ascii/image slide uses the SAME fixed Y coordinates
- Title always at Y=0.4, height=1.0in, vertical_anchor=MIDDLE (so wrapped
  titles stay centered in the same band)
- Underline always at Y=1.5in (bottom of title band)
- Subtitle slot reserved at Y=1.7in even when empty (body Y is constant)
- Body always at Y=2.4in
- Footer always at Y=7.05in
- All textboxes have auto_size=NONE to prevent textbox-resize-on-content

This produces visual consistency akin to using a real PowerPoint slide master.
"""
from __future__ import annotations

from dataclasses import dataclass, field

from pptx import Presentation
from pptx.util import Pt, Inches, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE


# ===========================================================================
# Layout constants — FIXED ACROSS ALL SLIDES (slide-master discipline)
# ===========================================================================
SLIDE_WIDTH = Inches(13.333)  # 16:9 widescreen
SLIDE_HEIGHT = Inches(7.5)

MARGIN_X = Inches(0.5)

# Title band: y=0.4 to y=1.4, vertical_anchor=MIDDLE
TITLE_TOP = Inches(0.4)
TITLE_HEIGHT = Inches(1.0)

# Underline at the bottom of title band
UNDERLINE_TOP = Inches(1.42)
UNDERLINE_HEIGHT = Inches(0.04)
UNDERLINE_WIDTH = Inches(2.4)

# Subtitle band: y=1.65 to y=2.25 (reserved even when empty)
SUBTITLE_TOP = Inches(1.65)
SUBTITLE_HEIGHT = Inches(0.55)

# Body: y=2.35 to y=6.85 (fixed across all slides)
BODY_TOP = Inches(2.35)
BODY_BOTTOM = Inches(6.85)
BODY_HEIGHT = BODY_BOTTOM - BODY_TOP

# Footer: y=7.05 to y=7.4
FOOTER_TOP = Inches(7.05)
FOOTER_HEIGHT = Inches(0.35)

# Font sizes (minimums, never go below — see reference/slide_design.md)
FONT_TITLE = Pt(28)
FONT_SUBTITLE = Pt(20)
FONT_BODY = Pt(18)
FONT_TABLE_HEADER = Pt(15)
FONT_TABLE_CELL = Pt(14)
FONT_FOOTER = Pt(11)
FONT_ASCII = Pt(15)
FONT_COVER_TITLE = Pt(54)
FONT_COVER_SUBTITLE = Pt(26)
FONT_CHAPTER_TITLE = Pt(58)

# Color palette
C_TITLE = RGBColor(0x1F, 0x4E, 0x78)
C_ACCENT = RGBColor(0xC0, 0x39, 0x2B)
C_POSITIVE = RGBColor(0x2E, 0x7D, 0x32)
C_BODY = RGBColor(0x20, 0x20, 0x20)
C_HELPER = RGBColor(0x66, 0x66, 0x66)
C_FOOTER = RGBColor(0x80, 0x80, 0x80)
C_HEADER_BG = RGBColor(0x1F, 0x4E, 0x78)
C_STRIPE = RGBColor(0xF5, 0xF5, 0xF5)
C_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
C_DIVIDER_BG = RGBColor(0x1F, 0x4E, 0x78)


@dataclass
class Slide:
    kind: str
    title: str = ""
    subtitle: str = ""
    bullets: list = field(default_factory=list)
    table: list = field(default_factory=list)
    table_headers: list = field(default_factory=list)
    ascii_art: str = ""
    image_placeholder: str = ""
    image_caption: str = ""
    footer: str = ""
    note: str = ""


@dataclass
class SlideDeckConfig:
    title: str
    subtitle: str
    presenter: str
    slides: list = field(default_factory=list)


# ===========================================================================
# Low-level helpers — all enforce auto_size=NONE for fixed positioning
# ===========================================================================

def _add_textbox(slide, left, top, width, height, text, *,
                 size=None, bold=False, color=None, align="left",
                 vertical_anchor=MSO_ANCHOR.TOP, font_name=None,
                 italic=False):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    # Disable auto-resize so the textbox stays at the specified coordinates
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.word_wrap = True
    tf.vertical_anchor = vertical_anchor
    tf.margin_left = Emu(0)
    tf.margin_right = Emu(0)
    tf.margin_top = Emu(0)
    tf.margin_bottom = Emu(0)
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
    run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color
    if font_name:
        run.font.name = font_name
    return tb


def _add_bullets(slide, left, top, width, height, bullets, *,
                 size=FONT_BODY, color=C_BODY, bullet_char="•"):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.margin_left = Emu(0)
    tf.margin_right = Emu(0)
    tf.margin_top = Emu(0)
    tf.margin_bottom = Emu(0)
    for i, b in enumerate(bullets):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = PP_ALIGN.LEFT
        para.line_spacing = 1.35
        para.space_after = Pt(6)
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


# ===========================================================================
# Slide-master-like fixed regions
# ===========================================================================

def _add_title_region(slide, title_text):
    """Render the title at the FIXED title band position.

    Uses vertical_anchor=MIDDLE so wrapped titles still occupy the same band.
    """
    _add_textbox(
        slide, MARGIN_X, TITLE_TOP,
        SLIDE_WIDTH - 2 * MARGIN_X, TITLE_HEIGHT,
        title_text, size=FONT_TITLE, bold=True, color=C_TITLE,
        vertical_anchor=MSO_ANCHOR.MIDDLE,
    )
    # Thin underline always at the same position
    _add_rect(slide, MARGIN_X, UNDERLINE_TOP,
              UNDERLINE_WIDTH, UNDERLINE_HEIGHT, C_TITLE)


def _add_subtitle_region(slide, subtitle_text):
    """Render the 1-message subtitle at the FIXED subtitle band.

    Even when subtitle_text is empty, we don't add anything — but the body
    region below is at a fixed Y regardless, so positions stay consistent.
    """
    if not subtitle_text:
        return
    _add_textbox(
        slide, MARGIN_X, SUBTITLE_TOP,
        SLIDE_WIDTH - 2 * MARGIN_X, SUBTITLE_HEIGHT,
        subtitle_text, size=FONT_SUBTITLE, bold=True, color=C_ACCENT,
        vertical_anchor=MSO_ANCHOR.MIDDLE,
    )


def _add_footer_region(slide, footer_text):
    """Render the source citation footer at the FIXED footer band."""
    if not footer_text:
        return
    _add_textbox(
        slide, MARGIN_X, FOOTER_TOP,
        SLIDE_WIDTH - 2 * MARGIN_X, FOOTER_HEIGHT,
        "出典: " + footer_text,
        size=FONT_FOOTER, color=C_FOOTER, italic=True,
        vertical_anchor=MSO_ANCHOR.MIDDLE,
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
    layout = pres.slide_layouts[6]
    slide = pres.slides.add_slide(layout)
    # Top deco band
    _add_rect(slide, 0, 0, SLIDE_WIDTH, Inches(0.4), C_TITLE)
    # Title (centered)
    _add_textbox(
        slide, MARGIN_X, Inches(2.3),
        SLIDE_WIDTH - 2 * MARGIN_X, Inches(1.6),
        deck.title, size=FONT_COVER_TITLE, bold=True,
        color=C_TITLE, align="center",
    )
    # Subtitle (multi-line)
    for i, line in enumerate(deck.subtitle.split("\n")):
        _add_textbox(
            slide, MARGIN_X, Inches(4.3 + i * 0.5),
            SLIDE_WIDTH - 2 * MARGIN_X, Inches(0.5),
            line, size=FONT_COVER_SUBTITLE, color=C_BODY, align="center",
        )
    _add_textbox(
        slide, MARGIN_X, Inches(6.5),
        SLIDE_WIDTH - 2 * MARGIN_X, Inches(0.5),
        deck.presenter, size=Pt(18), color=C_FOOTER, align="center",
    )
    # Bottom deco band
    _add_rect(slide, 0, SLIDE_HEIGHT - Inches(0.4), SLIDE_WIDTH,
              Inches(0.4), C_TITLE)
    return slide


def _build_chapter(pres, slide_def):
    """Full-bleed chapter divider."""
    layout = pres.slide_layouts[6]
    slide = pres.slides.add_slide(layout)
    _add_rect(slide, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT, C_DIVIDER_BG)
    _add_textbox(
        slide, MARGIN_X, Inches(2.8),
        SLIDE_WIDTH - 2 * MARGIN_X, Inches(1.8),
        slide_def.title, size=FONT_CHAPTER_TITLE, bold=True,
        color=C_WHITE, align="center",
        vertical_anchor=MSO_ANCHOR.MIDDLE,
    )
    if slide_def.subtitle:
        _add_textbox(
            slide, MARGIN_X, Inches(4.7),
            SLIDE_WIDTH - 2 * MARGIN_X, Inches(1.0),
            slide_def.subtitle, size=Pt(24),
            color=C_WHITE, align="center",
        )
    return slide


def _build_content(pres, slide_def):
    layout = pres.slide_layouts[6]
    slide = pres.slides.add_slide(layout)
    _add_title_region(slide, slide_def.title)
    _add_subtitle_region(slide, slide_def.subtitle)
    if slide_def.bullets:
        _add_bullets(slide, MARGIN_X + Inches(0.2), BODY_TOP,
                     SLIDE_WIDTH - 2 * MARGIN_X - Inches(0.4),
                     BODY_HEIGHT, slide_def.bullets, size=FONT_BODY)
    _add_footer_region(slide, slide_def.footer)
    _add_speaker_note(slide, slide_def.note)
    return slide


def _build_toc(pres, slide_def):
    layout = pres.slide_layouts[6]
    slide = pres.slides.add_slide(layout)
    _add_title_region(slide, slide_def.title)
    _add_subtitle_region(slide, slide_def.subtitle)
    if slide_def.bullets:
        _add_bullets(slide, MARGIN_X + Inches(0.4), BODY_TOP,
                     SLIDE_WIDTH - 2 * MARGIN_X - Inches(0.8),
                     BODY_HEIGHT, slide_def.bullets, size=Pt(22),
                     color=C_TITLE, bullet_char="▶")
    return slide


def _style_table(tbl, n_rows, n_cols, has_headers):
    """Apply header coloring, row striping, alignment."""
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
                    p.alignment = PP_ALIGN.LEFT
                for run in p.runs:
                    if is_header:
                        run.font.bold = True
                        run.font.size = FONT_TABLE_HEADER
                        run.font.color.rgb = C_WHITE
                    else:
                        run.font.size = FONT_TABLE_CELL
                        run.font.color.rgb = C_BODY
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
    _add_title_region(slide, slide_def.title)
    _add_subtitle_region(slide, slide_def.subtitle)

    rows = slide_def.table
    headers = slide_def.table_headers
    has_headers = bool(headers)
    n_rows = len(rows) + (1 if has_headers else 0)
    n_cols = max((len(r) for r in rows), default=0)
    if has_headers:
        n_cols = max(n_cols, len(headers))
    if n_rows == 0 or n_cols == 0:
        return slide

    table_width = SLIDE_WIDTH - 2 * MARGIN_X
    row_h = min(Inches(0.5), BODY_HEIGHT / n_rows)
    table_height = row_h * n_rows

    tbl_shape = slide.shapes.add_table(
        n_rows, n_cols, MARGIN_X, BODY_TOP, table_width, table_height,
    )
    tbl = tbl_shape.table
    if has_headers:
        for j, h in enumerate(headers):
            tbl.cell(0, j).text = str(h)
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            tbl.cell(i + (1 if has_headers else 0), j).text = (
                str(val) if val is not None else ""
            )
    _style_table(tbl, n_rows, n_cols, has_headers)
    _add_footer_region(slide, slide_def.footer)
    _add_speaker_note(slide, slide_def.note)
    return slide


def _build_image_placeholder(pres, slide_def):
    layout = pres.slide_layouts[6]
    slide = pres.slides.add_slide(layout)
    _add_title_region(slide, slide_def.title)
    _add_subtitle_region(slide, slide_def.subtitle)

    # Large bordered placeholder box at body region
    box_left = Inches(1.5)
    box_top = BODY_TOP
    box_width = SLIDE_WIDTH - Inches(3.0)
    box_height = BODY_BOTTOM - BODY_TOP - Inches(0.2)

    rect = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, box_left, box_top, box_width, box_height,
    )
    rect.fill.solid()
    rect.fill.fore_color.rgb = RGBColor(0xFA, 0xFA, 0xFA)
    rect.line.color.rgb = RGBColor(0xBB, 0xBB, 0xBB)

    tf = rect.text_frame
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.word_wrap = True
    tf.margin_left = Inches(0.3)
    tf.margin_right = Inches(0.3)
    tf.margin_top = Inches(0.4)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE

    p1 = tf.paragraphs[0]
    p1.alignment = PP_ALIGN.CENTER
    r1 = p1.add_run()
    r1.text = "[ 画像 ]"
    r1.font.size = Pt(28)
    r1.font.bold = True
    r1.font.color.rgb = C_HELPER

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

    _add_footer_region(slide, slide_def.footer)
    _add_speaker_note(slide, slide_def.note)
    return slide


def _build_ascii(pres, slide_def):
    layout = pres.slide_layouts[6]
    slide = pres.slides.add_slide(layout)
    _add_title_region(slide, slide_def.title)
    _add_subtitle_region(slide, slide_def.subtitle)

    tb = slide.shapes.add_textbox(
        MARGIN_X + Inches(0.3), BODY_TOP,
        SLIDE_WIDTH - 2 * MARGIN_X - Inches(0.6),
        BODY_HEIGHT,
    )
    tf = tb.text_frame
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.word_wrap = False
    tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.margin_left = Inches(0.1)
    tf.margin_top = Inches(0.1)

    p = tf.paragraphs[0]
    p.line_spacing = 1.2
    run = p.add_run()
    run.text = slide_def.ascii_art
    run.font.size = FONT_ASCII
    run.font.name = "Courier New"
    run.font.color.rgb = C_BODY

    _add_footer_region(slide, slide_def.footer)
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

        if slide_def.kind == "content" and len(slide_def.bullets) > 5:
            warnings.append(
                f"slide {i} '{slide_def.title}': has {len(slide_def.bullets)} "
                "bullets — split or trim (>5 too dense)"
            )
        total_body_chars = sum(len(str(b)) for b in slide_def.bullets)
        if total_body_chars > 250:
            warnings.append(
                f"slide {i} '{slide_def.title}': body has {total_body_chars} "
                "chars — consider trimming (>250 too dense)"
            )
        if slide_def.kind == "image_placeholder":
            if len(slide_def.image_placeholder) < 25:
                warnings.append(
                    f"slide {i} '{slide_def.title}': image_placeholder too "
                    "vague — describe what / from where / size hint"
                )
        if slide_def.kind in ("content", "table") and not slide_def.subtitle:
            warnings.append(
                f"slide {i} '{slide_def.title}': no subtitle (1メッセージ)"
            )
        # Title length check: warn if title might wrap
        if slide_def.kind not in ("chapter",) and len(slide_def.title) > 35:
            warnings.append(
                f"slide {i} '{slide_def.title}': title is long "
                f"({len(slide_def.title)} chars) — may wrap"
            )

    pres.save(output_path)

    if warnings:
        print("\n=== Visual-design warnings ===")
        for w in warnings:
            print(f"  - {w}")
    return warnings
