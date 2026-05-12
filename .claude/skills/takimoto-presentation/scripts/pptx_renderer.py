"""Reusable .pptx renderer enforcing 瀧本ゼミ slide style.

Input: a SlideDeckConfig containing a list of Slide objects.
Output: a .pptx file importable to Google Slides.

Rules enforced:
- 1 message per slide (no auto-validation, but layout is single-column)
- No image generation — image needs are surfaced as `[画像: ...]` text blocks
- Minimal decoration (white bg, simple title, single content area)
- Footer with source/citation on each slide if provided
- Optional 章タイトル スライド between chapters
"""
from __future__ import annotations

from dataclasses import dataclass, field

from pptx import Presentation
from pptx.util import Pt, Inches, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE


# Slide dimensions (16:9 widescreen)
SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)

# Colors
COLOR_TITLE = RGBColor(0x1F, 0x4E, 0x78)  # blue, used for title
COLOR_BODY = RGBColor(0x20, 0x20, 0x20)  # near-black
COLOR_FOOTER = RGBColor(0x80, 0x80, 0x80)  # gray
COLOR_HIGHLIGHT = RGBColor(0xC0, 0x39, 0x2B)  # red accent
COLOR_SECTION_BG = RGBColor(0x1F, 0x4E, 0x78)  # blue full bg for chapter dividers


@dataclass
class Slide:
    """One slide definition.

    `kind` controls layout:
      - "title": cover slide (large title, recommend center)
      - "chapter": full-color chapter divider
      - "content": standard content (title + bullets or table or image_placeholder)
      - "toc": table of contents
      - "table": title + a 2D table
      - "image_placeholder": title + a big [画像: ...] block
      - "ascii": title + monospaced ascii art (for structure diagrams)
    """

    kind: str
    title: str = ""
    subtitle: str = ""  # used by title/chapter slides
    bullets: list = field(default_factory=list)  # content slides
    table: list = field(default_factory=list)  # 2D list of rows, for kind="table"
    table_headers: list = field(default_factory=list)
    ascii_art: str = ""  # monospaced text for kind="ascii"
    image_placeholder: str = ""  # text describing the image to place
    footer: str = ""  # source citation
    note: str = ""  # speaker note (not shown on slide)


@dataclass
class SlideDeckConfig:
    """Configuration for the whole deck."""

    title: str
    subtitle: str  # e.g., 推奨/目標株価/投資期間 multi-line
    presenter: str
    slides: list = field(default_factory=list)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _set_text(frame, text: str, size: int = 18, bold: bool = False, color=None,
              align: str = "left"):
    """Set text on a textframe with given styling."""
    frame.text = ""
    para = frame.paragraphs[0]
    if align == "center":
        para.alignment = PP_ALIGN.CENTER
    run = para.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color
    else:
        run.font.color.rgb = COLOR_BODY


def _add_textbox(slide, left, top, width, height, text, size=18, bold=False,
                 color=None, align="left", word_wrap=True):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = word_wrap
    _set_text(tf, text, size=size, bold=bold, color=color, align=align)
    return tb


def _add_bullets(slide, left, top, width, height, bullets: list, size=18):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    for i, b in enumerate(bullets):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = PP_ALIGN.LEFT
        run = para.add_run()
        run.text = "• " + str(b)
        run.font.size = Pt(size)
        run.font.color.rgb = COLOR_BODY


def _add_footer(slide, text):
    if not text:
        return
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(7.0),
                                  SLIDE_WIDTH - Inches(1.0), Inches(0.4))
    tf = tb.text_frame
    _set_text(tf, "出典: " + text, size=10, color=COLOR_FOOTER)


def _add_note(slide, text):
    if not text:
        return
    nf = slide.notes_slide.notes_text_frame
    nf.text = text


# ---------------------------------------------------------------------------
# Slide builders
# ---------------------------------------------------------------------------

def _build_cover(pres, deck_config):
    blank_layout = pres.slide_layouts[6]
    slide = pres.slides.add_slide(blank_layout)
    # Big title
    _add_textbox(slide, Inches(1.0), Inches(2.5), SLIDE_WIDTH - Inches(2.0),
                 Inches(1.5), deck_config.title, size=48, bold=True,
                 color=COLOR_TITLE, align="center")
    # Subtitle (推奨/目標株価/投資期間)
    _add_textbox(slide, Inches(1.0), Inches(4.2), SLIDE_WIDTH - Inches(2.0),
                 Inches(1.5), deck_config.subtitle, size=24, align="center")
    # Presenter
    _add_textbox(slide, Inches(1.0), Inches(6.0), SLIDE_WIDTH - Inches(2.0),
                 Inches(0.6), deck_config.presenter, size=18,
                 color=COLOR_FOOTER, align="center")
    return slide


def _build_chapter(pres, slide_def):
    """Full-color chapter divider."""
    blank_layout = pres.slide_layouts[6]
    slide = pres.slides.add_slide(blank_layout)
    # Full-bleed background
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0,
                                SLIDE_WIDTH, SLIDE_HEIGHT)
    bg.fill.solid()
    bg.fill.fore_color.rgb = COLOR_SECTION_BG
    bg.line.fill.background()
    # Title centered, white
    _add_textbox(slide, Inches(1.0), Inches(3.0), SLIDE_WIDTH - Inches(2.0),
                 Inches(1.5), slide_def.title, size=54, bold=True,
                 color=RGBColor(0xFF, 0xFF, 0xFF), align="center")
    if slide_def.subtitle:
        _add_textbox(slide, Inches(1.0), Inches(4.7), SLIDE_WIDTH - Inches(2.0),
                     Inches(1.0), slide_def.subtitle, size=22,
                     color=RGBColor(0xFF, 0xFF, 0xFF), align="center")
    return slide


def _build_content(pres, slide_def):
    blank_layout = pres.slide_layouts[6]
    slide = pres.slides.add_slide(blank_layout)
    # Title bar
    _add_textbox(slide, Inches(0.5), Inches(0.3), SLIDE_WIDTH - Inches(1.0),
                 Inches(0.8), slide_def.title, size=28, bold=True,
                 color=COLOR_TITLE)
    # Optional subtitle (1スライド1メッセージの「メッセージ」)
    body_top = Inches(1.3)
    if slide_def.subtitle:
        _add_textbox(slide, Inches(0.5), Inches(1.1),
                     SLIDE_WIDTH - Inches(1.0), Inches(0.6),
                     slide_def.subtitle, size=18, bold=True,
                     color=COLOR_HIGHLIGHT)
        body_top = Inches(1.8)
    # Body bullets
    if slide_def.bullets:
        _add_bullets(slide, Inches(0.7), body_top,
                     SLIDE_WIDTH - Inches(1.4),
                     Inches(5.0), slide_def.bullets, size=18)
    _add_footer(slide, slide_def.footer)
    _add_note(slide, slide_def.note)
    return slide


def _build_table(pres, slide_def):
    blank_layout = pres.slide_layouts[6]
    slide = pres.slides.add_slide(blank_layout)
    _add_textbox(slide, Inches(0.5), Inches(0.3), SLIDE_WIDTH - Inches(1.0),
                 Inches(0.8), slide_def.title, size=28, bold=True,
                 color=COLOR_TITLE)
    if slide_def.subtitle:
        _add_textbox(slide, Inches(0.5), Inches(1.1),
                     SLIDE_WIDTH - Inches(1.0), Inches(0.6),
                     slide_def.subtitle, size=18, bold=True,
                     color=COLOR_HIGHLIGHT)
        table_top = Inches(1.8)
    else:
        table_top = Inches(1.3)
    # Table
    rows = slide_def.table
    headers = slide_def.table_headers
    n_rows = len(rows) + (1 if headers else 0)
    n_cols = max(len(r) for r in rows) if rows else len(headers)
    if n_rows == 0 or n_cols == 0:
        return slide
    table_width = SLIDE_WIDTH - Inches(1.0)
    table_height = Inches(min(0.4 * n_rows, 5.0))
    tbl_shape = slide.shapes.add_table(n_rows, n_cols, Inches(0.5),
                                       table_top, table_width, table_height)
    tbl = tbl_shape.table
    r_offset = 0
    if headers:
        for j, h in enumerate(headers):
            cell = tbl.cell(0, j)
            cell.text = str(h)
            for p in cell.text_frame.paragraphs:
                for run in p.runs:
                    run.font.bold = True
                    run.font.size = Pt(14)
        r_offset = 1
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = tbl.cell(i + r_offset, j)
            cell.text = str(val) if val is not None else ""
            for p in cell.text_frame.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(13)
    _add_footer(slide, slide_def.footer)
    _add_note(slide, slide_def.note)
    return slide


def _build_image_placeholder(pres, slide_def):
    blank_layout = pres.slide_layouts[6]
    slide = pres.slides.add_slide(blank_layout)
    _add_textbox(slide, Inches(0.5), Inches(0.3), SLIDE_WIDTH - Inches(1.0),
                 Inches(0.8), slide_def.title, size=28, bold=True,
                 color=COLOR_TITLE)
    # Big bordered placeholder box
    placeholder = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(1.5), Inches(1.8),
        SLIDE_WIDTH - Inches(3.0), Inches(4.5))
    placeholder.fill.solid()
    placeholder.fill.fore_color.rgb = RGBColor(0xF5, 0xF5, 0xF5)
    placeholder.line.color.rgb = RGBColor(0x80, 0x80, 0x80)
    tf = placeholder.text_frame
    tf.word_wrap = True
    _set_text(tf, "[画像]\n\n" + slide_def.image_placeholder,
              size=20, bold=True, color=COLOR_FOOTER, align="center")
    _add_footer(slide, slide_def.footer)
    _add_note(slide, slide_def.note)
    return slide


def _build_ascii(pres, slide_def):
    blank_layout = pres.slide_layouts[6]
    slide = pres.slides.add_slide(blank_layout)
    _add_textbox(slide, Inches(0.5), Inches(0.3), SLIDE_WIDTH - Inches(1.0),
                 Inches(0.8), slide_def.title, size=28, bold=True,
                 color=COLOR_TITLE)
    # Monospace ASCII art
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(1.3),
                                  SLIDE_WIDTH - Inches(1.0), Inches(5.5))
    tf = tb.text_frame
    tf.word_wrap = False
    para = tf.paragraphs[0]
    run = para.add_run()
    run.text = slide_def.ascii_art
    run.font.size = Pt(14)
    run.font.name = "Courier New"
    run.font.color.rgb = COLOR_BODY
    _add_footer(slide, slide_def.footer)
    _add_note(slide, slide_def.note)
    return slide


def _build_toc(pres, slide_def):
    return _build_content(pres, slide_def)  # same layout, just bulleted chapters


BUILDERS = {
    "title": lambda p, s: None,  # handled separately
    "chapter": _build_chapter,
    "content": _build_content,
    "toc": _build_toc,
    "table": _build_table,
    "image_placeholder": _build_image_placeholder,
    "ascii": _build_ascii,
}


# ---------------------------------------------------------------------------
# Entry
# ---------------------------------------------------------------------------

def render(config: SlideDeckConfig, output_path: str) -> list:
    """Render the deck to .pptx. Returns a list of warning strings."""
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

        # Validation: 1 message per slide
        if slide_def.kind == "content" and len(slide_def.bullets) > 7:
            warnings.append(
                f"slide {i} '{slide_def.title}': has {len(slide_def.bullets)} "
                "bullets — consider splitting (>7 cramped)"
            )

    pres.save(output_path)

    if warnings:
        print("\n=== Warnings ===")
        for w in warnings:
            print(f"  - {w}")
    return warnings
