"""
Script to generate a PowerPoint presentation (.pptx).
Run: python3 create_presentation.py
Requires: pip install python-pptx
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt


# ── colour palette ──────────────────────────────────────────────────────────
DARK_BLUE   = RGBColor(0x1F, 0x39, 0x7A)   # headings / title bg
MID_BLUE    = RGBColor(0x2E, 0x75, 0xB6)   # accent
LIGHT_BLUE  = RGBColor(0xBD, 0xD7, 0xEE)   # subtle bg
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
DARK_GREY   = RGBColor(0x40, 0x40, 0x40)


def set_text(tf, text, font_size=18, bold=False, color=DARK_GREY,
             align=PP_ALIGN.LEFT):
    tf.text = text
    for para in tf.paragraphs:
        para.alignment = align
        for run in para.runs:
            run.font.size  = Pt(font_size)
            run.font.bold  = bold
            run.font.color.rgb = color


def add_bg_rect(slide, fill_color, left=0, top=0,
                width=None, height=None, prs=None):
    """Add a filled rectangle as a slide background / band."""
    if width is None:
        width = prs.slide_width
    if height is None:
        height = prs.slide_height
    shape = slide.shapes.add_shape(
        1,   # MSO_SHAPE_TYPE.RECTANGLE
        left, top, width, height
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()   # no border
    # push to back
    slide.shapes._spTree.remove(shape._element)
    slide.shapes._spTree.insert(2, shape._element)
    return shape


# ── slide builders ──────────────────────────────────────────────────────────

def build_title_slide(prs):
    slide_layout = prs.slide_layouts[6]   # blank
    slide = prs.slides.add_slide(slide_layout)

    # full-bleed dark-blue background
    add_bg_rect(slide, DARK_BLUE, prs=prs)

    # decorative mid-blue band at the bottom
    band_h = Inches(1.2)
    add_bg_rect(slide, MID_BLUE,
                top=prs.slide_height - band_h,
                height=band_h, prs=prs)

    # title
    txBox = slide.shapes.add_textbox(
        Inches(0.8), Inches(1.8), Inches(8.4), Inches(1.6))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = "Presentation Title"
    run.font.size  = Pt(44)
    run.font.bold  = True
    run.font.color.rgb = WHITE

    # subtitle
    txBox2 = slide.shapes.add_textbox(
        Inches(0.8), Inches(3.6), Inches(8.4), Inches(0.9))
    tf2 = txBox2.text_frame
    p2 = tf2.paragraphs[0]
    p2.alignment = PP_ALIGN.CENTER
    run2 = p2.add_run()
    run2.text = "A concise subtitle or author / date"
    run2.font.size  = Pt(22)
    run2.font.color.rgb = LIGHT_BLUE

    return slide


def build_agenda_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg_rect(slide, WHITE, prs=prs)

    # header bar
    add_bg_rect(slide, DARK_BLUE,
                height=Inches(1.1), prs=prs)

    # slide title in header
    txBox = slide.shapes.add_textbox(
        Inches(0.4), Inches(0.15), Inches(9.2), Inches(0.8))
    set_text(txBox.text_frame, "Agenda",
             font_size=32, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

    # agenda items
    items = [
        "1. Introduction",
        "2. Background & Context",
        "3. Key Findings",
        "4. Solution / Proposal",
        "5. Next Steps",
        "6. Q & A",
    ]
    txBox2 = slide.shapes.add_textbox(
        Inches(0.8), Inches(1.4), Inches(8.4), Inches(4.8))
    tf = txBox2.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_before = Pt(8)
        run = p.add_run()
        run.text = item
        run.font.size  = Pt(22)
        run.font.color.rgb = DARK_GREY

    return slide


def build_content_slide(prs, title, bullets):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg_rect(slide, WHITE, prs=prs)
    add_bg_rect(slide, DARK_BLUE, height=Inches(1.1), prs=prs)

    # title
    txBox = slide.shapes.add_textbox(
        Inches(0.4), Inches(0.15), Inches(9.2), Inches(0.8))
    set_text(txBox.text_frame, title,
             font_size=28, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

    # accent bar on the left
    bar = slide.shapes.add_shape(
        1, Inches(0.3), Inches(1.3), Inches(0.07), Inches(4.8))
    bar.fill.solid()
    bar.fill.fore_color.rgb = MID_BLUE
    bar.line.fill.background()

    # bullets
    txBox2 = slide.shapes.add_textbox(
        Inches(0.6), Inches(1.3), Inches(9.0), Inches(4.8))
    tf = txBox2.text_frame
    tf.word_wrap = True
    for i, bullet in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_before = Pt(6)
        run = p.add_run()
        run.text = f"•  {bullet}"
        run.font.size  = Pt(20)
        run.font.color.rgb = DARK_GREY

    return slide


def build_closing_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg_rect(slide, MID_BLUE, prs=prs)

    txBox = slide.shapes.add_textbox(
        Inches(0.8), Inches(2.0), Inches(8.4), Inches(1.6))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = "Thank You!"
    run.font.size  = Pt(54)
    run.font.bold  = True
    run.font.color.rgb = WHITE

    txBox2 = slide.shapes.add_textbox(
        Inches(0.8), Inches(3.8), Inches(8.4), Inches(0.8))
    tf2 = txBox2.text_frame
    p2 = tf2.paragraphs[0]
    p2.alignment = PP_ALIGN.CENTER
    run2 = p2.add_run()
    run2.text = "Questions & Discussion"
    run2.font.size  = Pt(24)
    run2.font.color.rgb = LIGHT_BLUE

    return slide


# ── main ────────────────────────────────────────────────────────────────────

def create_presentation(output_path="presentation.pptx"):
    prs = Presentation()
    # standard 16:9 widescreen
    prs.slide_width  = Inches(10)
    prs.slide_height = Inches(5.625)

    build_title_slide(prs)
    build_agenda_slide(prs)

    build_content_slide(prs,
        title="Background & Context",
        bullets=[
            "Brief overview of the problem or opportunity",
            "Relevant market / industry data",
            "Stakeholders involved",
            "Timeline and constraints",
        ])

    build_content_slide(prs,
        title="Key Findings",
        bullets=[
            "Finding #1 — supported by data or research",
            "Finding #2 — supported by data or research",
            "Finding #3 — supported by data or research",
            "Summary of implications",
        ])

    build_content_slide(prs,
        title="Solution / Proposal",
        bullets=[
            "Proposed approach and rationale",
            "Expected benefits and outcomes",
            "Resource requirements (team, budget, tools)",
            "Risk mitigation strategy",
        ])

    build_content_slide(prs,
        title="Next Steps",
        bullets=[
            "Action item 1 — Owner: TBD, Due: Week 1",
            "Action item 2 — Owner: TBD, Due: Week 2",
            "Action item 3 — Owner: TBD, Due: Week 3",
            "Schedule follow-up meeting",
        ])

    build_closing_slide(prs)

    prs.save(output_path)
    print(f"Presentation saved to: {output_path}")


if __name__ == "__main__":
    create_presentation()
