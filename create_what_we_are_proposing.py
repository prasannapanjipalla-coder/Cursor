#!/usr/bin/env python3
"""Build a native (editable) PowerPoint slide for 'What We Are Proposing'.

Recreates the consulting MVP proposal slide with shapes, connectors, and
text frames — not an embedded screenshot.
"""

from __future__ import annotations

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

# --- Slide geometry (widescreen 16:9) ---
SLIDE_W = Inches(13.333333)
SLIDE_H = Inches(7.5)

# --- Palette (PwC-style) ---
ORANGE = RGBColor(0xD0, 0x4A, 0x02)
ORANGE_DARK = RGBColor(0xB8, 0x3E, 0x00)
ORANGE_MID = RGBColor(0xE8, 0x8A, 0x52)
ORANGE_BAR = RGBColor(0xF3, 0xC4, 0xA8)
ORANGE_SOFT = RGBColor(0xFB, 0xE8, 0xDC)
ORANGE_WASH = RGBColor(0xFD, 0xF3, 0xEC)
BLACK = RGBColor(0x00, 0x00, 0x00)
NEAR_BLACK = RGBColor(0x1A, 0x1A, 0x1A)
DARK_GRAY = RGBColor(0x4A, 0x4A, 0x4A)
MID_GRAY = RGBColor(0x6B, 0x6B, 0x6B)
LABEL_GRAY = RGBColor(0x5C, 0x5C, 0x5C)
LINE_GRAY = RGBColor(0xB0, 0xB0, 0xB0)
BOX_GRAY = RGBColor(0xE8, 0xE8, 0xE8)
CARD_BORDER = RGBColor(0xE4, 0xE4, 0xE4)
LIGHT_GRAY = RGBColor(0xF3, 0xF3, 0xF3)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

SERIF = "Georgia"
SANS = "Calibri"


def _in(x: float) -> int:
    return int(Inches(x))


def _set_fill(shape, color: RGBColor | None) -> None:
    if color is None:
        shape.fill.background()
    else:
        shape.fill.solid()
        shape.fill.fore_color.rgb = color


def _set_line(shape, color: RGBColor | None, width_pt: float = 1.0, dash=None) -> None:
    if color is None:
        shape.line.fill.background()
        return
    shape.line.color.rgb = color
    shape.line.width = Pt(width_pt)
    if dash is not None:
        shape.line.dash_style = dash


def add_rect(slide, l, t, w, h, fill, line=None, line_w=1.0, rounded=False, adj=0.08):
    kind = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    sh = slide.shapes.add_shape(kind, _in(l), _in(t), _in(w), _in(h))
    _set_fill(sh, fill)
    _set_line(sh, line, line_w)
    if rounded:
        try:
            sh.adjustments[0] = adj
        except Exception:
            pass
    sh.shadow.inherit = False
    return sh


def add_oval(slide, l, t, w, h, fill, line=None, line_w=1.0):
    sh = slide.shapes.add_shape(MSO_SHAPE.OVAL, _in(l), _in(t), _in(w), _in(h))
    _set_fill(sh, fill)
    _set_line(sh, line, line_w)
    sh.shadow.inherit = False
    return sh


def add_auto(slide, kind, l, t, w, h, fill, line=None, line_w=1.0):
    sh = slide.shapes.add_shape(kind, _in(l), _in(t), _in(w), _in(h))
    _set_fill(sh, fill)
    _set_line(sh, line, line_w)
    sh.shadow.inherit = False
    return sh


def add_line(slide, x1, y1, x2, y2, color, width_pt=2.0, dash=None, arrow=False):
    conn = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, _in(x1), _in(y1), _in(x2), _in(y2)
    )
    conn.line.color.rgb = color
    conn.line.width = Pt(width_pt)
    if dash is not None:
        conn.line.dash_style = dash
    if arrow:
        ln = conn.line._ln
        tail = ln.find(qn("a:tailEnd"))
        if tail is None:
            tail = etree.SubElement(ln, qn("a:tailEnd"))
        tail.set("type", "triangle")
        tail.set("w", "med")
        tail.set("len", "med")
    return conn


def apply_shadow(shape, blur=0.08, dist=0.03, alpha=18000) -> None:
    """Soft drop shadow via DrawingML (cards)."""
    sp_pr = shape._element.spPr
    # Remove existing effectLst if present
    for child in list(sp_pr):
        if child.tag == qn("a:effectLst"):
            sp_pr.remove(child)
    effect_lst = etree.SubElement(sp_pr, qn("a:effectLst"))
    shdw = etree.SubElement(effect_lst, qn("a:outerShdw"))
    shdw.set("blurRad", str(int(Inches(blur))))
    shdw.set("dist", str(int(Inches(dist))))
    shdw.set("dir", "2700000")
    shdw.set("algn", "tl")
    shdw.set("rotWithShape", "0")
    srgb = etree.SubElement(shdw, qn("a:srgbClr"))
    srgb.set("val", "000000")
    a = etree.SubElement(srgb, qn("a:alpha"))
    a.set("val", str(alpha))


def set_margins(tf, l=0.06, t=0.04, r=0.06, b=0.04) -> None:
    tf.margin_left = Inches(l)
    tf.margin_right = Inches(r)
    tf.margin_top = Inches(t)
    tf.margin_bottom = Inches(b)


def _style_run(run, name, size, color, bold=False, italic=False) -> None:
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rPr, qn(tag))
        el.set("typeface", name)


def textbox(
    slide,
    l,
    t,
    w,
    h,
    parts,
    font=SANS,
    size=11,
    color=NEAR_BLACK,
    bold=False,
    italic=False,
    align=PP_ALIGN.LEFT,
    anchor=MSO_ANCHOR.TOP,
    margins=(0.04, 0.02, 0.04, 0.02),
):
    tb = slide.shapes.add_textbox(_in(l), _in(t), _in(w), _in(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.anchor = anchor
    set_margins(tf, *margins)

    if isinstance(parts, str):
        lines = parts.split("\n")
        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            p.space_before = Pt(0)
            p.space_after = Pt(0)
            run = p.add_run()
            run.text = line
            _style_run(run, font, size, color, bold, italic)
        return tb

    p = tf.paragraphs[0]
    p.alignment = align
    p.space_before = Pt(0)
    p.space_after = Pt(0)
    for part in parts:
        text = part[0]
        b = part[1] if len(part) > 1 else bold
        it = part[2] if len(part) > 2 else italic
        c = part[3] if len(part) > 3 else color
        f = part[4] if len(part) > 4 else font
        s = part[5] if len(part) > 5 else size
        run = p.add_run()
        run.text = text
        _style_run(run, f, s, c, b, it)
    return tb


# ---------------------------------------------------------------------------
# Icons (composed from native AutoShapes so they stay editable)
# ---------------------------------------------------------------------------

def icon_clipboard(slide, x, y, s=0.42, color=ORANGE):
    # Board
    add_rect(slide, x + s * 0.12, y + s * 0.16, s * 0.76, s * 0.82, color, rounded=True, adj=0.12)
    # Clip
    add_rect(slide, x + s * 0.32, y + s * 0.02, s * 0.36, s * 0.22, color, rounded=True, adj=0.3)
    add_rect(slide, x + s * 0.38, y + s * 0.06, s * 0.24, s * 0.12, WHITE, rounded=True, adj=0.3)
    # Checklist lines
    ly = y + s * 0.38
    for i in range(3):
        add_rect(slide, x + s * 0.26, ly + i * s * 0.16, s * 0.10, s * 0.10, WHITE, rounded=True, adj=0.4)
        add_rect(slide, x + s * 0.42, ly + i * s * 0.16 + s * 0.02, s * 0.34, s * 0.06, WHITE, rounded=True, adj=0.5)


def icon_target(slide, x, y, s=0.42, color=ORANGE):
    add_oval(slide, x, y, s, s, color)
    m = s * 0.22
    add_oval(slide, x + m, y + m, s - 2 * m, s - 2 * m, WHITE)
    m2 = s * 0.36
    add_oval(slide, x + m2, y + m2, s - 2 * m2, s - 2 * m2, color)


def icon_arrows(slide, x, y, s=0.46, color=ORANGE):
    """Three stacked right-pointing arrows (continue / next phase)."""
    h = s * 0.26
    w = s * 1.08
    gap = s * 0.08
    for i in range(3):
        add_auto(slide, MSO_SHAPE.RIGHT_ARROW, x, y + i * (h + gap), w, h, color)


def icon_hex_target(slide, x, y, s=0.36, color=ORANGE):
    add_auto(slide, MSO_SHAPE.HEXAGON, x, y, s, s * 0.92, color)
    inset = s * 0.22
    add_auto(
        slide,
        MSO_SHAPE.HEXAGON,
        x + inset,
        y + inset * 0.9,
        s - 2 * inset,
        s * 0.92 - 2 * inset * 0.9,
        WHITE,
    )
    c = s * 0.36
    add_oval(slide, x + c, y + c * 0.85, s - 2 * c, s - 2 * c, color)


def icon_cube(slide, x, y, s=0.36, color=ORANGE):
    add_auto(slide, MSO_SHAPE.CUBE, x, y, s, s, color)


def icon_calendar(slide, x, y, s=0.28, color=ORANGE):
    add_rect(slide, x, y + s * 0.12, s, s * 0.88, WHITE, color, 1.0, rounded=True, adj=0.12)
    add_rect(slide, x, y + s * 0.12, s, s * 0.28, color, rounded=True, adj=0.12)
    # cover bottom rounding of header
    add_rect(slide, x, y + s * 0.28, s, s * 0.12, color)
    # rings
    add_rect(slide, x + s * 0.22, y, s * 0.10, s * 0.22, color, rounded=True, adj=0.4)
    add_rect(slide, x + s * 0.68, y, s * 0.10, s * 0.22, color, rounded=True, adj=0.4)
    # day cells
    add_rect(slide, x + s * 0.18, y + s * 0.52, s * 0.18, s * 0.16, ORANGE_SOFT, rounded=True, adj=0.3)
    add_rect(slide, x + s * 0.42, y + s * 0.52, s * 0.18, s * 0.16, color, rounded=True, adj=0.3)
    add_rect(slide, x + s * 0.66, y + s * 0.52, s * 0.18, s * 0.16, ORANGE_SOFT, rounded=True, adj=0.3)
    add_rect(slide, x + s * 0.18, y + s * 0.74, s * 0.18, s * 0.16, ORANGE_SOFT, rounded=True, adj=0.3)
    add_rect(slide, x + s * 0.42, y + s * 0.74, s * 0.18, s * 0.16, ORANGE_SOFT, rounded=True, adj=0.3)


def icon_star(slide, x, y, s=0.32, fill=MID_GRAY, ring=True):
    if ring:
        add_oval(slide, x, y, s, s, WHITE, fill, 1.5)
        pad = s * 0.18
        add_auto(slide, MSO_SHAPE.STAR_5_POINT, x + pad, y + pad, s - 2 * pad, s - 2 * pad, fill)
    else:
        add_auto(slide, MSO_SHAPE.STAR_5_POINT, x, y, s, s, fill)


def icon_check(slide, x, y, s=0.30):
    add_oval(slide, x, y, s, s, ORANGE)
    textbox(
        slide,
        x,
        y + 0.01,
        s,
        s,
        "✓",
        font=SANS,
        size=12,
        color=WHITE,
        bold=True,
        align=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.MIDDLE,
        margins=(0, 0, 0, 0),
    )


def numbered_badge(slide, x, y, n: str, d=0.30):
    add_oval(slide, x - 0.018, y - 0.018, d + 0.036, d + 0.036, WHITE)
    add_oval(slide, x, y, d, d, ORANGE)
    textbox(
        slide,
        x,
        y + 0.005,
        d,
        d - 0.01,
        n,
        font=SANS,
        size=12,
        color=WHITE,
        bold=True,
        align=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.MIDDLE,
        margins=(0, 0, 0, 0),
    )


# ---------------------------------------------------------------------------
# Slide construction
# ---------------------------------------------------------------------------

def build() -> Presentation:
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank

    # Background
    bg = add_rect(slide, 0, 0, 13.333333, 7.5, WHITE)
    sp_tree = slide.shapes._spTree
    sp = bg._element
    sp_tree.remove(sp)
    sp_tree.insert(2, sp)

    ML = 0.32  # left margin
    MR = 13.333333 - 0.32

    # ===== Header =====
    textbox(
        slide, ML, 0.12, 12.7, 0.42,
        "What We Are Proposing",
        font=SERIF, size=28, color=BLACK, bold=True,
        margins=(0, 0, 0, 0),
    )
    textbox(
        slide, ML, 0.52, 12.7, 0.32,
        "Start MVP definition and build, enabling early value and setting the foundation for continued capability and roadmap development.",
        font=SANS, size=12, color=DARK_GRAY,
        margins=(0, 0, 0, 0),
    )

    # ===== Three step cards =====
    card_y = 0.88
    card_h = 2.20
    gap = 0.14
    usable = MR - ML
    card_w = (usable - 2 * gap) / 3
    xs = [ML + i * (card_w + gap) for i in range(3)]

    cards = [
        {
            "n": "1",
            "title": "Temporarily Pause Business Requirements (BRs)",
            "icon": "clipboard",
            "body": [
                ("Complete labeling requirements while ", False),
                ("pausing", True),
                (" detailed requirements for other ", False),
                ("planning capabilities.", True),
            ],
            "callout": None,
            "timing": None,
        },
        {
            "n": "2",
            "title": "Define MVP Scope and Decision",
            "icon": "target",
            "body": [
                ("Prioritize high-value capabilities, develop functional and UX requirements, and define MVP scope.", False),
            ],
            "callout": {
                "kicker": "DECISION POINT  (~4 WEEKS*)",
                "text": "Confirm MVP scope and timeline. Proceed with planned scope or adjust features to meet target timeline.",
            },
            "timing": None,
        },
        {
            "n": "3",
            "title": "Complete Remaining BRs and Deliver MVP Build",
            "icon": "arrows",
            "body": [
                ("Complete requirements for remaining planning capabilities and prioritize future releases.", False),
            ],
            "callout": None,
            "timing": {
                "kicker": "~6 WEEKS*",
                "text": "Starts as soon as MVP definition is complete and runs in parallel with MVP build.",
            },
        },
    ]

    for i, (cx, spec) in enumerate(zip(xs, cards)):
        card = add_rect(slide, cx, card_y, card_w, card_h, WHITE, CARD_BORDER, 1.0, rounded=True, adj=0.04)
        apply_shadow(card)

        # Icon top-right
        ix = cx + card_w - 0.58
        iy = card_y + 0.16
        if spec["icon"] == "clipboard":
            icon_clipboard(slide, ix, iy, 0.40)
        elif spec["icon"] == "target":
            icon_target(slide, ix, iy, 0.40)
        else:
            icon_arrows(slide, ix - 0.04, iy + 0.04, 0.42)

        # Title
        textbox(
            slide, cx + 0.16, card_y + 0.16, card_w - 0.78, 0.52,
            spec["title"],
            font=SANS, size=13, color=BLACK, bold=True,
            anchor=MSO_ANCHOR.TOP,
            margins=(0, 0, 0, 0),
        )

        # Body
        body_h = 0.62 if spec["callout"] or spec["timing"] else 0.90
        tb = slide.shapes.add_textbox(_in(cx + 0.12), _in(card_y + 0.70), _in(card_w - 0.24), _in(body_h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.anchor = MSO_ANCHOR.TOP
        set_margins(tf, 0.04, 0.02, 0.04, 0.02)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        p.space_before = Pt(0)
        p.space_after = Pt(0)
        for part in spec["body"]:
            run = p.add_run()
            run.text = part[0]
            _style_run(run, SANS, 11, DARK_GRAY, bold=part[1])

        if spec["callout"]:
            cy = card_y + card_h - 0.78
            ch = 0.66
            add_rect(slide, cx + 0.12, cy, card_w - 0.24, ch, ORANGE_WASH, rounded=True, adj=0.08)
            icon_calendar(slide, cx + 0.20, cy + 0.16, 0.30)
            textbox(
                slide, cx + 0.56, cy + 0.04, card_w - 0.76, 0.22,
                spec["callout"]["kicker"],
                font=SANS, size=9, color=ORANGE, bold=True,
                margins=(0.02, 0.02, 0.04, 0),
            )
            textbox(
                slide, cx + 0.56, cy + 0.24, card_w - 0.76, 0.40,
                spec["callout"]["text"],
                font=SANS, size=9, color=DARK_GRAY,
                margins=(0.02, 0, 0.04, 0.04),
            )

        if spec["timing"]:
            ty = card_y + card_h - 0.62
            textbox(
                slide, cx + 0.16, ty, card_w - 0.32, 0.22,
                spec["timing"]["kicker"],
                font=SANS, size=12, color=ORANGE, bold=True,
                margins=(0, 0, 0, 0),
            )
            textbox(
                slide, cx + 0.16, ty + 0.20, card_w - 0.32, 0.38,
                spec["timing"]["text"],
                font=SANS, size=9, color=MID_GRAY,
                margins=(0, 0, 0, 0),
            )

        numbered_badge(slide, cx - 0.06, card_y - 0.08, spec["n"], 0.30)

    # ===== High-level roadmap =====
    rm_top = 3.18
    textbox(
        slide, ML, rm_top, usable, 0.22,
        "HIGH-LEVEL ROADMAP",
        font=SANS, size=10, color=LABEL_GRAY, bold=True,
        align=PP_ALIGN.CENTER, margins=(0, 0, 0, 0),
    )

    # Shared timeline coordinates so the three workstreams align
    label_x = ML
    label_w = 2.12
    tl0 = ML + 2.28                 # path start (dot)
    box_x = tl0 + 0.48              # definition-phase boxes
    box_w = 2.85
    box_r = box_x + box_w
    star_x = box_r + 0.18           # decision star
    dash_end = star_x + 0.55        # end of dashed gap after pause
    brs_arrow_end = dash_end + 2.55 # remaining-BRs line (shorter than MVP build)
    bar_x = star_x + 0.08           # MVP build starts at the decision
    build_end = 11.42               # MVP build is longer than remaining BRs
    delivered_x = build_end + 0.10

    row1_y = 3.40
    row2_y = 4.22
    row3_y = 5.04
    line_w = 2.5

    # -- Row 1: Business requirements (dot → pause box → dashed → solid arrow) --
    icon_clipboard(slide, label_x, row1_y + 0.16, 0.32)
    textbox(
        slide, label_x + 0.38, row1_y + 0.08, label_w - 0.42, 0.54,
        "BUSINESS REQUIREMENTS &\nSCALE-UP PLAN",
        font=SANS, size=8, color=LABEL_GRAY, bold=True,
        margins=(0, 0, 0, 0),
    )

    mid1 = row1_y + 0.38
    add_oval(slide, tl0, mid1 - 0.085, 0.17, 0.17, ORANGE)
    add_line(slide, tl0 + 0.17, mid1, box_x, mid1, ORANGE, line_w)
    add_rect(slide, box_x, row1_y + 0.10, box_w, 0.56, ORANGE_SOFT, ORANGE, 1.15, rounded=True, adj=0.10)
    textbox(
        slide, box_x + 0.06, row1_y + 0.12, box_w - 0.12, 0.52,
        "Temporarily pause detailed BRs for other planning capabilities (focus on labeling)",
        font=SANS, size=8, color=NEAR_BLACK,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        margins=(0.04, 0.02, 0.04, 0.02),
    )
    add_line(
        slide, box_r, mid1, dash_end, mid1, ORANGE, line_w, dash=MSO_LINE_DASH_STYLE.DASH
    )
    add_line(slide, dash_end, mid1, brs_arrow_end, mid1, ORANGE, line_w, arrow=True)
    textbox(
        slide, brs_arrow_end + 0.08, row1_y + 0.10, MR - (brs_arrow_end + 0.08), 0.56,
        "Complete remaining BRs and prioritize future releases",
        font=SANS, size=8, color=DARK_GRAY,
        anchor=MSO_ANCHOR.MIDDLE,
        margins=(0.02, 0.02, 0.02, 0.02),
    )

    # -- Row 2: MVP definition (line → box → star / decision) --
    icon_hex_target(slide, label_x, row2_y + 0.16, 0.32, color=MID_GRAY)
    textbox(
        slide, label_x + 0.38, row2_y + 0.12, label_w - 0.42, 0.50,
        "MVP DEFINITION\n& DECISION",
        font=SANS, size=8, color=LABEL_GRAY, bold=True,
        margins=(0, 0, 0, 0),
    )

    mid2 = row2_y + 0.36
    add_oval(slide, tl0 + 0.02, mid2 - 0.06, 0.12, 0.12, LINE_GRAY)
    add_line(slide, tl0 + 0.14, mid2, box_x, mid2, LINE_GRAY, 1.85)
    add_rect(slide, box_x, row2_y + 0.10, box_w, 0.52, LIGHT_GRAY, LINE_GRAY, 1.15, rounded=True, adj=0.10)
    textbox(
        slide, box_x + 0.06, row2_y + 0.12, box_w - 0.12, 0.48,
        "Define MVP scope and decision",
        font=SANS, size=10, color=NEAR_BLACK, bold=True,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        margins=(0.04, 0.02, 0.04, 0.02),
    )
    add_line(slide, box_r, mid2, star_x + 0.04, mid2, LINE_GRAY, 1.85)
    icon_star(slide, star_x, row2_y + 0.16, 0.38, fill=MID_GRAY, ring=True)
    textbox(
        slide, star_x - 0.22, row2_y + 0.54, 0.82, 0.20,
        "(~4 WEEKS*)",
        font=SANS, size=8, color=ORANGE, bold=True,
        align=PP_ALIGN.CENTER, margins=(0, 0, 0, 0),
    )

    # Vertical dashed drop from the decision star to the start of MVP build
    star_cx = star_x + 0.19
    add_line(
        slide,
        star_cx, row2_y + 0.54,
        star_cx, row3_y + 0.18,
        LINE_GRAY, 1.35, dash=MSO_LINE_DASH_STYLE.DASH,
    )

    # -- Row 3: MVP build (starts at decision; longer than remaining BRs) --
    icon_cube(slide, label_x, row3_y + 0.14, 0.34, color=ORANGE)
    textbox(
        slide, label_x + 0.38, row3_y + 0.16, label_w - 0.42, 0.36,
        "MVP BUILD",
        font=SANS, size=8, color=LABEL_GRAY, bold=True,
        margins=(0, 0, 0, 0),
    )

    bar_w = build_end - bar_x
    add_rect(slide, bar_x, row3_y + 0.26, bar_w, 0.36, ORANGE_BAR, rounded=True, adj=0.35)
    textbox(
        slide, bar_x, row3_y + 0.26, bar_w, 0.36,
        "MVP Build  (duration longer than BRs)",
        font=SANS, size=10, color=NEAR_BLACK, bold=True,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        margins=(0.08, 0, 0.08, 0),
    )
    textbox(
        slide, bar_x + 0.55, row3_y - 0.02, 5.4, 0.24,
        "~6 WEEKS*   Starts as soon as MVP definition is complete and runs in parallel with MVP build.",
        font=SANS, size=8, color=ORANGE,
        margins=(0.02, 0, 0, 0),
    )

    icon_check(slide, delivered_x, row3_y + 0.24, 0.36)
    textbox(
        slide, delivered_x + 0.40, row3_y + 0.16, MR - (delivered_x + 0.40), 0.48,
        "MVP BUILD\nDELIVERED",
        font=SANS, size=8, color=ORANGE, bold=True,
        anchor=MSO_ANCHOR.MIDDLE,
        margins=(0.02, 0, 0, 0),
    )

    # ===== Outcome banner =====
    out_y = 6.02
    out_h = 0.78
    banner = add_rect(slide, ML, out_y, usable, out_h, LIGHT_GRAY, rounded=True, adj=0.08)
    icon_hex_target(slide, ML + 0.18, out_y + 0.20, 0.40, color=ORANGE)
    tb = slide.shapes.add_textbox(_in(ML + 0.70), _in(out_y + 0.10), _in(usable - 0.90), _in(0.58))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.anchor = MSO_ANCHOR.MIDDLE
    set_margins(tf, 0.06, 0.04, 0.08, 0.04)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    p.space_before = Pt(0)
    p.space_after = Pt(0)
    r1 = p.add_run()
    r1.text = "Outcome: "
    _style_run(r1, SANS, 13, NEAR_BLACK, bold=True)
    r2 = p.add_run()
    r2.text = "Deliver early value with a focused MVP while building the foundation for continued capability and roadmap development."
    _style_run(r2, SANS, 13, DARK_GRAY, bold=False)

    # ===== Footer =====
    fy = 6.95
    # PwC wordmark
    textbox(
        slide, ML, fy, 1.1, 0.32,
        "PwC",
        font=SERIF, size=16, color=BLACK, bold=True,
        anchor=MSO_ANCHOR.MIDDLE, margins=(0, 0, 0, 0),
    )
    textbox(
        slide, 3.4, fy, 6.5, 0.32,
        "* Durations are estimates and will be refined.",
        font=SANS, size=9, color=MID_GRAY, italic=True,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margins=(0, 0, 0, 0),
    )
    textbox(
        slide, MR - 0.50, fy, 0.50, 0.32,
        "4",
        font=SANS, size=11, color=DARK_GRAY,
        align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE, margins=(0, 0, 0, 0),
    )

    prs.core_properties.title = "What We Are Proposing"
    prs.core_properties.subject = "MVP definition and build"
    return prs


def main() -> None:
    prs = build()
    out = "What_We_Are_Proposing.pptx"
    prs.save(out)
    print(f"Wrote {out}")


if __name__ == "__main__":
    main()
