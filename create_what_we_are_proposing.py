#!/usr/bin/env python3
"""Build an editable PowerPoint slide for 'What We Are Proposing'.

All copy is native PowerPoint text. Icons are small transparent PNGs so they
stay sharp; the slide itself is not a screenshot.
"""

from __future__ import annotations

import math
from pathlib import Path

from lxml import etree
from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

# ---------------------------------------------------------------------------
# Geometry and palette
# ---------------------------------------------------------------------------

SLIDE_W_IN = 13.333333
SLIDE_H_IN = 7.5

ORANGE = RGBColor(0xD0, 0x4A, 0x02)
ORANGE_SOFT = RGBColor(0xFB, 0xE6, 0xD8)
ORANGE_WASH = RGBColor(0xFD, 0xF1, 0xE9)
ORANGE_BAR = RGBColor(0xF2, 0xC2, 0xA3)
BLACK = RGBColor(0x00, 0x00, 0x00)
INK = RGBColor(0x2B, 0x2B, 0x2B)
BODY = RGBColor(0x4A, 0x4A, 0x4A)
MUTED = RGBColor(0x6F, 0x6F, 0x6F)
LABEL = RGBColor(0x5A, 0x5A, 0x5A)
HAIRLINE = RGBColor(0xE2, 0xE2, 0xE2)
LINE_GRAY = RGBColor(0xC8, 0xC8, 0xC8)
PANEL = RGBColor(0xF6, 0xF6, 0xF6)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

SERIF = "Georgia"
SANS = "Calibri"

ORANGE_RGBA = (0xD0, 0x4A, 0x02, 255)
GRAY_RGBA = (0x6B, 0x6B, 0x6B, 255)
WHITE_RGBA = (255, 255, 255, 255)
ICON_DIR = Path("/tmp/pptx-icons")


def _in(x: float) -> int:
    return int(Inches(x))


# ---------------------------------------------------------------------------
# Icons (drawn at 512px, then downscaled)
# ---------------------------------------------------------------------------

def _canvas(size: int = 512) -> tuple[Image.Image, ImageDraw.ImageDraw]:
    im = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    return im, ImageDraw.Draw(im)


def _save(im: Image.Image, name: str, out: int = 160) -> str:
    ICON_DIR.mkdir(parents=True, exist_ok=True)
    path = ICON_DIR / name
    im.resize((out, out), Image.Resampling.LANCZOS).save(path)
    return str(path)


def _line(d: ImageDraw.ImageDraw, a, b, fill, w: int) -> None:
    d.line([a, b], fill=fill, width=w)
    r = max(w // 2, 1)
    for p in (a, b):
        d.ellipse([p[0] - r, p[1] - r, p[0] + r, p[1] + r], fill=fill)


def _hex(cx, cy, r, pointy=True):
    start = 30 if pointy else 0
    pts = []
    for i in range(6):
        a = math.radians(start + i * 60)
        pts.append((cx + r * math.cos(a), cy + r * math.sin(a)))
    return pts


def _star(cx, cy, r_out, r_in, n=5):
    pts = []
    for i in range(n * 2):
        r = r_out if i % 2 == 0 else r_in
        a = math.radians(-90 + i * 180 / n)
        pts.append((cx + r * math.cos(a), cy + r * math.sin(a)))
    return pts


def _hex_ring(im: Image.Image, cx: float, cy: float, r: float, width: float, color) -> None:
    """Filled hexagonal ring with clean corners (no notched joints)."""
    mask = Image.new("L", im.size, 0)
    md = ImageDraw.Draw(mask)
    md.polygon(_hex(cx, cy, r + width / 2), fill=255)
    md.polygon(_hex(cx, cy, max(r - width / 2, 1)), fill=0)
    layer = Image.new("RGBA", im.size, (0, 0, 0, 0))
    layer.paste(Image.new("RGBA", im.size, color), mask=mask)
    im.alpha_composite(layer)


def make_icons() -> dict[str, str]:
    icons: dict[str, str] = {}
    S, C, sw = 512, 256, 28

    # Clipboard
    im, d = _canvas()
    d.rounded_rectangle([118, 118, 394, 430], radius=36, outline=ORANGE_RGBA, width=sw)
    d.rounded_rectangle([186, 70, 326, 150], radius=22, fill=ORANGE_RGBA)
    d.rounded_rectangle([210, 88, 302, 128], radius=12, fill=WHITE_RGBA)
    for y in (210, 278, 346):
        d.rounded_rectangle([168, y, 204, y + 28], radius=6, outline=ORANGE_RGBA, width=10)
        _line(d, (228, y + 14), (352, y + 14), ORANGE_RGBA, 12)
    icons["clipboard"] = _save(im, "clipboard.png")

    # Target / bullseye
    im, d = _canvas()
    d.ellipse([86, 86, 426, 426], outline=ORANGE_RGBA, width=sw)
    d.ellipse([156, 156, 356, 356], outline=ORANGE_RGBA, width=sw)
    d.ellipse([226, 226, 286, 286], fill=ORANGE_RGBA)
    icons["target"] = _save(im, "target.png")

    # Three right-pointing chevrons
    im, d = _canvas()

    def chevron(y0: int) -> None:
        h, w = 96, 400
        x0 = 56
        pts = [
            (x0, y0),
            (x0 + int(w * 0.58), y0),
            (x0 + w, y0 + h // 2),
            (x0 + int(w * 0.58), y0 + h),
            (x0, y0 + h),
            (x0 + int(w * 0.42), y0 + h // 2),
        ]
        d.polygon(pts, fill=ORANGE_RGBA)

    chevron(86)
    chevron(208)
    chevron(330)
    icons["arrows"] = _save(im, "arrows.png")

    # Calendar
    im, d = _canvas()
    d.rounded_rectangle([96, 118, 416, 430], radius=36, outline=ORANGE_RGBA, width=sw)
    d.rectangle([96, 118, 416, 200], fill=ORANGE_RGBA)
    d.pieslice([96, 118, 168, 190], 180, 270, fill=ORANGE_RGBA)
    d.pieslice([344, 118, 416, 190], 270, 360, fill=ORANGE_RGBA)
    for x in (176, 312):
        d.rounded_rectangle([x, 78, x + 28, 150], radius=12, fill=ORANGE_RGBA)
    for i, x in enumerate((156, 242, 328)):
        for j, y in enumerate((250, 334)):
            fill = ORANGE_RGBA if (i, j) == (1, 0) else (*ORANGE_RGBA[:3], 70)
            d.rounded_rectangle([x, y, x + 44, y + 44], radius=8, fill=fill)
    icons["calendar"] = _save(im, "calendar.png")

    # Hex target (orange)
    im, d = _canvas()
    _hex_ring(im, C, C, 176, 32, ORANGE_RGBA)
    _hex_ring(im, C, C, 108, 26, ORANGE_RGBA)
    d = ImageDraw.Draw(im)
    d.ellipse([C - 34, C - 34, C + 34, C + 34], fill=ORANGE_RGBA)
    icons["hex_orange"] = _save(im, "hex_orange.png")

    # Hex target (gray)
    im, d = _canvas()
    _hex_ring(im, C, C, 176, 32, GRAY_RGBA)
    _hex_ring(im, C, C, 108, 26, GRAY_RGBA)
    d = ImageDraw.Draw(im)
    d.ellipse([C - 34, C - 34, C + 34, C + 34], fill=GRAY_RGBA)
    icons["hex_gray"] = _save(im, "hex_gray.png")

    # Isometric cube
    im, d = _canvas()
    top = [(256, 90), (400, 168), (256, 246), (112, 168)]
    left = [(112, 168), (256, 246), (256, 422), (112, 344)]
    right = [(256, 246), (400, 168), (400, 344), (256, 422)]
    d.polygon(top, fill=(0xD0, 0x4A, 0x02, 255))
    d.polygon(left, fill=(0xE0, 0x78, 0x38, 255))
    d.polygon(right, fill=(0xB8, 0x3E, 0x00, 255))
    icons["cube"] = _save(im, "cube.png")

    # Star in circle
    im, d = _canvas()
    d.ellipse([64, 64, 448, 448], outline=GRAY_RGBA, width=26)
    d.polygon(_star(C, C, 140, 58), fill=GRAY_RGBA)
    icons["star"] = _save(im, "star.png")

    # Check in circle
    im, d = _canvas()
    d.ellipse([48, 48, 464, 464], fill=ORANGE_RGBA)
    _line(d, (150, 268), (228, 348), WHITE_RGBA, 36)
    _line(d, (228, 348), (370, 186), WHITE_RGBA, 36)
    icons["check"] = _save(im, "check.png")

    return icons


# ---------------------------------------------------------------------------
# PowerPoint helpers
# ---------------------------------------------------------------------------

def _fill(shape, color: RGBColor | None) -> None:
    if color is None:
        shape.fill.background()
    else:
        shape.fill.solid()
        shape.fill.fore_color.rgb = color


def _line_fmt(shape, color: RGBColor | None, width_pt: float = 1.0) -> None:
    if color is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = color
        shape.line.width = Pt(width_pt)


def add_rect(slide, l, t, w, h, fill, line=None, line_w=0.75, rounded=False, adj=0.06):
    kind = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    sh = slide.shapes.add_shape(kind, _in(l), _in(t), _in(w), _in(h))
    _fill(sh, fill)
    _line_fmt(sh, line, line_w)
    if rounded:
        try:
            sh.adjustments[0] = adj
        except Exception:
            pass
    sh.shadow.inherit = False
    return sh


def add_oval(slide, l, t, w, h, fill, line=None, line_w=0.75):
    sh = slide.shapes.add_shape(MSO_SHAPE.OVAL, _in(l), _in(t), _in(w), _in(h))
    _fill(sh, fill)
    _line_fmt(sh, line, line_w)
    sh.shadow.inherit = False
    return sh


def add_arrowhead(slide, x, y, color=ORANGE, w=0.14, h=0.11):
    sh = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, _in(x), _in(y - h / 2), _in(w), _in(h))
    _fill(sh, color)
    _line_fmt(sh, None)
    try:
        sh.adjustments[0] = 0.5
    except Exception:
        pass
    sh.shadow.inherit = False
    return sh


def apply_shadow(shape, blur=0.06, dist=0.02, alpha=12000) -> None:
    sp_pr = shape._element.spPr
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


def set_margins(tf, l=0.04, t=0.02, r=0.04, b=0.02) -> None:
    tf.margin_left = Inches(l)
    tf.margin_right = Inches(r)
    tf.margin_top = Inches(t)
    tf.margin_bottom = Inches(b)


def textbox(
    slide,
    l,
    t,
    w,
    h,
    parts,
    font=SANS,
    size=11,
    color=INK,
    bold=False,
    italic=False,
    align=PP_ALIGN.LEFT,
    anchor=MSO_ANCHOR.TOP,
    margins=(0.02, 0.01, 0.02, 0.01),
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


def add_pic(slide, path: str, x: float, y: float, s: float):
    pic = slide.shapes.add_picture(path, _in(x), _in(y), _in(s), _in(s))
    pic.line.fill.background()
    return pic


def hbar(slide, x1, x2, y, color, th=0.03):
    if x2 <= x1:
        return None
    return add_rect(slide, x1, y - th / 2, x2 - x1, th, color, rounded=True, adj=0.5)


def hdash(slide, x1, x2, y, color, th=0.03, dash=0.07, gap=0.05):
    x = x1
    while x < x2 - 0.01:
        w = min(dash, x2 - x)
        if w >= 0.03:
            add_rect(slide, x, y - th / 2, w, th, color, rounded=True, adj=0.5)
        x += dash + gap


def vdash(slide, x, y1, y2, color, th=0.018, dash=0.06, gap=0.045):
    y = y1
    while y < y2 - 0.01:
        h = min(dash, y2 - y)
        if h >= 0.03:
            add_rect(slide, x - th / 2, y, th, h, color, rounded=True, adj=0.5)
        y += dash + gap


def badge(slide, x, y, n: str, d=0.28):
    add_oval(slide, x, y, d, d, ORANGE)
    textbox(
        slide, x, y, d, d, n,
        font=SANS, size=11, color=WHITE, bold=True,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        margins=(0, 0, 0, 0),
    )


# ---------------------------------------------------------------------------
# Slide
# ---------------------------------------------------------------------------

def build() -> Presentation:
    icons = make_icons()
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W_IN)
    prs.slide_height = Inches(SLIDE_H_IN)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    add_rect(slide, 0, 0, SLIDE_W_IN, SLIDE_H_IN, WHITE)

    ML, MR = 0.38, 12.95
    usable = MR - ML

    # Header
    textbox(
        slide, ML, 0.16, usable, 0.40,
        "What We Are Proposing",
        font=SERIF, size=28, color=BLACK, bold=True, margins=(0, 0, 0, 0),
    )
    textbox(
        slide, ML, 0.56, usable, 0.30,
        "Start MVP definition and build, enabling early value and setting the foundation for continued capability and roadmap development.",
        font=SANS, size=13, color=BODY, margins=(0, 0, 0, 0),
    )

    # ----- Three cards -----
    card_y, card_h, gap = 0.92, 2.22, 0.16
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
            "kind": "plain",
        },
        {
            "n": "2",
            "title": "Define MVP Scope and Decision",
            "icon": "target",
            "body": [
                ("Prioritize high-value capabilities, develop functional and UX requirements, and define MVP scope.", False),
            ],
            "kind": "decision",
        },
        {
            "n": "3",
            "title": "Complete Remaining BRs and Deliver MVP Build",
            "icon": "arrows",
            "body": [
                ("Complete requirements for remaining planning capabilities and prioritize future releases.", False),
            ],
            "kind": "timing",
        },
    ]

    for cx, spec in zip(xs, cards):
        card = add_rect(slide, cx, card_y, card_w, card_h, WHITE, HAIRLINE, 0.75, rounded=True, adj=0.045)
        apply_shadow(card)

        badge(slide, cx + 0.14, card_y + 0.14, spec["n"], 0.28)
        add_pic(slide, icons[spec["icon"]], cx + card_w - 0.50, card_y + 0.14, 0.34)

        textbox(
            slide, cx + 0.50, card_y + 0.12, card_w - 1.08, 0.48,
            spec["title"], font=SANS, size=13, color=BLACK, bold=True,
            anchor=MSO_ANCHOR.MIDDLE, margins=(0.02, 0, 0.02, 0),
        )

        footer_h = 0.70 if spec["kind"] == "decision" else (0.56 if spec["kind"] == "timing" else 0.12)
        body_top = card_y + 0.66
        body_h = (card_y + card_h - footer_h - 0.08) - body_top
        tb = slide.shapes.add_textbox(
            _in(cx + 0.14), _in(body_top),
            _in(card_w - 0.28), _in(max(body_h, 0.40)),
        )
        tf = tb.text_frame
        tf.word_wrap = True
        tf.anchor = MSO_ANCHOR.TOP
        set_margins(tf, 0.02, 0.0, 0.02, 0.0)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        p.space_before = Pt(0)
        p.space_after = Pt(0)
        p.line_spacing = 1.10
        for part in spec["body"]:
            run = p.add_run()
            run.text = part[0]
            _style_run(run, SANS, 12, BODY, bold=part[1])

        if spec["kind"] == "decision":
            cy, ch = card_y + card_h - 0.70, 0.60
            add_rect(slide, cx + 0.12, cy, card_w - 0.24, ch, ORANGE_WASH, rounded=True, adj=0.10)
            add_pic(slide, icons["calendar"], cx + 0.20, cy + 0.16, 0.30)
            textbox(
                slide, cx + 0.54, cy + 0.05, card_w - 0.72, 0.20,
                "DECISION POINT  (~4 WEEKS*)",
                font=SANS, size=9, color=ORANGE, bold=True, margins=(0.02, 0, 0.04, 0),
            )
            textbox(
                slide, cx + 0.54, cy + 0.24, card_w - 0.72, 0.36,
                "Confirm MVP scope and timeline. Proceed with planned scope or adjust features to meet target timeline.",
                font=SANS, size=9, color=BODY, margins=(0.02, 0, 0.04, 0.04),
            )

        if spec["kind"] == "timing":
            ty = card_y + card_h - 0.58
            add_rect(slide, cx + 0.14, ty - 0.08, 0.28, 0.035, ORANGE)
            textbox(
                slide, cx + 0.14, ty, card_w - 0.28, 0.20,
                "~6 WEEKS*", font=SANS, size=12, color=ORANGE, bold=True, margins=(0, 0, 0, 0),
            )
            textbox(
                slide, cx + 0.14, ty + 0.20, card_w - 0.28, 0.30,
                "Starts as soon as MVP definition is complete and runs in parallel with MVP build.",
                font=SANS, size=10, color=MUTED, margins=(0, 0, 0, 0),
            )

    # ----- Roadmap panel -----
    pan_y, pan_h = 3.26, 2.62
    add_rect(slide, ML, pan_y, usable, pan_h, PANEL, rounded=True, adj=0.04)
    textbox(
        slide, ML, pan_y + 0.06, usable, 0.22,
        "HIGH-LEVEL ROADMAP",
        font=SANS, size=10, color=LABEL, bold=True,
        align=PP_ALIGN.CENTER, margins=(0, 0, 0, 0),
    )

    label_x = ML + 0.16
    tl0 = ML + 2.42
    box_x = tl0 + 0.40
    box_w = 3.05
    box_r = box_x + box_w
    star_x = box_r + 0.16
    dash_end = star_x + 0.50
    brs_end = dash_end + 2.35
    bar_x = star_x + 0.02
    build_end = 11.15
    delivered_x = build_end + 0.12

    row1_y, row2_y, row3_y = pan_y + 0.32, pan_y + 1.08, pan_y + 1.86

    # Row 1
    add_pic(slide, icons["clipboard"], label_x, row1_y + 0.14, 0.30)
    textbox(
        slide, label_x + 0.36, row1_y + 0.08, 1.70, 0.50,
        "BUSINESS REQUIREMENTS &\nSCALE-UP PLAN",
        font=SANS, size=8, color=LABEL, bold=True, margins=(0, 0, 0, 0),
    )
    mid1 = row1_y + 0.36
    add_oval(slide, tl0, mid1 - 0.07, 0.14, 0.14, ORANGE)
    hbar(slide, tl0 + 0.14, box_x, mid1, ORANGE, 0.032)
    add_rect(slide, box_x, row1_y + 0.08, box_w, 0.56, ORANGE_SOFT, ORANGE, 1.0, rounded=True, adj=0.12)
    textbox(
        slide, box_x + 0.08, row1_y + 0.10, box_w - 0.16, 0.52,
        "Temporarily pause detailed BRs for other planning capabilities (focus on labeling)",
        font=SANS, size=9, color=INK, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        margins=(0.04, 0.02, 0.04, 0.02),
    )
    hdash(slide, box_r + 0.04, dash_end, mid1, ORANGE, 0.032)
    hbar(slide, dash_end, brs_end, mid1, ORANGE, 0.032)
    add_arrowhead(slide, brs_end - 0.02, mid1, ORANGE, 0.15, 0.12)
    textbox(
        slide, brs_end + 0.16, row1_y + 0.08, MR - 0.18 - (brs_end + 0.16), 0.56,
        "Complete remaining BRs and prioritize future releases",
        font=SANS, size=9, color=BODY, anchor=MSO_ANCHOR.MIDDLE, margins=(0.02, 0.02, 0.02, 0.02),
    )

    # Row 2
    add_pic(slide, icons["hex_gray"], label_x, row2_y + 0.14, 0.30)
    textbox(
        slide, label_x + 0.36, row2_y + 0.10, 1.70, 0.46,
        "MVP DEFINITION\n& DECISION",
        font=SANS, size=8, color=LABEL, bold=True, margins=(0, 0, 0, 0),
    )
    mid2 = row2_y + 0.34
    add_oval(slide, tl0 + 0.01, mid2 - 0.06, 0.12, 0.12, LINE_GRAY)
    hbar(slide, tl0 + 0.13, box_x, mid2, LINE_GRAY, 0.026)
    add_rect(slide, box_x, row2_y + 0.08, box_w, 0.52, WHITE, LINE_GRAY, 1.0, rounded=True, adj=0.12)
    textbox(
        slide, box_x + 0.08, row2_y + 0.10, box_w - 0.16, 0.48,
        "Define MVP scope and decision",
        font=SANS, size=11, color=INK, bold=True,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        margins=(0.04, 0.02, 0.04, 0.02),
    )
    hbar(slide, box_r + 0.04, star_x + 0.04, mid2, LINE_GRAY, 0.026)
    add_pic(slide, icons["star"], star_x, row2_y + 0.12, 0.38)
    textbox(
        slide, star_x - 0.18, row2_y + 0.50, 0.76, 0.18,
        "(~4 WEEKS*)", font=SANS, size=8, color=ORANGE, bold=True,
        align=PP_ALIGN.CENTER, margins=(0, 0, 0, 0),
    )
    star_cx = star_x + 0.19
    vdash(slide, star_cx, row2_y + 0.50, row3_y + 0.22, LINE_GRAY)

    # Row 3
    add_pic(slide, icons["cube"], label_x, row3_y + 0.14, 0.30)
    textbox(
        slide, label_x + 0.36, row3_y + 0.16, 1.70, 0.28,
        "MVP BUILD", font=SANS, size=8, color=LABEL, bold=True, margins=(0, 0, 0, 0),
    )
    textbox(
        slide, bar_x + 0.46, row3_y + 0.00, 5.6, 0.20,
        "~6 WEEKS*   Starts as soon as MVP definition is complete and runs in parallel with MVP build.",
        font=SANS, size=8, color=ORANGE, margins=(0, 0, 0, 0),
    )
    add_rect(slide, bar_x, row3_y + 0.24, build_end - bar_x, 0.34, ORANGE_BAR, rounded=True, adj=0.45)
    textbox(
        slide, bar_x, row3_y + 0.24, build_end - bar_x, 0.34,
        "MVP Build   (duration longer than BRs)",
        font=SANS, size=11, color=INK, bold=True,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margins=(0.08, 0, 0.08, 0),
    )
    add_pic(slide, icons["check"], delivered_x, row3_y + 0.22, 0.36)
    textbox(
        slide, delivered_x + 0.40, row3_y + 0.16, MR - 0.16 - (delivered_x + 0.40), 0.46,
        "MVP BUILD\nDELIVERED",
        font=SANS, size=9, color=ORANGE, bold=True, anchor=MSO_ANCHOR.MIDDLE, margins=(0.02, 0, 0, 0),
    )

    # ----- Outcome -----
    out_y, out_h = 6.00, 0.72
    add_rect(slide, ML, out_y, usable, out_h, PANEL, rounded=True, adj=0.10)
    add_pic(slide, icons["hex_orange"], ML + 0.18, out_y + 0.16, 0.40)
    tb = slide.shapes.add_textbox(_in(ML + 0.68), _in(out_y + 0.08), _in(usable - 0.86), _in(0.56))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.anchor = MSO_ANCHOR.MIDDLE
    set_margins(tf, 0.04, 0.02, 0.08, 0.02)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    p.space_before = Pt(0)
    p.space_after = Pt(0)
    r1 = p.add_run()
    r1.text = "Outcome: "
    _style_run(r1, SANS, 14, INK, bold=True)
    r2 = p.add_run()
    r2.text = "Deliver early value with a focused MVP while building the foundation for continued capability and roadmap development."
    _style_run(r2, SANS, 14, BODY)

    # ----- Footer -----
    fy = 6.86
    textbox(
        slide, ML, fy, 1.1, 0.32, "PwC",
        font=SERIF, size=16, color=BLACK, bold=True,
        anchor=MSO_ANCHOR.MIDDLE, margins=(0, 0, 0, 0),
    )
    textbox(
        slide, 3.2, fy, 6.8, 0.32,
        "* Durations are estimates and will be refined.",
        font=SANS, size=9, color=MUTED, italic=True,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margins=(0, 0, 0, 0),
    )
    textbox(
        slide, MR - 0.40, fy, 0.40, 0.32, "4",
        font=SANS, size=11, color=MUTED,
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
