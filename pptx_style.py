"""Общий журнальный стиль презентаций модуля 1."""

from __future__ import annotations

from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

ASSETS = Path("/workspace/pptx_assets")

BG = RGBColor(0xC9, 0xD2, 0xDC)
CARD = RGBColor(0xF6, 0xF7, 0xF4)
NAVY = RGBColor(0x2C, 0x3D, 0x5B)
PALE = RGBColor(0xEE, 0xF1, 0xF4)
INK = RGBColor(0x2A, 0x33, 0x44)
MUTED = RGBColor(0x6B, 0x75, 0x84)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
SOFT = RGBColor(0xDD, 0xE2, 0xEA)

SERIF = "Georgia"
SANS = "Calibri"
W, H = Inches(13.333), Inches(7.5)


def _set_run(run, *, size=18, bold=False, color=INK, name=SANS, italic=False):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = name
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("a:rFonts"))
    if rFonts is None:
        rFonts = etree.SubElement(rPr, qn("a:rFonts"))
    for a in ("ascii", "hAnsi", "cs", "eastAsia"):
        rFonts.set(a, name)


def _tf(shape, margin=0.12):
    tf = shape.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(margin)
    tf.margin_right = Inches(margin)
    tf.margin_top = Inches(margin * 0.6)
    tf.margin_bottom = Inches(margin * 0.4)
    return tf


def add_text(slide, l, t, w, h, text, *, size=18, bold=False, color=INK, name=SANS,
             align=PP_ALIGN.LEFT, italic=False, margin=0.08):
    box = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = _tf(box, margin)
    p = tf.paragraphs[0]
    p.alignment = align
    p.space_after = Pt(0)
    p.space_before = Pt(0)
    run = p.add_run()
    run.text = text
    _set_run(run, size=size, bold=bold, color=color, name=name, italic=italic)
    return box


def add_paras(slide, l, t, w, h, lines, *, size=15, color=INK, name=SANS, bold=False, spacing=8):
    box = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = _tf(box, 0.06)
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(spacing)
        run = p.add_run()
        run.text = line
        _set_run(run, size=size, bold=bold, color=color, name=name)
    return box


def rect(slide, l, t, w, h, fill, *, rounded=True, adj=0.08):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE,
        Inches(l), Inches(t), Inches(w), Inches(h),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if rounded:
        try:
            shape.adjustments[0] = adj
        except Exception:
            pass
    shape.line.fill.background()
    shape.shadow.inherit = False
    return shape


def bar(slide, l, t, w, h=0.035, fill=NAVY):
    return rect(slide, l, t, w, h, fill, rounded=False)


def picture(slide, path, l, t, w, h):
    return slide.shapes.add_picture(str(path), Inches(l), Inches(t), Inches(w), Inches(h))


def new_prs():
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H
    return prs


def new_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, W, H)
    bg.fill.solid()
    bg.fill.fore_color.rgb = BG
    bg.line.fill.background()
    rect(slide, 0.28, 0.22, 12.78, 7.06, CARD, adj=0.04)
    return slide


def footer(slide, page, total, topic="ТЕМА 1.1"):
    add_text(slide, 0.55, 6.95, 7.2, 0.28, f"МОДУЛЬ 1  ·  {topic}", size=10, color=MUTED, bold=True)
    add_text(slide, 10.6, 6.95, 2.1, 0.28, f"{page:02d}  /  {total:02d}", size=10, color=MUTED, align=PP_ALIGN.RIGHT)


def label(slide, l, t, text="TRAINING"):
    add_text(slide, l, t, 5, 0.28, text, size=11, color=MUTED, bold=True)


def cover(prs, *, topic, title, subtitle, photo, total, kicker="TRAINING"):
    s = new_slide(prs)
    label(s, 0.7, 0.48, kicker)
    add_text(s, 0.7, 0.85, 7.2, 0.4, "МОДУЛЬ 1", size=16, color=NAVY, bold=True)
    bar(s, 0.72, 1.28, 2.4, 0.045)
    add_text(s, 0.7, 1.5, 7.4, 2.8, title, size=34, color=NAVY, name=SERIF)
    add_text(s, 0.72, 4.5, 6.8, 1.4, subtitle, size=18, color=INK)
    add_text(s, 0.72, 6.15, 6.8, 0.5, "Курс: рабочий по обслуживанию объектов", size=13, color=MUTED)
    picture(s, photo, 8.55, 0.55, 4.2, 6.4)
    footer(s, 1, total, topic)
    return s


def contents(prs, *, topic, items, total):
    s = new_slide(prs)
    add_text(s, 0.7, 0.5, 12, 0.9, "Содержание", size=36, color=NAVY, name=SERIF)
    bar(s, 0.72, 1.4, 2.1, 0.04)
    y = 1.7
    for n, t in items:
        add_text(s, 0.85, y, 1.2, 0.55, f"//{n}", size=20, color=NAVY, bold=True)
        add_text(s, 2.2, y, 9.5, 0.55, t, size=20, color=INK)
        y += 0.78
    footer(s, 2, total, topic)
    return s


def cards(slide, items, origin=(0.65, 1.45), size=(5.85, 2.35), gap=(6.05, 2.55), cols=2):
    for i, (n, title, body) in enumerate(items):
        x = origin[0] + (i % cols) * gap[0]
        y = origin[1] + (i // cols) * gap[1]
        navy = i % 2 == 1 or i in (1, 2, 5, 6)
        fill = NAVY if navy else PALE
        c1 = WHITE if navy else NAVY
        c2 = SOFT if navy else INK
        rect(slide, x, y, size[0], size[1], fill, adj=0.08)
        add_text(slide, x + 0.22, y + 0.14, 1.4, 0.35, n, size=14, color=c1, bold=True)
        add_text(slide, x + 0.22, y + 0.48, size[0] - 0.4, 0.4, title, size=18, color=c1, name=SERIF)
        add_text(slide, x + 0.22, y + 0.95, size[0] - 0.4, size[1] - 1.1, body, size=13, color=c2)


def numbered(slide, lines, l=0.7, t=1.35, w=11.8):
    y = t
    for i, text in enumerate(lines, 1):
        add_text(slide, l, y, 0.7, 0.5, f"{i:02d}", size=16, color=NAVY, bold=True)
        add_text(slide, l + 0.8, y, w, 0.55, text, size=16, color=INK)
        y += 0.68
    return y
