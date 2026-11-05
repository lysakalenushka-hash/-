#!/usr/bin/env python3
"""Презентация темы 1.1 в журнальном стиле референса (светлый фон, navy, крупные заголовки)."""

from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageEnhance, ImageFilter, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import nsmap, qn
from pptx.util import Emu, Inches, Pt
from lxml import etree

ASSETS = Path("/workspace/pptx_assets")
OUT = Path("/workspace/1.1_Общие_требования_промышленной_безопасности.pptx")

# Palette from the reference board
BG = RGBColor(0xC9, 0xD2, 0xDC)
CARD = RGBColor(0xF6, 0xF7, 0xF4)
NAVY = RGBColor(0x2C, 0x3D, 0x5B)
NAVY2 = RGBColor(0x3D, 0x51, 0x74)
INK = RGBColor(0x2A, 0x33, 0x44)
MUTED = RGBColor(0x6B, 0x75, 0x84)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LINE = RGBColor(0x2C, 0x3D, 0x5B)

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
    tf.auto_size = None
    tf.margin_left = Inches(margin)
    tf.margin_right = Inches(margin)
    tf.margin_top = Inches(margin * 0.6)
    tf.margin_bottom = Inches(margin * 0.4)
    return tf


def add_text(slide, l, t, w, h, text, *, size=18, bold=False, color=INK, name=SANS,
             align=PP_ALIGN.LEFT, italic=False, anchor=MSO_ANCHOR.TOP, margin=0.08):
    box = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = _tf(box, margin)
    tf.paragraphs[0].alignment = align
    try:
        box.text_frame._txBody.bodyPr.set("anchor", {MSO_ANCHOR.TOP: "t", MSO_ANCHOR.MIDDLE: "ctr", MSO_ANCHOR.BOTTOM: "b"}.get(anchor, "t"))
    except Exception:
        pass
    p = tf.paragraphs[0]
    p.alignment = align
    p.space_after = Pt(0)
    p.space_before = Pt(0)
    run = p.add_run()
    run.text = text
    _set_run(run, size=size, bold=bold, color=color, name=name, italic=italic)
    return box


def add_paras(slide, l, t, w, h, lines, *, size=15, color=INK, name=SANS, bold=False, spacing=8, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = _tf(box, 0.06)
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(spacing)
        p.level = 0
        run = p.add_run()
        run.text = line
        _set_run(run, size=size, bold=bold, color=color, name=name)
    return box


def rect(slide, l, t, w, h, fill, *, rounded=True, adj=0.08, line=None):
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
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
        shape.line.width = Pt(1)
    shape.shadow.inherit = False
    return shape


def bar(slide, l, t, w, h=0.035, fill=NAVY):
    return rect(slide, l, t, w, h, fill, rounded=False)


def picture(slide, path, l, t, w, h):
    return slide.shapes.add_picture(str(path), Inches(l), Inches(t), Inches(w), Inches(h))


def new_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    # cool grey canvas
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, W, H)
    bg.fill.solid()
    bg.fill.fore_color.rgb = BG
    bg.line.fill.background()
    # white card
    rect(slide, 0.28, 0.22, 12.78, 7.06, CARD, adj=0.04)
    return slide


def footer(slide, page, total=16):
    add_text(slide, 0.55, 6.95, 6, 0.28, "МОДУЛЬ 1  ·  ТЕМА 1.1", size=10, color=MUTED, name=SANS, bold=True)
    add_text(slide, 10.6, 6.95, 2.1, 0.28, f"{page:02d}  /  {total:02d}", size=10, color=MUTED, name=SANS, align=PP_ALIGN.RIGHT)


def label(slide, l, t, text="TRAINING"):
    add_text(slide, l, t, 4, 0.28, text, size=11, color=MUTED, name=SANS, bold=True)


def tone_photo(src: Path, dst: Path, size=(1600, 1200), crop="center"):
    im = Image.open(src).convert("RGB")
    tw, th = size
    r = max(tw / im.width, th / im.height)
    im = im.resize((int(im.width * r) + 1, int(im.height * r) + 1), Image.Resampling.LANCZOS)
    left = (im.width - tw) // 2
    top = (im.height - th) // 2
    if crop == "top":
        top = 0
    elif crop == "bottom":
        top = im.height - th
    im = im.crop((left, top, left + tw, top + th))
    im = ImageEnhance.Color(im).enhance(0.55)
    im = ImageEnhance.Contrast(im).enhance(1.08)
    im = ImageEnhance.Brightness(im).enhance(1.05)
    im.save(dst, quality=88)
    return dst


def draw_pentagon_ru(path: Path):
    W, H = 1400, 1320
    im = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    cx, cy, r = 700, 680, 430
    import math
    pts = []
    for i in range(5):
        a = math.radians(-90 + i * 72)
        pts.append((cx + r * math.cos(a), cy + r * math.sin(a)))
    # outer fill
    d.polygon(pts, fill=(44, 61, 91, 255))
    # inner pentagon
    r2 = r * 0.58
    pts2 = []
    for i in range(5):
        a = math.radians(-90 + i * 72)
        pts2.append((cx + r2 * math.cos(a), cy + r2 * math.sin(a)))
    d.polygon(pts2, fill=(246, 247, 244, 255))
    # spokes
    for p in pts:
        d.line([(cx, cy), p], fill=(246, 247, 244, 255), width=10)
    font_b = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 42)
    font_c = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 36)
    labels = [
        (0, "ТОПЛИВО\nпыль"),
        (1, "КИСЛОРОД"),
        (2, "ЗАЖИГАНИЕ"),
        (3, "ВЗВИХРЕНИЕ"),
        (4, "ЗАМКНУТЫЙ\nОБЪЁМ"),
    ]
    # place labels around ring midpoints
    for i, text in labels:
        a1 = math.radians(-90 + i * 72)
        a2 = math.radians(-90 + ((i + 1) % 5) * 72)
        mx = cx + (r * 0.78) * math.cos((a1 + a2) / 2)
        my = cy + (r * 0.78) * math.sin((a1 + a2) / 2)
        # simpler: put at vertex mid-ring using sequential vertices
    # vertex-based placement
    mids = []
    for i in range(5):
        x = (pts[i][0] + pts[(i + 1) % 5][0]) / 2
        y = (pts[i][1] + pts[(i + 1) % 5][1]) / 2
        # pull toward center a bit
        x = cx + (x - cx) * 0.82
        y = cy + (y - cy) * 0.82
        mids.append((x, y))
    texts = ["ПЫЛЬ", "КИСЛОРОД", "ЗАЖИГАНИЕ", "ВЗВИХРЕНИЕ", "ОБЪЁМ"]
    # remap: top vertex is fuel in our construction? vertex 0 is top.
    # sides: 0-1 right-top, 1-2 right-bot, 2-3 bot, 3-4 left-bot, 4-0 left-top
    order = ["ЗАЖИГАНИЕ", "ВЗВИХРЕНИЕ", "ОБЪЁМ", "ПЫЛЬ", "КИСЛОРОД"]
    for (x, y), t in zip(mids, order):
        bbox = d.textbbox((0, 0), t, font=font_b)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        d.text((x - tw / 2, y - th / 2), t, font=font_b, fill=(255, 255, 255, 255))
    center = "ВЗРЫВ\nПЫЛИ"
    bbox = d.textbbox((0, 0), center, font=font_c, align="center")
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    d.text((cx - tw / 2, cy - th / 2 - 8), center, font=font_c, fill=(44, 61, 91, 255), align="center")
    im.save(path)
    return path


def prepare_assets():
    tone_photo(ASSETS / "silos.jpg", ASSETS / "silos_ed.jpg", (1800, 1200), "center")
    tone_photo(ASSETS / "silos.jpg", ASSETS / "silos_cover.jpg", (1100, 1600), "center")
    tone_photo(ASSETS / "silos2.jpg", ASSETS / "silos2_ed.jpg", (1000, 1500), "center")
    tone_photo(ASSETS / "wheat.jpg", ASSETS / "wheat_ed.jpg", (1200, 1500), "center")
    tone_photo(ASSETS / "mill.jpg", ASSETS / "mill_ed.jpg", (1600, 1100), "center")
    tone_photo(ASSETS / "dryer.jpg", ASSETS / "dryer_ed.jpg", (1100, 1500), "center")
    tone_photo(ASSETS / "lockout.jpg", ASSETS / "lockout_ed.jpg", (1500, 1100), "center")
    tone_photo(ASSETS / "dust_demo.jpg", ASSETS / "dust_ed.jpg", (1600, 1100), "center")
    if (ASSETS / "becherwerk.jpg").exists():
        tone_photo(ASSETS / "becherwerk.jpg", ASSETS / "norie_ed.jpg", (700, 1200), "center")
    draw_pentagon_ru(ASSETS / "pentagon_ru.png")


def build():
    prepare_assets()
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H

    # 01 cover
    s = new_slide(prs)
    label(s, 0.7, 0.48)
    add_text(s, 0.7, 0.85, 7.2, 0.4, "МОДУЛЬ 1", size=16, color=NAVY, bold=True, name=SANS)
    bar(s, 0.72, 1.28, 2.4, 0.045)
    add_text(s, 0.7, 1.5, 7.4, 2.6, "Общие требования\nпромышленной\nбезопасности", size=36, bold=False, color=NAVY, name=SERIF)
    add_text(
        s, 0.72, 4.35, 6.8, 1.4,
        "На объектах хранения и переработки\nрастительного сырья",
        size=18, color=INK, name=SANS,
    )
    add_text(s, 0.72, 6.15, 6.8, 0.5, "Курс: рабочий по обслуживанию объектов", size=13, color=MUTED)
    picture(s, ASSETS / "silos_cover.jpg", 8.55, 0.55, 4.2, 6.4)
    footer(s, 1)

    # 02 contents — like актуализация знаний
    s = new_slide(prs)
    add_text(s, 0.7, 0.5, 12, 0.9, "Содержание", size=36, color=NAVY, name=SERIF)
    bar(s, 0.72, 1.4, 2.1, 0.04)
    items = [
        ("01", "Что такое промышленная безопасность"),
        ("02", "Пылевой взрыв и источники зажигания"),
        ("03", "Самосогревание сырья"),
        ("04", "Оборудование, сушилки, силосы"),
        ("05", "Пыль, организация работ, меры предупреждения"),
    ]
    y = 1.7
    for n, t in items:
        add_text(s, 0.85, y, 1.2, 0.55, f"//{n}", size=20, color=NAVY, name=SANS, bold=True)
        add_text(s, 2.2, y, 9.5, 0.55, t, size=20, color=INK, name=SANS)
        y += 0.78
    footer(s, 2)

    # 03 definition — rose analog = wheat
    s = new_slide(prs)
    label(s, 0.7, 0.42, "ОПРЕДЕЛЕНИЕ")
    add_text(s, 0.7, 0.75, 7.4, 1.2, "Промышленная\nбезопасность", size=32, color=NAVY, name=SERIF)
    bar(s, 0.72, 2.05, 1.8, 0.04)
    add_text(
        s, 0.72, 2.3, 7.0, 1.6,
        "Состояние защищённости от аварий на опасных производственных объектах и от последствий таких аварий.",
        size=18, color=INK,
    )
    add_text(s, 0.72, 4.05, 7.0, 0.4, "Для работника это значит:", size=14, color=MUTED, bold=True)
    add_paras(s, 0.72, 4.4, 7.0, 2.2, [
        "работать по установленным правилам;",
        "распознавать опасные признаки;",
        "своевременно сообщать о нарушениях.",
    ], size=16, spacing=10)
    picture(s, ASSETS / "wheat_ed.jpg", 8.4, 0.55, 4.35, 6.4)
    footer(s, 3)

    # 04 why this industry
    s = new_slide(prs)
    add_text(s, 0.65, 0.42, 12, 0.7, "Почему отрасль особая", size=32, color=NAVY, name=SERIF)
    bar(s, 0.67, 1.15, 2.0, 0.04)
    cards = [
        ("01", "Сырьё", "Зерно, мука, комбикорма, семена — сухие органические материалы. Сами по себе горючи."),
        ("02", "Пыль", "При пересыпании и дроблении образуется горючая пыль. Она оседает и снова поднимается в воздух."),
        ("03", "Ёмкости", "Силосы, бункеры, нории и сушилки — замкнутые объёмы, где авария развивается быстро."),
        ("04", "Сочетание", "Нарушение режима + неисправность + источник зажигания = пожар, взрыв, разрушение."),
    ]
    positions = [(0.65, 1.5), (6.7, 1.5), (0.65, 4.15), (6.7, 4.15)]
    for (n, title, body), (x, y) in zip(cards, positions):
        rect(s, x, y, 5.85, 2.4, NAVY if n in ("02", "03") else RGBColor(0xEE, 0xF1, 0xF4), adj=0.08)
        c_title = WHITE if n in ("02", "03") else NAVY
        c_body = RGBColor(0xE4, 0xE8, 0xEE) if n in ("02", "03") else INK
        add_text(s, x + 0.25, y + 0.18, 1.2, 0.4, n, size=18, color=c_title, bold=True)
        add_text(s, x + 0.25, y + 0.55, 5.3, 0.45, title, size=20, color=c_title, name=SERIF)
        add_text(s, x + 0.25, y + 1.1, 5.3, 1.05, body, size=14, color=c_body)
    footer(s, 4)

    # 05 pentagon
    s = new_slide(prs)
    add_text(s, 0.65, 0.4, 7, 0.7, "Пылевой взрыв", size=32, color=NAVY, name=SERIF)
    bar(s, 0.67, 1.12, 1.8, 0.04)
    add_text(
        s, 0.65, 1.35, 6.3, 1.3,
        "Взрыв возможен, только если одновременно есть все пять условий. Уберите одно — взрыва нет.",
        size=16, color=INK,
    )
    picture(s, ASSETS / "pentagon.png", 7.35, 1.15, 4.85, 4.6)
    add_paras(s, 0.65, 2.8, 6.3, 3.6, [
        "Пыль — топливо: зерно, мука, комбикорм.",
        "Кислород — воздух в помещении и оборудовании.",
        "Зажигание — искра, перегрев, разряд, огневые работы.",
        "Взвихрение — пыль во взвешенном состоянии.",
        "Замкнутый объём — силос, нория, фильтр, цех.",
    ], size=15, spacing=12)
    add_text(s, 7.2, 5.85, 5.4, 0.85, "Fuel — пыль. Oxygen — воздух. Ignition — зажигание. Dispersion — взвихрение. Confinement — замкнутый объём. Схема NIOSH, общественное достояние.", size=11, color=MUTED, italic=True)
    footer(s, 5)

    # 06 primary / secondary + explosion photo
    s = new_slide(prs)
    add_text(s, 0.65, 0.4, 12, 0.65, "Первичный и вторичный взрыв", size=30, color=NAVY, name=SERIF)
    bar(s, 0.67, 1.1, 2.0, 0.04)
    rect(s, 0.65, 1.45, 5.9, 3.55, RGBColor(0xEE, 0xF1, 0xF4), adj=0.08)
    add_text(s, 0.9, 1.65, 5.4, 0.4, "01  ·  Первичный", size=16, color=NAVY, bold=True)
    add_text(s, 0.9, 2.15, 5.4, 2.4, "Вспышка внутри нории, фильтра, силоса или другого замкнутого объёма. Ударная волна поднимает осевшую пыль.", size=16, color=INK)
    rect(s, 6.75, 1.45, 5.9, 3.55, NAVY, adj=0.08)
    add_text(s, 7.0, 1.65, 5.4, 0.4, "02  ·  Вторичный", size=16, color=WHITE, bold=True)
    add_text(s, 7.0, 2.15, 5.4, 2.4, "Поднятая пыль взрывается повторно. Именно вторичные взрывы разрушают здания и убивают людей.", size=16, color=WHITE)
    add_text(s, 0.7, 5.2, 8.5, 1.4, "Что повышает вероятность: накопление пыли, неисправности, искрение и перегрев, статика, огневые работы без порядка, отказ взрывозащиты.", size=15, color=INK)
    footer(s, 6)

    # 07 fire
    s = new_slide(prs)
    add_text(s, 0.65, 0.38, 8, 0.65, "Пожар и возгорание сырья", size=30, color=NAVY, name=SERIF)
    bar(s, 0.67, 1.08, 1.9, 0.04)
    picture(s, ASSETS / "iso_f001.png", 9.55, 0.5, 3.1, 3.1)
    add_text(s, 0.65, 1.3, 8.4, 1.1, "Сырьё и пыль горючи. Достаточно пламени, искры или нагретой поверхности. Огонь идёт по отложениям пыли и по технологическому тракту.", size=16, color=INK)
    srcs = [
        ("Электрика", "Неисправные двигатели, повреждённая изоляция"),
        ("Механика", "Перегретые подшипники, трение, удары металла"),
        ("Сушилка", "Нарушение температурного режима"),
        ("Люди", "Сварка без допуска, курение в запрещённых местах"),
    ]
    x0 = 0.65
    for i, (t, b) in enumerate(srcs):
        x = x0 + (i % 2) * 4.25
        y = 2.7 + (i // 2) * 1.7
        rect(s, x, y, 4.05, 1.5, NAVY if i % 2 == 0 else RGBColor(0xEE, 0xF1, 0xF4), adj=0.1)
        c1 = WHITE if i % 2 == 0 else NAVY
        c2 = RGBColor(0xDD, 0xE2, 0xEA) if i % 2 == 0 else INK
        add_text(s, x + 0.2, y + 0.18, 3.6, 0.4, t, size=16, color=c1, name=SERIF)
        add_text(s, x + 0.2, y + 0.65, 3.6, 0.65, b, size=13, color=c2)
    add_text(s, 9.3, 4.0, 3.5, 2.4, "Дым отравляет и сбивает с маршрута эвакуации. Пожар при пыли может перейти во взрыв.", size=14, color=INK)
    add_text(s, 9.3, 6.15, 3.5, 0.55, "Знак ISO 7010 F001, общественное достояние.", size=9, color=MUTED, italic=True)
    footer(s, 7)

    # 08 self heating
    s = new_slide(prs)
    add_text(s, 0.65, 0.4, 12, 0.65, "Самосогревание и самовозгорание", size=30, color=NAVY, name=SERIF)
    bar(s, 0.67, 1.1, 2.0, 0.04)
    add_text(s, 0.65, 1.3, 12, 0.8, "Влажность, слабая вентиляция и биохимические процессы поднимают температуру внутри насыпи. Если тепло не уходит — участок тлеет.", size=16, color=INK)
    signs = [
        ("01", "Рост температуры сырья или участка хранилища"),
        ("02", "Затхлый запах, запах перегрева или гари"),
        ("03", "Дым, признаки тления"),
        ("04", "Изменение цвета и структуры материала"),
        ("05", "Отклонение приборов контроля"),
    ]
    y = 2.25
    for n, t in signs:
        rect(s, 0.65, y, 7.7, 0.72, RGBColor(0xEE, 0xF1, 0xF4), adj=0.18)
        add_text(s, 0.85, y + 0.12, 1.0, 0.5, n, size=16, color=NAVY, bold=True)
        add_text(s, 1.9, y + 0.14, 6.2, 0.5, t, size=16, color=INK)
        y += 0.82
    rect(s, 8.6, 2.25, 4.05, 4.05, NAVY, adj=0.08)
    add_text(s, 8.85, 2.5, 3.6, 0.5, "Запрещено", size=20, color=WHITE, name=SERIF)
    add_text(s, 8.85, 3.15, 3.6, 2.8, "Самостоятельно вскрывать оборудование, перекладывать нагретое сырьё и «гасить» очаг без разрешения и мер безопасности.", size=15, color=WHITE)
    footer(s, 8)

    # 09 equipment
    s = new_slide(prs)
    add_text(s, 0.65, 0.38, 8, 0.6, "Технологическое оборудование", size=28, color=NAVY, name=SERIF)
    bar(s, 0.67, 1.05, 1.9, 0.04)
    picture(s, ASSETS / "mill_ed.jpg", 8.35, 0.5, 4.4, 3.15)
    risks = [
        ("Захват", "Ленты, валы, цепи, шнеки затягивают одежду и руки."),
        ("Перегрев", "Подшипники и приводы становятся источником зажигания."),
        ("Искры", "Металл в механизме, удары, трение."),
        ("Завалы", "Подпор продукта, сход ленты, заклинивание, перегруз."),
        ("Аспирация", "Слабая вытяжка = пыль в помещении и внутри машин."),
    ]
    y = 1.3
    for title, body in risks:
        add_text(s, 0.7, y, 2.1, 0.55, title, size=16, color=NAVY, name=SERIF, bold=True)
        add_text(s, 2.85, y, 5.2, 0.7, body, size=15, color=INK)
        y += 0.85
    add_text(s, 8.35, 3.85, 4.4, 2.5, "Слышите стук, скрежет, чувствуете запах гари или вибрацию — действуйте по инструкции и сообщите. Не снимайте ограждения на ходу.", size=15, color=INK)
    footer(s, 9)

    # 10 silos diagrams
    s = new_slide(prs)
    add_text(s, 0.55, 0.32, 12, 0.55, "Силосы и бункеры: три сценария гибели", size=26, color=NAVY, name=SERIF)
    bar(s, 0.57, 0.92, 1.8, 0.035)
    pics = [
        (ASSETS / "bin_img_p1_0.png", "Свод", "Стоять на «мосте» из слежавшегося зерна. Свод обрушивается — человек проваливается."),
        (ASSETS / "bin_img_p1_1.png", "Поток", "Выгрузка идёт — зерно как зыбучий песок. 2–3 секунды на реакцию."),
        (ASSETS / "bin_img_p1_2.png", "Стенка", "Масса со стенки обрушивается на того, кто пытается её обрушить."),
    ]
    for i, (p, title, body) in enumerate(pics):
        x = 0.5 + i * 4.2
        rect(s, x, 1.15, 4.0, 5.4, RGBColor(0xF0, 0xF2, 0xF0), adj=0.06)
        picture(s, p, x + 0.15, 1.3, 3.7, 1.55)
        add_text(s, x + 0.2, 3.0, 3.6, 0.45, f"0{i+1}  {title}", size=18, color=NAVY, name=SERIF)
        add_text(s, x + 0.2, 3.55, 3.6, 2.5, body, size=14, color=INK)
    add_text(s, 0.55, 6.6, 12, 0.3, "Схемы: OSHA, Worker Entry into Grain Storage Bins. Общественное достояние США. Учебная наглядность, не норма РФ.", size=10, color=MUTED, italic=True)
    footer(s, 10)

    # 11 silo rules + confined sign
    s = new_slide(prs)
    add_text(s, 0.6, 0.38, 8, 0.6, "Вход внутрь — не осмотр", size=30, color=NAVY, name=SERIF)
    bar(s, 0.62, 1.05, 1.8, 0.04)
    picture(s, ASSETS / "confined.png", 8.55, 0.55, 4.15, 2.85)
    rules = [
        "Не входить самовольно и в одиночку.",
        "Изолировать подачу, выгрузку и связанные механизмы.",
        "Проверить атмосферу, если это предусмотрено порядком.",
        "Наблюдающий остаётся снаружи. Он не бросается внутрь «спасать».",
        "Не стоять на сыпучем продукте и не ходить по своду.",
        "Импровизированное спасение часто даёт второго погибшего.",
    ]
    y = 1.3
    for i, t in enumerate(rules, 1):
        add_text(s, 0.7, y, 0.7, 0.5, f"{i:02d}", size=16, color=NAVY, bold=True)
        add_text(s, 1.5, y, 6.7, 0.55, t, size=16, color=INK)
        y += 0.72
    add_text(s, 8.55, 3.55, 4.2, 2.8, "Знак из приказа Минтруда России № 902н. Вход только по наряду-допуску.", size=14, color=INK)
    footer(s, 11)

    # 12 dryer
    s = new_slide(prs)
    picture(s, ASSETS / "dryer_ed.jpg", 0.5, 0.5, 5.15, 6.15)
    add_text(s, 6.0, 0.55, 6.6, 1.1, "Сушка сырья", size=34, color=NAVY, name=SERIF)
    bar(s, 6.02, 1.7, 1.8, 0.04)
    add_text(s, 6.0, 1.95, 6.5, 1.1, "Высокая температура, поток воздуха и горючий материал в одной установке.", size=16, color=INK)
    pts = [
        "перегрев сырья и узлов;",
        "воспламенение пыли в воздуховодах;",
        "сбой горелки и подачи топлива;",
        "отказ автоматики и датчиков температуры;",
        "ошибка при розжиге и обслуживании.",
    ]
    add_paras(s, 6.0, 3.15, 6.5, 2.6, pts, size=16, spacing=8)
    add_text(s, 6.0, 5.85, 6.5, 0.9, "Дым, запах гари, скачок температуры, отказ автоматики — сообщить и действовать по инструкции агрегата. Не открывать камеру самовольно.", size=14, color=NAVY, bold=True)
    footer(s, 12)

    # 13 ignition + lockout
    s = new_slide(prs)
    add_text(s, 0.6, 0.38, 7.5, 0.6, "Источники зажигания", size=30, color=NAVY, name=SERIF)
    bar(s, 0.62, 1.05, 1.8, 0.04)
    picture(s, ASSETS / "lockout_ed.jpg", 8.2, 0.5, 4.5, 3.4)
    add_text(s, 0.6, 1.3, 7.3, 0.9, "Сухое сырьё при пересыпании накапливает статику. Разряд — достаточная искра для пылевоздушной смеси.", size=16, color=INK)
    left = [
        "неисправное электрооборудование;",
        "повреждённая изоляция;",
        "искрение механизмов;",
        "трение и удары металла;",
        "открытый огонь и курение;",
        "огневые работы без порядка.",
    ]
    add_paras(s, 0.6, 2.35, 7.3, 3.0, left, size=16, spacing=8)
    add_text(s, 8.2, 4.05, 4.5, 2.4, "Ремонт — только после безопасного отключения и блокировки. Срочность линии не отменяет lockout. Фото: блокировка механизма, CC BY-SA, Wikimedia.", size=13, color=INK)
    footer(s, 13)

    # 14 dust health + people
    s = new_slide(prs)
    add_text(s, 0.6, 0.38, 12, 0.6, "Пыль бьёт дважды", size=30, color=NAVY, name=SERIF)
    bar(s, 0.62, 1.05, 1.8, 0.04)
    rect(s, 0.6, 1.4, 6.0, 5.05, NAVY, adj=0.07)
    add_text(s, 0.9, 1.65, 5.4, 0.5, "По здоровью", size=20, color=WHITE, name=SERIF)
    add_paras(s, 0.9, 2.3, 5.4, 3.8, [
        "раздражение глаз и слизистых;",
        "кашель, затруднение дыхания;",
        "раздражение кожи;",
        "аллергия у чувствительных;",
        "нарастание нагрузки при дырявом оборудовании и мёртвой аспирации.",
    ], size=16, color=WHITE, spacing=12)
    rect(s, 6.85, 1.4, 5.85, 5.05, RGBColor(0xEE, 0xF1, 0xF4), adj=0.07)
    add_text(s, 7.15, 1.65, 5.4, 0.5, "По аварии", size=20, color=NAVY, name=SERIF)
    add_paras(s, 7.15, 2.3, 5.3, 3.8, [
        "осевшая пыль — запас топлива для вторичного взрыва;",
        "уборка сжатым воздухом поднимает облако;",
        "нужны исправная аспирация, порядок уборки и СИЗ;",
        "пыль — и гигиена, и взрыв.",
    ], size=16, color=INK, spacing=12)
    footer(s, 14)

    # 15 human factors
    s = new_slide(prs)
    add_text(s, 0.6, 0.38, 12, 0.6, "Когда подводит организация", size=30, color=NAVY, name=SERIF)
    bar(s, 0.62, 1.05, 1.8, 0.04)
    facts = [
        ("01", "Инструкции", "Работа не по режиму и не по инструкции."),
        ("02", "Допуск", "Человек на месте без подготовки."),
        ("03", "Неисправности", "Видели — не сообщили, не устранили."),
        ("04", "Блокировки", "Сбросили сигнал и пошли дальше."),
        ("05", "Уборка", "Пыль подняли в воздух «чтобы быстрее»."),
        ("06", "Ремонт", "Полезли в механизм без отключения."),
        ("07", "Силос", "Самовольно ковыряют зависание."),
        ("08", "Молчание", "Дым, запах, стук — «само пройдёт»."),
    ]
    for i, (n, t, b) in enumerate(facts):
        x = 0.55 + (i % 4) * 3.15
        y = 1.4 + (i // 4) * 2.55
        fill = NAVY if i in (1, 2, 5, 6) else RGBColor(0xEE, 0xF1, 0xF4)
        c1 = WHITE if fill == NAVY else NAVY
        c2 = RGBColor(0xDD, 0xE2, 0xEA) if fill == NAVY else INK
        rect(s, x, y, 3.0, 2.35, fill, adj=0.1)
        add_text(s, x + 0.18, y + 0.15, 2.6, 0.35, n, size=14, color=c1, bold=True)
        add_text(s, x + 0.18, y + 0.5, 2.6, 0.45, t, size=16, color=c1, name=SERIF)
        add_text(s, x + 0.18, y + 1.05, 2.6, 1.05, b, size=13, color=c2)
    footer(s, 15)

    # 16 prevention
    s = new_slide(prs)
    add_text(s, 0.6, 0.32, 12, 0.55, "Что держит аварию", size=30, color=NAVY, name=SERIF)
    bar(s, 0.62, 0.95, 1.8, 0.035)
    steps = [
        ("01", "Исправность", "Оборудование, аспирация, защиты и блокировки — в работе, не «на бумаге»."),
        ("02", "Пыль", "Убирать по порядку предприятия. Не поднимать облако."),
        ("03", "Температура", "Смотреть сырьё и узлы. Признаки тления — сразу руководителю."),
        ("04", "Огонь", "Нет самовольных источников зажигания. Огневые работы — только по допуску."),
        ("05", "Инструкции", "Нории, дробилки, сушилки — как написано, не «как привыкли»."),
        ("06", "Ремонт", "Сначала безопасный вывод из работы, потом руки в механизм."),
        ("07", "Ёмкости", "В силос и бункер — только с допуском. Зависания не бить с ходу."),
        ("08", "Связь", "Дым, гарь, шум, вибрация, рост температуры — сообщение не откладывают."),
        ("09", "План", "Знать оповещение, выход, действия при пожаре на своём участке."),
    ]
    for i, (n, t, b) in enumerate(steps):
        x = 0.5 + (i % 3) * 4.2
        y = 1.2 + (i // 3) * 1.85
        rect(s, x, y, 4.0, 1.7, RGBColor(0xEE, 0xF1, 0xF4) if i % 2 == 0 else NAVY, adj=0.1)
        fill_navy = i % 2 == 1
        c1 = WHITE if fill_navy else NAVY
        c2 = RGBColor(0xDD, 0xE2, 0xEA) if fill_navy else INK
        add_text(s, x + 0.18, y + 0.1, 3.6, 0.32, n, size=13, color=c1, bold=True)
        add_text(s, x + 0.18, y + 0.4, 3.6, 0.35, t, size=16, color=c1, name=SERIF)
        add_text(s, x + 0.18, y + 0.8, 3.6, 0.75, b, size=12, color=c2)
    footer(s, 16)

    prs.save(str(OUT))
    print("saved", OUT, "slides", len(prs.slides))


if __name__ == "__main__":
    build()
