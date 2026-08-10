#!/usr/bin/env python3
"""
Презентация: «Признаки наружного кровотечения и кровопотери»

Оформление — как в «Организационно-правовые аспекты оказания первой помощи».
Содержание — по учебному пособию (Тема 2).
Редактируемый текст + отдельные Picture-объекты.
"""

from __future__ import annotations

from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

ROOT = Path("/workspace/presentations")
ASSETS = ROOT / "assets"
OUT = ROOT / "out"
OUT.mkdir(parents=True, exist_ok=True)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BG_LIGHT = RGBColor(0xF8, 0xF8, 0xF8)
BG_BAR = RGBColor(221, 221, 221)
LINE = RGBColor(0x99, 0x99, 0x99)
TEXT = RGBColor(0x33, 0x33, 0x33)
MUTED = RGBColor(0x66, 0x66, 0x66)
ACCENT_RED = RGBColor(0xED, 0x1C, 0x24)
CREAM = RGBColor(0xF5, 0xF0, 0xE1)
NUM_BG = RGBColor(0x55, 0x55, 0x55)
TABLE_HDR = RGBColor(0x44, 0x44, 0x44)
ROW_ALT = RGBColor(0xF5, 0xF5, 0xF5)

FONT = "Open Sans"
SUBTITLE = "Оказание первой помощи при наружных кровотечениях"
COURSE = "Тема 2 · Оказание первой помощи при наружных кровотечениях"


def asset(name: str) -> Path:
    return ASSETS / name


def new_prs() -> Presentation:
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    return prs


def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def font(run, size, bold=False, color=TEXT):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    for tag in ("latin", "ea", "cs"):
        el = rPr.find(qn(f"a:{tag}"))
        if el is None:
            el = etree.SubElement(rPr, qn(f"a:{tag}"))
        el.set("typeface", FONT)


def rect(slide, l, t, w, h, fill, line=None):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(1)
    return sh


def round_rect(slide, l, t, w, h, fill, line=None):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(1)
    return sh


def oval(slide, l, t, w, h, fill):
    sh = slide.shapes.add_shape(MSO_SHAPE.OVAL, l, t, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    sh.line.fill.background()
    return sh


def tbox(slide, l, t, w, h, text, *, size=20, bold=False, color=TEXT,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    try:
        tf._txBody.bodyPr.set(
            "anchor",
            {MSO_ANCHOR.TOP: "t", MSO_ANCHOR.MIDDLE: "ctr", MSO_ANCHOR.BOTTOM: "b"}[anchor],
        )
    except Exception:
        pass
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    font(r, size, bold, color)
    return box


def rich_tbox(slide, l, t, w, h, parts, *, size=20, color=TEXT,
              align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    try:
        tf._txBody.bodyPr.set(
            "anchor",
            {MSO_ANCHOR.TOP: "t", MSO_ANCHOR.MIDDLE: "ctr", MSO_ANCHOR.BOTTOM: "b"}[anchor],
        )
    except Exception:
        pass
    p = tf.paragraphs[0]
    p.alignment = align
    for text, bold in parts:
        r = p.add_run()
        r.text = text
        font(r, size, bold, color)
    return box


def bullets(slide, l, t, w, h, items, *, size=18, marker="•", alert=None):
    alert = alert or set()
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(8)
        r = p.add_run()
        r.text = f"{marker}  {item}"
        font(r, size, bold=(i in alert), color=ACCENT_RED if i in alert else TEXT)
    return box


def pic_fit(slide, name, l, t, max_w, max_h):
    from PIL import Image as PILImage
    path = asset(name)
    if not path.exists():
        rect(slide, l, t, max_w, max_h, WHITE, line=LINE)
        tbox(slide, l, t + max_h // 2 - Emu(200000), max_w, Emu(400000),
             f"[нет файла: {name}]", size=12, color=MUTED, align=PP_ALIGN.CENTER)
        return None
    with PILImage.open(path) as im:
        iw, ih = im.size
    scale = min(float(max_w) / iw, float(max_h) / ih)
    w, h = int(iw * scale), int(ih * scale)
    x = int(l + (max_w - w) / 2)
    y = int(t + (max_h - h) / 2)
    return slide.shapes.add_picture(str(path), x, y, width=w, height=h)


def slide_number(slide, n: int):
    size = Emu(420000)
    top = Emu(2800000)
    left = Emu(0)
    rect(slide, left, top, size, size, NUM_BG)
    tbox(slide, left, top, size, size, str(n), size=14, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def content_header(slide, title: str, num: int | None = None):
    rect(slide, 0, 0, SLIDE_W, SLIDE_H, WHITE)
    margin = Emu(700000)
    tbox(slide, margin, Emu(280000), Emu(11000000), Emu(550000),
         title, size=24, bold=True, color=TEXT)
    rect(slide, margin, Emu(880000), Emu(9000000), Emu(25000), NUM_BG)
    tbox(slide, margin, Emu(950000), Emu(11000000), Emu(350000),
         SUBTITLE, size=12, color=MUTED)
    if num is not None:
        slide_number(slide, num)


def slide_title(prs):
    slide = blank(prs)
    rect(slide, 0, 0, SLIDE_W, SLIDE_H, BG_LIGHT)
    rect(slide, Emu(8800000), 0, Emu(3400000), SLIDE_H, BG_BAR)
    rect(slide, Emu(0), Emu(2800000), Emu(2800000), Emu(1800000), BG_BAR)
    tbox(slide, Emu(700000), Emu(3000000), Emu(7800000), Emu(1400000),
         "ПРИЗНАКИ НАРУЖНОГО\nКРОВОТЕЧЕНИЯ И КРОВОПОТЕРИ",
         size=28, bold=True, color=TEXT)
    rect(slide, Emu(700000), Emu(4600000), Emu(5500000), Emu(30000), TEXT)
    tbox(slide, Emu(700000), Emu(4800000), Emu(7500000), Emu(500000),
         COURSE, size=14, color=MUTED)
    return slide


def slide_thanks(prs):
    slide = blank(prs)
    rect(slide, 0, 0, Emu(3200000), SLIDE_H, BG_BAR)
    rect(slide, Emu(3200000), 0, Emu(9000000), SLIDE_H, BG_LIGHT)
    tbox(slide, Emu(3800000), Emu(3000000), Emu(7500000), Emu(1000000),
         "БЛАГОДАРИМ ЗА ВНИМАНИЕ", size=32, bold=True, color=TEXT,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
    rect(slide, Emu(3800000), Emu(4100000), Emu(4500000), Emu(30000), TEXT)
    return slide


def slide_toc(prs, items, num=2):
    slide = blank(prs)
    content_header(slide, "СОДЕРЖАНИЕ", num)
    y = Emu(1550000)
    for i, item in enumerate(items, 1):
        oval(slide, Emu(700000), y, Emu(550000), Emu(550000), NUM_BG)
        tbox(slide, Emu(700000), y, Emu(550000), Emu(550000), str(i),
             size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE)
        tbox(slide, Emu(1450000), y, Emu(10000000), Emu(550000),
             item, size=17, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)
        y += Emu(850000)
    return slide


def slide_definition(prs, num=3):
    slide = blank(prs)
    content_header(slide, "ПОНЯТИЕ «КРОВОТЕЧЕНИЕ» И КРОВОПОТЕРЯ", num)
    pic_fit(slide, "def_man.png", Emu(500000), Emu(1500000),
            Emu(3000000), Emu(5200000))
    round_rect(slide, Emu(3300000), Emu(1600000), Emu(8600000), Emu(2800000), CREAM)
    rich_tbox(
        slide, Emu(3600000), Emu(1800000), Emu(8000000), Emu(2400000),
        [
            ("Под ", False),
            ("кровотечением понимают", True),
            (" ситуацию, когда кровь по разным причинам "
             "(чаще всего в результате травмы) покидает сосудистое русло, "
             "что приводит к ", False),
            ("кровопотере", True),
            (" — безвозвратной утрате части крови.", False),
        ],
        size=15, color=TEXT,
    )
    tbox(slide, Emu(3300000), Emu(4700000), Emu(8600000), Emu(1800000),
         "Это сопровождается снижением функции системы кровообращения "
         "по переносу кислорода и питательных веществ к органам, "
         "что ведет к ухудшению или прекращению их деятельности.",
         size=15, color=TEXT)
    return slide


def slide_signs_blood_loss(prs, num=4):
    slide = blank(prs)
    content_header(slide, "ПРИЗНАКИ КРОВОПОТЕРИ", num)
    left = [
        "резкая общая слабость;",
        "чувство жажды;",
        "головокружение;",
        "мелькание «мушек» перед глазами;",
    ]
    right = [
        "обморок (чаще при попытке встать);",
        "бледная, влажная и холодная кожа;",
        "учащённое сердцебиение;",
        "частое дыхание.",
    ]
    bullets(slide, Emu(700000), Emu(1600000), Emu(5400000), Emu(3200000),
            left, size=17, marker="☐")
    bullets(slide, Emu(6500000), Emu(1600000), Emu(5400000), Emu(3200000),
            right, size=17, marker="☐")
    round_rect(slide, Emu(700000), Emu(5000000), Emu(11000000), Emu(1400000), CREAM)
    rect(slide, Emu(700000), Emu(5000000), Emu(100000), Emu(1400000), ACCENT_RED)
    tbox(slide, Emu(1000000), Emu(5200000), Emu(10400000), Emu(1000000),
         "Эти признаки возможны при продолжающемся кровотечении, "
         "при уже остановленном кровотечении и при отсутствии видимой крови "
         "(в том числе при внутреннем кровотечении).",
         size=15, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)
    return slide


def slide_danger(prs, num=5):
    slide = blank(prs)
    content_header(slide, "ОСТРАЯ МАССИВНАЯ КРОВОПОТЕРЯ", num)
    round_rect(slide, Emu(700000), Emu(1550000), Emu(11000000), Emu(1400000), CREAM)
    tbox(slide, Emu(950000), Emu(1700000), Emu(10500000), Emu(1100000),
         "Наиболее опасно интенсивное кровотечение, приводящее к быстрой потере "
         "большого количества крови, — острая массивная кровопотеря.",
         size=17, bold=True, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)
    bullets(slide, Emu(700000), Emu(3300000), Emu(11000000), Emu(3200000), [
        "При повреждении крупных сосудов без остановки кровотечения "
        "гибель может наступить в течение нескольких минут.",
        "При кровотечении слабой и средней интенсивности организм обычно "
        "способен поддерживать жизнь, но кровотечение всё равно нужно остановить.",
        "Даже «неинтенсивная» кровопотеря может привести к поздним осложнениям "
        "травмы и ухудшить исход.",
    ], size=16)
    return slide


def slide_external_def(prs, num=6):
    slide = blank(prs)
    content_header(slide, "ЧТО ТАКОЕ НАРУЖНОЕ КРОВОТЕЧЕНИЕ", num)
    tbox(slide, Emu(700000), Emu(1600000), Emu(11000000), Emu(1200000),
         "Наружное кровотечение сопровождается повреждением кожных покровов "
         "и слизистых оболочек, при этом кровь изливается наружу "
         "в окружающую среду.",
         size=18, color=TEXT)
    tbox(slide, Emu(700000), Emu(3100000), Emu(11000000), Emu(500000),
         "По виду повреждённых сосудов кровотечения бывают:", size=16, bold=True)
    # four pills
    labels = ["Артериальное", "Венозное", "Капиллярное", "Смешанное"]
    x = Emu(700000)
    for lab in labels:
        round_rect(slide, x, Emu(3900000), Emu(2600000), Emu(900000), BG_BAR)
        tbox(slide, x, Emu(3900000), Emu(2600000), Emu(900000), lab,
             size=16, bold=True, color=TEXT, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE)
        x += Emu(2800000)
    tbox(slide, Emu(700000), Emu(5200000), Emu(11000000), Emu(1000000),
         "При оказании первой помощи вид определить сложно — "
         "ориентируйтесь на интенсивность кровотечения.",
         size=15, color=MUTED)
    return slide


def slide_types_visual(prs, num=7):
    slide = blank(prs)
    content_header(slide, "ПРИЗНАКИ ВИДОВ НАРУЖНОГО КРОВОТЕЧЕНИЯ", num)
    cols = [
        ("bleed_arterial.png", "Артериальное",
         "Пульсирующая алая струя; быстро растекающаяся лужа алого цвета; "
         "одежда быстро пропитывается кровью. Наиболее опасно."),
        ("bleed_venous.png", "Венозное",
         "Кровь тёмно-вишнёвая, вытекает «ручьём». "
         "Скорость кровопотери меньше, но остановка обязательна."),
        ("bleed_capillary.png", "Капиллярное",
         "Ссадины, порезы, царапины. "
         "Как правило, непосредственной угрозы жизни не представляет."),
    ]
    x = Emu(700000)
    cw = Emu(3600000)
    gap = Emu(300000)
    for img, title, desc in cols:
        tbox(slide, x, Emu(1450000), cw, Emu(450000), title, size=16, bold=True,
             color=TEXT, align=PP_ALIGN.CENTER)
        pic_fit(slide, img, x + Emu(150000), Emu(2000000), cw - Emu(300000), Emu(3000000))
        tbox(slide, x, Emu(5200000), cw, Emu(1400000), desc, size=12, color=TEXT,
             align=PP_ALIGN.CENTER)
        x += cw + gap
    return slide


def slide_mixed(prs, num=8):
    slide = blank(prs)
    content_header(slide, "СМЕШАННОЕ КРОВОТЕЧЕНИЕ", num)
    bullets(slide, Emu(700000), Emu(1600000), Emu(11000000), Emu(2800000), [
        "Одновременно артериальное, венозное и капиллярное кровотечение.",
        "Наблюдается, например, при отрыве конечности.",
        "Опасно вследствие наличия артериального компонента.",
    ], size=17)
    round_rect(slide, Emu(700000), Emu(4600000), Emu(11000000), Emu(1600000), CREAM)
    rect(slide, Emu(700000), Emu(4600000), Emu(100000), Emu(1600000), ACCENT_RED)
    tbox(slide, Emu(1000000), Emu(4850000), Emu(10400000), Emu(1100000),
         "При наличии кровотечения останавливайте его любым доступным способом "
         "или их комбинацией — не тратьте время на точное определение вида.",
         size=16, bold=True, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)
    return slide


def slide_compare_table(prs, num=9):
    slide = blank(prs)
    content_header(slide, "СРАВНЕНИЕ ВИДОВ КРОВОТЕЧЕНИЯ", num)
    headers = ["Признак", "Артериальное", "Венозное", "Капиллярное"]
    rows = [
        ["Цвет крови", "Алый, яркий", "Тёмный, вишнёвый", "Красный"],
        ["Характер", "Пульсирующей струёй", "Равномерным «ручьём»", "Сочится"],
        ["Скорость", "Очень быстрая", "Умеренная / быстрая", "Медленная"],
        ["Опасность", "Критическая — минуты", "Высокая", "Обычно низкая"],
    ]
    table = slide.shapes.add_table(
        len(rows) + 1, len(headers),
        Emu(700000), Emu(1550000), Emu(11000000), Emu(4500000)
    ).table
    for c, h in enumerate(headers):
        cell = table.cell(0, c)
        cell.fill.solid()
        cell.fill.fore_color.rgb = TABLE_HDR
        cell.text = h
        for p in cell.text_frame.paragraphs:
            for r in p.runs:
                font(r, 13, True, WHITE)
    for ri, row in enumerate(rows, 1):
        for c, val in enumerate(row):
            cell = table.cell(ri, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = ROW_ALT if ri % 2 == 0 else WHITE
            cell.text = val
            alert = any(k in val for k in ("Алый", "Пульсир", "Критическ", "Очень"))
            for p in cell.text_frame.paragraphs:
                for r in p.runs:
                    font(r, 13, alert, ACCENT_RED if alert else TEXT)
    return slide


def slide_intensity_signs(prs, num=10):
    slide = blank(prs)
    content_header(slide, "ПРИЗНАКИ ИНТЕНСИВНОГО КРОВОТЕЧЕНИЯ", num)
    tbox(slide, Emu(700000), Emu(1500000), Emu(11000000), Emu(800000),
         "Сигнал к немедленной остановке кровотечения:", size=16, bold=True)
    bullets(slide, Emu(700000), Emu(2400000), Emu(11000000), Emu(2500000), [
        "одежда, пропитанная кровью;",
        "скопление значительного количества крови на земле возле пострадавшего;",
        "видимые раны с интенсивно вытекающей из них кровью.",
    ], size=17, marker="☐")
    round_rect(slide, Emu(700000), Emu(5200000), Emu(11000000), Emu(1100000), CREAM)
    tbox(slide, Emu(950000), Emu(5350000), Emu(10500000), Emu(800000),
         "Обнаружив такие признаки — сразу приступайте к остановке "
         "всеми доступными способами.",
         size=16, bold=True, color=ACCENT_RED, anchor=MSO_ANCHOR.MIDDLE)
    return slide


def slide_summary(prs, num=11):
    slide = blank(prs)
    content_header(slide, "ГЛАВНОЕ ЗАПОМНИТЬ", num)
    points = [
        "Кровотечение → кровопотеря → риск нарушения работы органов.",
        "Признаки кровопотери: слабость, жажда, бледность, холодный пот, "
        "учащённый пульс и дыхание, обморок.",
        "Артериальное — алая пульсирующая струя; венозное — тёмный «ручей»; "
        "капиллярное — сочение.",
        "Ориентир для первой помощи — интенсивность, а не точный вид сосуда.",
    ]
    y = Emu(1550000)
    for pt in points:
        rect(slide, Emu(700000), y, Emu(11000000), Emu(1000000), WHITE, line=LINE)
        rect(slide, Emu(700000), y, Emu(90000), Emu(1000000), ACCENT_RED)
        tbox(slide, Emu(1000000), y + Emu(150000), Emu(10400000), Emu(700000),
             pt, size=15, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)
        y += Emu(1150000)
    return slide


def verify(path: Path):
    prs = Presentation(str(path))
    full = pics = texts = 0
    for s in prs.slides:
        for sh in s.shapes:
            if sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
                pics += 1
                if sh.width >= int(SLIDE_W * 0.95) and sh.height >= int(SLIDE_H * 0.95):
                    full += 1
            if sh.has_text_frame and sh.text_frame.text.strip():
                texts += 1
    if full:
        raise SystemExit(f"FAIL: full-slide images = {full}")
    print(f"OK: {path.name} · slides={len(prs.slides)} · pics={pics} · texts={texts} · "
          f"{path.stat().st_size // 1024} KB")


def build():
    prs = new_prs()
    slide_title(prs)
    slide_toc(prs, [
        "Понятие кровотечения и кровопотери",
        "Признаки кровопотери",
        "Острая массивная кровопотеря",
        "Виды наружного кровотечения и их признаки",
        "Интенсивность как главный ориентир",
    ], 2)
    slide_definition(prs, 3)
    slide_signs_blood_loss(prs, 4)
    slide_danger(prs, 5)
    slide_external_def(prs, 6)
    slide_types_visual(prs, 7)
    slide_mixed(prs, 8)
    slide_compare_table(prs, 9)
    slide_intensity_signs(prs, 10)
    slide_summary(prs, 11)
    slide_thanks(prs)

    name = "Признаки_наружного_кровотечения_и_кровопотери.pptx"
    path = OUT / name
    prs.save(path)
    verify(path)
    path2 = ROOT / name
    prs.save(path2)
    print(f"Saved: {path}")
    print(f"Saved: {path2}")
    return path


if __name__ == "__main__":
    build()
