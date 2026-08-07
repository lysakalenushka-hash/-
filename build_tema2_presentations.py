#!/usr/bin/env python3
"""
Тема 2 — 5 презентаций по наружным кровотечениям.
Макеты по принципу визуального баланса (9 типов слайдов).
Open Sans · редактируемый текст · иллюстрации как отдельные объекты.
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

OUT = Path("/workspace/tema2_bleeding")
ASSETS = OUT / "assets"
OUT.mkdir(parents=True, exist_ok=True)

SLIDE_W = Emu(24384000)
SLIDE_H = Emu(13716000)

# Поля ≈ 1 см
M = Emu(360000)
CONTENT_L = M
CONTENT_T = Emu(1600000)   # после зоны заголовка (~12%)
CONTENT_W = SLIDE_W - 2 * M
CONTENT_H = Emu(10500000)  # средняя зона до ~85%

WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BG_SOFT = RGBColor(0xF7, 0xF7, 0xF7)
BG_BAR = RGBColor(221, 221, 221)
LINE = RGBColor(225, 225, 225)
TEXT = RGBColor(0x33, 0x33, 0x33)
ACCENT_RED = RGBColor(0xCC, 0x00, 0x00)
ACCENT_BLUE = RGBColor(0x00, 0x55, 0xAA)
TABLE_HDR = RGBColor(0x44, 0x44, 0x44)
ROW_ALT = RGBColor(0xF5, 0xF5, 0xF5)

FONT = "Open Sans"
COURSE = "Тема 2 · Оказание первой помощи при наружных кровотечениях"


def asset(name: str) -> Path:
    return ASSETS / name


def new_prs() -> Presentation:
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    return prs


def blank(prs: Presentation):
    return prs.slides.add_slide(prs.slide_layouts[6])


def set_run_font(run, *, size, bold=False, color=TEXT):
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


def add_rect(slide, left, top, width, height, fill, line=None):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(1.25)
    return sh


def add_round(slide, left, top, width, height, fill, line=None):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(1.25)
    return sh


def add_oval(slide, left, top, width, height, fill):
    sh = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top, width, height)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    sh.line.fill.background()
    return sh


def add_textbox(slide, left, top, width, height, text, *, size=32, bold=False,
                color=TEXT, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(left, top, width, height)
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
    run = p.add_run()
    run.text = text
    set_run_font(run, size=size, bold=bold, color=color)
    return box


def add_bullets(slide, left, top, width, height, items, *, size=32, alert=None):
    alert = alert or set()
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(14)
        run = p.add_run()
        run.text = f"•  {item}"
        set_run_font(run, size=size, bold=(i in alert), color=ACCENT_RED if i in alert else TEXT)
    return box


def add_picture_fit(slide, image_path: Path, left, top, max_w, max_h):
    from PIL import Image as PILImage
    if not Path(image_path).exists():
        return None
    with PILImage.open(image_path) as im:
        iw, ih = im.size
    scale = min(float(max_w) / iw, float(max_h) / ih)
    w = int(iw * scale)
    h = int(ih * scale)
    x = int(left + (max_w - w) / 2)
    y = int(top + (max_h - h) / 2)
    return slide.shapes.add_picture(str(image_path), x, y, width=w, height=h)


def add_image_placeholder(slide, left, top, width, height, description: str):
    """Место под изображение, если в исходнике нет готовой иллюстрации."""
    add_rect(slide, left, top, width, height, WHITE, line=LINE)
    # пунктирная имитация — внутренний прямоугольник
    pad = Emu(80000)
    add_rect(slide, left + pad, top + pad, width - 2 * pad, height - 2 * pad, BG_SOFT, line=LINE)
    add_textbox(
        slide, left + Emu(200000), top + height // 2 - Emu(500000),
        width - Emu(400000), Emu(1000000),
        f"[МЕСТО ДЛЯ ИЗОБРАЖЕНИЯ: {description}]",
        size=18, bold=True, color=ACCENT_BLUE,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
    )


def bg(slide):
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, WHITE)


def header(slide, title: str, *, center=False):
    """Верхняя зона 0–15%: заголовок 46 pt."""
    bg(slide)
    add_textbox(
        slide, CONTENT_L, Emu(350000), CONTENT_W, Emu(1100000),
        title, size=46, bold=True, color=TEXT,
        align=PP_ALIGN.CENTER if center else PP_ALIGN.LEFT,
    )
    add_rect(slide, CONTENT_L, Emu(1450000), CONTENT_W, Emu(20000), LINE)


# ═══════════════ ТИПЫ МАКЕТОВ ═══════════════

def slide_title(prs, title: str, subtitle: str):
    """1. Титульный слайд."""
    slide = blank(prs)
    bg(slide)
    add_textbox(
        slide, CONTENT_L, Emu(4800000), CONTENT_W, Emu(1600000),
        title.upper(), size=46, bold=True, color=TEXT, align=PP_ALIGN.CENTER,
    )
    add_rect(slide, Emu(7000000), Emu(6600000), Emu(10000000), Emu(25000), ACCENT_BLUE)
    add_textbox(
        slide, CONTENT_L, Emu(6900000), CONTENT_W, Emu(800000),
        subtitle, size=26, bold=False, color=TEXT, align=PP_ALIGN.CENTER,
    )
    add_textbox(
        slide, CONTENT_L, Emu(7800000), CONTENT_W, Emu(500000),
        COURSE, size=20, bold=False, color=ACCENT_BLUE, align=PP_ALIGN.CENTER,
    )
    return slide


def slide_toc(prs, items: list[str]):
    """2. Оглавление."""
    slide = blank(prs)
    header(slide, "Содержание", center=True)
    icons = ["1", "2", "3", "4", "5"]
    y = CONTENT_T + Emu(400000)
    for i, item in enumerate(items[:5]):
        # иконка-номер
        add_oval(slide, CONTENT_L + Emu(2000000), y, Emu(700000), Emu(700000), ACCENT_BLUE)
        add_textbox(
            slide, CONTENT_L + Emu(2000000), y, Emu(700000), Emu(700000),
            icons[i], size=26, bold=True, color=WHITE,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        )
        add_textbox(
            slide, CONTENT_L + Emu(3000000), y, Emu(14000000), Emu(700000),
            item, size=32, bold=False, color=TEXT, anchor=MSO_ANCHOR.MIDDLE,
        )
        y += Emu(1400000)
    # декоративная линия справа
    add_rect(slide, Emu(22000000), CONTENT_T, Emu(80000), Emu(9000000), BG_BAR)
    return slide


def slide_info_icons(prs, title: str, items: list[tuple[str, str]], *, alert_last=False):
    """3. Информационный: тезисы + иконки слева.
    items: [(icon_label, text), ...] — max 5.
    """
    slide = blank(prs)
    header(slide, title)
    y = CONTENT_T + Emu(300000)
    for i, (icon, text) in enumerate(items[:5]):
        fill = ACCENT_RED if (alert_last and i == len(items) - 1) else ACCENT_BLUE
        add_round(slide, CONTENT_L, y, Emu(900000), Emu(900000), fill)
        add_textbox(
            slide, CONTENT_L, y, Emu(900000), Emu(900000),
            icon, size=22, bold=True, color=WHITE,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        )
        add_textbox(
            slide, CONTENT_L + Emu(1200000), y, CONTENT_W - Emu(1200000), Emu(900000),
            text, size=32, bold=False, color=TEXT, anchor=MSO_ANCHOR.MIDDLE,
        )
        y += Emu(1600000)
    return slide


def slide_scheme(prs, title: str, steps: list[str], *, key_idx: int | None = None,
                 horizontal: bool = False):
    """4. Слайд-схема (алгоритм)."""
    slide = blank(prs)
    header(slide, title)
    n = len(steps)
    if horizontal and n <= 4:
        gap = Emu(200000)
        arrow_w = Emu(350000)
        usable = CONTENT_W - arrow_w * (n - 1) - gap * (n - 1)
        bw = usable // n
        bh = Emu(3200000)
        y = CONTENT_T + Emu(2500000)
        x = CONTENT_L
        for i, step in enumerate(steps):
            accent = key_idx is not None and i == key_idx
            fill = ACCENT_BLUE if accent else BG_SOFT
            add_round(slide, x, y, bw, bh, fill, line=None if accent else LINE)
            add_textbox(
                slide, x + Emu(100000), y + Emu(200000), bw - Emu(200000), Emu(600000),
                str(i + 1), size=28, bold=True,
                color=WHITE if accent else ACCENT_BLUE, align=PP_ALIGN.CENTER,
            )
            add_textbox(
                slide, x + Emu(150000), y + Emu(1000000), bw - Emu(300000), Emu(1900000),
                step, size=24, bold=False,
                color=WHITE if accent else TEXT, align=PP_ALIGN.CENTER,
            )
            if i < n - 1:
                ax = x + bw + gap // 2
                sh = slide.shapes.add_shape(
                    MSO_SHAPE.RIGHT_ARROW, ax, y + bh // 2 - Emu(150000), arrow_w, Emu(300000)
                )
                sh.fill.solid()
                sh.fill.fore_color.rgb = ACCENT_BLUE
                sh.line.fill.background()
            x += bw + arrow_w + gap
    else:
        # vertical — занимает ~65% площади
        bh = min(Emu(1400000), CONTENT_H // max(n, 1) - Emu(200000))
        y = CONTENT_T + Emu(200000)
        for i, step in enumerate(steps):
            accent = key_idx is not None and i == key_idx
            fill = ACCENT_BLUE if accent else BG_SOFT
            add_round(slide, CONTENT_L + Emu(1500000), y, Emu(19000000), bh, fill,
                      line=None if accent else LINE)
            add_oval(slide, CONTENT_L + Emu(1800000), y + bh // 2 - Emu(400000),
                     Emu(800000), Emu(800000), WHITE if accent else ACCENT_BLUE)
            add_textbox(
                slide, CONTENT_L + Emu(1800000), y + bh // 2 - Emu(400000),
                Emu(800000), Emu(800000), str(i + 1), size=26, bold=True,
                color=ACCENT_BLUE if accent else WHITE,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
            )
            add_textbox(
                slide, CONTENT_L + Emu(2900000), y + Emu(200000),
                Emu(17000000), bh - Emu(400000),
                step, size=28, bold=False,
                color=WHITE if accent else TEXT, anchor=MSO_ANCHOR.MIDDLE,
            )
            if i < n - 1:
                sh = slide.shapes.add_shape(
                    MSO_SHAPE.DOWN_ARROW,
                    CONTENT_L + Emu(11000000), y + bh + Emu(20000),
                    Emu(350000), Emu(180000),
                )
                sh.fill.solid()
                sh.fill.fore_color.rgb = ACCENT_BLUE
                sh.line.fill.background()
            y += bh + Emu(250000)
    return slide


def slide_table(prs, title: str, headers: list[str], rows: list[list[str]], *, font_size=24):
    """5. Слайд-таблица (≤70% площади)."""
    slide = blank(prs)
    header(slide, title)
    n_rows = len(rows) + 1
    n_cols = len(headers)
    tw = Emu(20000000)
    th = min(Emu(7500000), Emu(1100000) * n_rows)
    left = CONTENT_L + (CONTENT_W - tw) // 2
    top = CONTENT_T + Emu(400000)
    table = slide.shapes.add_table(n_rows, n_cols, left, top, tw, th).table
    for c, h in enumerate(headers):
        cell = table.cell(0, c)
        cell.fill.solid()
        cell.fill.fore_color.rgb = TABLE_HDR
        cell.text = h
        for p in cell.text_frame.paragraphs:
            for r in p.runs:
                set_run_font(r, size=26, bold=True, color=WHITE)
    for ri, row in enumerate(rows, 1):
        for c, val in enumerate(row):
            cell = table.cell(ri, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = ROW_ALT if ri % 2 == 0 else WHITE
            cell.text = val
            alert = any(k in val for k in ("60", "30", "Алый", "Пульсир", "Критическ", "НЕ "))
            for p in cell.text_frame.paragraphs:
                for r in p.runs:
                    set_run_font(r, size=font_size, bold=alert,
                                 color=ACCENT_RED if alert else TEXT)
    return slide


def slide_image(prs, title: str, image_name: str | None, caption: str,
                *, placeholder: str | None = None, layout: str = "A"):
    """6. Слайд с изображением.
    layout A: текст 60% слева, изображение 40% справа
    layout V: изображение сверху 60%, подпись снизу
    """
    slide = blank(prs)
    header(slide, title)
    if layout == "V":
        img_box = (CONTENT_L, CONTENT_T + Emu(200000), CONTENT_W, Emu(6500000))
        cap_box = (CONTENT_L, CONTENT_T + Emu(7000000), CONTENT_W, Emu(2500000))
        add_rect(slide, *img_box, WHITE, line=LINE)
        if image_name and asset(image_name).exists():
            add_picture_fit(slide, asset(image_name), img_box[0] + Emu(150000),
                            img_box[1] + Emu(150000), img_box[2] - Emu(300000),
                            img_box[3] - Emu(300000))
        else:
            add_image_placeholder(slide, *img_box, placeholder or title)
        add_textbox(slide, *cap_box, caption, size=26, bold=False, color=TEXT,
                    align=PP_ALIGN.CENTER)
    else:  # A — text left, image right
        text_w = int(CONTENT_W * 0.55)
        img_w = int(CONTENT_W * 0.40)
        add_bullets(slide, CONTENT_L, CONTENT_T + Emu(400000), text_w, Emu(8500000),
                    [ln.strip() for ln in caption.split("\n") if ln.strip()], size=28)
        ix = CONTENT_L + text_w + Emu(400000)
        add_rect(slide, ix, CONTENT_T + Emu(400000), img_w, Emu(8500000), WHITE, line=LINE)
        if image_name and asset(image_name).exists():
            add_picture_fit(slide, asset(image_name), ix + Emu(100000),
                            CONTENT_T + Emu(500000), img_w - Emu(200000), Emu(8300000))
        else:
            add_image_placeholder(slide, ix, CONTENT_T + Emu(400000), img_w, Emu(8500000),
                                  placeholder or title)
    return slide


def slide_image_row(prs, title: str, cells: list[tuple[str | None, str]], *,
                    placeholders: list[str] | None = None):
    """Несколько иллюстраций в ряд + короткие подписи (тип 6, сетка)."""
    slide = blank(prs)
    header(slide, title)
    n = len(cells)
    gap = Emu(250000)
    cw = (CONTENT_W - gap * (n - 1)) // n
    ch = Emu(9000000)
    img_h = Emu(5800000)
    for i, (img, cap) in enumerate(cells):
        x = CONTENT_L + i * (cw + gap)
        y = CONTENT_T + Emu(300000)
        add_rect(slide, x, y, cw, ch, WHITE, line=LINE)
        if img and asset(img).exists():
            add_picture_fit(slide, asset(img), x + Emu(100000), y + Emu(100000),
                            cw - Emu(200000), img_h)
        else:
            desc = (placeholders[i] if placeholders else cap)[:80]
            add_image_placeholder(slide, x + Emu(100000), y + Emu(100000),
                                  cw - Emu(200000), img_h, desc)
        add_textbox(slide, x + Emu(150000), y + img_h + Emu(200000),
                    cw - Emu(300000), ch - img_h - Emu(300000),
                    cap, size=20, bold=False, color=TEXT, align=PP_ALIGN.CENTER)
    return slide


def slide_list(prs, title: str, items: list[str], *, alert: set[int] | None = None,
               highlight: str | None = None):
    """7. Слайд-список (4–6 коротких пунктов)."""
    slide = blank(prs)
    header(slide, title)
    items = items[:6]
    h = Emu(7500000) if not highlight else Emu(6500000)
    add_bullets(slide, CONTENT_L + Emu(500000), CONTENT_T + Emu(400000),
                CONTENT_W - Emu(1000000), h, items, size=32, alert=alert)
    if highlight:
        add_round(slide, CONTENT_L, Emu(11200000), CONTENT_W, Emu(1500000), BG_SOFT)
        add_rect(slide, CONTENT_L, Emu(11200000), Emu(120000), Emu(1500000), ACCENT_RED)
        add_textbox(slide, CONTENT_L + Emu(400000), Emu(11400000),
                    CONTENT_W - Emu(600000), Emu(1100000),
                    highlight, size=26, bold=True, color=ACCENT_RED,
                    anchor=MSO_ANCHOR.MIDDLE)
    return slide


def slide_list_image(prs, title: str, items: list[str], image_name: str | None,
                     *, placeholder: str | None = None, alert: set[int] | None = None,
                     highlight: str | None = None):
    """7+. Список слева (60%) + иллюстрация справа (40%)."""
    slide = blank(prs)
    header(slide, title)
    text_w = int(CONTENT_W * 0.55)
    img_w = int(CONTENT_W * 0.40)
    h = Emu(8000000) if not highlight else Emu(6800000)
    add_bullets(slide, CONTENT_L, CONTENT_T + Emu(300000), text_w, h,
                items[:6], size=28, alert=alert)
    ix = CONTENT_L + text_w + Emu(400000)
    iy = CONTENT_T + Emu(300000)
    add_rect(slide, ix, iy, img_w, h, WHITE, line=LINE)
    if image_name and asset(image_name).exists():
        add_picture_fit(slide, asset(image_name), ix + Emu(100000), iy + Emu(100000),
                        img_w - Emu(200000), h - Emu(200000))
    else:
        add_image_placeholder(slide, ix, iy, img_w, h, placeholder or title)
    if highlight:
        add_round(slide, CONTENT_L, Emu(11200000), CONTENT_W, Emu(1500000), BG_SOFT)
        add_rect(slide, CONTENT_L, Emu(11200000), Emu(120000), Emu(1500000), ACCENT_RED)
        add_textbox(slide, CONTENT_L + Emu(400000), Emu(11400000),
                    CONTENT_W - Emu(600000), Emu(1100000),
                    highlight, size=24, bold=True, color=ACCENT_RED,
                    anchor=MSO_ANCHOR.MIDDLE)
    return slide


def slide_summary(prs, points: list[str], callout: str | None = None):
    """8. Итоговый слайд."""
    slide = blank(prs)
    header(slide, "Главное запомнить", center=True)
    y = CONTENT_T + Emu(400000)
    for i, pt in enumerate(points[:4]):
        add_oval(slide, CONTENT_L + Emu(1500000), y + Emu(150000),
                 Emu(600000), Emu(600000), ACCENT_BLUE)
        add_textbox(
            slide, CONTENT_L + Emu(1500000), y + Emu(150000),
            Emu(600000), Emu(600000), "OK", size=24, bold=True, color=WHITE,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        )
        add_textbox(
            slide, CONTENT_L + Emu(2500000), y, Emu(18000000), Emu(900000),
            pt, size=28, bold=False, color=TEXT, anchor=MSO_ANCHOR.MIDDLE,
        )
        y += Emu(1600000)
    if callout:
        add_textbox(
            slide, CONTENT_L, Emu(11200000), CONTENT_W, Emu(1200000),
            callout, size=26, bold=True, color=ACCENT_RED, align=PP_ALIGN.CENTER,
        )
    return slide


def slide_thanks(prs):
    """9. Заключительный слайд."""
    slide = blank(prs)
    bg(slide)
    add_textbox(
        slide, CONTENT_L, Emu(5500000), CONTENT_W, Emu(1200000),
        "Благодарим за внимание", size=46, bold=True, color=TEXT, align=PP_ALIGN.CENTER,
    )
    add_rect(slide, Emu(7000000), Emu(7000000), Emu(10000000), Emu(25000), ACCENT_BLUE)
    add_textbox(
        slide, CONTENT_L, Emu(7400000), CONTENT_W, Emu(500000),
        COURSE, size=20, bold=False, color=TEXT, align=PP_ALIGN.CENTER,
    )
    return slide


def slide_section(prs, title: str):
    """Разделитель раздела (минималистичный)."""
    slide = blank(prs)
    bg(slide)
    add_textbox(
        slide, CONTENT_L, Emu(5500000), CONTENT_W, Emu(1400000),
        title, size=46, bold=True, color=TEXT, align=PP_ALIGN.CENTER,
    )
    add_rect(slide, Emu(7000000), Emu(7200000), Emu(10000000), Emu(25000), ACCENT_BLUE)
    return slide


# ═══════════════════ ПРЕЗЕНТАЦИЯ 1 ═══════════════════

def build_pres1():
    prs = new_prs()
    slide_title(prs, "Кровотечение и обзорный осмотр пострадавшего",
                "Теоретический модуль · п. 2.1")
    slide_toc(prs, [
        "Определение кровотечения и признаки кровопотери",
        "Цель и порядок обзорного осмотра",
        "Обзорный и подробный осмотр: различия",
        "Когда обзорный осмотр критически важен",
    ])
    slide_info_icons(prs, "Понятие «кровотечение»", [
        ("1", "Выход крови из сосудистого русла"),
        ("2", "Приводит к острой кровопотере"),
        ("3", "Наружное — видно снаружи (рана)"),
        ("4", "Требует немедленной остановки"),
    ])
    slide_info_icons(prs, "Признаки острой кровопотери", [
        ("!", "Слабость, жажда, головокружение"),
        ("•", "«Мушки» перед глазами, обморок"),
        ("•", "Бледная, влажная, холодная кожа"),
        ("•", "Учащённое сердцебиение и дыхание"),
    ])
    slide_list_image(
        prs, "Обзорный осмотр",
        [
            "Быстрая оценка с головы до ног",
            "Цель — найти наружное кровотечение",
            "Длительность: около 1–2 секунд",
            "Сразу после оценки безопасности",
            "При крови — сразу начать остановку",
        ],
        "overview_exam.png",
        placeholder="обзорный осмотр пострадавшего с головы до ног",
        highlight="Приоритет: сначала угрожающее жизни кровотечение",
    )
    slide_scheme(prs, "Порядок обзорного осмотра", [
        "Безопасность себе и пострадавшему",
        "Осмотр с головы до ног (1–2 с)",
        "Искать кровь, лужи, пульсирующую струю",
        "Остановить кровотечение",
        "Затем — сознание, дыхание, подробный осмотр",
    ], key_idx=3)
    # сравнение двух колонок
    slide = blank(prs)
    header(slide, "Обзорный и подробный осмотр")
    add_round(slide, CONTENT_L, CONTENT_T + Emu(300000), Emu(10800000), Emu(9000000), WHITE, line=LINE)
    add_round(slide, CONTENT_L + Emu(11800000), CONTENT_T + Emu(300000), Emu(10800000), Emu(9000000), WHITE, line=LINE)
    add_rect(slide, CONTENT_L, CONTENT_T + Emu(300000), Emu(10800000), Emu(1000000), ACCENT_BLUE)
    add_rect(slide, CONTENT_L + Emu(11800000), CONTENT_T + Emu(300000), Emu(10800000), Emu(1000000), BG_BAR)
    add_textbox(slide, CONTENT_L + Emu(200000), CONTENT_T + Emu(450000), Emu(10400000), Emu(700000),
                "Обзорный", size=26, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_textbox(slide, CONTENT_L + Emu(12000000), CONTENT_T + Emu(450000), Emu(10400000), Emu(700000),
                "Подробный", size=26, bold=True, color=TEXT, align=PP_ALIGN.CENTER)
    add_bullets(slide, CONTENT_L + Emu(400000), CONTENT_T + Emu(1600000), Emu(10000000), Emu(7000000), [
        "Найти наружное кровотечение",
        "1–2 секунды",
        "С головы до ног",
        "Сразу после оценки обстановки",
    ], size=26)
    add_bullets(slide, CONTENT_L + Emu(12200000), CONTENT_T + Emu(1600000), Emu(10000000), Emu(7000000), [
        "Выявить травмы и угрозы",
        "Несколько минут",
        "Голова → шея → грудь → …",
        "После остановки кровотечения",
    ], size=26)

    slide_info_icons(prs, "Когда осмотр критически важен", [
        ("1", "ДТП — раны под одеждой"),
        ("2", "Падение с высоты"),
        ("3", "Производственные травмы"),
        ("4", "Несколько пострадавших"),
    ])
    slide_summary(prs, [
        "Кровотечение — выход крови с риском острой кровопотери",
        "Обзорный осмотр — быстрый поиск кровотечения (1–2 с)",
        "Сначала останавливаем кровь, затем подробный осмотр",
        "Особенно важен при ДТП и падении с высоты",
    ], callout="Время — критический фактор!")
    slide_thanks(prs)
    path = OUT / "Кровотечение_и_обзорный_осмотр.pptx"
    prs.save(path)
    return path


# ═══════════════════ ПРЕЗЕНТАЦИЯ 2 ═══════════════════

def build_pres2():
    prs = new_prs()
    slide_title(prs, "Признаки наружного кровотечения и кровопотери",
                "Теоретический модуль · п. 2.2")
    slide_toc(prs, [
        "Виды наружного кровотечения",
        "Сравнительная характеристика",
        "Признаки кровопотери и объём",
        "Скрытое (внутреннее) кровотечение",
    ])
    slide_image_row(prs, "Виды наружного кровотечения", [
        ("bleed_arterial.png", "Артериальное\nАлая кровь, пульсирующая струя"),
        ("bleed_venous.png", "Венозное\nТёмная кровь, равномерная струя"),
        ("bleed_capillary.png", "Капиллярное\nСочится по поверхности"),
    ], placeholders=[
        "артериальное кровотечение — струя",
        "венозное кровотечение — тёмная струя",
        "капиллярное кровотечение — сочение",
    ])
    slide_table(prs, "Сравнение видов кровотечения",
                ["Признак", "Артериальное", "Венозное", "Капиллярное"],
                [
                    ["Цвет", "Алый", "Тёмный", "Красный"],
                    ["Характер", "Пульсирующей струёй", "Равномерной струёй", "Сочится"],
                    ["Скорость", "Очень быстрая", "Умеренная", "Медленная"],
                    ["Опасность", "Критическая", "Высокая", "Обычно низкая"],
                ], font_size=22)
    slide_info_icons(prs, "Признаки острой кровопотери", [
        ("1", "Слабость, жажда, головокружение"),
        ("2", "«Мушки», обморок при попытке встать"),
        ("3", "Бледная холодная влажная кожа"),
        ("4", "Тахикардия и частое дыхание"),
    ])
    slide_table(prs, "Объём кровопотери и состояние",
                ["Объём", "Доля ОЦК", "Состояние"],
                [
                    ["До ~500 мл", "≈ 10%", "Слабость, жажда"],
                    ["500–1000 мл", "≈ 10–20%", "Бледность, тахикардия"],
                    ["1000–1500 мл", "≈ 20–30%", "Обмороки, холодный пот"],
                    [">1500–2000 мл", "> 30%", "Шок, угроза жизни"],
                ], font_size=22)
    slide_list(prs, "Скрытое (внутреннее) кровотечение", [
        "Снаружи крови может не быть видно",
        "Бледность, холодный пот, слабость, жажда",
        "Слабый учащённый пульс, частое дыхание",
        "Боль и напряжение живота",
        "Действия: СМП, положение, холод, не поить",
    ], highlight="Внутреннее кровотечение на месте не останавливают — нужна СМП")
    slide_summary(prs, [
        "Артериальное — алая пульсирующая струя",
        "Венозное — тёмная струя; капиллярное — сочение",
        "Потеря >30% ОЦК угрожает жизни",
        "При подозрении на внутреннее — срочный вызов СМП",
    ], callout="Ориентируйтесь на признаки, а не на подсчёт миллилитров")
    slide_thanks(prs)
    path = OUT / "Признаки_наружного_кровотечения.pptx"
    prs.save(path)
    return path


# ═══════════════════ ПРЕЗЕНТАЦИЯ 3 ═══════════════════

def build_pres3():
    prs = new_prs()
    slide_title(prs, "Способы временной остановки наружного кровотечения",
                "Теоретический модуль · п. 2.3")
    slide_toc(prs, [
        "Прямое давление и давящая повязка",
        "Пальцевое прижатие и сгибание",
        "Жгут и подручные средства",
        "Алгоритм, записка, ошибки",
    ])

    slide_section(prs, "Прямое давление и давящая повязка")
    slide_image_row(prs, "Прямое давление и давящая повязка", [
        ("direct_pressure.png", "Прямое давление на рану\nОсновной способ (приказ № 220н)"),
        ("pressure_bandage.png", "Давящая повязка\nПосле остановки или если давление невозможно"),
    ])

    slide_section(prs, "Пальцевое прижатие артерий")
    slide_image_row(prs, "Сонная артерия", [
        ("carotid_point.png", "Точка: шея снаружи от гортани"),
        ("carotid_4fingers.png", "Давление четырьмя пальцами"),
        ("carotid_thumb.png", "Вариант: большим пальцем"),
    ], placeholders=[
        "точка прижатия сонной артерии",
        "прижатие сонной артерии четырьмя пальцами",
        "прижатие сонной артерии большим пальцем",
    ])
    slide_image_row(prs, "Подключичная артерия", [
        ("subclavian_1.png", "Ямка над ключицей к I ребру"),
        ("subclavian_2.png", "Четыре выпрямленных пальца"),
        ("subclavian_3.png", "Вариант: согнутыми пальцами"),
    ], placeholders=[
        "точка подключичной артерии",
        "прижатие подключичной артерии",
        "прижатие согнутыми пальцами",
    ])
    slide_image_row(prs, "Плечевая артерия", [
        ("brachial_1.png", "Внутренняя сторона плеча,\nсредняя треть"),
        ("brachial_2.png", "Четыре пальца,\nобхватывающие плечо"),
    ])
    slide_image_row(prs, "Подмышечная и бедренная", [
        ("axillary_2.png", "Подмышечная: впадина\nк плечевой кости"),
        ("femoral_2.png", "Бедренная: ниже паха,\nкулаком, весом тела"),
    ])
    slide_table(prs, "Точки пальцевого прижатия — сводка",
                ["Артерия", "Место", "Как давить"],
                [
                    ["Сонная", "Шея снаружи от гортани", "4 пальца / большой"],
                    ["Подключичная", "Над ключицей к I ребру", "4 пальца"],
                    ["Подмышечная", "Подмышечная впадина", "Жёсткие пальцы"],
                    ["Плечевая", "Между бицепсом и трицепсом", "Обхват плеча"],
                    ["Бедренная", "Ниже паховой складки", "Кулак, вес тела"],
                ], font_size=20)

    slide_image_row(prs, "Максимальное сгибание конечности", [
        ("flexion_arm.png", "Предплечье:\nвалик в локтевой сгиб"),
        ("flexion_leg.png", "Голень:\nвалик в подколенную ямку"),
        ("flexion_thigh.png", "Бедро:\nвалик в пах, колено к груди"),
    ])

    slide_section(prs, "Кровоостанавливающий жгут")
    slide_list_image(
        prs, "Показания к жгуту",
        [
            "Обширное повреждение конечности",
            "Отрыв конечности",
            "Давление и повязка неэффективны",
            "Накладывать на плечо или бедро",
            "Только на одежду / подкладку",
        ],
        "tourniquet_2.png",
        placeholder="наложение кровоостанавливающего жгута",
        highlight="t max ≤ 60 мин (тепло) · ≤ 30 мин (холод)",
    )
    slide_scheme(prs, "Алгоритм наложения жгута", [
        "Проверить показания",
        "Выше раны, на плечо/бедро",
        "На одежду; первый тур — макс. натяжение",
        "Кровь остановилась",
        "Записка со временем; жгут виден",
        "Лимит: 60 мин / 30 мин",
    ], key_idx=5)
    slide_image_row(prs, "Наложение жгута", [
        ("tourniquet_1.png", "1. Выше раны на одежду"),
        ("tourniquet_2.png", "2. Максимальное натяжение"),
        ("tourniquet_3.png", "3. Записка со временем"),
    ])
    slide_list_image(
        prs, "Подручные средства",
        [
            "Тесьма, платок, галстук, ремень",
            "Закрутка для натяжения",
            "Не использовать проволоку и леску",
            "Критерий: кровь остановилась",
            "Обязательна записка со временем",
        ],
        "improvised_tq.png",
        placeholder="импровизированный жгут из подручных средств",
    )
    slide_info_icons(prs, "Записка под жгут", [
        ("1", "Время наложения: ЧЧ:ММ"),
        ("2", "Кто наложил (по возможности)"),
        ("3", "Кратко — место происшествия"),
        ("•", "Записка на видном месте"),
    ], alert_last=False)
    slide_list(prs, "Ошибки при наложении жгута", [
        "Слабое натяжение — усиление кровотечения",
        "Слишком далеко от раны",
        "На голую кожу без подкладки",
        "Скрытие жгута одеждой",
        "Нет записки со временем",
        "Жгут без показаний",
    ])
    slide_list(prs, "Когда НЕЛЬЗЯ накладывать жгут", [
        "Кровь останавливается давлением / повязкой",
        "Капиллярное и умеренное венозное кровотечение",
        "Раны шеи, груди, живота, головы",
        "Нет конечности проксимальнее раны",
    ], alert={2}, highlight="На шею, грудь, живот и голову жгут не накладывают")
    slide_summary(prs, [
        "Приоритет: давление → повязка → жгут",
        "Пальцевое прижатие и сгибание — временные приёмы",
        "Жгут: на одежду, с запиской; ≤60 / ≤30 мин",
        "Гемостатики усиливают давление там, где уместно",
    ], callout="Время наложения жгута — критический фактор!")
    slide_thanks(prs)
    path = OUT / "Способы_временной_остановки_кровотечения.pptx"
    prs.save(path)
    return path


# ═══════════════════ ПРЕЗЕНТАЦИЯ 4 ═══════════════════

def build_pres4():
    prs = new_prs()
    slide_title(prs, "Последовательность мероприятий по остановке кровотечения",
                "Теоретический модуль · п. 2.4")
    slide_toc(prs, [
        "Полный алгоритм действий",
        "Приоритетность способов",
        "Профилактика травматического шока",
        "Подробный осмотр и контроль",
    ])
    slide_scheme(prs, "Полный алгоритм (чек-лист)", [
        "Безопасность (СИЗ)",
        "Обзорный осмотр",
        "Остановка кровотечения",
        "Признаки жизни",
        "СЛР при необходимости",
        "Вызов СМП",
        "Подробный осмотр",
        "Положение, тепло, передача СМП",
    ], key_idx=2)
    # разбить длинную схему — второй слайд с приоритетом способов
    slide_scheme(prs, "Приоритетность способов остановки", [
        "Прямое давление",
        "Давящая повязка",
        "Жгут по показаниям",
    ], key_idx=0, horizontal=True)

    slide_list(prs, "Оценка состояния пострадавшего", [
        "При массивном кровотечении — сначала кровь",
        "Сознание: обратитесь, потрясите за плечи",
        "Дыхание: смотрю–слушаю–ощущаю ≤10 с",
        "Сначала то, что угрожает жизни сейчас",
    ])
    slide_list(prs, "Приоритет при множественных травмах", [
        "Угрожающее жизни кровотечение",
        "Дыхательные пути и дыхание / СЛР",
        "Остальные кровотечения и раны",
        "Иммобилизация и положение",
        "При нескольких пострадавших — дети и массивные кровотечения",
    ])
    slide_info_icons(prs, "Травматический шок — признаки", [
        ("!", "Тяжёлая травма и сильное кровотечение"),
        ("2", "Учащённое дыхание и сердцебиение"),
        ("•", "Бледная холодная влажная кожа"),
        ("4", "Возбуждение → апатия"),
    ])
    slide_list_image(
        prs, "Профилактика травматического шока",
        [
            "Остановить кровотечение",
            "Оптимальное положение тела",
            "Иммобилизация конечностей",
            "Защита от переохлаждения",
            "Не кормить и не поить",
        ],
        "shock_prev.png",
        placeholder="меры профилактики травматического шока",
    )
    slide_image_row(prs, "Подробный осмотр — часть 1", [
        ("exam_head.png", "1. Голова"),
        ("exam_neck.png", "2. Шея"),
        ("exam_chest.png", "3. Грудь"),
    ])
    slide_image_row(prs, "Подробный осмотр — часть 2", [
        ("exam_abdomen.png", "4. Живот и таз"),
        ("exam_legs.png", "5. Ноги"),
        ("exam_arms.png", "6. Руки"),
    ])
    slide_list(prs, "После остановки кровотечения", [
        "Контроль повязки (при промокании — усилить)",
        "Сознание, дыхание, цвет кожи",
        "Жгут: время, видимость, записка",
        "Согреть, обеспечить покой",
        "Передать бригаде: что сделано и время жгута",
    ])
    slide_list(prs, "Когда и как ослаблять жгут", [
        "Планово ослаблять на месте не рекомендуется",
        "Если превышен лимит 60 / 30 мин и СМП нет:",
        "Подготовить давление / повязку",
        "Медленно ослабить, оценивая кровь",
        "При струе — сразу затянуть и обновить время",
    ], alert={1}, highlight="Лимит жгута: ≤ 60 мин (тепло) · ≤ 30 мин (холод)")
    slide_summary(prs, [
        "Безопасность → осмотр → кровь → признаки жизни → СМП",
        "Способы: давление → повязка → жгут",
        "Профилактика шока: кровь, положение, тепло",
        "После остановки — контроль до передачи СМП",
    ], callout="Сначала угрожающее жизни кровотечение!")
    slide_thanks(prs)
    path = OUT / "Последовательность_остановки_кровотечения.pptx"
    prs.save(path)
    return path


# ═══════════════════ ПРЕЗЕНТАЦИЯ 5 ═══════════════════

def build_pres5():
    prs = new_prs()
    slide_title(prs, "Остановка кровотечения при ранениях различных областей тела",
                "Теоретический модуль · п. 2.5")
    slide_toc(prs, [
        "Голова, глаза, нос",
        "Шея, грудь, инородный предмет",
        "Живот, таз, конечности",
        "Позвоночник и ампутация",
    ])

    slide_section(prs, "Голова · глаза · нос")
    slide_image_row(prs, "Травма головы", [
        ("head_trauma_1.png", "Боковое положение\nпри потере сознания"),
        ("head_trauma_2.png", "Прямое давление\n(не при открытой ЧМТ)"),
        ("head_trauma_3.png", "Давящая повязка\nна рану головы"),
    ])
    slide_image_row(prs, "Травмы глаза и носа", [
        ("eye_injury.png", "Глаз: повязка на оба глаза"),
        ("nose_injury.png", "Нос: вперёд, зажим, холод"),
    ])
    slide_image_row(prs, "Носовое кровотечение", [
        ("nose_1.png", "1. Усадить, голова вперёд"),
        ("nose_2.png", "2. Зажать на 10–15 мин"),
        ("nose_3.png", "3. Холод на переносицу"),
        ("nose_4.png", "4. Не высмаркиваться"),
    ])

    slide_section(prs, "Шея и грудь")
    slide_image_row(prs, "Травма шеи", [
        ("neck_1.png", "1. Остановка кровотечения"),
        ("neck_2.png", "2. Фиксация при перемещении"),
        ("neck_3.png", "3. Фиксация шейного отдела"),
    ])
    slide_list(prs, "Шея: артерия или вена", [
        "Артериальное: алая пульсирующая струя",
        "Венозное: тёмная обильная струя",
        "Прямое давление (не пережимать обе сонные)",
        "Повязка через плечо — не круговая тугая",
        "Жгут на шею НЕ накладывают",
    ], alert={4}, highlight="Не сдавливайте дыхательные пути повязкой")
    slide_image_row(prs, "Травма груди", [
        ("chest_1.png", "1. Полусидячее положение"),
        ("chest_2.png", "2. Окклюзионная повязка"),
        ("chest_3.png", "3. Предмет: обложить, повязка"),
    ])
    slide_scheme(prs, "Окклюзионная повязка — схема", [
        "Полусидячее положение",
        "Воздухонепроницаемый материал на рану",
        "Фиксация пластырем",
        "При сквозном — вход и выход",
        "Контроль дыхания",
    ], key_idx=1)
    slide_list(prs, "Инородный предмет в ране", [
        "НЕ ИЗВЛЕКАТЬ",
        "Обложить валиками вокруг",
        "Давящая повязка с фиксацией",
        "Не продавливать предмет вглубь",
        "Вызвать СМП",
    ], highlight="Правило: зафиксировать — не извлекать!")

    slide_section(prs, "Живот · таз · конечности")
    slide_list_image(
        prs, "Закрытая травма живота",
        [
            "Вызвать СМП",
            "Холод на живот",
            "На спине, ноги полусогнуты",
            "Не кормить и не поить",
        ],
        "abdomen_closed.png",
        placeholder="закрытая травма живота — холод и положение",
    )
    slide_list_image(
        prs, "Открытая травма живота",
        [
            "Органы НЕ вправлять",
            "Накрыть влажной тканью",
            "Повязка без давления на органы",
            "Срочный вызов СМП",
        ],
        "abdomen_open.png",
        placeholder="открытая травма живота — органы не вправлять",
        highlight="Органы брюшной полости не вправлять!",
    )
    slide_image_row(prs, "Травма конечностей", [
        ("limb_1.png", "1. Остановить кровотечение"),
        ("limb_2.png", "2. Иммобилизация 2 суставов"),
        ("limb_3.png", "3. Не на отломки"),
    ])
    slide_list_image(
        prs, "Иммобилизация",
        [
            "Неподвижность повреждённой части",
            "Шины поверх одежды",
            "Два соседних сустава",
            "Отломки не вправлять",
            "Сначала кровь — затем шина",
        ],
        "immobilization.png",
        placeholder="иммобилизация конечности шинами",
    )
    slide_list(prs, "Травматическая ампутация", [
        "Жгут выше отрыва (плечо / бедро)",
        "Давящая повязка на культю",
        "Записка со временем",
        "Сегмент: ткань → пакет → холод",
        "Передать сегмент вместе с пострадавшим",
    ], alert={0}, highlight="Жгут при отрыве показан. t max ≤ 60 / 30 мин")

    slide_section(prs, "Позвоночник")
    slide_image_row(prs, "Травма позвоночника", [
        ("spine_1.png", "Жёсткая ровная поверхность"),
        ("spine_2.png", "Несколько человек + фиксация шеи"),
    ])
    slide_list(prs, "Перемещение при травме позвоночника", [
        "Только при угрозе жизни или эвакуации",
        "Нужно 3–5 человек",
        "Один держит голову и шею постоянно",
        "Одно движение, без скручивания",
        "Фиксация шеи вручную / воротником",
    ])
    slide_summary(prs, [
        "Голова/нос: повязки; при потере сознания — боковое положение",
        "Шея и грудь: без кругового жгута; окклюзионная повязка",
        "Живот: органы не вправлять; конечности — кровь, затем шина",
        "Позвоночник — жёсткая поверхность и команда с фиксацией шеи",
    ], callout="Инородный предмет — зафиксировать, не извлекать!")
    slide_thanks(prs)
    path = OUT / "Кровотечение_при_ранениях_областей_тела.pptx"
    prs.save(path)
    return path


def main():
    paths = []
    for fn in (build_pres1, build_pres2, build_pres3, build_pres4, build_pres5):
        p = fn()
        print(f"OK: {p.name} · {p.stat().st_size // 1024} KB")
        paths.append(p)

    for p in paths:
        prs = Presentation(str(p))
        pics = full = 0
        for s in prs.slides:
            for sh in s.shapes:
                if sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    pics += 1
                    if sh.width >= SLIDE_W * 0.95 and sh.height >= SLIDE_H * 0.95:
                        full += 1
        print(f"  {p.name}: {len(prs.slides)} slides, pics={pics}, full-slide={full}")
        if full:
            raise SystemExit(f"Full-slide image in {p.name}")

    (OUT / "README.md").write_text(
        "# Тема 2. Наружные кровотечения\n\n"
        "Пять презентаций с макетами визуального баланса "
        "(титул, оглавление, иконки, схемы, таблицы, иллюстрации, резюме).\n\n"
        "Сборка: `python build_tema2_presentations.py`\n",
        encoding="utf-8",
    )
    print("Done.")


if __name__ == "__main__":
    main()
