#!/usr/bin/env python3
"""
Тема 2 — 5 редактируемых презентаций по наружным кровотечениям.
Стиль: Open Sans, белый фон, без логотипов.
Слайд ≠ картинка: иллюстрации из исходника — отдельные объекты рядом с текстом.
"""

from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt
from lxml import etree

OUT = Path("/workspace/tema2_bleeding")
OUT.mkdir(parents=True, exist_ok=True)
ASSETS = OUT / "assets"


def asset(name: str) -> Path:
    """Путь к вырезанной иллюстрации из исходного PDF."""
    return ASSETS / name


# --- Размер ---
SLIDE_W = Emu(24384000)  # 16:9
SLIDE_H = Emu(13716000)

# --- Цвета ---
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BG_SOFT = RGBColor(0xFA, 0xFA, 0xFA)
BG_BAR = RGBColor(221, 221, 221)       # только фоновые элементы
LINE = RGBColor(225, 225, 225)        # рамки / разделители
TEXT = RGBColor(0x33, 0x33, 0x33)     # #333333
BLACK = RGBColor(0x00, 0x00, 0x00)
ACCENT_RED = RGBColor(0xCC, 0x00, 0x00)
ACCENT_BLUE = RGBColor(0x00, 0x55, 0xAA)
TABLE_HDR = RGBColor(0x00, 0x55, 0xAA)
ROW_ALT = RGBColor(0xF5, 0xF5, 0xF5)

FONT = "Open Sans"
COURSE = "Тема 2 · Оказание первой помощи при наружных кровотечениях"


# ───────────────────── helpers ─────────────────────

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
    # East Asian / latin fallback hint for Open Sans
    rPr = run._r.get_or_add_rPr()
    for tag in ("latin", "ea", "cs"):
        el = rPr.find(qn(f"a:{tag}"))
        if el is None:
            el = etree.SubElement(rPr, qn(f"a:{tag}"))
        el.set("typeface", FONT)


def add_rect(slide, left, top, width, height, fill: RGBColor, line: RGBColor | None = None):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(1)
    return sh


def add_round(slide, left, top, width, height, fill: RGBColor):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    sh.line.fill.background()
    return sh


def add_oval(slide, left, top, width, height, fill: RGBColor):
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


def chrome(slide):
    """Белый фон + тонкая верхняя полоса (не логотип)."""
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, WHITE)
    add_rect(slide, 0, 0, SLIDE_W, Emu(120000), BG_BAR)


def title_bar(slide, title: str, subtitle: str | None = None):
    chrome(slide)
    add_textbox(
        slide, Emu(800000), Emu(350000), Emu(22800000), Emu(1000000),
        title, size=46, bold=True, color=TEXT, align=PP_ALIGN.LEFT,
    )
    add_rect(slide, Emu(800000), Emu(1450000), Emu(22800000), Emu(20000), LINE)
    if subtitle:
        add_textbox(
            slide, Emu(800000), Emu(1550000), Emu(22800000), Emu(600000),
            subtitle, size=26, bold=True, color=ACCENT_BLUE, align=PP_ALIGN.LEFT,
        )
        return Emu(2300000)  # content top
    return Emu(1750000)


def bullets(slide, left, top, width, height, items: list[str], *, size=32,
            alert_idx: set[int] | None = None):
    alert_idx = alert_idx or set()
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(10)
        p.level = 0
        run = p.add_run()
        is_alert = i in alert_idx
        run.text = f"•  {item}"
        set_run_font(
            run,
            size=size,
            bold=is_alert,
            color=ACCENT_RED if is_alert else TEXT,
        )
    return box


def add_icon_circle(slide, left, top, size, label: str, fill=ACCENT_BLUE):
    """Минималистичная «иконка»: круг + буква/цифра."""
    add_oval(slide, left, top, size, size, fill)
    add_textbox(
        slide, left, top, size, size, label,
        size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
    )


def add_flow_step(slide, left, top, width, height, num: str, text: str, accent=False):
    fill = ACCENT_BLUE if accent else BG_BAR
    add_round(slide, left, top, width, height, fill)
    add_textbox(
        slide, left + Emu(80000), top + Emu(60000), Emu(500000), height - Emu(120000),
        num, size=24, bold=True, color=WHITE if accent else TEXT, anchor=MSO_ANCHOR.MIDDLE,
    )
    add_textbox(
        slide, left + Emu(600000), top + Emu(80000), width - Emu(700000), height - Emu(160000),
        text, size=20, bold=False, color=WHITE if accent else TEXT, anchor=MSO_ANCHOR.MIDDLE,
    )


def add_arrow_right(slide, left, top):
    sh = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, left, top, Emu(400000), Emu(250000))
    sh.fill.solid()
    sh.fill.fore_color.rgb = ACCENT_BLUE
    sh.line.fill.background()


def add_table(slide, left, top, width, height, headers: list[str], rows: list[list[str]],
              font_size=24):
    n_rows = len(rows) + 1
    n_cols = len(headers)
    table = slide.shapes.add_table(n_rows, n_cols, left, top, width, height).table
    for c, h in enumerate(headers):
        cell = table.cell(0, c)
        cell.fill.solid()
        cell.fill.fore_color.rgb = TABLE_HDR
        cell.text = h
        for p in cell.text_frame.paragraphs:
            p.alignment = PP_ALIGN.LEFT
            for r in p.runs:
                set_run_font(r, size=font_size, bold=True, color=WHITE)
    for ri, row in enumerate(rows, 1):
        for c, val in enumerate(row):
            cell = table.cell(ri, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = ROW_ALT if ri % 2 == 0 else WHITE
            cell.text = val
            alert = any(k in val for k in ("60", "30", "Алый", "Пульсир", "Критическ", "НЕ "))
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.LEFT
                for r in p.runs:
                    set_run_font(
                        r,
                        size=font_size,
                        bold=alert,
                        color=ACCENT_RED if alert else TEXT,
                    )
    return table



def add_picture_fit(slide, image_path: Path, left, top, max_w, max_h):
    """Вставить изображение с сохранением пропорций внутри max_w x max_h."""
    from PIL import Image as PILImage
    if not Path(image_path).exists():
        return None
    with PILImage.open(image_path) as im:
        iw, ih = im.size
    scale = min(max_w / iw, max_h / ih)
    w = int(iw * scale)
    h = int(ih * scale)
    # center inside box
    x = int(left + (max_w - w) / 2)
    y = int(top + (max_h - h) / 2)
    return slide.shapes.add_picture(str(image_path), x, y, width=w, height=h)


def slide_text_image(prs, title: str, items: list[str], image_name: str,
                     *, subtitle: str | None = None, highlight: str | None = None,
                     alert_idx: set[int] | None = None, image_right: bool = True):
    """Редактируемый текст + иллюстрация сбоку (не фон слайда)."""
    slide = blank(prs)
    y0 = title_bar(slide, title, subtitle)
    text_w = Emu(12000000)
    img_w = Emu(10000000)
    if image_right:
        tx_left, img_left = Emu(800000), Emu(13500000)
    else:
        tx_left, img_left = Emu(12000000), Emu(800000)
    h = Emu(7500000) if not highlight else Emu(6000000)
    bullets(slide, tx_left, y0 + Emu(200000), text_w, h, items, size=28, alert_idx=alert_idx)
    add_rect(slide, img_left, y0 + Emu(200000), img_w, Emu(8500000), WHITE, line=LINE)
    add_picture_fit(slide, asset(image_name), img_left + Emu(150000), y0 + Emu(350000),
                    img_w - Emu(300000), Emu(8200000))
    if highlight:
        add_round(slide, Emu(800000), Emu(11200000), Emu(22800000), Emu(1600000), BG_SOFT)
        add_rect(slide, Emu(800000), Emu(11200000), Emu(120000), Emu(1600000), ACCENT_RED)
        add_textbox(slide, Emu(1200000), Emu(11400000), Emu(21800000), Emu(1200000),
                    highlight, size=24, bold=True, color=ACCENT_RED, anchor=MSO_ANCHOR.MIDDLE)
    return slide


def slide_image_grid(prs, title: str, cells: list[tuple[str, str]], *, cols: int = 3,
                     subtitle: str | None = None, caption_size: int = 20):
    """Сетка: иллюстрация сверху + редактируемый текст снизу в каждой ячейке."""
    slide = blank(prs)
    y0 = title_bar(slide, title, subtitle)
    n = len(cells)
    cols = min(cols, n)
    rows = (n + cols - 1) // cols
    gap = Emu(250000)
    total_w = Emu(22800000)
    cell_w = int((total_w - gap * (cols - 1)) / cols)
    avail_h = Emu(10500000) - (y0 - Emu(1750000))
    cell_h = int((avail_h - gap * (rows - 1)) / rows)
    img_h = int(cell_h * 0.62)
    for i, (img_name, caption) in enumerate(cells):
        r, c = divmod(i, cols)
        x = Emu(800000) + c * (cell_w + gap)
        y = y0 + Emu(150000) + r * (cell_h + gap)
        add_rect(slide, x, y, cell_w, cell_h, WHITE, line=LINE)
        add_picture_fit(slide, asset(img_name), x + Emu(100000), y + Emu(100000),
                        cell_w - Emu(200000), img_h - Emu(150000))
        add_textbox(slide, x + Emu(150000), y + img_h, cell_w - Emu(300000),
                    cell_h - img_h - Emu(100000), caption, size=caption_size,
                    bold=False, color=TEXT)
    return slide


def title_slide(prs, title: str, point: str):
    slide = blank(prs)
    chrome(slide)
    add_rect(slide, 0, Emu(4200000), SLIDE_W, Emu(2800000), BG_SOFT)
    add_textbox(
        slide, Emu(800000), Emu(4600000), Emu(22800000), Emu(1400000),
        title.upper(), size=46, bold=True, color=TEXT,
    )
    add_rect(slide, Emu(800000), Emu(6200000), Emu(12000000), Emu(20000), ACCENT_BLUE)
    add_textbox(
        slide, Emu(800000), Emu(6500000), Emu(22800000), Emu(600000),
        point, size=26, bold=False, color=ACCENT_BLUE,
    )
    add_textbox(
        slide, Emu(800000), Emu(7200000), Emu(22800000), Emu(500000),
        COURSE, size=20, bold=False, color=TEXT,
    )
    return slide


def toc_slide(prs, items: list[str]):
    slide = blank(prs)
    y0 = title_bar(slide, "Содержание")
    y = y0 + Emu(200000)
    for i, item in enumerate(items, 1):
        add_round(slide, Emu(800000), y, Emu(900000), Emu(900000), ACCENT_BLUE)
        add_textbox(
            slide, Emu(800000), y, Emu(900000), Emu(900000),
            str(i), size=32, bold=True, color=WHITE,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        )
        add_rect(slide, Emu(1900000), y, Emu(20500000), Emu(900000), WHITE, line=LINE)
        add_textbox(
            slide, Emu(2200000), y, Emu(19500000), Emu(900000),
            item, size=32, bold=False, color=TEXT, anchor=MSO_ANCHOR.MIDDLE,
        )
        y += Emu(1200000)
    return slide


def content_slide(prs, title: str, items: list[str], *, subtitle: str | None = None,
                  alert_idx: set[int] | None = None, highlight: str | None = None):
    slide = blank(prs)
    y0 = title_bar(slide, title, subtitle)
    h = Emu(8500000) if not highlight else Emu(7000000)
    bullets(slide, Emu(800000), y0 + Emu(200000), Emu(22800000), h, items,
            size=32, alert_idx=alert_idx)
    if highlight:
        add_round(slide, Emu(800000), Emu(11000000), Emu(22800000), Emu(1800000), BG_SOFT)
        add_rect(slide, Emu(800000), Emu(11000000), Emu(120000), Emu(1800000), ACCENT_RED)
        add_textbox(
            slide, Emu(1200000), Emu(11300000), Emu(21800000), Emu(1400000),
            highlight, size=26, bold=True, color=ACCENT_RED, anchor=MSO_ANCHOR.MIDDLE,
        )
    return slide


def summary_slide(prs, points: list[str]):
    slide = blank(prs)
    y0 = title_bar(slide, "Краткое резюме")
    y = y0 + Emu(200000)
    for pt in points:
        add_rect(slide, Emu(800000), y, Emu(22800000), Emu(1600000), WHITE, line=LINE)
        add_rect(slide, Emu(800000), y, Emu(120000), Emu(1600000), ACCENT_BLUE)
        add_textbox(
            slide, Emu(1200000), y + Emu(300000), Emu(21800000), Emu(1100000),
            pt, size=26, bold=False, color=TEXT, anchor=MSO_ANCHOR.MIDDLE,
        )
        y += Emu(1900000)
    return slide


def thanks_slide(prs):
    slide = blank(prs)
    chrome(slide)
    add_textbox(
        slide, Emu(800000), Emu(5500000), Emu(22800000), Emu(1200000),
        "Спасибо за внимание", size=46, bold=True, color=TEXT,
    )
    add_rect(slide, Emu(800000), Emu(6900000), Emu(10000000), Emu(20000), ACCENT_BLUE)
    add_textbox(
        slide, Emu(800000), Emu(7200000), Emu(22800000), Emu(500000),
        COURSE, size=20, bold=False, color=TEXT,
    )
    return slide


def section_slide(prs, title: str):
    slide = blank(prs)
    chrome(slide)
    add_textbox(
        slide, Emu(800000), Emu(5500000), Emu(22800000), Emu(1200000),
        title, size=46, bold=True, color=TEXT,
    )
    add_rect(slide, Emu(800000), Emu(6900000), Emu(12000000), Emu(20000), ACCENT_BLUE)
    return slide


# ───────────────────── ПРЕЗЕНТАЦИЯ 1 ─────────────────────

def build_pres1():
    prs = new_prs()
    title_slide(
        prs,
        "Кровотечение и обзорный осмотр пострадавшего",
        "Теоретический модуль · п. 2.1",
    )
    toc_slide(prs, [
        "Определение кровотечения и признаки кровопотери",
        "Цель и порядок обзорного осмотра",
        "Обзорный и подробный осмотр: различия",
        "Когда обзорный осмотр критически важен",
    ])

    content_slide(prs, "Понятие «кровотечение»", [
        "Кровотечение — выход крови из сосудистого русла.",
        "Приводит к острой кровопотере и угрозе жизни.",
        "Наружное кровотечение видно снаружи (рана, истечение крови).",
        "Требует немедленной временной остановки.",
    ])

    # Признаки с иконками-кругами
    slide = blank(prs)
    y0 = title_bar(slide, "Основные признаки острой кровопотери")
    signs = [
        ("1", "Слабость, жажда, головокружение"),
        ("2", "«Мушки» перед глазами, обморок"),
        ("3", "Бледная, влажная, холодная кожа"),
        ("4", "Учащённое сердцебиение"),
        ("5", "Частое дыхание"),
        ("6", "Возбуждение → апатия"),
    ]
    positions = [
        (Emu(800000), y0 + Emu(300000)),
        (Emu(8500000), y0 + Emu(300000)),
        (Emu(16200000), y0 + Emu(300000)),
        (Emu(800000), y0 + Emu(4200000)),
        (Emu(8500000), y0 + Emu(4200000)),
        (Emu(16200000), y0 + Emu(4200000)),
    ]
    for (num, txt), (x, y) in zip(signs, positions):
        add_rect(slide, x, y, Emu(7000000), Emu(3200000), WHITE, line=LINE)
        add_icon_circle(slide, x + Emu(250000), y + Emu(400000), Emu(900000), num)
        add_textbox(
            slide, x + Emu(250000), y + Emu(1600000), Emu(6500000), Emu(1300000),
            txt, size=24, bold=False, color=TEXT,
        )

    slide_text_image(
        prs, "Что такое обзорный осмотр",
        [
            "Быстрая визуальная оценка с головы до ног.",
            "Цель — выявить наружное кровотечение.",
            "Сразу после оценки обстановки и безопасности.",
            "Длительность: около 1–2 секунд.",
            "При кровотечении — немедленно начать остановку.",
            "По Приказу № 220н — до подробного осмотра.",
        ],
        "overview_exam.png",
        highlight="Приоритет: сначала угрожающее жизни кровотечение.",
    )

    content_slide(prs, "Порядок обзорного осмотра", [
        "Убедиться в собственной безопасности (перчатки / СИЗ).",
        "Быстро осмотреть пострадавшего с головы до ног.",
        "Искать кровь на одежде, лужи крови, пульсирующую струю.",
        "При нескольких пострадавших — осмотреть всех.",
        "Приоритет — детям и пострадавшим с массивным кровотечением.",
        "После остановки крови — оценка сознания / дыхания и подробный осмотр.",
    ])

    # Сравнительная схема двух колонок
    slide = blank(prs)
    y0 = title_bar(slide, "Обзорный и подробный осмотр")
    add_rect(slide, Emu(800000), y0 + Emu(200000), Emu(10800000), Emu(9500000), WHITE, line=LINE)
    add_rect(slide, Emu(12800000), y0 + Emu(200000), Emu(10800000), Emu(9500000), WHITE, line=LINE)
    add_rect(slide, Emu(800000), y0 + Emu(200000), Emu(10800000), Emu(900000), ACCENT_BLUE)
    add_rect(slide, Emu(12800000), y0 + Emu(200000), Emu(10800000), Emu(900000), BG_BAR)
    add_textbox(
        slide, Emu(1000000), y0 + Emu(350000), Emu(10000000), Emu(600000),
        "Обзорный осмотр", size=26, bold=True, color=WHITE,
    )
    add_textbox(
        slide, Emu(13000000), y0 + Emu(350000), Emu(10000000), Emu(600000),
        "Подробный осмотр", size=26, bold=True, color=TEXT,
    )
    left_items = [
        "Цель: найти наружное кровотечение",
        "Время: 1–2 секунды",
        "Объём: быстрый взгляд с головы до ног",
        "Когда: сразу после оценки обстановки",
        "Результат: останавливать кровь или нет",
    ]
    right_items = [
        "Цель: выявить травмы и другие угрозы",
        "Время: несколько минут",
        "Объём: голова → шея → грудь → … → руки",
        "Когда: после остановки кровотечения",
        "Результат: повязки, иммобилизация",
    ]
    bullets(slide, Emu(1000000), y0 + Emu(1400000), Emu(10000000), Emu(8000000), left_items, size=24)
    bullets(slide, Emu(13000000), y0 + Emu(1400000), Emu(10000000), Emu(8000000), right_items, size=24)

    content_slide(prs, "Когда обзорный осмотр критически важен", [
        "ДТП — раны под одеждой, кровотечение из конечностей.",
        "Падение с высоты — множественные повреждения.",
        "Производственные травмы — риск артериального кровотечения.",
        "ЧС с несколькими пострадавшими — быстрая сортировка.",
        "Огнестрельные и колото-резаные ранения.",
        "Правило: видна кровь или одежда пропитана — действуйте сразу.",
    ])

    summary_slide(prs, [
        "Кровотечение — выход крови из сосудов с риском острой кровопотери.",
        "Обзорный осмотр — быстрый поиск кровотечения с головы до ног (1–2 с).",
        "Сначала останавливаем кровь, затем проводим подробный осмотр.",
        "Особенно важен при ДТП, падении с высоты и множественных пострадавших.",
    ])
    thanks_slide(prs)

    path = OUT / "Кровотечение_и_обзорный_осмотр.pptx"
    prs.save(path)
    return path


# ───────────────────── ПРЕЗЕНТАЦИЯ 2 ─────────────────────

def build_pres2():
    prs = new_prs()
    title_slide(
        prs,
        "Признаки наружного кровотечения и кровопотери",
        "Теоретический модуль · п. 2.2",
    )
    toc_slide(prs, [
        "Виды наружного кровотечения",
        "Сравнительная характеристика",
        "Признаки острой кровопотери и объём",
        "Скрытое (внутреннее) кровотечение",
    ])

    # Три вида — иллюстрации из исходника + редактируемые подписи
    slide_image_grid(prs, "Виды наружного кровотечения", [
        ("bleed_arterial.png", "Артериальное\nАлая кровь пульсирующей струёй\nОпасность: критическая"),
        ("bleed_venous.png", "Венозное\nТёмная кровь равномерной струёй\nОпасность: высокая"),
        ("bleed_capillary.png", "Капиллярное\nСочится по поверхности\nОпасность: обычно низкая"),
    ], cols=3, caption_size=20)

    slide = blank(prs)
    y0 = title_bar(slide, "Сравнительная таблица видов кровотечения")
    add_table(
        slide, Emu(800000), y0 + Emu(300000), Emu(22800000), Emu(9000000),
        ["Признак", "Артериальное", "Венозное", "Капиллярное"],
        [
            ["Цвет крови", "Алый, яркий", "Тёмный, вишнёвый", "Красный"],
            ["Характер", "Пульсирующей струёй", "Равномерной струёй", "Сочится"],
            ["Скорость", "Очень быстрая", "Умеренная / быстрая", "Медленная"],
            ["Опасность", "Критическая — минуты", "Высокая при крупных венах", "Обычно низкая"],
            ["Приоритет", "Немедленно", "Быстро", "По ситуации"],
        ],
        font_size=22,
    )

    content_slide(prs, "Признаки острой кровопотери", [
        "Резкая общая слабость, жажда, головокружение.",
        "Мелькание «мушек» перед глазами; обморок при попытке встать.",
        "Бледная, влажная, холодная кожа.",
        "Учащённое сердцебиение и частое дыхание.",
        "Тревога, возбуждение → затем апатия.",
        "Слабый / нитевидный пульс при тяжёлой кровопотере.",
    ])

    slide = blank(prs)
    y0 = title_bar(slide, "Оценка объёма кровопотери")
    add_table(
        slide, Emu(800000), y0 + Emu(300000), Emu(22800000), Emu(7500000),
        ["Объём", "Доля ОЦК", "Состояние пострадавшего"],
        [
            ["До ~500 мл", "≈ 10%", "Слабость, жажда"],
            ["500–1000 мл", "≈ 10–20%", "Бледность, тахикардия"],
            ["1000–1500 мл", "≈ 20–30%", "Головокружение, холодный пот, обмороки"],
            ["Более 1500–2000 мл", "> 30%", "Угроза жизни, шок, потеря сознания"],
        ],
        font_size=22,
    )
    add_textbox(
        slide, Emu(800000), Emu(11500000), Emu(22800000), Emu(1200000),
        "Ориентиры условные; у детей и пожилых критический порог ниже.",
        size=24, bold=True, color=ACCENT_RED,
    )

    content_slide(
        prs, "Признаки скрытого (внутреннего) кровотечения",
        [
            "Кровь изливается в полости тела — снаружи может не быть видно.",
            "Бледность, холодный пот, нарастающая слабость, жажда.",
            "Учащённый слабый пульс, частое дыхание.",
            "Боль и напряжение живота (кровотечение в брюшную полость).",
            "При травме груди — одышка, слабость, бледность.",
            "Действия: вызвать СМП, положение, холод на живот, не кормить и не поить.",
        ],
        highlight="Внутреннее кровотечение на месте не «останавливают» — нужна СМП.",
    )

    summary_slide(prs, [
        "Артериальное — алая пульсирующая струя; венозное — тёмная струя; капиллярное — сочение.",
        "Признаки кровопотери: слабость, жажда, бледность, холодный пот, тахикардия.",
        "Потеря >30% ОЦК угрожает жизни; у детей порог ниже.",
        "При подозрении на внутреннее кровотечение — срочный вызов СМП.",
    ])
    thanks_slide(prs)

    path = OUT / "Признаки_наружного_кровотечения.pptx"
    prs.save(path)
    return path


# ───────────────────── ПРЕЗЕНТАЦИЯ 3 ─────────────────────

def build_pres3():
    prs = new_prs()
    title_slide(
        prs,
        "Способы временной остановки наружного кровотечения",
        "Теоретический модуль · п. 2.3",
    )
    toc_slide(prs, [
        "Прямое давление и давящая повязка",
        "Пальцевое прижатие и максимальное сгибание",
        "Жгут и подручные средства",
        "Алгоритм, записка, ошибки и ограничения",
    ])

    section_slide(prs, "Прямое давление и давящая повязка")

    slide_image_grid(prs, "Прямое давление и давящая повязка", [
        ("direct_pressure.png", "Прямое давление на рану\nОсновной способ по Приказу № 220н.\nНадавите салфеткой и удерживайте."),
        ("pressure_bandage.png", "Давящая повязка\nПосле остановки крови давлением.\nПри промокании — усилить, не снимая."),
    ], cols=2, caption_size=22)

    section_slide(prs, "Пальцевое прижатие артерий")

    slide_image_grid(prs, "Сонная артерия", [
        ("carotid_point.png", "Точка: шея снаружи от гортани,\nна стороне повреждения, к позвоночнику."),
        ("carotid_4fingers.png", "Давление четырьмя пальцами\nодновременно к позвоночнику."),
        ("carotid_thumb.png", "Вариант: большим пальцем\nв ту же точку к позвоночнику."),
    ], cols=3, caption_size=18)

    slide_image_grid(prs, "Подключичная артерия", [
        ("subclavian_1.png", "Прижимается в ямке над ключицей\nк I ребру."),
        ("subclavian_2.png", "Давление четырьмя\nвыпрямленными пальцами."),
        ("subclavian_3.png", "Вариант: давление\nсогнутыми пальцами."),
    ], cols=3, caption_size=18)

    slide_image_grid(prs, "Плечевая артерия", [
        ("brachial_1.png", "К плечевой кости с внутренней стороны\nмежду бицепсом и трицепсом, средняя треть."),
        ("brachial_2.png", "Четыре пальца кисти,\nобхватывающей плечо сверху или снизу."),
    ], cols=2, caption_size=20)

    slide_image_grid(prs, "Подмышечная и бедренная артерии", [
        ("axillary_1.png", "Подмышечная: в подмышечной впадине\nк плечевой кости, жёсткими пальцами."),
        ("axillary_2.png", "При кровотечении из раны плеча\nниже плечевого сустава."),
        ("femoral_1.png", "Бедренная: ниже паховой складки\nпри кровотечении из ран бедра."),
        ("femoral_2.png", "Давление кулаком,\nзафиксированным второй рукой, весом тела."),
    ], cols=2, caption_size=18)

    # Сводка таблицей
    slide = blank(prs)
    y0 = title_bar(slide, "Точки пальцевого прижатия — сводка")
    add_table(
        slide, Emu(800000), y0 + Emu(300000), Emu(22800000), Emu(9500000),
        ["Артерия", "Место прижатия", "Как давить"],
        [
            ["Сонная", "Шея снаружи от гортани, к позвоночнику", "4 пальца или большой палец"],
            ["Подключичная", "Ямка над ключицей к I ребру", "4 пальца / согнутые пальцы"],
            ["Подмышечная", "Подмышечная впадина к плечевой кости", "Прямые жёсткие пальцы"],
            ["Плечевая", "Внутренняя сторона плеча (бицепс–трицепс)", "4 пальца, обхватывая плечо"],
            ["Бедренная", "Ниже паховой складки", "Кулак, вес тела"],
        ],
        font_size=20,
    )

    slide_image_grid(prs, "Максимальное сгибание конечности", [
        ("flexion_arm.png", "Предплечье: валик в локтевой сгиб,\nсогнуть, фиксировать к плечу."),
        ("flexion_leg.png", "Голень/стопа: валик в подколенную ямку,\nсогнуть в колене, фиксировать."),
        ("flexion_thigh.png", "Бедро: валик в пах, колено к груди,\nфиксировать руками или бинтом."),
    ], cols=3, caption_size=18)

    section_slide(prs, "Кровоостанавливающий жгут")

    slide_text_image(
        prs, "Показания к наложению жгута",
        [
            "Обширное повреждение конечности.",
            "Отрыв (травматическая ампутация).",
            "Давление и давящая повязка неэффективны.",
            "Накладывать на плечо или бедро.",
            "Обязательно на одежду или подкладку.",
            "Записка со временем обязательна.",
        ],
        "tourniquet_2.png",
        highlight="t max ≤ 60 мин в тёплое время · ≤ 30 мин в холодное!",
    )

    slide_image_grid(prs, "Наложение кровоостанавливающего жгута", [
        ("tourniquet_1.png", "1. Наложить выше раны\nна одежду / подкладку."),
        ("tourniquet_2.png", "2. Первый тур — максимальное\nнатяжение; кровь остановилась."),
        ("tourniquet_3.png", "3. Записка со временем;\nжгут должен быть виден."),
    ], cols=3, caption_size=18)

    slide_text_image(
        prs, "Подручные средства в качестве жгута",
        [
            "Тесьма, платок, галстук, ремень.",
            "Используйте закрутку для натяжения.",
            "Не использовать проволоку, леску, узкий шнур.",
            "Критерий: кровотечение остановилось.",
            "Обязательна записка со временем.",
        ],
        "improvised_tq.png",
    )

    content_slide(prs, "Записка под жгут", [
        "Укажите точное время наложения: ЧЧ:ММ (и дату при необходимости).",
        "Фамилия / кто наложил — по возможности.",
        "Кратко — место происшествия.",
        "Вложите записку под жгут или прикрепите к одежде на видном месте.",
        "Сообщите время бригаде скорой помощи при передаче.",
        "Не полагайтесь только на память.",
    ], highlight="Без записки легко превысить допустимое время жгута!")

    content_slide(prs, "Ошибки при наложении жгута", [
        "Слабое натяжение — венозный застой, усиление кровотечения.",
        "Слишком далеко от раны — излишняя ишемия.",
        "На голую кожу без подкладки — повреждение кожи и нервов.",
        "Скрытие жгута одеждой — бригада СМП может не заметить.",
        "Нет записки со временем — риск превышения срока.",
        "Жгут «на всякий случай» при капиллярном кровотечении.",
    ])

    content_slide(prs, "Когда НЕЛЬЗЯ накладывать жгут", [
        "Кровотечение останавливается давлением или давящей повязкой.",
        "Капиллярное и умеренное венозное кровотечение.",
        "Раны шеи, груди, живота, головы — жгут туда не накладывают.",
        "Нет конечности проксимальнее раны для размещения жгута.",
        "Инородное тело без массивного артериального кровотечения — фиксация + повязка.",
    ], alert_idx={2})

    content_slide(prs, "Подручные средства и гемостатики", [
        "Импровизированный жгут: тесьма, платок, галстук, ремень + закрутка.",
        "Не использовать проволоку, леску, узкий шнур.",
        "Гемостатические салфетки / бинты из аптечки первой помощи.",
        "Плотно вложить в рану / на рану, сильно придавить, затем повязка.",
        "Применяют там, где жгут невозможен, или как усиление давления.",
        "После остановки — контроль, согревание, ожидание СМП.",
    ])

    summary_slide(prs, [
        "Приоритет: прямое давление → давящая повязка → жгут по показаниям.",
        "Пальцевое прижатие и сгибание — временные приёмы.",
        "Жгут: на одежду, выше раны, с запиской; ≤60 мин / ≤30 мин.",
        "Гемостатические средства усиливают давление там, где это уместно.",
    ])
    thanks_slide(prs)

    path = OUT / "Способы_временной_остановки_кровотечения.pptx"
    prs.save(path)
    return path


# ───────────────────── ПРЕЗЕНТАЦИЯ 4 ─────────────────────

def build_pres4():
    prs = new_prs()
    title_slide(
        prs,
        "Последовательность мероприятий по остановке кровотечения",
        "Теоретический модуль · п. 2.4",
    )
    toc_slide(prs, [
        "Полный алгоритм действий",
        "Приоритетность способов остановки",
        "Профилактика травматического шока",
        "Подробный осмотр и действия после остановки",
    ])

    # Чек-лист алгоритма схемой
    slide = blank(prs)
    y0 = title_bar(slide, "Полный алгоритм (чек-лист)", "Приказ Минздрава РФ № 220н")
    steps = [
        ("1", "Оценка обстановки → безопасность (СИЗ)"),
        ("2", "Обзорный осмотр — поиск кровотечения"),
        ("3", "Временная остановка кровотечения"),
        ("4", "Признаки жизни: сознание → дыхание"),
        ("5", "СЛР при отсутствии признаков жизни"),
        ("6", "Вызов скорой медицинской помощи"),
        ("7", "Подробный осмотр и помощь при травмах"),
        ("8", "Положение, тепло, контроль, передача СМП"),
    ]
    # 2 columns of steps
    for i, (num, txt) in enumerate(steps):
        col = i % 2
        row = i // 2
        x = Emu(800000) + Emu(col * 11800000)
        y = y0 + Emu(200000) + Emu(row * 2000000)
        accent = num in ("2", "3")
        add_flow_step(slide, x, y, Emu(11000000), Emu(1600000), num, txt, accent=accent)

    content_slide(prs, "Оценка состояния пострадавшего", [
        "При массивном кровотечении сначала остановите кровь.",
        "Сознание: громко обратитесь, осторожно потрясите за плечи.",
        "Дыхание: смотрю–слушаю–ощущаю не более 10 секунд.",
        "При нескольких травмах — сначала то, что угрожает жизни сейчас.",
        "Не тратьте время на второстепенные раны при угрожающем кровотечении.",
    ])

    # Приоритетность — схема стрелками
    slide = blank(prs)
    y0 = title_bar(slide, "Приоритетность способов остановки")
    labels = [
        ("1", "Прямое\nдавление"),
        ("2", "Давящая\nповязка"),
        ("3", "Жгут\nпо показаниям"),
    ]
    x = Emu(1200000)
    for i, (num, lab) in enumerate(labels):
        add_round(slide, x, y0 + Emu(2500000), Emu(5000000), Emu(3500000), ACCENT_BLUE if i == 0 else BG_BAR)
        add_textbox(
            slide, x, y0 + Emu(2700000), Emu(5000000), Emu(800000),
            num, size=32, bold=True, color=WHITE if i == 0 else TEXT,
            align=PP_ALIGN.CENTER,
        )
        add_textbox(
            slide, x + Emu(200000), y0 + Emu(3600000), Emu(4600000), Emu(2000000),
            lab, size=26, bold=True, color=WHITE if i == 0 else TEXT,
            align=PP_ALIGN.CENTER,
        )
        if i < 2:
            add_arrow_right(slide, x + Emu(5200000), y0 + Emu(4000000))
        x += Emu(7000000)
    add_textbox(
        slide, Emu(800000), Emu(11000000), Emu(22800000), Emu(1500000),
        "Дополнительно (по учебным программам): пальцевое прижатие, максимальное сгибание.",
        size=24, bold=False, color=TEXT,
    )

    content_slide(prs, "Приоритет при множественных травмах", [
        "Сначала — угрожающее жизни кровотечение.",
        "Затем — проходимость дыхательных путей и дыхание / СЛР.",
        "Затем — остальные кровотечения и раны.",
        "Затем — иммобилизация, повязки, оптимальное положение.",
        "При нескольких пострадавших — приоритет детям и массивным кровотечениям.",
    ])

    content_slide(prs, "Травматический шок: понятие и признаки", [
        "Причины: тяжёлые травмы и сильные кровотечения.",
        "Наличие тяжёлой травмы и сильного кровотечения.",
        "Нарушения дыхания и кровообращения (учащение).",
        "Бледная холодная влажная кожа.",
        "Возбуждение, сменяющееся апатией.",
    ])

    slide_text_image(
        prs, "Профилактика травматического шока",
        [
            "Остановить кровотечение как можно раньше.",
            "Придать оптимальное положение тела.",
            "Иммобилизировать повреждённые конечности.",
            "Защитить от переохлаждения.",
            "Не давать есть и пить при тяжёлой травме.",
            "Покой, поддержка, контроль сознания и дыхания.",
        ],
        "shock_prev.png",
    )

    slide_image_grid(prs, "Подробный осмотр — голова, шея, грудь", [
        ("exam_head.png", "Шаг 1. Осмотр головы"),
        ("exam_neck.png", "Шаг 2. Осмотр шеи"),
        ("exam_chest.png", "Шаг 3. Осмотр груди"),
    ], cols=3, caption_size=22)

    slide_image_grid(prs, "Подробный осмотр — живот, ноги, руки", [
        ("exam_abdomen.png", "Шаг 4. Осмотр живота и таза"),
        ("exam_legs.png", "Шаг 5. Осмотр ног"),
        ("exam_arms.png", "Шаг 6. Осмотр рук"),
    ], cols=3, caption_size=22)

    content_slide(prs, "Порядок подробного осмотра", [
        "Голова → шея → грудь → живот и таз → ноги → руки.",
        "Спина — при возможности безопасно повернуть / осмотреть.",
        "При сознании — короткий опрос о жалобах и механизме травмы.",
        "Ищите раны, деформации, болезненность, признаки внутреннего кровотечения.",
        "Устраняйте угрозы по мере выявления (повязка, фиксация, холод).",
    ])

    content_slide(prs, "Действия после остановки кровотечения", [
        "Контролируйте повязку: при промокании усильте, не снимая нижние слои.",
        "Проверяйте признаки жизни: сознание, дыхание, цвет кожи.",
        "Следите за жгутом: время, видимость, записка.",
        "Согрейте пострадавшего, обеспечьте покой.",
        "При ухудшении — повторный вызов СМП с динамикой.",
        "Передайте бригаде: что случилось, что сделано, время жгута.",
    ])

    content_slide(
        prs, "Когда и как ослаблять жгут",
        [
            "Планово ослаблять жгут на месте не рекомендуется.",
            "Если время превысило 60 мин / 30 мин и СМП задерживается:",
            "подготовьте прямое давление / давящую повязку;",
            "медленно ослабьте жгут, оценивая кровотечение;",
            "если кровь снова бьёт струёй — немедленно затяните и обновите время;",
            "решение о снятии, как правило, принимает медработник.",
        ],
        alert_idx={1},
        highlight="Лимит жгута: ≤ 60 мин (тепло) · ≤ 30 мин (холод)!",
    )

    summary_slide(prs, [
        "Алгоритм: безопасность → обзорный осмотр → остановка крови → признаки жизни → СМП.",
        "Способы: от простого к сложному — давление → повязка → жгут.",
        "Профилактика шока: кровь, положение, иммобилизация, тепло.",
        "После остановки — контроль повязки / жгута и состояния до передачи СМП.",
    ])
    thanks_slide(prs)

    path = OUT / "Последовательность_остановки_кровотечения.pptx"
    prs.save(path)
    return path


# ───────────────────── ПРЕЗЕНТАЦИЯ 5 ─────────────────────

def build_pres5():
    prs = new_prs()
    title_slide(
        prs,
        "Остановка кровотечения при ранениях различных областей тела",
        "Теоретический модуль · п. 2.5",
    )
    toc_slide(prs, [
        "Голова, глаза, нос, носовое кровотечение",
        "Шея, грудь, инородный предмет",
        "Живот, таз, конечности",
        "Позвоночник, отрыв конечности, перемещение",
    ])

    section_slide(prs, "Голова · глаза · нос")

    slide_image_grid(prs, "Травма головы", [
        ("head_trauma_1.png", "Боковое положение\nпри потере сознания."),
        ("head_trauma_2.png", "Прямое давление на рану\n(не при открытой ЧМТ)."),
        ("head_trauma_3.png", "Давящая повязка\nна рану головы."),
    ], cols=3, caption_size=20)

    slide_image_grid(prs, "Травмы глаза и носа", [
        ("eye_injury.png", "Глаз: стерильная повязка на оба глаза.\nИнородные тела не извлекать."),
        ("nose_injury.png", "Нос: голова вперёд, зажать нос,\nхолод на переносицу."),
    ], cols=2, caption_size=22)

    slide_image_grid(prs, "Носовое кровотечение — шаги", [
        ("nose_1.png", "Шаг 1. Усадить,\nголова слегка вперёд."),
        ("nose_2.png", "Шаг 2. Зажать крылья носа\nна 10–15 минут."),
        ("nose_3.png", "Шаг 3. Холод\nна переносицу."),
        ("nose_4.png", "Шаг 4. Не высмаркиваться.\nПри обильном — СМП."),
    ], cols=4, caption_size=16)

    section_slide(prs, "Шея и грудь")

    slide_image_grid(prs, "Травма шеи", [
        ("neck_1.png", "Шаг 1. Остановка кровотечения\nпрямым давлением."),
        ("neck_2.png", "Шаг 2. Фиксация шеи\nпри перемещении."),
        ("neck_3.png", "Шаг 3. Фиксация\nшейного отдела."),
    ], cols=3, caption_size=18)

    content_slide(
        prs, "Травма шеи: артерия или вена",
        [
            "Артериальное (сонная): алая кровь, пульсирующая струя.",
            "Венозное (яремные): тёмная обильная струя; риск воздушной эмболии.",
            "Прямое давление на рану (не пережимать обе сонные артерии).",
            "Давящая повязка через плечо — не тугая круговая на шее.",
            "Срочный вызов СМП.",
        ],
        highlight="На шею жгут НЕ накладывают. Не сдавливайте дыхательные пути.",
    )

    slide_image_grid(prs, "Травма груди", [
        ("chest_1.png", "Шаг 1. Полусидячее\nположение."),
        ("chest_2.png", "Шаг 2. Окклюзионная\n(герметизирующая) повязка."),
        ("chest_3.png", "Шаг 3. Инородный предмет:\nобложить, повязка поверх."),
    ], cols=3, caption_size=18)

    # Схема окклюзионной повязки (фигуры)
    slide = blank(prs)
    y0 = title_bar(slide, "Окклюзионная повязка — схема")
    steps = [
        ("1", "Полусидячее положение"),
        ("2", "Воздухонепроницаемый материал на рану"),
        ("3", "Фиксация пластырем (клапан / герметично)"),
        ("4", "При сквозном — закрыть вход и выход"),
        ("5", "Контроль дыхания"),
    ]
    y = y0 + Emu(300000)
    for num, txt in steps:
        add_flow_step(slide, Emu(800000), y, Emu(22800000), Emu(1400000), num, txt, accent=(num == "2"))
        y += Emu(1650000)

    content_slide(
        prs, "Инородный предмет в ране",
        [
            "НЕ ИЗВЛЕКАТЬ — предмет может тампонировать сосуд.",
            "Обложить салфетками / бинтами валиками вокруг.",
            "Давящая повязка поверх валиков с фиксацией предмета.",
            "Остановить кровотечение вокруг, не продавливая предмет вглубь.",
            "Иммобилизировать область, вызвать СМП.",
        ],
        highlight="Правило: зафиксировать — не извлекать!",
    )

    section_slide(prs, "Живот · таз · конечности")

    slide_text_image(
        prs, "Закрытая травма живота и таза",
        [
            "Вызвать скорую помощь.",
            "Положить холод на живот.",
            "На спине, валик под полусогнутыми разведёнными ногами.",
            "Не кормить и не поить.",
            "Контроль признаков внутреннего кровотечения и шока.",
        ],
        "abdomen_closed.png",
    )

    slide_text_image(
        prs, "Открытая травма живота",
        [
            "Выпавшие органы НЕ вправлять.",
            "Накрыть влажной тканью, обложить валиками.",
            "Повязка без давления на органы.",
            "Холод — рядом, без прямого давления на органы.",
            "Срочный вызов СМП.",
        ],
        "abdomen_open.png",
        highlight="Органы брюшной полости не вправлять!",
    )

    slide_image_grid(prs, "Травма конечностей", [
        ("limb_1.png", "Шаг 1. Остановить\nкровотечение."),
        ("limb_2.png", "Шаг 2. Иммобилизация\nдвух соседних суставов."),
        ("limb_3.png", "Шаг 3. Не класть шину\nна выступающие отломки."),
    ], cols=3, caption_size=18)

    slide_text_image(
        prs, "Иммобилизация",
        [
            "Создание неподвижности повреждённой части тела.",
            "Шины или подручные средства поверх одежды.",
            "Обездвижить два соседних сустава.",
            "Костные отломки не вправлять.",
            "Сначала кровь — затем иммобилизация.",
        ],
        "immobilization.png",
    )

    content_slide(
        prs, "Травматическая ампутация",
        [
            "Немедленно жгут выше места отрыва (плечо / бедро).",
            "Давящая повязка на культю.",
            "Записка со временем наложения жгута.",
            "Сегмент: ткань → пакет → пакет со холодом (не мочить).",
            "Передать сегмент вместе с пострадавшим бригаде СМП.",
        ],
        alert_idx={0},
        highlight="Жгут при отрыве показан. t max ≤ 60 мин / 30 мин.",
    )

    section_slide(prs, "Позвоночник")

    slide_image_grid(prs, "Травма позвоночника", [
        ("spine_1.png", "Ровная жёсткая горизонтальная\nповерхность при перемещении."),
        ("spine_2.png", "Несколько человек; постоянная\nфиксация шейного отдела."),
    ], cols=2, caption_size=22)

    content_slide(prs, "Правила перемещения при травме позвоночника", [
        "Перемещать только при угрозе жизни или для эвакуации.",
        "Нужно несколько человек (обычно 3–5).",
        "Один постоянно удерживает голову и шейный отдел.",
        "Перекладывание одним движением, без скручивания.",
        "Фиксация шеи вручную / подручными средствами / воротником.",
    ])

    summary_slide(prs, [
        "Голова / глаза / нос: повязки; при потере сознания — боковое положение.",
        "Шея и грудь: давление без кругового жгута; окклюзионная повязка; предмет не извлекать.",
        "Живот: органы не вправлять; конечности — кровь, затем иммобилизация двух суставов.",
        "Позвоночник — жёсткая поверхность и команда с фиксацией шеи.",
    ])
    thanks_slide(prs)

    path = OUT / "Кровотечение_при_ранениях_областей_тела.pptx"
    prs.save(path)
    return path


def main():
    paths = []
    for fn in (build_pres1, build_pres2, build_pres3, build_pres4, build_pres5):
        p = fn()
        print(f"OK: {p.name} · {p.stat().st_size // 1024} KB")
        paths.append(p)

    # Verify pictures are separate objects (not whole-slide backgrounds)
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    for p in paths:
        prs = Presentation(str(p))
        pics = 0
        texts = 0
        for s in prs.slides:
            for sh in s.shapes:
                if sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    pics += 1
                    # picture must not cover entire slide
                    if sh.width >= SLIDE_W * 0.95 and sh.height >= SLIDE_H * 0.95:
                        raise SystemExit(f"FULL-SLIDE IMAGE in {p.name}")
                if sh.has_text_frame and sh.text_frame.text.strip():
                    texts += 1
        print(f"  {p.name}: {len(prs.slides)} slides, pictures={pics}, textboxes≈{texts}")

    readme = OUT / "README.md"
    readme.write_text(
        "# Тема 2. Оказание первой помощи при наружных кровотечениях\n\n"
        "Пять **редактируемых** презентаций (Open Sans, единый стиль).\n\n"
        "| Файл | Пункт |\n|---|---|\n"
        "| `Кровотечение_и_обзорный_осмотр.pptx` | 2.1 |\n"
        "| `Признаки_наружного_кровотечения.pptx` | 2.2 |\n"
        "| `Способы_временной_остановки_кровотечения.pptx` | 2.3 |\n"
        "| `Последовательность_остановки_кровотечения.pptx` | 2.4 |\n"
        "| `Кровотечение_при_ранениях_областей_тела.pptx` | 2.5 |\n\n"
        "Слайды не вставлены картинками: текст, таблицы и схемы — объекты PowerPoint.\n"
        "Сборка: `build_tema2_presentations.py`.\n",
        encoding="utf-8",
    )
    print("Done.")


if __name__ == "__main__":
    main()
