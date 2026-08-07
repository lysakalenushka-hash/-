#!/usr/bin/env python3
"""
Тема 2 — 5 презентаций.

ПРАВИЛО:
- Слайды из исходного PDF переносятся ЦЕЛИКОМ без изменений
  (точный вид страницы исходника).
- НОВЫЕ слайды — только для информации, которой нет в исходнике
  (Open Sans, 46/26/32 pt, RGB 221/225, без логотипа).
"""

from __future__ import annotations

from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

OUT = Path("/workspace/tema2_bleeding")
PAGES = Path("/tmp/bleed_pages")
OUT.mkdir(parents=True, exist_ok=True)

SLIDE_W = Emu(24384000)
SLIDE_H = Emu(13716000)

WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BG_SOFT = RGBColor(0xF7, 0xF7, 0xF7)
BG_BAR = RGBColor(221, 221, 221)   # фоновые элементы
LINE = RGBColor(225, 225, 225)    # рамки / разделители
TEXT = RGBColor(0x33, 0x33, 0x33)
ACCENT_RED = RGBColor(0xCC, 0x00, 0x00)
ACCENT_BLUE = RGBColor(0x00, 0x55, 0xAA)
TABLE_HDR = RGBColor(0x44, 0x44, 0x44)
ROW_ALT = RGBColor(0xF5, 0xF5, 0xF5)
FONT = "Open Sans"


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
        sh.line.width = Pt(1)
    return sh


def add_round(slide, left, top, width, height, fill, line=None):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(1)
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
        p.space_after = Pt(12)
        run = p.add_run()
        run.text = f"•  {item}"
        set_run_font(run, size=size, bold=(i in alert),
                     color=ACCENT_RED if i in alert else TEXT)
    return box


# ─── ИСХОДНЫЙ СЛАЙД (без изменений) ───

def add_original(prs: Presentation, page_no: int):
    """Перенос слайда из исходного PDF целиком, без правок."""
    png = PAGES / f"page_{page_no:02d}.png"
    if not png.exists():
        raise FileNotFoundError(png)
    slide = blank(prs)
    slide.shapes.add_picture(str(png), 0, 0, width=SLIDE_W, height=SLIDE_H)
    return slide


# ─── НОВЫЕ СЛАЙДЫ (только пробелы) ───

def new_header(slide, title: str):
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, WHITE)
    add_textbox(slide, Emu(800000), Emu(400000), Emu(22800000), Emu(1000000),
                title, size=46, bold=True, color=TEXT)
    add_rect(slide, Emu(800000), Emu(1500000), Emu(22800000), Emu(20000), LINE)


def new_toc(prs, items: list[str]):
    """Оглавление — в исходнике отдельного слайда нет."""
    slide = blank(prs)
    new_header(slide, "Содержание")
    y = Emu(2000000)
    for i, item in enumerate(items[:5], 1):
        add_oval(slide, Emu(800000), y, Emu(800000), Emu(800000), ACCENT_BLUE)
        add_textbox(slide, Emu(800000), y, Emu(800000), Emu(800000),
                    str(i), size=26, bold=True, color=WHITE,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_rect(slide, Emu(1900000), y, Emu(20500000), Emu(800000), WHITE, line=LINE)
        add_textbox(slide, Emu(2200000), y, Emu(19500000), Emu(800000),
                    item, size=32, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)
        y += Emu(1400000)
    return slide


def new_list(prs, title: str, items: list[str], *, highlight: str | None = None,
             alert: set[int] | None = None):
    slide = blank(prs)
    new_header(slide, title)
    h = Emu(8000000) if not highlight else Emu(6500000)
    add_bullets(slide, Emu(800000), Emu(1900000), Emu(22800000), h,
                items[:6], size=32, alert=alert)
    if highlight:
        add_round(slide, Emu(800000), Emu(11000000), Emu(22800000), Emu(1600000), BG_SOFT)
        add_rect(slide, Emu(800000), Emu(11000000), Emu(120000), Emu(1600000), ACCENT_RED)
        add_textbox(slide, Emu(1200000), Emu(11200000), Emu(21800000), Emu(1200000),
                    highlight, size=26, bold=True, color=ACCENT_RED,
                    anchor=MSO_ANCHOR.MIDDLE)
    return slide


def new_scheme(prs, title: str, steps: list[str], *, key_idx: int | None = None):
    slide = blank(prs)
    new_header(slide, title)
    y = Emu(1900000)
    bh = Emu(1300000)
    for i, step in enumerate(steps):
        accent = key_idx is not None and i == key_idx
        fill = ACCENT_BLUE if accent else BG_SOFT
        add_round(slide, Emu(800000), y, Emu(22800000), bh, fill,
                  line=None if accent else LINE)
        add_oval(slide, Emu(1100000), y + Emu(250000), Emu(800000), Emu(800000),
                 WHITE if accent else ACCENT_BLUE)
        add_textbox(slide, Emu(1100000), y + Emu(250000), Emu(800000), Emu(800000),
                    str(i + 1), size=26, bold=True,
                    color=ACCENT_BLUE if accent else WHITE,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(slide, Emu(2200000), y + Emu(250000), Emu(20500000), Emu(800000),
                    step, size=28, color=WHITE if accent else TEXT,
                    anchor=MSO_ANCHOR.MIDDLE)
        y += bh + Emu(200000)
    return slide


def new_table(prs, title: str, headers: list[str], rows: list[list[str]], *, fs=22):
    slide = blank(prs)
    new_header(slide, title)
    n_rows, n_cols = len(rows) + 1, len(headers)
    tw, th = Emu(22800000), min(Emu(8000000), Emu(1200000) * n_rows)
    table = slide.shapes.add_table(n_rows, n_cols, Emu(800000), Emu(1900000), tw, th).table
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
                    set_run_font(r, size=fs, bold=alert,
                                 color=ACCENT_RED if alert else TEXT)
    return slide


def new_compare(prs, title: str, left_title: str, left_items: list[str],
                right_title: str, right_items: list[str]):
    slide = blank(prs)
    new_header(slide, title)
    add_rect(slide, Emu(800000), Emu(1900000), Emu(10800000), Emu(9500000), WHITE, line=LINE)
    add_rect(slide, Emu(12800000), Emu(1900000), Emu(10800000), Emu(9500000), WHITE, line=LINE)
    add_rect(slide, Emu(800000), Emu(1900000), Emu(10800000), Emu(1000000), ACCENT_BLUE)
    add_rect(slide, Emu(12800000), Emu(1900000), Emu(10800000), Emu(1000000), BG_BAR)
    add_textbox(slide, Emu(1000000), Emu(2050000), Emu(10400000), Emu(700000),
                left_title, size=26, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_textbox(slide, Emu(13000000), Emu(2050000), Emu(10400000), Emu(700000),
                right_title, size=26, bold=True, color=TEXT, align=PP_ALIGN.CENTER)
    add_bullets(slide, Emu(1100000), Emu(3200000), Emu(10000000), Emu(7800000),
                left_items, size=26)
    add_bullets(slide, Emu(13100000), Emu(3200000), Emu(10000000), Emu(7800000),
                right_items, size=26)
    return slide


def new_summary(prs, points: list[str], callout: str | None = None):
    """Резюме — в исходнике нет отдельного слайда итогов по теме."""
    slide = blank(prs)
    new_header(slide, "Главное запомнить")
    y = Emu(2000000)
    for pt in points[:4]:
        add_rect(slide, Emu(800000), y, Emu(22800000), Emu(1600000), WHITE, line=LINE)
        add_rect(slide, Emu(800000), y, Emu(120000), Emu(1600000), ACCENT_BLUE)
        add_textbox(slide, Emu(1300000), y + Emu(350000), Emu(21500000), Emu(900000),
                    pt, size=28, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)
        y += Emu(1900000)
    if callout:
        add_textbox(slide, Emu(800000), Emu(11500000), Emu(22800000), Emu(1000000),
                    callout, size=26, bold=True, color=ACCENT_RED, align=PP_ALIGN.CENTER)
    return slide


def new_divider(prs, title: str):
    """Разделитель раздела — в исходнике нет; короткий новый слайд."""
    slide = blank(prs)
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, WHITE)
    add_rect(slide, 0, Emu(5000000), SLIDE_W, Emu(2500000), BG_BAR)
    add_textbox(slide, Emu(800000), Emu(5500000), Emu(22800000), Emu(1400000),
                title, size=46, bold=True, color=TEXT, align=PP_ALIGN.CENTER)
    return slide


# ═══════════════════ ПРЕЗЕНТАЦИЯ 1 ═══════════════════
# Исходник: 1 (титул), 2 (понятие), 4 (обзорный), 28 (спасибо)
# Новое: оглавление, определение обзорного, различия, критические ситуации, резюме

def build_pres1():
    prs = new_prs()
    add_original(prs, 1)   # титул исходника — без изменений
    new_toc(prs, [
        "Определение кровотечения и признаки кровопотери",
        "Цель и порядок обзорного осмотра",
        "Обзорный и подробный осмотр: различия",
        "Когда обзорный осмотр критически важен",
    ])
    add_original(prs, 2)   # понятие «кровотечение» + признаки
    add_original(prs, 4)   # обзорный осмотр

    # --- НОВОЕ: в исходнике нет отдельного определения цели/критичности ---
    new_list(prs, "Что такое обзорный осмотр и зачем он нужен", [
        "Быстрая визуальная оценка пострадавшего с головы до ног",
        "Цель — выявить продолжающееся наружное кровотечение",
        "Выполняется сразу после оценки обстановки и безопасности",
        "При кровотечении — немедленно начать временную остановку",
        "По Приказу Минздрава № 220н — до подробного осмотра",
    ], highlight="Приоритет: сначала угрожающее жизни кровотечение")

    new_compare(
        prs, "Обзорный и подробный осмотр — различия",
        "Обзорный осмотр",
        [
            "Цель: найти наружное кровотечение",
            "Время: около 1–2 секунд",
            "Объём: быстрый взгляд с головы до ног",
            "Когда: сразу после оценки обстановки",
        ],
        "Подробный осмотр",
        [
            "Цель: выявить травмы и другие угрозы",
            "Время: несколько минут",
            "Объём: голова → шея → грудь → … → руки",
            "Когда: после остановки кровотечения",
        ],
    )

    new_list(prs, "Когда обзорный осмотр критически важен", [
        "ДТП — раны под одеждой, кровотечение из конечностей",
        "Падение с высоты — множественные повреждения",
        "Производственные травмы — риск артериального кровотечения",
        "ЧС с несколькими пострадавшими — быстрая сортировка",
        "Огнестрельные и колото-резаные ранения",
    ])

    new_summary(prs, [
        "Кровотечение — выход крови из сосудов с риском кровопотери",
        "Обзорный осмотр — быстрый поиск кровотечения (1–2 с)",
        "Сначала останавливаем кровь, затем подробный осмотр",
        "Особенно важен при ДТП и падении с высоты",
    ], callout="Время — критический фактор!")

    add_original(prs, 28)  # «Благодарим за внимание» исходника
    path = OUT / "Кровотечение_и_обзорный_осмотр.pptx"
    prs.save(path)
    return path


# ═══════════════════ ПРЕЗЕНТАЦИЯ 2 ═══════════════════
# Исходник: 1, 2, 3, 28
# Новое: оглавление, сравнительная таблица, объём кровопотери, внутреннее, резюме

def build_pres2():
    prs = new_prs()
    add_original(prs, 1)
    new_toc(prs, [
        "Виды наружного кровотечения",
        "Сравнительная характеристика",
        "Признаки кровопотери и объём",
        "Скрытое (внутреннее) кровотечение",
    ])
    add_original(prs, 2)   # признаки кровопотери (дубль уместен)
    add_original(prs, 3)   # три вида с иллюстрациями

    new_table(prs, "Сравнительная таблица видов кровотечения",
              ["Признак", "Артериальное", "Венозное", "Капиллярное"],
              [
                  ["Цвет крови", "Алый, яркий", "Тёмный, вишнёвый", "Красный"],
                  ["Характер", "Пульсирующей струёй", "Равномерной струёй", "Сочится"],
                  ["Скорость", "Очень быстрая", "Умеренная / быстрая", "Медленная"],
                  ["Опасность", "Критическая — минуты", "Высокая", "Обычно низкая"],
              ])

    new_table(prs, "Оценка объёма кровопотери",
              ["Объём", "Доля ОЦК", "Состояние"],
              [
                  ["До ~500 мл", "≈ 10%", "Слабость, жажда"],
                  ["500–1000 мл", "≈ 10–20%", "Бледность, тахикардия"],
                  ["1000–1500 мл", "≈ 20–30%", "Обмороки, холодный пот"],
                  [">1500–2000 мл", "> 30%", "Шок, угроза жизни"],
              ])

    new_list(prs, "Признаки скрытого (внутреннего) кровотечения", [
        "Снаружи крови может не быть видно",
        "Бледность, холодный пот, нарастающая слабость, жажда",
        "Слабый учащённый пульс, частое дыхание",
        "Боль и напряжение живота",
        "При травме груди — одышка, слабость, бледность",
        "Действия: СМП, положение, холод, не кормить и не поить",
    ], highlight="Внутреннее кровотечение на месте не останавливают — нужна СМП")

    new_summary(prs, [
        "Артериальное — алая пульсирующая струя",
        "Венозное — тёмная струя; капиллярное — сочение",
        "Потеря >30% ОЦК угрожает жизни",
        "При подозрении на внутреннее — срочный вызов СМП",
    ])
    add_original(prs, 28)
    path = OUT / "Признаки_наружного_кровотечения.pptx"
    prs.save(path)
    return path


# ═══════════════════ ПРЕЗЕНТАЦИЯ 3 ═══════════════════
# Исходник: 1, 5–13, 28
# Новое: оглавление, алгоритм жгута, записка, ошибки, когда нельзя, гемостатики, резюме

def build_pres3():
    prs = new_prs()
    add_original(prs, 1)
    new_toc(prs, [
        "Прямое давление и давящая повязка",
        "Пальцевое прижатие и максимальное сгибание",
        "Жгут и подручные средства",
        "Алгоритм, записка, ошибки и ограничения",
    ])

    add_original(prs, 5)   # прямое давление + давящая повязка
    add_original(prs, 6)   # сонная
    add_original(prs, 7)   # подключичная
    add_original(prs, 8)   # плечевая
    add_original(prs, 9)   # подмышечная
    add_original(prs, 10)  # бедренная
    add_original(prs, 11)  # максимальное сгибание
    add_original(prs, 12)  # жгут (с 60/30 мин)
    add_original(prs, 13)  # подручные средства

    # --- НОВОЕ ---
    new_scheme(prs, "Пошаговый алгоритм наложения жгута", [
        "Показания: давление и повязка неэффективны / отрыв",
        "Место: выше раны, на плечо или бедро",
        "На одежду / подкладку; первый тур — максимальное натяжение",
        "Кровь остановилась; пульс дистальнее не определяется",
        "Записка со временем; жгут должен быть виден",
        "Лимит: 60 мин (тепло) / 30 мин (холод)",
    ], key_idx=5)

    new_list(prs, "Записка под жгут — что указать", [
        "Точное время наложения: ЧЧ:ММ (и дата при необходимости)",
        "Фамилия / кто наложил — по возможности",
        "Кратко — место происшествия",
        "Вложить под жгут или прикрепить к одежде на видном месте",
        "Сообщить время бригаде СМП при передаче",
    ], highlight="Без записки легко превысить допустимое время жгута")

    new_list(prs, "Ошибки при наложении жгута", [
        "Слабое натяжение — венозный застой, усиление кровотечения",
        "Слишком далеко от раны — излишняя ишемия",
        "На голую кожу без подкладки",
        "Скрытие жгута одеждой — СМП может не заметить",
        "Нет записки со временем",
        "Жгут без показаний (при капиллярном кровотечении)",
    ])

    new_list(prs, "Когда НЕЛЬЗЯ накладывать жгут", [
        "Кровотечение останавливается давлением или давящей повязкой",
        "Капиллярное и умеренное венозное кровотечение",
        "Раны шеи, груди, живота, головы",
        "Нет конечности проксимальнее раны для размещения жгута",
    ], alert={2}, highlight="На шею, грудь, живот и голову жгут не накладывают")

    new_list(prs, "Гемостатические средства", [
        "Гемостатические салфетки / бинты из аптечки первой помощи",
        "Плотно вложить в рану / на рану и сильно придавить",
        "Затем — давящая повязка",
        "Применяют там, где жгут невозможен, или как усиление давления",
        "Не заменяют вызов скорой помощи",
    ])

    new_summary(prs, [
        "Приоритет: давление → повязка → жгут по показаниям",
        "Пальцевое прижатие и сгибание — временные приёмы",
        "Жгут: на одежду, с запиской; ≤60 мин / ≤30 мин",
        "Гемостатики усиливают давление там, где уместно",
    ], callout="t max ≤ 60 мин (тепло) · ≤ 30 мин (холод)")
    add_original(prs, 28)
    path = OUT / "Способы_временной_остановки_кровотечения.pptx"
    prs.save(path)
    return path


# ═══════════════════ ПРЕЗЕНТАЦИЯ 4 ═══════════════════
# Исходник: 1, 15, 16, 17, 18, 28 (+ дубли 2, 5 по желанию — не обязательно)
# Новое: оглавление, полный алгоритм, оценка состояния, приоритеты, после остановки, жгут, резюме

def build_pres4():
    prs = new_prs()
    add_original(prs, 1)
    new_toc(prs, [
        "Полный алгоритм действий",
        "Приоритетность способов остановки",
        "Профилактика травматического шока",
        "Подробный осмотр и контроль после остановки",
    ])

    new_scheme(prs, "Полный алгоритм (чек-лист)", [
        "Оценка обстановки → безопасность (СИЗ)",
        "Обзорный осмотр — поиск кровотечения",
        "Временная остановка кровотечения",
        "Признаки жизни: сознание → дыхание",
        "СЛР при отсутствии признаков жизни",
        "Вызов скорой медицинской помощи",
        "Подробный осмотр и помощь при травмах",
        "Положение, тепло, контроль, передача СМП",
    ], key_idx=2)

    new_list(prs, "Оценка состояния пострадавшего", [
        "При массивном кровотечении сначала остановите кровь",
        "Сознание: громко обратитесь, осторожно потрясите за плечи",
        "Дыхание: смотрю–слушаю–ощущаю не более 10 секунд",
        "При нескольких травмах — сначала то, что угрожает жизни",
    ])

    new_scheme(prs, "Приоритетность способов остановки", [
        "Прямое давление на рану",
        "Давящая повязка",
        "Жгут — по показаниям",
    ], key_idx=0)

    new_list(prs, "Приоритет при множественных травмах", [
        "Сначала — угрожающее жизни кровотечение",
        "Затем — дыхательные пути и дыхание / СЛР",
        "Затем — остальные кровотечения и раны",
        "Затем — иммобилизация и положение",
        "При нескольких пострадавших — приоритет детям и массивным кровотечениям",
    ])

    add_original(prs, 15)  # травматический шок
    add_original(prs, 16)  # профилактика шока
    add_original(prs, 17)  # подробный осмотр 1
    add_original(prs, 18)  # подробный осмотр 2

    new_list(prs, "Действия после остановки кровотечения", [
        "Контролируйте повязку: при промокании усильте, не снимая",
        "Проверяйте сознание, дыхание, цвет кожи",
        "Следите за жгутом: время, видимость, записка",
        "Согрейте пострадавшего, обеспечьте покой",
        "Передайте бригаде: что сделано и время наложения жгута",
    ])

    new_list(prs, "Когда и как ослаблять жгут", [
        "Планово ослаблять жгут на месте не рекомендуется",
        "Если время превысило 60 мин / 30 мин и СМП задерживается:",
        "Подготовьте прямое давление / давящую повязку",
        "Медленно ослабьте жгут, оценивая кровотечение",
        "Если кровь снова бьёт струёй — немедленно затяните и обновите время",
    ], alert={1}, highlight="Лимит жгута: ≤ 60 мин (тепло) · ≤ 30 мин (холод)")

    new_summary(prs, [
        "Безопасность → осмотр → кровь → признаки жизни → СМП",
        "Способы: давление → повязка → жгут",
        "Профилактика шока: кровь, положение, иммобилизация, тепло",
        "После остановки — контроль до передачи СМП",
    ], callout="Сначала угрожающее жизни кровотечение!")
    add_original(prs, 28)
    path = OUT / "Последовательность_остановки_кровотечения.pptx"
    prs.save(path)
    return path


# ═══════════════════ ПРЕЗЕНТАЦИЯ 5 ═══════════════════
# Исходник: 1, 14, 19–27, 28
# Новое: оглавление, шея артерия/вена, окклюзионная схема, инородный предмет,
#        иммобилизация с кровотечением, ампутация, перемещение при позвоночнике, резюме

def build_pres5():
    prs = new_prs()
    add_original(prs, 1)
    new_toc(prs, [
        "Голова, глаза, нос, носовое кровотечение",
        "Шея, грудь, инородный предмет",
        "Живот, таз, конечности",
        "Позвоночник и травматическая ампутация",
    ])

    add_original(prs, 19)  # голова
    add_original(prs, 20)  # глаз и нос
    add_original(prs, 14)  # носовое кровотечение (отдельный блок)

    add_original(prs, 21)  # шея
    new_list(prs, "Кровотечение при травме шеи: артерия или вена", [
        "Артериальное (сонная): алая кровь, пульсирующая струя",
        "Венозное (яремные): тёмная обильная струя; риск воздушной эмболии",
        "Прямое давление на рану (не пережимать обе сонные артерии)",
        "Давящая повязка через плечо — не тугая круговая на шее",
        "Жгут на шею НЕ накладывают",
    ], alert={4}, highlight="Не сдавливайте дыхательные пути повязкой")

    add_original(prs, 22)  # грудь
    new_scheme(prs, "Техника окклюзионной повязки", [
        "Придать полусидячее положение",
        "Воздухонепроницаемый материал на рану",
        "Зафиксировать пластырем (клапан или герметично)",
        "При сквозном ранении закрыть вход и выход",
        "Контролировать дыхание",
    ], key_idx=1)

    new_list(prs, "Инородный предмет в ране", [
        "НЕ ИЗВЛЕКАТЬ — предмет может тампонировать сосуд",
        "Обложить салфетками / бинтами валиками вокруг",
        "Давящая повязка поверх валиков с фиксацией предмета",
        "Не продавливать предмет вглубь",
        "Иммобилизировать область, вызвать СМП",
    ], highlight="Правило: зафиксировать — не извлекать!")

    add_original(prs, 23)  # живот закрытая
    add_original(prs, 24)  # живот открытая
    add_original(prs, 25)  # конечности
    add_original(prs, 26)  # иммобилизация

    new_list(prs, "Иммобилизация при переломе с кровотечением", [
        "Сначала остановите кровотечение",
        "Затем иммобилизация двух соседних суставов",
        "Шины — поверх одежды",
        "При открытом переломе не класть шину на выступающие отломки",
        "Костные отломки не вправлять",
    ])

    new_list(prs, "Первая помощь при отрыве конечности", [
        "Немедленно наложите жгут выше места отрыва (плечо / бедро)",
        "Давящая повязка на культю",
        "Записка со временем наложения жгута",
        "Сегмент: чистая ткань → пакет → пакет со холодом (не мочить)",
        "Передать сегмент вместе с пострадавшим бригаде СМП",
    ], alert={0}, highlight="Жгут при отрыве показан. t max ≤ 60 мин / 30 мин")

    add_original(prs, 27)  # позвоночник
    new_list(prs, "Правила перемещения при травме позвоночника", [
        "Перемещать только при угрозе жизни или для эвакуации",
        "Поверхность: ровная, жёсткая, горизонтальная",
        "Нужно несколько человек (обычно 3–5)",
        "Один постоянно удерживает голову и шейный отдел",
        "Перекладывание одним движением, без скручивания",
    ])

    new_summary(prs, [
        "Голова / нос: повязки; при потере сознания — боковое положение",
        "Шея и грудь: без кругового жгута; окклюзионная повязка",
        "Живот: органы не вправлять; конечности — кровь, затем шина",
        "Позвоночник — жёсткая поверхность и фиксация шеи",
    ], callout="Инородный предмет — зафиксировать, не извлекать!")
    add_original(prs, 28)
    path = OUT / "Кровотечение_при_ранениях_областей_тела.pptx"
    prs.save(path)
    return path


def main():
    paths = []
    for fn in (build_pres1, build_pres2, build_pres3, build_pres4, build_pres5):
        p = fn()
        prs = Presentation(str(p))
        print(f"OK: {p.name} · {len(prs.slides)} slides · {p.stat().st_size // 1024} KB")
        paths.append(p)

    (OUT / "README.md").write_text(
        "# Тема 2. Наружные кровотечения\n\n"
        "Исходные слайды из PDF перенесены **без изменений**.\n"
        "Новые слайды (Open Sans) — только там, где в исходнике не было информации.\n\n"
        "| Файл | Пункт |\n|---|---|\n"
        "| `Кровотечение_и_обзорный_осмотр.pptx` | 2.1 |\n"
        "| `Признаки_наружного_кровотечения.pptx` | 2.2 |\n"
        "| `Способы_временной_остановки_кровотечения.pptx` | 2.3 |\n"
        "| `Последовательность_остановки_кровотечения.pptx` | 2.4 |\n"
        "| `Кровотечение_при_ранениях_областей_тела.pptx` | 2.5 |\n",
        encoding="utf-8",
    )
    print("Done.")


if __name__ == "__main__":
    main()
