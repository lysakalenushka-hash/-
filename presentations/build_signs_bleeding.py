#!/usr/bin/env python3
"""
2) Признаки наружного кровотечения и кровопотери

Угол: распознавание — как выглядит наружное кровотечение, какие признаки
кровопотери у организма, как оценить интенсивность и объём.
Без алгоритма обзорного осмотра и чек-листа действий (это в 1-й презентации).

Стиль: «Организационно-правовые аспекты оказания первой помощи».
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu

from legal_style import (
    ACCENT_RED, BG_BAR, CREAM, LINE, MUTED, NUM_BG, OUT, ROOT, ROW_ALT,
    TABLE_HDR, TEXT, WHITE, blank, bullets, content_header, cream_note, font,
    new_prs, oval, pic_fit, rect, rich_tbox, round_rect, slide_thanks,
    slide_title, slide_toc, tbox, verify,
)


def slide_blood_loss_state(prs, num=3):
    """Кровопотеря как состояние организма — не повтор «что такое кровотечение» в том же виде."""
    slide = blank(prs)
    content_header(slide, "КРОВОПОТЕРЯ — ЧТО ПРОИСХОДИТ С ОРГАНИЗМОМ", num)
    pic_fit(slide, "def_man.png", Emu(500000), Emu(1500000),
            Emu(3000000), Emu(5000000))
    cream_note(
        slide, Emu(3400000), Emu(1600000), Emu(8600000), Emu(2200000),
        "Когда кровь покидает сосуды, снижается доставка кислорода и питательных "
        "веществ к органам. Это и есть кровопотеря — безвозвратная утрата части крови, "
        "из‑за которой органы начинают хуже работать или перестают функционировать.",
        size=14,
    )
    bullets(slide, Emu(3400000), Emu(4200000), Emu(8600000), Emu(2200000), [
        "Тяжесть зависит от объёма потери, вида сосуда и органа, который он питал.",
        "Самое опасное — острая массивная кровопотеря (быстрая большая потеря).",
        "Даже «умеренная» потеря без остановки ведёт к поздним осложнениям.",
    ], size=13)
    return slide


def slide_signs_groups(prs, num=4):
    slide = blank(prs)
    content_header(slide, "ПРИЗНАКИ КРОВОПОТЕРИ", num)
    groups = [
        ("Самочувствие", [
            "резкая общая слабость;",
            "чувство жажды;",
            "головокружение;",
            "мелькание «мушек»;",
            "обморок при попытке встать.",
        ]),
        ("Кожа", [
            "бледная;",
            "влажная;",
            "холодная на ощупь.",
        ]),
        ("Дыхание и сердце", [
            "учащённое сердцебиение;",
            "частое дыхание.",
        ]),
    ]
    x = Emu(700000)
    widths = [Emu(4000000), Emu(3200000), Emu(3400000)]
    for (title, items), w in zip(groups, widths):
        rect(slide, x, Emu(1500000), w, Emu(4200000), WHITE, line=LINE)
        rect(slide, x, Emu(1500000), w, Emu(700000), NUM_BG)
        tbox(slide, x + Emu(100000), Emu(1600000), w - Emu(200000), Emu(500000),
             title, size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        bullets(slide, x + Emu(150000), Emu(2450000), w - Emu(300000), Emu(3000000),
                items, size=13, marker="☐")
        x += w + Emu(250000)
    cream_note(
        slide, Emu(700000), Emu(5900000), Emu(11000000), Emu(800000),
        "Эти признаки возможны при продолжающемся, уже остановленном кровотечении "
        "и даже при отсутствии видимой крови.",
        size=12,
    )
    return slide


def slide_hidden(prs, num=5):
    slide = blank(prs)
    content_header(slide, "КОГДА КРОВИ «НЕ ВИДНО»", num)
    tbox(slide, Emu(700000), Emu(1500000), Emu(11000000), Emu(800000),
         "Наружное кровотечение — кровь изливается наружу. Но кровопотеря бывает "
         "и скрытой (внутренней): снаружи крови нет, а состояние ухудшается.",
         size=15, color=TEXT)
    bullets(slide, Emu(700000), Emu(2600000), Emu(11000000), Emu(2400000), [
        "Настораживают: удар в живот/грудь, падение с высоты, ДТП.",
        "Смотрите на слабость, бледность, холодный пот, жажду, боль в животе.",
        "Внутреннее кровотечение на месте не останавливают — нужны СМП, покой, контроль.",
    ], size=14)
    cream_note(
        slide, Emu(700000), Emu(5400000), Emu(11000000), Emu(1000000),
        "Нет видимой крови ≠ нет угрозы. Оценивайте состояние пострадавшего.",
        size=14,
    )
    return slide


def slide_types(prs, num=6):
    slide = blank(prs)
    content_header(slide, "НАРУЖНОЕ КРОВОТЕЧЕНИЕ: КАК РАСПОЗНАТЬ ВИД", num)
    cols = [
        ("bleed_arterial.png", "Артериальное",
         "Пульсирующая алая струя; лужа алого цвета; одежда быстро пропитывается. Наиболее опасно."),
        ("bleed_venous.png", "Венозное",
         "Кровь тёмно-вишнёвая, вытекает «ручьём». Скорость меньше, но остановка обязательна."),
        ("bleed_capillary.png", "Капиллярное",
         "Ссадины, порезы, царапины. Как правило, без непосредственной угрозы жизни."),
    ]
    x = Emu(700000)
    cw = Emu(3600000)
    for img, title, desc in cols:
        tbox(slide, x, Emu(1450000), cw, Emu(450000), title, size=15, bold=True,
             color=TEXT, align=PP_ALIGN.CENTER)
        pic_fit(slide, img, x + Emu(150000), Emu(2000000), cw - Emu(300000), Emu(3000000))
        tbox(slide, x, Emu(5200000), cw, Emu(1400000), desc, size=11, color=TEXT,
             align=PP_ALIGN.CENTER)
        x += cw + Emu(250000)
    return slide


def slide_table(prs, num=7):
    slide = blank(prs)
    content_header(slide, "СРАВНЕНИЕ ПРИЗНАКОВ ПО ВИДАМ", num)
    headers = ["Признак", "Артериальное", "Венозное", "Капиллярное"]
    rows = [
        ["Цвет", "Алый, яркий", "Тёмный, вишнёвый", "Красный"],
        ["Характер", "Пульсирующая струя", "Равномерный «ручей»", "Сочится"],
        ["Скорость", "Очень быстрая", "Умеренная / быстрая", "Медленная"],
        ["Опасность", "Критическая — минуты", "Высокая", "Обычно низкая"],
    ]
    table = slide.shapes.add_table(
        len(rows) + 1, len(headers),
        Emu(700000), Emu(1600000), Emu(11000000), Emu(4000000)
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
                    font(r, 12, alert, ACCENT_RED if alert else TEXT)
    cream_note(
        slide, Emu(700000), Emu(5900000), Emu(11000000), Emu(800000),
        "На месте вид определить сложно — ориентируйтесь на интенсивность истечения.",
        size=13,
    )
    return slide


def slide_mixed_intensity(prs, num=8):
    slide = blank(prs)
    content_header(slide, "СМЕШАННОЕ И ИНТЕНСИВНОСТЬ", num)
    # left mixed
    rect(slide, Emu(700000), Emu(1500000), Emu(5400000), Emu(4800000), WHITE, line=LINE)
    rect(slide, Emu(700000), Emu(1500000), Emu(5400000), Emu(700000), NUM_BG)
    tbox(slide, Emu(900000), Emu(1600000), Emu(5000000), Emu(500000),
         "Смешанное кровотечение", size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    bullets(slide, Emu(950000), Emu(2450000), Emu(4900000), Emu(3500000), [
        "Одновременно артерия + вена + капилляры",
        "Часто при отрыве конечности",
        "Опасно из‑за артериального компонента",
    ], size=13)
    # right intensity
    rect(slide, Emu(6500000), Emu(1500000), Emu(5400000), Emu(4800000), WHITE, line=LINE)
    rect(slide, Emu(6500000), Emu(1500000), Emu(5400000), Emu(700000), BG_BAR)
    tbox(slide, Emu(6700000), Emu(1600000), Emu(5000000), Emu(500000),
         "Интенсивность", size=15, bold=True, color=TEXT, align=PP_ALIGN.CENTER)
    bullets(slide, Emu(6750000), Emu(2450000), Emu(4900000), Emu(3500000), [
        "Слабая — сочение, ссадины",
        "Средняя — стойкий «ручей», мокнет повязка",
        "Интенсивная — струя, лужа, быстро мокнет одежда",
    ], size=13)
    return slide


def slide_masking(prs, num=9):
    slide = blank(prs)
    content_header(slide, "ЧТО МАСКИРУЕТ ПРИЗНАКИ", num)
    items = [
        ("Тёмная одежда", "Кровь плохо видна — ощупайте, приподнимите ткань."),
        ("Кровь под телом", "Лужа сзади / под пострадавшим — осмотрите вокруг."),
        ("Холод, дождь, грязь", "Цвет кожи обманчив — смотрите слабость, пульс, дыхание."),
        ("Самопомощь", "Пострадавший может прижать рану — кровь «скрыта» рукой/тканью."),
    ]
    y = Emu(1500000)
    for title, text in items:
        rect(slide, Emu(700000), y, Emu(11000000), Emu(1100000), WHITE, line=LINE)
        tbox(slide, Emu(950000), y + Emu(150000), Emu(10400000), Emu(350000),
             title, size=14, bold=True, color=TEXT)
        tbox(slide, Emu(950000), y + Emu(550000), Emu(10400000), Emu(400000),
             text, size=13, color=MUTED)
        y += Emu(1200000)
    return slide


def slide_volume(prs, num=10):
    slide = blank(prs)
    content_header(slide, "ОРИЕНТИРЫ ПО ОБЪЁМУ КРОВОПОТЕРИ", num)
    headers = ["Объём", "Доля ОЦК", "Типичное состояние"]
    rows = [
        ["До ~500 мл", "≈ 10%", "Слабость, жажда"],
        ["500–1000 мл", "≈ 10–20%", "Бледность, учащённый пульс"],
        ["1000–1500 мл", "≈ 20–30%", "Обмороки, холодный пот"],
        [">1500–2000 мл", "> 30%", "Шок, угроза жизни"],
    ]
    table = slide.shapes.add_table(
        len(rows) + 1, len(headers),
        Emu(700000), Emu(1550000), Emu(11000000), Emu(3800000)
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
            alert = ">" in val or "Шок" in val
            for p in cell.text_frame.paragraphs:
                for r in p.runs:
                    font(r, 13, alert, ACCENT_RED if alert else TEXT)
    tbox(slide, Emu(700000), Emu(5600000), Emu(11000000), Emu(800000),
         "Точный объём на месте не измерить — таблица помогает понять тяжесть "
         "по признакам состояния.",
         size=13, color=MUTED)
    return slide


def slide_summary(prs, num=11):
    slide = blank(prs)
    content_header(slide, "ГЛАВНОЕ ЗАПОМНИТЬ", num)
    points = [
        "Кровопотеря — угроза органам из‑за нехватки кислорода и питательных веществ.",
        "Признаки: слабость, жажда, бледность, холодный пот, частый пульс и дыхание.",
        "Артериальное / венозное / капиллярное различают по цвету и характеру струи.",
        "Интенсивность важнее «точного вида»; крови может не быть видно снаружи.",
    ]
    y = Emu(1550000)
    for pt in points:
        rect(slide, Emu(700000), y, Emu(11000000), Emu(1000000), WHITE, line=LINE)
        rect(slide, Emu(700000), y, Emu(90000), Emu(1000000), ACCENT_RED)
        tbox(slide, Emu(1000000), y + Emu(200000), Emu(10400000), Emu(600000),
             pt, size=13, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)
        y += Emu(1150000)
    return slide


def build():
    prs = new_prs()
    slide_title(prs, "ПРИЗНАКИ НАРУЖНОГО\nКРОВОТЕЧЕНИЯ И КРОВОПОТЕРИ")
    slide_toc(prs, [
        "Кровопотеря: что происходит с организмом",
        "Признаки кровопотери и скрытая потеря крови",
        "Как распознать виды наружного кровотечения",
        "Смешанное кровотечение и интенсивность",
        "Маскировка признаков и ориентиры по объёму",
    ], 2)
    slide_blood_loss_state(prs, 3)
    slide_signs_groups(prs, 4)
    slide_hidden(prs, 5)
    slide_types(prs, 6)
    slide_table(prs, 7)
    slide_mixed_intensity(prs, 8)
    slide_masking(prs, 9)
    slide_volume(prs, 10)
    slide_summary(prs, 11)
    slide_thanks(prs)

    name = "Признаки_наружного_кровотечения_и_кровопотери.pptx"
    path = OUT / name
    prs.save(path)
    verify(path)
    path2 = ROOT / name
    prs.save(path2)
    print("Saved:", path2)
    return path2


if __name__ == "__main__":
    build()
