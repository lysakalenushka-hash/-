#!/usr/bin/env python3
"""
Признаки наружного кровотечения и кровопотери — стиль СУОТ.

Дополняет презентацию «Обзорный осмотр» (без дублирования её слайдов).
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu

from suot_style import (
    ACCENT_RED, BG_PANEL, LINE, MUTED, NUM_BG, OUT, ROOT, ROW_ALT,
    TABLE_HDR, TEAL, TEAL_DARK, TEAL_TEXT, TEXT, WHITE,
    blank, bullets, content_header, font, new_prs, oval, rect, round_rect,
    slide_thanks_suot, slide_title_suot, slide_toc_suot, tbox, teal_bubble,
    verify,
)


def slide_bridge(prs, num=3):
    slide = blank(prs)
    content_header(slide, "ЧЕМ ЭТА ТЕМА ДОПОЛНЯЕТ ОБЗОРНЫЙ ОСМОТР", num)
    round_rect(slide, Emu(700000), Emu(1500000), Emu(5400000), Emu(4800000), WHITE, line=LINE)
    rect(slide, Emu(700000), Emu(1500000), Emu(5400000), Emu(700000), BG_PANEL)
    tbox(slide, Emu(900000), Emu(1600000), Emu(5000000), Emu(500000),
         "Уже в теме «Обзорный осмотр»", size=14, bold=True, color=TEXT,
         align=PP_ALIGN.CENTER)
    bullets(slide, Emu(950000), Emu(2450000), Emu(4900000), Emu(3500000), [
        "понятие кровотечения;",
        "как выглядят виды кровотечения;",
        "как провести обзорный осмотр;",
        "что делать при обнаружении крови.",
    ], size=14)
    round_rect(slide, Emu(6500000), Emu(1500000), Emu(5400000), Emu(4800000), WHITE, line=TEAL)
    rect(slide, Emu(6500000), Emu(1500000), Emu(5400000), Emu(700000), TEAL)
    tbox(slide, Emu(6700000), Emu(1600000), Emu(5000000), Emu(500000),
         "Разбираем сейчас", size=14, bold=True, color=TEAL_TEXT,
         align=PP_ALIGN.CENTER)
    bullets(slide, Emu(6750000), Emu(2450000), Emu(4900000), Emu(3500000), [
        "как читать признаки кровопотери;",
        "когда крови «не видно», а угроза есть;",
        "как оценить интенсивность;",
        "сравнение видов и смешанное кровотечение;",
        "что маскирует признаки.",
    ], size=14)
    return slide


def slide_signs_groups(prs, num=4):
    slide = blank(prs)
    content_header(slide, "ПРИЗНАКИ КРОВОПОТЕРИ — ПО ГРУППАМ", num)
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
        round_rect(slide, x, Emu(1500000), w, Emu(4800000), WHITE, line=TEAL)
        rect(slide, x, Emu(1500000), w, Emu(700000), TEAL)
        tbox(slide, x + Emu(100000), Emu(1600000), w - Emu(200000), Emu(500000),
             title, size=14, bold=True, color=TEAL_TEXT, align=PP_ALIGN.CENTER)
        bullets(slide, x + Emu(150000), Emu(2450000), w - Emu(300000), Emu(3500000),
                items, size=13, marker="☐")
        x += w + Emu(250000)
    return slide


def slide_when_signs(prs, num=5):
    slide = blank(prs)
    content_header(slide, "КОГДА ПОЯВЛЯЮТСЯ ПРИЗНАКИ КРОВОПОТЕРИ", num)
    cards = [
        ("Кровотечение продолжается",
         "Признаки нарастают на фоне видимой крови."),
        ("Кровотечение уже остановлено",
         "Признаки могут сохраняться — потерянная кровь не возвращается сразу."),
        ("Видимой крови нет",
         "Возможно внутреннее (скрытое) кровотечение — оценивайте состояние."),
    ]
    y = Emu(1550000)
    for title, text in cards:
        round_rect(slide, Emu(700000), y, Emu(11000000), Emu(1300000), BG_PANEL)
        oval(slide, Emu(900000), y + Emu(350000), Emu(550000), Emu(550000), TEAL)
        tbox(slide, Emu(1700000), y + Emu(200000), Emu(9600000), Emu(400000),
             title, size=15, bold=True, color=TEAL_DARK)
        tbox(slide, Emu(1700000), y + Emu(650000), Emu(9600000), Emu(450000),
             text, size=13, color=MUTED)
        y += Emu(1500000)
    return slide


def slide_hidden(prs, num=6):
    slide = blank(prs)
    content_header(slide, "СКРЫТАЯ (ВНУТРЕННЯЯ) КРОВОПОТЕРЯ", num)
    tbox(slide, Emu(700000), Emu(1500000), Emu(11000000), Emu(900000),
         "Снаружи крови может не быть видно, но состояние ухудшается. "
         "Ориентируйтесь на признаки кровопотери и механизм травмы.",
         size=15, color=TEXT)
    bullets(slide, Emu(700000), Emu(2600000), Emu(11000000), Emu(2400000), [
        "Настораживающие ситуации: удар в живот / грудь, падение с высоты, ДТП.",
        "Боль и напряжение живота, нарастающая слабость, холодный пот, жажда.",
        "На месте внутреннее кровотечение не останавливают — СМП, покой, контроль.",
    ], size=14)
    teal_bubble(
        slide, Emu(700000), Emu(5300000), Emu(11000000), Emu(1100000),
        "Нет видимой крови ≠ нет угрозы. Смотрите на состояние пострадавшего.",
        size=14,
    )
    return slide


def slide_intensity_scale(prs, num=7):
    slide = blank(prs)
    content_header(slide, "КАК ОЦЕНИТЬ ИНТЕНСИВНОСТЬ КРОВОТЕЧЕНИЯ", num)
    levels = [
        (TEAL, "Слабая",
         "Капиллярное сочение, небольшие ссадины/порезы. Угрозы жизни обычно нет."),
        (TEAL_DARK, "Средняя",
         "Стойкий «ручей», пропитывание повязки/ткани. Без остановки — значительная потеря."),
        (ACCENT_RED, "Интенсивная",
         "Струя / быстрое пропитывание одежды, лужа крови — останавливать немедленно."),
    ]
    y = Emu(1500000)
    for color, title, text in levels:
        oval(slide, Emu(700000), y + Emu(150000), Emu(550000), Emu(550000), color)
        tbox(slide, Emu(700000), y + Emu(150000), Emu(550000), Emu(550000),
             title[0], size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE)
        tbox(slide, Emu(1450000), y, Emu(10000000), Emu(400000),
             title, size=15, bold=True, color=TEXT)
        tbox(slide, Emu(1450000), y + Emu(450000), Emu(10000000), Emu(700000),
             text, size=13, color=MUTED)
        y += Emu(1450000)
    return slide


def slide_compare_table(prs, num=8):
    slide = blank(prs)
    content_header(slide, "СРАВНЕНИЕ ПРИЗНАКОВ ПО ВИДАМ", num)
    tbox(slide, Emu(700000), Emu(1400000), Emu(11000000), Emu(350000),
         "Детализация к видам из темы «Обзорный осмотр».", size=12, color=MUTED)
    headers = ["Признак", "Артериальное", "Венозное", "Капиллярное"]
    rows = [
        ["Цвет", "Алый, яркий", "Тёмный, вишнёвый", "Красный"],
        ["Характер", "Пульсирующая струя", "Равномерный «ручей»", "Сочится"],
        ["Скорость", "Очень быстрая", "Умеренная / быстрая", "Медленная"],
        ["Опасность", "Критическая — минуты", "Высокая", "Обычно низкая"],
        ["Картина", "Лужа алого, одежда мокнет", "Стойкое истечение", "Ссадина, порез"],
    ]
    table = slide.shapes.add_table(
        len(rows) + 1, len(headers),
        Emu(700000), Emu(1850000), Emu(11000000), Emu(4300000)
    ).table
    for c, h in enumerate(headers):
        cell = table.cell(0, c)
        cell.fill.solid()
        cell.fill.fore_color.rgb = TABLE_HDR
        cell.text = h
        for p in cell.text_frame.paragraphs:
            for r in p.runs:
                font(r, 12, True, WHITE)
    for ri, row in enumerate(rows, 1):
        for c, val in enumerate(row):
            cell = table.cell(ri, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = ROW_ALT if ri % 2 == 0 else WHITE
            cell.text = val
            alert = any(k in val for k in ("Алый", "Пульсир", "Критическ", "Очень"))
            for p in cell.text_frame.paragraphs:
                for r in p.runs:
                    font(r, 11, alert, ACCENT_RED if alert else TEXT)
    return slide


def slide_mixed(prs, num=9):
    slide = blank(prs)
    content_header(slide, "СМЕШАННОЕ КРОВОТЕЧЕНИЕ — КОГДА ПОДОЗРЕВАТЬ", num)
    bullets(slide, Emu(700000), Emu(1550000), Emu(11000000), Emu(2800000), [
        "Одновременно признаки артериального, венозного и капиллярного кровотечения.",
        "Частый пример — отрыв конечности или обширное размозжение.",
        "Опасно из‑за артериального компонента: быстрая массивная кровопотеря.",
        "Не тратьте время на «точную классификацию» — останавливайте по интенсивности.",
    ], size=14)
    teal_bubble(
        slide, Emu(700000), Emu(4700000), Emu(11000000), Emu(1500000),
        "Правило: есть интенсивное истечение крови — действуйте как при угрозе жизни.",
        size=15,
    )
    return slide


def slide_masking(prs, num=10):
    slide = blank(prs)
    content_header(slide, "ЧТО МАСКИРУЕТ ПРИЗНАКИ КРОВОТЕЧЕНИЯ", num)
    items = [
        ("Тёмная / плотная одежда",
         "Кровь плохо видна — ощупайте, приподнимите одежду при осмотре."),
        ("Кровь под пострадавшим",
         "Лужа может быть сзади / под телом — осмотрите вокруг."),
        ("Холод, дождь, грязь",
         "Сложно оценить цвет кожи — опирайтесь на слабость, пульс, дыхание."),
        ("Несколько пострадавших",
         "Интенсивное кровотечение ищите в первую очередь у каждого."),
    ]
    y = Emu(1450000)
    for title, text in items:
        round_rect(slide, Emu(700000), y, Emu(11000000), Emu(1100000), BG_PANEL)
        tbox(slide, Emu(950000), y + Emu(120000), Emu(10400000), Emu(350000),
             title, size=14, bold=True, color=TEAL_DARK)
        tbox(slide, Emu(950000), y + Emu(500000), Emu(10400000), Emu(450000),
             text, size=12, color=MUTED)
        y += Emu(1200000)
    return slide


def slide_volume(prs, num=11):
    slide = blank(prs)
    content_header(slide, "ОРИЕНТИРЫ ПО ОБЪЁМУ КРОВОПОТЕРИ", num)
    headers = ["Объём (ориентир)", "Доля ОЦК", "Типичное состояние"]
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
         "На месте точный объём измерить нельзя — таблица помогает понять тяжесть "
         "по признакам состояния.",
         size=13, color=MUTED)
    return slide


def slide_summary(prs, num=12):
    slide = blank(prs)
    content_header(slide, "ГЛАВНОЕ ЗАПОМНИТЬ", num)
    points = [
        "Признаки кровопотери — по самочувствию, коже, дыханию и пульсу.",
        "Признаки возможны и без видимой крови — думайте о скрытой кровопотере.",
        "Интенсивность важнее «точного вида сосуда» для решения об остановке.",
        "Одежда, положение тела и погода могут маскировать кровь — осматривайте внимательно.",
    ]
    y = Emu(1500000)
    for pt in points:
        round_rect(slide, Emu(700000), y, Emu(11000000), Emu(1000000), BG_PANEL)
        oval(slide, Emu(900000), y + Emu(250000), Emu(500000), Emu(500000), TEAL)
        tbox(slide, Emu(1600000), y + Emu(200000), Emu(9800000), Emu(600000),
             pt, size=13, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)
        y += Emu(1150000)
    return slide


def build():
    prs = new_prs()
    slide_title_suot(prs, "ПРИЗНАКИ НАРУЖНОГО\nКРОВОТЕЧЕНИЯ И КРОВОПОТЕРИ")
    slide_toc_suot(prs, [
        "Чем тема дополняет обзорный осмотр",
        "Признаки кровопотери по группам",
        "Когда признаки есть без видимой крови",
        "Оценка интенсивности и сравнение видов",
        "Смешанное кровотечение, маскировка, объём",
    ], 2)
    slide_bridge(prs, 3)
    slide_signs_groups(prs, 4)
    slide_when_signs(prs, 5)
    slide_hidden(prs, 6)
    slide_intensity_scale(prs, 7)
    slide_compare_table(prs, 8)
    slide_mixed(prs, 9)
    slide_masking(prs, 10)
    slide_volume(prs, 11)
    slide_summary(prs, 12)
    slide_thanks_suot(prs)

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
