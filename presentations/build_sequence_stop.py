#!/usr/bin/env python3
"""
Последовательность выполнения мероприятий по остановке кровотечения

Угол: порядок решений и ветвления алгоритма.
Не дублирует техники (давление/повязка/жгут подробно — другая презентация)
и не разбирает обзорный осмотр и признаки кровопотери.

Стиль: «Организационно-правовые аспекты оказания первой помощи».
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu

from legal_style import (
    ACCENT_RED, CREAM, LINE, MUTED, NUM_BG, OUT, ROOT, TEXT, WHITE,
    blank, bullets, content_header, cream_note, new_prs, oval, pic_fit, rect,
    slide_thanks, slide_title, slide_toc, tbox, verify,
)


def slide_three_start(prs, num=3):
    slide = blank(prs)
    content_header(slide, "ТРИ ПЕРВЫХ ШАГА", num)
    tbox(slide, Emu(700000), Emu(1500000), Emu(11000000), Emu(700000),
         "При травме человек, оказывающий первую помощь, действует в фиксированном порядке.",
         size=15, color=TEXT)
    cards = [
        ("1", "Безопасность", "Обеспечить безопасные условия для оказания помощи"),
        ("2", "Обзорный осмотр", "Определить наличие и интенсивность кровотечения"),
        ("3", "Выбор способа", "Остановить кровь наиболее подходящим приёмом"),
    ]
    x = Emu(700000)
    for n, title, desc in cards:
        rect(slide, x, Emu(2500000), Emu(3600000), Emu(3400000), WHITE, line=LINE)
        oval(slide, x + Emu(1400000), Emu(2750000), Emu(800000), Emu(800000), NUM_BG)
        tbox(slide, x + Emu(1400000), Emu(2750000), Emu(800000), Emu(800000), n,
             size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE)
        tbox(slide, x + Emu(150000), Emu(3800000), Emu(3300000), Emu(600000),
             title, size=16, bold=True, color=TEXT, align=PP_ALIGN.CENTER)
        tbox(slide, x + Emu(200000), Emu(4500000), Emu(3200000), Emu(1100000),
             desc, size=13, color=MUTED, align=PP_ALIGN.CENTER)
        x += Emu(3800000)
    return slide


def slide_pressure_first(prs, num=4):
    slide = blank(prs)
    content_header(slide, "ПЕРВЫЙ ВЫБОР — ПРЯМОЕ ДАВЛЕНИЕ", num)
    pic_fit(slide, "direct_pressure.png", Emu(500000), Emu(1500000),
            Emu(5500000), Emu(4800000))
    bullets(slide, Emu(6400000), Emu(1600000), Emu(5800000), Emu(3200000), [
        "При интенсивном кровотечении в первую очередь — прямое давление на рану.",
        "Давление — в перчатках или через ткань.",
        "Цель шага: быстро снизить потерю крови, затем решить — повязка или жгут.",
    ], size=14)
    cream_note(
        slide, Emu(6400000), Emu(5000000), Emu(5800000), Emu(1300000),
        "Давление — старт алгоритма, а не единственный способ.",
        size=13,
    )
    return slide


def slide_when_skip_pressure(prs, num=5):
    slide = blank(prs)
    content_header(slide, "КОГДА ДАВЛЕНИЕ НЕ ПОДХОДИТ", num)
    tbox(slide, Emu(700000), Emu(1500000), Emu(11000000), Emu(800000),
         "Если прямое давление невозможно, опасно или явно неэффективно — "
         "сразу давящая повязка и/или жгут.",
         size=15, color=TEXT)
    cases = [
        ("Инородное тело в ране", "Давление на предмет опасно; фиксируют повязкой"),
        ("Открытый перелом", "Костные отломки в ране — давление неэффективно/опасно"),
        ("Нельзя прижать рану", "Анатомия / доступ — переход к повязке или жгуту"),
    ]
    y = Emu(2500000)
    for title, desc in cases:
        rect(slide, Emu(700000), y, Emu(11000000), Emu(1100000), WHITE, line=LINE)
        rect(slide, Emu(700000), y, Emu(90000), Emu(1100000), ACCENT_RED)
        tbox(slide, Emu(1000000), y + Emu(150000), Emu(10400000), Emu(400000),
             title, size=15, bold=True, color=TEXT)
        tbox(slide, Emu(1000000), y + Emu(550000), Emu(10400000), Emu(400000),
             desc, size=13, color=MUTED)
        y += Emu(1250000)
    return slide


def slide_amputation(prs, num=6):
    slide = blank(prs)
    content_header(slide, "ОБШИРНОЕ ПОВРЕЖДЕНИЕ И ОТРЫВ", num)
    pic_fit(slide, "tourniquet_1.png", Emu(500000), Emu(1500000),
            Emu(5200000), Emu(4800000))
    bullets(slide, Emu(6000000), Emu(1600000), Emu(6200000), Emu(3500000), [
        "Обширное или множественное повреждение конечности.",
        "Разрушение или отрыв (ампутация).",
        "В этих случаях жгут накладывают немедленно — "
        "не тратя время на попытки давления «по учебнику».",
    ], size=14, alert={2})
    cream_note(
        slide, Emu(6000000), Emu(5300000), Emu(6200000), Emu(1200000),
        "Отрыв крупных частей (кисть, стопа и выше) — показание к жгуту.",
        size=13,
    )
    return slide


def slide_if_pressure_fails(prs, num=7):
    slide = blank(prs)
    content_header(slide, "ДАВЛЕНИЕ НЕ ОСТАНОВИЛО КРОВЬ", num)
    steps = [
        ("1", "Продолжается кровотечение", "Прямое давление оказалось недостаточным"),
        ("2", "Жгут выше раны", "Между раной и сердцем — на конечность"),
        ("3", "Контроль эффекта", "Кровотечение должно остановиться"),
    ]
    y = Emu(1600000)
    for n, title, desc in steps:
        oval(slide, Emu(700000), y, Emu(700000), Emu(700000),
             ACCENT_RED if n == "2" else NUM_BG)
        tbox(slide, Emu(700000), y, Emu(700000), Emu(700000), n,
             size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE)
        rect(slide, Emu(1600000), y, Emu(10000000), Emu(700000), WHITE, line=LINE)
        tbox(slide, Emu(1800000), y + Emu(80000), Emu(9600000), Emu(300000),
             title, size=15, bold=True, color=TEXT)
        tbox(slide, Emu(1800000), y + Emu(380000), Emu(9600000), Emu(250000),
             desc, size=13, color=MUTED)
        y += Emu(1100000)
    cream_note(
        slide, Emu(700000), Emu(5200000), Emu(11000000), Emu(1200000),
        "Жгут — следующий шаг после неудачи давления, а не «запасной» приём в конце.",
        size=13,
    )
    return slide


def slide_if_pressure_works(prs, num=8):
    slide = blank(prs)
    content_header(slide, "ДАВЛЕНИЕ ПОМОГЛО — ЧТО ДАЛЬШЕ", num)
    pic_fit(slide, "pressure_bandage.png", Emu(500000), Emu(1500000),
            Emu(5200000), Emu(4800000))
    cards = [
        ("А", "Наложить давящую повязку",
         "Фиксирует результат давления, освобождает руки"),
        ("Б", "Повязка неэффективна",
         "Быстро промокает / не держит — наложить жгут"),
        ("В", "Нет повязки и жгута",
         "Продолжать прямое давление до медицинской помощи"),
    ]
    y = Emu(1550000)
    for n, title, desc in cards:
        rect(slide, Emu(6000000), y, Emu(6200000), Emu(1400000), WHITE, line=LINE)
        oval(slide, Emu(6200000), y + Emu(300000), Emu(600000), Emu(600000), NUM_BG)
        tbox(slide, Emu(6200000), y + Emu(300000), Emu(600000), Emu(600000), n,
             size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE)
        tbox(slide, Emu(7000000), y + Emu(250000), Emu(5000000), Emu(400000),
             title, size=14, bold=True, color=TEXT)
        tbox(slide, Emu(7000000), y + Emu(700000), Emu(5000000), Emu(500000),
             desc, size=12, color=MUTED)
        y += Emu(1600000)
    return slide


def slide_flowchart(prs, num=9):
    slide = blank(prs)
    content_header(slide, "СВОДНЫЙ АЛГОРИТМ", num)
    rows = [
        ("Старт", "Безопасность → обзорный осмотр → выбор способа"),
        ("Интенсивное кровотечение", "Прямое давление на рану"),
        ("Давление нельзя / опасно", "Повязка и/или жгут сразу"),
        ("Отрыв / разрушение конечности", "Жгут немедленно"),
        ("Давление не помогло", "Жгут выше раны"),
        ("Давление помогло", "Давящая повязка; если нет — жгут"),
        ("Нет средств", "Продолжать давление до медпомощи"),
    ]
    y = Emu(1450000)
    for left, right in rows:
        rect(slide, Emu(700000), y, Emu(4200000), Emu(650000), NUM_BG)
        tbox(slide, Emu(850000), y, Emu(3900000), Emu(650000), left,
             size=12, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
        rect(slide, Emu(5000000), y, Emu(7000000), Emu(650000), WHITE, line=LINE)
        tbox(slide, Emu(5200000), y, Emu(6600000), Emu(650000), right,
             size=13, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)
        y += Emu(720000)
    return slide


def slide_checklist(prs, num=10):
    slide = blank(prs)
    content_header(slide, "КОНТРОЛЬНЫЙ СПИСОК НА МЕСТЕ", num)
    left = [
        "Место безопасно для оказания помощи?",
        "Кровотечение найдено и оценено по интенсивности?",
        "Начато прямое давление (если можно)?",
        "Есть показания сразу к жгуту?",
    ]
    right = [
        "Давление остановило кровь → повязка?",
        "Повязка держит / не промокает насквозь?",
        "При жгуте — время отмечено, жгут на виду?",
        "Давление продолжается, если нет повязки/жгута?",
    ]
    tbox(slide, Emu(700000), Emu(1500000), Emu(5500000), Emu(400000),
         "До остановки", size=14, bold=True, color=TEXT)
    tbox(slide, Emu(6800000), Emu(1500000), Emu(5500000), Emu(400000),
         "После / параллельно", size=14, bold=True, color=TEXT)
    bullets(slide, Emu(700000), Emu(2000000), Emu(5500000), Emu(4000000), left, size=13)
    bullets(slide, Emu(6800000), Emu(2000000), Emu(5500000), Emu(4000000), right, size=13)
    return slide


def slide_summary(prs, num=11):
    slide = blank(prs)
    content_header(slide, "ГЛАВНОЕ ЗАПОМНИТЬ", num)
    points = [
        "Порядок: безопасность → обзорный осмотр → способ остановки.",
        "Старт при интенсивном кровотечении — прямое давление.",
        "Нельзя давить / отрыв конечности — сразу повязка и/или жгут.",
        "Давление помогло → повязка; не помогло / повязка не держит → жгут.",
        "Нет средств — не отпускать давление до передачи медикам.",
    ]
    y = Emu(1500000)
    for pt in points:
        rect(slide, Emu(700000), y, Emu(11000000), Emu(850000), WHITE, line=LINE)
        rect(slide, Emu(700000), y, Emu(90000), Emu(850000), ACCENT_RED)
        tbox(slide, Emu(1000000), y + Emu(150000), Emu(10400000), Emu(550000),
             pt, size=14, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)
        y += Emu(950000)
    return slide


def build():
    prs = new_prs()
    slide_title(prs, "ПОСЛЕДОВАТЕЛЬНОСТЬ МЕРОПРИЯТИЙ\nПО ОСТАНОВКЕ КРОВОТЕЧЕНИЯ")
    slide_toc(prs, [
        "Три первых шага",
        "Прямое давление как старт",
        "Ветвления: нельзя давить / отрыв / не помогло",
        "Сводный алгоритм и контроль",
    ], 2)
    slide_three_start(prs, 3)
    slide_pressure_first(prs, 4)
    slide_when_skip_pressure(prs, 5)
    slide_amputation(prs, 6)
    slide_if_pressure_fails(prs, 7)
    slide_if_pressure_works(prs, 8)
    slide_flowchart(prs, 9)
    slide_checklist(prs, 10)
    slide_summary(prs, 11)
    slide_thanks(prs)

    name = "Последовательность_остановки_кровотечения.pptx"
    path = OUT / name
    prs.save(path)
    verify(path)
    path2 = ROOT / name
    prs.save(path2)
    print("Saved:", path2)
    return path2


if __name__ == "__main__":
    build()
