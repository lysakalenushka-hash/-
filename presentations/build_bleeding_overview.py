#!/usr/bin/env python3
"""
1) Кровотечение. Обзорный осмотр пострадавшего (пострадавших)

Угол: процесс оказания помощи — зачем и как искать кровотечение,
что делать при обнаружении. Без углублённой клиники признаков кровопотери
и детальной классификации видов (это во 2-й презентации).

Стиль: «Организационно-правовые аспекты оказания первой помощи».
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu

from legal_style import (
    ACCENT_RED, BG, BG_BAR, COURSE, CREAM, LINE, MUTED, NUM_BG, OUT, ROOT,
    TEXT, WHITE, blank, bullets, content_header, cream_note, new_prs, oval,
    pic_fit, rect, rich_tbox, round_rect, slide_thanks, slide_title, slide_toc,
    tbox, verify,
)


def slide_why(prs, num=3):
    """Зачем вообще смотреть на кровотечение — процессный угол."""
    slide = blank(prs)
    content_header(slide, "КРОВОТЕЧЕНИЕ В АЛГОРИТМЕ ПЕРВОЙ ПОМОЩИ", num)
    tbox(slide, Emu(700000), Emu(1550000), Emu(11000000), Emu(900000),
         "После оценки обстановки и безопасности участник оказания первой помощи "
         "сначала ищет угрожающее жизни наружное кровотечение — и только потом "
         "переходит к подробному осмотру и другим мероприятиям.",
         size=15, color=TEXT)
    steps = [
        ("1", "Безопасность", "Оценить обстановку, защитить себя"),
        ("2", "Обзорный осмотр", "Быстро найти интенсивное кровотечение"),
        ("3", "Остановка крови", "Начать временную остановку сразу"),
        ("4", "Далее по алгоритму", "Признаки жизни, СМП, подробный осмотр"),
    ]
    x = Emu(700000)
    for n, title, desc in steps:
        round_rect(slide, x, Emu(2800000), Emu(2700000), Emu(3200000), WHITE, line=LINE)
        oval(slide, x + Emu(950000), Emu(3000000), Emu(800000), Emu(800000), NUM_BG)
        tbox(slide, x + Emu(950000), Emu(3000000), Emu(800000), Emu(800000), n,
             size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE)
        tbox(slide, x + Emu(150000), Emu(4000000), Emu(2400000), Emu(600000),
             title, size=14, bold=True, color=TEXT, align=PP_ALIGN.CENTER)
        tbox(slide, x + Emu(150000), Emu(4700000), Emu(2400000), Emu(1000000),
             desc, size=12, color=MUTED, align=PP_ALIGN.CENTER)
        x += Emu(2900000)
    return slide


def slide_goal(prs, num=4):
    slide = blank(prs)
    content_header(slide, "ЦЕЛЬ ОБЗОРНОГО ОСМОТРА", num)
    cream_note(
        slide, Emu(700000), Emu(1550000), Emu(11000000), Emu(1600000),
        "Обзорный осмотр пострадавшего (пострадавших) нужен, чтобы определить "
        "наличие и расположение ранений с интенсивным наружным кровотечением, "
        "требующим немедленной остановки.",
        size=15,
    )
    tbox(slide, Emu(700000), Emu(3500000), Emu(11000000), Emu(450000),
         "Сигналы «остановить кровь сейчас»:", size=15, bold=True, color=TEXT)
    cards = [
        "одежда, пропитанная кровью;",
        "лужа / скопление крови возле пострадавшего;",
        "рана с интенсивно вытекающей кровью.",
    ]
    x = Emu(700000)
    for c in cards:
        rect(slide, x, Emu(4200000), Emu(3600000), Emu(1800000), WHITE, line=LINE)
        rect(slide, x, Emu(4200000), Emu(90000), Emu(1800000), ACCENT_RED)
        tbox(slide, x + Emu(250000), Emu(4600000), Emu(3200000), Emu(1000000),
             c, size=14, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)
        x += Emu(3800000)
    return slide


def slide_how(prs, num=5):
    slide = blank(prs)
    content_header(slide, "КАК ПРОВОДИТСЯ ОБЗОРНЫЙ ОСМОТР", num)
    pic_fit(slide, "overview_exam.png", Emu(500000), Emu(1500000),
            Emu(6000000), Emu(5200000))
    rich_tbox(
        slide, Emu(6800000), Emu(2200000), Emu(5600000), Emu(2000000),
        [
            ("Обзорный осмотр", True),
            (" производится очень быстро, в течение ", False),
            ("нескольких секунд", True),
            (", с головы до ног.", False),
        ],
        size=18, color=TEXT,
    )
    cream_note(
        slide, Emu(6800000), Emu(4600000), Emu(5600000), Emu(1800000),
        "Как только найдены признаки угрожающего жизни кровотечения — "
        "сразу переходите к его остановке всеми доступными способами.",
        size=13,
    )
    return slide


def slide_multiple(prs, num=6):
    """Несколько пострадавших — угол обзора, не клиника."""
    slide = blank(prs)
    content_header(slide, "НЕСКОЛЬКО ПОСТРАДАВШИХ", num)
    bullets(slide, Emu(700000), Emu(1600000), Emu(11000000), Emu(3500000), [
        "Сначала быстро «пройдите» всех глазами: у кого кровь течёт сильнее.",
        "Приоритет — тот, у кого интенсивное наружное кровотечение.",
        "По возможности дайте указание на самопомощь "
        "(прямое давление на рану), пока помогаете другому.",
        "Не застревайте на подробном осмотре одного, пока не закрыли "
        "угрожающие жизни кровотечения у остальных.",
    ], size=15)
    cream_note(
        slide, Emu(700000), Emu(5400000), Emu(11000000), Emu(1000000),
        "Обзорный осмотр — инструмент сортировки: сначала жизнь, потом детали.",
        size=14,
    )
    return slide


def slide_actions(prs, num=7):
    slide = blank(prs)
    content_header(slide, "ДЕЙСТВИЯ ПРИ ОБНАРУЖЕНИИ КРОВОТЕЧЕНИЯ", num)
    steps = [
        "Убедиться в безопасности для себя и пострадавшего.",
        "Найти источник интенсивного наружного кровотечения (обзорный осмотр).",
        "Начать временную остановку: давление на рану → повязка → жгут по показаниям.",
        "Вызвать скорую медицинскую помощь (если ещё не вызвана).",
        "Контролировать состояние до прибытия бригады СМП.",
    ]
    y = Emu(1550000)
    for i, step in enumerate(steps, 1):
        oval(slide, Emu(700000), y, Emu(500000), Emu(500000),
             ACCENT_RED if i <= 3 else NUM_BG)
        tbox(slide, Emu(700000), y, Emu(500000), Emu(500000), str(i),
             size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE)
        rect(slide, Emu(1400000), y, Emu(10000000), Emu(500000), WHITE, line=LINE)
        tbox(slide, Emu(1600000), y, Emu(9600000), Emu(500000),
             step, size=14, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)
        y += Emu(850000)
    return slide


def slide_compare(prs, num=8):
    slide = blank(prs)
    content_header(slide, "ОБЗОРНЫЙ И ПОДРОБНЫЙ ОСМОТР", num)
    rect(slide, Emu(700000), Emu(1550000), Emu(5400000), Emu(4800000), WHITE, line=LINE)
    rect(slide, Emu(700000), Emu(1550000), Emu(5400000), Emu(700000), NUM_BG)
    tbox(slide, Emu(900000), Emu(1650000), Emu(5000000), Emu(500000),
         "Обзорный осмотр", size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    bullets(slide, Emu(950000), Emu(2500000), Emu(4900000), Emu(3500000), [
        "Цель: найти интенсивное кровотечение",
        "Время: несколько секунд",
        "Объём: быстро с головы до ног",
        "Когда: сразу после оценки обстановки",
    ], size=14)
    rect(slide, Emu(6500000), Emu(1550000), Emu(5400000), Emu(4800000), WHITE, line=LINE)
    rect(slide, Emu(6500000), Emu(1550000), Emu(5400000), Emu(700000), BG_BAR)
    tbox(slide, Emu(6700000), Emu(1650000), Emu(5000000), Emu(500000),
         "Подробный осмотр", size=16, bold=True, color=TEXT, align=PP_ALIGN.CENTER)
    bullets(slide, Emu(6750000), Emu(2500000), Emu(4900000), Emu(3500000), [
        "Цель: выявить травмы и другие угрозы",
        "Время: несколько минут",
        "Объём: голова → шея → грудь → … → руки",
        "Когда: после остановки кровотечения",
    ], size=14)
    return slide


def slide_summary(prs, num=9):
    slide = blank(prs)
    content_header(slide, "ГЛАВНОЕ ЗАПОМНИТЬ", num)
    points = [
        "В алгоритме первой помощи угрожающее жизни кровотечение ищут сразу.",
        "Обзорный осмотр — за несколько секунд, с головы до ног.",
        "Нашли интенсивное истечение крови — сразу останавливайте.",
        "Подробный осмотр — только после остановки угрожающего жизни кровотечения.",
    ]
    y = Emu(1550000)
    for pt in points:
        rect(slide, Emu(700000), y, Emu(11000000), Emu(1000000), WHITE, line=LINE)
        rect(slide, Emu(700000), y, Emu(90000), Emu(1000000), ACCENT_RED)
        tbox(slide, Emu(1000000), y + Emu(200000), Emu(10400000), Emu(600000),
             pt, size=14, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)
        y += Emu(1150000)
    return slide


def build():
    prs = new_prs()
    slide_title(prs, "КРОВОТЕЧЕНИЕ.\nОБЗОРНЫЙ ОСМОТР ПОСТРАДАВШЕГО\n(ПОСТРАДАВШИХ)")
    slide_toc(prs, [
        "Кровотечение в алгоритме первой помощи",
        "Цель обзорного осмотра и сигналы к действию",
        "Как проводится обзорный осмотр",
        "Несколько пострадавших и порядок действий",
        "Обзорный и подробный осмотр",
    ], 2)
    slide_why(prs, 3)
    slide_goal(prs, 4)
    slide_how(prs, 5)
    slide_multiple(prs, 6)
    slide_actions(prs, 7)
    slide_compare(prs, 8)
    slide_summary(prs, 9)
    slide_thanks(prs)

    name = "Кровотечение_Обзорный_осмотр_пострадавшего.pptx"
    path = OUT / name
    prs.save(path)
    verify(path)
    path2 = ROOT / name
    prs.save(path2)
    print("Saved:", path2)
    return path2


if __name__ == "__main__":
    build()
