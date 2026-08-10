#!/usr/bin/env python3
"""Кровотечение. Обзорный осмотр — стиль СУОТ."""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu

from suot_style import (
    ACCENT_RED, ASSETS, BG, BG_PANEL, COURSE, LINE, MUTED, NUM_BG, OUT, ROOT,
    TEAL, TEAL_DARK, TEAL_TEXT, TEXT, WHITE,
    blank, bullets, content_header, new_prs, oval, pic_fit, rect, rich_tbox,
    round_rect, slide_thanks_suot, slide_title_suot, slide_toc_suot, tbox,
    teal_bubble, verify,
)


def slide_definition(prs, num=3):
    slide = blank(prs)
    content_header(slide, "ПОНЯТИЕ «КРОВОТЕЧЕНИЕ»", num)
    # левая серая зона + маскот
    rect(slide, 0, Emu(1300000), Emu(4200000), Emu(5550000), BG_PANEL)
    pic_fit(slide, "suot_mascot.png", Emu(400000), Emu(1600000),
            Emu(2800000), Emu(4800000))
    teal_bubble(
        slide, Emu(3000000), Emu(1800000), Emu(4000000), Emu(2400000),
        "Под кровотечением понимают ситуацию, когда кровь покидает сосудистое русло, "
        "что приводит к кровопотере — безвозвратной утрате части крови.",
        size=12,
    )
    tbox(slide, Emu(7400000), Emu(1600000), Emu(5000000), Emu(450000),
         "Признаки острой кровопотери:", size=15, bold=True, color=TEAL_DARK)
    bullets(slide, Emu(7400000), Emu(2150000), Emu(5000000), Emu(4200000), [
        "резкая общая слабость;",
        "чувство жажды;",
        "головокружение;",
        "мелькание «мушек» перед глазами;",
        "обморок (чаще при попытке встать);",
        "бледная, влажная и холодная кожа;",
        "учащённое сердцебиение;",
        "частое дыхание.",
    ], size=13, marker="☐")
    return slide


def slide_mass_loss(prs, num=4):
    slide = blank(prs)
    content_header(slide, "ОСТРАЯ МАССИВНАЯ КРОВОПОТЕРЯ", num)
    teal_bubble(
        slide, Emu(700000), Emu(1550000), Emu(11000000), Emu(1400000),
        "Наиболее опасно интенсивное кровотечение, приводящее к быстрой потере "
        "большого количества крови, — острая массивная кровопотеря.",
        size=15,
    )
    bullets(slide, Emu(700000), Emu(3300000), Emu(11000000), Emu(3000000), [
        "При повреждении крупных сосудов без остановки гибель может наступить "
        "в течение нескольких минут.",
        "Признаки кровопотери возможны и при остановленном кровотечении, "
        "и при отсутствии видимой крови.",
        "Кровотечения слабой и средней интенсивности тоже нужно останавливать.",
    ], size=15)
    return slide


def slide_types(prs, num=5):
    slide = blank(prs)
    content_header(slide, "ВИДЫ НАРУЖНОГО КРОВОТЕЧЕНИЯ", num)
    cols = [
        ("bleed_arterial.png", "Артериальное",
         "Пульсирующая алая струя; быстро растекающаяся лужа; одежда быстро мокнет."),
        ("bleed_venous.png", "Венозное",
         "Кровь тёмно-вишнёвая, вытекает «ручьём». Остановка обязательна."),
        ("bleed_capillary.png", "Капиллярное",
         "Ссадины, порезы, царапины. Обычно без угрозы жизни."),
    ]
    x = Emu(700000)
    cw = Emu(3600000)
    for img, title, desc in cols:
        round_rect(slide, x, Emu(1500000), cw, Emu(5000000), WHITE, line=TEAL)
        oval(slide, x + Emu(1400000), Emu(1650000), Emu(800000), Emu(400000), TEAL)
        tbox(slide, x + Emu(100000), Emu(2150000), cw - Emu(200000), Emu(450000),
             title, size=14, bold=True, color=TEAL_TEXT, align=PP_ALIGN.CENTER)
        pic_fit(slide, img, x + Emu(200000), Emu(2700000), cw - Emu(400000), Emu(2400000))
        tbox(slide, x + Emu(150000), Emu(5300000), cw - Emu(300000), Emu(1000000),
             desc, size=11, color=TEXT, align=PP_ALIGN.CENTER)
        x += cw + Emu(250000)
    return slide


def slide_intensity(prs, num=6):
    slide = blank(prs)
    content_header(slide, "ОРИЕНТИР — ИНТЕНСИВНОСТЬ", num)
    bullets(slide, Emu(700000), Emu(1600000), Emu(11000000), Emu(2500000), [
        "Смешанные кровотечения часто бывают при отрыве конечности "
        "и опасны из‑за артериального компонента.",
        "Вид кровотечения на месте определить сложно.",
        "Ориентируйтесь на интенсивность и останавливайте любым доступным способом.",
    ], size=15)
    round_rect(slide, Emu(700000), Emu(4500000), Emu(11000000), Emu(1600000), TEAL)
    tbox(slide, Emu(1000000), Emu(4800000), Emu(10400000), Emu(1000000),
         "Приоритет: угрожающее жизни кровотечение останавливают немедленно.",
         size=16, bold=True, color=TEAL_TEXT, anchor=MSO_ANCHOR.MIDDLE)
    return slide


def slide_overview_goal(prs, num=7):
    slide = blank(prs)
    content_header(slide, "ЦЕЛЬ ОБЗОРНОГО ОСМОТРА", num)
    tbox(slide, Emu(700000), Emu(1550000), Emu(11000000), Emu(1000000),
         "Обзорный осмотр пострадавшего (пострадавших) — чтобы найти ранения "
         "с интенсивным наружным кровотечением, требующим немедленной остановки.",
         size=15, color=TEXT)
    tbox(slide, Emu(700000), Emu(2800000), Emu(11000000), Emu(400000),
         "Признаки интенсивного наружного кровотечения:", size=15, bold=True,
         color=TEAL_DARK)
    # три бирюзовые карточки
    cards = [
        "одежда, пропитанная кровью;",
        "скопление крови на земле возле пострадавшего;",
        "видимые раны с интенсивно вытекающей кровью.",
    ]
    x = Emu(700000)
    for c in cards:
        round_rect(slide, x, Emu(3500000), Emu(3600000), Emu(2200000), TEAL)
        tbox(slide, x + Emu(200000), Emu(3900000), Emu(3200000), Emu(1400000),
             c, size=14, bold=True, color=TEAL_TEXT, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE)
        x += Emu(3800000)
    return slide


def slide_overview_how(prs, num=8):
    slide = blank(prs)
    content_header(slide, "ОБЗОРНЫЙ ОСМОТР ПОСТРАДАВШЕГО", num)
    pic_fit(slide, "overview_exam.png", Emu(500000), Emu(1500000),
            Emu(5800000), Emu(5000000))
    teal_bubble(
        slide, Emu(6600000), Emu(2200000), Emu(5600000), Emu(2200000),
        "Обзорный осмотр производится очень быстро, в течение нескольких секунд, "
        "с головы до ног.",
        size=15,
    )
    tbox(slide, Emu(6600000), Emu(4800000), Emu(5600000), Emu(1400000),
         "Сразу после обнаружения признаков угрожающего жизни кровотечения "
         "приступают к его остановке всеми доступными способами.",
         size=13, color=MUTED)
    return slide


def slide_actions(prs, num=9):
    slide = blank(prs)
    content_header(slide, "ДЕЙСТВИЯ ПРИ ОБНАРУЖЕНИИ КРОВОТЕЧЕНИЯ", num)
    steps = [
        "Оценить обстановку и безопасность.",
        "Провести обзорный осмотр — найти интенсивное кровотечение.",
        "Немедленно начать временную остановку кровотечения.",
        "Давление на рану → давящая повязка → жгут (по показаниям).",
        "Вызвать скорую медицинскую помощь.",
        "Контролировать состояние до прибытия СМП.",
    ]
    y = Emu(1500000)
    for i, step in enumerate(steps, 1):
        oval(slide, Emu(700000), y, Emu(480000), Emu(480000), TEAL, line=TEAL_DARK)
        tbox(slide, Emu(700000), y, Emu(480000), Emu(480000), str(i),
             size=14, bold=True, color=TEAL_TEXT, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE)
        round_rect(slide, Emu(1350000), y, Emu(10000000), Emu(480000), BG_PANEL)
        tbox(slide, Emu(1550000), y, Emu(9600000), Emu(480000),
             step, size=14, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)
        y += Emu(720000)
    return slide


def slide_compare(prs, num=10):
    slide = blank(prs)
    content_header(slide, "ОБЗОРНЫЙ И ПОДРОБНЫЙ ОСМОТР", num)
    # left teal header
    round_rect(slide, Emu(700000), Emu(1550000), Emu(5400000), Emu(4800000), WHITE, line=TEAL)
    rect(slide, Emu(700000), Emu(1550000), Emu(5400000), Emu(700000), TEAL)
    tbox(slide, Emu(900000), Emu(1650000), Emu(5000000), Emu(500000),
         "Обзорный осмотр", size=16, bold=True, color=TEAL_TEXT, align=PP_ALIGN.CENTER)
    bullets(slide, Emu(950000), Emu(2500000), Emu(4900000), Emu(3500000), [
        "Цель: найти наружное кровотечение",
        "Время: несколько секунд",
        "Объём: быстро с головы до ног",
        "Когда: сразу после оценки обстановки",
    ], size=14)
    round_rect(slide, Emu(6500000), Emu(1550000), Emu(5400000), Emu(4800000), WHITE, line=LINE)
    rect(slide, Emu(6500000), Emu(1550000), Emu(5400000), Emu(700000), BG_PANEL)
    tbox(slide, Emu(6700000), Emu(1650000), Emu(5000000), Emu(500000),
         "Подробный осмотр", size=16, bold=True, color=TEXT, align=PP_ALIGN.CENTER)
    bullets(slide, Emu(6750000), Emu(2500000), Emu(4900000), Emu(3500000), [
        "Цель: выявить травмы и другие угрозы",
        "Время: несколько минут",
        "Объём: голова → шея → грудь → … → руки",
        "Когда: после остановки кровотечения",
    ], size=14)
    return slide


def slide_summary(prs, num=11):
    slide = blank(prs)
    content_header(slide, "ГЛАВНОЕ ЗАПОМНИТЬ", num)
    points = [
        "Кровотечение — выход крови из сосудов с риском кровопотери.",
        "Обзорный осмотр — быстрый поиск интенсивного наружного кровотечения.",
        "Осмотр — за несколько секунд, с головы до ног.",
        "Сначала останавливаем угрожающее жизни кровотечение, затем — подробный осмотр.",
    ]
    y = Emu(1550000)
    for pt in points:
        round_rect(slide, Emu(700000), y, Emu(11000000), Emu(1000000), BG_PANEL)
        oval(slide, Emu(900000), y + Emu(250000), Emu(500000), Emu(500000), TEAL)
        tbox(slide, Emu(1600000), y + Emu(200000), Emu(9800000), Emu(600000),
             pt, size=14, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)
        y += Emu(1150000)
    return slide


def build():
    prs = new_prs()
    slide_title_suot(
        prs,
        "КРОВОТЕЧЕНИЕ.\nОБЗОРНЫЙ ОСМОТР ПОСТРАДАВШЕГО\n(ПОСТРАДАВШИХ)",
    )
    slide_toc_suot(prs, [
        "Понятие «кровотечение» и признаки кровопотери",
        "Виды наружного кровотечения",
        "Цель и признаки для обзорного осмотра",
        "Как проводится обзорный осмотр",
        "Действия при обнаружении кровотечения",
    ], 2)
    slide_definition(prs, 3)
    slide_mass_loss(prs, 4)
    slide_types(prs, 5)
    slide_intensity(prs, 6)
    slide_overview_goal(prs, 7)
    slide_overview_how(prs, 8)
    slide_actions(prs, 9)
    slide_compare(prs, 10)
    slide_summary(prs, 11)
    slide_thanks_suot(prs)

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
