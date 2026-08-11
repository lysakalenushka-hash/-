#!/usr/bin/env python3
"""
Остановка кровотечения при ранениях различных областей тела

Угол: что меняется в зависимости от зоны ранения.
Не дублирует общий алгоритм последовательности и подробные техники
наложения давления / повязки / жгута (другие презентации).

Стиль: «Организационно-правовые аспекты оказания первой помощи».
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu

from legal_style import (
    ACCENT_RED, LINE, MUTED, NUM_BG, OUT, ROOT, ROW_ALT, TEXT, WHITE,
    blank, bullets, content_header, cream_note, new_prs, oval, pic_fit, rect,
    slide_thanks, slide_title, slide_toc, tbox, verify,
)


def slide_map(prs, num=3):
    slide = blank(prs)
    content_header(slide, "ЗОНЫ, КОТОРЫЕ РАЗБИРАЕМ", num)
    zones = [
        ("1", "Голова", "Давление, повязка, осторожность при черепе"),
        ("2", "Шея", "Давление + повязка через противоположную подмышку"),
        ("3", "Грудь и спина", "Только поверхностные сосуды"),
        ("4", "Живот и таз", "Не давить на выпавшие органы"),
        ("5", "Конечности", "Все способы; при отрыве — жгут"),
        ("6", "Смежные зоны", "Предпочтительно прямое давление"),
    ]
    positions = [
        (Emu(700000), Emu(1500000)),
        (Emu(4600000), Emu(1500000)),
        (Emu(8500000), Emu(1500000)),
        (Emu(700000), Emu(4200000)),
        (Emu(4600000), Emu(4200000)),
        (Emu(8500000), Emu(4200000)),
    ]
    for (n, title, desc), (x, y) in zip(zones, positions):
        rect(slide, x, y, Emu(3400000), Emu(2200000), WHITE, line=LINE)
        oval(slide, x + Emu(150000), y + Emu(250000), Emu(550000), Emu(550000), NUM_BG)
        tbox(slide, x + Emu(150000), y + Emu(250000), Emu(550000), Emu(550000), n,
             size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE)
        tbox(slide, x + Emu(850000), y + Emu(300000), Emu(2350000), Emu(450000),
             title, size=15, bold=True, color=TEXT)
        tbox(slide, x + Emu(200000), y + Emu(1000000), Emu(3000000), Emu(900000),
             desc, size=12, color=MUTED)
    return slide


def slide_head(prs, num=4):
    slide = blank(prs)
    content_header(slide, "РАНЕНИЕ ГОЛОВЫ", num)
    pic_fit(slide, "head_bandage.jpeg", Emu(500000), Emu(1500000),
            Emu(5000000), Emu(4800000))
    bullets(slide, Emu(5800000), Emu(1550000), Emu(6400000), Emu(3200000), [
        "Ранения головы (особенно волосистой части) часто дают сильное кровотечение.",
        "Кровотечение из скальпа часто не останавливается само — действовать быстро.",
        "Обычный порядок: прямое давление на рану, при необходимости — повязка.",
        "Инородный предмет в ране — зафиксировать валиками, не извлекать.",
    ], size=13)
    cream_note(
        slide, Emu(5800000), Emu(5000000), Emu(6400000), Emu(1400000),
        "При возможном повреждении костей черепа — не давить на рану и не "
        "усиливать давление тампоном: только циркулярная давящая повязка.",
        size=12,
    )
    return slide


def slide_nose(prs, num=5):
    slide = blank(prs)
    content_header(slide, "НОСОВОЕ КРОВОТЕЧЕНИЕ", num)
    cards = [
        ("В сознании",
         "Усадить со слегка наклонённой вперёд головой. "
         "Зажать крылья носа на 15–20 минут. Холод на переносицу."),
        ("Без сознания",
         "Устойчивое боковое положение, контроль дыхательных путей, "
         "вызвать скорую медицинскую помощь."),
        ("Если не остановилось",
         "Вызвать СМП; до приезда продолжать те же мероприятия. "
         "Самостоятельное вправление переломов носа недопустимо."),
    ]
    y = Emu(1550000)
    for title, desc in cards:
        rect(slide, Emu(700000), y, Emu(11000000), Emu(1400000), WHITE, line=LINE)
        rect(slide, Emu(700000), y, Emu(90000), Emu(1400000), ACCENT_RED)
        tbox(slide, Emu(1000000), y + Emu(200000), Emu(10400000), Emu(400000),
             title, size=15, bold=True, color=TEXT)
        tbox(slide, Emu(1000000), y + Emu(650000), Emu(10400000), Emu(550000),
             desc, size=13, color=MUTED)
        y += Emu(1600000)
    return slide


def slide_neck(prs, num=6):
    slide = blank(prs)
    content_header(slide, "РАНЕНИЕ ШЕИ", num)
    pic_fit(slide, "neck_pressure.jpeg", Emu(500000), Emu(1500000),
            Emu(4200000), Emu(4800000))
    pic_fit(slide, "neck_bandage_axilla.jpeg", Emu(4800000), Emu(1500000),
            Emu(3200000), Emu(4800000))
    bullets(slide, Emu(8300000), Emu(1550000), Emu(4000000), Emu(4000000), [
        "Повреждение крупных сосудов шеи (в т.ч. сонных артерий) — угроза жизни.",
        "Сразу после обнаружения — остановка крови.",
        "Быстрый способ: прямое давление на рану.",
        "Затем давящая повязка через противоположную подмышку.",
        "Можно использовать жгут достаточного размера как материал повязки.",
    ], size=12)
    return slide


def slide_chest(prs, num=7):
    slide = blank(prs)
    content_header(slide, "ГРУДЬ И СПИНА", num)
    pic_fit(slide, "multi_pressure.jpeg", Emu(500000), Emu(1500000),
            Emu(5500000), Emu(4800000))
    bullets(slide, Emu(6400000), Emu(1600000), Emu(5800000), Emu(3200000), [
        "В грудной полости — сердце, лёгкие, крупные сосуды.",
        "Кровотечение из внутренних крупных сосудов на этапе первой помощи "
        "остановить нельзя — нужна хирургия.",
        "Поверхностные сосуды: прямое давление на рану или давящая повязка.",
    ], size=14)
    cream_note(
        slide, Emu(6400000), Emu(5100000), Emu(5800000), Emu(1300000),
        "Задача первой помощи при груди — поверхностное кровотечение и "
        "быстрый вызов медпомощи при тяжёлой травме.",
        size=12,
    )
    return slide


def slide_abdomen(prs, num=8):
    slide = blank(prs)
    content_header(slide, "ЖИВОТ И ТАЗ", num)
    items = [
        ("Закрытая травма",
         "Может оставаться незамеченной, пока внутреннее кровотечение "
         "не вызовет резкого ухудшения состояния."),
        ("Открытая травма",
         "Может сопровождаться выпадением внутренних органов и кровотечением."),
        ("Остановка крови",
         "Прямое давление на рану и/или наложение давящей повязки."),
        ("Важно",
         "При выпадении внутренних органов накладывать на них "
         "давящую повязку не рекомендуется."),
    ]
    y = Emu(1500000)
    for i, (title, desc) in enumerate(items):
        fill_line = ACCENT_RED if i == 3 else NUM_BG
        rect(slide, Emu(700000), y, Emu(11000000), Emu(1100000), WHITE, line=LINE)
        rect(slide, Emu(700000), y, Emu(90000), Emu(1100000), fill_line)
        tbox(slide, Emu(1000000), y + Emu(150000), Emu(10400000), Emu(350000),
             title, size=15, bold=True, color=TEXT)
        tbox(slide, Emu(1000000), y + Emu(550000), Emu(10400000), Emu(450000),
             desc, size=13, color=MUTED)
        y += Emu(1250000)
    return slide


def slide_limbs(prs, num=9):
    slide = blank(prs)
    content_header(slide, "РАНЕНИЯ КОНЕЧНОСТЕЙ", num)
    tbox(slide, Emu(700000), Emu(1450000), Emu(11000000), Emu(600000),
         "Применяются все способы: давление → повязка → жгут — "
         "по общей последовательности.",
         size=14, color=TEXT)
    factors = [
        ("Вид / интенсивность", "Артериальное, венозное, капиллярное"),
        ("Место ранения", "Где расположена рана на конечности"),
        ("Срок прибытия СМП", "Близко — достаточно простых способов"),
        ("Оснащение", "Табельный жгут или подручные средства"),
        ("Состояние кровотечения", "Остановилось или продолжается"),
    ]
    y = Emu(2200000)
    for title, desc in factors:
        rect(slide, Emu(700000), y, Emu(11000000), Emu(700000), WHITE, line=LINE)
        tbox(slide, Emu(900000), y, Emu(4500000), Emu(700000), title,
             size=13, bold=True, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)
        tbox(slide, Emu(5600000), y, Emu(5900000), Emu(700000), desc,
             size=13, color=MUTED, anchor=MSO_ANCHOR.MIDDLE)
        y += Emu(800000)
    return slide


def slide_amputation_region(prs, num=10):
    slide = blank(prs)
    content_header(slide, "ОТРЫВ ЧАСТИ КОНЕЧНОСТИ", num)
    pic_fit(slide, "tourniquet_2.png", Emu(500000), Emu(1500000),
            Emu(5200000), Emu(4800000))
    bullets(slide, Emu(6000000), Emu(1600000), Emu(6200000), Emu(3500000), [
        "Одно из наиболее тяжёлых повреждений конечностей.",
        "Сначала — остановить кровотечение.",
        "Отрыв крупных частей (кисть или стопа и выше) — "
        "накладывать кровоостанавливающий жгут.",
        "В остальных случаях допустимы прямое давление или давящая повязка.",
    ], size=14, alert={2})
    cream_note(
        slide, Emu(6000000), Emu(5300000), Emu(6200000), Emu(1200000),
        "Выбор способа зависит от уровня ампутации и интенсивности кровотечения.",
        size=12,
    )
    return slide


def slide_adjacent(prs, num=11):
    slide = blank(prs)
    content_header(slide, "СМЕЖНЫЕ ЗОНЫ", num)
    tbox(slide, Emu(700000), Emu(1500000), Emu(11000000), Emu(800000),
         "Смежные зоны — места сочленения конечностей и шеи с туловищем "
         "(подмышки, пах, основание шеи).",
         size=15, color=TEXT)
    cards = [
        ("Почему опасно",
         "Там проходят крупные сосуды — ранения часто дают сильное кровотечение."),
        ("Почему сложно",
         "Давящую повязку или жгут в этих местах трудно наложить и зафиксировать."),
        ("Что предпочесть",
         "Для остановки кровотечения предпочтительно прямое давление на рану."),
    ]
    y = Emu(2500000)
    for title, desc in cards:
        rect(slide, Emu(700000), y, Emu(11000000), Emu(1100000), WHITE, line=LINE)
        rect(slide, Emu(700000), y, Emu(90000), Emu(1100000), ACCENT_RED)
        tbox(slide, Emu(1000000), y + Emu(150000), Emu(10400000), Emu(350000),
             title, size=15, bold=True, color=TEXT)
        tbox(slide, Emu(1000000), y + Emu(550000), Emu(10400000), Emu(400000),
             desc, size=13, color=MUTED)
        y += Emu(1250000)
    return slide


def slide_table(prs, num=12):
    slide = blank(prs)
    content_header(slide, "СВОДКА ПО ОБЛАСТЯМ", num)
    rows = [
        ("Область", "Ключевой приём", "Особая оговорка"),
        ("Голова", "Давление / повязка", "При риске перелома черепа — без усиленного давления"),
        ("Нос", "Зажим крыльев 15–20 мин", "Без сознания — боковое положение"),
        ("Шея", "Давление → повязка", "Повязка через противоположную подмышку"),
        ("Грудь / спина", "Давление / повязка", "Внутренние крупные сосуды — только СМП"),
        ("Живот / таз", "Давление / повязка", "Не давить на выпавшие органы"),
        ("Конечности", "Все способы", "Отрыв крупных частей — жгут"),
        ("Смежные зоны", "Прямое давление", "Повязка и жгут трудно зафиксировать"),
    ]
    y = Emu(1400000)
    for i, (a, b, c) in enumerate(rows):
        fill = NUM_BG if i == 0 else (ROW_ALT if i % 2 == 0 else WHITE)
        tc = WHITE if i == 0 else TEXT
        rect(slide, Emu(700000), y, Emu(2800000), Emu(580000), fill, line=LINE)
        rect(slide, Emu(3500000), y, Emu(3500000), Emu(580000), fill, line=LINE)
        rect(slide, Emu(7000000), y, Emu(5000000), Emu(580000), fill, line=LINE)
        tbox(slide, Emu(850000), y, Emu(2500000), Emu(580000), a,
             size=11, bold=True, color=tc, anchor=MSO_ANCHOR.MIDDLE)
        tbox(slide, Emu(3650000), y, Emu(3200000), Emu(580000), b,
             size=11, bold=(i == 0), color=tc, anchor=MSO_ANCHOR.MIDDLE)
        tbox(slide, Emu(7150000), y, Emu(4700000), Emu(580000), c,
             size=11, bold=(i == 0), color=tc, anchor=MSO_ANCHOR.MIDDLE)
        y += Emu(580000)
    return slide


def slide_summary(prs, num=13):
    slide = blank(prs)
    content_header(slide, "ГЛАВНОЕ ЗАПОМНИТЬ", num)
    points = [
        "Голова: быстро останавливать; при риске перелома черепа — без усиленного давления.",
        "Шея: давление, затем повязка через противоположную подмышку.",
        "Грудь: останавливаем только поверхностное кровотечение.",
        "Живот: не накладывать давящую повязку на выпавшие органы.",
        "Конечности: полный набор способов; крупные отрывы — жгут.",
        "Смежные зоны: предпочтительно прямое давление на рану.",
    ]
    y = Emu(1450000)
    for pt in points:
        rect(slide, Emu(700000), y, Emu(11000000), Emu(750000), WHITE, line=LINE)
        rect(slide, Emu(700000), y, Emu(90000), Emu(750000), ACCENT_RED)
        tbox(slide, Emu(1000000), y + Emu(120000), Emu(10400000), Emu(500000),
             pt, size=13, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)
        y += Emu(850000)
    return slide


def build():
    prs = new_prs()
    slide_title(prs, "ОСТАНОВКА КРОВОТЕЧЕНИЯ ПРИ РАНЕНИЯХ\nРАЗЛИЧНЫХ ОБЛАСТЕЙ ТЕЛА")
    slide_toc(prs, [
        "Карта зон",
        "Голова и нос",
        "Шея, грудь, живот",
        "Конечности, смежные зоны, сводка",
    ], 2)
    slide_map(prs, 3)
    slide_head(prs, 4)
    slide_nose(prs, 5)
    slide_neck(prs, 6)
    slide_chest(prs, 7)
    slide_abdomen(prs, 8)
    slide_limbs(prs, 9)
    slide_amputation_region(prs, 10)
    slide_adjacent(prs, 11)
    slide_table(prs, 12)
    slide_summary(prs, 13)
    slide_thanks(prs)

    name = "Остановка_кровотечения_области_тела.pptx"
    path = OUT / name
    prs.save(path)
    verify(path)
    path2 = ROOT / name
    prs.save(path2)
    print("Saved:", path2)
    return path2


if __name__ == "__main__":
    build()
