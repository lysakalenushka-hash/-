#!/usr/bin/env python3
"""
Способы временной остановки наружного кровотечения

Угол: техники остановки — давление, повязка, жгут, импровизация, порядок выбора.
Не дублирует обзорный осмотр и разбор признаков кровопотери (другие презентации).

Стиль: «Организационно-правовые аспекты оказания первой помощи».
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu

from legal_style import (
    ACCENT_RED, BG_BAR, CREAM, LINE, MUTED, NUM_BG, OUT, ROOT, TEXT, WHITE,
    blank, bullets, content_header, cream_note, new_prs, oval, pic_fit, rect,
    rich_tbox, round_rect, slide_thanks, slide_title, slide_toc, tbox, verify,
)


def slide_three_methods(prs, num=3):
    slide = blank(prs)
    content_header(slide, "ТРИ ОСНОВНЫХ СПОСОБА", num)
    tbox(slide, Emu(700000), Emu(1500000), Emu(11000000), Emu(800000),
         "Для временной остановки наружного кровотечения используют прямое давление "
         "на рану, давящую повязку и кровоостанавливающий жгут — по отдельности "
         "или в комбинации.",
         size=15, color=TEXT)
    cards = [
        ("1", "Прямое давление", "Самый простой и приоритетный приём"),
        ("2", "Давящая повязка", "Фиксирует давление, освобождает руки"),
        ("3", "Жгут", "Когда давление и повязка невозможны или неэффективны"),
    ]
    x = Emu(700000)
    for n, title, desc in cards:
        rect(slide, x, Emu(2700000), Emu(3600000), Emu(3200000), WHITE, line=LINE)
        oval(slide, x + Emu(1400000), Emu(2950000), Emu(800000), Emu(800000), NUM_BG)
        tbox(slide, x + Emu(1400000), Emu(2950000), Emu(800000), Emu(800000), n,
             size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE)
        tbox(slide, x + Emu(150000), Emu(4000000), Emu(3300000), Emu(600000),
             title, size=15, bold=True, color=TEXT, align=PP_ALIGN.CENTER)
        tbox(slide, x + Emu(150000), Emu(4700000), Emu(3300000), Emu(900000),
             desc, size=12, color=MUTED, align=PP_ALIGN.CENTER)
        x += Emu(3800000)
    return slide


def slide_direct_pressure(prs, num=4):
    slide = blank(prs)
    content_header(slide, "ПРЯМОЕ ДАВЛЕНИЕ НА РАНУ", num)
    pic_fit(slide, "direct_pressure.png", Emu(500000), Emu(1500000),
            Emu(5500000), Emu(5000000))
    bullets(slide, Emu(6400000), Emu(1600000), Emu(5800000), Emu(3500000), [
        "Рана закрывается салфетками, бинтом или тканью.",
        "Давление рукой — с силой, достаточной для остановки крови.",
        "При отсутствии средств допустимо давление рукой "
        "(лучше в перчатках / через ткань).",
        "Можно рекомендовать пострадавшему самопомощь — "
        "прижать рану самостоятельно.",
    ], size=14)
    cream_note(
        slide, Emu(6400000), Emu(5300000), Emu(5800000), Emu(1200000),
        "Это первый выбор при интенсивном кровотечении.",
        size=13,
    )
    return slide


def slide_bandage(prs, num=5):
    slide = blank(prs)
    content_header(slide, "НАЛОЖЕНИЕ ДАВЯЩЕЙ ПОВЯЗКИ", num)
    pic_fit(slide, "pressure_bandage.png", Emu(500000), Emu(1500000),
            Emu(5500000), Emu(5000000))
    bullets(slide, Emu(6400000), Emu(1550000), Emu(5800000), Emu(4200000), [
        "Задача повязки — остановить кровь, поэтому её накладывают с усилием.",
        "На рану — салфетки / бинт / свёрнутая ткань, затем тугие туры бинта "
        "(с периодическими перекрутами).",
        "Закрепить свободный конец бинта.",
        "Слабо промокает — сверху ещё одну давящую повязку.",
        "Быстро промокает / вторая не помогла — нужен жгут.",
    ], size=13)
    return slide


def slide_foreign_body(prs, num=6):
    slide = blank(prs)
    content_header(slide, "ИНОРОДНОЕ ТЕЛО В РАНЕ", num)
    # use image32 if present else text-only
    from pathlib import Path
    img = "image32.jpeg" if (ROOT / "assets" / "image32.jpeg").exists() else None
    if img:
        pic_fit(slide, img, Emu(6500000), Emu(1600000), Emu(5800000), Emu(4500000))
    bullets(slide, Emu(700000), Emu(1550000), Emu(5500000), Emu(4000000), [
        "НЕ ИЗВЛЕКАТЬ инородный предмет (осколок, стекло и т.п.).",
        "При интенсивном кровотечении: обложить края раны и предмет "
        "бинтами-валиками, затем давящая повязка без давления на предмет.",
        "При отсутствии интенсивного кровотечения — оставить предмет в ране, "
        "ограничить движения, вызвать СМП.",
        "При открытом переломе с отломками — не давить на костные выступы.",
    ], size=13)
    cream_note(
        slide, Emu(700000), Emu(5600000), Emu(11000000), Emu(900000),
        "Правило: зафиксировать — не извлекать.",
        size=14,
    )
    return slide


def slide_tq_when(prs, num=7):
    slide = blank(prs)
    content_header(slide, "ЖГУТ — КОГДА НУЖЕН", num)
    cream_note(
        slide, Emu(700000), Emu(1500000), Emu(11000000), Emu(1400000),
        "Жгут (табельный или импровизированный) — когда прямое давление или "
        "давящая повязка невозможны / неэффективны, а также при отрыве конечности.",
        size=14,
    )
    bullets(slide, Emu(700000), Emu(3200000), Emu(11000000), Emu(3200000), [
        "Только при кровотечении из конечностей.",
        "Выше раны (между раной и сердцем), обычно на 5–7 см от раны.",
        "Если место раны неизвестно — максимально близко к туловищу.",
        "При отрыве — на 5–7 см выше зоны отрыва, без прямого давления на рану.",
        "На обнажённую кожу — нельзя (кроме моделей по инструкции, напр. турникет).",
    ], size=14)
    return slide


def slide_tq_rules(prs, num=8):
    slide = blank(prs)
    content_header(slide, "ПРАВИЛА НАЛОЖЕНИЯ ЖГУТА", num)
    # images
    for i, name in enumerate(["tourniquet_1.png", "tourniquet_2.png", "tourniquet_3.png"]):
        x = Emu(700000) + i * Emu(3800000)
        pic_fit(slide, name, x, Emu(1450000), Emu(3600000), Emu(2800000))
    bullets(slide, Emu(700000), Emu(4500000), Emu(11000000), Emu(2200000), [
        "Кровотечение останавливается первым (растянутым) туром; следующие — "
        "с перекрытием примерно наполовину.",
        "После жгута — давящая повязка на рану; сам жгут должен быть на виду.",
        "Указать точное время наложения (маркер на коже или записка под жгутом).",
        "Конечность иммобилизировать и укутать.",
    ], size=13)
    return slide


def slide_tq_time(prs, num=9):
    slide = blank(prs)
    content_header(slide, "ВРЕМЯ ЖГУТА И ОСЛАБЛЕНИЕ", num)
    round_rect(slide, Emu(700000), Emu(1550000), Emu(11000000), Emu(1400000), CREAM)
    tbox(slide, Emu(1000000), Emu(1750000), Emu(10400000), Emu(1000000),
         "Относительно безопасный срок — до 2 часов "
         "(независимо от температуры окружающей среды). "
         "Снимать жгут вне медорганизации при сроке > 2 ч не рекомендуется.",
         size=15, bold=True, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)
    tbox(slide, Emu(700000), Emu(3200000), Emu(11000000), Emu(450000),
         "Если эвакуация задерживается (> 2 ч) — попытка ослабления через 1–1,5 ч:",
         size=14, bold=True, color=TEXT)
    bullets(slide, Emu(700000), Emu(3800000), Emu(11000000), Emu(2000000), [
        "Осуществить прямое давление на рану.",
        "Ослабить жгут на 15 минут.",
        "Повторно наложить жгут и не снимать до передачи медикам.",
    ], size=14)
    tbox(slide, Emu(700000), Emu(6000000), Emu(11000000), Emu(700000),
         "Внимание! Если кровь возобновилась при давлении на рану — сразу затянуть жгут.",
         size=14, bold=True, color=ACCENT_RED)
    return slide


def slide_improvised(prs, num=10):
    slide = blank(prs)
    content_header(slide, "ИМПРОВИЗИРОВАННЫЙ ЖГУТ", num)
    bullets(slide, Emu(700000), Emu(1550000), Emu(5500000), Emu(4000000), [
        "Подручные средства: тесьма, платок, галстук и подобные вещи.",
        "Делается петля, закручивается прочным предметом "
        "(прут, палка) до остановки или значительного ослабления крови.",
        "Прут фиксируют к конечности бинтом.",
        "Правила те же, что у табельного жгута "
        "(место, время, видимость).",
        "Эффективность и удобство ниже, чем у табельных жгутов.",
    ], size=13)
    pic_fit(slide, "improvised_tq.png", Emu(6500000), Emu(1500000),
            Emu(5800000), Emu(5000000))
    return slide


def slide_priority(prs, num=11):
    slide = blank(prs)
    content_header(slide, "ПОРЯДОК ВЫБОРА СПОСОБА", num)
    steps = [
        "Безопасность → обзорный осмотр (интенсивность кровотечения).",
        "Интенсивное кровотечение → сначала прямое давление на рану.",
        "Давление невозможно / опасно / явно неэффективно "
        "(инородное тело, открытый перелом) → повязка и/или жгут.",
        "Обширное повреждение, разрушение или отрыв конечности → сразу жгут.",
        "Давление помогло → давящая повязка; повязка не держит → жгут выше раны.",
    ]
    y = Emu(1500000)
    for i, step in enumerate(steps, 1):
        oval(slide, Emu(700000), y, Emu(500000), Emu(500000),
             ACCENT_RED if i in (2, 4) else NUM_BG)
        tbox(slide, Emu(700000), y, Emu(500000), Emu(500000), str(i),
             size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE)
        rect(slide, Emu(1400000), y, Emu(10000000), Emu(500000), WHITE, line=LINE)
        tbox(slide, Emu(1600000), y, Emu(9600000), Emu(500000),
             step, size=13, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)
        y += Emu(850000)
    return slide


def slide_errors(prs, num=12):
    slide = blank(prs)
    content_header(slide, "ТИПИЧНЫЕ ОШИБКИ", num)
    bullets(slide, Emu(700000), Emu(1550000), Emu(11000000), Emu(4500000), [
        "Слабое давление на рану / слабая повязка — кровь продолжается.",
        "Жгут без показаний или слишком далеко от раны.",
        "Жгут на голую кожу (когда модель этого не предусматривает).",
        "Жгут скрыт одеждой или повязкой — его не видят.",
        "Нет записи времени наложения.",
        "Извлечение инородного предмета из раны.",
    ], size=15, alert={5})
    cream_note(
        slide, Emu(700000), Emu(5800000), Emu(11000000), Emu(900000),
        "Жгут должен быть виден; время — указано; предмет в ране — не извлекать.",
        size=13,
    )
    return slide


def slide_summary(prs, num=13):
    slide = blank(prs)
    content_header(slide, "ГЛАВНОЕ ЗАПОМНИТЬ", num)
    points = [
        "Приоритет: прямое давление → давящая повязка → жгут по показаниям.",
        "Инородное тело в ране фиксируют валиками — не извлекают.",
        "Жгут: на конечность, выше раны, на подкладку, на виду, с временем.",
        "Срок жгута — до 2 часов; снятие вне медорганизации при просрочке не рекомендуется.",
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
    slide_title(prs, "СПОСОБЫ ВРЕМЕННОЙ ОСТАНОВКИ\nНАРУЖНОГО КРОВОТЕЧЕНИЯ")
    slide_toc(prs, [
        "Три основных способа",
        "Прямое давление и давящая повязка",
        "Инородное тело в ране",
        "Жгут: показания, правила, время",
        "Импровизированный жгут и порядок выбора",
    ], 2)
    slide_three_methods(prs, 3)
    slide_direct_pressure(prs, 4)
    slide_bandage(prs, 5)
    slide_foreign_body(prs, 6)
    slide_tq_when(prs, 7)
    slide_tq_rules(prs, 8)
    slide_tq_time(prs, 9)
    slide_improvised(prs, 10)
    slide_priority(prs, 11)
    slide_errors(prs, 12)
    slide_summary(prs, 13)
    slide_thanks(prs)

    name = "Способы_временной_остановки_кровотечения.pptx"
    path = OUT / name
    prs.save(path)
    verify(path)
    path2 = ROOT / name
    prs.save(path2)
    print("Saved:", path2)
    return path2


if __name__ == "__main__":
    build()
