#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Переработка «Общие положения ПБ на ОПО.pptx» в красный стиль (как Б.7.5 / котлы)."""

from __future__ import annotations

import re
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor as PptxRGB
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Pt as PptxPt

SOURCE = Path("Общие_положения_ПБ_на_ОПО_исходник.pptx")
TEMPLATE = Path("/tmp/red_style_sample.pptx")
OUTPUT = Path("Общие_положения_ПБ_на_ОПО.pptx")

RED = PptxRGB(0xE3, 0x06, 0x13)
GRAY = PptxRGB(0x6B, 0x72, 0x80)
DARK = PptxRGB(0x1A, 0x1A, 0x1A)
WHITE = PptxRGB(0xFF, 0xFF, 0xFF)
CREAM = PptxRGB(0xFE, 0xF3, 0xC7)
AMBER = PptxRGB(0xF5, 0x9E, 0x0B)
LIGHT_BG = PptxRGB(0xFA, 0xFA, 0xFA)

SLIDE_W = Emu(12192000)
SLIDE_H = Emu(6858000)
MARGIN_L = Emu(640080)
CONTENT_W = Emu(10911535)
FONT = "Inter"

# Текст слайда 16 был только в изображении — восстанавливаем по 116-ФЗ
SLIDE_16_ITEMS = [
    "Немедленно информировать федеральный орган исполнительной власти в области "
    "промышленной безопасности и иные органы о каждой аварии на ОПО.",
    "Принимать меры по локализации и ликвидации последствий аварии.",
    "Оказывать содействие государственным органам в расследовании причин аварии.",
    "Принимать меры по устранению причин аварии и профилактике подобных аварий.",
    "Не допускать сокрытия информации об авариях и инцидентах на ОПО.",
]

SLIDE_6_ITEMS = [
    "Проектирование, изготовление, монтаж, наладка, обслуживание, ремонт, "
    "консервация и ликвидация технических устройств, применяемых на ОПО.",
    "Эксплуатация опасных производственных объектов.",
    "Проведение экспертизы промышленной безопасности.",
    "Подготовка и аттестация работников в области промышленной безопасности.",
    "Разработка декларации промышленной безопасности и обоснования безопасности ОПО.",
]


def set_run_font(run, size=14, bold=False, color=DARK, name=FONT):
    run.font.name = name
    run.font.size = PptxPt(size)
    run.font.bold = bold
    run.font.color.rgb = color


def add_textbox(slide, left, top, width, height, text="", size=14, bold=False, color=DARK, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    set_run_font(run, size=size, bold=bold, color=color)
    return box


def add_bg(slide, color=LIGHT_BG):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    sp = shape._element
    slide.shapes._spTree.remove(sp)
    slide.shapes._spTree.insert(2, sp)


def add_accent_bar(slide, top=Emu(6355080)):
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, MARGIN_L, top, CONTENT_W, Emu(91440))
    bar.fill.solid()
    bar.fill.fore_color.rgb = RED
    bar.line.fill.background()


def add_page_num(slide, num, total):
    add_textbox(
        slide,
        Emu(10180015),
        Emu(6492240),
        Emu(1188720),
        Emu(228600),
        f"{num:02d} / {total:02d}",
        size=11,
        color=GRAY,
        align=PP_ALIGN.RIGHT,
    )


def add_title_slide(prs, title, subtitle, tags, page, total):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide)
    oval = slide.shapes.add_shape(MSO_SHAPE.OVAL, Emu(9753600), Emu(-457200), Emu(2743200), Emu(2743200))
    oval.fill.solid()
    oval.fill.fore_color.rgb = CREAM
    oval.line.fill.background()
    add_textbox(slide, MARGIN_L, Emu(731520), Emu(4572000), Emu(228600), "КУРС", size=12, bold=True, color=RED)
    add_textbox(slide, MARGIN_L, Emu(1188720), CONTENT_W, Emu(640080), title, size=28, bold=True, color=DARK)
    add_textbox(slide, MARGIN_L, Emu(2103120), CONTENT_W, Emu(914400), subtitle, size=18, color=GRAY)
    add_textbox(slide, MARGIN_L, Emu(3200400), CONTENT_W, Emu(457200), "116-ФЗ · Промышленная безопасность · ОПО", size=14, color=GRAY)
    add_accent_bar(slide)
    add_textbox(slide, MARGIN_L, Emu(5486400), CONTENT_W, Emu(365760), tags, size=12, color=GRAY)
    add_page_num(slide, page, total)


def add_items_slide(prs, title, items, page, total, subtitle=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide)
    add_textbox(slide, MARGIN_L, Emu(731520), CONTENT_W, Emu(640080), title, size=24, bold=True, color=DARK)
    y = Emu(1463040) if subtitle else Emu(1650720)
    if subtitle:
        add_textbox(slide, MARGIN_L, Emu(1463040), CONTENT_W, Emu(365760), subtitle, size=14, color=GRAY)
        y = Emu(1920240)
    max_items = 6
    shown = items[:max_items]
    step = Emu(780000) if len(shown) > 4 else Emu(960000)
    for i, item in enumerate(shown, 1):
        add_textbox(slide, MARGIN_L, y, Emu(640080), Emu(365760), f"{i:02d}", size=14, bold=True, color=RED)
        add_textbox(slide, Emu(1371600), y, Emu(9539980), Emu(720000), item, size=13, color=GRAY)
        y += step
    if len(items) > max_items:
        add_textbox(
            slide,
            MARGIN_L,
            Emu(6120000),
            CONTENT_W,
            Emu(228600),
            f"(+ ещё {len(items) - max_items} положений на исходном слайде)",
            size=11,
            color=GRAY,
        )
    add_accent_bar(slide)
    add_page_num(slide, page, total)


def add_detail_slide(prs, title, section, body, note, page, total, subtitle=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide)
    add_textbox(slide, MARGIN_L, Emu(731520), CONTENT_W, Emu(640080), title, size=24, bold=True, color=DARK)
    top = Emu(1463040)
    if subtitle:
        add_textbox(slide, MARGIN_L, top, CONTENT_W, Emu(365760), subtitle, size=16, color=GRAY)
        top = Emu(1920240)
    add_textbox(slide, MARGIN_L, top, Emu(3200400), Emu(365760), section, size=14, bold=True, color=RED)
    add_textbox(slide, MARGIN_L, top + Emu(457200), CONTENT_W, Emu(2103120), body, size=13, color=GRAY)
    if note:
        add_textbox(slide, MARGIN_L, Emu(4754880), CONTENT_W, Emu(914400), note, size=12, color=AMBER)
    add_accent_bar(slide)
    add_page_num(slide, page, total)


def add_summary_slide(prs, items, page, total, closing=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide)
    add_textbox(slide, MARGIN_L, Emu(731520), Emu(4572000), Emu(228600), "Итоги", size=14, bold=True, color=RED)
    add_textbox(
        slide,
        MARGIN_L,
        Emu(1005840),
        CONTENT_W,
        Emu(640080),
        "Ключевые выводы",
        size=22,
        bold=True,
        color=DARK,
    )
    y = Emu(2103120)
    for title, body in items:
        add_textbox(slide, MARGIN_L, y, Emu(3200400), Emu(365760), title, size=13, bold=True, color=RED)
        add_textbox(slide, Emu(3840480), y, Emu(7711135), Emu(640080), body, size=12, color=GRAY)
        y += Emu(960000)
    if closing:
        add_textbox(slide, MARGIN_L, Emu(5486400), CONTENT_W, Emu(548640), closing, size=12, color=GRAY)
    add_accent_bar(slide)
    add_page_num(slide, page, total)


def extract_blocks(slide):
    blocks = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            text = shape.text_frame.text.strip()
            if text:
                blocks.append(re.sub(r"\x0b", " ", text))
    return blocks


def is_number(s: str) -> bool:
    return s.strip().isdigit()


def parse_pairs(blocks):
    """Парсинг блоков вида «заголовок — описание» или «01 / заголовок / текст»."""
    items = []
    i = 0
    while i < len(blocks):
        b = blocks[i]
        if is_number(b) and i + 2 < len(blocks) and not is_number(blocks[i + 1]):
            label = blocks[i + 1]
            desc = blocks[i + 2]
            items.append(f"{label}: {desc}")
            i += 3
        elif is_number(b) and i + 1 < len(blocks):
            items.append(f"{blocks[i + 1]}")
            i += 2
        elif i + 1 < len(blocks) and len(b) < 55 and len(blocks[i + 1]) > len(b) + 10:
            items.append(f"{b} — {blocks[i + 1]}")
            i += 2
        else:
            for line in b.split("\n"):
                line = line.strip()
                if line and line not in items:
                    items.append(line)
            i += 1
    return items


def slide_to_spec(blocks, slide_num: int) -> dict:
    if not blocks:
        return {"type": "items", "title": f"Слайд {slide_num}", "items": []}

    if slide_num == 1:
        return {
            "type": "title",
            "title": blocks[0],
            "subtitle": blocks[1] if len(blocks) > 1 else "",
            "tags": "116-ФЗ · ОПО · Промышленная безопасность · Аттестация",
        }

    if blocks[0] == "Ключевые выводы":
        pairs = []
        rest = blocks[1:]
        i = 0
        while i < len(rest):
            if is_number(rest[i]) and i + 2 < len(rest):
                pairs.append((rest[i + 1], rest[i + 2]))
                i += 3
            else:
                i += 1
        return {
            "type": "summary",
            "items": pairs,
            "closing": "Соблюдение требований ФЗ № 116-ФЗ — основа безопасной эксплуатации ОПО.",
        }

    title = blocks[0]
    rest = blocks[1:]

    if slide_num == 6:
        return {"type": "items", "title": title, "items": SLIDE_6_ITEMS, "subtitle": rest[0] if rest else None}

    if slide_num == 16:
        return {"type": "items", "title": title, "items": SLIDE_16_ITEMS}

    if slide_num == 28 and len(rest) == 1:
        return {
            "type": "detail",
            "title": title,
            "section": "Сущность технического перевооружения",
            "body": rest[0],
            "note": "При техническом перевооружении может потребоваться обновление обоснования безопасности ОПО "
            "и экспертиза промышленной безопасности.",
        }

    items = parse_pairs(rest)

    # Один длинный абзац — detail
    if len(items) == 1 and len(items[0]) > 180:
        return {
            "type": "detail",
            "title": title,
            "section": "Основные положения",
            "body": items[0],
            "note": "",
        }

    if len(items) >= 2:
        return {"type": "items", "title": title, "items": items}

    body = rest[0] if rest else ""
    return {
        "type": "detail",
        "title": title,
        "section": "Основные положения",
        "body": body,
        "note": "",
    }


def build():
    if not SOURCE.exists():
        raise FileNotFoundError(SOURCE)
    if not TEMPLATE.exists():
        raise FileNotFoundError(TEMPLATE)

    src = Presentation(str(SOURCE))
    specs = []
    for i, slide in enumerate(src.slides, 1):
        specs.append(slide_to_spec(extract_blocks(slide), i))

    prs = Presentation(str(TEMPLATE))
    while len(prs.slides) > 0:
        r_id = prs.slides._sldIdLst[0].rId
        prs.part.drop_rel(r_id)
        del prs.slides._sldIdLst[0]

    total = len(specs)
    for page, spec in enumerate(specs, 1):
        t = spec["type"]
        if t == "title":
            add_title_slide(prs, spec["title"], spec["subtitle"], spec["tags"], page, total)
        elif t == "summary":
            add_summary_slide(prs, spec["items"], page, total, spec.get("closing"))
        elif t == "detail":
            add_detail_slide(
                prs,
                spec["title"],
                spec["section"],
                spec["body"],
                spec.get("note", ""),
                page,
                total,
            )
        else:
            add_items_slide(
                prs,
                spec["title"],
                spec["items"],
                page,
                total,
                spec.get("subtitle"),
            )

    prs.save(OUTPUT)
    print(f"Created: {OUTPUT} ({total} slides, {OUTPUT.stat().st_size // 1024} KB)")


if __name__ == "__main__":
    build()
