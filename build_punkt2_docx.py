#!/usr/bin/env python3
"""Пункт 2. Четыре таблицы, как на экране: τ = 0 и τ = 0,5."""

from __future__ import annotations

import shutil
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

OUT_DIR = Path("/workspace/reaktory_sistemy_variant2")
CA0 = 30.0

# Как на экране: две строки — вход (τ = 0) и выход (τ = 0,5).
# Выход — системы целиком (последний аппарат / смешение параллели).
TABLES = [
    (
        "Таблица 1. Система 1 — четыре РИС-н последовательно",
        "Vi = 25 л, vi = 50 л/мин, τi = 0,5 мин. Две строки с экрана: вход системы и выход реактора 4.",
        (0.0, 0.0000, 0.0000, 0.0000),
        (0.5, 0.6498, 0.3964, 0.2535),
    ),
    (
        "Таблица 2. Система 2 — четыре РИС-н параллельно",
        "Vi = 25 л, vi = 12,5 л/мин. Аппараты одинаковы. На экране те же две строки; выход как у единичного РИС-н.",
        (0.0, 0.0000, 0.0000, 0.0000),
        (0.5, 0.5454, 0.3030, 0.2424),
    ),
    (
        "Таблица 3. Система 3 — четыре РИВ последовательно",
        "Vi = 25 л, vi = 50 л/мин, τi = 0,5 мин. Вход системы и выход реактора 4.",
        (0.0, 0.0000, 0.0000, 0.0000),
        (0.5, 0.6988, 0.4444, 0.2544),
    ),
    (
        "Таблица 4. Система 4 — четыре РИВ параллельно",
        "Vi = 25 л, vi = 12,5 л/мин. Аппараты одинаковы. Выход как у единичного РИВ.",
        (0.0, 0.0000, 0.0000, 0.0000),
        (0.5, 0.6988, 0.4444, 0.2544),
    ),
]


def rec(tau, x, y, z):
    return (tau, x, y, z, CA0 * (1.0 - x), CA0 * y, CA0 * z)


def comma(v, nd):
    return f"{v:.{nd}f}".replace(".", ",")


def set_run(run, size=14, bold=False, color=None):
    run.font.name = "Times New Roman"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)


def add_p(doc, text, size=14, bold=False, color=None, space_after=6, space_before=0):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.line_spacing = 1.15
    run = p.add_run(text)
    set_run(run, size=size, bold=bold, color=color)
    return p


def shade(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:shd")):
        tcPr.remove(old)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def set_cell_border(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:color"), "808080")
        tcBorders.append(el)
    tcPr.append(tcBorders)


def fill_cell(cell, text, bold=False, size=11, fill=None, color=None):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    run = p.add_run(text)
    set_run(run, size=size, bold=bold, color=color)
    set_cell_border(cell)
    if fill:
        shade(cell, fill)


def add_table(doc, inlet, outlet):
    headers = ["τ, мин", "X", "Y", "Z", "CA, моль/л", "CR, моль/л", "CS, моль/л"]
    tbl = doc.add_table(rows=3, cols=7)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = True
    for j, h in enumerate(headers):
        fill_cell(tbl.rows[0].cells[j], h, bold=True, size=11, fill="1F4E79", color=(255, 255, 255))
    for i, raw in enumerate((inlet, outlet)):
        tau, x, y, z, ca, cr, cs = rec(*raw)
        vals = [comma(tau, 2), comma(x, 4), comma(y, 4), comma(z, 4), comma(ca, 2), comma(cr, 2), comma(cs, 2)]
        last = i == 1
        fill = "FFF2CC" if last else "FFFFFF"
        for j, v in enumerate(vals):
            fill_cell(tbl.rows[i + 1].cells[j], v, bold=last, size=11, fill=fill)
    doc.add_paragraph()
    return tbl


def build():
    doc = Document()
    for sec in doc.sections:
        sec.page_width = Cm(29.7)
        sec.page_height = Cm(21.0)
        sec.left_margin = Cm(1.5)
        sec.right_margin = Cm(1.5)
        sec.top_margin = Cm(1.4)
        sec.bottom_margin = Cm(1.4)

    add_p(doc, "Лабораторная работа «Реакторные системы». Вариант 2", size=16, bold=True, space_after=2)
    add_p(doc, "Пункт 2. Таблицы результатов моделирования", size=16, bold=True, space_after=8)
    add_p(
        doc,
        "Четыре системы — четыре таблицы. На экране у каждой две строки: τ = 0 (вход) и τ = 0,5 мин (выход). "
        "CA0 = 30 моль/л: CA = 30·(1−X), CR = 30·Y, CS = 30·Z. Жёлтая строка — выход.",
        size=12,
        space_after=10,
    )

    for title, note, inlet, outlet in TABLES:
        add_p(doc, title, size=14, bold=True, color=(31, 78, 121), space_before=8)
        add_p(doc, note, size=12)
        add_table(doc, inlet, outlet)

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    cyr = OUT_DIR / "Пункт2_таблицы_моделирования.docx"
    lat = OUT_DIR / "Punkt2_modeling_tables.docx"
    doc.save(cyr)
    shutil.copyfile(cyr, lat)
    print("saved", cyr, lat)


if __name__ == "__main__":
    build()
