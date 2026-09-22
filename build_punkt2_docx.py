#!/usr/bin/env python3
"""Пункт 2. Таблицы моделирования по четырём системам, вариант 2."""

from __future__ import annotations

import math
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

OUT_DIR = Path("/workspace/reaktory_sistemy_variant2")
CA0 = 30.0
K1, K2 = 0.6, 0.4

CSTR_SERIES = {
    1: [(0.0, 0.0000, 0.0000, 0.0000), (0.5, 0.2308, 0.1923, 0.0385)],
    2: [(0.0, 0.2308, 0.1923, 0.0385), (0.5, 0.4083, 0.3082, 0.1001)],
    3: [(0.0, 0.4083, 0.3082, 0.1001), (0.5, 0.5448, 0.3706, 0.1742)],
    4: [(0.0, 0.5448, 0.3706, 0.1742), (0.5, 0.6498, 0.3964, 0.2535)],
}
ROWS_CSTR_PAR = [(0.0, 0.0000, 0.0000, 0.0000), (2.0, 0.5454, 0.3030, 0.2424)]


def xyz_from_c(ca, cr, cs):
    return 1.0 - ca, cr, cs


def pfr_step(tau, ca_in, cr_in, cs_in):
    if tau == 0:
        return ca_in, cr_in, cs_in
    ca = ca_in * math.exp(-K1 * tau)
    cr = cr_in * math.exp(-K2 * tau) + (K1 / (K2 - K1)) * ca_in * (
        math.exp(-K1 * tau) - math.exp(-K2 * tau)
    )
    cs = ca_in + cr_in + cs_in - ca - cr
    return ca, cr, cs


def pfr_profile(tau_max, n_steps, ca_in=1.0, cr_in=0.0, cs_in=0.0):
    rows = []
    last = (ca_in, cr_in, cs_in)
    for i in range(n_steps + 1):
        t = round(tau_max * i / n_steps, 4)
        ca, cr, cs = pfr_step(t, ca_in, cr_in, cs_in)
        x, y, z = xyz_from_c(ca, cr, cs)
        rows.append((t, round(x, 4), round(y, 4), round(z, 4)))
        last = (ca, cr, cs)
    return rows, last[0], last[1], last[2]


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


def add_table(doc, rows, highlight_last=True):
    headers = [
        "τ, мин",
        "X",
        "Y",
        "Z",
        "CA, моль/л",
        "CR, моль/л",
        "CS, моль/л",
    ]
    tbl = doc.add_table(rows=1 + len(rows), cols=7)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = True
    for j, h in enumerate(headers):
        fill_cell(tbl.rows[0].cells[j], h, bold=True, size=11, fill="1F4E79", color=(255, 255, 255))
    for i, raw in enumerate(rows):
        tau, x, y, z, ca, cr, cs = rec(*raw)
        vals = [
            comma(tau, 2),
            comma(x, 4),
            comma(y, 4),
            comma(z, 4),
            comma(ca, 2),
            comma(cr, 2),
            comma(cs, 2),
        ]
        last = highlight_last and i == len(rows) - 1
        fill = "FFF2CC" if last else ("D6EAF8" if i % 2 else "FFFFFF")
        for j, v in enumerate(vals):
            fill_cell(tbl.rows[i + 1].cells[j], v, bold=last, size=11, fill=fill)
    doc.add_paragraph()
    return tbl


def build():
    ca_in, cr_in, cs_in = 1.0, 0.0, 0.0
    pfr_series = []
    for _ in range(4):
        rows, ca_in, cr_in, cs_in = pfr_profile(0.5, 5, ca_in, cr_in, cs_in)
        pfr_series.append(rows)
    rows_pfr_par, *_ = pfr_profile(2.0, 10, 1.0, 0.0, 0.0)

    doc = Document()
    for sec in doc.sections:
        sec.orientation = sec.orientation
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
        "CA0 = 30 моль/л.  CA = 30·(1−X),  CR = 30·Y,  CS = 30·Z.  "
        "Жёлтая строка — выход аппарата.  "
        "Для параллельных схем все четыре аппарата одинаковы, поэтому таблица одна.",
        size=12,
        space_after=10,
    )

    add_p(doc, "Система 1. Четыре РИС-н последовательно", size=14, bold=True, color=(31, 78, 121), space_before=6)
    add_p(
        doc,
        "Vi = 25 л, vi = 50 л/мин, τi = 0,5 мин. Данные с экрана программы (идеальное смешение, изотермический режим).",
        size=12,
    )
    for i in range(1, 5):
        add_p(doc, f"Реактор {i}", size=13, bold=True, color=(46, 117, 182), space_before=4, space_after=4)
        add_table(doc, CSTR_SERIES[i])

    add_p(doc, "Система 2. Четыре РИС-н параллельно", size=14, bold=True, color=(31, 78, 121), space_before=8)
    add_p(
        doc,
        "Vi = 25 л, vi = 12,5 л/мин, τi = 2 мин, равная нагрузка. "
        "Выход каждого аппарата совпадает с единичным РИС-н (л/р №3, n1 = n2 = 1).",
        size=12,
    )
    add_p(doc, "Реакторы 1–4 (одинаковые)", size=13, bold=True, color=(46, 117, 182), space_before=4, space_after=4)
    add_table(doc, ROWS_CSTR_PAR)

    add_p(doc, "Система 3. Четыре РИВ последовательно", size=14, bold=True, color=(31, 78, 121), space_before=8)
    add_p(
        doc,
        "Vi = 25 л, vi = 50 л/мин, τi = 0,5 мин. Суммарное время пребывания Στ = 2 мин.",
        size=12,
    )
    for i, rows in enumerate(pfr_series, 1):
        add_p(doc, f"Реактор {i}", size=13, bold=True, color=(46, 117, 182), space_before=4, space_after=4)
        add_table(doc, rows)

    add_p(doc, "Система 4. Четыре РИВ параллельно", size=14, bold=True, color=(31, 78, 121), space_before=8)
    add_p(
        doc,
        "Vi = 25 л, vi = 12,5 л/мин, τi = 2 мин. "
        "Выход каждого аппарата совпадает с единичным РИВ (л/р №3, n1 = n2 = 1).",
        size=12,
    )
    add_p(doc, "Реакторы 1–4 (одинаковые)", size=13, bold=True, color=(46, 117, 182), space_before=4, space_after=4)
    add_table(doc, rows_pfr_par)

    add_p(doc, "Краткие выходы систем", size=14, bold=True, color=(31, 78, 121), space_before=8)
    summary = [
        ("1. 4 РИС-н последовательно", 0.6498, 0.3964, 0.2535),
        ("2. 4 РИС-н параллельно", 0.5454, 0.3030, 0.2424),
        ("3. 4 РИВ последовательно", 0.6988, 0.4444, 0.2544),
        ("4. 4 РИВ параллельно", 0.6988, 0.4444, 0.2544),
    ]
    tbl = doc.add_table(rows=5, cols=7)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    heads = ["Система", "X", "Y", "Z", "CA, моль/л", "CR, моль/л", "CS, моль/л"]
    for j, h in enumerate(heads):
        fill_cell(tbl.rows[0].cells[j], h, bold=True, size=11, fill="1F4E79", color=(255, 255, 255))
    for i, (name, x, y, z) in enumerate(summary):
        vals = [name, comma(x, 4), comma(y, 4), comma(z, 4), comma(CA0 * (1 - x), 2), comma(CA0 * y, 2), comma(CA0 * z, 2)]
        fill = "E2EFDA" if i >= 2 else "FFF2CC"
        for j, v in enumerate(vals):
            fill_cell(tbl.rows[i + 1].cells[j], v, bold=(j == 0), size=11, fill=fill)
            if j == 0:
                tbl.rows[i + 1].cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT

    add_p(
        doc,
        "1. 4 РИС-н в ряду: CA = 10,51, CR = 11,89, CS = 7,61 моль/л (X = 0,6498).  "
        "2. 4 РИС-н параллельно: CA = 13,64, CR = 9,09, CS = 7,27 моль/л (X = 0,5454).  "
        "3–4. 4 РИВ в ряду и параллельно: CA = 9,04, CR = 13,33, CS = 7,63 моль/л (X = 0,6988).",
        size=12,
        space_before=10,
    )

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    cyr = OUT_DIR / "Пункт2_таблицы_моделирования.docx"
    lat = OUT_DIR / "Punkt2_modeling_tables.docx"
    doc.save(cyr)
    # ASCII copy for iPhone
    import shutil

    shutil.copyfile(cyr, lat)
    print("saved", cyr, lat)


if __name__ == "__main__":
    build()
