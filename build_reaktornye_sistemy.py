#!/usr/bin/env python3
"""Л/р «Реакторные системы». Вариант 2. n1 = n2 = 1. Четыре схемы по 4 реактора."""

from __future__ import annotations

import math
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUT = Path("/workspace/reaktory_sistemy_variant2/Реакторные_системы_вариант2.xlsx")

CA0 = 30.0
V_SYS = 100.0  # л, суммарный объём системы
V0_SYS = 50.0  # л/мин через систему
N = 4
V_I = V_SYS / N  # 25 л
K1, K2 = 0.6, 0.4  # мин⁻¹, как л/р №3 случай n1 = n2 = 1
PI_FACTOR = V0_SYS * 60 / 1000.0  # ΠR = 3·CR кмоль/ч при полном расходе системы

HEAD_FILL = PatternFill("solid", fgColor="1F4E79")
HEAD_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(name="Calibri", bold=True, size=14, color="1F4E79")
SUB_FONT = Font(name="Calibri", bold=True, size=12, color="2E75B6")
CELL_FONT = Font(name="Calibri", size=11)
NUM_FMT = "0.0000"
CONC_FMT = "0.00"


def thin():
    return Border(
        left=Side(style="thin", color="B0B0B0"),
        right=Side(style="thin", color="B0B0B0"),
        top=Side(style="thin", color="B0B0B0"),
        bottom=Side(style="thin", color="B0B0B0"),
    )


def style_header(cell):
    cell.fill = HEAD_FILL
    cell.font = HEAD_FONT
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = thin()


def style_cell(cell, fmt=None):
    cell.font = CELL_FONT
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = thin()
    if fmt:
        cell.number_format = fmt


def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def xyz_from_c(ca, cr, cs):
    x = 1.0 - ca
    y = cr
    z = cs
    s = (1.0 - K2 * cr / (K1 * ca)) if ca > 1e-12 else 0.0
    return x, y, z, s


def rec(tau, x, y, z):
    return {
        "tau": tau,
        "X": x,
        "CA": CA0 * (1.0 - x),
        "Y": y,
        "CR": CA0 * y,
        "Z": z,
        "CS": CA0 * z,
        "PiR": PI_FACTOR * CA0 * y,
    }


# --- РИС-н в ряду: четыре экрана (τ_i = 0,5 мин) ---
CSTR_SERIES = {
    1: [
        (0.0, 0.0000, 0.0000, 0.0000),
        (0.5, 0.2308, 0.1923, 0.0385),
    ],
    2: [
        (0.0, 0.2308, 0.1923, 0.0385),
        (0.5, 0.4083, 0.3082, 0.1001),
    ],
    3: [
        (0.0, 0.4083, 0.3082, 0.1001),
        (0.5, 0.5448, 0.3706, 0.1742),
    ],
    4: [
        (0.0, 0.5448, 0.3706, 0.1742),
        (0.5, 0.6498, 0.3964, 0.2535),
    ],
}

# л/р №3, единичные, n1 = n2 = 1, τ = 2 мин
SINGLE_CSTR = (2.0, 0.5454, 0.3030, 0.2424)
SINGLE_PFR = (2.0, 0.6988, 0.4444, 0.2544)


def cstr_from_inlet(tau, ca_in, cr_in, cs_in):
    """Один РИС-н, безразмерные концентрации (C0 = 1)."""
    ca = ca_in / (1.0 + K1 * tau)
    cr = (cr_in + K1 * ca * tau) / (1.0 + K2 * tau)
    cs = cs_in + K2 * cr * tau
    return ca, cr, cs


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
    for i in range(n_steps + 1):
        t = round(tau_max * i / n_steps, 4)
        ca, cr, cs = pfr_step(t, ca_in, cr_in, cs_in)
        x, y, z, _s = xyz_from_c(ca, cr, cs)
        rows.append((t, round(x, 4), round(y, 4), round(z, 4)))
    return rows, ca, cr, cs


def write_block(ws, r0, title, rows):
    headers = [
        "τ, мин",
        "X",
        "CA, моль/л",
        "Y",
        "CR, моль/л",
        "Z",
        "CS, моль/л",
    ]
    ws.merge_cells(start_row=r0, start_column=1, end_row=r0, end_column=7)
    t = ws.cell(r0, 1, title)
    t.font = SUB_FONT
    t.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[r0].height = 22
    hr = r0 + 1
    for i, h in enumerate(headers):
        style_header(ws.cell(hr, i + 1, h))
    ws.row_dimensions[hr].height = 28
    for i, raw in enumerate(rows):
        d = rec(*raw)
        vals = [d["tau"], d["X"], d["CA"], d["Y"], d["CR"], d["Z"], d["CS"]]
        fmts = ["0.00", NUM_FMT, CONC_FMT, NUM_FMT, CONC_FMT, NUM_FMT, CONC_FMT]
        for j, (v, fmt) in enumerate(zip(vals, fmts), 1):
            cell = ws.cell(hr + 1 + i, j, v)
            style_cell(cell, fmt)
        if i % 2 == 1:
            for j in range(1, 8):
                ws.cell(hr + 1 + i, j).fill = PatternFill("solid", fgColor="D6EAF8")
    return hr + 1 + len(rows)


def build():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()

    # ----- исходные -----
    w0 = wb.active
    w0.title = "Исходные"
    w0["A1"] = "Реакторные системы. Изотермические процессы. Вариант 2"
    w0["A1"].font = TITLE_FONT
    w0.merge_cells("A1:B1")
    lines = [
        ("Реакция", "A → R → S, n1 = n2 = 1 (как л/р №3, случай n1 = n2)"),
        ("k1, k2", "0,6 мин⁻¹ и 0,4 мин⁻¹; T = 150 °C; Ei = 0"),
        ("CA0", "30 моль/л"),
        ("V системы", "100 л (суммарно, одинаково для всех схем)"),
        ("v0 через систему", "50 л/мин"),
        ("Число аппаратов", "4, объёмы равны: Vi = 25 л"),
        ("1. РИС-н послед.", "последовательно; расход через каждый = 50 л/мин; τi = 25/50 = 0,5 мин"),
        ("2. РИС-н паралл.", "параллельно, одинаковая нагрузка; vi = 12,5 л/мин; τi = 25/12,5 = 2 мин"),
        ("3. РИВ послед.", "последовательно; τi = 0,5 мин; суммарно τ = 2 мин"),
        ("4. РИВ паралл.", "параллельно, vi = 12,5 л/мин; τi = 2 мин"),
        ("Концентрации", "CA = 30·(1−X); CR = 30·Y; CS = 30·Z"),
        ("ΠR системы", "v0·CR_вых = 3·CR_вых, кмоль/ч (CR в моль/л)"),
    ]
    for i, (a, b) in enumerate(lines, 3):
        w0.cell(i, 1, a).font = Font(name="Calibri", bold=True, size=11)
        w0.cell(i, 2, b).font = CELL_FONT
        w0.cell(i, 2).alignment = Alignment(wrap_text=True)
        w0.row_dimensions[i].height = 20
    set_widths(w0, [22, 92])

    # ----- 1. четыре РИС-н последовательно -----
    w1 = wb.create_sheet("1_РИСн_последовательно")
    w1["A1"] = "Схема 1. Четыре РИС-н одинакового объёма, последовательно. Данные с экрана программы."
    w1["A1"].font = TITLE_FONT
    w1.merge_cells("A1:G1")
    r = 3
    for i in range(1, 5):
        r = write_block(w1, r, f"Реактор {i} (РИС-н), τi = 0,5 мин", CSTR_SERIES[i]) + 2
    note = w1.cell(
        r,
        1,
        "На входе следующего аппарата X, Y, Z равны выходу предыдущего. "
        "S на экране на входе каждого реактора сбрасывается к 1 — это локальный счётчик программы, не физическая селективность.",
    )
    note.alignment = Alignment(wrap_text=True)
    w1.merge_cells(start_row=r, start_column=1, end_row=r + 1, end_column=7)
    set_widths(w1, [12, 12, 14, 12, 14, 12, 14])
    w1.page_setup.orientation = "landscape"
    w1.freeze_panes = "A3"

    # ----- 2. четыре РИС-н параллельно -----
    w2 = wb.create_sheet("2_РИСн_параллельно")
    w2["A1"] = (
        "Схема 2. Четыре РИС-н одинакового объёма, параллельно, одинаковая нагрузка. "
        "Каждый: Vi = 25 л, vi = 12,5 л/мин, τi = 2 мин — как единичный РИС-н л/р №3."
    )
    w2["A1"].font = TITLE_FONT
    w2.merge_cells("A1:G1")
    w2.row_dimensions[1].height = 32
    ca, cr, cs = cstr_from_inlet(2.0, 1.0, 0.0, 0.0)
    x, y, z, _ = xyz_from_c(ca, cr, cs)
    # экран единичного РИС-н л/р №3
    rows_par = [
        (0.0, 0.0000, 0.0000, 0.0000),
        (2.0, 0.5454, 0.3030, 0.2424),
    ]
    r = 3
    for i in range(1, 5):
        r = write_block(
            w2,
            r,
            f"Реактор {i} (РИС-н, параллель) — состав тот же, что у остальных трёх",
            rows_par,
        ) + 2
    w2.cell(
        r,
        1,
        f"Проверка баланса РИС-н: X = {x:.4f}, Y = {y:.4f}, Z = {z:.4f} (совпадает с экраном 0,5454 / 0,3030 / 0,2424).",
    ).alignment = Alignment(wrap_text=True)
    w2.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    set_widths(w2, [12, 12, 14, 12, 14, 12, 14])
    w2.page_setup.orientation = "landscape"

    # ----- 3. четыре РИВ последовательно -----
    w3 = wb.create_sheet("3_РИВ_последовательно")
    w3["A1"] = (
        "Схема 3. Четыре РИВ одинакового объёма, последовательно. "
        "Каждый τi = 0,5 мин; вместе эквивалентны одному РИВ V = 100 л, τ = 2 мин."
    )
    w3["A1"].font = TITLE_FONT
    w3.merge_cells("A1:G1")
    w3.row_dimensions[1].height = 32
    r = 3
    ca_in, cr_in, cs_in = 1.0, 0.0, 0.0
    pfr_series_outlets = []
    for i in range(1, 5):
        rows, ca_in, cr_in, cs_in = pfr_profile(0.5, 5, ca_in, cr_in, cs_in)
        pfr_series_outlets.append(rows[-1])
        r = write_block(w3, r, f"Реактор {i} (РИВ), τi = 0…0,5 мин (шаг 0,1)", rows) + 2
    x, y, z = pfr_series_outlets[-1][1], pfr_series_outlets[-1][2], pfr_series_outlets[-1][3]
    w3.cell(
        r,
        1,
        f"Выход системы: X = {x:.4f}, Y = {y:.4f}, Z = {z:.4f} — как единичный РИВ л/р №3 (0,6988 / 0,4444 / 0,2544).",
    ).alignment = Alignment(wrap_text=True)
    w3.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    set_widths(w3, [12, 12, 14, 12, 14, 12, 14])
    w3.page_setup.orientation = "landscape"
    w3.freeze_panes = "A3"

    # ----- 4. четыре РИВ параллельно -----
    w4 = wb.create_sheet("4_РИВ_параллельно")
    w4["A1"] = (
        "Схема 4. Четыре РИВ одинакового объёма, параллельно, одинаковая нагрузка. "
        "Каждый τi = 2 мин — как единичный РИВ л/р №3."
    )
    w4["A1"].font = TITLE_FONT
    w4.merge_cells("A1:G1")
    w4.row_dimensions[1].height = 32
    rows_pfr, *_rest = pfr_profile(2.0, 10, 1.0, 0.0, 0.0)
    # подставить экранные точки л/р №3 на узлах 0,2
    r = 3
    for i in range(1, 5):
        r = write_block(
            w4,
            r,
            f"Реактор {i} (РИВ, параллель) — профиль тот же, что у остальных трёх",
            rows_pfr,
        ) + 2
    set_widths(w4, [12, 12, 14, 12, 14, 12, 14])
    w4.page_setup.orientation = "landscape"
    w4.freeze_panes = "A3"

    # ----- ΠR, сравнение, вывод -----
    w5 = wb.create_sheet("Производительность_вывод")
    w5["A1"] = "Производительность систем по R и сравнение (п. 3–4 задания)"
    w5["A1"].font = TITLE_FONT
    w5.merge_cells("A1:F1")

    headers = ["Схема", "X вых.", "Y вых.", "CR, моль/л", "ΠR, кмоль/ч", "Примечание"]
    for i, h in enumerate(headers, 1):
        style_header(w5.cell(3, i, h))
    w5.row_dimensions[3].height = 28

    y_ser = CSTR_SERIES[4][-1][2]
    y_cstr = SINGLE_CSTR[2]
    y_pfr = SINGLE_PFR[2]
    data = [
        ("Единичный РИС-н (л/р №3)", SINGLE_CSTR[1], y_cstr, "τ = 2 мин, V = 100 л"),
        ("Единичный РИВ (л/р №3)", SINGLE_PFR[1], y_pfr, "τ = 2 мин, V = 100 л"),
        ("4 РИС-н последовательно", CSTR_SERIES[4][-1][1], y_ser, "τi = 0,5 мин, Στ = 2 мин"),
        ("4 РИС-н параллельно", SINGLE_CSTR[1], y_cstr, "как один РИС-н: τi = 2 мин"),
        ("4 РИВ последовательно", SINGLE_PFR[1], y_pfr, "как один РИВ: Στ = 2 мин"),
        ("4 РИВ параллельно", SINGLE_PFR[1], y_pfr, "как один РИВ: τi = 2 мин"),
    ]
    for i, (name, x, y, note) in enumerate(data):
        cr = CA0 * y
        pir = PI_FACTOR * cr
        vals = [name, x, y, cr, pir, note]
        fmts = [None, NUM_FMT, NUM_FMT, CONC_FMT, "0.00", None]
        for j, (v, fmt) in enumerate(zip(vals, fmts), 1):
            cell = w5.cell(4 + i, j, v)
            style_cell(cell, fmt)
            if j in (1, 6):
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        if "РИВ" in name:
            for j in range(1, 7):
                w5.cell(4 + i, j).fill = PatternFill("solid", fgColor="E2EFDA")

    chart = BarChart()
    chart.type = "col"
    chart.title = "ΠR, кмоль/ч"
    chart.y_axis.title = "ΠR, кмоль/ч"
    chart.legend = None
    data_ref = Reference(w5, min_col=5, min_row=3, max_row=9)
    cats = Reference(w5, min_col=1, min_row=4, max_row=9)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats)
    chart.shape = 4
    chart.width = 18
    chart.height = 9
    w5.add_chart(chart, "A12")

    w5["A28"] = "Вывод"
    w5["A28"].font = SUB_FONT
    conclusion = (
        "Кинетика n1 = n2 = 1, k1 = 0,6 мин⁻¹, k2 = 0,4 мин⁻¹, CA0 = 30 моль/л, "
        "суммарный объём 100 л, расход системы 50 л/мин.\n\n"
        "Четыре РИС-н параллельно дают тот же выход, что единичный РИС-н (X = 0,5454, "
        "ΠR = 27,27 кмоль/ч): нагрузка и τ на каждый аппарат те же (τi = 2 мин). "
        "Четыре РИВ — и последовательно, и параллельно — совпадают с единичным РИВ "
        "(X = 0,6988, ΠR = 40,00 кмоль/ч): в ряду складываются времена пребывания, "
        "в параллели каждое τi = 2 мин при том же кинетическом режиме.\n\n"
        "Четыре РИС-н последовательно лучше одного РИС-н и хуже РИВ: X = 0,6498, "
        "Y = 0,3964, ΠR = 35,68 кмоль/ч. Каскад смешения приближается к вытеснению "
        "(модель ёмкостей в ряду).\n\n"
        "Гидравлическое сопротивление. В последовательных схемах через каждый аппарат "
        "идёт полный расход 50 л/мин, перепады складываются — сопротивление системы выше. "
        "В параллельных схемах расход на аппарат 12,5 л/мин, общий ΔP определяется одной "
        "ветвью и заметно меньше. По ΠR вытеснение (ряд или параллель) равно и максимально; "
        "по гидравлике выгоднее параллельный РИВ: та же производительность 40 кмоль/ч при "
        "меньшем сопротивлении."
    )
    w5.merge_cells("A29:F36")
    c = w5["A29"]
    c.value = conclusion
    c.alignment = Alignment(wrap_text=True, vertical="top")
    c.font = Font(name="Calibri", size=12)
    w5.row_dimensions[29].height = 80
    set_widths(w5, [32, 12, 12, 14, 16, 42])
    w5.page_setup.orientation = "landscape"

    wb.save(OUT)
    print("saved", OUT)
    print("4 CSTR series Y", y_ser, "Pi", PI_FACTOR * CA0 * y_ser)
    print("PFR series last", pfr_series_outlets[-1])


if __name__ == "__main__":
    build()
