#!/usr/bin/env python3
"""Графики Ci(τ) в стиле Kniga1.xlsx: маркеры square / star / x, линии smoothMarker."""

from __future__ import annotations

import math
from copy import copy
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import ScatterChart, Reference, Series
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.legend import Legend
from openpyxl.chart.marker import Marker
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.series import SeriesLabel
from openpyxl.chart.text import RichText
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import (
    CharacterProperties,
    Paragraph,
    ParagraphProperties,
    RegularTextRun,
)
from openpyxl.drawing.colors import ColorChoice, SchemeColor
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart.layout import Layout, ManualLayout

OUT = Path("/workspace/reaktory_variant2/Реакторы_вариант2_графики.xlsx")
PREVIEW_DIR = Path("/workspace/reaktory_variant2")

CA0 = 30.0  # моль/л
V0_L_MIN = 50.0
# ΠR = v0·CR : 50 л/мин · CR моль/л · 60 мин/ч / 1000 = 3·CR кмоль/ч
PI_FACTOR = V0_L_MIN * 60 / 1000.0

# РИВ — точки с экрана программы (без промежуточной 1.85)
PFR = {
    1: [  # n1=2, n2=1
        (0.0, 0.0000, 0.0000, 0.0000, 1.0000),
        (0.2, 0.1071, 0.1028, 0.0043, 0.9596),
        (0.4, 0.1935, 0.1778, 0.0157, 0.9188),
        (0.6, 0.2647, 0.2325, 0.0322, 0.8782),
        (0.8, 0.3243, 0.2718, 0.0525, 0.8381),
        (1.0, 0.3750, 0.2996, 0.0754, 0.7989),
        (1.2, 0.4186, 0.3184, 0.1002, 0.7606),
        (1.4, 0.4565, 0.3303, 0.1262, 0.7236),
        (1.6, 0.4898, 0.3369, 0.1529, 0.6878),
        (1.8, 0.5192, 0.3393, 0.1800, 0.6534),
        (2.0, 0.5455, 0.3384, 0.2071, 0.6203),
    ],
    2: [  # n1=n2=1
        (0.0, 0.0000, 0.0000, 0.0000, 1.0000),
        (0.2, 0.1131, 0.1086, 0.0045, 0.9603),
        (0.4, 0.2134, 0.1965, 0.0168, 0.9211),
        (0.6, 0.3023, 0.2669, 0.0355, 0.8827),
        (0.8, 0.3812, 0.3221, 0.0591, 0.8449),
        (1.0, 0.4512, 0.3645, 0.0867, 0.8079),
        (1.2, 0.5132, 0.3961, 0.1172, 0.7717),
        (1.4, 0.5683, 0.4185, 0.1498, 0.7364),
        (1.6, 0.6171, 0.4332, 0.1839, 0.7020),
        (1.8, 0.6604, 0.4415, 0.2189, 0.6685),
        (2.0, 0.6988, 0.4444, 0.2544, 0.6359),
    ],
    3: [  # n1=1, n2=2
        (0.0, 0.0000, 0.0000, 0.0000, 1.0000),
        (0.2, 0.1131, 0.1127, 0.0004, 0.9969),
        (0.4, 0.2134, 0.2108, 0.0025, 0.9881),
        (0.6, 0.3023, 0.2946, 0.0077, 0.9744),
        (0.8, 0.3812, 0.3647, 0.0165, 0.9566),
        (1.0, 0.4512, 0.4222, 0.0290, 0.9357),
        (1.2, 0.5132, 0.4683, 0.0449, 0.9124),
        (1.4, 0.5683, 0.5044, 0.0639, 0.8875),
        (1.6, 0.6171, 0.5317, 0.0855, 0.8615),
        (1.8, 0.6604, 0.5514, 0.1090, 0.8350),
        (2.0, 0.6988, 0.5649, 0.1339, 0.8083),
    ],
}

# РИС-н на τ=2 с экрана
CSTR_OUT = {
    1: (0.4132, 0.2296, 0.1836, 0.5556),
    2: (0.5454, 0.3030, 0.2424, 0.5556),
    3: (0.5454, 0.4106, 0.1349, 0.7528),
}

CASE_TITLES = {
    1: "случай 1: n1=2, n2=1; k1=0,6 л/(моль·мин); k2=0,4 мин⁻¹",
    2: "случай 2: n1=n2=1; k1=0,6 мин⁻¹; k2=0,4 мин⁻¹",
    3: "случай 3: n1=1, n2=2; k1=0,6 мин⁻¹; k2=0,4 л/(моль·мин)",
}


def cstr_case1_profile():
    """Профиль РИС-н для графика случая 1 (кинетика как в программе, C0=1)."""
    k1, k2 = 0.6, 0.4
    rows = []
    for i in range(11):
        tau = round(i * 0.2, 1)
        if tau == 0:
            x = y = z = 0.0
        else:
            ca = (-1.0 + math.sqrt(1.0 + 4.0 * k1 * tau)) / (2.0 * k1 * tau)
            x = 1.0 - ca
            cr = k1 * ca**2 * tau / (1.0 + k2 * tau)
            z = k2 * cr * tau
            y = cr
        s = y / x if x else 1.0
        rows.append((tau, x, y, z, s))
    # подставить экранные значения на τ=2
    rows[-1] = (2.0, *CSTR_OUT[1])
    return rows


def conc_row(tau, x, y, z, s):
    return {
        "tau": tau,
        "X": x,
        "CA": CA0 * (1.0 - x),
        "Y": y,
        "CR": CA0 * y,
        "Z": z,
        "CS": CA0 * z,
        "S": s,
        "SR": (y / x) if x else None,
        "PiR": PI_FACTOR * CA0 * y,
    }


def thin():
    return Border(
        left=Side(style="thin", color="B0B0B0"),
        right=Side(style="thin", color="B0B0B0"),
        top=Side(style="thin", color="B0B0B0"),
        bottom=Side(style="thin", color="B0B0B0"),
    )


HEAD_FILL = PatternFill("solid", fgColor="1F4E79")
HEAD_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(name="Calibri", bold=True, size=14, color="1F4E79")
SUB_FONT = Font(name="Calibri", bold=True, size=12, color="2E75B6")
CELL_FONT = Font(name="Calibri", size=11)
NUM_FMT = "0.0000"
CONC_FMT = "0.00"


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


def axis_title(text: str, rot: int | None = None):
    """Подпись оси в том же духе, что в Kniga1 (RichText)."""
    from openpyxl.chart.title import Title
    from openpyxl.chart.text import Text
    from openpyxl.drawing.text import RichTextProperties, ListStyle

    body = RichTextProperties(
        rot=rot if rot is not None else 0,
        spcFirstLastPara=True,
        vertOverflow="ellipsis",
        vert="horz",
        wrap="square",
        anchor="ctr",
        anchorCtr=True,
    )
    def_rpr = CharacterProperties(sz=1100, b=False, kern=1200, baseline=0)
    ppr = ParagraphProperties(defRPr=def_rpr)
    run = RegularTextRun(t=text)
    run.rPr = CharacterProperties(lang="ru-RU")
    p = Paragraph(pPr=ppr, r=[run])
    rich = RichText(bodyPr=body, p=[p], lstStyle=ListStyle())
    tx = Text()
    tx.rich = rich
    title = Title()
    title.tx = tx
    title.overlay = False
    gp = GraphicalProperties()
    gp.noFill = True
    gp.line = LineProperties(noFill=True, prstDash="solid")
    title.spPr = gp
    return title


def scheme_line(accent: str, width=19050, no_fill=False) -> GraphicalProperties:
    gp = GraphicalProperties()
    ln = LineProperties(w=width, cap="rnd", prstDash="solid", round=True)
    if no_fill:
        ln.noFill = True
    else:
        ln.solidFill = ColorChoice(schemeClr=SchemeColor(val=accent))
    gp.ln = ln
    return gp


def scheme_marker(symbol: str, accent: str, size=10.0, hollow=False) -> Marker:
    m = Marker(symbol=symbol, size=size)
    sp = GraphicalProperties()
    if hollow:
        sp.noFill = True
    else:
        sp.solidFill = ColorChoice(schemeClr=SchemeColor(val=accent))
    ln = LineProperties(w=9525, prstDash="solid")
    ln.solidFill = ColorChoice(schemeClr=SchemeColor(val=accent))
    sp.ln = ln
    m.spPr = sp
    return m


def make_scatter(ws, title, x_col, y_cols, start_row, n_pts, anchor, y_max=32.0):
    """
    y_cols: list of (col_idx, series_title, marker_symbol, accent, size)
    x_col: 1-based column of τ
    """
    chart = ScatterChart()
    chart.scatterStyle = "smoothMarker"
    chart.title = title
    chart.style = 10
    chart.x_axis.title = axis_title("τ, мин")
    chart.y_axis.title = axis_title("Ci, моль/л", rot=-5400000)
    chart.x_axis.scaling.min = 0.0
    chart.x_axis.scaling.max = 2.0
    chart.y_axis.scaling.min = 0.0
    chart.y_axis.scaling.max = y_max
    chart.y_axis.majorUnit = 5.0
    chart.x_axis.majorUnit = 0.25
    chart.x_axis.majorGridlines = ChartLines()
    chart.y_axis.majorGridlines = ChartLines()
    chart.x_axis.axId = 1
    chart.y_axis.axId = 2
    chart.legend = Legend()
    chart.legend.position = "b"
    chart.legend.overlay = False
    # поле точек чуть выше нижней легенды
    chart.plot_area.layout = Layout(
        manualLayout=ManualLayout(
            layoutTarget="inner",
            xMode="edge",
            yMode="edge",
            x=0.10,
            y=0.04,
            w=0.84,
            h=0.78,
        )
    )

    xvalues = Reference(ws, min_col=x_col, min_row=start_row, max_row=start_row + n_pts - 1)
    for col, ser_title, symbol, accent, size in y_cols:
        yvalues = Reference(ws, min_col=col, min_row=start_row, max_row=start_row + n_pts - 1)
        ser = Series(yvalues, xvalues, title=ser_title)
        ser.smooth = True
        ser.marker = scheme_marker(symbol, accent, size=size, hollow=(symbol == "x"))
        ser.graphicalProperties = scheme_line(accent)
        ser.spPr = ser.graphicalProperties
        chart.series.append(ser)

    chart.anchor = anchor
    chart.width = 18
    chart.height = 11
    ws.add_chart(chart)
    return chart


def write_table(ws, r0, c0, title, rows):
    headers = [
        "τ, мин",
        "X",
        "CA, моль/л",
        "Y",
        "CR, моль/л",
        "Z",
        "CS, моль/л",
        "S (дифф.)",
        "SR = Y/X",
        "ΠR, кмоль/ч",
    ]
    ws.merge_cells(start_row=r0, start_column=c0, end_row=r0, end_column=c0 + len(headers) - 1)
    tcell = ws.cell(r0, c0, title)
    tcell.font = SUB_FONT
    tcell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[r0].height = 20

    hr = r0 + 1
    for i, h in enumerate(headers):
        cell = ws.cell(hr, c0 + i, h)
        style_header(cell)
    ws.row_dimensions[hr].height = 32

    for i, raw in enumerate(rows):
        rec = conc_row(*raw) if not isinstance(raw, dict) else raw
        vals = [
            rec["tau"],
            rec["X"],
            rec["CA"],
            rec["Y"],
            rec["CR"],
            rec["Z"],
            rec["CS"],
            rec["S"],
            rec["SR"] if rec["SR"] is not None else 1.0,
            rec["PiR"],
        ]
        fmts = ["0.00", NUM_FMT, CONC_FMT, NUM_FMT, CONC_FMT, NUM_FMT, CONC_FMT, NUM_FMT, NUM_FMT, "0.00"]
        for j, (v, fmt) in enumerate(zip(vals, fmts)):
            cell = ws.cell(hr + 1 + i, c0 + j, v)
            style_cell(cell, fmt)
        if i % 2 == 1:
            for j in range(len(headers)):
                ws.cell(hr + 1 + i, c0 + j).fill = PatternFill("solid", fgColor="D6EAF8")
    return hr + 1, len(rows)  # data start row, n


def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def build():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()

    # --- Лист графиков (как Kniga1: данные слева, график справа) ---
    ws = wb.active
    ws.title = "Графики_случай1"

    ws["A1"] = "Вариант 2. Графики Ci(τ) — оформление как в Kniga1.xlsx (квадраты и звёздочки)"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:J1")
    ws["A2"] = (
        "Маркеры: CA — квадрат (square), CR — звезда (star), CS — звёздочка-крест (x). "
        "Линии сглаженные (smoothMarker), как в образце."
    )
    ws["A2"].font = Font(name="Calibri", italic=True, size=10, color="666666")
    ws.merge_cells("A2:J2")

    # Данные РИВ
    pfr_rows = PFR[1]
    data_start, n = write_table(ws, 4, 1, "РИВ, " + CASE_TITLES[1], pfr_rows)

    make_scatter(
        ws,
        "РИВ, случай 1: Ci = f(τ)",
        x_col=1,
        y_cols=[
            (3, "CA", "square", "accent1", 10.0),
            (5, "CR (целевой)", "star", "accent2", 12.0),
            (7, "CS (побочный)", "x", "accent3", 10.0),
        ],
        start_row=data_start,
        n_pts=n,
        anchor="L4",
    )

    # Данные РИС
    cstr_rows = cstr_case1_profile()
    r2 = data_start + n + 3
    data_start2, n2 = write_table(ws, r2, 1, "РИС-н, " + CASE_TITLES[1] + " (профиль по τ)", cstr_rows)
    make_scatter(
        ws,
        "РИС-н, случай 1: Ci = f(τ)",
        x_col=1,
        y_cols=[
            (3, "CA", "square", "accent1", 10.0),
            (5, "CR (целевой)", "star", "accent2", 12.0),
            (7, "CS (побочный)", "x", "accent3", 10.0),
        ],
        start_row=data_start2,
        n_pts=n2,
        anchor="L24",
    )

    note = ws.cell(
        data_start2 + n2 + 2,
        1,
        "Примечание: X, Y, Z — с экрана программы (безразмерные, C0=1). "
        "CA = 30·(1−X), CR = 30·Y, CS = 30·Z, моль/л. "
        "ΠR = v0·CR = 3·CR, кмоль/ч. Для РИС-н на графике — расчёт при тех же τ, "
        "на τ=2 подставлены значения с экрана (X=0,4132; Y=0,2296; Z=0,1836).",
    )
    note.font = Font(name="Calibri", size=10, color="333333")
    note.alignment = Alignment(wrap_text=True)
    ws.merge_cells(start_row=note.row, start_column=1, end_row=note.row + 2, end_column=10)
    ws.row_dimensions[note.row].height = 36

    set_widths(ws, [12, 12, 14, 12, 14, 12, 14, 12, 12, 14, 12, 12])
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:2"

    # --- Лист все таблицы ---
    wt = wb.create_sheet("Таблицы_все_случаи")
    wt["A1"] = "Результаты моделирования. Вариант 2. CA0 = 30 моль/л; V = 100 л; v0 = 50 л/мин; τ = V/v0 = 2 мин; T = 150 °C"
    wt["A1"].font = TITLE_FONT
    wt.merge_cells("A1:J1")
    wt.row_dimensions[1].height = 24

    r = 3
    for case in (1, 2, 3):
        r_start, n = write_table(wt, r, 1, f"РИВ, {CASE_TITLES[case]}", PFR[case])
        r = r_start + n + 3
        x, y, z, s = CSTR_OUT[case]
        cstr_two = [(0.0, 0.0, 0.0, 0.0, 1.0), (2.0, x, y, z, s)]
        r_start, n = write_table(wt, r, 1, f"РИС-н (экран программы), {CASE_TITLES[case]}", cstr_two)
        r = r_start + n + 4

    set_widths(wt, [12, 12, 14, 12, 14, 12, 14, 12, 12, 14])
    wt.page_setup.orientation = "landscape"
    wt.page_setup.fitToPage = True
    wt.page_setup.fitToWidth = 1
    wt.page_setup.fitToHeight = 0
    wt.sheet_properties.pageSetUpPr.fitToPage = True

    # --- Сводка ΠR, SR, выводы ---
    ws2 = wb.create_sheet("Производительность_выводы")
    ws2["A1"] = "Производительность по R, интегральная селективность и выводы"
    ws2["A1"].font = TITLE_FONT
    ws2.merge_cells("A1:G1")

    headers = ["Случай", "Реактор", "X (τ=2)", "Y", "CR, моль/л", "ΠR, кмоль/ч", "SR = Y/X"]
    for i, h in enumerate(headers, 1):
        style_header(ws2.cell(3, i, h))
    ws2.row_dimensions[3].height = 28

    row = 4
    summary = []
    for case in (1, 2, 3):
        pfr = conc_row(*PFR[case][-1])
        x, y, z, s = CSTR_OUT[case]
        cstr = conc_row(2.0, x, y, z, s)
        for name, rec in (("РИС-н", cstr), ("РИВ", pfr)):
            vals = [
                CASE_TITLES[case].split(":")[0],
                name,
                rec["X"],
                rec["Y"],
                rec["CR"],
                rec["PiR"],
                rec["SR"],
            ]
            fmts = [None, None, NUM_FMT, NUM_FMT, CONC_FMT, "0.00", NUM_FMT]
            for j, (v, fmt) in enumerate(zip(vals, fmts), 1):
                cell = ws2.cell(row, j, v)
                style_cell(cell, fmt)
                if j == 1:
                    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            if name == "РИВ":
                for j in range(1, 8):
                    ws2.cell(row, j).fill = PatternFill("solid", fgColor="E2EFDA")
            summary.append((case, name, rec))
            row += 1

    ws2.merge_cells("A11:G11")
    ws2["A11"] = "Вывод"
    ws2["A11"].font = SUB_FONT

    text = (
        "Моделирование последовательной реакции A → R → S выполнено в изотермическом режиме "
        "(T = 150 °C) для реакторов идеального смешения непрерывного действия (РИС-н) и "
        "идеального вытеснения (РИВ) одинакового объёма: V = 100 л, v0 = 50 л/мин, "
        "условное время пребывания τ = V/v0 = 2 мин, начальная концентрация CA0 = 30 моль/л. "
        "Реальные концентрации рассчитывали по данным программы: CA = CA0·(1−X), CR = CA0·Y, "
        "CS = CA0·Z; производительность по целевому продукту ΠR = v0·CR = 3·CR (кмоль/ч); "
        "интегральная селективность SR = Y/X.\n\n"
        "По степени превращения X и по производительности ΠR во всех трёх кинетических случаях "
        "эффективнее РИВ. При том же объёме и расходе в вытеснении средняя концентрация реагента A "
        "вдоль аппарата выше, чем выходная, поэтому скорость целевой стадии A → R больше. "
        "В РИС-н состав во всём объёме равен выходному: реакция идёт при уже низкой CA, скорость "
        "целевой стадии падает. При τ = 2 мин получено: случай 1 — X = 0,546 (РИВ) против 0,413 "
        "(РИС-н), ΠR = 30,46 против 20,66 кмоль/ч; случай 2 — X = 0,699 против 0,545, "
        "ΠR = 40,00 против 27,27 кмоль/ч; случай 3 — X = 0,699 против 0,545, "
        "ΠR = 50,84 против 36,95 кмоль/ч.\n\n"
        "По селективности РИВ также предпочтителен: SR = 0,620 против 0,556 (случай 1), "
        "0,636 против 0,556 (случай 2), 0,808 против 0,753 (случай 3). В РИС-н интегральная "
        "селективность совпадает с дифференциальной на выходе, так как весь объём работает "
        "при выходном составе. В РИВ целевой продукт образуется при более высокой CA на начальном "
        "участке, доля побочного превращения R → S меньше.\n\n"
        "Среди кинетических вариантов наименее выгоден случай 1 (n1 = 2, n2 = 1): второй порядок "
        "по A замедляет целевую реакцию по мере расходования реагента, поэтому X и ΠR минимальны. "
        "Случаи 2 и 3 имеют одинаковый первый порядок стадии A → R, поэтому степени превращения "
        "в каждом типе реактора совпадают. Наилучший результат даёт случай 3 (n2 = 2): побочная "
        "реакция второго порядка по R замедлена при малых CR, селективность и выход целевого "
        "продукта максимальны (в РИВ CR = 16,95 моль/л, ΠR = 50,84 кмоль/ч).\n\n"
        "Итог: при заданных V и v0 процесс целесообразнее вести в реакторе идеального вытеснения; "
        "оптимальная кинетика из трёх рассмотренных — случай 3. График Ci(τ) построен для случая 1: "
        "CA монотонно снижается, CR выходит на пологий максимум к τ ≈ 1,8–2 мин, CS непрерывно растёт — "
        "типичная картина последовательной схемы A → R → S."
    )
    ws2.merge_cells("A12:G28")
    c = ws2["A12"]
    c.value = text
    c.alignment = Alignment(wrap_text=True, vertical="top")
    c.font = Font(name="Calibri", size=12)
    ws2.row_dimensions[12].height = 220
    set_widths(ws2, [18, 12, 14, 12, 16, 16, 14])
    ws2.page_setup.orientation = "landscape"

    # freeze
    ws.freeze_panes = "A4"
    wt.freeze_panes = "A3"

    wb.save(OUT)
    print("saved", OUT)

    preview_pngs()


def preview_pngs():
    import matplotlib

    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    def plot_one(rows, title, fname):
        tau = [r[0] for r in rows]
        ca = [CA0 * (1 - r[1]) for r in rows]
        cr = [CA0 * r[2] for r in rows]
        cs = [CA0 * r[3] for r in rows]
        fig, ax = plt.subplots(figsize=(9.5, 6.2), dpi=140)
        ax.plot(tau, ca, color="#5B9BD5", lw=2.0, marker="s", ms=8, label="CA")
        ax.plot(tau, cr, color="#ED7D31", lw=2.0, marker="*", ms=14, label="CR (целевой)")
        ax.plot(tau, cs, color="#70AD47", lw=2.0, marker="x", ms=9, mew=2, label="CS (побочный)")
        ax.set_xlabel("τ, мин", fontsize=12)
        ax.set_ylabel("Ci, моль/л", fontsize=12)
        ax.set_title(title, fontsize=13)
        ax.set_xlim(0, 2)
        ax.set_ylim(0, 32)
        ax.grid(True, ls="-", color="#D0D0D0")
        ax.legend(
            loc="upper center",
            bbox_to_anchor=(0.5, -0.16),
            ncol=3,
            frameon=True,
            fancybox=False,
            borderaxespad=0.4,
            fontsize=11,
        )
        fig.tight_layout()
        fig.subplots_adjust(bottom=0.20)
        fig.savefig(PREVIEW_DIR / fname, bbox_inches="tight")
        plt.close(fig)

    plot_one(PFR[1], "РИВ, случай 1 — маркеры как в Kniga1 (квадрат / звезда / x)", "graf_riv_sluchaj1.png")
    plot_one(cstr_case1_profile(), "РИС-н, случай 1 — маркеры как в Kniga1 (квадрат / звезда / x)", "graf_ris_sluchaj1.png")
    print("pngs ok")


if __name__ == "__main__":
    build()
