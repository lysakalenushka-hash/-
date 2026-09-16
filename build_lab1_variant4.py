#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Отчёт лабораторной №1, вариант 4 — схема пользователя без правок, остальное по примеру."""

from __future__ import annotations

import shutil
from pathlib import Path

import matplotlib.pyplot as plt
import matplotlib.patches as mp
from matplotlib.patches import FancyBboxPatch
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

OUT_DIR = Path("лаба1_вариант4")
DOCX = Path("Лабораторная_работа_1_гидравлика_вариант_4.docx")
USER_SCHEME = Path(
    "/home/ubuntu/.cursor/projects/workspace/assets/e7e088f7-1bd5-40fb-9fad-cd36bb97a557.png"
)


def set_run_font(run, name="Times New Roman", size=14, bold=False, italic=False):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic


def add_p(doc, text, *, size=14, bold=False, italic=False, align="left", space_after=6, space_before=0):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.line_spacing = 1.15
    if align == "center":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "right":
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    elif align == "justify":
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, italic=italic)
    return p


def shade_header(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.makeelement(qn("w:shd"), {qn("w:fill"): "1F4E79", qn("w:val"): "clear"})
    tcPr.append(shd)


def draw_matrix(path: Path) -> None:
    cols = [
        "V1", "V2", "V3", "V4", "V5", "V6", "V7", "V8",
        "P7", "P8", "P9", "P10", "P11", "P12",
        "H1", "H2", "H3",
    ]
    rows = ["1", "2", "3", "4", "5", "6", "7", "8", "9'", "10'", "11'", "12", "13", "14", "15", "16", "17"]
    det = [0, 1, 2, 3, 4, 5, 6, 7, 14, 15, 16, 11, 8, 12, 9, 13, 10]
    used = {
        0: [0, 9],
        1: [1, 10],
        2: [2, 8, 9],
        3: [3, 9, 10],
        4: [4, 8],
        5: [5, 9],
        6: [6, 9],
        7: [7, 10],
        8: [2, 4, 14],
        9: [0, 2, 3, 5, 6, 15],
        10: [1, 3, 7, 16],
        11: [11, 14],
        12: [8, 11, 14],
        13: [12, 15],
        14: [9, 12, 15],
        15: [13, 16],
        16: [10, 13, 16],
    }
    step_no = [8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 1, 2, 3, 4, 5, 6]

    fig, ax = plt.subplots(figsize=(13.8, 8.2))
    ax.set_xlim(-1.7, len(cols) + 2.3)
    ax.set_ylim(-1.0, len(rows) + 2.2)
    ax.axis("off")
    ax.set_title("Информационная матрица системы уравнений", fontsize=13)
    ax.text(-0.9, len(rows) + 0.55, "№ ур.", ha="center", va="center", fontsize=8, weight="bold")
    ax.text(len(cols) + 1.15, len(rows) + 0.55, "№ шага", ha="center", va="center", fontsize=8, weight="bold")
    for j, c in enumerate(cols):
        ax.text(j + 0.5, len(rows) + 0.55, c, ha="center", va="center", fontsize=8, weight="bold")
        ax.plot([j, j], [0, len(rows) + 1.05], color="#cccccc", lw=0.45)
    ax.plot([0, len(cols)], [len(rows) + 1.05, len(rows) + 1.05], "k", lw=0.8)
    ax.plot([0, len(cols)], [0, 0], "k", lw=0.8)

    for i, num in enumerate(rows):
        y = len(rows) - i - 0.5
        ax.text(-0.9, y, num, ha="center", va="center", fontsize=8)
        for j in used[i]:
            ax.add_patch(plt.Circle((j + 0.5, y), 0.15, fill=False, lw=1.05, ec="k"))
        jdet = det[i]
        ax.plot(
            [jdet + 0.5, jdet + 0.72, jdet + 0.5, jdet + 0.28, jdet + 0.5],
            [y + 0.20, y, y - 0.20, y, y + 0.20],
            "k",
            lw=1.2,
        )
        ax.text(len(cols) + 1.15, y, str(step_no[i]), ha="center", va="center", fontsize=8)

    ax.text(
        len(cols) / 2,
        -0.5,
        "○ — переменная входит в уравнение;  ◇ — определяемая переменная.  "
        "P1–P6, ki, Si, Hgi, PN, ρ, g — известны.",
        ha="center",
        fontsize=8,
    )
    fig.tight_layout()
    fig.savefig(path, dpi=170, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def box(ax, x, y, w, h, text, fs=9):
    ax.add_patch(
        FancyBboxPatch(
            (x, y), w, h,
            boxstyle="round,pad=0.02,rounding_size=0.08",
            fc="#f4f7fb", ec="k", lw=1.15,
        )
    )
    ax.text(x + w / 2, y + h / 2, text, ha="center", va="center", fontsize=fs)


def draw_flow(path: Path) -> None:
    fig, ax = plt.subplots(figsize=(8.6, 10.6))
    ax.set_xlim(0, 9)
    ax.set_ylim(0, 10.8)
    ax.axis("off")
    ax.set_title("Блок-схема алгоритма расчётов стационарного режима", fontsize=12)

    def arr(x, y0, y1):
        ax.annotate("", xy=(x, y1), xytext=(x, y0), arrowprops=dict(arrowstyle="->", lw=1.25))

    cx, w = 2.55, 3.9
    ax.add_patch(mp.FancyBboxPatch((cx, 9.85), w, 0.65, boxstyle="round,pad=0.03", fc="#e8f5e9", ec="k"))
    ax.text(cx + w / 2, 10.17, "Старт", ha="center", va="center", fontsize=11, weight="bold")

    seq = [
        (8.15, 1.15, "Ввод: P1–P6, k1–k8,\nS, Hg, PN, ρ, g, H1(0), H2(0), H3(0)"),
        (6.90, 0.80, "12–13. P10, P7  (ёмкость 1)"),
        (5.75, 0.80, "14–15. P11, P8  (ёмкость 2)"),
        (4.60, 0.80, "16–17. P12, P9  (ёмкость 3)"),
        (3.40, 0.85, "1–8. Скорости V1…V8\nпо ур. Бернулли"),
        (2.15, 0.90, "9'–11'. Невязки балансов\nкоррекция H1, H2, H3"),
    ]
    prev = 9.85
    for y, h, text in seq:
        box(ax, cx, y, w, h, text, 9)
        arr(4.5, prev, y + h)
        prev = y

    dy, dx = 0.40, 0.92
    y0 = 0.18
    ax.add_patch(
        mp.Polygon(
            [[4.5, y0 + 2 * dy], [4.5 + dx, y0 + dy], [4.5, y0], [4.5 - dx, y0 + dy]],
            closed=True, fc="#fff8e1", ec="k", lw=1.15,
        )
    )
    ax.text(4.5, y0 + dy, "|f| < ε ?", ha="center", va="center", fontsize=9)
    arr(4.5, 2.15, y0 + 2 * dy)

    ax.add_patch(mp.FancyBboxPatch((6.6, 0.15), 2.1, 0.78, boxstyle="round,pad=0.03", fc="#e8f5e9", ec="k"))
    ax.text(7.65, 0.54, "Вывод H, V, P\nСтоп", ha="center", va="center", fontsize=9)
    ax.annotate("", xy=(6.6, 0.54), xytext=(4.5 + dx, y0 + dy), arrowprops=dict(arrowstyle="->", lw=1.2))
    ax.text(6.15, 0.88, "да", fontsize=8)
    ax.annotate(
        "нет",
        xy=(6.55, 8.55),
        xytext=(4.5 - dx, y0 + dy),
        fontsize=8,
        arrowprops=dict(arrowstyle="->", lw=1.15, connectionstyle="arc3,rad=-0.42"),
    )
    fig.tight_layout()
    fig.savefig(path, dpi=170, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def build_doc(scheme: Path, matrix: Path, flow: Path) -> None:
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Cm(2)
    sec.bottom_margin = Cm(2)
    sec.left_margin = Cm(2.5)
    sec.right_margin = Cm(1.5)

    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(14)
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")

    for t in [
        "Министерство науки и высшего образования Российской Федерации",
        "Федеральное государственное бюджетное образовательное учреждение высшего образования",
        "«Российский химико-технологический университет им. Д.И. Менделеева»",
        "Кафедра информатики и компьютерного проектирования",
    ]:
        add_p(doc, t, size=12, align="center", space_after=0)

    add_p(doc, "", size=12, space_after=16)
    add_p(doc, "ОТЧЕТ", size=16, bold=True, align="center", space_after=8)
    add_p(doc, "По лабораторной работе №1 на тему:", size=14, bold=True, align="center", space_after=4)
    add_p(
        doc,
        "«Моделирование простой гидравлической системы в статическом режиме»",
        size=14,
        align="center",
        space_after=8,
    )
    add_p(doc, "Вариант – 4", size=14, bold=True, align="center", space_after=16)
    add_p(doc, "Выполнил: ______________________________", size=14, align="right", space_after=2)
    add_p(doc, "Группа: ________________________________", size=14, align="right", space_after=2)
    add_p(doc, "Проверила: _____________________________", size=14, align="right", space_after=16)
    add_p(doc, "г. Москва, 2026", size=14, align="center", space_after=16)

    add_p(doc, "Схематическое изображение гидравлической модели с закрытыми ёмкостями:", size=14, bold=True)
    doc.add_picture(str(scheme), width=Cm(16.2))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    add_p(doc, "Условные обозначения:", size=14, bold=True, space_before=10)
    for line in [
        "P1, P2 – давление жидкости на входе в систему;",
        "P3, P4, P5, P6 – давление жидкости на выходе из системы;",
        "P7, P8, P9 – давление жидкости на дно ёмкости;",
        "P10, P11, P12 – давление газа внутри ёмкости;",
        "V1 – V8 – скорость истечения жидкости через клапан;",
        "k1 – k8 – коэффициент пропускной способности клапана;",
        "H1, H2, H3 – высоты накапливаемой жидкости в ёмкости;",
        "Hg1, Hg2, Hg3 (H1G, H2G, H3G) – геометрическая высота ёмкости.",
    ]:
        add_p(doc, line, space_after=2)

    add_p(doc, "Допущения", size=14, bold=True, space_before=12)
    for line in [
        "1. Газ идеален.",
        "2. Во всех трубах протекает однофазный поток жидкости, температура которого одинакова на всех участках.",
        "3. Форма закрытой ёмкости цилиндрическая с площадью поперечного сечения S.",
        "4. Все трубы располагаются на одном уровне, в системе нет рециклических потоков, или рециклов, не учитываются местные сопротивления и перепады давлений в трубах, т.е. рассматриваются так называемые короткие трубопроводы.",
        "5. В емкостях, не занятых жидкостью, давление газа PN · Hgi / (Hgi − Hi).",
        "6. Системы включают только клапаны (вентили) с постоянными, не изменяющимися коэффициентами пропускной способности и закрытые емкости.",
    ]:
        add_p(doc, line, align="justify", space_after=3)

    add_p(doc, "Построение системы уравнений математического описания гидравлической системы:", size=14, bold=True, space_before=12)
    add_p(doc, "ур. Бернулли:  V = k * sign (Pвх – Pвых) * √|Pвх – Pвых|", align="center", italic=True)
    add_p(doc, "sign – функция знака: −1, 0, +1", align="center", space_after=8)
    add_p(doc, "ур. Паскаля:  Pж = Pг + ρ*g*H", align="center", italic=True, space_after=10)

    add_p(doc, "Определение скорости потоков жидкости:", size=14, bold=True)
    for line in [
        "1)  V1 = k1 * sign (P1 – P8) * √|P1 – P8|",
        "2)  V2 = k2 * sign (P2 – P9) * √|P2 – P9|",
        "3)  V3 = k3 * sign (P7 – P8) * √|P7 – P8|",
        "4)  V4 = k4 * sign (P8 – P9) * √|P8 – P9|",
        "5)  V5 = k5 * sign (P7 – P3) * √|P7 – P3|",
        "6)  V6 = k6 * sign (P8 – P4) * √|P8 – P4|",
        "7)  V7 = k7 * sign (P8 – P5) * √|P8 – P5|",
        "8)  V8 = k8 * sign (P9 – P6) * √|P9 – P6|",
    ]:
        add_p(doc, line, space_after=3)

    add_p(doc, "Уравнения массового баланса:", size=14, bold=True, space_before=10)
    add_p(doc, "9’)  (V3 – V5) / S1 = (H1(t(k)) – H1(t(0))) / Δt = f9", space_after=2)
    add_p(doc, "9*)  H1(t(0)) = H1(0)", space_after=6)
    add_p(doc, "10’)  (V1 + V3 – V4 – V6 – V7) / S2 = (H2(t(k)) – H2(t(0))) / Δt = f10", space_after=2)
    add_p(doc, "10*)  H2(t(0)) = H2(0)", space_after=6)
    add_p(doc, "11’)  (V2 + V4 – V8) / S3 = (H3(t(k)) – H3(t(0))) / Δt = f11", space_after=2)
    add_p(doc, "11*)  H3(t(0)) = H3(0)", space_after=8)
    add_p(
        doc,
        "Для стационарного режима правые части 9’–11’ равны нулю: dHi/dt = 0.",
        align="justify",
    )

    add_p(doc, "Определение давлений жидкости и газа в закрытых ёмкостях:", size=14, bold=True, space_before=10)
    for line in [
        "12)  P10 = PN * Hg1 / (Hg1 – H1)",
        "13)  P7  = P10 + ρ*g*H1",
        "14)  P11 = PN * Hg2 / (Hg2 – H2)",
        "15)  P8  = P11 + ρ*g*H2",
        "16)  P12 = PN * Hg3 / (Hg3 – H3)",
        "17)  P9  = P12 + ρ*g*H3",
    ]:
        add_p(doc, line, space_after=3)

    add_p(doc, "Информационная матрица системы уравнений:", size=14, bold=True, space_before=12)
    add_p(doc, "Условные обозначения:", space_after=4)
    add_p(doc, "◇ – определяемая переменная;", space_after=2)
    add_p(doc, "○ – известная переменная (входит в уравнение);", space_after=2)
    add_p(doc, "В столбце «№ шага» отражается последовательность вычислений согласно выбранному алгоритму расчетов.", align="justify")
    doc.add_picture(str(matrix), width=Cm(16.6))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    add_p(doc, "Блок-схема алгоритма расчётов стационарного режима гидравлической системы:", size=14, bold=True, space_before=12)
    doc.add_picture(str(flow), width=Cm(14.5))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    add_p(
        doc,
        "Компьютерная программа на языке MATLAB в данную лабораторную работу не входит.",
        italic=True,
        space_before=12,
    )

    doc.save(DOCX)


def main() -> None:
    OUT_DIR.mkdir(exist_ok=True)
    scheme = OUT_DIR / "схема_вариант4.png"
    shutil.copy2(USER_SCHEME, scheme)
    matrix = OUT_DIR / "информационная_матрица.png"
    flow = OUT_DIR / "блок_схема.png"
    draw_matrix(matrix)
    draw_flow(flow)
    build_doc(scheme, matrix, flow)
    print("saved", DOCX.resolve())


if __name__ == "__main__":
    main()
