#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Отчёт лабораторной №1, вариант 4 — до MATLAB."""

from __future__ import annotations

from pathlib import Path

import matplotlib.pyplot as plt
import matplotlib.patches as mp
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

OUT_DIR = Path("лаба1_вариант4")
DOCX = Path("Лабораторная_работа_1_гидравлика_вариант_4.docx")


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


def valve(ax, x, y, w=0.55, h=0.32, vertical=False):
    if vertical:
        xs = [x, x + h / 2, x, x - h / 2, x]
        ys = [y - w / 2, y, y + w / 2, y, y - w / 2]
    else:
        xs = [x - w / 2, x, x + w / 2, x, x - w / 2]
        ys = [y, y + h / 2, y, y - h / 2, y]
    ax.plot(xs, ys, "k-", lw=1.4)
    ax.plot([min(xs), max(xs)], [y, y] if not vertical else [min(ys), max(ys)], "k-", lw=1.0)


def tank(ax, x, y, w, h, liquid=0.42):
    ax.add_patch(mp.Rectangle((x, y), w, h, fill=False, lw=1.6, ec="k"))
    ax.plot([x, x + w], [y + h * liquid, y + h * liquid], "k-", lw=1.2)


def draw_scheme(path: Path) -> None:
    fig, ax = plt.subplots(figsize=(11.5, 8.2))
    ax.set_xlim(-0.4, 11.2)
    ax.set_ylim(-0.3, 8.4)
    ax.set_aspect("equal")
    ax.axis("off")
    ax.set_title("Вариант 4. Гидравлическая система с тремя закрытыми ёмкостями", fontsize=13, pad=8)

    # tanks: top, mid, bot
    tank(ax, 4.55, 5.55, 1.7, 2.15, 0.40)
    tank(ax, 4.35, 2.85, 1.55, 2.05, 0.42)
    tank(ax, 4.35, 0.35, 1.55, 1.95, 0.45)

    ax.text(4.72, 6.95, "P10", fontsize=9)
    ax.text(4.72, 6.05, "P7", fontsize=9)
    ax.text(6.35, 7.35, "Hg1", fontsize=9)
    ax.text(3.55, 6.35, "H1", fontsize=9)
    ax.text(4.50, 7.85, "ёмкость 1", fontsize=8)

    ax.text(4.52, 4.20, "P11", fontsize=9)
    ax.text(4.55, 3.35, "P8", fontsize=9)
    ax.text(6.05, 4.55, "Hg2", fontsize=9)
    ax.text(3.45, 3.70, "H2", fontsize=9)
    ax.text(4.35, 5.05, "ёмкость 2", fontsize=8)

    ax.text(4.52, 1.65, "P12", fontsize=9)
    ax.text(4.55, 0.80, "P9", fontsize=9)
    ax.text(6.05, 2.00, "Hg3", fontsize=9)
    ax.text(3.45, 1.20, "H3", fontsize=9)
    ax.text(4.35, 2.45, "ёмкость 3", fontsize=8)

    # top tank left-right pipes
    ax.plot([2.35, 4.55], [6.40, 6.40], "k-", lw=1.3)
    ax.plot([6.25, 9.55], [6.40, 6.40], "k-", lw=1.3)
    valve(ax, 3.35, 6.40)
    valve(ax, 7.55, 6.40)
    ax.text(3.05, 6.72, "V3, k3", fontsize=8)
    ax.text(7.25, 6.72, "V5, k5", fontsize=8)
    ax.annotate("", xy=(9.95, 6.40), xytext=(9.55, 6.40), arrowprops=dict(arrowstyle="->", lw=1.2))
    ax.text(10.0, 6.55, "P3", fontsize=10)

    # стояк: разрыв на пересечении со средней линией (нет врезки)
    ax.plot([2.35, 2.35], [6.40, 4.12], "k-", lw=1.3)
    ax.plot([2.35, 2.35], [3.64, 1.32], "k-", lw=1.3)
    hop = mp.Arc((2.35, 3.88), 0.42, 0.42, theta1=0, theta2=180, lw=1.3, color="k")
    ax.add_patch(hop)
    valve(ax, 2.35, 2.05, vertical=True)
    ax.text(1.05, 2.05, "V4, k4", fontsize=8)
    ax.text(2.55, 4.85, "P13", fontsize=9)

    # mid left inlet
    ax.plot([0.15, 4.35], [3.88, 3.88], "k-", lw=1.3)
    valve(ax, 1.35, 3.88)
    ax.text(0.95, 4.18, "V1, k1", fontsize=8)
    ax.annotate("", xy=(0.15, 3.88), xytext=(-0.05, 3.88), arrowprops=dict(arrowstyle="-", lw=0))
    ax.text(-0.28, 4.05, "P1", fontsize=10)
    ax.annotate("", xy=(0.55, 3.88), xytext=(0.15, 3.88), arrowprops=dict(arrowstyle="->", lw=1.2))

    # mid right two outlets
    ax.plot([5.90, 9.55], [3.88, 3.88], "k-", lw=1.3)
    ax.plot([7.15, 7.15], [3.88, 4.55], "k-", lw=1.3)
    ax.plot([7.15, 9.55], [4.55, 4.55], "k-", lw=1.3)
    valve(ax, 8.25, 4.55)
    valve(ax, 8.25, 3.88)
    ax.text(8.00, 4.88, "V6, k6", fontsize=8)
    ax.text(8.00, 3.48, "V7, k7", fontsize=8)
    ax.annotate("", xy=(9.95, 4.55), xytext=(9.55, 4.55), arrowprops=dict(arrowstyle="->", lw=1.2))
    ax.annotate("", xy=(9.95, 3.88), xytext=(9.55, 3.88), arrowprops=dict(arrowstyle="->", lw=1.2))
    ax.text(10.0, 4.68, "P4", fontsize=10)
    ax.text(10.0, 3.55, "P5", fontsize=10)

    # bottom left-right
    ax.plot([0.15, 4.35], [1.32, 1.32], "k-", lw=1.3)
    ax.plot([5.90, 9.55], [1.32, 1.32], "k-", lw=1.3)
    valve(ax, 1.35, 1.32)
    valve(ax, 8.25, 1.32)
    ax.text(0.95, 1.62, "V2, k2", fontsize=8)
    ax.text(8.00, 0.92, "V8, k8", fontsize=8)
    ax.text(-0.28, 1.48, "P2", fontsize=10)
    ax.annotate("", xy=(0.55, 1.32), xytext=(0.15, 1.32), arrowprops=dict(arrowstyle="->", lw=1.2))
    ax.annotate("", xy=(9.95, 1.32), xytext=(9.55, 1.32), arrowprops=dict(arrowstyle="->", lw=1.2))
    ax.text(10.0, 1.45, "P6", fontsize=10)

    # join downcomer to bottom header
    ax.plot([2.35, 2.35], [1.32, 1.32], "k-", lw=1.3)

    note = (
        "Стояк от ёмкости 1 пересекает линию ёмкости 2 без врезки\n"
        "и соединяется с линией ёмкости 3 (вентили V3 и V4 последовательно)."
    )
    ax.text(0.1, -0.15, note, fontsize=8, color="#333333")

    fig.tight_layout()
    fig.savefig(path, dpi=180, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def draw_matrix(path: Path) -> None:
    # columns: variables
    cols = [
        "V1", "V2", "V3", "V4", "V5", "V6", "V7", "V8",
        "P7", "P8", "P9", "P10", "P11", "P12", "P13",
        "H1", "H2", "H3",
    ]
    # known boundary not in matrix as unknowns: P1-P6, k, S, Hg, PN
    rows = [
        ("1", "V1", 0),
        ("2", "V2", 1),
        ("3", "V3", 2),
        ("4", "V4", 3),
        ("5", "V5", 4),
        ("6", "V6", 5),
        ("7", "V7", 6),
        ("8", "V8", 7),
        ("9", "V3=V4", None),
        ("10", "баланс 1", None),
        ("11", "баланс 2", None),
        ("12", "баланс 3", None),
        ("13", "P10", 11),
        ("14", "P7", 8),
        ("15", "P11", 12),
        ("16", "P8", 9),
        ("17", "P12", 13),
        ("18", "P9", 10),
        ("19", "P13", 14),
    ]
    # incidence: which unknowns appear in equation (for display we mark determined variable)
    fig, ax = plt.subplots(figsize=(14.2, 8.6))
    ax.set_xlim(-1.6, len(cols) + 2.2)
    ax.set_ylim(-1.2, len(rows) + 2.4)
    ax.axis("off")
    ax.set_title("Информационная матрица системы уравнений (вариант 4)", fontsize=13)

    # header
    ax.text(-0.9, len(rows) + 0.55, "№ ур.", ha="center", va="center", fontsize=8, weight="bold")
    ax.text(len(cols) + 1.1, len(rows) + 0.55, "шаг", ha="center", va="center", fontsize=8, weight="bold")
    for j, c in enumerate(cols):
        ax.text(j + 0.5, len(rows) + 0.55, c, ha="center", va="center", fontsize=8, weight="bold")
        ax.plot([j, j], [0, len(rows) + 1.1], color="#bbbbbb", lw=0.5)
    ax.plot([0, len(cols)], [len(rows) + 1.1, len(rows) + 1.1], "k", lw=0.8)
    ax.plot([0, len(cols)], [0, 0], "k", lw=0.8)

    # determined variable per equation (diamond = solved here, circle = used known/previously found)
    det = {
        0: 0, 1: 1, 2: 2, 3: 3, 4: 4, 5: 5, 6: 6, 7: 7,
        8: 2,  # V3=V4 uses V3,V4; mark V4 as closing
        9: 15, 10: 16, 11: 17,  # H from balances in static iteration — actually H are iterated
        12: 11, 13: 8, 14: 12, 15: 9, 16: 13, 17: 10, 18: 14,
    }
    steps = [8, 9, 10, 11, 12, 13, 14, 15, 7, 16, 17, 18, 1, 2, 3, 4, 5, 6, 7]
    # better step order matching algorithm:
    step_no = {
        12: 1, 13: 2, 14: 3, 15: 4, 16: 5, 17: 6, 18: 7,
        0: 8, 1: 9, 2: 10, 3: 11, 4: 12, 5: 13, 6: 14, 7: 15,
        8: 11, 9: 16, 10: 17, 11: 18,
    }

    used = {
        0: [0, 9],  # V1, P8 (and P1 known)
        1: [1, 10],
        2: [2, 8, 14],
        3: [3, 10, 14],
        4: [4, 8],
        5: [5, 9],
        6: [6, 9],
        7: [7, 10],
        8: [2, 3],
        9: [2, 4, 15],
        10: [0, 5, 6, 16],
        11: [1, 3, 7, 17],
        12: [11, 15],
        13: [8, 11, 15],
        14: [12, 16],
        15: [9, 12, 16],
        16: [13, 17],
        17: [10, 13, 17],
        18: [8, 10, 14],
    }

    for i, (num, _name, _) in enumerate(rows):
        y = len(rows) - i - 0.5
        ax.text(-0.9, y, num, ha="center", va="center", fontsize=8)
        ax.plot([-1.5, len(cols) + 0.05], [len(rows) - i, len(rows) - i], color="#dddddd", lw=0.4)
        for j in used.get(i, []):
            circ = plt.Circle((j + 0.5, y), 0.16, fill=False, lw=1.1, ec="k")
            ax.add_patch(circ)
        jdet = det.get(i)
        if jdet is not None:
            d = mp.FancyBboxPatch(
                (jdet + 0.5 - 0.18, y - 0.18),
                0.36,
                0.36,
                boxstyle="darrow,pad=0.02",
                mutation_aspect=1,
                fill=False,
                lw=0,
            )
            # diamond
            ax.plot(
                [jdet + 0.5, jdet + 0.5 + 0.22, jdet + 0.5, jdet + 0.5 - 0.22, jdet + 0.5],
                [y + 0.22, y, y - 0.22, y, y + 0.22],
                "k",
                lw=1.25,
            )
        ax.text(len(cols) + 1.1, y, str(step_no.get(i, "")), ha="center", va="center", fontsize=8)

    ax.text(
        len(cols) / 2,
        -0.55,
        "○ — переменная входит в уравнение; ◇ — переменная, определяемая из данного уравнения. "
        "P1–P6, ki, Si, Hgi, PN, ρ, g — известны (в матрицу не включены).",
        ha="center",
        fontsize=8,
    )
    fig.tight_layout()
    fig.savefig(path, dpi=170, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def box(ax, x, y, w, h, text, fs=8):
    ax.add_patch(
        FancyBboxPatch(
            (x, y),
            w,
            h,
            boxstyle="round,pad=0.02,rounding_size=0.08",
            fill=True,
            facecolor="#f4f7fb",
            edgecolor="k",
            lw=1.15,
        )
    )
    ax.text(x + w / 2, y + h / 2, text, ha="center", va="center", fontsize=fs, wrap=True)


def draw_flow(path: Path) -> None:
    fig, ax = plt.subplots(figsize=(8.8, 11.2))
    ax.set_xlim(0, 9)
    ax.set_ylim(0, 11.4)
    ax.axis("off")
    ax.set_title("Блок-схема алгоритма расчёта стационарного режима (вариант 4)", fontsize=12)

    def arr(x, y0, y1):
        ax.annotate("", xy=(x, y1), xytext=(x, y0), arrowprops=dict(arrowstyle="->", lw=1.25))

    cx, w = 2.7, 3.6
    ax.add_patch(mp.FancyBboxPatch((cx, 10.35), w, 0.7, boxstyle="round,pad=0.03", fc="#e8f5e9", ec="k"))
    ax.text(cx + w / 2, 10.7, "Старт", ha="center", va="center", fontsize=11, weight="bold")
    arr(4.5, 10.35, 9.85)

    seq = [
        (8.55, 1.2, "Ввод данных: P1–P6, k1–k8,\nS, Hg, PN, ρ, g, H1(0), H2(0), H3(0)"),
        (7.15, 0.85, "P10, P7  по ёмкости 1"),
        (5.95, 0.85, "P11, P8  по ёмкости 2"),
        (4.75, 0.85, "P12, P9  по ёмкости 3"),
        (3.50, 0.90, "P13 из условия V3 = V4"),
        (2.25, 0.90, "Скорости V1…V8 (Бернулли)"),
        (1.45, 0.95, "Невязки балансов f1, f2, f3\nи коррекция H1, H2, H3"),
    ]
    prev_bottom = 10.35
    for y, h, text in seq:
        box(ax, cx, y, w, h, text, 9)
        arr(4.5, prev_bottom, y + h)
        prev_bottom = y

    # diamond
    dy, dx = 0.42, 0.95
    y0 = 0.08
    ax.add_patch(
        mp.Polygon(
            [[4.5, y0 + 2 * dy], [4.5 + dx, y0 + dy], [4.5, y0], [4.5 - dx, y0 + dy]],
            closed=True,
            fc="#fff8e1",
            ec="k",
            lw=1.15,
        )
    )
    ax.text(4.5, y0 + dy, "|f| < ε ?", ha="center", va="center", fontsize=9)
    arr(4.5, 1.45, y0 + 2 * dy)

    ax.add_patch(mp.FancyBboxPatch((6.55, 0.12), 2.15, 0.80, boxstyle="round,pad=0.03", fc="#e8f5e9", ec="k"))
    ax.text(7.62, 0.52, "Вывод H, V, P\nСтоп", ha="center", va="center", fontsize=9)
    ax.annotate("", xy=(6.55, 0.52), xytext=(4.5 + dx, y0 + dy), arrowprops=dict(arrowstyle="->", lw=1.2))
    ax.text(6.05, 0.88, "да", fontsize=8)

    ax.annotate(
        "нет",
        xy=(6.55, 9.05),
        xytext=(4.5 - dx, y0 + dy),
        fontsize=8,
        arrowprops=dict(arrowstyle="->", lw=1.15, connectionstyle="arc3,rad=-0.45"),
    )

    fig.tight_layout()
    fig.savefig(path, dpi=170, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def shade_header(cell):
    tc = cell._tePr if False else cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = tcPr.makeelement(qn("w:shd"), {qn("w:fill"): "1F4E79", qn("w:val"): "clear"})
    tcPr.append(shd)


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

    add_p(doc, "", size=12, space_after=18)
    add_p(doc, "ОТЧЁТ", size=16, bold=True, align="center", space_after=8)
    add_p(doc, "по лабораторной работе № 1", size=14, bold=True, align="center", space_after=4)
    add_p(
        doc,
        "по курсу «Моделирование химико-технологических процессов»",
        size=14,
        align="center",
        space_after=8,
    )
    add_p(
        doc,
        "Тема: «Моделирование простой гидравлической системы в статическом режиме»",
        size=14,
        align="center",
        space_after=8,
    )
    add_p(doc, "Вариант – 4", size=14, bold=True, align="center", space_after=16)

    add_p(doc, "Выполнил: ______________________________", size=14, align="right", space_after=2)
    add_p(doc, "Группа: ________________________________", size=14, align="right", space_after=2)
    add_p(doc, "Проверил: ______________________________", size=14, align="right", space_after=18)
    add_p(doc, "г. Москва, 2026", size=14, align="center", space_after=18)

    add_p(doc, "1. Схематическое изображение гидравлической модели", size=14, bold=True, space_before=8)
    add_p(
        doc,
        "Схема соответствует выданному варианту 4: три закрытые цилиндрические ёмкости, "
        "два входных коллектора слева, четыре выхода справа. Стояк от верхней ёмкости "
        "пересекает линию средней ёмкости без гидравлической врезки и соединяется с линией нижней ёмкости.",
        align="justify",
    )
    doc.add_picture(str(scheme), width=Cm(16.5))
    last = doc.paragraphs[-1]
    last.alignment = WD_ALIGN_PARAGRAPH.CENTER

    add_p(doc, "Условные обозначения:", size=14, bold=True, space_before=10)
    legends = [
        "P1, P2 — давление жидкости на входах в систему (слева);",
        "P3, P4, P5, P6 — давление жидкости на выходах из системы (справа);",
        "P7, P8, P9 — давление жидкости на дно ёмкостей 1, 2 и 3;",
        "P10, P11, P12 — давление газа в свободном объёме ёмкостей 1, 2 и 3;",
        "P13 — давление в узле стояка между последовательными вентилями V3 и V4;",
        "V1 – V8 — объёмные скорости истечения через вентили (м³/с);",
        "k1 – k8 — коэффициенты пропускной способности вентилей;",
        "H1, H2, H3 — высоты уровня жидкости в ёмкостях, м;",
        "Hg1, Hg2, Hg3 (H1G, H2G, H3G) — геометрические высоты ёмкостей, м;",
        "S1, S2, S3 — площади поперечного сечения ёмкостей, м²;",
        "PN — начальное (атмосферное) давление газа при пустой ёмкости, Па;",
        "ρ — плотность жидкости, кг/м³;  g — ускорение свободного падения, м/с².",
    ]
    for line in legends:
        add_p(doc, line, align="justify", space_after=2)

    add_p(doc, "2. Допущения", size=14, bold=True, space_before=12)
    assumptions = [
        "1. Газ в свободном объёме ёмкостей идеален; процесс сжатия/расширения газа изотермический.",
        "2. Во всех трубах — однофазный поток жидкости; температура одинакова на всех участках.",
        "3. Форма закрытых ёмкостей — цилиндрическая, с постоянной площадью поперечного сечения Si.",
        "4. Трубы короткие: местные сопротивления и потери давления по длине не учитываются, кроме сосредоточенных сопротивлений вентилей. Рециклы отсутствуют.",
        "5. Все горизонтальные участки расположены на одном геометрическом уровне относительно днища соответствующей ёмкости.",
        "6. В части ёмкости, не занятой жидкостью, давление газа связано с уровнем жидкости: Pi газ = PN · Hgi / (Hgi − Hi).",
        "7. Вентили имеют постоянные (не зависящие от времени) коэффициенты пропускной способности ki.",
        "8. В стояке между V3 и V4 накопление жидкости не рассматривается, поэтому в любой момент V3 = V4.",
        "9. Пересечение стояка с линией средней ёмкости — «холостое»: гидравлической связи ёмкостей 1 и 2 нет.",
    ]
    for line in assumptions:
        add_p(doc, line, align="justify", space_after=3)

    add_p(doc, "3. Система уравнений математического описания", size=14, bold=True, space_before=12)
    add_p(doc, "Уравнение Бернулли для вентиля (короткий трубопровод):", align="justify")
    add_p(
        doc,
        "Vi = ki · sign(Pвх − Pвых) · √|Pвх − Pвых|,    i = 1…8.",
        align="center",
        italic=True,
    )
    add_p(
        doc,
        "sign — функция знака: −1, 0, +1. Положительное направление потока — от «вх» к «вых» в записи ниже.",
        align="justify",
    )
    add_p(doc, "Уравнение Паскаля для столба жидкости в ёмкости:", align="justify")
    add_p(doc, "Pжидк = Pгаз + ρ · g · H.", align="center", italic=True)

    add_p(doc, "3.1. Скорости потоков через вентили", size=14, bold=True, space_before=8)
    eqs = [
        "1)  V1 = k1 · sign(P1 − P8) · √|P1 − P8|     (вход в ёмкость 2)",
        "2)  V2 = k2 · sign(P2 − P9) · √|P2 − P9|     (вход в ёмкость 3)",
        "3)  V3 = k3 · sign(P7 − P13) · √|P7 − P13|   (из ёмкости 1 в стояк)",
        "4)  V4 = k4 · sign(P13 − P9) · √|P13 − P9|   (из стояка в ёмкость 3)",
        "5)  V5 = k5 · sign(P7 − P3) · √|P7 − P3|     (выход из ёмкости 1)",
        "6)  V6 = k6 · sign(P8 − P4) · √|P8 − P4|     (верхний выход из ёмкости 2)",
        "7)  V7 = k7 · sign(P8 − P5) · √|P8 − P5|     (основной выход из ёмкости 2)",
        "8)  V8 = k8 · sign(P9 − P6) · √|P9 − P6|     (выход из ёмкости 3)",
        "9)  V3 = V4     (нет ёмкости на стояке)",
    ]
    for line in eqs:
        add_p(doc, line, space_after=3)

    add_p(doc, "3.2. Уравнения материального баланса", size=14, bold=True, space_before=8)
    add_p(
        doc,
        "Для стационарного режима производные уровней равны нулю. Для последующего динамического расчёта "
        "(лабораторная с MATLAB) те же балансы записываются как ОДУ. Ниже — оба вида.",
        align="justify",
    )
    add_p(doc, "Ёмкость 1 (верхняя):", bold=True, space_after=2)
    add_p(doc, "10)  S1 · dH1/dt = −V3 − V5     (стационар:  −V3 − V5 = 0)", space_after=2)
    add_p(doc, "10*)  H1(t0) = H1(0)", space_after=6)
    add_p(doc, "Ёмкость 2 (средняя):", bold=True, space_after=2)
    add_p(doc, "11)  S2 · dH2/dt = V1 − V6 − V7     (стационар:  V1 − V6 − V7 = 0)", space_after=2)
    add_p(doc, "11*)  H2(t0) = H2(0)", space_after=6)
    add_p(doc, "Ёмкость 3 (нижняя):", bold=True, space_after=2)
    add_p(doc, "12)  S3 · dH3/dt = V2 + V4 − V8     (стационар:  V2 + V4 − V8 = 0)", space_after=2)
    add_p(doc, "12*)  H3(t0) = H3(0)", space_after=8)

    add_p(doc, "3.3. Давления жидкости и газа в закрытых ёмкостях", size=14, bold=True)
    peqs = [
        "13)  P10 = PN · Hg1 / (Hg1 − H1)",
        "14)  P7  = P10 + ρ · g · H1",
        "15)  P11 = PN · Hg2 / (Hg2 − H2)",
        "16)  P8  = P11 + ρ · g · H2",
        "17)  P12 = PN · Hg3 / (Hg3 − H3)",
        "18)  P9  = P12 + ρ · g · H3",
        "19)  P13 находится из равенства (9) совместно с (3) и (4): при известных P7 и P9 это алгебраическое уравнение относительно P13 "
        "(P13 лежит между P7 и P9, если поток направлен из ёмкости 1 в ёмкость 3).",
    ]
    for line in peqs:
        add_p(doc, line, align="justify", space_after=3)

    add_p(
        doc,
        "Ограничения: 0 < Hi < Hgi, иначе знаменатель (Hgi − Hi) теряет смысл (ёмкость пуста или полностью залита).",
        align="justify",
        space_before=6,
    )

    add_p(doc, "4. Информационная матрица системы уравнений", size=14, bold=True, space_before=12)
    add_p(
        doc,
        "Матрица показывает, какие переменные входят в каждое уравнение и в каком порядке их можно вычислить "
        "при известных граничных давлениях P1–P6, коэффициентах ki и геометрических параметрах. "
        "Ромб — переменная, определяемая из строки; круг — переменная, уже известная или входящая в уравнение.",
        align="justify",
    )
    doc.add_picture(str(matrix), width=Cm(16.8))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    add_p(doc, "Последовательность шагов расчёта (столбец «шаг»):", bold=True, space_before=8)
    steps = [
        "Шаги 1–2. По текущему H1 вычислить P10 (газ) и P7 (жидкость) в ёмкости 1.",
        "Шаги 3–4. По H2 вычислить P11 и P8 в ёмкости 2.",
        "Шаги 5–6. По H3 вычислить P12 и P9 в ёмкости 3.",
        "Шаг 7. Из V3 = V4 найти P13.",
        "Шаги 8–15. По Бернулли найти V1…V8.",
        "Шаги 16–18. Вычислить невязки балансов ёмкостей; в стационаре подобрать H1, H2, H3 так, чтобы невязки стали нулевыми "
        "(итерации). Начальные H1(0), H2(0), H3(0) задают стартовую точку.",
    ]
    for line in steps:
        add_p(doc, line, align="justify", space_after=3)

    add_p(doc, "Структура неизвестных и известных величин", size=14, bold=True, space_before=10)
    table = doc.add_table(rows=1, cols=3)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = ["Тип", "Величины", "Количество"]
    for i, h in enumerate(hdr):
        cell = table.rows[0].cells[i]
        cell.text = h
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if p.runs:
                set_run_font(p.runs[0], size=12, bold=True)
                p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
        shade_header(cell)
    data = [
        ("Известные (граничные)", "P1, P2, P3, P4, P5, P6", "6"),
        ("Известные (параметры)", "k1…k8, S1…S3, Hg1…Hg3, PN, ρ, g", "8+3+3+3"),
        ("Искомые давления", "P7…P13", "7"),
        ("Искомые скорости", "V1…V8", "8"),
        ("Искомые уровни", "H1, H2, H3", "3"),
        ("Уравнения модели", "Бернулли (8) + связь V3=V4 + балансы (3) + Паскаль/газ (6)", "18"),
    ]
    for row in data:
        cells = table.add_row().cells
        for i, val in enumerate(row):
            cells[i].text = val
            for p in cells[i].paragraphs:
                if p.runs:
                    set_run_font(p.runs[0], size=12)

    add_p(
        doc,
        "Система замкнута: 8 скоростей + 7 давлений + 3 уровня = 18 неизвестных при стационаре "
        "(начальные Hi(0) нужны только как стартовые значения итерации или как НУ динамики).",
        align="justify",
        space_before=8,
    )

    add_p(doc, "5. Блок-схема алгоритма расчёта стационарного режима", size=14, bold=True, space_before=12)
    add_p(
        doc,
        "Алгоритм — последовательный по информационной матрице с внешним циклом по уровням Hi. "
        "Критерий останова: модуль невязок материальных балансов всех трёх ёмкостей меньше заданной точности ε. "
        "Программа MATLAB в данную работу не входит.",
        align="justify",
    )
    doc.add_picture(str(flow), width=Cm(16.5))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    add_p(doc, "6. Вывод", size=14, bold=True, space_before=12)
    add_p(
        doc,
        "Для варианта 4 построена математическая модель гидравлической системы с тремя закрытыми ёмкостями "
        "и восемью вентилями. Верхняя и нижняя ёмкости связаны стояком с двумя последовательными вентилями; "
        "средняя ёмкость гидравлически независима (линия пересекается стояком без врезки) и имеет один вход и два выхода. "
        "Модель включает уравнения Бернулли, Паскаля, изотермическое сжатие газа и материальные балансы. "
        "По информационной матрице определён порядок расчёта стационарного режима. "
        "Численное интегрирование и построение графиков H(t) выполняются на следующем этапе (MATLAB) и в отчёт № 1 не включались.",
        align="justify",
    )

    doc.save(DOCX)


def main() -> None:
    OUT_DIR.mkdir(exist_ok=True)
    scheme = OUT_DIR / "схема_вариант4.png"
    matrix = OUT_DIR / "информационная_матрица.png"
    flow = OUT_DIR / "блок_схема.png"
    draw_scheme(scheme)
    draw_matrix(matrix)
    draw_flow(flow)
    build_doc(scheme, matrix, flow)
    print("saved", DOCX.resolve())


if __name__ == "__main__":
    main()
