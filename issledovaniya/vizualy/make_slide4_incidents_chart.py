#!/usr/bin/env python3
"""Chart for slide 4: safety incidents, high vs low engagement (Gallup Q12)."""

from pathlib import Path

import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

OUT = Path(__file__).with_name("slide4-incidents-vovlechennost.png")

TEAL = "#1A9B8A"
NAVY = "#0B1F3A"
INK = "#1A2332"
MUTED = "#5B6778"
RED = "#C0453C"
TRACK = "#EEF2F6"
WHITE = "#FFFFFF"
GOLD = "#D4A017"

BOTTOM = 100
TOP = 37


def setup_fonts():
    for p in (
        "/usr/share/fonts/truetype/macos/Inter-Regular.ttf",
        "/usr/share/fonts/truetype/macos/Inter-Bold.ttf",
        "/usr/share/fonts/truetype/macos/Inter-SemiBold.ttf",
        "/usr/share/fonts/truetype/macos/Inter-Medium.ttf",
    ):
        font_manager.fontManager.addfont(p)
    plt.rcParams["font.family"] = "Inter"
    plt.rcParams["axes.unicode_minus"] = False


def round_rect(ax, x, y, w, h, color, radius=0.04, z=2):
    p = FancyBboxPatch(
        (x, y),
        w,
        h,
        boxstyle=f"round,pad=0,rounding_size={radius}",
        linewidth=0,
        facecolor=color,
        zorder=z,
    )
    ax.add_patch(p)
    return p


def main():
    setup_fonts()
    fig, ax = plt.subplots(figsize=(6.6, 5.9), dpi=240)
    fig.patch.set_facecolor(WHITE)
    ax.set_facecolor(WHITE)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis("off")

    ax.text(6, 94, "Инциденты по безопасности", fontsize=17, color=NAVY, fontweight="bold", va="center")
    ax.text(
        6,
        86.5,
        "Сравнение команд: нижний квартиль вовлечённости = 100",
        fontsize=10,
        color=MUTED,
        va="center",
    )

    # plot area
    y0, y1 = 18, 72
    plot_h = y1 - y0
    col_w = 22
    x_low, x_high = 18, 58

    # faint grid lines
    for frac, lab in ((0.25, "25"), (0.5, "50"), (0.75, "75"), (1.0, "100")):
        yy = y0 + plot_h * frac
        ax.plot([12, 86], [yy, yy], color=TRACK, lw=1.1, zorder=1)
        ax.text(10.5, yy, lab, fontsize=7.5, color="#9AA3B0", ha="right", va="center")

    h_low = plot_h * (BOTTOM / 100)
    h_high = plot_h * (TOP / 100)
    round_rect(ax, x_low, y0, col_w, h_low, RED, radius=1.6, z=3)
    round_rect(ax, x_high, y0, col_w, h_high, TEAL, radius=1.6, z=3)

    ax.text(x_low + col_w / 2, y0 + h_low + 3.2, "100", fontsize=18, color=RED, fontweight="bold", ha="center", va="center")
    ax.text(x_high + col_w / 2, y0 + h_high + 3.2, "37", fontsize=18, color=TEAL, fontweight="bold", ha="center", va="center")

    ax.text(x_low + col_w / 2, 12.2, "Низкая", fontsize=11, color=INK, fontweight="bold", ha="center")
    ax.text(x_low + col_w / 2, 7.4, "вовлечённость", fontsize=9, color=MUTED, ha="center")
    ax.text(x_high + col_w / 2, 12.2, "Высокая", fontsize=11, color=INK, fontweight="bold", ha="center")
    ax.text(x_high + col_w / 2, 7.4, "вовлечённость", fontsize=9, color=MUTED, ha="center")

    # −63% badge to the right of the drop
    ax.add_patch(
        FancyArrowPatch(
            (x_low + col_w + 1.5, y0 + h_low - 1),
            (x_high - 1.5, y0 + h_high + 4),
            arrowstyle="-|>",
            mutation_scale=12,
            lw=1.8,
            color=TEAL,
            zorder=4,
        )
    )
    round_rect(ax, 82.2, 40, 14.5, 11.5, TEAL, radius=1.4, z=5)
    ax.text(89.45, 45.7, "−63%", fontsize=13.5, color=WHITE, fontweight="bold", ha="center", va="center", zorder=6)

    ax.text(6, 2.4, "Gallup Q12 Meta-Analysis, 11th ed., 2024", fontsize=8.5, color=MUTED, va="center")

    fig.subplots_adjust(left=0, right=1, top=1, bottom=0)
    fig.savefig(OUT, dpi=240, facecolor=WHITE)
    plt.close()
    print("wrote", OUT)


if __name__ == "__main__":
    main()
