#!/usr/bin/env python3
"""Slide 5: illuminated cost scale for cost of inaction."""

from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
from matplotlib import font_manager
from matplotlib.colors import LinearSegmentedColormap
from matplotlib.patches import Circle, FancyBboxPatch, Rectangle

OUT = Path(__file__).with_name("slide5-shkala-poter.png")

NAVY = "#0B1F3A"
NAVY2 = "#122844"
INK = "#E8EEF5"
MUTED = "#9AA8B8"
TEAL = "#2EE0C5"
GOLD = "#F0C14B"
RED = "#FF6B5A"
WHITE = "#FFFFFF"


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


def glow_circle(ax, x, y, r, color, layers=8):
    for i in range(layers, 0, -1):
        ax.add_patch(
            Circle(
                (x, y),
                r * (1 + i * 0.38),
                facecolor=color,
                edgecolor="none",
                alpha=0.045 * i,
                zorder=3,
            )
        )
    ax.add_patch(Circle((x, y), r, facecolor=color, edgecolor="none", zorder=5))
    ax.add_patch(Circle((x, y), r * 0.38, facecolor=WHITE, edgecolor="none", alpha=0.9, zorder=6))


def main():
    setup_fonts()
    fig, ax = plt.subplots(figsize=(13.2, 4.55), dpi=220)
    fig.patch.set_facecolor(NAVY)
    ax.set_facecolor(NAVY)
    ax.set_xlim(0, 132)
    ax.set_ylim(0, 45.5)
    ax.axis("off")

    ax.add_patch(
        FancyBboxPatch(
            (1.2, 1.0),
            129.6,
            43.4,
            boxstyle="round,pad=0,rounding_size=1.4",
            linewidth=0,
            facecolor=NAVY2,
            zorder=0,
        )
    )

    ax.text(6.2, 40.4, "ШКАЛА ПОТЕРЬ", fontsize=10, color=TEAL, fontweight="bold", va="center", zorder=7)
    ax.text(
        6.2,
        36.6,
        "Компания ~1000 человек  ·  что стоит «как сейчас»",
        fontsize=11,
        color=MUTED,
        va="center",
        zorder=7,
    )

    # Track
    y = 22.8
    x0, x1 = 10, 122
    ax.add_patch(FancyBboxPatch((x0, y - 1.15), x1 - x0, 2.3, boxstyle="round,pad=0,rounding_size=1.1", linewidth=0, facecolor="#0A1A30", zorder=1))

    cmap = LinearSegmentedColormap.from_list("lights", [TEAL, GOLD, RED])
    xs = np.linspace(x0 + 1.2, x1 - 1.2, 56)
    for i, x in enumerate(xs):
        c = cmap(i / (len(xs) - 1))
        h = 0.55 if i % 4 else 1.35
        w = 0.28 if i % 4 else 0.42
        ax.add_patch(Rectangle((x - w / 2, y - h / 2), w, h, facecolor=c, edgecolor="none", alpha=0.55 if i % 4 else 0.95, zorder=2))

    # User order: 16 млн → 3 млн (before NS) → 500 тыс. НС
    marks = [
        {
            "x": 28,
            "color": TEAL,
            "value": "16 млн ₽",
            "title": "Потеря рабочего времени",
            "caption": "Оплата времени работников\nна обучение",
        },
        {
            "x": 66,
            "color": GOLD,
            "value": "3 млн ₽",
            "title": "Организация обучения",
            "caption": "Договоры и администрирование\n«для галочки»",
        },
        {
            "x": 104,
            "color": RED,
            "value": "от 500 тыс. ₽",
            "title": "Потеря смены / НС",
            "caption": "Больничный, простой,\nрасследование",
        },
    ]

    for m in marks:
        glow_circle(ax, m["x"], y, 1.55, m["color"])
        ax.plot([m["x"], m["x"]], [y + 2.1, y + 5.6], color=m["color"], lw=1.4, alpha=0.85, zorder=4)
        ax.add_patch(
            FancyBboxPatch(
                (m["x"] - 14.6, 4.4),
                29.2,
                12.6,
                boxstyle="round,pad=0,rounding_size=0.7",
                linewidth=0,
                facecolor="#0A1A30",
                zorder=4,
            )
        )
        ax.text(m["x"], 14.6, m["value"], fontsize=16, color=m["color"], fontweight="bold", ha="center", va="center", zorder=7)
        ax.text(m["x"], 10.8, m["title"], fontsize=10.5, color=WHITE, fontweight="bold", ha="center", va="center", zorder=7)
        ax.text(m["x"], 7.0, m["caption"], fontsize=8.6, color=MUTED, ha="center", va="center", zorder=7, linespacing=1.25)

    fig.subplots_adjust(left=0, right=1, top=1, bottom=0)
    fig.savefig(OUT, dpi=220, facecolor=NAVY)
    plt.close()
    print("wrote", OUT)


if __name__ == "__main__":
    main()
