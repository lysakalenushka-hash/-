#!/usr/bin/env python3
"""Slide 5: illuminated cost scale, segment length and circle area ∝ rubles."""

from pathlib import Path

import numpy as np
from matplotlib import font_manager
from matplotlib.patches import Circle, FancyBboxPatch, Rectangle
import matplotlib.pyplot as plt

OUT = Path(__file__).with_name("slide5-shkala-poter.png")

NAVY = "#0B1F3A"
NAVY2 = "#122844"
MUTED = "#9AA8B8"
TEAL = "#2EE0C5"
GOLD = "#F0C14B"
RED = "#FF6B5A"
WHITE = "#FFFFFF"
TRACK = "#0A1A30"

# Shares from the slide amounts, mln ₽
ITEMS = [
    {
        "mln": 16.0,
        "color": TEAL,
        "value": "16 млн ₽",
        "title": "Потеря рабочего времени",
        "caption": "Оплата времени работников\nна обучение",
        "value_short": "16 млн",
    },
    {
        "mln": 3.0,
        "color": GOLD,
        "value": "3 млн ₽",
        "title": "Организация обучения",
        "caption": "Договоры и администрирование\n«для галочки»",
        "value_short": "3 млн",
    },
    {
        "mln": 0.5,
        "color": RED,
        "value": "от 500 тыс. ₽",
        "title": "Потеря смены / НС",
        "caption": "Больничный, простой,\nрасследование",
        "value_short": "0,5 млн",
    },
]
TOTAL = sum(i["mln"] for i in ITEMS)  # 19.5


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


def glow_circle(ax, x, y, r, color, layers=7):
    for i in range(layers, 0, -1):
        ax.add_patch(
            Circle(
                (x, y),
                r * (1 + i * 0.28),
                facecolor=color,
                edgecolor="none",
                alpha=0.04 * i,
                zorder=4,
            )
        )
    ax.add_patch(Circle((x, y), r, facecolor=color, edgecolor="none", zorder=6))
    ax.add_patch(
        Circle((x, y), max(r * 0.34, 0.12), facecolor=WHITE, edgecolor="none", alpha=0.92, zorder=7)
    )


def main():
    setup_fonts()
    fig, ax = plt.subplots(figsize=(13.2, 4.55), dpi=220)
    fig.patch.set_facecolor(NAVY)
    ax.set_facecolor(NAVY)
    ax.set_xlim(0, 132)
    ax.set_ylim(0, 45.5)
    ax.axis("off")
    ax.set_aspect("equal", adjustable="box")
    # equal aspect makes the figure letterbox; drop it so the banner fills 16:9-ish
    ax.set_aspect("auto")

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

    ax.text(6.2, 41.0, "ШКАЛА ПОТЕРЬ", fontsize=10, color=TEAL, fontweight="bold", va="center", zorder=8)
    ax.text(
        6.2,
        37.4,
        "Длина участка и площадь круга = доля от 19,5 млн ₽",
        fontsize=10.5,
        color=MUTED,
        va="center",
        zorder=8,
    )

    bar_y = 24.6
    bar_h = 3.15
    x0, x1 = 7.2, 124.8
    bar_w = x1 - x0

    ax.add_patch(
        FancyBboxPatch(
            (x0 - 0.35, bar_y - bar_h / 2 - 0.35),
            bar_w + 0.7,
            bar_h + 0.7,
            boxstyle="round,pad=0,rounding_size=1.35",
            linewidth=0,
            facecolor=TRACK,
            zorder=1,
        )
    )

    # Circle area ∝ amount ⇒ radius ∝ sqrt(mln)
    r_max = 5.15
    cursor = x0
    for item in ITEMS:
        share = item["mln"] / TOTAL
        w = bar_w * share
        item["x0"] = cursor
        item["x1"] = cursor + w
        item["mid"] = cursor + w / 2
        item["share"] = share
        item["r"] = r_max * np.sqrt(item["mln"] / ITEMS[0]["mln"])
        ax.add_patch(
            Rectangle(
                (cursor, bar_y - bar_h / 2),
                w,
                bar_h,
                facecolor=item["color"],
                edgecolor="none",
                alpha=0.22,
                zorder=2,
            )
        )
        n_ticks = max(2, int(round(48 * share)))
        xs = np.linspace(cursor + 0.35, cursor + w - 0.35, n_ticks) if w > 0.8 else np.array([item["mid"]])
        for i, x in enumerate(xs):
            major = i % 4 == 0 or n_ticks <= 3
            h = bar_h * (0.92 if major else 0.42)
            tw = 0.38 if major else 0.22
            ax.add_patch(
                Rectangle(
                    (x - tw / 2, bar_y - h / 2),
                    tw,
                    h,
                    facecolor=item["color"],
                    edgecolor="none",
                    alpha=0.95 if major else 0.55,
                    zorder=3,
                )
            )
        cursor += w

    for item in ITEMS:
        glow_circle(ax, item["mid"], bar_y, item["r"], item["color"])

    a, b, c = ITEMS

    ax.text(a["mid"], 15.4, f"{a['value']}  ·  {a['share']*100:.0f}%", fontsize=17, color=TEAL, fontweight="bold", ha="center", zorder=8)
    ax.text(a["mid"], 12.0, a["title"], fontsize=11, color=WHITE, fontweight="bold", ha="center", zorder=8)
    ax.text(a["mid"], 8.2, a["caption"], fontsize=8.8, color=MUTED, ha="center", zorder=8, linespacing=1.25)

    ax.text(104.0, 15.4, f"{b['value']}  ·  {b['share']*100:.0f}%", fontsize=13, color=GOLD, fontweight="bold", ha="center", zorder=8)
    ax.text(104.0, 12.0, b["title"], fontsize=9.6, color=WHITE, fontweight="bold", ha="center", zorder=8)
    ax.text(104.0, 8.2, b["caption"], fontsize=8.0, color=MUTED, ha="center", zorder=8, linespacing=1.25)

    ax.annotate(
        "",
        xy=(c["mid"], bar_y + c["r"] + 0.15),
        xytext=(126.8, 40.6),
        arrowprops=dict(arrowstyle="-|>", color=RED, lw=1.15, mutation_scale=9),
        zorder=5,
    )
    ax.text(129.0, 41.2, c["value"], fontsize=12, color=RED, fontweight="bold", ha="right", va="center", zorder=8)
    ax.text(129.0, 38.4, f"{c['title']}  ·  {c['share']*100:.1f}%", fontsize=8.3, color=WHITE, ha="right", va="center", zorder=8)
    ax.text(129.0, 36.0, "Больничный, простой, расследование", fontsize=7.5, color=MUTED, ha="right", va="center", zorder=8)

    fig.subplots_adjust(left=0, right=1, top=1, bottom=0)
    fig.savefig(OUT, dpi=220, facecolor=NAVY)
    plt.close()
    print("wrote", OUT)
    for i in ITEMS:
        print(
            f"{i['value_short']}: share={i['share']*100:.2f}%  "
            f"len={i['x1']-i['x0']:.2f}  r={i['r']:.2f}  r^2={i['r']**2:.2f}"
        )


if __name__ == "__main__":
    main()
