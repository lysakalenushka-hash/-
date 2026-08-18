#!/usr/bin/env python3
"""Crop illustration-only regions from source PDF page renders (no captions)."""
from pathlib import Path
from PIL import Image

PAGES = Path("/tmp/bleed_pages")
OUT = Path("/workspace/tema2_bleeding/assets")
OUT.mkdir(parents=True, exist_ok=True)

CROPS = {
    # page: [(box, name), ...]
    2: [((240, 540, 960, 2040), "def_man.png")],
    3: [
        ((300, 820, 1260, 2040), "bleed_arterial.png"),
        ((1420, 820, 2380, 2040), "bleed_venous.png"),
        ((2580, 820, 3620, 2040), "bleed_capillary.png"),
    ],
    4: [((200, 540, 2000, 2080), "overview_exam.png")],
    5: [
        ((220, 780, 1820, 2040), "direct_pressure.png"),
        ((1980, 780, 3680, 2040), "pressure_bandage.png"),
    ],
    6: [
        ((220, 680, 1260, 1550), "carotid_point.png"),
        ((1360, 680, 2420, 1550), "carotid_4fingers.png"),
        ((2520, 680, 3660, 1550), "carotid_thumb.png"),
    ],
    7: [
        ((220, 680, 1260, 1550), "subclavian_1.png"),
        ((1360, 680, 2420, 1550), "subclavian_2.png"),
        ((2520, 680, 3660, 1550), "subclavian_3.png"),
    ],
    8: [
        ((280, 650, 1820, 1650), "brachial_1.png"),
        ((2020, 650, 3580, 1650), "brachial_2.png"),
    ],
    9: [
        ((300, 1200, 1800, 2050), "axillary_1.png"),
        ((2050, 1200, 3600, 2050), "axillary_2.png"),
    ],
    10: [
        ((280, 1000, 1850, 2050), "femoral_1.png"),
        ((2000, 1000, 3620, 2050), "femoral_2.png"),
    ],
    11: [
        ((220, 650, 1260, 1600), "flexion_arm.png"),
        ((1360, 650, 2420, 1600), "flexion_leg.png"),
        ((2520, 650, 3660, 1600), "flexion_thigh.png"),
    ],
    12: [
        ((220, 820, 1260, 2050), "tourniquet_1.png"),
        ((1360, 820, 2420, 2050), "tourniquet_2.png"),
        ((2520, 820, 3660, 2050), "tourniquet_3.png"),
    ],
    13: [((1950, 560, 3680, 2050), "improvised_tq.png")],
    14: [
        ((250, 720, 1800, 1280), "nose_1.png"),
        ((2050, 720, 3600, 1280), "nose_2.png"),
        ((250, 1450, 1800, 2050), "nose_3.png"),
        ((2050, 1450, 3600, 2050), "nose_4.png"),
    ],
    16: [((1980, 560, 3680, 2050), "shock_prev.png")],
    17: [
        ((220, 860, 1260, 2050), "exam_head.png"),
        ((1360, 860, 2420, 2050), "exam_neck.png"),
        ((2520, 860, 3660, 2050), "exam_chest.png"),
    ],
    18: [
        ((220, 860, 1260, 2050), "exam_abdomen.png"),
        ((1360, 860, 2420, 2050), "exam_legs.png"),
        ((2520, 860, 3660, 2050), "exam_arms.png"),
    ],
    19: [
        ((2350, 750, 3600, 1080), "head_trauma_1.png"),
        ((220, 1080, 1700, 1480), "head_trauma_2.png"),
        ((2350, 1650, 3600, 2080), "head_trauma_3.png"),
    ],
    20: [
        ((280, 780, 1750, 1650), "eye_injury.png"),
        ((2100, 780, 3600, 1650), "nose_injury.png"),
    ],
    21: [
        ((220, 880, 1260, 2050), "neck_1.png"),
        ((1360, 880, 2420, 2050), "neck_2.png"),
        ((2520, 880, 3660, 2050), "neck_3.png"),
    ],
    22: [
        ((220, 900, 1260, 2050), "chest_1.png"),
        ((1360, 900, 2420, 2050), "chest_2.png"),
        ((2520, 900, 3660, 2050), "chest_3.png"),
    ],
    23: [((350, 1280, 3500, 2050), "abdomen_closed.png")],
    24: [((350, 1280, 3500, 2050), "abdomen_open.png")],
    25: [
        ((200, 780, 1260, 1650), "limb_1.png"),
        ((1340, 780, 2440, 1650), "limb_2.png"),
        ((2500, 780, 3680, 1650), "limb_3.png"),
    ],
    26: [((1950, 560, 3680, 2050), "immobilization.png")],
    27: [
        ((220, 520, 1850, 1100), "spine_1.png"),
        ((2000, 1300, 3650, 2050), "spine_2.png"),
    ],
}


def main():
    for page, items in CROPS.items():
        src = PAGES / f"page_{page:02d}.png"
        im = Image.open(src).convert("RGB")
        for box, name in items:
            im.crop(box).save(OUT / name, optimize=True)
            print(f"p{page:02d} -> {name}")
    print("Done.")


if __name__ == "__main__":
    main()
