#!/usr/bin/env python3
"""Универсальный сборщик презентаций в legal-стиле из структурированного контента."""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu

from legal_style import (
    ACCENT_RED, ASSETS, LINE, MUTED, NUM_BG, OUT, ROOT, TEXT, WHITE,
    blank, bullets, content_header, cream_note, new_prs, oval, pic_fit, rect,
    set_context, slide_thanks, slide_title, slide_toc, tbox, verify,
)

INFO = ROOT / "infographics"
INFO.mkdir(parents=True, exist_ok=True)


def _slide_intro(prs, s, num):
    slide = blank(prs)
    content_header(slide, s["title"], num)
    tbox(slide, Emu(700000), Emu(1600000), Emu(11000000), Emu(4200000),
         s["text"], size=16, color=TEXT)
    if s.get("note"):
        cream_note(slide, Emu(700000), Emu(5600000), Emu(11000000), Emu(1000000),
                   s["note"], size=13)
    return slide


def _slide_bullets(prs, s, num):
    slide = blank(prs)
    content_header(slide, s["title"], num)
    items = s["items"]
    alert = set(s.get("alert") or [])
    bullets(slide, Emu(700000), Emu(1500000), Emu(11000000), Emu(4200000),
            items, size=14, alert=alert)
    if s.get("note"):
        cream_note(slide, Emu(700000), Emu(5600000), Emu(11000000), Emu(1000000),
                   s["note"], size=13)
    return slide


def _slide_alert_bullets(prs, s, num):
    s = dict(s)
    if "alert" not in s:
        # подсветить только первый пункт как акцент
        s["alert"] = [0] if s.get("items") else []
    return _slide_bullets(prs, s, num)


def _slide_cards(prs, s, num):
    slide = blank(prs)
    content_header(slide, s["title"], num)
    cards = s["cards"]
    n = len(cards)
    if n <= 3:
        cols, rows = n, 1
    elif n == 4:
        cols, rows = 2, 2
    elif n == 5:
        cols, rows = 3, 2
    else:
        cols, rows = 3, 2
    gap_x, gap_y = Emu(250000), Emu(200000)
    top = Emu(1450000)
    left = Emu(700000)
    avail_w = Emu(11000000)
    avail_h = Emu(4800000)
    cw = int((avail_w - gap_x * (cols - 1)) / cols)
    rh = int((avail_h - gap_y * (rows - 1)) / rows)
    for i, c in enumerate(cards):
        r, col = divmod(i, cols)
        if r >= rows:
            break
        x = left + col * (cw + gap_x)
        y = top + r * (rh + gap_y)
        rect(slide, x, y, cw, rh, WHITE, line=LINE)
        oval(slide, x + Emu(150000), y + Emu(200000), Emu(500000), Emu(500000), NUM_BG)
        tbox(slide, x + Emu(150000), y + Emu(200000), Emu(500000), Emu(500000),
             str(c.get("n", i + 1)), size=14, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        tbox(slide, x + Emu(750000), y + Emu(250000), cw - Emu(900000), Emu(450000),
             c["title"], size=13, bold=True, color=TEXT)
        tbox(slide, x + Emu(150000), y + Emu(850000), cw - Emu(300000), rh - Emu(1000000),
             c["desc"], size=12, color=MUTED)
    return slide


def _slide_steps(prs, s, num):
    slide = blank(prs)
    content_header(slide, s["title"], num)
    steps = s["steps"]
    y = Emu(1450000)
    h = min(Emu(750000), int(Emu(5000000) / max(len(steps), 1)))
    for i, step in enumerate(steps, 1):
        oval(slide, Emu(700000), y, Emu(550000), Emu(550000),
             ACCENT_RED if i == 1 else NUM_BG)
        tbox(slide, Emu(700000), y, Emu(550000), Emu(550000), str(i),
             size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE)
        rect(slide, Emu(1450000), y, Emu(10500000), Emu(550000), WHITE, line=LINE)
        tbox(slide, Emu(1650000), y, Emu(10100000), Emu(550000), step,
             size=13, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)
        y += h + Emu(120000)
    if s.get("note"):
        cream_note(slide, Emu(700000), Emu(5800000), Emu(11000000), Emu(900000),
                   s["note"], size=12)
    return slide


def _slide_two_col(prs, s, num):
    slide = blank(prs)
    content_header(slide, s["title"], num)
    tbox(slide, Emu(700000), Emu(1450000), Emu(5400000), Emu(400000),
         s.get("left_title", ""), size=14, bold=True, color=TEXT)
    tbox(slide, Emu(6600000), Emu(1450000), Emu(5400000), Emu(400000),
         s.get("right_title", ""), size=14, bold=True, color=TEXT)
    bullets(slide, Emu(700000), Emu(1950000), Emu(5400000), Emu(4000000),
            s.get("left", []), size=13)
    bullets(slide, Emu(6600000), Emu(1950000), Emu(5400000), Emu(4000000),
            s.get("right", []), size=13)
    return slide


def _slide_summary(prs, points, num):
    slide = blank(prs)
    content_header(slide, "ГЛАВНОЕ ЗАПОМНИТЬ", num)
    n = len(points)
    h = min(Emu(850000), int(Emu(5000000) / max(n, 1)))
    y = Emu(1450000)
    for pt in points:
        rect(slide, Emu(700000), y, Emu(11000000), h, WHITE, line=LINE)
        rect(slide, Emu(700000), y, Emu(90000), h, ACCENT_RED)
        tbox(slide, Emu(1000000), y + Emu(100000), Emu(10400000), h - Emu(200000),
             pt, size=13, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)
        y += h + Emu(100000)
    return slide


def _slide_memo(prs, memo_name, num):
    slide = blank(prs)
    content_header(slide, "ПАМЯТКА", num)
    candidates = [
        INFO / memo_name,
        Path("/opt/cursor/artifacts/assets") / memo_name,
        ASSETS / memo_name,
    ]
    path = next((p for p in candidates if p.exists()), None)
    if path is None:
        cream_note(slide, Emu(700000), Emu(2500000), Emu(11000000), Emu(1500000),
                   "Памятка будет добавлена рядом с презентацией.", size=14)
        return slide
    from PIL import Image as PILImage
    with PILImage.open(path) as im:
        iw, ih = im.size
    max_w, max_h = Emu(11000000), Emu(5200000)
    l, t = Emu(700000), Emu(1400000)
    scale = min(float(max_w) / iw, float(max_h) / ih)
    w, h = int(iw * scale), int(ih * scale)
    x = int(l + (max_w - w) / 2)
    y = int(t + (max_h - h) / 2)
    slide.shapes.add_picture(str(path), x, y, width=w, height=h)
    return slide


HANDLERS = {
    "intro": _slide_intro,
    "bullets": _slide_bullets,
    "alert_bullets": _slide_alert_bullets,
    "cards": _slide_cards,
    "steps": _slide_steps,
    "two_col": _slide_two_col,
}


def build_deck(deck: dict) -> Path:
    prs = new_prs()
    slide_title(prs, deck["title"])
    toc = deck.get("toc") or []
    # keep toc to max 5 for layout
    slide_toc(prs, toc[:5], 2)
    num = 3
    for s in deck["slides"]:
        handler = HANDLERS.get(s["type"])
        if handler is None:
            raise ValueError(f"Unknown slide type: {s['type']}")
        handler(prs, s, num)
        num += 1
    _slide_summary(prs, deck["memo_points"], num)
    num += 1
    if deck.get("memo_file"):
        _slide_memo(prs, deck["memo_file"], num)
        num += 1
    slide_thanks(prs)

    name = deck["filename"]
    path = OUT / name
    prs.save(path)
    verify(path)
    path2 = ROOT / name
    prs.save(path2)
    print("Saved:", path2)
    return path2


def build_all(decks, *, subtitle: str, course: str):
    set_context(subtitle, course)
    paths = []
    for deck in decks:
        paths.append(build_deck(deck))
    return paths
