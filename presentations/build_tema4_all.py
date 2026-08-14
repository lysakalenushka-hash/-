#!/usr/bin/env python3
"""Собрать все презентации темы 4."""

from __future__ import annotations

from deck_builder import build_all
from tema4_content import DECKS, TEMA4_COURSE, TEMA4_SUBTITLE


if __name__ == "__main__":
    paths = build_all(DECKS, subtitle=TEMA4_SUBTITLE, course=TEMA4_COURSE)
    print(f"Built {len(paths)} presentations")
