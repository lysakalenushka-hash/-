#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Скачать последовательные билеты ЭБ 1260.25 с prombez24 (ключи в hidden-поле correct)."""

from __future__ import annotations

import html
import json
import re
import time
import urllib.request
from pathlib import Path

OUT = Path("/tmp/prombez24_eb1260.json")
UA = {"User-Agent": "Mozilla/5.0 (compatible; exam-prep-research/1.0)"}


def fetch(url: str) -> str:
    req = urllib.request.Request(url, headers=UA)
    with urllib.request.urlopen(req, timeout=30) as r:
        return r.read().decode("utf-8", "replace")


def unescape(s: str) -> str:
    return html.unescape(s or "").strip()


def parse_page(text: str) -> list[dict]:
    qs = []
    for b in re.split(r'<div class="question row">', text)[1:]:
        qm = re.search(r'name="questionList\[\d+\]\.questionText"\s+value="([^"]*)"', b)
        if not qm:
            continue
        ntd = re.search(r'name="questionList\[\d+\]\.ntdLink"\s+value="([^"]*)"', b)
        opts = []
        for am in re.finditer(
            r'name="questionList\[\d+\]\.answers\[\d+\]\.answerText"\s+value="([^"]*)"[\s\S]*?'
            r'name="questionList\[\d+\]\.answers\[\d+\]\.correct"\s+value="(true|false)"',
            b,
        ):
            opts.append({"text": unescape(am.group(1)), "correct": am.group(2) == "true"})
        q = unescape(qm.group(1))
        if q and opts:
            qs.append(
                {
                    "q": q,
                    "ntd": unescape(ntd.group(1) if ntd else ""),
                    "opts": opts,
                    "correct": [o["text"] for o in opts if o["correct"]],
                }
            )
    return qs


def main():
    allq = []
    seen = set()
    for page in range(0, 80):
        url = f"https://prombez24.com/ticket/ordered/?testId=212&page={page}&size=10"
        items = parse_page(fetch(url))
        new = 0
        for it in items:
            if it["q"] not in seen:
                seen.add(it["q"])
                allq.append(it)
                new += 1
        print(f"page {page}: {len(items)} parsed, {new} new, total {len(allq)}")
        if not items:
            break
        time.sleep(0.25)
    OUT.write_text(json.dumps(allq, ensure_ascii=False, indent=2), encoding="utf-8")
    print("saved", len(allq), OUT)


if __name__ == "__main__":
    main()
