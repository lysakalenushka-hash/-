#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Download indtec.ru electro tests (sequential pages) into JSON."""

from __future__ import annotations

import json
import re
import time
import urllib.request
from html.parser import HTMLParser
from pathlib import Path

OUT = Path("/tmp/indtec_elektro.json")
CACHE = Path("/tmp/indtec_pages")
CACHE.mkdir(exist_ok=True)
UA = {"User-Agent": "Mozilla/5.0 (compatible; question-bank-sync/1.0)"}

TESTS = {
    "elektro-5v-neprom": {"lms": 199, "kind": "непромышленные", "above": True, "group": 5, "pages": 9},
    "elektro-2d-neprom": {"lms": 200, "kind": "непромышленные", "above": False, "group": 2, "pages": 5},
    "elektro-3d-neprom": {"lms": 201, "kind": "непромышленные", "above": False, "group": 3, "pages": 8},
    "elektro-4d-neprom": {"lms": 202, "kind": "непромышленные", "above": False, "group": 4, "pages": 8},
    "elektro-5d-neprom": {"lms": 203, "kind": "непромышленные", "above": False, "group": 5, "pages": 9},
    "elektro-2v": {"lms": 204, "kind": "промышленные", "above": True, "group": 2, "pages": 11},
    "elektro-3v": {"lms": 205, "kind": "промышленные", "above": True, "group": 3, "pages": 36},
    "elektro-4v": {"lms": 206, "kind": "промышленные", "above": True, "group": 4, "pages": 42},
    "elektro-5": {"lms": 207, "kind": "промышленные", "above": True, "group": 5, "pages": 42},
    "elektro-2d": {"lms": 208, "kind": "промышленные", "above": False, "group": 2, "pages": 10},
    "elektro-3d": {"lms": 209, "kind": "промышленные", "above": False, "group": 3, "pages": 29},
    "elektro-4d": {"lms": 210, "kind": "промышленные", "above": False, "group": 4, "pages": 31},
    "elektro-5d": {"lms": 211, "kind": "промышленные", "above": False, "group": 5, "pages": 31},
}

Q_RE = re.compile(
    r'<div class="question card mb-3" data-qid="(\d+)">'
    r'<div class="card-header"><strong>Вопрос \d+:</strong>\s*(.*?)</div>'
    r'<div class="card-body">(.*?)</div></div>',
    re.S,
)
A_RE = re.compile(
    r'<label class="form-check-label" for="[^"]+">(.*?)</label>',
    re.S,
)
C_RE = re.compile(r'name="correct\[(\d+)\]" value="([^"]+)"')


def fetch(url: str, cache: Path) -> str:
    if cache.exists() and cache.stat().st_size > 1000:
        return cache.read_text(encoding="utf-8", errors="replace")
    req = urllib.request.Request(url, headers=UA)
    html = urllib.request.urlopen(req, timeout=45).read().decode("utf-8", "replace")
    cache.write_text(html, encoding="utf-8")
    time.sleep(0.12)
    return html


def unescape(t: str) -> str:
    t = re.sub(r"<[^>]+>", "", t)
    t = (
        t.replace("&nbsp;", " ")
        .replace("&quot;", '"')
        .replace("&#039;", "'")
        .replace("&amp;", "&")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
    )
    return re.sub(r"\s+", " ", t).strip()


def parse_page(html: str) -> list[dict]:
    corrects = {int(i): v for i, v in C_RE.findall(html)}
    items = []
    for qid, qtext, body in Q_RE.findall(html):
        qid = int(qid)
        answers = [unescape(a) for a in A_RE.findall(body)]
        mask = corrects.get(qid, "")
        flags = []
        for i, _ in enumerate(answers):
            flags.append(1 if i < len(mask) and mask[i] == "1" else 0)
        items.append({"qid": qid, "q": unescape(qtext), "answers": answers, "flags": flags})
    return items


def main():
    data = {}
    for slug, meta in TESTS.items():
        bank = []
        seen = set()
        for p in range(1, meta["pages"] + 1):
            url = f"https://indtec.ru/tests/{slug}/{p}/"
            cache = CACHE / f"{slug}_{p}.html"
            try:
                html = fetch(url, cache)
            except Exception as e:
                print("FAIL", slug, p, e)
                continue
            for item in parse_page(html):
                key = item["q"]
                if key in seen:
                    continue
                seen.add(key)
                bank.append(item)
        data[slug] = {"meta": meta, "questions": bank}
        print(slug, "got", len(bank))
    OUT.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    print("wrote", OUT)


if __name__ == "__main__":
    main()
