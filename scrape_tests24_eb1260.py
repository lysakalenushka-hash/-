#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Пройти билеты ЭБ 1260.25 на tests24.su и сохранить ключи из разбора после сдачи."""

from __future__ import annotations

import html
import json
import re
import time
import urllib.parse
import urllib.request
from pathlib import Path

CATALOG = "https://tests24.su/eb-5-gruppa-vyshe-1000-v/"
AJAX = "https://tests24.su/wp-admin/admin-ajax.php"
OUT = Path("/tmp/tests24_eb1260.json")
UA = {
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36"
}


def fetch(url: str, data: dict | None = None, referer: str | None = None) -> str:
    headers = dict(UA)
    if referer:
        headers["Referer"] = referer
        headers["Origin"] = "https://tests24.su"
    body = None
    if data is not None:
        body = urllib.parse.urlencode(data, doseq=True).encode()
        headers["Content-Type"] = "application/x-www-form-urlencoded"
    req = urllib.request.Request(url, data=body, headers=headers)
    with urllib.request.urlopen(req, timeout=60) as r:
        return r.read().decode("utf-8", "replace")


def unescape(s: str) -> str:
    s = html.unescape(s or "")
    s = re.sub(r"<[^>]+>", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def parse_ticket(page: str) -> dict:
    quiz = re.search(r'name="quiz_id"\s+value="(\d+)"', page)
    post = re.search(r"WatuPRO\.post_id\s*=\s*(\d+)", page)
    if not post:
        m = re.search(r"WatuPRO\.post_id\s*=\s*(\d+)", html.unescape(page))
        post = m
    # init script is base64; also try decoded blobs
    watu_q = re.search(r'name="watupro_questions"\s+value="([^"]+)"', page)
    start = re.search(r'name="start_time"\s+id="startTime"\s+value="([^"]+)"', page)
    questions = []
    wq = html.unescape(watu_q.group(1)) if watu_q else ""
    for part in [p.strip() for p in wq.split("|") if p.strip()]:
        qid, _, aids = part.partition(":")
        qid = qid.strip()
        ids = [x.strip() for x in aids.split(",") if x.strip()]
        wrap = re.search(
            rf"watupro-question-id-{qid}'>([\s\S]*?)</div><!-- end questionWrap-->",
            page,
        )
        qtext = ""
        answers = []
        atype = "radio"
        tm = re.search(rf"id='answerType{qid}'[^>]*value='(\w+)'", page)
        if tm:
            atype = tm.group(1)
        if wrap:
            sm = re.search(r"<strong>(.*?)</strong>", wrap.group(1), re.S)
            if sm:
                qtext = unescape(sm.group(1))
            for aid in ids:
                am = re.search(
                    rf"id='answer-id-{aid}'[\s\S]*?<span>(.*?)</span>",
                    wrap.group(1),
                )
                answers.append({"id": aid, "text": unescape(am.group(1) if am else "")})
        else:
            answers = [{"id": aid, "text": ""} for aid in ids]
        questions.append({"id": qid, "q": qtext, "atype": atype, "answers": answers})
    post_id = None
    if post:
        post_id = post.group(1)
    else:
        # decode base64 init scripts
        for b64 in re.findall(r'src="data:text/javascript;base64,([^"]+)"', page):
            try:
                import base64

                js = base64.b64decode(b64).decode("utf-8", "replace")
            except Exception:
                continue
            m = re.search(r"WatuPRO\.post_id\s*=\s*(\d+)", js)
            if m:
                post_id = m.group(1)
                break
    return {
        "quiz_id": quiz.group(1) if quiz else None,
        "post_id": post_id,
        "start_time": start.group(1) if start else "",
        "watupro_questions": html.unescape(watu_q.group(1)) if watu_q else "",
        "questions": questions,
    }


def parse_submit(html_text: str) -> list[dict]:
    out = []
    chunks = re.split(r"show-question-content", html_text)[1:]
    for ch in chunks:
        qm = re.search(r"<strong>(.*?)</strong>", ch, re.S)
        q = unescape(qm.group(1)) if qm else ""
        correct = []
        for cm in re.finditer(
            r"class='answer correct-answer'>[\s\S]*?class='answer'>(.*?)</span>",
            ch,
        ):
            correct.append(unescape(cm.group(1)))
        if not correct:
            for cm in re.finditer(
                r"correct-answerWATUEMAIL-->(.*?)</span>",
                ch,
            ):
                correct.append(unescape(cm.group(1)))
        opts = []
        for om in re.finditer(r"<!--WATUEMAILanswer[^>]*-->(.*?)</span>", ch):
            opts.append(unescape(om.group(1)))
        if not opts:
            for om in re.finditer(r"<li class='answer[^']*'[\s\S]*?class='answer'>(.*?)</span>", ch):
                opts.append(unescape(om.group(1)))
        if q:
            out.append({"q": q, "correct": correct, "opts": opts})
    return out


def ticket_urls() -> list[str]:
    cat = fetch(CATALOG)
    nums = sorted({int(n) for n in re.findall(r"eb-1260-bilet-(\d+)", cat)})
    return [f"https://tests24.su/eb-1260-bilet-{n}/" for n in nums]


def main():
    urls = ticket_urls()
    print("tickets", len(urls))
    allq = []
    seen = set()
    for i, url in enumerate(urls, 1):
        page = fetch(url)
        t = parse_ticket(page)
        if not t["quiz_id"] or not t["questions"]:
            print(i, url, "PARSE FAIL quiz", t["quiz_id"], "nq", len(t["questions"]))
            continue
        payload = {
            "action": "watupro_submit",
            "quiz_id": t["quiz_id"],
            "post_id": t["post_id"] or "",
            "start_time": t["start_time"],
            "watupro_questions": t["watupro_questions"],
        }
        qids = []
        for q in t["questions"]:
            qids.append(q["id"])
            payload.setdefault("question_id[]", []).append(q["id"])
            if q["answers"]:
                payload[f"answer-{q['id']}[]"] = q["answers"][0]["id"]
        result = fetch(AJAX, payload, referer=url)
        keyed = parse_submit(result)
        new = 0
        for it in keyed:
            if it["q"] not in seen:
                seen.add(it["q"])
                allq.append({**it, "ticket": url})
                new += 1
        print(
            f"{i}/{len(urls)} {url} quiz={t['quiz_id']} q={len(t['questions'])} "
            f"keyed={len(keyed)} new={new} total={len(allq)} corr_marks={result.count('correct-answer')}"
        )
        time.sleep(0.35)
    OUT.write_text(json.dumps(allq, ensure_ascii=False, indent=2), encoding="utf-8")
    print("saved", len(allq), OUT)


if __name__ == "__main__":
    main()
