#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Сопоставить ключи бесплатных тренажёров ЭБ 1260.25 (prombez24 и tests24.su) с листом V приложения 2 РТН."""

from __future__ import annotations

import html
import json
import re
from difflib import SequenceMatcher
from pathlib import Path

from docx import Document
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

RTN_PATH = Path("/tmp/Prilozhenie_2_RTN.xlsx")
TRAINER_FILES = [
    ("prombez24", Path("/tmp/prombez24_eb1260.json")),
    ("tests24.su", Path("/tmp/tests24_eb1260.json")),
    ("prombez24 спецразделы", Path("/tmp/prombez24_spec.json")),
]
DOCX_FILES = [
    (
        "24тест V выше 1000",
        Path("/home/ubuntu/.cursor/projects/workspace/uploads/PT_PR_1_5_01_09_2026v2_NTD_8e0d.docx"),
    ),
    (
        "24тест IV до 1000",
        Path("/home/ubuntu/.cursor/projects/workspace/uploads/PT_PR_0_4_01_09_2026v2_NTD_f473.docx"),
    ),
]
OUT = Path("электробезопасность_V_ключи_тренажёр.xlsx")

FONT = Font(name="Times New Roman", size=11)
HEAD = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
FILL_H = PatternFill("solid", fgColor="1F4E79")
FILL_OK = PatternFill("solid", fgColor="C6EFCE")
FILL_NO = PatternFill("solid", fgColor="FCE4D6")
ALIGN = Alignment(wrap_text=True, vertical="top")


def norm(s: str) -> str:
    s = html.unescape(s or "")
    s = s.replace("\xa0", " ").replace("ё", "е").replace("Ё", "Е")
    s = re.sub(r"<[^>]+>", " ", s)
    s = re.sub(r"[«»“”]", '"', s)
    return re.sub(r"\s+", " ", s).strip().lower()


def stem(s: str) -> str:
    s = norm(s)
    s = re.split(r",\s*согласно|\s+согласно\s+|,\s*утвержд|утвержд[её]нн", s, maxsplit=1)[0]
    return s.strip(" ?.,;")


def parse_rtn_v():
    wb = load_workbook(RTN_PATH, data_only=True)
    ws = wb["V"]
    npa = ""
    qs = []
    cur = None
    for row in ws.iter_rows(min_row=1, values_only=True):
        a, b, c, d, e, f, g, h = (list(row) + [None] * 8)[:8]
        if a == "№":
            continue
        is_num = isinstance(a, int) or (isinstance(a, str) and str(a).strip().isdigit())
        if a and not is_num and b is None:
            npa = str(a).strip()
            continue
        if is_num and b:
            if cur:
                qs.append(cur)
            cur = {
                "n": int(a),
                "q": str(b).strip(),
                "npa": npa,
                "v1000": c == "+",
                "v_above": d == "+",
                "opts": [],
            }
        elif cur and b and (a is None or a == ""):
            cur["opts"].append(str(b).strip())
    if cur:
        qs.append(cur)
    wb.close()
    return qs


def best_match(qtext: str, bank: list[dict], threshold: float = 0.84):
    nq = norm(qtext)
    st = stem(qtext)
    exact = {norm(x["q"]): x for x in bank}
    if nq in exact:
        return exact[nq], 1.0, "exact"
    by_stem = {stem(x["q"]): x for x in bank}
    if st in by_stem:
        return by_stem[st], 0.97, "stem"
    best, best_sc = None, 0.0
    nq_s = st[:220]
    for x in bank:
        sc = SequenceMatcher(None, nq_s, stem(x["q"])[:220]).ratio()
        if sc > best_sc:
            best, best_sc = x, sc
    if best_sc >= threshold:
        return best, best_sc, "fuzzy"
    return None, best_sc, "miss"


def map_correct(rtn_opts: list[str], trainer_correct: list[str]) -> list[int]:
    flags = [0] * len(rtn_opts)
    if not trainer_correct:
        return flags
    tnorm = [norm(t) for t in trainer_correct]

    def score(a: str, b: str) -> float:
        if a == b:
            return 1.0
        if a and b:
            if re.search(r"(^|[^а-яa-z0-9])" + re.escape(a) + r"([^а-яa-z0-9]|$)", b) and len(a) <= 8:
                return 0.95
            if a in b or b in a:
                if min(len(a), len(b)) >= 2:
                    return 0.93
        return SequenceMatcher(None, a[:200], b[:200]).ratio()

    for i, opt in enumerate(rtn_opts):
        no = norm(opt)
        for tn in tnorm:
            if score(no, tn) >= 0.82:
                flags[i] = 1
                break
    if sum(flags) == 0 and len(trainer_correct) == 1 and rtn_opts:
        t = norm(trainer_correct[0])
        j = max(range(len(rtn_opts)), key=lambda k: score(norm(rtn_opts[k]), t))
        if score(norm(rtn_opts[j]), t) >= 0.68:
            flags[j] = 1
    return flags


def autosize(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def header(ws, titles):
    for c, t in enumerate(titles, 1):
        cell = ws.cell(1, c, t)
        cell.font = HEAD
        cell.fill = FILL_H
        cell.alignment = ALIGN


def parse_24test_docx(path: Path) -> list[dict]:
    doc = Document(str(path))
    qs = []
    for table in doc.tables:
        rows = [r.cells[0].text.strip() for r in table.rows]
        if len(rows) < 4:
            continue
        q = rows[1].strip()
        corr = re.sub(r"^Правильный ответ:\s*", "", rows[3], flags=re.I).strip()
        ntd = rows[4].strip() if len(rows) > 4 else ""
        if not q or not corr:
            continue
        qs.append({"q": q, "correct": [corr], "ntd": ntd, "opts": []})
    return qs


def join_src(prev: str, name: str) -> str:
    parts = [p.strip() for p in (prev or "").split("+") if p.strip()]
    if name not in parts:
        parts.append(name)
    return " + ".join(parts)


def load_trainers():
    merged = {}
    counts = {}

    def ingest(name: str, items: list[dict]):
        counts[name] = len(items)
        for it in items:
            rec = {**it, "src": name}
            key = stem(rec["q"])
            prev = merged.get(key)
            n_new = len([c for c in (rec.get("correct") or []) if c])
            n_old = len([c for c in (prev.get("correct") or []) if c]) if prev else -1
            if not prev or n_new > n_old:
                if prev:
                    rec["src"] = join_src(prev.get("src", ""), name)
                    if prev.get("ntd") and not rec.get("ntd"):
                        rec["ntd"] = prev["ntd"]
                merged[key] = rec
            elif prev:
                prev["src"] = join_src(prev.get("src", ""), name)

    for name, path in TRAINER_FILES:
        if path.exists():
            ingest(name, json.loads(path.read_text(encoding="utf-8")))
    npa_path = Path("npa_keys_остаток_V.json")
    if npa_path.exists():
        ingest("ПТЭЭП № 811 (открытый текст)", json.loads(npa_path.read_text(encoding="utf-8")))
    for name, path in DOCX_FILES:
        if path.exists():
            ingest(name, parse_24test_docx(path))
    return list(merged.values()), counts


def main():
    rtn = parse_rtn_v()
    trainer, src_counts = load_trainers()
    keyed_rows = []
    missing_rows = []
    how_c = {"exact": 0, "stem": 0, "fuzzy": 0, "miss": 0}
    src_hit = {}
    keyed = 0
    for idx, q in enumerate(rtn, 1):
        hit, sc, how = best_match(q["q"], trainer)
        how_c[how] += 1
        flags = map_correct(q["opts"], hit["correct"] if hit else [])
        ok = sum(flags) > 0
        if ok:
            keyed += 1
            src = hit.get("src", "") if hit else ""
            src_hit[src] = src_hit.get(src, 0) + 1
            keyed_rows.append(
                {
                    "idx": idx,
                    "n": q["n"],
                    "npa": q["npa"],
                    "q": q["q"],
                    "how": how,
                    "sc": sc,
                    "correct": [opt for opt, f in zip(q["opts"], flags) if f],
                    "opts": q["opts"],
                    "flags": flags,
                    "v1000": q["v1000"],
                    "v_above": q["v_above"],
                    "ntd": hit.get("ntd", "") if hit else "",
                    "src": src,
                }
            )
        else:
            missing_rows.append(
                {
                    "idx": idx,
                    "n": q["n"],
                    "npa": q["npa"],
                    "q": q["q"],
                    "best": sc,
                    "v1000": q["v1000"],
                    "v_above": q["v_above"],
                }
            )

    wb = Workbook()
    ws = wb.active
    ws.title = "Сводка"
    header(ws, ["Показатель", "Значение"])
    rows = [
        (
            "Источник ключей",
            "ЭБ 1260.25 (prombez24/tests24/24тест) плюс спецразделы prombez24: выше 6000 В, краны, КЛ, "
            "сварка, электродвигатели, ЭТЛ, электротермия, техэлектростанции (V до и выше 1000 В).",
        ),
        ("Банк вопросов", "Приложение 2 РТН, лист V, промышленные потребители, 701 вопрос."),
        ("Вопросов в тренажёре prombez24", src_counts.get("prombez24", 0)),
        ("Вопросов в спецразделах prombez24", src_counts.get("prombez24 спецразделы", 0)),
        ("Вопросов в дампе 24тест V выше 1000", src_counts.get("24тест V выше 1000", 0)),
        ("Вопросов в дампе 24тест IV до 1000", src_counts.get("24тест IV до 1000", 0)),
        ("Совпало с РТН и проставлен ключ", keyed),
        ("Разбивка по источникам (уникальные совпадения РТН)", "; ".join(f"{k}: {v}" for k, v in sorted(src_hit.items(), key=lambda x: -x[1])) or "—"),
        ("Без ключа", len(missing_rows)),
        ("Точное совпадение формулировки", how_c["exact"]),
        ("Совпадение без хвоста «согласно НПА»", how_c["stem"]),
        ("Нечёткое совпадение (опечатки РТН)", how_c["fuzzy"]),
        ("Нет пары", how_c["miss"]),
        ("Память чата", "Ключи хранятся в этом файле, а не в контексте диалога."),
    ]
    for i, (a, b) in enumerate(rows, 2):
        ws.cell(i, 1, a).font = FONT
        ws.cell(i, 2, b).font = FONT
        ws.cell(i, 1).alignment = ALIGN
        ws.cell(i, 2).alignment = ALIGN
        ws.row_dimensions[i].height = 48 if i in (2, 15, 16) else 22
    autosize(ws, [42, 110])
    ws.row_dimensions[1].height = 22

    ws_k = wb.create_sheet("Ключи")
    header(
        ws_k,
        [
            "№ п/п РТН",
            "№ в разделе",
            "НПА",
            "Вопрос",
            "Правильный ответ (по тренажёру)",
            "Как совпало",
            "Сходство",
            "До 1000 В",
            "Выше 1000 В",
            "Сайт",
            "Ссылка НПА в тренажёре",
        ],
    )
    for i, r in enumerate(keyed_rows, 2):
        vals = [
            r["idx"],
            r["n"],
            r["npa"],
            r["q"],
            " | ".join(r["correct"]),
            r["how"],
            round(r["sc"], 3),
            "+" if r["v1000"] else "",
            "+" if r["v_above"] else "",
            r.get("src", ""),
            r["ntd"],
        ]
        for c, v in enumerate(vals, 1):
            cell = ws_k.cell(i, c, v)
            cell.font = FONT
            cell.alignment = ALIGN
            if c == 5:
                cell.fill = FILL_OK
        ws_k.row_dimensions[i].height = 48
    autosize(ws_k, [12, 12, 40, 70, 70, 12, 12, 12, 14, 28, 40])

    ws_m = wb.create_sheet("Без ключа")
    header(ws_m, ["№ п/п РТН", "№ в разделе", "НПА", "Вопрос", "Лучшее сходство с тренажёром", "До 1000 В", "Выше 1000 В"])
    for i, r in enumerate(missing_rows, 2):
        vals = [r["idx"], r["n"], r["npa"], r["q"], round(r["best"], 3), "+" if r["v1000"] else "", "+" if r["v_above"] else ""]
        for c, v in enumerate(vals, 1):
            cell = ws_m.cell(i, c, v)
            cell.font = FONT
            cell.alignment = ALIGN
            if c == 4:
                cell.fill = FILL_NO
        ws_m.row_dimensions[i].height = 40
    autosize(ws_m, [12, 12, 40, 90, 18, 12, 14])

    ws_d = wb.create_sheet("Ключи_развёрнуто")
    header(ws_d, ["№ п/п РТН", "№ в разделе", "Тип", "Текст", "Ключ 1/0"])
    row = 2
    for r in keyed_rows:
        ws_d.cell(row, 1, r["idx"]).font = FONT
        ws_d.cell(row, 2, r["n"]).font = FONT
        ws_d.cell(row, 3, "Вопрос").font = FONT
        ws_d.cell(row, 4, r["q"]).font = FONT
        ws_d.cell(row, 4).alignment = ALIGN
        ws_d.row_dimensions[row].height = 36
        row += 1
        for opt, fl in zip(r["opts"], r["flags"]):
            ws_d.cell(row, 1, r["idx"]).font = FONT
            ws_d.cell(row, 2, r["n"]).font = FONT
            ws_d.cell(row, 3, "Ответ").font = FONT
            ws_d.cell(row, 4, opt).font = FONT
            ws_d.cell(row, 4).alignment = ALIGN
            cell = ws_d.cell(row, 5, fl)
            cell.font = FONT
            if fl:
                cell.fill = FILL_OK
            row += 1
    autosize(ws_d, [12, 12, 12, 100, 12])

    wb.save(OUT)
    print(
        f"saved {OUT} rtn={len(rtn)} trainer={len(trainer)} keyed={keyed} "
        f"missing={len(missing_rows)} exact={how_c['exact']} stem={how_c['stem']} "
        f"fuzzy={how_c['fuzzy']} miss={how_c['miss']} src={src_hit}"
    )


if __name__ == "__main__":
    main()
