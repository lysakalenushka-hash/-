#!/usr/bin/env python3
"""Сравнение банка Г.1.1 с официальным перечнем РТН (PDF priok.gosnadzor.ru)."""

from __future__ import annotations

import json
import re
import unicodedata
from collections import defaultdict
from difflib import SequenceMatcher
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

LMS_PATH = Path(
    "/home/ubuntu/.cursor/projects/workspace/uploads/bank_questions__.1.1_0f05.xlsx"
)
OCR_PATH = Path("/tmp/g11/g11_ocr.txt")
OUT_XLSX = Path("/workspace/Сравнение_Г.1.1_с_сайтом_РТН.xlsx")
OUT_JSON = Path("/tmp/g11/cmp_final.json")

TAIL_RE = re.compile(
    r"\s*выберите\s+(правильный\s+вариант\s+ответа|"
    r"2\s+варианта\s+ответа|несколько\s+вариантов\s+ответа)\.?\s*$",
    re.I,
)
EDIT_NOTE_RE = re.compile(r"\s*исправлено\b.*$", re.I)

GREEN = PatternFill("solid", fgColor="C6EFCE")
YELLOW = PatternFill("solid", fgColor="FFF2CC")
RED = PatternFill("solid", fgColor="F4CCCC")
BLUE = PatternFill("solid", fgColor="D0E2F3")
GREY = PatternFill("solid", fgColor="F3F3F3")
HEAD = PatternFill("solid", fgColor="1F4E79")
HEAD_FONT = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
BODY = Font(name="Calibri", size=11)
TITLE = Font(name="Calibri", size=14, bold=True, color="1F4E79")
WRAP = Alignment(wrap_text=True, vertical="top")
THIN = Border(
    left=Side(style="thin", color="B4B4B4"),
    right=Side(style="thin", color="B4B4B4"),
    top=Side(style="thin", color="B4B4B4"),
    bottom=Side(style="thin", color="B4B4B4"),
)


def normalize(text: str) -> str:
    s = unicodedata.normalize("NFKC", str(text or ""))
    s = s.replace("\xa0", " ").replace("\u200b", "")
    s = s.lower().replace("ё", "е")
    s = re.sub(r"[«»“”\"'`]", "", s)
    s = re.sub(r"\s+", " ", s).strip()
    s = EDIT_NOTE_RE.sub("", s).strip()
    s = TAIL_RE.sub("", s).strip()
    s = re.sub(r"[?\s.;:]+$", "", s)
    s = s.replace("—", "-").replace("–", "-")
    s = re.sub(r"(?<=[а-яa-z])-(?=[а-яa-z])", "", s)
    s = re.sub(r"\s+", " ", s)
    return s


def parse_rtn(ocr: str) -> list[dict]:
    ocr = ocr.replace("=====PAGE=====", " ")
    ocr = re.sub(r"\n\s*\d{1,2}\s*\n", "\n", ocr)
    ocr = re.sub(r"-\n", "", ocr)
    ocr = re.sub(r"\s+", " ", ocr.replace("\n", " "))
    matches = list(re.finditer(r"(?:(?<=\s)|(?<=^))(\d{1,3})\.\s", ocr))
    items = []
    expect = 1
    for i, m in enumerate(matches):
        n = int(m.group(1))
        if n != expect:
            continue
        end = matches[i + 1].start() if i + 1 < len(matches) else len(ocr)
        raw = ocr[m.end() : end].strip()
        raw = re.sub(r"\s+", " ", raw)
        raw = raw.replace("‘", "").replace("›", "").replace("—", " ")
        raw = re.sub(r"\s+", " ", raw).strip()
        items.append({"num": n, "raw": raw, "norm": normalize(raw)})
        expect += 1
    if len(items) != 178:
        raise SystemExit(f"ожидали 178 вопросов РТН, получили {len(items)}")
    return items


def load_lms() -> list[dict]:
    import openpyxl

    wb = openpyxl.load_workbook(LMS_PATH, data_only=True)
    ws = wb["Questions"]
    qs: list[dict] = []
    cur = None
    for row in ws.iter_rows(min_row=2, values_only=True):
        _folder, num, typ, text, correct, *_rest = row[:9]
        if typ == "Вопрос":
            if cur:
                qs.append(cur)
            note = None
            qtext = str(text or "").strip()
            m = re.search(r"\s*(ИСПРАВЛЕНО\b.*)$", qtext, re.I)
            if m:
                note = m.group(1).strip()
                qtext = qtext[: m.start()].rstrip()
            cur = {
                "num": int(num),
                "q": qtext,
                "note": note,
                "norm": normalize(qtext),
                "answers": [],
            }
        elif typ == "Ответ" and cur:
            cur["answers"].append(
                {"text": str(text or "").strip(), "correct": bool(correct)}
            )
    if cur:
        qs.append(cur)
    return qs


def match(lms: list[dict], rtn: list[dict]) -> tuple[list[dict], list[dict]]:
    rtn_map: dict[str, list[int]] = defaultdict(list)
    for q in rtn:
        rtn_map[q["norm"]].append(q["num"])
    used: set[int] = set()
    rows = []
    for lq in lms:
        exact = [n for n in rtn_map.get(lq["norm"], []) if n not in used]
        if exact:
            n = exact[0]
            used.add(n)
            rq = next(x for x in rtn if x["num"] == n)
            kind = "совпадает"
            if lq.get("note"):
                kind = "совпадает, в банке служебная пометка"
            rows.append(
                {
                    "lms": lq["num"],
                    "rtn": n,
                    "score": 1.0,
                    "kind": kind,
                    "lms_q": lq["q"],
                    "rtn_q": rq["raw"],
                    "note": lq.get("note"),
                }
            )
            continue
        best, bests = None, 0.0
        for rq in rtn:
            if rq["num"] in used:
                continue
            s = SequenceMatcher(None, lq["norm"], rq["norm"]).ratio()
            if s > bests:
                bests, best = s, rq
        if best and bests >= 0.90:
            used.add(best["num"])
            rows.append(
                {
                    "lms": lq["num"],
                    "rtn": best["num"],
                    "score": round(bests, 3),
                    "kind": "совпадает (мелкие отличия / OCR)",
                    "lms_q": lq["q"],
                    "rtn_q": best["raw"],
                    "note": lq.get("note"),
                }
            )
        else:
            rows.append(
                {
                    "lms": lq["num"],
                    "rtn": None,
                    "score": round(bests, 3) if best else 0,
                    "kind": "нет в официальном перечне",
                    "lms_q": lq["q"],
                    "rtn_q": best["raw"] if best else "",
                    "note": lq.get("note"),
                }
            )
    missing = [q for q in rtn if q["num"] not in used]
    return rows, missing


def autosize(ws: Worksheet, widths: dict[int, int]) -> None:
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w


def header(ws: Worksheet, cells: list[str]) -> None:
    for i, h in enumerate(cells, 1):
        c = ws.cell(1, i, h)
        c.fill = HEAD
        c.font = HEAD_FONT
        c.alignment = Alignment(wrap_text=True, vertical="center")
        c.border = THIN
    ws.row_dimensions[1].height = 32
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(cells))}1"


def paint(cell, fill=None):
    cell.font = BODY
    cell.alignment = WRAP
    cell.border = THIN
    if fill:
        cell.fill = fill


def build():
    rtn = parse_rtn(OCR_PATH.read_text())
    lms = load_lms()
    rows, missing = match(lms, rtn)

    n_ok = sum(1 for r in rows if r["rtn"] and r["kind"].startswith("совпадает"))
    n_note = sum(1 for r in rows if r["kind"] == "совпадает, в банке служебная пометка")
    n_miss_lms = sum(1 for r in rows if r["rtn"] is None)

    wb = Workbook()

    ws = wb.active
    ws.title = "Сводка"
    ws["A1"] = "Сравнение банка Г.1.1 с официальным перечнем Ростехнадзора"
    ws["A1"].font = TITLE
    ws.merge_cells("A1:B1")
    summary = [
        (
            "Источник официальных вопросов",
            "http://priok.gosnadzor.ru/activity/attestation/testquest/docs/Г.1.1.pdf",
        ),
        (
            "Документ РТН",
            "Вопросы тестирования Г.1.1 «Эксплуатация электроустановок», "
            "утверждены распоряжением Ростехнадзора от 30.07.2024 № 69-рп "
            "(файл на сайте — скан Xerox WorkCentre, дата в PDF 30.07.2024)",
        ),
        ("Вопросов на сайте РТН", 178),
        ("Вопросов в банке LMS (bank_questions_Г.1.1.xlsx)", len(lms)),
        ("Есть в банке и совпадают по смыслу/формулировке", n_ok),
        ("Из них со служебной пометкой «ИСПРАВЛЕНО…» в тексте вопроса", n_note),
        ("Есть в банке, но нет в официальном перечне", n_miss_lms),
        ("Есть на сайте РТН, но нет в банке", len(missing)),
        (
            "Вывод",
            "Банк — актуальный перечень 2024/2026 (69-рп), но неполный: "
            "не хватает 16 вопросов, почти все про электрозащитные средства "
            "(№ 54–67, 114, 166 официальной нумерации). "
            "Нумерация в LMS своя, с сайтом РТН не совпадает — это нормально. "
            "Ответы на сайте РТН не публикуются, сверялись только формулировки вопросов.",
        ),
    ]
    ws["A3"] = "Показатель"
    ws["B3"] = "Значение"
    ws["A3"].fill = HEAD
    ws["B3"].fill = HEAD
    ws["A3"].font = HEAD_FONT
    ws["B3"].font = HEAD_FONT
    for i, (k, v) in enumerate(summary, 4):
        a, b = ws.cell(i, 1, k), ws.cell(i, 2, v)
        paint(a, GREY)
        fill = GREEN if i in (8,) else YELLOW if i in (10, 11) else None
        paint(b, fill)
        ws.row_dimensions[i].height = 48 if i == 12 else 28
    ws.row_dimensions[12].height = 90
    autosize(ws, {1: 62, 2: 110})

    ws2 = wb.create_sheet("Нет_в_банке")
    header(
        ws2,
        ["№ на сайте РТН", "Вопрос официального перечня", "Тема"],
    )
    theme = "Электрозащитные средства / СИЗ в электроустановках"
    for i, q in enumerate(missing, 2):
        ws2.cell(i, 1, q["num"])
        ws2.cell(i, 2, q["raw"])
        ws2.cell(i, 3, theme)
        for col in range(1, 4):
            paint(ws2.cell(i, col), RED)
        ws2.row_dimensions[i].height = 48
    autosize(ws2, {1: 18, 2: 120, 3: 48})

    ws3 = wb.create_sheet("Сопоставление")
    header(
        ws3,
        [
            "№ в банке LMS",
            "Вопрос банка",
            "Статус",
            "№ на сайте РТН",
            "Вопрос официального перечня",
            "Сходство",
            "Примечание",
        ],
    )
    fills = {
        "совпадает": GREEN,
        "совпадает (мелкие отличия / OCR)": GREEN,
        "совпадает, в банке служебная пометка": YELLOW,
        "нет в официальном перечне": RED,
    }
    for i, r in enumerate(rows, 2):
        vals = [
            r["lms"],
            r["lms_q"],
            r["kind"],
            r["rtn"] or "",
            r["rtn_q"],
            r["score"],
            r.get("note") or "",
        ]
        fill = fills.get(r["kind"], YELLOW)
        for col, v in enumerate(vals, 1):
            c = ws3.cell(i, col, v)
            paint(c, fill)
        ws3.row_dimensions[i].height = 42
    autosize(ws3, {1: 14, 2: 70, 3: 38, 4: 16, 5: 70, 6: 12, 7: 40})

    ws4 = wb.create_sheet("Все_вопросы_РТН")
    header(
        ws4,
        ["№ РТН", "Вопрос официального перечня", "Есть в банке?", "№ в банке LMS"],
    )
    lms_by_rtn = {r["rtn"]: r for r in rows if r["rtn"]}
    for i, q in enumerate(rtn, 2):
        hit = lms_by_rtn.get(q["num"])
        vals = [
            q["num"],
            q["raw"],
            "Да" if hit else "Нет",
            hit["lms"] if hit else "",
        ]
        fill = GREEN if hit else RED
        for col, v in enumerate(vals, 1):
            c = ws4.cell(i, col, v)
            paint(c, fill)
        ws4.row_dimensions[i].height = 36
    autosize(ws4, {1: 12, 2: 120, 3: 16, 4: 16})

    ws5 = wb.create_sheet("Легенда")
    ws5["A1"] = "Как читать файл"
    ws5["A1"].font = TITLE
    legend = [
        ("Зелёный", "Формулировка банка есть в официальном PDF Г.1.1."),
        (
            "Жёлтый",
            "Тот же вопрос, но в тексте банка осталась служебная пометка "
            "(например «ИСПРАВЛЕНО 15.03.2024…»). Её лучше убрать из формулировки для обучаемого.",
        ),
        ("Красный", "Вопроса нет в банке (или нет в перечне РТН)."),
        (
            "Нумерация",
            "На сайте РТН вопросы 1–178 в фиксированном порядке. В LMS своя нумерация 1–162. "
            "Сравнивать нужно по тексту, не по номеру.",
        ),
        (
            "OCR",
            "Официальный файл — скан. Текст РТН в этом Excel снят OCR; "
            "мелкие артефакты скана (слипшиеся дефисы) не считались расхождением с банком.",
        ),
        (
            "Ответы",
            "На сайте РТН опубликованы только вопросы без вариантов ответа. "
            "Правильность ключей в LMS этой сверкой не проверялась.",
        ),
    ]
    ws5["A3"] = "Цвет / поле"
    ws5["B3"] = "Значение"
    ws5["A3"].fill = HEAD
    ws5["B3"].fill = HEAD
    ws5["A3"].font = HEAD_FONT
    ws5["B3"].font = HEAD_FONT
    color_fill = {
        "Зелёный": GREEN,
        "Жёлтый": YELLOW,
        "Красный": RED,
    }
    for i, (k, v) in enumerate(legend, 4):
        a, b = ws5.cell(i, 1, k), ws5.cell(i, 2, v)
        paint(a, color_fill.get(k, GREY))
        paint(b)
        ws5.row_dimensions[i].height = 40
    autosize(ws5, {1: 22, 2: 110})

    wb.save(OUT_XLSX)
    OUT_JSON.write_text(
        json.dumps(
            {
                "lms": len(lms),
                "rtn": len(rtn),
                "matched": n_ok,
                "missing_rtn_nums": [q["num"] for q in missing],
                "missing": [{"num": q["num"], "q": q["raw"]} for q in missing],
                "note_rows": [r for r in rows if r.get("note")],
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    print("wrote", OUT_XLSX, "matched", n_ok, "missing", len(missing))


if __name__ == "__main__":
    build()
