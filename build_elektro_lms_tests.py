#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Собрать 13 тестов электробезопасности: вопросы РТН, ключи только при совпадении с indtec."""

from __future__ import annotations

import json
import re
import shutil
from collections import defaultdict
from copy import copy
from difflib import SequenceMatcher
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font

RTN_PATH = Path("Приложение_2_РТН.xlsx")
INDTEC_JSON = Path("/tmp/indtec_elektro.json")
TEMPLATE = Path("/home/ubuntu/.cursor/projects/workspace/uploads/____________________________f0a6.xlsx")
OUTDIR = Path("тесты_электробезопасность_РТН")

FONT = Font(name="Times New Roman", size=11)
ALIGN = Alignment(wrap_text=True, vertical="top")


def norm(s: str) -> str:
    s = (s or "").lower().replace("\xa0", " ").replace("ё", "е")
    s = re.sub(r"[«»“”]", '"', s)
    return re.sub(r"\s+", " ", s).strip()


def stem(s: str) -> str:
    s = norm(s)
    s = re.split(r",\s*согласно|\s+согласно\s+|,\s*утвержд|утвержд[её]нн", s, maxsplit=1)[0]
    return s.strip(" ?.,;")


def tokens(s: str) -> list[str]:
    return [w for w in re.findall(r"[а-яa-z0-9]+", stem(s)) if len(w) > 3]


def parse_rtn(ws):
    npa = None
    questions = []
    cur = None
    for r in range(1, ws.max_row + 1):
        a, b = ws.cell(r, 1).value, ws.cell(r, 2).value
        if a == "№":
            continue
        is_num = isinstance(a, int) or (isinstance(a, str) and str(a).strip().isdigit())
        if a and not is_num:
            npa = str(a).strip()
            continue
        if is_num and b:
            if cur:
                questions.append(cur)

            def plus(c):
                return ws.cell(r, c).value == "+"

            cur = {
                "q": str(b).strip(),
                "ans": [],
                "npa": npa or "",
                "u1": plus(3),
                "u2": plus(4),
                "g2": plus(5),
                "g3": plus(6),
                "g4": plus(7),
                "g5": plus(8),
            }
            continue
        if cur and b and (a is None or a == ""):
            cur["ans"].append(str(b).strip())
    if cur:
        questions.append(cur)
    return questions


def filt(qs, above: bool, group: int):
    key = f"g{group}"
    out = []
    for q in qs:
        if not q[key]:
            continue
        if above and q["u2"]:
            out.append(q)
        if (not above) and q["u1"]:
            out.append(q)
    return out


class Matcher:
    def __init__(self, bank: list[dict]):
        self.bank = bank
        self.by_exact = {norm(x["q"]): x for x in bank}
        self.by_stem = {}
        self.inv = defaultdict(list)
        for i, x in enumerate(bank):
            st = stem(x["q"])
            self.by_stem.setdefault(st, x)
            for t in tokens(x["q"])[:10]:
                self.inv[t].append(i)

    def find(self, qtext: str):
        nq = norm(qtext)
        if nq in self.by_exact:
            return self.by_exact[nq], 1.0, "exact"
        st = stem(qtext)
        if st in self.by_stem:
            return self.by_stem[st], 0.97, "stem"
        cand = []
        seen = set()
        for t in tokens(qtext)[:8]:
            for i in self.inv.get(t, []):
                if i not in seen:
                    seen.add(i)
                    cand.append(i)
        if not cand:
            cand = range(min(len(self.bank), 40))
        best = None
        best_sc = 0.0
        nq_s = st[:200]
        for i in cand:
            x = self.bank[i]
            sc = SequenceMatcher(None, nq_s, stem(x["q"])[:200]).ratio()
            if sc > best_sc:
                best_sc = sc
                best = x
        if best_sc >= 0.82:
            return best, best_sc, "fuzzy"
        return None, best_sc, "miss"


def map_flags(rtn_ans, ind):
    flags = [0] * len(rtn_ans)
    if not ind:
        return flags, 0
    ind_ans = [norm(a) for a in ind["answers"]]
    ind_flags = ind["flags"]
    mapped = 0
    for i, a in enumerate(rtn_ans):
        na = norm(a)
        best_j, best_sc = None, 0
        for j, ia in enumerate(ind_ans):
            sc = 1.0 if na == ia else SequenceMatcher(None, na[:180], ia[:180]).ratio()
            if sc > best_sc:
                best_sc = sc
                best_j = j
        if best_j is not None and best_sc >= 0.86:
            flags[i] = ind_flags[best_j] if best_j < len(ind_flags) else 0
            mapped += 1
    if mapped < max(2, min(3, len(rtn_ans) - 1)) and sum(flags) == 0:
        return [0] * len(rtn_ans), mapped
    return flags, mapped


def copy_style(src, dst):
    if src.has_style:
        dst.font = copy(src.font)
        dst.border = copy(src.border)
        dst.fill = copy(src.fill)
        dst.number_format = src.number_format
        dst.protection = copy(src.protection)
        dst.alignment = copy(src.alignment)


def write_test(template: Path, dest: Path, folder: str, program: str, tag: str, items: list[dict]):
    shutil.copy2(template, dest)
    wb = load_workbook(dest)
    ws = wb["Questions"]
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)
    proto_q = [ws.cell(1, c) for c in range(1, 10)]  # header only; use row1 fonts
    row = 2
    for num, it in enumerate(items, 1):
        flags = it["flags"]
        n_ok = sum(flags) or 0
        min_ok = n_ok if n_ok else 1
        qvals = [folder, num, "Вопрос", it["q"], min_ok, "Выбор ответа", it.get("comment"), None, tag]
        for c, v in enumerate(qvals, 1):
            cell = ws.cell(row, c, v)
            cell.font = FONT
            cell.alignment = ALIGN
        ws.row_dimensions[row].height = 36
        row += 1
        for text, flag in zip(it["ans"], flags):
            avals = [folder, num, "Ответ", text, int(flag)]
            for c, v in enumerate(avals, 1):
                cell = ws.cell(row, c, v)
                cell.font = FONT
                cell.alignment = ALIGN
            ws.row_dimensions[row].height = 22
            row += 1
    ws_p = wb["Теги программ"]
    ws_p.cell(2, 1).value = program
    ws_p.cell(2, 2).value = tag
    wb.save(dest)


def program_name(meta):
    volt = "До и выше 1000 В" if meta["above"] else "До 1000 В"
    kind = "Промышленные" if meta["kind"] == "промышленные" else "Непромышленные"
    g = {2: "II", 3: "III", 4: "IV", 5: "V"}[meta["group"]]
    return f"Потребители электроэнергии. {kind}. {volt} {g} Группа"


def file_name(meta):
    volt = "dovyshe1000" if meta["above"] else "do1000"
    kind = "prom" if meta["kind"] == "промышленные" else "neprom"
    return f"{meta['lms']}_{kind}_{volt}_G{meta['group']}.xlsx"


def main():
    rtn_wb = load_workbook(RTN_PATH, data_only=True)
    industrial = parse_rtn(rtn_wb["V"])
    nonind = parse_rtn(rtn_wb["Vн"])
    indtec = json.loads(INDTEC_JSON.read_text(encoding="utf-8"))
    OUTDIR.mkdir(exist_ok=True)

    report_rows = []
    for slug, pack in indtec.items():
        meta = pack["meta"]
        src = nonind if meta["kind"] == "непромышленные" else industrial
        rtn_qs = filt(src, meta["above"], meta["group"])
        matcher = Matcher(pack["questions"])
        items = []
        stats = defaultdict(int)
        for q in rtn_qs:
            hit, sc, how = matcher.find(q["q"])
            flags, mapped = map_flags(q["ans"], hit if how != "miss" else None)
            keyed = sum(flags) > 0
            stats[how] += 1
            if keyed:
                stats["keyed"] += 1
            comment = q["npa"]
            if keyed:
                comment = f"{q['npa']}. Ключ: indtec ({how}, сходство {sc:.2f})"
            else:
                comment = f"{q['npa']}. Ключ не проставлен: нет надёжного совпадения с indtec"
            items.append({"q": q["q"], "ans": q["ans"], "flags": flags, "comment": comment})
        prog = program_name(meta)
        tag = prog
        dest = OUTDIR / file_name(meta)
        write_test(TEMPLATE, dest, folder=prog, program=prog, tag=tag, items=items)
        report_rows.append(
            {
                "lms": meta["lms"],
                "file": dest.name,
                "program": prog,
                "rtn": len(rtn_qs),
                "indtec": len(pack["questions"]),
                "exact": stats["exact"],
                "stem": stats["stem"],
                "fuzzy": stats["fuzzy"],
                "miss": stats["miss"],
                "keyed": stats["keyed"],
            }
        )
        print(
            f"{meta['lms']} {dest.name} RTN={len(rtn_qs)} keyed={stats['keyed']} "
            f"exact={stats['exact']} stem={stats['stem']} fuzzy={stats['fuzzy']} miss={stats['miss']}"
        )

    # summary xlsx
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Сводка"
    headers = [
        "Код LMS",
        "Файл",
        "Программа",
        "Вопросов РТН",
        "Вопросов indtec",
        "Точное совпадение",
        "Совпадение без хвоста НПА",
        "Нечёткое совпадение",
        "Нет пары",
        "Проставлен ключ",
    ]
    ws.append(headers)
    for r in sorted(report_rows, key=lambda x: x["lms"]):
        ws.append(
            [r["lms"], r["file"], r["program"], r["rtn"], r["indtec"], r["exact"], r["stem"], r["fuzzy"], r["miss"], r["keyed"]]
        )
    ws["A13"] = (
        "Вопросы всегда из приложения 2 РТН (листы V / Vн). "
        "Ответы-ключи взяты с https://indtec.ru/test-24-elektrobezopasnost/ только если формулировка вопроса совпала. "
        "Страница energobezopasnost — аттестация Г1/Г2, к этим 13 тестам потребителей не относится."
    )
    ws.merge_cells("A13:J16")
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 36
    ws.column_dimensions["C"].width = 70
    for col in "DEFGHIJ":
        ws.column_dimensions[col].width = 18
    wb.save(OUTDIR / "Сводка_совпадение_РТН_и_indtec.xlsx")
    print("done", OUTDIR)


if __name__ == "__main__":
    main()
