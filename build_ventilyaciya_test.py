#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Собрать тест «Вентиляция и отопление» по актуальному листу III РТН (01.09.2026)."""

from __future__ import annotations

import json
import re
from copy import copy
from difflib import SequenceMatcher
from html import unescape
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font

RTN_PATH = Path("/workspace/Приложение_2_РТН.xlsx")
JULY_PATH = Path("/tmp/july_vent_template.xlsx")
OUT_PATH = Path("/workspace/Вентиляция_и_отопление_тест.xlsx")
TSSMART_DIR = Path("/tmp/ts1309")
REPORT = Path("/tmp/vent_build_report.json")

FOLDER = "Вентиляция и отопление"
PROGRAM = "Теплопотребители. Вентиляция и отопление"
TAG = "Вентиляция и отопление"
FONT = Font(name="Times New Roman", size=11)
ALIGN = Alignment(wrap_text=True, vertical="top")

NO_KEY_COMMENT = "Ключ не проставлен: формулировка не совпала с источником ответов"


def norm(s: str | None) -> str:
    s = unescape(str(s or "")).replace("\xa0", " ").replace("ё", "е").replace("Ё", "Е")
    s = re.sub(r"[«»“”„]", '"', s)
    s = re.sub(r"[–—−]", "-", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def n2(s: str | None) -> str:
    s = norm(s).lower().rstrip(" .;:")
    s = re.sub(r"[?]+$", "", s)
    return s.strip()


def stem(s: str | None) -> str:
    s = n2(s)
    s = re.split(
        r",\s*согласно|\s+согласно\s+|,\s*утвержд|утвержд[её]нн|в соответствии с приказом|в соответствии с правилами",
        s,
        maxsplit=1,
    )[0]
    return s.strip(" ?.,;")


def parse_rtn_iii() -> list[dict]:
    wb = load_workbook(RTN_PATH, data_only=True)
    ws = wb["III"]
    qs: list[dict] = []
    cur = None
    section = ""
    for r in range(1, ws.max_row + 1):
        a, b = ws.cell(r, 1).value, ws.cell(r, 2).value
        if a == "№":
            continue
        if isinstance(a, str) and not str(a).strip().isdigit() and a.strip():
            section = str(a).strip()
            continue
        if a is None and b and cur is None:
            section = str(b).strip()
            continue
        is_num = isinstance(a, (int, float)) and not isinstance(a, bool)
        if is_num and b:
            if cur:
                qs.append(cur)
            cur = {
                "seq": len(qs) + 1,
                "n": int(a),
                "section": section,
                "q": str(b).strip(),
                "ans": [],
            }
            continue
        if cur is not None and b:
            cur["ans"].append(str(b).strip())
    if cur:
        qs.append(cur)
    return qs


def parse_july() -> list[dict]:
    wb = load_workbook(JULY_PATH, data_only=True)
    ws = wb["Questions"]
    qs: list[dict] = []
    cur = None
    for r in range(2, ws.max_row + 1):
        typ = ws.cell(r, 3).value
        txt = ws.cell(r, 4).value
        flag = ws.cell(r, 5).value
        if typ == "Вопрос":
            if cur:
                qs.append(cur)
            cur = {"n": ws.cell(r, 2).value, "q": str(txt or "").strip(), "ans": [], "flags": []}
        elif typ == "Ответ" and cur is not None:
            cur["ans"].append(str(txt or "").strip())
            try:
                fv = int(float(flag)) if flag not in (None, "") else 0
            except (TypeError, ValueError):
                fv = 1 if str(flag).strip() in {"1", "да", "Да"} else 0
            cur["flags"].append(1 if fv == 1 else 0)
    if cur:
        qs.append(cur)
    return qs


def parse_testsmart() -> list[dict]:
    items: list[dict] = []
    if not TSSMART_DIR.exists():
        return items
    for i in range(1, 31):
        p = TSSMART_DIR / f"full{i}.html"
        if not p.exists():
            continue
        html = p.read_text(errors="ignore")
        m = re.search(r"var ans0 = new Array\((.*?)\);", html, re.S)
        if not m:
            continue
        keys = [x.strip().strip("'\"") for x in m.group(1).split(",") if x.strip()]
        blocks = re.findall(
            r'<div class="quest">(.*?)<div class="ans">(.*?)</div></div></div>',
            html,
            re.S,
        )
        for (qhtml, ahtml), key in zip(blocks, keys):
            q = norm(re.sub(r"<img[^>]*>", "", qhtml))
            q = re.sub(r"^\d+\s*", "", q)
            opts = [
                norm(x)
                for x in re.findall(
                    r'<label class="form-check-label"[^>]*>(.*?)</label>', ahtml, re.S
                )
            ]
            flags = [1 if j < len(key) and key[j] == "1" else 0 for j in range(len(opts))]
            items.append({"q": q, "ans": opts, "flags": flags, "key": key})
    # unique by stem, keep first
    uniq: dict[str, dict] = {}
    for it in items:
        uniq.setdefault(stem(it["q"]), it)
        uniq.setdefault(n2(it["q"]), it)
    return list({id(v): v for v in uniq.values()}.values())


def map_flags(rtn_ans: list[str], src: dict | None, min_map: int = 2) -> list[int]:
    flags = [0] * len(rtn_ans)
    if not src:
        return flags
    src_ans = src["ans"]
    src_flags = src["flags"]
    mapped = 0
    for i, a in enumerate(rtn_ans):
        na = n2(a)
        best_j, best_sc = None, 0.0
        for j, ia in enumerate(src_ans):
            nia = n2(ia)
            sc = 1.0 if na == nia else SequenceMatcher(None, na[:220], nia[:220]).ratio()
            if sc > best_sc:
                best_sc = sc
                best_j = j
        if best_j is not None and best_sc >= 0.86:
            flags[i] = src_flags[best_j] if best_j < len(src_flags) else 0
            mapped += 1
    if mapped < min(min_map, max(1, len(rtn_ans) - 1)) and sum(flags) == 0:
        return [0] * len(rtn_ans)
    if sum(flags) == len(flags) and len(flags) > 1:
        return [0] * len(rtn_ans)
    return flags


def opt_fp(ans: list[str]) -> frozenset[str]:
    return frozenset(n2(a) for a in ans)


class BankMatcher:
    def __init__(self, bank: list[dict]):
        self.bank = bank
        self.by_n2 = {n2(x["q"]): x for x in bank}
        self.by_stem = {stem(x["q"]): x for x in bank}
        self.by_opts: dict[frozenset[str], list[dict]] = {}
        for x in bank:
            self.by_opts.setdefault(opt_fp(x["ans"]), []).append(x)

    def find(self, qtext: str, ans: list[str] | None = None) -> tuple[dict | None, float, str]:
        nq = n2(qtext)
        if nq in self.by_n2:
            return self.by_n2[nq], 1.0, "exact"
        st = stem(qtext)
        if st in self.by_stem:
            return self.by_stem[st], 0.97, "stem"
        if ans:
            fp = opt_fp(ans)
            hits = self.by_opts.get(fp) or []
            if len(hits) == 1:
                return hits[0], 0.93, "options"
            if len(hits) > 1:
                best = max(hits, key=lambda x: SequenceMatcher(None, st, stem(x["q"])).ratio())
                sc = SequenceMatcher(None, st, stem(best["q"])).ratio()
                if sc >= 0.45:
                    return best, sc, "options"
        best = None
        best_sc = 0.0
        nq_s = st[:220]
        for x in self.bank:
            sc = SequenceMatcher(None, nq_s, stem(x["q"])[:220]).ratio()
            if sc > best_sc:
                best_sc = sc
                best = x
        if best_sc >= 0.72:
            return best, best_sc, "fuzzy"
        return None, best_sc, "miss"


def law_924n_flags(q: dict) -> list[int] | None:
    """Ключи раздела ПОТ 924н по тексту правил (ред. 29.04.2025)."""
    qt = n2(q["q"])
    if "924н" not in qt and "министерства труда" not in qt:
        return None
    ans = q["ans"]
    nans = [n2(a) for a in ans]

    def pick(*needles: str) -> list[int] | None:
        flags = [0] * len(ans)
        hit = False
        for i, a in enumerate(nans):
            if all(n in a for n in needles):
                flags[i] = 1
                hit = True
        return flags if hit and 0 < sum(flags) < len(flags) else None

    # Q1 workers
    if "какие требования предъявляются к работникам" in qt and "эксплуатации" in qt:
        return pick("обучение безопасным", "стажировку")
    if "с какой периодичностью проводится повторный инструктаж" in qt:
        return pick("3 месяца")
    if "с какой периодичностью проводится проверка знаний требований охраны труда" in qt:
        return pick("12 месяцев")
    if "какие работники допускаются к выполнению работ по техническому обслуживанию" in qt:
        return pick("профессиональную подготовку", "охране труда")
    if "металлические площадки" in qt and "минимальной высоты" in qt:
        return pick("1,1")
    if "минимальная ширина площадок" in qt:
        return pick("0,8")
    if "гладких или изготовленных из прутковой" in qt:
        return pick("запрещается в любом случае")
    if "минимальная ширина лестниц" in qt:
        return pick("0,6")
    if "относительно лестниц, предназначенных для систематического" in qt:
        # 1,5 м и не более 50°
        flags = [0] * len(ans)
        for i, a in enumerate(nans):
            if "1,5" in a and "50" in a:
                flags[i] = 1
        return flags if sum(flags) == 1 else None
    if "расстояние от пола до низа площадок" in qt:
        return pick("2 м")
    if "в соответствии с каким документом выполняются работы повышенной опасности" in qt:
        return pick("нарядом-допуском")
    if "кем утверждается и может быть дополнен перечень работ" in qt:
        return pick("работодателем")
    if "какой документ оформляется при выполнении ремонтных и других работ подрядными" in qt:
        return pick("акт-допуск")
    if "не требуется разрешение технического руководителя" in qt:
        return pick("45")
    if "снятия давления" in qt and "неверным" in qt:
        # неверно: запорная арматура дренажей закрытого типа ... должна быть открыта
        flags = [0] * len(ans)
        for i, a in enumerate(nans):
            if "закрытого типа" in a and "открыта" in a:
                flags[i] = 1
        return flags if sum(flags) == 1 else None
    if "переносные электрические светильники" in qt and "повышенной опасностью" in qt:
        return pick("50 в")
    if "переносные электрические светильники" in qt and "особо неблагоприятных" in qt:
        return pick("12 в")
    if "температуру наружной поверхности" in qt:
        return pick("+45") or pick("45")
    if "в каком случае допускается эксплуатировать" in qt and "манометр" in " ".join(nans):
        return pick("пломба") or pick("клеймо")
    if "светильники во взрывозащищенном" in qt:
        return pick("12 в")
    return None


def law_220n_flags(q: dict) -> list[int] | None:
    qt = n2(q["q"])
    if "220н" not in qt:
        return None
    ans = q["ans"]
    nans = [n2(a) for a in ans]
    if "какие действия не допускаются" in qt:
        flags = [0] * len(ans)
        for i, a in enumerate(nans):
            if "препятствуя судорожным" in a and "не препятствуя" not in a:
                flags[i] = 1
        return flags if sum(flags) == 1 else None
    if "в каком месте не может оказываться" in qt:
        return _pick_ans(ans, nans, "угрожающих факторов")
    if "в каком случае не оказывается первая помощь" in qt:
        return _pick_ans(ans, nans, "отказа гражданина")
    if "какие средства не могут использоваться" in qt:
        return _pick_ans(ans, nans, "гражданской обороне") or _pick_ans(ans, nans, "чрезвычайным ситуациям")
    if "последовательности осуществления мероприятий" in qt or "какая последовательности осуществления" in qt:
        # полный перечень приложения 2 п.1-9
        flags = [0] * len(ans)
        for i, a in enumerate(nans):
            if ("оценки обстановки" in a or "оценка обстановки" in a) and "обзорного осмотра" in a and "лекарственных препаратов" in a:
                flags[i] = 1
        return flags if sum(flags) == 1 else None
    if "оценк" in qt and "безопасных условий" in qt:
        flags = [0] * len(ans)
        for i, a in enumerate(nans):
            if "перчатки медицинские" in a and "инородным телом" in a:
                flags[i] = 1
        return flags if sum(flags) == 1 else None
    if "временной остановке наружного кровотечения" in qt or "способы по временной остановке" in qt:
        flags = [0] * len(ans)
        for i, a in enumerate(nans):
            if "прямое давление на рану" in a and "кровоостанавливающего жгута" in a and "обширном повреждении" in a:
                flags[i] = 1
        return flags if sum(flags) == 1 else None
    if "отсутствии у пострадавшего признаков жизни" in qt:
        flags = [0] * len(ans)
        for i, a in enumerate(nans):
            if "твердой ровной поверхности" in a and "дефибриллятор" in a:
                flags[i] = 1
        return flags if sum(flags) == 1 else None
    if "наличии у пострадавшего признаков жизни" in qt and "отсутствии сознания" in qt:
        flags = [0] * len(ans)
        for i, a in enumerate(nans):
            if "устойчивого бокового положения" in a and "запрокидывание" in a:
                flags[i] = 1
        return flags if sum(flags) == 1 else None
    if "подробный осмотр и опрос" in qt:
        flags = [0] * len(ans)
        for i, a in enumerate(nans):
            if "осмотр головы" in a and "осмотр шеи" in a and "живота и таза" in a:
                flags[i] = 1
        return flags if sum(flags) == 1 else None
    return None


def _pick_ans(ans: list[str], nans: list[str], needle: str) -> list[int] | None:
    flags = [0] * len(ans)
    for i, a in enumerate(nans):
        if needle in a:
            flags[i] = 1
    return flags if 0 < sum(flags) < len(flags) else None


def law_511_flags(q: dict) -> list[int] | None:
    """Ключи по ПТЭ ОТиТПУ (приказ Минэнерго 511)."""
    qt = n2(q["q"])
    ans = q["ans"]
    nans = [n2(a) for a in ans]
    if "не устанавливается обязательная форма работы с персоналом" in qt and "стажировка" in qt:
        return _pick_ans(ans, nans, "руководящих работников")
    if "не проводится первичная проверка знаний" in qt:
        return _pick_ans(ans, nans, "менее 3 лет")
    if "продолжительности дублирования является верным" in qt:
        return _pick_ans(ans, nans, "12 рабочих смен")
    if "регулированию температуры сетевой воды" in qt:
        return _pick_ans(ans, nans, "30")
    if "технического руководителя" in qt and "ликвидации аварийных" in qt:
        return _pick_ans(ans, nans, "оперативном журнале")
    if "бак-аккумулятор должен подвергаться техническому диагностированию" in qt:
        return _pick_ans(ans, nans, "3 года")
    if "пригодность бака-аккумулятора" in qt:
        flags = [0] * len(ans)
        for i, a in enumerate(nans):
            if "50%" in a and "30%" in a and "20%" in a and a.find("50%") < a.find("30%"):
                flags[i] = 1
        return flags if sum(flags) == 1 else None
    if "появление капель в разъемных" in qt:
        return _pick_ans(ans, nans, "не увеличиваются")
    return None


def write_xlsx(template: Path, dest: Path, items: list[dict]) -> None:
    wb = load_workbook(template)
    ws = wb["Questions"]
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)
    row = 2
    for num, it in enumerate(items, 1):
        flags = it["flags"]
        n_ok = sum(flags)
        min_ok = n_ok if n_ok else 1
        qvals = [
            FOLDER,
            num,
            "Вопрос",
            it["q"],
            min_ok,
            "Выбор ответа",
            it.get("comment"),
            None,
            TAG,
        ]
        for c, v in enumerate(qvals, 1):
            cell = ws.cell(row, c, v)
            cell.font = FONT
            cell.alignment = ALIGN
        ws.row_dimensions[row].height = 42
        row += 1
        for text, flag in zip(it["ans"], flags):
            for c, v in enumerate([FOLDER, num, "Ответ", text, int(flag)], 1):
                cell = ws.cell(row, c, v)
                cell.font = FONT
                cell.alignment = ALIGN
            ws.row_dimensions[row].height = 22
            row += 1
    ws_p = wb["Теги программ"]
    ws_p.cell(2, 1).value = PROGRAM
    ws_p.cell(2, 2).value = TAG
    dest.parent.mkdir(parents=True, exist_ok=True)
    wb.save(dest)


def main() -> None:
    rtn = parse_rtn_iii()
    july = parse_july()
    ts = parse_testsmart()
    jm = BankMatcher(july)
    tm = BankMatcher(ts)

    items = []
    stats = {"keyed": 0, "law": 0, "ts": 0, "july": 0, "none": 0}
    new_vs_july = []
    for q in rtn:
        jhit, jsc, jhow = jm.find(q["q"], q["ans"])
        thit, tsc, thow = tm.find(q["q"], q["ans"])
        law = law_924n_flags(q) or law_220n_flags(q) or law_511_flags(q)
        flags = [0] * len(q["ans"])
        src = None
        comment = q["section"]
        if law and sum(law) > 0:
            flags = law
            src = "закон (924н/220н)"
            stats["law"] += 1
        if sum(flags) == 0 and thow != "miss":
            flags = map_flags(q["ans"], thit)
            if sum(flags) > 0:
                src = f"testsmart ({thow}, {tsc:.2f})"
                stats["ts"] += 1
        if sum(flags) == 0 and jhow != "miss":
            flags = map_flags(q["ans"], jhit)
            if sum(flags) > 0:
                src = f"июльский тест ({jhow}, {jsc:.2f})"
                stats["july"] += 1
        keyed = sum(flags) > 0
        if keyed:
            stats["keyed"] += 1
            comment = f"{q['section']}. Ключ: {src}"
        else:
            stats["none"] += 1
            comment = f"{q['section']}. {NO_KEY_COMMENT}"
        if jhow == "miss":
            new_vs_july.append({"seq": q["seq"], "q": q["q"][:200], "keyed": keyed, "src": src})
        items.append(
            {
                "q": q["q"],
                "ans": q["ans"],
                "flags": flags,
                "comment": comment,
                "seq": q["seq"],
                "section": q["section"],
            }
        )

    write_xlsx(JULY_PATH, OUT_PATH, items)
    REPORT.write_text(
        json.dumps(
            {
                "rtn": len(rtn),
                "july": len(july),
                "testsmart_unique": len(ts),
                "stats": stats,
                "unkeyed": [
                    {"seq": it["seq"], "q": it["q"][:180]}
                    for it in items
                    if sum(it["flags"]) == 0
                ],
                "new_vs_july": new_vs_july,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    print("RTN", len(rtn), "written", OUT_PATH)
    print("stats", stats)
    print("unkeyed", stats["none"], "new_vs_july", len(new_vs_july))


if __name__ == "__main__":
    main()
