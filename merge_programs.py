#!/usr/bin/env python3
"""Сравнение двух PDF-файлов с программами и формирование общей таблицы."""

import re
import fitz
import pandas as pd
from collections import defaultdict

FILE1 = "/home/ubuntu/.cursor/projects/workspace/uploads/_________________________2026__1__653e.pdf"
FILE2 = "/home/ubuntu/.cursor/projects/workspace/uploads/________________0c24.pdf"
OUTPUT = "/workspace/общая_таблица_программ.xlsx"

FILE1_LABEL = "План разработки 2026"
FILE2_LABEL = "Каталог курсов (СДО)"

DIRECTIONS = {"ОТ", "ПБ", "БДД", "ПП"}
STATUSES = {"в разработке", "разработан", "запрос на разработку"}
WORK_TYPES = {"актуализация", "разработка"}
MONTHS = {
    "январь", "февраль", "март", "апрель", "май", "июнь",
    "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь",
}


def normalize_name(name: str) -> str:
    if not name:
        return ""
    s = name.lower().strip()
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"[«»\"'„""]", "", s)
    s = re.sub(r"\s*;\s*", ";", s)
    return s


def is_date(s: str) -> bool:
    return bool(re.fullmatch(r"\d{2}\.\d{2}\.\d{4}", s))


def is_id(s: str) -> bool:
    return bool(re.fullmatch(r"\d{4,5}", s))


def is_hours(s: str) -> bool:
    return bool(re.fullmatch(r"\d{1,4};?", s))


def parse_file1(path: str) -> list[dict]:
    doc = fitz.open(path)
    text = doc[0].get_text("text")
    lines = [l.strip() for l in text.split("\n") if l.strip()]

    start = 0
    for i, line in enumerate(lines):
        if line in WORK_TYPES or line == "ФМСМ":
            start = i
            break

    records = []
    i = start
    n = len(lines)

    def peek(k=0):
        return lines[i + k] if i + k < n else None

    def take():
        nonlocal i
        val = lines[i]
        i += 1
        return val

    while i < n:
        line = lines[i]

        if line == "Октябрь":
            break
        if line in {"+", "-", "готово", "Готово", "загрузка", "в работе", "не тест", "16ч"}:
            i += 1
            continue
        if re.match(r"^[\+\-\s]+$", line):
            i += 1
            continue
        if line.startswith(("1.", "2.", "3.", "4.", "5.")):
            i += 1
            continue

        company = ""
        work_type = ""
        if line == "ФМСМ":
            company = take()
            line = take()
        if line in WORK_TYPES:
            work_type = take()
        else:
            i += 1
            continue

        name_parts = []
        while i < n:
            cur = lines[i]
            if cur in WORK_TYPES or cur == "ФМСМ" or cur == "Октябрь":
                break
            if cur in DIRECTIONS:
                break
            if cur in STATUSES:
                break
            if is_date(cur) or is_id(cur):
                break
            if cur in MONTHS or re.match(r"^[а-яё]+-[а-яё]+$", cur):
                break
            if cur in {"+", "-", "готово", "Готово", "загрузка", "в работе"}:
                break
            name_parts.append(take())

        if not name_parts:
            continue

        direction = ""
        if i < n and lines[i] in DIRECTIONS:
            direction = take()

        hours = ""
        if i < n and is_hours(lines[i]):
            hours = take().rstrip(";")

        training_type = ""
        if i < n and lines[i] == "повышение":
            take()
            if i < n and lines[i] == "квалификации":
                take()
                training_type = "повышение квалификации"
        elif i < n and lines[i] == "проф.":
            take()
            if i < n and lines[i] == "переподготовка":
                take()
                training_type = "проф. переподготовка"
        elif i < n and lines[i] == "прф.переподготовка":
            take()
            training_type = "проф. переподготовка"

        customer = ""
        contact = ""
        note = ""
        status = ""
        author = ""
        dates = []
        course_id = ""

        while i < n:
            cur = lines[i]
            if cur in WORK_TYPES or cur == "ФМСМ" or cur == "Октябрь":
                break
            if cur in STATUSES:
                status = take()
                continue
            if is_id(cur):
                course_id = take()
                break
            if is_date(cur):
                dates.append(take())
                continue
            if cur in MONTHS or re.match(r"^[а-яё]+-[а-яё]+$", cur):
                dates.append(take())
                continue
            if re.match(r"^[\+\-\s]+$", cur):
                i += 1
                continue

            if not customer:
                customer = take()
                if i < n and lines[i] == "труд»":
                    customer += " " + take()
                elif i < n and lines[i].startswith("НЕДРА"):
                    customer += " " + take()
                continue

            if not contact and not status:
                nxt = peek()
                if nxt in STATUSES or nxt in WORK_TYPES or nxt == "ФМСМ":
                    pass
                elif "запрос" in cur or (nxt and "разработку" in nxt):
                    contact = take()
                    if i < n and "разработку" in lines[i]:
                        contact += " " + take()
                    continue

            if not contact:
                contact = take()
                continue

            if not note and ("замечаниям" in cur or "актуализир" in cur or "тесты" in cur):
                note_parts = [take()]
                while i < n:
                    nxt = lines[i]
                    if nxt in WORK_TYPES or nxt == "ФМСМ" or nxt in STATUSES:
                        break
                    if nxt in MONTHS or is_date(nxt) or is_id(nxt):
                        break
                    note_parts.append(take())
                note = " ".join(note_parts)
                continue

            if not author and re.match(r"^[А-Яа-яЁё]+\s+[А-Яа-яЁё]\.?$", cur):
                author = take()
                continue

            if not status and cur in STATUSES:
                status = take()
                continue

            if is_id(cur):
                course_id = take()
                break

            i += 1

        name = " ".join(name_parts)
        records.append({
            "Наименование программы": name,
            "Тип работы": work_type,
            "Компания": company,
            "Направление": direction,
            "Часов": hours,
            "Вид обучения": training_type,
            "Заказчик": customer,
            "Контактное лицо": contact,
            "Примечание": note,
            "Статус": status,
            "Автор курса": author,
            "Даты": "; ".join(dates),
            "ID курса": course_id,
            "Источник": FILE1_LABEL,
        })

    return records


def col_for_x_file2(x: float) -> str:
    if x < 90:
        return "ID"
    if x < 230:
        return "Наименование курса"
    if x < 268:
        return "Описание"
    if x < 276:
        return "Цена"
    if x < 296:
        return "Бейдж"
    if x < 306:
        return "Часы"
    if x < 360:
        return "Тип обучения"
    if x < 376:
        return "ФРДО/РОЛ"
    if x < 390:
        return "ProgramId"
    if x < 404:
        return "Код ЕИСОТ"
    if x < 418:
        return "Программа ЕИС"
    if x < 430:
        return "Часы ЕИС"
    return "Ссылка/год"


def parse_file2(path: str) -> list[dict]:
    doc = fitz.open(path)
    records = []

    for page in doc:
        words = page.get_text("words")
        by_y = defaultdict(list)
        for x0, y0, x1, y1, word, block, line, wno in words:
            yk = round(y0, 0)
            by_y[yk].append((x0, word))

        for y in sorted(by_y.keys()):
            items = sorted(by_y[y], key=lambda t: t[0])
            if not items or items[0][0] >= 95:
                continue
            if not re.fullmatch(r"\d{4,5}", items[0][1]):
                continue

            cols = defaultdict(list)
            for x, word in items:
                cols[col_for_x_file2(x)].append(word)

            row = {k: " ".join(v) for k, v in cols.items()}
            name = row.get("Наименование курса", "").strip()
            if not name or name.upper() == "NULL":
                continue

            records.append({
                "ID": row.get("ID", ""),
                "Наименование программы": name,
                "Описание": row.get("Описание", ""),
                "Цена": row.get("Цена", ""),
                "Бейдж": row.get("Бейдж", ""),
                "Часы": row.get("Часы", ""),
                "Тип обучения": row.get("Тип обучения", ""),
                "ФРДО/РОЛ": row.get("ФРДО/РОЛ", ""),
                "ProgramId": row.get("ProgramId", ""),
                "Код ЕИСОТ": row.get("Код ЕИСОТ", ""),
                "Программа ЕИС": row.get("Программа ЕИС", ""),
                "Часы ЕИС": row.get("Часы ЕИС", ""),
                "Ссылка/год": row.get("Ссылка/год", ""),
                "Источник": FILE2_LABEL,
            })

    return records


def names_match(n1: str, n2: str) -> bool:
    if not n1 or not n2:
        return False
    if n1 == n2:
        return True
    if len(n1) >= 20 and (n1 in n2 or n2 in n1):
        return True
    # Сравнение по ключевым словам (без служебных слов)
    stop = {"программа", "профессионального", "обучения", "переподготовки", "повышения", "квалификации", "по"}
    w1 = {w for w in re.split(r"\W+", n1) if len(w) > 3 and w not in stop}
    w2 = {w for w in re.split(r"\W+", n2) if len(w) > 3 and w not in stop}
    if len(w1) >= 3 and len(w2) >= 3:
        common = w1 & w2
        if len(common) >= min(3, len(w1) * 0.6):
            return True
    return False


def merge_records(rec1: list[dict], rec2: list[dict]) -> pd.DataFrame:
    for r in rec1:
        r["_norm"] = normalize_name(r["Наименование программы"])
    for r in rec2:
        r["_norm"] = normalize_name(r["Наименование программы"])

    used1 = set()
    used2 = set()
    pairs = []

    # 1. Сопоставление по ID
    id2_map = {}
    for idx, r in enumerate(rec2):
        cid = str(r.get("ID", "")).strip()
        if cid:
            id2_map.setdefault(cid, []).append(idx)

    for i1, r1 in enumerate(rec1):
        cid = str(r1.get("ID курса", "")).strip()
        if not cid or cid not in id2_map:
            continue
        for i2 in id2_map[cid]:
            if i2 not in used2:
                pairs.append((i1, i2, "ID"))
                used1.add(i1)
                used2.add(i2)
                break

    # 2. Точное совпадение названия
    for i1, r1 in enumerate(rec1):
        if i1 in used1:
            continue
        for i2, r2 in enumerate(rec2):
            if i2 in used2:
                continue
            if r1["_norm"] and r1["_norm"] == r2["_norm"]:
                pairs.append((i1, i2, "название"))
                used1.add(i1)
                used2.add(i2)
                break

    # 3. Частичное совпадение названия
    for i1, r1 in enumerate(rec1):
        if i1 in used1:
            continue
        for i2, r2 in enumerate(rec2):
            if i2 in used2:
                continue
            if names_match(r1["_norm"], r2["_norm"]):
                pairs.append((i1, i2, "название (частичное)"))
                used1.add(i1)
                used2.add(i2)
                break

    rows = []

    def build_row(r1, r2, source):
        row = {
            "Наименование программы": (
                r1["Наименование программы"] if r1 is not None else r2["Наименование программы"]
            ),
            "Источник": source,
        }
        f1_cols = [
            "Тип работы", "Компания", "Направление", "Часов", "Вид обучения",
            "Заказчик", "Контактное лицо", "Примечание", "Статус", "Автор курса",
            "Даты", "ID курса",
        ]
        for col in f1_cols:
            row[f"[План 2026] {col}"] = r1.get(col, "") if r1 is not None else ""
        f2_cols = [
            "ID", "Описание", "Цена", "Бейдж", "Часы", "Тип обучения",
            "ФРДО/РОЛ", "ProgramId", "Код ЕИСОТ", "Программа ЕИС", "Часы ЕИС", "Ссылка/год",
        ]
        for col in f2_cols:
            row[f"[Каталог] {col}"] = r2.get(col, "") if r2 is not None else ""
        return row

    for i1, i2, _ in pairs:
        rows.append(build_row(rec1[i1], rec2[i2], "Оба файла"))

    for i1, r1 in enumerate(rec1):
        if i1 not in used1:
            rows.append(build_row(r1, None, FILE1_LABEL))

    for i2, r2 in enumerate(rec2):
        if i2 not in used2:
            rows.append(build_row(None, r2, FILE2_LABEL))

    rows.sort(key=lambda r: r["Наименование программы"].lower())
    df = pd.DataFrame(rows)
    cols = ["Наименование программы", "Источник"] + [
        c for c in df.columns if c not in ("Наименование программы", "Источник")
    ]
    return df[cols]


def main():
    print("Парсинг файла 1...")
    rec1 = parse_file1(FILE1)
    print(f"  Найдено программ: {len(rec1)}")

    print("Парсинг файла 2...")
    rec2 = parse_file2(FILE2)
    print(f"  Найдено программ: {len(rec2)}")

    print("Объединение...")
    merged = merge_records(rec1, rec2)
    print(f"  Всего в общей таблице: {len(merged)}")
    print(f"  Только в плане 2026: {(merged['Источник'] == FILE1_LABEL).sum()}")
    print(f"  Только в каталоге: {(merged['Источник'] == FILE2_LABEL).sum()}")
    print(f"  В обоих файлах: {(merged['Источник'] == 'Оба файла').sum()}")

    with pd.ExcelWriter(OUTPUT, engine="openpyxl") as writer:
        merged.to_excel(writer, sheet_name="Общая таблица", index=False)
        pd.DataFrame([{k: v for k, v in r.items() if not k.startswith("_")} for r in rec1]).to_excel(
            writer, sheet_name="План разработки 2026", index=False
        )
        pd.DataFrame([{k: v for k, v in r.items() if not k.startswith("_")} for r in rec2]).to_excel(
            writer, sheet_name="Каталог курсов", index=False
        )

        both = merged[merged["Источник"] == "Оба файла"][
            ["Наименование программы", "Источник", "[План 2026] ID курса", "[Каталог] ID", "[План 2026] Статус"]
        ]
        both.to_excel(writer, sheet_name="Совпадения", index=False)

    csv_path = OUTPUT.replace(".xlsx", ".csv")
    merged.to_csv(csv_path, index=False, encoding="utf-8-sig")

    print(f"\nРезультат сохранён:")
    print(f"  Excel: {OUTPUT}")
    print(f"  CSV:   {csv_path}")

    print("\n--- Совпадения (оба файла) ---")
    for _, r in merged[merged["Источник"] == "Оба файла"].iterrows():
        print(f"  • {r['Наименование программы'][:70]}")


if __name__ == "__main__":
    main()
