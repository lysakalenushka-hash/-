#!/usr/bin/env python3
"""Сортирует населённые пункты из Города_по_частоте.xlsx по численности населения."""

from __future__ import annotations

import json
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "Города_по_размеру.xlsx"
WD = Path("/tmp/wd_cities.json")

NAMES = [
    "Москва",
    "Санкт-Петербург",
    "Екатеринбург",
    "Ростов-на-Дону",
    "Самара",
    "Краснодар",
    "Нижний Новгород",
    "Ярославль",
    "Воронеж",
    "Казань",
    "Новосибирск",
    "Омск",
    "Пермь",
    "Тюмень",
    "Волгоград",
    "Ижевск",
    "Саратов",
    "Уфа",
    "Чебоксары",
    "Челябинск",
    "Архангельск",
    "Астрахань",
    "Барнаул",
    "Белгород",
    "Брянск",
    "Владикавказ",
    "Владимир",
    "Иркутск",
    "Калуга",
    "Кемерово",
    "Киров",
    "Красноярск",
    "Курск",
    "Липецк",
    "Магнитогорск",
    "Набережные Челны",
    "Орел",
    "Оренбург",
    "Подольск",
    "Рязань",
    "Смоленск",
    "Сочи",
    "Ставрополь",
    "Сургут",
    "Тверь",
    "Ульяновск",
    "Иваново",
    "Калининград",
    "Котельники",
    "Новокузнецк",
    "Новороссийск",
    "Пенза",
    "Пятигорск",
    "Стерлитамак",
    "Тольятти",
    "Томск",
    "Тула",
    "Великий Новгород",
    "Вологда",
    "Железнодорожный",
    "Лобня",
    "Махачкала",
    "Мурманск",
    "Новая Адыгея",
    "Ноябрьск",
    "Петрозаводск",
    "Псков",
    "Саранск",
    "Севастополь",
    "Серпухов",
    "Совхоз имени Ленина",
    "Сыктывкар",
    "Тамбов",
    "Томилино",
    "Черная грязь",
    "Абакан",
    "Абрау-Дюрсо",
    "Адыгея",
    "Апаринки",
    "Армавир",
    "Балашиха",
    "Батайск",
    "Белоярский",
    "Березники",
    "Биробиджан",
    "Благовещенск",
    "Братск",
    "Владивосток",
    "Волгодонск",
    "Волжский",
    "Гайдук",
    "Грозный",
    "Дмитров",
    "Дрезна",
    "Дубовое",
    "Ессентуки",
    "Засечное",
    "Звездное",
    "Зеленодольск",
    "Ивантеевка",
    "Игнатово",
    "Ишим",
    "Йошкар-Ола",
    "Камень-на-Оби",
    "Камышин",
    "Каневская",
    "Карабулак",
    "Кисловодск",
    "Коломна",
    "Колпино",
    "Комсомольск-на-Амуре",
    "Копейск",
    "Кострома",
    "Кропоткин",
    "Кстово",
    "Курган",
    "Ликино",
    "Ликино-Дулево",
    "Люберцы",
    "Магадан",
    "Марий Эл",
    "Микунь",
    "Мирное",
    "Московский",
    "Надым",
    "Нальчик",
    "Нерюнгри",
    "Нижневартовск",
    "Нижний Тагил",
    "Новые Псарьки",
    "Ногинск",
    "Нягань",
    "Одинцово-Ликино",
    "Озёры",
    "Орск",
    "Петропавловск-Камчатский",
    "Пойковский",
    "Прудное",
    "Салехард",
    "Семилуки",
    "Сергиев Посад",
    "Симферополь",
    "Славянск",
    "Солонцы",
    "Тельмана",
    "Тобольск",
    "Улан-Удэ",
    "Уренгой",
    "Хабаровск",
    "Ханты-Мансийск",
    "Черкесск",
    "Чита",
    "Шахты",
    "Шушары",
    "Элиста",
    "Энгельс",
    "Югорск",
    "Южно-Сахалинск",
    "Якутск",
]

# Население, дата, тип, пояснение. Эти пункты нельзя выбрать как «самый крупный одноимённый».
OVERRIDES = {
    "Орел": (289503, "2025-01-01", "город", "Орловская область"),
    "Славянск": (60780, "2025-01-01", "город", "Славянск-на-Кубани"),
    "Черная грязь": (179, "2010-01-01", "деревня", "Солнечногорский район, с.п. Лунёвское"),
    "Ликино-Дулево": (33917, "2025-01-01", "город", "Орехово-Зуевский округ"),
    "Одинцово-Ликино": (1621, "2010-01-01", "деревня", "деревня Ликино, с.п. Жаворонковское"),
    "Ликино": (1621, "2010-01-01", "деревня", "с.п. Жаворонковское, Одинцовский округ"),
    "Игнатово": (610, "2010-01-01", "деревня", "Ивановский район"),
    "Прудное": (274, "2010-01-01", "деревня", "Тула"),
    "Мирное": (10921, "2021-01-01", "село", "Симферопольский район"),
    "Солонцы": (3786, "2010-01-01", "посёлок", "Емельяновский район"),
    "Железнодорожный": (151985, "2015-01-01", "микрорайон", "Балашиха, бывший город"),
    "Дубовое": (16101, "2021-01-01", "посёлок", "Белгородский район"),
    "Засечное": (26878, "2021-01-01", "село", "Пензенский район"),
    "Московский": (59218, "2020-01-01", "город", "п. Московский в Москве"),
    "Белоярский": (19994, "2021-01-01", "город", "Ханты-Мансийский автономный округ"),
    "Благовещенск": (239932, "2025-01-01", "город", "Амурская область"),
    "Волжский": (313034, "2025-01-01", "город", "Волгоградская область"),
    "Киров": (475871, "2025-01-01", "город", "Кировская область"),
    "Коломна": (132772, "2025-01-01", "город", "без городского округа"),
    "Люберцы": (238534, "2025-01-01", "город", "без городского округа"),
    "Тула": (456813, "2025-01-01", "город", "без городского округа"),
    "Сыктывкар": (219094, "2025-01-01", "город", "без городского округа"),
    "Колпино": (143914, "2023-01-01", "город в составе Санкт-Петербурга", ""),
    "Шушары": (132978, "2023-01-01", "муниципальный округ", "Санкт-Петербург"),
    "Совхоз имени Ленина": (7234, "2021-01-01", "посёлок", "Ленинский район Московской области"),
    "Тельмана": (23407, "2025-01-01", "город", "Тосненский район Ленинградской области"),
    "Адыгея": (500731, "2025-01-01", "республика", ""),
    "Марий Эл": (669854, "2024-01-01", "республика", ""),
    "Севастополь": (485386, "2022-01-01", "город", ""),
    "Надым": (44046, "2025-01-01", "город", "не река"),
    "Уренгой": (5845, "2021-01-01", "посёлок", "не Новый Уренгой"),
    "Новая Адыгея": (16968, "2024-01-01", "посёлок", "Тахтамукайский район"),
    "Гайдук": (10450, "2021-01-01", "село", "Новороссийск"),
    "Апаринки": (78, "2021-01-01", "деревня", "Ленинский район Московской области"),
    "Новые Псарьки": (88, "2010-01-01", "деревня", "Богородский округ"),
    "Звездное": (None, "", "", "по названию филиала населённый пункт не определён"),
}

CITY_TYPES = {
    "город",
    "большой город",
    "мегагород",
    "мегаполис",
    "метрополис",
    "столица",
    "бывшая столица страны",
    "город россии",
    "город федерального значения в россии",
    "город областного значения",
    "город областного подчинения рф",
    "город регионального значения в субъектах рф",
    "город окружного значения",
}


def latest_population(pops: list[tuple[str, int]]):
    if not pops:
        return None
    return sorted(pops)[-1]


def is_city(types: set[str]) -> bool:
    for type_name in types:
        lowered = type_name.lower()
        if "округ" in lowered:
            continue
        if lowered in CITY_TYPES or lowered.startswith("город "):
            return True
    return False


def is_place(types: set[str]) -> bool:
    if not types:
        return False
    text = " ".join(types).lower()
    markers = ("город", "село", "дерев", "посёл", "поселен", "хутор", "станиц", "республик", "нейборхуд", "населён", "мегаполис", "столиц")
    return any(marker in text for marker in markers)


def load_frequency():
    return [{"name": name} for name in NAMES]


def load_wikidata():
    data = json.loads(WD.read_text())
    items: dict[tuple[str, str], dict] = {}
    for binding in data["results"]["bindings"]:
        name = binding["name"]["value"]
        item_id = binding["item"]["value"].rsplit("/", 1)[-1]
        record = items.setdefault((name, item_id), {"types": set(), "pops": []})
        if "typeLabel" in binding:
            record["types"].add(binding["typeLabel"]["value"])
        if "pop" in binding:
            date = binding.get("date", {}).get("value", "")[:10]
            record["pops"].append((date, int(float(binding["pop"]["value"]))))
    return items


def choose_auto(name: str, items: dict):
    candidates = []
    for (item_name, _item_id), record in items.items():
        if item_name != name or not is_place(record["types"]):
            continue
        latest = latest_population(record["pops"])
        if not latest:
            continue
        candidates.append((latest[1], latest[0], is_city(record["types"]), record["types"]))
    if not candidates:
        return None
    cities = [item for item in candidates if item[2]]
    pool = cities or candidates
    population, date, _is_city, types = max(pool, key=lambda item: item[0])
    type_name = "город" if _is_city else sorted(types)[0]
    return population, date, type_name, ""


def build_rows():
    frequency = load_frequency()
    items = load_wikidata()
    result = []
    for row in frequency:
        name = row["name"]
        if name in OVERRIDES:
            population, date, type_name, note = OVERRIDES[name]
        else:
            chosen = choose_auto(name, items)
            if chosen is None:
                population, date, type_name, note = None, "", "", "численность не найдена"
            else:
                population, date, type_name, note = chosen
        result.append({**row, "population": population, "date": date, "type": type_name, "note": note})
    result.sort(key=lambda item: (item["population"] is None, -(item["population"] or 0), item["name"]))
    return result


FILL_HEADER = PatternFill("solid", fgColor="1F4E79")
FILL_TITLE = PatternFill("solid", fgColor="0D2A4A")
FILL_ZEBRA = PatternFill("solid", fgColor="F7F9FC")
FONT_HEADER = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
FONT = Font(name="Calibri", size=11)
FONT_TITLE = Font(name="Calibri", bold=True, color="FFFFFF", size=16)
THIN = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)
LEFT = Alignment(vertical="center", wrap_text=True)


def write_book(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Города"
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 120
    ws.merge_cells("A1:F1")
    ws["A1"] = "Населённые пункты по размеру"
    ws["A1"].font = FONT_TITLE
    ws["A1"].fill = FILL_TITLE
    ws["A1"].alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 28
    ws.merge_cells("A2:F2")
    ws["A2"] = (
        "Размер — численность населения по Викиданным. Для большинства городов это оценка на 1 января 2025 года, "
        "для части посёлков и сёл — перепись 2010 или 2021 года. Если название встречается несколько раз, "
        "взят пункт из адреса исходных файлов. Звездное оставлено без численности: по названию филиала посёлок не определяется."
    )
    ws["A2"].font = Font(name="Calibri", size=11, italic=True, color="595959")
    ws["A2"].alignment = LEFT
    ws.row_dimensions[2].height = 48

    headers = ["№", "Населённый пункт", "Население, чел.", "Дата", "Тип", "Пояснение"]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(4, col, header)
        cell.fill = FILL_HEADER
        cell.font = FONT_HEADER
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = THIN
    ws.row_dimensions[4].height = 30

    for index, row in enumerate(rows, 1):
        values = [index, row["name"], row["population"], row["date"], row["type"], row["note"]]
        for col, value in enumerate(values, 1):
            cell = ws.cell(4 + index, col, value if value != "" else None)
            cell.font = FONT
            cell.border = THIN
            cell.alignment = LEFT
            if index % 2 == 0:
                cell.fill = FILL_ZEBRA
        if row["population"] is not None:
            ws.cell(4 + index, 3).number_format = "#,##0"

    last = 4 + len(rows)
    ws.auto_filter.ref = f"A4:F{last}"
    ws.freeze_panes = "A5"
    for idx, width in enumerate([8, 32, 20, 14, 36, 52], 1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:4"
    ws.page_setup.horizontalCentered = True
    ws.oddFooter.right.text = "Города по размеру, стр. &P из &N"
    ws.page_setup.horizontalDpi = 300
    wb.save(OUT)


def main():
    rows = build_rows()
    write_book(rows)
    missing = [row["name"] for row in rows if row["population"] is None]
    print(f"rows {len(rows)} missing {missing}")
    for row in rows:
        pop = f"{row['population']:>10}" if row["population"] is not None else "         —"
        print(f"{pop}  {row['date']:10}  {row['type'][:22]:22}  {row['name']}")


if __name__ == "__main__":
    main()
