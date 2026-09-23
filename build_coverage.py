#!/usr/bin/env python3
"""Сопоставляет три файла:

- горда.xlsx — города и адреса торговых центров (столбцы в исходнике съехали);
- Магнит_города для ПП_ Приложение 1.xlsx — объёмы обучения по филиалам;
- География.xlsx — города МегаФона для очного обучения.

Результат: Покрытие_ПП_Магнит.xlsx
"""

from __future__ import annotations

import re
from collections import defaultdict
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "исходные"
OUT = ROOT / "Покрытие_ПП_Магнит.xlsx"

# Филиал Магнита → (город площадки, пояснение). Считается покрытием агломерации.
SATELLITE = {
    "Абрау-Дюрсо": ("Новороссийск", "около 15 км, площадка в с. Гайдук"),
    "Адыгея": ("Новая Адыгея", "площадка в а. Новая Адыгея"),
    "Батайск": ("Ростов-на-Дону", "город-спутник Ростова-на-Дону"),
    "Дмитров": ("Лобня", "север Московской области, около 30 км"),
    "Зеленодольск": ("Казань", "около 40 км"),
    "Ивантеевка": ("Москва", "северо-восток Московской области"),
    "Колпино": ("Санкт-Петербург", "в составе агломерации"),
    "Кстово": ("Нижний Новгород", "около 30 км"),
    "Шушары": ("Санкт-Петербург", "в составе агломерации"),
    "Энгельс": ("Саратов", "город-спутник через Волгу"),
}

# Ближайший город с площадкой — только справка, в покрытие не входит.
NEAREST = {
    "Армавир": "Краснодар",
    "Березники": "Пермь",
    "Великий Новгород": "Санкт-Петербург",
    "Волгодонск": "Ростов-на-Дону",
    "Вологда": "Ярославль",
    "Дрезна": "Железнодорожный",
    "Камень-на-Оби": "Барнаул",
    "Камышин": "Волгоград",
    "Каневская": "Краснодар",
    "Коломна": "Подольск",
    "Кропоткин": "Краснодар",
    "Ликино-Дулево": "Ногинск",
    "Нижний Тагил": "Екатеринбург",
    "Ноябрьск": "Сургут",
    "Озёры": "Подольск",
    "Орск": "Оренбург",
    "Петрозаводск": "Санкт-Петербург",
    "Псков": "Санкт-Петербург",
    "Сергиев Посад": "Москва",
    "Славянск": "Краснодар",
    "Тамбов": "Липецк",
    "Шахты": "Ростов-на-Дону",
}


def squash(value) -> str:
    return re.sub(r"\s+", " ", str(value or "").replace("\xa0", " ")).strip()


def round1(value):
    if value is None or value == "":
        return None
    return round(float(value) + 1e-9, 1)


def base_branch(name: str) -> str:
    cleaned = re.sub(r"\s+\d+$", "", squash(name))
    for suffix in (" Восток", " Запад", " Север", " Юг", " Центр"):
        if cleaned.endswith(suffix):
            cleaned = cleaned[: -len(suffix)]
    return cleaned


# Площадка в другом населённом пункте, который относим к городу слева.
PLACE_TO_CITY = {
    "копейск": ("Челябинск", "г. Копейск"),
    "волжский": ("Волгоград", "г. Волжский"),
    "семилуки": ("Воронеж", "г. Семилуки"),
    "дубовое": ("Белгород", "п. Дубовое"),
    "засечное": ("Пенза", "с. Засечное"),
    "солонцы": ("Красноярск", "п. Солонцы"),
    "мирное": ("Симферополь", "с. Мирное"),
    "гайдук": ("Новороссийск", "с. Гайдук"),
    "прудное": ("Тула", "д. Прудное"),
    "игнатово": ("Иваново", "д. Игнатово"),
    "московский": ("Москва", "п. Московский"),
}

# Населённые пункты, которые не сворачиваются в более крупный город.
OWN_PLACES = {
    "черная грязь": ("Черная грязь", "д. Черная грязь"),
    "лобня": ("Лобня", "г. Лобня"),
    "томилино": ("Томилино", "рп. Томилино"),
    "железнодорожный": ("Железнодорожный", "г. Балашиха, мкр. Железнодорожный"),
    "новые псарьки": ("Ногинск", "д. Новые Псарьки"),
    "котельники": ("Котельники", "г. Котельники"),
    "подольск": ("Подольск", "г. Подольск"),
    "ликино": ("Одинцово", "д. Ликино, Одинцовский район"),
    "имени ленина": ("Апаринки", "пос. совхоза им. Ленина"),
    "серпухов": ("Серпухов", "г. Серпухов"),
    "новая адыгея": ("Новая Адыгея", "а. Новая Адыгея"),
}


def _word(needle: str, text: str) -> bool:
    return re.search(rf"(?<![0-9а-яё]){re.escape(needle)}(?![0-9а-яё])", text) is not None


def classify_address(addr: str) -> tuple[str, str]:
    """Город площадки и фактический населённый пункт из текста адреса.

    Город берётся из «г. …» или из отдельного населённого пункта.
    Названия улиц и областей («Московский проспект», «Челябинская область»,
    «ул. Волгоградская») не становятся городом.
    """
    text = squash(addr)
    low = text.lower().replace("ё", "е")

    for needle, (city, place) in OWN_PLACES.items():
        phrase = needle in {"новые псарьки", "черная грязь", "новая адыгея", "имени ленина"}
        if (phrase and needle in low) or _word(needle, low):
            return city, place

    for needle, (city, place) in PLACE_TO_CITY.items():
        if needle == "московский":
            if _word("п. московский", low) or "п. московский" in low:
                return city, place
            continue
        if _word(needle, low):
            return city, place

    titled = re.findall(
        r"г\.\s*([А-ЯЁ][а-яё]+(?:-на-[А-ЯЁ][а-яё]+|-[А-ЯЁ][а-яё]+)?(?:\s+[А-ЯЁ][а-яё]+)?)",
        text,
    )
    if titled:
        name = titled[-1]
        key = name.lower()
        if key in PLACE_TO_CITY:
            return PLACE_TO_CITY[key]
        return name, f"г. {name}"

    for name in (
        "Санкт-Петербург",
        "Нижний Новгород",
        "Ростов-на-Дону",
        "Набережные Челны",
        "Новая Адыгея",
        "Москва",
        "Владимир",
        "Казань",
        "Сочи",
    ):
        if _word(name.lower().replace("ё", "е"), low):
            return name, f"г. {name}"

    raise ValueError(f"Не удалось определить город: {text}")


def load_sites():
    wb = load_workbook(SRC / "горда.xlsx", data_only=True)
    ws = wb.active
    sites = []
    for idx, row in enumerate(ws.iter_rows(min_row=3, values_only=True), 3):
        label, addr = row[1], row[2]
        if not addr:
            continue
        city, place = classify_address(addr)
        sites.append(
            {
                "row": idx,
                "label": squash(label),
                "address": squash(addr),
                "city": city,
                "place": place,
                "aligned": squash(label).replace("МО (", "").replace(")", "").strip() == city
                or squash(label) == city,
            }
        )
    wb.close()
    return sites


def load_magnit():
    wb = load_workbook(SRC / "Магнит_города для ПП_ Приложение 1.xlsx", data_only=True)
    ws = wb.active
    rows = []
    for idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
        name = squash(row[0])
        if not name or name in {"Общий итог", "Кол-во тестов всего"}:
            continue
        rows.append(
            {
                "row": idx,
                "branch": name,
                "city": base_branch(name),
                "ot": round1(row[1]),
                "harm": round1(row[2]),
                "pp": round1(row[3]),
                "siz": round1(row[4]),
                "danger": round1(row[5]),
            }
        )
    wb.close()
    return rows


def load_geography():
    wb = load_workbook(SRC / "География.xlsx", data_only=True)
    ws = wb["Прил.1.3.1_География услуг "]
    rows = []
    filial = None
    for row in ws.iter_rows(min_row=7, values_only=True):
        if row[0]:
            filial = squash(row[0])
        city = squash(row[1]) if row[1] else ""
        if not city or city.isdigit():
            continue
        flag = squash(row[2])
        rows.append(
            {
                "filial": filial,
                "city": city.rstrip(),
                "in_person": flag.lower() == "да",
            }
        )
    wb.close()
    return rows


def geo_index(geography):
    index = {}
    for item in geography:
        index[item["city"].replace("ё", "е")] = item
    return index


def match_branch(city: str, polygons: dict, geo: dict):
    if city in polygons:
        status = "есть полигон"
        polygon_city = city
        note = ""
    elif city in SATELLITE:
        status = "полигон агломерации"
        polygon_city, note = SATELLITE[city]
    else:
        status = "нет полигона"
        polygon_city = ""
        if city in NEAREST:
            note = f"ближайшая площадка: {NEAREST[city]}, в покрытие не включена"
        else:
            note = "в списке площадок соседнего города нет"

    lookup = polygon_city or city
    info = geo.get(lookup.replace("ё", "е"))
    if info is None and city.replace("ё", "е") in geo:
        info = geo[city.replace("ё", "е")]
    return status, polygon_city, note, info


def copy_sources_from_uploads():
    """Кладёт в исходные/ только листы, нужные для сверки, без анкеты и реквизитов."""
    uploads = Path("/home/ubuntu/.cursor/projects/workspace/uploads")
    if not uploads.exists():
        return
    SRC.mkdir(exist_ok=True)
    mapping = {
        "______7342.xlsx": SRC / "горда.xlsx",
        "_________________________________1_453e.xlsx": SRC / "Магнит_города для ПП_ Приложение 1.xlsx",
    }
    for src_name, dest in mapping.items():
        source = uploads / src_name
        if source.exists() and not dest.exists():
            dest.write_bytes(source.read_bytes())

    geo_src = uploads / "__________34c0.xlsx"
    geo_dest = SRC / "География.xlsx"
    if geo_src.exists() and not geo_dest.exists():
        full = load_workbook(geo_src, data_only=True)
        sheet = full["Прил.1.3.1_География услуг "]
        thin = Workbook()
        out = thin.active
        out.title = "Прил.1.3.1_География услуг "
        for row in sheet.iter_rows(values_only=True):
            out.append(list(row[:3]))
        thin.save(geo_dest)
        full.close()


FILL = {
    "header": PatternFill("solid", fgColor="1F4E79"),
    "title": PatternFill("solid", fgColor="0D2A4A"),
    "green": PatternFill("solid", fgColor="C6EFCE"),
    "amber": PatternFill("solid", fgColor="FFE699"),
    "red": PatternFill("solid", fgColor="F8CBAD"),
    "blue": PatternFill("solid", fgColor="D6EAF8"),
    "gray": PatternFill("solid", fgColor="F2F2F2"),
    "white": PatternFill("solid", fgColor="FFFFFF"),
    "total": PatternFill("solid", fgColor="1F4E79"),
}
FONT_HEADER = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
FONT = Font(name="Calibri", size=11)
FONT_BOLD = Font(name="Calibri", bold=True, size=11)
FONT_TITLE = Font(name="Calibri", bold=True, color="FFFFFF", size=16)
FONT_WHITE = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
THIN = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)
WRAP = Alignment(wrap_text=True, vertical="center")
LEFT = Alignment(vertical="center", wrap_text=True)


def style_header(ws, row, columns):
    for col in range(1, columns + 1):
        cell = ws.cell(row, col)
        cell.fill = FILL["header"]
        cell.font = FONT_HEADER
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = THIN
    ws.auto_filter.ref = f"A{row}:{get_column_letter(columns)}{ws.max_row}"
    ws.freeze_panes = f"A{row + 1}"
    ws.auto_filter.ref = f"A{row}:{get_column_letter(columns)}{max(ws.max_row, row + 1)}"
    ws.row_dimensions[row].height = 32


def paint(cell, status):
    cell.alignment = LEFT
    cell.border = THIN
    cell.font = FONT
    if status == "есть полигон":
        cell.fill = FILL["green"]
    elif status == "полигон агломерации":
        cell.fill = FILL["amber"]
    elif status == "нет полигона":
        cell.fill = FILL["red"]
    elif status == "да":
        cell.fill = FILL["green"]
    elif status == "нет":
        cell.fill = FILL["red"]


def set_widths(ws, widths):
    for idx, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = width


def build():
    copy_sources_from_uploads()
    sites = load_sites()
    magnit = load_magnit()
    geography = load_geography()
    geo = geo_index(geography)

    by_city = defaultdict(list)
    for site in sites:
        by_city[site["city"]].append(site)

    polygons = {city: sites_ for city, sites_ in by_city.items()}

    matched = []
    for item in magnit:
        status, polygon_city, note, info = match_branch(item["city"], polygons, geo)
        addresses = polygons.get(polygon_city, [])
        matched.append(
            {
                **item,
                "status": status,
                "polygon_city": polygon_city,
                "note": note,
                "addresses": "\n".join(s["address"] for s in addresses),
                "places": ", ".join(dict.fromkeys(s["place"] for s in addresses)),
                "n_sites": len(addresses),
                "mf_filial": info["filial"] if info else "",
                "mf_city": info["city"] if info else "",
                "mf_in_person": ("Да" if info["in_person"] else "город есть, очно не требуется")
                if info
                else "нет в географии МегаФон",
            }
        )

    wb = Workbook()

    # --- Сводка ---
    summary = wb.active
    summary.title = "Сводка"
    summary.sheet_view.showGridLines = False
    summary.merge_cells("A1:F1")
    summary["A1"] = "Покрытие очной первой помощи Магнита площадками"
    summary["A1"].font = FONT_TITLE
    summary["A1"].fill = FILL["title"]
    summary["A1"].alignment = Alignment(vertical="center")
    summary.row_dimensions[1].height = 28

    summary.merge_cells("A2:F2")
    summary["A2"] = (
        "Адреса из горда.xlsx привязаны к городу по тексту адреса: в исходном файле подпись города "
        "и адрес стоят в разных строках. Покрытие считается по городу филиала Магнита."
    )
    summary["A2"].font = Font(name="Calibri", size=11, italic=True, color="595959")
    summary["A2"].alignment = Alignment(wrap_text=True, vertical="center")
    summary.row_dimensions[2].height = 36

    pp_total = round(sum(x["pp"] or 0 for x in matched), 1)
    siz_total = round(sum(x["siz"] or 0 for x in matched), 1)
    buckets = {
        "есть полигон": [x for x in matched if x["status"] == "есть полигон"],
        "полигон агломерации": [x for x in matched if x["status"] == "полигон агломерации"],
        "нет полигона": [x for x in matched if x["status"] == "нет полигона"],
    }

    headers = ["Показатель", "Филиалов", "Первая помощь, очно", "Доля ПП", "СИЗ, очно", "Комментарий"]
    for col, header in enumerate(headers, 1):
        cell = summary.cell(4, col, header)
        cell.fill = FILL["header"]
        cell.font = FONT_HEADER
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = THIN
    summary.row_dimensions[4].height = 30

    comments = {
        "есть полигон": "Город филиала совпадает с городом в адресе торгового центра",
        "полигон агломерации": "Площадка в соседнем городе, см. пояснение в листе филиалов",
        "нет полигона": "Своего адреса нет. Ближайший город указан справочно и в долю не входит",
    }
    order = ["есть полигон", "полигон агломерации", "нет полигона"]
    for offset, key in enumerate(order):
        rows = buckets[key]
        pp = round(sum(x["pp"] or 0 for x in rows), 1)
        siz = round(sum(x["siz"] or 0 for x in rows), 1)
        values = [key, len(rows), pp, pp / pp_total if pp_total else 0, siz, comments[key]]
        for col, value in enumerate(values, 1):
            cell = summary.cell(5 + offset, col, value)
            cell.font = FONT_BOLD if col == 1 else FONT
            cell.border = THIN
            cell.alignment = LEFT
            paint(cell, key if col == 1 else None)
            if col == 1:
                paint(cell, key)
        summary.cell(5 + offset, 3).number_format = "#,##0.0"
        summary.cell(5 + offset, 4).number_format = "0.0%"
        summary.cell(5 + offset, 5).number_format = "#,##0.0"

    total_row = 8
    totals = ["Итого", len(matched), pp_total, 1, siz_total, "Сумма филиалов без строки «Общий итог»"]
    for col, value in enumerate(totals, 1):
        cell = summary.cell(total_row, col, value)
        cell.font = FONT_WHITE
        cell.fill = FILL["total"]
        cell.border = THIN
        cell.alignment = LEFT
    summary.cell(total_row, 3).number_format = "#,##0.0"
    summary.cell(total_row, 4).number_format = "0.0%"
    summary.cell(total_row, 5).number_format = "#,##0.0"

    mf_yes = [g for g in geography if g["in_person"]]
    mf_gap = [g for g in mf_yes if g["city"].replace("ё", "е") not in polygons and g["city"] not in polygons]
    # Ессентуки покрывается Пятигорском
    mf_gap = [g for g in mf_gap if g["city"] != "Ессентуки"]

    summary["A10"] = "Площадки"
    summary["A10"].font = FONT_BOLD
    facts = [
        (11, "Адресов торговых центров", len(sites)),
        (12, "Городов, в которых адрес распознан", len(polygons)),
        (13, "Строк, где подпись города в горда.xlsx не совпала с адресом", sum(1 for s in sites if not _label_matches(s))),
        (14, "Городов МегаФона с отметкой «Да» (очно-заочно, нужен полигон)", len(mf_yes)),
        (15, "Из них без адреса площадки", len(mf_gap)),
        (16, "Ессентуки", "площадка есть в Пятигорске"),
    ]
    for row_idx, label, value in facts:
        summary.cell(row_idx, 1, label).font = FONT
        summary.cell(row_idx, 1).alignment = LEFT
        cell = summary.cell(row_idx, 2, value)
        cell.font = FONT_BOLD
        cell.alignment = LEFT

    summary["A18"] = "Города МегаФона с отметкой «Да», для которых адреса площадки нет"
    summary["A18"].font = FONT_BOLD
    summary.merge_cells("A19:F19")
    summary["A19"] = ", ".join(g["city"].strip() for g in mf_gap) if mf_gap else "—"
    summary["A19"].alignment = Alignment(wrap_text=True, vertical="center")
    summary.row_dimensions[19].height = 48

    summary["A21"] = "Как читать статусы"
    summary["A21"].font = FONT_BOLD
    notes = [
        "Зелёный — в этом городе есть адрес торгового центра.",
        "Жёлтый — филиал отнесён к площадке соседнего города (Батайск → Ростов-на-Дону, Колпино и Шушары → Санкт-Петербург, Энгельс → Саратов, Зеленодольск → Казань, Кстово → Нижний Новгород, Дмитров → Лобня, Ивантеевка → Москва, Абрау-Дюрсо → Новороссийск, Адыгея → Новая Адыгея).",
        "Красный — площадки нет. Город в колонке «Ближайшая площадка» не означает, что обучение можно провести там.",
        "Ликино-Дулево и деревня Ликино в Одинцово — разные населённые пункты. Площадка на Минском шоссе закрывает Одинцово, не Ликино-Дулево.",
        "Копейск отнесён к Челябинску, Семилуки — к Воронежу, Волжский — к Волгограду, Солонцы — к Красноярску, Дубовое — к Белгороду, Засечное — к Пензе, Мирное — к Симферополю, Гайдук — к Новороссийску.",
        "Тельмана и Звездное не сопоставлены: по названию филиала нельзя выбрать город.",
    ]
    for idx, note in enumerate(notes):
        summary.merge_cells(start_row=22 + idx, start_column=1, end_row=22 + idx, end_column=6)
        summary.cell(22 + idx, 1, note).font = FONT
        summary.cell(22 + idx, 1).alignment = Alignment(wrap_text=True, vertical="center")
        summary.row_dimensions[22 + idx].height = 32

    set_widths(summary, [78, 32, 28, 14, 16, 78])
    summary.page_setup.orientation = "landscape"
    summary.page_setup.fitToPage = True
    summary.page_setup.fitToWidth = 1
    summary.page_setup.fitToHeight = 1
    summary.sheet_properties.pageSetUpPr.fitToPage = True
    summary.oddHeader.left.text = "Покрытие ПП Магнит"
    summary.oddFooter.right.text = "Стр. &P из &N"
    summary.print_title_rows = "1:4"
    summary.page_setup.paperSize = summary.PAPERSIZE_A4
    summary.sheet_view.zoomScale = 110

    chart = BarChart()
    chart.type = "col"
    chart.title = "Первая помощь, очно"
    chart.y_axis.title = None
    chart.x_axis.title = None
    data = Reference(summary, min_col=3, min_row=4, max_row=7)
    cats = Reference(summary, min_col=1, min_row=5, max_row=7)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.shape = 4
    chart.legend = None
    chart.style = 10
    chart.y_axis.numFmt = "#,##0"
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showVal = True
    chart.width = 18
    chart.height = 8
    summary.add_chart(chart, "A29")

    # --- Филиалы ---
    branches = wb.create_sheet("Филиалы Магнит")
    branch_headers = [
        "Филиал",
        "Город",
        "Статус площадки",
        "Город площадки",
        "Площадок",
        "Населённый пункт в адресе",
        "Первая помощь, очно",
        "СИЗ, очно",
        "Повышенная опасность",
        "Охрана труда, дистант",
        "Вредные факторы, дистант",
        "Филиал МегаФон",
        "Город в географии МегаФон",
        "Очно-заочно у МегаФона",
        "Пояснение",
        "Адреса площадок",
    ]
    branches.append(branch_headers)
    for item in matched:
        branches.append(
            [
                item["branch"],
                item["city"],
                item["status"],
                item["polygon_city"],
                item["n_sites"] or None,
                item["places"],
                item["pp"],
                item["siz"],
                item["danger"],
                item["ot"],
                item["harm"],
                item["mf_filial"],
                item["mf_city"],
                item["mf_in_person"],
                item["note"],
                item["addresses"],
            ]
        )
    style_header(branches, 1, len(branch_headers))
    for row in branches.iter_rows(min_row=2, max_row=branches.max_row, max_col=len(branch_headers)):
        status = row[2].value
        for cell in row:
            cell.font = FONT
            cell.border = THIN
            cell.alignment = LEFT
        paint(row[2], status)
        for col in (7, 8, 9, 10, 11):
            row[col - 1].number_format = "#,##0.0"
        branches.row_dimensions[row[0].row].height = 30 if row[15].value else 18
    branches.auto_filter.ref = f"A1:P{branches.max_row}"
    set_widths(branches, [32, 24, 24, 22, 12, 36, 16, 14, 18, 18, 20, 28, 26, 32, 42, 70])
    branches.page_setup.orientation = "landscape"
    branches.page_setup.fitToPage = True
    branches.page_setup.fitToWidth = 1
    branches.page_setup.fitToHeight = 0
    branches.page_setup.paperSize = branches.PAPERSIZE_A4
    branches.page_setup.horizontalCentered = True
    branches.sheet_properties.pageSetUpPr.fitToPage = True
    branches.print_title_rows = "1:1"
    branches.page_setup.horizontalDpi = 300
    branches.oddFooter.right.text = "Филиалы Магнит, стр. &P из &N"
    branches.auto_filter.ref = f"A1:{get_column_letter(len(branch_headers))}{branches.max_row}"
    branches.freeze_panes = "A2"

    # --- Без площадки ---
    gaps = wb.create_sheet("Без полигона")
    gap_headers = ["Филиал", "Город", "Первая помощь, очно", "СИЗ, очно", "Филиал МегаФон", "Очно-заочно у МегаФона", "Пояснение"]
    gaps.append(gap_headers)
    gap_rows = sorted(
        [x for x in matched if x["status"] == "нет полигона"],
        key=lambda x: -(x["pp"] or 0),
    )
    for item in gap_rows:
        gaps.append(
            [
                item["branch"],
                item["city"],
                item["pp"],
                item["siz"],
                item["mf_filial"],
                item["mf_in_person"],
                item["note"],
            ]
        )
    style_header(gaps, 1, len(gap_headers))
    for row in gaps.iter_rows(min_row=2, max_row=gaps.max_row, max_col=len(gap_headers)):
        for cell in row:
            cell.font = FONT
            cell.border = THIN
            cell.alignment = LEFT
            cell.fill = FILL["red"]
        row[2].number_format = "#,##0.0"
        row[3].number_format = "#,##0.0"
    if gap_rows:
        last = gaps.max_row + 1
        gaps.cell(last, 1, "Итого").font = FONT_WHITE
        for col in range(1, 8):
            gaps.cell(last, col).fill = FILL["total"]
            gaps.cell(last, col).font = FONT_WHITE
            gaps.cell(last, col).border = THIN
        gaps.cell(last, 3, round(sum(x["pp"] or 0 for x in gap_rows), 1)).number_format = "#,##0.0"
        gaps.cell(last, 3).font = FONT_WHITE
        gaps.cell(last, 3).fill = FILL["total"]
        gaps.cell(last, 4, round(sum(x["siz"] or 0 for x in gap_rows), 1)).number_format = "#,##0.0"
        gaps.cell(last, 4).font = FONT_WHITE
        gaps.cell(last, 4).fill = FILL["total"]
    gaps.auto_filter.ref = f"A1:G{max(gaps.max_row, 2)}"
    gaps.freeze_panes = "A2"
    set_widths(gaps, [32, 24, 22, 14, 36, 36, 55])
    gaps.page_setup.orientation = "landscape"
    gaps.page_setup.fitToPage = True
    gaps.page_setup.fitToWidth = 1
    gaps.page_setup.fitToHeight = 0
    gaps.page_setup.paperSize = gaps.PAPERSIZE_A4
    gaps.sheet_properties.pageSetUpPr.fitToPage = True
    gaps.print_title_rows = "1:1"
    gaps.oddFooter.right.text = "Без полигона, стр. &P из &N"
    gaps.page_setup.horizontalCentered = True

    # --- Площадки ---
    places = wb.create_sheet("Площадки")
    place_headers = [
        "Город площадки",
        "Населённый пункт в адресе",
        "Адрес",
        "Строка в горда.xlsx",
        "Подпись в исходном файле",
        "Подпись совпала с адресом",
    ]
    places.append(place_headers)
    for site in sites:
        label_ok = _label_matches(site)
        places.append(
            [
                site["city"],
                site["place"],
                site["address"],
                site["row"],
                site["label"],
                "да" if label_ok else "нет",
            ]
        )
    style_header(places, 1, len(place_headers))
    for row in places.iter_rows(min_row=2, max_row=places.max_row, max_col=6):
        for cell in row:
            cell.font = FONT
            cell.border = THIN
            cell.alignment = LEFT
        paint(row[5], "да" if row[5].value == "да" else "нет")
        places.row_dimensions[row[0].row].height = 32
    places.auto_filter.ref = f"A1:F{places.max_row}"
    places.freeze_panes = "A2"
    set_widths(places, [24, 42, 88, 20, 28, 28])
    places.page_setup.orientation = "landscape"
    places.page_setup.fitToPage = True
    places.page_setup.fitToWidth = 1
    places.page_setup.fitToHeight = 0
    places.page_setup.paperSize = places.PAPERSIZE_A4
    places.sheet_properties.pageSetUpPr.fitToPage = True
    places.print_title_rows = "1:1"
    places.oddFooter.right.text = "Площадки, стр. &P из &N"
    places.page_setup.horizontalCentered = True

    # --- География МегаФон ---
    geo_sheet = wb.create_sheet("География МегаФон")
    geo_headers = [
        "Филиал МегаФон",
        "Город",
        "Очно-заочно (нужен полигон)",
        "Статус площадки",
        "Город площадки",
        "Адреса",
    ]
    geo_sheet.append(geo_headers)
    for item in geography:
        city = item["city"].strip()
        key = city.replace("ё", "е")
        if key in polygons or city in polygons:
            status = "есть полигон"
            polygon_city = key if key in polygons else city
        elif city == "Ессентуки":
            status = "полигон агломерации"
            polygon_city = "Пятигорск"
        else:
            status = "нет полигона"
            polygon_city = ""
        addresses = "\n".join(s["address"] for s in polygons.get(polygon_city, []))
        geo_sheet.append(
            [
                item["filial"],
                city,
                "Да" if item["in_person"] else "",
                status,
                polygon_city,
                addresses,
            ]
        )
    style_header(geo_sheet, 1, len(geo_headers))
    for row in geo_sheet.iter_rows(min_row=2, max_row=geo_sheet.max_row, max_col=6):
        for cell in row:
            cell.font = FONT
            cell.border = THIN
            cell.alignment = LEFT
        paint(row[3], row[3].value)
        if row[2].value == "Да" and row[3].value == "нет полигона":
            row[2].fill = FILL["red"]
            row[2].font = FONT_BOLD
        geo_sheet.row_dimensions[row[0].row].height = 28 if row[5].value else 18
    geo_sheet.auto_filter.ref = f"A1:F{geo_sheet.max_row}"
    geo_sheet.freeze_panes = "A2"
    set_widths(geo_sheet, [42, 28, 28, 24, 22, 78])
    geo_sheet.page_setup.orientation = "landscape"
    geo_sheet.page_setup.fitToPage = True
    geo_sheet.page_setup.fitToWidth = 1
    geo_sheet.page_setup.fitToHeight = 0
    geo_sheet.page_setup.paperSize = geo_sheet.PAPERSIZE_A4
    geo_sheet.sheet_properties.pageSetUpPr.fitToPage = True
    geo_sheet.print_title_rows = "1:1"
    geo_sheet.oddFooter.right.text = "География МегаФон, стр. &P из &N"
    geo_sheet.page_setup.horizontalCentered = True

    wb.save(OUT)
    return {
        "sites": len(sites),
        "cities": len(polygons),
        "branches": len(matched),
        "pp_total": pp_total,
        "buckets": {k: (len(v), round(sum(x["pp"] or 0 for x in v), 1)) for k, v in buckets.items()},
        "mf_gap": [g["city"].strip() for g in mf_gap],
        "misaligned": sum(1 for s in sites if not _label_matches(s)),
    }


def _label_matches(site) -> bool:
    label = site["label"].replace("ё", "е")
    city = site["city"]
    if not label:
        return False
    if label == city:
        return True
    inner = label.replace("МО (", "").replace(")", "").strip()
    aliases = {
        "Одинцово-Ликино": "Одинцово",
        "Стерлитомак": "Стерлитамак",
    }
    return aliases.get(inner, inner) == city


if __name__ == "__main__":
    stats = build()
    for key, value in stats.items():
        print(f"{key}: {value}")
