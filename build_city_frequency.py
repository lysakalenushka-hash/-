#!/usr/bin/env python3
"""Уникальные населённые пункты из всех столбцов трёх файлов, по убыванию частоты."""

from __future__ import annotations

import re
from collections import defaultdict
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
UPLOADS = Path("/home/ubuntu/.cursor/projects/workspace/uploads")
OUT = ROOT / "Города_по_частоте.xlsx"

FILES = [
    (
        "горда.xlsx",
        [
            UPLOADS / "______7342.xlsx",
            ROOT / "исходные" / "горда.xlsx",
        ],
    ),
    (
        "География.xlsx",
        [
            UPLOADS / "__________34c0.xlsx",
            ROOT / "исходные" / "География.xlsx",
        ],
    ),
    (
        "Магнит_города для ПП_ Приложение 1.xlsx",
        [
            UPLOADS / "_________________________________1_453e.xlsx",
            ROOT / "исходные" / "Магнит_города для ПП_ Приложение 1.xlsx",
        ],
    ),
]

# Длинные названия раньше коротких, чтобы «Нижний Новгород» не распался на «Новгород».
MULTI = [
    "Петропавловск-Камчатский",
    "Комсомольск-на-Амуре",
    "Набережные Челны",
    "Ростов-на-Дону",
    "Санкт-Петербург",
    "Великий Новгород",
    "Южно-Сахалинск",
    "Ханты-Мансийск",
    "Ликино-Дулево",
    "Сергиев Посад",
    "Камень-на-Оби",
    "Нижний Новгород",
    "Абрау-Дюрсо",
    "Новая Адыгея",
    "Нижний Тагил",
    "Новые Псарьки",
    "Черная грязь",
    "Йошкар-Ола",
    "Улан-Удэ",
    "Марий Эл",
    "Совхоз имени Ленина",
]

# Написания одного и того же пункта.
ALIASES = {
    "стерлитомак": "Стерлитамак",
    "москвы": "Москва",
    "орел": "Орел",
    "орёл": "Орел",
    "озеры": "Озёры",
    "озёры": "Озёры",
    "совхоза им.ленина": "Совхоз имени Ленина",
    "совхоза им. ленина": "Совхоз имени Ленина",
}

SKIP_LABELS = {
    "город",
    "филиал",
    "филиалы",
    "общий итог",
    "кол-во тестов всего",
    "адрес торгового центра",
}

SKIP_BRANCH_SUFFIXES = (" Восток", " Запад", " Север", " Юг", " Центр")

# Не населённые пункты, даже если стоят с заглавной.
NOT_PLACES = {
    "приложение",
    "форма",
    "лот",
    "лоты",
    "итого",
    "да",
    "нет",
    "участник",
    "заказчик",
    "мегафон",
    "смарта",
    "пригородный",
    "ленин",
    "область",
    "край",
    "республика",
    "район",
    "округ",
    "шоссе",
    "улица",
    "проспект",
    "дом",
    "строение",
    "корпус",
    "офис",
    "московский проспект",
}


def squash(value) -> str:
    return re.sub(r"\s+", " ", str(value or "").replace("\xa0", " ")).strip()


def key_of(name: str) -> str:
    return name.lower().replace("ё", "е").replace("—", "-").strip()


def canonical(name: str) -> str:
    cleaned = squash(name).replace("—", "-")
    mapped = ALIASES.get(key_of(cleaned))
    if mapped:
        return mapped
    return cleaned


def resolve_path(candidates: list[Path]) -> Path:
    for path in candidates:
        if path.exists():
            return path
    raise FileNotFoundError(candidates[0])


def base_branch(name: str) -> str:
    cleaned = re.sub(r"\s+\d+$", "", squash(name))
    for suffix in SKIP_BRANCH_SUFFIXES:
        if cleaned.endswith(suffix):
            cleaned = cleaned[: -len(suffix)].strip()
    return cleaned


def add_name(gazetteer: dict[str, str], name: str) -> None:
    name = canonical(name)
    if not name or key_of(name) in NOT_PLACES or key_of(name) in SKIP_LABELS:
        return
    if len(name) < 3 or len(name) > 40 or re.search(r"\d", name) or "," in name:
        return
    gazetteer.setdefault(key_of(name), name)


def discover_marked(text: str, gazetteer: dict[str, str]) -> None:
    """Населённые пункты, которые в тексте названы прямо: г., деревня, село, п., с., рп., аул."""
    raw = squash(text)
    lowered = raw.lower().replace("ё", "е")
    for phrase in MULTI:
        if key_of(phrase) in lowered or phrase.lower().replace("ё", "е") in lowered:
            add_name(gazetteer, phrase)

    # «c.Гайдук» с латинской c
    normalized = re.sub(r"(?<![A-Za-zА-Яа-яЁё])c\.(?=[А-ЯЁ])", "с.", raw)

    patterns = [
        r"г\.(?!\s*о\b)\s*([А-ЯЁ][а-яё]+(?:-на-[А-ЯЁ][а-яё]+|-[А-ЯЁ][а-яё]+)?)",
        r"(?<![А-Яа-яЁё])([А-ЯЁ][а-яё]+(?:-на-[А-ЯЁ][а-яё]+|-[А-ЯЁ][а-яё]+)?)\s+г\b",
        r"деревн[яеи]\s+([А-ЯЁ][а-яё]+(?:\s+[А-ЯЁа-яё]+){0,2})",
        r"село\s+([А-ЯЁ][а-яё]+)",
        r"мкр\.\s*([А-ЯЁ][а-яё]+)",
        r"([А-ЯЁ][а-яё]+)\s+рп\.",
        r"(?<![А-Яа-яЁё])а\.\s*([А-ЯЁ][а-яё]+(?:\s+[А-ЯЁ][а-яё]+)?)",
        r"(?<!с\.)(?<![А-Яа-яЁё])п\.\s*([А-ЯЁ][а-яё]+)",
        r"(?<![А-Яа-яЁё])с\.\s*([А-ЯЁ][а-яё]+)",
        r"(?<![А-Яа-яЁё])д\.\s*([А-ЯЁ][а-яё]+(?:\s+[а-яё]+){0,1})",
        r"города\s+([А-ЯЁ][а-яё]+)",
        r"добыча\s+([А-ЯЁ][а-яё]+)",
    ]
    for pattern in patterns:
        for match in re.findall(pattern, normalized):
            piece = squash(match)
            # «д. Черная грязь» и похожие пары, где второе слово с маленькой буквы.
            if key_of(piece).startswith("черная"):
                piece = "Черная грязь"
            if piece.lower().startswith("совхоза"):
                piece = "Совхоз имени Ленина"
            # «Нижний» из «г. Нижний Новгород» раскрывается отдельным списком MULTI.
            if piece in {"Нижний", "Набережные", "Великий", "Сергиев", "Новая", "Новые", "Черная"}:
                continue
            if re.search(r"\d", piece) or "," in piece:
                continue
            if " " in piece and key_of(piece) not in {key_of(item) for item in MULTI}:
                continue
            add_name(gazetteer, piece)


def discover_label(text: str, gazetteer: dict[str, str]) -> None:
    raw = squash(text)
    if not raw:
        return
    paren = re.fullmatch(r"МО\s*\((.+)\)", raw)
    if paren:
        add_name(gazetteer, paren.group(1))
        return
    if key_of(raw) in SKIP_LABELS:
        return
    # Филиал целиком («Дальневосточный филиал (ДВФ)») — не город.
    if "филиал" in raw.lower():
        return
    add_name(gazetteer, base_branch(raw))


def is_place_label(title: str, sheet: str, column: int, text: str) -> bool:
    """Короткое имя в столбце городов, а не адрес и не заголовок анкеты."""
    if re.search(r"\d", text) or len(text) > 40 or "," in text:
        return False
    lowered = text.lower()
    if lowered.startswith(("приложение", "к ", "адреса ", "обучение", "форма", "оказание", "использование", "общие", "безопасн")):
        return False
    if title.startswith("горда") and sheet == "Лист1" and column == 2:
        return True
    if title.startswith("Магнит") and sheet == "Список" and column == 1:
        return True
    if title.startswith("География") and sheet.startswith("Прил.1.3.1") and column == 2:
        return True
    return False


def load_cells():
    cells = []
    for title, candidates in FILES:
        path = resolve_path(candidates)
        wb = load_workbook(path, data_only=True, read_only=True)
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.strip():
                        cells.append((title, sheet, cell.column, cell.coordinate, squash(cell.value)))
        wb.close()
    return cells


def build_gazetteer(cells):
    gazetteer: dict[str, str] = {}
    for phrase in MULTI:
        add_name(gazetteer, phrase)
    for title, sheet, column, _coord, text in cells:
        discover_marked(text, gazetteer)
        if is_place_label(title, sheet, column, text):
            discover_label(text, gazetteer)
    gazetteer.pop("пригородный", None)
    gazetteer.pop("жаворонковское", None)
    gazetteer.pop("луневское", None)
    return gazetteer


def search_forms(gazetteer: dict[str, str]) -> list[tuple[str, str]]:
    """Пары (как написано в файле, каноническое имя), длинные формы раньше коротких."""
    forms = [(name, name) for name in gazetteer.values()]
    by_canon = {key_of(name): name for name in gazetteer.values()}
    for alias, canon in ALIASES.items():
        target = by_canon.get(key_of(canon))
        if target and key_of(alias) != key_of(target):
            forms.append((alias, target))
    forms.sort(key=lambda item: len(item[0].replace("ё", "е")), reverse=True)
    return forms


def find_mentions(text: str, forms: list[tuple[str, str]]) -> list[str]:
    """Сколько раз каждое название встречается отдельным словом. Длинные имена закрывают короткие."""
    hay = text.replace("ё", "е").replace("Ё", "Е")
    occupied = [False] * (len(hay) + 1)
    found = []
    for surface, canonical_name in forms:
        probe = surface.replace("ё", "е")
        pattern = re.compile(
            rf"(?<![0-9A-Za-zА-Яа-яЁё\-]){re.escape(probe)}(?![0-9A-Za-zА-Яа-яЁё\-])",
            re.IGNORECASE,
        )
        for match in pattern.finditer(hay):
            if any(occupied[match.start() : match.end()]):
                continue
            if not accept_mention(hay, match.start(), canonical_name):
                continue
            for index in range(match.start(), match.end()):
                occupied[index] = True
            found.append(canonical_name)
    return found


def accept_mention(text: str, start: int, name: str) -> bool:
    before = text[max(0, start - 24) : start].lower()
    if name == "Московский":
        return bool(re.search(r"п\.\s*$", before))
    if name == "Адыгея":
        return not before.rstrip().endswith("республика")
    return True


def snippet(text: str, name: str) -> str:
    hay = text.replace("ё", "е")
    probe = name.replace("ё", "е")
    match = re.search(re.escape(probe), hay, re.IGNORECASE)
    if not match:
        short = text if len(text) <= 140 else text[:137] + "…"
        return short
    start = max(0, match.start() - 40)
    end = min(len(text), match.end() + 40)
    piece = text[start:end]
    if start:
        piece = "…" + piece
    if end < len(text):
        piece = piece + "…"
    return piece


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


def write_book(rows, details, file_titles):
    wb = Workbook()
    ws = wb.active
    ws.title = "Города"
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:G1")
    ws["A1"] = "Населённые пункты по частоте упоминания"
    ws["A1"].font = FONT_TITLE
    ws["A1"].fill = FILL_TITLE
    ws["A1"].alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 28
    ws.merge_cells("A2:G2")
    ws["A2"] = (
        "Считается каждое отдельное упоминание во всех столбцах горда.xlsx, География.xlsx "
        "и Магнит_города для ПП_ Приложение 1.xlsx. «Москва 1» и «Белгород Восток» — это Москва и Белгород. "
        "«Стерлитомак» учтён как Стерлитамак. Названия улиц и районов («Московский проспект», «Челябинская область») не считаются городом."
    )
    ws["A2"].font = Font(name="Calibri", size=11, italic=True, color="595959")
    ws["A2"].alignment = LEFT
    ws.row_dimensions[2].height = 48

    headers = ["№", "Населённый пункт", "Упоминаний", *file_titles]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(4, col, header)
        cell.fill = FILL_HEADER
        cell.font = FONT_HEADER
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = THIN
    ws.row_dimensions[4].height = 32

    for index, row in enumerate(rows, 1):
        values = [index, row["name"], row["count"], *[row["by_file"].get(title, 0) for title in file_titles]]
        for col, value in enumerate(values, 1):
            cell = ws.cell(4 + index, col, value)
            cell.font = FONT
            cell.border = THIN
            cell.alignment = LEFT
            if index % 2 == 0:
                cell.fill = FILL_ZEBRA
        ws.cell(4 + index, 3).number_format = "#,##0"
        for col in range(4, 4 + len(file_titles)):
            ws.cell(4 + index, col).number_format = "#,##0"

    ws.auto_filter.ref = f"A4:{get_column_letter(3 + len(file_titles))}{4 + len(rows)}"
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:{get_column_letter(3 + len(file_titles))}{4 + len(rows)}"
    widths = [8, 32, 16, 42, 22, 48]
    for idx, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:4"
    ws.page_setup.horizontalCentered = True
    ws.oddFooter.right.text = "Города по частоте, стр. &P из &N"
    ws.page_setup.horizontalDpi = 300

    detail = wb.create_sheet("Упоминания")
    detail_headers = ["Населённый пункт", "Файл", "Лист", "Ячейка", "Фрагмент"]
    detail.append(detail_headers)
    for item in details:
        detail.append([item["name"], item["file"], item["sheet"], item["cell"], item["snippet"]])
    for col in range(1, 6):
        cell = detail.cell(1, col)
        cell.fill = FILL_HEADER
        cell.font = FONT_HEADER
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = THIN
    for row in detail.iter_rows(min_row=2, max_row=detail.max_row, max_col=5):
        for cell in row:
            cell.font = FONT
            cell.border = THIN
            cell.alignment = LEFT
        detail.row_dimensions[row[0].row].height = 30
    detail.auto_filter.ref = f"A1:E{detail.max_row}"
    detail.freeze_panes = "A2"
    for idx, width in enumerate([32, 48, 36, 12, 80], 1):
        detail.column_dimensions[get_column_letter(idx)].width = width
    detail.page_setup.orientation = "landscape"
    detail.page_setup.fitToPage = True
    detail.page_setup.fitToWidth = 1
    detail.page_setup.fitToHeight = 0
    detail.page_setup.paperSize = detail.PAPERSIZE_A4
    detail.sheet_properties.pageSetUpPr.fitToPage = True
    detail.print_title_rows = "1:1"
    detail.oddFooter.right.text = "Упоминания, стр. &P из &N"
    detail.page_setup.horizontalCentered = True

    wb.save(OUT)


def main():
    cells = load_cells()
    gazetteer = build_gazetteer(cells)
    forms = search_forms(gazetteer)
    counts = defaultdict(int)
    by_file = defaultdict(lambda: defaultdict(int))
    details = []
    for title, sheet, _column, coord, text in cells:
        mentions = find_mentions(text, forms)
        for name in mentions:
            counts[name] += 1
            by_file[name][title] += 1
            details.append(
                {
                    "name": name,
                    "file": title,
                    "sheet": sheet,
                    "cell": coord,
                    "snippet": snippet(text, name),
                }
            )

    file_titles = [title for title, _candidates in FILES]
    rows = []
    for name, count in counts.items():
        rows.append({"name": name, "count": count, "by_file": by_file[name]})
    rows.sort(key=lambda row: (-row["count"], row["name"]))
    details.sort(key=lambda item: (-counts[item["name"]], item["name"], item["file"], item["sheet"], item["cell"]))
    write_book(rows, details, file_titles)
    print(f"places {len(rows)} mentions {sum(counts.values())}")
    for row in rows:
        parts = " | ".join(f"{title.split('.')[0][:12]} {row['by_file'].get(title, 0)}" for title in file_titles)
        print(f"{row['count']:4}  {row['name']:28} {parts}")


if __name__ == "__main__":
    main()
