#!/usr/bin/env python3
"""Сравнение состава данных нарядов-допусков из DOCX."""

from __future__ import annotations

import re
from pathlib import Path

from docx import Document
from docx.document import Document as Doc
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

INPUT = Path("/home/ubuntu/.cursor/projects/workspace/uploads/_____-_______629d.docx")
OUTPUT = Path("/workspace/Сравнение_нарядов-допусков.xlsx")

# Унифицированные поля для сравнения состава данных
FIELDS = [
    ("Номер наряда-допуска", [r"наряд[-\s]?допуск\s*n", r"наряд[-\s]?допуск\s*№", r"наряд[-\s]?допуск\s*№"]),
    ("Организация / наименование организации", [r"организац", r"наименование организации"]),
    ("Подразделение", [r"подразделени"]),
    ("Место / объект / содержание работ", [r"поручается", r"наименование работ", r"место производ", r"содержание.*работ", r"характеристика.*работ", r"объект"]),
    ("Дата и время начала работ", [r"начать", r"начало работ", r"начать работ"]),
    ("Дата и время окончания работ", [r"закончить", r"окончание работ", r"окончить работ"]),
    ("Срок действия / действителен до", [r"действителен до", r"срок действия"]),
    ("Выдающий наряд (ФИО, должность)", [r"наряд[-\s]?допуск выдал", r"наряд выдал", r"лицо,?\s*выдавшее"]),
    ("Производитель работ", [r"производител"]),
    ("Ответственный руководитель работ", [r"ответственн\w+\s+руководител", r"руководител\w*\s+работ"]),
    ("Допускающий", [r"допускающ"]),
    ("Наблюдающий", [r"наблюдающ"]),
    ("Состав бригады / исполнителей", [r"членами бригады", r"состав.*бригад", r"состав исполнител", r"бригадой в составе", r"бригаде в составе"]),
    ("Группа по электробезопасности", [r"электробезопасност"]),
    ("Профессия / должность / квалификация членов бригады", [r"професси", r"квалификац"]),
    ("Мероприятия по подготовке рабочих мест", [r"подготовк\w+\s+рабоч", r"мероприяти.*подготов", r"до начала производства"]),
    ("Меры безопасности при производстве работ", [r"меры безопасности", r"в процессе производства", r"обеспечить следующие меры", r"мероприяти.*безопасност"]),
    ("Вредные и опасные производственные факторы", [r"вредн\w+\s+и\s+опасн", r"опасн\w+\s+производственн"]),
    ("Отдельные указания", [r"отдельн\w+\s+указан"]),
    ("Системы обеспечения безопасности работ на высоте", [r"систем\w+\s+обеспечения безопасности работ на высоте", r"работы на высоте"]),
    ("Целевой инструктаж (выдающий)", [r"целев\w+[\s\S]{0,80}инструктаж[\s\S]{0,80}выдающ", r"инструктаж[\s\S]{0,40}выдающ"]),
    ("Целевой инструктаж (допускающий)", [r"целев\w+[\s\S]{0,80}инструктаж[\s\S]{0,80}допускающ"]),
    ("Целевой инструктаж (руководитель / производитель)", [r"целев\w+[\s\S]{0,80}инструктаж[\s\S]{0,80}ответственн", r"целев\w+[\s\S]{0,80}инструктаж[\s\S]{0,80}производител"]),
    ("Инструктаж по охране труда (номера инструкций)", [r"инструктаж по охране труда", r"наименования или номера инструкций", r"объеме инструкций"]),
    ("Разрешение на подготовку рабочих мест / допуск", [r"разрешение на подготовку", r"разрешаю приступить", r"допуск к выполнению", r"допуск к производству"]),
    ("Ежедневный допуск / оформление начала работ", [r"ежедневн\w+\s+допуск", r"оформление начала производства", r"оформление ежедневного"]),
    ("Изменения в составе бригады", [r"изменен.*состав", r"введен в состав", r"выведен из состава"]),
    ("Продление наряда-допуска", [r"продл"]),
    ("Закрытие наряда / окончание работ", [r"наряд[-\s]?допуск закрыт", r"работа полностью закончена", r"работы завершены", r"работа выполнена"]),
    ("Осталось под напряжением / особые условия после подготовки", [r"под напряжением остались", r"рабочие места подготовлены"]),
    ("Согласование / разрешение эксплуатирующей организации", [r"эксплуатирующ", r"согласован"]),
    ("Принятие наряда / получил наряд", [r"наряд[-\s]?допуск принял", r"наряд[-\s]?допуск получил", r"наряд получил"]),
    ("Подписи об ознакомлении с условиями работ", [r"ознакомлен", r"подпись лица, прошедшего инструктаж", r"с условиями"]),
    ("Два экземпляра наряда", [r"двух экземплярах", r"в двух экземплярах"]),
]


def iter_block_items(parent):
    parent_elm = parent.element.body if isinstance(parent, Doc) else parent._tc
    for child in parent_elm.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, parent)
        elif child.tag == qn("w:tbl"):
            yield Table(child, parent)


ORDER_NAMES = {
    "903н": "Электроустановки",
    "883н": "Строительство",
    "882н": "Дорожные работы",
    "871н": "Автотранспорт",
    "875н": "Горэлектротранспорт",
    "867н": "Объекты связи",
    "866н": "Пищевая продукция",
    "858н": "Водные биоресурсы",
    "835н": "Инструмент",
    "833н": "Технологическое оборудование",
    "832н": "Полиграфия",
    "782н": "Работы на высоте",
    "776н": "Металлопокрытия",
    "814н": "Промтранспорт",
    "758н": "ЖКХ",
    "746н": "Сельское хозяйство",
    "721н": "Метрополитен",
    "644н": "Лес / деревообработка",
    "924н": "Теплоснабжение",
    "922н": "Водолазные работы",
    "915н": "Нефтепродукты",
    "849н": "Окрасочные работы",
    "834н": "Химвещества / чистка",
    "781н": "Производство цемента",
}


def extract_order_meta(text: str) -> tuple[str, str, str]:
    """Return (topic, order_no, order_date)."""
    compact = re.sub(r"\s+", " ", text)
    m = re.search(
        r"Правилам по охране труда (.+?),?\s*утвержденн",
        compact,
        flags=re.I,
    )
    topic = m.group(1).strip(" ,") if m else ""
    topic = re.sub(r"\s+", " ", topic)
    if len(topic) > 110:
        topic = topic[:107] + "..."

    m_order = re.search(r"(?:N|№)\s*([0-9]+н)", compact, flags=re.I)
    order_no = m_order.group(1) if m_order else ""

    m_date = re.search(
        r"от\s+(\d{1,2}\s+[а-яё]+\s+\d{4}\s*г\.?)",
        compact,
        flags=re.I,
    )
    order_date = m_date.group(1) if m_date else ""
    return topic, order_no, order_date


def short_form_name(title: str, topic: str, order_no: str, appendix: str = "") -> str:
    title_l = title.lower()
    base = ORDER_NAMES.get(order_no)
    if order_no == "922н":
        if "судовых" in title_l:
            return "Водолазные судовые (922н пр.4)"
        if "наряд-задание" in title_l:
            return "Водолазные наряд-задание (922н пр.5)"
        return "Водолазные работы (922н пр.3)"
    if base:
        return f"{base} ({order_no})"
    return (topic[:40] or title[:40]) + (f" ({order_no})" if order_no else "")


def build_forms(doc: Document) -> list[dict]:
    blocks: list[tuple[str, str]] = []
    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            t = block.text.strip()
            if t:
                blocks.append(("p", t))
        else:
            cells = []
            for row in block.rows:
                for cell in row.cells:
                    ct = cell.text.strip()
                    if ct:
                        cells.append(ct)
            if cells:
                blocks.append(("t", "\n".join(cells)))

    starts = [i for i, (k, t) in enumerate(blocks) if k == "p" and re.match(r"^Приложение", t)]
    forms = []
    for si, start in enumerate(starts):
        end = starts[si + 1] if si + 1 < len(starts) else len(blocks)
        chunk = blocks[start:end]
        full_text = "\n".join(t for _, t in chunk)
        low = full_text.lower()

        appendix = chunk[0][1]
        topic, order_no, order_date = extract_order_meta(appendix)

        title = ""
        for kind, text in chunk[:12]:
            if re.search(r"НАРЯД|Наряд-допуск|Наряд-задание", text):
                title = re.sub(r"\s+", " ", text)
                break
        if not title:
            title = "Наряд-допуск"

        name = short_form_name(title, topic, order_no, appendix)
        forms.append(
            {
                "idx": si + 1,
                "name": name,
                "title": title[:120],
                "topic": topic,
                "order_no": order_no,
                "order_date": order_date,
                "appendix": re.sub(r"\s+", " ", appendix)[:220],
                "text": full_text,
                "low": low,
            }
        )

    # Disambiguate identical short names
    counts: dict[str, int] = {}
    for f in forms:
        counts[f["name"]] = counts.get(f["name"], 0) + 1
    seen: dict[str, int] = {}
    for f in forms:
        if counts[f["name"]] > 1:
            seen[f["name"]] = seen.get(f["name"], 0) + 1
            f["name"] = f'{f["name"]} #{seen[f["name"]]}'
    return forms


def field_present(low: str, patterns: list[str]) -> bool:
    # DOTALL: в бланках формулировки часто разорваны переносами строк
    return any(re.search(p, low, flags=re.I | re.DOTALL) for p in patterns)


def style_header(cell, fill_color="1F4E79"):
    cell.fill = PatternFill("solid", fgColor=fill_color)
    cell.font = Font(bold=True, color="FFFFFF", size=10)
    cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")


def autosize(ws, max_width=42):
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        width = 10
        for cell in col:
            if cell.value:
                width = max(width, min(len(str(cell.value)) + 2, max_width))
        ws.column_dimensions[letter].width = width


def main():
    doc = Document(str(INPUT))
    forms = build_forms(doc)

    # Presence matrix
    matrix = []
    for field_name, patterns in FIELDS:
        row = {"field": field_name}
        for f in forms:
            row[f["name"]] = "Да" if field_present(f["low"], patterns) else "—"
        matrix.append(row)

    # Count presence per form
    for f in forms:
        f["fields_count"] = sum(1 for r in matrix if r[f["name"]] == "Да")
        f["fields_total"] = len(FIELDS)

    wb = Workbook()

    # --- Sheet 1: comparison matrix ---
    ws = wb.active
    ws.title = "Сравнение состава данных"
    thin = Border(
        left=Side(style="thin", color="B0B0B0"),
        right=Side(style="thin", color="B0B0B0"),
        top=Side(style="thin", color="B0B0B0"),
        bottom=Side(style="thin", color="B0B0B0"),
    )

    headers = ["Поле / блок данных"] + [f["name"] for f in forms] + ["В скольких формах"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(1, col, h)
        style_header(cell)
    ws.row_dimensions[1].height = 55
    ws.freeze_panes = "B2"

    yes_fill = PatternFill("solid", fgColor="C6EFCE")
    no_fill = PatternFill("solid", fgColor="F2F2F2")
    yes_font = Font(color="006100", bold=True)
    no_font = Font(color="808080")

    for ri, row in enumerate(matrix, 2):
        cell = ws.cell(ri, 1, row["field"])
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        cell.border = thin
        yes_count = 0
        for ci, f in enumerate(forms, 2):
            val = row[f["name"]]
            c = ws.cell(ri, ci, val)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = thin
            if val == "Да":
                c.fill = yes_fill
                c.font = yes_font
                yes_count += 1
            else:
                c.fill = no_fill
                c.font = no_font
        c = ws.cell(ri, len(forms) + 2, yes_count)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = thin
        c.font = Font(bold=True)

    # totals row
    total_row = len(matrix) + 2
    ws.cell(total_row, 1, "Итого полей (из списка)").font = Font(bold=True)
    for ci, f in enumerate(forms, 2):
        c = ws.cell(total_row, ci, f["fields_count"])
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center")
        c.fill = PatternFill("solid", fgColor="DDEBF7")

    ws.column_dimensions["A"].width = 48
    for i in range(2, len(forms) + 3):
        ws.column_dimensions[get_column_letter(i)].width = 14

    # --- Sheet 2: registry ---
    ws2 = wb.create_sheet("Реестр форм", 0)
    reg_headers = [
        "№",
        "Краткое название",
        "Полное название формы",
        "Область ПОТ",
        "Приказ",
        "Дата приказа",
        "Источник (приложение)",
        "Полей из списка",
        "% покрытия списка",
    ]
    for col, h in enumerate(reg_headers, 1):
        style_header(ws2.cell(1, col, h))
    for ri, f in enumerate(forms, 2):
        pct = round(100 * f["fields_count"] / f["fields_total"], 1)
        values = [
            f["idx"],
            f["name"],
            f["title"],
            f["topic"],
            f["order_no"],
            f["order_date"],
            f["appendix"],
            f["fields_count"],
            pct,
        ]
        for ci, v in enumerate(values, 1):
            cell = ws2.cell(ri, ci, v)
            cell.alignment = Alignment(wrap_text=True, vertical="center")
            cell.border = thin
    ws2.freeze_panes = "A2"
    ws2.column_dimensions["A"].width = 5
    ws2.column_dimensions["B"].width = 34
    ws2.column_dimensions["C"].width = 40
    ws2.column_dimensions["D"].width = 40
    ws2.column_dimensions["E"].width = 10
    ws2.column_dimensions["F"].width = 22
    ws2.column_dimensions["G"].width = 55
    ws2.column_dimensions["H"].width = 14
    ws2.column_dimensions["I"].width = 14

    # --- Sheet 3: summary by field frequency ---
    ws3 = wb.create_sheet("Частота полей")
    style_header(ws3.cell(1, 1, "Поле / блок данных"))
    style_header(ws3.cell(1, 2, "Есть в формах"))
    style_header(ws3.cell(1, 3, "% форм"))
    style_header(ws3.cell(1, 4, "Категория распространённости"))
    freq_rows = []
    for row in matrix:
        cnt = sum(1 for f in forms if row[f["name"]] == "Да")
        pct = round(100 * cnt / len(forms), 1)
        if pct >= 80:
            cat = "Общее почти для всех"
        elif pct >= 50:
            cat = "Частое"
        elif pct >= 20:
            cat = "Частичное"
        else:
            cat = "Редкое / специфичное"
        freq_rows.append((row["field"], cnt, pct, cat))
    freq_rows.sort(key=lambda x: (-x[2], x[0]))
    for ri, (field, cnt, pct, cat) in enumerate(freq_rows, 2):
        ws3.cell(ri, 1, field).alignment = Alignment(wrap_text=True)
        ws3.cell(ri, 2, cnt).alignment = Alignment(horizontal="center")
        ws3.cell(ri, 3, pct).alignment = Alignment(horizontal="center")
        c = ws3.cell(ri, 4, cat)
        if cat.startswith("Общее"):
            c.fill = PatternFill("solid", fgColor="C6EFCE")
        elif cat == "Частое":
            c.fill = PatternFill("solid", fgColor="FFF2CC")
        elif cat == "Частичное":
            c.fill = PatternFill("solid", fgColor="FCE4D6")
        else:
            c.fill = PatternFill("solid", fgColor="F2F2F2")
    ws3.column_dimensions["A"].width = 55
    ws3.column_dimensions["B"].width = 14
    ws3.column_dimensions["C"].width = 10
    ws3.column_dimensions["D"].width = 28

    # --- Sheet 4: notes ---
    ws4 = wb.create_sheet("Пояснения")
    notes = [
        "Источник: файл «Наряд-допуск.docx» — сборник рекомендуемых образцов нарядов-допусков из ПОТ Минтруда РФ (2020).",
        f"Всего форм в файле: {len(forms)}.",
        "Сравнение сделано по унифицированному списку блоков данных (не посимвольное сравнение бланков).",
        "«Да» — блок явно присутствует в тексте/таблицах формы; «—» — не найден по типовым формулировкам.",
        "Формы ЖКХ #1 и #2 — два почти идентичных приложения из приказа 758н (как в исходном файле).",
        "Приложение 4 к 922н (судовые водолазные) в файле почти пустое (только заголовок) — покрытие низкое.",
        "У электроустановок (903н) и работ на высоте (782н) — наиболее «богатый» состав данных (целевые инструктажи, спец. блоки).",
        "Лист «Реестр форм» — перечень с реквизитами приказов.",
        "Лист «Сравнение состава данных» — матрица присутствия полей.",
        "Лист «Частота полей» — какие блоки общие, а какие специфичны.",
    ]
    ws4["A1"] = "Пояснения к сравнению"
    ws4["A1"].font = Font(bold=True, size=14)
    for i, note in enumerate(notes, 3):
        ws4.cell(i, 1, f"• {note}")
        ws4.cell(i, 1).alignment = Alignment(wrap_text=True)
    ws4.column_dimensions["A"].width = 120

    wb.save(OUTPUT)
    print(f"Saved: {OUTPUT}")
    print(f"Forms: {len(forms)}")
    for f in forms:
        print(f"  {f['idx']:2}. {f['name']}: {f['fields_count']}/{f['fields_total']}")


if __name__ == "__main__":
    main()
