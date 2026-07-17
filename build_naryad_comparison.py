#!/usr/bin/env python3
"""Сравнение: какую составляющую нужно указывать в каком наряде-допуске."""

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

# Составляющие наряда: (название, группа, паттерны поиска)
COMPONENTS = [
    ("Группа", "Составляющая", "Паттерны"),
]

# Реальные составляющие для заполнения
FIELDS: list[tuple[str, str, list[str]]] = [
    ("Шапка", "Номер наряда-допуска", [r"наряд[-\s]?допуск\s*[n№n]", r"наряд[-\s]?допуск\s*n\s*[_\.]"]),
    ("Шапка", "Организация (наименование)", [r"организац", r"наименование организации"]),
    ("Шапка", "Подразделение", [r"подразделени"]),
    ("Шапка", "Место / содержание / характеристика работ", [r"поручается", r"наименование работ", r"место производ", r"содержание работ", r"содержание[,\s]*характеристика", r"объем работ", r"необходимые для производства работ"]),
    ("Сроки", "Дата и время начала работ", [r"начать", r"начало работ", r"начать работ"]),
    ("Сроки", "Дата и время окончания работ", [r"закончить", r"окончание работ", r"окончить работ"]),
    ("Сроки", "Срок действия / «действителен до»", [r"действителен до", r"срок действия"]),
    ("Роли", "Кто выдал наряд (ФИО, должность, подпись)", [r"наряд[-\s]?допуск выдал", r"наряд выдал", r"лицо,?\s*выдавшее"]),
    ("Роли", "Производитель работ", [r"производител\w*\s+работ", r"производителю работ", r"производителю\s"]),
    ("Роли", "Ответственный руководитель работ", [r"ответственн\w+\s+руководител", r"руководител\w*\s+работ", r"руководителю работ"]),
    ("Роли", "Допускающий", [r"допускающ"]),
    ("Роли", "Наблюдающий", [r"наблюдающ"]),
    ("Бригада", "Состав бригады / исполнителей", [r"членами бригады", r"состав.*бригад", r"состав исполнител", r"бригадой в составе", r"бригаде в составе"]),
    ("Бригада", "ФИО членов бригады", [r"фамилия,\s*инициалы", r"фамилия и инициалы", r"фамилия, имя, отчество"]),
    ("Бригада", "Профессия / должность членов бригады", [r"професси"]),
    ("Бригада", "Группа по электробезопасности", [r"электробезопасност"]),
    ("Бригада", "Изменения в составе бригады", [r"изменен[\s\S]{0,40}состав", r"введен в состав", r"выведен из состава"]),
    ("Меры", "Мероприятия до начала работ / подготовка рабочих мест", [r"подготовк[\s\S]{0,30}рабоч", r"до начала[\s\S]{0,40}мероприят", r"до начала производства", r"мероприяти[\s\S]{0,40}подготов"]),
    ("Меры", "Мероприятия в процессе работ / меры безопасности", [r"в процессе производства", r"меры безопасности", r"обеспечить следующие меры", r"мероприяти[\s\S]{0,40}безопасност"]),
    ("Меры", "Вредные и опасные производственные факторы (ВОПФ)", [r"вредн[\s\S]{0,20}опасн", r"опасн[\s\S]{0,20}производственн\w+\s+фактор"]),
    ("Меры", "Отдельные / особые указания", [r"отдельн\w+\s+указан", r"особые условия"]),
    ("Меры", "Системы обеспечения безопасности работ на высоте", [r"систем\w+\s+обеспечения безопасности работ на высоте"]),
    ("Инструктаж", "Целевой инструктаж выдающего наряд", [r"целев[\s\S]{0,80}инструктаж[\s\S]{0,80}выдающ"]),
    ("Инструктаж", "Целевой инструктаж допускающего", [r"целев[\s\S]{0,80}инструктаж[\s\S]{0,80}допускающ"]),
    ("Инструктаж", "Целевой инструктаж руководителя / производителя", [r"целев[\s\S]{0,80}инструктаж[\s\S]{0,100}(ответственн|производител|наблюдающ)"]),
    ("Инструктаж", "Инструктаж по ОТ (номера инструкций) + подписи бригады", [r"инструктаж по охране труда", r"номера инструкций", r"наименования или номера инструкций", r"подпись лица, получившего инструктаж"]),
    ("Инструктаж", "Подпись об ознакомлении с условиями работ", [r"ознакомлен", r"прошедшего инструктаж", r"с условиями работ"]),
    ("Допуск", "Разрешение на подготовку РМ / допуск к работам", [r"разрешение на подготовку", r"разрешаю приступить", r"допуск к выполнению", r"допуск к производству", r"допускающий к работе"]),
    ("Допуск", "Ежедневный допуск / оформление начала и окончания смены", [r"ежедневн[\s\S]{0,20}допуск", r"оформление начала производства", r"оформление ежедневного"]),
    ("Допуск", "Отметка: рабочие места подготовлены / осталось под напряжением", [r"рабочие места подготовлены", r"под напряжением остались"]),
    ("Допуск", "Согласование / разрешение эксплуатирующей организации", [r"эксплуатирующ", r"акт-допуск", r"письменное\s+разрешение"]),
    ("Закрытие", "Принятие наряда / получил наряд", [r"наряд[-\s]?допуск принял", r"наряд[-\s]?допуск получил", r"наряд получил"]),
    ("Закрытие", "Продление наряда-допуска", [r"продл"]),
    ("Закрытие", "Закрытие наряда / окончание работ (подписи)", [r"наряд[-\s]?допуск закрыт", r"работа полностью закончена", r"работы завершены", r"работа выполнена"]),
]

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
    "922н": "Водолазные",
    "915н": "Нефтепродукты",
    "849н": "Окрасочные работы",
    "834н": "Химвещества / чистка",
    "781н": "Производство цемента",
}

YES = "Нужно"
NO = "Нет"
YES_FILL = PatternFill("solid", fgColor="C6EFCE")
NO_FILL = PatternFill("solid", fgColor="F2F2F2")
YES_FONT = Font(color="006100", bold=True, size=10)
NO_FONT = Font(color="A0A0A0", size=10)
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=9)
GROUP_FILLS = {
    "Шапка": PatternFill("solid", fgColor="D6EAF8"),
    "Сроки": PatternFill("solid", fgColor="D5F5E3"),
    "Роли": PatternFill("solid", fgColor="FCF3CF"),
    "Бригада": PatternFill("solid", fgColor="FADBD8"),
    "Меры": PatternFill("solid", fgColor="E8DAEF"),
    "Инструктаж": PatternFill("solid", fgColor="FDEBD0"),
    "Допуск": PatternFill("solid", fgColor="D1F2EB"),
    "Закрытие": PatternFill("solid", fgColor="D5D8DC"),
}
THIN = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)


def iter_block_items(parent):
    parent_elm = parent.element.body if isinstance(parent, Doc) else parent._tc
    for child in parent_elm.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, parent)
        elif child.tag == qn("w:tbl"):
            yield Table(child, parent)


def present(low: str, patterns: list[str]) -> bool:
    return any(re.search(p, low, flags=re.I | re.DOTALL) for p in patterns)


def extract_meta(text: str) -> tuple[str, str, str]:
    compact = re.sub(r"\s+", " ", text)
    m = re.search(r"Правилам по охране труда (.+?),?\s*утвержденн", compact, flags=re.I)
    topic = re.sub(r"\s+", " ", m.group(1).strip(" ,")) if m else ""
    m_order = re.search(r"(?:N|№)\s*([0-9]+н)", compact, flags=re.I)
    order_no = m_order.group(1) if m_order else ""
    m_date = re.search(r"от\s+(\d{1,2}\s+[а-яё]+\s+\d{4}\s*г\.?)", compact, flags=re.I)
    order_date = m_date.group(1) if m_date else ""
    return topic, order_no, order_date


def short_name(title: str, order_no: str) -> str:
    title_l = title.lower()
    if order_no == "922н":
        if "судовых" in title_l:
            return "Водолазные судовые (922н)"
        if "наряд-задание" in title_l:
            return "Водолазные наряд-задание (922н)"
        return "Водолазные работы (922н)"
    base = ORDER_NAMES.get(order_no, title[:30])
    return f"{base} ({order_no})" if order_no else base


def build_forms(doc: Document) -> list[dict]:
    blocks: list[str] = []
    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            t = block.text.strip()
            if t:
                blocks.append(t)
        else:
            cells = []
            for row in block.rows:
                for cell in row.cells:
                    ct = cell.text.strip()
                    if ct:
                        cells.append(ct)
            if cells:
                blocks.append("\n".join(cells))

    starts = [i for i, t in enumerate(blocks) if t.startswith("Приложение")]
    forms = []
    for si, start in enumerate(starts):
        end = starts[si + 1] if si + 1 < len(starts) else len(blocks)
        chunk = blocks[start:end]
        full = "\n".join(chunk)
        appendix = chunk[0]
        topic, order_no, order_date = extract_meta(appendix)
        title = next(
            (re.sub(r"\s+", " ", t) for t in chunk[:12] if re.search(r"НАРЯД|Наряд-допуск|Наряд-задание", t)),
            "Наряд-допуск",
        )
        forms.append(
            {
                "idx": si + 1,
                "name": short_name(title, order_no),
                "title": title[:140],
                "topic": topic,
                "order_no": order_no,
                "order_date": order_date,
                "low": full.lower(),
            }
        )

    counts: dict[str, int] = {}
    for f in forms:
        counts[f["name"]] = counts.get(f["name"], 0) + 1
    seen: dict[str, int] = {}
    for f in forms:
        if counts[f["name"]] > 1:
            seen[f["name"]] = seen.get(f["name"], 0) + 1
            f["name"] = f'{f["name"]} #{seen[f["name"]]}'
    return forms


def style_header(cell):
    cell.fill = HEADER_FILL
    cell.font = HEADER_FONT
    cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
    cell.border = THIN


def main():
    forms = build_forms(Document(str(INPUT)))

    # matrix: field -> form -> YES/NO
    matrix: list[dict] = []
    for group, field, patterns in FIELDS:
        row = {"group": group, "field": field}
        for f in forms:
            row[f["name"]] = YES if present(f["low"], patterns) else NO
        matrix.append(row)

    for f in forms:
        f["need_count"] = sum(1 for r in matrix if r[f["name"]] == YES)

    wb = Workbook()

    # ========== 1. Главный лист: составляющая × наряд ==========
    ws = wb.active
    ws.title = "Что указывать (матрица)"
    headers = ["Группа", "Составляющая наряда"] + [f["name"] for f in forms] + ["В скольких нарядах нужно"]
    for col, h in enumerate(headers, 1):
        style_header(ws.cell(1, col, h))
    ws.row_dimensions[1].height = 70
    ws.freeze_panes = "C2"

    for ri, row in enumerate(matrix, 2):
        g = ws.cell(ri, 1, row["group"])
        g.fill = GROUP_FILLS.get(row["group"], PatternFill())
        g.alignment = Alignment(vertical="center", horizontal="center")
        g.border = THIN
        g.font = Font(bold=True, size=9)

        fcell = ws.cell(ri, 2, row["field"])
        fcell.alignment = Alignment(wrap_text=True, vertical="center")
        fcell.border = THIN
        fcell.fill = GROUP_FILLS.get(row["group"], PatternFill())

        need_n = 0
        for ci, form in enumerate(forms, 3):
            val = row[form["name"]]
            c = ws.cell(ri, ci, val)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = THIN
            if val == YES:
                c.fill = YES_FILL
                c.font = YES_FONT
                need_n += 1
            else:
                c.fill = NO_FILL
                c.font = NO_FONT
        total = ws.cell(ri, len(forms) + 3, need_n)
        total.alignment = Alignment(horizontal="center", vertical="center")
        total.font = Font(bold=True)
        total.border = THIN

    # totals
    tr = len(matrix) + 2
    ws.cell(tr, 1, "").border = THIN
    ws.cell(tr, 2, "Сколько составляющих нужно заполнять").font = Font(bold=True)
    ws.cell(tr, 2).border = THIN
    for ci, form in enumerate(forms, 3):
        c = ws.cell(tr, ci, form["need_count"])
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center")
        c.fill = PatternFill("solid", fgColor="DDEBF7")
        c.border = THIN

    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 48
    for i in range(3, len(forms) + 4):
        ws.column_dimensions[get_column_letter(i)].width = 13

    # ========== 2. По составляющей: где нужно / где нет ==========
    ws2 = wb.create_sheet("По составляющей")
    for col, h in enumerate(
        ["Группа", "Составляющая", "Нужно указывать (наряды)", "Не нужно (наряды)", "Нужно, шт.", "Нет, шт."],
        1,
    ):
        style_header(ws2.cell(1, col, h))
    ws2.row_dimensions[1].height = 30
    ws2.freeze_panes = "A2"

    for ri, row in enumerate(matrix, 2):
        need = [f["name"] for f in forms if row[f["name"]] == YES]
        skip = [f["name"] for f in forms if row[f["name"]] == NO]
        values = [
            row["group"],
            row["field"],
            "; ".join(need) if need else "—",
            "; ".join(skip) if skip else "—",
            len(need),
            len(skip),
        ]
        for ci, v in enumerate(values, 1):
            c = ws2.cell(ri, ci, v)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            c.border = THIN
            if ci == 1:
                c.fill = GROUP_FILLS.get(row["group"], PatternFill())
                c.font = Font(bold=True, size=9)
            if ci == 3 and need:
                c.fill = YES_FILL
            if ci == 4 and skip:
                c.fill = NO_FILL
        ws2.row_dimensions[ri].height = max(30, 15 * (1 + max(len(need), len(skip)) // 3))

    ws2.column_dimensions["A"].width = 12
    ws2.column_dimensions["B"].width = 45
    ws2.column_dimensions["C"].width = 55
    ws2.column_dimensions["D"].width = 55
    ws2.column_dimensions["E"].width = 12
    ws2.column_dimensions["F"].width = 10

    # ========== 3. Чек-лист по каждому наряду ==========
    ws3 = wb.create_sheet("Чек-лист по наряду")
    style_header(ws3.cell(1, 1, "Наряд-допуск"))
    style_header(ws3.cell(1, 2, "Приказ"))
    style_header(ws3.cell(1, 3, "Что НУЖНО указывать"))
    style_header(ws3.cell(1, 4, "Чего НЕТ в бланке"))
    style_header(ws3.cell(1, 5, "Нужно, шт."))
    ws3.row_dimensions[1].height = 30
    ws3.freeze_panes = "A2"

    for ri, form in enumerate(forms, 2):
        need = [r["field"] for r in matrix if r[form["name"]] == YES]
        skip = [r["field"] for r in matrix if r[form["name"]] == NO]
        ws3.cell(ri, 1, form["name"]).alignment = Alignment(wrap_text=True, vertical="top")
        ws3.cell(ri, 2, form["order_no"]).alignment = Alignment(horizontal="center", vertical="top")
        c3 = ws3.cell(ri, 3, "\n".join(f"• {x}" for x in need))
        c3.alignment = Alignment(wrap_text=True, vertical="top")
        c3.fill = YES_FILL
        c4 = ws3.cell(ri, 4, "\n".join(f"• {x}" for x in skip))
        c4.alignment = Alignment(wrap_text=True, vertical="top")
        c4.fill = NO_FILL
        ws3.cell(ri, 5, len(need)).alignment = Alignment(horizontal="center", vertical="top")
        for ci in range(1, 6):
            ws3.cell(ri, ci).border = THIN
        ws3.row_dimensions[ri].height = max(60, 12 * max(len(need), len(skip), 4))

    ws3.column_dimensions["A"].width = 32
    ws3.column_dimensions["B"].width = 10
    ws3.column_dimensions["C"].width = 55
    ws3.column_dimensions["D"].width = 55
    ws3.column_dimensions["E"].width = 12

    # ========== 4. Реестр ==========
    ws4 = wb.create_sheet("Реестр нарядов", 0)
    for col, h in enumerate(
        ["№", "Наряд-допуск", "Область ПОТ", "Приказ", "Дата", "Составляющих нужно указать"],
        1,
    ):
        style_header(ws4.cell(1, col, h))
    for ri, f in enumerate(forms, 2):
        for ci, v in enumerate(
            [f["idx"], f["name"], f["topic"], f["order_no"], f["order_date"], f["need_count"]],
            1,
        ):
            c = ws4.cell(ri, ci, v)
            c.alignment = Alignment(wrap_text=True, vertical="center")
            c.border = THIN
    ws4.column_dimensions["A"].width = 5
    ws4.column_dimensions["B"].width = 36
    ws4.column_dimensions["C"].width = 50
    ws4.column_dimensions["D"].width = 10
    ws4.column_dimensions["E"].width = 22
    ws4.column_dimensions["F"].width = 18

    # ========== 5. Пояснения ==========
    ws5 = wb.create_sheet("Как читать")
    ws5["A1"] = "Как читать таблицу"
    ws5["A1"].font = Font(bold=True, size=14)
    notes = [
        "Зелёное «Нужно» — в бланке этого наряда-допуска есть поле/блок, куда это указывают.",
        "Серое «Нет» — в бланке такого поля нет (указывать не требуется по форме).",
        "Лист «Что указывать (матрица)» — быстрый обзор: строка = составляющая, столбец = наряд.",
        "Лист «По составляющей» — для каждой составляющей списки нарядов «нужно» и «не нужно».",
        "Лист «Чек-лист по наряду» — для каждого наряда: что заполнять и чего в бланке нет.",
        "Сравнение по рекомендуемым образцам из ПОТ Минтруда (файл Наряд-допуск.docx), 27 форм.",
        "ЖКХ #1 и #2 — два почти одинаковых бланка из приказа 758н (как в исходном файле).",
        "Водолазные судовые (922н) в файле почти пустые (только заголовок) — у них мало «Нужно».",
    ]
    for i, n in enumerate(notes, 3):
        ws5.cell(i, 1, f"• {n}").alignment = Alignment(wrap_text=True)
    ws5.column_dimensions["A"].width = 120

    wb.save(OUTPUT)
    print(f"Saved {OUTPUT}")
    print(f"Forms: {len(forms)}, components: {len(FIELDS)}")
    rare = [r for r in matrix if sum(1 for f in forms if r[f["name"]] == YES) <= 3]
    common = [r for r in matrix if sum(1 for f in forms if r[f["name"]] == YES) >= 24]
    print("Common (>=24):", [r["field"] for r in common])
    print("Rare (<=3):", [r["field"] for r in rare])


if __name__ == "__main__":
    main()
