#!/usr/bin/env python3
"""Сравнение нарядов-допусков: составляющие слово в слово из бланков."""

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

# Сравнительные позиции: для каждой — список regex, которыми ищем ДОСЛОВНЫЙ фрагмент в бланке.
# В ячейку попадает найденный текст из документа (не моя формулировка).
COMPARE_SLOTS: list[tuple[str, list[str]]] = [
    ("Номер наряда-допуска", [r"НАРЯД[-\s]?ДОПУСК\s*[N№n]?\s*[_.…—\-]*", r"Наряд-допуск\s*[N№n]?\s*[_.…—\-]*"]),
    ("Организация", [r"Организация\s*[:_]*.{0,40}", r"\(наименование организации\)"]),
    ("Подразделение", [r"Подразделение\s*[:_]*.{0,40}"]),
    ("Ответственный руководитель работ", [
        r"Ответственному руководителю\s*работ",
        r"Ответственному\s*руководителю работ",
        r"1\.\s*Руководителю работ",
        r"Руководитель работ",
        r"Ответственному\s*руководителю работ:",
    ]),
    ("Производитель / ответственный исполнитель работ", [
        r"Производителю\s*работ",
        r"1\.1\.\s*Производителю работ",
        r"Ответственному\s*исполнителю \(производителю\) работ",
        r"Производитель работ",
    ]),
    ("Допускающий", [r"допускающему", r"Допускающий", r"Допускающий к работе"]),
    ("Наблюдающий", [r"наблюдающему", r"наблюдающему", r"наблюдающий"]),
    ("Члены бригады / состав исполнителей", [
        r"с членами бригады",
        r"Состав исполнителей работ",
        r"с бригадой в составе",
        r"6\.\s*Состав исполнителей работ",
    ]),
    ("Группа по электробезопасности (дословно)", [
        r"фамилия,\s*инициалы,\s*группа\s*по\s*электробезопасности",
        r"Профессия \(должность\), квалификация, группа по электробезопасности",
        r"\(фамилия, инициалы, группа по электробезопасности\)",
    ]),
    ("Поручается / наименование / содержание работ", [
        r"поручается",
        r"2\.\s*Наименование работ",
        r"\(наименование работ, место, условия их выполнения\)",
        r"Наименование работ",
        r"Содержание работ",
        r"поручается произвести следующие работы",
        r"\(содержание, характеристика, место производства и объем работ\)",
    ]),
    ("Место выполнения работ", [r"Место выполнения работ", r"место производства"]),
    ("Начало работ (дата/время)", [
        r"Работу начать:\s*дата",
        r"Начало работ",
        r"1\.3\.\s*Начать работы",
        r"Начать работы",
    ]),
    ("Окончание работ (дата/время)", [
        r"Работу закончить:\s*дата",
        r"Окончание работ",
        r"1\.4\.\s*Окончить работы",
        r"Окончить работы",
    ]),
    ("Срок действия / выдан / действителен до", [
        r"Выдан\s*[\"«]?",
        r"Действителен до",
        r"Срок действия наряда",
    ]),
    ("Вредные и опасные производственные факторы", [
        r"Вредные\s+и\s+опасные\s+производственные\s+факторы",
        r"Опасные и вредные производственные",
        r"вредные и опасные производственные факторы",
    ]),
    ("Мероприятия по подготовке рабочих мест / до начала работ", [
        r"Мероприятия по подготовке рабочих мест к выполнению работ",
        r"До\s+начала\s+производства\s+работ\s+необходимо\s+выполнить\s+следующие\s+мероприятия",
        r"До начала работ следует выполнить следующие мероприятия",
        r"При\s+подготовке\s+и\s+производстве\s+работ\s+обеспечить\s+следующие\s+меры\s+безопасности",
        r"1\.2\.\s*При\s+подготовке",
    ]),
    ("Мероприятия в процессе производства работ", [
        r"В\s+процессе\s+производства\s+работ\s+необходимо\s+выполнить\s+следующие\s+мероприятия",
        r"Наименование мероприятия по безопасности работ на высоте",
    ]),
    ("Отдельные указания / особые условия", [
        r"Отдельные указания",
        r"Особые условия проведения работ",
    ]),
    ("Системы обеспечения безопасности работ на высоте", [
        r"Системы обеспечения безопасности работ на высоте",
    ]),
    ("Наряд-допуск выдал / Наряд выдал", [
        r"Наряд-допуск выдал",
        r"Наряд выдал",
        r"1\.5\.\s*Наряд выдал",
        r"7\.\s*Наряд-допуск выдал",
    ]),
    ("Наряд-допуск принял / получил", [
        r"Наряд-допуск принял",
        r"наряд-допуск получил",
        r"С условиями производства работ ознакомлен, наряд-допуск получил",
        r"С условиями работы ознакомлен, наряд-допуск получил",
        r"С условиями работ ознакомлен и наряд-допуск получил",
    ]),
    ("Наряд-допуск продлил / продлен", [
        r"Наряд-допуск продлил",
        r"Наряд продлил",
        r"Наряд-допуск продлен",
        r"Наряд допуск продлен",
    ]),
    ("Регистрация целевого инструктажа, проводимого выдающим наряд-допуск", [
        r"Регистрация целевого инструктажа,\s*проводимого выдающим наряд-допуск",
        r"Регистрация целевого инструктажа, проводимого выдающим наряд-допуск",
    ]),
    ("Регистрация целевого инструктажа, проводимого допускающим при первичном допуске", [
        r"Регистрация целевого инструктажа,[\s\n]*проводимого допускающим при первичном допуске",
    ]),
    ("Регистрация целевого инструктажа, проводимого ответственным руководителем работ", [
        r"Регистрация целевого инструктажа,[\s\n]*проводимого ответственным руководителем работ",
        r"Регистрация целевого инструктажа при первичном допуске",
    ]),
    ("Инструктаж по охране труда (номера инструкций) / таблица подписей бригады", [
        r"Инструктаж по охране труда в объеме инструкций",
        r"Подпись лица, получившего инструктаж",
        r"\(указать наименования или номера инструкций",
    ]),
    ("Разрешение на подготовку рабочих мест и на допуск к выполнению работ", [
        r"Разрешение на подготовку рабочих мест и на допуск к выполнению работ",
        r"Разрешаю приступить к выполнению работ",
        r"Разрешаю приступить\s*к выполнению работ",
        r"2\.2\.\s*Мероприятия,\s*обеспечивающие\s*безопасность\s*работ",
    ]),
    ("Рабочие места подготовлены. Под напряжением остались", [
        r"Рабочие места подготовлены\.\s*Под напряжением остались",
        r"Рабочие места подготовлены\.",
    ]),
    ("Ежедневный допуск к работе и время ее окончания", [
        r"Ежедневный допуск к работе и время ее окончания",
        r"Оформление ежедневного допуска к производству работ",
        r"Оформление начала производства работ",
    ]),
    ("Изменения в составе бригады / исполнителей", [
        r"Изменения в составе бригады",
        r"Изменения в составе исполнителей работ",
        r"Введен в состав бригады",
        r"Введен в состав исполнителей работ",
    ]),
    ("Письменное разрешение эксплуатирующей организации / согласование", [
        r"Письменное\s+разрешение\s+эксплуатирующей\s+организации",
        r"Письменное\s+разрешение\s+\(акт-допуск\)\s+действующего\s+предприятия",
        r"Мероприятия по обеспечению безопасности строительного производства\s+согласованы",
    ]),
    ("Закрытие наряда-допуска", [
        r"Наряд-допуск закрыт",
        r"Работа полностью закончена, бригада удалена",
        r"Работа\s+выполнена\s+в\s+полном\s+объеме",
        r"Работы\s+завершены,\s+рабочие\s+места\s+убраны",
    ]),
]

YES_FILL = PatternFill("solid", fgColor="C6EFCE")
NO_FILL = PatternFill("solid", fgColor="F2F2F2")
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=9)
THIN = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)


def norm_space(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip()


def iter_block_items(parent):
    parent_elm = parent.element.body if isinstance(parent, Doc) else parent._tc
    for child in parent_elm.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, parent)
        elif child.tag == qn("w:tbl"):
            yield Table(child, parent)


def extract_order(text: str) -> tuple[str, str]:
    compact = norm_space(text)
    m = re.search(r"(?:N|№)\s*([0-9]+н)", compact, flags=re.I)
    order = m.group(1) if m else ""
    m2 = re.search(r"Правилам по охране труда (.+?),?\s*утвержденн", compact, flags=re.I)
    topic = norm_space(m2.group(1)) if m2 else ""
    return order, topic


def short_name(title: str, order: str) -> str:
    tl = title.lower()
    if order == "922н":
        if "судовых" in tl:
            return "Водолазные судовые (922н)"
        if "наряд-задание" in tl:
            return "Водолазные наряд-задание (922н)"
        return "Водолазные работы (922н)"
    base = ORDER_NAMES.get(order, title[:35])
    return f"{base} ({order})" if order else base


def extract_literal_components(chunk: list[tuple[str, object]]) -> list[str]:
    """Дословные составляющие бланка: поля, заголовки разделов, заголовки таблиц."""
    comps: list[str] = []
    skip_re = re.compile(
        r"^(Приложение|Рекомендуемый образец|Лицевая сторона|Оборотная сторона|"
        r"ДЛЯ РАБОТЫ|ЗАПОЛНЕНИЮ|Примечание\.?|НАРЯД-ДОПУСК$|Наряд-допуск$|"
        r"НА ПРОИЗВОДСТВО|на производство работ)",
        re.I,
    )

    for kind, payload in chunk:
        if kind == "p":
            t = norm_space(str(payload))
            if not t or skip_re.search(t):
                # keep title НАРЯД-ДОПУСК №...
                if re.search(r"НАРЯД[-\s]?ДОПУСК\s*[N№n]", t, re.I) or re.search(r"Наряд-допуск\s+N", t, re.I):
                    comps.append(t)
                continue
            # numbered sections / labeled fields / blanks
            if (
                re.match(r"^\d+(\.\d+)*\.", t)
                or "____" in t
                or t.endswith(":")
                or re.search(r"(выдал|принял|продлил|закрыт|поручается|начало|окончание|состав|мероприяти|инструктаж|разреш|согласован)", t, re.I)
            ):
                # skip pure underscores continuation lines
                if re.fullmatch(r"[_\s.…—\-]+", t):
                    continue
                # skip parenthetical-only hints unless they contain key terms
                if t.startswith("(") and t.endswith(")") and "электробезопасн" not in t.lower():
                    # keep electro group hints and role hints
                    if any(k in t.lower() for k in ("фамилия", "должность", "подпись", "группа")):
                        comps.append(t)
                    continue
                comps.append(t)
        else:
            rows = payload
            if not rows:
                continue
            header = " | ".join(norm_space(c) for c in rows[0])
            # skip empty/numeric-only headers
            if re.fullmatch(r"[\d\s\|]+", header):
                continue
            comps.append(f"[таблица] {header}")

    # de-dupe preserving order
    out, seen = [], set()
    for c in comps:
        key = c.lower()
        if key not in seen:
            seen.add(key)
            out.append(c)
    return out


def find_literal(text: str, patterns: list[str]) -> str | None:
    for pat in patterns:
        m = re.search(pat, text, flags=re.I | re.DOTALL)
        if m:
            frag = norm_space(m.group(0))
            # trim trailing underscores for readability but keep wording
            frag = re.sub(r"[_…—\-]{3,}$", "", frag).strip(" :")
            if len(frag) > 180:
                frag = frag[:177] + "..."
            return frag
    return None


def build_forms(doc: Document) -> list[dict]:
    blocks: list[tuple[str, object]] = []
    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            t = block.text.strip()
            if t:
                blocks.append(("p", t))
        else:
            rows = []
            for row in block.rows:
                cells, seen = [], set()
                for cell in row.cells:
                    ct = norm_space(cell.text)
                    if ct and ct not in seen:
                        cells.append(ct)
                        seen.add(ct)
                if cells:
                    rows.append(cells)
            if rows:
                blocks.append(("t", rows))

    starts = [i for i, (k, t) in enumerate(blocks) if k == "p" and str(t).startswith("Приложение")]
    forms = []
    for si, start in enumerate(starts):
        end = starts[si + 1] if si + 1 < len(starts) else len(blocks)
        chunk = blocks[start:end]
        appendix = str(chunk[0][1])
        order, topic = extract_order(appendix)
        title = "Наряд-допуск"
        for k, t in chunk[:12]:
            if k == "p" and re.search(r"НАРЯД|Наряд-допуск|Наряд-задание", str(t)):
                title = norm_space(str(t))
                break
        # full text for search (paragraphs + table headers)
        parts = []
        for k, t in chunk:
            if k == "p":
                parts.append(str(t))
            else:
                for row in t:
                    parts.append(" | ".join(row))
        full = "\n".join(parts)
        forms.append(
            {
                "idx": si + 1,
                "name": short_name(title, order),
                "title": title,
                "order": order,
                "topic": topic,
                "full": full,
                "components": extract_literal_components(chunk),
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

    # Comparison matrix with LITERAL text in cells
    matrix = []
    for slot_name, patterns in COMPARE_SLOTS:
        row = {"slot": slot_name}
        for f in forms:
            lit = find_literal(f["full"], patterns)
            row[f["name"]] = lit  # exact fragment or None
        matrix.append(row)

    wb = Workbook()

    # ---- 1. Сравнение: в ячейках дословный текст ----
    ws = wb.active
    ws.title = "Сравнение (дословно)"
    headers = ["Составляющая (для сравнения)"] + [f["name"] for f in forms] + ["Есть в N нарядах"]
    for col, h in enumerate(headers, 1):
        style_header(ws.cell(1, col, h))
    ws.row_dimensions[1].height = 55
    ws.freeze_panes = "B2"

    for ri, row in enumerate(matrix, 2):
        c0 = ws.cell(ri, 1, row["slot"])
        c0.alignment = Alignment(wrap_text=True, vertical="center")
        c0.border = THIN
        c0.font = Font(bold=True, size=9)
        cnt = 0
        for ci, f in enumerate(forms, 2):
            val = row[f["name"]]
            if val:
                cell = ws.cell(ri, ci, val)
                cell.fill = YES_FILL
                cell.font = Font(size=8, color="006100")
                cnt += 1
            else:
                cell = ws.cell(ri, ci, "—")
                cell.fill = NO_FILL
                cell.font = Font(size=8, color="A0A0A0")
            cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
            cell.border = THIN
        tot = ws.cell(ri, len(forms) + 2, cnt)
        tot.alignment = Alignment(horizontal="center", vertical="center")
        tot.font = Font(bold=True)
        tot.border = THIN
        ws.row_dimensions[ri].height = 45

    ws.column_dimensions["A"].width = 42
    for i in range(2, len(forms) + 3):
        ws.column_dimensions[get_column_letter(i)].width = 22

    # ---- 2. Группа по ЭБ отдельно (проверка) ----
    ws_e = wb.create_sheet("Проверка группы по ЭБ")
    style_header(ws_e.cell(1, 1, "Наряд"))
    style_header(ws_e.cell(1, 2, "Приказ"))
    style_header(ws_e.cell(1, 3, "Есть группа по электробезопасности?"))
    style_header(ws_e.cell(1, 4, "Дословная формулировка из бланка"))
    ri = 2
    for f in forms:
        hits = []
        for line in f["full"].split("\n"):
            if "электробезопасн" in line.lower():
                hits.append(norm_space(line))
        ws_e.cell(ri, 1, f["name"]).border = THIN
        ws_e.cell(ri, 2, f["order"]).border = THIN
        if hits:
            ws_e.cell(ri, 3, "ДА").fill = YES_FILL
            ws_e.cell(ri, 3).font = Font(bold=True, color="006100")
            ws_e.cell(ri, 4, "\n".join(hits)).alignment = Alignment(wrap_text=True)
            ws_e.row_dimensions[ri].height = 20 * max(1, len(hits))
        else:
            ws_e.cell(ri, 3, "НЕТ").fill = NO_FILL
            ws_e.cell(ri, 4, "—")
        for col in range(1, 5):
            ws_e.cell(ri, col).border = THIN
            ws_e.cell(ri, col).alignment = Alignment(wrap_text=True, vertical="top")
        ri += 1
    ws_e.column_dimensions["A"].width = 34
    ws_e.column_dimensions["B"].width = 10
    ws_e.column_dimensions["C"].width = 18
    ws_e.column_dimensions["D"].width = 90

    note = (
        "Итог перепроверки: формулировка «группа по электробезопасности» есть только в бланках "
        "903н (электроустановки), 883н (строительство — графа таблицы состава исполнителей) "
        "и 746н (сельское хозяйство — графа таблицы инструктажа). В остальных нарядах — нет."
    )
    ws_e.cell(ri + 1, 1, note)
    ws_e.merge_cells(start_row=ri + 1, start_column=1, end_row=ri + 1, end_column=4)
    ws_e.cell(ri + 1, 1).alignment = Alignment(wrap_text=True)
    ws_e.row_dimensions[ri + 1].height = 50

    # ---- 3. Дословный перечень по каждому наряду ----
    ws2 = wb.create_sheet("Дословно по нарядам")
    style_header(ws2.cell(1, 1, "№"))
    style_header(ws2.cell(1, 2, "Наряд"))
    style_header(ws2.cell(1, 3, "Приказ"))
    style_header(ws2.cell(1, 4, "Составляющие бланка (слово в слово)"))
    style_header(ws2.cell(1, 5, "Кол-во"))
    for ri, f in enumerate(forms, 2):
        ws2.cell(ri, 1, f["idx"]).border = THIN
        ws2.cell(ri, 2, f["name"]).border = THIN
        ws2.cell(ri, 3, f["order"]).border = THIN
        text = "\n".join(f"• {c}" for c in f["components"])
        c = ws2.cell(ri, 4, text)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        c.border = THIN
        ws2.cell(ri, 5, len(f["components"])).border = THIN
        ws2.row_dimensions[ri].height = min(400, max(60, 10 * len(f["components"])))
        for col in range(1, 6):
            ws2.cell(ri, col).alignment = Alignment(wrap_text=True, vertical="top")
            ws2.cell(ri, col).border = THIN
    ws2.column_dimensions["A"].width = 5
    ws2.column_dimensions["B"].width = 32
    ws2.column_dimensions["C"].width = 10
    ws2.column_dimensions["D"].width = 100
    ws2.column_dimensions["E"].width = 10

    # ---- 4. Реестр ----
    ws3 = wb.create_sheet("Реестр", 0)
    for col, h in enumerate(["№", "Наряд", "Область ПОТ", "Приказ", "Составляющих в бланке"], 1):
        style_header(ws3.cell(1, col, h))
    for ri, f in enumerate(forms, 2):
        for ci, v in enumerate([f["idx"], f["name"], f["topic"], f["order"], len(f["components"])], 1):
            c = ws3.cell(ri, ci, v)
            c.border = THIN
            c.alignment = Alignment(wrap_text=True, vertical="center")
    ws3.column_dimensions["A"].width = 5
    ws3.column_dimensions["B"].width = 34
    ws3.column_dimensions["C"].width = 55
    ws3.column_dimensions["D"].width = 10
    ws3.column_dimensions["E"].width = 16

    # ---- 5. Как читать ----
    ws4 = wb.create_sheet("Как читать")
    ws4["A1"] = "Как читать (после перепроверки)"
    ws4["A1"].font = Font(bold=True, size=14)
    notes = [
        "В листе «Сравнение (дословно)» в зелёных ячейках — фрагмент текста ИЗ бланка, не пересказ.",
        "«—» значит, что такой составляющей в бланке нет.",
        "Лист «Проверка группы по ЭБ» — отдельная перепроверка вашего вопроса.",
        "Лист «Дословно по нарядам» — полный перечень полей/разделов/заголовков таблиц каждого бланка.",
        "Группа по электробезопасности: ДА только в 903н, 883н и 746н (см. дословные формулировки).",
    ]
    for i, n in enumerate(notes, 3):
        ws4.cell(i, 1, f"• {n}")
    ws4.column_dimensions["A"].width = 110

    wb.save(OUTPUT)

    # Console verification for electro group
    print("=== ГРУППА ПО ЭЛЕКТРОБЕЗОПАСНОСТИ ===")
    for f in forms:
        hits = [norm_space(l) for l in f["full"].split("\n") if "электробезопасн" in l.lower()]
        if hits:
            print(f"YES {f['name']}:")
            for h in hits:
                print("   ", h[:160])
        else:
            print(f"NO  {f['name']}")
    print(f"\nSaved {OUTPUT}, forms={len(forms)}")


if __name__ == "__main__":
    main()
