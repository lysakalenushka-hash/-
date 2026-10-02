#!/usr/bin/env python3
"""Build Lab-Intellekt sales deck from the agreed slide structure."""

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import nsmap, qn
from pptx.util import Emu, Inches, Pt
from lxml import etree

NAVY = RGBColor(0x0B, 0x1F, 0x3A)
NAVY2 = RGBColor(0x12, 0x32, 0x58)
TEAL = RGBColor(0x1A, 0x9B, 0x8A)
TEAL_DK = RGBColor(0x0E, 0x6B, 0x5E)
GOLD = RGBColor(0xD4, 0xA0, 0x17)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
OFF = RGBColor(0xF5, 0xF7, 0xFA)
INK = RGBColor(0x1A, 0x23, 0x32)
MUTED = RGBColor(0x5B, 0x67, 0x78)
LINE = RGBColor(0xD8, 0xDE, 0xE6)
RED = RGBColor(0xC0, 0x45, 0x3C)
CARD_NAVY = RGBColor(0x14, 0x3A, 0x62)

FONT = "Calibri"
W = Inches(13.333)
H = Inches(7.5)


def set_run(run, size=18, bold=False, color=INK, font=FONT, italic=False):
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    for tag in ("latin", "ea", "cs"):
        el = rPr.find(qn(f"a:{tag}"))
        if el is None:
            el = etree.SubElement(rPr, qn(f"a:{tag}"))
        el.set("typeface", font)


def add_rect(slide, l, t, w, h, fill, line=None):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
    sh.shadow.inherit = False
    return sh


def add_round(slide, l, t, w, h, fill):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    sh.line.fill.background()
    sh.shadow.inherit = False
    return sh


def tb(slide, l, t, w, h, text, size=18, bold=False, color=INK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, italic=False):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    try:
        tf._txBody.bodyPr.set("anchor", {MSO_ANCHOR.TOP: "t", MSO_ANCHOR.MIDDLE: "ctr", MSO_ANCHOR.BOTTOM: "b"}[anchor])
    except Exception:
        pass
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    set_run(run, size=size, bold=bold, color=color, italic=italic)
    return box


def bullets(slide, l, t, w, h, items, size=18, color=INK, spacing=10, bold_first=False):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(spacing)
        p.level = 0
        run = p.add_run()
        run.text = "•  " + item
        set_run(run, size=size, bold=(bold_first and i == 0), color=color)
    return box


def footer(slide, n, total, light=False):
    c = RGBColor(0x9A, 0xA3, 0xB0) if not light else RGBColor(0x8A, 0xA0, 0xB8)
    tb(slide, Inches(0.5), Inches(7.18), Inches(9), Inches(0.28),
       "Lab-Интеллект  ·  СМАРТА  ·  конфиденциально", size=11, color=c)
    tb(slide, Inches(11.4), Inches(7.18), Inches(1.4), Inches(0.28),
       f"{n}  /  {total}", size=11, color=c, align=PP_ALIGN.RIGHT)


def accent_bar(slide):
    add_rect(slide, Inches(0), Inches(0), Inches(0.12), H, TEAL)


def header(slide, kicker, title, subtitle=None, dark=False):
    accent_bar(slide)
    add_rect(slide, Inches(0), Inches(0), W, Inches(0.08), TEAL if not dark else GOLD)
    kc = TEAL if not dark else GOLD
    tc = INK if not dark else WHITE
    sc = MUTED if not dark else RGBColor(0xB8, 0xC5, 0xD6)
    tb(slide, Inches(0.55), Inches(0.22), Inches(12), Inches(0.32), kicker.upper(), size=12, bold=True, color=kc)
    tb(slide, Inches(0.55), Inches(0.5), Inches(12.2), Inches(0.9), title, size=28, bold=True, color=tc)
    if subtitle:
        tb(slide, Inches(0.55), Inches(1.35), Inches(12.2), Inches(0.45), subtitle, size=16, color=sc)


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def card(slide, l, t, w, h, fill=WHITE, line=LINE):
    sh = add_round(slide, l, t, w, h, fill)
    if line:
        sh.line.fill.solid()
        sh.line.color.rgb = line
        sh.line.width = Pt(1)
    return sh


def add_table(slide, l, t, w, h, rows, col_w=None, header=True):
    n_rows = len(rows)
    n_cols = len(rows[0])
    table_shape = slide.shapes.add_table(n_rows, n_cols, l, t, w, h)
    table = table_shape.table
    if col_w:
        for i, cw in enumerate(col_w):
            table.columns[i].width = cw
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            cell = table.cell(r, c)
            cell.text = ""
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            run = p.add_run()
            run.text = val
            is_h = header and r == 0
            set_run(run, size=13 if not is_h else 13, bold=is_h or (c == 0 and not is_h),
                    color=WHITE if is_h else INK)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            fill = cell.fill
            fill.solid()
            if is_h:
                fill.fore_color.rgb = NAVY
            elif r % 2 == 0:
                fill.fore_color.rgb = OFF
            else:
                fill.fore_color.rgb = WHITE
    return table_shape


def build():
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H
    blank = prs.slide_layouts[6]
    slides = []

    def ns():
        s = prs.slides.add_slide(blank)
        slides.append(s)
        return s

    total = 28  # 20 + demo + 7 backup

    # 1 Title
    s = ns()
    add_rect(s, 0, 0, W, H, NAVY)
    add_rect(s, 0, 0, Inches(0.18), H, TEAL)
    add_rect(s, 0, Inches(7.15), W, Inches(0.35), TEAL)
    tb(s, Inches(0.7), Inches(1.35), Inches(11.5), Inches(0.4),
       "LAB-ИНТЕЛЛЕКТ  ·  СМАРТА", size=14, bold=True, color=TEAL)
    tb(s, Inches(0.7), Inches(1.85), Inches(12), Inches(2.0),
       "Микрообучение по охране труда,\nкоторое сотрудники реально проходят",
       size=36, bold=True, color=WHITE)
    tb(s, Inches(0.7), Inches(4.15), Inches(11), Inches(0.7),
       "Продающая встреча  ·  20 минут + демо  ·  цель: пилот 14–30 дней",
       size=18, color=RGBColor(0xC5, 0xD0, 0xDC))
    tb(s, Inches(0.7), Inches(5.4), Inches(5), Inches(0.7),
       "Клиент: ______________________\nДата: ________________________",
       size=14, color=RGBColor(0x9A, 0xB0, 0xC4))
    tb(s, Inches(7.2), Inches(5.4), Inches(5), Inches(0.7),
       "Спикер: ______________________\nРоль: ОТ / HR / закупки",
       size=14, color=RGBColor(0x9A, 0xB0, 0xC4))
    notes(s, "Не начинать с истории компании. Назвать клиента. Цель встречи — решение по пилоту.")

    # 2 Stake
    s = ns()
    add_rect(s, 0, 0, W, H, OFF)
    header(s, "Ставка встречи", "За 20 минут — понять, стоит ли запускать пилот")
    items = [
        ("01", "Почему формальное обучение не снижает травмы и просрочки"),
        ("02", "Как ежедневные микрокурсы дают вовлечённость и отчёт к проверке"),
        ("03", "Что вы получите за 14–30 дней пилота на своей группе"),
    ]
    for i, (num, txt) in enumerate(items):
        x = Inches(0.55) + i * Inches(4.15)
        card(s, x, Inches(2.15), Inches(3.95), Inches(3.15))
        tb(s, x + Inches(0.25), Inches(2.4), Inches(3.4), Inches(0.5), num, size=28, bold=True, color=TEAL)
        tb(s, x + Inches(0.25), Inches(3.1), Inches(3.45), Inches(1.8), txt, size=18, color=INK)
    tb(s, Inches(0.55), Inches(5.55), Inches(12.2), Inches(0.7),
       "Если цифры не сойдутся — пилот не берём. Если сойдутся — считаем КП под вашу численность.",
       size=16, italic=True, color=NAVY)
    footer(s, 2, total)
    notes(s, "Проговорить правило встречи вслух. Снять давление «нас сейчас продают».")

    # 3 Mirror
    s = ns()
    add_rect(s, 0, 0, W, H, OFF)
    header(s, "Зеркало клиента", "Как обычно выглядит обучение ОТ у компаний вашего масштаба",
           "Заполните до встречи. Если данных нет — спросите: какой сценарий ближе?")
    cols = [
        ("Сейчас", ["Численность / филиалы: ______", "Обучаем в год: ______ чел.", "Формат: очно / СМАРТА / LMS / Excel"]),
        ("Где болит", ["Просрочки и «догоняем к проверке»", "ЕИСОТ и формальный допуск", "Текучка и повторное обучение"]),
        ("Типовой сценарий", ["A. Крупный промышленный, филиалы", "B. 200–1000 чел., нет своей LMS", "C. Уже клиент СМАРТА, нет контура между аттестациями"]),
    ]
    for i, (title, pts) in enumerate(cols):
        x = Inches(0.55) + i * Inches(4.15)
        card(s, x, Inches(2.05), Inches(3.95), Inches(4.4))
        add_rect(s, x, Inches(2.05), Inches(3.95), Inches(0.08), TEAL)
        tb(s, x + Inches(0.25), Inches(2.3), Inches(3.45), Inches(0.45), title, size=18, bold=True, color=NAVY)
        bullets(s, x + Inches(0.2), Inches(2.85), Inches(3.55), Inches(3.2), pts, size=15, spacing=12)
    footer(s, 3, total)

    # 4 Problem 1
    s = ns()
    add_rect(s, 0, 0, W, H, OFF)
    header(s, "Проблема 1", "Сотрудник «прошёл курс» — и на следующий день работает как раньше")
    left = [
        "8–40 часов раз в год ≠ навык в момент риска",
        "Знания выветриваются между аттестациями",
        "Куратор видит факт прохождения, не видит вовлечённость",
    ]
    card(s, Inches(0.55), Inches(2.05), Inches(6.3), Inches(4.35))
    bullets(s, Inches(0.85), Inches(2.3), Inches(5.8), Inches(3.8), left, size=18, spacing=18)
    card(s, Inches(7.1), Inches(2.05), Inches(5.7), Inches(4.35), fill=NAVY)
    tb(s, Inches(7.4), Inches(2.4), Inches(5.2), Inches(0.4), "ДВА ФОРМАТА", size=12, bold=True, color=TEAL)
    tb(s, Inches(7.4), Inches(2.9), Inches(5.2), Inches(1.4),
       "Длинный курс\nраз в год — галочка", size=22, bold=True, color=WHITE)
    tb(s, Inches(7.4), Inches(4.5), Inches(5.2), Inches(1.4),
       "5–15 минут\nв рабочий день — привычка", size=22, bold=True, color=GOLD)
    footer(s, 4, total)
    notes(s, "Связка: формат длинного курса против ежедневного микроблока.")

    # 5 Problem 2
    s = ns()
    add_rect(s, 0, 0, W, H, OFF)
    header(s, "Проблема 2", "Если обучение не измеряется каждый день, им нельзя управлять")
    metrics = [
        ("20%", "работников в мире вовлечены", "Gallup, State of the Global Workplace 2026, данные 2025"),
        ("−63%", "инцидентов у команд верхнего квартиля вовлечённости vs нижнего", "Gallup Q12 Meta-Analysis, 11th ed., 2024"),
        ("64%", "индекс вовлечённости в России — зона риска", "ЭКОПСИ, мониторинг 2024"),
    ]
    for i, (n, cap, src) in enumerate(metrics):
        x = Inches(0.55) + i * Inches(4.15)
        card(s, x, Inches(2.05), Inches(3.95), Inches(3.55))
        tb(s, x + Inches(0.25), Inches(2.25), Inches(3.45), Inches(0.9), n, size=40, bold=True, color=TEAL)
        tb(s, x + Inches(0.25), Inches(3.2), Inches(3.45), Inches(1.3), cap, size=16, color=INK)
        tb(s, x + Inches(0.25), Inches(4.6), Inches(3.45), Inches(0.7), src, size=11, color=MUTED)
    tb(s, Inches(0.55), Inches(5.8), Inches(12.2), Inches(0.7),
       "Для охраны труда: вовлечённость — ведущий индикатор. Травма и штраф — запаздывающий.",
       size=16, italic=True, color=NAVY)
    footer(s, 5, total)
    notes(s, "Не сравнивать 20% Gallup и 64% ЭКОПСИ как одну шкалу. Источники — запасной слайд F.")

    # 6 Cost of inaction
    s = ns()
    add_rect(s, 0, 0, W, H, OFF)
    header(s, "Цена бездействия", "Сколько стоит «как сейчас» — на одном филиале, не в триллионах")
    cols = [
        ("Штраф / предписание", "Предписание ГИТ, приостановка, повторная проверка", "Ваша цифра: ________ ₽"),
        ("Потеря смены / НС", "Больничный, простой, расследование, репутация", "Ваша цифра: ________ ₽"),
        ("Обучение «для галочки»", "Повторные курсы, отрыв от смены, работа куратора в Excel", "Ваша цифра: ________ ₽"),
    ]
    for i, (t, d, f) in enumerate(cols):
        x = Inches(0.55) + i * Inches(4.15)
        card(s, x, Inches(2.05), Inches(3.95), Inches(4.35))
        add_rect(s, x, Inches(2.05), Inches(3.95), Inches(0.1), RED)
        tb(s, x + Inches(0.25), Inches(2.35), Inches(3.45), Inches(1.0), t, size=18, bold=True, color=NAVY)
        tb(s, x + Inches(0.25), Inches(3.4), Inches(3.45), Inches(1.4), d, size=15, color=MUTED)
        tb(s, x + Inches(0.25), Inches(5.2), Inches(3.45), Inches(0.7), f, size=15, bold=True, color=RED)
    footer(s, 6, total)
    notes(s, "Спросить вслух их цифру. Если нет — оставить пропуски и вернуться в КП.")

    # 7 Insight
    s = ns()
    add_rect(s, 0, 0, W, H, OFF)
    header(s, "Инсайт", "Между аттестациями нужен второй контур — короткий, ежедневный, с рейтингом")
    loops = [
        ("1", "Аттестация / ЕИСОТ", "Раз в год.\nФормальный допуск.\nМожно оставить текущему провайдеру или СМАРТА."),
        ("2", "Микрообучение", "Каждый рабочий день.\n5–15 минут.\nНавык и привычка — это Lab-Интеллект."),
        ("3", "Аналитика руководителя", "Просрочки, баллы,\nвовлечённость.\nОтчёт к проверке — это Lab-Интеллект."),
    ]
    for i, (n, t, d) in enumerate(loops):
        x = Inches(0.55) + i * Inches(4.15)
        card(s, x, Inches(2.05), Inches(3.95), Inches(4.0), fill=NAVY if i else WHITE)
        tc, dc = (WHITE, RGBColor(0xC5, 0xD0, 0xDC)) if i else (INK, MUTED)
        tb(s, x + Inches(0.25), Inches(2.25), Inches(3.45), Inches(0.5), n, size=28, bold=True, color=TEAL if i else GOLD)
        tb(s, x + Inches(0.25), Inches(2.85), Inches(3.45), Inches(0.7), t, size=18, bold=True, color=tc)
        tb(s, x + Inches(0.25), Inches(3.6), Inches(3.45), Inches(2.1), d, size=15, color=dc)
    footer(s, 7, total)

    # 8 Product
    s = ns()
    add_rect(s, 0, 0, W, H, OFF)
    header(s, "Продукт", "Корпоративная платформа микрокурсов по охране труда")
    feats = [
        ("Микрокурсы", "Видео, PDF, презентации, вопросы проверки знаний"),
        ("Программы и назначения", "Группы, каждый день или по рабочим дням, последовательный трек"),
        ("Баллы и рейтинг", "Вовлечённость, своевременность, место в таблице"),
        ("Кабинет компании", "Штатка, филиалы, роли: admin / curator / user"),
        ("Нормативка и учёт", "Поля СНИЛС / ЕИСОТ, нормативные документы в курсах"),
        ("Отчёты", "Прогресс, просрочки, выгрузка без Excel-рутины"),
    ]
    for i, (t, d) in enumerate(feats):
        r, c = divmod(i, 3)
        x = Inches(0.55) + c * Inches(4.15)
        y = Inches(2.0) + r * Inches(2.25)
        card(s, x, y, Inches(3.95), Inches(2.05))
        tb(s, x + Inches(0.25), y + Inches(0.25), Inches(3.45), Inches(0.5), t, size=16, bold=True, color=NAVY)
        tb(s, x + Inches(0.25), y + Inches(0.8), Inches(3.45), Inches(1.0), d, size=14, color=MUTED)
    footer(s, 8, total)
    notes(s, "Не обещать ИИ-адаптацию, симуляции и сертификаты 24/7, если не покажете в демо.")

    # 9 Employee
    s = ns()
    add_rect(s, 0, 0, W, H, OFF)
    header(s, "Глазами сотрудника", "5–15 минут, телефон или компьютер. Один блок — один навык")
    pts = [
        "Актуальные микрокурсы на сегодня",
        "Последовательный трек: следующий курс откроется после текущего",
        "Вопрос проверки знаний и баллы за попытку",
        "Статус «все курсы на сегодня пройдены»",
    ]
    card(s, Inches(0.55), Inches(2.05), Inches(6.5), Inches(4.35))
    bullets(s, Inches(0.85), Inches(2.35), Inches(6.0), Inches(3.6), pts, size=18, spacing=16)
    card(s, Inches(7.3), Inches(2.05), Inches(5.5), Inches(4.35), fill=TEAL_DK)
    tb(s, Inches(7.6), Inches(2.5), Inches(5.0), Inches(1.2), "Моё обучение", size=22, bold=True, color=WHITE)
    tb(s, Inches(7.6), Inches(3.8), Inches(5.0), Inches(2.0),
       "Это не ещё одна LMS.\nЭто смена, которую можно закрыть до обеда.",
       size=18, color=WHITE)
    footer(s, 9, total)

    # 10 Manager
    s = ns()
    add_rect(s, 0, 0, W, H, OFF)
    header(s, "Глазами ОТ / HR", "Кто не прошёл, кто просрочил, кто вовлечён — без Excel")
    weights = [
        ("Прохождение блока", "10"),
        ("Вопрос проверки", "40"),
        ("Вовлечённость", "20"),
        ("Доп. информация", "20"),
        ("Своевременность", "10"),
    ]
    for i, (t, wgt) in enumerate(weights):
        x = Inches(0.55) + i * Inches(2.5)
        card(s, x, Inches(2.05), Inches(2.35), Inches(2.35))
        tb(s, x + Inches(0.15), Inches(2.25), Inches(2.05), Inches(0.7), wgt + "%", size=28, bold=True, color=TEAL, align=PP_ALIGN.CENTER)
        tb(s, x + Inches(0.1), Inches(3.1), Inches(2.15), Inches(1.0), t, size=13, color=INK, align=PP_ALIGN.CENTER)
    pts = ["Прогресс программ", "Своевременность и просрочки", "Рейтинг сотрудников и баллы", "Выгрузка отчёта к проверке"]
    bullets(s, Inches(0.55), Inches(4.6), Inches(12.2), Inches(1.8), pts, size=16, spacing=8)
    footer(s, 10, total)
    notes(s, "Веса по умолчанию в продукте. Компания может пересчитать критерии.")

    # 11 Differentiation
    s = ns()
    add_rect(s, 0, 0, W, H, OFF)
    header(s, "Отличие", "Не «ещё одна LMS», а контур ОТ + микроформат + СМАРТА")
    rows = [
        ["", "Классическая LMS", "Разовое дистанционное ОТ", "Lab-Интеллект"],
        ["Формат", "Длинные курсы", "Галочка к проверке", "Ежедневные микрокурсы"],
        ["Вовлечённость", "Слабая", "Нет рейтинга", "Баллы и рейтинг"],
        ["Внедрение", "Проект IT", "Подрядчик обучения", "Кабинет, импорт CSV/XLSX"],
        ["Контент", "Общий", "Только аттестация", "Нормативка + поддерживающее обучение"],
        ["Место в контуре", "Замена всего", "Раз в год", "364 дня между аттестациями"],
    ]
    add_table(s, Inches(0.45), Inches(1.95), Inches(12.4), Inches(4.35), rows,
              col_w=[Inches(2.2), Inches(3.3), Inches(3.4), Inches(3.5)])
    footer(s, 11, total)
    notes(s, "Если уже клиент СМАРТА: «Вы уже учите людей. Мы закрываем 364 дня между аттестациями.»")

    # DEMO interstitial
    s = ns()
    add_rect(s, 0, 0, W, H, NAVY)
    add_rect(s, 0, 0, Inches(0.18), H, GOLD)
    tb(s, Inches(0.7), Inches(2.0), Inches(12), Inches(0.4), "ДЕМО  ·  8–10 МИНУТ", size=14, bold=True, color=GOLD)
    tb(s, Inches(0.7), Inches(2.5), Inches(12), Inches(1.4),
       "Смотрим продукт, не слайды", size=36, bold=True, color=WHITE)
    tb(s, Inches(0.7), Inches(4.1), Inches(12), Inches(1.4),
       "1. Сотрудник проходит один микрокурс с телефона\n2. Руководитель видит рейтинг и просрочки\n3. Вопросы — после, не во время кликов",
       size=18, color=RGBColor(0xC5, 0xD0, 0xDC))
    notes(s, "Не заменять демо скриншотами. Второй экран — кабинет руководителя.")

    # 12 How it works
    s = ns()
    add_rect(s, 0, 0, W, H, OFF)
    header(s, "Как это работает", "Назначили → проходят каждый день → видите отчёт")
    steps = [
        ("01", "Назначение", "Программа на группу или должность. Срок, рабочие дни, последовательный трек."),
        ("02", "Микрообучение", "Ежедневные блоки: материал + вопрос проверки. Баллы за попытку."),
        ("03", "Контроль", "Рейтинг, просрочки, вовлечённость. Отчёт к проверке и разбор с куратором."),
    ]
    for i, (n, t, d) in enumerate(steps):
        x = Inches(0.55) + i * Inches(4.15)
        card(s, x, Inches(2.1), Inches(3.95), Inches(4.2))
        tb(s, x + Inches(0.25), Inches(2.35), Inches(3.45), Inches(0.55), n, size=26, bold=True, color=TEAL)
        tb(s, x + Inches(0.25), Inches(3.0), Inches(3.45), Inches(0.55), t, size=20, bold=True, color=NAVY)
        tb(s, x + Inches(0.25), Inches(3.7), Inches(3.45), Inches(2.1), d, size=16, color=MUTED)
    footer(s, 12, total)
    notes(s, "Персонализация — через программы, должности и назначения, не через обещание ИИ.")

    # 13 Proof
    s = ns()
    add_rect(s, 0, 0, W, H, OFF)
    header(s, "Доказательства формата", "Микрообучение и вовлечённость уже работают в охране труда")
    cards = [
        ("Gallup Q12, 2024", "Подразделения с высокой вовлечённостью: на 63% меньше инцидентов по безопасности vs нижний квартиль.", "Метаанализ 3,35 млн сотрудников, 183 806 подразделений"),
        ("FSC / Signal Gold", "Микрообучение на шахте без отрыва от смены. Гибкая подача + поведение в H&S.", "Оценка программы, Канада, 2024"),
        ("Кейсы формата, не наши клиенты", "Walmart: ежедневные 3–5 мин квизы по safety. Bloomingdale’s — снижение claims.", "Помечать явно: зарубежный формат, не кейс СМАРТА"),
    ]
    for i, (t, d, n) in enumerate(cards):
        x = Inches(0.55) + i * Inches(4.15)
        card(s, x, Inches(2.05), Inches(3.95), Inches(4.35))
        tb(s, x + Inches(0.25), Inches(2.25), Inches(3.45), Inches(0.8), t, size=16, bold=True, color=NAVY)
        tb(s, x + Inches(0.25), Inches(3.15), Inches(3.45), Inches(2.0), d, size=15, color=INK)
        tb(s, x + Inches(0.25), Inches(5.3), Inches(3.45), Inches(0.8), n, size=12, color=MUTED)
    footer(s, 13, total)
    notes(s, "Свои кейсы СМАРТА ставить первыми, если есть цифры. Чужие — только как доказательство формата.")

    # 14 Mini case
    s = ns()
    add_rect(s, 0, 0, W, H, OFF)
    header(s, "Сценарий на 30 дней", "Как это выглядит на 500 сотрудниках в двух филиалах")
    left_pts = [
        "Программа «Базовая безопасность» — 10 микрокурсов",
        "10 минут × рабочие дни",
        "Должности: стропка / склад / водители / ИТР — подставить ваши",
        "Пилотная группа 50–100, затем масштаб",
    ]
    card(s, Inches(0.55), Inches(2.05), Inches(6.4), Inches(4.35))
    tb(s, Inches(0.8), Inches(2.25), Inches(5.9), Inches(0.45), "Вход", size=14, bold=True, color=TEAL)
    bullets(s, Inches(0.8), Inches(2.75), Inches(5.9), Inches(3.3), left_pts, size=16, spacing=12)
    card(s, Inches(7.2), Inches(2.05), Inches(5.6), Inches(4.35), fill=NAVY)
    tb(s, Inches(7.5), Inches(2.25), Inches(5.1), Inches(0.45), "Через 30 дней смотрим", size=14, bold=True, color=GOLD)
    tb(s, Inches(7.5), Inches(2.85), Inches(5.1), Inches(3.2),
       "% завершения назначенного\n\n% прошедших вовремя\n\nТоп просрочек по отделам\n\nВовлечённость и баллы",
       size=18, color=WHITE)
    footer(s, 14, total)

    # 15 Implementation
    s = ns()
    add_rect(s, 0, 0, W, H, OFF)
    header(s, "Внедрение", "Пилот за 2 недели, не IT-проект на полгода")
    steps = [
        ("Неделя 1", ["Кабинет компании и роли", "Импорт сотрудников CSV / XLSX", "1 готовая программа + 1 ваша тема"]),
        ("Неделя 2", ["Назначение группе 50–100 чел.", "Kick-off 1 час с куратором", "Старт ежедневных микрокурсов"]),
        ("День 14 / 30", ["Отчёт: завершение, вовремя, вовлечённость", "Разбор с куратором", "Решение: КП и масштаб"]),
    ]
    for i, (t, pts) in enumerate(steps):
        x = Inches(0.55) + i * Inches(4.15)
        card(s, x, Inches(2.05), Inches(3.95), Inches(4.35))
        tb(s, x + Inches(0.25), Inches(2.25), Inches(3.45), Inches(0.6), t, size=18, bold=True, color=TEAL)
        bullets(s, x + Inches(0.2), Inches(3.0), Inches(3.55), Inches(3.0), pts, size=16, spacing=14)
    footer(s, 15, total)

    # 16 Commercial
    s = ns()
    add_rect(s, 0, 0, W, H, OFF)
    header(s, "Коммерческая логика", "Считаем от риска и объёма, не от «стоимости LMS»")
    tb(s, Inches(0.55), Inches(1.95), Inches(12.2), Inches(0.55),
       "Стоимость штрафа / простоя / повторного обучения   vs   лицензия на сотрудника в год",
       size=18, bold=True, color=NAVY)
    pkgs = [
        ("Старт", "до 100 человек", "Базовые программы\nКабинет и назначения\nОтчёт куратора"),
        ("Бизнес", "до 500 человек", "Безлимит программ\nРейтинг и веса критериев\nНесколько групп и отделов"),
        ("Корпоратив", "500+ / филиалы", "Кабинеты филиалов\nЕИСОТ-контур, SLA\nКастомные программы"),
    ]
    for i, (t, s2, d) in enumerate(pkgs):
        x = Inches(0.55) + i * Inches(4.15)
        fill = NAVY if i == 1 else WHITE
        tc, dc = (WHITE, RGBColor(0xC5, 0xD0, 0xDC)) if i == 1 else (INK, MUTED)
        card(s, x, Inches(2.65), Inches(3.95), Inches(3.35), fill=fill)
        tb(s, x + Inches(0.25), Inches(2.85), Inches(3.45), Inches(0.45), t, size=20, bold=True, color=TEAL if i == 1 else NAVY)
        tb(s, x + Inches(0.25), Inches(3.35), Inches(3.45), Inches(0.4), s2, size=14, bold=True, color=GOLD if i == 1 else TEAL)
        tb(s, x + Inches(0.25), Inches(3.9), Inches(3.45), Inches(1.8), d, size=15, color=dc)
    tb(s, Inches(0.55), Inches(6.15), Inches(12.2), Inches(0.5),
       "Точную цифру дадим после пилота и вашей численности. Сегодня фиксируем логику.",
       size=14, italic=True, color=MUTED)
    footer(s, 16, total)

    # 17 Objections
    s = ns()
    add_rect(s, 0, 0, W, H, OFF)
    header(s, "Возражения", "Три страха, которые обычно останавливают сделку")
    rows = [
        ["Страх", "Как закрываем"],
        ["Сотрудники не будут проходить", "5–15 минут, рейтинг, назначения на рабочие дни, отчёт руководителю смены"],
        ["Не примут на проверке", "Нормативка в курсах, контур СМАРТА, отчётность, поля ЕИСОТ / СНИЛС"],
        ["Уже есть LMS", "Не замена аттестации. Второй контур. Может жить рядом с текущей LMS"],
    ]
    add_table(s, Inches(0.55), Inches(2.05), Inches(12.2), Inches(4.2), rows, col_w=[Inches(4.4), Inches(7.8)])
    footer(s, 17, total)

    # 18 Offer
    s = ns()
    add_rect(s, 0, 0, W, H, OFF)
    header(s, "Офер встречи", "Пилот 14–30 дней на 50–100 сотрудников")
    card(s, Inches(0.55), Inches(2.05), Inches(6.3), Inches(4.4))
    tb(s, Inches(0.85), Inches(2.25), Inches(5.8), Inches(0.45), "Что входит", size=16, bold=True, color=TEAL)
    bullets(s, Inches(0.8), Inches(2.8), Inches(5.8), Inches(3.3), [
        "Кабинет компании",
        "1 готовая программа",
        "1 кастомная тема — по возможности",
        "Отчёт: вовлечённость, завершение, своевременность",
        "Разбор с куратором",
    ], size=16, spacing=10)
    card(s, Inches(7.1), Inches(2.05), Inches(5.7), Inches(4.4), fill=NAVY)
    tb(s, Inches(7.4), Inches(2.25), Inches(5.2), Inches(0.45), "Что нужно от вас", size=16, bold=True, color=GOLD)
    tb(s, Inches(7.4), Inches(2.9), Inches(5.2), Inches(2.2),
       "Список пилотной группы\nКуратор со стороны клиента\n1 час kick-off",
       size=18, color=WHITE)
    tb(s, Inches(7.4), Inches(5.3), Inches(5.2), Inches(0.8),
       "Один вопрос: дата старта пилота", size=16, bold=True, color=TEAL)
    footer(s, 18, total)

    # 19 Close
    s = ns()
    add_rect(s, 0, 0, W, H, OFF)
    header(s, "Закрытие", "Три решения на этой встрече")
    dec = [
        ("1", "Пилот", "Да / нет / какие условия"),
        ("2", "Куратор", "ФИО и роль со стороны клиента"),
        ("3", "Kick-off", "Дата и час запуска"),
    ]
    for i, (n, t, d) in enumerate(dec):
        x = Inches(0.55) + i * Inches(4.15)
        card(s, x, Inches(2.15), Inches(3.95), Inches(3.5))
        tb(s, x + Inches(0.25), Inches(2.4), Inches(3.45), Inches(0.6), n, size=32, bold=True, color=TEAL)
        tb(s, x + Inches(0.25), Inches(3.15), Inches(3.45), Inches(0.55), t, size=22, bold=True, color=NAVY)
        tb(s, x + Inches(0.25), Inches(3.85), Inches(3.45), Inches(1.2), d, size=16, color=MUTED)
    tb(s, Inches(0.55), Inches(5.9), Inches(12.2), Inches(0.5),
       "Не «мы вам перезвоним». Фиксируем дату до выхода из комнаты.",
       size=16, italic=True, color=NAVY)
    footer(s, 19, total)

    # 20 Thanks
    s = ns()
    add_rect(s, 0, 0, W, H, NAVY)
    add_rect(s, 0, 0, Inches(0.18), H, TEAL)
    tb(s, Inches(0.7), Inches(2.0), Inches(12), Inches(0.4), "LAB-ИНТЕЛЛЕКТ  ·  СМАРТА", size=14, bold=True, color=TEAL)
    tb(s, Inches(0.7), Inches(2.5), Inches(12), Inches(1.0), "Спасибо. Следующий шаг — дата пилота.", size=32, bold=True, color=WHITE)
    tb(s, Inches(0.7), Inches(4.0), Inches(12), Inches(1.6),
       "Спикер: ______________________\nТелефон: _____________________\nПочта: ________________________\nКабинет: adddev.smarta-ot.ru",
       size=18, color=RGBColor(0xC5, 0xD0, 0xDC))
    notes(s, "Не добавлять новый смысл. Только контакты и пауза под вопросы.")

    # Backup A
    s = ns()
    add_rect(s, 0, 0, W, H, OFF)
    header(s, "Запас A  ·  по вопросу", "Роли в кабинете")
    roles = [
        ("Superadmin", "Платформа, кабинеты компаний, управление курсами"),
        ("Admin", "Компания: сотрудники, структура, настройки"),
        ("Curator", "Программы, группы, назначения, контроль прохождения"),
        ("User", "Моё обучение: ежедневные микрокурсы и рейтинг"),
    ]
    for i, (t, d) in enumerate(roles):
        y = Inches(1.95) + i * Inches(1.15)
        card(s, Inches(0.55), y, Inches(12.2), Inches(1.05))
        tb(s, Inches(0.85), y + Inches(0.28), Inches(3.2), Inches(0.5), t, size=18, bold=True, color=TEAL)
        tb(s, Inches(4.2), y + Inches(0.28), Inches(8.2), Inches(0.55), d, size=16, color=INK)
    footer(s, 21, total)

    # Backup B
    s = ns()
    add_rect(s, 0, 0, W, H, OFF)
    header(s, "Запас B  ·  по вопросу", "Импорт штатки и структура компании")
    bullets(s, Inches(0.55), Inches(2.1), Inches(12.2), Inches(4.5), [
        "Файл CSV, XLSX или XLS — без ручного заведения сотен людей",
        "Должность, отдел, контакты, генерация пароля",
        "Группы и назначения программ поверх штатки",
        "Филиалы / кабинеты — для распределённых компаний",
        "IT не нужен: импорт делает куратор или админ компании",
    ], size=18, spacing=14)
    footer(s, 22, total)

    # Backup C
    s = ns()
    add_rect(s, 0, 0, W, H, OFF)
    header(s, "Запас C  ·  по вопросу", "Чем не iSpring / WebSoft / «просто LMS»")
    rows = [
        ["Критерий", "Универсальная LMS", "Lab-Интеллект"],
        ["Ниша", "Любое корпоративное обучение", "Охрана труда, поддерживающий контур"],
        ["Формат", "Курсы любой длины", "Ежедневные микрокурсы"],
        ["Метрики", "Прохождение курса", "Вовлечённость, своевременность, рейтинг"],
        ["Учёт ОТ", "Настраивается отдельно", "СНИЛС, ЕИСОТ, нормативка в модели"],
        ["Подрядчик", "IT / L&D", "Экосистема СМАРТА: аттестация + микрообучение"],
    ]
    add_table(s, Inches(0.45), Inches(1.95), Inches(12.4), Inches(4.5), rows,
              col_w=[Inches(2.4), Inches(5.0), Inches(5.0)])
    footer(s, 23, total)

    # Backup D
    s = ns()
    add_rect(s, 0, 0, W, H, OFF)
    header(s, "Запас D  ·  по вопросу", "Контур, кабинет, данные")
    bullets(s, Inches(0.55), Inches(2.1), Inches(12.2), Inches(4.5), [
        "Отдельный кабинет компании (субдомен)",
        "Роли и доступ: сотрудник не видит чужие назначения и аналитику компании",
        "Авторизация, восстановление пароля, учётные записи штатки",
        "Площадка в контуре СМАРТА / smarta-ot.ru",
        "Детали ИБ, ДПДн и размещения серверов — в приложении к договору, не на этом слайде",
    ], size=18, spacing=14)
    footer(s, 24, total)

    # Backup E
    s = ns()
    add_rect(s, 0, 0, W, H, OFF)
    header(s, "Запас E  ·  по вопросу", "Стартовый контент для пилота")
    packs = [
        ("Базовая охрана труда", "Короткие блоки под рабочие ситуации, не 40-часовой курс"),
        ("Пожарная безопасность", "Повторение критичных действий между аттестациями"),
        ("Первая помощь", "Навык, который выветривается без регулярного касания"),
        ("Ваша тема", "1 кастомный микрокурс под инцидент или профессию клиента"),
    ]
    for i, (t, d) in enumerate(packs):
        r, c = divmod(i, 2)
        x = Inches(0.55) + c * Inches(6.3)
        y = Inches(2.05) + r * Inches(2.2)
        card(s, x, y, Inches(6.05), Inches(2.0))
        tb(s, x + Inches(0.3), y + Inches(0.3), Inches(5.45), Inches(0.5), t, size=18, bold=True, color=NAVY)
        tb(s, x + Inches(0.3), y + Inches(0.9), Inches(5.45), Inches(0.8), d, size=15, color=MUTED)
    footer(s, 25, total)

    # Backup F
    s = ns()
    add_rect(s, 0, 0, W, H, OFF)
    header(s, "Запас F  ·  источники", "Цифры, на которые можно ссылаться")
    rows = [
        ["Цифра", "Источник"],
        ["20% вовлечённых в мире (2025)", "Gallup. State of the Global Workplace: 2026 Report"],
        ["−63% инцидентов, верх vs низ квартиля", "Gallup Q12 Meta-Analysis, 11th Edition, 2024"],
        ["64% индекс вовлечённости в РФ (2024)", "ЭКОПСИ, Всероссийский мониторинг вовлечённости"],
        ["Микрообучение на шахте без отрыва от смены", "FSC Canada / Signal Gold, evaluation 2024"],
        ["Ежедневные safety-квизы (формат, не наш клиент)", "Публичные кейсы Walmart / Bloomingdale’s — проверять первоисточник"],
    ]
    add_table(s, Inches(0.4), Inches(1.95), Inches(12.5), Inches(4.5), rows, col_w=[Inches(5.6), Inches(6.9)])
    footer(s, 26, total)

    # Backup G
    s = ns()
    add_rect(s, 0, 0, W, H, OFF)
    header(s, "Запас G  ·  клиенты СМАРТА", "Как стыкуется с централизованным обучением")
    bullets(s, Inches(0.55), Inches(2.1), Inches(12.2), Inches(4.5), [
        "Аттестация и программы по ОТ — как сейчас, у СМАРТА",
        "Lab-Интеллект — ежедневное поддерживающее обучение между циклами",
        "Один подрядчик: контент, кабинет, куратор, отчётность",
        "Кросс-продажа в действующий договор, без нового тендера там, где это допустимо",
        "Пилот на уже обученной группе: быстрее импорт и понятнее KPI",
    ], size=18, spacing=14)
    footer(s, 27, total)

    out = "/workspace/issledovaniya/Lab-Intellekt_prodayushchaya-prezentaciya.pptx"
    prs.save(out)
    print(out, "slides", len(prs.slides))


if __name__ == "__main__":
    build()
