#!/usr/bin/env python3
"""Build presentation about steam/hot water pipelines from LMS slide screenshots."""

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Pt

OUT = "/workspace/Презентация_Трубопроводы.pptx"

RED = RGBColor(0xE3, 0x06, 0x13)
DARK = RGBColor(0x1A, 0x1A, 0x1A)
GRAY = RGBColor(0x6B, 0x72, 0x80)
AMBER = RGBColor(0xF5, 0x9E, 0x0B)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_BG = RGBColor(0xF5, 0xF8, 0xFC)
LIGHT_GRAY = RGBColor(0xE5, 0xE7, 0xEB)

W = Emu(12191695)
H = Emu(6858000)
ML = Emu(548640)
CW = Emu(11094415)
TOTAL = 14


def run(p, text, size=13, bold=False, color=DARK):
    r = p.add_run()
    r.text = text
    r.font.name = "Inter"
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    return r


def add_tb(slide, left, top, width, height, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(left, top, width, height)
    box.text_frame.word_wrap = True
    box.text_frame.vertical_anchor = MSO_ANCHOR.TOP
    box.text_frame.paragraphs[0].alignment = align
    return box


def add_bg(slide):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, W, H)
    bg.fill.solid()
    bg.fill.fore_color.rgb = LIGHT_BG
    bg.line.fill.background()


def add_line(slide):
    ln = slide.shapes.add_connector(1, ML, Emu(6355080), ML + CW, Emu(6355080))
    ln.line.color.rgb = RED
    ln.line.width = Pt(1.5)


def add_page(slide, num):
    box = add_tb(slide, Emu(10271455), Emu(6492240), Emu(1188720), Emu(228600), PP_ALIGN.RIGHT)
    p = box.text_frame.paragraphs[0]
    run(p, f"{num:02d} / {TOTAL:02d}", 9, True, RED)


def add_note(slide, text, top=Emu(5852160)):
    box = add_tb(slide, Emu(731520), top, Emu(10728655), Emu(365760))
    p = box.text_frame.paragraphs[0]
    run(p, text, 14, True, AMBER)


def header(slide, section, title, intro=None):
    s = add_tb(slide, ML, Emu(548640), Emu(9000000), Emu(228600))
    run(s.text_frame.paragraphs[0], section, 11, True, RED)
    t = add_tb(slide, ML, Emu(850000), CW, Emu(700000))
    run(t.text_frame.paragraphs[0], title, 28, True, DARK)
    if intro:
        i = add_tb(slide, ML, Emu(1600000), CW, Emu(400000))
        run(i.text_frame.paragraphs[0], intro, 12, False, GRAY)


def numbered_list(slide, items, top=Emu(2150000)):
    y = top
    for idx, item in enumerate(items, 1):
        n = add_tb(slide, ML, y, Emu(411480), Emu(274320))
        run(n.text_frame.paragraphs[0], f"{idx:02d}", 14, True, RED)
        b = add_tb(slide, Emu(1005840), y, Emu(10637215), Emu(520000))
        run(b.text_frame.paragraphs[0], item, 13, False, DARK)
        y += Emu(580000)


def body_text(slide, text, top=Emu(2150000), height=Emu(3600000), size=13, color=DARK):
    box = add_tb(slide, ML, top, CW, height)
    p = box.text_frame.paragraphs[0]
    run(p, text, size, False, color)


def title_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide)

    block = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(365760), Emu(1600000), Emu(6400800), Emu(3200000))
    block.fill.solid()
    block.fill.fore_color.rgb = LIGHT_GRAY
    block.line.fill.background()

    side = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(9144000), 0, Emu(3048000), H)
    side.fill.solid()
    side.fill.fore_color.rgb = LIGHT_GRAY
    side.line.fill.background()

    s = add_tb(slide, Emu(640080), Emu(2000000), Emu(8000000), Emu(228600))
    run(s.text_frame.paragraphs[0], "Промышленная безопасность", 11, True, RED)

    t = add_tb(slide, Emu(640080), Emu(2500000), Emu(8000000), Emu(1800000))
    p = t.text_frame.paragraphs[0]
    run(p, "УСТРОЙСТВО И НАЗНАЧЕНИЕ", 28, True, DARK)
    p2 = t.text_frame.add_paragraph()
    run(p2, "ТРУБОПРОВОДОВ ПАРА", 28, True, DARK)
    p3 = t.text_frame.add_paragraph()
    run(p3, "И ГОРЯЧЕЙ ВОДЫ", 28, True, DARK)

    ln = slide.shapes.add_connector(1, Emu(640080), Emu(4500000), Emu(8500000), Emu(4500000))
    ln.line.color.rgb = DARK
    ln.line.width = Pt(3)
    dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, Emu(8450000), Emu(4470000), Emu(60000), Emu(60000))
    dot.fill.solid()
    dot.fill.fore_color.rgb = RED
    dot.line.fill.background()

    sub = add_tb(slide, Emu(640080), Emu(4700000), Emu(8000000), Emu(600000))
    run(sub.text_frame.paragraphs[0], "ФНП при использовании оборудования, работающего под избыточным давлением", 12, False, GRAY)

    add_line(slide)
    add_page(slide, 1)


def content_slide(prs, num, section, title, intro, items=None, body=None, note=None, highlight=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide)
    header(slide, section, title, intro)

    if highlight:
        box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, ML, Emu(2100000), CW, Emu(700000))
        box.fill.solid()
        box.fill.fore_color.rgb = WHITE
        box.line.color.rgb = RED
        box.line.width = Pt(1.5)
        tb = add_tb(slide, Emu(700000), Emu(2200000), Emu(10500000), Emu(550000))
        p = tb.text_frame.paragraphs[0]
        for part, bold, color in highlight:
            run(p, part, 13, bold, color)

    if items:
        top = Emu(3000000) if highlight else Emu(2150000)
        numbered_list(slide, items, top)
    elif body:
        top = Emu(3000000) if highlight else Emu(2150000)
        body_text(slide, body, top)

    if note:
        add_note(slide, note)

    add_line(slide)
    add_page(slide, num)


def main():
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H

    title_slide(prs)

    content_slide(
        prs, 2,
        "Нормативная база",
        "Федеральные нормы и правила",
        "Приказ Ростехнадзора от 15.12.2020 № 536",
        body=(
            "Об утверждении федеральных норм и правил в области промышленной "
            "безопасности «Правила промышленной безопасности при использовании "
            "оборудования, работающего под избыточным давлением» (далее — ФНП)."
        ),
    )

    content_slide(
        prs, 3,
        "Общие положения",
        "Назначение ФНП",
        "Требования промышленной безопасности при эксплуатации оборудования под давлением",
        body=(
            "ФНП устанавливают единые требования к безопасной эксплуатации "
            "оборудования, работающего под избыточным давлением, и распространяются "
            "на проектирование, монтаж, наладку, эксплуатацию, ремонт, "
            "реконструкцию и техническое освидетельствование такого оборудования."
        ),
    )

    content_slide(
        prs, 4,
        "Общие положения",
        "Область применения ФНП",
        "Требования обязательны при следующих видах работ",
        items=[
            "при разработке технологических процессов",
            "техническом перевооружении опасного производственного объекта (ОПО)",
            "при размещении, монтаже и ремонте",
            "реконструкции (модернизации) оборудования",
            "наладке и эксплуатации",
            "техническом освидетельствовании",
            "техническом диагностировании и экспертизе промышленной безопасности",
        ],
    )

    content_slide(
        prs, 5,
        "Общие положения",
        "К какому оборудованию применяются ФНП",
        "Оборудование, работающее под избыточным давлением",
        items=[
            "паровые и водогрейные котлы",
            "сосуды, работающие под давлением",
            "трубопроводы пара и горячей воды",
            "газовые баллоны и резервуары для сжатых газов",
            "арматура и предохранительные устройства",
        ],
        note="ФНП обязательны для всех стадий жизненного цикла оборудования под давлением.",
    )

    content_slide(
        prs, 6,
        "Устройство трубопроводов",
        "Конструкция трубопроводов пара и горячей воды",
        "Назначение и общие требования к устройству",
        body=(
            "Трубопроводы пара и горячей воды предназначены для транспортировки "
            "теплоносителя от источника тепла к потребителям. Конструкция должна "
            "обеспечивать прочность, герметичность, компенсацию температурных "
            "деформаций, возможность дренажа и отвода воздуха, а также безопасную "
            "эксплуатацию в заданных параметрах давления и температуры."
        ),
    )

    content_slide(
        prs, 7,
        "Устройство трубопроводов",
        "Состав трубопровода (1/2)",
        "Трубопровод состоит из основных элементов",
        items=[
            "плотно соединённых между собой прямых участков труб",
            "фасонных деталей — отводы, переходники, тройники",
            "крепёжных элементов — фланцы, болты, шпильки",
            "арматуры — краны, вентили, задвижки, регулирующие клапаны",
            "редуцирующих и предохранительных клапанов",
        ],
    )

    content_slide(
        prs, 8,
        "Устройство трубопроводов",
        "Приборы КИПиА",
        "Контрольно-измерительные приборы и автоматика",
        items=[
            "манометры — контроль давления в трубопроводе",
            "термометры — контроль температуры теплоносителя",
            "расходомеры — учёт и контроль расхода среды",
            "диафрагмы — измерение расхода по перепаду давления",
        ],
        note="Приборы КИПиА обеспечивают контроль параметров и безопасную эксплуатацию.",
    )

    content_slide(
        prs, 9,
        "Устройство трубопроводов",
        "Опорно-подвесная система",
        "Крепление и поддержание трубопровода",
        items=[
            "неподвижные опоры — фиксируют положение трубопровода",
            "подвижные опоры — допускают перемещение при температурных деформациях",
            "скользящие, катковые и подвесные опоры",
            "пружинные и жёсткие подвески",
            "направляющие и ограничители перемещения",
        ],
    )

    content_slide(
        prs, 10,
        "Устройство трубопроводов",
        "Конденсатоотводчики, дренажи и патрубки",
        "Устройства для отвода конденсата и воздуха",
        body=(
            "В нижних точках каждого отключаемого задвижками участка трубопровода "
            "должны предусматриваться спускные штуцера с запорной арматурой для "
            "опорожнения. В верхних точках устанавливаются воздушники. Нижние "
            "концевые точки паропроводов и изгибов снабжаются устройствами для "
            "продувки. Непрерывный отвод конденсата через конденсационные горшки "
            "или другие устройства обязателен для паропроводов насыщенного пара "
            "и тупиковых участков паропроводов перегретого пара."
        ),
    )

    content_slide(
        prs, 11,
        "Устройство трубопроводов",
        "Компенсаторы температурных деформаций",
        "Для компенсации удлинения трубопровода при нагреве",
        highlight=[
            ("При нагреве трубопровода на ", False, DARK),
            ("100 °C", True, RED),
            (" удлинение стальной трубы составляет около ", False, DARK),
            ("1,2 мм", True, RED),
            (" на 1 м длины.", False, DARK),
        ],
        items=[
            "резиновые компенсаторы с фланцами",
            "П-образные компенсаторы (Z-образные участки)",
            "U-образные (омегаобразные) компенсаторы",
            "Г-образные компенсаторы",
            "сильфонные компенсаторы",
            "линзовые компенсаторы",
            "сальниковые компенсаторы",
        ],
    )

    content_slide(
        prs, 12,
        "Устройство трубопроводов",
        "Заглушки",
        "Элементы для глухого запечатывания выходных отверстий",
        body=(
            "Заглушка — элемент трубопровода, предназначенный для глухого "
            "запечатывания его выходных отверстий; как правило используется при "
            "проведении пневмо- и гидроиспытаний. Металлические концевые заглушки "
            "различаются по виду исполнения и типу крепления."
        ),
        note="Типы: плоские сварные, фланцевые, эллиптические, сферические, быстросъёмные, межфланцевые.",
    )

    content_slide(
        prs, 13,
        "Устройство трубопроводов",
        "Теплоизоляция",
        "Защита от потерь тепла и ожогов",
        body=(
            "Теплоизоляция трубопроводов предназначена для снижения теплопотерь, "
            "предотвращения конденсации влаги на поверхности, защиты персонала от "
            "термических ожогов и обеспечения стабильности технологических "
            "параметров. Применяются минераловатные, пенополиуретановые и другие "
            "изоляционные материалы с защитной оболочкой."
        ),
    )

    content_slide(
        prs, 14,
        "Устройство трубопроводов",
        "Опознавательная окраска",
        "Маркировка трубопроводов по ГОСТ 14202-69",
        items=[
            "коммуникации делятся на 10 основных групп по транспортируемым веществам",
            "зелёный цвет — 1 группа, транспортирует воду",
            "красный цвет — 2 группа, транспортирует пар",
            "предупреждающие кольца: красные — легковоспламеняющиеся и взрывоопасные",
            "жёлтые — токсичные и радиоактивные вещества",
            "зелёные с белой каймой — внутренняя безопасность",
        ],
        note="Количество предупреждающих колец зависит от давления и температуры среды.",
    )

    prs.save(OUT)
    print(f"Saved {OUT} ({len(prs.slides)} slides)")


if __name__ == "__main__":
    main()
