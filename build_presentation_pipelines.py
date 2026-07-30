#!/usr/bin/env python3
"""Build 1:1 pipeline presentation matching LMS screenshot style with images."""

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Pt

OUT = "/workspace/Презентация_Трубопроводы.pptx"
IMG = "/workspace/assets/pipeline_images/generated"

W = Emu(12191695)
H = Emu(6858000)
ML = Emu(457200)
TEXT_W = Emu(6400800)
IMG_LEFT = Emu(7000000)
IMG_W = Emu(5000000)
TOTAL = 14

ORANGE = RGBColor(0xFF, 0x66, 0x00)
DARK = RGBColor(0x33, 0x33, 0x33)
GRAY = RGBColor(0x66, 0x66, 0x66)
RED = RGBColor(0xE3, 0x06, 0x13)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GRAY = RGBColor(0xE5, 0xE7, 0xEB)
FONT = "Arial"


def run(p, text, size=12, bold=False, color=DARK, italic=False):
    r = p.add_run()
    r.text = text
    r.font.name = FONT
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = color
    return r


def tb(slide, left, top, width, height, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.paragraphs[0].alignment = align
    return box


def bg(slide):
    rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, W, H)
    rect.fill.solid()
    rect.fill.fore_color.rgb = WHITE
    rect.line.fill.background()


def header_line(slide, top=Emu(1050000)):
    orange = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, ML, top, Emu(4800000), Emu(36000))
    orange.fill.solid()
    orange.fill.fore_color.rgb = ORANGE
    orange.line.fill.background()
    gray = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(5257200), top, Emu(5600000), Emu(18000))
    gray.fill.solid()
    gray.fill.fore_color.rgb = DARK
    gray.line.fill.background()


def section_header(slide, title, subtitle=None):
    t = tb(slide, ML, Emu(350000), TEXT_W, Emu(500000))
    run(t.text_frame.paragraphs[0], title, 20, True, DARK)
    header_line(slide)
    if subtitle:
        s = tb(slide, ML, Emu(1250000), TEXT_W, Emu(350000))
        run(s.text_frame.paragraphs[0], subtitle, 13, True, DARK)


def slide_num(slide, num, top=Emu(2200000)):
    sq = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(180000), top, Emu(420000), Emu(420000))
    sq.fill.solid()
    sq.fill.fore_color.rgb = ORANGE
    sq.line.fill.background()
    n = tb(slide, Emu(180000), top + Emu(40000), Emu(420000), Emu(340000), PP_ALIGN.CENTER)
    run(n.text_frame.paragraphs[0], str(num), 18, True, WHITE)


def add_image(slide, path, left=IMG_LEFT, top=Emu(1200000), width=IMG_W, height=Emu(5200000)):
    slide.shapes.add_picture(path, left, top, width=width, height=height)


def add_image_bottom(slide, path, top=Emu(3600000), height=Emu(2800000)):
    slide.shapes.add_picture(path, Emu(900000), top, width=Emu(10400000), height=height)


def bullets(slide, items, top=Emu(2200000), numbered=False, highlight_first=False):
    y = top
    for i, item in enumerate(items, 1):
        if numbered:
            nbox = tb(slide, ML + Emu(500000), y, Emu(350000), Emu(280000))
            run(nbox.text_frame.paragraphs[0], f"{i})", 12, True, ORANGE)
            body_left = ML + Emu(900000)
        else:
            dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, ML + Emu(520000), y + Emu(80000), Emu(90000), Emu(90000))
            dot.fill.solid()
            dot.fill.fore_color.rgb = ORANGE
            dot.line.fill.background()
            body_left = ML + Emu(900000)
        bbox = tb(slide, body_left, y, TEXT_W - Emu(900000), Emu(520000))
        p = bbox.text_frame.paragraphs[0]
        if highlight_first and i == 1 and " — " in item:
            label, text = item.split(" — ", 1)
            run(p, label + " — ", 12, False, ORANGE)
            run(p, text, 12, False, DARK)
        else:
            run(p, item, 12, False, DARK)
        y += Emu(520000)


def body_paragraph(slide, text, top=Emu(2200000), width=TEXT_W, size=12):
    box = tb(slide, ML + Emu(500000), top, width, Emu(4200000))
    run(box.text_frame.paragraphs[0], text, size, False, GRAY)


def highlight_box(slide, parts, top=Emu(2100000)):
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, ML + Emu(400000), top, TEXT_W, Emu(650000))
    box.fill.solid()
    box.fill.fore_color.rgb = WHITE
    box.line.color.rgb = RED
    box.line.width = Pt(1.5)
    tbox = tb(slide, ML + Emu(550000), top + Emu(120000), TEXT_W - Emu(200000), Emu(450000))
    p = tbox.text_frame.paragraphs[0]
    for text, bold, color in parts:
        run(p, text, 12, bold, color)


def title_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg(slide)
    left_blk = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(365760), Emu(1600000), Emu(6400800), Emu(3200000))
    left_blk.fill.solid()
    left_blk.fill.fore_color.rgb = LIGHT_GRAY
    left_blk.line.fill.background()
    right_blk = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(9144000), 0, Emu(3048000), H)
    right_blk.fill.solid()
    right_blk.fill.fore_color.rgb = LIGHT_GRAY
    right_blk.line.fill.background()

    t = tb(slide, Emu(640080), Emu(2200000), Emu(8000000), Emu(2000000))
    lines = ["УСТРОЙСТВО И НАЗНАЧЕНИЕ", "ТРУБОПРОВОДОВ ПАРА", "И ГОРЯЧЕЙ ВОДЫ"]
    for i, line in enumerate(lines):
        p = t.text_frame.paragraphs[0] if i == 0 else t.text_frame.add_paragraph()
        run(p, line, 24, True, DARK)

    ln = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(640080), Emu(4300000), Emu(7800000), Emu(36000))
    ln.fill.solid()
    ln.fill.fore_color.rgb = DARK
    ln.line.fill.background()
    dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, Emu(8400000), Emu(4270000), Emu(60000), Emu(60000))
    dot.fill.solid()
    dot.fill.fore_color.rgb = ORANGE
    dot.line.fill.background()


def content_slide(prs, num, header, intro, body=None, items=None, note_item=None,
                  image=None, image_bottom=False, highlight=None, subtitle=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg(slide)
    section_header(slide, header, subtitle if subtitle else None)
    slide_num(slide, num)

    y = Emu(1350000)
    if intro:
        ibox = tb(slide, ML + Emu(500000), y, TEXT_W, Emu(350000))
        p = ibox.text_frame.paragraphs[0]
        if note_item and ". " in intro:
            prefix, rest = intro.split(". ", 1)
            run(p, prefix + ". ", 12, False, DARK)
            run(p, note_item, 12, True, ORANGE)
            run(p, " " + rest, 12, False, DARK)
        else:
            run(p, intro, 12, False, DARK)
        y = Emu(1750000)

    top = Emu(2200000)
    if highlight:
        highlight_box(slide, highlight, Emu(2100000))
        top = Emu(2900000)

    if body:
        body_paragraph(slide, body, top)
    if items:
        bullets(slide, items, top, numbered=("состоит" in (intro or "")), highlight_first=bool(note_item))

    if image:
        if image_bottom:
            add_image_bottom(slide, image)
        else:
            add_image(slide, image)
    return slide


def main():
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H

    title_slide(prs)

    # 2 FNP order
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    section_header(s, "ФЕДЕРАЛЬНЫЕ НОРМЫ И ПРАВИЛА В ОБЛАСТИ ПРОМЫШЛЕННОЙ БЕЗОПАСНОСТИ")
    sub = tb(s, ML, Emu(1250000), TEXT_W, Emu(300000))
    run(sub.text_frame.paragraphs[0], "ПРИКАЗ РОСТЕХНАДЗОРА ОТ 15.12.2020 № 536", 12, True, GRAY)
    slide_num(s, 2, Emu(2800000))
    body = tb(s, Emu(900000), Emu(3000000), Emu(10400000), Emu(2500000), PP_ALIGN.CENTER)
    run(body.text_frame.paragraphs[0],
        "Об утверждении федеральных норм и правил в области промышленной "
        "безопасности «Правила промышленной безопасности при использовании "
        "оборудования, работающего под избыточным давлением» (далее — ФНП).",
        14, False, DARK)

    content_slide(
        prs, 3, "ОБЩИЕ ПОЛОЖЕНИЯ", intro=None,
        body="ФНП устанавливают требования промышленной безопасности при "
             "разработке, размещении, монтаже, наладке, эксплуатации, ремонте, "
             "реконструкции, техническом освидетельствовании и диагностировании "
             "оборудования, работающего под избыточным давлением.",
        image=f"{IMG}/industrial_steam_sidebar.png",
    )

    content_slide(
        prs, 4, "ОБЩИЕ ПОЛОЖЕНИЯ",
        "ФНП устанавливают требования промышленной безопасности, обязательные:",
        items=[
            "при разработке технологических процессов",
            "при техническом перевооружении опасного производственного объекта (далее — ОПО)",
            "при размещении",
            "при монтаже",
            "при ремонте",
            "при реконструкции (модернизации)",
            "при наладке и эксплуатации",
            "при техническом освидетельствовании",
            "при техническом диагностировании и экспертизе промышленной безопасности "
            "оборудования, работающего под избыточным давлением",
        ],
        image=f"{IMG}/pipeline_system_intro.png",
    )

    content_slide(
        prs, 5, "ОБЩИЕ ПОЛОЖЕНИЯ",
        "ФНП распространяются на следующее оборудование, работающее под избыточным давлением:",
        items=[
            "паровые и водогрейные котлы",
            "сосуды, работающие под давлением",
            "трубопроводы пара и горячей воды",
            "газовые баллоны",
            "резервуары для сжатых газов",
            "арматуру и предохранительные устройства",
        ],
        image=f"{IMG}/industrial_steam_sidebar.png",
    )

    content_slide(
        prs, 6, "УСТРОЙСТВО ТРУБОПРОВОДОВ", intro=None,
        subtitle="КОНСТРУКЦИЯ ТРУБОПРОВОДОВ ПАРА И ГОРЯЧЕЙ ВОДЫ",
        body="Трубопроводы пара и горячей воды предназначены для транспортировки "
             "теплоносителя от источника тепла к потребителям. Конструкция должна "
             "обеспечивать прочность, герметичность, компенсацию температурных "
             "деформаций и безопасную эксплуатацию.",
        image=f"{IMG}/pipeline_system_intro.png",
    )

    content_slide(
        prs, 7, "УСТРОЙСТВО ТРУБОПРОВОДОВ",
        "Трубопровод состоит из:",
        items=[
            "Плотно соединённых между собой прямых участков труб",
            "Фасонных деталей — отводы, переходники, тройники",
            "Крепёжных элементов — фланцы, болты, шпильки",
            "Арматуры — краны, вентили, задвижки, регулирующие клапаны",
            "Редуцирующих и предохранительных клапанов",
        ],
        image=f"{IMG}/pipeline_assembly_3d.png",
    )

    content_slide(
        prs, 8, "УСТРОЙСТВО ТРУБОПРОВОДОВ",
        "Трубопровод состоит из: 6. Приборы КИПиА — манометры, термометры, расходомеры, диафрагмы.",
        note_item="6.",
        body="На трубопроводах устанавливаются приборы контроля и автоматики для "
             "измерения давления, температуры и расхода теплоносителя.",
        image=f"{IMG}/kipia_instruments.png",
        image_bottom=True,
    )

    content_slide(
        prs, 9, "УСТРОЙСТВО ТРУБОПРОВОДОВ",
        "Трубопровод состоит из: 7. Опорно-подвесной системы",
        note_item="7.",
        items=[
            "неподвижные опоры — фиксируют положение трубопровода",
            "подвижные опоры — обеспечивают перемещение при температурных деформациях",
            "опоры скольжения, катковые, роликовые",
            "опоры подвесные, пружинные и жёсткие подвески",
        ],
        image=f"{IMG}/pipe_supports_3d.png",
    )

    content_slide(
        prs, 10, "УСТРОЙСТВО ТРУБОПРОВОДОВ",
        "Трубопровод состоит из: 8. Конденсатоотводчиков — дренажи, патрубки",
        note_item="8.",
        body="В нижних точках каждого отключаемого задвижками участка трубопровода "
             "должны предусматриваться спускные штуцера с запорной арматурой для "
             "опорожнения. В верхних точках устанавливаются воздушники. Нижние "
             "концевые точки паропроводов и изгибов снабжаются устройствами для "
             "продувки. Непрерывный отвод конденсата обязателен для паропроводов "
             "насыщенного пара и тупиковых участков паропроводов перегретого пара.",
        image=f"{IMG}/condensate_drainage.png",
    )

    content_slide(
        prs, 11, "УСТРОЙСТВО ТРУБОПРОВОДОВ",
        "Для компенсации удлинения трубопровода при нагреве:",
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
        image=f"{IMG}/compensators_7types.png",
    )

    content_slide(
        prs, 12, "УСТРОЙСТВО ТРУБОПРОВОДОВ",
        "Трубопровод состоит из: 10. Заглушек",
        note_item="10.",
        body="Заглушка — элемент трубопровода, предназначенный для глухого "
             "запечатывания его выходных отверстий; как правило используется при "
             "проведении пневмо- и гидроиспытаний. Металлические концевые заглушки "
             "различаются по виду исполнения и типу крепления.",
        items=[
            "плоские заглушки, закрепляемые сваркой",
            "фланцевые стальные (1)",
            "эллиптические (2)",
            "сферические",
            "быстросъёмные (3)",
            "межфланцевые (4)",
        ],
        image=f"{IMG}/pipe_caps_4types.png",
    )

    content_slide(
        prs, 13, "УСТРОЙСТВО ТРУБОПРОВОДОВ",
        "Трубопровод состоит из: 11. Теплоизоляции",
        note_item="11.",
        body="Теплоизоляция трубопроводов предназначена для снижения теплопотерь, "
             "предотвращения конденсации влаги на поверхности, защиты персонала от "
             "термических ожогов и обеспечения стабильности технологических "
             "параметров. Применяются минераловатные и другие материалы с защитной оболочкой.",
        image=f"{IMG}/thermal_insulation_cutaway.png",
    )

    content_slide(
        prs, 14, "УСТРОЙСТВО ТРУБОПРОВОДОВ",
        "Трубопровод состоит из: 12. Опознавательной окраски",
        note_item="12.",
        body="Коммуникации делятся на 10 основных групп по транспортируемым веществам, "
             "что требует их идентификации и маркировки.",
        items=[
            "Цветовая градация при разметке трубопроводов (ГОСТ 14202-69)",
            "зелёный цвет соответствует 1 группе, транспортирует воду",
            "красный цвет соответствует 2 группе, транспортирует пар",
            "предупреждающие кольца: красные — легковоспламеняющиеся и взрывоопасные",
            "жёлтые — токсичные и радиоактивные вещества",
            "зелёные с белой каймой — внутренняя безопасность",
        ],
        image=f"{IMG}/pipe_identification_gost.png",
    )

    prs.save(OUT)
    print(f"Saved {OUT} ({len(prs.slides)} slides)")


if __name__ == "__main__":
    main()
