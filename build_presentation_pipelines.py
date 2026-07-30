#!/usr/bin/env python3
"""Build pipeline presentation in red Inter style with images."""

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Pt

OUT = "/workspace/Презентация_Трубопроводы.pptx"
IMG = "/workspace/assets/pipeline_images/generated"

W = Emu(12191695)
H = Emu(6858000)
ML = Emu(548640)
CW = Emu(11094415)
TEXT_W = Emu(5800000)
IMG_LEFT = Emu(6600000)
IMG_W = Emu(5200000)
TOTAL = 14

RED = RGBColor(0xE3, 0x06, 0x13)
DARK = RGBColor(0x1A, 0x1A, 0x1A)
GRAY = RGBColor(0x6B, 0x72, 0x80)
AMBER = RGBColor(0xF5, 0x9E, 0x0B)
LIGHT_BG = RGBColor(0xF5, 0xF8, 0xFC)
FONT = "Inter"


def run(p, text, size=13, bold=False, color=DARK, italic=False):
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
    rect.fill.fore_color.rgb = LIGHT_BG
    rect.line.fill.background()


def footer_line(slide):
    ln = slide.shapes.add_connector(1, ML, Emu(6355080), ML + CW, Emu(6355080))
    ln.line.color.rgb = RED
    ln.line.width = Pt(1.5)


def page_num(slide, num):
    box = tb(slide, Emu(10271455), Emu(6492240), Emu(1188720), Emu(228600), PP_ALIGN.RIGHT)
    run(box.text_frame.paragraphs[0], f"{num:02d} / {TOTAL:02d}", 9, True, RED)


def header(slide, section, title, intro=None):
    s = tb(slide, ML, Emu(548640), Emu(9000000), Emu(228600))
    run(s.text_frame.paragraphs[0], section, 11, True, RED)
    t = tb(slide, ML, Emu(850000), CW, Emu(700000))
    run(t.text_frame.paragraphs[0], title, 28, True, DARK)
    if intro:
        i = tb(slide, ML, Emu(1600000), TEXT_W, Emu(400000))
        run(i.text_frame.paragraphs[0], intro, 12, False, GRAY)


def add_image(slide, path, left=IMG_LEFT, top=Emu(1200000), width=IMG_W, height=Emu(4800000)):
    slide.shapes.add_picture(path, left, top, width=width, height=height)


def add_image_bottom(slide, path, top=Emu(3600000), height=Emu(2500000)):
    slide.shapes.add_picture(path, Emu(731520), top, width=Emu(10700000), height=height)


def numbered_list(slide, items, top=Emu(2150000)):
    y = top
    for idx, item in enumerate(items, 1):
        n = tb(slide, ML, y, Emu(411480), Emu(274320))
        run(n.text_frame.paragraphs[0], f"{idx:02d}", 14, True, RED)
        b = tb(slide, Emu(1005840), y, Emu(5200000), Emu(520000))
        run(b.text_frame.paragraphs[0], item, 13, False, DARK)
        y += Emu(580000)


def body_text(slide, text, top=Emu(2150000), width=TEXT_W):
    box = tb(slide, ML, top, width, Emu(3600000))
    run(box.text_frame.paragraphs[0], text, 13, False, DARK)


def note(slide, text, top=Emu(5852160)):
    box = tb(slide, Emu(731520), top, Emu(10728655), Emu(365760))
    run(box.text_frame.paragraphs[0], text, 14, True, AMBER)


def highlight_box(slide, parts, top=Emu(2100000)):
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, ML, top, TEXT_W, Emu(650000))
    box.fill.solid()
    box.fill.fore_color.rgb = LIGHT_BG
    box.line.color.rgb = RED
    box.line.width = Pt(1.5)
    tbox = tb(slide, Emu(700000), top + Emu(120000), TEXT_W - Emu(150000), Emu(450000))
    p = tbox.text_frame.paragraphs[0]
    for text, bold, color in parts:
        run(p, text, 13, bold, color)


def finish_slide(slide, num, note_text=None):
    if note_text:
        note(slide, note_text)
    footer_line(slide)
    page_num(slide, num)


def title_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg(slide)
    s = tb(slide, ML, Emu(1800000), CW, Emu(228600))
    run(s.text_frame.paragraphs[0], "Промышленная безопасность", 14, True, RED)
    t = tb(slide, ML, Emu(2300000), CW, Emu(900000))
    run(t.text_frame.paragraphs[0], "Устройство и назначение трубопроводов пара и горячей воды", 36, True, DARK)
    sub = tb(slide, ML, Emu(3400000), CW, Emu(700000))
    run(sub.text_frame.paragraphs[0],
        "ФНП при использовании оборудования, работающего под избыточным давлением",
        12, False, GRAY)
    finish_slide(slide, 1)


def content_slide(prs, num, section, title, intro=None, body=None, items=None,
                  image=None, image_bottom=False, highlight=None, note_text=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg(slide)
    header(slide, section, title, intro)

    top = Emu(2150000)
    if highlight:
        highlight_box(slide, highlight, Emu(2100000))
        top = Emu(2900000)

    if body:
        body_text(slide, body, top)
    if items:
        numbered_list(slide, items, top)

    if image:
        if image_bottom:
            add_image_bottom(slide, image)
        else:
            add_image(slide, image)

    finish_slide(slide, num, note_text)
    return slide


def main():
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H

    title_slide(prs)

    # 2 FNP order
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "Нормативная база", "Федеральные нормы и правила",
         "Приказ Ростехнадзора от 15.12.2020 № 536")
    body = tb(s, ML, Emu(2800000), CW, Emu(2800000))
    run(body.text_frame.paragraphs[0],
        "Об утверждении федеральных норм и правил в области промышленной "
        "безопасности «Правила промышленной безопасности при использовании "
        "оборудования, работающего под избыточным давлением» (далее — ФНП).",
        14, False, DARK)
    finish_slide(s, 2)

    content_slide(
        prs, 3, "Общие положения", "Назначение ФНП",
        intro="Требования промышленной безопасности при эксплуатации оборудования под давлением",
        body="ФНП устанавливают требования промышленной безопасности при разработке, "
             "размещении, монтаже, наладке, эксплуатации, ремонте, реконструкции, "
             "техническом освидетельствовании и диагностировании оборудования, "
             "работающего под избыточным давлением.",
        image=f"{IMG}/industrial_steam_sidebar.png",
    )

    content_slide(
        prs, 4, "Общие положения", "Область применения ФНП",
        intro="Требования обязательны при следующих видах работ",
        items=[
            "при разработке технологических процессов",
            "при техническом перевооружении опасного производственного объекта (ОПО)",
            "при размещении, монтаже и ремонте",
            "при реконструкции (модернизации) оборудования",
            "при наладке и эксплуатации",
            "при техническом освидетельствовании",
            "при техническом диагностировании и экспертизе промышленной безопасности",
        ],
        image=f"{IMG}/pipeline_system_intro.png",
    )

    content_slide(
        prs, 5, "Общие положения", "К какому оборудованию применяются ФНП",
        intro="Оборудование, работающее под избыточным давлением",
        items=[
            "паровые и водогрейные котлы",
            "сосуды, работающие под давлением",
            "трубопроводы пара и горячей воды",
            "газовые баллоны и резервуары для сжатых газов",
            "арматуру и предохранительные устройства",
        ],
        note_text="ФНП обязательны для всех стадий жизненного цикла оборудования под давлением.",
        image=f"{IMG}/industrial_steam_sidebar.png",
    )

    content_slide(
        prs, 6, "Устройство трубопроводов", "Конструкция трубопроводов пара и горячей воды",
        intro="Назначение и общие требования к устройству",
        body="Трубопроводы пара и горячей воды предназначены для транспортировки "
             "теплоносителя от источника тепла к потребителям. Конструкция должна "
             "обеспечивать прочность, герметичность, компенсацию температурных "
             "деформаций и безопасную эксплуатацию.",
        image=f"{IMG}/pipeline_system_intro.png",
    )

    content_slide(
        prs, 7, "Устройство трубопроводов", "Состав трубопровода (1/2)",
        intro="Трубопровод состоит из основных элементов",
        items=[
            "плотно соединённых между собой прямых участков труб",
            "фасонных деталей — отводы, переходники, тройники",
            "крепёжных элементов — фланцы, болты, шпильки",
            "арматуры — краны, вентили, задвижки, регулирующие клапаны",
            "редуцирующих и предохранительных клапанов",
        ],
        image=f"{IMG}/pipeline_assembly_3d.png",
    )

    content_slide(
        prs, 8, "Устройство трубопроводов", "Приборы КИПиА",
        intro="Контрольно-измерительные приборы и автоматика",
        items=[
            "манометры — контроль давления в трубопроводе",
            "термометры — контроль температуры теплоносителя",
            "расходомеры — учёт и контроль расхода среды",
            "диафрагмы — измерение расхода по перепаду давления",
        ],
        note_text="Приборы КИПиА обеспечивают контроль параметров и безопасную эксплуатацию.",
        image=f"{IMG}/kipia_instruments.png",
        image_bottom=True,
    )

    content_slide(
        prs, 9, "Устройство трубопроводов", "Опорно-подвесная система",
        intro="Крепление и поддержание трубопровода",
        items=[
            "неподвижные опоры — фиксируют положение трубопровода",
            "подвижные опоры — допускают перемещение при температурных деформациях",
            "скользящие, катковые и подвесные опоры",
            "пружинные и жёсткие подвески",
            "направляющие и ограничители перемещения",
        ],
        image=f"{IMG}/pipe_supports_3d.png",
    )

    content_slide(
        prs, 10, "Устройство трубопроводов", "Конденсатоотводчики, дренажи и патрубки",
        intro="Устройства для отвода конденсата и воздуха",
        body="В нижних точках каждого отключаемого задвижками участка трубопровода "
             "должны предусматриваться спускные штуцера с запорной арматурой для "
             "опорожнения. В верхних точках устанавливаются воздушники. Нижние "
             "концевые точки паропроводов и изгибов снабжаются устройствами для "
             "продувки. Непрерывный отвод конденсата обязателен для паропроводов "
             "насыщенного пара и тупиковых участков паропроводов перегретого пара.",
        image=f"{IMG}/condensate_drainage.png",
    )

    content_slide(
        prs, 11, "Устройство трубопроводов", "Компенсаторы температурных деформаций",
        intro="Для компенсации удлинения трубопровода при нагреве",
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
        prs, 12, "Устройство трубопроводов", "Заглушки",
        intro="Элементы для глухого запечатывания выходных отверстий",
        body="Заглушка — элемент трубопровода, предназначенный для глухого "
             "запечатывания его выходных отверстий; как правило используется при "
             "проведении пневмо- и гидроиспытаний.",
        note_text="Типы: плоские сварные, фланцевые, эллиптические, сферические, "
                  "быстросъёмные, межфланцевые.",
        image=f"{IMG}/pipe_caps_4types.png",
    )

    content_slide(
        prs, 13, "Устройство трубопроводов", "Теплоизоляция",
        intro="Защита от потерь тепла и ожогов",
        body="Теплоизоляция трубопроводов предназначена для снижения теплопотерь, "
             "предотвращения конденсации влаги на поверхности, защиты персонала от "
             "термических ожогов и обеспечения стабильности технологических параметров.",
        image=f"{IMG}/thermal_insulation_cutaway.png",
    )

    content_slide(
        prs, 14, "Устройство трубопроводов", "Опознавательная окраска",
        intro="Маркировка трубопроводов по ГОСТ 14202-69",
        items=[
            "коммуникации делятся на 10 основных групп по транспортируемым веществам",
            "зелёный цвет — 1 группа, транспортирует воду",
            "красный цвет — 2 группа, транспортирует пар",
            "предупреждающие кольца: красные — легковоспламеняющиеся и взрывоопасные",
            "жёлтые — токсичные и радиоактивные вещества",
            "зелёные с белой каймой — внутренняя безопасность",
        ],
        note_text="Количество предупреждающих колец зависит от давления и температуры среды.",
        image=f"{IMG}/pipe_identification_gost.png",
    )

    prs.save(OUT)
    print(f"Saved {OUT} ({len(prs.slides)} slides)")


if __name__ == "__main__":
    main()
