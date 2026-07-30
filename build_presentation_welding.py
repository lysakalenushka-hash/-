#!/usr/bin/env python3
"""Build welding works requirements presentation from 6.docx."""

from copy import deepcopy

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Emu, Pt

TEMPLATE = "/home/ubuntu/.cursor/projects/workspace/uploads/____________54c7.pptx"
OUT = "/workspace/Презентация_Сварочные_работы.pptx"
IMG = "/workspace/assets/operation_images"

RED = RGBColor(0xE3, 0x06, 0x13)
DARK = RGBColor(0x1A, 0x1A, 0x1A)
GRAY = RGBColor(0x6B, 0x72, 0x80)
AMBER = RGBColor(0xF5, 0x9E, 0x0B)

IMG_LEFT = Emu(6600000)
IMG_W = Emu(5200000)


def set_run(run, size_pt, bold=False, color=DARK):
    run.font.name = "Inter"
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.color.rgb = color


def set_text(shape, text, size_pt=13, bold=False, color=DARK):
    tf = shape.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = text
    set_run(r, size_pt, bold=bold, color=color)


def text_shapes(slide):
    return [s for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip()]


def clear_slides(prs):
    while len(prs.slides) > 0:
        r_id = prs.slides._sldIdLst[0].rId
        prs.part.drop_rel(r_id)
        del prs.slides._sldIdLst[0]


def make_slide(prs, shape_elements):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    for shape in list(slide.shapes):
        shape.element.getparent().remove(shape.element)
    for element in shape_elements:
        slide.shapes._spTree.insert_element_before(deepcopy(element), "p:extLst")
    return slide


def update_pages(prs):
    total = len(prs.slides)
    for i, slide in enumerate(prs.slides, 1):
        for sh in text_shapes(slide):
            t = sh.text_frame.text.strip()
            if " / " in t:
                left, right = [x.strip() for x in t.split(" / ", 1)]
                if left.isdigit() and right.isdigit():
                    set_text(sh, f"{i:02d} / {total:02d}", 9, True, RED)


def fill_text_slide(slide, section, title, intro, body, note=None):
    shapes = sorted(text_shapes(slide), key=lambda s: (s.top, s.left))
    set_text(shapes[0], section, 11, True, RED)
    set_text(shapes[1], title, 28, True, DARK)
    set_text(shapes[2], intro, 12, False, GRAY)
    set_text(shapes[3], body, 13, False, DARK)
    if len(shapes) > 4:
        set_text(shapes[4], "", 11, False, GRAY)
    if note:
        for sh in shapes:
            if sh.top > Emu(5800000) and sh.left < Emu(5000000):
                set_text(sh, note, 13, True, AMBER)
                break


def fill_list_slide(slide, section, title, intro, list_title, items, note=None):
    shapes = sorted(text_shapes(slide), key=lambda s: (s.top, s.left))
    set_text(shapes[0], section, 11, True, RED)
    set_text(shapes[1], title, 28, True, DARK)
    set_text(shapes[2], intro, 12, False, GRAY)
    set_text(shapes[3], list_title, 11, True, RED)

    pairs = []
    more_shape = None
    for sh in shapes[5:]:
        t = sh.text_frame.text.strip()
        if t in {f"{i:02d}" for i in range(1, 10)}:
            pairs.append([sh, None])
        elif pairs and pairs[-1][1] is None:
            pairs[-1][1] = sh
        elif t.startswith("…"):
            more_shape = sh

    for i, item in enumerate(items[:4]):
        if i < len(pairs):
            set_text(pairs[i][0], f"{i+1:02d}", 14, True, RED)
            set_text(pairs[i][1], item, 13, False, DARK)

    for i in range(len(items[:4]), len(pairs)):
        set_text(pairs[i][0], "", 14, True, RED)
        set_text(pairs[i][1], "", 14, False, DARK)

    if more_shape:
        rest = max(0, len(items) - 4)
        set_text(more_shape, f"… ещё {rest}" if rest else "", 9, False, GRAY)

    if note:
        for sh in shapes:
            if sh.top > Emu(5900000) and sh.left > Emu(700000):
                set_text(sh, note, 13, True, AMBER)
                break


def add_image(slide, path, bottom=False):
    if bottom:
        slide.shapes.add_picture(path, Emu(731520), Emu(3600000), width=Emu(10700000), height=Emu(2500000))
    else:
        slide.shapes.add_picture(path, IMG_LEFT, Emu(1200000), width=IMG_W, height=Emu(4800000))


def main():
    tpl = Presentation(TEMPLATE)
    text_tpl = [deepcopy(s.element) for s in tpl.slides[0].shapes]
    list_tpl = [deepcopy(s.element) for s in tpl.slides[2].shapes]

    prs = Presentation(TEMPLATE)
    clear_slides(prs)

    # 1. Title
    s = make_slide(prs, text_tpl)
    fill_text_slide(
        s,
        "ФНП · Сварочное производство",
        "Общие требования к производству сварочных работ на ОПО",
        "Персонал, организация работ, ПТД, контроль качества",
        "Модуль охватывает определение сварочных работ, требования к персоналу "
        "и аттестации, проверку готовности технологий, содержание "
        "производственно-технологической документации, организацию работ, "
        "входной, операционный и приёмочный контроль.",
        note="Сварочные работы на ОПО — только по аттестованным технологиям.",
    )
    add_image(s, f"{IMG}/pipeline_system_intro.png")

    # 2. Definition and personnel duties
    s = make_slide(prs, list_tpl)
    fill_list_slide(
        s,
        "Общие требования",
        "Сварочные работы и персонал сварочного производства",
        "Персонал (сварщики, операторы, специалисты, контролёры) должен обеспечивать:",
        "Обязанности персонала",
        [
            "техническую и технологическую подготовку и выполнение сварочных работ",
            "безопасную эксплуатацию, обслуживание и ремонт сварочного оборудования",
            "соблюдение технологий сварки",
            "контроль качества сварных соединений",
        ],
        note="Производство сварочных работ — деятельность с применением сварочных процессов, материалов и оборудования.",
    )

    # 3. Certification
    s = make_slide(prs, list_tpl)
    fill_list_slide(
        s,
        "Аттестация",
        "Аттестация сварщиков и технологий сварки",
        "Допуск к сварочным работам на ОПО",
        "Требования к аттестации",
        [
            "квалификация должна соответствовать видам работ и технологиям сварки",
            "аттестация по способам сварки, видам конструкций, положениям, материалам",
            "допуск — по положительным результатам аттестационных испытаний",
            "организация должна пройти проверку готовности к применению аттестованных технологий",
            "контрольные сварные соединения выполняют на месте производства работ",
            "положительные результаты оформляются документом с областью допуска",
        ],
    )
    add_image(s, f"{IMG}/industrial_steam_sidebar.png")

    # 4. Organization of welding works
    s = make_slide(prs, list_tpl)
    fill_list_slide(
        s,
        "Организация работ",
        "Организация сварочных работ",
        "Руководство и производственно-технологическая документация",
        "Ключевые требования",
        [
            "организацию обеспечивает руководитель организации или уполномоченное лицо",
            "работы выполняются в соответствии с ПТД, утверждённой руководителем",
            "ПТД разрабатывается специалистом сварочного производства на основе проекта и НД",
            "ПТД включает технологические инструкции и технологические карты сварки",
        ],
        note="Аттестационные процедуры обеспечивает руководитель независимого аттестационного центра.",
    )

    # 5. PTD content
    s = make_slide(prs, list_tpl)
    fill_list_slide(
        s,
        "ПТД",
        "Содержание производственно-технологической документации",
        "В ПТД применительно к выполняемым работам устанавливают:",
        "Что должно быть в ПТД",
        [
            "способы сварки, режимы, пространственные положения",
            "требования к квалификации и допускным испытаниям сварщиков",
            "требования к сборке, прихваткам и временным креплениям",
            "сварочные материалы, оборудование, род и полярность тока",
            "подогрев, защита зоны сварки, маркировка соединений",
            "методы и объёмы НК, требования к исправлению дефектов",
        ],
    )

    # 6. Assembly requirements
    s = make_slide(prs, list_tpl)
    fill_list_slide(
        s,
        "Сборка",
        "Требования к сборке деталей под сварку",
        "В требованиях по сборке в ПТД должны быть приведены:",
        "Параметры сборки",
        [
            "способы подготовки поверхностей деталей под сварку",
            "приспособления, оборудование, порядок и последовательность сборки",
            "способы сварки, материалы и режимы при выполнении прихваток",
            "размеры, количество и расположение прихваток",
            "методы контроля качества сборки",
            "стыковые кольцевые швы — соосное позиционирование и фиксация",
        ],
    )
    add_image(s, f"{IMG}/pipeline_system_intro.png")

    # 7. Equipment and supervisor duties
    s = make_slide(prs, list_tpl)
    fill_list_slide(
        s,
        "Оборудование и руководство",
        "Сварочное оборудование и обязанности руководителя работ",
        "Перед выполнением сварочных работ руководитель обязан:",
        "Подготовка к работам",
        [
            "проверить состав и квалификацию персонала, оборудование и материалы",
            "ознакомить сварщиков с технологическими картами под подпись",
            "организовать проведение операционного контроля",
            "оборудование и материалы должны соответствовать аттестованным технологиям",
            "место сварки — защита от осадков, влаги, сквозняков",
            "допускные соединения — при первом допуске или после длительного перерыва",
        ],
        note="Сварочное оборудование содержат в исправном состоянии по указаниям производителя.",
    )

    # 8. Types of control
    s = make_slide(prs, list_tpl)
    fill_list_slide(
        s,
        "Контроль",
        "Виды контроля сварочных работ",
        "При подготовке и выполнении сварочных работ осуществляют:",
        "Виды контроля",
        [
            "входной контроль — партии свариваемых и сварочных материалов до применения",
            "операционный контроль — подготовка кромок, сборка, прихватка, сварка",
            "приёмочный контроль — все выполненные сварные соединения",
            "в процессе сварки — режимы, очерёдность швов, отсутствие видимых дефектов",
        ],
        note="При обнаружении трещин или недопустимых дефектов работы останавливают.",
    )

    # 9. Incoming and operational control details
    s = make_slide(prs, list_tpl)
    fill_list_slide(
        s,
        "Контроль",
        "Входной и операционный контроль",
        "Что проверяется на этапах контроля",
        "Основные проверки",
        [
            "входной: сертификаты качества, маркировка, целостность упаковки",
            "электроды, проволока, лента, флюс — размеры и состояние поверхности",
            "операционный: разделка кромок, фиксация, зазоры, смещение кромок",
            "размеры и расположение прихваток, чистота поверхностей",
        ],
    )
    add_image(s, f"{IMG}/kipia_instruments.png", bottom=True)

    # 10. Documentation and marking
    s = make_slide(prs, list_tpl)
    fill_list_slide(
        s,
        "Документация",
        "Маркировка, исправление дефектов и оформление",
        "Руководитель сварочных работ обеспечивает:",
        "Документирование и качество",
        [
            "маркировка швов — шифр клейма сварщика для однозначной идентификации",
            "исправление дефектов — по ПТД; число исправлений не выше указанного",
            "идентификация материалов, оборудования и мест сварных соединений",
            "регистрация сведений о сварщиках и результатов контроля качества",
            "оформление журналов сварочных работ, актов и заключений по НК",
            "протоколы испытаний сварных соединений",
        ],
        note="Исполнительная и эксплуатационная документация оформляется в процессе работ.",
    )

    update_pages(prs)
    prs.save(OUT)
    print(f"Saved {OUT} ({len(prs.slides)} slides)")


if __name__ == "__main__":
    main()
