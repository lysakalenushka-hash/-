#!/usr/bin/env python3
"""Update Tema 3 presentation: prune extras, add missing from 3.docx, fix numbering."""

from copy import deepcopy

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Emu, Pt

SRC = "/home/ubuntu/.cursor/projects/workspace/uploads/____________e977.pptx"
TMP = "/tmp/tema3_step1.pptx"
OUT = "/workspace/Презентация_Тема_3_Персонал.pptx"
IMG = "/workspace/assets/personnel_images"

RED = RGBColor(0xE3, 0x06, 0x13)
DARK = RGBColor(0x1A, 0x1A, 0x1A)
GRAY = RGBColor(0x6B, 0x72, 0x80)
AMBER = RGBColor(0xF5, 0x9E, 0x0B)

# 0-based: quizzes 11-13,25; dups 4,14,33
DELETE_INDICES = {3, 10, 11, 12, 13, 24, 32}


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


def delete_slides(prs, indices):
    sldIdLst = prs.slides._sldIdLst
    for idx in sorted(indices, reverse=True):
        sldId = sldIdLst[idx]
        rId = sldId.get(
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
        )
        prs.part.drop_rel(rId)
        sldIdLst.remove(sldId)


def update_pages(prs):
    total = len(prs.slides)
    for i, slide in enumerate(prs.slides, 1):
        for sh in slide.shapes:
            if not sh.has_text_frame:
                continue
            if sh.top > Emu(6200000) and sh.left > Emu(8000000):
                set_text(sh, f"{i:02d} / {total:02d}", 9, True, RED)
                break
            t = sh.text_frame.text.strip()
            if " / " in t:
                left, right = [x.strip() for x in t.split(" / ", 1)]
                if left.isdigit() and right.isdigit():
                    set_text(sh, f"{i:02d} / {total:02d}", 9, True, RED)
                    break


def fill_text_slide(slide, section, title, intro, body, note=None):
    shapes = sorted(text_shapes(slide), key=lambda s: (s.top, s.left))
    content = [s for s in shapes if not (s.top > Emu(6200000) and s.left > Emu(8000000))]
    set_text(content[0], section, 11, True, RED)
    set_text(content[1], title, 26, True, DARK)
    set_text(content[2], intro, 12, False, GRAY)
    set_text(content[3], body, 13, False, DARK)
    if len(content) > 4 and content[4].top < Emu(6200000):
        set_text(content[4], note or "", 12, True, AMBER if note else GRAY)


def fill_list_slide(slide, section, title, intro, list_title, items, note=None):
    shapes = sorted(text_shapes(slide), key=lambda s: (s.top, s.left))
    content = [s for s in shapes if not (s.top > Emu(6200000) and s.left > Emu(8000000))]
    set_text(content[0], section, 11, True, RED)
    set_text(content[1], title, 26, True, DARK)
    set_text(content[2], intro, 12, False, GRAY)
    set_text(content[3], list_title, 11, True, RED)

    pairs, more_shape, note_shape = [], None, None
    for sh in content[4:]:
        t = sh.text_frame.text.strip()
        if t in {f"{i:02d}" for i in range(1, 10)} or t in {str(i) for i in range(1, 10)}:
            pairs.append([sh, None])
        elif pairs and pairs[-1][1] is None and sh.left > Emu(700000):
            pairs[-1][1] = sh
        elif t.startswith("…"):
            more_shape = sh
        elif sh.top > Emu(5600000) and sh.left < Emu(9000000):
            note_shape = sh

    for i, item in enumerate(items[:4]):
        if i < len(pairs):
            set_text(pairs[i][0], f"{i+1:02d}", 14, True, RED)
            if pairs[i][1] is not None:
                set_text(pairs[i][1], item, 12, False, DARK)
    for i in range(len(items[:4]), len(pairs)):
        set_text(pairs[i][0], "", 14, True, RED)
        if pairs[i][1] is not None:
            set_text(pairs[i][1], "", 12, False, DARK)
    if more_shape:
        rest = max(0, len(items) - 4)
        set_text(more_shape, f"… ещё {rest}" if rest else "", 9, False, GRAY)
    if note and note_shape:
        set_text(note_shape, note, 12, True, AMBER)


def add_image(slide, path, bottom=False):
    if bottom:
        slide.shapes.add_picture(path, Emu(731520), Emu(3800000), width=Emu(10700000), height=Emu(2300000))
    else:
        slide.shapes.add_picture(path, Emu(6600000), Emu(1200000), width=Emu(5000000), height=Emu(4500000))


def find_templates(prs):
    text_tpl = list_tpl = None
    for slide in prs.slides:
        n = sum(1 for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip())
        if text_tpl is None and 5 <= n <= 7:
            text_tpl = slide
        if list_tpl is None and n >= 10:
            texts = [s.text_frame.text.strip() for s in slide.shapes if s.has_text_frame]
            if any(t in {"01", "02", "1", "2"} for t in texts):
                list_tpl = slide
        if text_tpl and list_tpl:
            break
    return text_tpl, list_tpl


def clone_slide(prs, elems):
    blank = prs.slides.add_slide(prs.slide_layouts[6])
    for shape in list(blank.shapes):
        shape.element.getparent().remove(shape.element)
    for el in elems:
        blank.shapes._spTree.insert_element_before(deepcopy(el), "p:extLst")
    return blank


def append_missing(prs):
    text_tpl, list_tpl = find_templates(prs)
    text_elems = [deepcopy(s.element) for s in text_tpl.shapes]
    list_elems = [deepcopy(s.element) for s in list_tpl.shapes]

    def nt():
        return clone_slide(prs, text_elems)

    def nl():
        return clone_slide(prs, list_elems)

    # 3.2 Forms of work with personnel
    s = nl()
    fill_list_slide(
        s,
        "3.2 · Работа с персоналом",
        "Обязательные формы работы с персоналом",
        "Приказ Минэнерго № 796 — формы по категориям персонала",
        "Формы работы",
        [
            "АТП: предэкзаменационная подготовка, проверка знаний, производственный инструктаж",
            "диспетчерский / оперативный / ОРп: стажировка, новая должность, проверка знаний, дублирование, тренировки, спецподготовка, инструктаж",
            "ремонтный: стажировка, новая должность, проверка знаний, производственный инструктаж",
            "вспомогательный: предэкзаменационная подготовка и проверка знаний",
            "персонал РЗА — дополнительно подготовка и допуск к ТО устройств РЗА",
            "в организации — порядок работы с персоналом, утверждённый руководителем",
        ],
        note="Требования не распространяются на потребителей — физических лиц.",
    )

    s = nl()
    fill_list_slide(
        s,
        "3.2 · Специальная подготовка",
        "Специальная подготовка диспетчерского и оперативного персонала",
        "Не менее 5% и не более 20% рабочего времени с отрывом от основных обязанностей",
        "Объём специальной подготовки",
        [
            "учебные противоаварийные и противопожарные тренировки",
            "теоретические занятия по устройству и эксплуатации оборудования",
            "плановый производственный инструктаж",
            "дополнительно: изменения схем, аварийность, РЗА, ошибки по тренировкам",
            "отражается в графике оперативных дежурств (сменности)",
            "помещения для работы с персоналом: кабинеты, полигоны, тренажёры, библиотека",
        ],
    )
    add_image(s, f"{IMG}/switchgear_room.png", bottom=True)

    # 3.3 New position
    s = nl()
    fill_list_slide(
        s,
        "3.3 · Новая должность",
        "Подготовка по новой должности (рабочему месту)",
        "Обязательна для диспетчерского, оперативного, ОРп и ремонтного персонала",
        "Последовательность этапов",
        [
            "теоретическая подготовка (самоподготовка / иные способы)",
            "стажировка (для диспетчеров — также ознакомление с объектами)",
            "предэкзаменационная подготовка и проверка знаний",
            "дублирование (диспетчерский, оперативный, ОРп)",
            "противоаварийные и противопожарные тренировки",
            "допуск к самостоятельной работе",
        ],
        note="Аттестация по безопасности — до допуска к дублированию.",
    )

    # 3.4 Internship
    s = nl()
    fill_list_slide(
        s,
        "3.4 · Стажировка",
        "Стажировка в организациях",
        "Для диспетчерского, оперативного, ОРп и ремонтного персонала",
        "Ключевые требования",
        [
            "проводится под руководством назначенного ответственного за стажировку",
            "минимальный срок — 2 рабочих дня (смены) на каждом рабочем месте",
            "максимальный срок — не более 14 рабочих дней (смен)",
            "допуск к стажировке — ОРД с указанием сроков и ответственных",
            "ознакомление диспетчеров с объектами — не менее 1 рабочего дня",
            "уведомление о ознакомлении — не позднее чем за 3 рабочих дня",
        ],
    )
    add_image(s, f"{IMG}/electrical.png", bottom=True)

    # 3.5 Knowledge check details
    s = nl()
    fill_list_slide(
        s,
        "3.5 · Группы I–V",
        "Группы по электробезопасности (допуск)",
        "Пять классов допуска по опыту и правам работы",
        "Группы",
        [
            "I группа — неэлектротехнический персонал (запись в журнале)",
            "II группа — начинающие специалисты, переносной инструмент, водители в ЭУ",
            "III группа — единоличная работа до 1000 В; в бригаде — и выше",
            "IV группа — обслуживание > 1000 В; старший смены / ответственный до 1000 В",
            "V группа — ответственность за электроустановки > 1000 В",
        ],
    )

    s = nl()
    fill_list_slide(
        s,
        "3.5 · Проверка знаний",
        "Периодичность и основания внеочередной проверки",
        "По Приказу Минэнерго № 796",
        "Сроки и основания",
        [
            "очередная: диспетчерский/оперативный/ОРп и др. — не реже 1 раза в 12 месяцев",
            "очередная: иные работники — не реже 1 раза в 3 года",
            "первичная АТП/вспомогательного — не позднее 1 месяца после назначения",
            "внеочередная: новые НПА, перевод, новое оборудование, нарушения, аварии",
            "внеочередная: предписание надзора; перерыв в работе > 6 месяцев",
            "ознакомление с графиком — в течение месяца, но не позднее чем за 14 дней",
        ],
        note="Допуск к самостоятельной работе без очередной проверки знаний запрещён.",
    )

    s = nl()
    fill_list_slide(
        s,
        "3.5 · Комиссия и оформление",
        "Комиссия по проверке знаний и оформление результатов",
        "Постоянно действующие комиссии организации и филиалов",
        "Требования",
        [
            "комиссия — не менее 5 человек (председатель и заместитель)",
            "при проверке присутствует не менее 3 членов, включая председателя (зам.)",
            "при присвоении группы ЭБ — ≥ 3 с группой ЭБ, один не ниже присваиваемой",
            "неудовлетворительно — > 30% неверных ответов",
            "протокол, журнал учёта, удостоверение; автоэкзаменатор допускается",
            "повторная проверка при «неуд.» — не более 1 месяца; иначе — по ТК РФ",
        ],
    )
    add_image(s, f"{IMG}/kipia_instruments.png", bottom=True)

    # 3.6 Duplication
    s = nl()
    fill_list_slide(
        s,
        "3.6 · Дублирование",
        "Дублирование диспетчерского и оперативного персонала",
        "Этап подготовки по новой должности и после перерыва в работе",
        "Сроки и условия",
        [
            "при новой должности — после проверки знаний, не менее 12 рабочих смен",
            "после перерыва > 30 и < 60 дней — по порядку организации, ≥ 1 смена",
            "после перерыва от 60 дней до 6 месяцев — дублирование обязательно",
            "в период дублирования — контрольные противоаварийная и противопожарная тренировки",
            "продление дублирования допускается, но не более основной продолжительности",
            "ответственность несут и дублёр, и руководящий работник в равной мере",
        ],
    )

    # 3.7 Independent work
    s = nl()
    fill_list_slide(
        s,
        "3.7 · Допуск к самостоятельной работе",
        "Допуск к самостоятельной работе и его отзыв",
        "Для диспетчерского, оперативного, ОРп и ремонтного персонала",
        "Ключевые правила",
        [
            "допуск после этапов индивидуальной программы подготовки по новой должности",
            "действует до срока очередной проверки знаний",
            "повторный допуск после очередной проверки не требуется",
            "отзыв: «неуд.» по знаниям / повторной тренировке, нарушения НПА",
            "отзыв: акты расследования НС/аварий, предписания энергонадзора",
            "после перерыва — ознакомление с изменениями оборудования, схем и НПА",
        ],
    )

    # 3.8 Briefing
    s = nl()
    fill_list_slide(
        s,
        "3.8 · Производственный инструктаж",
        "Плановый и внеплановый производственный инструктаж",
        "Для диспетчерского, оперативного, ОРп и ремонтного персонала",
        "Периодичность и порядок",
        [
            "плановый для диспетчерского/оперативного/ОРп — ежемесячно",
            "плановый для ремонтного — не реже одного раза в квартал",
            "программа планового инструктажа актуализируется ежегодно",
            "внеплановый — при новых/изменённых документах, схемах, нарушениях",
            "качество усвоения проверяется опросом; регистрация в журнале",
            "допускается совмещать с инструктажем по охране труда",
        ],
    )
    add_image(s, f"{IMG}/ppe.png", bottom=True)

    # 3.9 Fire drills
    s = nl()
    fill_list_slide(
        s,
        "3.9 · Тренировки",
        "Противоаварийные и противопожарные тренировки",
        "Учебные и контрольные тренировки персонала",
        "Периодичность",
        [
            "учебная противопожарная — 1 раз в 3 календарных месяца",
            "контрольная противопожарная — 1 раз в 6 календарных месяцев",
            "допускается совмещать контрольные противоаварийные и противопожарные",
            "проводятся в свободное от дежурства время (или в смену по решению)",
            "при «неуд.» по контрольной ППТ — повтор не позднее 1 месяца",
            "повторный «неуд.» — отстранение до внеочередной проверки знаний",
        ],
        note="Время тренировок включается в рабочее время тренирующихся.",
    )

    # 3.10 DPO
    s = nl()
    fill_list_slide(
        s,
        "3.10 · Повышение квалификации",
        "Дополнительное профессиональное образование",
        "Непрерывный характер повышения квалификации персонала",
        "Требования",
        [
            "ДПО для АТП, диспетчерского, оперативного, ОРп и ремонтного персонала",
            "не реже одного раза в 5 лет с отрывом от основных обязанностей",
            "в образовательном подразделении организации или специализированной организации",
            "краткосрочное обучение АТП — по порядку, установленному руководителем",
            "диспетчерский персонал — периодическое ознакомление с объектами",
            "руководитель обязан организовать повышение квалификации персонала",
        ],
    )

    # 3.11 Walkthroughs
    s = nl()
    fill_list_slide(
        s,
        "3.11 · Обходы и осмотры",
        "Обходы и осмотры рабочих мест",
        "Для диспетчерского, оперативного и оперативно-ремонтного персонала",
        "Что проверяется",
        [
            "выполнение требований НПА и поддержание технологического режима",
            "порядок приёма-сдачи смены, оперативная документация, дисциплина",
            "выявление дефектов и принятие мер к их устранению",
            "правильное применение нарядов-допусков",
            "гигиена труда, исправность инструмента и средств защиты",
            "результаты фиксируются; при нарушениях — меры по устранению",
        ],
        note="Периодичность и ответственные определяются руководителем организации.",
    )
    add_image(s, f"{IMG}/switchgear_room.png", bottom=True)

    # Final summary
    s = nl()
    fill_list_slide(
        s,
        "Итоги темы 3",
        "Ключевые требования к персоналу и его подготовке",
        "Что должен знать каждый специалист",
        "Главные выводы",
        [
            "пять категорий персонала + обязательные формы работы по Приказу № 796",
            "группа по ЭБ — по результатам проверки знаний (кроме I группы)",
            "новая должность: теория → стажировка → проверка → дублирование → допуск",
            "проверка знаний: 12 месяцев / 3 года; комиссия ≥ 5 человек",
            "инструктажи, тренировки, ДПО раз в 5 лет, обходы рабочих мест",
            "допуск без подготовки и проверки знаний запрещён",
        ],
    )


def main():
    prs = Presentation(SRC)
    delete_slides(prs, DELETE_INDICES)
    try:
        add_image(prs.slides[0], f"{IMG}/ppe.png")
    except Exception:
        pass
    prs.save(TMP)

    prs = Presentation(TMP)
    append_missing(prs)
    update_pages(prs)
    prs.save(OUT)
    print(f"Saved {OUT} ({len(prs.slides)} slides)")


if __name__ == "__main__":
    main()
