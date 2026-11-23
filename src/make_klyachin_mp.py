# -*- coding: utf-8 -*-
"""Новая редакция отчёта по УП.01: «Дефектация механизма манометра МТИ-1232»,
преподаватель — Чепик Н.А. Оформление приведено к МУ (Мухина Ю.Н., ТИУ, 2024)."""
import copy
import io
import os
import re

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from PIL import Image
from lxml import etree

SRC = "Отчёт по УП.01 Клячин-3.docx"
WORK_INDEX = "УП.02"
OUT = "Отчёт по УП.02 Клячин.docx"

# ------------------------------------------------------------------ текст
TEXT = {
    # ---------- введение
    63: "Тема «Дефектация механизма манометра МТИ-1232» направлена на формирование навыков "
        "разборки, контроля и оценки технического состояния показывающих приборов давления.",
    64: "Актуальность темы определяется тем, что на каждом производственном объекте эксплуатируется "
        "большое число манометров. Своевременная дефектация позволяет выявить износ передаточного "
        "механизма до появления грубых погрешностей и продлить срок службы прибора.",
    65: "Манометр точных измерений МТИ-1232 — показывающий прибор с одновитковой трубчатой пружиной "
        "и диаметром корпуса 160 мм.",
    66: "Цель работы — закрепить навыки дефектации манометра: освоить последовательность разборки "
        "прибора, контроль его узлов и оценку пригодности к дальнейшей эксплуатации. Для этого "
        "решаются задачи:",
    67: "– изучить нормативные требования к дефектации манометров;",
    68: "– определить последовательность операций разборки и контроля прибора;",
    69: "– рассмотреть применяемые инструменты, приспособления и средства поверки;",
    70: "– освоить методику контроля трубчатой пружины, передаточного механизма и оценку "
        "погрешности показаний;",
    71: "– изучить типовые неисправности манометров и требования безопасности при выполнении работ.",
    72: "Объектом рассмотрения является манометр точных измерений МТИ-1232, предметом — технология "
        "его дефектации. Работа выполнена в период практики с 23 по 28 ноября 2026 г. "
        "в Многопрофильном колледже. Нормативной базой послужили ГОСТ 2405-88 [2] "
        "и Федеральный закон № 102-ФЗ [6].",
    # ---------- раздел 1
    73: "1 Нормативные требования к дефектации манометров",
    74: "Дефектация — это комплекс работ по оценке технического состояния прибора, определению "
        "причин отказов и назначению вида восстановления. Порядок работ и требования к приборам "
        "устанавливают нормативные документы.",
    75: "Общие технические требования к манометрам, вакуумметрам и мановакуумметрам, а также "
        "правила их приёмки определены ГОСТ 2405-88 [2]. Документ устанавливает требования "
        "к шкале, погрешности и герметичности корпуса.",
    76: "В зависимости от состояния прибора по результатам дефектации принимается одно из решений:",
    77: "– прибор годен к эксплуатации — отклонения показаний и состояние деталей соответствуют "
        "требованиям;",
    78: "– прибор подлежит ремонту — выявлен устранимый дефект: ослабление крепления, засорение "
        "канала штуцера, износ зубчатого сектора;",
    79: "– прибор не подлежит ремонту — повреждены пружина, корпус или резьба штуцера, ремонт "
        "нецелесообразен;",
    80: "– прибор подлежит поверке после ремонта — восстановленный прибор допускается "
        "к применению после подтверждения характеристик.",
    81: "Классы точности установлены ГОСТ 8.401-80 [3]. Порядок обращения средств измерений "
        "и требования к поверке определяет Федеральный закон № 102-ФЗ [6], поверка манометров "
        "выполняется по МИ 2124-90 [5], межповерочный интервал МТИ-1232 — один год [4].",
    82: "Прибор признаётся пригодным при соответствии маркировки и диапазона условиям "
        "эксплуатации, отсутствии трещин корпуса и пружины, герметичности соединения "
        "и погрешности в пределах класса точности.",
    # ---------- раздел 2
    83: "2 Порядок и этапы дефектации манометра МТИ-1232",
    84: "Дефектация выполняется в определённой последовательности: от внешнего осмотра, "
        "не требующего разборки, к контролю внутренних узлов и оценке погрешности. Нарушение "
        "последовательности приводит к потере информации о причинах отказа.",
    85: "Подготовительный этап и внешний осмотр",
    86: "Изучается паспорт и техническое описание прибора [4]: диапазон измерений, класс точности, "
        "присоединительная резьба, дата выпуска и отметка о поверке. Проверяются комплектность, "
        "наличие пломбы, состояние стекла, корпуса и присоединительного штуцера.",
    87: "Особое внимание уделяется штуцеру и трубчатой пружине: следы травления, трещины и "
        "деформация свидетельствуют о перегрузке прибора давлением. Общий вид манометра показан "
        "в соответствии с рисунком 1.",
    88: "", 89: "",
    90: "Рисунок 1 – Манометр показывающий с трубчатой пружиной",
    91: "Разборка прибора",
    92: "Снимается защитное кольцо и стекло, стрелка снимается съёмником строго по оси, "
        "отвинчиваются крепёжные винты механизма. Детали укладываются в лоток в порядке разборки, "
        "чтобы при сборке не перепутать крепёж. Прокладки и уплотнительные кольца заменяются "
        "новыми независимо от внешнего вида, поскольку после разборки они теряют герметичность.",
    93: "Контроль трубчатой пружины и передаточного механизма",
    94: "Пружина осматривается на наличие трещин, коррозии и остаточной деформации; проверяется "
        "плавность её перемещения. В передаточном механизме контролируются зацепление зубчатого "
        "сектора с трибкой, износ опор и люфт стрелки, отсутствие засорения камневых опор. "
        "Люфт определяется покачиванием стрелки при неподвижной трибке: заметное перемещение "
        "свидетельствует о износе зацепления и необходимости замены механизма.",
    95: "Проверка герметичности",
    96: "Прибор подключается к поверочному прессу, давление поднимается до верхнего предела "
        "и выдерживается не менее пяти минут. Появление капель на штуцере и в месте пайки "
        "трубчатой пружины не допускается.",
    97: "Оценка погрешности показаний",
    98: "Показания прибора сравниваются с показаниями образцового манометра в пяти точках шкалы "
        "при прямом и обратном ходе. Рассчитывается приведённая погрешность и определяется "
        "соответствие классу точности.",
    99: "Оформление результатов дефектации",
    100: "Результаты заносятся в ведомость дефектации: данные прибора, перечень проверенных "
         "параметров, выявленные дефекты, заключение о пригодности и рекомендуемый вид "
         "восстановления. Ведомость подписывается исполнителем и руководителем работ.",
    # ---------- раздел 3
    101: "3 Инструменты, приспособления и средства поверки",
    102: "Комплект инструмента подбирается по перечню операций [4], средства поверки — "
         "по МИ 2124-90 [5]. Основной инструмент — набор "
         "отвёрток с изолированными рукоятками для снятия кольца, стекла и крепёжных винтов, "
         "а также ключ под присоединительную резьбу М20×1,5 для монтажа прибора на прессе. "
         "Для оценки состояния деталей применяются лупа с увеличением не менее четырёх крат "
         "и штангенциркуль.",
    103: "Инструмент должен быть исправным и соответствовать операции: для снятия стрелки "
         "используется съёмник, для затяжки соединений — ключи нужного профиля, для промывки "
         "деталей — ультразвуковая ванна. Инструмент показан в соответствии с рисунком 2, "
         "его назначение приведено в таблице 1.",
    104: "", 105: "",
    106: "Рисунок 2 – Набор отвёрток для разборки корпуса прибора",
    107: "",
    108: "Таблица 1 – Инструменты и приспособления для дефектации",
    # ---------- раздел 4
    110: "4 Методика выполнения основных операций дефектации",
    111: "Разборка корпуса и снятие стрелки",
    112: "Работа выполняется на верстаке с мягким покрытием, чтобы не повредить стекло и корпус. "
         "Кольцо снимается отвёрткой по периметру, стрелка — съёмником; усилие прикладывается "
         "по оси, перекос не допускается, поскольку посадка стрелки на оси выполнена с натягом.",
    113: "Промывка и контроль передаточного механизма",
    114: "Снятый механизм промывается в ультразвуковой ванне, продувается сжатым воздухом "
         "и осматривается. Проверяется плавность перемещения сектора, отсутствие люфта "
         "и заеданий, состояние опор. Схема устройства прибора приведена в соответствии "
         "с рисунком 3.",
    115: "", 116: "",
    117: "Рисунок 3 – Устройство манометра с трубчатой пружиной",
    118: "",
    119: "Основные технические характеристики прибора, которые проверяются при дефектации, "
         "приведены в таблице 2.",
    120: "Таблица 2 – Технические характеристики манометра МТИ-1232",
    121: "", 122: "",
    123: "Проверка герметичности и опрессовка",
    124: "Прибор закрепляется на поверочном прессе, полость заполняется рабочей средой "
         "и плавно нагружается до верхнего предела измерений. Давление выдерживается не менее "
         "пяти минут, после чего осматриваются штуцер, корпус и место пайки трубчатой пружины; "
         "падение давления по образцовому манометру свидетельствует о негерметичности.",
    125: "Определение погрешности по образцовому манометру",
    126: "Проверка выполняется в пяти точках шкалы при прямом и обратном ходе:",
    127: "– при нулевом значении давления стрелка устанавливается на нулевую отметку корректором;",
    128: "– при 25 % верхнего предела показания сравниваются с образцовым прибором;",
    129: "– при 50 % и 75 % предела фиксируются отклонения прямого и обратного хода;",
    130: "– при верхнем пределе прибор выдерживается пять минут для оценки остаточной деформации.",
    131: "Приведённая погрешность рассчитывается как отношение наибольшего отклонения к верхнему "
         "пределу измерений. Прибор считается годным, если погрешность не превышает предела, "
         "допускаемого для класса точности 0,6 или 1,0 [3].",
    # ---------- раздел 5
    132: "5 Типовые неисправности механизма манометра",
    133: "Большинство дефектов относится к ограниченному кругу типовых неисправностей, знание "
         "которых ускоряет поиск причины отказа. Частые неисправности манометров и способы "
         "их устранения приведены в таблице 3 [7].",
    134: "Таблица 3 – Типовые неисправности и способы их устранения",
    135: "", 136: "",
    137: "После устранения дефектов прибор проверяется повторно в полном объёме, поскольку "
         "изменения в одном узле могут повлиять на показания в целом. Отремонтированный прибор "
         "допускается к применению только после поверки.",
    # ---------- раздел 6
    138: "6 Требования безопасности при дефектации манометров",
    139: "Работы по дефектации относятся к работам повышенной опасности, поскольку связаны "
         "с сосудами под давлением и слесарным инструментом. Общие требования безопасности "
         "производственных процессов установлены ГОСТ 12.3.002-2014 [1], до начала работ "
         "проводится инструктаж на рабочем месте.",
    140: "Перед разборкой прибор отсоединяется от технологической линии, давление сбрасывается, "
          "остатки среды сливаются в поддон. При работе на поверочном прессе применяются защитные "
          "очки и перчатки, нагружение выполняется плавно, без рывков; рабочее место оборудуется "
          "местным освещением и вытяжной вентиляцией, поскольку рабочая среда может быть "
          "токсичной.",
    141: "Запрещается разбирать прибор, находящийся под давлением, применять открытый огонь "
         "и нагревать узлы, работать неисправным инструментом, оставлять на верстаке снятые "
         "детали без лотка.",
    142: "Персонал должен уметь оказывать первую помощь при травмах и попадании рабочей среды "
         "на кожу, знать расположение средств пожаротушения и порядок вызова медицинской помощи.",
    # ---------- заключение
    143: "Заключение",
    144: "В ходе учебной практики рассмотрена технология дефектации манометра точных измерений "
         "МТИ-1232. Изучены нормативные документы и определена последовательность операций — "
         "от внешнего осмотра и разборки до оценки погрешности и оформления ведомости дефектации.",
    145: "Установлено, что пригодность прибора определяется совокупностью признаков: состоянием "
         "трубчатой пружины и передаточного механизма, герметичностью соединения и погрешностью "
         "показаний в пределах класса точности. Наибольшее влияние на точность оказывают износ "
         "зубчатого сектора и остаточная деформация пружины.",
    146: "Рассмотрены инструменты и приспособления, методика контроля узлов, типовые неисправности "
         "и способы их устранения (таблица 3), а также требования безопасности. Выполненные "
         "операции по разборке, промывке и оценке состояния прибора позволили закрепить "
         "практические навыки и понять, что качественная дефектация продлевает срок службы "
         "средства измерений.",
    # ---------- список источников
    147: "Список использованных источников",
    148: "1. ГОСТ 12.3.002-2014. Система стандартов безопасности труда. Процессы "
         "производственные. Общие требования безопасности. – Текст : электронный. – URL: "
         "http://docs.cntd.ru/document/1200124407 (дата обращения: 24.11.2026).",
    149: "2. ГОСТ 2405-88. Манометры, вакуумметры, мановакуумметры, напоромеры, тягомеры "
         "и тягонапоромеры. Общие технические условия. – Текст : электронный. – URL: "
         "http://docs.cntd.ru/document/1200001405 (дата обращения: 25.11.2026).",
    150: "3. ГОСТ 8.401-80. Классы точности средств измерений. Общие требования. – "
         "Текст : электронный. – URL: http://docs.cntd.ru/document/gost-8-401-80-gsi "
         "(дата обращения: 23.11.2026).",
    151: "4. Манометр точных измерений МТИ-1232 : техническое описание и инструкция "
         "по эксплуатации. – Энгельс : ООО «Манометр», 2021. – 8 с. – "
         "Текст : непосредственный.",
    152: "5. Манометры точных измерений МТИ-1232 : [сайт]. – Москва, 2026. – URL: "
         "https://manometry.ru/mti-1232 (дата обращения: 26.11.2026). – Текст : электронный.",
    153: "6. МИ 2124-90. Методика поверки манометров, вакуумметров и мановакуумметров. – "
         "Текст : электронный. – URL: http://www.1bm.ru/techdocs/kgs/gost/934/info/28850/ "
         "(дата обращения: 27.11.2026).",
    154: "7. Об обеспечении единства измерений : Федеральный закон № 102-ФЗ : [принят "
         "Государственной Думой 26 июня 2008 г.]. – Текст : электронный // "
         "КонсультантПлюс : [сайт]. – URL: http://www.consultant.ru "
         "(дата обращения: 28.11.2026).",
}

TOC = {
    52: ("Введение", "3"),
    53: ("1 Нормативные требования к дефектации манометров", "4"),
    54: ("2 Порядок и этапы дефектации манометра МТИ-1232", "5"),
    55: ("3 Инструменты, приспособления и средства поверки", "7"),
    56: ("4 Методика выполнения основных операций дефектации", "8"),
    57: ("5 Типовые неисправности механизма манометра", "10"),
    58: ("6 Требования безопасности при дефектации манометров", "11"),
    59: ("Заключение", "12"),
    60: ("Список использованных источников", "13"),
}

# ------------------------------------------------------------------ таблицы
TABLES = [
    [   # таблица 1 — инструменты
        ["Инструмент, приспособление", "Назначение", "Основные характеристики"],
        ["Набор отвёрток", "Снятие кольца, стекла, крепёжных винтов", "Шлицевые и крестовые, рукоятки изолированные"],
        ["Ключ под резьбу М20×1,5", "Монтаж прибора на поверочном прессе", "Рожковый или торцевой, сталь"],
        ["Съёмник стрелок", "Снятие стрелки с оси трибки", "Усилие по оси, без перекоса"],
        ["Лупа и штангенциркуль", "Осмотр деталей, контроль размеров", "Увеличение 4 крат, цена деления 0,05 мм"],
        ["Ультразвуковая ванна", "Промывка механизма", "Частота 40 кГц, 3–5 мин"],
        ["Поверочный пресс", "Опрессовка и оценка погрешности", "Образцовый прибор класса 0,4, до 60 МПа"],
    ],
    [   # таблица 2 — характеристики МТИ-1232
        ["Параметр", "Значение", "Нормативный документ"],
        ["Диапазон измерений", "0–16; 25; 40; 60 МПа (160–600 кгс/см²)", "ТУ 25.05.1481-77"],
        ["Класс точности", "0,6 или 1,0", "ГОСТ 8.401-80"],
        ["Диаметр корпуса, резьба", "160 мм; М20×1,5", "ГОСТ 2405-88"],
        ["Степень защиты, межповерочный интервал", "IP40; 1 год", "ГОСТ 14254-96; МИ 2124-90"],
    ],
    [   # таблица 3 — неисправности
        ["Неисправность", "Причина", "Способ устранения"],
        ["Стрелка не возвращается к нулю", "Остаточная деформация трубчатой пружины", "Заменить пружину, отрегулировать нуль корректором"],
        ["Завышенная погрешность показаний", "Износ зубчатого сектора и трибки, люфт", "Заменить механизм или сектор, отрегулировать тягу"],
        ["Негерметичность по штуцеру", "Ослабление резьбы, износ уплотнения", "Подтянуть штуцер, заменить уплотнительную прокладку"],
        ["Засорение канала штуцера", "Загрязнение рабочей среды", "Продуть канал, промыть и просушить прибор"],
        ["Помутнение или трещина стекла", "Механическое воздействие, старение", "Заменить стекло и уплотнительное кольцо"],
    ],
]

# ------------------------------------------------------------------ рисунки
IMAGE_REPL = {
    0: "photos/manometr.jpg",
    1: "photos/otvertki.jpg",
    2: "photos/ustroystvo.jpg",
}

TITUL_INSERT = [
    # (после абзаца, текст) — вставляется в обратном порядке, чтобы не сбить индексы
    (29, "ОТ ПРОФИЛЬНОЙ ОРГАНИЗАЦИИ __________________ ________________________________"),
    (25, "В качестве ______________________________________________________________"),
]

# ------------------------------------------------------------------ штамп
HEADER_REPL = [
    ("02403", "0240"),                    # шифр по МУ: 4 последние цифры зачётки
    ("УП.01.-15.01.37-0240-2026", "15.01.37.0240.2026.УП.01"),
    ("Марковских", "Чепик"),
    ("А.В.", "Н.А."),
    ("Проверка электрических схем ", "Дефектация механизма манометра "),
    ("низковольтных цепей при монтаже", "МТИ-1232"),
    ("Листов", "Листов"),                 # заглушка, число правится отдельно
]


# список источников переупорядочен по МУ: НПА, затем ГОСТ/МИ в хронологическом
# порядке, затем техническая документация, затем сайт.
# старый номер ссылки -> новый номер
REF_MAP = {1: 1, 2: 2, 3: 3, 4: 4, 5: 6, 6: 7, 7: 5}


def remap_refs(text):
    """Пересчитывает ссылки вида [n] под новый порядок списка источников."""
    return re.sub(r"\[(\d)\]", lambda m: "[%d]" % REF_MAP.get(int(m.group(1)), int(m.group(1))), text)


def set_text(par, text):
    runs = par.runs
    if not runs:
        par.add_run(text)
        return
    runs[0].text = text
    for r in runs[1:]:
        r.text = ""


def drop_src_rect(shape):
    """В исходных рамках стояла обрезка a:srcRect от старых картинок.
    После замены изображения она режет новый рисунок — убираем."""
    for el in shape._inline.iter(
        "{http://schemas.openxmlformats.org/drawingml/2006/main}srcRect"
    ):
        el.getparent().remove(el)


def replace_image(doc, index, path):
    shape = doc.inline_shapes[index]
    drop_src_rect(shape)
    rid = shape._inline.graphic.graphicData.pic.blipFill.blip.get(qn("r:embed"))
    part = doc.part.related_parts[rid]
    ratio = shape.width / shape.height
    im = Image.open(path)
    w, h = im.size
    if w / h > ratio:
        nw = int(h * ratio)
        im = im.crop(((w - nw) // 2, 0, (w - nw) // 2 + nw, h))
    else:
        nh = int(w / ratio)
        im = im.crop((0, (h - nh) // 2, w, (h - nh) // 2 + nh))
    buf = io.BytesIO()
    im.convert("RGB").save(buf, "JPEG", quality=88, optimize=True, progressive=True)
    part._blob = buf.getvalue()


def strip_bold(doc):
    """МУ: полужирный шрифт не применяется."""
    body = doc.element.body
    for el in body.iter():
        if el.tag == qn("w:b") or el.tag == qn("w:bCs"):
            el.getparent().remove(el)
    for part in doc.part.package.parts:
        pn = str(part.partname)
        if pn.startswith("/word/header") or pn.startswith("/word/footer"):
            for el in list(part.element.iter()):
                if el.tag == qn("w:b") or el.tag == qn("w:bCs"):
                    el.getparent().remove(el)



def fix_titul(doc):
    """Титульный лист: пробелы в датах, «Обучающегося» по приложению А МУ,
    и запас по высоте, чтобы «Тюмень / 2026 г.» оставались на первой странице."""
    for par in doc.paragraphs:
        if not par.runs:
            continue
        joined = "".join(r.text or "" for r in par.runs)
        if "2026г." in joined or "Обучающийся " in joined:
            new_text = joined.replace("2026г.", "2026 г.").replace(
                "Обучающийся ", "Обучающегося "
            )
            par.runs[0].text = new_text
            for r in par.runs[1:]:
                r.text = ""
    ps = doc.paragraphs
    # титульный лист набран 12 пт, как в принятом образце (приложение А МУ):
    # при 14 пт строки формы переносятся и титул не укладывается на страницу
    # строка «ОТ ПРОФИЛЬНОЙ ОРГАНИЗАЦИИ» должна укладываться в одну строку
    for par in ps:
        if par.text.strip().startswith("В период"):
            joined = "".join(r.text or "" for r in par.runs)
            par.runs[0].text = (
                joined.replace("«09»", "«23»").replace("«21»", "«28»")
            )
            par.runs[0].text = " ".join(par.runs[0].text.split("  ")).strip()
            for r in par.runs[1:]:
                r.text = ""
    for par in ps:
        if par.text.strip().startswith("ОТ ПРОФИЛЬНОЙ ОРГАНИЗАЦИИ"):
            joined = " ".join(par.text.split())
            head = "ОТ ПРОФИЛЬНОЙ ОРГАНИЗАЦИИ "
            par.runs[0].text = head + "_" * 28
            for r in par.runs[1:]:
                r.text = ""
    # индекс практики на титульном листе
    for par in ps:
        if par.text.strip().startswith("ПО УП."):
            par.runs[0].text = f"ПО {WORK_INDEX} УЧЕБНОЙ ПРАКТИКЕ"
            for r in par.runs[1:]:
                r.text = ""

    # содержание открывает отдельную секцию — разрыв страницы даёт сама секция
    idx = next((i for i, p in enumerate(ps) if p.text.strip() == "Тюмень"), None)
    if idx is None:
        return
    removed = 0
    for j in range(idx - 1, -1, -1):
        if removed >= 4:
            break
        par = ps[j]
        xml = par._p.xml
        if par.text.strip() or "w:br" in xml or "graphicData" in xml:
            continue
        par._p.getparent().remove(par._p)
        removed += 1


def keep_tables_whole(doc):
    """Таблица не разрывается между страницами: МУ разрешает перенос таблицы
    на следующую страницу, но требует надпись «Продолжение таблицы N»,
    поэтому проще держать таблицу целиком."""
    for table in doc.tables:
        rows = table.rows
        for ri, row in enumerate(rows):
            trPr = row._tr.get_or_add_trPr()
            if trPr.find(qn("w:cantSplit")) is None:
                cant = trPr.makeelement(qn("w:cantSplit"), {})
                trPr.append(cant)
            if ri == 0 and trPr.find(qn("w:tblHeader")) is None:
                trPr.append(trPr.makeelement(qn("w:tblHeader"), {}))
            if ri < len(rows) - 1:
                for cell in row.cells:
                    for par in cell.paragraphs:
                        par.paragraph_format.keep_with_next = True


def link_urls(doc):
    """Превращает адреса URL в списке источников в настоящие гиперссылки."""
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    from docx.opc.constants import RELATIONSHIP_TYPE as RT

    for par in doc.paragraphs:
        text = par.text
        m = re.search(r"https?://[^\s]+", text)
        if not m or "URL" not in text:
            continue
        url = m.group(0)
        before, after = text[: m.start()], text[m.end() :]
        rpr = ""
        if par.runs:
            r = par.runs[0]
            rpr = (
                '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
                f'<w:sz w:val="{int((r.font.size.pt if r.font.size else 14) * 2)}"/>'
                '<w:szCs w:val="28"/></w:rPr>'
            )
        for r in par.runs[1:]:
            r.text = ""
        if par.runs:
            par.runs[0].text = before
        else:
            par.add_run(before)
        anchor = par.runs[0]._r
        rid = doc.part.relate_to(url, RT.HYPERLINK, is_external=True)
        hl = parse_xml(
            f'<w:hyperlink {nsdecls("w", "r")} r:id="{rid}">'
            f'<w:r {nsdecls("w")}>{rpr}<w:t xml:space="preserve">{url}</w:t></w:r>'
            "</w:hyperlink>"
        )
        anchor.addnext(hl)
        if after:
            tail = parse_xml(
                f'<w:r {nsdecls("w")}>{rpr}<w:t xml:space="preserve">{after}</w:t></w:r>'
            )
            hl.addnext(tail)


def tidy_figures(doc):
    """Убираем пустые абзацы вокруг рисунков и подписей (они «съедали» по 1–2 см
    и выталкивали таблицы на следующий лист) и слегка уменьшаем рисунок 2."""
    ps = doc.paragraphs
    def is_fig(par):
        return "graphicData" in par._p.xml

    def is_caption(par):
        return par.text.strip().startswith("Рисунок ")

    def is_empty(par):
        xml = par._p.xml
        return (
            not par.text.strip()
            and "graphicData" not in xml
            and "w:sectPr" not in xml
            and 'w:type="page"' not in xml
        )

    for i in range(len(ps) - 1, -1, -1):
        par = ps[i]
        if not is_empty(par):
            continue
        prev_par = ps[i - 1] if i else None
        next_par = ps[i + 1] if i + 1 < len(ps) else None
        near = any(
            x is not None and (is_fig(x) or is_caption(x))
            for x in (prev_par, next_par)
        )
        if near:
            par._p.getparent().remove(par._p)

    for par in doc.paragraphs:
        if is_fig(par):
            par.paragraph_format.line_spacing = 1.0
    # рисунок 2 — чуть меньше, чтобы раздел 3 уложился на один лист
    shapes = doc.inline_shapes
    if len(shapes) > 1:
        sh = shapes[1]
        ratio = sh.width / sh.height
        sh.width = Cm(4.0)
        sh.height = Cm(round(4.0 / ratio, 2))


def compact_tables(doc):
    """Таблицы должны укладываться на лист целиком: кегль 11 (МУ допускает
    не менее 10) и минимальные внутренние отступы ячеек."""
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls

    for table in doc.tables:
        tblPr = table._tbl.tblPr
        old = tblPr.find(qn("w:tblCellMar"))
        if old is not None:
            tblPr.remove(old)
        tblPr.append(
            parse_xml(
                f'<w:tblCellMar {nsdecls("w")}>'
                '<w:top w:w="14" w:type="dxa"/><w:left w:w="80" w:type="dxa"/>'
                '<w:bottom w:w="14" w:type="dxa"/><w:right w:w="80" w:type="dxa"/>'
                "</w:tblCellMar>"
            )
        )
        for row in table.rows:
            for cell in row.cells:
                for par in cell.paragraphs:
                    par.paragraph_format.space_before = Pt(0)
                    par.paragraph_format.space_after = Pt(0)
                    for r in par.runs:
                        r.font.size = Pt(11)


def add_toc_pagerefs(doc):
    """Номера страниц в содержании — поля PAGEREF на закладки заголовков,
    чтобы при любом сдвиге вёрстки они пересчитывались сами."""
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls

    RPR = (
        f'<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
        f'<w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>'
    )
    ps = doc.paragraphs
    start = next((i for i, p in enumerate(ps) if p.text.strip() == "СОДЕРЖАНИЕ"), None)
    if start is None:
        return
    rows = []
    for i in range(start + 1, len(ps)):
        if "\t" in ps[i].text:
            rows.append(i)
        elif ps[i].text.strip() and rows:
            break
    bm_id = 900
    for n, i in enumerate(rows, 1):
        par = ps[i]
        title, _, page = par.text.partition("\t")
        title = title.strip()
        page = re.sub(r"\D", "", page) or "1"
        target = next(
            (p for p in ps[rows[-1] + 1 :] if p.text.strip() == title), None
        )
        if target is None:
            continue
        bm = "_RefHeading%d" % n
        bm_id += 1
        bs = parse_xml(f'<w:bookmarkStart {nsdecls("w")} w:id="{bm_id}" w:name="{bm}"/>')
        pPr = target._p.find(qn("w:pPr"))
        if pPr is not None:
            pPr.addnext(bs)
        else:
            target._p.insert(0, bs)
        target._p.append(
            parse_xml(f'<w:bookmarkEnd {nsdecls("w")} w:id="{bm_id}"/>')
        )
        runs = par.runs
        idx = next((k for k, r in enumerate(runs) if "\t" in (r.text or "")), None)
        if idx is None:
            continue
        tabrun = runs[idx]
        tabrun.text = "\t"
        for r in runs[idx + 1 :]:
            r.text = ""
        anchor = tabrun._r
        parts = [
            f'<w:r {nsdecls("w")}>{RPR}<w:fldChar w:fldCharType="begin"/></w:r>',
            f'<w:r {nsdecls("w")}>{RPR}<w:instrText xml:space="preserve"> PAGEREF {bm} \\h </w:instrText></w:r>',
            f'<w:r {nsdecls("w")}>{RPR}<w:fldChar w:fldCharType="separate"/></w:r>',
            f'<w:r {nsdecls("w")}>{RPR}<w:t>{page}</w:t></w:r>',
            f'<w:r {nsdecls("w")}>{RPR}<w:fldChar w:fldCharType="end"/></w:r>',
        ]
        for xml in parts:
            el = parse_xml(xml)
            anchor.addnext(el)
            anchor = el


def auto_page_total(doc):
    """Графа «Листов» в основной надписи — поле NUMPAGES вместо числа."""
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls

    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    for part in doc.part.package.parts:
        if not str(part.partname).startswith("/word/header"):
            continue
        ts = list(part.element.iter(W + "t"))
        for k, t in enumerate(ts[:-1]):
            if (t.text or "").strip() != "Листов":
                continue
            nxt = ts[k + 1]
            run = nxt.getparent()
            rpr = run.find(qn("w:rPr"))
            rpr_xml = (
                etree.tostring(rpr, encoding="unicode") if rpr is not None else ""
            )
            cached = (nxt.text or "13").strip()
            anchor = run
            for xml in (
                f'<w:r {nsdecls("w")}>{rpr_xml}<w:fldChar w:fldCharType="begin"/></w:r>',
                f'<w:r {nsdecls("w")}>{rpr_xml}<w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r>',
                f'<w:r {nsdecls("w")}>{rpr_xml}<w:fldChar w:fldCharType="separate"/></w:r>',
                f'<w:r {nsdecls("w")}>{rpr_xml}<w:t>{cached}</w:t></w:r>',
                f'<w:r {nsdecls("w")}>{rpr_xml}<w:fldChar w:fldCharType="end"/></w:r>',
            ):
                el = parse_xml(xml)
                anchor.addnext(el)
                anchor = el
            run.getparent().remove(run)
            break


def enable_update_fields(doc):
    """Word пересчитывает поля (содержание, «Листов») при открытии файла."""
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls

    settings = doc.settings.element
    if settings.find(qn("w:updateFields")) is not None:
        return
    el = parse_xml(f'<w:updateFields {nsdecls("w")} w:val="true"/>')
    anchor = settings.find(qn("w:hdrShapeDefaults")) or settings.find(qn("w:compat"))
    if anchor is not None:
        anchor.addprevious(el)
    else:
        settings.append(el)


def merge_titul_and_toc(doc):
    """Титул + содержание — одна секция без колонтитула (МУ: содержание без номера страницы)."""
    body = doc.element.body
    sect_prs = []
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            pPr = child.find(qn("w:pPr"))
            if pPr is not None:
                sp = pPr.find(qn("w:sectPr"))
                if sp is not None:
                    sect_prs.append((pPr, sp))
    if len(sect_prs) < 2:
        return
    (pPr_t, sp_t) = sect_prs[0]      # секция титула
    (pPr_c, sp_c) = sect_prs[1]      # секция содержания
    # переносим параметры страницы титула в секцию содержания
    for tag in ("w:pgMar", "w:pgSz", "w:cols", "w:docGrid"):
        old = sp_c.find(qn(tag))
        if old is not None:
            sp_c.remove(old)
        new = sp_t.find(qn(tag))
        if new is not None:
            sp_c.append(copy.deepcopy(new))
    # убираем колонтитул (номер страницы) и нумерацию
    for ref in sp_c.findall(qn("w:headerReference")):
        sp_c.remove(ref)
    for tag in ("w:pgNumType", "w:titlePg"):
        el = sp_c.find(qn(tag))
        if el is not None:
            sp_c.remove(el)
    # удаляем разрыв секции после титула
    sp_t.getparent().remove(sp_t)


def main():
    doc = Document(SRC)
    ps = doc.paragraphs

    # 1. текст по индексам (абзацы с рисунками не трогаем — иначе пропадёт картинка)
    for idx, text in TEXT.items():
        if "graphicData" in ps[idx]._p.xml:
            continue
        set_text(ps[idx], remap_refs(text))

    # 2. содержание
    for idx, (title, page) in TOC.items():
        par = ps[idx]
        runs = par.runs
        if not runs:
            par.add_run(title)
            continue
        runs[0].text = title
        if len(runs) > 1:
            runs[1].text = "\t" + page
        for r in runs[2:]:
            r.text = ""

    # 3. таблицы
    for table, rows in zip(doc.tables, TABLES):
        for row, values in zip(table.rows, rows):
            for cell, value in zip(row.cells, values):
                par = cell.paragraphs[0]
                if par.runs:
                    set_text(par, value)
                else:
                    par.add_run(value)

    # 4. рисунки
    for idx, path in IMAGE_REPL.items():
        replace_image(doc, idx, path)

    # 5. оформление по МУ
    strip_bold(doc)

    headings = [62, 73, 83, 101, 110, 132, 138, 143, 147]
    subheads = [85, 91, 93, 95, 97, 99, 111, 113, 123, 125]
    captions_tbl = [108, 120, 134]
    captions_fig = [90, 106, 117]

    for i, par in enumerate(ps):
        pf = par.paragraph_format
        if i in headings:
            pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            pf.first_line_indent = Cm(1.25)
            pf.space_before = pf.space_before
            pf.keep_with_next = True          # МУ: заголовок не оставлять внизу страницы
        elif i in subheads:
            pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            pf.first_line_indent = Cm(1.25)
            pf.keep_with_next = True
        elif i in captions_tbl:
            pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            pf.first_line_indent = Cm(1.25)
            pf.space_before = Pt(12)
            pf.keep_with_next = True          # подпись не отрывается от таблицы
            for r in par.runs:                # МУ: кегль текста 14, в таблице — не менее 10
                r.font.size = Pt(14)
        elif i in captions_fig:
            pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf.first_line_indent = Cm(0)
            for r in par.runs:
                r.font.size = Pt(14)

    # подпись рисунка не отрывается от самого рисунка
    for par in ps:
        if "graphicData" in par._p.xml:
            par.paragraph_format.keep_with_next = True

    # 6. вставки на титульном листе (с конца, чтобы индексы не сдвигались)
    for after, text in sorted(TITUL_INSERT, reverse=True):
        base = ps[after]
        newp = copy.deepcopy(base._p)
        base._p.addnext(newp)
        from docx.text.paragraph import Paragraph
        set_text(Paragraph(newp, base._parent), text)

    # 6б. титульный лист: пробелы, «Обучающегося», запас по высоте
    fix_titul(doc)

    # 7. штамп и титульный лист
    for part in doc.part.package.parts:
        pn = str(part.partname)
        if not pn.startswith("/word/header"):
            continue
        for t in part.element.iter(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"
        ):
            if not t.text:
                continue
            for old, new in HEADER_REPL:
                if old in t.text:
                    t.text = t.text.replace(old, new)

    # число листов в штампе — фактическое
    for part in doc.part.package.parts:
        if not str(part.partname).startswith("/word/header"):
            continue
        ts = [
            t
            for t in part.element.iter(
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"
            )
        ]
        for k, t in enumerate(ts[:-1]):
            if t.text and t.text.strip() == "Листов":
                nxt = ts[k + 1]
                if nxt.text and nxt.text.strip().isdigit():
                    nxt.text = "13"


    # 8. шифр в штампе состоит из нескольких коротких прогонов — собираем их в один
    for part in doc.part.package.parts:
        pn = str(part.partname)
        if not pn.startswith("/word/header"):
            continue
        texts = [
            t
            for t in part.element.iter(
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"
            )
        ]
        i = 0
        while i < len(texts):
            if (texts[i].text or "").startswith("УП.0"):
                acc, j = "", i
                while j < len(texts) and len(acc) < 34 and not (
                    len(acc) >= 20 and acc.endswith("2026")
                ):
                    acc += texts[j].text or ""
                    j += 1
                texts[i].text = "15.01.37.0240.2026." + WORK_INDEX
                for t in texts[i + 1 : j]:
                    t.text = ""
                i = j
            else:
                i += 1

    # 9. ФИО руководителя от университета на титульном листе
    for par in doc.element.body.iter(
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"
    ):
        texts = [
            t
            for t in par.iter(
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"
            )
        ]
        joined = "".join(t.text or "" for t in texts)
        if "Марковских" in joined:
            texts[0].text = joined.replace("Марковских Андрей Викторович",
                                           "Чепик Никита Александрович")
            for t in texts[1:]:
                t.text = ""

    # 9б. таблицы не разрываются между страницами
    keep_tables_whole(doc)

    # 10. содержание без номера страницы: объединяем титул и содержание в одну секцию без колонтитула
    # титул, содержание и основная часть — три секции (как в исходном файле):
    # содержание и все листы основной части с рамкой и штампом
    # merge_titul_and_toc(doc)  # не объединяем: тогда лист содержания остаётся без рамки

    tidy_figures(doc)
    compact_tables(doc)
    link_urls(doc)
    add_toc_pagerefs(doc)
    auto_page_total(doc)
    enable_update_fields(doc)

    for part in doc.part.package.parts:
        if not str(part.partname).startswith("/word/") or not hasattr(part, "element"):
            continue
        for t in part.element.iter(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"
        ):
            if t.text and "УП.01" in t.text:
                t.text = t.text.replace("УП.01", WORK_INDEX)

    doc.save(OUT)
    print("сохранено:", OUT, os.path.getsize(OUT), "байт")


if __name__ == "__main__":
    main()
