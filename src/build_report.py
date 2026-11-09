# -*- coding: utf-8 -*-
"""Сборка отчёта УП.01 на оригинальном шаблоне (рамка и штамп из «Пример отчёта.docx»)."""
import os
import sys

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src import content as C                 # noqa: E402
from src.layout_estimate import estimate    # noqa: E402

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(BASE, "google_drive", "Пример отчёта.docx")
OUT = os.path.join(BASE, "Отчет_УП.01_Клячин_К_Е.docx")

FONT = "Times New Roman"
OLD_NAME = "Рыбьяков Максим Евгеньевич"
NEW_NAME = "Клячин Кирилл Евгеньевич"
OLD_SHORT = "СВОЕ ФИО"
NEW_SHORT = "Клячин К.Е."
OLD_NUM = "21574"
NEW_NUM = "02403"
OLD_TOPIC_1 = "Монтаж элек"
NEW_TOPIC_1 = "Проверка электрических схем "
OLD_TOPIC_2 = "трических схем приборов для измерения расхода жидкости"
NEW_TOPIC_2 = "низковольтных цепей при монтаже"

# ------------------------------------------------------------------ форматирование


def set_font(run, size=14, bold=False):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), FONT)


def fmt_par(p, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_indent=1.25, before=0, after=0,
            spacing=1.5, keep_next=False):
    pf = p.paragraph_format
    pf.alignment = align
    pf.first_line_indent = Cm(first_indent)
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = spacing
    pf.keep_with_next = keep_next
    return p


def add_par(doc, text="", size=14, bold=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
            first_indent=1.25, before=0, after=0, spacing=1.5):
    p = doc.add_paragraph()
    fmt_par(p, align, first_indent, before, after, spacing)
    if text:
        set_font(p.add_run(text), size=size, bold=bold)
    return p


# ------------------------------------------------------------------ служебное


def split_breaks(body):
    """Абзацы, содержащие sectPr (границы разделов шаблона)."""
    out = []
    for p in body.findall(qn("w:p")):
        ppr = p.find(qn("w:pPr"))
        if ppr is not None and ppr.find(qn("w:sectPr")) is not None:
            out.append(p)
    return out


def replace_in_runs(paragraphs, old, new):
    """Замена текста в абзацах (в пределах одного run)."""
    done = 0
    for p in paragraphs:
        for r in p.iter(qn("w:t")):
            if r.text and old in r.text:
                r.text = r.text.replace(old, new)
                done += 1
    return done


def edit_header(section, pairs):
    """Замена текста внутри фигур колонтитула (рамка со штампом)."""
    hdr = section.header
    el = hdr._element
    for t in el.iter(qn("w:t")):
        if not t.text:
            continue
        for old, new in pairs:
            if t.text == old:
                t.text = new
                if new != new.strip():
                    t.set(qn("xml:space"), "preserve")
            elif old in t.text and old not in ("11",):
                t.text = t.text.replace(old, new)


WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
VML = "{urn:schemas-microsoft-com:vml}"

# поля рамки по ГОСТ 2.104: слева 20 мм, сверху/справа/снизу 5 мм
FRAME_X = 720000            # 2,0 см (20 мм)
FRAME_Y = 180000            # 0,5 см (5 мм)
FRAME_CX = 6660000          # 21 - 2,0 - 0,5 = 18,5 см
FRAME_CY = 10332000         # 29,7 - 0,5 - 0,5 = 28,7 см


def fix_frame(section):
    """Ставит рамку колонтитула ровно по ГОСТ 2.104: 20/5/5/5 мм."""
    import re as _re
    el = section.header._element
    for a in el.iter(WP + "anchor"):
        ext = a.find(WP + "extent")
        if ext is not None:
            ext.set("cx", str(FRAME_CX))
            ext.set("cy", str(FRAME_CY))
        ph = a.find(WP + "positionH")
        if ph is not None and ph.find(WP + "posOffset") is not None:
            ph.find(WP + "posOffset").text = str(FRAME_X)
        pv = a.find(WP + "positionV")
        if pv is not None and pv.find(WP + "posOffset") is not None:
            pv.find(WP + "posOffset").text = str(FRAME_Y)
    for g in el.iter(VML + "group"):
        st = g.get("style") or ""
        st = _re.sub(r"margin-left:[\d.]+pt", "margin-left:56.7pt", st)
        st = _re.sub(r"margin-top:[\d.]+pt", "margin-top:14.2pt", st)
        st = _re.sub(r"width:[\d.]+pt", "width:524.4pt", st)
        st = _re.sub(r"height:[\d.]+pt", "height:813.5pt", st)
        g.set("style", st)


def _num_run(text=None, instr=None):
    """Ран с настройками шрифта для поля номера страницы."""
    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rf = OxmlElement("w:rFonts")
    for a in ("w:ascii", "w:hAnsi", "w:cs"):
        rf.set(qn(a), FONT)
    it = OxmlElement("w:i")
    it.set(qn("w:val"), "0")
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), "24")
    szcs = OxmlElement("w:szCs")
    szcs.set(qn("w:val"), "24")
    for e in (rf, it, sz, szcs):
        rPr.append(e)
    r.append(rPr)
    if instr is not None:
        ie = OxmlElement("w:instrText")
        ie.set(qn("xml:space"), "preserve")
        ie.text = instr
        r.append(ie)
    if text is not None:
        t = OxmlElement("w:t")
        t.text = text
        r.append(t)
    return r


def _fld(kind):
    r = OxmlElement("w:r")
    fc = OxmlElement("w:fldChar")
    fc.set(qn("w:fldCharType"), kind)
    r.append(fc)
    return r


def add_page_number(section, cached="2"):
    """Номер страницы в правом верхнем углу листа (поле PAGE, шрифт 14 пт)."""
    section.header_distance = Cm(0.6)
    el = section.header._element
    p = el.find(qn("w:p"))
    pPr = p.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p.insert(0, pPr)
    for old in pPr.findall(qn("w:jc")):
        pPr.remove(old)
    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), "right")
    style_rpr = pPr.find(qn("w:rPr"))
    if style_rpr is not None:
        style_rpr.addprevious(jc)
    else:
        pPr.append(jc)
    for r in (_fld("begin"), _num_run(instr=" PAGE  \\* MERGEFORMAT "),
              _fld("separate"), _num_run(text=cached), _fld("end")):
        p.append(r)


# ------------------------------------------------------------------ таблицы и рисунки


def add_table(doc, spec):
    add_par(doc, spec["caption"], size=12, align=WD_ALIGN_PARAGRAPH.LEFT, first_indent=0,
            before=6, after=3, spacing=1.0)
    table = doc.add_table(rows=1 + len(spec["rows"]), cols=len(spec["cols"]))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    # фиксированные ширины столбцов (сумма = ширина полосы набора 17,5 см)
    widths = spec.get("widths") or [17.5 / len(spec["cols"])] * len(spec["cols"])
    grid = table._tbl.find(qn("w:tblGrid"))
    for gc, w in zip(grid.findall(qn("w:gridCol")), widths):
        gc.set(qn("w:w"), str(int(round(w * 567))))
    for row in table.rows:
        for cell, w in zip(row.cells, widths):
            cell.width = Cm(w)
    for j, name in enumerate(spec["cols"]):
        cell = table.cell(0, j)
        cell.text = ""
        p = cell.paragraphs[0]
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        set_font(p.add_run(name), size=12, bold=True)
    for i, row in enumerate(spec["rows"], start=1):
        for j, value in enumerate(row):
            cell = table.cell(i, j)
            cell.text = ""
            p = cell.paragraphs[0]
            p.paragraph_format.alignment = (WD_ALIGN_PARAGRAPH.CENTER if j == 0
                                            else WD_ALIGN_PARAGRAPH.LEFT)
            p.paragraph_format.line_spacing = 1.0
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            set_font(p.add_run(value), size=12)
    add_par(doc, "", size=6, spacing=1.0, first_indent=0)


def add_figure(doc, path, caption, width_cm):
    p = doc.add_paragraph()
    fmt_par(p, WD_ALIGN_PARAGRAPH.CENTER, first_indent=0, before=6, after=0, spacing=1.0)
    p.paragraph_format.keep_with_next = True
    p.add_run().add_picture(os.path.join(BASE, path), width=Cm(width_cm))
    add_par(doc, caption, size=14, align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=0,
            after=6, spacing=1.0)


# ------------------------------------------------------------------ сборка


def main():
    toc, total_pages = estimate(C.BLOCKS, start_page=3)

    doc = Document(SRC)
    body = doc.element.body
    brk1, brk2 = split_breaks(body)

    # --- титульный лист: меняем только ФИО
    title_pars = []
    for p in body.findall(qn("w:p")):
        if p is brk1:
            break
        title_pars.append(p)
    replace_in_runs(title_pars, "Рыбьяков", "Клячин")
    replace_in_runs(title_pars, "Максим Евгеньевич", "Кирилл Евгеньевич")
    replace_in_runs(title_pars, "Обучающейся", "Обучающийся")

    # --- удаляем старое содержание и старый текст
    children = list(body)
    i1, i2 = children.index(brk1), children.index(brk2)
    for c in children[i1 + 1:]:
        if c is brk2 or c.tag == qn("w:sectPr"):
            continue
        body.remove(c)

    # --- новое содержание (перед вторым разрывом раздела)
    titled = add_par(doc, "СОДЕРЖАНИЕ", size=14, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
                     first_indent=0, after=12, spacing=1.5)
    brk2.addprevious(titled._p)
    for title, page in toc:
        p = doc.add_paragraph()
        fmt_par(p, WD_ALIGN_PARAGRAPH.LEFT, first_indent=0)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(17.4), WD_TAB_ALIGNMENT.RIGHT,
                                                  WD_TAB_LEADER.DOTS)
        set_font(p.add_run(title), size=14)
        set_font(p.add_run("\t" + str(page)), size=14)
        brk2.addprevious(p._p)

    # --- основной текст
    for block in C.BLOCKS:
        kind = block[0]
        if kind == "title_h1":
            add_par(doc, block[1], size=14, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
                    first_indent=0, before=6, after=6)
        elif kind == "h1":
            add_par(doc, block[1], size=14, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
                    first_indent=0, before=6, after=6)
        elif kind == "h2":
            add_par(doc, block[1], size=14, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT,
                    first_indent=1.25, before=6, after=0)
        elif kind == "p":
            add_par(doc, block[1], size=14)
        elif kind == "li":
            add_par(doc, "– " + block[1], size=14, first_indent=1.25)
        elif kind == "src":
            n = sum(1 for b in C.BLOCKS[:C.BLOCKS.index(block)] if b[0] == "src") + 1
            p = doc.add_paragraph()
            fmt_par(p, WD_ALIGN_PARAGRAPH.JUSTIFY, first_indent=1.25)
            set_font(p.add_run(f"{n}. {block[1]}"), size=14)
        elif kind == "fig":
            add_figure(doc, block[1], block[2], block[3])
        elif kind == "table":
            add_table(doc, block[1])

    # --- каждый раздел начинается с новой страницы (рисунки стоят в конце разделов)
    heads = [b[1] for b in C.BLOCKS if b[0] in ("h1", "title_h1")]
    first_done = False
    for p in doc.paragraphs:
        if p.text in heads:
            if not first_done:
                first_done = True          # «Введение» и так с новой страницы
                continue
            p.paragraph_format.page_break_before = True

    # --- штамп: ФИО, номер работы, тема, число листов
    edit_header(doc.sections[1], [(OLD_SHORT, NEW_SHORT), (OLD_NUM, NEW_NUM),
                                  (OLD_TOPIC_1, NEW_TOPIC_1), (OLD_TOPIC_2, NEW_TOPIC_2),
                                  ("11", str(total_pages - 1))])
    edit_header(doc.sections[2], [(OLD_NUM, NEW_NUM)])
    for sec in (doc.sections[1], doc.sections[2]):
        fix_frame(sec)
        # текст строго внутри рамки: по 5 мм от её линий (рамка 20/5/5/5 мм)
        sec.left_margin = Cm(2.5)
        sec.right_margin = Cm(1.0)
        sec.top_margin = Cm(1.2)
        sec.bottom_margin = Cm(5.0)

    # --- номер страницы в правом верхнем углу (титул без номера: у него нет колонтитула)
    add_page_number(doc.sections[1], "2")
    add_page_number(doc.sections[2], "3")

    for t in doc.sections[1].header._element.iter(qn("w:t")):
        if t.text in (NEW_TOPIC_1, NEW_TOPIC_2):
            run = t.getparent()
            rpr = run.find(qn("w:rPr"))
            if rpr is not None:
                for sz in rpr.findall(qn("w:sz")):
                    sz.set(qn("w:val"), "20")
                for sz in rpr.findall(qn("w:szCs")):
                    sz.set(qn("w:val"), "20")

    doc.save(OUT)
    print("сохранено:", OUT)
    print("страниц (оценка):", total_pages, "| листов без титула:", total_pages - 1)
    for t, p in toc:
        print(f"   {p:>3}  {t}")


if __name__ == "__main__":
    main()
