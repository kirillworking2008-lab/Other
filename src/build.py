# -*- coding: utf-8 -*-
"""Сборка отчёта по учебной практике УП.01 в формате .docx."""
import os
import sys

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src import content as C                     # noqa: E402
from src.layout_estimate import estimate        # noqa: E402

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FONT = "Times New Roman"

STUDENT = "Клячин Кирилл Евгеньевич"
STUDENT_SHORT = "Клячин К.Е."
TEACHER_SHORT = "Марковских А.В."
TEACHER = "Марковских Андрей Викторович"
DOC_NUMBER = "УП.01.-15.01.37-02403-2026"

# ---------------------------------------------------------------- вспомогательные


def set_font(run, size=14, bold=False, italic=False):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = RGBColor(0, 0, 0)
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), FONT)


def para(doc_or_cell, text="", size=14, bold=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
         first_indent=0.0, left_indent=0.0, before=0, after=0, spacing=1.5, keep_next=False):
    p = doc_or_cell.add_paragraph()
    pf = p.paragraph_format
    pf.alignment = align
    pf.first_line_indent = Cm(first_indent)
    pf.left_indent = Cm(left_indent)
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = spacing
    pf.keep_with_next = keep_next
    if text:
        set_font(p.add_run(text), size=size, bold=bold)
    return p


def add_field(paragraph, instr, size=9, placeholder="1"):
    """Вставляет поле Word (PAGE/NUMPAGES) с отображаемым значением."""
    run = paragraph.add_run()
    set_font(run, size=size)
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    run._r.append(fld_begin)

    run2 = paragraph.add_run()
    set_font(run2, size=size)
    instr_el = OxmlElement("w:instrText")
    instr_el.set(qn("xml:space"), "preserve")
    instr_el.text = f" {instr} "
    run2._r.append(instr_el)

    run3 = paragraph.add_run()
    set_font(run3, size=size)
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    run3._r.append(fld_sep)

    set_font(paragraph.add_run(placeholder), size=size)

    run5 = paragraph.add_run()
    set_font(run5, size=size)
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run5._r.append(fld_end)


def cell_text(cell, text, size=9, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER):
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    if text:
        set_font(p.add_run(text), size=size, bold=bold)
    return p


TBLPR_ORDER = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
               "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd",
               "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook",
               "tblCaption", "tblDescription", "tblPrChange"]


def tblpr_set(tblPr, el):
    """Вставляет элемент в tblPr, соблюдая порядок элементов схемы."""
    tag = el.tag.split("}")[-1]
    idx = TBLPR_ORDER.index(tag)
    for child in list(tblPr):
        ctag = child.tag.split("}")[-1]
        if ctag in TBLPR_ORDER and TBLPR_ORDER.index(ctag) > idx:
            child.addprevious(el)
            return
    tblPr.append(el)


def set_cell_margins(table, left=0.1, right=0.1, top=0.0, bottom=0.0):
    tblPr = table._tbl.tblPr
    old = tblPr.find(qn("w:tblCellMar"))
    if old is not None:
        tblPr.remove(old)
    mar = OxmlElement("w:tblCellMar")
    for name, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        el = OxmlElement(f"w:{name}")
        el.set(qn("w:w"), str(int(val * 567)))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tblpr_set(tblPr, mar)


def set_table_width(table, width_cm, indent_cm=0.0):
    tblPr = table._tbl.tblPr
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:type"), "dxa")
    tblW.set(qn("w:w"), str(int(width_cm * 567)))
    old = tblPr.find(qn("w:tblInd"))
    if old is not None:
        tblPr.remove(old)
    ind = OxmlElement("w:tblInd")
    ind.set(qn("w:type"), "dxa")
    ind.set(qn("w:w"), str(int(indent_cm * 567)))
    tblpr_set(tblPr, ind)


def set_col_widths(table, widths_cm):
    table.autofit = False
    tblPr = table._tbl.tblPr
    old = tblPr.find(qn("w:tblLayout"))
    if old is not None:
        tblPr.remove(old)
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblpr_set(tblPr, layout)
    grid = table._tbl.find(qn("w:tblGrid"))
    if grid is not None:
        for gc, w in zip(grid.findall(qn("w:gridCol")), widths_cm):
            gc.set(qn("w:w"), str(int(w * 567)))
    for row in table.rows:
        for cell, w in zip(row.cells, widths_cm):
            cell.width = Cm(w)


def set_row_height(row, height_cm, exact=True):
    trPr = row._tr.get_or_add_trPr()
    trH = OxmlElement("w:trHeight")
    trH.set(qn("w:val"), str(int(height_cm * 567)))
    trH.set(qn("w:hRule"), "exact" if exact else "atLeast")
    trPr.append(trH)


def add_page_borders(section, left_pt=57, other_pt=14):
    sectPr = section._sectPr
    pgBorders = OxmlElement("w:pgBorders")
    pgBorders.set(qn("w:offsetFrom"), "page")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "12")
        el.set(qn("w:space"), str(left_pt if edge == "left" else other_pt))
        el.set(qn("w:color"), "000000")
        pgBorders.append(el)
    # порядок элементов в sectPr: pgBorders идёт после pgMar/pgSz и до docGrid
    successors = ["w:lnNumType", "w:pgNumType", "w:cols", "w:formProt", "w:vAlign",
                  "w:noEndnote", "w:titlePg", "w:textDirection", "w:bidi", "w:rtlGutter",
                  "w:docGrid", "w:printerSettings"]
    anchor = None
    for tag in successors:
        found = sectPr.find(qn(tag))
        if found is not None:
            anchor = found
            break
    if anchor is not None:
        anchor.addprevious(pgBorders)
    else:
        sectPr.append(pgBorders)


def setup_page(section, left, right, top, bottom, footer=1.0, header=1.0):
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(left)
    section.right_margin = Cm(right)
    section.top_margin = Cm(top)
    section.bottom_margin = Cm(bottom)
    section.footer_distance = Cm(footer)
    section.header_distance = Cm(header)


# ---------------------------------------------------------------- штамп (основная надпись)


def build_stamp(section):
    footer = section.footer
    footer.is_linked_to_previous = False
    # убираем содержимое по умолчанию
    for p in list(footer.paragraphs):
        p._element.getparent().remove(p._element)

    top_par = footer.add_paragraph()
    top_par.paragraph_format.space_before = Pt(0)
    top_par.paragraph_format.space_after = Pt(0)
    top_par.paragraph_format.line_spacing = Pt(2)
    set_font(top_par.add_run(""), size=2)

    widths = [1.8, 4.4, 7.2, 5.1]      # 18,5 см — по ширине рамки
    table = footer.add_table(rows=4, cols=4, width=Cm(18.5))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    set_table_width(table, 18.5, indent_cm=-0.5)
    set_col_widths(table, widths)
    set_cell_margins(table, 0.1, 0.1, 0.0, 0.0)
    for row in table.rows:
        set_row_height(row, 0.5)

    labels = ["Разраб.", "Пров.", "Н. контр.", "Утв."]
    names = [STUDENT_SHORT, TEACHER_SHORT, "", ""]
    for i in range(4):
        cell_text(table.cell(i, 0), labels[i], size=8)
        cell_text(table.cell(i, 1), names[i], size=8)

    # номер документа — объединённая ячейка на 4 строки
    num_cell = table.cell(0, 2).merge(table.cell(3, 2))
    num_cell.text = ""
    p1 = num_cell.paragraphs[0]
    p1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.paragraph_format.space_before = Pt(0)
    p1.paragraph_format.space_after = Pt(0)
    p1.paragraph_format.line_spacing = 1.0
    set_font(p1.add_run(DOC_NUMBER), size=9, bold=True)
    p2 = num_cell.add_paragraph()
    p2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after = Pt(0)
    p2.paragraph_format.line_spacing = 1.0
    set_font(p2.add_run("Отчет по учебной практике"), size=8)

    # «Лист» с номером страницы
    list_cell = table.cell(0, 3).merge(table.cell(1, 3))
    list_cell.text = ""
    pl = list_cell.paragraphs[0]
    pl.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pl.paragraph_format.space_before = Pt(0)
    pl.paragraph_format.space_after = Pt(0)
    pl.paragraph_format.line_spacing = 1.0
    set_font(pl.add_run("Лист "), size=8)
    add_field(pl, "PAGE", size=9, placeholder="2")

    # «Листов» с общим числом страниц
    pages_cell = table.cell(2, 3).merge(table.cell(3, 3))
    pages_cell.text = ""
    pp = pages_cell.paragraphs[0]
    pp.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pp.paragraph_format.space_before = Pt(0)
    pp.paragraph_format.space_after = Pt(0)
    pp.paragraph_format.line_spacing = 1.0
    set_font(pp.add_run("Листов "), size=8)
    add_field(pp, "NUMPAGES", size=9, placeholder="15")

    bottom_par = footer.add_paragraph()
    bottom_par.paragraph_format.space_before = Pt(0)
    bottom_par.paragraph_format.space_after = Pt(0)
    bottom_par.paragraph_format.line_spacing = Pt(2)
    set_font(bottom_par.add_run(""), size=2)


# ---------------------------------------------------------------- титульный лист


def build_title(doc):
    para(doc, "МИНИСТЕРСТВО НАУКИ И ВЫСШЕГО ОБРАЗОВАНИЯ РОССИЙСКОЙ ФЕДЕРАЦИИ",
         size=13, align=WD_ALIGN_PARAGRAPH.CENTER, spacing=1.0)
    para(doc, "ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ "
              "ВЫСШЕГО ОБРАЗОВАНИЯ", size=13, align=WD_ALIGN_PARAGRAPH.CENTER, spacing=1.0)
    para(doc, "«ТЮМЕНСКИЙ ИНДУСТРИАЛЬНЫЙ УНИВЕРСИТЕТ»", size=13, bold=True,
         align=WD_ALIGN_PARAGRAPH.CENTER, spacing=1.0)
    para(doc, "", size=10, spacing=1.0)
    para(doc, "МНОГОПРОФИЛЬНЫЙ КОЛЛЕДЖ", size=13, bold=True,
         align=WD_ALIGN_PARAGRAPH.CENTER, spacing=1.0)

    for _ in range(5):
        para(doc, "", size=12, spacing=1.0)

    para(doc, "ОТЧЕТ", size=20, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, spacing=1.0)
    para(doc, "ПО УП.01 УЧЕБНОЙ ПРАКТИКЕ", size=16, bold=True,
         align=WD_ALIGN_PARAGRAPH.CENTER, spacing=1.0, after=12)

    para(doc, "в ______Многопрофильном колледже_______", size=14,
         align=WD_ALIGN_PARAGRAPH.CENTER, spacing=1.0)
    para(doc, "(наименование организации/предприятия)", size=10,
         align=WD_ALIGN_PARAGRAPH.CENTER, spacing=1.0, after=18)

    para(doc, f"Обучающийся ____________{STUDENT}_______________", size=14, spacing=1.0)
    para(doc, "Курса_1_", size=14, spacing=1.0)
    para(doc, "Группы_КИПр-26-(11)-1_", size=14, spacing=1.0)
    para(doc, "Профессия 15.01.37 Слесарь-наладчик контрольно-измерительных приборов "
              "и автоматики", size=14, spacing=1.0, after=12)
    para(doc, "В период     с   «09» ноября 2026г. по  «21» ноября 2026г.", size=14,
         spacing=1.0, after=12)
    para(doc, "РУКОВОДИТЕЛИ:", size=14, spacing=1.0, after=6)
    para(doc, f"ОТ УНИВЕРСИТЕТА __________________  {TEACHER}", size=14, spacing=1.0)

    for _ in range(6):
        para(doc, "", size=12, spacing=1.0)

    para(doc, "Тюмень", size=14, align=WD_ALIGN_PARAGRAPH.CENTER, spacing=1.0)
    para(doc, "", size=14, align=WD_ALIGN_PARAGRAPH.CENTER, spacing=1.0)
    para(doc, "2026г.", size=14, align=WD_ALIGN_PARAGRAPH.CENTER, spacing=1.0)


# ---------------------------------------------------------------- содержание


def fld_char(kind):
    el = OxmlElement("w:fldChar")
    el.set(qn("w:fldCharType"), kind)
    return el


def build_toc(doc, entries):
    para(doc, "СОДЕРЖАНИЕ", size=14, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
         after=12, keep_next=True)

    first = True
    for i, (title, page) in enumerate(entries):
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf.first_line_indent = Cm(0)
        pf.left_indent = Cm(0)
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = 1.5
        pf.tab_stops.add_tab_stop(Cm(17.4), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)

        if first:
            run = p.add_run()
            set_font(run, size=14)
            run._r.append(fld_char("begin"))
            run2 = p.add_run()
            set_font(run2, size=14)
            instr = OxmlElement("w:instrText")
            instr.set(qn("xml:space"), "preserve")
            instr.text = r' TOC \o "1-1" \h \z \u '
            run2._r.append(instr)
            run3 = p.add_run()
            set_font(run3, size=14)
            run3._r.append(fld_char("separate"))
            first = False

        set_font(p.add_run(title), size=14)
        set_font(p.add_run("\t"), size=14)
        set_font(p.add_run(str(page)), size=14)

        if i == len(entries) - 1:
            end_run = p.add_run()
            set_font(end_run, size=14)
            end_run._r.append(fld_char("end"))


# ---------------------------------------------------------------- таблицы и рисунки


def add_table(doc, spec):
    para(doc, spec["caption"], size=12, align=WD_ALIGN_PARAGRAPH.LEFT, before=6, after=3,
         spacing=1.0, keep_next=True)
    rows = spec["rows"]
    cols = spec["cols"]
    table = doc.add_table(rows=1 + len(rows), cols=len(cols))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_col_widths(table, spec["widths"])
    set_table_width(table, sum(spec["widths"]))
    set_cell_margins(table, 0.15, 0.15, 0.05, 0.05)

    # шапка повторяется при переходе таблицы на следующую страницу
    trPr = table.rows[0]._tr.get_or_add_trPr()
    header = OxmlElement("w:tblHeader")
    trPr.append(header)

    for j, name in enumerate(cols):
        cell_text(table.cell(0, j), name, size=12, bold=True)
    for i, row in enumerate(rows, start=1):
        for j, value in enumerate(row):
            cell_text(table.cell(i, j), value, size=12,
                      align=WD_ALIGN_PARAGRAPH.CENTER if j == 0 else WD_ALIGN_PARAGRAPH.LEFT)

    para(doc, "", size=6, spacing=1.0, after=0)


def add_figure(doc, path, caption, width_cm):
    p = doc.add_paragraph()
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.keep_with_next = True
    p.add_run().add_picture(os.path.join(BASE, path), width=Cm(width_cm))
    para(doc, caption, size=12, align=WD_ALIGN_PARAGRAPH.CENTER, after=8, spacing=1.0)


# ---------------------------------------------------------------- документ


def main():
    toc_entries, total_pages = estimate(C.BLOCKS)

    doc = Document()

    # базовые стили
    style = doc.styles["Normal"]
    style.font.name = FONT
    style.font.size = Pt(14)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    style.element.rPr.rFonts.set(qn("w:cs"), FONT)
    pf = style.paragraph_format
    pf.line_spacing = 1.5
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.first_line_indent = Cm(1.25)
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    for name, size in (("Heading 1", 14), ("Heading 2", 14), ("Heading 3", 14)):
        st = doc.styles[name]
        st.font.name = FONT
        st.font.size = Pt(size)
        st.font.bold = True
        st.font.italic = False
        st.font.color.rgb = RGBColor(0, 0, 0)
        rpr = st.element.get_or_add_rPr()
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            rfonts = OxmlElement("w:rFonts")
            rpr.insert(0, rfonts)
        for attr in ("w:ascii", "w:hAnsi", "w:csw", "w:eastAsia"):
            if attr in ("w:ascii", "w:hAnsi", "w:eastAsia"):
                rfonts.set(qn(attr), FONT)
        st.paragraph_format.first_line_indent = Cm(0)
        st.paragraph_format.line_spacing = 1.5
        st.paragraph_format.space_before = Pt(6)
        st.paragraph_format.space_after = Pt(6)
        st.paragraph_format.keep_with_next = True

    # ---------------- титульный лист (раздел 1)
    sec1 = doc.sections[0]
    setup_page(sec1, 3.0, 1.5, 2.0, 2.0, footer=1.0, header=1.0)
    build_title(doc)

    # ---------------- основной текст (раздел 2)
    sec2 = doc.add_section(WD_SECTION.NEW_PAGE)
    setup_page(sec2, 2.5, 1.0, 1.5, 2.5, footer=0.5, header=0.5)
    add_page_borders(sec2)
    build_stamp(sec2)

    # ---------------- содержание
    build_toc(doc, toc_entries)
    doc.add_page_break()

    # ---------------- основной текст
    for block in C.BLOCKS:
        kind = block[0]

        if kind == "title_h1":
            p = doc.add_paragraph(style="Heading 1")
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_font(p.add_run(block[1]), size=14, bold=True)
        elif kind == "h1":
            p = doc.add_paragraph(style="Heading 1")
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            set_font(p.add_run(block[1]), size=14, bold=True)
        elif kind == "h2":
            p = doc.add_paragraph(style="Heading 2")
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.first_line_indent = Cm(1.25)
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            set_font(p.add_run(block[1]), size=14, bold=True)
        elif kind == "p":
            para(doc, block[1], first_indent=1.25)
        elif kind == "li":
            para(doc, "– " + block[1], left_indent=1.25)
        elif kind == "src":
            n = sum(1 for b in C.BLOCKS[:C.BLOCKS.index(block)] if b[0] == "src") + 1
            p = doc.add_paragraph()
            pf2 = p.paragraph_format
            pf2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            pf2.left_indent = Cm(0.75)
            pf2.first_line_indent = Cm(-0.75)
            pf2.space_before = Pt(0)
            pf2.space_after = Pt(0)
            pf2.line_spacing = 1.5
            pf2.tab_stops.add_tab_stop(Cm(0.75), WD_TAB_ALIGNMENT.LEFT)
            set_font(p.add_run(f"{n}.\t"), size=14)
            set_font(p.add_run(block[1]), size=14)
        elif kind == "fig":
            add_figure(doc, block[1], block[2], block[3])
        elif kind == "table":
            add_table(doc, block[1])

    out = os.path.join(BASE, "Отчет_УП.01_Клячин_К_Е.docx")
    doc.save(out)
    print("сохранено:", out)
    print("оценка объёма: страниц всего", total_pages, "| листов без титула", total_pages - 1)
    for name, pg in toc_entries:
        print(f"   {pg:>3}  {name}")


if __name__ == "__main__":
    main()
