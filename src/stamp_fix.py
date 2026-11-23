"""Правка штампа (рамка по ГОСТ 2.104) в колонтитулах отчёта.

Зачем: номер листа в штампе задан полем PAGE внутри группы фигур DrawingML.
Word не пересчитывает поля внутри группы и показывает одно и то же число на всех
листах (в отчёте Клячина застывало «3»). Поэтому фигуру с номером убираем, а номер
ставим отдельной надписью в колонтитуле, вне группы: такое поле Word обновляет на
каждой странице. Заодно подписи «Изм.», «Лист», «Подпись» в форме 2 выравниваются
по своим ячейкам — в исходном шаблоне они были уже ячейки и обрезались.
"""

from docx.oxml import parse_xml

_MM = 36000  # EMU в миллиметре

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
_V = "{urn:schemas-microsoft-com:vml}"
_MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
_WPS = {"wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"}

# подпись -> (x, ширина) в единицах группы (20000 единиц = 185 мм по ширине)
_LABEL_FIX = {
    # header: {подпись: (x_old_min, x_old_max, новый x, новая ширина)}
    "header1": {
        "Изм.": (0, 900, 10, 983),
        "Лист": (900, 1200, 1000, 1260),
        "Подпись": (4500, 5500, 5119, 1438),
    },
    "header2": {
        "Изм.": (0, 900, 10, 1083),
        "Лист": (1000, 1300, 1100, 1080),
        "№ документа": (2000, 2600, 2190, 2720),
        "Подпись": (4500, 5500, 4919, 1638),
        "Дата": (6400, 6800, 6600, 1050),
    },
}

# x, y, ширина, высота в мм — ячейка «Лист» штампа
_NUM_BOX = {
    "header1": (166.6, 271.7, 15.2, 5.0),
    "header2": (195.0, 277.0, 9.7, 14.4),
}

# диапазон x ячейки «Лист» в единицах группы
_NUM_CELL = {
    "header1": (15848, 17487),
    "header2": (18905, 19977),
}

# номер первой страницы секции — кешированное значение поля
_NUM_FIRST = {"header1": "2", "header2": "3"}

_ANCHOR = '''<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
<w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="{did}"
 behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>
 <wp:positionH relativeFrom="page"><wp:posOffset>{x}</wp:posOffset></wp:positionH>
 <wp:positionV relativeFrom="page"><wp:posOffset>{y}</wp:posOffset></wp:positionV>
 <wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>
 <wp:docPr id="{did}" name="НомерЛиста"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>
 <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
  <wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/>
   <a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
   <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>
   <wps:txbx><w:txbxContent><w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
    <w:jc w:val="center"/><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="20"/><w:lang w:val="ru-RU"/></w:rPr></w:pPr>
    <w:fldSimple w:instr=" PAGE  \\* MERGEFORMAT " w:dirty="true"><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="20"/><w:lang w:val="ru-RU"/></w:rPr><w:t>{text}</w:t></w:r></w:fldSimple>
   </w:p></w:txbxContent></wps:txbx>
   <wps:bodyPr rot="0" vert="horz" wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr" anchorCtr="0"><a:noAutofit/></wps:bodyPr>
  </wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>'''


def _strip_spaces(shape):
    ts = list(shape.iter(_W + "t"))
    if not ts:
        return
    joined = "".join(t.text or "" for t in ts).strip()
    ts[0].text = joined
    for t in ts[1:]:
        t.text = ""


def _vml_left(el):
    try:
        return int((el.get("style") or "").split("left:")[1].split(";")[0])
    except (IndexError, ValueError):
        return -1


def _resize_vml(el, x, width):
    st = el.get("style") or ""
    try:
        st = st.replace("left:%d;" % _vml_left(el), "left:%d;" % x)
        st = st.replace("width:%d;" % int(st.split("width:")[1].split(";")[0]),
                        "width:%d;" % width)
    except (IndexError, ValueError):
        return
    el.set("style", st)


def rebuild_page_number(doc):
    for part in doc.part.package.parts:
        tag = str(part.partname).split("/")[-1].replace(".xml", "")
        if tag not in _NUM_BOX:
            continue
        root = part.element
        lo, hi = _NUM_CELL[tag]

        # слой DrawingML — его читает Word
        for shape in list(root.findall(".//wps:wsp", _WPS)):
            txt = "".join(t.text or "" for t in shape.iter(_W + "t")).strip()
            off = shape.find(".//" + _A + "off")
            if off is None or int(off.get("y")) < 16000:
                continue
            if txt.isdigit():
                if lo <= int(off.get("x")) <= hi:
                    shape.getparent().remove(shape)
                continue
            fix = _LABEL_FIX[tag].get(txt)
            if fix and fix[0] <= int(off.get("x")) <= fix[1]:
                _, _, x, cx = fix
                off.set("x", str(x))
                shape.find(".//" + _A + "ext").set("cx", str(cx))
                bp = shape.find(".//wps:bodyPr", _WPS)
                if bp is not None:
                    for attr in ("lIns", "tIns", "rIns", "bIns"):
                        bp.set(attr, "0")
                    bp.set("wrap", "none")
                _strip_spaces(shape)

        # слой VML — резервный, его показывают конвертеры и предпросмотр
        fb = root.find(".//" + _MC + "Fallback")
        if fb is not None:
            for el in list(fb.iter()):
                if not el.tag.startswith(_V) or el.tag.split("}")[1] not in ("rect", "shape"):
                    continue
                txt = "".join(t.text or "" for t in el.iter(_W + "t")).strip()
                if not txt:
                    continue
                if txt.isdigit():
                    if lo <= _vml_left(el) <= hi:
                        el.getparent().remove(el)
                    continue
                fix = _LABEL_FIX[tag].get(txt)
                if fix and fix[0] <= _vml_left(el) <= fix[1]:
                    _resize_vml(el, fix[2], fix[3])
                    for tb in el.iter(_V + "textbox"):
                        tb.set("inset", "0,0,0,0")
                    _strip_spaces(el)

        x, y, w, h = _NUM_BOX[tag]
        did = 9100 if tag == "header2" else 9101
        root.find(".//" + _W + "p").append(parse_xml(_ANCHOR.format(
            did=did, x=int(x * _MM), y=int(y * _MM),
            cx=int(w * _MM), cy=int(h * _MM), text=_NUM_FIRST[tag])))
