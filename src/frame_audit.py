# -*- coding: utf-8 -*-
"""Аудит геометрии рамки и штампа (ГОСТ 2.104) в собранном docx."""
import sys
import zipfile

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
WPS = "{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}"
WPG = "{http://schemas.microsoft.com/office/word/2010/wordprocessingGroup}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def cm(v):
    return v / 360000.0


def audit(path):
    z = zipfile.ZipFile(path)
    for f in ("word/header1.xml", "word/header2.xml"):
        root = etree.fromstring(z.read(f))
        print("=" * 12, f)
        ext = root.find(".//" + WP + "extent")
        ex_cm, ey_cm = cm(int(ext.get("cx"))), cm(int(ext.get("cy")))
        ph = root.find(".//" + WP + "positionH/" + WP + "posOffset")
        pv = root.find(".//" + WP + "positionV/" + WP + "posOffset")
        print(f"рамка: x={cm(int(ph.text)):.2f} см  y={cm(int(pv.text)):.2f} см  "
              f"ширина={ex_cm:.2f}  высота={ey_cm:.2f}")
        print(f"      поля до края листа: слева {cm(int(ph.text)) * 10:.1f} мм, "
              f"сверху {cm(int(pv.text)) * 10:.1f} мм, "
              f"справа {(21 - cm(int(ph.text)) - ex_cm) * 10:.1f} мм, "
              f"снизу {(29.7 - cm(int(pv.text)) - ey_cm) * 10:.1f} мм")
        g = root.find(".//" + WPG + "wgp")
        if g is None:
            print("  (группа не найдена)")
            continue
        xf = g.find(WPG + "grpSpPr/" + A + "xfrm")
        che = xf.find(A + "chExt")
        cw, chh = int(che.get("cx")), int(che.get("cy"))
        kx = ex_cm / cw      # см на единицу сетки
        ky = ey_cm / chh
        print(f"  сетка {cw}x{chh}; 1 ед = {kx * 10:.3f} x {ky * 10:.3f} мм")
        rows = []
        for sp in g:
            tag = etree.QName(sp).localname
            if tag not in ("wsp", "grpSp", "pic", "cxnSp"):
                continue
            x = sp.find(".//" + WPS + "spPr/" + A + "xfrm")
            if x is None:
                continue
            o, e = x.find(A + "off"), x.find(A + "ext")
            if o is None or e is None:
                continue
            ox, oy = int(o.get("x")), int(o.get("y"))
            ex, ey = int(e.get("cx")), int(e.get("cy"))
            txt = "".join(t.text or "" for t in sp.iter(W + "t"))[:26]
            rows.append((oy, ox, ox * kx, oy * ky, ex * kx, ey * ky, tag, txt))
        print(f"  дочерних фигур: {len(rows)}")
        for oy, ox, x_cm, y_cm, w_cm, h_cm, tag, txt in sorted(rows):
            print(f"   y={y_cm:6.2f} x={x_cm:6.2f} w={w_cm:5.2f} h={h_cm:5.2f} см  {tag:5s} {txt}")
        # границы по Y: верх/низ фигур
        ys = sorted({round(r[3], 2) for r in rows})
        print("  y-координаты начал фигур (см):", ys[-8:])


if __name__ == "__main__":
    audit(sys.argv[1] if len(sys.argv) > 1 else "Отчет_УП.01_Клячин_К_Е.docx")
