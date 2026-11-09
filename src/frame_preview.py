# -*- coding: utf-8 -*-
"""Предпросмотр рамки/штампа: рисует линии из колонтитула в PNG (для визуальной проверки)."""
import sys
import zipfile

from lxml import etree
from PIL import Image, ImageDraw, ImageFont

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
WPS = "{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}"
WPG = "{http://schemas.microsoft.com/office/word/2010/wordprocessingGroup}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

DPI = 150.0
PX = DPI / 2.54
FONT_PATH = "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf"


def cm(v):
    return v / 360000.0


def shapes(root):
    ext = root.find(".//" + WP + "extent")
    ex, ey = cm(int(ext.get("cx"))), cm(int(ext.get("cy")))
    ph = root.find(".//" + WP + "positionH/" + WP + "posOffset")
    pv = root.find(".//" + WP + "positionV/" + WP + "posOffset")
    ox, oy = cm(int(ph.text)), cm(int(pv.text))
    g = root.find(".//" + WPG + "wgp")
    che = g.find(WPG + "grpSpPr/" + A + "xfrm/" + A + "chExt")
    kx, ky = ex / int(che.get("cx")), ey / int(che.get("cy"))
    out = []
    for sp in g:
        if etree.QName(sp).localname != "wsp":
            continue
        x = sp.find(".//" + WPS + "spPr/" + A + "xfrm")
        if x is None:
            continue
        o, e = x.find(A + "off"), x.find(A + "ext")
        if o is None or e is None:
            continue
        out.append((
            ox + int(o.get("x")) * kx, oy + int(o.get("y")) * ky,
            int(e.get("cx")) * kx, int(e.get("cy")) * ky,
            "".join(t.text or "" for t in sp.iter(W + "t")).strip(),
        ))
    return out


def render(path, header, out_png, title="", margins=(2.0, 1.5, 2.0, 3.0)):
    """margins: (top, right, bottom, left) в см."""
    z = zipfile.ZipFile(path)
    root = etree.fromstring(z.read(header))
    img = Image.new("RGB", (int(21 * PX), int(29.7 * PX)), "white")
    d = ImageDraw.Draw(img)
    f10 = ImageFont.truetype(FONT_PATH, 13)
    f8 = ImageFont.truetype(FONT_PATH, 9)
    top, right, bottom, left = margins
    x0, x1 = left * PX, (21 - right) * PX
    y0, y1 = top * PX, (29.7 - bottom) * PX
    lines = []
    for X, Y, Wd, Ht, txt in shapes(root):
        if Ht < 0.05 and Wd > 0.05:
            lines.append(("h", X * PX, Y * PX, (X + Wd) * PX, Y * PX))
        elif Wd < 0.05 and Ht > 0.05:
            lines.append(("v", X * PX, Y * PX, X * PX, (Y + Ht) * PX))
        elif Wd > 18.0 and Ht > 28.0:
            lines.append(("frame", X * PX, Y * PX, (X + Wd) * PX, (Y + Ht) * PX))
        elif txt:
            d.text((X * PX + 2, Y * PX + 1), txt[:30], fill=(20, 20, 20), font=f8)
    for kind, a, b, c, e in lines:
        if kind == "frame":
            d.rectangle([a, b, c, e], outline=(0, 0, 0), width=3)
        elif kind == "h":
            d.line([a, b, c, e], fill=(0, 0, 0), width=2)
        else:
            d.line([a, b, c, e], fill=(0, 0, 0), width=2)
    # область набора текста
    d.rectangle([x0, y0, x1, y1], outline=(120, 160, 255))
    d.text((8, 8), title, fill=(150, 0, 0), font=f10)
    img.save(out_png)
    print("сохранено:", out_png)


def margins_of(path, idx):
    """Поля страницы нужной секции (в см): (верх, право, низ, лево)."""
    z = zipfile.ZipFile(path)
    doc = etree.fromstring(z.read("word/document.xml"))
    sects = list(doc.iter(W + "sectPr"))
    m = sects[idx].find(W + "pgMar")
    return tuple(int(m.get(W + k)) / 567.0 for k in ("top", "right", "bottom", "left"))


if __name__ == "__main__":
    src = sys.argv[1]
    render(src, "word/header1.xml", "/home/user/Other/frame_header1.png",
           "форма 2 — лист «СОДЕРЖАНИЕ»", margins_of(src, 1))
    render(src, "word/header2.xml", "/home/user/Other/frame_header2.png",
           "форма 2а — листы тела", margins_of(src, 2))
