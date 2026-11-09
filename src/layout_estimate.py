# -*- coding: utf-8 -*-
"""Оценка разбивки отчёта на страницы.

Модель повторяет поведение Word: абзацы переносятся построчно (с разрывом между
страницами), рисунки и таблицы целиком.
"""
import math
import os

from PIL import Image, ImageFont

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# ---------------------------------------------------------------- параметры вёрстки (см)
TEXT_W = 17.5          # ширина полосы набора (поля 2,5 / 1,0 см)
TEXT_H = 23.5          # высота полосы набора (поля 1,2 / 5,0 см)
LINE_H = 0.852         # высота строки: 14 пт, интервал 1,5
CHARS_K = 0.885        # поправка ширины DejaVu Serif -> Times New Roman

_font_cache = {}


def _font(size_pt=14):
    if size_pt not in _font_cache:
        path = "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf"
        try:
            _font_cache[size_pt] = ImageFont.truetype(path, int(size_pt * 4))
        except Exception:
            _font_cache[size_pt] = ImageFont.load_default()
    return _font_cache[size_pt]


def text_width_cm(text, size_pt=14):
    f = _font(size_pt)
    return f.getlength(text) / 4.0 * 2.54 / 72.0 * CHARS_K


def wrap_lines(text, size_pt=14, width_cm=TEXT_W, first_indent=1.25, indent=0.0):
    """Число строк, которое займёт абзац."""
    words = text.split()
    lines, cur = 1, ""
    avail = width_cm - first_indent
    for w in words:
        trial = (cur + " " + w).strip()
        if text_width_cm(trial, size_pt) <= avail:
            cur = trial
        else:
            lines += 1
            cur = w
            avail = width_cm - indent
    return lines


class Flow:
    def __init__(self, start_page):
        self.page = start_page
        self.used = 0.0

    def new_page(self):
        self.page += 1
        self.used = 0.0

    def add_lines(self, n):
        """Добавляет n строк с переносом между страницами."""
        left = n
        while left > 0:
            space = TEXT_H - self.used
            fit = int(space // LINE_H + 1e-6)
            if fit <= 0:
                self.new_page()
                continue
            take = min(fit, left)
            self.used += take * LINE_H
            left -= take
            if left > 0:
                self.new_page()

    def add_block(self, h):
        """Добавляет неразрывный блок (рисунок, таблица, заголовок)."""
        if self.used + h > TEXT_H + 1e-6:
            self.new_page()
        self.used += h


def table_height(t):
    h = LINE_H + 0.18                       # заголовок таблицы
    for row in [t["cols"]] + t["rows"]:
        lines = 1
        for cell, w in zip(row, t["widths"]):
            if cell:
                avail = max(0.8, w - 0.4)
                lines = max(lines, math.ceil(text_width_cm(cell, 12) / avail))
        h += lines * 0.44 + 0.14
    return h


def fig_height(path, w_cm):
    with Image.open(path) as im:
        ratio = im.height / im.width
    return w_cm * ratio + 0.35 + LINE_H          # рисунок + отступ + подпись


def estimate(blocks, start_page=3):
    """Возвращает ([(заголовок, страница)], последняя_страница)."""
    fl = Flow(start_page)
    toc = []
    for b in blocks:
        kind = b[0]
        if kind == "pagebreak":
            fl.new_page()
        elif kind in ("h1", "title_h1"):
            n = max(1, wrap_lines(b[1], 14, TEXT_W, 0, 0))
            h = n * LINE_H + 0.42
            if kind in ("h1", "title_h1") and fl.used + h + 2 * LINE_H > TEXT_H:
                fl.new_page()
            fl.add_block(h)
            toc.append((b[1], fl.page))
        elif kind == "h2":
            n = max(1, wrap_lines(b[1], 14, TEXT_W, 1.25))
            h = n * LINE_H + 0.2
            if fl.used + h + 2 * LINE_H > TEXT_H:
                fl.new_page()
            fl.add_block(h)
        elif kind == "p":
            fl.add_lines(wrap_lines(b[1]))
        elif kind == "li":
            fl.add_lines(wrap_lines(b[1], 14, TEXT_W, 0.0, 1.25))
        elif kind == "src":
            fl.add_lines(wrap_lines(b[1], 14, TEXT_W - 0.75, 0.0, 0.0))
        elif kind == "fig":
            fl.add_block(fig_height(os.path.join(BASE, b[1]), b[3]))
        elif kind == "table":
            fl.add_block(table_height(b[1]))
    return toc, fl.page


if __name__ == "__main__":
    from src import content as C
    toc, total = estimate(C.BLOCKS)
    for name, pg in toc:
        print(f"{pg:>3}  {name}")
    print("всего страниц (с титульным):", total, "| листов без титула:", total - 1)
