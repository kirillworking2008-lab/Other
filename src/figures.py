# -*- coding: utf-8 -*-
"""Построение схем (рисунков) для отчёта по учебной практике УП.01."""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch, Rectangle, Circle

plt.rcParams["font.family"] = "DejaVu Sans"

FS = 11          # размер шрифта подписей на схемах
LW = 1.6


def canvas(w, h, scale=1.0):
    fig = plt.figure(figsize=(w * scale, h * scale), dpi=300)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, w)
    ax.set_ylim(0, h)
    ax.set_aspect("equal")
    ax.axis("off")
    return fig, ax


def box(ax, x, y, w, h, text, fs=FS, round_=0.08):
    p = FancyBboxPatch((x, y), w, h, boxstyle=f"round,pad=0,rounding_size={round_}",
                       linewidth=LW, edgecolor="black", facecolor="white")
    ax.add_patch(p)
    ax.text(x + w / 2, y + h / 2, text, ha="center", va="center", fontsize=fs, wrap=True)


def arrow(ax, xy1, xy2, style="-|>", lw=LW, ls="-"):
    a = FancyArrowPatch(xy1, xy2, arrowstyle=style, mutation_scale=14,
                        linewidth=lw, linestyle=ls, color="black", shrinkA=0, shrinkB=0)
    ax.add_patch(a)


# ---------------------------------------------------------------- Рисунок 1
def fig1(path):
    """Структурная схема последовательности проверки низковольтной цепи."""
    W, H = 7.6, 3.9
    fig, ax = canvas(W, H, scale=2.2)
    bw, bh = 3.35, 0.72
    x1, x2 = 0.25, W - bw - 0.25
    ys = [2.95, 2.05, 1.15, 0.25]
    texts_left = [
        "1. Изучение проектной схемы\nи маркировки проводников",
        "2. Внешний осмотр\nи проверка качества монтажа",
        "3. Прозвонка и проверка\nцелостности цепей",
        "4. Измерение сопротивления\nизоляции",
    ]
    texts_right = [
        "5. Проверка полярности\nи фазировки",
        "6. Функциональные испытания\nпод напряжением (УЗО, автоматы)",
        "7. Устранение выявленных\nдефектов монтажа",
        "8. Оформление протокола\nи сдача работы",
    ]
    for y, t in zip(ys, texts_left):
        box(ax, x1, y, bw, bh, t)
    for y, t in zip(ys[::-1], texts_right):
        box(ax, x2, y, bw, bh, t)

    # стрелки сверху вниз в левом столбце
    for i in range(3):
        arrow(ax, (x1 + bw / 2, ys[i]), (x1 + bw / 2, ys[i + 1] + bh))
    # переход из левого нижнего блока в правый нижний
    by = ys[3] + bh / 2
    ax.plot([x1 + bw, (x1 + bw + x2) / 2, (x1 + bw + x2) / 2, x2],
            [by, by, by, by], color="black", lw=LW)
    arrow(ax, ((x1 + bw + x2) / 2, by), (x2, by))
    # стрелки снизу вверх в правом столбце
    for i in range(3):
        arrow(ax, (x2 + bw / 2, ys[i + 1] + bh), (x2 + bw / 2, ys[i]))

    fig.savefig(path, dpi=300, facecolor="white")
    plt.close(fig)


# ---------------------------------------------------------------- Рисунок 2
def fig2(path):
    """Схема измерения сопротивления изоляции мегаомметром."""
    W, H = 7.4, 3.5
    fig, ax = canvas(W, H, scale=2.2)
    x0, x1 = 1.35, 5.55
    yL, yN, yPE = 2.95, 2.15, 1.35

    for y, name, col in ((yL, "L", "black"), (yN, "N", "black"), (yPE, "PE", "black")):
        ax.plot([x0, x1], [y, y], color=col, lw=LW)
        ax.add_patch(Circle((x0, y), 0.055, color="black"))
        ax.add_patch(Circle((x1, y), 0.055, color="black"))
        ax.text(x0 - 0.18, y, name, ha="right", va="center", fontsize=FS)

    # сопротивление изоляции между L и N (условное обозначение)
    xr = 3.1
    step = (yL - yN) / 8.0
    pts_y = [yL - i * step for i in range(9)]
    pts_x = [xr] * 9
    for i in range(1, 8):
        pts_x[i] = xr + (0.16 if i % 2 else -0.16)
    ax.plot([xr, xr], [yL, pts_y[0]], color="black", lw=LW)
    ax.plot(pts_x, pts_y, color="black", lw=LW)
    ax.plot([xr, xr], [pts_y[-1], yN], color="black", lw=LW)
    ax.text(xr - 0.22, (yL + yN) / 2 - 0.02, "Rиз", ha="right", va="center", fontsize=FS)
    ax.text(xr + 0.28, (yL + yN) / 2 - 0.02, "изоляция\nцепи", ha="left", va="center",
            fontsize=FS - 1.5)

    # мегаомметр
    cx, cy, r = 6.35, 2.15, 0.32
    ax.add_patch(Circle((cx, cy), r, facecolor="white", edgecolor="black", lw=LW))
    ax.text(cx, cy, "M", ha="center", va="center", fontsize=FS + 2)
    # вывод к L
    ax.plot([cx - 0.24, cx - 0.24, x1], [cy + 0.21, yL, yL], color="black", lw=LW)
    # вывод к PE
    ax.plot([cx, cx, x1], [cy - r, yPE, yPE], color="black", lw=LW)
    ax.text(cx + 0.45, cy + 0.22, "Мегаомметр\n500 В", ha="left", va="bottom", fontsize=FS - 1.5)

    # знак заземления на PE
    gx = x1
    ax.plot([gx, gx], [yPE, yPE - 0.22], color="black", lw=LW)
    for i, wlen in enumerate((0.34, 0.22, 0.10)):
        yy = yPE - 0.22 - i * 0.09
        ax.plot([gx - wlen / 2, gx + wlen / 2], [yy, yy], color="black", lw=LW)

    ax.text(x0, 0.45, "Проверяемая цепь должна быть отключена,\nпотребители и УЗИП — отсоединены",
            ha="left", va="center", fontsize=FS - 1.5)

    fig.savefig(path, dpi=300, facecolor="white")
    plt.close(fig)


# ---------------------------------------------------------------- Рисунок 3
def fig3(path):
    """Схема прозвонки цепи мультиметром."""
    W, H = 7.2, 2.5
    fig, ax = canvas(W, H, scale=2.2)

    # мультиметр
    mx, my, mw, mh = 0.25, 0.55, 1.95, 1.45
    ax.add_patch(FancyBboxPatch((mx, my), mw, mh, boxstyle="round,pad=0,rounding_size=0.12",
                                facecolor="white", edgecolor="black", lw=LW))
    ax.add_patch(Rectangle((mx + 0.16, my + 0.70), mw - 0.32, 0.50,
                           facecolor="white", edgecolor="black", lw=1.0))
    ax.text(mx + mw / 2, my + 0.95, "≈ 0 Ом", ha="center", va="center", fontsize=FS + 2)
    ax.text(mx + mw / 2, my + 0.34, "Мультиметр\n(режим прозвонки)", ha="center",
            va="center", fontsize=FS - 2)

    # проверяемая жила
    wx0, wx1, wy = 3.45, 6.75, 1.75
    ax.plot([wx0, wx1], [wy, wy], color="black", lw=3.2)
    for x in (wx0, wx1):
        ax.add_patch(Circle((x, wy), 0.062, color="black"))
    ax.text((wx0 + wx1) / 2, wy + 0.22, "Проверяемая жила (цепь)", ha="center",
            va="bottom", fontsize=FS)

    # щуп 1
    y1 = my + mh - 0.35
    ax.plot([mx + mw, 2.95, 2.95, wx0], [y1, y1, wy, wy], color="black", lw=LW)
    ax.text(3.05, y1 + 0.14, "щуп 1", ha="left", va="bottom", fontsize=FS - 2)
    # щуп 2
    y2 = my + 0.30
    yb = 0.30
    ax.plot([mx + mw, 3.20, 3.20, wx1], [y2, y2, yb, yb], color="black", lw=LW)
    ax.plot([wx1, wx1], [yb, wy], color="black", lw=LW)
    ax.text(3.30, yb - 0.16, "щуп 2", ha="left", va="top", fontsize=FS - 2)

    fig.savefig(path, dpi=300, facecolor="white")
    plt.close(fig)


if __name__ == "__main__":
    import os
    out = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "fig")
    os.makedirs(out, exist_ok=True)
    fig1(os.path.join(out, "ris1.png"))
    fig2(os.path.join(out, "ris2.png"))
    fig3(os.path.join(out, "ris3.png"))
    print("figures done:", out)
