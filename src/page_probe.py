# -*- coding: utf-8 -*-
"""Перебор вариантов полосы набора: сколько страниц получается."""
import sys

from src import content as C
from src import layout_estimate as L


def run(w, h, show=False):
    L.TEXT_W, L.TEXT_H = w, h
    toc, total = L.estimate(C.BLOCKS)
    if show:
        for name, pg in toc:
            print(f"   {pg:>3}  {name}")
    return total


if __name__ == "__main__":
    for w, h in [(17.5, 23.8), (17.5, 24.8), (17.5, 25.0), (17.6, 25.2),
                 (17.0, 24.0), (18.0, 24.0), (17.5, 26.2)]:
        print(f"W={w} H={h}: страниц {run(w, h)}")
    if len(sys.argv) > 1:
        print("--- детально для", sys.argv[1:])
        run(float(sys.argv[1]), float(sys.argv[2]), show=True)
