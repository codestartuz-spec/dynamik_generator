# -*- coding: utf-8 -*-
"""
Jarayon foizini oynali ilovaga (app.py) yetkazish.

GEN_PROGRESS=1 bo'lsagina ishlaydi (oyna shuni o'rnatadi), shuning uchun generator.bat orqali
oddiy ishga tushirganda jurnalda qo'shimcha qatorlar chiqmaydi.
Qator formati:  @@PROGRESS <foiz>|<bosqich matni>     va yakunda  @@DONE
"""
import os

ENABLED = os.environ.get("GEN_PROGRESS") == "1"
PREFIX = "@@PROGRESS "
DONE = "@@DONE"


def _clamp(x, lo=0.0, hi=1.0):
    return lo if x < lo else hi if x > hi else x


def emit(pct, text=""):
    if ENABLED:
        print("%s%.1f|%s" % (PREFIX, _clamp(float(pct), 0.0, 100.0), text), flush=True)


def done():
    if ENABLED:
        print(DONE, flush=True)


class Span(object):
    """Umumiy 0..100% ning bir oralig'i. at(0..1) shu oraliq ichidagi ulushni foizga aylantiradi."""

    def __init__(self, lo=0.0, hi=100.0):
        self.lo, self.hi = float(lo), float(hi)

    def at(self, frac, text=""):
        emit(self.lo + (self.hi - self.lo) * _clamp(frac), text)

    def sub(self, a, b):
        return Span(self.lo + (self.hi - self.lo) * a, self.lo + (self.hi - self.lo) * b)


def files_fraction(index, total, rows, prev_rows):
    """
    Ko'p fayl ketma-ket o'qilganda umumiy ulush (0..1). Fayl ichidagi qatorlar soni oldindan noma'lum,
    shuning uchun oldingi fayllarning o'rtacha qatorlariga taqqoslanadi (birinchi faylda - silliq taxmin).
    """
    if total <= 0:
        return 0.0
    if prev_rows:
        expected = float(sum(prev_rows)) / len(prev_rows)
        f = rows / expected if expected > 0 else 0.0
    else:
        f = 1.0 - 1.0 / (1.0 + rows / 150000.0)
    return (index + _clamp(f, 0.0, 0.97)) / total
