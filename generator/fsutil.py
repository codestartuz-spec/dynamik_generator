# -*- coding: utf-8 -*-
"""Fayl yo'llari bilan ishlash yordamchilari."""
import re


def _natural(name):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", name)]


def list_xlsx(folder):
    """Papkadagi barcha .xlsx fayllar (tabiiy tartibda). Excel vaqtinchalik (~$) fayllari o'tkazib yuboriladi."""
    if not folder.is_dir():
        return []
    files = [p for p in folder.iterdir()
             if p.is_file() and p.suffix.lower() == ".xlsx"
             and not p.name.startswith("~$") and not p.name.startswith(".")]
    return sorted(files, key=lambda p: _natural(p.name))
