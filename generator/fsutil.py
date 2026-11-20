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


def find_reports(folder, prefixes):
    """
    Nomi prefiks bilan boshlanadigan .xlsx fayllar, eng yangisi (o'zgargan vaqti bo'yicha) birinchi.
    Katta-kichik harf farqlanmaydi: DirectMeter.xlsx, DirectMeter(2).xlsx, DirectMeter (1).xlsx ... hammasi mos.
    """
    if not folder.is_dir():
        return []
    pre = tuple(p.casefold() for p in prefixes)
    files = [p for p in folder.iterdir()
             if p.is_file() and p.suffix.lower() == ".xlsx"
             and not p.name.startswith("~$") and not p.name.startswith(".")
             and p.stem.casefold().startswith(pre)]
    return sorted(files, key=lambda p: p.stat().st_mtime, reverse=True)


def latest_report(folder, prefixes):
    """(eng yangi fayl yoki None, boshqa mos fayllar ro'yxati)."""
    files = find_reports(folder, prefixes)
    return (files[0] if files else None), files[1:]


def clear_upload(folder, only=None, log=print):
    """
    Yangi generatsiyadan oldin chiqish papkasidagi eski fayllarni o'chiradi (.gitkeep qoladi).
    only='area' / 'details' bo'lsa faqat shu hisobotning fayllari o'chiriladi.
    Qaytaradi: o'chirib bo'lmagan fayllar ro'yxati (masalan, Excel da ochiq turgan fayl).
    """
    prefix = {"area": "area_", "details": "details_"}.get(only)
    failed, removed = [], 0
    if not folder.is_dir():
        return failed
    for p in sorted(folder.iterdir()):
        if not p.is_file() or p.name == ".gitkeep":
            continue
        bare = p.name[2:] if p.name.startswith("~$") else p.name
        if prefix and not bare.casefold().startswith(prefix):
            continue
        try:
            p.unlink()
            removed += 1
        except OSError:
            failed.append(p)
    log("upload papkasi tozalandi: %d ta fayl o'chirildi" % removed)
    return failed
