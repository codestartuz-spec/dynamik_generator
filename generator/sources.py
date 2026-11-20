# -*- coding: utf-8 -*-
"""CAS hisobot fayllaridan kerakli ma'lumotlarni yuklash (ustunlar tilga bog'liq emas - schema.py)."""
import config
import specs
from schema import open_table, to_datetime, norm
from xlsx_reader import meter_key, col_to_index


def load_direct_meters(path, log=print):
    """DirectMeter fayldagi hisoblagich raqamlari to'plami."""
    wanted = set()
    tbl = open_table(path, specs.DIRECT, wanted)
    if tbl.note:
        log("   %s: %s" % (path.name, tbl.note))
    col = tbl.cols["meter"]
    wanted.add(col)
    result = set()
    for _rnum, cells in tbl.rows:
        key = meter_key(cells.get(col))
        if key:
            result.add(key)
    return result


def _value_after_label(preface, labels):
    """
    Sarlavha qatorlarida 'Count:' / 'Date:' / 'Total:' kabi yozuvdan keyingi qiymatni topadi.
    labels - bir nechta tildagi variantlar (kichik harfda, boshlanishi bo'yicha solishtiriladi).
    """
    if isinstance(labels, str):
        labels = (labels,)
    labels = tuple(l.lower() for l in labels)
    for _rnum, cells in preface:
        for col in sorted(cells, key=col_to_index):
            v = cells[col]
            if isinstance(v, str) and v.strip().lower().startswith(labels):
                for c2 in sorted(cells, key=col_to_index):
                    if col_to_index(c2) > col_to_index(col) and not isinstance(cells[c2], str):
                        return cells[c2]
    return None


COUNT_LABELS = ("count", "количество", "кол-во", "soni", "miqdori")
DATE_LABELS = ("date", "дата", "sana")
TOTAL_LABELS = ("total", "всего", "итого", "jami", "umumiy")

_METER_HEADER_NAMES = specs.METERING.by_key["meter"].aliases | specs.ARCHIVE.by_key["meter"].aliases


def load_poll_file(path, progress=None):
    """
    Bitta MeteringDevicesReport fayli: ({raqam: obxod vaqti}, ma'lumot).
    ma'lumot: declared (faylning o'zidagi 'Count'), rows (haqiqiy qatorlar), date (hisobot sanasi), note.
    """
    wanted = set()
    tbl = open_table(path, specs.METERING, wanted)
    mcol, pcol = tbl.cols["meter"], tbl.cols["poll"]
    wanted.update([mcol, pcol])
    result = {}
    count = 0
    duplicates = 0
    for _rnum, cells in tbl.rows:
        key = meter_key(cells.get(mcol))
        if not key or norm(key) in _METER_HEADER_NAMES:
            continue
        count += 1
        dt = to_datetime(cells.get(pcol))
        if key in result:
            duplicates += 1
        if dt is not None:
            old = result.get(key)
            if old is None or dt > old:
                result[key] = dt
        if progress and count % 200000 == 0:
            progress(count)
    declared = _value_after_label(tbl.preface, COUNT_LABELS)
    info = {
        "declared": int(declared) if isinstance(declared, (int, float)) else None,
        "rows": count,
        "with_poll": len(result),
        "duplicates": duplicates,
        "date": to_datetime(_value_after_label(tbl.preface, DATE_LABELS)),
        "note": tbl.note,
    }
    return result, info


def load_poll_times(paths, log=print):
    """
    Bir nechta MeteringDevicesReport bo'laklarini o'qiydi, tekshiradi va birlashtiradi.
    Qaytaradi: ({raqam: obxod vaqti}, [ogohlantirishlar]).
    """
    merged = {}
    warnings = []
    seen_rows = 0
    for p in paths:
        log("   %s o'qilmoqda..." % p.name)
        polls, info = load_poll_file(p, progress=lambda n: log("      o'qildi: %d qator" % n))
        if info["note"]:
            log("      " + info["note"])
        decl = info["declared"]
        mark = "OK"
        if decl is not None and decl != info["rows"]:
            mark = "MOS EMAS"
            warnings.append("%s: faylda Count=%d, lekin %d qator o'qildi (fayl to'liq yuklanmagan bo'lishi mumkin)"
                            % (p.name, decl, info["rows"]))
        log("      qatorlar: %d, Count: %s [%s], hisobot vaqti: %s"
            % (info["rows"], decl if decl is not None else "?", mark,
               info["date"].strftime("%d.%m.%Y %H:%M") if info["date"] else "?"))
        if info["duplicates"]:
            warnings.append("%s: ichida %d ta takroriy hisoblagich" % (p.name, info["duplicates"]))
        overlap = 0
        for k, dt in polls.items():
            if k in merged:
                overlap += 1
                if dt <= merged[k]:
                    continue
            merged[k] = dt
        if overlap:
            warnings.append("%s: oldingi bo'laklar bilan %d ta hisoblagich takrorlangan (yangisi olindi)"
                            % (p.name, overlap))
        seen_rows += info["rows"]
    return merged, warnings


class ArchiveFile(object):
    """
    Bitta Meter Archive Details fayli: tanlangan ustunlar va qatorlar.
    Ustun nomlari chiqishda KANONIK (rus tilidagi) nomlarga keltiriladi, shuning uchun CAS tili o'zgarsa
    ham natijadagi sarlavhalar bir xil bo'ladi. Tanilmagan qo'shimcha ustunlar o'z nomi bilan saqlanadi.
    """

    def __init__(self, path):
        self.path = path
        tbl = open_table(path, specs.ARCHIVE)
        self.note = tbl.note
        self._rows = tbl.rows
        self.meter_col = tbl.cols["meter"]
        self.all_columns = []
        self.columns = []
        for name, letter in tbl.header:
            f = tbl.letter_field.get(letter)
            out_name = f.name if f else name
            self.all_columns.append(out_name)
            if f is not None and f.drop:
                continue
            self.columns.append((out_name, letter))

    def rows(self):
        """Har bir qator uchun ({ustun nomi: qiymat}, hisoblagich_kaliti) qaytaradi."""
        for _rnum, cells in self._rows:
            key = meter_key(cells.get(self.meter_col))
            if not key or norm(key) in _METER_HEADER_NAMES:
                continue
            yield {n: cells.get(c) for n, c in self.columns}, key
