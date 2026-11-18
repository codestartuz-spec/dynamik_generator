# -*- coding: utf-8 -*-
"""CAS hisobot fayllaridan kerakli ma'lumotlarni yuklash."""
from datetime import datetime, timedelta

import config
from xlsx_reader import iter_rows, read_header, meter_key, col_to_index

_EXCEL_EPOCH = datetime(1899, 12, 30)
_TEXT_FORMATS = ("%d/%m/%Y %H:%M:%S", "%d.%m.%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M")


def to_datetime(v):
    """Excel sana raqami yoki matnni datetime ga aylantiradi; bo'lmasa None."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        if v <= 0:
            return None
        return _EXCEL_EPOCH + timedelta(days=float(v))
    s = str(v).strip()
    for fmt in _TEXT_FORMATS:
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def load_direct_meters(path):
    """DirectMeter.xlsx dagi hisoblagich raqamlari to'plami."""
    wanted = set()
    rows = iter_rows(path, wanted)
    header = dict(read_header(rows, [config.DIRECT_METER_COL]))
    col = header[config.DIRECT_METER_COL]
    wanted.add(col)
    result = set()
    for _rnum, cells in rows:
        key = meter_key(cells.get(col))
        if key:
            result.add(key)
    return result


def _value_after_label(preface, label):
    """Sarlavha qatorlarida 'Count:' / 'Date:' kabi yozuvdan keyingi qiymatni topadi."""
    for _rnum, cells in preface:
        for col in sorted(cells, key=col_to_index):
            v = cells[col]
            if isinstance(v, str) and v.strip().lower().startswith(label):
                for c2 in sorted(cells, key=col_to_index):
                    if col_to_index(c2) > col_to_index(col) and not isinstance(cells[c2], str):
                        return cells[c2]
    return None


def load_poll_file(path, progress=None):
    """
    Bitta MeteringDevicesReport fayli: ({raqam: obxod vaqti}, ma'lumot).
    ma'lumot: declared (faylning o'zidagi 'Count'), rows (haqiqiy qatorlar), date (hisobot sanasi).
    """
    wanted = set()
    preface = []
    rows = iter_rows(path, wanted)
    try:
        header = dict(read_header(rows, [config.METERING_METER_COL, config.METERING_POLL_COL], preface))
    except ValueError:
        raise ValueError("%s: MeteringDevicesReport hisobotiga o'xshamaydi ('%s' / '%s' ustunlari topilmadi). "
                         "Faylni papkadan olib tashlang." % (path.name, config.METERING_METER_COL,
                                                              config.METERING_POLL_COL))
    mcol = header[config.METERING_METER_COL]
    pcol = header[config.METERING_POLL_COL]
    wanted.update([mcol, pcol])
    result = {}
    count = 0
    duplicates = 0
    for _rnum, cells in rows:
        key = meter_key(cells.get(mcol))
        if not key or key == config.METERING_METER_COL:
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
    declared = _value_after_label(preface, "count")
    info = {
        "declared": int(declared) if isinstance(declared, (int, float)) else None,
        "rows": count,
        "with_poll": len(result),
        "duplicates": duplicates,
        "date": to_datetime(_value_after_label(preface, "date")),
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
    """Bitta Meter Archive Details fayli: tanlangan ustunlar va qatorlar."""

    def __init__(self, path):
        self.path = path
        self._rows = iter_rows(path)
        try:
            header = read_header(self._rows, [config.ARCHIVE_METER_COL])
        except ValueError:
            raise ValueError("%s: Meter Archive Details hisobotiga o'xshamaydi ('%s' ustuni topilmadi). "
                             "Faylni papkadan olib tashlang." % (path.name, config.ARCHIVE_METER_COL))
        drop = set(n.strip().casefold() for n in config.DROP_COLUMNS)
        self.all_columns = [n for n, _c in header]
        self.columns = [(n, c) for n, c in header if n.strip().casefold() not in drop]
        self.meter_col = dict(header)[config.ARCHIVE_METER_COL]

    def rows(self):
        """Har bir qator uchun ({ustun nomi: qiymat}, hisoblagich_kaliti) qaytaradi."""
        for _rnum, cells in self._rows:
            key = meter_key(cells.get(self.meter_col))
            if not key or key == config.ARCHIVE_METER_COL:
                continue
            yield {n: cells.get(c) for n, c in self.columns}, key
