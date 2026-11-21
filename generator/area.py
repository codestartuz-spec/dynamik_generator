# -*- coding: utf-8 -*-
"""
2-qism: Area (transformatorlar) hisoboti.

AreaInformationReport + CommissioningStatus + Online_Device_Monitroing_Report birlashtirilib,
upload\\ ga Area_Qashqadaryo HETK-<sana>.xlsx va .csv.gz sifatida yoziladi.
"""
import csv
import gzip
import os
import re
from datetime import datetime

import xlsxwriter

import config
import specs
from fsutil import latest_report
from progress import Span
from schema import open_table, to_datetime, norm
from sources import _value_after_label, TOTAL_LABELS
from xlsx_reader import meter_key


# ---------------------------------------------------------------- ismlar ----
_APOS = {ord(c): "'" for c in "`‘’ʻʼ´′"}
_CYR = {
    "А": "A", "Б": "B", "В": "V", "Г": "G", "Д": "D", "Ё": "YO", "Ж": "J", "З": "Z", "И": "I",
    "Й": "Y", "К": "K", "Л": "L", "М": "M", "Н": "N", "О": "O", "П": "P", "Р": "R", "С": "S",
    "Т": "T", "У": "U", "Ф": "F", "Х": "X", "Ц": "TS", "Ч": "CH", "Ш": "SH", "Щ": "SH",
    "Ъ": "'", "Ы": "I", "Ь": "", "Э": "E", "Ю": "YU", "Я": "YA",
    "Ў": "O'", "Ғ": "G'", "Қ": "Q", "Ҳ": "H",
}
_VOWELS = set("AEIOUЕАЁИОУЭЮЯЎ")


def _translit(s):
    """Kirill harflarni lotinga o'tkazadi (lotin harflar o'zgarmaydi). s allaqachon katta harfda."""
    out = []
    for i, ch in enumerate(s):
        if ch == "Е":
            prev = s[i - 1] if i else " "
            out.append("YE" if (not prev.isalpha() or prev in _VOWELS or prev == "'") else "E")
        else:
            out.append(_CYR.get(ch, ch))
    return "".join(out)


def normalize_name(v):
    """Ism-familiya: lotin, KATTA harf, apostrof belgilari bir xil ('), ortiqcha bo'sh joylarsiz."""
    if v is None:
        return None
    s = re.sub(r"\s+", " ", str(v)).strip()
    if not s:
        return None
    s = _translit(s.upper()).translate(_APOS)
    return re.sub(r"\s+", " ", s).strip() or None


# ------------------------------------------------------------- o'qish -------
def _join_key(esp):
    """ESP nomini solishtirish uchun: apostrof/harf katta-kichikligi/kirill-lotin farqlari e'tiborga olinmaydi."""
    return norm(normalize_name(esp) or "")


def _open_table(path, spec, log=print):
    """Hisobotni ochadi: (preface, {ichki nom: qiymat} qatorlari generatori)."""
    tbl = open_table(path, spec)
    if tbl.note:
        log("   %s: %s" % (path.name, tbl.note))
    cols = tbl.cols

    def gen():
        for _rn, cells in tbl.rows:
            yield {k: cells.get(c) for k, c in cols.items()}
    return tbl.preface, gen()


def _s(v):
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    v = str(v).strip()
    return v or None


def _num(v):
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        try:
            return float(v.strip().replace(",", ".").rstrip("%"))
        except ValueError:
            return None
    return None


def load_area(path, log=print):
    preface, rows = _open_table(path, specs.AREA, log)
    declared = _value_after_label(preface, TOTAL_LABELS)
    fields = specs.AREA.fields
    result = []
    for r in rows:
        if _s(r.get("tr")) is None and _s(r.get("tr_code")) is None:
            continue
        result.append({f.name: _s(r.get(f.key)) for f in fields})
    return result, (int(declared) if isinstance(declared, (int, float)) else None)


def load_commissioning(path, log=print):
    """{(ESP, TR No.): (dcu, installed, daily)}"""
    _p, rows = _open_table(path, specs.COMMISSIONING, log)
    result = {}
    for r in rows:
        esp, tr = _s(r.get("esp")), _s(r.get("tr"))
        if not esp or not tr:
            continue
        result[(_join_key(esp), tr)] = (meter_key(r.get("dcu")), _num(r.get("installed")), _num(r.get("daily")))
    return result


def is_offline(status):
    """Holat 'offline' ekanligi (Offline / Оффлайн / Офлайн / Oflayn ...)."""
    return bool(status) and norm(status).startswith(config.STATUS_OFFLINE_PREFIXES)


def load_online(path, log=print):
    """{dcu raqami: (holat, holat vaqti)}"""
    _p, rows = _open_table(path, specs.ONLINE, log)
    result = {}
    for r in rows:
        typ = _s(r.get("type"))
        if typ is not None and typ.upper() != config.ONLINE_TYPE_DCU:
            continue
        key = meter_key(r.get("dcu"))
        if not key:
            continue
        result[key] = (_s(r.get("status")), to_datetime(r.get("time")))
    return result


# --------------------------------------------------------- hisoblash --------
def daily_bins(value):
    """Daily Data % bo'yicha 5 ta oraliq ustuni uchun [1/0, ...]; qiymat yo'q bo'lsa [None]*5."""
    if value is None:
        return [None] * len(config.DAILY_BINS)
    out = []
    for _name, kind, lo, hi in config.DAILY_BINS:
        if kind == "eq":
            ok = value == lo
        else:
            ok = value > lo and (hi is None or value <= hi)
        out.append(1 if ok else 0)
    return out


def compute_rows(area_rows, comm, online, today):
    """Yakuniy qatorlar ro'yxati: har biri uchun ustunlar tartibida qiymatlar. today - bugun 00:00."""
    now = today          # generatsiya qilingan vaqt: offline muddati shunga nisbatan hisoblanadi
    stats = {"rows": 0, "no_comm": 0, "dcu": 0, "online_missing": 0, "offline": 0, "offline_long": 0, "offline_all": 0, "offline_no_time": 0,
             "daily_missing": 0, "managers_blank": 0, "bins": [0] * len(config.DAILY_BINS), "used_comm": set()}
    out = []
    for a in area_rows:
        esp, tr, code = _join_key(a[config.AREA_ESP_COL]), a[config.AREA_TR_COL], a[config.AREA_TRCODE_COL]
        c = comm.get((esp, tr)) or comm.get((esp, code))
        key = (esp, tr) if (esp, tr) in comm else ((esp, code) if (esp, code) in comm else None)
        base = [a[k] for k in config.AREA_KEEP_COLUMNS]
        mi = config.AREA_KEEP_COLUMNS.index(config.AREA_MANAGER_COL)
        base[mi] = normalize_name(a[config.AREA_MANAGER_COL])
        if base[mi] is None:
            stats["managers_blank"] += 1
        stats["rows"] += 1
        if c is None:
            stats["no_comm"] += 1
        else:
            stats["used_comm"].add(key)
        dcu = c[0] if c else None
        new = [None] * (6 + len(config.DAILY_BINS))
        if dcu:
            stats["dcu"] += 1
            installed, daily = c[1], c[2]
            new[0] = dcu
            new[1] = 1 if (installed is not None and installed > 0) else None   # 0 -> bo'sh
            new[2] = daily
            o = online.get(dcu)
            if o is None:
                stats["online_missing"] += 1
            else:
                status, t = o
                new[3] = status
                if is_offline(status):
                    stats["offline_all"] += 1
                    if t is None:
                        stats["offline_no_time"] += 1
                    else:
                        hours = (now - t).total_seconds() / 3600.0
                        if hours > config.OFFLINE_MIN_HOURS:
                            new[4] = 1
                            stats["offline"] += 1
                        if hours > config.OFFLINE_LONG_HOURS:
                            new[5] = 1
                            stats["offline_long"] += 1
            bins = daily_bins(daily)
            if daily is None:
                stats["daily_missing"] += 1
            for i, b in enumerate(bins):
                if b == 1:                       # 1 bo'lsa yoziladi, 0 va yo'q qiymat bo'sh qoladi
                    new[6 + i] = 1
                    stats["bins"][i] += 1
        out.append(base + new)
    return out, stats


def header_names():
    return (list(config.AREA_KEEP_COLUMNS)
            + [config.AREA_DCU_COL, config.AREA_CONSUMER_COL, config.AREA_LEVEL_COL, config.AREA_STATE_COL,
               config.AREA_OFFLINE_COL, config.AREA_OFFLINE_LONG_COL]
            + [b[0] for b in config.DAILY_BINS])


# ------------------------------------------------------------- yozish -------
_WIDTHS = {"ESP Code": 10, "ESP name": 18, "Substation Code": 12, "Substation": 28, "Feeder Code": 10,
           "Feeder": 22, "TR Code": 10, "TR": 9, "Manager Name": 46, config.AREA_DCU_COL: 15,
           config.AREA_CONSUMER_COL: 14, config.AREA_LEVEL_COL: 12, config.AREA_STATE_COL: 13,
           config.AREA_OFFLINE_COL: 9, config.AREA_OFFLINE_LONG_COL: 13}


def _csv_cell(v):
    if v is None:
        return ""
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else repr(v)
    return v


def write_outputs(rows, xlsx_path, csv_path):
    header = header_names()
    tmp_x, tmp_c = str(xlsx_path) + ".tmp", str(csv_path) + ".tmp"
    try:
        wb = xlsxwriter.Workbook(tmp_x, {"strings_to_numbers": False, "strings_to_formulas": False,
                                         "strings_to_urls": False})
        ws = wb.add_worksheet(config.AREA_SHEET)
        f_head = wb.add_format({"bold": True, "bg_color": "#1F4E78", "font_color": "#FFFFFF", "border": 1,
                                "align": "center", "valign": "vcenter", "text_wrap": True})
        f_center = wb.add_format({"align": "center"})
        f_pct = wb.add_format({"align": "center", "num_format": "0.00"})
        ws.set_row(0, 32)
        for i, name in enumerate(header):
            ws.set_column(i, i, _WIDTHS.get(name, 12))
            ws.write_string(0, i, name, f_head)
        n_base = len(config.AREA_KEEP_COLUMNS)
        with gzip.open(tmp_c, "wt", encoding="utf-8", newline="", compresslevel=config.CSV_GZIP_LEVEL) as cf:
            cw = csv.writer(cf)
            cw.writerow(header)
            for r, row in enumerate(rows, start=1):
                for i, v in enumerate(row):
                    if v is None:
                        continue
                    if i < n_base or i in (n_base, n_base + 3):          # matn: kodlar, ism, DCU, holat
                        ws.write_string(r, i, str(v))
                    elif i == n_base + 2:                                # Obxod darajasi (foiz)
                        ws.write_number(r, i, v, f_pct)
                    else:
                        ws.write_number(r, i, v, f_center)
                cw.writerow([_csv_cell(v) for v in row])
        ws.freeze_panes(1, 0)
        ws.autofilter(0, 0, len(rows), len(header) - 1)
        wb.close()
    except BaseException:
        for p in (tmp_x, tmp_c):
            if os.path.exists(p):
                try:
                    os.remove(p)
                except OSError:
                    pass
        raise
    os.replace(tmp_c, str(csv_path))
    os.replace(tmp_x, str(xlsx_path))
    return os.path.getsize(str(csv_path))


# --------------------------------------------------------------- asosiy -----
def run(today, out_dir, log=print, span=None):
    """Area hisobotini tayyorlaydi. 0 - muvaffaqiyatli, 1 - xato."""
    span = span or Span()
    span.at(0.02, "Area: fayllar tekshirilmoqda")
    found = [latest_report(config.FILE_DIR, p) for p in
             (config.AREA_PREFIXES, config.COMMISSIONING_PREFIXES, config.ONLINE_PREFIXES)]
    area_f, comm_f, online_f = (f for f, _o in found)
    missing = [pre[0] + "*.xlsx" for f, pre in zip((area_f, comm_f, online_f),
               (config.AREA_PREFIXES, config.COMMISSIONING_PREFIXES, config.ONLINE_PREFIXES)) if f is None]
    if missing:
        log("XATO (Area): fayllar topilmadi: " + ", ".join(str(config.FILE_DIR / m) for m in missing))
        return 1
    for f, others in found:
        if others:
            log("   DIQQAT: '%s' kabi bir nechta fayl bor, eng yangisi olindi: %s (olinmadi: %s)"
                % (f.name, f.name, ", ".join(o.name for o in others)))
    date_s = today.strftime(config.OUTPUT_DATE_FORMAT)
    xlsx_path = out_dir / config.AREA_OUTPUT_NAME.format(date=date_s)
    csv_path = out_dir / config.AREA_CSV_NAME.format(date=date_s)

    log("=== AREA (transformatorlar) hisoboti ===")
    for f in (area_f, comm_f, online_f):
        log("   " + f.name)
    span.at(0.05, "Area: transformatorlar o'qilmoqda")
    area_rows, declared = load_area(area_f, log)
    span.at(0.45, "Area: Commissioning o'qilmoqda")
    log("   Area: %d ta transformator (faylda Total: %s)" % (len(area_rows), declared if declared is not None else "?"))
    if declared is not None and declared != len(area_rows):
        log("   DIQQAT: Total=%d, lekin %d qator o'qildi (fayl to'liq yuklanmagan bo'lishi mumkin)"
            % (declared, len(area_rows)))
    comm = load_commissioning(comm_f, log)
    span.at(0.6, "Area: Online holat o'qilmoqda")
    online = load_online(online_f, log)
    span.at(0.7, "Area: hisoblanmoqda")
    log("   Commissioning: %d, Online (DCU): %d" % (len(comm), len(online)))

    rows, st = compute_rows(area_rows, comm, online, today)
    span.at(0.8, "Area: fayllar yozilmoqda")
    if st["no_comm"]:
        log("   DIQQAT: Commissioning da topilmagan transformatorlar: %d" % st["no_comm"])
    comm_unused = len(comm) - len(st["used_comm"])
    if comm_unused:
        log("   DIQQAT: Commissioning da bor, lekin Area da yo'q transformatorlar: %d" % comm_unused)
    if st["online_missing"]:
        log("   DIQQAT: Online faylda topilmagan konsentratorlar: %d (holati bo'sh)" % st["online_missing"])
    dcu_all = set(r[0] for r in comm.values() if r[0])
    extra = len(set(online) - dcu_all)
    if extra:
        log("   DIQQAT: Online faylda bor, lekin Commissioning da yo'q konsentratorlar: %d" % extra)

    try:
        size = write_outputs(rows, xlsx_path, csv_path)
    except PermissionError:
        log("XATO (Area): faylni yozib bo'lmadi. Excel da ochiq bo'lsa yoping: %s" % xlsx_path)
        return 1
    span.at(1.0, "Area tayyor")
    log("TAYYOR (Area Excel): %s" % xlsx_path)
    log("TAYYOR (Area CSV+gzip): %s  (%.1f MB)" % (csv_path, size / 1e6))
    log("   Qatorlar: %d, konsentratori bor: %d, ismi bo'sh: %d" % (st["rows"], st["dcu"], st["managers_blank"]))
    log("   Offline holatida jami: %d; shundan %d soatdan ortiq (Offline=1): %d, %d soatdan ortiq (3 kundan ortiq offline=1): %d"
        % (st["offline_all"], config.OFFLINE_MIN_HOURS, st["offline"], config.OFFLINE_LONG_HOURS, st["offline_long"]))
    if st["offline_no_time"]:
        log("   DIQQAT: offline holatidagi %d ta konsentratorning vaqti noma'lum (belgilanmadi)" % st["offline_no_time"])
    log("   Daily Data %% yo'q: %d" % st["daily_missing"])
    for (name, _k, _l, _h), cnt in zip(config.DAILY_BINS, st["bins"]):
        log("   %-14s: %d" % (name, cnt))
    return 0
