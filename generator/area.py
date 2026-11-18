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
from sources import _value_after_label
from xlsx_reader import iter_rows, read_header, meter_key


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
def _latest(pattern):
    files = [p for p in config.FILE_DIR.glob(pattern) if not p.name.startswith("~$")]
    return max(files, key=lambda p: p.stat().st_mtime) if files else None


def _open_table(path, required):
    preface = []
    rows = iter_rows(path)
    try:
        header = read_header(rows, required, preface)
    except ValueError:
        raise ValueError("%s: kutilgan hisobotga o'xshamaydi (kerakli ustunlar: %s)" % (path.name, ", ".join(required)))

    def gen():
        for _rn, cells in rows:
            yield {n: cells.get(c) for n, c in header}
    return preface, gen()


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


def load_area(path):
    required = [config.AREA_ESP_COL, config.AREA_TR_COL, config.AREA_MANAGER_COL] + config.AREA_KEEP_COLUMNS
    preface, rows = _open_table(path, sorted(set(required)))
    declared = _value_after_label(preface, "total")
    result = []
    for r in rows:
        if _s(r.get(config.AREA_TR_COL)) is None and _s(r.get(config.AREA_TRCODE_COL)) is None:
            continue
        result.append({c: _s(r.get(c)) for c in config.AREA_KEEP_COLUMNS})
    return result, (int(declared) if isinstance(declared, (int, float)) else None)


def load_commissioning(path):
    """{(ESP, TR No.): (dcu, installed, daily)}"""
    _p, rows = _open_table(path, [config.COMM_ESP_COL, config.COMM_TR_COL, config.COMM_DCU_COL])
    result = {}
    for r in rows:
        esp, tr = _s(r.get(config.COMM_ESP_COL)), _s(r.get(config.COMM_TR_COL))
        if not esp or not tr:
            continue
        result[(esp, tr)] = (meter_key(r.get(config.COMM_DCU_COL)),
                             _num(r.get(config.COMM_INSTALLED_COL)),
                             _num(r.get(config.COMM_DAILY_COL)))
    return result


def load_online(path):
    """{dcu raqami: (holat, holat vaqti)}"""
    _p, rows = _open_table(path, [config.ONLINE_DCU_COL, config.ONLINE_STATUS_COL, config.ONLINE_TIME_COL])
    result = {}
    for r in rows:
        typ = _s(r.get(config.ONLINE_TYPE_COL))
        if typ is not None and typ.upper() != config.ONLINE_TYPE_DCU:
            continue
        key = meter_key(r.get(config.ONLINE_DCU_COL))
        if not key:
            continue
        t = None
        raw = _s(r.get(config.ONLINE_TIME_COL))
        if raw:
            try:
                t = datetime.strptime(raw, config.ONLINE_TIME_FORMAT)
            except ValueError:
                t = None
        result[key] = (_s(r.get(config.ONLINE_STATUS_COL)), t)
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
    limit_time = today.replace(hour=0, minute=0, second=0, microsecond=0)
    from datetime import timedelta
    limit_time = limit_time - timedelta(hours=config.OFFLINE_LONG_HOURS)
    stats = {"rows": 0, "no_comm": 0, "dcu": 0, "online_missing": 0, "offline": 0, "offline_long": 0,
             "daily_missing": 0, "managers_blank": 0, "bins": [0] * len(config.DAILY_BINS), "used_comm": set()}
    out = []
    for a in area_rows:
        esp, tr, code = a[config.AREA_ESP_COL], a[config.AREA_TR_COL], a[config.AREA_TRCODE_COL]
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
            new[1] = (1 if installed > 0 else 0) if installed is not None else None
            new[2] = daily
            o = online.get(dcu)
            if o is None:
                stats["online_missing"] += 1
            else:
                status, t = o
                new[3] = status
                if status and status.strip().lower() == config.STATUS_OFFLINE:
                    new[4] = 1
                    stats["offline"] += 1
                    if t is not None and t < limit_time:
                        new[5] = 1
                        stats["offline_long"] += 1
            bins = daily_bins(daily)
            if daily is None:
                stats["daily_missing"] += 1
            for i, b in enumerate(bins):
                new[6 + i] = b
                if b == 1:
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
def run(today, out_dir, log=print):
    """Area hisobotini tayyorlaydi. 0 - muvaffaqiyatli, 1 - xato."""
    area_f, comm_f, online_f = (_latest(config.AREA_GLOB), _latest(config.COMMISSIONING_GLOB),
                                _latest(config.ONLINE_GLOB))
    missing = [pat for f, pat in ((area_f, config.AREA_GLOB), (comm_f, config.COMMISSIONING_GLOB),
                                  (online_f, config.ONLINE_GLOB)) if f is None]
    if missing:
        log("XATO (Area): fayllar topilmadi: " + ", ".join(str(config.FILE_DIR / m) for m in missing))
        return 1
    date_s = today.strftime(config.OUTPUT_DATE_FORMAT)
    xlsx_path = out_dir / config.AREA_OUTPUT_NAME.format(date=date_s)
    csv_path = out_dir / config.AREA_CSV_NAME.format(date=date_s)

    log("=== AREA (transformatorlar) hisoboti ===")
    for f in (area_f, comm_f, online_f):
        log("   " + f.name)
    area_rows, declared = load_area(area_f)
    log("   Area: %d ta transformator (faylda Total: %s)" % (len(area_rows), declared if declared is not None else "?"))
    if declared is not None and declared != len(area_rows):
        log("   DIQQAT: Total=%d, lekin %d qator o'qildi (fayl to'liq yuklanmagan bo'lishi mumkin)"
            % (declared, len(area_rows)))
    comm = load_commissioning(comm_f)
    online = load_online(online_f)
    log("   Commissioning: %d, Online (DCU): %d" % (len(comm), len(online)))

    rows, st = compute_rows(area_rows, comm, online, today)
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
    log("TAYYOR (Area Excel): %s" % xlsx_path)
    log("TAYYOR (Area CSV+gzip): %s  (%.1f MB)" % (csv_path, size / 1e6))
    log("   Qatorlar: %d, konsentratori bor: %d, ismi bo'sh: %d" % (st["rows"], st["dcu"], st["managers_blank"]))
    log("   Offline: %d, 3 kundan ortiq offline: %d, Daily Data %% yo'q: %d"
        % (st["offline"], st["offline_long"], st["daily_missing"]))
    for (name, _k, _l, _h), cnt in zip(config.DAILY_BINS, st["bins"]):
        log("   %-14s: %d" % (name, cnt))
    return 0
