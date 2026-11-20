# -*- coding: utf-8 -*-
"""Birlashgan Details_Qashqadaryo HETK faylini yig'ish va Excel hamda CSV(gzip) ga yozish."""
import csv
import gzip
import math
import os

import xlsxwriter

import config
from geo import RegionChecker
from sources import ArchiveFile

_WIDTHS = {
    "ЭСП": 18, "ТП": 12, "Лицевой счет потребителя": 20, "Номер счетчика": 18,
    "Тип счетчика": 20, "Тип связи": 11, "Тип потребителя": 15, "Коэффициент Т/Т": 12,
    "Коэффициент Т/Н": 12, "широта": 12, "долгота": 12, "Статус оборудования": 16,
    "Имя потребителя": 34, "Адрес установки": 40, config.POLL_COL_NAME: 20, config.DAYS_COL_NAME: 11,
}


def lag_seconds(poll, ref):
    """Hisoblash vaqtidan oxirgi obxodgacha o'tgan soniyalar (butun)."""
    return int((ref - poll).total_seconds())


def lag_days(poll, ref):
    """Obxod kuni: o'tgan to'liq sutkalar (24 soatdan kam -> 0, 24-48 soat -> 1, 48-72 soat -> 2 ...)."""
    return max(0, lag_seconds(poll, ref) // 86400)


def bucket_values(poll, ref):
    """Obxod vaqtiga ko'ra [48 soat, 2-7, 7-30, 30-90, 90+] qiymatlari (1/0 yoki '-'): pastki*24 soat <= o'tgan < yuqori*24 soat."""
    if poll is None:
        return [config.NO_VALUE] * len(config.BUCKETS)
    lag = lag_seconds(poll, ref)
    values = []
    for _name, lo, hi in config.BUCKETS:
        ok = (lo is None or lag >= lo * 86400) and (hi is None or lag < hi * 86400)
        values.append(1 if ok else 0)
    return values


def to_float(v):
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        try:
            return float(v.strip().replace(",", "."))
        except ValueError:
            return None
    return None


def fix_coords(lat_raw, lon_raw, region):
    """
    Qaytaradi (lat, lon, holat). holat: 'none' (yo'q edi), 'ok', 'fixed' (o'rni almashtirildi),
    'blank' (noto'g'ri yoki viloyatdan tashqarida - bo'sh qoldiriladi).
    """
    if lat_raw is None and lon_raw is None:
        return None, None, "none"
    lat, lon = to_float(lat_raw), to_float(lon_raw)
    if lat is None or lon is None or not (math.isfinite(lat) and math.isfinite(lon)):
        return None, None, "blank"
    if region.contains(lat, lon):
        return lat, lon, "ok"
    if region.contains(lon, lat):      # shirota va uzunlik o'rni almashib qolgan
        return lon, lat, "fixed"
    return None, None, "blank"


def _csv_val(v):
    """CSV uchun qiymat: bo'sh -> '', butun son float -> int ko'rinishida."""
    if v is None:
        return config.CSV_EMPTY
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else repr(v)
    return v


def build(archive_paths, poll_times, direct_set, out_path, ref, limit=None, log=print, csv_path=None):
    archives = [ArchiveFile(p) for p in archive_paths]
    for a in archives:
        if a.note:
            log("   %s: %s" % (a.path.name, a.note))
    base_cols = [n for n, _c in archives[0].columns]
    for a in archives[1:]:
        if set(n for n, _c in a.columns) != set(base_cols):
            raise ValueError("Ustunlar mos emas: %s" % a.path.name)

    bucket_names = [b[0] for b in config.BUCKETS]
    header = base_cols + [config.POLL_COL_NAME, config.DAYS_COL_NAME] + bucket_names + [config.DIRECT_COL_NAME]
    region = RegionChecker()
    lat_name = config.LAT_COL if config.LAT_COL in base_cols else None
    lon_name = config.LON_COL if config.LON_COL in base_cols else None

    tmp_path = str(out_path) + ".tmp"
    wb = xlsxwriter.Workbook(tmp_path, {
        "constant_memory": True,
        "strings_to_numbers": False,
        "strings_to_formulas": False,
        "strings_to_urls": False,
    })
    csv_tmp = (str(csv_path) + ".tmp") if csv_path else None
    csv_file = csv_w = None
    if csv_path:
        csv_file = gzip.open(csv_tmp, "wt", encoding="utf-8", newline="", compresslevel=config.CSV_GZIP_LEVEL)
        csv_w = csv.writer(csv_file)
        csv_w.writerow(header)
    ws = wb.add_worksheet(config.OUTPUT_SHEET)
    f_head = wb.add_format({"bold": True, "bg_color": "#1F4E78", "font_color": "#FFFFFF",
                            "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True})
    f_date = wb.add_format({"num_format": "dd/mm/yyyy hh:mm:ss"})
    f_center = wb.add_format({"align": "center"})

    for i, name in enumerate(header):
        ws.set_column(i, i, _WIDTHS.get(name, 11))
    ws.set_row(0, 30)
    for i, name in enumerate(header):
        ws.write_string(0, i, name, f_head)

    stats = {"rows": 0, "duplicates": 0, "poll_found": 0, "poll_missing": 0, "direct": 0,
             "geo": {"none": 0, "ok": 0, "fixed": 0, "blank": 0},
             "buckets": dict((n, 0) for n in bucket_names)}
    stats["files"] = []
    esp_name = "ЭСП" if "ЭСП" in base_cols else None
    seen = set()
    n_base = len(base_cols)
    r = 0
    try:
        for a in archives:
            f_rows, f_esp = 0, set()
            for row, key in a.rows():
                f_rows += 1
                if esp_name:
                    f_esp.add(row[esp_name])
                if key in seen:
                    stats["duplicates"] += 1
                    continue
                seen.add(key)
                r += 1
                if r >= config.EXCEL_MAX_ROWS:
                    raise ValueError("Excel qator limiti (1 048 576) dan oshib ketdi")
                if lat_name and lon_name:
                    row[lat_name], row[lon_name], gstate = fix_coords(row[lat_name], row[lon_name], region)
                    stats["geo"][gstate] += 1
                for i, name in enumerate(base_cols):
                    v = row[name]
                    if v is None:
                        continue
                    if isinstance(v, float):
                        ws.write_number(r, i, v)
                    else:
                        ws.write_string(r, i, str(v))
                poll = poll_times.get(key)
                # Quyidagi 8 ustunda faqat mazmunli qiymat yoziladi: "-" va 0 bo'sh katak bo'lib qoladi.
                if poll is not None:
                    ws.write_datetime(r, n_base, poll, f_date)
                    stats["poll_found"] += 1
                else:
                    stats["poll_missing"] += 1
                days = lag_days(poll, ref) if poll is not None else None
                if days is not None and not (days == 0 and config.BLANK_ZERO_DAYS):
                    ws.write_number(r, n_base + 1, days, f_center)
                else:
                    days = None
                vals = bucket_values(poll, ref)
                for j, v in enumerate(vals):
                    if v == 1:
                        ws.write_number(r, n_base + 2 + j, 1, f_center)
                        stats["buckets"][bucket_names[j]] += 1
                is_direct = 1 if key in direct_set else 0
                stats["direct"] += is_direct
                if is_direct:
                    ws.write_number(r, n_base + 2 + len(bucket_names), 1, f_center)
                if csv_w is not None:
                    csv_w.writerow(
                        [_csv_val(row[name]) for name in base_cols]
                        + [poll.strftime(config.CSV_DATETIME_FORMAT) if poll is not None else config.CSV_EMPTY,
                           days if days is not None else config.CSV_EMPTY]
                        + [1 if v == 1 else config.CSV_EMPTY for v in vals]
                        + [1 if is_direct else config.CSV_EMPTY])
                if r % 100000 == 0:
                    log("     %d qator yozildi" % r)
                if limit and r >= limit:
                    break
            stats["files"].append((a.path.name, f_rows, sorted(f_esp)))
            log("  -> %s: %d qator, ESP: %s" % (a.path.name, f_rows, ", ".join(sorted(f_esp)) or "?"))
            if limit and r >= limit:
                break
        stats["metering_not_in_archive"] = sum(1 for k in poll_times if k not in seen)
        ws.freeze_panes(1, 0)
        ws.autofilter(0, 0, r, len(header) - 1)
        wb.close()
        if csv_file is not None:
            csv_file.close()
            csv_file = None
    except BaseException:
        try:
            wb.close()
        except Exception:
            pass
        if csv_file is not None:
            try:
                csv_file.close()
            except Exception:
                pass
        for tp in (tmp_path, csv_tmp):
            if tp and os.path.exists(tp):
                os.remove(tp)
        raise

    if csv_path:
        os.replace(csv_tmp, str(csv_path))
        stats["csv_path"] = str(csv_path)
        stats["csv_size"] = os.path.getsize(str(csv_path))
        log("  CSV tayyor: %s" % csv_path)
    os.replace(tmp_path, str(out_path))
    stats["rows"] = r
    stats["columns"] = header
    return stats
