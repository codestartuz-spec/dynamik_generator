# -*- coding: utf-8 -*-
"""
Details_Qashqadaryo HETK generatori.

1) Meter Archive Details_*.xlsx (ESP bo'yicha 17 ta fayl) birlashtiriladi,
   keraksiz 6 ustun olib tashlanadi (1-qoida).
2) MeteringDevicesReport*.xlsx (bo'laklar) dan "Oxirgi obxod vaqti", "Obxod kuni" va 5 ta oraliq ustuni,
   DirectMeter.xlsx dan DirectMeter ustuni qo'shiladi (2-qoida).
3) Koordinatalar Qashqadaryo hududiga tekshiriladi (o'rni almashganlari to'g'rilanadi, noto'g'rilari bo'sh).
4) upload\ ga ikkita fayl yoziladi: Details_Qashqadaryo HETK-<sana>.xlsx (ko'rish uchun)
   va Details_Qashqadaryo HETK-<sana>.csv.gz (web ilovaga yuklash uchun).

Test uchun:  python main.py --limit 5000
"""
import argparse
import os
import sys
import time
import zipfile
from datetime import datetime

import area
import config
import builder
from fsutil import list_xlsx
from sources import load_direct_meters, load_poll_times


def log(msg):
    print("[%s] %s" % (datetime.now().strftime("%H:%M:%S"), msg), flush=True)


def run_details(args, today, out_dir):
    ref = today.replace(hour=config.REFERENCE_HOUR, minute=0, second=0, microsecond=0)
    date_s = today.strftime(config.OUTPUT_DATE_FORMAT)
    out_path = out_dir / config.OUTPUT_NAME.format(date=date_s)
    csv_path = out_dir / config.CSV_NAME.format(date=date_s)

    archives = list_xlsx(config.ARCHIVE_DIR)
    meterings = list_xlsx(config.METERING_DIR)
    direct = config.FILE_DIR / config.DIRECT_FILE

    missing = []
    if not archives:
        missing.append("%s  (papkada .xlsx fayl yo'q)" % config.ARCHIVE_DIR)
    if not meterings:
        missing.append("%s  (papkada .xlsx fayl yo'q)" % config.METERING_DIR)
    if not direct.exists():
        missing.append(str(direct))
    if missing:
        log("XATO: quyidagi papka/fayllar topilmadi yoki bo'sh:")
        for m in missing:
            log("   " + m)
        return 1

    t0 = time.time()
    log("Hisoblash vaqti: %s" % ref.strftime("%d.%m.%Y %H:%M"))
    log("Meter Archive Details fayllari (%d ta): %s" % (len(archives), config.ARCHIVE_DIR))
    for p in archives:
        log("   " + p.name)

    log("DirectMeter o'qilmoqda...")
    direct_set = load_direct_meters(direct)
    log("   DirectMeter: %d ta hisoblagich" % len(direct_set))

    log("MeteringDevicesReport fayllari (%d ta) o'qilmoqda: %s" % (len(meterings), config.METERING_DIR))
    log("   (bir necha daqiqa ketishi mumkin)")
    poll_times, warns = load_poll_times(meterings, log=log)
    log("   Metering jami: %d ta hisoblagich" % len(poll_times))
    for w in warns:
        log("   DIQQAT: " + w)

    log("Birlashgan fayl yig'ilmoqda -> %s va %s" % (out_path.name, csv_path.name))
    try:
        stats = builder.build(archives, poll_times, direct_set, out_path, ref,
                              limit=args.limit or None, log=log, csv_path=csv_path)
    except PermissionError:
        log("XATO: faylni yozib bo'lmadi. Excel da ochiq bo'lsa yoping: %s" % out_path)
        return 1

    log("TAYYOR (Excel): %s" % out_path)
    log("TAYYOR (CSV+gzip): %s  (%.1f MB)" % (csv_path, stats["csv_size"] / 1e6))
    log("   Qatorlar: %d  (takroriy tashlab yuborilgan: %d)" % (stats["rows"], stats["duplicates"]))
    log("   Obxod vaqti topildi: %d, topilmadi ('-'): %d" % (stats["poll_found"], stats["poll_missing"]))
    for name, cnt in stats["buckets"].items():
        log("   %-10s: %d" % (name, cnt))
    log("   DirectMeter = 1: %d" % stats["direct"])
    esps = set(e for _n, _r, es in stats["files"] for e in es)
    log("   ESP (tuman) soni: %d - %s" % (len(esps), ", ".join(sorted(esps))))
    if args.limit:
        log("   (test rejimi: to'liqlik tekshiruvlari o'tkazilmadi)")
    if stats["metering_not_in_archive"] and not args.limit:
        log("   DIQQAT: Metering da bor, lekin Archive da yo'q hisoblagichlar: %d (Archive to'liq emas bo'lishi mumkin)"
            % stats["metering_not_in_archive"])
    g = stats["geo"]
    log("   Koordinata: to'g'ri %d, o'rni almashtirib to'g'rilandi %d, noto'g'ri (bo'shatildi) %d, yo'q edi %d"
        % (g["ok"], g["fixed"], g["blank"], g["none"]))
    log("   Ustunlar: %s" % ", ".join(stats["columns"]))
    log("Jami vaqt: %.0f soniya" % (time.time() - t0))
    return 0




def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=0, help="faqat test: Details da shuncha qator yoziladi")
    ap.add_argument("--date", help="sana YYYY-MM-DD (default: bugun)")
    ap.add_argument("--out-dir", help="chiqish papkasi (default: upload)")
    ap.add_argument("--only", choices=["details", "area"], help="faqat bitta hisobotni tayyorlash")
    args = ap.parse_args()

    today = datetime.strptime(args.date, "%Y-%m-%d") if args.date else datetime.now()
    out_dir = type(config.UPLOAD_DIR)(args.out_dir) if args.out_dir else config.UPLOAD_DIR
    out_dir.mkdir(parents=True, exist_ok=True)

    mode = os.environ.get("GEN_NETWORK")
    if mode:
        log("Tarmoq rejimi: %s" % mode)

    rc = 0
    jobs = []
    if args.only in (None, "area"):
        jobs.append(lambda: area.run(today, out_dir, log=log))
    if args.only in (None, "details"):
        jobs.append(lambda: run_details(args, today, out_dir))
    for job in jobs:
        try:
            rc |= job()
        except (ValueError, zipfile.BadZipFile) as e:
            log("XATO: %s" % e)
            rc |= 1
    return rc


if __name__ == "__main__":
    sys.exit(main())
