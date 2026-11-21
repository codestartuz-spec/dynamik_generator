# -*- coding: utf-8 -*-
"""Oyna ilovasining oynasiz qismi: fayllarni tekshirish, generatorni ishga tushirish buyrug'i, jurnalni o'qish.
(Oynadan ajratilgan, shunda alohida sinab ko'rish mumkin.)"""
import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path

import config
import progress
from fsutil import list_xlsx, find_reports

UZ_WEEKDAYS = ["Dushanba", "Seshanba", "Chorshanba", "Payshanba", "Juma", "Shanba", "Yakshanba"]

# (kalit, nomi, qaysi rejimda kerak)
INPUTS = [
    ("archive", "Meter Archive Details", "details"),
    ("metering", "MeteringDevicesReport", "details"),
    ("direct", "DirectMeter", "details"),
    ("area", "AreaInformationReport", "area"),
    ("comm", "CommissioningStatus", "area"),
    ("online", "Online_Device_Monitroing_Report", "area"),
]


def _date_str(ts):
    return datetime.fromtimestamp(ts).strftime("%d.%m.%Y")


def _scan_files(files, day):
    """(detail matni, sana farq qiladimi) - fayllarning o'zgargan sanasi tanlangan sanaga mosmi."""
    stamps = [p.stat().st_mtime for p in files]
    lo, hi = _date_str(min(stamps)), _date_str(max(stamps))
    dates = lo if lo == hi else "%s - %s" % (lo, hi)
    mismatch = day is not None and any(datetime.fromtimestamp(t).date() != day.date() for t in stamps)
    return dates, mismatch


def scan_inputs(only, day):
    """
    Kiruvchi fayllar holati. Har element: key, title, needed (tanlangan rejimda kerakmi), ok, detail, warn.
    only: None / 'area' / 'details'; day: tanlangan sana (datetime) yoki None.
    """
    prefixes = {"direct": config.DIRECT_PREFIXES, "area": config.AREA_PREFIXES,
                "comm": config.COMMISSIONING_PREFIXES, "online": config.ONLINE_PREFIXES}
    out = []
    for key, title, group in INPUTS:
        needed = only is None or only == group
        item = {"key": key, "title": title, "needed": needed, "ok": False, "detail": "", "warn": False}
        try:
            if key in ("archive", "metering"):
                folder = config.ARCHIVE_DIR if key == "archive" else config.METERING_DIR
                files = list_xlsx(folder)
                if files:
                    dates, mism = _scan_files(files, day)
                    item.update(ok=True, warn=mism, detail="%d ta fayl  ·  yuklangan: %s" % (len(files), dates))
                else:
                    item["detail"] = "papkada .xlsx fayl yo'q  (%s)" % folder.name
            else:
                files = find_reports(config.FILE_DIR, prefixes[key])
                if files:
                    dates, mism = _scan_files(files[:1], day)
                    extra = "  (yana %d ta bor, eng yangisi olinadi)" % (len(files) - 1) if len(files) > 1 else ""
                    item.update(ok=True, warn=mism, detail="%s  ·  yuklangan: %s%s" % (files[0].name, dates, extra))
                else:
                    item["detail"] = "%s*.xlsx topilmadi" % prefixes[key][0]
        except OSError as e:
            item["detail"] = "o'qib bo'lmadi: %s" % e
        out.append(item)
    return out


def inputs_ready(items):
    return all(i["ok"] for i in items if i["needed"])


def app_script():
    return str(Path(__file__).resolve().parent / "app.py")


def build_command(day, only, log_path):
    """Generatorni alohida jarayon sifatida ishga tushirish: (buyruq, muhit)."""
    args = ["--cli", "--date", day.strftime("%Y-%m-%d"), "--log-file", str(log_path)]
    if only:
        args += ["--only", only]
    if getattr(sys, "frozen", False):
        cmd = [sys.executable] + args
    else:
        cmd = [sys.executable, app_script()] + args
    env = dict(os.environ)
    env["GEN_PROGRESS"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUNBUFFERED"] = "1"
    return cmd, env


def popen_kwargs():
    kw = {"stdin": subprocess.DEVNULL, "stdout": subprocess.DEVNULL, "stderr": subprocess.DEVNULL}
    if sys.platform.startswith("win"):
        kw["creationflags"] = 0x08000000      # CREATE_NO_WINDOW: qora oyna chiqmasin
    return kw


class LogTail(object):
    """Jurnal faylidan faqat yangi, TO'LIQ qatorlarni o'qiydi (yarim yozilgan qator keyingi safar o'qiladi)."""

    def __init__(self, path):
        self.path = str(path)
        self._pos = 0
        self._buf = b""

    def read_new(self, final=False):
        lines = []
        try:
            with open(self.path, "rb") as f:
                f.seek(self._pos)
                chunk = f.read()
                self._pos += len(chunk)
        except OSError:
            return lines
        self._buf += chunk
        parts = self._buf.split(b"\n")
        self._buf = parts.pop()
        if final and self._buf:
            parts.append(self._buf)
            self._buf = b""
        for raw in parts:
            lines.append(raw.decode("utf-8", "replace").rstrip("\r"))
        return lines


def classify(line):
    """('progress', foiz, matn) | ('done',) | ('log', teg, matn). Teg: err / warn / ok / None."""
    if line.startswith(progress.PREFIX):
        body = line[len(progress.PREFIX):]
        pct, _sep, text = body.partition("|")
        try:
            return ("progress", float(pct), text)
        except ValueError:
            return ("log", None, line)
    if line.startswith(progress.DONE):
        return ("done",)
    low = line.lower()
    if "xato" in low or "traceback" in low or low.startswith("  file ") or "error" in low[:40]:
        return ("log", "err", line)
    if "diqqat" in low:
        return ("log", "warn", line)
    if "tayyor" in low:
        return ("log", "ok", line)
    return ("log", None, line)


def result_files(upload_dir, day):
    """Shu sanaga tegishli tayyor fayllar: [(nomi, hajmi bayt)]."""
    date_s = day.strftime(config.OUTPUT_DATE_FORMAT)
    out = []
    try:
        for p in sorted(Path(upload_dir).iterdir()):
            if p.is_file() and date_s in p.name and not p.name.endswith(".tmp"):
                out.append((p.name, p.stat().st_size))
    except OSError:
        pass
    return out


def cleanup_tmp(upload_dir):
    n = 0
    try:
        for p in Path(upload_dir).glob("*.tmp"):
            try:
                p.unlink()
                n += 1
            except OSError:
                pass
    except OSError:
        pass
    return n


def fmt_size(n):
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return ("%d %s" % (n, unit)) if unit == "B" else ("%.1f %s" % (n, unit))
        n /= 1024.0


def fmt_elapsed(sec):
    sec = int(sec)
    return "%02d:%02d" % (sec // 60, sec % 60) if sec < 3600 else "%d:%02d:%02d" % (sec // 3600, sec % 3600 // 60, sec % 60)


def open_path(path):
    """Papka yoki faylni tizim dasturida ochadi."""
    path = str(path)
    if sys.platform.startswith("win"):
        os.startfile(path)                                 # noqa: Windows
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])
