# -*- coding: utf-8 -*-
"""
CAS hisobot generatori - oynali ilova (Qashqadaryo HETK).

Sanani tanlaysiz, "Generatsiya qilish" ni bosasiz; jarayon foizi, jurnal, xatolar va yakuniy xabar oynada ko'rinadi.
Generatorning o'zi alohida jarayonda ishlaydi (main.py), shuning uchun oyna qotmaydi va "Bekor qilish" ishlaydi.

Ishga tushirish:   python generator\\app.py        yoki  build.bat orqali yig'ilgan  CAS Generator.exe
"""
import sys

if __name__ == "__main__" and "--cli" in sys.argv:       # oyna ichidan chaqiriladigan generator rejimi
    import main as _gen_main
    sys.exit(_gen_main.main(sys.argv[sys.argv.index("--cli") + 1:]))

import os
import socket
import subprocess
import time
from datetime import datetime, timedelta
from tkinter import messagebox

import customtkinter as ctk

import config
import guicore as gc
from main import parse_date

APP_TITLE = "CAS hisobot generatori  ·  Qashqadaryo HETK"
BLUE, GREEN, RED, AMBER, GREY = "#1F4E78", "#2E7D32", "#C62828", "#B26A00", "#6B7280"
MODES = {"Hammasi": None, "Faqat Area": "area", "Faqat Details": "details"}
LOCK_PORT = 47653


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        self.title(APP_TITLE)
        h = min(900, max(640, self.winfo_screenheight() - 90))
        self.geometry("980x%d" % h)
        self.minsize(900, 600)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self.proc = None
        self.tail = None
        self.started = 0.0
        self.canceled = False
        self.running = False
        self.warn_count = 0
        self.err_lines = []
        self.log_path = None
        self.day = None
        self.rows = {}

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(5, weight=1, minsize=110)
        self._build_header()
        self._build_date_card()
        self._build_inputs_card()
        self._build_run_card()
        self._build_banner()
        self._build_log()

        config.FILE_DIR.mkdir(parents=True, exist_ok=True)
        config.UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
        self.set_day(datetime.now().replace(hour=0, minute=0, second=0, microsecond=0))
        self.after(4000, self._idle_refresh)

    # ------------------------------------------------------------------ qurilish
    def _build_header(self):
        f = ctk.CTkFrame(self, fg_color=BLUE, corner_radius=0)
        f.grid(row=0, column=0, sticky="ew")
        ctk.CTkLabel(f, text="CAS hisobot generatori", font=ctk.CTkFont(size=20, weight="bold"),
                     text_color="white").pack(anchor="w", padx=22, pady=(10, 0))
        ctk.CTkLabel(f, text="Qashqadaryo HETK  ·  Area va Details hisobotlarini web ilova uchun tayyorlaydi",
                     text_color="#CFE0F1").pack(anchor="w", padx=22, pady=(0, 10))

    def _card(self, row, title):
        card = ctk.CTkFrame(self, corner_radius=10, fg_color="#F4F6F9", border_width=1, border_color="#DDE3EA")
        card.grid(row=row, column=0, sticky="ew", padx=18, pady=(10 if row == 1 else 6, 0))
        card.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(card, text=title, font=ctk.CTkFont(size=14, weight="bold"), text_color=BLUE
                     ).grid(row=0, column=0, columnspan=6, sticky="w", padx=14, pady=(8, 2))
        return card

    def _build_date_card(self):
        card = self._card(1, "1.  Qaysi sana bo'yicha generatsiya qilinsin?")
        row = ctk.CTkFrame(card, fg_color="transparent")
        row.grid(row=1, column=0, columnspan=6, sticky="ew", padx=14, pady=(0, 4))
        self.btn_prev = ctk.CTkButton(row, text="◀", width=36, command=lambda: self.shift_day(-1))
        self.btn_prev.pack(side="left")
        self.date_var = ctk.StringVar()
        self.date_var.trace_add("write", lambda *_: self.on_date_typed())
        self.entry = ctk.CTkEntry(row, width=150, height=34, justify="center", textvariable=self.date_var,
                                  font=ctk.CTkFont(size=17, weight="bold"))
        self.entry.pack(side="left", padx=6)
        self.btn_next = ctk.CTkButton(row, text="▶", width=36, command=lambda: self.shift_day(1))
        self.btn_next.pack(side="left")
        self.btn_today = ctk.CTkButton(row, text="Bugun", width=70, fg_color="#5B7FA6",
                                       command=lambda: self.set_day(self._today()))
        self.btn_today.pack(side="left", padx=(14, 4))
        self.btn_yest = ctk.CTkButton(row, text="Kecha", width=70, fg_color="#5B7FA6",
                                      command=lambda: self.set_day(self._today() - timedelta(days=1)))
        self.btn_yest.pack(side="left")
        self.mode_var = ctk.StringVar(value="Hammasi")
        self.seg = ctk.CTkSegmentedButton(row, values=list(MODES), variable=self.mode_var,
                                          command=lambda _v: self.refresh_inputs())
        self.seg.pack(side="right")
        ctk.CTkLabel(row, text="Hisobot:", text_color=GREY).pack(side="right", padx=(0, 8))
        self.date_info = ctk.CTkLabel(card, text="", anchor="w")
        self.date_info.grid(row=2, column=0, columnspan=6, sticky="w", padx=16, pady=(0, 8))

    def _build_inputs_card(self):
        card = self._card(2, "2.  Kiruvchi fayllar")
        card.grid_columnconfigure(0, weight=0)
        card.grid_columnconfigure(2, weight=1)
        for i, (key, title, _g) in enumerate(gc.INPUTS):
            icon = ctk.CTkLabel(card, text="●", width=24, font=ctk.CTkFont(size=15))
            name = ctk.CTkLabel(card, text=title, anchor="w", width=210, font=ctk.CTkFont(weight="bold"))
            det = ctk.CTkLabel(card, text="", anchor="w", text_color=GREY)
            icon.grid(row=i + 1, column=0, padx=(14, 2), pady=0)
            name.grid(row=i + 1, column=1, sticky="w")
            det.grid(row=i + 1, column=2, sticky="w", padx=8)
            self.rows[key] = (icon, name, det)
        bar = ctk.CTkFrame(card, fg_color="transparent")
        bar.grid(row=len(gc.INPUTS) + 1, column=0, columnspan=6, sticky="w", padx=12, pady=(4, 8))
        ctk.CTkButton(bar, text="Fayllar papkasini ochish", width=190, fg_color="#5B7FA6",
                      command=lambda: self.open_dir(config.FILE_DIR)).pack(side="left", padx=2)
        ctk.CTkButton(bar, text="Yangilash", width=100, fg_color="#5B7FA6",
                      command=self.refresh_inputs).pack(side="left", padx=6)
        self.inputs_hint = ctk.CTkLabel(bar, text="", text_color=GREY)
        self.inputs_hint.pack(side="left", padx=8)

    def _build_run_card(self):
        card = self._card(3, "3.  Generatsiya")
        card.grid_columnconfigure(0, weight=1)
        top = ctk.CTkFrame(card, fg_color="transparent")
        top.grid(row=1, column=0, sticky="ew", padx=14)
        top.grid_columnconfigure(0, weight=1)
        self.btn_run = ctk.CTkButton(top, text="▶   Generatsiya qilish", height=42, fg_color=GREEN, hover_color="#1B5E20",
                                     font=ctk.CTkFont(size=16, weight="bold"), command=self.start)
        self.btn_run.grid(row=0, column=0, sticky="ew")
        self.btn_cancel = ctk.CTkButton(top, text="Bekor qilish", height=42, width=130, fg_color=RED,
                                        hover_color="#8E0000", command=self.cancel, state="disabled")
        self.btn_cancel.grid(row=0, column=1, padx=(10, 0))
        info = ctk.CTkFrame(card, fg_color="transparent")
        info.grid(row=2, column=0, sticky="ew", padx=14, pady=(8, 0))
        info.grid_columnconfigure(0, weight=1)
        self.stage = ctk.CTkLabel(info, text="Tayyor. Sanani tanlang va tugmani bosing.", anchor="w")
        self.stage.grid(row=0, column=0, sticky="w")
        self.elapsed = ctk.CTkLabel(info, text="", text_color=GREY)
        self.elapsed.grid(row=0, column=1, sticky="e", padx=(0, 14))
        self.pct = ctk.CTkLabel(info, text="0%", font=ctk.CTkFont(size=15, weight="bold"), text_color=BLUE, width=46)
        self.pct.grid(row=0, column=2, sticky="e")
        self.bar = ctk.CTkProgressBar(card, height=16, progress_color=GREEN)
        self.bar.set(0)
        self.bar.grid(row=3, column=0, sticky="ew", padx=14, pady=(4, 12))

    def _build_banner(self):
        self.banner = ctk.CTkFrame(self, corner_radius=10, border_width=2)
        self.banner.grid_columnconfigure(0, weight=1)
        self.banner_title = ctk.CTkLabel(self.banner, text="", font=ctk.CTkFont(size=18, weight="bold"), anchor="w")
        self.banner_title.grid(row=0, column=0, sticky="w", padx=16, pady=(8, 0))
        self.banner_text = ctk.CTkLabel(self.banner, text="", anchor="w", justify="left", wraplength=880)
        self.banner_text.grid(row=1, column=0, sticky="w", padx=16, pady=(0, 4))
        bb = ctk.CTkFrame(self.banner, fg_color="transparent")
        bb.grid(row=2, column=0, sticky="w", padx=12, pady=(0, 8))
        ctk.CTkButton(bb, text="Natijalar papkasi (upload)", width=200,
                      command=lambda: self.open_dir(config.UPLOAD_DIR)).pack(side="left", padx=4)
        ctk.CTkButton(bb, text="Jurnal faylini ochish", width=160, fg_color="#5B7FA6",
                      command=self.open_log).pack(side="left", padx=4)

    def _build_log(self):
        f = ctk.CTkFrame(self, fg_color="transparent")
        f.grid(row=5, column=0, sticky="nsew", padx=18, pady=(8, 14))
        f.grid_columnconfigure(0, weight=1)
        f.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(f, text="Jurnal", font=ctk.CTkFont(size=13, weight="bold"), text_color=BLUE
                     ).grid(row=0, column=0, sticky="w")
        self.log = ctk.CTkTextbox(f, font=ctk.CTkFont(family="Consolas", size=12), wrap="none",
                                  fg_color="#101820", text_color="#D6DEE7", border_width=0)
        self.log.grid(row=1, column=0, sticky="nsew")
        for tag, color in (("err", "#FF6B6B"), ("warn", "#FFC857"), ("ok", "#6BE08B"), ("dim", "#7C8A99")):
            self.log.tag_config(tag, foreground=color)
        self.log.configure(state="disabled")

    # ------------------------------------------------------------------ sana
    @staticmethod
    def _today():
        return datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

    def set_day(self, day):
        self.date_var.set(day.strftime("%d.%m.%Y"))

    def shift_day(self, delta):
        base = self.day or self._today()
        self.set_day(min(base + timedelta(days=delta), self._today()))

    def on_date_typed(self):
        text = self.date_var.get()
        self.day = None
        try:
            day = parse_date(text)
            if day.date() > datetime.now().date():
                raise ValueError("kelajak sanasi")
            self.day = day
            wd = gc.UZ_WEEKDAYS[day.weekday()]
            ago = (self._today().date() - day.date()).days
            when = "bugun" if ago == 0 else ("kecha" if ago == 1 else "%d kun oldin" % ago)
            self.date_info.configure(text="%s, %s  ·  %s" % (wd, day.strftime("%d.%m.%Y"), when), text_color=GREEN)
            self.entry.configure(border_color="#8FB996")
        except ValueError as e:
            msg = "kelajak sanasini tanlab bo'lmaydi" if "kelajak" in str(e) else \
                "sana tushunarsiz. Format: KK.OO.YYYY  (masalan %s)" % self._today().strftime("%d.%m.%Y")
            self.date_info.configure(text=msg, text_color=RED)
            self.entry.configure(border_color=RED)
        self.refresh_inputs()

    # ------------------------------------------------------------------ fayllar
    def only(self):
        return MODES[self.mode_var.get()]

    def refresh_inputs(self):
        items = gc.scan_inputs(self.only(), self.day)
        self.items = items
        for it in items:
            icon, name, det = self.rows[it["key"]]
            if not it["needed"]:
                icon.configure(text="–", text_color=GREY)
                name.configure(text_color=GREY)
                det.configure(text="bu rejimda kerak emas", text_color="#A3AAB5")
                continue
            name.configure(text_color="black")
            if not it["ok"]:
                icon.configure(text="✖", text_color=RED)
                det.configure(text=it["detail"], text_color=RED)
            elif it["warn"]:
                icon.configure(text="⚠", text_color=AMBER)
                det.configure(text=it["detail"], text_color=AMBER)
            else:
                icon.configure(text="✔", text_color=GREEN)
                det.configure(text=it["detail"], text_color=GREY)
        warn = any(i["ok"] and i["warn"] for i in items if i["needed"])
        if not gc.inputs_ready(items):
            self.inputs_hint.configure(text="Yetishmayotgan fayllarni papkaga joylang", text_color=RED)
        elif warn:
            self.inputs_hint.configure(text="⚠ Ba'zi fayllar boshqa kunda yuklangan: tanlangan sanaga mosligini tekshiring",
                                       text_color=AMBER)
        else:
            self.inputs_hint.configure(text="Hamma fayllar joyida", text_color=GREEN)
        self.update_run_state()

    def _idle_refresh(self):
        if not self.running:
            self.refresh_inputs()
        self.after(4000, self._idle_refresh)

    def update_run_state(self):
        ok = (not self.running) and self.day is not None and gc.inputs_ready(getattr(self, "items", []))
        self.btn_run.configure(state="normal" if ok else "disabled")

    def set_running(self, running):
        self.running = running
        state = "disabled" if running else "normal"
        for w in (self.entry, self.btn_prev, self.btn_next, self.btn_today, self.btn_yest, self.seg):
            w.configure(state=state)
        self.btn_cancel.configure(state="normal" if running else "disabled", text="Bekor qilish")
        self.update_run_state()

    # ------------------------------------------------------------------ jurnal
    def log_add(self, text, tag=None):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n", tag) if tag else self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def log_clear(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    # ------------------------------------------------------------------ ishga tushirish
    def start(self):
        if self.running or self.day is None:
            return
        self.refresh_inputs()
        if not gc.inputs_ready(self.items):
            return
        mode = self.mode_var.get()
        d = self.day.strftime("%d.%m.%Y")
        if not messagebox.askyesno(
                "Tasdiqlang",
                "%s sanasi bo'yicha generatsiya qilinadi  (%s).\n\n"
                "Diqqat: upload papkasidagi eski fayllar o'chiriladi va qaytadan yaratiladi.\n\nDavom etamizmi?"
                % (d, mode.lower())):
            return
        logs = config.BASE_DIR / "logs"
        logs.mkdir(parents=True, exist_ok=True)
        self.log_path = logs / ("generatsiya_%s_(%s).log" % (datetime.now().strftime("%Y%m%d_%H%M%S"),
                                                             self.day.strftime("%d-%m-%Y")))
        cmd, env = gc.build_command(self.day, self.only(), self.log_path)
        try:
            self.proc = subprocess.Popen(cmd, env=env, cwd=str(config.BASE_DIR), **gc.popen_kwargs())
        except OSError as e:
            messagebox.showerror("Xato", "Generatorni ishga tushirib bo'lmadi:\n%s" % e)
            return
        self.tail = gc.LogTail(self.log_path)
        self.started = time.time()
        self.canceled = False
        self.warn_count = 0
        self.err_lines = []
        self.done_seen = False
        self.banner.grid_remove()
        self.log_clear()
        self.log_add("Sana: %s  ·  hisobot: %s" % (d, mode), "dim")
        self.bar.configure(progress_color=GREEN)
        self.bar.set(0)
        self.pct.configure(text="0%")
        self.stage.configure(text="Ishga tushirilmoqda...")
        self.set_running(True)
        self.after(150, self._tick)

    def _handle(self, lines):
        for line in lines:
            if not line.strip():
                continue
            kind = gc.classify(line)
            if kind[0] == "progress":
                self.bar.set(kind[1] / 100.0)
                self.pct.configure(text="%d%%" % int(kind[1]))
                if kind[2]:
                    self.stage.configure(text=kind[2])
            elif kind[0] == "done":
                self.done_seen = True
            else:
                _k, tag, text = kind
                if tag == "warn":
                    self.warn_count += 1
                if tag == "err" and "XATO" in text.upper():
                    self.err_lines.append(text.split("] ", 1)[-1])
                self.log_add(text, tag)

    def _tick(self):
        if self.proc is None:
            return
        self._handle(self.tail.read_new())
        self.elapsed.configure(text="Vaqt: %s" % gc.fmt_elapsed(time.time() - self.started))
        rc = self.proc.poll()
        if rc is None:
            self.after(150, self._tick)
            return
        self._handle(self.tail.read_new(final=True))
        self.finish(rc)

    def cancel(self):
        if not self.running or self.proc is None or self.canceled:
            return
        if not messagebox.askyesno("Bekor qilish", "Generatsiya to'xtatilsinmi?\n\n"
                                   "upload papkasi to'liq bo'lmay qoladi, keyin qaytadan generatsiya qilish kerak."):
            return
        self.canceled = True
        self.btn_cancel.configure(state="disabled", text="To'xtatilmoqda...")
        try:
            self.proc.terminate()
        except OSError:
            pass

    def finish(self, rc):
        elapsed = gc.fmt_elapsed(time.time() - self.started)
        self.proc = None
        gc.cleanup_tmp(config.UPLOAD_DIR)
        self.set_running(False)
        self.elapsed.configure(text="Umumiy vaqt: %s" % elapsed)
        if self.canceled:
            self.bar.configure(progress_color=AMBER)
            self.stage.configure(text="Bekor qilindi")
            self.log_add("Bekor qilindi.", "warn")
            self.show_banner(AMBER, "■  Generatsiya bekor qilindi",
                             "upload papkasi to'liq emas. Qaytadan generatsiya qiling.")
        elif rc == 0 and self.done_seen:
            files = gc.result_files(config.UPLOAD_DIR, self.day)
            self.bar.set(1)
            self.pct.configure(text="100%")
            self.stage.configure(text="Generatsiya qilindi")
            lines = ["%s   (%s)" % (n, gc.fmt_size(s)) for n, s in files] or ["(natija fayllari topilmadi)"]
            extra = "\n\n%d ta ogohlantirish bor - jurnalda sariq rangda ko'rinadi." % self.warn_count if self.warn_count else ""
            self.show_banner(GREEN, "✔  Generatsiya qilindi",
                             "Sana: %s   ·   vaqt: %s\n%s%s" % (self.day.strftime("%d.%m.%Y"), elapsed, "\n".join(lines), extra))
            messagebox.showinfo("Tayyor", "Generatsiya qilindi ✔\n\nSana: %s\nVaqt: %s\n\nFayllar upload papkasida."
                                % (self.day.strftime("%d.%m.%Y"), elapsed))
        else:
            self.bar.configure(progress_color=RED)
            self.stage.configure(text="Xato bilan tugadi")
            why = "\n".join(self.err_lines[-4:]) if self.err_lines else \
                "Dastur kutilmaganda to'xtadi (kod %s). Jurnalning oxirgi qatorlarini ko'ring." % rc
            self.show_banner(RED, "✖  Generatsiya bajarilmadi", why)
            messagebox.showerror("Xato", "Generatsiya bajarilmadi.\n\n%s" % why)
        self.refresh_inputs()

    def show_banner(self, color, title, text):
        self.banner.configure(border_color=color, fg_color="#FFFFFF")
        self.banner_title.configure(text=title, text_color=color)
        self.banner_text.configure(text=text)
        self.banner.grid(row=4, column=0, sticky="ew", padx=18, pady=(8, 0))

    # ------------------------------------------------------------------ yordamchi
    def open_dir(self, path):
        try:
            path.mkdir(parents=True, exist_ok=True)
            gc.open_path(path)
        except Exception as e:
            messagebox.showerror("Xato", "Papkani ochib bo'lmadi:\n%s" % e)

    def open_log(self):
        try:
            if self.log_path and self.log_path.exists():
                gc.open_path(self.log_path)
        except Exception as e:
            messagebox.showerror("Xato", "Jurnalni ochib bo'lmadi:\n%s" % e)

    def on_close(self):
        if self.running and self.proc is not None:
            if not messagebox.askyesno("Chiqish", "Generatsiya ketmoqda. To'xtatib chiqilsinmi?"):
                return
            try:
                self.proc.terminate()
            except OSError:
                pass
            gc.cleanup_tmp(config.UPLOAD_DIR)
        self.destroy()


def _single_instance():
    """Ikkinchi nusxa ochilsa, upload papkasi bir vaqtda ikki marta tozalanib ketmasin."""
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.bind(("127.0.0.1", LOCK_PORT))
    except OSError:
        return None
    return s


def _crash(exc_text):
    """Kutilmagan xatoni logs\\oyna_xato.log ga yozadi va oynada ko'rsatadi (yopilib qolmasligi uchun)."""
    path = None
    try:
        logs = config.BASE_DIR / "logs"
        logs.mkdir(parents=True, exist_ok=True)
        path = logs / "oyna_xato.log"
        with open(path, "a", encoding="utf-8") as f:
            f.write("=== %s ===\n%s\n" % (datetime.now().strftime("%d.%m.%Y %H:%M:%S"), exc_text))
    except Exception:
        pass
    try:
        import tkinter
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror("Oyna ishga tushmadi", "Xato:\n\n%s\n\nBatafsil: %s" % (exc_text[-900:], path or "-"))
        root.destroy()
    except Exception:
        pass


def main():
    lock = _single_instance()
    if lock is None:
        root = ctk.CTk()
        root.withdraw()
        messagebox.showwarning("Allaqachon ochiq", "CAS generatori allaqachon ishlayapti.")
        return 1
    try:
        app = App()
        app.mainloop()
    except Exception:
        import traceback
        _crash(traceback.format_exc())
        return 2
    finally:
        lock.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
