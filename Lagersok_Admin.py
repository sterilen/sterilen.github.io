#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Lagersök Admin - Databasuppdatering (fönsterprogram)
Uppdaterad med Hjälp-flik och Avvikelser-flik uppdelad per kategori.

Samma motor som uppdatera_databas.py, men med flikar:
- Huvud: befintlig funktionalitet
- Hjälp: förklarar hur programmet fungerar
- Avvikelser: visar avvikelser uppdelat i Fabrikssterilt / Instrumentgaller / Bulk
"""
import argparse
import datetime
import glob
import os
import queue
import re
import subprocess
import sys
import threading
import tkinter as tk
import tkinter.font as tkfont
import webbrowser
from tkinter import filedialog, messagebox, ttk

import uppdatera_databas as U

# ===== FIX: Krypterade inställningar (installningar.dat bredvid exe-filen) =====
# Första starten skapas filen från STANDARD nedan. Filen är krypterad (kan inte läsas i Anteckningar)
# men inte knuten till dator/användare - program + installningar.dat kan kopieras till en annan dator.
import base64
import hashlib
import hmac
import json

# ----- Standardvärden (följer med i exe-filen) -----
STANDARD = {
    "pdf": r"G:\Sos\PVI_Steril\Förrådet\Pelle\projekt\LagerApp\lager.pdf",
    "galler": r"G:\Sos\PVI_Steril\Förrådet\Mallar förrådet\Galler på Reol",
    "bulk": r"G:\Sos\PVI_Steril\Förrådet\Pelle\projekt\LagerApp\artiklar_bulk.xlsx",
    "ut": r"G:\Sos\PVI_Steril\Förrådet\Pelle\projekt\LagerApp\data.js",
    "sparr_procent": "20",
    "publicera": "api",
    "github_repo": "sterilen/Sterilen",
    "github_branch": "main",
    "statistik_url": "https://sterilsok-stat.pelle-holmstrom.workers.dev",
    "statistik_nyckel": "Dce0nlyz!Dce0nlyz!",
    "github_token": "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",
    "stad_intervall": "7",
    "stad_rader": "A,B,C,D,E,F,G,H,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,BG,BPA,GALLER",
    "stadlogg": r"G:\Sos\PVI_Steril\Förrådet\Pelle\projekt\LagerApp\stadlogg.json",
    "bilder_github_path": "bilder",
    "etikett_metod": "usb",
    "etikett_skrivare": "",
    "etikett_ip": "",
    "etikett_dpi": "203",
    "etikett_bredd": "70.0",
    "etikett_hojd": "50.8",
    "skannade_mapp": r"G:\Sos\PVI_Steril\Förrådet\Scannade_Följesedlar",
    "larande_mapp": r"G:\Sos\PVI_Steril\Förrådet\Följesedel_Learning",
    "bild_url": "https://lagersok-bilder.pelle-holmstrom.workers.dev",          # t.ex. https://lagersok-bilder.dittnamn.workers.dev
    "bild_nyckel": "h7Kq2mXv9Rt4Wz8Ld3Pn6Bc5",                   # samma värde som ADMIN_KEY i Cloudflare-tjänsten
}
# ------------------------------------------------------------------

_NYCKLAR = tuple(STANDARD)
_MAGI = b"LSK2"
_HEMLIGHET = b"Lagersok/Sterilcentrum/installningar/v2"


def _get_base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _fil():
    return os.path.join(_get_base_dir(), "installningar.dat")


def _get_fernet():
    try:
        from cryptography.fernet import Fernet
    except ImportError:
        raise ImportError("cryptography saknas - kör: pip install cryptography")
    raw = hashlib.sha256(_HEMLIGHET).digest()
    key = base64.urlsafe_b64encode(raw)
    return Fernet(key)

def _ström(nyckel, nonce, n):
    ut = b""
    i = 0
    while len(ut) < n:
        ut += hashlib.sha256(nyckel + nonce + i.to_bytes(4, "big")).digest()
        i += 1
    return ut[:n]

def _dekryptera_gammal(rå):
    if not rå.startswith(_MAGI) or len(rå) < 52:
        raise ValueError("okänt filformat")
    nonce, mac, ct = rå[4:20], rå[20:52], rå[52:]
    nyckel = hashlib.sha256(_HEMLIGHET).digest()
    if not hmac.compare_digest(mac, hmac.new(nyckel, nonce + ct, hashlib.sha256).digest()):
        raise ValueError("filen är skadad eller ändrad")
    return bytes(a ^ b for a, b in zip(ct, _ström(nyckel, nonce, len(ct)))).decode("utf-8")

def _kryptera(text):
    f = _get_fernet()
    return f.encrypt(text.encode("utf-8"))

def _dekryptera(rå):
    try:
        f = _get_fernet()
        return f.decrypt(rå).decode("utf-8")
    except Exception:
        try:
            return _dekryptera_gammal(rå)
        except Exception as e:
            raise ValueError(f"Kunde inte dekryptera: {e}")



def _las_txt(sokvag):
    d = {}
    with open(sokvag, "r", encoding="utf-8-sig", errors="ignore") as fh:
        for rad in fh:
            rad = rad.strip()
            if rad and not rad.startswith("#") and "=" in rad:
                k, v = rad.split("=", 1)
                d[k.strip()] = v.strip().strip('"')
    return d


def _skriv(data_dict):
    d = {k: str(data_dict.get(k, "") or "") for k in _NYCKLAR}
    d["sparr_procent"] = d["sparr_procent"] or "20"
    d["github_branch"] = d["github_branch"] or "main"
    fil = _fil()
    tmp = fil + ".tmp"
    with open(tmp, "wb") as fh:
        fh.write(_kryptera(json.dumps(d, ensure_ascii=False)))
    os.replace(tmp, fil)
    return fil



def _las_installningar_dict_fix():
    fil = _fil()
    if os.path.isfile(fil):
        try:
            with open(fil, "rb") as fh:
                return json.loads(_dekryptera(fh.read()))
        except Exception as e:
            logga(f"Kunde inte läsa installningar.dat ({e!r}) - använder standardvärden")
            return dict(STANDARD)
    gammal = os.path.join(_get_base_dir(), "installningar.txt")
    if os.path.isfile(gammal):
        try:
            d = _las_txt(gammal)
            _skriv(d)
            os.remove(gammal)
            logga("installningar.txt flyttades till krypterad installningar.dat och raderades")
            return d
        except Exception as e:
            logga(f"Kunde inte flytta gamla inställningar: {e!r}")
    try:                                  # första starten: skapa från standardvärdena
        _skriv(STANDARD)
        logga("installningar.dat skapades från standardvärden")
    except Exception as e:
        logga(f"Kunde inte skapa installningar.dat: {e!r}")
    return dict(STANDARD)


U.las_installningar_dict = _las_installningar_dict_fix
U.skriv_installningar = _skriv

# ===== Felsökningslogg: AVSTÄNGD – ingen debug-fil skapas =====
# _LOGGFIL är borttagen medvetet, logga() skriver ingenting till disk.

def logga(text):
    return

def logg_start():
    return
# ===== SLUT FIX =====


VERSION = "v1.0"

# ------------------------------------------------------------------ färger
BAKGRUND = '#eef1f5'
KORT = '#ffffff'
RAM = '#dde2ea'
TEXT = '#1f2937'
DAMPAD = '#6b7280'
MORK = '#18222f'
GRON, BLA, ROSA = '#7fb046', '#4687b0', '#b04671'
GRON_MORK = '#5e8f2c'
ROD, GUL = '#c0392b', '#c98a0b'
LOGG_BG, LOGG_TEXT = '#111827', '#d1d5db'


def dpi_pa():
    if os.name == 'nt':
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass


def oppna(sokvag):
    try:
        if os.name == 'nt':
            os.startfile(sokvag)          # noqa
        elif sys.platform == 'darwin':
            subprocess.Popen(['open', sokvag])
        else:
            subprocess.Popen(['xdg-open', sokvag])
    except Exception:
        pass


def tid(sek):
    return datetime.datetime.fromtimestamp(sek).strftime('%Y-%m-%d %H:%M')


def hitta(sokvag, monster):
    """Visar vilken fil som skulle användas. Ger (text, färg)."""
    sokvag = (sokvag or '').strip().strip('"')
    if not sokvag:
        return "Ingen vald", DAMPAD
    try:
        if os.path.isdir(sokvag):
            filer = [f for f in glob.glob(os.path.join(sokvag, monster))
                     if not os.path.basename(f).startswith('~$')]
            if not filer:
                return f"Hittade ingen {monster}-fil i mappen", ROD
            nyast = max(filer, key=os.path.getmtime)
            return f"Använder nyaste filen: {os.path.basename(nyast)}  (ändrad {tid(os.path.getmtime(nyast))})", GRON_MORK
        if os.path.isfile(sokvag):
            return f"{os.path.basename(sokvag)}  (ändrad {tid(os.path.getmtime(sokvag))})", GRON_MORK
    except OSError as e:
        return "Kan inte läsa: " + str(e), ROD
    return "Hittar inte filen eller mappen", ROD


def hitta_galler(mapp):
    mapp = (mapp or '').strip().strip('"')
    if not mapp:
        return "Ingen vald", DAMPAD
    if not os.path.isdir(mapp):
        return "Hittar inte mappen", ROD
    filer = []
    for mon in ('*.xlsx', '*.xlsm'):
        filer += glob.glob(os.path.join(mapp, '**', mon), recursive=True)
    filer = [f for f in filer if not os.path.basename(f).startswith('~$')]
    if not filer:
        return "Hittade inga Excel-filer i mappen", ROD
    return f"{len(filer)} Excel-filer hittades (även i undermappar)", GRON_MORK


def _norm_stat(typ, v):
    """Slår ihop stavningsvarianter i enhetsstatistiken (t.ex. dator/Dator, Mac/macOS, Okänd/Annat)."""
    s = str(v or '').strip()
    if not s:
        return 'Okänd'
    k = s.lower()
    if typ == 'enhet':
        return {'mobil': 'Mobil', 'dator': 'Dator', 'surfplatta': 'Surfplatta'}.get(k, s[:1].upper() + s[1:].lower())
    if typ == 'os':
        return {'windows': 'Windows', 'android': 'Android', 'ios': 'iOS', 'mac': 'macOS', 'macos': 'macOS',
                'linux': 'Linux', 'chromeos': 'ChromeOS', 'okänd': 'Annat', 'annat': 'Annat'}.get(k, s)
    if typ == 'webb':
        return {'chrome': 'Chrome', 'edge': 'Edge', 'safari': 'Safari', 'firefox': 'Firefox', 'opera': 'Opera',
                'samsung': 'Samsung', 'okänd': 'Annan', 'annan': 'Annan'}.get(k, s)
    return s


class Knapp(tk.Label):
    """Platt knapp med hover-effekt."""
    def __init__(self, foralder, text, kommando, bg, fg='#ffffff', hover=None, stor=False, **kw):
        fs = (FONT_FET, 12 if stor else 10)
        super().__init__(foralder, text=text, bg=bg, fg=fg, font=fs, cursor='hand2',
                         padx=22 if stor else 14, pady=10 if stor else 6, **kw)
        self.bg, self.fg, self.hov = bg, fg, hover or bg
        self.kommando = kommando
        self.aktiv = True
        self.bind('<Enter>', lambda e: self.aktiv and self.config(bg=self.hov))
        self.bind('<Leave>', lambda e: self.aktiv and self.config(bg=self.bg))
        self.bind('<Button-1>', lambda e: self.aktiv and self.kommando())

    def satt_aktiv(self, pa):
        self.aktiv = pa
        self.config(bg=self.bg if pa else '#c5ccd6', fg='#ffffff' if pa else '#8a94a3',
                    cursor='hand2' if pa else 'arrow')



class ScrollableFrame(tk.Frame):
    """Scrollbar som funkar med mus och laptop-touchpad.

    Tk på Windows skickar mushjulshändelser till det widget som har tangentbordsfokus,
    inte till det som musen står över. Därför fångas alla hjul-/touchpad-händelser här
    på programnivå och skickas till det som ligger under muspekaren."""
    _global_installerad = False
    _acc = 0.0
    _rot = None

    def __init__(self, parent, bg=BAKGRUND, visa_scroll=True, **kw):
        super().__init__(parent, bg=bg, **kw)
        self.canvas = tk.Canvas(self, bg=bg, highlightthickness=0, bd=0)
        self.v_scroll = ttk.Scrollbar(self, orient='vertical', command=self.canvas.yview)
        self.inner = tk.Frame(self.canvas, bg=bg)

        self.inner_id = self.canvas.create_window((0, 0), window=self.inner, anchor='nw')
        self.canvas.configure(yscrollcommand=self.v_scroll.set)

        self.canvas.pack(side='left', fill='both', expand=True)
        if visa_scroll:
            self.v_scroll.pack(side='right', fill='y')

        self.inner.bind("<Configure>", self._on_inner_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        ScrollableFrame._installera_global(self)

    def _on_inner_configure(self, e):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, e):
        self.canvas.itemconfig(self.inner_id, width=e.width)

    @property
    def scrollable_frame(self):
        return self.inner

    # ---- global hantering av mushjul / touchpad
    @staticmethod
    def _installera_global(widget):
        if ScrollableFrame._global_installerad:
            return
        ScrollableFrame._global_installerad = True
        rot = widget.winfo_toplevel()
        ScrollableFrame._rot = rot
        for hj, fn in (("<MouseWheel>", ScrollableFrame._hjul), ("<Button-4>", ScrollableFrame._hjul),
                       ("<Button-5>", ScrollableFrame._hjul), ("<TouchpadScroll>", ScrollableFrame._touchpad)):
            try:
                rot.bind_all(hj, fn)
                # Text och rullist har egen standardhantering som skulle ge dubbel scroll
                for klass in ("Text", "TScrollbar", "Scrollbar"):
                    rot.bind_class(klass, hj, fn)
            except tk.TclError:
                pass   # äldre Tk saknar <TouchpadScroll> (finns från Tk 9)

    @staticmethod
    def _under_pekaren(event):
        try:
            nod = ScrollableFrame._rot.winfo_containing(event.x_root, event.y_root)
        except Exception:
            nod = None
        if nod is None and not isinstance(event.widget, str):
            nod = event.widget
        return nod

    @staticmethod
    def _scrolla(nod, rader=0, pixlar=0):
        """Går uppåt från widgeten under pekaren tills något kan scrolla. rader = rader, pixlar = bildpunkter (+ = nedåt)."""
        while nod is not None:
            if isinstance(nod, tk.Text):
                try:
                    if nod.yview() != (0.0, 1.0):
                        if pixlar:
                            nod.yview_scroll(int(pixlar), "pixels")
                        else:
                            nod.yview_scroll(rader, "units")
                        return
                except Exception:
                    pass
            elif isinstance(nod, ScrollableFrame):
                c = nod.canvas
                bb = c.bbox("all")
                total = (bb[3] - bb[1]) if bb else 0
                if total > c.winfo_height():
                    if pixlar:
                        c.yview_moveto(c.yview()[0] + pixlar / float(total))
                    else:
                        c.yview_scroll(rader, "units")
                return
            nod = getattr(nod, 'master', None)

    @staticmethod
    def _hjul(event):
        """Vanligt mushjul (och touchpad i Tk 8.x)."""
        try:
            num = getattr(event, 'num', None)
            if num == 4:
                steg = -3
            elif num == 5:
                steg = 3
            else:
                d = event.delta
                if d == 0:
                    return "break"
                if sys.platform == 'darwin':
                    steg = -int(d)
                else:
                    # 120 = ett mushjulsklick = 3 rader. Små värden samlas ihop.
                    ScrollableFrame._acc += -d / 40.0
                    steg = int(ScrollableFrame._acc)
                    ScrollableFrame._acc -= steg
            nod = ScrollableFrame._under_pekaren(event)
            if steg:
                ScrollableFrame._scrolla(nod, rader=steg)
        except Exception as e:
            logga(f"Scrollfel (hjul): {e!r}")
        return "break"

    @staticmethod
    def _touchpad(event):
        """Tvåfingersscroll på precisionstouchpad (Tk 9 skickar <TouchpadScroll>, inte <MouseWheel>)."""
        try:
            d = int(event.delta)
            try:
                res = ScrollableFrame._rot.tk.splitlist(ScrollableFrame._rot.tk.call('tk::PreciseScrollDeltas', d))
                dy = int(float(res[1]))
            except Exception:
                dy = d & 0xffff
                if dy >= 0x8000:
                    dy -= 0x10000
            nod = ScrollableFrame._under_pekaren(event)
            if dy:
                # i Tk är positivt dy = innehållet flyttas nedåt, dvs scrolla uppåt
                ScrollableFrame._scrolla(nod, pixlar=-dy)
        except Exception as e:
            logga(f"Scrollfel (touchpad): {e!r}")
        return "break"


class SplashScreen(tk.Toplevel):
    """Splash screen"""
    def __init__(self, parent, on_close):
        super().__init__(parent)
        self.parent = parent
        self.on_close = on_close
        self.overrideredirect(True)
        self.configure(bg=MORK)
        self.attributes('-topmost', True)
        # Centrera
        w, h = 460, 300
        sw = parent.winfo_screenwidth()
        sh = parent.winfo_screenheight()
        x = (sw - w)//2
        y = (sh - h)//2
        self.geometry(f"{w}x{h}+{x}+{y}")

        # Skugga / ram
        outer = tk.Frame(self, bg='#0f1720')
        outer.pack(fill='both', expand=True, padx=1, pady=1)
        inner = tk.Frame(outer, bg=MORK)
        inner.pack(fill='both', expand=True)

        # Toppfärg
        tk.Frame(inner, bg=GRON, height=4).pack(fill='x')

        center = tk.Frame(inner, bg=MORK)
        center.pack(fill='both', expand=True, padx=30, pady=20)

        # Ikon / logo
        icon_fr = tk.Frame(center, bg=MORK)
        icon_fr.pack(pady=(10,10))
        # Cirkel med bokstav
        canv = tk.Canvas(icon_fr, width=70, height=70, bg=MORK, highlightthickness=0)
        canv.pack()
        canv.create_oval(2,2,68,68, fill=GRON, outline='')
        canv.create_text(35, 36, text="L", fill='white', font=(FONT_FET, 28, 'bold'))

        tk.Label(center, text="Lagersök Admin", bg=MORK, fg='white', font=(FONT_FET, 26, 'bold')).pack()
        tk.Label(center, text=f"Databasuppdatering {VERSION}", bg=MORK, fg='#9fb0c3', font=(FONT, 11)).pack(pady=(4,0))

        # Progress bar animation
        self.prog_fr = tk.Frame(center, bg='#1e2a38', height=4)
        self.prog_fr.pack(fill='x', pady=(30,10))
        self.prog_bar = tk.Frame(self.prog_fr, bg=GRON, width=0, height=4)
        self.prog_bar.place(x=0, y=0, relheight=1)

        self.status = tk.Label(center, text="Laddar...", bg=MORK, fg=DAMPAD, font=(FONT, 9))
        self.status.pack()


        self.progress = 0
        self.animate()

    def animate(self):
        msgs = ["Laddar...", "Läser inställningar...", "Initierar...", "Klart!"]
        if self.progress < 100:
            self.progress += 4 if self.progress < 80 else 8
            if self.progress > 100:
                self.progress = 100
            # uppdatera bar
            self.prog_bar.config(width=int(400 * self.progress/100))
            if self.progress < 30:
                self.status.config(text=msgs[0])
            elif self.progress < 60:
                self.status.config(text=msgs[1])
            elif self.progress < 90:
                self.status.config(text=msgs[2])
            else:
                self.status.config(text=msgs[3])
            self.after(30, self.animate)
        else:
            self.after(250, self.fade_out)

    def fade_out(self):
        # Fade out animation
        try:
            alpha = self.attributes('-alpha')
        except Exception:
            alpha = 1.0
            try:
                self.attributes('-alpha', 1.0)
            except Exception:
                self.close_final()
                return
        if alpha <= 0.05:
            self.close_final()
        else:
            try:
                self.attributes('-alpha', max(0, alpha - 0.08))
            except Exception:
                self.close_final()
                return
            self.after(25, self.fade_out)

    def close_final(self):
        try:
            self.destroy()
        except Exception:
            pass
        self.on_close()

    def close(self):
        self.close_final()


FONT = 'TkDefaultFont'
FONT_FET = 'TkDefaultFont'


class App:
    def __init__(self, rot):
        self.rot = rot
        rot.title(f"Lagersök Admin - Databasuppdatering {VERSION}")
        rot.configure(bg=BAKGRUND)
        rot.geometry("1024x700")
        rot.minsize(960, 640)
        self.kor_pagar = False
        self.avbryt = False
        self.ko = queue.Queue()
        self.steg_nr, self.steg_tot = 0, 1
        self.sparr_svar = None
        self.v = {}
        self.byggt = False
        self.fonter()
        self.vars()
        self.bygg()
        self.las_in(U.las_installningar_dict())
        try:
            import os as _os
            _fil_dat = os.path.join(_get_base_dir(), "installningar.dat")
            if not _os.path.isfile(_fil_dat):
                U.skriv_installningar(self.till_dict())
        except Exception:
            pass
        self.uppdatera_status()
        rot.protocol("WM_DELETE_WINDOW", self.stang)
        rot.after(60, self.pumpa)

    # ---------------------------------------------------------------- uppsättning
    def fonter(self):
        global FONT, FONT_FET
        fam = set(tkfont.families())
        FONT = FONT_FET = tkfont.nametofont('TkDefaultFont').actual('family')
        for kandidat in ('Segoe UI', 'Calibri', 'Arial', 'Helvetica Neue', 'Inter', 'Noto Sans', 'DejaVu Sans'):
            if kandidat in fam:
                FONT = FONT_FET = kandidat
                break
        self.f_norm = (FONT, 11)
        self.f_fet = (FONT_FET, 10, 'bold')
        self.f_liten = (FONT, 10)
        self.f_rubrik = (FONT_FET, 13, 'bold')
        self.f_rubrik_stor = (FONT_FET, 15, 'bold')
        self.f_mono = ('Consolas' if 'Consolas' in fam else 'DejaVu Sans Mono' if 'DejaVu Sans Mono' in fam else 'Courier', 10)

    def vars(self):
        S = tk.StringVar
        self.src = S(value='pdf')  # endast PDF nu
        self.p_forbruk = S()
        self.p_galler = S()
        self.bulk_pa = tk.BooleanVar(value=True)
        self.p_bulk = S()
        self.p_ut = S()
        self.p_skannade = S(value=r"G:\Sos\PVI_Steril\Förrådet\Scannade_Följesedlar")
        self.p_larande = S(value=STANDARD["larande_mapp"])
        self.p_bild_url = S(value="")
        self.p_bild_nyckel = S(value="")        # fylls i från installningar.dat / STANDARD, har ingen ruta i fönstret
        self.sparr = S(value='20')
        self.publ = S(value='api')
        self.repo = S()
        self.gren = S(value='main')
        self.token = S()
        self.visa_token = tk.BooleanVar(value=False)
        self.stat_url = S()
        self.stat_key = S()
        self.stat_dagar = S(value='30')
        self.p_stad_intervall = S(value='7')
        self.p_et_metod = S(value='usb')
        self.p_et_skrivare = S()
        self.p_et_ip = S()
        self.p_et_dpi = S(value='203')
        self.p_et_bredd = S(value='70.0')
        self.p_et_hojd = S(value='50.8')
        self.p_stad_rader = S(value=STANDARD.get("stad_rader",""))
        self.p_stadlogg = S(value=STANDARD.get("stadlogg",""))
        for var in (self.src, self.p_forbruk, self.p_galler, self.bulk_pa, self.p_bulk):
            var.trace_add('write', lambda *a: self.uppdatera_status())
        self.nytt_stod = None
        # Avvikelser-data
        self.avv_data = {"fabrikssterilt": [], "instrumentgaller": [], "bulk": [], "utanplats": [], "ovrigt": []}

    def las_in(self, d):
        # Endast PDF-läsning för fabrikssterilt
        self.src.set('pdf')
        # Bakåtkompatibel: om gammal artiklar-nyckel finns, använd den som fallback
        pdf_val = d.get('pdf') or d.get('artiklar') or ''
        self.p_forbruk.set(pdf_val)
        self.p_galler.set(d.get('galler', ''))
        self.p_bulk.set(d.get('bulk', ''))
        self.bulk_pa.set(True)  # bulk alltid med
        self.p_ut.set(d.get('ut', ''))
        self.p_skannade.set(d.get('skannade_mapp', r"G:\Sos\PVI_Steril\Förrådet\Scannade_Följesedlar") or r"G:\Sos\PVI_Steril\Förrådet\Scannade_Följesedlar")
        self.p_larande.set(d.get('larande_mapp', STANDARD['larande_mapp']) or STANDARD['larande_mapp'])
        _bu = d.get('bild_url', STANDARD['bild_url']) or STANDARD['bild_url']      # saknas i äldre installningar.dat: ta värdet ur STANDARD
        _bn = d.get('bild_nyckel', STANDARD['bild_nyckel']) or STANDARD['bild_nyckel']
        self.p_bild_url.set('' if _bu.startswith('KLISTRA_IN') else _bu)
        self.p_bild_nyckel.set('' if _bn.startswith('KLISTRA_IN') else _bn)
        self.sparr.set(d.get('sparr_procent', '20') or '20')
        pub = d.get('publicera', '').lower()
        self.publ.set('api' if pub == 'api' else 'git' if pub in ('ja', 'yes', 'true', '1', 'git') else 'ingen' if d else 'api')
        self.repo.set(d.get('github_repo', ''))
        self.gren.set(d.get('github_branch', 'main') or 'main')
        tok = d.get('github_token', '')
        self.token.set('' if tok.startswith('KLISTRA_IN') else tok)
        self.stat_url.set(d.get('statistik_url', ''))
        _sn = d.get('statistik_nyckel', '')
        self.stat_key.set('' if _sn.startswith('KLISTRA_IN') else _sn)
        self.p_stad_intervall.set(d.get('stad_intervall', '7') or '7')
        self.p_et_metod.set('usb')   # nätverksutskrift är borttagen ur gränssnittet
        self.p_et_skrivare.set(d.get('etikett_skrivare', ''))
        self.p_et_ip.set(d.get('etikett_ip', ''))
        self.p_et_dpi.set(d.get('etikett_dpi', '203') or '203')
        self.p_et_bredd.set(d.get('etikett_bredd', '70.0') or '70.0')
        self.p_et_hojd.set(d.get('etikett_hojd', '50.8') or '50.8')
        self.p_stad_rader.set(d.get('stad_rader', STANDARD.get("stad_rader","")))
        self.p_stadlogg.set(d.get('stadlogg', STANDARD.get("stadlogg","")))
        self.byggt = True

    def till_dict(self):
        d = {'galler': self.p_galler.get().strip().strip('"'),
             'ut': self.p_ut.get().strip().strip('"'),
             'skannade_mapp': self.p_skannade.get().strip().strip('"') if hasattr(self, 'p_skannade') else r"G:\Sos\PVI_Steril\Förrådet\Scannade_Följesedlar",
             'larande_mapp': (self.p_larande.get().strip().strip('"') if hasattr(self, 'p_larande') else ''),
             'bild_url': (self.p_bild_url.get().strip() if hasattr(self, 'p_bild_url') else ''),
             'bild_nyckel': (self.p_bild_nyckel.get().strip() if hasattr(self, 'p_bild_nyckel') else ''),
             'sparr_procent': self.sparr.get().strip() or '20',
             'publicera': {'api': 'api', 'git': 'ja'}.get(self.publ.get(), ''),
             'github_repo': self.repo.get().strip(), 'github_branch': self.gren.get().strip() or 'main',
             'github_token': self.token.get().strip(),
             'statistik_url': self.stat_url.get().strip(), 'statistik_nyckel': self.stat_key.get().strip(),
             'stad_intervall': (self.p_stad_intervall.get().strip() if hasattr(self, 'p_stad_intervall') else '7'),
             'stad_rader': (self.p_stad_rader.get().strip() if hasattr(self, 'p_stad_rader') else ''),
             'stadlogg': (self.p_stadlogg.get().strip() if hasattr(self, 'p_stadlogg') else '')}
        sokv = self.p_forbruk.get().strip().strip('"')
        d['pdf'] = sokv  # endast PDF
        d.update({'etikett_metod': self.p_et_metod.get(), 'etikett_skrivare': self.p_et_skrivare.get().strip(),
                  'etikett_ip': self.p_et_ip.get().strip(), 'etikett_dpi': self.p_et_dpi.get().strip() or '203',
                  'etikett_bredd': self.p_et_bredd.get().strip() or '70.0', 'etikett_hojd': self.p_et_hojd.get().strip() or '50.8'})
        d['bulk'] = self.p_bulk.get().strip().strip('"')  # bulk alltid med
        if hasattr(self, 'bilder_gh_path_var'):
            d['bilder_github_path'] = self.bilder_gh_path_var.get().strip() or 'bilder'
        return d

    def spara(self, tyst=False):
        try:
            fil = U.skriv_installningar(self.till_dict())
        except Exception as e:   # t.ex. skrivskyddad mapp (OSError) eller att kryptering saknas (ImportError)
            messagebox.showerror("Kunde inte spara", f"Kunde inte skriva inställningarna:\n{e}", parent=self.rot)
            return False
        if not tyst:
            self.status_rad("Inställningarna sparades i " + fil, GRON_MORK)
        return True

    # ---------------------------------------------------------------- bygga fönstret
    def kort(self, foralder, rubrik, farg, under=None):
        yttre = tk.Frame(foralder, bg=RAM)
        inre = tk.Frame(yttre, bg=KORT)
        inre.pack(fill='both', expand=True, padx=1, pady=1)
        tk.Frame(inre, bg=farg, width=6).pack(side='left', fill='y')
        inhall = tk.Frame(inre, bg=KORT)
        inhall.pack(side='left', fill='both', expand=True, padx=14, pady=10)
        rad = tk.Frame(inhall, bg=KORT)
        rad.pack(fill='x')
        tk.Label(rad, text=rubrik, bg=KORT, fg=TEXT, font=self.f_rubrik).pack(side='left')
        if under:
            tk.Label(rad, text=under, bg=KORT, fg=DAMPAD, font=self.f_liten).pack(side='left', padx=(10, 0), pady=(3, 0))
        return yttre, inhall

    def sokvagsrad(self, foralder, var, fil_typer=None, vals_fil=True, vals_mapp=True, spara_fil=False):
        rad = tk.Frame(foralder, bg=KORT)
        ent = ttk.Entry(rad, textvariable=var, font=self.f_norm)
        ent.pack(side='left', fill='x', expand=True, ipady=3)
        if vals_fil:
            Knapp(rad, "Välj fil…", lambda: self.valj(var, 'fil', fil_typer, spara_fil), '#e5e9f0', TEXT, '#d5dbe5').pack(side='left', padx=(6, 0))
        if vals_mapp:
            Knapp(rad, "Välj mapp…", lambda: self.valj(var, 'mapp', None, False), '#e5e9f0', TEXT, '#d5dbe5').pack(side='left', padx=(6, 0))
        return rad

    def valj(self, var, typ, fil_typer, spara):
        start = var.get().strip().strip('"')
        startmapp = start if os.path.isdir(start) else os.path.dirname(start) if start else ''
        if not os.path.isdir(startmapp or ''):
            startmapp = U.basmapp()
        if typ == 'mapp':
            val = filedialog.askdirectory(parent=self.rot, initialdir=startmapp, title="Välj mapp")
        elif spara:
            val = filedialog.asksaveasfilename(parent=self.rot, initialdir=startmapp, title="Var ska data.js sparas?",
                                               initialfile='data.js', defaultextension='.js',
                                               filetypes=[('data.js', '*.js'), ('Alla filer', '*.*')])
        else:
            val = filedialog.askopenfilename(parent=self.rot, initialdir=startmapp, title="Välj fil",
                                             filetypes=fil_typer or [('Alla filer', '*.*')])
        if val:
            var.set(os.path.normpath(val))

    def statusetikett(self, foralder):
        l = tk.Label(foralder, text="", bg=KORT, fg=DAMPAD, font=self.f_liten, anchor='w', justify='left', wraplength=470)
        l.pack(fill='x', pady=(4, 0))
        return l

    def bygg(self):
        # --- rubrikband
        topp = tk.Frame(self.rot, bg=MORK)
        topp.pack(fill='x')
        logo = tk.Canvas(topp, width=54, height=54, bg=MORK, highlightthickness=0)
        logo.pack(side='left', padx=(22, 12), pady=14)
        for i, c in enumerate((GRON, BLA, ROSA)):
            logo.create_rectangle(6 + i * 14, 40 - (i + 1) * 9 - 4, 18 + i * 14, 44, fill=c, outline='')
        tx = tk.Frame(topp, bg=MORK)
        tx.pack(side='left', pady=14)
        tk.Label(tx, text="Lagersök Admin", bg=MORK, fg='#ffffff', font=(FONT_FET, 20, 'bold')).pack(anchor='w')
        tk.Label(tx, text="Databasuppdatering för sterilförrådet", bg=MORK, fg='#9fb0c3', font=self.f_norm).pack(anchor='w')
        tk.Label(topp, text=VERSION, bg=MORK, fg='#6f8196', font=self.f_liten).pack(side='right', padx=22)

        # --- NOTEBOOK MED SNYGGARE FLIKAR ---
        style = ttk.Style()
        try:
            style.theme_use('clam')
        except tk.TclError:
            pass
        # Modern pill-style
        style.configure('TNotebook', background=BAKGRUND, borderwidth=0, tabmargins=(12, 6, 12, 0))
        style.configure('TNotebook.Tab',
                        background='#dde3ed',
                        foreground='#5b6b83',
                        padding=(22, 10),
                        font=(FONT_FET, 10),
                        borderwidth=0,
                        focuscolor=BAKGRUND)
        style.map('TNotebook.Tab',
                  background=[('disabled', '#eceff4'), ('selected', KORT), ('active', '#e6ecf5'), ('!selected', '#dde3ed')],
                  foreground=[('disabled', '#aab3c2'), ('selected', MORK), ('active', TEXT), ('!selected', '#5b6b83')],
                  padding=[('selected', (26, 12))])

        self.notebook = ttk.Notebook(self.rot, style='TNotebook')
        self.notebook.pack(fill='both', expand=True, padx=12, pady=(0, 4))

        # Skapa flikar i ny ordning: Huvud -> Avvikelser -> Städning -> Statistik -> Hjälp -> Om
        flik_huvud = tk.Frame(self.notebook, bg=BAKGRUND)
        flik_inst = tk.Frame(self.notebook, bg=BAKGRUND)
        flik_avv = tk.Frame(self.notebook, bg=BAKGRUND)
        flik_stat = tk.Frame(self.notebook, bg=BAKGRUND)
        flik_folj = tk.Frame(self.notebook, bg=BAKGRUND)
        flik_bilder = tk.Frame(self.notebook, bg=BAKGRUND)
        flik_stad = tk.Frame(self.notebook, bg=BAKGRUND)
        flik_hjalp = tk.Frame(self.notebook, bg=BAKGRUND)
        flik_om = tk.Frame(self.notebook, bg=BAKGRUND)

        self.notebook.add(flik_huvud, text="  HEM  ")
        self.notebook.add(flik_avv, text="  AVVIKELSER  ")
        self.notebook.add(flik_folj, text="  FÖLJESEDLAR  ")
        self.notebook.add(flik_stat, text="  STATISTIK  ")
        self.notebook.add(flik_bilder, text="  BILDER  ")
        self.notebook.add(flik_stad, text="  STÄDNING  ")
        self.notebook.add(flik_inst, text="  INSTÄLLNINGAR  ")
        self.notebook.add(flik_hjalp, text="  HJÄLP  ")
        self.notebook.add(flik_om, text="  OM  ")

        self._flik_installn = flik_inst
        self._flik_avv = flik_avv
        self.bygg_installningar_flik(flik_inst)     # före HEM: uppdatera_status använder fälten där
        self.bygg_huvud_flik(flik_huvud)
        self.bygg_avvikelser_flik(flik_avv)
        self.bygg_foljesedel_flik(flik_folj)
        self.bygg_statistik_flik(flik_stat)
        self.bygg_bilder_flik(flik_bilder)
        self.bygg_stadning_flik(flik_stad)
        # STÄDNING är under utveckling: fliken är grå och går inte att öppna. Ta bort raden nedan för att slå på den igen.
        try:
            self.notebook.tab(flik_stad, state='disabled')
        except Exception as e:
            logga(f"Kunde inte stänga av STÄDNING-fliken: {e!r}")
        self.bygg_hjalp_flik(flik_hjalp)
        self.bygg_om_flik(flik_om)

        # --- Auto-hämta statistik när STATISTIK-fliken öppnas ---
        self._stat_auto_hamtad = False
        def _on_tab_changed(e=None):
            try:
                idx = 3 if str(self.notebook.select()) == str(flik_stat) else -1
                # STATISTIK är index 3 (HEM=0, AVVIKELSER=1, STÄDNING=2, STATISTIK=3)
                if idx == 3:
                    if not self._stat_auto_hamtad:
                        self._stat_auto_hamtad = True
                        self.rot.after(400, self.hamta_statistik)
                    else:
                        # uppdatera diskret varje gång man går tillbaka, om det gått >5 min
                        import time
                        now = time.time()
                        last = getattr(self, '_stat_senast', 0)
                        if now - last > 300:
                            self.rot.after(300, self.hamta_statistik)
            except Exception:
                pass
        self.notebook.bind('<<NotebookTabChanged>>', _on_tab_changed)
        # Spara tidpunkt vid lyckad hämtning
        _orig_h_stat = self.h_stat
        def _h_stat_wrapper(data, fel):
            import time
            if not fel and data:
                self._stat_senast = time.time()
            return _orig_h_stat(data, fel)
        self.h_stat = _h_stat_wrapper


        # --- fot (utanför flikar)
        fot = tk.Frame(self.rot, bg=BAKGRUND)
        fot.pack(fill='x', padx=20, pady=(2, 10))
        self.fot_etikett = tk.Label(fot, text="", bg=BAKGRUND, fg=DAMPAD, font=self.f_liten, anchor='w')
        self.fot_etikett.pack(side='left')

    # ===================== INSTÄLLNINGAR (alla sökvägar och publicering) =====================
    def bygg_installningar_flik(self, parent):
        scroll = ScrollableFrame(parent, bg=BAKGRUND)
        scroll.pack(fill='both', expand=True)
        kropp = tk.Frame(scroll.scrollable_frame, bg=BAKGRUND)
        kropp.pack(fill='both', expand=True, padx=18, pady=14)
        tk.Label(kropp, text="Här ställer du in var filerna och mapparna finns. Resultatet av kontrollen visas på fliken HEM.", bg=BAKGRUND, fg=DAMPAD,
                 font=self.f_norm, anchor='w', justify='left', wraplength=900).pack(fill='x', pady=(0, 10))

        # Fabrikssterila - endast PDF
        k, h = self.kort(kropp, "Fabrikssterila artiklar", GRON, "(PDF)")
        k.pack(fill='x', pady=(0, 10))
        self.sokvagsrad(h, self.p_forbruk, [('PDF-fil', '*.pdf'), ('Alla filer', '*.*')], vals_mapp=False).pack(fill='x')

        # Galler
        k, h = self.kort(kropp, "Instrumentgaller", BLA, "mapp med Excel-filer")
        k.pack(fill='x', pady=(0, 10))
        self.sokvagsrad(h, self.p_galler, vals_fil=False).pack(fill='x', pady=(6, 0))

        # Bulk - ALLTID MED (obligatorisk)
        k, h = self.kort(kropp, "Bulkplatser", ROSA, "(Excel-fil)")
        k.pack(fill='x', pady=(0, 10))
        self.bulk_rad = self.sokvagsrad(h, self.p_bulk, [('Excel-fil', '*.xlsx *.xlsm'), ('Alla filer', '*.*')], vals_mapp=False)
        self.bulk_rad.pack(fill='x')

        # Skannade följesedlar
        k, h = self.kort(kropp, "Skannade följesedlar", BLA, "mapp för PDF:er")
        k.pack(fill='x', pady=(0, 10))
        tk.Label(h, text="Mapp där skannade följesedlar sparas direkt som SOxxxxxxxx.pdf (inga undermappar)", bg=KORT, fg=DAMPAD, font=self.f_liten, wraplength=460, justify='left').pack(anchor='w', pady=(0,4))
        self.sokvagsrad(h, self.p_skannade, vals_fil=False).pack(fill='x')

        # Inlärning av följesedlar
        k, h = self.kort(kropp, "Inlärning av följesedlar", GRON, "mapp för inlärd data")
        k.pack(fill='x', pady=(0, 10))
        tk.Label(h, text="Mapp där programmet sparar texten från följesedlar (andra leverantörer än MediCarrier), dina rättelser och leverantörsprofilerna (leverantorsprofiler.json). Ändrar du mappen flyttas inte det som redan är sparat, så flytta mappen först och peka sedan om sökvägen här.", bg=KORT, fg=DAMPAD, font=self.f_liten, wraplength=460, justify='left').pack(anchor='w', pady=(0,4))
        self.sokvagsrad(h, self.p_larande, vals_fil=False).pack(fill='x')

        # Publicering
        k, h = self.kort(kropp, "Publicering och inställningar", '#6b7280')
        k.pack(fill='x', pady=(0, 4))
        val = tk.Frame(h, bg=KORT)
        val.pack(fill='x', pady=(4, 4))
        for txt, v in (("Ladda upp till GitHub", 'api'), ("Bara spara data.js", 'ingen')):
            tk.Radiobutton(val, text=txt, variable=self.publ, value=v, bg=KORT, activebackground=KORT, fg=TEXT,
                           font=self.f_norm, selectcolor=KORT, highlightthickness=0, bd=0,
                           command=self.uppdatera_status).pack(side='left', padx=(0, 14))
        # GitHub-fälten (Arkiv/Gren/Åtkomstnyckel) är borttagna från vyn.
        # Värdena ligger kvar i installningar.dat och används fortfarande vid uppladdning.
        self.gh = tk.Frame(h, bg=KORT)
        self.gh.pack(fill='x')
        self.gh.columnconfigure(1, weight=1)
        # Testa-knappen längst till vänster så att resultattexten får full bredd.
        tst = tk.Frame(self.gh, bg=KORT)
        tst.pack(fill='x', pady=(6, 0))
        Knapp(tst, "Testa anslutningen", self.testa, '#e5e9f0', TEXT, '#d5dbe5').pack(side='left')
        self.test_res = tk.Label(tst, text="", bg=KORT, fg=DAMPAD, font=self.f_liten, wraplength=520, justify='left')
        self.test_res.pack(side='left', padx=10, fill='x', expand=True)

        tk.Frame(h, bg=RAM, height=1).pack(fill='x', pady=8)
        g2 = tk.Frame(h, bg=KORT)
        g2.pack(fill='x')
        g2.columnconfigure(1, weight=1)
        tk.Label(g2, text="Spara data.js här", bg=KORT, fg=DAMPAD, font=self.f_liten).grid(row=0, column=0, sticky='w', padx=(0, 10))
        self.sokvagsrad(g2, self.p_ut, vals_mapp=False, spara_fil=True).grid(row=0, column=1, sticky='ew')
        tk.Label(g2, text="Spärr (procent)", bg=KORT, fg=DAMPAD, font=self.f_liten).grid(row=1, column=0, sticky='w', pady=(6, 0), padx=(0, 10))
        sp = tk.Frame(g2, bg=KORT)
        sp.grid(row=1, column=1, sticky='w', pady=(6, 0))
        ttk.Spinbox(sp, from_=0, to=100, increment=5, textvariable=self.sparr, width=5, font=self.f_norm).pack(side='left', ipady=2)
        tk.Label(sp, text="0 = ingen spärr", bg=KORT, fg=DAMPAD, font=self.f_liten).pack(side='left', padx=8)
        tk.Label(g2, text="Stoppar uppdateringen om antalet artiklar sjunker mer än så här jämfört med förra gången.",
                 bg=KORT, fg=DAMPAD, font=self.f_liten, wraplength=440, justify='left', anchor='w').grid(row=2, column=0, columnspan=2, sticky='w', pady=(4, 0))

        Knapp(kropp, "Spara inställningar", self.spara, '#6b7280', hover='#5a5a5a', stor=True).pack(anchor='w', pady=(10, 4))

    # ===================== HUVUDFLIK =====================
    def bygg_huvud_flik(self, parent):
        # Ingen rullningslist på HEM: loggen längst ner (hela bredden) har sin egen och tar resten av höjden.
        kropp = tk.Frame(parent, bg=BAKGRUND)
        kropp.pack(fill='both', expand=True, padx=18, pady=14)
        kropp.columnconfigure(0, weight=11, uniform='k')
        kropp.columnconfigure(1, weight=10, uniform='k')
        kropp.rowconfigure(0, weight=0)
        kropp.rowconfigure(1, weight=1)
        vanster = tk.Frame(kropp, bg=BAKGRUND)
        vanster.grid(row=0, column=0, sticky='nsew', padx=(0, 9))
        hoger = tk.Frame(kropp, bg=BAKGRUND)
        hoger.grid(row=0, column=1, sticky='nsew', padx=(9, 0))

        # Källor: bara statusfält. Sökvägarna ändras under fliken INSTÄLLNINGAR.
        k, h = self.kort(vanster, "Källor och status", GRON, "sökvägar ändras under INSTÄLLNINGAR")
        k.pack(fill='x', pady=(0, 10))
        for txt, attr in (("Fabrikssterila artiklar (PDF)", 'st_forbruk'), ("Instrumentgaller (mapp med Excel-filer)", 'st_galler'),
                          ("Bulkplatser (Excel-fil)", 'st_bulk')):
            tk.Label(h, text=txt, bg=KORT, fg=TEXT, font=self.f_fet, anchor='w').pack(fill='x', pady=(6, 0))
            setattr(self, attr, self.statusetikett(h))

        # --- höger: körning
        knappar = tk.Frame(hoger, bg=BAKGRUND)
        knappar.pack(fill='x', pady=(4,2))
        for i in range(3):
            knappar.columnconfigure(i, weight=1, uniform='knappar')

        self.b_kor = Knapp(knappar, "Uppdatera databasen", lambda: self.starta(False), GRON_MORK, hover='#4f7a24', stor=True)
        self.b_kor.grid(row=0, column=0, sticky='ew', padx=(0,6), pady=4)

        self.b_torr = Knapp(knappar, "Torrkörning", lambda: self.starta(True), BLA, hover='#3a7396', stor=True)
        self.b_torr.grid(row=0, column=1, sticky='ew', padx=3, pady=4)

        self.b_spara = Knapp(knappar, "Spara inställningar", self.spara, '#6b7280', hover='#5a5a5a', stor=True)
        self.b_spara.grid(row=0, column=2, sticky='ew', padx=(6,0), pady=4)

        self.b_avbryt = Knapp(knappar, "Avbryt", self.avbryt_kor, ROD, hover='#a02f23', stor=True)
        tk.Label(hoger, text="Torrkörning läser alla filer och visar resultatet, men skriver och skickar ingenting.",
                 bg=BAKGRUND, fg=DAMPAD, font=self.f_liten, anchor='w').pack(fill='x', pady=(6, 8))

        stilen = ttk.Style()
        stilen.configure('Lager.Horizontal.TProgressbar', troughcolor=RAM, background=GRON, bordercolor=RAM,
                         lightcolor=GRON, darkcolor=GRON, thickness=12)
        stilen.configure('TEntry', fieldbackground='#ffffff', bordercolor='#c5ccd6', lightcolor='#c5ccd6', darkcolor='#c5ccd6', padding=3)
        stilen.map('TEntry', bordercolor=[('focus', BLA)], lightcolor=[('focus', BLA)], darkcolor=[('focus', BLA)])
        self.steg_etikett = tk.Label(hoger, text="Redo.", bg=BAKGRUND, fg=TEXT, font=self.f_fet, anchor='w')
        self.steg_etikett.pack(fill='x')
        self.prog = ttk.Progressbar(hoger, style='Lager.Horizontal.TProgressbar', mode='determinate', maximum=1000)
        self.prog.pack(fill='x', pady=(4, 10))
        self._res_anker = tk.Frame(hoger, bg=BAKGRUND, height=1)      # resultatrutorna packas före den här
        self._res_anker.pack(fill='x')

        # resultatkort
        self.res = tk.Frame(hoger, bg=BAKGRUND)
        self.res_tiles = []
        for i, (txt, farg) in enumerate((("Fabrikssterila", GRON), ("Instrumentgaller", BLA), ("Med bulkplats", ROSA))):
            self.res.columnconfigure(i, weight=1, uniform='t')
            ram = tk.Frame(self.res, bg=RAM)
            ram.grid(row=0, column=i, sticky='nsew', padx=(0 if i == 0 else 6, 0))
            inre = tk.Frame(ram, bg=KORT)
            inre.pack(fill='both', expand=True, padx=1, pady=1)
            tk.Frame(inre, bg=farg, height=4).pack(fill='x')
            n = tk.Label(inre, text="0", bg=KORT, fg=farg, font=(FONT_FET, 24, 'bold'))
            n.pack(pady=(6, 0))
            tk.Label(inre, text=txt, bg=KORT, fg=TEXT, font=self.f_liten).pack()
            d = tk.Label(inre, text="", bg=KORT, fg=DAMPAD, font=self.f_liten)
            d.pack(pady=(0, 6))
            self.res_tiles.append((n, d))
        self.res_info = tk.Frame(hoger, bg=BAKGRUND)
        self.avv_etikett = tk.Label(self.res_info, text="", bg=BAKGRUND, fg=TEXT, font=self.f_norm)
        self.avv_etikett.pack(side='left')
        self.avv_lank = tk.Label(self.res_info, text="Visa avvikelser", bg=BAKGRUND, fg=BLA, font=(FONT, 10, 'underline'), cursor='hand2')
        self.app_lank = tk.Label(self.res_info, text="", bg=BAKGRUND, fg=BLA, font=(FONT, 10, 'underline'), cursor='hand2')

        # logg
        lram = tk.Frame(kropp, bg=RAM)
        lram.grid(row=1, column=0, columnspan=2, sticky='nsew', pady=(12, 0))
        self.logg = tk.Text(lram, bg=LOGG_BG, fg=LOGG_TEXT, font=self.f_mono, relief='flat', wrap='word', height=8,
                            padx=10, pady=8, state='disabled', insertbackground=LOGG_TEXT, bd=0, highlightthickness=0)
        rull = ttk.Scrollbar(lram, command=self.logg.yview)
        self.logg.configure(yscrollcommand=rull.set)
        rull.pack(side='right', fill='y')
        self.logg.pack(side='left', fill='both', expand=True, padx=1, pady=1)
        for tagg, c in (('steg', '#67d4e8'), ('ok', '#8fd16a'), ('varning', '#f0c25a'), ('fel', '#ff7b72'), ('info', '#8b97a6')):
            self.logg.tag_configure(tagg, foreground=c)
        self.logg.tag_configure('steg', font=(self.f_mono[0], 9, 'bold'))
        self.skriv_logg("Kontrollera källorna till vänster (ändras under INSTÄLLNINGAR) och tryck på Uppdatera databasen.\n", 'info')

    # ===================== HJÄLPFLIK =====================
    def bygg_hjalp_flik(self, parent):
        wrapper = ScrollableFrame(parent, bg=BAKGRUND)
        wrapper.pack(fill='both', expand=True)
        content = wrapper.scrollable_frame
        canvas = wrapper.canvas  # för kompatibilitet, används inte direkt

        header = tk.Frame(content, bg=MORK)
        header.pack(fill='x', padx=18, pady=(18, 16))
        tk.Label(header, text="Lagersök Admin", bg=MORK, fg='#ffffff', font=(FONT_FET, 22, 'bold')).pack(anchor='w', padx=18, pady=(14,0))
        tk.Label(header, text=f"Databasuppdatering  {VERSION} Så här funkar det", bg=MORK, fg='#9fb0c3', font=self.f_norm).pack(anchor='w', padx=18, pady=(2,14))

        def card(parent, accent=GRON):
            outer = tk.Frame(parent, bg=RAM)
            outer.pack(fill='x', padx=18, pady=(0,12))
            inner = tk.Frame(outer, bg=KORT)
            inner.pack(fill='both', expand=True, padx=1, pady=1)
            tk.Frame(inner, bg=accent, height=4).pack(fill='x')
            body = tk.Frame(inner, bg=KORT)
            body.pack(fill='both', expand=True, padx=16, pady=14)
            return body

        def title_row(body, icon, titel):
            r = tk.Frame(body, bg=KORT)
            r.pack(fill='x', pady=(0,6))
            tk.Label(r, text=icon, bg=KORT, fg=TEXT, font=(FONT_FET, 14)).pack(side='left')
            tk.Label(r, text=titel, bg=KORT, fg=TEXT, font=self.f_rubrik_stor).pack(side='left', padx=(8,0))

        def bullet(body, text):
            r = tk.Frame(body, bg=KORT)
            r.pack(fill='x', pady=2)
            tk.Label(r, text="•", bg=KORT, fg=DAMPAD, font=self.f_norm).pack(side='left', anchor='n')
            tk.Label(r, text=text, bg=KORT, fg=TEXT, font=self.f_norm, justify='left', wraplength=760, anchor='w').pack(side='left', padx=(8,0), fill='x', expand=True)

        def small(body, text):
            tk.Label(body, text=text, bg=KORT, fg=DAMPAD, font=self.f_liten, justify='left', wraplength=760, anchor='w').pack(fill='x', pady=(4,0))

        b = card(content, GRON_MORK)
        title_row(b, "📦", "Vad gör programmet?")
        bullet(b, "Läser 3 obligatoriska källor och bygger ihop dem till en databas:")
        bullet(b, "  1. Fabrikssterilt – PDF-etiketter med artikelnummer och lagerplats")
        bullet(b, "  2. Instrumentgaller – Excel-filer med gallerinnehåll")
        bullet(b, "  3. Bulkplatser – Excel-fil som kopplar artikelnummer till bulkplats")
        bullet(b, "Skapar filen data.js som Lagersök-webbappen (index.html) använder.")
        bullet(b, "Kan ladda upp data.js direkt till GitHub så webben uppdateras automatiskt för alla användare.")
        small(b, "Smart filval: Väljer du en mapp används automatiskt den nyaste filen i mappen – du behöver aldrig rensa gamla filer.")

        b = card(content, BLA)
        title_row(b, "🚀", "Snabbstart i 3 steg")
        steps = tk.Frame(b, bg=KORT)
        steps.pack(fill='x', pady=(6,0))
        for nr, (t, desc) in enumerate([
            ("Välj filer", "1. Fabrikssterilt: Välj mapp eller PDF med etiketter\n2. Instrumentgaller: Välj mappen med alla galler-Excel\n3. Bulkplatser: Välj Excel-filen med bulkplatser"),
            ("Torrkörning", "Klicka 'Torrkörning' först. Inget sparas. Du ser antal artiklar och avvikelser direkt i fliken Avvikelser, och en jämförelse med förra gången."),
            ("Uppdatera", "Om allt ser bra ut: Klicka 'Uppdatera databasen'. Då skrivs data.js och laddas upp (om GitHub är valt)."),
        ], 1):
            sf = tk.Frame(steps, bg='#f1f5f9')
            sf.pack(fill='x', pady=4)
            tk.Label(sf, text=str(nr), bg=BLA, fg='white', font=self.f_fet, width=3, pady=4).pack(side='left', padx=6, pady=6)
            inner_sf = tk.Frame(sf, bg='#f1f5f9')
            inner_sf.pack(side='left', fill='x', expand=True, padx=8, pady=6)
            tk.Label(inner_sf, text=t, bg='#f1f5f9', fg=TEXT, font=self.f_fet, anchor='w').pack(fill='x')
            tk.Label(inner_sf, text=desc, bg='#f1f5f9', fg=TEXT, font=self.f_norm, justify='left', anchor='w').pack(fill='x', pady=(2,0))

        b = card(content, GRON)
        title_row(b, "📄", "Fabrikssterilt – PDF-etiketter (obligatorisk)")
        bullet(b, "Vad: Varje etikett i PDF:en innehåller Artnr, Lev.artnr, Produktnamn och Plats (t.ex. A-01-02-03).")
        bullet(b, "Hur: Välj en mapp med PDF-filer eller en enskild PDF. Programmet väljer automatiskt den senast ändrade filen i mappen.")
        bullet(b, "Krav: Plats måste ha formatet BOKSTAV-XX-XX-XX. Ogiltig plats = avvikelse och raden hoppas över.")
        bullet(b, "Avvikelser: Dubbletter, etiketter utan artnr/namn eller ogiltig plats hamnar i fliken Avvikelser > Fabrikssterilt och i avvikelser.txt.")
        small(b, "Format: Endast PDF stöds. Tre etiketter per rad i PDF:en läses automatiskt.")

        b = card(content, BLA)
        title_row(b, "🧰", "Instrumentgaller (obligatorisk)")
        bullet(b, "Vad: En mapp med Excel-filer (.xlsx/.xlsm), en fil per galler. Filnamnet blir gallernamnet i sökappen.")
        bullet(b, "Struktur: Varje flik i Excel-filen ska ha rubrik 'Rad X Reol Y'. Under rubriken listas artikelnummer. Tomma galler flaggas som avvikelse.")
        bullet(b, "Sökning: Programmet letar rekursivt även i undermappar – du kan organisera galler i undermappar.")
        bullet(b, "Kvalitetskontroll: Samma artikelnummer med olika benämning i olika galler, eller ogiltiga artikelnummer, hamnar i Avvikelser > Instrumentgaller.")
        small(b, "Tips: Lägg nya exporter i samma mapp – programmet hittar nyaste automatiskt via ändringsdatum.")

        b = card(content, ROSA)
        title_row(b, "📦", "Bulkplatser (obligatorisk)")
        bullet(b, "Vad: En Excel-fil med tre kolumner: Artikelnummer, Plockplats, Bulkplats. Kopplar en fabrikssteril artikel till dess bulkplats.")
        bullet(b, "Hur: Välj Excel-filen. Alla tre kolumnerna måste finnas. Plockplats och Bulkplats valideras mot formatet XX-XX-XX eller liknande.")
        bullet(b, "Varför obligatorisk: Sökappen visar bulkplats direkt när du söker en artikel – därför måste den alltid vara med, precis som Fabrikssterilt och Instrumentgaller.")
        bullet(b, "Avvikelser: Artikelnummer i bulkfilen som inte finns i fabrikssterilt, dubbla bulkplatser för samma artikel, eller ogiltiga platsformat hamnar i Avvikelser > Bulk.")
        small(b, "Exempel: Artikel 12345 finns på A-01-02-03 men bulk finns på BULK-1 – då visas båda platserna i sökappen.")

        b = card(content, "#6b7280")
        title_row(b, "☁️", "Publicering och filer")
        bullet(b, "Var sparas data.js: Välj sökväg i fältet 'Spara data.js här'. Filen sparas alltid lokalt först. Bredvid skapas även data.js.bak (föregående version) och avvikelser.txt.")
        bullet(b, "Ladda upp till GitHub (API) – Rekommenderat: Använder arkiv, gren och åtkomstnyckel som redan är sparade i installningar.dat. Klicka 'Testa anslutningen' för att verifiera att anslutningen fungerar.")
        bullet(b, "Bara spara data.js: Välj detta om du vill ladda upp filen manuellt eller testa lokalt. Ingen GitHub-uppladdning sker.")
        small(b, "Säkerhet: Inställningarna, inklusive nyckeln, sparas krypterade i installningar.dat och skickas bara till GitHub. Använd en nyckel som är begränsad till just det här arkivet med enbart skrivrätt till innehåll.")

        b = card(content, "#6b7280")
        title_row(b, "⚙️", "Fliken INSTÄLLNINGAR")
        bullet(b, "Alla sökvägar och uppgifter ligger samlade här: källfiler, var data.js sparas, GitHub, skannade följesedlar och inlärningsmappen. Flikarna HEM, AVVIKELSER med flera visar bara knappar och status.")
        bullet(b, "Allt sparas krypterat i installningar.dat. Tryck 'Spara' efter ändringar. Bidragsbilder (adress och nyckel till bildtjänsten) och bildmappen på GitHub ('bilder') är inbyggda och visas inte.")
        small(b, "Fliken STÄDNING är under utveckling och är avstängd tills vidare.")

        b = card(content, GRON)
        title_row(b, "🧾", "Följesedlar och etiketter")
        bullet(b, "Lägg till PDF:er… läser följesedlar. MediCarrier-PDF:er läses direkt. Övriga leverantörer (PDF eller bild) läses med textlagret eller OCR, och programmet lär sig med tiden hur varje leverantörs följesedel ser ut.")
        bullet(b, "Skanna följesedel…: lägg ett A4 i skannern och tryck. Ingen dialogruta visas. Därefter får du frågan om du vill skanna fler sidor. Sidorna sparas som en PDF i mappen för skannade följesedlar.")
        bullet(b, "Listan: klicka på en rad för att välja bort eller in den, eller på orderraden för hela ordern. Klicka på en kolumnrubrik för att sortera inom varje order. Etiketterna skrivs alltid ut i den ordning raderna visas.")
        bullet(b, "Rensa (till höger ovanför kolumnen 'Ta bort') tömmer listan. Du kan också lägga till en etikett manuellt med artikelnummer.")
        bullet(b, "Inlärning (leverantörer utom MediCarrier): programmet lär sig av hur du markerar raderna. 'Spara som träning' sparar följesedlarna i inlärningsmappen utan att skriva ut. 'Inlärning…' visar vad programmet har lärt sig. Ett mönster används först när det setts på minst 2 följesedlar och 3 rader.")
        bullet(b, "Etikettskrivare: välj Windows-skrivaren (Zebra ZD421), upplösning och etikettstorlek (standard 70,0 x 50,8 mm). 'Testetikett' skriver en provetikett och 'Spara som ZPL-fil' sparar utskriften som fil.")

        b = card(content, ROSA)
        title_row(b, "🖼️", "Fliken BILDER")
        bullet(b, "Bilder konverteras till webp (800 px, kvalitet 80) plus en miniatyr och laddas upp till GitHub. Tryck 'Ladda artiklar + kolla bilder' först så vet programmet vilka artiklar som finns och vilka som saknar bild.")
        bullet(b, "Välj bild för markerad: en bild till en artikel. Välj flera bilder…: många på en gång. Heter filen som artikelnumret går den rakt igenom, annars får du en ruta med förhandsvisning där du anger artikelnummer.")
        bullet(b, "Spärr: en bild kan bara laddas upp till en artikel som finns i databasen. Okända artikelnummer avvisas.")
        bullet(b, "Inkorg…: bilder som personalen skickat från sökappen. Markera en bild för att se artikelnummer, produktnamn, vem som skickat in och om det är en ny bild eller en ersättning (då visas även den nuvarande bilden bredvid). 'Godkänn' konverterar och publicerar, 'Avvisa' raderar ur inkorgen.")
        bullet(b, "Koder…: skapa en personlig kod per bidragsgivare, stäng av eller slå på en kod, eller ta bort den. Koden skriver personen in i sökappen första gången. Det finns ingen gräns för hur många bilder som kan skickas.")
        bullet(b, "Visa online, Orphans (bilder utan artikel) och 'Ta bort bild online' hanterar det som redan ligger på GitHub. Loggen längst ner visar vad som hänt.")

        b = card(content, BLA)
        title_row(b, "📊", "Statistik – vad söker personalen efter?")
        bullet(b, "Sökappen kan skicka varje sökning (söktext, tid och om det gav träff) till en egen liten tjänst. Ingen information om vem som sökte sparas.")
        bullet(b, "I fliken STATISTIK trycker du på 'Uppdatera statistik' för att hämta. Du kan välja period (7, 14, 30, 90 eller 365 dagar) i rullgardinen. Adress och nyckel till statistik-tjänsten är sparade i installningar.dat – det finns inga fält för dem i fönstret.")
        bullet(b, "Överst: Totalt sökningar, Snitt / dag, Träffsäkerhet och Utan träff. Röd färg används bara där något behöver åtgärdas, resten är blått.")
        bullet(b, "Utan träff och Mest sökta ligger direkt under toppkorten. Utan träff är det du kan göra något åt: rätta stavning, lägg till synonymer eller lägg till artikeln i databasen. Staplarna visar de 10 vanligaste sökningarna med antal och procent.")
        bullet(b, "Diagram: Sökningar per dag, Fördelning över dygnet och Veckodag.")
        bullet(b, "Nyttig insikt – korten: Artiklar vs Instrumentgaller, Utanför arbetstid 21-07, Trend mot föregående period, Återkommande sökningar, Kontinuitet och Aktiva dagar.")
        bullet(b, "Enheter: Enhetstyp, Operativsystem och Installerad som app, med en stapel, antal och procent per rad.")
        bullet(b, "Kameran: ett flöde från Kameran öppnades, via hur många av dem som ledde till en sökning, till antal sökningar via kameran och deras andel av alla sökningar. Under finns en rad om när kameran används mest. Längst ner finns Så söker man: andel artikelnummer mot ord, och hur långa sökningarna är.")
        small(b, "Tips: Kolla Utan träff regelbundet – lägg till synonymer, rätta stavning eller lägg till artikeln i databasen för de vanligaste orden där. En hög andel sökningar utanför arbetstid kan tyda på behov av bättre ljus eller skyltning i förrådet.")

        b = card(content, GUL)
        title_row(b, "🛡️", "Spärr – skydd mot att råka välja fel fil")
        bullet(b, "Vad: Jämför antalet artiklar i nya data.js med förra data.js. Om antalet minskat med mer än X % stoppas uppdateringen.")
        bullet(b, "Varför: Skyddar mot att du råkat välja en tom eller gammal mapp, eller att PDF-läsningen misslyckats och nästan inga artiklar hittats.")
        bullet(b, "Inställning: Standard 20 %. Sätt 0 för att stänga av spärren helt (då stoppas aldrig uppdateringen). Ändra i fältet 'Spärr (procent)'.")
        bullet(b, "När den triggas: Du får en varningsruta med före/efter antal och diff. Välj Avbryt om du valt fel mapp, eller Fortsätt ändå om minskningen är korrekt (t.ex. storstädning i förrådet).")
        small(b, "Spärren kollar bara första gången efter att data.js redan finns. Finns ingen tidigare data.js körs ingen spärrkontroll.")

        b = card(content, BLA)
        title_row(b, "⚠️", "Avvikelser-fliken – din kvalitetskontroll (läs alltid denna)")
        bullet(b, "Var: Efter varje körning (även torrkörning) uppdateras fliken AVVIKELSER uppe i fönstret med totalt antal i titeln.")
        bullet(b, "Uppdelat i 3 kategorier: Fabrikssterilt (PDF-problem), Instrumentgaller (Excel-struktur, namnkonflikter), Bulk (bulk utan match, dubbletter). Etiketter som saknar hyllplats helt (t.ex. PASSIV eller en kod som 02999999) visas för sig under kolumnerna, dolda tills du klickar, och räknas inte i flikens siffra.")
        bullet(b, "Sök och filtrera: Använd sökfältet överst i Avvikelser-fliken för att filtrera på text, t.ex. ett artikelnummer.")
        bullet(b, "Kopiera: Knapparna 'Kopiera filtrerade' och 'Kopiera alla' kopierar till urklipp så du kan klistra in i mail/Excel.")
        bullet(b, "Fil: Samma avvikelser sparas även i avvikelser.txt bredvid data.js – bra att skicka till kollegor eller spara som dokumentation.")
        small(b, "Arbetsflöde: 1) Torrkör 2) Kolla Avvikelser 3) Åtgärda ev. fel i Excel/PDF 4) Torrkör igen 5) Uppdatera på riktigt när Avvikelser ser OK ut.")

        b = card(content, DAMPAD)
        title_row(b, "💡", "Tips & felsökning")
        bullet(b, "Organisera: Lägg alltid nya exporter i samma mappar – programmet hittar nyaste automatiskt via ändringsdatum. Du behöver aldrig radera gammalt.")
        bullet(b, "Inställningar: Sparas automatiskt när du stänger programmet (och när du trycker på Spara inställningar eller kör). De ligger krypterade i installningar.dat bredvid exe-filen och kan inte redigeras med anteckningar. Filen kan kopieras tillsammans med programmet till en annan dator.")
        bullet(b, "Loggen: I fliken HEM, längst ner över hela bredden, ser du steg-för-steg vad som händer. Röd text = fel som stoppar, gul = varning/avvikelse, grön = klart.")
        bullet(b, "Vanliga fel: 'Hittar ingen PDF' → Kolla att mappen innehåller .pdf. 'Hittade inga Excel-filer' → Kolla att mappen innehåller .xlsx/.xlsm. 'Bulkplatser' röd → Kolla att Excel-filen har kolumnerna Artikelnummer, Plockplats, Bulkplats.")
        bullet(b, "Efter uppdatering: Öppna data.js i webbläsaren eller testa sökappen (index.html) lokalt för att se att allt ser rätt ut innan du litar på GitHub-uppladdningen.")

        # Scroll hanteras nu av ScrollableFrame automatiskt



    def bygg_om_flik(self, parent):
        """Om-flik - SoS"""
        scroll = ScrollableFrame(parent, bg=BAKGRUND)
        scroll.pack(fill='both', expand=True)
        wrapper = tk.Frame(scroll.scrollable_frame, bg=BAKGRUND)
        wrapper.pack(fill='both', expand=True, padx=18, pady=18)

        kort_yttre = tk.Frame(wrapper, bg=RAM)
        kort_yttre.pack(fill='x', pady=(0,12))
        kort_inre = tk.Frame(kort_yttre, bg=KORT)
        kort_inre.pack(fill='both', expand=True, padx=1, pady=1)
        tk.Frame(kort_inre, bg=MORK, height=5).pack(fill='x')
        body = tk.Frame(kort_inre, bg=KORT)
        body.pack(fill='both', expand=True, padx=26, pady=24)

        tk.Label(body, text="Lagersök Admin", bg=KORT, fg=TEXT, font=(FONT_FET, 28, 'bold')).pack(anchor='w')
        tk.Label(body, text=f"Databasuppdatering {VERSION}", bg=KORT, fg=BLA, font=self.f_rubrik_stor).pack(anchor='w', pady=(4,0))

        tk.Frame(body, bg=RAM, height=1).pack(fill='x', pady=16)

        tk.Label(body, text="Ett komplett verktyg för Sterilförrådet på Södersjukhuset för att:", bg=KORT, fg=TEXT, font=self.f_fet, wraplength=650, justify='left').pack(anchor='w', pady=(0,12))

        def punkt(t):
            r = tk.Frame(body, bg=KORT)
            r.pack(fill='x', pady=5, padx=(8,0))
            tk.Label(r, text="•", bg=KORT, fg=BLA, font=(FONT_FET, 14, 'bold')).pack(side='left', anchor='n', padx=(0,2))
            tk.Label(r, text=t, bg=KORT, fg=TEXT, font=self.f_norm, wraplength=600, justify='left', anchor='w').pack(side='left', padx=(8,0), fill='x', expand=True)

        punkt("Hålla Lagersök-webbappen uppdaterad.")
        punkt("Skriva ut platsetiketter från följesedlar – även bulkplatser, med RAD, SO och datum.")
        punkt("Hantera artikelbilder – 800px webp + thumb som laddas upp till GitHub.")
        punkt("Få ut statistik kring användandet av webbappen – sökningar, träffar, enheter, kamera.")

        tk.Frame(body, bg=RAM, height=1).pack(fill='x', pady=18)

        tk.Label(body, text="Utvecklat av Pelle Holmström tillsammans med personalen i Sterilförrådet 2026.", bg=KORT, fg=TEXT, font=(FONT_FET, 12, 'bold'), wraplength=650, justify='left').pack(anchor='w', pady=(0,8))
        tk.Label(body, text="För Sterilförrådet, Södersjukhuset. Detaljerad manual finns i fliken HJÄLP.", bg=KORT, fg=DAMPAD, font=self.f_liten, wraplength=650, justify='left').pack(anchor='w')

        tk.Frame(body, bg=RAM, height=1).pack(fill='x', pady=16)
        tk.Label(body, text="Krav: Python v3.11.x för full OCR (RapidOCR + Windows OCR + Tesseract). Python 3.15 saknar wheels för RapidOCR/winrt.", bg=KORT, fg=ROD, font=self.f_liten, wraplength=650, justify='left').pack(anchor='w', pady=(0,8))
        tk.Label(body, text=f"Version {VERSION}", bg=KORT, fg=DAMPAD, font=self.f_liten).pack(anchor='w')


    # ===================== STATISTIKFLIK - NYTTIG VERSION =====================
    # ------------------------------------------------------------------
    # FÖLJESEDLAR: läs MediCarrier-PDF:er, jämför med databasen, lista artiklar som ska ha platsetikett
    # ------------------------------------------------------------------
    def bygg_foljesedel_flik(self, parent):
        self.fs_sedlar = []      # inlästa följesedlar
        self.fs_kallor = {}      # sökväg -> ('medi', följesedel) eller ('text', text från PDF/OCR)
        self.fs_ordning = []
        self.fs_nya, self.fs_fel, self.fs_klara = {}, [], 0
        self.fs_poster = []      # en post per orderrad
        self.fs_valda = {}       # radnyckel -> True/False (kryssruta)
        self.fs_sort = None
        self.fs_markerad = {}    # radnyckel -> 'pos' | 'neg' (dina rättelser, används som facit för inlärningen)
        topp = tk.Frame(parent, bg=BAKGRUND)
        topp.pack(fill='x', padx=18, pady=(12, 4))
        tk.Label(topp, text="Följesedlar och platsetiketter", bg=BAKGRUND, fg=TEXT, font=self.f_rubrik_stor).pack(side='left')
        knappar = tk.Frame(topp, bg=BAKGRUND)
        knappar.pack(side='right')
        Knapp(knappar, "Skanna följesedel…", self.fs_skanna, BLA, 'white', '#1e40af').pack(side='left', padx=4)
        Knapp(knappar, "Lägg till PDF:er…", self.fs_valj_pdf, GRON, 'white', GRON_MORK).pack(side='left', padx=4)
        Knapp(knappar, "Exportera lista", self.fs_exportera, '#64748b', 'white', '#475569').pack(side='left', padx=4)
        Knapp(knappar, "Inlärning…", self.fs_inlarning_visa, '#64748b', 'white', '#475569').pack(side='left', padx=4)
        Knapp(knappar, "Spara som träning", self.fs_spara_traning, BLA, 'white', '#1e40af').pack(side='left', padx=4)
        Knapp(knappar, "Skriv ut etiketter", self.fs_skriv_ut, GRON, 'white', GRON_MORK).pack(side='left', padx=4)

        info = tk.Frame(parent, bg='#eef6ff', highlightbackground='#bcd7f5', highlightthickness=1)
        info.pack(fill='x', padx=18, pady=(0, 8))
        tk.Label(info, text="Välj en eller flera följesedlar (PDF från MediCarrier, eller PDF/bild från andra leverantörer). Varje levererad orderrad får en etikett, även artiklar med bulkplats (bulkplats får egen etikett). Endast rader under \"Beställningsvaror/Restorder\" hoppas över. Listan är delad per order (SO-nummer). Klicka på en rad för att välja bort/in den, eller på orderraden för att välja bort/in hela ordern. Klicka på en kolumnrubrik för att sortera inom varje order, t.ex. på Plats.",
                 bg='#eef6ff', fg='#1e3a5f', font=self.f_norm, wraplength=1000, justify='left').pack(side='left', padx=12, pady=8)

        manuell_rad = tk.Frame(parent, bg=BAKGRUND)
        manuell_rad.pack(fill='x', padx=18, pady=(0, 8))
        tk.Label(manuell_rad, text="Manuellt artikelnummer:", bg=BAKGRUND, fg=TEXT, font=self.f_norm).pack(side='left')
        self.fs_manuell_var = tk.StringVar()
        ent = ttk.Entry(manuell_rad, textvariable=self.fs_manuell_var, width=20, font=self.f_norm)
        ent.pack(side='left', padx=(8, 4))
        ent.bind('<Return>', lambda e: self.fs_lagg_till_manuell())
        Knapp(manuell_rad, "Lägg till etikett", self.fs_lagg_till_manuell, BLA, 'white', '#1e40af').pack(side='left', padx=4)
        tk.Label(manuell_rad, text="Skriv art.nr och tryck Retur – raden hamnar i listan nedan.", bg=BAKGRUND, fg=DAMPAD, font=self.f_liten).pack(side='left', padx=(10,0))


        skr = tk.Frame(parent, bg=KORT, highlightbackground=RAM, highlightthickness=1)
        skr.pack(fill='x', padx=18, pady=(0, 8))
        sb = tk.Frame(skr, bg=KORT)
        sb.pack(fill='x', padx=12, pady=8)
        tk.Label(sb, text="Etikettskrivare (Zebra ZD421, direkttermiska etiketter):", bg=KORT, fg=TEXT, font=self.f_fet).grid(row=0, column=0, columnspan=6, sticky='w')
        tk.Label(sb, text="Windows-skrivare (USB):", bg=KORT, fg=TEXT, font=self.f_norm).grid(row=1, column=0, sticky='w')
        self.fs_skrivarval = ttk.Combobox(sb, textvariable=self.p_et_skrivare, values=U.lista_skrivare(), width=34, font=self.f_norm)
        self.fs_skrivarval.grid(row=1, column=1, padx=6, sticky='w')
        Knapp(sb, "Hämta lista", lambda: self.fs_skrivarval.config(values=U.lista_skrivare()), '#94a3b8', 'white', '#64748b').grid(row=1, column=2, padx=4)
        tk.Label(sb, text="Upplösning (dpi):", bg=KORT, fg=TEXT, font=self.f_norm).grid(row=1, column=3, padx=(18, 4))
        ttk.Combobox(sb, textvariable=self.p_et_dpi, values=('203', '300'), width=6, state='readonly', font=self.f_norm).grid(row=1, column=4, sticky='w')
        tk.Label(sb, text="Etikett (mm):", bg=KORT, fg=TEXT, font=self.f_norm).grid(row=2, column=3, padx=(18, 4))
        em = tk.Frame(sb, bg=KORT)
        em.grid(row=2, column=4, sticky='w')
        ttk.Entry(em, textvariable=self.p_et_bredd, width=6, font=self.f_norm).pack(side='left')
        tk.Label(em, text=" x ", bg=KORT, fg=TEXT, font=self.f_norm).pack(side='left')
        ttk.Entry(em, textvariable=self.p_et_hojd, width=6, font=self.f_norm).pack(side='left')
        Knapp(sb, "Testetikett", self.fs_testetikett, '#0ea5e9', 'white', '#0284c7').grid(row=1, column=5, padx=4, sticky='w')
        Knapp(sb, "Spara som ZPL-fil", self.fs_spara_zpl, '#8b5cf6', 'white', '#7c3aed').grid(row=2, column=5, padx=4, sticky='w')

        summa_rad = tk.Frame(parent, bg=BAKGRUND)
        summa_rad.pack(fill='x', padx=18, pady=(0, 4))
        # Rensa ligger till höger, rakt ovanför kolumnen "Ta bort" i listan
        Knapp(summa_rad, "Rensa", self.fs_rensa, '#94a3b8', 'white', '#64748b').pack(side='right', padx=(8, 22))
        self.fs_summa = tk.Label(summa_rad, text="Inga följesedlar inlästa än.", bg=BAKGRUND, fg=DAMPAD, font=self.f_fet, anchor='w', wraplength=950, justify='left')
        self.fs_summa.pack(side='left', fill='x', expand=True)

        ram = tk.Frame(parent, bg=BAKGRUND)
        ram.pack(fill='both', expand=True, padx=18, pady=(0, 12))
        kol = ('ok', 'rad', 'artnr', 'namn', 'plats', 'status', 'del')
        stil = ttk.Style()
        stil.configure('Folj.Treeview', font=(self.f_norm[0], 13), rowheight=34)
        stil.configure('Folj.Treeview.Heading', font=(self.f_fet[0], 12, 'bold'))
        self.fs_tree = ttk.Treeview(ram, columns=kol, show='headings', selectmode='browse', style='Folj.Treeview')
        rubr = {'ok': ('Etikett', 85), 'rad': ('Rad', 55), 'artnr': ('Art.nr', 100), 'namn': ('Namn', 470), 'plats': ('Plats', 130), 'status': ('Status', 250), 'del': ('Ta bort', 80)}
        for k in kol:
            self.fs_tree.heading(k, text=rubr[k][0], command=lambda c=k: self.fs_sortera(c))
            self.fs_tree.column(k, width=rubr[k][1], anchor='center' if k in ('ok', 'rad') else 'w', stretch=(k in ('namn', 'status')))
        rull = ttk.Scrollbar(ram, orient='vertical', command=self.fs_tree.yview)
        self.fs_tree.configure(yscrollcommand=rull.set)
        self.fs_tree.pack(side='left', fill='both', expand=True)
        rull.pack(side='right', fill='y')
        self.fs_tree.tag_configure('order', background='#dbe7f5', foreground='#1e3a5f', font=(self.f_fet[0], 13, 'bold'))
        self.fs_tree.tag_configure('etikett', background='#ffffff')
        self.fs_tree.tag_configure('bort', foreground='#94a3b8')
        self.fs_tree.tag_configure('bulk', background='#fbeef3', foreground='#9d3a63')
        self.fs_tree.tag_configure('saknas', background='#fdecea', foreground='#a02c22')
        self.fs_tree.tag_configure('ingenplats', background='#fff7e0', foreground='#8a5a00')
        self.fs_tree.bind('<Button-1>', self.fs_klick)
        self.fs_tree.bind('<Button-3>', self.fs_hogerklick)
        self.fs_tree.bind('<Button-2>', self.fs_hogerklick)

    def fs_valj_pdf(self):
        filer = filedialog.askopenfilenames(parent=self.rot, title="Välj följesedlar (PDF eller bild)",
                                            filetypes=[('PDF och bilder', '*.pdf *.png *.jpg *.jpeg *.tif *.tiff *.bmp'), ('PDF-filer', '*.pdf'), ('Alla filer', '*.*')])
        nya = [f for f in filer if f not in self.fs_kallor]
        if not nya:
            return
        self.fs_fel, self.fs_klara = [], 0
        self.fs_summa.config(text=f"Läser {len(nya)} filer… (scanningar tar några sekunder per sida)", fg=DAMPAD)

        def jobb():
            for fil in nya:
                try:
                    sedel = None
                    if fil.lower().endswith('.pdf'):
                        sedel = U.las_foljesedel(fil)          # MediCarrier-layouten (textlager)
                    if sedel and sedel['rader']:
                        res = ('medi', sedel)
                    else:
                        res = ('text', U.hamta_text(fil))      # övriga leverantörer: textlager eller OCR
                except Exception as e:
                    self.fs_fel.append(f"{os.path.basename(fil)}: {e}")
                    res = None
                if res:
                    self.fs_nya[fil] = res
                self.fs_klara += 1
        self.fs_nya = {}
        threading.Thread(target=jobb, daemon=True).start()
        self.rot.after(300, lambda: self.fs_vanta(len(nya), nya))

    def fs_vanta(self, antal, nya):
        if self.fs_klara < antal:
            self.fs_summa.config(text=f"Läser följesedlar… {self.fs_klara}/{antal}", fg=DAMPAD)
            self.rot.after(300, lambda: self.fs_vanta(antal, nya))
            return
        for fil in nya:
            if fil in self.fs_nya:
                self.fs_kallor[fil] = self.fs_nya[fil]
                self.fs_ordning.append(fil)
        if self.fs_fel:
            messagebox.showwarning("Följesedlar", "Kunde inte läsa:\n\n" + "\n".join(self.fs_fel), parent=self.rot)
        self.fs_uppdatera()

    def fs_uppdatera(self):
        db, kalla_text = None, ''
        repo, token, gren = self.repo.get().strip(), self.token.get().strip(), self.gren.get().strip() or 'main'
        filnamn = os.path.basename(self.p_ut.get().strip().strip('"')) or 'data.js'
        try:
            db = U.las_databas_text(U.hamta_databas_github(repo, token, gren, filnamn), 'GitHub')
            kalla_text = 'GitHub (publicerad ' + (getattr(U.las_databas_text, 'senaste_stampel', '') or 'okänd tid') + ')'
        except Exception as e:
            self.fs_summa.config(text=f"Kunde inte hämta data.js från GitHub: {e}. Kontrollera att GitHub-uppgifterna i HEM är korrekta.", fg=ROD)
            return
        self.fs_kalla_text = kalla_text
        self.fs_sedlar = []
        lardir = self.fs_lardir()
        try:
            profiler = U.ladda_profiler(lardir) if lardir else {}
        except Exception:
            profiler = {}
        for fil in self.fs_ordning:
            typ, val = self.fs_kallor[fil]
            if typ == 'medi':
                self.fs_sedlar.append(val)
            else:
                lev = U.gissa_leverantor(val, profiler)
                sedel = U.generisk_foljesedel_lard(val, os.path.basename(fil), db, profiler.get(lev), lev)
                sedel['sokvag'] = fil
                self.fs_sedlar.append(sedel)
        self.fs_poster = U.matcha(self.fs_sedlar, db)
        for p in self.fs_poster:
            nyckel = (p['fil'], p['so'], p['rad'])
            p['nyckel'] = nyckel
            self.fs_valda.setdefault(nyckel, p['status'] in ('etikett', 'bulk') and not p.get('osaker'))
        self.fs_rita()

    def fs_status_text(self, p):
        mark = {'neg': 'Markerad som felträff · ', 'pos': 'Markerad som rätt · '}.get(self.fs_markerad.get(p.get('nyckel')), '')
        return mark + {'etikett': 'Ska ha etikett', 'bulk': 'Bulk – ska ha etikett (bulkplats ' + p['bulk'] + ')',
                'saknas': 'Finns inte i databasen', 'ingenplats': 'Saknar plats i databasen'}[p['status']] + ((' · ' + p['notering']) if p['notering'] else '')

    def fs_rita(self):
        t = self.fs_tree
        t.delete(*t.get_children())
        for n, sedel in enumerate(self.fs_sedlar):
            poster = [p for p in self.fs_poster if p['fil'] == sedel['fil'] and p['so'] == sedel['so']]
            if self.fs_sort:
                kol, fall = self.fs_sort
                poster.sort(key=lambda p: (len(p['artnr']), p['artnr']) if kol == 'artnr' else p[kol], reverse=fall)
            ant = sum(1 for p in poster if p['status'] in ('etikett', 'bulk') and self.fs_valda.get(p['nyckel']))
            rubrik = (f"  Order {sedel['so']}   ·   följesedel {sedel['ordernr']}   ·   {sedel['datum']}   ·   " if sedel['so'] else f"  {sedel['fil']}   ·   " + (f"leverantör {sedel['leverantor']}   ·   " if sedel.get('leverantor') else '')) + f"{len(poster)} rader, {ant} etiketter"
            t.insert('', 'end', iid=f"order|{n}", values=('', '', '', rubrik, '', '', ''), tags=('order',))
            for p in poster:
                valt = self.fs_valda.get(p['nyckel'], False)
                kan = p['status'] in ('etikett', 'bulk')
                ruta = ('☑' if valt else '☐') if kan else '–'
                tagg = p['status'] if (valt or not kan) else 'bort'
                t.insert('', 'end', iid=f"{n}|{p['rad']}", values=(ruta, p['rad'], p['artnr'], p['namn'], p['plats'] or '–', self.fs_status_text(p), '✕'), tags=(tagg,))
        s = U.sammanfatta(self.fs_sedlar, self.fs_poster)
        ant = sum(1 for p in self.fs_poster if p['status'] in ('etikett', 'bulk') and self.fs_valda.get(p['nyckel']))
        delar = [f"{s['filer']} ordrar", f"{s['rader']} levererade rader", f"{ant} etiketter att skriva ut (inkl. {s['bulk']} bulk)"]
        if s['saknas']:
            delar.append(f"{s['saknas']} saknas i databasen")
        if s['ingenplats']:
            delar.append(f"{s['ingenplats']} utan plats")
        if s['rest']:
            delar.append(f"{s['rest']} restorderrader hoppades över")
        tomma = [x['fil'] for x in self.fs_sedlar if not x['rader']]
        if tomma:
            delar.append(f"inga artiklar hittades i {len(tomma)} fil(er): " + ', '.join(tomma[:3]))
        if any(x.get('generisk') for x in self.fs_sedlar):
            delar.append('OBS: vid scannade följesedlar hoppas restorder inte över, bocka ur dem själv')
        delar.append('databas: ' + getattr(self, 'fs_kalla_text', ''))
        self.fs_summa.config(text="  ·  ".join(delar), fg=(ROD if s['saknas'] else TEXT))

    def fs_klick(self, ev):
        if self.fs_tree.identify_region(ev.x, ev.y) != 'cell':
            return
        iid = self.fs_tree.identify_row(ev.y)
        col = self.fs_tree.identify_column(ev.x)
        if not iid:
            return
        forst, rest = iid.split('|', 1)
        if col == '#7':
            if forst == 'order':
                return
            sedel = self.fs_sedlar[int(forst)]
            p = next((x for x in self.fs_poster if x['fil'] == sedel['fil'] and x['so'] == sedel['so'] and str(x['rad']) == rest), None)
            if p:
                fil_key = p['fil']
                so_key = p['so']
                self.fs_poster.remove(p)
                self.fs_valda.pop(p.get('nyckel'), None)
                kvar = [x for x in self.fs_poster if x['fil'] == fil_key and x['so'] == so_key]
                if not kvar:
                    self.fs_sedlar = [s for s in self.fs_sedlar if not (s['fil'] == fil_key and s['so'] == so_key)]
                    if fil_key in self.fs_kallor:
                        del self.fs_kallor[fil_key]
                    if fil_key in self.fs_ordning:
                        self.fs_ordning = [f for f in self.fs_ordning if f != fil_key]
                self.fs_rita()
            return
        if forst == 'order':
            sedel = self.fs_sedlar[int(rest)]
            att = [p for p in self.fs_poster if p['fil'] == sedel['fil'] and p['so'] == sedel['so'] and p['status'] in ('etikett', 'bulk')]
            nytt = not all(self.fs_valda.get(p['nyckel']) for p in att)
            for p in att:
                self.fs_valda[p['nyckel']] = nytt
        else:
            sedel = self.fs_sedlar[int(forst)]
            p = next((x for x in self.fs_poster if x['fil'] == sedel['fil'] and x['so'] == sedel['so'] and str(x['rad']) == rest), None)
            if not p or p['status'] not in ('etikett', 'bulk'):
                return
            self.fs_valda[p['nyckel']] = not self.fs_valda.get(p['nyckel'], False)
        self.fs_rita()

    def fs_sortera(self, kol):
        if kol == 'ok':
            return
        fall = bool(self.fs_sort and self.fs_sort[0] == kol and not self.fs_sort[1])
        self.fs_sort = (kol, fall)
        self.fs_rita()

    def fs_rensa(self):
        self.fs_sedlar, self.fs_poster, self.fs_valda, self.fs_kallor, self.fs_ordning = [], [], {}, {}, []
        self.fs_markerad = {}
        self.fs_tree.delete(*self.fs_tree.get_children())
        self.fs_summa.config(text="Inga följesedlar inlästa än.", fg=DAMPAD)

    def fs_hamta_db_for_manuell(self):
        # Alltid GitHub sedan valet GitHub/lokal togs bort
        repo, token, gren = self.repo.get().strip(), self.token.get().strip(), self.gren.get().strip() or 'main'
        filnamn = self.p_ut.get().strip().strip('"') or 'data.js'
        filnamn = __import__('os').path.basename(filnamn)
        db = U.las_databas_text(U.hamta_databas_github(repo, token, gren, filnamn), 'GitHub')
        return db

    def fs_lagg_till_manuell(self):
        art = self.fs_manuell_var.get().strip()
        if not art:
            return
        art = art.upper().replace(' ', '')
        try:
            db = self.fs_hamta_db_for_manuell()
        except Exception as e:
            from tkinter import messagebox
            messagebox.showerror("Manuell etikett", f"Kunde inte lasa databasen: {e}", parent=self.rot)
            return
        idx = U.bygg_index(db)
        traffar = idx.get(art, [])
        manu_fil = 'Manuellt tillagda'
        manu_so = 'MANUELL'
        sedel = next((s for s in self.fs_sedlar if s['fil'] == manu_fil), None)
        if not sedel:
            import datetime
            sedel = {'fil': manu_fil, 'so': manu_so, 'ordernr': 'Manuell', 'datum': datetime.datetime.now().strftime('%Y-%m-%d'), 'rader': [], 'rest': [], 'generisk': False}
            self.fs_sedlar.append(sedel)
            self.fs_ordning.append(manu_fil)
            self.fs_kallor[manu_fil] = ('medi', sedel)
        bef = [p for p in self.fs_poster if p['fil'] == manu_fil]
        naesta_rad = max([p['rad'] for p in bef if isinstance(p['rad'], int)], default=0) + 1 if bef else 1
        def bygg_post_fran_val(val):
            typ, plats, bulk, forsta = val
            if typ == 'bulk':
                return {'fil': manu_fil, 'so': manu_so, 'ordernr': 'Manuell', 'rad': naesta_rad, 'artnr': forsta['artnr'], 'levererat': 1, 'enhet': '', 'fnamn': forsta['namn'], 'namn': forsta['namn'], 'datum': sedel['datum'], 'lev': forsta['lev'], 'plats': plats, 'bulk': bulk, 'status': 'bulk', 'notering': 'Manuellt tillagd'}
            else:
                return {'fil': manu_fil, 'so': manu_so, 'ordernr': 'Manuell', 'rad': naesta_rad, 'artnr': forsta['artnr'], 'levererat': 1, 'enhet': '', 'fnamn': forsta['namn'], 'namn': forsta['namn'], 'datum': sedel['datum'], 'lev': forsta['lev'], 'plats': plats, 'bulk': bulk if bulk else '', 'status': 'etikett' if plats else 'ingenplats', 'notering': 'Manuellt tillagd'}
        if not traffar:
            post = {'fil': manu_fil, 'so': manu_so, 'ordernr': 'Manuell', 'rad': naesta_rad, 'artnr': art, 'levererat': 1, 'enhet': '', 'fnamn': art, 'namn': art + ' (finns ej i databasen)', 'datum': sedel['datum'], 'lev': '', 'plats': '', 'bulk': '', 'status': 'saknas', 'notering': 'Manuellt tillagd, saknas i DB'}
        else:
            alternativ = []
            for t in traffar:
                if t.get('plats'):
                    alternativ.append(('forrad', t['plats'], t.get('bulk',''), t))
                if t.get('bulk'):
                    alternativ.append(('bulk', t['plats'], t['bulk'], t))
            unika = []
            sedd = set()
            for alt in alternativ:
                nyckel = (alt[0], alt[1], alt[2])
                if nyckel not in sedd:
                    sedd.add(nyckel)
                    unika.append(alt)
            alternativ = unika
            if len(alternativ) == 0:
                forsta = traffar[0]
                post = {'fil': manu_fil, 'so': manu_so, 'ordernr': 'Manuell', 'rad': naesta_rad, 'artnr': forsta['artnr'], 'levererat': 1, 'enhet': '', 'fnamn': forsta['namn'], 'namn': forsta['namn'], 'datum': sedel['datum'], 'lev': forsta['lev'], 'plats': '', 'bulk': '', 'status': 'ingenplats', 'notering': 'Manuellt tillagd'}
            elif len(alternativ) == 1:
                post = bygg_post_fran_val(alternativ[0])
            else:
                valt = self.fs_visa_platsval_dialog(art, alternativ)
                if valt is None:
                    return
                post = bygg_post_fran_val(valt)
        post['nyckel'] = (post['fil'], post['so'], post['rad'])
        self.fs_poster.append(post)
        self.fs_valda[post['nyckel']] = post['status'] in ('etikett', 'bulk')
        sedel['rader'].append({'artnr': post['artnr'], 'rad': post['rad'], 'levererat': 1, 'enhet': '', 'namn': post['namn']})
        self.fs_manuell_var.set('')
        self.fs_rita()
        self.status_rad(f"Lade till {post['artnr']} som rad {naesta_rad} – {self.fs_status_text(post)}", GRON_MORK)

    def fs_visa_platsval_dialog(self, artnr, alternativ):
        import tkinter as tk
        val_holder = {'val': None}
        top = tk.Toplevel(self.rot)
        top.title(f"Valj plats for {artnr}")
        top.transient(self.rot)
        top.grab_set()
        top.resizable(False, False)
        frm = tk.Frame(top, bg=KORT, padx=12, pady=12)
        frm.pack(fill='both', expand=True)
        tk.Label(frm, text=f"{artnr} finns pa flera platser – valj vilken etikett som ska skrivas:", bg=KORT, fg=TEXT, font=self.f_norm, wraplength=400, justify='left').pack(anchor='w', pady=(0,10))
        var = tk.StringVar(value="0")
        for i, (typ, plats, bulk, forsta) in enumerate(alternativ):
            if typ == 'bulk':
                txt = f"BULKPLATS: {bulk}  (forrad: {plats or '-'}) – {forsta['namn']}"
            else:
                txt = f"FORRADSPLATS: {plats}  (bulk: {bulk or '-'}) – {forsta['namn']}"
            rb = tk.Radiobutton(frm, text=txt, variable=var, value=str(i), bg=KORT, fg=TEXT, font=self.f_norm, anchor='w', justify='left', wraplength=420)
            rb.pack(fill='x', anchor='w', pady=2)
        btn_rad = tk.Frame(frm, bg=KORT)
        btn_rad.pack(fill='x', pady=(12,0))
        def ok():
            try:
                idx = int(var.get())
                val_holder['val'] = alternativ[idx]
            except Exception:
                val_holder['val'] = None
            top.destroy()
        def avbryt():
            val_holder['val'] = None
            top.destroy()
        tk.Button(btn_rad, text="Avbryt", command=avbryt, bg=KORT, fg=TEXT).pack(side='right', padx=(6,0))
        tk.Button(btn_rad, text="OK", command=ok, bg=BLA, fg='white', font=self.f_fet).pack(side='right')
        top.bind('<Return>', lambda e: ok())
        top.bind('<Escape>', lambda e: avbryt())
        top.update_idletasks()
        x = self.rot.winfo_rootx() + (self.rot.winfo_width() - top.winfo_width())//2
        y = self.rot.winfo_rooty() + (self.rot.winfo_height() - top.winfo_height())//3
        top.geometry(f"+{x}+{y}")
        self.rot.wait_window(top)
        return val_holder['val']

    def fs_skanna(self):
        """Lägg en A4 i scannern och tryck: sidan skannas utan dialogruta, sedan frågar programmet om det finns fler sidor."""
        basmapp = self.p_skannade.get().strip().strip('"') if hasattr(self, 'p_skannade') else r"G:\Sos\PVI_Steril\Förrådet\Scannade_Följesedlar"
        if basmapp.lower().endswith('.pdf'):
            basmapp = os.path.dirname(basmapp)
        self._fs_sidor = []
        self._fs_basmapp = basmapp
        self.fs_skanna_sida()

    def fs_skanna_sida(self):
        n = len(self._fs_sidor) + 1
        self.fs_summa.config(text=f"Skannar sida {n}… vänta tills scannern är klar", fg=BLA)
        self._fs_sida_res = None

        def jobb():
            try:
                self._fs_sida_res = (U.skanna_sida(), None)
            except Exception as e:
                self._fs_sida_res = (None, str(e))
        threading.Thread(target=jobb, daemon=True).start()
        self.rot.after(300, self.fs_skanna_sida_vanta)

    def fs_skanna_sida_vanta(self):
        if self._fs_sida_res is None:
            self.rot.after(300, self.fs_skanna_sida_vanta)
            return
        bild, fel = self._fs_sida_res
        self._fs_sida_res = None
        if fel:
            if not self._fs_sidor:
                messagebox.showerror("Skanning", f"Kunde inte skanna:\n{fel}", parent=self.rot)
                self.fs_summa.config(text="Skanning misslyckades", fg=ROD)
                return
            messagebox.showwarning("Skanning", f"Sida {len(self._fs_sidor) + 1} kunde inte skannas ({fel}).\nDe {len(self._fs_sidor)} sidor som redan är skannade används.", parent=self.rot)
        else:
            self._fs_sidor.append(bild)
            if messagebox.askyesno("Skanna fler sidor?", f"Sida {len(self._fs_sidor)} är skannad.\n\nHar följesedeln fler sidor?\n"
                                   "Lägg nästa sida i scannern och välj Ja. Välj Nej när alla sidor är skannade.", parent=self.rot):
                self.fs_skanna_sida()
                return
        self.fs_skanna_klar()

    def fs_skanna_klar(self):
        sidor, basmapp = self._fs_sidor, self._fs_basmapp
        self.fs_summa.config(text=f"Läser texten på {len(sidor)} sidor…", fg=BLA)
        self._fs_scan_result = None
        self._fs_ocr_fel = ''

        def jobb():
            try:
                delar, fel_ocr = [], ''
                for sida in sidor:
                    try:
                        delar.append(U.ocr_bild(sida))
                    except Exception as e:
                        fel_ocr = str(e)
                text = '\n'.join(delar)
                self._fs_ocr_fel = fel_ocr
                so_num = U.extrahera_so_från_text(text)
                pdf_path = U.spara_skannad_pdf(sidor, so_num, basmapp=basmapp)    # PDF:en sparas även om texten inte gick att läsa
                self._fs_scan_result = (pdf_path, text, so_num, sidor, basmapp, None)
            except Exception as e:
                import traceback
                self._fs_scan_result = (None, None, None, None, None, f"{e}\n{traceback.format_exc()}")
        threading.Thread(target=jobb, daemon=True).start()
        self.rot.after(300, self.fs_skanna_vanta)

    def fs_skanna_vanta(self):
        if self._fs_scan_result is None:
            self.rot.after(300, self.fs_skanna_vanta)
            return
        res = self._fs_scan_result
        if len(res) == 6:
            pdf_path, ocr_text, so_num, pil_img, basmapp, fel = res
        else:
            pdf_path, ocr_text, so_num, pil_img, fel = res
            basmapp = self.p_skannade.get().strip().strip('"') if hasattr(self, 'p_skannade') else r"G:\Sos\PVI_Steril\Förrådet\Scannade_Följesedlar"
            if basmapp.lower().endswith('.pdf'):
                import os
                basmapp = os.path.dirname(basmapp)
        self._fs_scan_result = None
        if fel:
            from tkinter import messagebox
            messagebox.showerror("Skanning", f"Kunde inte skanna:\n{fel}", parent=self.rot)
            self.fs_summa.config(text="Skanning misslyckades", fg=ROD)
            return
        if not so_num:
            import tkinter.simpledialog as sd
            so_num = sd.askstring("SO-nummer", "Kunde inte hitta SO-nummer automatiskt (SOxxxxxxxx).\nAnge SO-nummer manuellt eller lämna tomt för OKAND:", parent=self.rot)
            if so_num is None:
                self.fs_summa.config(text="Skanning avbruten", fg=DAMPAD)
                return
            so_num = so_num.strip().upper()
            if so_num:
                try:
                    import os
                    ny_pdf = U.spara_skannad_pdf(pil_img, so_num, basmapp=basmapp)
                    if os.path.abspath(ny_pdf) != os.path.abspath(pdf_path):
                        try:
                            os.remove(pdf_path)
                        except Exception:
                            pass
                    pdf_path = ny_pdf
                except Exception:
                    pass
        if getattr(self, '_fs_ocr_fel', ''):
            messagebox.showwarning("Skanning", "Följesedeln är sparad som PDF men texten kunde inte läsas:\n" + self._fs_ocr_fel[:600] +
                                   "\n\nNär OCR fungerar kan du lägga till PDF:en med 'Lägg till PDF:er'.", parent=self.rot)
        self.fs_kallor[pdf_path] = ('text', ocr_text)
        if pdf_path not in self.fs_ordning:
            self.fs_ordning.append(pdf_path)
        self.fs_summa.config(text=f"Skannad {so_num or 'OKAND'} sparad som {pdf_path} – läser mot databas...", fg=BLA)
        try:
            self.fs_uppdatera()
            self.status_rad(f"Skannad följesedel {so_num or 'OKAND'} tillagd ({len([p for p in self.fs_poster if p['fil']==pdf_path])} rader) – sparad i {pdf_path}", GRON_MORK)
        except Exception as e:
            from tkinter import messagebox
            messagebox.showwarning("Skanning", f"PDF sparad i {pdf_path} men kunde inte läsas mot databas:\n{e}", parent=self.rot)


    # ------------------------------------------------------------ Inlärning (andra leverantörer än MediCarrier)
    def fs_lardir(self):
        """Mapp för inlärd data (inställd på fliken HEM, kortet Inlärning av följesedlar)."""
        return (self.p_larande.get().strip().strip('"') if hasattr(self, 'p_larande') else '')

    def fs_hogerklick(self, ev):
        iid = self.fs_tree.identify_row(ev.y)
        if not iid or iid.startswith('order|'):
            return
        forst, rest = iid.split('|', 1)
        try:
            sedel = self.fs_sedlar[int(forst)]
        except (ValueError, IndexError):
            return
        p = next((x for x in self.fs_poster if x['fil'] == sedel['fil'] and x['so'] == sedel['so'] and str(x['rad']) == rest), None)
        if not p or not sedel.get('sokvag'):
            return            # inlärning gäller bara inlästa/scannade följesedlar, inte MediCarrier eller manuella rader
        meny = tk.Menu(self.rot, tearoff=0)
        meny.add_command(label="Rätt artikelnummer – lär in", command=lambda: self.fs_markera(p, 'pos'))
        meny.add_command(label="Fel träff – lär bort", command=lambda: self.fs_markera(p, 'neg'))
        meny.add_command(label="Ta bort markering", command=lambda: self.fs_markera(p, None))
        try:
            meny.tk_popup(ev.x_root, ev.y_root)
        finally:
            meny.grab_release()

    def fs_markera(self, p, vad):
        if vad:
            self.fs_markerad[p['nyckel']] = vad
        else:
            self.fs_markerad.pop(p['nyckel'], None)
        if vad == 'neg':
            self.fs_valda[p['nyckel']] = False
        elif vad == 'pos' and p['status'] in ('etikett', 'bulk'):
            self.fs_valda[p['nyckel']] = True
        self.fs_rita()

    def fs_registrera_inlarning(self):
        """Anropas när etiketter skrivs ut: sparar texten och facit (godkända och felmarkerade nummer) för varje inläst/scannad följesedel
        och lär om leverantörens profil. Får aldrig stoppa utskriften, så alla fel fångas."""
        lardir = self.fs_lardir()
        if not lardir:
            return 0
        sparade = 0
        for sedel in self.fs_sedlar:
            sokv = sedel.get('sokvag')
            if not sokv or sokv not in self.fs_kallor or not sedel.get('generisk'):
                continue
            try:
                text = self.fs_kallor[sokv][1]
                poster = [p for p in self.fs_poster if p['fil'] == sedel['fil'] and p['so'] == sedel['so']]
                pos = {p['lasta'] for p in poster if (self.fs_valda.get(p['nyckel']) and p['status'] in ('etikett', 'bulk')) or self.fs_markerad.get(p['nyckel']) == 'pos'}
                neg = {p['lasta'] for p in poster if self.fs_markerad.get(p['nyckel']) == 'neg'}
                if not pos:
                    continue
                lev = sedel.get('leverantor') or ''
                kt = []
                if not lev:
                    import tkinter.simpledialog as sd
                    gissning = U.foresla_leverantor(text)
                    svar = sd.askstring("Ny leverantör", "Vilken leverantör kommer följesedeln " + sedel['fil'] + " från?\n"
                                        "Programmet lär sig hur just den leverantörens följesedlar ser ut.\n"
                                        "Skriv ett namn som också står på följesedeln (lämna tomt för att hoppa över).",
                                        initialvalue=gissning, parent=self.rot)
                    lev = (svar or '').strip()
                    if not lev:
                        continue
                    kt = [lev] if lev.lower() in text.lower() else ([gissning] if gissning else [])
                    for s2 in self.fs_sedlar:               # samma leverantör på övriga okända följesedlar i listan
                        if s2.get('generisk') and not s2.get('leverantor') and s2.get('sokvag') in self.fs_kallor and lev.lower() in self.fs_kallor[s2['sokvag']][1].lower():
                            s2['leverantor'] = lev
                    sedel['leverantor'] = lev
                U.registrera_inlarning(lardir, lev, sokv, text, pos, neg)
                U.lar_profil(lardir, lev, kt)
                sparade += 1
            except Exception as e:
                self.status_rad(f"Kunde inte spara inlärning: {e}", ROD)
                return sparade
        if sparade:
            self.status_rad(f"Inlärning: sparade {sparade} följesedlar i {lardir}", GRON_MORK)
        return sparade

    def fs_spara_traning(self):
        """Sparar texten och dina rättelser för de inlästa/scannade följesedlarna som träning, utan att skriva ut något."""
        if not any(s.get('sokvag') and s.get('generisk') for s in self.fs_sedlar):
            messagebox.showinfo("Spara som träning", "Det finns inga inlästa eller scannade följesedlar att träna på. (MediCarrier och manuella rader används inte till inlärning.)", parent=self.rot)
            return
        if not self.fs_lardir():
            messagebox.showinfo("Spara som träning", "Ange mappen för inlärning på fliken HEM först.", parent=self.rot)
            return
        if not messagebox.askyesno("Spara som träning", "Spara följesedlarna som träning utan att skriva ut?\n\nDet som är ikryssat räknas som rätt, och det du markerat som felträff räknas som fel. Gå igenom listan först.", parent=self.rot):
            return
        n = self.fs_registrera_inlarning()
        if not n:
            messagebox.showinfo("Spara som träning", "Inget sparades. Det måste finnas minst en ikryssad rad per följesedel, och leverantörsnamn behövs för nya leverantörer.", parent=self.rot)
            return
        if messagebox.askyesno("Spara som träning", f"Sparade {n} följesedlar som träning.\n\nRensa listan så att du kan läsa in nästa?", parent=self.rot):
            self.fs_rensa()

    def fs_inlarning_visa(self):
        lardir = self.fs_lardir()
        if not lardir or not os.path.isdir(lardir):
            messagebox.showinfo("Inlärning", "Ingen inlärning sparad än. Den börjar när du skriver ut etiketter från en scannad eller inläst följesedel (inte MediCarrier).\n\n"
                                "Sparas i: " + (lardir or "(ange mappen för inlärning på fliken HEM)"), parent=self.rot)
            return
        try:
            db = self.fs_hamta_db_for_manuell()
            rader, _ = U.utvardera_inlarning(lardir, db)
        except Exception as e:
            messagebox.showerror("Inlärning", f"Kunde inte köra testet:\n{e}", parent=self.rot)
            return
        prof = U.ladda_profiler(lardir)
        extra = [f"{lev}: {p.get('antal_sedlar', 0)} följesedlar, kännetecken: {', '.join(p.get('kannetecken', [])) or '(inget än)'}" for lev, p in sorted(prof.items())]
        messagebox.showinfo("Inlärning", ("Testet kör om alla sparade följesedlar med och utan inlärning:\n\n" + "\n\n".join(rader) if rader else "Inga sparade följesedlar att testa.")
                            + "\n\n" + "\n".join(extra) + "\n\nMapp: " + lardir
                            + "\n\nEtt mönster används först när det setts på minst 2 följesedlar.", parent=self.rot)

    def fs_etiketter(self):
        return [p for p in self.fs_poster if p['status'] in ('etikett', 'bulk') and self.fs_valda.get(p['nyckel'])]

    def fs_exportera(self):
        if not self.fs_poster:
            messagebox.showinfo("Följesedlar", "Lägg till minst en följesedel först.", parent=self.rot)
            return
        sokv = filedialog.asksaveasfilename(parent=self.rot, title="Spara lista", defaultextension='.csv', initialfile='etiketter.csv',
                                            filetypes=[('CSV (öppnas i Excel)', '*.csv')])
        if not sokv:
            return
        import csv
        with open(sokv, 'w', newline='', encoding='utf-8-sig') as f:
            w = csv.writer(f, delimiter=';')
            w.writerow(['Order (SO)', 'Följesedel', 'Rad', 'Artikelnummer', 'Namn', 'Plats', 'Bulkplats', 'Levererat', 'Enhet', 'Etikett', 'Status'])
            for p in self.fs_poster:
                w.writerow([p['so'], p['ordernr'], p['rad'], p['artnr'], p['namn'], p['plats'], p['bulk'], p['levererat'], p['enhet'],
                            'ja' if (p['status'] == 'etikett' and self.fs_valda.get(p['nyckel'])) else 'nej', self.fs_status_text(p)])
        messagebox.showinfo("Följesedlar", f"Sparade {len(self.fs_poster)} rader i\n{sokv}", parent=self.rot)

    def fs_mal_text(self):
        if self.p_et_metod.get() == 'natverk':
            return "skrivaren på " + (self.p_et_ip.get().strip() or "(IP saknas)")
        return "skrivaren '" + (self.p_et_skrivare.get().strip() or "(ingen vald)") + "'"

    def fs_dpi(self):
        try:
            return int(self.p_et_dpi.get())
        except ValueError:
            return 203

    def fs_mm(self):
        try:
            return float(self.p_et_bredd.get().replace(',', '.')), float(self.p_et_hojd.get().replace(',', '.'))
        except ValueError:
            return 70.0, 50.8

    def fs_skicka(self, zpl):
        U.skriv_ut_zpl(zpl, self.p_et_metod.get(), self.p_et_skrivare.get().strip(), self.p_et_ip.get().strip())

    def fs_testetikett(self):
        try:
            self.spara(tyst=True)
            self.fs_skicka(U.zpl_testetikett(self.fs_dpi(), *self.fs_mm()))
            self.status_rad("Testetikett skickad till " + self.fs_mal_text(), GRON_MORK)
        except Exception as e:
            messagebox.showerror("Testetikett", f"Kunde inte skriva ut:\n{e}", parent=self.rot)

    def fs_zpl_for_valda(self):
        dpi = self.fs_dpi()
        delar = []
        for sedel in self.fs_sedlar:
            poster = [p for p in self.fs_etiketter() if p['fil'] == sedel['fil'] and p['so'] == sedel['so']]
            for p in sorted(poster, key=lambda p: p['rad']):
                rad_for_etikett = None if p['fil'] == 'Manuellt tillagda' else p['rad']
                delar.append(U.zpl_etikett(p['plats'], p['namn'], p['artnr'], p['lev'], p['so'], p['datum'], dpi, *self.fs_mm(), rad=rad_for_etikett))
        return ''.join(delar), len(delar)

    def fs_spara_zpl(self):
        zpl, n = self.fs_zpl_for_valda() if self.fs_etiketter() else (U.zpl_testetikett(self.fs_dpi(), *self.fs_mm()), 0)
        sokv = filedialog.asksaveasfilename(parent=self.rot, title="Spara ZPL-fil", defaultextension='.zpl', initialfile='etiketter.zpl',
                                            filetypes=[('ZPL-fil', '*.zpl'), ('Textfil', '*.txt')])
        if not sokv:
            return
        with open(sokv, 'w', encoding='utf-8', newline='\n') as fh:
            fh.write(zpl)
        messagebox.showinfo("ZPL-fil", (f"Sparade {n} etiketter i\n{sokv}" if n else f"Sparade en testetikett i\n{sokv}") +
                            "\n\nFilen kan klistras in på labelary.com för att se hur etiketterna ser ut.", parent=self.rot)

    def fs_skriv_ut(self):
        att = self.fs_etiketter()
        if not att:
            messagebox.showinfo("Skriv ut etiketter", "Inga etiketter är valda. Lägg till en följesedel och bocka i de rader som ska skrivas ut.", parent=self.rot)
            return
        if not messagebox.askyesno("Skriv ut etiketter", f"Skriva ut {len(att)} etiketter på {self.fs_mal_text()}?", parent=self.rot):
            return
        try:
            self.spara(tyst=True)
            zpl, n = self.fs_zpl_for_valda()
            self.fs_skicka(zpl)
        except Exception as e:
            messagebox.showerror("Skriv ut etiketter", f"Kunde inte skriva ut:\n{e}", parent=self.rot)
            return
        self.status_rad(f"{n} etiketter skickade till {self.fs_mal_text()}", GRON_MORK)
        try:
            self.fs_registrera_inlarning()
        except Exception:
            pass
        try:
            top = tk.Toplevel(self.rot)
            top.title("Skriv ut etiketter")
            top.transient(self.rot)
            top.resizable(False, False)
            top.update_idletasks()
            w, h = 360, 110
            rx = self.rot.winfo_rootx() + (self.rot.winfo_width() - w)//2
            ry = self.rot.winfo_rooty() + (self.rot.winfo_height() - h)//3
            top.geometry(f"{w}x{h}+{max(rx,0)}+{max(ry,0)}")
            frm = tk.Frame(top, bg=KORT, highlightbackground=RAM, highlightthickness=1)
            frm.pack(fill='both', expand=True, padx=1, pady=1)
            tk.Label(frm, text="🖨️", bg=KORT, font=(self.f_rubrik_stor[0], 22)).pack(pady=(12,2))
            tk.Label(frm, text=f"Skriver ut {n} etiketter...", bg=KORT, fg=TEXT, font=self.f_fet).pack()
            tk.Label(frm, text=self.fs_mal_text(), bg=KORT, fg=DAMPAD, font=self.f_liten).pack(pady=(2,0))
            top.after(2500, top.destroy)
        except Exception:
            pass

    def bygg_statistik_flik(self, parent):
        scroll = ScrollableFrame(parent, bg=BAKGRUND)
        scroll.pack(fill='both', expand=True)
        kropp = tk.Frame(scroll.scrollable_frame, bg=BAKGRUND)
        kropp.pack(fill='both', expand=True, padx=18, pady=14)
        self.stat_status = tk.Label(kropp, text="", bg=BAKGRUND, fg=DAMPAD, font=self.f_liten)
        self.stat_status.pack(fill='x', pady=(0, 6))

        # Kontrollrad
        ctrl = tk.Frame(kropp, bg=KORT)
        ctrl.pack(fill='x', pady=(0,12), padx=1)
        tk.Frame(ctrl, bg=BLA, height=3).pack(fill='x')
        inner_ctrl = tk.Frame(ctrl, bg=KORT)
        inner_ctrl.pack(fill='x', padx=12, pady=10)
        tk.Label(inner_ctrl, text="Statistik", bg=KORT, fg=TEXT, font=self.f_rubrik_stor).pack(side='left')
        tk.Label(inner_ctrl, text="Period:", bg=KORT, fg=DAMPAD, font=self.f_liten).pack(side='right', padx=(12,6))
        ttk.Combobox(inner_ctrl, textvariable=self.stat_dagar, values=('7','14','30','90','365'), width=6, state='readonly').pack(side='right')
        Knapp(inner_ctrl, "↻ Uppdatera statistik", self.hamta_statistik, BLA, hover='#3a6d8f').pack(side='right', padx=12)

        # KPI-rad 4 kort
        self.stat_tiles = []
        tr = tk.Frame(kropp, bg=BAKGRUND)
        tr.pack(fill='x', pady=(0, 10))
        for i in range(4):
            tr.columnconfigure(i, weight=1, uniform='st')
        kpi_defs = [("Totalt sökningar", BLA), ("Snitt / dag", BLA), ("Träffsäkerhet", BLA), ("Utan träff", ROD)]   # färg = betydelse: röd bara för problem
        for i, (txt, farg) in enumerate(kpi_defs):
            ram = tk.Frame(tr, bg=RAM)
            ram.grid(row=0, column=i, sticky='nsew', padx=(0 if i==0 else 6, 0))
            inre = tk.Frame(ram, bg=KORT)
            inre.pack(fill='both', expand=True, padx=1, pady=1)
            tk.Frame(inre, bg=farg, height=4).pack(fill='x')
            n = tk.Label(inre, text="–", bg=KORT, fg=farg, font=(FONT_FET, 28, 'bold'), anchor='w')
            n.pack(fill='x', padx=12, pady=(8,0))
            sub = tk.Label(inre, text=txt, bg=KORT, fg=DAMPAD, font=(FONT, 11), anchor='w', wraplength=300)
            sub.pack(fill='x', padx=12, pady=(0,10))
            # dynamisk wraplength
            def _bind_w(e, lbl=sub):
                lbl.config(wraplength=e.width-24)
            inre.bind('<Configure>', _bind_w)
            self.stat_tiles.append(n)

        # Diagram rad 1 - tre staplar
        dg = tk.Frame(kropp, bg=BAKGRUND)
        dg.pack(fill='x', pady=(0, 10))
        for i in range(3):
            dg.columnconfigure(i, weight=1, uniform='dg')
        self.stat_c_dag = self._diagram_kort(dg, 0, "Sökningar per dag (senaste perioden)", BLA, typ='linje')
        self.stat_c_tim = self._diagram_kort(dg, 1, "Fördelning över dygnet", BLA, typ='stapel')
        self.stat_c_vdag = self._diagram_kort(dg, 2, "Veckodag", BLA, typ='stapel')

        # Diagram rad 2 - nyttiga fördelningar
        dg2 = tk.Frame(kropp, bg=BAKGRUND)
        # Raden visas inte längre (dubblerade toppkorten och enhetsrutorna). Rutorna skapas kvar så uppdateringskoden fungerar.
        for i in range(3):
            dg2.columnconfigure(i, weight=1, uniform='dg2')
        self.stat_c_traff = self._diagram_kort(dg2, 0, "Träff vs utan träff", GRON_MORK, typ='donut')
        self.stat_c_enhet = self._diagram_kort(dg2, 1, "Enhet / plattform", BLA, typ='horisontell')
        self.stat_c_kamera = self._diagram_kort(dg2, 2, "Sökmetod", GUL, typ='horisontell')

        # Nyttig fakta - ej rolig
        fk, fh = self.kort(kropp, "Nyttig insikt", BLA, "")
        fk.pack(fill='x', pady=(0, 10))
        self.stat_fakta_frame = tk.Frame(fh, bg=KORT)
        self.stat_fakta_frame.pack(fill='x', pady=(10,0))
        self.stat_fakta_placeholder = tk.Label(self.stat_fakta_frame, text="Hämta statistik för att se nyttiga insikter.", bg=KORT, fg=DAMPAD, font=self.f_norm, anchor='w', wraplength=800, justify='left')
        self.stat_fakta_placeholder.pack(fill='x', pady=6)

        def _visa_fakta(kort_lista, kamera_text=""):
            for w in self.stat_fakta_frame.winfo_children():
                w.destroy()
            if not kort_lista:
                tk.Label(self.stat_fakta_frame, text="Inga sökningar än.", bg=KORT, fg=DAMPAD, font=self.f_norm, anchor='w', wraplength=700, justify='left').pack(fill='x')
                return
            rad = None
            slots_kvar = 0
            for k in kort_lista:
                stor = k.get('stor', False)
                cspan = 2 if stor else 1
                if rad is None or slots_kvar < cspan:
                    rad = tk.Frame(self.stat_fakta_frame, bg=KORT)
                    rad.pack(fill='x')
                    slots_kvar = 3
                kort_w = self._fakta_kort(rad, k['ikon'], k['rubrik'], k['varde'], k['undertext'], k['farg'], k['accent'], progress=k.get('progress'), sparkline=k.get('sparkline'), stor=stor)
                kort_w.pack(side='left', fill='both', expand=True, padx=4, pady=4)
                slots_kvar -= cspan
            if kamera_text:
                banner = tk.Frame(self.stat_fakta_frame, bg="#f0f9ff", highlightbackground="#bae6fd", highlightthickness=1)
                banner.pack(fill='x', pady=(14,0), ipady=10)
                tk.Label(banner, text="📷", bg="#f0f9ff", font=(FONT, 18)).pack(side='left', padx=(12,6))
                lbl = tk.Label(banner, text=kamera_text.strip(), bg="#f0f9ff", fg="#0c4a6e", font=self.f_norm, anchor='w', justify='left')
                lbl.pack(side='left', fill='x', expand=True, padx=(0,12))
                def _wrap_banner(e):
                    lbl.config(wraplength=e.width-60)
                banner.bind('<Configure>', _wrap_banner)

        self._visa_fakta = _visa_fakta

        # Enheter: tre rutor med staplar (Webbläsare är borttagen, den gick sällan att agera på)
        ek = tk.Frame(kropp, bg=BAKGRUND)
        ek.pack(fill='x', pady=(0, 10))
        for i in range(3):
            ek.columnconfigure(i, weight=1, uniform='en')
        self.stat_enh = []
        for i, rub in enumerate(("Enhetstyp", "Operativsystem", "Installerad som app")):
            t = self._lista_kort(ek, i, rub, BLA, ikon="")
            t._stapel = True
            self.stat_enh.append(t)
            self.stat_text(t, "–")

        # Kamera: ett flöde i stället för fyra separata siffror
        kram = tk.Frame(kropp, bg=RAM)
        kram.pack(fill='x', pady=(0, 10))
        kin = tk.Frame(kram, bg=KORT)
        kin.pack(fill='both', expand=True, padx=1, pady=1)
        tk.Frame(kin, bg=BLA, height=4).pack(fill='x')
        tk.Label(kin, text="📷  Kameran", bg=KORT, fg=TEXT, font=self.f_rubrik, anchor='w').pack(fill='x', padx=12, pady=(8, 2))
        flode = tk.Frame(kin, bg=KORT)
        flode.pack(fill='x', padx=12, pady=(2, 12))
        self.stat_kam = [None] * 4
        def _steg(kol, idx, text):
            f = tk.Frame(flode, bg='#fbfcfe', highlightbackground='#e2e8f0', highlightthickness=1)
            f.grid(row=0, column=kol, sticky='nsew', padx=2)
            n = tk.Label(f, text="–", bg='#fbfcfe', fg=BLA, font=(FONT_FET, 26, 'bold'))
            n.pack(padx=14, pady=(10, 0))
            tk.Label(f, text=text, bg='#fbfcfe', fg=DAMPAD, font=(FONT, 10), wraplength=170, justify='center').pack(padx=10, pady=(0, 10))
            self.stat_kam[idx] = n
        def _pil(kol, text):
            tk.Label(flode, text="→\n" + text, bg=KORT, fg=DAMPAD, font=(FONT, 10), justify='center').grid(row=0, column=kol, padx=4)
        for kol in (0, 2, 4, 6):
            flode.columnconfigure(kol, weight=1, uniform='kamflode')
        _steg(0, 0, "Kameran öppnades")
        _pil(1, "")
        _steg(2, 3, "av dem ledde till en sökning")
        _pil(3, "")
        _steg(4, 1, "sökningar via kameran")
        _pil(5, "")
        _steg(6, 2, "av alla sökningar görs med kameran")

        # 4. Topplistor som fadeade stapeldiagram
        dg3 = tk.Frame(kropp, bg=BAKGRUND)
        dg3.pack(fill='x', pady=(0,10), before=dg)   # högst upp: det här är det du kan agera på
        for i in range(2):
            dg3.columnconfigure(i, weight=1, uniform='dg3')
        self.stat_c_utan = self._diagram_kort(dg3, 0, "Utan träff – rätta eller lägg till i databasen", ROD, typ='horisontell')
        self.stat_c_utan.config(height=340)
        self.stat_c_mest = self._diagram_kort(dg3, 1, "Mest sökta", BLA, typ='horisontell')
        self.stat_c_mest.config(height=340)

        # Så söker man: Sökt som + Söklängd i ett kort
        li = tk.Frame(kropp, bg=BAKGRUND)
        li.pack(fill='x', pady=(0, 10))
        li.columnconfigure(0, weight=1)
        kort_sok = self._lista_kort(li, 0, "Så söker man", BLA, ikon="🔍")
        kort_sok.columnconfigure(0, weight=1, uniform='sm')
        kort_sok.columnconfigure(1, weight=1, uniform='sm')
        def _delruta(kol, rubrik):
            d = tk.Frame(kort_sok, bg=KORT)
            d.grid(row=0, column=kol, sticky='nsew', padx=(0 if kol == 0 else 8, 0))
            tk.Label(d, text=rubrik, bg=KORT, fg=DAMPAD, font=(FONT_FET, 10, 'bold'), anchor='w').pack(fill='x', padx=4, pady=(0, 2))
            u = tk.Frame(d, bg=KORT)
            u.pack(fill='both', expand=True)
            u._grafisk = True
            u._rubrik = rubrik
            return u
        self.stat_t_mest = _delruta(0, "Sökt som")
        self.stat_t_utan = _delruta(1, "Söklängd")
        self.stat_text(self.stat_t_mest, "Hämta statistik.")
        self.stat_text(self.stat_t_utan, "Hämta statistik.")

    def _diagram_kort(self, foralder, kol, rubrik, farg, typ='stapel'):
        yttre = tk.Frame(foralder, bg=RAM)
        yttre.grid(row=0, column=kol, sticky='nsew', padx=(0 if kol==0 else 6, 0))
        inre = tk.Frame(yttre, bg=KORT)
        inre.pack(fill='both', expand=True, padx=1, pady=1)
        tk.Frame(inre, bg=farg, height=4).pack(fill='x')
        # rubrik med wraplength
        rub_lbl = tk.Label(inre, text=rubrik, bg=KORT, fg=TEXT, font=self.f_rubrik, anchor='w', justify='left')
        rub_lbl.pack(fill='x', padx=12, pady=(8,4))
        def _wrap_rub(e):
            rub_lbl.config(wraplength=e.width-24)
        inre.bind('<Configure>', _wrap_rub)
        c = tk.Canvas(inre, bg=KORT, height=200, highlightthickness=0)
        c.pack(fill='both', expand=True, padx=10, pady=(0, 10))
        c._farg = farg
        c._data = None
        c._typ = typ
        c.bind('<Configure>', lambda e, c=c: self._rita_diagram(c))
        return c

    def _lista_kort(self, foralder, kol, rubrik, farg, ikon=None):
        yttre = tk.Frame(foralder, bg=RAM)
        yttre.grid(row=0, column=kol, sticky='nsew', padx=(0 if kol==0 else 6, 0))
        inre = tk.Frame(yttre, bg=KORT)
        inre.pack(fill='both', expand=True, padx=1, pady=1)
        tk.Frame(inre, bg=farg, height=4).pack(fill='x')
        if ikon is None:
            ikon = "🔢" if "Sökt som" in rubrik else "📏" if "Söklängd" in rubrik else "✕"
        lbl = tk.Label(inre, text=(f"{ikon}  {rubrik}" if ikon else rubrik), bg=KORT, fg=TEXT, font=self.f_rubrik, anchor='w', justify='left')
        lbl.pack(fill='x', padx=12, pady=(8,4))
        def _wrap(e):
            lbl.config(wraplength=e.width-24)
        inre.bind('<Configure>', _wrap)
        t = tk.Frame(inre, bg=KORT)
        t.pack(fill='both', expand=True, padx=8, pady=(0, 8))
        t._grafisk = True
        t._rubrik = rubrik
        t._farg = farg
        t._inre = inre
        return t

    def stat_text(self, t, text):
        if getattr(t, '_grafisk', False):
            for w in list(t.winfo_children()):
                w.destroy()
            linjer = [l for l in (text or "").split("\n") if l.strip()]
            if not linjer:
                tk.Label(t, text="–", bg=KORT, fg=DAMPAD, font=(FONT, 11)).pack(fill='x', padx=8, pady=4)
                return
            import re as _re
            if getattr(t, '_stapel', False) or "Sökt som" in getattr(t, '_rubrik', '') or "Söklängd" in getattr(t, '_rubrik', ''):
                for i, linje in enumerate(linjer):
                    m = _re.match(r'^(.*?)\s+(\d+)\s+\((\d+)%\)', linje.strip())
                    if not m:
                        tk.Label(t, text=linje, bg=KORT, fg=TEXT, font=(FONT, 11)).pack(fill='x', padx=8, pady=2)
                        continue
                    etikett, antal, pct = m.group(1).strip(), int(m.group(2)), int(m.group(3))
                    rad = tk.Frame(t, bg='#fbfcfe', highlightbackground='#e2e8f0', highlightthickness=1)
                    rad.pack(fill='x', padx=4, pady=4)
                    ovre = tk.Frame(rad, bg='#fbfcfe')
                    ovre.pack(fill='x', padx=12, pady=(10, 4))
                    tk.Label(ovre, text=etikett, bg='#fbfcfe', fg=TEXT, font=(FONT, 13, 'bold'), anchor='w').pack(side='left')
                    tk.Label(ovre, text=f"{pct} %", bg='#fbfcfe', fg=TEXT, font=(FONT, 15, 'bold'), anchor='e').pack(side='right')
                    tk.Label(ovre, text=(f"{antal}" if getattr(t, '_stapel', False) else f"{antal} sök"), bg='#fbfcfe', fg=DAMPAD, font=(FONT, 10), anchor='e').pack(side='right', padx=(0, 12))
                    stapel_f = tk.Frame(rad, bg='#e8edf3', height=10)
                    stapel_f.pack(fill='x', padx=12, pady=(0, 12))
                    stapel_f.pack_propagate(False)
                    stapel_farg = BLA   # en enda färg: färg används bara för att peka ut problem
                    fill = tk.Frame(stapel_f, bg=stapel_farg, height=10)
                    fill.pack(side='left', fill='y')
                    def _satt(e=None, fill=fill, pct=pct, stapel_f=stapel_f):
                        w_tot = stapel_f.winfo_width()
                        if w_tot > 5:
                            fill.config(width=max(2, int(w_tot * pct / 100)))
                    stapel_f.bind('<Configure>', _satt)
                    t.after(120, _satt)
            else:
                for i, linje in enumerate(linjer):
                    m = _re.match(r'^(\d+)\s+(.*)$', linje.strip())
                    antal, q = (m.group(1), m.group(2)) if m else ("", linje.strip())
                    bg = '#fbfcfe' if i % 2 == 0 else '#fdf0f3'
                    rad = tk.Frame(t, bg=bg)
                    rad.pack(fill='x', padx=2, pady=3)
                    tk.Label(rad, text="✕", bg=bg, fg=ROD, font=(FONT_FET, 13, 'bold'), width=2, anchor='center').pack(side='left', padx=(4,2), pady=4)
                    tk.Label(rad, text=q, bg=bg, fg=TEXT, font=(FONT, 12), anchor='w').pack(side='left', fill='x', expand=True, padx=(2,6), pady=4)
                    if antal:
                        badge = tk.Label(rad, text=f"{antal} ggr", bg=ROD, fg='white', font=(FONT, 10, 'bold'), padx=8, pady=2)
                        badge.pack(side='right', padx=(6,8), pady=4)
            return
        t.config(state='normal')
        t.delete('1.0', 'end')
        t.insert('end', text)
        if getattr(t, '_raster', None):
            linjer = text.split('\n')
            for i, _linje in enumerate(linjer):
                tagg = 'rad_udda' if i % 2 == 1 else 'rad_jamn'
                t.tag_add(tagg, f'{i+1}.0', f'{i+1}.end')
        t.config(state='disabled')

    def _hex_to_rgb(self, hexcol):
        h=hexcol.lstrip('#')
        if len(h)==3:
            h=''.join([c*2 for c in h])
        try:
            return tuple(int(h[i:i+2],16) for i in (0,2,4))
        except Exception:
            return (124,176,70)
    def _rgb_to_hex(self, r,g,b):
        return f"#{int(r):02x}{int(g):02x}{int(b):02x}"
    def _fade_color(self, base, factor=0.45):
        r,g,b=self._hex_to_rgb(base)
        r2 = r + (255-r)*factor
        g2 = g + (255-g)*factor
        b2 = b + (255-b)*factor
        return self._rgb_to_hex(r2,g2,b2)
    def _draw_gradient_rect(self, canvas, x0, y0, x1, y1, base_color, vertical=True):
        r0,g0,b0=self._hex_to_rgb(base_color)
        r1,g1,b1=self._hex_to_rgb(self._fade_color(base_color, 0.58))
        steps=16
        if vertical:
            h = y1 - y0
            if h<=0:
                return
            for i in range(steps):
                t=i/steps
                r = r1 + (r0 - r1)*t
                g = g1 + (g0 - g1)*t
                b = b1 + (b0 - b1)*t
                col=self._rgb_to_hex(r,g,b)
                yy0 = y0 + h * t
                yy1 = y0 + h * (t+1/steps)
                canvas.create_rectangle(x0, yy0, x1, yy1+1, fill=col, outline=col, width=0)
        else:
            w = x1 - x0
            if w<=0:
                return
            for i in range(steps):
                t=i/steps
                r = r0 + (r1 - r0)*t
                g = g0 + (g1 - g0)*t
                b = b0 + (b1 - b0)*t
                col=self._rgb_to_hex(r,g,b)
                xx0 = x0 + w * t
                xx1 = x0 + w * (t+1/steps)
                canvas.create_rectangle(xx0, y0, xx1+1, y1, fill=col, outline=col, width=0)

    def _rita_diagram(self, c):
        typ = getattr(c, '_typ', 'stapel')
        if typ == 'stapel':
            self._rita_staplar(c)
        elif typ == 'horisontell':
            self._rita_horisontell(c)
        elif typ == 'donut':
            self._rita_donut(c)
        elif typ == 'linje':
            self._rita_linje(c)
        else:
            self._rita_staplar(c)

    def _rita_staplar(self, c):
        c.delete('all')
        data = c._data
        w, h = c.winfo_width(), c.winfo_height()
        if w < 30:
            w = 300
        if not data or w < 50:
            c.create_text(w//2, h//2, text="Ingen data än", fill=DAMPAD, font=self.f_liten, width=w-20)
            return
        etiketter, varden = data
        n = len(varden)
        if n == 0:
            return
        topp, botten, left, right = 30, 50, 50, 12
        maxv = max(varden) or 1
        for k in range(4):
            y = topp + (h - topp - botten) * k / 3
            c.create_line(left, y, w-right, y, fill='#eef1f5', dash=(2,4))
            val = int(maxv * (1 - k/3))
            c.create_text(left-6, y, text=str(val), fill=DAMPAD, font=(FONT, 9), anchor='e')
        bredd = (w - left - right) / n
        for i, v in enumerate(varden):
            x0 = left + i * bredd + bredd*0.15
            x1 = left + (i+1)*bredd - bredd*0.15
            hojd = (h - topp - botten) * v / maxv if maxv else 0
            y1 = h - botten
            y0 = y1 - hojd
            if v>0:
                self._draw_gradient_rect(c, x0, y0, x1, y1, c._farg, vertical=True)
                c.create_rectangle(x0, y0, x1, y1, outline='white', width=1)
                if bredd > 28 and hojd > 14:
                    c.create_text((x0+x1)/2, y0-10, text=str(v), fill=TEXT, font=(FONT, 9, 'bold'))
        steg = max(1, n//8) if n>12 else 1
        for i in range(0, n, steg):
            txt = str(etiketter[i])[:12]
            c.create_text(left + i*bredd + bredd/2, h - botten + 16, text=txt, fill=DAMPAD, font=(FONT, 9), anchor='n', width=int(bredd+10))
        c.create_line(left, h-botten, w-right, h-botten, fill=RAM, width=1)

    def _rita_horisontell(self, c):
        c.delete('all')
        data = c._data
        w, h = c.winfo_width(), c.winfo_height()
        if w < 30:
            w = 300
        if not data:
            c.create_text(w//2, h//2, text="Ingen data", fill=DAMPAD, font=self.f_liten, width=w-20)
            return
        etiketter, varden = data
        if not varden:
            return
        maxv = max(varden) or 1
        total = sum(varden) or 1
        topp, botten, left = 10, 10, 110
        hog_marginal = 76
        palett = [GRON_MORK, BLA, ROSA, GUL, "#7c3aed", "#ea580c", "#0ea5e9", "#dc2626", "#16a34a", "#0891b2"]
        rad_hojd = max(18, (h - topp - botten) / max(1, len(varden)))
        for i, (lab, v) in enumerate(zip(etiketter, varden)):
            y = topp + i*rad_hojd
            if y > h - botten:
                break
            farg = c._farg   # en färg per diagram
            c.create_text(8, y+rad_hojd/2, text=str(lab)[:18], fill=TEXT, font=(FONT, 10), anchor='w', width=left-16)
            c.create_rectangle(left, y+5, w-hog_marginal, y+rad_hojd-5, fill='#f1f5f9', outline='')
            fill_w = (w-hog_marginal-left) * v / maxv if maxv else 0
            if fill_w>0:
                self._draw_gradient_rect(c, left, y+5, left+fill_w, y+rad_hojd-5, farg, vertical=False)
            c.create_text(w-8, y+rad_hojd/2, text=f"{v} ({v*100//total if total else 0}%)" if total>10 else str(v), fill=TEXT, font=(FONT, 10, 'bold'), anchor='e')

    def _rita_donut(self, c):
        c.delete('all')
        data = c._data
        w, h = c.winfo_width(), c.winfo_height()
        if w < 30:
            w=300
        if not data:
            c.create_text(w//2, h//2, text="Ingen data", fill=DAMPAD, font=self.f_liten, width=w-20)
            return
        etiketter, varden = data
        total = sum(varden) or 1
        legend_h = 42
        h_chart = h - legend_h
        cx, cy = w//2, h_chart//2 + 5
        r_outer = min(w*0.9, h_chart*0.8) * 0.38
        r_inner = r_outer * 0.60
        start = -90
        colors = [GRON_MORK, ROSA, BLA, GUL, "#7c3aed", "#ea580c"]
        for idx, (lab, v) in enumerate(zip(etiketter, varden)):
            extent = 360 * v / total if total else 0
            if extent < 1:
                continue
            base = colors[idx % len(colors)] if len(etiketter)>1 else c._farg
            c.create_arc(cx-r_outer, cy-r_outer, cx+r_outer, cy+r_outer, start=start, extent=extent, fill=base, outline='white', width=2, style='pieslice')
            try:
                light = self._fade_color(base, 0.35)
                c.create_arc(cx-r_outer+4, cy-r_outer+4, cx+r_outer-4, cy+r_outer-4, start=start, extent=extent, fill=light, outline='', style='pieslice')
            except Exception:
                pass
            start += extent
        c.create_oval(cx-r_inner, cy-r_inner, cx+r_inner, cy+r_inner, fill=KORT, outline='white', width=2)
        c.create_text(cx, cy, text=f"{total}\nsök", fill=TEXT, font=(FONT_FET, 12, 'bold'), justify='center')
        ly = h_chart + 14
        lx = 12
        for idx, (lab, v) in enumerate(zip(etiketter, varden)):
            pct = v*100//total if total else 0
            col = colors[idx % len(colors)] if len(etiketter)>1 else c._farg
            c.create_rectangle(lx, ly-8, lx+12, ly+2, fill=col, outline='')
            c.create_text(lx+16, ly, text=f"{lab}: {v} ({pct}%)", fill=TEXT, font=(FONT, 9, 'bold'), anchor='w')
            lx += 140
            if lx > w-110:
                lx = 12
                ly += 18


    def _rita_linje(self, c):
        c.delete('all')
        data = c._data
        w, h = c.winfo_width(), c.winfo_height()
        if w < 30:
            w=300
        if not data:
            c.create_text(w//2, h//2, text="Ingen data", fill=DAMPAD, font=self.f_liten)
            return
        etiketter, varden = data
        if len(varden) < 2:
            self._rita_staplar(c)
            return
        topp, botten, left, right = 20, 40, 40, 12
        maxv = max(varden) or 1
        minv = min(varden)
        rng = maxv - minv or 1
        # grid
        for k in range(4):
            y = topp + (h - topp - botten) * k / 3
            c.create_line(left, y, w-right, y, fill='#eef1f5', dash=(2,4))
            val = int(maxv - k/3*rng)
            c.create_text(left-6, y, text=str(val), fill=DAMPAD, font=(FONT, 8), anchor='e')
        pts = []
        for i, v in enumerate(varden):
            x = left + i * (w-left-right) / max(1, len(varden)-1)
            y = topp + (h - topp - botten) * (1 - (v - minv)/rng)
            pts.append((x,y))
        # fyllning under
        if len(pts)>1:
            poly = pts + [(pts[-1][0], h-botten), (pts[0][0], h-botten)]
            flat = [coord for p in poly for coord in p]
            c.create_polygon(flat, fill=self._fade_color(c._farg, 0.78), outline='')
            for (x1,y1),(x2,y2) in zip(pts, pts[1:]):
                c.create_line(x1,y1,x2,y2, fill=c._farg, width=2.5, smooth=True, capstyle='round')
            for x,y in pts:
                c.create_oval(x-3, y-3, x+3, y+3, fill=c._farg, outline='white', width=1.5)
        # x-etiketter
        steg = max(1, len(etiketter)//8) if len(etiketter)>12 else 1
        for i in range(0, len(etiketter), steg):
            x = left + i * (w-left-right) / max(1, len(etiketter)-1)
            c.create_text(x, h-botten+14, text=str(etiketter[i])[:8], fill=DAMPAD, font=(FONT, 8), anchor='n')

    def _fakta_kort(self, parent, ikon, rubrik, varde, undertext, farg, accent_bg, progress=None, sparkline=None, stor=False):
        ram = tk.Frame(parent, bg=RAM)
        inre = tk.Frame(ram, bg=KORT)
        inre.pack(fill='both', expand=True, padx=1, pady=1)
        tk.Frame(inre, bg=farg, height=5 if stor else 4).pack(fill='x')

        topp = tk.Frame(inre, bg=KORT)
        topp.pack(fill='x', padx=12, pady=(10,2))
        size = 40 if stor else 34
        cv = tk.Canvas(topp, width=size, height=size, bg=KORT, highlightthickness=0)
        cv.create_oval(2,2,size-2,size-2, fill=accent_bg, outline='')
        cv.create_text(size//2, size//2, text=ikon, font=(FONT, 16 if not stor else 18))
        cv.pack(side='left')
        rub_lbl = tk.Label(topp, text=rubrik.upper(), bg=KORT, fg=DAMPAD, font=(FONT, 8 if not stor else 9, 'bold'), anchor='w', justify='left')
        rub_lbl.pack(side='left', padx=(10,0), fill='x', expand=True)
        def _wrap_rub(e):
            rub_lbl.config(wraplength=e.width-10)
        topp.bind('<Configure>', _wrap_rub)

        var_lbl = tk.Label(inre, text=varde, bg=KORT, fg=TEXT, font=(FONT_FET, 13 if not stor else 16, 'bold'), anchor='w', justify='left', wraplength=240)
        var_lbl.pack(fill='x', padx=14, pady=(8,2))
        under_lbl = tk.Label(inre, text="", bg=KORT)
        if (undertext or "").strip():
            if (undertext or "").strip().startswith(("🥇", "🥈", "🥉")):
                farger = {"🥇": "#b8860b", "🥈": "#64748b", "🥉": "#a16207"}
                for rad in undertext.strip().split("\n"):
                    if not rad.strip():
                        continue
                    medalj = rad[0]
                    rest = rad[1:].strip()
                    import re as _re
                    m = _re.match(r'^(.*?)\s+(\d+\s+ggr)$', rest)
                    text_del, antal_del = (m.group(1), m.group(2)) if m else (rest, "")
                    rad_f = tk.Frame(inre, bg=KORT)
                    rad_f.pack(fill='x', padx=14, pady=(2,2))
                    tk.Label(rad_f, text=medalj, bg=KORT, fg=farger.get(medalj, TEXT), font=(FONT_FET, 14, 'bold'), width=2, anchor='w').pack(side='left')
                    tk.Label(rad_f, text=text_del, bg=KORT, fg=TEXT, font=(FONT_FET, 11 if not stor else 12, 'bold'), anchor='w').pack(side='left', fill='x', expand=True, padx=(4,6))
                    if antal_del:
                        tk.Label(rad_f, text=antal_del, bg=KORT, fg=DAMPAD, font=(FONT, 10, 'bold'), anchor='e').pack(side='right')
            elif '|' in (undertext or ""):
                for rad in undertext.strip().split("\n"):
                    if not rad.strip() or '|' not in rad:
                        continue
                    delar = [d.strip() for d in rad.split('|')]
                    namn, antal = delar[0], delar[1]
                    pct = int(delar[2]) if len(delar) > 2 and delar[2].isdigit() else None
                    topp_rad = len(delar) > 3 and delar[3] == '1'
                    rad_f = tk.Frame(inre, bg=KORT)
                    rad_f.pack(fill='x', padx=14, pady=(2, 2))
                    tk.Label(rad_f, text=namn, bg=KORT, fg=TEXT, font=(FONT_FET if topp_rad else FONT, 10 if not stor else 11, 'bold' if topp_rad else 'normal'), anchor='w', width=8).pack(side='left')
                    tk.Label(rad_f, text=antal, bg=KORT, fg=TEXT, font=(FONT, 10 if not stor else 11, 'bold'), anchor='e', width=4).pack(side='right')
                    if pct is not None:
                        sp = tk.Frame(rad_f, bg='#f1f5f9', height=8)
                        sp.pack(side='left', fill='x', expand=True, padx=(6, 8))
                        sp.pack_propagate(False)
                        sf = tk.Frame(sp, bg=farg if topp_rad else self._fade_color(farg, 0.45), height=8)
                        sf.pack(side='left', fill='y')
                        def _sb(e=None, sp=sp, sf=sf, pct=pct):
                            w_tot = sp.winfo_width()
                            if w_tot > 5:
                                sf.config(width=max(2 if pct else 0, int(w_tot * min(100, pct) / 100)))
                        sp.bind('<Configure>', _sb)
            else:
                for rad in undertext.strip().split("\n"):
                    if rad.strip():
                        tk.Label(inre, text=rad, bg=KORT, fg=TEXT, font=(FONT, 10 if not stor else 11), anchor='w', justify='left', wraplength=260).pack(fill='x', padx=14, pady=(1,1))
            tk.Frame(inre, bg=KORT, height=4).pack(fill='x')
        def _wrap_vals(e):
            w = e.width - 28
            var_lbl.config(wraplength=w)
        inre.bind('<Configure>', _wrap_vals)

        if progress is not None:
            p_frame = tk.Frame(inre, bg='#f1f5f9', height=8)
            p_frame.pack(fill='x', padx=12, pady=(4,10))
            p_frame.pack_propagate(False)
            fill = tk.Frame(p_frame, bg=farg, height=8)
            fill.pack(side='left', fill='y')
            def _satt_bredd(e=None):
                try:
                    w_tot = p_frame.winfo_width()
                    if w_tot > 5:
                        fill.config(width=int(w_tot * min(100, progress) / 100))
                except Exception:
                    pass
            p_frame.bind('<Configure>', _satt_bredd)
            inre.after(120, _satt_bredd)

        if sparkline and len(sparkline)>1:
            sh, sw = 38, 240
            sc = tk.Canvas(inre, width=sw, height=sh, bg=KORT, highlightthickness=0)
            sc.pack(fill='x', padx=12, pady=(2,10))
            max_v = max(sparkline) or 1
            min_v = min(sparkline)
            rng = max_v - min_v or 1
            pts = []
            for i, v in enumerate(sparkline):
                x = 6 + i * (sw-12) / (len(sparkline)-1)
                y = sh - 6 - (v - min_v) / rng * (sh-12)
                pts.append((x,y))
            poly = pts + [(pts[-1][0], sh), (pts[0][0], sh)]
            flat = [c for p in poly for c in p]
            sc.create_polygon(flat, fill=accent_bg, outline='')
            for (x1,y1),(x2,y2) in zip(pts, pts[1:]):
                sc.create_line(x1,y1,x2,y2, fill=farg, width=2, smooth=True)
            sc.create_oval(pts[-1][0]-3, pts[-1][1]-3, pts[-1][0]+3, pts[-1][1]+3, fill=farg, outline='white', width=2)

        return ram

    def bygg_nyttig_fakta(self, ev, tider, utan, per_dag, per_tim, vd, dagar, fore, mest_sokta, utan_counter, enh_list):
        from collections import Counter
        import re as _re
        kort = []
        if not ev:
            return kort
        total = len(ev)
        n_utan = len(utan)
        traff_pct = 100 - n_utan*100//total if total else 0

        # 1. Träffsäkerhet hero
        spark = [per_dag.get(d,0) for d in sorted(per_dag)[-14:]]
        pass  # Träffsäkerhet visas redan i toppkorten
        # --- NY RUTA 1: Artiklar vs Instrumentgaller, exakt via träffkälla (fältet s från index.html) ---
        try:
            def _norm_kalla(e):
                if isinstance(e, dict):
                    v = str(e.get('s') or e.get('kalla') or '').strip().lower()
                    return v if v in ('artikel','galler','båda','utan') else None
                if len(e) > 3 and isinstance(e[3], str) and e[3].strip().lower() in ('artikel','galler','båda','utan'):
                    return e[3].strip().lower()
                if len(e) > 9 and isinstance(e[9], str) and e[9].strip().lower() in ('artikel','galler','båda','utan'):
                    return e[9].strip().lower()
                return None
            kalla = [_norm_kalla(e) for e in ev]
            n_med_kalla = sum(1 for k in kalla if k)
            if n_med_kalla * 2 >= total:
                n_art = sum(1 for k in kalla if k == "artikel")
                n_galler = sum(1 for k in kalla if k == "galler")
                n_bada = sum(1 for k in kalla if k == "båda")
                n_utan_k = sum(1 for k in kalla if k == "utan")
                kort.append(dict(ikon="🧰", rubrik="Artiklar vs Instrumentgaller", varde=f"Artiklar {n_art*100//total if total else 0}%", undertext="\n".join([f"Artiklar {n_art} sök", f"Galler {n_galler} sök", f"Båda {n_bada} sök"] + ([f"Utan träff {n_utan_k} sök"] if n_utan_k else [])), farg=BLA, accent="#dbeaf2", progress=n_art*100//total if total else 0, stor=False))
            else:
                import re as _re2
                def _ar_galler(q):
                    ql = str(q).strip().lower()
                    return 'galler' in ql or bool(_re2.match(r'^[a-zåäö]+\s*\d', ql))
                n_galler = len([e for e in ev if _ar_galler(e[1] if not isinstance(e, dict) else e.get('q',''))])
                n_art = total - n_galler
                kort.append(dict(ikon="🧰", rubrik="Artiklar vs Instrumentgaller", varde=f"Artiklar {n_art*100//total if total else 0}% *", undertext=f"Artiklar {n_art} sök\nInstrumentgaller {n_galler} sök (* uppskattat)", farg=BLA, accent="#dbeaf2", progress=n_art*100//total if total else 0, stor=False))
        except Exception:
            pass
        # --- NY RUTA 2: Utanför arbetstid (21-07) ---
        try:
            utan_arb = [x for x in tider if x.hour >= 21 or x.hour < 7]
            n_utan_arb = len(utan_arb)
            pct_utan_arb = n_utan_arb*100//total if total else 0
            if n_utan_arb:
                from collections import Counter as _C2
                kh = _C2(x.hour for x in utan_arb)
                topp_h, topp_n = kh.most_common(1)[0]
                under = f"Flest kl {topp_h:02d}:00 ({topp_n} sök). Kväll/natt kan tyda på behov av bättre ljus eller skyltning."
            else:
                under = "Inga sökningar mellan 21 och 07 under perioden."
            kort.append(dict(ikon="🌙", rubrik="Utanför arbetstid (21-07)", varde=f"{n_utan_arb} sök ({pct_utan_arb}%)", undertext=under, farg=BLA, accent="#dbeaf2", progress=None, stor=False))
        except Exception:
            pass

        # 2. Mest aktiv dag
        snitt = total / max(1, len(per_dag))
        topp_dag, n_topp = per_dag.most_common(1)[0]
        pass  # borttagen: visas redan i diagram

        # 3. Timme
        topp_tim, n_tim = per_tim.most_common(1)[0]
        lunch = sum(per_tim.get(h,0) for h in (11,12,13))
        pass  # borttagen: visas redan i diagram

        # 4. Veckodag bemanning
        v_namn = ['Måndag','Tisdag','Onsdag','Torsdag','Fredag','Lördag','Söndag']
        topp_v, n_v = vd.most_common(1)[0]
        pass  # borttagen: visas redan i diagram

        # 5. Mest sökta artiklar - åtgärd
        if mest_sokta:
            topp3 = mest_sokta.most_common(3)
            topp_q, topp_n = topp3[0]
            medaljer = ["🥇", "🥈", "🥉"]
            under3 = "\n".join(f"{medaljer[i]} \"{q[:22]}\" {n} ggr" for i, (q, n) in enumerate(topp3))
            pass  # borttagen: visas redan i diagram

        # 6. Utan träff - prioritering
        if utan_counter:
            topp_u, n_u = utan_counter.most_common(1)[0]
            pass  # borttagen: visas redan i diagram

        # 7. Variation / återkommande
        unika = len({str(e[1]).strip().lower() for e in ev})
        kort.append(dict(ikon="🔁", rubrik="Återkommande sökningar", varde=f"{unika} unika av {total}", undertext=f"Varje sökning upprepas i snitt {total/max(1,unika):.1f} ggr. Bra underlag för favoritlista.", farg=BLA, accent="#dbeaf2", progress=None, stor=False))

        # 8. Kameraeffektivitet
        kam_sok = [e for e in ev if len(e)>8 and e[8]==1]
        if kam_sok:
            kam_traff = len([e for e in kam_sok if e[2]])
            pct = kam_traff*100//len(kam_sok) if kam_sok else 0
            pass  # kamera visas i kamerarutan

        # 9. Enhet
        if enh_list:
            def _cap_enh(s):
                s=str(s).strip()
                return s[:1].upper()+s[1:].lower() if s else 'Okänd'
            c_enh = Counter(_cap_enh(e[4]) for e in enh_list if len(e)>4 and e[4])
            c_enh = Counter({('Mobil' if k.lower()=='mobil' else 'Dator' if k.lower()=='dator' else k): v for k,v in c_enh.items()})
            if c_enh:
                topp_enh, n_enh = c_enh.most_common(1)[0]
                # Bygg undertext med fördelning: t.ex. Dator 187 (69%), Mobil 84 (31%)
                total_enh = len(enh_list) or 1
                fordelning = ", ".join(f"{k} {v} ({v*100//total_enh}%)" for k,v in c_enh.most_common())
                pass  # borttagen: visas redan i diagram

        # 10. Trend
        if fore>0:
            diff = (total - fore)*100//fore
            pil = "↗" if diff>=0 else "↘"
            kort.append(dict(ikon="📈", rubrik="Trend vs föreg. period", varde=f"{pil} {'+' if diff>=0 else ''}{diff}%", undertext=f"{total} nu mot {fore} innan. Ökning = fler sök = behov av bättre datakvalitet.", farg=GRON_MORK if diff>=0 else ROD, accent="#dcfce7" if diff>=0 else "#fee2e2", progress=50+max(-50,min(50,diff)), sparkline=[fore, total], stor=False))

        # 11. Svit dagar
        dagar_med = sorted(per_dag)
        bast = svit = 1
        for a,b in zip(dagar_med, dagar_med[1:]):
            svit = svit+1 if (b-a).days==1 else 1
            bast = max(bast, svit)
        kort.append(dict(ikon="🔥", rubrik="Kontinuitet", varde=f"{bast} dagar i rad med sök", undertext=f"Aktivt använd varje arbetsdag? Längsta svit visar adoption.", farg=BLA, accent="#dbeaf2", progress=None, stor=False))

        # EXTRA 2 nya för komplett rad
        try:
            avg_len = sum(len(str(e[1])) for e in ev) / max(1,len(ev))
            pass  # Söklängd visas i rutan Så söker man
        except Exception:
            pass
        try:
            total_dagar_period = dagar
            aktiva_dagar = len(per_dag)
            pct_aktiva = aktiva_dagar*100//max(1,total_dagar_period)
            kort.append(dict(ikon="📊", rubrik="Aktiva dagar", varde=f"{aktiva_dagar} av {total_dagar_period} dagar", undertext=f"{pct_aktiva}% av perioden har sökningar. Visar hur ofta Lagersök används.", farg=BLA, accent="#dbeaf2", progress=None, stor=False))
        except Exception:
            pass

        return kort

    def hamta_statistik(self):
        url = self.stat_url.get().strip().rstrip('/')
        nyckel = self.stat_key.get().strip()
        if not url or not nyckel:
            try:
                self.stat_status.config(text="Fyll i adress och nyckel först.", fg=ROD)
            except Exception:
                pass
            return
        try:
            self.stat_status.config(text="Hämtar…", fg=DAMPAD)
        except Exception:
            pass
        dagar = self.stat_dagar.get() or '30'
        def jobb():
            import json as _json
            import urllib.request
            try:
                req = urllib.request.Request(f"{url}/stats?dagar={int(dagar)*2}", headers={'X-Admin-Key': nyckel, 'User-Agent': 'lagersok-uppdatering'})
                with urllib.request.urlopen(req, timeout=30) as r:
                    data = _json.loads(r.read().decode('utf-8'))
                self.ko.put(('stat', data, None))
            except Exception as e:
                self.ko.put(('stat', None, str(e)))
        import threading
        threading.Thread(target=jobb, daemon=True).start()

    def h_stat(self, data, fel):
        if fel or data is None:
            try:
                self.stat_status.config(text="✗ Kunde inte hämta: " + str(fel), fg=ROD)
            except Exception:
                pass
            return
        try:
            self.spara(tyst=True)
        except Exception:
            pass
        from collections import Counter
        import datetime as _dt
        ev = data.get('events', [])
        dagar = int(self.stat_dagar.get() or 30)
        grans0 = (_dt.datetime.now() - _dt.timedelta(days=dagar)).timestamp()*1000
        oppn = [e for e in ev if len(e)>8 and e[8]==2 and e[0]>=grans0]
        ev_alla = [e for e in ev if not (len(e)>8 and e[8]==2)]
        ev = [e for e in ev_alla if e[0]>=grans0]
        fore_antal = len(ev_alla) - len(ev)
        if not ev:
            try:
                self.stat_status.config(text="Inga sökningar för perioden.", fg=DAMPAD)
            except Exception:
                pass
        else:
            try:
                self.stat_status.config(text=f"✓ {len(ev)} sökningar hämtade ({_dt.datetime.now().strftime('%H:%M')})", fg=GRON_MORK)
            except Exception:
                pass
        tider = [_dt.datetime.fromtimestamp(e[0]/1000.0) for e in ev]
        utan = [e for e in ev if not e[2]]
        mest = Counter(str(e[1]).strip().lower() for e in ev if str(e[1]).strip())
        utanr = Counter(str(e[1]).strip().lower() for e in utan if str(e[1]).strip())
        per_dag = Counter(t.date() for t in tider)
        per_tim = Counter(t.hour for t in tider)
        idag = _dt.date.today()
        dagar_lista = [idag - _dt.timedelta(days=i) for i in range(min(dagar,60)-1, -1, -1)]
        # diagram data
        self.stat_c_dag._data = ([d.strftime('%d/%m') for d in dagar_lista], [per_dag.get(d,0) for d in dagar_lista])
        self.stat_c_tim._data = ([f"{h:02d}" for h in range(24)], [per_tim.get(h,0) for h in range(24)])
        vd = Counter(t.weekday() for t in tider)
        self.stat_c_vdag._data = (['Måndag','Tisdag','Onsdag','Torsdag','Fredag','Lördag','Söndag'], [vd.get(i,0) for i in range(7)])
        # träff vs utan
        self.stat_c_traff._data = (['Träff','Utan träff'], [len(ev)-len(utan), len(utan)])
        # enhet
        har_enh = [e for e in ev if len(e)>4 and e[4]]
        if har_enh:
            def _cap2(s):
                s=str(s).strip()
                return s[:1].upper()+s[1:].lower() if s else 'Okänd'
            enh_c = Counter(_cap2(e[4]) for e in har_enh)
            enh_c = Counter({('Mobil' if k.lower()=='mobil' else 'Dator' if k.lower()=='dator' else k): v for k,v in enh_c.items()})
            self.stat_c_enhet._data = (list(enh_c.keys())[:6], [enh_c[k] for k in list(enh_c.keys())[:6]])
        else:
            self.stat_c_enhet._data = (['Ingen data'], [0])
        # kamera vs text
        kam = [e for e in ev if len(e)>8 and e[8]==1]
        text_sok = len(ev)-len(kam)
        self.stat_c_kamera._data = (['Text','Kamera'], [text_sok, len(kam)])
        try:
            mest_top = mest.most_common(10)
            self.stat_c_mest._data = ([q[:22] for q,_ in mest_top], [n for _,n in mest_top])
        except Exception:
            self.stat_c_mest._data = (['Ingen data'], [0])
        try:
            utan_top = utanr.most_common(10)
            self.stat_c_utan._data = ([q[:22] for q,_ in utan_top], [n for _,n in utan_top])
        except Exception:
            self.stat_c_utan._data = (['Ingen data'], [0])
        # rita alla - inkluderar nya topp10 och med try
        for c in (self.stat_c_dag, self.stat_c_tim, self.stat_c_vdag, self.stat_c_traff, self.stat_c_enhet, self.stat_c_kamera, self.stat_c_mest, self.stat_c_utan):
            try:
                self._rita_diagram(c)
            except Exception as ex:
                print("rita fel", ex)
        def _cap_st(s):
            s=str(s).strip()
            return s[:1].upper()+s[1:].lower() if s else 'Okänd'
        def stapel(raknare, tot):
            rader=[]
            for namn,n in raknare.most_common(8):
                namn_c=_cap_st(namn)
                if namn_c.lower().startswith('nej, i webbläsare'):
                    namn_c = 'Nej, webbläsare'
                if namn_c.lower().startswith('ja, som app'):
                    namn_c = 'Ja, app'
                andel=n*100//tot if tot else 0
                rader.append(f"{namn_c[:20]}  {n} ({andel}%)")
            return "\n".join(rader) or "–"
        if har_enh:
            tot=len(har_enh)
            self.stat_text(self.stat_enh[0], stapel(Counter(_norm_stat('enhet', e[4]) for e in har_enh), tot))
            self.stat_text(self.stat_enh[1], stapel(Counter(_norm_stat('os', e[5]) for e in har_enh), tot))
            self.stat_text(self.stat_enh[2], stapel(Counter('Ja, som app' if e[6] else 'Nej, i webbläsare' for e in har_enh if len(e)>6), tot))
        else:
            for t in self.stat_enh:
                self.stat_text(t, "Ingen enhetsdata än.")
        self.stat_kam[0].config(text=str(len(oppn)))
        self.stat_kam[1].config(text=str(len(kam)))
        self.stat_kam[2].config(text=(f"{len(kam)*100//len(ev)} %" if ev else "–"))
        # Lyckad = kamerasökningar / antal gånger kameran öppnades. (En kamerasökning loggas bara när ett nummer redan hittats i databasen,
        # så "andel med träff" skulle alltid bli 100 %.) Högst 100 % ifall äldre data saknar "kameran öppnades".
        self.stat_kam[3].config(text=(f"{min(100, len(kam)*100//len(oppn))} %" if oppn else "–"))
        kam_fakta=""
        if kam:
            kh=Counter(_dt.datetime.fromtimestamp(e[0]/1000.0).hour for e in kam)
            kd=Counter(_norm_stat('enhet', e[4]) for e in kam if len(e)>4 and e[4])
            kt=100 - len([e for e in kam if not e[2]])*100//len(kam) if kam else 0
            kam_fakta=f"Kamera används mest kl {kh.most_common(1)[0][0]:02d}. Kameran öppnades {len(oppn)} gånger och gav {len(kam)} sökningar."
        try:
            mest = Counter(str(e[1]).strip().lower() for e in ev if str(e[1]).strip())
            utanr = Counter(str(e[1]).strip().lower() for e in utan if str(e[1]).strip())
            fakta_kort=self.bygg_nyttig_fakta(ev, tider, utan, per_dag, per_tim, vd, dagar, fore_antal, mest, utanr, har_enh)
            self._visa_fakta(fakta_kort, kam_fakta)
        except Exception as e:
            import traceback
            traceback.print_exc()
        aktiva=len(per_dag) or 1
        mest_tim=per_tim.most_common(1)
        self.stat_tiles[0].config(text=str(len(ev)))
        self.stat_tiles[1].config(text=f"{len(ev)/max(1,min(dagar,aktiva)):.1f}".replace('.',','))
        self.stat_tiles[2].config(text=f"{(100 - len(utan)*100//len(ev)) if ev else 0} %")
        self.stat_tiles[3].config(text=(f"{len(utan)} ({len(utan)*100//len(ev)} %)" if ev else "–"))
        import re as _re_num
        def _ar_nummer(q):
            q2 = str(q).strip()
            return bool(_re_num.match(r'^\d+$', q2)) or bool(_re_num.match(r'^\d+\s', q2)) or bool(_re_num.match(r'^[a-zåäö]+\s*\d', q2.lower()))
        n_num = sum(n for q,n in mest.most_common(100) if _ar_nummer(q))
        n_ord = sum(n for q,n in mest.most_common(100) if not _ar_nummer(q))
        tot_m = n_num + n_ord or 1
        self.stat_text(self.stat_t_mest, (f"Artikelnummer  {n_num}  ({n_num*100//tot_m}%)\nOrd  {n_ord}  ({n_ord*100//tot_m}%)") if (n_num + n_ord) else "Inga sökningar.")
        from collections import Counter as _C_l
        langd = _C_l()
        for q,n in mest.most_common(100):
            ord_antal = len(str(q).strip().split())
            langd['1 ord' if ord_antal==1 else '2 ord' if ord_antal==2 else '3+ ord'] += n
        tot_l = sum(langd.values()) or 1
        soklangd_text = "\n".join(f"{etikett}  {langd.get(etikett,0)}  ({langd.get(etikett,0)*100//tot_l}%)" for etikett in ['1 ord','2 ord','3+ ord'] if langd.get(etikett,0)>0)
        self.stat_text(self.stat_t_utan, soklangd_text or "Inga sökningar.")


        # ===================== AVVIKELSERFLIK =====================


    def bygg_avvikelser_flik(self, parent):
        topp = tk.Frame(parent, bg=BAKGRUND)
        topp.pack(fill='x', padx=18, pady=(12,6))

        tk.Label(topp, text="Avvikelser uppdelat per kategori", bg=BAKGRUND, fg=TEXT, font=self.f_rubrik_stor).pack(side='left')
        self.avv_sum_label = tk.Label(topp, text="Ingen körning än", bg=BAKGRUND, fg=DAMPAD, font=self.f_liten)
        self.avv_sum_label.pack(side='left', padx=(12,0))

        btn_rad = tk.Frame(topp, bg=BAKGRUND)
        btn_rad.pack(side='right')
        Knapp(btn_rad, "Öppna avvikelser.txt", self.oppna_avv_fil, '#e5e9f0', TEXT, '#d5dbe5').pack(side='left', padx=4)
        Knapp(btn_rad, "Ladda fil…", self.ladda_avv_fil, '#e5e9f0', TEXT, '#d5dbe5').pack(side='left', padx=4)
        Knapp(btn_rad, "Rensa", self.rensa_avv, '#e5e9f0', TEXT, '#d5dbe5').pack(side='left', padx=4)

        # Sök + Kopiera rad
        sok_rad = tk.Frame(parent, bg=BAKGRUND)
        sok_rad.pack(fill='x', padx=18, pady=(0,8))
        tk.Label(sok_rad, text="🔍 Sök:", bg=BAKGRUND, fg=DAMPAD, font=self.f_liten).pack(side='left')
        self.avv_sok_var = tk.StringVar()
        self.avv_sok_var.trace_add('write', lambda *a: self.filtrera_avvikelser())
        ent = ttk.Entry(sok_rad, textvariable=self.avv_sok_var, font=self.f_norm, width=30)
        ent.pack(side='left', padx=(6,10), ipady=2)
        ent.insert(0, "")
        Knapp(sok_rad, "Rensa sökning", lambda: self.avv_sok_var.set(""), '#e5e9f0', TEXT, '#d5dbe5').pack(side='left', padx=2)
        Knapp(sok_rad, "📋 Kopiera filtrerade", self.kopiera_filtrerade_avv, '#e5e9f0', TEXT, '#d5dbe5').pack(side='left', padx=6)
        Knapp(sok_rad, "📋 Kopiera alla", self.kopiera_alla_avv, '#e5e9f0', TEXT, '#d5dbe5').pack(side='left', padx=2)
        self.avv_filter_label = tk.Label(sok_rad, text="", bg=BAKGRUND, fg=BLA, font=self.f_liten)
        self.avv_filter_label.pack(side='left', padx=10)

        # tre kolumner - nu i scrollable wrapper för små skärmar
        avv_scroll = ScrollableFrame(parent, bg=BAKGRUND, visa_scroll=False)
        avv_scroll.pack(fill='both', expand=True, padx=0, pady=0)
        kol_frame = tk.Frame(avv_scroll.scrollable_frame, bg=BAKGRUND)
        kol_frame.pack(fill='both', expand=True, padx=18, pady=(6,14))
        for i in range(3):
            kol_frame.columnconfigure(i, weight=1, uniform='avv')
        kol_frame.rowconfigure(0, weight=1)

        self.avv_widgets = {}
        kategorier = [
            ("fabrikssterilt", "Fabrikssterilt", GRON, "Dubbletter, saknade fält, ogiltiga nr"),
            ("instrumentgaller", "Instrumentgaller", BLA, "Tomma galler, namnkonflikter"),
            ("bulk", "Bulk", ROSA, "Bulk utan match, dubbla platser"),
        ]

        for idx, (key, titel, farg, under) in enumerate(kategorier):
            yttre = tk.Frame(kol_frame, bg=RAM)
            yttre.grid(row=0, column=idx, sticky='nsew', padx=(0 if idx==0 else 8, 0))
            inre = tk.Frame(yttre, bg=KORT)
            inre.pack(fill='both', expand=True, padx=1, pady=1)

            header = tk.Frame(inre, bg=KORT)
            header.pack(fill='x', padx=12, pady=(10,6))
            tk.Frame(header, bg=farg, height=4).pack(fill='x', pady=(0,6))
            tk.Label(header, text=titel, bg=KORT, fg=TEXT, font=self.f_rubrik).pack(anchor='w')
            tk.Label(header, text=under, bg=KORT, fg=DAMPAD, font=self.f_liten).pack(anchor='w')
            count_lbl = tk.Label(header, text="0 avvikelser", bg=KORT, fg=farg, font=self.f_fet)
            count_lbl.pack(anchor='w', pady=(4,0))

            txt_frame = tk.Frame(inre, bg=KORT)
            txt_frame.pack(fill='both', expand=True, padx=8, pady=(0,8))
            txt = tk.Text(txt_frame, bg='#fbfcfe', fg=TEXT, font=self.f_mono, wrap='word',
                          relief='flat', bd=0, highlightthickness=0, padx=8, pady=6, height=18)
            sb = ttk.Scrollbar(txt_frame, command=txt.yview)
            txt.configure(yscrollcommand=sb.set)
            sb.pack(side='right', fill='y')
            txt.pack(side='left', fill='both', expand=True)
            txt.insert('1.0', f"Inga avvikelser för {titel.lower()} än.\nKör en torrkörning eller uppdatering för att se resultat.")
            txt.config(state='disabled')

            self.avv_widgets[key] = (txt, count_lbl)

        # Utan hyllplats - egen tystare grupp, räknas inte i flikens siffra
        up_frame = tk.Frame(parent, bg=BAKGRUND)
        up_frame.pack(fill='x', padx=18, pady=(0, 6))
        self.up_lank = tk.Label(up_frame, text="", bg=BAKGRUND, fg=BLA, font=(FONT, 10, 'underline'), cursor='hand2', anchor='w')
        self.up_lank.pack(fill='x')
        self.up_lank.bind('<Button-1>', lambda e: self.vaxla_utanplats())
        self.up_text = tk.Text(up_frame, bg='#fbfcfe', fg=TEXT, font=self.f_mono, wrap='word', relief='flat', bd=0,
                               highlightthickness=1, highlightbackground=RAM, padx=8, pady=6, height=7, state='disabled')
        self.up_synlig = False

        # Övrigt längst ner (om något inte kategoriseras)
        ovr_frame = tk.Frame(parent, bg=BAKGRUND)
        ovr_frame.pack(fill='x', padx=18, pady=(0,12))
        self.ovr_label = tk.Label(ovr_frame, text="", bg=BAKGRUND, fg=DAMPAD, font=self.f_liten, anchor='w', justify='left')
        self.ovr_label.pack(fill='x')

    def kategorisera_avvikelser(self, avv_lista):
        """Delar upp avvikelser efter det prefix motorn själv ger varje rad:
        'PDF ...' / 'Artikelfil ...' = fabrikssterilt, 'Galler: ...' = instrumentgaller,
        'Bulk: ...' / 'Bulkfil ...' = bulk. Allt annat hamnar under Övrigt."""
        buckets = {"fabrikssterilt": [], "instrumentgaller": [], "bulk": [], "utanplats": [], "ovrigt": []}
        for item in avv_lista:
            s = str(item).strip()
            if not s or s.startswith("Körning "):
                continue
            low = s.lower()
            if low.startswith("pdf utan hyllplats"):
                buckets["utanplats"].append(s)
            elif low.startswith(("bulk:", "bulkfil")):
                buckets["bulk"].append(s)
            elif low.startswith("galler"):
                buckets["instrumentgaller"].append(s)
            elif low.startswith(("pdf", "artikelfil")):
                buckets["fabrikssterilt"].append(s)
            else:
                buckets["ovrigt"].append(s)
        return buckets

    def visa_avvikelser_i_flik(self, avv_lista):
        buckets = self.kategorisera_avvikelser(avv_lista)
        self.avv_data = buckets
        total = sum(len(v) for k, v in buckets.items() if k != 'utanplats')
        n_utan = len(buckets['utanplats'])
        self.avv_sum_label.config(text=(f"Totalt {total} avvikelser" if total else "Inga avvikelser") + (f"  (+ {n_utan} utan hyllplats)" if n_utan else ""))
        # Uppdatera flik-räknare
        try:
            if total > 0:
                self.notebook.tab(self._flik_avv, text=f"  AVVIKELSER ({total})  ")
            else:
                self.notebook.tab(self._flik_avv, text="  AVVIKELSER  ")
        except Exception:
            pass

        for key in ["fabrikssterilt", "instrumentgaller", "bulk"]:
            txt, cnt_lbl = self.avv_widgets[key]
            items = buckets[key]
            txt.config(state='normal')
            txt.delete('1.0', 'end')
            if items:
                txt.insert('end', "\n\n".join(f"• {it}" for it in items))
                cnt_lbl.config(text=f"{len(items)} avvikelser")
            else:
                txt.insert('end', f"✓ Inga avvikelser för denna kategori.")
                cnt_lbl.config(text="0 avvikelser")
            txt.config(state='disabled')

        self._fyll_utanplats()
        ovr = buckets["ovrigt"]
        if ovr:
            self.ovr_label.config(text=f"Övrigt ({len(ovr)}): " + " | ".join(ovr[:10]))
        else:
            self.ovr_label.config(text="")

        # Växla till avvikelser-fliken automatiskt om det finns avvikelser
        if total > 0:
            # markera fliken visuellt via statusrad
            self.status_rad(f"{total} avvikelser uppdelade i fliken Avvikelser", GUL)

    def vaxla_utanplats(self):
        self.up_synlig = not self.up_synlig
        self._fyll_utanplats()

    def _fyll_utanplats(self):
        """Etiketter utan hyllplats (t.ex. PASSIV eller operationssalskoder): visas kompakt och dolda som standard."""
        items = self.avv_data.get("utanplats", [])
        query = (self.avv_sok_var.get() if hasattr(self, 'avv_sok_var') else "").strip().lower()
        if query:
            items = [it for it in items if query in it.lower()]
        if not self.avv_data.get("utanplats"):
            self.up_lank.config(text="")
            self.up_text.pack_forget()
            return
        pil = "▾" if self.up_synlig else "▸"
        self.up_lank.config(text=f"{pil} Utan hyllplats ({len(items)}): etiketter med t.ex. PASSIV eller kod i stället för plats - hoppades över, ingen åtgärd behövs normalt")
        self.up_text.config(state='normal')
        self.up_text.delete('1.0', 'end')
        self.up_text.insert('end', "\n".join("• " + it for it in items) if items else "Inga träffar.")
        self.up_text.config(state='disabled')
        if self.up_synlig:
            self.up_text.pack(fill='x', pady=(4, 0))
        else:
            self.up_text.pack_forget()

    def oppna_avv_fil(self):
        ut = self.p_ut.get().strip().strip('"')
        if not ut:
            messagebox.showinfo("Ingen sökväg", "Välj var data.js sparas först. avvikelser.txt ligger bredvid.", parent=self.rot)
            return
        base = os.path.dirname(ut) if os.path.isfile(ut) or ut.endswith('.js') else ut
        fil = os.path.join(base, 'avvikelser.txt')
        if os.path.isfile(fil):
            oppna(fil)
        else:
            # prova bredvid basmapp
            fil2 = os.path.join(U.basmapp(), 'avvikelser.txt')
            if os.path.isfile(fil2):
                oppna(fil2)
            else:
                messagebox.showinfo("Hittar inte", f"Hittade ingen avvikelser.txt i {base}", parent=self.rot)

    def ladda_avv_fil(self):
        fil = filedialog.askopenfilename(parent=self.rot, title="Välj avvikelser.txt eller annan textfil",
                                         filetypes=[('Text', '*.txt'), ('Alla', '*.*')])
        if not fil:
            return
        try:
            with open(fil, 'r', encoding='utf-8', errors='ignore') as f:
                rader = [r.strip() for r in f if r.strip()]
            self.visa_avvikelser_i_flik(rader)
            self.notebook.select(self._flik_avv)
        except Exception as e:
            messagebox.showerror("Kunde inte läsa", str(e), parent=self.rot)

    def rensa_avv(self):
        self.avv_data = {"fabrikssterilt": [], "instrumentgaller": [], "bulk": [], "utanplats": [], "ovrigt": []}
        for key in self.avv_widgets:
            txt, cnt = self.avv_widgets[key]
            txt.config(state='normal')
            txt.delete('1.0','end')
            txt.insert('end', "Rensad.")
            txt.config(state='disabled')
            cnt.config(text="0 avvikelser")
        self.avv_sum_label.config(text="Ingen körning än")
        self.ovr_label.config(text="")
        self._fyll_utanplats()
        try:
            self.notebook.tab(self._flik_avv, text="  AVVIKELSER  ")
        except Exception:
            pass

    def filtrera_avvikelser(self):
        """Sök & filtrera avvikelser"""
        if not hasattr(self, 'avv_data'):
            return
        query = (self.avv_sok_var.get() if hasattr(self, 'avv_sok_var') else "").strip().lower()
        if not hasattr(self, 'avv_widgets'):
            return
        total_visade = 0
        for key in ["fabrikssterilt", "instrumentgaller", "bulk"]:
            txt, cnt_lbl = self.avv_widgets.get(key, (None, None))
            if not txt:
                continue
            items = self.avv_data.get(key, [])
            if query:
                filtrerade = [it for it in items if query in it.lower()]
            else:
                filtrerade = items
            txt.config(state='normal')
            txt.delete('1.0', 'end')
            if filtrerade:
                txt.insert('end', "\n\n".join(f"• {it}" for it in filtrerade))
            else:
                if query:
                    txt.insert('end', f"Inga träffar för '{query}' i denna kategori.")
                else:
                    txt.insert('end', "✓ Inga avvikelser för denna kategori.")
            txt.config(state='disabled')
            if query:
                cnt_lbl.config(text=f"{len(filtrerade)}/{len(items)}")
            else:
                cnt_lbl.config(text=f"{len(items)} avvikelser" if items else "0 avvikelser")
            total_visade += len(filtrerade)
        self._fyll_utanplats()
        if hasattr(self, 'avv_filter_label'):
            if query:
                self.avv_filter_label.config(text=f"Visar {total_visade} träffar för '{query}'")
            else:
                self.avv_filter_label.config(text="")

    def kopiera_alla_avv(self):
        if not hasattr(self, 'avv_data'):
            return
        alla = []
        for key in ["fabrikssterilt", "instrumentgaller", "bulk"]:
            for it in self.avv_data.get(key, []):
                alla.append(it)
        if not alla:
            messagebox.showinfo("Inget att kopiera", "Inga avvikelser att kopiera.", parent=self.rot)
            return
        text = "\n".join(alla)
        self.rot.clipboard_clear()
        self.rot.clipboard_append(text)
        self.status_rad(f"{len(alla)} avvikelser kopierade till urklipp", GRON_MORK)

    def kopiera_filtrerade_avv(self):
        if not hasattr(self, 'avv_data'):
            return
        query = (self.avv_sok_var.get() if hasattr(self, 'avv_sok_var') else "").strip().lower()
        alla = []
        for key in ["fabrikssterilt", "instrumentgaller", "bulk"]:
            items = self.avv_data.get(key, [])
            if query:
                filtrerade = [it for it in items if query in it.lower()]
            else:
                filtrerade = items
            alla.extend(filtrerade)
        if not alla:
            messagebox.showinfo("Inget att kopiera", "Inga filtrerade avvikelser.", parent=self.rot)
            return
        text = "\n".join(alla)
        self.rot.clipboard_clear()
        self.rot.clipboard_append(text)
        self.status_rad(f"{len(alla)} filtrerade avvikelser kopierade", GRON_MORK)

        # ===================== STÄDNINGFLIK - RAD-VIS, QR, VECKOVIS =====================
    def _stadlogg_fil(self):
        p = self.p_stadlogg.get().strip() if hasattr(self, 'p_stadlogg') else ""
        if p:
            return p
        try:
            return os.path.join(_get_base_dir(), "stadlogg.json")
        except Exception:
            return "stadlogg.json"

    def _stad_rader_fil(self):
        # lista över rader som ska städas, t.ex. E, Z, BG
        # kan ligga bredvid stadlogg.json som stad_rader.txt
        base = os.path.dirname(self._stadlogg_fil())
        return os.path.join(base, "stad_rader.txt") if base else "stad_rader.txt"

    def _las_stadlogg(self):
        fil = self._stadlogg_fil()
        if os.path.isfile(fil):
            try:
                with open(fil, "r", encoding="utf-8") as fh:
                    data = json.load(fh)
                    return data if isinstance(data, dict) else {}
            except Exception:
                return {}
        return {}

    def _spara_stadlogg(self, data):
        fil = self._stadlogg_fil()
        try:
            os.makedirs(os.path.dirname(fil), exist_ok=True)
            tmp = fil + ".tmp"
            with open(tmp, "w", encoding="utf-8") as fh:
                json.dump(data, fh, ensure_ascii=False, indent=2)
            os.replace(tmp, fil)
            return True
        except Exception as e:
            messagebox.showerror("Kunde inte spara", f"Kunde inte spara {fil}: {e}", parent=self.rot)
            return False

    def _hamta_rader(self):
        # Prioritet: 1) fil stad_rader.txt 2) inställning stad_rader 3) default
        # Filformat: en rad per rad, t.ex. E, Z, BG
        fil = self._stad_rader_fil()
        rader = []
        if os.path.isfile(fil):
            try:
                with open(fil, "r", encoding="utf-8") as fh:
                    for line in fh:
                        line=line.strip()
                        if line and not line.startswith("#"):
                            # tillåt komma-separerat också
                            for part in re.split(r"[,\s]+", line):
                                part=part.strip().upper()
                                if part:
                                    rader.append(part)
            except Exception:
                pass
        if not rader:
            raw = self.p_stad_rader.get().strip() if hasattr(self, 'p_stad_rader') else STANDARD.get("stad_rader","")
            if raw:
                for part in re.split(r"[,\s]+", raw):
                    part=part.strip().upper()
                    if part:
                        rader.append(part)
        if not rader:
            rader = list("ABCDEFGHJKLMNPQRSTUVWXYZ") + ["BG","BPA","GALLER"]
        # unika, behåll ordning
        seen=set()
        uniq=[]
        for r in rader:
            if r not in seen:
                seen.add(r)
                uniq.append(r)
        return uniq

    def _vecka_str(self, dt):
        # ISO-vecka YYYY-Www
        try:
            iso = dt.isocalendar()
            return f"{iso[0]}-W{iso[1]:02d}"
        except Exception:
            return dt.strftime("%Y-W%W")

    def bygg_stadning_flik(self, parent):
        topp = tk.Frame(parent, bg=BAKGRUND)
        topp.pack(fill='x', padx=18, pady=(12,8))
        tk.Label(topp, text="Städning per rad – veckovis", bg=BAKGRUND, fg=TEXT, font=self.f_rubrik_stor).pack(side='left')
        self.stad_sum_label = tk.Label(topp, text="", bg=BAKGRUND, fg=DAMPAD, font=self.f_liten)
        self.stad_sum_label.pack(side='left', padx=(12,0))
        btn_rad = tk.Frame(topp, bg=BAKGRUND)
        btn_rad.pack(side='right')
        Knapp(btn_rad, "Ladda om", self.stad_ladda, '#e5e9f0', TEXT, '#d5dbe5').pack(side='left', padx=4)
        Knapp(btn_rad, "QR-koder", self.stad_skapa_qr, '#e5e9f0', TEXT, '#d5dbe5').pack(side='left', padx=4)
        Knapp(btn_rad, "Exportera", self.stad_exportera, '#e5e9f0', TEXT, '#d5dbe5').pack(side='left', padx=4)

        # Info-rad – ingen inmatning i Admin
        info = tk.Frame(parent, bg='#fef9c3', highlightbackground='#facc15', highlightthickness=1)
        info.pack(fill='x', padx=18, pady=(0,8))
        tk.Label(info, text="Admin är endast visning. Inrapportering sker i Lagersök-appen: öppna kameran → scanna QR på utsidan av raden → skriv namn → Klar. Veckan registreras automatiskt och visas här (tryck Ladda om). En rad per vecka.", bg='#fef9c3', fg='#854d0e', font=self.f_liten, wraplength=900, justify='left').pack(side='left', padx=12, pady=8)

        # Inställningar
        inst = tk.Frame(parent, bg=KORT, highlightbackground=RAM, highlightthickness=1)
        inst.pack(fill='x', padx=18, pady=(0,8))
        inner = tk.Frame(inst, bg=KORT)
        inner.pack(fill='x', padx=12, pady=8)
        tk.Label(inner, text="Rader att städa (kommasep):", bg=KORT, fg=TEXT, font=self.f_norm).pack(side='left')
        ttk.Entry(inner, textvariable=self.p_stad_rader, width=50, font=self.f_liten).pack(side='left', padx=(8,4), ipady=2)
        Knapp(inner, "Spara rader", self.stad_spara_rader, '#7fb046', '#fff').pack(side='left', padx=8)
        tk.Label(inner, text="Logg:", bg=KORT, fg=DAMPAD, font=self.f_liten).pack(side='left', padx=(12,4))
        ttk.Entry(inner, textvariable=self.p_stadlogg, width=30, font=self.f_liten).pack(side='left', padx=2)

        # filter
        sok_rad = tk.Frame(parent, bg=BAKGRUND)
        sok_rad.pack(fill='x', padx=18, pady=(0,8))
        tk.Label(sok_rad, text="Filter:", bg=BAKGRUND, fg=DAMPAD, font=self.f_liten).pack(side='left')
        self.stad_filter_var = tk.StringVar(value="Alla")
        ttk.Combobox(sok_rad, textvariable=self.stad_filter_var, values=["Alla","Ej städad denna vecka","Städad denna vecka","Över 2 veckor"], width=20, state="readonly").pack(side='left', padx=8)
        self.stad_filter_var.trace_add('write', lambda *a: self.stad_filtrera())
        Knapp(sok_rad, "Visa historik", self.stad_visa_historik, BLA, '#fff').pack(side='right', padx=4)

        # tabell – nu rad-vis
        tab_frame = tk.Frame(parent, bg=BAKGRUND)
        tab_frame.pack(fill='both', expand=True, padx=18, pady=(0,12))
        cols = ("rad","vecka","datum","vem","status")
        self.stad_tree = ttk.Treeview(tab_frame, columns=cols, show='headings', height=20)
        self.stad_tree.heading("rad", text="Rad")
        self.stad_tree.heading("vecka", text="Vecka")
        self.stad_tree.heading("datum", text="Datum")
        self.stad_tree.heading("vem", text="Städad av")
        self.stad_tree.heading("status", text="Status")
        self.stad_tree.column("rad", width=80, anchor='center')
        self.stad_tree.column("vecka", width=100, anchor='center')
        self.stad_tree.column("datum", width=130, anchor='center')
        self.stad_tree.column("vem", width=150, anchor='w')
        self.stad_tree.column("status", width=200, anchor='w')
        vsb = ttk.Scrollbar(tab_frame, orient='vertical', command=self.stad_tree.yview)
        self.stad_tree.configure(yscrollcommand=vsb.set)
        self.stad_tree.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')
        self.stad_tree.tag_configure('ok', background='#e8f5d8')
        self.stad_tree.tag_configure('ej', background='#fee2e2')
        self.stad_tree.tag_configure('varning', background='#fef9c3')
        self.rot.after(600, self.stad_ladda)

    def stad_spara_rader(self):
        try:
            d = self.till_dict()
            U.skriv_installningar(d)
            # spara även till fil för enkel delning
            try:
                fil = self._stad_rader_fil()
                os.makedirs(os.path.dirname(fil), exist_ok=True)
                with open(fil, "w", encoding="utf-8") as fh:
                    fh.write("# En rad per rad som ska städas\n")
                    for r in self._hamta_rader():
                        fh.write(r + "\n")
            except Exception:
                pass
            self.status_rad(f"Rader sparade: {', '.join(self._hamta_rader())}", GRON_MORK)
            self.stad_ladda()
        except Exception as e:
            messagebox.showerror("Fel", str(e), parent=self.rot)

    def stad_ladda(self):
        # försök även hämta från Worker om url finns
        try:
            self._hamta_fran_worker()
        except Exception:
            pass
        self._stad_rader = self._hamta_rader()
        self.stad_filtrera()

    def _hamta_fran_worker(self):
        url = self.stat_url.get().strip() if hasattr(self, 'stat_url') else ""
        nyckel = self.stat_key.get().strip() if hasattr(self, 'stat_key') else ""
        if not url:
            return
        import threading, json, urllib.request
        def jobb():
            try:
                req = urllib.request.Request(url.rstrip("/")+"/stad", headers={"X-API-Key": nyckel, "X-Admin-Key": nyckel, "User-Agent": "lagersok-uppdatering"}, method="GET")
                with urllib.request.urlopen(req, timeout=5) as resp:
                    data = json.loads(resp.read().decode("utf-8"))
                    # förväntar {E:{datum, vem, vecka, historik:[]}, ...}
                    if isinstance(data, dict):
                        # merga med lokal fil
                        lokal = self._las_stadlogg()
                        for k,v in data.items():
                            if k not in lokal or (v.get("datum","") > lokal.get(k,{}).get("datum","")):
                                lokal[k]=v
                        self._spara_stadlogg(lokal)
                        self.rot.after(0, self.stad_filtrera)
            except Exception:
                pass
        threading.Thread(target=jobb, daemon=True).start()

    def stad_filtrera(self):
        if not hasattr(self, 'stad_tree'):
            return
        logg = self._las_stadlogg()
        filter_val = self.stad_filter_var.get() if hasattr(self, 'stad_filter_var') else "Alla"
        for item in self.stad_tree.get_children():
            self.stad_tree.delete(item)
        from datetime import datetime, date, timedelta
        idag = date.today()
        iso_year, iso_week, _ = idag.isocalendar()
        denna_vecka = f"{iso_year}-W{iso_week:02d}"
        total=len(self._hamta_rader()) if hasattr(self, '_hamta_rader') else 0
        ok=0
        ej=0
        for rad in getattr(self, '_stad_rader', self._hamta_rader()):
            info = logg.get(rad, {})
            vecka = info.get("vecka","") if isinstance(info, dict) else ""
            datum = info.get("datum","")[:10] if isinstance(info, dict) and info.get("datum") else ""
            vem = info.get("vem","") if isinstance(info, dict) else ""
            # räkna dagar sedan
            dagar_sedan=None
            if info.get("datum"):
                try:
                    d = datetime.strptime(info.get("datum")[:10], "%Y-%m-%d").date()
                    dagar_sedan=(idag-d).days
                except Exception:
                    pass
            if vecka==denna_vecka:
                status="Städad denna vecka"
                tagg="ok"
                ok+=1
            else:
                if not info:
                    status="Aldrig städad"
                    tagg="ej"
                    ej+=1
                elif dagar_sedan is not None and dagar_sedan>14:
                    status=f"Över 2 veckor ({dagar_sedan}d) – {vecka}"
                    tagg="ej"
                    ej+=1
                elif dagar_sedan is not None and dagar_sedan>7:
                    status=f"Ej denna vecka, senast {vecka} ({dagar_sedan}d)"
                    tagg="varning"
                    ej+=1
                else:
                    status=f"Ej denna vecka – senast {vecka}"
                    tagg="varning"
                    ej+=1
            if filter_val=="Städad denna vecka" and tagg!="ok":
                continue
            if filter_val=="Ej städad denna vecka" and tagg=="ok":
                continue
            if filter_val=="Över 2 veckor" and not (dagar_sedan and dagar_sedan>14):
                continue
            self.stad_tree.insert('', 'end', values=(rad, vecka or "-", datum or "-", vem or "-", status), tags=(tagg,))
        self.stad_sum_label.config(text=f"{len(self._stad_rader)} rader | {ok} städade v{iso_week} | {ej} ej klara denna vecka")

    def stad_visa_historik(self):
        sel = self.stad_tree.selection()
        if not sel:
            messagebox.showinfo("Välj", "Välj en rad först", parent=self.rot)
            return
        rad = self.stad_tree.item(sel[0], 'values')[0]
        logg = self._las_stadlogg()
        info = logg.get(rad, {})
        hist = info.get("historik", []) if isinstance(info, dict) else []
        dialog = tk.Toplevel(self.rot)
        dialog.title(f"Historik {rad}")
        dialog.geometry("480x400")
        dialog.transient(self.rot)
        txt = tk.Text(dialog, font=self.f_mono, wrap='word')
        txt.pack(fill='both', expand=True, padx=10, pady=10)
        txt.insert('end', f"Rad: {rad}\nSenast: {info.get('vecka','-')} {info.get('datum','-')} av {info.get('vem','-')}\n\nHistorik (senaste 30):\n")
        if not hist:
            txt.insert('end', "Ingen historik ännu\n")
        else:
            for h in reversed(hist[-30:]):
                txt.insert('end', f"{h.get('vecka','')}  {h.get('datum','')[:10]}  {h.get('vem','')}\n")
        txt.config(state='disabled')

    def stad_skapa_qr(self):
        # Clean – bara QR, rad-nummer och snygg ram – tjock ram 5px
        rader = getattr(self, '_stad_rader', self._hamta_rader())
        # QR-koden pekar på själva Lagersök-appen med ?stad=RAD. Adressen härleds från GitHub-inställningen
        # (samma regel som länken "Öppna appen" efter uppdatering).
        webb_base = "https://sterilen.github.io/"
        _repo = self.repo.get().strip() if hasattr(self, 'repo') else ""
        if "/" in _repo:
            _agare, _arkiv = _repo.split("/", 1)
            webb_base = f"https://{_agare.lower()}.github.io/" + ("" if _arkiv.lower() == f"{_agare.lower()}.github.io" else _arkiv + "/")
        webb_base = webb_base.rstrip("/") + "/"
        import os, webbrowser, datetime
        try:
            html_path = os.path.join(_get_base_dir(), "qr_stad.html")
            with open(html_path, "w", encoding="utf-8") as fh:
                fh.write("<!doctype html>\n<html lang=\"sv\"><head><meta charset='utf-8'><meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">\n")
                fh.write("<title>QR STÄDNING</title>\n<style>\n")
                fh.write(":root{--bg:#f7f5f0;--card:#fff;--ink:#0f172a;--line:#111;}\n")
                fh.write("*{box-sizing:border-box;margin:0;padding:0}\n")
                fh.write("body{font-family:-apple-system,Inter,Segoe UI,sans-serif;background:var(--bg);color:var(--ink);padding:24px}\n")
                fh.write(".grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:18px;max-width:1200px;margin:0 auto}\n")
                fh.write(".kort{background:var(--card);border:5px solid var(--line);border-radius:18px;padding:14px 12px 12px;text-align:center}\n")
                fh.write(".rad{font-size:46px;font-weight:900;letter-spacing:-1.5px;line-height:1;margin-bottom:10px}\n")
                fh.write(".qr{display:flex;justify-content:center}\n")
                fh.write(".foot{margin-top:10px;font-size:9px;letter-spacing:1px;font-weight:700;opacity:0.35;text-transform:uppercase}\n")
                fh.write("@media print{body{background:#fff;padding:8mm} .grid{gap:10px} .kort{break-inside:avoid} @page{margin:10mm}}\n")
                fh.write("</style>\n")
                fh.write('<script src="https://cdn.jsdelivr.net/npm/qrcodejs@1.0.0/qrcode.min.js"></script>\n')
                fh.write("</head><body>\n<div class=\"grid\">\n")
                for rad in rader:
                    url = webb_base + "?stad=" + rad
                    fh.write("<div class='kort'>\n")
                    fh.write("  <div class='rad'>" + rad + "</div>\n")
                    fh.write("  <div class='qr' id='q-" + rad + "'></div>\n")
                    fh.write("  <div class='foot'>Scanna i Lagersök</div>\n")
                    fh.write("</div>\n")
                fh.write("</div>\n<script>\n")
                for rad in rader:
                    url = webb_base + "?stad=" + rad
                    fh.write("new QRCode(document.getElementById('q-" + rad + "'), {text:'" + url + "', width:150, height:150, correctLevel: QRCode.CorrectLevel.M, colorDark:'#000000', colorLight:'#ffffff'});\n")
                fh.write("</script></body></html>")
            webbrowser.open("file:///" + html_path.replace(chr(92), "/"))
            self.status_rad("Clean QR-sida skapad: " + html_path + " (" + str(len(rader)) + " rader) - tjock ram 5px", GRON_MORK)
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Fel", str(e), parent=self.rot)


    def stad_exportera(self):
        try:
            import openpyxl
        except ImportError:
            messagebox.showerror("Saknas", "openpyxl saknas - pip install openpyxl", parent=self.rot)
            return
        fil = filedialog.asksaveasfilename(parent=self.rot, title="Spara Excel", defaultextension=".xlsx", filetypes=[("Excel","*.xlsx")], initialfile="stadlogg.xlsx")
        if not fil:
            return
        logg = self._las_stadlogg()
        from datetime import datetime
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Städlogg rad"
        ws.append(["Rad","Vecka","Datum","Av vem","Status"])
        for rad in getattr(self, '_stad_rader', self._hamta_rader()):
            info = logg.get(rad, {})
            ws.append([rad, info.get("vecka",""), info.get("datum","")[:10], info.get("vem",""), ""])
        try:
            wb.save(fil)
            oppna(fil)
            self.status_rad(f"Exporterad till {fil}", GRON_MORK)
        except Exception as e:
            messagebox.showerror("Fel", str(e), parent=self.rot)


    def bygg_bilder_flik(self, parent):
        """Förenklad: 800px webp, 3 kolumner, lokalt endast vid uppladdning."""
        self.bilder_lista = []
        self._bilder_artikel_index = {}
        self._bilder_github_bilder = set()

        topp = tk.Frame(parent, bg=MORK)
        topp.pack(fill='x')
        tx = tk.Frame(topp, bg=MORK)
        tx.pack(side='left', padx=18, pady=12)
        tk.Label(tx, text="Bilder för artiklar (800px webp)", bg=MORK, fg='#ffffff', font=(FONT_FET, 14, 'bold')).pack(anchor='w')
        tk.Label(tx, text="Välj bild lokalt -> konverteras till 800px webp (kvalitet 80) -> laddas upp till GitHub. Sparas inte lokalt efteråt.", bg=MORK, fg='#9fb0c3', font=self.f_norm).pack(anchor='w')

        # Bildmappen på GitHub är alltid 'bilder', så den visas inte längre. Variabeln finns kvar för koden som använder den.
        self.bilder_gh_path_var = tk.StringVar(value='bilder')

        sok_fr = tk.Frame(parent, bg=BAKGRUND)
        sok_fr.pack(fill='x', padx=18, pady=(12,8))
        tk.Label(sok_fr, text="Sök:", bg=BAKGRUND, fg=TEXT, font=self.f_norm).pack(side='left')
        self.bilder_sok_var = tk.StringVar()
        ent = tk.Entry(sok_fr, textvariable=self.bilder_sok_var, bg=KORT, fg=TEXT, font=self.f_norm, width=30, relief='flat', highlightbackground=RAM, highlightthickness=1)
        ent.pack(side='left', padx=6, ipady=4)
        ent.bind('<KeyRelease>', lambda e: self.bilder_uppdatera_tree())

        tb = tk.Frame(parent, bg=BAKGRUND)
        tb.pack(fill='x', padx=18, pady=(0,8))
        def kn(text, cmd, bgc, fg='#ffffff', hov=None):
            return Knapp(tb, text, cmd, bgc, fg, hov).pack(side='left', padx=(0,8))
        kn("Ladda artiklar + kolla bilder", self.bilder_ladda_allt, BLA, '#ffffff', '#1e40af')
        kn("Välj bild för markerad", self.bilder_valj_och_ladda_upp, BLA, '#ffffff', '#1e40af')
        kn("Välj flera bilder...", self.bilder_valj_flera_och_ladda_upp, GRON_MORK, '#ffffff', '#14532d')
        kn("Visa online", self.bilder_visa_online, '#0ea5e9', '#ffffff', '#0284c7')
        kn("Orphans", self.bilder_visa_orphans, '#f59e0b', '#ffffff', '#d97706')
        kn("Ta bort bild online", self.bilder_ta_bort_online, ROSA, '#ffffff', '#be185d')
        kn("Inkorg…", self.inkorg_oppna, '#0f766e', '#ffffff', '#115e59')
        kn("Koder…", self.koder_oppna, '#64748b', '#ffffff', '#475569')

        # Logg längst ner först så den alltid syns
        bot = tk.Frame(parent, bg=BAKGRUND)
        bot.pack(fill='x', padx=18, pady=(0,8), side='bottom')
        log_fr = tk.Frame(bot, bg=KORT, highlightbackground=RAM, highlightthickness=1, height=225)
        log_fr.pack(fill='x', expand=False)
        log_fr.pack_propagate(False)
        tk.Label(log_fr, text="Logg (viktigt - fel visas här)", bg=KORT, fg=ROD, font=self.f_fet).pack(anchor='w', padx=10, pady=(6,0))
        log_inner = tk.Frame(log_fr, bg=KORT)
        log_inner.pack(fill='both', expand=True, padx=4, pady=4)
        self.bilder_logg = tk.Text(log_inner, height=8, bg=LOGG_BG, fg=LOGG_TEXT, font=('Consolas', 9), relief='flat', wrap='word')
        self.bilder_logg.pack(side='left', fill='both', expand=True, padx=(8,0), pady=8)
        sb_log = ttk.Scrollbar(log_inner, orient='vertical', command=self.bilder_logg.yview)
        sb_log.pack(side='right', fill='y', padx=(0,8), pady=8)
        self.bilder_logg.config(yscrollcommand=sb_log.set)
        self.bilder_status_label = tk.Label(log_fr, text="", bg=KORT, fg=BLA, font=self.f_liten, anchor='w')
        self.bilder_status_label.pack(fill='x', padx=10, pady=(0,6))

        tab_frame = tk.Frame(parent, bg=BAKGRUND)
        tab_frame.pack(fill='both', expand=True, padx=18, pady=(0,8))
        cols = ('artnr','namn','status')
        self.bilder_tree = ttk.Treeview(tab_frame, columns=cols, show='headings', height=10)
        self._bilder_sort_kolumn = 'artnr'
        self._bilder_sort_fallande = False
        def _sortera(kol):
            if self._bilder_sort_kolumn == kol:
                self._bilder_sort_fallande = not self._bilder_sort_fallande
            else:
                self._bilder_sort_kolumn = kol
                self._bilder_sort_fallande = False
            self.bilder_uppdatera_tree()
        self.bilder_tree.heading('artnr', text='Artikelnummer ▲', command=lambda: _sortera('artnr'))
        self.bilder_tree.heading('namn', text='Produktnamn', command=lambda: _sortera('namn'))
        self.bilder_tree.heading('status', text='Status', command=lambda: _sortera('status'))
        self.bilder_tree.column('artnr', width=150, anchor='w')
        self.bilder_tree.column('namn', width=350, anchor='w')
        self.bilder_tree.column('status', width=160, anchor='center')
        vsb = ttk.Scrollbar(tab_frame, orient='vertical', command=self.bilder_tree.yview)
        self.bilder_tree.configure(yscrollcommand=vsb.set)
        self.bilder_tree.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')

        self.rot.after(600, self.bilder_ladda_allt)

    def bilder_status(self, txt, color=TEXT):
        try:
            self.bilder_status_label.config(text=txt, fg=color)
            self.bilder_logg.insert('end', txt+"\n")
            self.bilder_logg.see('end')
        except Exception:
            pass

    def bilder_ladda_allt(self):
        """Hämtar artikellistan ur den PUBLICERADE data.js på GitHub (samma som appen visar) och listan över bilder online."""
        import threading
        inst = U.las_installningar_dict()
        repo = self.repo.get().strip() if hasattr(self, 'repo') else inst.get('github_repo', '')
        token = self.token.get().strip() if hasattr(self, 'token') else inst.get('github_token', '')
        gren = self.gren.get().strip() if hasattr(self, 'gren') else inst.get('github_branch', 'main')
        gren = gren or 'main'
        gh_path = self.bilder_gh_path_var.get().strip() or 'bilder'
        ut = self.p_ut.get().strip().strip('"') if hasattr(self, 'p_ut') else inst.get('ut', '')
        filnamn = os.path.basename(ut) or 'data.js'

        def hamta():
            import json, urllib.request, urllib.error
            # 1. Artikellistan: publicerad data.js på GitHub, annars (om det inte går) filen på disk
            idx, kalla = None, ''
            try:
                db = U.las_databas_text(U.hamta_databas_github(repo, token, gren, filnamn), 'GitHub')
                idx = {a['artnr']: a['namn'] for a in db if a['artnr']}
                kalla = 'GitHub (publicerad ' + (getattr(U.las_databas_text, 'senaste_stampel', '') or 'okänd tid') + ')'
            except Exception as e:
                self.ko.put(('bilder_log', f"Kunde inte hämta data.js från GitHub: {e}"))
                try:
                    db = U.las_databas(ut)
                    idx = {a['artnr']: a['namn'] for a in db if a['artnr']}
                    kalla = 'fil på datorn (reserv, kan skilja sig från appen)'
                except Exception as e2:
                    self.ko.put(('bilder_log', f"Kunde inte läsa data.js på disk heller: {e2}"))
            if idx is not None:
                self.ko.put(('bilder_index', idx))
                self.ko.put(('bilder_log', f"{len(idx)} artiklar inlästa från {kalla}"))
            # 2. Vilka bilder finns online
            try:
                if not repo or not token:
                    self.ko.put(('bilder_log', "GitHub repo/token saknas - kan inte kolla befintliga bilder"))
                    return
                req = urllib.request.Request(f"https://api.github.com/repos/{repo}/contents/{gh_path}?ref={gren}",
                                             headers={'Authorization': f'Bearer {token}', 'Accept': 'application/vnd.github.v3+json', 'User-Agent': 'lagersok-bilder'})
                with urllib.request.urlopen(req, timeout=30) as r:
                    data = json.loads(r.read().decode('utf-8'))
                bilder = set()
                if isinstance(data, list):
                    for f in data:
                        namn = f.get('name', '')
                        if namn.lower().endswith('.webp') and not namn.lower().endswith('_thumb.webp'):
                            bilder.add(os.path.splitext(namn)[0])
                self._bilder_github_bilder = bilder
                self.ko.put(('bilder_log', f"{len(bilder)} bilder hittades online i {gh_path}/"))
            except urllib.error.HTTPError as he:
                if he.code == 404:
                    self.ko.put(('bilder_log', f"GitHub mappen {gh_path} finns inte än"))
                    self._bilder_github_bilder = set()
                else:
                    self.ko.put(('bilder_log', f"Fel vid hämtning av bilder: {he}"))
            except Exception as e:
                self.ko.put(('bilder_log', f"Fel: {e}"))
            self.ko.put(('bilder_refresh', True))

        threading.Thread(target=hamta, daemon=True).start()
        self.bilder_status("Hämtar artiklar från den publicerade data.js och kollar vilka bilder som finns online...", BLA)
        self.rot.after(300, self.bilder_poll_queue)

    def bilder_bygg_lista_fran_index(self):
        self.bilder_lista = []
        for artnr, namn in (self._bilder_artikel_index or {}).items():
            har = artnr in self._bilder_github_bilder
            self.bilder_lista.append({"artnr": artnr, "namn": namn, "har_bild": har})
        self.bilder_lista.sort(key=lambda x: x['artnr'])
        self.bilder_uppdatera_tree()

    def bilder_uppdatera_tree(self):
        for i in self.bilder_tree.get_children():
            self.bilder_tree.delete(i)
        sok = self.bilder_sok_var.get().lower().strip() if hasattr(self, 'bilder_sok_var') else ""
        # Filtrera först
        filtrerad = []
        for it in self.bilder_lista:
            if sok and sok not in it['artnr'].lower() and sok not in it.get('namn','').lower():
                continue
            filtrerad.append(it)
        # Sortera enligt vald kolumn
        kol = getattr(self, '_bilder_sort_kolumn', 'artnr')
        fallande = getattr(self, '_bilder_sort_fallande', False)
        def sort_key(it):
            if kol == 'artnr':
                return it['artnr']
            elif kol == 'namn':
                return (it['namn'] or '').lower()
            elif kol == 'status':
                # Bild finns först om stigande
                return 0 if it['har_bild'] else 1
            return it['artnr']
        try:
            filtrerad.sort(key=sort_key, reverse=fallande)
        except Exception:
            pass
        # Uppdatera rubriker med pil
        try:
            pil = ' ▼' if fallande else ' ▲'
            for k, label in [('artnr','Artikelnummer'),('namn','Produktnamn'),('status','Status')]:
                txt = label + (pil if k == kol else '')
                self.bilder_tree.heading(k, text=txt)
        except Exception:
            pass
        for it in filtrerad:
            status_text = "Bild finns" if it['har_bild'] else "Bild finns inte"
            tag = 'finns' if it['har_bild'] else 'saknas'
            self.bilder_tree.insert('', 'end', values=(it['artnr'], it['namn'], status_text), tags=(tag,))
        try:
            self.bilder_tree.tag_configure('finns', foreground=GRON_MORK)
            self.bilder_tree.tag_configure('saknas', foreground=DAMPAD)
        except:
            pass

    def bilder_valj_och_ladda_upp(self):
        sel = self.bilder_tree.selection()
        if not sel:
            messagebox.showinfo("Välj artikel", "Markera en artikel först.", parent=self.rot)
            return
        artnr = self.bilder_tree.item(sel[0], 'values')[0]
        fil = filedialog.askopenfilename(title=f"Välj bild för {artnr}", filetypes=[("Bilder","*.jpg *.jpeg *.png *.webp *.bmp"),("Alla filer","*.*")])
        if not fil:
            return
        self.bilder_konvertera_och_ladda_upp([(artnr, fil)])

    def _fraga_artnr(self, fil, nr, tot, forslag):
        """Dialog med förhandsvisning av bilden och fält för artikelnummer.
        Returnerar artikelnummer (str), None = hoppa över, 'AVBRYT' = sluta fråga."""
        import re
        index = self._bilder_artikel_index or {}
        top = tk.Toplevel(self.rot)
        top.title("Vilken artikel är bilden på?")
        top.transient(self.rot)
        top.resizable(False, False)
        res = {'v': 'AVBRYT'}
        ram = ttk.Frame(top, padding=14)
        ram.pack()
        ttk.Label(ram, text=f"Bild {nr} av {tot}:  {os.path.basename(fil)}", font=(FONT_FET, 10, 'bold')).pack(anchor='w', pady=(0, 8))
        bildruta = tk.Label(ram, bg='#f1f5f9', relief='solid', bd=1)
        bildruta.pack(fill='both')
        try:
            from PIL import Image, ImageOps, ImageTk
            im = Image.open(fil)
            try:
                im = ImageOps.exif_transpose(im)
            except Exception:
                pass
            im = im.convert('RGB')
            im.thumbnail((420, 380))
            foto = ImageTk.PhotoImage(im)
            bildruta.configure(image=foto)
            bildruta.image = foto  # håll referensen så bilden inte försvinner
        except Exception as e:
            logga(f"Förhandsvisning misslyckades för {fil}: {e!r}")
            bildruta.configure(text="(Ingen förhandsvisning)", width=40, height=8)
        ttk.Label(ram, text="Artikelnummer:").pack(anchor='w', pady=(10, 2))
        var = tk.StringVar(value=forslag)
        ent = ttk.Entry(ram, textvariable=var, font=(FONT_FET, 13), width=24)
        ent.pack(anchor='w')
        namnlbl = tk.Label(ram, text="", anchor='w', justify='left', wraplength=420)
        namnlbl.pack(anchor='w', pady=(4, 0), fill='x')

        def rensa(t):
            return re.sub(r'[^A-Za-z0-9]', '', t.strip()).upper()

        def uppdatera(*_):
            a = rensa(var.get())
            namn = index.get(a)
            if namn:
                namnlbl.configure(text=f"\u2714 {namn}", fg='#15803d')
            elif a:
                namnlbl.configure(text="Finns inte i artikellistan. Bilden måste höra till en artikel.", fg='#b91c1c')
            else:
                namnlbl.configure(text="")
        var.trace_add('write', uppdatera)
        uppdatera()

        def ok(_e=None):
            a = rensa(var.get())
            if not a or a not in index:
                top.bell()
                return
            res['v'] = a
            top.destroy()

        def hoppa():
            res['v'] = None
            top.destroy()

        def avbryt(_e=None):
            res['v'] = 'AVBRYT'
            top.destroy()
        kn_ram = ttk.Frame(ram)
        kn_ram.pack(fill='x', pady=(12, 0))
        ttk.Button(kn_ram, text="OK (Enter)", command=ok).pack(side='left')
        ttk.Button(kn_ram, text="Hoppa över bilden", command=hoppa).pack(side='left', padx=8)
        ttk.Button(kn_ram, text="Avbryt allt", command=avbryt).pack(side='right')
        top.bind('<Return>', ok)
        top.bind('<Escape>', avbryt)
        top.protocol("WM_DELETE_WINDOW", avbryt)
        top.update_idletasks()
        x = self.rot.winfo_rootx() + (self.rot.winfo_width() - top.winfo_width()) // 2
        y = self.rot.winfo_rooty() + max(20, (self.rot.winfo_height() - top.winfo_height()) // 3)
        top.geometry(f"+{max(0, x)}+{max(0, y)}")
        try:
            top.grab_set()
        except tk.TclError:
            pass
        ent.focus_set()
        ent.select_range(0, 'end')
        self.rot.wait_window(top)
        return res['v']

    def bilder_valj_flera_och_ladda_upp(self):
        filer = filedialog.askopenfilenames(title="Välj bilder", filetypes=[("Bilder","*.jpg *.jpeg *.png *.webp *.bmp"),("Alla filer","*.*")])
        if not filer:
            return
        import re
        index = self._bilder_artikel_index or {}
        if not index:
            messagebox.showinfo("Artikellistan är inte inläst",
                                "Bilder får bara laddas upp till artiklar som finns i databasen.\n"
                                "Tryck först på \"Ladda artiklar + kolla bilder\" och försök sedan igen.", parent=self.rot)
            return
        # Filer som redan heter som en känd artikel går rakt igenom. Alla andra frågar vi om
        # (med förhandsvisning). Bara kända artikelnummer godtas.
        auto, fraga_lista = [], []
        for f in filer:
            base = os.path.splitext(os.path.basename(f))[0]
            artnr = re.sub(r'[^A-Za-z0-9]', '', base).upper()
            if index and artnr in index:
                auto.append((artnr, f))
            else:
                fraga_lista.append((artnr, f))
        par = list(auto)
        for i, (forslag, f) in enumerate(fraga_lista, 1):
            svar = self._fraga_artnr(f, i, len(fraga_lista), forslag)
            if svar == 'AVBRYT':
                if not messagebox.askyesno("Avbryt", "Sluta fråga? Bilderna du redan angett artikelnummer för laddas fortfarande upp (om du bekräftar uppladdningen).\n\nSvara Nej för att avbryta helt.", parent=self.rot):
                    return
                break
            if svar:
                par.append((svar, f))
        if not par:
            return
        self.bilder_konvertera_och_ladda_upp(par)

    def bilder_konvertera_och_ladda_upp(self, par_lista, ok_cb=None, fraga=True):
        try:
            from PIL import Image
        except ImportError:
            messagebox.showerror("Pillow saknas", "Kör: pip install Pillow", parent=self.rot)
            return
        repo = self.repo.get().strip() if hasattr(self, 'repo') else U.las_installningar_dict().get('github_repo', '')
        token = self.token.get().strip() if hasattr(self, 'token') else U.las_installningar_dict().get('github_token', '')
        branch = self.gren.get().strip() if hasattr(self, 'gren') else U.las_installningar_dict().get('github_branch', 'main')
        gh_path = self.bilder_gh_path_var.get().strip() or 'bilder'
        if not repo or not token:
            messagebox.showerror("GitHub saknas", f"Repo='{repo}' Token finns={'ja' if token else 'nej'}\nFyll i GitHub repo och token i HEM-fliken (de sparas i installningar.dat).", parent=self.rot)
            self.bilder_status(f"FEL: GitHub repo/token saknas - repo='{repo}'", ROD)
            return
        import re
        rensad = []
        for artnr, kalla in par_lista:
            a = re.sub(r'[^A-Za-z0-9]', '', str(artnr)).upper()
            if not a:
                continue
            rensad.append((a, kalla))
        par_lista = rensad
        if not par_lista:
            messagebox.showinfo("Inget giltigt", "Inga giltiga artikelnummer (endast A-Z 0-9).", parent=self.rot)
            return
        # Spärr: en bild måste höra till en artikel som finns i databasen.
        index = self._bilder_artikel_index or {}
        if not index:
            messagebox.showerror("Artikellistan är inte inläst",
                                 "Bilder får bara laddas upp till artiklar som finns i databasen.\n"
                                 "Tryck först på \"Ladda artiklar + kolla bilder\" och försök sedan igen.", parent=self.rot)
            return
        okanda = [a for a, _k in par_lista if a not in index]
        if okanda:
            messagebox.showwarning("Okända artiklar",
                                   "Dessa artikelnummer finns inte i databasen och hoppas över:\n" + ", ".join(sorted(set(okanda))[:20]), parent=self.rot)
            par_lista = [(a, k) for a, k in par_lista if a in index]
            if not par_lista:
                return
        if fraga and not messagebox.askyesno("Ladda upp?", f"Konvertera {len(par_lista)} bilder till 800px + 200px thumb webp och ladda upp till {repo}/{gh_path}/ ?\nFilnamn lokalt spelar ingen roll - sparas som ARTNR.webp", parent=self.rot):
            return
        def jobb():
            import base64, json, urllib.request, urllib.error, tempfile, time, random, os
            ok = 0
            fel_lista = []
            for idx, (artnr, kalla) in enumerate(par_lista):
                try:
                    im = Image.open(kalla)
                    if im.mode in ('RGBA','LA','P'):
                        bg = Image.new('RGB', im.size, (255,255,255))
                        if im.mode == 'P':
                            im = im.convert('RGBA')
                        bg.paste(im, mask=im.split()[-1] if im.mode in ('RGBA','LA') else None)
                        im = bg
                    else:
                        im = im.convert('RGB')
                    im_full = im.copy()
                    im_full.thumbnail((800,800), Image.LANCZOS)
                    im_thumb = im.copy()
                    im_thumb.thumbnail((200,200), Image.LANCZOS)

                    def upload_one(pil_img, suffix):
                        fd, tmp_webp = tempfile.mkstemp(suffix='.webp')
                        os.close(fd)
                        pil_img.save(tmp_webp, 'WEBP', quality=80, method=6)
                        with open(tmp_webp, 'rb') as fh:
                            b64 = base64.b64encode(fh.read()).decode('ascii')
                        try:
                            os.remove(tmp_webp)
                        except Exception:
                            pass
                        remote_path = f"{gh_path.strip('/')}/{artnr}{suffix}.webp"
                        api_url = f"https://api.github.com/repos/{repo}/contents/{remote_path}"
                        for attempt in range(5):
                            try:
                                sha = None
                                for get_try in range(3):
                                    try:
                                        req_get = urllib.request.Request(api_url, headers={'Authorization': f'Bearer {token}', 'Accept': 'application/vnd.github.v3+json', 'User-Agent': 'lagersok-bilder'})
                                        with urllib.request.urlopen(req_get, timeout=30) as r:
                                            j = json.loads(r.read().decode('utf-8'))
                                            sha = j.get('sha')
                                        break
                                    except urllib.error.HTTPError as he_get:
                                        body_get = ""
                                        try:
                                            body_get = he_get.read().decode('utf-8')
                                        except Exception:
                                            pass
                                        if he_get.code in (301, 302, 307, 308):
                                            try:
                                                j = json.loads(body_get) if body_get else {}
                                                new_url = j.get('url')
                                                if new_url and new_url != api_url:
                                                    self.ko.put(('bilder_log', f"Redirect {he_get.code} för {artnr}{suffix} -> {new_url}"))
                                                    api_url = new_url
                                                    continue
                                                else:
                                                    try:
                                                        dir_url = f"https://api.github.com/repos/{repo}/contents/{gh_path.strip('/')}?ref={branch}"
                                                        req_dir = urllib.request.Request(dir_url, headers={'Authorization': f'Bearer {token}', 'Accept': 'application/vnd.github.v3+json', 'User-Agent': 'lagersok-bilder'})
                                                        with urllib.request.urlopen(req_dir, timeout=30) as r_dir:
                                                            data_dir = json.loads(r_dir.read().decode('utf-8'))
                                                            if isinstance(data_dir, list):
                                                                for f in data_dir:
                                                                    if f.get('name','').lower() == f"{artnr}{suffix}.webp".lower():
                                                                        sha = f.get('sha')
                                                                        break
                                                        if sha:
                                                            break
                                                    except Exception:
                                                        pass
                                                    sha = None
                                                    break
                                            except Exception:
                                                sha = None
                                                break
                                        if he_get.code == 404:
                                            sha = None
                                            break
                                        elif he_get.code in (429, 403) and attempt < 4:
                                            raise he_get
                                        elif he_get.code == 401:
                                            raise RuntimeError(f"401 Unauthorized - token fel eller utgången för {repo}")
                                        else:
                                            sha = None
                                            break
                                payload = {"message": f"Bild {artnr}{suffix}.webp via Admin", "content": b64, "branch": branch}
                                if sha:
                                    payload["sha"] = sha
                                req_put = urllib.request.Request(api_url, data=json.dumps(payload).encode('utf-8'), headers={'Authorization': f'Bearer {token}', 'Accept': 'application/vnd.github.v3+json', 'Content-Type': 'application/json', 'User-Agent': 'lagersok-bilder'}, method='PUT')
                                with urllib.request.urlopen(req_put, timeout=60) as r:
                                    r.read()
                                return True
                            except urllib.error.HTTPError as he:
                                body = ""
                                try:
                                    body = he.read().decode('utf-8')[:800]
                                except Exception:
                                    pass
                                if he.code in (301, 302, 307, 308):
                                    try:
                                        j = json.loads(body) if body else {}
                                        new_url = j.get('url')
                                        if new_url and new_url != api_url:
                                            self.ko.put(('bilder_log', f"PUT redirect {he.code} -> {new_url}"))
                                            api_url = new_url
                                            continue
                                    except Exception:
                                        pass
                                if he.code in (429, 403) and attempt < 4:
                                    wait = (2 ** attempt) + random.random()
                                    self.ko.put(('bilder_log', f"Rate limit {artnr}{suffix} - väntar {wait:.1f}s (försök {attempt+1}/5) {he.code} {body[:200]}"))
                                    time.sleep(wait)
                                    continue
                                elif he.code == 422:
                                    self.ko.put(('bilder_log', f"422 för {artnr}{suffix} - försöker hämta sha via mapp-listning"))
                                    try:
                                        dir_url = f"https://api.github.com/repos/{repo}/contents/{gh_path.strip('/')}?ref={branch}"
                                        req_dir = urllib.request.Request(dir_url, headers={'Authorization': f'Bearer {token}', 'Accept': 'application/vnd.github.v3+json', 'User-Agent': 'lagersok-bilder'})
                                        with urllib.request.urlopen(req_dir, timeout=30) as r_dir:
                                            data_dir = json.loads(r_dir.read().decode('utf-8'))
                                            if isinstance(data_dir, list):
                                                for f in data_dir:
                                                    if f.get('name','').lower() == f"{artnr}{suffix}.webp".lower():
                                                        sha = f.get('sha')
                                                        payload = {"message": f"Bild {artnr}{suffix}.webp via Admin", "content": b64, "branch": branch, "sha": sha}
                                                        req_put2 = urllib.request.Request(api_url, data=json.dumps(payload).encode('utf-8'), headers={'Authorization': f'Bearer {token}', 'Accept': 'application/vnd.github.v3+json', 'Content-Type': 'application/json', 'User-Agent': 'lagersok-bilder'}, method='PUT')
                                                        with urllib.request.urlopen(req_put2, timeout=60) as r2:
                                                            r2.read()
                                                        return True
                                    except Exception as e2:
                                        self.ko.put(('bilder_log', f"Kunde inte hämta sha via lista: {e2}"))
                                    raise RuntimeError(f"GitHub 422 för {artnr}{suffix}: {body}")
                                else:
                                    raise RuntimeError(f"GitHub fel {he.code} för {artnr}{suffix}: {body}")
                        return False

                    if not upload_one(im_full, ''):
                        raise RuntimeError("Kunde inte ladda upp fullbild")
                    time.sleep(0.6)
                    try:
                        if not upload_one(im_thumb, '_thumb'):
                            self.ko.put(('bilder_log', f"Varning {artnr}: thumb kunde inte laddas upp, full finns"))
                    except Exception as e_thumb:
                        self.ko.put(('bilder_log', f"Varning thumb {artnr}: {e_thumb}"))
                    ok += 1
                    self._bilder_github_bilder.add(artnr)
                    if ok_cb:
                        try:
                            ok_cb(artnr)
                        except Exception:
                            pass
                    self.ko.put(('bilder_log', f"OK {artnr} -> {artnr}.webp + {artnr}_thumb.webp till {repo}/{gh_path}/"))
                    if idx < len(par_lista)-1:
                        time.sleep(0.8)
                except Exception as e:
                    err = f"FEL {artnr}: {e}"
                    self.ko.put(('bilder_log', err))
                    fel_lista.append(err)
            self.ko.put(('bilder_log', f"Klart: {ok}/{len(par_lista)} uppladdade (800px + thumb)"))
            if fel_lista:
                self.ko.put(('bilder_log', "FEL-översikt: " + " | ".join(fel_lista[:5])))
            self.ko.put(('bilder_refresh', True))
        import threading
        threading.Thread(target=jobb, daemon=True).start()
        self.bilder_status(f"Konverterar och laddar upp {len(par_lista)} bilder (800px + thumb)... kolla loggen nedan", BLA)


    # ===================== BIDRAGSBILDER: inkorg och koder (Cloudflare-tjänsten bild_worker.js) =====================
    def bild_api(self, metod, sokvag, data=None, rå=False, tid=30):
        """Anrop till bildtjänsten med den hemliga nyckeln. Returnerar JSON (eller råa bytes om rå=True). Kastar RuntimeError med ett begripligt meddelande."""
        import json, urllib.request, urllib.error
        url = self.p_bild_url.get().strip().strip('"').rstrip('/')
        if url and not url.lower().startswith(('http://', 'https://')):
            url = 'https://' + url                      # adressen får skrivas utan https://
        nyckel = self.p_bild_nyckel.get().strip()
        if not url or not nyckel:
            raise RuntimeError("Bildtjänstens adress och nyckel saknas. Fyll i bild_url och bild_nyckel i STANDARD överst i Lagersok_Admin.py och bygg om (eller lägg dem i installningar.dat).")
        hdr = {'x-admin-key': nyckel, 'User-Agent': 'lagersok-admin'}
        kropp = None
        if data is not None:
            hdr['Content-Type'] = 'application/json'
            kropp = json.dumps(data).encode('utf-8')
        req = urllib.request.Request(url + sokvag, data=kropp, method=metod, headers=hdr)
        try:
            with urllib.request.urlopen(req, timeout=tid) as r:
                b = r.read()
        except urllib.error.HTTPError as e:
            if e.code == 401:
                raise RuntimeError("Fel nyckel till bildtjänsten.")
            raise RuntimeError(f"Bildtjänsten svarade {e.code} {e.reason}")
        except urllib.error.URLError as e:
            raise RuntimeError("Kunde inte nå bildtjänsten (" + str(e.reason) + ")")
        return b if rå else json.loads(b.decode('utf-8'))

    def bilder_pages_bas(self):
        """Adressen där bilderna ligger på den publika sidan (GitHub Pages), t.ex. https://sterilen.github.io/bilder/ , eller ''."""
        repo = self.repo.get().strip() if hasattr(self, 'repo') else ''
        if '/' not in repo:
            return ''
        owner, name = repo.split('/', 1)
        bas = '' if name.lower() == f"{owner.lower()}.github.io" else name + "/"
        sokv = (self.bilder_gh_path_var.get().strip() if hasattr(self, 'bilder_gh_path_var') else '') or 'bilder'
        return f"https://{owner.lower()}.github.io/{bas}{sokv.strip('/')}/"

    def inkorg_oppna(self):
        try:
            from PIL import Image, ImageTk
        except ImportError:
            messagebox.showerror("Pillow saknas", "Kör: pip install Pillow", parent=self.rot)
            return
        import io, urllib.parse, urllib.request, datetime, tempfile
        top = tk.Toplevel(self.rot)
        top.title("Inkorg – bilder från personalen")
        top.geometry("1180x700")
        top.transient(self.rot)
        top.configure(bg=BAKGRUND)
        # Större text i listan än i övriga tabeller
        _stil = ttk.Style()
        _stil.configure('Inkorg.Treeview', font=(FONT, 12), rowheight=34)
        _stil.configure('Inkorg.Treeview.Heading', font=(FONT_FET, 11, 'bold'))
        tb = tk.Frame(top, bg=BAKGRUND)
        tb.pack(fill='x', padx=14, pady=(12, 6))
        status = tk.Label(top, text="", bg=BAKGRUND, fg=DAMPAD, font=self.f_liten, anchor='w', justify='left', wraplength=1000)
        status.pack(side='bottom', fill='x', padx=14, pady=(4, 10))
        mitt = tk.Frame(top, bg=BAKGRUND)
        mitt.pack(fill='both', expand=True, padx=14)
        vans = tk.Frame(mitt, bg=BAKGRUND)
        vans.pack(side='left', fill='both', expand=True)
        kol = ('tid', 'art', 'namn', 'fran', 'typ')
        tree = ttk.Treeview(vans, columns=kol, show='headings', selectmode='extended', height=16, style='Inkorg.Treeview')
        for k, (txt, bredd) in {'tid': ('Inkom', 150), 'art': ('Artikel', 100), 'namn': ('Produkt', 260), 'fran': ('Från', 110), 'typ': ('Typ', 110)}.items():
            tree.heading(k, text=txt)
            tree.column(k, width=bredd, anchor='w', stretch=(k == 'namn'))
        rull = ttk.Scrollbar(vans, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=rull.set)
        tree.pack(side='left', fill='both', expand=True)
        rull.pack(side='right', fill='y')
        hog = tk.Frame(mitt, bg=KORT, highlightbackground=RAM, highlightthickness=1, width=400)
        hog.pack(side='right', fill='y', padx=(12, 0))
        hog.pack_propagate(False)
        # Informationskort överst: artikelnummer stort, produktnamn, vem som skickat in och när
        kort_fr = tk.Frame(hog, bg=KORT)
        kort_fr.pack(fill='x', padx=12, pady=(12, 6))
        info = tk.Label(kort_fr, text="Markera en bild i listan.", bg=KORT, fg=DAMPAD, font=self.f_norm, anchor='w', justify='left', wraplength=370)
        info.pack(fill='x')
        info_art = tk.Label(kort_fr, text="", bg=KORT, fg=MORK, font=(FONT_FET, 20, 'bold'), anchor='w')
        info_namn = tk.Label(kort_fr, text="", bg=KORT, fg=TEXT, font=(FONT, 12), anchor='w', justify='left', wraplength=370)
        info_rad = tk.Frame(kort_fr, bg=KORT)
        info_typ = tk.Label(info_rad, text="", font=(FONT_FET, 9, 'bold'), fg='#ffffff', padx=8, pady=2)
        info_fran = tk.Label(info_rad, text="", bg=KORT, fg=DAMPAD, font=(FONT, 11), anchor='w')
        info_varn = tk.Label(kort_fr, text="", bg=KORT, fg=ROD, font=(FONT_FET, 10, 'bold'), anchor='w', justify='left', wraplength=370)
        tk.Frame(hog, bg=RAM, height=1).pack(fill='x', padx=12, pady=(2, 6))

        def info_tom(text="Markera en bild i listan."):
            for w in (info_art, info_namn, info_rad, info_varn):
                w.pack_forget()
            info.config(text=text)
            info.pack(fill='x')

        def info_satt(artnr, namn, person, tid, ersatt, varning):
            info.pack_forget()
            for w in (info_art, info_namn, info_rad, info_varn):
                w.pack_forget()
            info_art.config(text=artnr or "?")
            info_art.pack(fill='x')
            info_namn.config(text=namn or "(okänd produkt)")
            info_namn.pack(fill='x', pady=(0, 6))
            info_typ.config(text=" ERSÄTTNING " if ersatt else " NY BILD ", bg='#d97706' if ersatt else GRON_MORK)
            info_fran.config(text=f"Från {person or 'okänd'}" + (f"  ·  {tid}" if tid else ""))
            info_rad.pack(fill='x')
            info_typ.pack(side='left')
            info_fran.pack(side='left', padx=(10, 0))
            if varning:
                info_varn.config(text=varning)
                info_varn.pack(fill='x', pady=(6, 0))
        tk.Label(hog, text="Förslag", bg=KORT, fg=DAMPAD, font=self.f_liten).pack(anchor='w', padx=12)
        bild_ny = tk.Label(hog, bg='#ffffff')
        bild_ny.pack(padx=10, pady=(2, 8))
        lbl_gammal = tk.Label(hog, text="Nuvarande bild", bg=KORT, fg=DAMPAD, font=self.f_liten)
        lbl_gammal.pack(anchor='w', padx=10)
        bild_gammal = tk.Label(hog, bg='#ffffff')
        bild_gammal.pack(padx=10, pady=(2, 8))
        st = {'bilder': {}}

        def satt_status(t, farg=DAMPAD):
            status.config(text=t, fg=farg)

        def hamta():
            try:
                lista = self.bild_api('GET', '/admin/inkorg')['bilder']
            except Exception as e:
                satt_status(str(e), ROD)
                return
            tree.delete(*tree.get_children())
            st['bilder'] = {b['id']: b for b in lista}
            namn_index = self._bilder_artikel_index or {}
            for b in lista:
                tid = datetime.datetime.fromtimestamp((b.get('tid') or 0) / 1000).strftime('%Y-%m-%d %H:%M')
                tree.insert('', 'end', iid=b['id'], values=(tid, b.get('artnr', ''), namn_index.get(b.get('artnr', ''), ''), b.get('person', ''),
                                                              'Ersättning' if b.get('ersatt') else 'Ny bild'))
            satt_status(f"{len(lista)} bilder i inkorgen." if lista else "Inkorgen är tom.", GRON_MORK if not lista else TEXT)
            info_tom()
            bild_ny.config(image='', text='')
            bild_gammal.config(image='', text='')

        def till_bild(bytes_, max_px=340):
            im = Image.open(io.BytesIO(bytes_))
            im.load()
            im.thumbnail((max_px, max_px))
            return ImageTk.PhotoImage(im.convert('RGB'))

        def visa(ev=None):
            sel = tree.selection()
            if not sel:
                return
            b = st['bilder'].get(sel[0])
            if not b:
                return
            artnr = b.get('artnr', '')
            index = self._bilder_artikel_index or {}
            varn = ''
            if index and artnr not in index:
                varn = "⚠ Artikelnumret finns inte i databasen."
            tid_txt = datetime.datetime.fromtimestamp((b.get('tid') or 0) / 1000).strftime('%Y-%m-%d %H:%M') if b.get('tid') else ''
            info_satt(artnr, index.get(artnr, ''), b.get('person', ''), tid_txt, bool(b.get('ersatt')), varn)
            try:
                data = self.bild_api('GET', '/admin/bild?id=' + urllib.parse.quote(b['id']), rå=True, tid=60)
                bild_ny.image = till_bild(data)
                bild_ny.config(image=bild_ny.image, text='')
            except Exception as e:
                bild_ny.config(image='', text=f"Kunde inte visa bilden: {e}")
            bild_gammal.config(image='', text='')
            lbl_gammal.config(text="Nuvarande bild" if b.get('ersatt') else "Nuvarande bild (artikeln ska sakna bild)")
            bas = self.bilder_pages_bas()
            if b.get('ersatt') and bas:
                try:
                    req = urllib.request.Request(f"{bas}{artnr}.webp", headers={'User-Agent': 'lagersok-admin'})
                    with urllib.request.urlopen(req, timeout=20) as r:
                        bild_gammal.image = till_bild(r.read(), 220)
                    bild_gammal.config(image=bild_gammal.image, text='')
                except Exception:
                    bild_gammal.config(image='', text="Ingen nuvarande bild hittades.")

        def valda():
            return [st['bilder'][i] for i in tree.selection() if i in st['bilder']]

        def godkann():
            val = valda()
            if not val:
                messagebox.showinfo("Godkänn", "Markera en eller flera bilder först.", parent=top)
                return
            arts = [b['artnr'] for b in val]
            if len(set(arts)) != len(arts):
                messagebox.showwarning("Godkänn", "Du har markerat flera bilder för samma artikel. Godkänn en i taget.", parent=top)
                return
            if not messagebox.askyesno("Godkänn", f"Godkänna {len(val)} bilder? De konverteras (800 px + miniatyr) och publiceras på GitHub direkt.\n"
                                       "Finns det redan en bild på artikeln ersätts den.", parent=top):
                return
            par, id_for = [], {}
            try:
                for b in val:
                    data = self.bild_api('GET', '/admin/bild?id=' + urllib.parse.quote(b['id']), rå=True, tid=60)
                    fd, tmp = tempfile.mkstemp(suffix='.jpg')
                    os.close(fd)
                    with open(tmp, 'wb') as fh:
                        fh.write(data)
                    par.append((b['artnr'], tmp))
                    id_for[b['artnr']] = (b['id'], tmp)
            except Exception as e:
                messagebox.showerror("Godkänn", f"Kunde inte hämta bilderna:\n{e}", parent=top)
                return

            def klar(artnr):       # körs när just den bilden är uppladdad: ta bort den ur inkorgen
                id_, tmp = id_for.get(artnr, (None, None))
                try:
                    if id_:
                        self.bild_api('DELETE', '/admin/bild?id=' + urllib.parse.quote(id_))
                        self.ko.put(('bilder_log', f"{artnr}: borttagen ur inkorgen"))
                except Exception as e:
                    self.ko.put(('bilder_log', f"Kunde inte ta bort {artnr} ur inkorgen: {e}"))
                try:
                    if tmp:
                        os.remove(tmp)
                except OSError:
                    pass
            self.bilder_konvertera_och_ladda_upp(par, ok_cb=klar, fraga=False)
            satt_status("Uppladdning pågår. Följ loggen under BILDER. Tryck på Hämta när den är klar för att uppdatera listan.", BLA)

        def avvisa():
            val = valda()
            if not val:
                messagebox.showinfo("Avvisa", "Markera en eller flera bilder först.", parent=top)
                return
            if not messagebox.askyesno("Avvisa", f"Radera {len(val)} bilder ur inkorgen utan att publicera dem?", parent=top):
                return
            fel = []
            for b in val:
                try:
                    self.bild_api('DELETE', '/admin/bild?id=' + urllib.parse.quote(b['id']))
                except Exception as e:
                    fel.append(str(e))
            hamta()
            if fel:
                satt_status("Kunde inte radera alla: " + fel[0], ROD)

        tree.bind('<<TreeviewSelect>>', visa)
        Knapp(tb, "Hämta", hamta, BLA, '#ffffff', '#1e40af').pack(side='left', padx=(0, 8))
        Knapp(tb, "Godkänn markerade", godkann, GRON_MORK, '#ffffff', '#14532d').pack(side='left', padx=(0, 8))
        Knapp(tb, "Avvisa markerade", avvisa, ROSA, '#ffffff', '#be185d').pack(side='left', padx=(0, 8))
        Knapp(tb, "Stäng", top.destroy, '#94a3b8', '#ffffff', '#64748b').pack(side='right')
        top.after(200, hamta)

    def koder_oppna(self):
        import datetime
        top = tk.Toplevel(self.rot)
        top.title("Koder för bidragsgivare")
        top.geometry("700x440")
        top.transient(self.rot)
        top.configure(bg=BAKGRUND)
        tb = tk.Frame(top, bg=BAKGRUND)
        tb.pack(fill='x', padx=14, pady=(12, 6))
        status = tk.Label(top, text="", bg=BAKGRUND, fg=DAMPAD, font=self.f_liten, anchor='w', justify='left', wraplength=660)
        status.pack(side='bottom', fill='x', padx=14, pady=(4, 10))
        kol = ('namn', 'kod', 'status', 'skapad')
        tree = ttk.Treeview(top, columns=kol, show='headings', selectmode='browse', height=12)
        for k, (txt, bredd) in {'namn': ('Namn', 220), 'kod': ('Kod', 140), 'status': ('Status', 100), 'skapad': ('Skapad', 140)}.items():
            tree.heading(k, text=txt)
            tree.column(k, width=bredd, anchor='w')
        tree.pack(fill='both', expand=True, padx=14)
        st = {'koder': {}}

        def satt_status(t, farg=DAMPAD):
            status.config(text=t, fg=farg)

        def hamta():
            try:
                lista = self.bild_api('GET', '/admin/koder')['koder']
            except Exception as e:
                satt_status(str(e), ROD)
                return
            tree.delete(*tree.get_children())
            st['koder'] = {k['kod']: k for k in lista}
            for k in lista:
                skapad = datetime.datetime.fromtimestamp((k.get('skapad') or 0) / 1000).strftime('%Y-%m-%d') if k.get('skapad') else ''
                tree.insert('', 'end', iid=k['kod'], values=(k.get('namn', ''), k['kod'], 'Aktiv' if k.get('aktiv', True) else 'Avstängd', skapad))
            satt_status(f"{len(lista)} koder.")

        def vald():
            sel = tree.selection()
            return st['koder'].get(sel[0]) if sel else None

        def kopiera(kod):
            try:
                self.rot.clipboard_clear()
                self.rot.clipboard_append(kod)
            except Exception:
                pass

        def ny():
            import tkinter.simpledialog as sd
            namn = sd.askstring("Ny kod", "Namn på personen som ska få en kod:", parent=top)
            if not namn or not namn.strip():
                return
            try:
                r = self.bild_api('POST', '/admin/kod', {'namn': namn.strip()})
            except Exception as e:
                satt_status(str(e), ROD)
                return
            kopiera(r['kod'])
            hamta()
            messagebox.showinfo("Ny kod", f"Kod för {r['namn']}:\n\n{r['kod']}\n\nKoden är kopierad till urklipp. Ge den till personen.", parent=top)

        def vaxla():
            k = vald()
            if not k:
                messagebox.showinfo("Koder", "Markera en kod först.", parent=top)
                return
            try:
                self.bild_api('POST', '/admin/kod/status', {'kod': k['kod'], 'aktiv': not k.get('aktiv', True)})
            except Exception as e:
                satt_status(str(e), ROD)
                return
            hamta()

        def radera():
            k = vald()
            if not k:
                messagebox.showinfo("Koder", "Markera en kod först.", parent=top)
                return
            if not messagebox.askyesno("Radera", f"Radera koden för {k.get('namn', '')}? Den kan inte användas igen.", parent=top):
                return
            try:
                self.bild_api('DELETE', '/admin/kod?kod=' + k['kod'])
            except Exception as e:
                satt_status(str(e), ROD)
                return
            hamta()

        def kop():
            k = vald()
            if k:
                kopiera(k['kod'])
                satt_status(f"Koden för {k.get('namn', '')} kopierad.", GRON_MORK)

        Knapp(tb, "Hämta", hamta, BLA, '#ffffff', '#1e40af').pack(side='left', padx=(0, 8))
        Knapp(tb, "Ny kod…", ny, GRON_MORK, '#ffffff', '#14532d').pack(side='left', padx=(0, 8))
        Knapp(tb, "Kopiera kod", kop, '#64748b', '#ffffff', '#475569').pack(side='left', padx=(0, 8))
        Knapp(tb, "Stäng av / slå på", vaxla, '#f59e0b', '#ffffff', '#d97706').pack(side='left', padx=(0, 8))
        Knapp(tb, "Radera", radera, ROSA, '#ffffff', '#be185d').pack(side='left', padx=(0, 8))
        top.after(200, hamta)

    def bilder_visa_online(self):
        sel = self.bilder_tree.selection()
        if not sel:
            messagebox.showinfo("Välj", "Markera en artikel som har bild.", parent=self.rot)
            return
        varden = self.bilder_tree.item(sel[0], 'values')
        artnr = varden[0]
        if len(varden) > 2 and 'inte' in str(varden[2]):
            messagebox.showinfo("Ingen bild", "Artikel " + str(artnr) + " har ingen bild uppladdad än.", parent=self.rot)
            return
        import re
        artnr = re.sub(r'[^A-Za-z0-9]', '', artnr).upper()
        repo = self.repo.get().strip() if hasattr(self, 'repo') else U.las_installningar_dict().get('github_repo', '')
        gren = self.gren.get().strip() if hasattr(self, 'gren') else U.las_installningar_dict().get('github_branch', 'main')
        gren = gren or 'main'
        gh_path = self.bilder_gh_path_var.get().strip() if hasattr(self, 'bilder_gh_path_var') else 'bilder'
        gh_path = gh_path.strip('/') or 'bilder'
        if '/' in repo:
            owner, name = repo.split('/',1)
            raw_url = f"https://raw.githubusercontent.com/{repo}/{gren}/{gh_path}/{artnr}.webp"
            bas = "" if name.lower() == f"{owner.lower()}.github.io" else name + "/"
            pages_url = f"https://{owner.lower()}.github.io/{bas}{gh_path}/{artnr}.webp"
            url = raw_url
        else:
            url = f"{gh_path}/{artnr}.webp"
            pages_url = url
            raw_url = url
        import webbrowser
        webbrowser.open(url)
        self.bilder_status(f"Oppnar {url}  (404? Testa aven Pages: {pages_url} - nyuppladdade bilder tar 1-2 min)", BLA)









    def bilder_visa_orphans(self):
        # Punkt 3: visa bilder som finns online men saknar artikel
        try:
            import json, urllib.request
            repo = self.repo.get().strip() if hasattr(self, 'repo') else U.las_installningar_dict().get('github_repo', '')
            token = self.token.get().strip() if hasattr(self, 'token') else U.las_installningar_dict().get('github_token', '')
            gh_path = self.bilder_gh_path_var.get().strip() or 'bilder'
            if not repo or not token:
                messagebox.showinfo("GitHub saknas", "Fyll i repo/token.", parent=self.rot)
                return
            self.bilder_status("Hämtar lista för orphan-koll...", BLA)
            def fetch():
                try:
                    api_url = f"https://api.github.com/repos/{repo}/contents/{gh_path}"
                    req = urllib.request.Request(api_url, headers={'Authorization': f'Bearer {token}', 'Accept': 'application/vnd.github.v3+json'})
                    with urllib.request.urlopen(req) as r:
                        data = json.loads(r.read().decode('utf-8'))
                    online = {}
                    for f in data:
                        name = f.get('name','')
                        if name.lower().endswith('.webp') and not name.lower().endswith('_thumb.webp'):
                            artnr = name[:-5]
                            online[artnr] = f.get('download_url','')
                    orphans = [a for a in online if a not in (self._bilder_artikel_index or {})]
                    if not orphans:
                        self.ko.put(('bilder_log', "Inga orphan-bilder hittades"))
                    else:
                        self.ko.put(('bilder_log', f"{len(orphans)} orphan-bilder: {', '.join(orphans[:20])}"))
                        # Visa i separat dialog
                        self.ko.put(('bilder_orphans', (orphans, online)))
                except Exception as e:
                    self.ko.put(('bilder_log', f"Fel orphan-koll: {e}"))
            import threading
            threading.Thread(target=fetch, daemon=True).start()
        except Exception as e:
            self.bilder_status(f"Fel: {e}", ROD)


    def bilder_ta_bort_online(self):
        sel = self.bilder_tree.selection()
        if not sel:
            messagebox.showinfo("Välj", "Markera artiklar att ta bort bild för.", parent=self.rot)
            return
        artnrs = [self.bilder_tree.item(s,'values')[0] for s in sel]
        if not messagebox.askyesno("Ta bort?", f"Ta bort {len(artnrs)} bilder från GitHub?", parent=self.rot):
            return
        repo = self.repo.get().strip() if hasattr(self, 'repo') else U.las_installningar_dict().get('github_repo', '')
        token = self.token.get().strip() if hasattr(self, 'token') else U.las_installningar_dict().get('github_token', '')
        branch = self.gren.get().strip() if hasattr(self, 'gren') else U.las_installningar_dict().get('github_branch', 'main')
        gh_path = self.bilder_gh_path_var.get().strip() or 'bilder'
        def jobb():
            import json, urllib.request, urllib.error
            for artnr in artnrs:
                try:
                    remote_path = f"{gh_path.strip('/')}/{artnr}.webp"
                    api_url = f"https://api.github.com/repos/{repo}/contents/{remote_path}"
                    req_get = urllib.request.Request(api_url, headers={'Authorization': f'Bearer {token}', 'Accept': 'application/vnd.github.v3+json'})
                    with urllib.request.urlopen(req_get) as r:
                        j = json.loads(r.read().decode('utf-8'))
                        sha = j.get('sha')
                    payload = {"message": f"Ta bort bild {artnr}.webp", "sha": sha, "branch": branch}
                    req_del = urllib.request.Request(api_url, data=json.dumps(payload).encode('utf-8'), headers={'Authorization': f'Bearer {token}', 'Accept': 'application/vnd.github.v3+json'}, method='DELETE')
                    with urllib.request.urlopen(req_del) as r:
                        r.read()
                    self._bilder_github_bilder.discard(artnr)
                    self.ko.put(('bilder_log', f"Borttagen {artnr}"))
                except Exception as e:
                    self.ko.put(('bilder_log', f"Fel ta bort {artnr}: {e}"))
            self.ko.put(('bilder_refresh', True))
        import threading
        threading.Thread(target=jobb, daemon=True).start()

    def bilder_poll_queue(self):
        try:
            while True:
                typ, data = self.ko.get_nowait()
                if typ == 'bilder_log':
                    self.bilder_status(data, TEXT)
                elif typ == 'bilder_index':
                    self._bilder_artikel_index = data
                elif typ == 'bilder_refresh':
                    self.bilder_bygg_lista_fran_index()
        except queue.Empty:
            pass
        if hasattr(self, 'bilder_tree'):
            self.rot.after(800, self.bilder_poll_queue)


    def backup_data_js(self, ut_sokvag):
        """Auto-backup"""
        try:
            import shutil
            from datetime import datetime
            if not ut_sokvag or not os.path.isfile(ut_sokvag):
                return None
            dir_name = os.path.dirname(ut_sokvag)
            base = os.path.basename(ut_sokvag)
            backup_dir = os.path.join(dir_name, "backups")
            os.makedirs(backup_dir, exist_ok=True)
            ts = datetime.now().strftime("%Y-%m-%d_%H%M%S")
            backup_name = f"{os.path.splitext(base)[0]}_{ts}.js"
            backup_path = os.path.join(backup_dir, backup_name)
            shutil.copy2(ut_sokvag, backup_path)
            # Håll bara 20 senaste
            filer = sorted(glob.glob(os.path.join(backup_dir, "*.js")), key=os.path.getmtime, reverse=True)
            for gammal in filer[20:]:
                try:
                    os.remove(gammal)
                except Exception:
                    pass
            return backup_path
        except Exception as e:
            print("Backup fel:", e)
            return None

    # ---------------------------------------------------------------- hjälpare
    def status_rad(self, text, farg=DAMPAD):
        self.fot_etikett.config(text=text, fg=farg)

    def uppdatera_status(self):
        if not hasattr(self, 'st_bulk'):
            return
        mon = '*.pdf'  # endast PDF för fabrikssterilt
        for etikett, (t, c) in ((self.st_forbruk, hitta(self.p_forbruk.get(), mon)),
                                (self.st_galler, hitta_galler(self.p_galler.get())),
                                (self.st_bulk, hitta(self.p_bulk.get(), '*.xlsx'))):
            etikett.config(text="●  " + t, fg=c)
        if self.publ.get() == 'api':
            self.gh.pack(fill='x')
        else:
            self.gh.pack_forget()

    def skriv_logg(self, text, tagg='info'):
        self.logg.config(state='normal')
        self.logg.insert('end', text, tagg)
        self.logg.see('end')
        self.logg.config(state='disabled')

    def satt_framsteg(self, andel):
        self.prog['value'] = max(0, min(1, andel)) * 1000

    # ---------------------------------------------------------------- GitHub-test
    def testa(self):
        self.test_res.config(text="Testar…", fg=DAMPAD)
        repo, tok, gren = self.repo.get().strip(), self.token.get().strip(), self.gren.get().strip() or 'main'

        def jobb():
            r = U.testa_github(repo, tok, gren)
            self.ko.put(('test', r))
        threading.Thread(target=jobb, daemon=True).start()

    # ---------------------------------------------------------------- körning
    def kontrollera(self, torr):
        fel = []
        if self.src.get() == 'pdf':
            t, c = hitta(self.p_forbruk.get(), '*.pdf')
        else:
            t, c = hitta(self.p_forbruk.get(), '*.xlsx')
        if c in (ROD, DAMPAD):
            fel.append("Fabrikssterila artiklar: " + t)
        t, c = hitta_galler(self.p_galler.get())
        if c in (ROD, DAMPAD):
            fel.append("Instrumentgaller: " + t)
        t, c = hitta(self.p_bulk.get(), '*.xlsx')
        if c in (ROD, DAMPAD):
            fel.append("Bulkplatser: " + t)
        if not self.p_ut.get().strip():
            fel.append("Du måste ange var data.js ska sparas.")
        try:
            float(self.sparr.get().replace(',', '.') or '0')
        except ValueError:
            fel.append("Spärren ska vara ett tal, t.ex. 20.")
        if not torr and self.publ.get() == 'api' and (not self.repo.get().strip() or not self.token.get().strip()):
            fel.append("Sparad GitHub-anslutning saknas (arkiv/åtkomstnyckel i installningar.dat) – välj Bara spara data.js istället.")
        return fel

    def starta(self, torr):
        if self.kor_pagar:
            return
        fel = self.kontrollera(torr)
        if fel:
            messagebox.showwarning("Något saknas", "Rätta detta först:\n\n• " + "\n• ".join(fel), parent=self.rot)
            return
        self.spara(tyst=True)
        d = self.till_dict()
        a = argparse.Namespace(
            pdf=d.get('pdf', '') or None, artiklar=d.get('artiklar', '') or None, galler=d['galler'],
            bulk=d.get('bulk', ''), ut=d['ut'], torrkorning=torr,
            sparr_procent=float(d['sparr_procent'].replace(',', '.')), tvinga=False,
            publicera=d['publicera'] == 'ja', publicera_api=d['publicera'] == 'api',
            github_repo=d['github_repo'], github_token=d['github_token'], github_branch=d['github_branch'],
            github_api='https://api.github.com')
        self.kor_pagar = True
        self.avbryt = False
        logga("Körning startas (" + ("torr" if torr else "skarp") + "): " + ", ".join(f"{k}={v}" for k, v in d.items() if k != 'github_token'))
        self.res.pack_forget()
        self.res_info.pack_forget()
        self.logg.config(state='normal')
        self.logg.delete('1.0', 'end')
        self.logg.config(state='disabled')
        self.satt_framsteg(0)
        self.steg_nr, self.steg_tot = 0, 1
        self.steg_etikett.config(text="Startar…", fg=TEXT)
        self.b_kor.grid_forget()
        self.b_torr.grid_forget()
        self.b_spara.grid_forget()
        self.b_avbryt.grid(row=0, column=0, columnspan=3, sticky='ew', padx=0, pady=4)
        threading.Thread(target=self.jobb, args=(a,), daemon=True).start()

    def avbryt_kor(self):
        self.avbryt = True
        logga("Användaren tryckte Avbryt")
        self.steg_etikett.config(text="Avbryter…", fg=GUL)


    def jobb(self, a):
        q = self.ko
        import traceback

        def steg(nr, tot, text):
            q.put(('steg', nr, tot, text))

        def prog(i, n, text=''):
            if not self.avbryt:
                q.put(('prog', i, n, text))

        def ok(text):
            q.put(('ok', text))

        def info(text):
            q.put(('info', text))

        def varning(text):
            q.put(('varning', text))

        def fel(text):
            q.put(('fel', text))

        def sparr(problem, handelse, svar):
            q.put(('sparr', problem, handelse, svar))

        def bekrafta(problem):
            handelse = threading.Event()
            svar = {}
            sparr(problem, handelse, svar)
            handelse.wait()
            if not svar.get('fortsatt'):
                raise SystemExit("Avbrot: databasen hade blivit mycket mindre.")
            q.put(('varning', "Du valde att fortsatta anda."))

        if not getattr(a, 'torrkorning', False):
            try:
                q.put(('info', 'Skapar backup av befintlig data.js...'))
                bp = self.backup_data_js(a.ut)
                if bp:
                    q.put(('ok', 'Backup sparad: ' + os.path.basename(bp)))
            except Exception as e:
                q.put(('varning', 'Kunde inte skapa backup: ' + str(e)))

        try:
            q.put(('info', 'Startar ' + ('torrkorning' if getattr(a,'torrkorning',False) else 'uppdatering') + '...'))
            if getattr(a,'pdf',None) and a.pdf:
                q.put(('info', 'Kalla PDF: ' + str(a.pdf)))
            if getattr(a,'galler',None) and a.galler:
                q.put(('info', 'Galler: ' + str(a.galler)))
            if getattr(a,'bulk',None) and a.bulk:
                q.put(('info', 'Bulk: ' + str(a.bulk)))

            try:
                res = U.kor(a, bekrafta_fn=bekrafta, steg_cb=steg, prog_cb=prog, ok_cb=ok, info_cb=info, varning_cb=varning, fel_cb=fel,
                            avbryt_fn=lambda: self.avbryt)
            except TypeError:
                try:
                    res = U.kor(a, steg, prog, ok, info, varning, fel, bekrafta)
                except TypeError:
                    q.put(('info', 'Kor med aldre motor utan live-progress'))
                    res = U.kor(a, bekrafta)

            q.put(('klar', res))

        except SystemExit as e:
            q.put(('stopp', str(e.code) if e.code not in (None, 0) else 'Avbruten.'))
        except Exception as e:
            tb = traceback.format_exc()
            q.put(('stopp', "Nagot gick fel: " + str(e) + chr(10) + tb))

    def pumpa(self):
        try:
            while True:
                try:
                    k, *rest = self.ko.get_nowait()
                except queue.Empty:
                    break
                if k == 'logg':
                    self.skriv_logg(rest[0], rest[1] if len(rest)>1 else 'info')
                elif k == 'steg':
                    # rest: nr, tot, text
                    if len(rest) >= 3:
                        self.h_steg(rest[0], rest[1], rest[2])
                    else:
                        self.steg_etikett.config(text=str(rest[0]) if rest else "")
                elif k == 'prog' or k == 'framsteg':
                    # rest: i, n, text
                    if len(rest) >= 2:
                        self.h_prog(rest[0], rest[1], rest[2] if len(rest) > 2 else "")
                elif k == 'ok':
                    self.h_ok(rest[0] if rest else "")
                elif k == 'info':
                    self.h_info(rest[0] if rest else "")
                elif k == 'varning':
                    self.h_varning(rest[0] if rest else "")
                elif k == 'fel':
                    self.h_fel(rest[0] if rest else "")
                elif k == 'sparr':
                    # rest: problem, handelse, svar
                    if len(rest) >= 3:
                        self.h_sparr(rest[0], rest[1], rest[2])
                elif k == 'klar':
                    self.h_klar(rest[0])
                elif k == 'stopp':
                    self.h_stopp(rest[0])
                elif k == 'stat':
                    self.h_stat(rest[0], rest[1] if len(rest)>1 else None)
                elif k == 'test':
                    self.h_test(rest[0])
                elif k == 'bilder_log':        # BILDER-fliken delar kö med huvudfönstret: båda måste hantera dess meddelanden
                    self.bilder_status(rest[0], TEXT)
                elif k == 'bilder_index':
                    self._bilder_artikel_index = rest[0]
                elif k == 'bilder_refresh':
                    self.bilder_bygg_lista_fran_index()
        except queue.Empty:
            pass
        except Exception as e:
            logga(f"Pumpa fel: {e!r}")
        finally:
            self.rot.after(60, self.pumpa)

    # ---- händelser från körningen
    def h_steg(self, nr, tot, text):
        self.steg_nr, self.steg_tot = nr, tot
        self.steg_etikett.config(text=f"Steg {nr} av {tot}:  {text}", fg=TEXT)
        self.satt_framsteg((nr - 1) / tot)
        self.skriv_logg(f"[{nr}/{tot}] {text}\n", 'steg')

    def h_prog(self, i, n, text):
        if n > 0:
            self.satt_framsteg((self.steg_nr - 1 + i / n) / self.steg_tot)
            self.steg_etikett.config(text=f"Steg {self.steg_nr} av {self.steg_tot}:  {i} av {n} {text}")

    def h_ok(self, t):
        self.skriv_logg("   ✓ " + t + "\n", 'ok')

    def h_info(self, t):
        self.skriv_logg("   " + t + "\n", 'info')

    def h_varning(self, t):
        self.skriv_logg("   ! " + t + "\n", 'varning')

    def h_fel(self, t):
        self.skriv_logg("   ✗ " + t + "\n", 'fel')

    def h_test(self, r):
        ok, text = r
        self.test_res.config(text=("✓ " if ok else "✗ ") + text, fg=GRON_MORK if ok else ROD)

    def h_sparr(self, problem, handelse, svar):
        d = tk.Toplevel(self.rot)
        d.title("Spärr")
        d.configure(bg=KORT)
        d.transient(self.rot)
        d.resizable(False, False)
        topp = tk.Frame(d, bg=GUL)
        topp.pack(fill='x')
        tk.Label(topp, text="Databasen har blivit mycket mindre", bg=GUL, fg='#ffffff', font=(FONT_FET, 13, 'bold'),
                 padx=20, pady=12).pack(anchor='w')
        h = tk.Frame(d, bg=KORT)
        h.pack(padx=22, pady=16)
        tk.Label(h, text="\n".join("•  " + p for p in problem), bg=KORT, fg=TEXT, font=self.f_norm, justify='left',
                 wraplength=470, anchor='w').pack(anchor='w')
        tk.Label(h, text="\nDet kan betyda att fel fil eller mapp valts, eller att en flik saknas.\nDen gamla databasen ligger kvar om du avbryter.",
                 bg=KORT, fg=DAMPAD, font=self.f_norm, justify='left', anchor='w').pack(anchor='w')
        kn = tk.Frame(h, bg=KORT)
        kn.pack(fill='x', pady=(18, 0))

        def klart(fortsatt):
            svar['fortsatt'] = fortsatt
            d.grab_release()
            d.destroy()
            handelse.set()
        Knapp(kn, "Avbryt (rekommenderas)", lambda: klart(False), GRON_MORK, hover='#4f7a24', stor=True).pack(side='left')
        Knapp(kn, "Fortsätt ändå", lambda: klart(True), '#e5e9f0', TEXT, '#d5dbe5', stor=True).pack(side='right')
        d.protocol("WM_DELETE_WINDOW", lambda: klart(False))
        d.update_idletasks()
        x = self.rot.winfo_rootx() + (self.rot.winfo_width() - d.winfo_width()) // 2
        y = self.rot.winfo_rooty() + (self.rot.winfo_height() - d.winfo_height()) // 3
        d.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        d.grab_set()
        d.focus_set()

    def klar_ui(self):
        self.kor_pagar = False
        self.b_avbryt.grid_forget()
        self.b_kor.grid(row=0, column=0, sticky='ew', padx=(0,6), pady=4)
        self.b_torr.grid(row=0, column=1, sticky='ew', padx=3, pady=4)
        self.b_spara.grid(row=0, column=2, sticky='ew', padx=(6,0), pady=4)

    def visa_diff_dialog(self, r):
        """Diff-vy före uppdatering"""
        forra = r.get('forra')
        if not forra:
            return
        try:
            gammal_f, gammal_g = forra[0], forra[1]
        except Exception:
            return
        ny_f, ny_g, ny_b = r['forbruk'], r['galler'], r['bulk']
        diff_f = ny_f - gammal_f
        diff_g = ny_g - gammal_g
        total_gammal = gammal_f + gammal_g
        total_ny = ny_f + ny_g
        total_diff = total_ny - total_gammal
        pct = (total_diff / total_gammal * 100) if total_gammal else 0

        d = tk.Toplevel(self.rot)
        d.title(f"Diff före uppdatering - {VERSION}")
        d.configure(bg=KORT)
        d.transient(self.rot)
        d.resizable(False, False)

        # Header
        topp = tk.Frame(d, bg=MORK)
        topp.pack(fill='x')
        tk.Label(topp, text="Jämförelse före uppdatering", bg=MORK, fg='white', font=(FONT_FET, 14, 'bold'), padx=20, pady=(14,2)).pack(anchor='w')
        tk.Label(topp, text=f"Förra data.js  →  Ny torrkörning  ({r.get('stampel','')})", bg=MORK, fg='#9fb0c3', font=self.f_norm, padx=20, pady=(0,12)).pack(anchor='w')

        body = tk.Frame(d, bg=KORT)
        body.pack(fill='both', expand=True, padx=20, pady=16)

        # Tabell
        hdr = tk.Frame(body, bg=KORT)
        hdr.pack(fill='x', pady=(0,6))
        for col, txt in enumerate(["Kategori", "Förut", "Nu", "Diff"]):
            tk.Label(hdr, text=txt, bg=KORT, fg=DAMPAD, font=self.f_liten, width=12, anchor='w' if col==0 else 'center').grid(row=0, column=col, sticky='w', padx=4)

        tk.Frame(body, bg=RAM, height=1).pack(fill='x', pady=4)

        def rad(cat, gammal, ny, diff, farg=TEXT):
            fr = tk.Frame(body, bg=KORT)
            fr.pack(fill='x', pady=3)
            tk.Label(fr, text=cat, bg=KORT, fg=TEXT, font=self.f_fet, width=14, anchor='w').pack(side='left')
            tk.Label(fr, text=str(gammal), bg=KORT, fg=TEXT, font=self.f_norm, width=10).pack(side='left')
            tk.Label(fr, text=str(ny), bg=KORT, fg=TEXT, font=self.f_fet, width=10).pack(side='left')
            txt = f"{'+' if diff>0 else ''}{diff}"
            col = GRON_MORK if diff>0 else ROD if diff<0 else DAMPAD
            if abs(diff) > 20:
                col = ROD if diff<0 else GRON_MORK
            tk.Label(fr, text=txt, bg=KORT, fg=col, font=self.f_fet, width=10).pack(side='left')
            return fr

        rad("Fabrikssterilt", gammal_f, ny_f, diff_f, GRON)
        rad("Instrumentgaller", gammal_g, ny_g, diff_g, BLA)
        # Bulk alltid med
        gammal_b = forra[2] if forra and len(forra)>2 and forra[2] is not None else 0
        rad("Bulk", gammal_b, ny_b, ny_b-gammal_b, ROSA)

        tk.Frame(body, bg=RAM, height=1).pack(fill='x', pady=8)

        # Totalt
        tot_fr = tk.Frame(body, bg='#f1f5f9')
        tot_fr.pack(fill='x', pady=4, ipady=6)
        tk.Label(tot_fr, text="Totalt", bg='#f1f5f9', fg=TEXT, font=self.f_rubrik, width=14, anchor='w').pack(side='left', padx=8)
        tk.Label(tot_fr, text=str(total_gammal), bg='#f1f5f9', fg=TEXT, font=self.f_norm, width=10).pack(side='left')
        tk.Label(tot_fr, text=str(total_ny), bg='#f1f5f9', fg=TEXT, font=self.f_rubrik, width=10).pack(side='left')
        col = ROD if pct < -10 else GRON_MORK if pct>0 else DAMPAD
        tk.Label(tot_fr, text=f"{'+' if total_diff>0 else ''}{total_diff} ({pct:+.1f}%)", bg='#f1f5f9', fg=col, font=self.f_rubrik, width=18).pack(side='left')

        # Varning om stor minskning
        if pct < -5:
            warn = tk.Frame(body, bg='#fef2f2')
            warn.pack(fill='x', pady=(12,0), ipady=8)
            tk.Label(warn, text=f"⚠️  Stor minskning: {pct:.1f}% färre artiklar. Kolla att du valt rätt mapp!", bg='#fef2f2', fg=ROD, font=self.f_fet, anchor='w').pack(fill='x', padx=10)

        n_avv = len(r.get('avvik', []))
        if n_avv:
            av = tk.Frame(body, bg='#fffbeb')
            av.pack(fill='x', pady=(10,0), ipady=6)
            tk.Label(av, text=f"⚠️ {n_avv} avvikelser hittades – kolla fliken AVVIKELSER ({n_avv})", bg='#fffbeb', fg=GUL, font=self.f_norm, anchor='w').pack(fill='x', padx=10)

        # Knappar
        kn = tk.Frame(body, bg=KORT)
        kn.pack(fill='x', pady=(18,0))
        def close():
            d.grab_release()
            d.destroy()

        Knapp(kn, "Stäng och granska", close, '#e5e9f0', TEXT, '#d5dbe5').pack(side='left')
        Knapp(kn, "Uppdatera nu", lambda: (close(), self.starta(False)), GRON_MORK, hover='#4f7a24', stor=True).pack(side='right')

        d.update_idletasks()
        x = self.rot.winfo_rootx() + (self.rot.winfo_width() - d.winfo_width()) // 2
        y = self.rot.winfo_rooty() + (self.rot.winfo_height() - d.winfo_height()) // 3
        d.geometry(f"+{max(x,0)}+{max(y,0)}")
        d.grab_set()
        d.focus_set()

    def h_klar(self, r):
        self.klar_ui()
        self.satt_framsteg(1)
        torr = r['torr']
        forra = r.get('forra')
        vardena = (r['forbruk'], r['galler'], r['bulk'])
        gamla = (forra[0], forra[1], None) if forra else (None, None, None)
        for (n, d), nytt, gammalt in zip(self.res_tiles, vardena, gamla):
            n.config(text=str(nytt))
            if gammalt is None:
                d.config(text=" ", fg=DAMPAD)
            else:
                diff = nytt - gammalt
                d.config(text=f"förut {gammalt}  ({'+' if diff > 0 else ''}{diff})", fg=DAMPAD if diff == 0 else (GRON_MORK if diff > 0 else ROD))
        self.res.pack(fill='x', pady=(0, 8), before=self._res_anker)
        for w in (self.avv_lank, self.app_lank):
            w.pack_forget()
        n_avv = len(r['avvik'])
        self.avv_etikett.config(text=("Avvikelser: " if n_avv else "Inga avvikelser. ") + (str(n_avv) + "  " if n_avv else ""),
                                fg=GUL if n_avv else TEXT)
        if n_avv and not torr:
            self.avv_lank.pack(side='left')
            fil = os.path.join(os.path.dirname(r['ut']), 'avvikelser.txt')
            self.avv_lank.bind('<Button-1>', lambda e: oppna(fil))
        if r['adress'] and not torr:
            self.app_lank.config(text="Öppna appen  " + r['adress'])
            self.app_lank.pack(side='right')
            self.app_lank.bind('<Button-1>', lambda e: webbrowser.open(r['adress']))
        self.res_info.pack(fill='x', pady=(0, 8), before=self._res_anker)

        try:
            self.visa_avvikelser_i_flik(r.get('avvik', []))
        except Exception:
            pass

        if torr:
            self.steg_etikett.config(text="Torrkörning klar - inget skrevs eller skickades.", fg=BLA)
            self.skriv_logg("\nTorrkörning klar. Tryck på Uppdatera databasen när det ser rätt ut.\n", 'steg')
            if n_avv:
                self.skriv_logg(f"Se fliken Avvikelser för {n_avv} avvikelser uppdelade per kategori.\n", 'varning')
            # Visa diff-dialog automatiskt om det finns tidigare data
            if forra:
                try:
                    self.rot.after(300, lambda: self.visa_diff_dialog(r))
                except Exception:
                    pass
        else:
            self.steg_etikett.config(text="Klart! Databasen är uppdaterad.", fg=GRON_MORK)
            self.skriv_logg(f"\nKLART {r['stampel']}  -  {r['forbruk']} fabrikssterila, {r['galler']} galler, {r['bulk']} med bulkplats.\n", 'ok')

    def h_stopp(self, text):
        self.klar_ui()
        self.satt_framsteg(0)
        self.steg_etikett.config(text="Stoppad - inget har ändrats.", fg=ROD)
        self.skriv_logg("\n✗ " + text + "\n", 'fel')

    def spara_vid_stangning(self):
        """Sparar inställningarna utan dialogrutor. Får aldrig hindra att programmet stängs."""
        try:
            U.skriv_installningar(self.till_dict())
            return True
        except Exception:
            return False

    def stang(self):
        logga("Stänger programmet")
        self.spara_vid_stangning()
        if self.kor_pagar and not messagebox.askyesno("Uppdatering pågår", "Uppdateringen pågår. Vill du avsluta ändå?", parent=self.rot):
            return
        self.spara_vid_stangning()
        try:
            self.avbryt = True
            self.kor_pagar = False
            try:
                while not self.ko.empty():
                    self.ko.get_nowait()
            except Exception:
                pass
            self.rot.quit()
            self.rot.destroy()
        except Exception as e:
            logga(f"Fel vid stängning: {e!r}")
            try:
                self.rot.destroy()
            except Exception:
                pass


def main():
    dpi_pa()
    logg_start()
    import traceback
    sys.excepthook = lambda t, v, tb: logga("OHANTERAT FEL:\n" + "".join(traceback.format_exception(t, v, tb)))
    try:
        _las_installningar_dict_fix()
    except Exception:
        pass
    rot = tk.Tk()
    rot.report_callback_exception = lambda t, v, tb: logga("FEL I FÖNSTRET:\n" + "".join(traceback.format_exception(t, v, tb)))
    rot.withdraw()  # göm huvudfönstret medan splash visas

    def maximera():
        try:
            rot.state('zoomed')                # Windows
        except tk.TclError:
            try:
                rot.attributes('-zoomed', True)    # Linux
            except Exception:
                pass

    def start_app():
        rot.deiconify()
        App(rot)
        maximera()                              # alltid maximerat vid start (efter att App satt sin startstorlek)
        rot.after(150, maximera)

    splash = SplashScreen(rot, start_app)
    # Startbilden från exe-filen (PyInstaller --splash) täcker uppstarten. Den stängs först när
    # programmets egen startbild har ritats, så det aldrig blir en tom lucka emellan.
    try:
        splash.update()
        import pyi_splash  # finns bara i den byggda exe-filen
        pyi_splash.close()
    except Exception:
        pass
    rot.mainloop()


if __name__ == '__main__':
    main()
