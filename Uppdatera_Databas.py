#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Bygger databasfilen data.js åt sökappen (index.html) från Excel.

  * Förbrukningsartiklar: antingen PDF-filen med produktetiketterna (--pdf)
    eller en Excel-fil med en lista (--artiklar: Artnr, Lev. artnr, Produktnamn, Plats)
  * Bulkplatser (valfritt): Excel-fil med Artikelnummer, Plockplats, Bulkplats (--bulk)
  * Galler: en mapp med Excel-filer, en flik per reol med rubriken "Rad X Reol Y"

Användning:
    python uppdatera_databas.py --pdf lager.pdf --galler C:\\Galler --ut data.js
    python uppdatera_databas.py --artiklar lager_artiklar.xlsx --galler C:\\Galler --ut data.js
    python uppdatera_databas.py ... --torrkorning      (skriver ingenting, visar bara resultatet)

Krav (en gång):  pip install openpyxl pdfplumber

Skriptet skriver data.js (och sparar förra versionen som data.js.bak) samt avvikelser.txt
bredvid, med allt som hoppades över eller ser konstigt ut.
"""
import argparse, base64, glob, json, os, re, shutil, subprocess, sys, tempfile, pathlib
import urllib.error, urllib.request
from pathlib import Path

VERSION = 'v1.0'
from datetime import datetime

try:
    from openpyxl import load_workbook
except ImportError:
    sys.exit("openpyxl saknas. Kör:  pip install openpyxl")

HEADER = re.compile(r'rad\s*(\d+)\s*reol\s*(\d+)', re.I)
# PLATS: tillåt A-01-02-03, A-1-2-3, A010203, A 01 02 03, med eller utan bindestreck, 1-2 siffror
PLATS = re.compile(r'^[A-Z]\s*[-\s]?\s*\d{1,2}\s*[-\s]?\s*\d{1,2}\s*[-\s]?\s*\d{1,2}$', re.I)

def normalisera_plats(plats: str) -> str:
    """Normaliserar A-1-2-3, A 01 02 03, A010203 -> A-01-02-03."""
    if not plats:
        return ''
    import re
    p = plats.strip().upper()
    # Ta bort alla mellanslag och bindestreck, behåll första bokstav + siffror
    m = re.match(r'^([A-Z])\s*[-\s]*([\d\s\-]+)$', p)
    if not m:
        return p
    bokstav = m.group(1)
    siffror = re.findall(r'\d+', m.group(2))
    if len(siffror) >= 3:
        return f"{bokstav}-{int(siffror[0]):02d}-{int(siffror[1]):02d}-{int(siffror[2]):02d}"
    elif len(siffror) == 1 and len(siffror[0]) >= 6:
        s = siffror[0]
        return f"{bokstav}-{s[0:2]}-{s[2:4]}-{s[4:6]}"
    return p

PLATS_STRICT = re.compile(r'^[A-Z]-\d{2}-\d{2}-\d{2}$')
# BULK: BG5, BG-05, BPA-27, BULK-1, med/utan bindestreck och nolla
BULK = re.compile(r'^[A-Z]{1,4}\s*[-\s]?\s*\d{1,3}([A-Z])?(\s*[-\s]?\s*\d{1,3})*$', re.I)
BULK_NORM = re.compile(r'([A-Z]+)[-\s]*(\d+)', re.I)


# ---------------------------------------------------------------- utseende i fönstret
FARG = False
BREDD = 62


def starta_konsol():
    """Slår på färger och UTF-8 i Windows-konsolen. Faller tillbaka till vanlig text om det inte går."""
    global FARG
    ok = True
    try:
        if os.name == 'nt':
            import ctypes
            k = ctypes.windll.kernel32
            k.SetConsoleOutputCP(65001)
            h = k.GetStdHandle(-11)
            lage = ctypes.c_ulong()
            ok = bool(k.GetConsoleMode(h, ctypes.byref(lage))) and bool(k.SetConsoleMode(h, lage.value | 4))
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        ok = False
    FARG = bool(ok and (sys.stdout.isatty() or os.environ.get('FORCE_COLOR')) and not os.environ.get('NO_COLOR'))


def f(text, *koder):
    return ("\033[" + ";".join(koder) + "m" + text + "\033[0m") if FARG else text


GRON, GUL, ROD, CYAN, GRA, FET = '32', '33', '31', '36', '90', '1'
BANNER = [
    ' ___  _____  ___  ___  ___  _     ___  _  _ ',
    '/ __||_   _|| __|| _ \\|_ _|| |   | __|| \\| |',
    '\\__ \\  | |  | _| |   / | | | |__ | _| | .` |',
    '|___/  |_|  |___||_|_\\|___||____||___||_|\\_|',
]


def visa_banner(version):
    print()
    for rad in BANNER:
        print("  " + f(rad, CYAN, FET))
    print("  " + f("Lagersok - databasuppdatering", FET) + f("   " + version, GRA))
    print("  " + f("-" * (BREDD - 2), GRA))
    print()


def steg(nr, tot, text):
    print(f(f"[{nr}/{tot}] ", CYAN, FET) + f(text, FET))


def ok(text):
    print("      " + f("[ OK ]", GRON, FET) + " " + text)


def varning(text):
    print("      " + f("[ !! ]", GUL, FET) + " " + f(text, GUL))


def info(text):
    print("      " + f(text, GRA))


def fel(text):
    print()
    print(f("  [ FEL ] ", ROD, FET) + f(str(text), ROD))


def framsteg(i, n, text=""):
    """Enkel förloppsmätare (bara i ett riktigt konsolfönster)."""
    if not sys.stdout.isatty() or n <= 0:
        return
    b = 24
    k = int(b * i / n)
    rad = f"      [{f('#' * k, GRON)}{f('-' * (b - k), GRA)}] {i}/{n} {text}"
    print("\r" + rad[:BREDD + 60].ljust(BREDD + 20), end='', flush=True)
    if i >= n:
        print("\r" + " " * (BREDD + 40), end='\r', flush=True)


def synlig(t):
    return len(re.sub(r'\x1b\[[0-9;]*m', '', t))


def ruta(rader, farg=GRON):
    """Ritar en ram med ASCII-tecken runt några rader text."""
    inre = max(synlig(r) for r in rader) + 4
    kant = "+" + "-" * inre + "+"
    print("  " + f(kant, farg))
    for r in rader:
        print("  " + f("|", farg) + "  " + r + " " * (inre - 4 - synlig(r)) + "  " + f("|", farg))
    print("  " + f(kant, farg))


def text(v) -> str:
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).replace('\xa0', ' ').replace(';', ',')
    return re.sub(r'\s+', ' ', s).strip()


def heltal(v) -> int | None:
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)) and float(v).is_integer():
        return int(v)
    if isinstance(v, str) and v.strip().isdigit():
        return int(v.strip())
    return None


def las_pdf(fil: str, avvik: list, framsteg=None) -> list:
    """Läser etikettsidorna: tre etiketter per rad, varje med Artnr., Lev.artnr.,
    namn (1-3 rader), Art.typ och Plats. Robust mot små X-förskjutningar."""
    try:
        import pdfplumber
    except ImportError:
        sys.exit("pdfplumber saknas. Kör:  pip install pdfplumber")
    from typing import List, Dict
    etiketter: List[Dict] = []
    with pdfplumber.open(fil) as pdf:
        for sida, pg in enumerate(pdf.pages, 1):
            if framsteg:
                framsteg(sida, len(pdf.pages), 'sidor')
            # Öka tolerans för att fånga ord som ligger nära varandra
            ord_ = pg.extract_words(keep_blank_chars=True, x_tolerance=2.5, y_tolerance=2, use_text_flow=True, extra_attrs=["size"])
            # Hitta ankare - tillåt både Artnr. och Artnr (utan punkt) samt case-insensitive
            ankare_raw = []
            for w in ord_:
                txt = w['text'].strip().lower().rstrip('.')
                if txt in ('artnr', 'art.nr'):
                    ankare_raw.append(w['x0'])
            # Klustra ankare som ligger inom 20px som samma kolumn (fixar 1px glapp)
            ankare_raw.sort()
            klustrade = []
            for x in ankare_raw:
                if not klustrade or abs(x - klustrade[-1]) > 20:
                    klustrade.append(x)
                else:
                    # medelvärde
                    klustrade[-1] = (klustrade[-1] + x) / 2
            ankare = sorted({round(x) for x in klustrade})
            if not ankare:
                continue
            kolumner = [[] for _ in ankare]
            for w in ord_:
                # Hitta närmaste ankare till vänster, med tolerans 10px
                idx = [i for i, a in enumerate(ankare) if w['x0'] >= a - 10]
                if idx:
                    kolumner[idx[-1]].append(w)
            for kol in kolumner:
                rader = []
                for w in sorted(kol, key=lambda w: w['top']):
                    if rader and abs(rader[-1][0] - w['top']) < 3:
                        rader[-1][1].append(w)
                    else:
                        rader.append([w['top'], [w]])
                cur = None
                for _, ws in rader:
                    ws.sort(key=lambda w: w['x0'])
                    forst = ws[0]['text']
                    rest = ' '.join(w['text'] for w in ws[1:]).strip()
                    if forst == 'Artnr.':
                        cur = {'sida': sida, 'art': rest, 'lev': '', 'namn': [], 'plats': ''}
                        etiketter.append(cur)
                    elif cur is None:
                        continue
                    elif forst == 'Lev.artnr.':
                        cur['lev'] = rest
                    elif forst == 'Plats':
                        cur['plats'] = rest
                    elif forst == 'Art.typ:':
                        pass
                    else:
                        cur['namn'].append(''.join(w['text'] for w in ws))
    ut, sedda = [], set()
    for e in etiketter:
        art, lev, plats = text(e['art']), text(e['lev']), text(e['plats'])
        namn = text(''.join(e['namn']))   # radbrytningarna ligger mitt i ord, så de läggs ihop utan mellanslag
        ref = f"sida {e['sida']}: {art} {namn}"
        if not namn or not art:
            avvik.append(f"PDF {ref} - saknar artikelnummer eller namn - hoppades över")
            continue
        plats_norm = normalisera_plats(plats)
        if not PLATS.match(plats_norm) and not PLATS_STRICT.match(plats_norm):
            if plats and not re.search(r'[A-Za-z]-\d', plats):
                avvik.append(f"PDF utan hyllplats {ref} - plats '{plats}' - hoppades över")
            else:
                avvik.append(f"PDF {ref} - ogiltig plats '{plats}' - hoppades över")
            continue
        k = (art, lev, namn.lower(), plats)
        if k in sedda:
            avvik.append(f"PDF {ref} - dubblett - hoppades över")
            continue
        sedda.add(k)
        ut.append(f"{art};{lev or '-'};{namn};{plats}")
    if not etiketter:
        sys.exit("Hittade inga etiketter i PDF-filen.")
    return ut


def las_artiklar(fil, avvik):
    wb = load_workbook(fil, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    rader = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rader:
        sys.exit("Artikelfilen är tom.")
    rubrik = [text(c).lower() for c in rader[0]]
    def kol(namn, standard):
        for i, r in enumerate(rubrik):
            if namn in r:
                return i
        return standard
    ci = (kol('artnr', 0) if 'lev' not in rubrik[0] else 0,
          next((i for i, r in enumerate(rubrik) if 'lev' in r), 1),
          next((i for i, r in enumerate(rubrik) if 'namn' in r), 2),
          next((i for i, r in enumerate(rubrik) if 'plats' in r), 3))
    ut, sedda = [], set()
    for n, r in enumerate(rader[1:], start=2):
        r = list(r) + [None] * 4
        art, lev, namn, plats = (text(r[i]) for i in ci)
        if not (art or lev or namn or plats):
            continue
        if not namn:
            avvik.append(f"Artikelfil rad {n}: saknar namn ({art}) - hoppades över")
            continue
        plats_norm = normalisera_plats(plats)
        if not PLATS.match(plats_norm) and not PLATS_STRICT.match(plats_norm):
            avvik.append(f"Artikelfil rad {n}: ogiltig plats '{plats}' - {art} {namn} - hoppades över")
            continue
        k = (art, lev, namn.lower(), plats)
        if k in sedda:
            avvik.append(f"Artikelfil rad {n}: dubblett - {art} {namn} - hoppades över")
            continue
        sedda.add(k)
        ut.append(f"{art or '-'};{lev or '-'};{namn};{plats}")
    return ut


def las_bulk(fil, avvik):
    """Läser Excel-filen med bulkplatser: Artikelnummer, Plockplats, Bulkplats (första fliken)."""
    wb = load_workbook(fil, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    rader = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rader:
        sys.exit("Bulkfilen är tom.")
    rubrik = [text(c).lower() for c in rader[0]]

    def kol(*ordlista, standard):
        for i, r in enumerate(rubrik):
            if any(o in r for o in ordlista):
                return i
        return standard
    ia, ip, ib = kol('artikelnummer', 'artnr', standard=0), kol('plock', standard=2), kol('bulk', standard=3)
    poster = {}
    for n, r in enumerate(rader[1:], start=2):
        r = list(r) + [None] * 8
        art, pick, bulk = text(r[ia]), text(r[ip]), text(r[ib])
        if not (art or bulk):
            continue
        if not art:
            avvik.append(f"Bulkfil rad {n}: saknar artikelnummer ({bulk}) - hoppades över")
            continue
        if not bulk:
            avvik.append(f"Bulkfil rad {n}: {art} saknar bulkplats - hoppades över")
            continue
        if not BULK.match(bulk):
            avvik.append(f"Bulkfil rad {n}: bulkplatsen '{bulk}' ser ovanlig ut ({art}) - används ändå")
        poster.setdefault(art, []).append((pick, bulk))
    return poster


def lagg_pa_bulk(artiklar, poster, avvik):
    """Lägger bulkplatsen sist på artikelraderna (art;lev;namn;plats;bulk). Matchar på artikelnummer."""
    rader_per_art = {}
    for i, rad in enumerate(artiklar):
        rader_per_art.setdefault(rad.split(';')[0], []).append(i)
    ut = list(artiklar)
    antal = 0
    for art, kandidater in poster.items():
        index = rader_per_art.get(art)
        if not index:
            avvik.append(f"Bulk: artikel {art} finns inte bland förbrukningsartiklarna (bulkplats {kandidater[0][1]}) - inte tillagd")
            continue
        for i in index:
            plats = ut[i].split(';')[3]
            val = next((k for k in kandidater if k[0] == plats), None)
            if val is None:
                if len(index) > 1:
                    avvik.append(f"Bulk: artikel {art} ligger på flera platser och ingen matchar plockplatsen i bulkfilen - ingen bulkplats lades till på {plats}")
                    continue
                val = kandidater[0]
                if val[0]:
                    avvik.append(f"Bulk: artikel {art} har plockplats {val[0]} i bulkfilen men {plats} på etiketten - etiketten gäller")
            if len({k[1] for k in kandidater if k[0] == val[0]}) > 1:
                avvik.append(f"Bulk: artikel {art} finns flera gånger i bulkfilen med olika bulkplatser - första används")
            ut[i] = ut[i] + ';' + val[1]
            antal += 1
    return ut, antal


def las_galler(mapp, avvik, framsteg=None):
    filer = []
    for mon in ('*.xlsx', '*.xlsm'):
        filer += glob.glob(os.path.join(mapp, '**', mon), recursive=True)
    filer = [f for f in sorted(filer) if not os.path.basename(f).startswith('~$')]
    if not filer:
        sys.exit("Hittade inga Excel-filer i " + mapp)
    poster = []
    for fnr, f in enumerate(filer, 1):
        if framsteg:
            framsteg(fnr, len(filer), 'filer')
        wb = load_workbook(f, data_only=True, read_only=True)
        for ws in wb.worksheets:
            rad = reol = None
            antal = 0
            for celler in ws.iter_rows(values_only=True):
                celler = list(celler)
                if rad is None:
                    for c in celler:
                        m = HEADER.search(str(c)) if c is not None else None
                        if m:
                            rad, reol = int(m[1]), int(m[2])
                            break
                    continue
                for i, c in enumerate(celler):
                    if heltal(c) is None:
                        continue
                    namn = next((x for x in celler[i + 1:] if x is not None and str(x).strip()), None)
                    if namn is not None and heltal(namn) is None:
                        n = re.sub(r'^modul (\d)', r'Modul \1', text(namn), flags=re.I)
                        if n and not n.startswith('*'):
                            poster.append((n, rad, reol))
                            antal += 1
                    break
            nm = os.path.basename(f)
            if rad is None:
                if ws.max_row and ws.max_row > 1:
                    avvik.append(f"Galler: hoppade över flik '{ws.title}' i {nm} (ingen 'Rad/Reol'-rubrik)")
            elif antal == 0:
                avvik.append(f"Galler: tom reol Rad {rad} Reol {reol} ({nm}, flik '{ws.title}')")
        wb.close()
    sedda, unika = set(), []
    for n, rad, reol in poster:
        k = (n.lower(), rad, reol)
        if k not in sedda:
            sedda.add(k)
            unika.append((n, rad, reol))
    unika.sort(key=lambda t: (t[1], t[2]))
    return [f";;{n};{rad}:{reol}" for n, rad, reol in unika]


def js_sträng(s):
    return s.replace('\\', '\\\\').replace('`', '\\`').replace('${', '\\${')


def valj_fil(sokvag: str, monster: str) -> str:
    """Är sökvägen en mapp används den senast ändrade filen i mappen (t.ex. nyaste PDF-filen)."""
    if os.path.isdir(sokvag):
        filer = [f for f in glob.glob(os.path.join(sokvag, monster))
                 if not os.path.basename(f).startswith('~$')]
        if not filer:
            sys.exit(f"Hittade ingen {monster}-fil i {sokvag}")
        return max(filer, key=os.path.getmtime)
    if not os.path.exists(sokvag):
        sys.exit("Hittar inte filen: " + sokvag)
    return sokvag


def basmapp():
    return os.path.dirname(os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else __file__))


def las_installningar():
    """Läser installningar.txt (bredvid programmet) och gör om till kommandoradsargument."""
    fil = os.path.join(basmapp(), 'installningar.txt')
    if not os.path.exists(fil):
        sys.exit(f"Hittar inte {fil}\nSkapa den med raderna:\n  pdf=C:\\Lager\\lager.pdf\n  galler=C:\\Galler\n  ut=\\\\server\\delat\\lager\\data.js")
    v = {}
    with open(fil, encoding='utf-8-sig') as fh:
        for rad in fh:
            rad = rad.strip()
            if rad and not rad.startswith('#') and '=' in rad:
                k, val = rad.split('=', 1)
                v[k.strip().lower()] = val.strip().strip('"')
    args = []
    for nyckel in ('pdf', 'artiklar', 'galler', 'bulk', 'ut'):
        if v.get(nyckel):
            args += ['--' + nyckel, v[nyckel]]
    pub = v.get('publicera', '').lower()
    if pub in ('ja', 'yes', 'true', '1', 'git'):
        args.append('--publicera')
    elif pub == 'api':
        args.append('--publicera-api')
        for nyckel in ('github_repo', 'github_token', 'github_branch', 'github_api'):
            if v.get(nyckel):
                args += ['--' + nyckel.replace('_', '-'), v[nyckel]]
    return args


def git(mapp, *args):
    try:
        return subprocess.run(['git', '-C', mapp] + list(args), capture_output=True, text=True)
    except FileNotFoundError:
        sys.exit("Git hittades inte. Installera Git for Windows fran https://git-scm.com/download/win "
                 "(data.js ar annars redan skriven).")


def publicera(ut, stämpel):
    mapp = os.path.dirname(ut)
    if git(mapp, 'rev-parse', '--is-inside-work-tree').returncode != 0:
        sys.exit("Kan inte publicera: " + mapp + " är inte en git-mapp (klonad GitHub-mapp). "
                 "data.js är ändå skriven.")
    git(mapp, 'add', os.path.basename(ut))
    if git(mapp, 'diff', '--cached', '--quiet').returncode == 0:
        info("Ingen ändring i data.js sedan förra gången - inget att publicera.")
        return
    ident = []   # git kräver namn och e-post för commit; används bara om de inte redan är inställda
    if not git(mapp, 'config', 'user.name').stdout.strip():
        ident += ['-c', 'user.name=Lagersok']
    if not git(mapp, 'config', 'user.email').stdout.strip():
        ident += ['-c', 'user.email=lagersok@users.noreply.github.com']
    c = git(mapp, *ident, 'commit', '-m', f'Uppdaterad databas {stämpel}')
    if c.returncode != 0:
        sys.exit("git commit misslyckades:\n" + c.stdout + c.stderr)
    p = git(mapp, 'push')
    if p.returncode != 0:
        p = git(mapp, 'push', '-u', 'origin', 'HEAD')   # forsta uppladdningen i ett nytt arkiv
    if p.returncode != 0:
        sys.exit("git push misslyckades (är du inloggad mot GitHub?):\n" + p.stdout + p.stderr)
    ok("Publicerat via Git. Det kan ta någon minut innan sidan visar nya datan.")


def publicera_api(ut, stämpel, repo, token, gren, api):
    """Laddar upp data.js till GitHub via webb-API:t. Kräver inte Git, bara internet och en åtkomstnyckel."""
    if not repo or not token:
        sys.exit("Publicering via API kraver github_repo och github_token i installningar.txt")
    fil = os.path.basename(ut)
    url = f"{api}/repos/{repo}/contents/{fil}"
    sidhuvud = {'Authorization': 'Bearer ' + token, 'Accept': 'application/vnd.github+json',
                'User-Agent': 'lagersok-uppdatering', 'X-GitHub-Api-Version': '2022-11-28'}

    def anrop(metod, adress, data=None):
        for _ in range(5):
            req = urllib.request.Request(adress, data=data, method=metod, headers=sidhuvud)
            if data is not None:
                req.add_header('Content-Type', 'application/json')
            try:
                with urllib.request.urlopen(req, timeout=60) as r:
                    return json.loads(r.read().decode('utf-8'))
            except urllib.error.HTTPError as e:
                if e.code in (301, 302, 303, 307, 308):
                    loc = e.headers.get('Location') or e.headers.get('location')
                    if loc:
                        # GitHub kan skicka 307 till /repositories/{id}/contents/...
                        # Följ för både GET och PUT om redirect går till api.github.com
                        if 'api.github.com' in loc or metod == 'GET':
                            adress = loc
                            continue
                raise

    try:
        sha = None
        try:
            sha = anrop('GET', f"{url}?ref={gren}").get('sha')
        except urllib.error.HTTPError as e:
            if e.code != 404:
                raise
        with open(ut, 'rb') as fh:
            innehall = fh.read()
        body = {'message': f'Uppdaterad databas {stämpel}', 'branch': gren,
                'content': base64.b64encode(innehall).decode('ascii')}
        if sha:
            body['sha'] = sha
        anrop('PUT', url, json.dumps(body).encode('utf-8'))
    except urllib.error.HTTPError as e:
        hjalp = {401: "Nyckeln (token) ar fel eller har gatt ut.",
                 403: "Nyckeln saknar ratt att skriva, eller natverket blockerar.",
                 404: "Arkivet hittades inte. Kontrollera github_repo och att nyckeln har atkomst till arkivet.",
                 409: "Konflikt - kor igen.", 422: "Grenen finns inte. Kontrollera github_branch.",
                 307: "307 redirect - följer nu redirect automatiskt till api.github.com/repositories/...",
                 302: "302 redirect - följer redirect"}.get(e.code, "")
        loc = ""
        try:
            loc = e.headers.get('Location','') or e.headers.get('location','')
        except Exception:
            pass
        if loc:
            hjalp += f" Redirect till: {loc[:200]}"
        sys.exit(f"GitHub svarade {e.code} {e.reason}. {hjalp}")
    except urllib.error.URLError as e:
        sys.exit("Kunde inte na GitHub (" + str(e.reason) + "). Kontrollera internet, proxy och brandvagg.")
    ok("Publicerat pa GitHub. Det kan ta nagon minut innan sidan visar nya datan.")


def las_installningar_dict():
    """Returnerar installningar.txt som dict - används av GUI. Fix för PyInstaller."""
    fil = os.path.join(basmapp(), 'installningar.txt')
    d = {}
    if not os.path.isfile(fil):
        return d
    try:
        with open(fil, encoding='utf-8-sig') as fh:
            for rad in fh:
                rad = rad.strip()
                if not rad or rad.startswith('#') or '=' not in rad:
                    continue
                k, v = rad.split('=', 1)
                d[k.strip().lower()] = v.strip().strip('"')
                # behåll även original-case för GUI som förväntar sig lower-case nycklar som github_repo
                d[k.strip()] = v.strip().strip('"')
    except Exception:
        pass
    # Normalisera nycklar till det GUI förväntar sig
    # GUI vill ha: pdf, galler, bulk, ut, sparr_procent, publicera, github_repo, github_branch, github_token
    out = {}
    out['pdf'] = d.get('pdf', '')
    out['artiklar'] = d.get('artiklar', '')
    out['galler'] = d.get('galler', '')
    out['bulk'] = d.get('bulk', '')
    out['ut'] = d.get('ut', '')
    out['sparr_procent'] = d.get('sparr_procent', d.get('sparr', '20'))
    out['publicera'] = d.get('publicera', '')
    out['github_repo'] = d.get('github_repo', '')
    out['github_branch'] = d.get('github_branch', 'main')
    out['github_token'] = d.get('github_token', '')
    # ta även med alla andra
    for k,v in d.items():
        if k not in out:
            out[k] = v
    return out

def skriv_installningar(data_dict):
    """Skriver installningar.txt bredvid exe - FIX för PyInstaller onefile"""
    fil = os.path.join(basmapp(), 'installningar.txt')
    os.makedirs(os.path.dirname(fil), exist_ok=True)
    rader = []
    rader.append("# Sökvägar för uppdatera_databas / Lagersok")
    rader.append(f"# Uppdaterad: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    rader.append("")
    # Viktig ordning
    for nyckel in ['pdf', 'artiklar', 'galler', 'bulk', 'ut', 'sparr_procent', 'publicera', 'github_repo', 'github_branch', 'github_token']:
        if nyckel in data_dict and data_dict[nyckel] not in (None, ''):
            rader.append(f"{nyckel}={data_dict[nyckel]}")
        elif nyckel in data_dict:
            # skriv även tomma för tydlighet
            rader.append(f"{nyckel}={data_dict[nyckel]}")
    # Extra nycklar
    for k,v in data_dict.items():
        if k not in ['pdf', 'artiklar', 'galler', 'bulk', 'ut', 'sparr_procent', 'publicera', 'github_repo', 'github_branch', 'github_token']:
            rader.append(f"{k}={v}")
    rader.append("")
    with open(fil, 'w', encoding='utf-8') as fh:
        fh.write("\n".join(rader))
    return fil

def testa_github(repo, token, branch='main'):
    """Testar GitHub API - används av GUI"""
    if not repo or not token:
        return (False, "Fyll i repo och token först")
    if '/' not in repo:
        return (False, "Repo ska vara formatet ägare/repo, t.ex. sterilen/sterilen.github.io")
    try:
        url = f"https://api.github.com/repos/{repo}"
        req = urllib.request.Request(url, method='GET', headers={
            'Authorization': 'Bearer ' + token,
            'Accept': 'application/vnd.github+json',
            'User-Agent': 'lagersok-test'
        })
        with urllib.request.urlopen(req, timeout=15) as r:
            data = json.loads(r.read().decode('utf-8'))
            return (True, f"OK! Hittade {data.get('full_name')} - branch {branch}")
    except urllib.error.HTTPError as e:
        if e.code == 401:
            return (False, "401: Token fel eller har gått ut")
        elif e.code == 404:
            return (False, "404: Repo hittades inte - kolla namn och att token har access")
        elif e.code == 403:
            return (False, f"403: Ingen rättighet - {e.reason}")
        else:
            return (False, f"{e.code}: {e.reason}")
    except Exception as e:
        return (False, f"Kunde inte nå GitHub: {e}")

def _count_from_data_js(sokvag):
    """Räknar artiklar i befintlig data.js för spärrkontroll"""
    try:
        if not os.path.isfile(sokvag):
            return None
        txt = pathlib.Path(sokvag).read_text(encoding='utf-8', errors='ignore')
        # Leta efter window.LAGER_CSV = `...`
        m = re.search(r'window\.LAGER_CSV\s*=\s*`([^`]*)`', txt, re.DOTALL)
        if not m:
            return None
        csv_content = m.group(1)
        lines = [l for l in csv_content.split('\n') if l.strip() and not l.startswith('Artikelnummer')]
        # Försök separera fabrikssterilt vs galler: galler-rader börjar med ;;
        fab = len([l for l in lines if not l.startswith(';;')])
        gal = len([l for l in lines if l.startswith(';;')])
        return (fab, gal, len(lines))
    except Exception:
        return None

def kor(a, bekrafta_fn=None, steg_cb=None, prog_cb=None, ok_cb=None, info_cb=None, varning_cb=None, fel_cb=None, **kwargs):
    """Kör uppdateringen - används av GUI. Stöder både gammal och ny signatur."""
    # Stöd för olika sätt GUI anropar: kor(a, bekrafta) eller kor(a, steg, prog, ...)
    # kwargs kan innehålla steg_cb, prog_cb etc från vår nya GUI-fix
    
    # Om anropet kom som positionella callbacks (gammal GUI-variant)
    # a är namespace, andra arg kan vara steg-funktion om bekrafta_fn ser ut som steg
    # Vi normaliserar: om bekrafta_fn inte är callable som bekrafta utan är steg, skifta
    # Men vi har nu kwargs, så vi låter bekrafta_fn vara bekrafta
    
    # Hämta callbacks från kwargs om de finns (nya GUI:t)
    if 'steg_cb' in kwargs: steg_cb = kwargs['steg_cb']
    if 'prog_cb' in kwargs: prog_cb = kwargs['prog_cb']
    if 'ok_cb' in kwargs: ok_cb = kwargs['ok_cb']
    if 'info_cb' in kwargs: info_cb = kwargs['info_cb']
    if 'varning_cb' in kwargs: varning_cb = kwargs['varning_cb']
    if 'fel_cb' in kwargs: fel_cb = kwargs['fel_cb']
    if 'bekrafta_fn' in kwargs: bekrafta_fn = kwargs['bekrafta_fn']
    avbryt_fn = kwargs.get('avbryt_fn')

    def kolla():
        if avbryt_fn and avbryt_fn():
            raise SystemExit("Avbruten av användaren - inget har ändrats.")
    
    # Fallback till globala funktioner om ingen callback
    def _call(cb, *args, **kw):
        if cb:
            try:
                cb(*args, **kw)
            except Exception:
                pass
    
    # Wrapper för att både logga till konsol och GUI
    def do_steg(nr, tot, text):
        if nr < tot - 1:
            kolla()   # läs-stegen kan avbrytas; skriv/publicera-stegen gör det inte
        _call(steg_cb, nr, tot, text)
        try: steg(nr, tot, text)
        except Exception: pass
    
    def do_ok(text):
        _call(ok_cb, text)
        try: ok(text)
        except Exception: pass
    
    def do_info(text):
        _call(info_cb, text)
        try: info(text)
        except Exception: pass
    
    def do_varning(text):
        _call(varning_cb, text)
        try: varning(text)
        except Exception: pass
    
    def do_framsteg(i, n, t=""):
        kolla()
        _call(prog_cb, i, n, t)
        try: framsteg(i, n, t)
        except Exception: pass
    
    nr_total = 4 + (1 if getattr(a, 'bulk', '') else 0)
    avvik = []
    
    # 1. Förbrukning
    if getattr(a, 'pdf', None):
        pdf_path = valj_fil(a.pdf, '*.pdf')
        do_steg(1, nr_total, "Läser förbrukningsartiklar från PDF")
        do_info(pdf_path)
        artiklar = las_pdf(pdf_path, avvik, do_framsteg)
    else:
        art_path = valj_fil(getattr(a, 'artiklar', ''), '*.xlsx')
        do_steg(1, nr_total, "Läser förbrukningsartiklar från Excel")
        do_info(art_path)
        artiklar = las_artiklar(art_path, avvik)
    do_ok(f"{len(artiklar)} artiklar inlästa")
    
    # 2. Galler
    do_steg(2, nr_total, "Läser galler från Excel-filerna")
    do_info(a.galler)
    galler = las_galler(a.galler, avvik, do_framsteg)
    do_ok(f"{len(galler)} galler inlästa")
    
    # 3. Bulk
    bulkantal = 0
    if getattr(a, 'bulk', ''):
        bulk_path = valj_fil(a.bulk, '*.xlsx')
        do_steg(3, nr_total, "Läser bulkplatser från Excel")
        do_info(bulk_path)
        artiklar, bulkantal = lagg_pa_bulk(artiklar, las_bulk(bulk_path, avvik), avvik)
        do_ok(f"{bulkantal} artiklar fick en bulkplats")
    
    # Spärrkontroll - jämför med tidigare data.js
    forra = None
    ut_path = os.path.abspath(a.ut)
    if os.path.isfile(ut_path):
        forra = _count_from_data_js(ut_path)
        if forra:
            total_forra = forra[2] if len(forra) > 2 else forra[0] + forra[1]
            total_nu = len(artiklar) + len(galler)
            _sp = getattr(a, 'sparr_procent', 20)
            sparr_pct = 20.0 if _sp is None else float(_sp)   # 0 = ingen spärr
            if sparr_pct > 0 and total_forra > 0:
                minskning = (total_nu - total_forra) / total_forra * 100
                if minskning < -sparr_pct:
                    problem = f"Databasen skulle minska från {total_forra} till {total_nu} artiklar ({minskning:.1f}%)."
                    if bekrafta_fn:
                        bekrafta_fn(problem)
                    else:
                        sys.exit(f"SPÄRR: {problem} Kör med --tvinga för att kringgå.")
    
    stampel = datetime.now().strftime('%Y-%m-%d %H:%M')
    
    if getattr(a, 'torrkorning', False):
        return {
            'forbruk': len(artiklar),
            'galler': len(galler),
            'bulk': bulkantal,
            'avvik': avvik,
            'torr': True,
            'ut': ut_path,
            'stampel': stampel,
            'forra': forra,
            'adress': ''
        }
    
    kolla()
    # Skriv data.js
    do_steg(nr_total-1, nr_total, "Skriver data.js")
    csv = "Artikelnummer;Leverantörsnummer;Produktnamn;Plats;Bulkplats\n" + "\n".join(artiklar + galler)
    innehall = ("// Skapad automatiskt av uppdatera_databas.py - ändra inte för hand.\n"
                f"window.LAGER_UPPDATERAD = '{stampel}';\n"
                f"window.LAGER_CSV = `{js_sträng(csv)}`;\n")
    ut = ut_path
    if os.path.exists(ut):
        shutil.copy2(ut, ut + '.bak')
    os.makedirs(os.path.dirname(ut), exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=os.path.dirname(ut), suffix='.tmp')
    with os.fdopen(fd, 'w', encoding='utf-8', newline='\n') as fh:
        fh.write(innehall)
    os.replace(tmp, ut)
    with open(os.path.join(os.path.dirname(ut), 'avvikelser.txt'), 'w', encoding='utf-8') as fh:
        fh.write(f"Körning {stampel}\n" + "\n".join(avvik) + "\n")
    do_ok(ut)
    if avvik:
        do_varning(f"{len(avvik)} rader hoppades över - se avvikelser.txt")
    
    do_steg(nr_total, nr_total, "Publicerar")
    adress = ""
    if getattr(a, 'publicera', False):
        publicera(ut, stampel)
    elif getattr(a, 'publicera_api', False):
        publicera_api(ut, stampel, a.github_repo, a.github_token, a.github_branch, a.github_api)
        if '/' in getattr(a, 'github_repo', ''):
            agare, arkiv = a.github_repo.split('/', 1)
            adress = f"https://{agare.lower()}.github.io/" + ("" if arkiv.lower() == f"{agare.lower()}.github.io" else arkiv + "/")
    else:
        do_info("Hoppar över publicering")
    
    return {
        'forbruk': len(artiklar),
        'galler': len(galler),
        'bulk': bulkantal,
        'avvik': avvik,
        'torr': False,
        'ut': ut_path,
        'stampel': stampel,
        'forra': forra,
        'adress': adress
    }


# ------------------------------------------------------------------ Följesedlar (MediCarrier-PDF) -> platsetiketter
# Läser följesedlar (PDF med textlager), matchar artikelnumren mot data.js och ger en lista över vilka
# artiklar som ska ha platsetikett. Används av fliken FÖLJESEDLAR i Lagersok Admin.
# Radnummer, artikelnummer, beställt, levererat, enhet, återstår. Artikelnumret kan vara 1-8 siffror.
RAD_RE = re.compile(r'^\s*(\d{1,3})\s+(\d+)\s+(\d+)\s+(\d+)\s+(\S.*?)\s+(\d+)\s*$')
REST_RAD_RE = re.compile(r'^\s*(\d{1,3})\s+(\d+)\s+(\S.*\s\d+\s*[A-Za-z]{1,4})\s*$')
REST_START_RE = re.compile(r'Best[äa]llningsvaror|Restorder', re.I)
FOTRADER = ('Bankgironr', 'Plusgiro', 'Momsreg', 'Telefon', 'Fax', 'Innehar', 'MediCarrier', 'Finspångsgatan', 'SE - 163', 'Säljarens', 'Kundtjänst', 'Märkning', 'Återstår att', 'Beställd', 'kvantitet')


def las_foljesedel(sokvag):
    """Returnerar dict: fil, ordernr, datum, rader (levererade), rest (restorder, hoppas över)."""
    try:
        import pdfplumber
    except ImportError:
        sys.exit("pdfplumber saknas. Kör:  pip install pdfplumber")
    rader, rest = [], []
    ordernr = datum = so = ''
    i_rest = False
    nuvarande = None
    with pdfplumber.open(sokvag) as pdf:
        for sida in pdf.pages:
            text = sida.extract_text() or ''
            for rad in text.splitlines():
                if not so:
                    m = re.search(r'\bSO\d{6,}\b', rad)
                    if m:
                        so = m.group(0)
                if not ordernr:
                    m = re.match(r'^\s*(\d{6,9})\s+(\d{4}-\d{2}-\d{2})\s*$', rad)
                    if m:
                        ordernr, datum = m.group(1), m.group(2)
                if REST_START_RE.search(rad) and not RAD_RE.match(rad):
                    i_rest = True
                    nuvarande = None
                    continue
                if i_rest:
                    m = REST_RAD_RE.match(rad)
                    if m:
                        rest.append({'rad': int(m.group(1)), 'artnr': m.group(2)})
                    continue
                m = RAD_RE.match(rad)
                if m:
                    nuvarande = {'rad': int(m.group(1)), 'artnr': m.group(2), 'bestalld': int(m.group(3)),
                                 'levererat': int(m.group(4)), 'enhet': m.group(5), 'namn': ''}
                    rader.append(nuvarande)
                elif rad.lstrip().startswith(FOTRADER):
                    nuvarande = None      # sidfot: namnet på raden ovanför är slut
                elif nuvarande is not None and rad.strip():
                    # radens namn fortsätter på nästa rad/rader (men inte rubriker och sidfot)
                    if re.match(r'^\s*(Sida|Ordernr|Godsm|Leverans|Stockholm|varumott|Mottagare|Södersjukhuset|Steriltekniskt|Hiss|Internpost|Köpare|\d+\s+\d+\s+Stockholm)', rad):
                        nuvarande = None
                    else:
                        nuvarande['namn'] = (nuvarande['namn'] + ' ' + rad.strip()).strip()
    return {'fil': os.path.basename(sokvag), 'so': so, 'ordernr': ordernr, 'datum': datum, 'rader': rader, 'rest': rest}


def las_databas(datajs):
    """Läser artiklarna ur en data.js på disk."""
    with open(datajs, encoding='utf-8-sig') as fh:
        return las_databas_text(fh.read(), datajs)


def hamta_databas_github(repo, token, gren='main', fil='data.js', api='https://api.github.com'):
    """Hämtar den publicerade data.js från GitHub (samma fil som appen visar). Returnerar filens text."""
    if not repo or not token:
        raise ValueError("arkiv eller åtkomstnyckel saknas i inställningarna")
    req = urllib.request.Request(f"{api}/repos/{repo}/contents/{fil}?ref={gren}", headers={
        'Authorization': 'Bearer ' + token, 'Accept': 'application/vnd.github.raw+json',
        'User-Agent': 'lagersok-foljesedel', 'X-GitHub-Api-Version': '2022-11-28'})
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return r.read().decode('utf-8-sig')
    except urllib.error.HTTPError as e:
        raise ValueError(f"GitHub svarade {e.code} {e.reason}")
    except urllib.error.URLError as e:
        raise ValueError("kunde inte nå GitHub (" + str(e.reason) + ")")


def las_databas_text(txt, kalla='data.js'):
    """Plockar artikelraderna ur texten i en data.js. Returnerar lista med dict: artnr, lev, namn, plats, bulk."""
    m = re.search(r'window\.LAGER_CSV\s*=\s*(`|")((?:\\.|(?!\1)[^\\])*)\1', txt, re.S)
    if not m:
        raise ValueError("Hittade ingen artikeldata i " + kalla)
    csv = m.group(2)
    if m.group(1) == '"':
        csv = csv.replace('\\n', '\n')
    csv = re.sub(r'\\(.)', r'\1', csv, flags=re.S)   # tar bort js_sträng-escapes (\\, \`, \${)
    ut = []
    for i, rad in enumerate(csv.splitlines()):
        d = rad.split(';')
        if i == 0 and d[0].strip().lower() in ('artikelnummer', 'artnr'):
            continue
        d += [''] * (5 - len(d))
        if not any(x.strip() for x in d):
            continue
        ut.append({'artnr': d[0].strip(), 'lev': d[1].strip(), 'namn': d[2].strip(), 'plats': d[3].strip(), 'bulk': d[4].strip()})
    stamp = re.search(r"LAGER_UPPDATERAD\s*=\s*['\"]([^'\"]*)", txt)
    las_databas_text.senaste_stampel = stamp.group(1) if stamp else ''
    return ut


def bygg_index(db):
    idx = {}
    for a in db:
        for nyckel in {a['artnr'], a['lev']}:
            if nyckel and nyckel != '-':
                idx.setdefault(nyckel, []).append(a)
    return idx


def matcha(foljesedlar, db):
    """En post per orderrad. status: 'etikett' | 'bulk' | 'saknas' | 'ingenplats'."""
    idx = bygg_index(db)
    ut = []
    for f in foljesedlar:
        for r in f['rader']:
            traffar = idx.get(r['artnr'], [])
            post = {'fil': f['fil'], 'so': f['so'], 'ordernr': f['ordernr'], 'rad': r['rad'], 'artnr': r['artnr'], 'levererat': r['levererat'],
                    'enhet': r['enhet'], 'fnamn': r['namn'], 'namn': r['namn'], 'datum': f['datum'], 'lev': '', 'plats': '', 'bulk': '', 'status': 'saknas', 'notering': ''}
            if traffar:
                forsta = traffar[0]
                post['namn'] = forsta['namn'] or r['namn']
                post['lev'] = forsta['lev']
                if forsta['artnr']:
                    post['artnr'] = forsta['artnr']      # etiketten visar alltid vårt eget artikelnummer
                if r.get('kort'):
                    post['notering'] = 'Kort nummer, kontrollera att det är rätt artikel'
                if any(t['bulk'] for t in traffar):
                    b = next(t for t in traffar if t['bulk'])
                    post.update(status='bulk', plats=b['plats'], bulk=b['bulk'], namn=b['namn'] or post['namn'])
                else:
                    plats = []
                    for t in traffar:
                        if t['plats'] and t['plats'] not in plats:
                            plats.append(t['plats'])
                    if plats:
                        post.update(status='etikett', plats=plats[0])
                        if len(plats) > 1:
                            post['notering'] = 'Flera platser: ' + ', '.join(plats)
                    else:
                        post.update(status='ingenplats')
            ut.append(post)
    return ut


def sammanfatta(foljesedlar, poster):
    antal = {}
    for p in poster:
        antal[p['status']] = antal.get(p['status'], 0) + 1
    return {'filer': len(foljesedlar), 'rader': len(poster), 'etikett': antal.get('etikett', 0), 'bulk': antal.get('bulk', 0),
            'saknas': antal.get('saknas', 0), 'ingenplats': antal.get('ingenplats', 0),
            'rest': sum(len(f['rest']) for f in foljesedlar)}



# ------------------------------------------------------------------ Andra leverantörer: scanning/foto/PDF utan fast layout
# 1. Har PDF:en ett textlager (t.ex. sökbar PDF från skannern) används det. 2. Annars OCR: först Windows egen OCR (inget att installera),
# sedan, som reserv, RapidOCR om det paketet följer med programmet. Därefter plockas alla tal ut som finns i databasen.
OCR_BILD_EXT = ('.png', '.jpg', '.jpeg', '.tif', '.tiff', '.bmp')


def _ocr_windows(png):
    if os.name != 'nt':
        raise OSError("Windows OCR finns bara i Windows")
    import asyncio
    from winrt.windows.media.ocr import OcrEngine
    from winrt.windows.graphics.imaging import BitmapDecoder
    from winrt.windows.storage.streams import InMemoryRandomAccessStream, DataWriter

    async def kor():
        ström = InMemoryRandomAccessStream()
        skrivare = DataWriter(ström)
        skrivare.write_bytes(bytearray(png))
        await skrivare.store_async()
        await skrivare.flush_async()
        skrivare.detach_stream()
        ström.seek(0)
        avkodare = await BitmapDecoder.create_async(ström)
        bild = await avkodare.get_software_bitmap_async()
        motor = OcrEngine.try_create_from_user_profile_languages()
        if motor is None:
            raise RuntimeError("inget OCR-språk finns installerat i Windows")
        res = await motor.recognize_async(bild)
        return '\n'.join(rad.text for rad in res.lines)
    return asyncio.run(kor())


def _ocr_rapid(png):
    import io
    import numpy as np
    from PIL import Image
    from rapidocr_onnxruntime import RapidOCR
    if not hasattr(_ocr_rapid, 'motor'):
        _ocr_rapid.motor = RapidOCR()
    bild = np.array(Image.open(io.BytesIO(png)).convert('RGB'))
    res, _ = _ocr_rapid.motor(bild)
    if not res:
        return ''
    rader = sorted(res, key=lambda r: (round(r[0][0][1] / 12), r[0][0][0]))      # uppifrån och ner, vänster till höger
    return '\n'.join(r[1] for r in rader)



def _ocr_tesseract(png_bytes):
    """Tesseract fallback - kräver tesseract.exe och pytesseract"""
    import io
    from PIL import Image
    import pytesseract
    # Försök hitta tesseract.exe automatiskt
    import shutil, os
    # Leta i PATH och vanliga Windows-sökvägar
    tesseract_cmd = shutil.which("tesseract")
    if not tesseract_cmd:
        for cand in [r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                     r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
                     os.path.join(os.path.dirname(__file__), "tesseract", "tesseract.exe"),
                     os.path.join(os.getcwd(), "tesseract", "tesseract.exe")]:
            if os.path.exists(cand):
                tesseract_cmd = cand
                break
    if tesseract_cmd:
        pytesseract.pytesseract.tesseract_cmd = tesseract_cmd
    im = Image.open(io.BytesIO(png_bytes))
    # Svenska + engelska
    try:
        return pytesseract.image_to_string(im, lang="swe+eng")
    except Exception:
        return pytesseract.image_to_string(im, lang="eng")



def ocr_bild(img):
    """Läser texten i en bild (PIL). Provar Windows OCR, RapidOCR och Tesseract."""
    import io
    img = img.convert('L')
    img.thumbnail((4800, 4800))
    buf = io.BytesIO()
    img.save(buf, 'PNG')
    png = buf.getvalue()
    fel = []
    for namn, fn in (('Windows OCR', _ocr_windows), ('RapidOCR', _ocr_rapid), ('Tesseract', _ocr_tesseract)):
        try:
            t = fn(png)
            if t and t.strip():
                return t
            fel.append(namn + ': ingen text hittades')
        except Exception as ex:
            fel.append(f"{namn}: {ex}")
    raise RuntimeError("OCR misslyckades (" + '; '.join(fel) + "). Installera Tesseract OCR (https://github.com/UB-Mannheim/tesseract/wiki) eller pip install rapidocr_onnxruntime winrt")




def hamta_text(sokvag, framsteg=None):
    """All text i en PDF (textlager, annars OCR sida för sida) eller bildfil (png, jpg, tif)."""
    ext = os.path.splitext(sokvag)[1].lower()
    delar = []
    if ext == '.pdf':
        import pdfplumber
        with pdfplumber.open(sokvag) as pdf:
            for n, sida in enumerate(pdf.pages, 1):
                if framsteg:
                    framsteg(n, len(pdf.pages))
                t = sida.extract_text() or ''
                if len(t.strip()) < 40:                       # ingen riktig text: sidan är en bild
                    t = ocr_bild(sida.to_image(resolution=300).original)
                delar.append(t)
    elif ext in OCR_BILD_EXT:
        from PIL import Image, ImageSequence
        with Image.open(sokvag) as im:
            for n, sida in enumerate(ImageSequence.Iterator(im), 1):
                delar.append(ocr_bild(sida.copy()))
    else:
        raise ValueError("Filtypen stöds inte: " + ext)
    return '\n'.join(delar)


def generisk_foljesedel(text, fil, db):
    """Plockar ut alla tal/koder i texten som är artikelnummer eller leverantörsnummer i databasen (varje en gång, i den ordning de står).
    Layouten spelar ingen roll, men restorder hoppas inte över automatiskt."""
    idx = bygg_index(db)
    sedda, rader = set(), []
    for tok in re.findall(r'[A-Za-z0-9][A-Za-z0-9\-_/]*', text or ''):
        tok = tok.rstrip('-_/')
        if len(tok) < 3 or tok in sedda or tok not in idx:
            continue
        sedda.add(tok)
        rader.append({'rad': len(rader) + 1, 'artnr': tok, 'bestalld': 0, 'levererat': 0, 'enhet': '', 'namn': '',
                      'kort': tok.isdigit() and len(tok) <= 4})
    return {'fil': fil, 'so': '', 'ordernr': '', 'datum': '', 'rader': rader, 'rest': [], 'generisk': True}


# ------------------------------------------------------------------ Etiketter: ZPL till Zebra ZD421
# Etikett 76,2 x 50,8 mm (Gilford/Zebra 2000D, direkttermisk). Skrivaren pratar ZPL. Utskrift går antingen som "rå" utskrift till en Windows-skrivare (USB),
# direkt över nätverk (port 9100) eller sparas som .zpl-fil (för test, t.ex. i labelary.com).
def _mm(mm, dpi):
    return int(round(float(mm) * float(dpi) / 25.4))


def _zt(s, maxlen=None):
    """Gör text säker att lägga i ett ZPL-fält (^ och ~ är styrtecken)."""
    s = re.sub(r'[\^~\r\n\t]+', ' ', str(s or '')).strip()
    return s[:maxlen] if maxlen else s


def zpl_etikett(plats, namn, artnr='', lev='', sedel='', datum='', dpi=203, bredd_mm=76.2, hojd_mm=50.8, rad=None):
    """En platsetikett med bara tre uppgifter: PLATS (störst), produktnamn och artikelnummer.
    (lev, sedel och datum tas emot men skrivs inte ut.) Layouten är ritad för höjden 72,6 mm och skalas efter
    etikettens höjd; texten anpassas efter bredden."""
    s = float(hojd_mm) / 72.6
    W = _mm(bredd_mm, dpi)
    m = _mm(max(3.0, 5 * s), dpi)
    bred = W - 2 * m

    def M(mm):                      # designmått (mm på en 72,6 mm hög etikett) -> punkter
        return _mm(mm * s, dpi)
    plats = _zt(plats, 14) or '-'
    w_plats = int(bred / (max(len(plats), 1) * 0.62))                # bredd per tecken så att raden ryms (ca 0,6 x höjden)
    h_plats = min(M(28), int(w_plats * 1.4))                          # något smalare (högre) skrift ger större plats
    w_plats = min(w_plats, int(h_plats * 0.75))                       # korta platser sträcks inte ut i bredd
    tjock = max(2, M(0.8))
    rader = ['^XA', '^CI28', '^MTD', f'^PW{W}', f'^LL{_mm(hojd_mm, dpi)}', '^LH0,0', '^PR4', '^MMT']
    rader.append(f'^FO{m},{M(4) + (M(28) - h_plats) // 2}^FB{bred},1,0,C,0^A0N,{h_plats},{w_plats}^FD{plats}^FS')
    rader.append(f'^FO{m},{M(34)}^GB{bred},{tjock},{tjock}^FS')
    rader.append(f'^FO{m},{M(37)}^FB{bred},3,{M(1)},L,0^A0N,{M(7)},{M(7)}^FD{_zt(namn, 75)}^FS')
    if artnr:
        art = 'Art.nr ' + _zt(artnr, 14)
        h_art = min(M(8), int(bred / (len(art) * 0.55)))        # anpassas så att raden ryms i bredd
        rader.append(f'^FO{m},{M(63)}^FB{bred},1,0,L,0^A0N,{h_art},{h_art}^FD{art}^FS')
    rader.append('^XZ')
    return '\n'.join(rader) + '\n'


def zpl_testetikett(dpi=203, bredd_mm=76.2, hojd_mm=50.8):
    return zpl_etikett('A-01-02-03', 'Testetikett - Endotrakealtub SHILEY HI-LO ID 7.0mm', '100132', '12345', 'SO00000000',
                       datetime.now().strftime('%Y-%m-%d'), dpi, bredd_mm, hojd_mm)


def lista_skrivare():
    """Windows-skrivare som är installerade (tom lista på andra system)."""
    if os.name != 'nt':
        return []
    try:
        import ctypes
        from ctypes import wintypes
        sp = ctypes.WinDLL('winspool.drv', use_last_error=True)

        class PINFO4(ctypes.Structure):
            _fields_ = [('pPrinterName', wintypes.LPWSTR), ('pServerName', wintypes.LPWSTR), ('Attributes', wintypes.DWORD)]
        flaggor = 2 | 4   # lokala + anslutna
        behov, antal = wintypes.DWORD(0), wintypes.DWORD(0)
        sp.EnumPrintersW(flaggor, None, 4, None, 0, ctypes.byref(behov), ctypes.byref(antal))
        if not behov.value:
            return []
        buf = ctypes.create_string_buffer(behov.value)
        if not sp.EnumPrintersW(flaggor, None, 4, buf, behov, ctypes.byref(behov), ctypes.byref(antal)):
            return []
        arr = ctypes.cast(buf, ctypes.POINTER(PINFO4))
        return sorted(arr[i].pPrinterName for i in range(antal.value) if arr[i].pPrinterName)
    except Exception:
        return []


def skicka_ra_windows(skrivare, data):
    """Skickar bytes oförändrade (RAW) till en Windows-skrivare, t.ex. Zebra ZD421 via USB."""
    if os.name != 'nt':
        raise OSError("Utskrift via Windows-skrivare fungerar bara i Windows")
    import ctypes
    from ctypes import wintypes
    sp = ctypes.WinDLL('winspool.drv', use_last_error=True)

    class DOCINFO1(ctypes.Structure):
        _fields_ = [('pDocName', wintypes.LPWSTR), ('pOutputFile', wintypes.LPWSTR), ('pDatatype', wintypes.LPWSTR)]
    sp.OpenPrinterW.argtypes = [wintypes.LPWSTR, ctypes.POINTER(ctypes.c_void_p), ctypes.c_void_p]
    sp.StartDocPrinterW.argtypes = [ctypes.c_void_p, wintypes.DWORD, ctypes.c_void_p]
    sp.StartPagePrinter.argtypes = [ctypes.c_void_p]
    sp.WritePrinter.argtypes = [ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD, ctypes.POINTER(wintypes.DWORD)]
    sp.EndPagePrinter.argtypes = [ctypes.c_void_p]
    sp.EndDocPrinter.argtypes = [ctypes.c_void_p]
    sp.ClosePrinter.argtypes = [ctypes.c_void_p]
    h = ctypes.c_void_p()
    if not sp.OpenPrinterW(skrivare, ctypes.byref(h), None):
        raise OSError(f"Kunde inte öppna skrivaren '{skrivare}' (fel {ctypes.get_last_error()})")
    try:
        if not sp.StartDocPrinterW(h, 1, ctypes.byref(DOCINFO1('Lagersok etiketter', None, 'RAW'))):
            raise OSError(f"StartDocPrinter misslyckades (fel {ctypes.get_last_error()})")
        try:
            sp.StartPagePrinter(h)
            skrivet = wintypes.DWORD(0)
            if not sp.WritePrinter(h, data, len(data), ctypes.byref(skrivet)) or skrivet.value != len(data):
                raise OSError(f"Kunde inte skicka data till skrivaren (fel {ctypes.get_last_error()})")
            sp.EndPagePrinter(h)
        finally:
            sp.EndDocPrinter(h)
    finally:
        sp.ClosePrinter(h)


def skicka_ra_natverk(ip, data, port=9100):
    import socket
    try:
        with socket.create_connection((ip, int(port)), timeout=10) as s:
            s.sendall(data)
    except OSError as e:
        raise OSError(f"Kunde inte nå skrivaren på {ip}:{port} ({e})")


def skriv_ut_zpl(zpl, metod, skrivare='', ip=''):
    """metod: 'usb' (Windows-skrivare) eller 'natverk' (IP, port 9100)."""
    data = zpl.encode('utf-8')
    if metod == 'natverk':
        if not ip:
            raise ValueError("IP-adress till skrivaren saknas")
        skicka_ra_natverk(ip.strip(), data)
    else:
        if not skrivare:
            raise ValueError("Ingen skrivare vald")
        skicka_ra_windows(skrivare, data)


def main():
    ap = argparse.ArgumentParser()
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument('--pdf', help="PDF-filen med produktetiketterna, eller en mapp (då används nyaste PDF-filen)")
    g.add_argument('--artiklar', help="Excel-fil med förbrukningsartiklar")
    ap.add_argument('--galler', required=True, help="Mapp med gallerfilerna")
    ap.add_argument('--bulk', default='', help="Excel-fil med bulkplatser (Artikelnummer, Plockplats, Bulkplats) - valfri")
    ap.add_argument('--ut', default='data.js', help="Sökväg till data.js som ska skrivas")
    ap.add_argument('--torrkorning', action='store_true')
    ap.add_argument('--publicera', action='store_true',
                    help="efter körningen: git commit + push av data.js (mappen måste vara en klonad GitHub-mapp)")
    ap.add_argument('--publicera-api', action='store_true', help="ladda upp data.js via GitHubs webb-API (kräver inte Git)")
    ap.add_argument('--github-repo', default='')
    ap.add_argument('--github-token', default='')
    ap.add_argument('--github-branch', default='main')
    ap.add_argument('--github-api', default='https://api.github.com', help=argparse.SUPPRESS)
    a = ap.parse_args(las_installningar() if len(sys.argv) == 1 else None)

    nr = 4 + (1 if a.bulk else 0)
    visa_banner(VERSION)
    avvik = []
    if a.pdf:
        a.pdf = valj_fil(a.pdf, '*.pdf')
        steg(1, nr, "Laser forbrukningsartiklar fran PDF")
        info(a.pdf)
        artiklar = las_pdf(a.pdf, avvik, framsteg)
    else:
        a.artiklar = valj_fil(a.artiklar, '*.xlsx')
        steg(1, nr, "Laser forbrukningsartiklar fran Excel")
        info(a.artiklar)
        artiklar = las_artiklar(a.artiklar, avvik)
    ok(f"{len(artiklar)} artiklar inlasta")

    steg(2, nr, "Laser galler fran Excel-filerna")
    info(a.galler)
    galler = las_galler(a.galler, avvik, framsteg)
    ok(f"{len(galler)} galler inlasta")

    bulkantal = 0
    if a.bulk:
        a.bulk = valj_fil(a.bulk, '*.xlsx')
        steg(3, nr, "Laser bulkplatser fran Excel")
        info(a.bulk)
        artiklar, bulkantal = lagg_pa_bulk(artiklar, las_bulk(a.bulk, avvik), avvik)
        ok(f"{bulkantal} artiklar fick en bulkplats")

    stämpel = datetime.now().strftime('%Y-%m-%d %H:%M')
    if a.torrkorning:
        print()
        ruta([f(f"TORRKORNING - inget skrevs", GUL, FET), "",
              f"Forbrukningsartiklar: {len(artiklar)}", f"Galler:               {len(galler)}",
              f"Med bulkplats:        {bulkantal}", f"Avvikelser:           {len(avvik)}"], GUL)
        return

    steg(nr - 1, nr, "Skriver data.js")
    csv = "Artikelnummer;Leverantörsnummer;Produktnamn;Plats;Bulkplats\n" + "\n".join(artiklar + galler)
    innehåll = ("// Skapad automatiskt av uppdatera_databas.py - ändra inte för hand.\n"
                f"window.LAGER_UPPDATERAD = '{stämpel}';\n"
                f"window.LAGER_CSV = `{js_sträng(csv)}`;\n")
    ut = os.path.abspath(a.ut)
    if os.path.exists(ut):
        shutil.copy2(ut, ut + '.bak')
    fd, tmp = tempfile.mkstemp(dir=os.path.dirname(ut), suffix='.tmp')
    with os.fdopen(fd, 'w', encoding='utf-8', newline='\n') as fh:
        fh.write(innehåll)
    os.replace(tmp, ut)
    with open(os.path.join(os.path.dirname(ut), 'avvikelser.txt'), 'w', encoding='utf-8') as fh:
        fh.write(f"Körning {stämpel}\n" + "\n".join(avvik) + "\n")
    ok(ut)
    if avvik:
        varning(f"{len(avvik)} rader hoppades over - se avvikelser.txt")

    steg(nr, nr, "Publicerar")
    adress = ""
    if a.publicera:
        publicera(ut, stämpel)
    elif a.publicera_api:
        publicera_api(ut, stämpel, a.github_repo, a.github_token, a.github_branch, a.github_api)
        if '/' in a.github_repo:
            agare, arkiv = a.github_repo.split('/', 1)
            adress = f"https://{agare.lower()}.github.io/" + ("" if arkiv.lower() == f"{agare.lower()}.github.io" else arkiv + "/")
    else:
        info("Hoppar over publicering (publicera= ar inte paslaget).")

    print()
    rader = [f("KLART!", GRON, FET), "",
             f"Forbrukningsartiklar: {len(artiklar)}", f"Galler:               {len(galler)}",
             f"Med bulkplats:        {bulkantal}",
             f"Totalt:               {len(artiklar) + len(galler)}", f"Tid:                  {stämpel}"]
    if avvik:
        rader.append(f(f"Avvikelser:           {len(avvik)}", GUL))
    if adress:
        rader += ["", "Appen:  " + f(adress, CYAN, FET)]
    ruta(rader, GRON if not avvik else GUL)


if __name__ == '__main__':
    starta_konsol()
    try:
        main()
    except SystemExit as e:
        if e.code not in (None, 0):
            fel(e.code)
        status = e.code
    except Exception as e:
        fel("Något gick fel: " + str(e))
        status = 1
    else:
        status = 0
    if getattr(sys, 'frozen', False) and len(sys.argv) == 1:
        input("\nTryck Retur for att stanga...")
    sys.exit(1 if status not in (None, 0) else 0)
