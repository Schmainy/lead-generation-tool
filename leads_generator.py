#!/usr/bin/env python3
"""
Google Maps Lead Generator
Sucht Unternehmen per Google Places API und exportiert sie als Excel oder CSV.
"""

import sys
import time
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ---------------------------------------------------------------------------
# ANSI color codes
# ---------------------------------------------------------------------------

RESET  = "\033[0m"
BOLD   = "\033[1m"
GREEN  = "\033[92m"
YELLOW = "\033[93m"
RED    = "\033[91m"
CYAN   = "\033[96m"
WHITE  = "\033[97m"

# ---------------------------------------------------------------------------
# Manuelle PLZ-Mappings für Berliner Bezirke
# ---------------------------------------------------------------------------

BERLIN_BEZIRKE = {
    "spandau":           ["13581","13583","13585","13587","13589",
                          "13591","13593","13595","13597","13599"],
    "mitte":             ["10115","10117","10119","10135","10178",
                          "10179"],
    "charlottenburg":    ["10585","10587","10589","10623","10625",
                          "10627","10629"],
    "schöneberg":        ["10777","10779","10781","10783","10785",
                          "10787","10789"],
    "kreuzberg":         ["10961","10963","10965","10967","10969",
                          "10999"],
    "prenzlauer berg":   ["10405","10407","10409","10435","10437",
                          "10439"],
    "friedrichshain":    ["10243","10245","10247","10249"],
    "neukölln":          ["12043","12045","12047","12049","12051",
                          "12053","12055","12057","12059"],
    "tempelhof":         ["12099","12101","12103","12105","12107",
                          "12109"],
    "steglitz":          ["12157","12159","12161","12163","12165",
                          "12167","12169"],
    "zehlendorf":        ["14109","14163","14165","14167","14169",
                          "14193","14195"],
    "reinickendorf":     ["13403","13405","13407","13409","13435",
                          "13437","13439","13465","13467","13469"],
    "pankow":            ["13125","13127","13129","13156","13158",
                          "13159","13187","13189"],
    "lichtenberg":       ["10315","10317","10318","10319","10365",
                          "10367","10369"],
    "treptow":           ["12435","12437","12439","12459","12487",
                          "12489"],
    "köpenick":          ["12555","12557","12559","12587","12589"],
    "marzahn":           ["12679","12681","12683","12685","12687",
                          "12689"],
    "hellersdorf":       ["12619","12621","12623","12627","12629"],
    "hohenschönhausen":  ["13051","13053","13055","13057","13059"],
}


def check_dependencies():
    packages = {
        "googlemaps": "googlemaps",
        "pandas":     "pandas",
        "openpyxl":   "openpyxl",
        "pgeocode":   "pgeocode",
        "numpy":      "numpy",
        "reportlab":  "reportlab",
    }
    missing = []
    for module, pip_name in packages.items():
        try:
            __import__(module)
        except ImportError:
            missing.append(pip_name)

    if missing:
        print(f"\n{RED}[!] Fehlende Pakete gefunden. Bitte installieren mit:{RESET}")
        print(f"    pip install {' '.join(missing)}")
        print("\nDanach dieses Script erneut starten.")
        sys.exit(1)

check_dependencies()

import html
import json
import googlemaps
import requests
import pandas as pd
import pgeocode
import numpy as np
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors as rl_colors
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from datetime import datetime
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# Konsolenausgabe auf UTF-8 zwingen. Ohne das brechen Box-Zeichen und '✓' mit
# UnicodeEncodeError ab, sobald die Ausgabe umgeleitet wird (Windows -> cp1252).
sys.stdout.reconfigure(encoding="utf-8", errors="replace")


# ---------------------------------------------------------------------------
# Einheitliche Export-Funktion (Excel + PDF) – Step 3
# ---------------------------------------------------------------------------

# Feste Spaltenbreiten fuer den Excel-Export (Zeicheneinheiten)
EXCEL_SPALTENBREITEN = {
    "Name":      35,
    "Adresse":   35,
    "Telefon":   18,
    "Website":   45,
    "Bewertung": 12,
    "PLZ":       10,
}
EXCEL_STANDARDBREITE = 18  # fuer alle nicht eigens festgelegten Spalten


def _excel_nachbearbeiten(xlsx_path, titel):
    """Fuegt der fertigen xlsx eine Titelzeile hinzu und formatiert Spalten/Umbrueche."""
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active

    ws.insert_rows(1)
    n_spalten = ws.max_column

    kopf = ws.cell(row=1, column=1, value=titel)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_spalten)
    kopf.font      = Font(bold=True, color="FFFFFF", size=12)
    kopf.fill      = PatternFill("solid", fgColor="1F4E79")
    kopf.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 22

    # Spaltentitel stehen nach dem Einfuegen in Zeile 2
    for spalte in range(1, n_spalten + 1):
        zelle = ws.cell(row=2, column=spalte)
        zelle.font = Font(bold=True)
        name = str(zelle.value or "")
        ws.column_dimensions[get_column_letter(spalte)].width = \
            EXCEL_SPALTENBREITEN.get(name, EXCEL_STANDARDBREITE)

    for zeile in ws.iter_rows(min_row=3, max_row=ws.max_row, max_col=n_spalten):
        for zelle in zeile:
            zelle.alignment = Alignment(wrap_text=True, vertical="top")

    ws.freeze_panes = "A3"  # Titelzeile + Spaltentitel bleiben stehen
    wb.save(xlsx_path)


def export_leads(df, business_type, category, city, state, country, base_dir,
                 filter_info=None):
    print(f"\n[*] Starte Export – {len(df)} Leads...")

    if df is None or len(df) == 0:
        print("[!] Keine Leads vorhanden – Export übersprungen.")
        return

    # Datei-/Ordnernamen bereinigen
    safe_type    = business_type.strip().replace(" ", "_")
    safe_city    = city.strip().replace(" ", "_")
    safe_state   = state.strip().replace(" ", "_") if state else "Unknown"
    safe_country = country.strip().replace(" ", "_") if country else "Germany"
    safe_cat     = category.strip().replace(" ", "_") if category else "Sonstige"

    # Zielpfad aufbauen: BASE_LEADS_DIR/country/state/city/category/
    folder = os.path.join(
        base_dir, safe_country, safe_state, safe_city, safe_cat
    )

    # Ordner anlegen
    try:
        os.makedirs(folder, exist_ok=True)
        print(f"[*] Ordner erstellt/gefunden: {folder}")
    except Exception as e:
        print(f"[!] Ordner konnte nicht erstellt werden: {e}")
        return

    # Suchdatum: haengt an beiden Dateinamen und steht in beiden Kopfzeilen
    datum    = datetime.now().strftime("%Y-%m-%d")
    datum_de = _datum_de(datum)

    # ── EXCEL ──────────────────────────────────────────────────────────────
    xlsx_path = os.path.join(folder, f"{safe_type}_{safe_city}_{datum}.xlsx")
    try:
        # PLZ-Spalte bleibt erhalten – sie ist fuer die Zuordnung im Export noetig
        export_df = df.copy()
        print(f"[*] Speichere nach: {xlsx_path}")
        export_df.to_excel(xlsx_path, index=False)
        _excel_nachbearbeiten(
            xlsx_path,
            f"Lead-Export: {business_type} | {city} | Suche vom {datum_de} | {len(df)} Leads",
        )
        print(f"[✓] Excel gespeichert: {xlsx_path}")
    except Exception as e:
        print(f"[!] Excel Fehler: {e}")

    # ── PDF ────────────────────────────────────────────────────────────────
    pdf_path = os.path.join(folder, f"{safe_type}_{safe_city}_{datum}.pdf")
    print(f"[*] Speichere nach: {pdf_path}")
    try:
        # Nach PLZ sortieren wenn Spalte vorhanden – _make_pdf erwartet sortierte Daten
        if 'PLZ' in df.columns:
            df_sorted = df.sort_values('PLZ').reset_index(drop=True)
        else:
            df_sorted = df.copy()

        _make_pdf(
            pdf_path,
            business_type,
            city,
            datum,
            df_sorted,
            gesamt=len(df),
            filter_info=filter_info,
        )
        print(f"[✓] PDF gespeichert: {pdf_path}")

    except Exception as e:
        print(f"[!] PDF Fehler: {e}")
        import traceback
        traceback.print_exc()


# ---------------------------------------------------------------------------
# Google Place Types mapping
# ---------------------------------------------------------------------------

GOOGLE_PLACE_TYPES = {
    # German
    "klempner": "plumber",
    "sanitär": "plumber",
    "elektriker": "electrician",
    "elektro": "electrician",
    "maler": "painter",
    "dachdecker": "roofing_contractor",
    "umzugsunternehmen": "moving_company",
    "umzug": "moving_company",
    "schlüsseldienst": "locksmith",
    "kfz-werkstatt": "car_repair",
    "autowerkstatt": "car_repair",
    "bäckerei": "bakery",
    "friseur": "hair_care",
    "arzt": "doctor",
    "zahnarzt": "dentist",
    "apotheke": "pharmacy",
    "restaurant": "restaurant",
    "café": "cafe",
    "hotel": "lodging",
    "rechtsanwalt": "lawyer",
    "steuerberater": "accounting",
    "buchhalter": "accounting",
    "tierarzt": "veterinary_care",
    "supermarkt": "supermarket",
    "physiotherapeut": "physiotherapist",
    "fitnessstudio": "gym",
    "reinigung": "laundry",
    "reinigungsservice": "laundry",
    # English
    "plumber": "plumber",
    "electrician": "electrician",
    "painter": "painter",
    "roofer": "roofing_contractor",
    "locksmith": "locksmith",
    "mechanic": "car_repair",
    "bakery": "bakery",
    "hairdresser": "hair_care",
    "doctor": "doctor",
    "dentist": "dentist",
    "pharmacy": "pharmacy",
    "lawyer": "lawyer",
    "accountant": "accounting",
    "vet": "veterinary_care",
    "gym": "gym",
}


# ---------------------------------------------------------------------------
# Hilfsfunktionen
# ---------------------------------------------------------------------------

def is_ja(s):
    """Gibt True zurück für deutsches/englisches Ja: j, ja, y, yes."""
    return isinstance(s, str) and s.strip().lower() in ("j", "ja", "y", "yes")


def is_nein(s):
    """Gibt True zurück für deutsches/englisches Nein: n, nein, no."""
    return isinstance(s, str) and s.strip().lower() in ("n", "nein", "no")


def load_dotenv_file(path=None):
    """Liest KEY=VALUE-Zeilen aus .env in os.environ (bestehende Variablen haben Vorrang)."""
    path = path or os.path.join(BASE_DIR, ".env")
    if not os.path.isfile(path):
        return
    try:
        with open(path, "r", encoding="utf-8") as fh:
            for line in fh:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                name, value = line.split("=", 1)
                os.environ.setdefault(name.strip(), value.strip().strip('"').strip("'"))
    except OSError as e:
        print(f"{YELLOW}[!] .env konnte nicht gelesen werden ({e}){RESET}")


def get_api_key():
    """Fragt den API-Key manuell ab."""
    print(f"{YELLOW}[!] API-Key nicht gefunden. Bitte manuell eingeben:{RESET}")
    key = input("API-Key: ").strip()
    if not key:
        print(f"{RED}[!] Kein API-Key eingegeben. Abbruch.{RESET}")
        sys.exit(1)
    return key


def load_keywords_database():
    """Lädt keywords.json aus demselben Verzeichnis wie das Script."""
    path = os.path.join(BASE_DIR, "keywords.json")  # BUGFIX: script_dir war redundant; BASE_DIR ist identisch
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"{YELLOW}[!] keywords.json nicht gefunden – Suche ohne Schlüsselwortdatenbank{RESET}")
        return {}
    except Exception as e:
        print(f"{YELLOW}[!] Fehler beim Laden von keywords.json: {e} – Suche ohne Schlüsselwortdatenbank{RESET}")
        return {}


def print_numbered_grid(items, cols=3, cell_width=28):
    """Gibt Elemente als nummeriertes Raster mit `cols` Spalten aus."""
    print()
    for i, item in enumerate(items):
        label = f"[{i + 1}]  {item}"
        if (i + 1) % cols == 0 or i == len(items) - 1:
            print(label)
        else:
            print(f"{label:<{cell_width}}", end="")
    print()


def select_business_type(keywords_db):
    """Interaktive Unternehmenstyp-Auswahl. Gibt (label, entry_or_none) oder BACK_SENTINEL zurück.
    entry_or_none ist ein dict {"keywords": [...], "related": [...]} oder None.
    BACK_SENTINEL wird zurückgegeben wenn der Nutzer 'z' beim ersten Prompt eingibt.
    """
    _CUSTOM = "Eigene Eingabe"

    def resolve(user_input, options):
        """Gibt das Element aus options zurück das per 1-basierter Nummer oder Name übereinstimmt."""
        if user_input.isdigit():
            idx = int(user_input) - 1
            if 0 <= idx < len(options):
                return options[idx]
            return None
        return next((o for o in options if o.lower() == user_input.lower()), None)

    while True:
        custom_term = None

        if keywords_db:
            categories    = list(keywords_db.keys())
            display_cats  = categories + [_CUSTOM]
            print_numbered_grid(display_cats)
            cat_raw = input("\nKategorie eingeben (Name oder Nummer, oder Enter zum Überspringen):\n(z = zurück)\n> ")
            if is_back(cat_raw):
                return BACK_SENTINEL  # zurück von Schritt 1 (Kategorie) → Aufrufer entscheidet
            cat_input = cat_raw.strip()

            if cat_input:
                resolved_cat = resolve(cat_input, display_cats)

                if resolved_cat is None:
                    print(f"{YELLOW}[!] Kategorie '{cat_input}' nicht gefunden.{RESET}")
                    continue

                if resolved_cat == _CUSTOM:
                    e_raw = input("Kategorie eingeben:\n(z = zurück)\n> ")
                    if is_back(e_raw):
                        print(f"{CYAN}[←] Zurück zum vorherigen Schritt...{RESET}")
                        continue  # zurück zur Kategorie-Liste
                    entered = e_raw.strip()
                    if not entered:
                        continue
                    # Prüfen ob der Text einer echten Kategorie entspricht; sonst freie Eingabe
                    resolved_cat = next((c for c in categories if c.lower() == entered.lower()), None)

                if resolved_cat is not None:
                    typ_back_to_kategorie = False
                    while True:
                        subcats       = list(keywords_db[resolved_cat].keys())
                        display_types = subcats + [_CUSTOM]
                        print_numbered_grid(display_types)
                        t_raw = input("\nUnternehmenstyp eingeben (Name oder Nummer, oder Enter für freie Eingabe):\n(z = zurück)\n> ")
                        if is_back(t_raw):
                            print(f"{CYAN}[←] Zurück zum vorherigen Schritt...{RESET}")
                            typ_back_to_kategorie = True
                            break  # zurück zur Kategorie-Liste
                        type_input = t_raw.strip()

                        if type_input:
                            matched_type = resolve(type_input, display_types)

                            if matched_type == _CUSTOM:
                                et_raw = input("Unternehmenstyp eingeben:\n(z = zurück)\n> ")
                                if is_back(et_raw):
                                    print(f"{CYAN}[←] Zurück zum vorherigen Schritt...{RESET}")
                                    continue  # zurück zur Unternehmenstyp-Liste
                                entered_type = et_raw.strip()
                                if not entered_type:
                                    continue
                                # DB-weite Suche nach dem eingegebenen Begriff
                                for cat, types in keywords_db.items():
                                    found = next((t for t in types if t.lower() == entered_type.lower()), None)
                                    if found:
                                        return (found, types[found])
                                print(f"{YELLOW}[!] '{entered_type}' nicht in der Datenbank gefunden. Trotzdem suchen? (j/n): {RESET}", end="")
                                if is_ja(input().strip()):
                                    return (entered_type, None)
                                continue

                            elif matched_type is not None:
                                return (matched_type, keywords_db[resolved_cat][matched_type])

                            else:
                                custom_term = type_input  # eingegeben aber nicht in Liste → freie Eingabe
                                break  # Typ-Schleife beenden, weiter zu freier Eingabe
                        else:
                            # leere Eingabe → weiter zu freier Eingabe
                            break

                    if typ_back_to_kategorie:
                        continue  # zurück zur Kategorie-Liste

                # resolved_cat ist None nach Eigene Eingabe → weiter zu freier Eingabe

        if custom_term is None:
            ct_raw = input("\nUnternehmenstyp eingeben:\n(z = zurück)\n> ")
            if is_back(ct_raw):
                if keywords_db:
                    print(f"{CYAN}[←] Zurück zum vorherigen Schritt...{RESET}")
                    continue  # zurück zur Kategorie-Liste
                else:
                    return BACK_SENTINEL  # keine keywords_db → dies ist der erste Schritt
            custom_term = ct_raw.strip()
            if not custom_term:
                continue

        # Exakte Übereinstimmung in der gesamten DB prüfen
        if keywords_db:
            for cat, types in keywords_db.items():
                matched = next((t for t in types if t.lower() == custom_term.lower()), None)
                if matched:
                    return (matched, types[matched])
            print(f"{YELLOW}[!] '{custom_term}' wurde nicht in der Schlüsselwortdatenbank gefunden.{RESET}")
            confirm = input(f"{YELLOW}Trotzdem mit diesem Begriff suchen? (j/n): {RESET}").strip()
            # BUGFIX: früher nur "y" akzeptiert; jetzt auch "j"/"ja" als Ja-Antwort
            if is_ja(confirm):
                return (custom_term, None)
            # else: Schleife wiederholen
        else:
            return (custom_term, None)


def eingabe_mit_default(frage, default=None):
    hinweis = f" [{default}]" if default else ""
    antwort = input(f"{frage}{hinweis}: ").strip()
    return antwort if antwort else default


# ---------------------------------------------------------------------------
# Zurück-Navigation Hilfsfunktionen
# ---------------------------------------------------------------------------

BACK_TOKENS   = ("z", "zurück", "zuruck")
BACK_SENTINEL = "__BACK__"

# Zähler für Einträge außerhalb Deutschlands (als Liste damit er aus Unterfunktionen heraus mutierbar ist)
_non_de_zaehler = [0]
# PLZ → Stadtname; vermeidet doppelte Geocoding-API-Aufrufe innerhalb einer Sitzung
_geocode_cache  = {}


def is_back(s):
    return isinstance(s, str) and s.strip().lower() in BACK_TOKENS


def input_with_back(prompt):
    """input() mit 'z = zurück' Hinweis. Gibt BACK_SENTINEL oder bereinigten Text zurück."""
    raw = input(f"{prompt}\n(z = zurück)\n> ")
    if is_back(raw):
        return BACK_SENTINEL
    return raw.strip()


def eingabe_mit_default_back(frage, default=None):
    """eingabe_mit_default mit Zurück-Option. Gibt BACK_SENTINEL, Wert oder Default zurück."""
    hinweis = f" [{default}]" if default else ""
    raw = input(f"{frage}{hinweis}:\n(z = zurück)\n> ")
    if is_back(raw):
        return BACK_SENTINEL
    antwort = raw.strip()
    return antwort if antwort else default


def haversine_array(lat1, lon1, lat2_arr, lon2_arr):
    """Haversine-Distanz von einem Punkt zu einem Array von Punkten (km)."""
    R = 6371.0
    phi1  = np.radians(lat1)
    phi2  = np.radians(lat2_arr)
    dphi  = phi2 - phi1
    dlam  = np.radians(lon2_arr) - np.radians(lon1)
    a = np.sin(dphi / 2)**2 + np.cos(phi1) * np.cos(phi2) * np.sin(dlam / 2)**2
    return 2 * R * np.arcsin(np.sqrt(np.clip(a, 0.0, 1.0)))  # clamp: float rounding can push a slightly above 1, making arcsin(sqrt(a)) return NaN


# ---------------------------------------------------------------------------
# PLZ-Ermittlung per pgeocode
# ---------------------------------------------------------------------------

def get_official_city_plz(city, country="DE"):
    """Findet alle PLZ für eine Stadt/einen Bezirk.
    Reihenfolge: 1) Berliner Bezirke-Dict, 2) exakter place_name-Match,
    3) place_name enthält Suchbegriff, 4) Breitsuche über place_name + state_name.
    """
    # 1) Berliner Bezirke: manuelle Zuordnung hat Vorrang
    if city.lower() in BERLIN_BEZIRKE:
        plz_list = BERLIN_BEZIRKE[city.lower()]
        print(f"[✓] Berliner Bezirk erkannt: {len(plz_list)} PLZ gefunden")
        return plz_list, None, None

    try:
        nomi = pgeocode.Nominatim(country)
        data = nomi._data.dropna(subset=["postal_code", "place_name"]).copy()
    except Exception as e:
        print(f"{RED}[!] Fehler beim Laden der PLZ-Datenbank: {e}{RESET}")
        return [], None, None

    # 2) Exakter Match (ursprüngliches Verhalten)
    mask    = data["place_name"].str.lower() == city.lower()
    matches = data[mask]

    # 3) Fallback: place_name enthält den Suchbegriff
    if matches.empty:
        mask    = data["place_name"].str.contains(city, case=False, na=False)
        matches = data[mask]

    # 4) Fallback: Breitsuche über place_name UND state_name
    if matches.empty:
        mask_place = data["place_name"].str.contains(city, case=False, na=False)
        if "state_name" in data.columns:
            mask_state = data["state_name"].str.contains(city, case=False, na=False)
            matches = data[mask_place | mask_state]
        else:
            matches = data[mask_place]

    if matches.empty:
        return [], None, None

    plz_list = sorted(set(
        str(p).split(".")[0].zfill(5)
        for p in matches["postal_code"].values
        if str(p).lower() not in ("nan", "none", "")
    ))

    coords = matches.dropna(subset=["latitude", "longitude"])
    if not coords.empty:
        center_lat = float(coords["latitude"].mean())
        center_lon = float(coords["longitude"].mean())
    else:
        center_lat, center_lon = None, None

    return plz_list, center_lat, center_lon


def _get_plz_by_radius(lat, lon, radius_km, country="DE"):
    """Gibt alle PLZ innerhalb von radius_km um (lat, lon) per pgeocode zurück."""
    try:
        nomi = pgeocode.Nominatim(country)
        data = nomi._data.dropna(subset=["latitude", "longitude"]).copy()
    except Exception as e:
        print(f"{RED}[!] Fehler beim Laden der PLZ-Datenbank: {e}{RESET}")
        return []

    distances = haversine_array(lat, lon, data["latitude"].values, data["longitude"].values)
    in_radius = data.loc[distances <= radius_km, "postal_code"].dropna().unique()
    return sorted(set(
        str(p).split(".")[0].zfill(5)
        for p in in_radius
        if str(p).lower() not in ("nan", "none", "")
    ))


def get_additional_plz_in_radius(gmaps, lat, lon, radius_km, exclude_plz, country="DE"):
    """Findet PLZ innerhalb radius_km die nicht in exclude_plz sind; gruppiert nach Stadt per Google Geocoding.
    Gibt dict zurück: Stadtname -> [plz, ...]
    """
    try:
        nomi = pgeocode.Nominatim(country)
        data = nomi._data.dropna(subset=["latitude", "longitude"]).copy()
    except Exception as e:
        print(f"{RED}[!] Fehler beim Laden der PLZ-Datenbank: {e}{RESET}")
        return {}

    distances = haversine_array(lat, lon, data["latitude"].values, data["longitude"].values)
    in_radius = data[distances <= radius_km].copy()

    additional = sorted(set(
        str(row["postal_code"]).split(".")[0].zfill(5)
        for _, row in in_radius.iterrows()
        if str(row["postal_code"]).lower() not in ("nan", "none", "")
        and str(row["postal_code"]).split(".")[0].zfill(5) not in exclude_plz
    ))

    plz_by_city = {}
    for plz in additional:
        if plz in _geocode_cache:
            plz_by_city.setdefault(_geocode_cache[plz], []).append(plz)
            continue
        try:
            result    = gmaps.geocode(plz + " Deutschland")
            city_name = "Unbekannt"
            if result:
                for comp in result[0].get("address_components", []):
                    ctypes = comp.get("types", [])
                    if "locality" in ctypes:
                        city_name = comp["long_name"]
                        break
                    elif "sublocality" in ctypes:
                        city_name = comp["long_name"]
                        break
                    elif "administrative_area_level_3" in ctypes:
                        city_name = comp["long_name"]
            _geocode_cache[plz] = city_name
            plz_by_city.setdefault(city_name, []).append(plz)
            time.sleep(0.05)
        except Exception as e:
            # BUGFIX: stille Exception; Nutzer sah keinen Hinweis auf fehlgeschlagenes Geocoding
            print(f"{YELLOW}[!] Geocoding-Fehler für PLZ {plz}: {e}{RESET}")
            _geocode_cache[plz] = "Unbekannt"
            plz_by_city.setdefault("Unbekannt", []).append(plz)

    return plz_by_city


# ---------------------------------------------------------------------------
# Known Businesses Datenbank
# ---------------------------------------------------------------------------

KNOWN_BUSINESSES_FILE = os.path.join(BASE_DIR, "Known_businesses.json")


def load_known_businesses():
    """Lädt Known_businesses.json; erstellt leere DB wenn Datei fehlt, leer oder ungültig."""
    # BUGFIX: fehlende Behandlung von leerer Datei und ungültigem JSON; nur FileNotFoundError wurde vorher abgefangen
    try:
        with open(KNOWN_BUSINESSES_FILE, "r", encoding="utf-8") as f:
            content = f.read().strip()
            if not content:
                raise ValueError("Leere Datei")
            data = json.loads(content)
    except (FileNotFoundError, ValueError, json.JSONDecodeError):
        data = {}
        try:
            with open(KNOWN_BUSINESSES_FILE, "w", encoding="utf-8") as f:
                json.dump({}, f)
        except OSError:
            pass
    return data


def save_known_businesses(leads, known_db):
    """Fügt neue Unternehmen aus leads in known_db ein (in-place) und schreibt auf Disk.
    Überspringt bereits vorhandene Place_IDs. Gibt Anzahl neu hinzugefügter Einträge zurück.
    """
    today     = pd.Timestamp.now().strftime("%Y-%m-%d")
    new_count = 0
    for lead in leads:
        pid = lead.get("Place_ID", "")
        if not pid or pid in known_db:
            continue
        address = lead.get("Adresse", "")
        city    = ""
        for part in reversed(address.split(",")):
            part = part.strip()
            if part and "Deutschland" not in part and "Germany" not in part:
                words = part.split()
                city  = " ".join(words[1:]) if (words and words[0].isdigit()) else part
                break
        known_db[pid] = {
            "name":                   lead.get("Name", ""),
            "formatted_address":      lead.get("Adresse", ""),
            "formatted_phone_number": lead.get("Telefon", "Keine Nummer"),
            "website":                lead.get("Website", "KEINE WEBSITE"),
            "rating":                 lead.get("Bewertung", ""),
            "user_ratings_total":     lead.get("Anzahl_Bewertungen", 0),
            "city":                   city,
            "date_added":             today,
        }
        new_count += 1
    try:
        with open(KNOWN_BUSINESSES_FILE, "w", encoding="utf-8") as f:
            json.dump(known_db, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"{RED}[!] Fehler beim Speichern der Known_businesses.json: {e}{RESET}")
    return new_count


# ---------------------------------------------------------------------------
# Suchverlauf Datenbank
# ---------------------------------------------------------------------------

SEARCH_HISTORY_FILE = os.path.join(BASE_DIR, "Search_history.json")


def load_search_history():
    """Lädt Search_history.json; erstellt leeren Verlauf wenn Datei fehlt, leer oder ungültig."""
    # BUGFIX: fehlende Behandlung von leerer Datei und ungültigem JSON; nur FileNotFoundError wurde vorher abgefangen
    try:
        with open(SEARCH_HISTORY_FILE, "r", encoding="utf-8") as f:
            content = f.read().strip()
            if not content:
                raise ValueError("Leere Datei")
            data = json.loads(content)
    except (FileNotFoundError, ValueError, json.JSONDecodeError):
        data = {}
        try:
            with open(SEARCH_HISTORY_FILE, "w", encoding="utf-8") as f:
                json.dump({}, f)
        except OSError:
            pass
    return data


def save_search_history(search_history, zusammenfassung):
    """Schreibt Suchergebnisse aus zusammenfassung in search_history und speichert auf Disk.
    Aktualisiert bestehende Einträge (datum + treffer). Gibt dict {term: anzahl_plz} zurück.
    """
    today  = pd.Timestamp.now().strftime("%Y-%m-%d")
    counts = {}
    seen   = set()
    for z in zusammenfassung:
        term = z["Branche"]
        plz  = str(z["PLZ"])
        key  = (term, plz)
        if key in seen:
            continue
        seen.add(key)
        search_history.setdefault(term, {})[plz] = {
            "datum":   today,
            "treffer": z["Gefunden"],
        }
        counts[term] = counts.get(term, 0) + 1
    try:
        with open(SEARCH_HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(search_history, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"{RED}[!] Fehler beim Speichern der Search_history.json: {e}{RESET}")
    return counts


def check_history_for_term(label, plz_liste, ort, search_history):
    """Prüft Suchverlauf für (label, plz_liste). Zeigt Warnungen und fragt Nutzer.
    Gibt gefilterte plz_liste zurück (alle, nur neue, oder leer) oder BACK_SENTINEL.
    """
    term_history     = search_history.get(label, {})
    already_searched = []   # [(plz, datum, treffer), ...]
    new_plz          = []

    for plz in plz_liste:
        if plz in term_history:
            e = term_history[plz]
            already_searched.append((plz, e["datum"], e["treffer"]))
        else:
            new_plz.append(plz)

    if not already_searched:
        return plz_liste   # kein Verlauf → direkt weiter

    def fmt(d):
        try:
            p = d.split("-")
            return f"{p[2]}.{p[1]}.{p[0]}"
        except Exception:
            return d

    # Stadt-Level-Prüfung: alle offiziellen Stadt-PLZ bereits durchsucht?
    official_plz, _, _ = get_official_city_plz(ort)
    if official_plz:
        searched_set = {plz for plz, _, _ in already_searched}
        if all(p in searched_set for p in official_plz):
            total  = sum(term_history[p]["treffer"] for p in official_plz if p in term_history)
            oldest = min(
                (term_history[p]["datum"] for p in official_plz if p in term_history),
                default=None,
            )
            if oldest:
                print(f"\n{YELLOW}[!] '{label}' wurde in {ort} bereits vollständig durchsucht")
                print(f"     ({fmt(oldest)}, {total} Treffer gesamt){RESET}")

    # PLZ-Level-Details
    all_already = (len(already_searched) == len(plz_liste))

    print(f"\n{YELLOW}[!] Folgende PLZ wurden für '{label}' bereits durchsucht:{RESET}")
    for plz, datum, treffer in already_searched:
        print(f"     {plz}  ({fmt(datum)}, {treffer} Treffer)")

    if new_plz:
        print(f"\n     Neue PLZ in dieser Suche: {', '.join(new_plz)}")

    print()

    if all_already:
        while True:
            raw = input(f"{YELLOW}     Trotzdem erneut suchen? (j/n): {RESET}").strip()
            if is_back(raw):
                return BACK_SENTINEL
            if is_ja(raw):
                return plz_liste
            if is_nein(raw):
                return []   # Term komplett überspringen
            print(f"{YELLOW}[!] Ungültige Eingabe.{RESET}")
    else:
        while True:
            raw = input(f"{YELLOW}     Bereits durchsuchte PLZ ebenfalls abfragen? (j/n): {RESET}").strip()
            if is_back(raw):
                return BACK_SENTINEL
            if is_ja(raw):
                return plz_liste
            if is_nein(raw):
                return new_plz
            print(f"{YELLOW}[!] Ungültige Eingabe.{RESET}")


# ---------------------------------------------------------------------------
# Suchgebiete Datenbank
# ---------------------------------------------------------------------------

SEARCH_AREAS_FILE = os.path.join(BASE_DIR, "Search_areas.json")


def load_search_areas():
    """Lädt Search_areas.json; gibt leeres dict zurück wenn Datei fehlt oder ungültig."""
    try:
        with open(SEARCH_AREAS_FILE, "r", encoding="utf-8") as f:
            content = f.read().strip()
            if not content:
                return {}
            return json.loads(content)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def save_search_area(name, city, radius_km, plz_list, city_groups, cities=None):
    """Speichert ein Suchgebiet in Search_areas.json.
    Wenn cities (dict city→[plz]) mit mehr als einem Eintrag übergeben wird,
    wird das kombinierte Format ohne radius_km gespeichert.
    """
    areas = load_search_areas()
    today = datetime.now().strftime("%Y-%m-%d")
    if cities and len(cities) > 1:
        areas[name] = {
            "cities":     {c: sorted(p) for c, p in cities.items()},
            "plz_list":   sorted(plz_list),
            "date_saved": today,
        }
    else:
        areas[name] = {
            "city":        city,
            "radius_km":   int(radius_km) if radius_km % 1 == 0 else radius_km,
            "plz_list":    sorted(plz_list),
            "city_groups": city_groups,
            "date_saved":  today,
        }
    try:
        with open(SEARCH_AREAS_FILE, "w", encoding="utf-8") as f:
            json.dump(areas, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"{RED}[!] Fehler beim Speichern der Search_areas.json: {e}{RESET}")


def _select_saved_area(areas):
    """Zeigt gespeicherte Suchgebiete zur Auswahl.
    Gibt (ort, plz_liste, include_known, city_display) oder None zurück.
    Unterstützt altes Format (city-Key) und kombiniertes Format (cities-Key).
    """
    area_names = list(areas.keys())
    while True:
        print(f"\nGespeicherte Suchgebiete:")
        for i, name in enumerate(area_names, 1):
            a = areas[name]
            try:
                date_str = datetime.strptime(a["date_saved"], "%Y-%m-%d").strftime("%d.%m.%Y")
            except Exception:
                date_str = a.get("date_saved", "")
            n_plz = len(a.get("plz_list", []))
            r_km  = a.get("radius_km", 0)
            city  = a.get("city", "") or " + ".join(a.get("cities", {}).keys())
            print(f" [{i}]  {name:<24} ({city}, {r_km}km,  {n_plz} PLZ,  gespeichert: {date_str})")
        print(f" (z = zurück)")
        sel = input("> ").strip()

        if is_back(sel):
            return None

        if not sel.isdigit() or not (1 <= int(sel) <= len(area_names)):
            print(f"{YELLOW}[!] Ungültige Eingabe.{RESET}")
            continue

        idx           = int(sel) - 1
        selected_name = area_names[idx]
        selected_area = areas[selected_name]

        # Kombiniertes Format (cities-Key) oder altes Format (city-Key)
        if "cities" in selected_area:
            cities_dict  = selected_area["cities"]
            city_keys    = list(cities_dict.keys())
            ort          = city_keys[0]
            city_display = " + ".join(city_keys)
            city_groups  = cities_dict
        else:
            ort          = selected_area.get("city", "")
            city_display = ort
            city_groups  = selected_area.get("city_groups", {ort: selected_area.get("plz_list", [])})

        if not ort:
            print(f"{YELLOW}[!] Suchgebiet '{selected_name}' hat keine gültige Stadt – übersprungen.{RESET}")
            continue
        plz_list = selected_area.get("plz_list", [])

        print(f"\n{GREEN}[✓] Suchgebiet geladen: {selected_name} ({len(plz_list)} PLZ){RESET}")
        if city_groups:
            max_len = max(len(c) for c in city_groups)
            for cname, cplz in city_groups.items():
                print(f"     {cname:<{max_len}}  {', '.join(cplz)}")

        conf = input("Suche starten? (j/n): ").strip()
        if is_back(conf) or is_nein(conf):
            continue
        if not is_ja(conf):
            print(f"{YELLOW}[!] Ungültige Eingabe.{RESET}")
            continue

        while True:
            inc_raw = input(
                "Sollen bereits bekannte Unternehmen ebenfalls angezeigt werden? (j/n): "
            ).strip()
            if is_back(inc_raw):
                break  # zurück zur Gebietsliste
            if is_ja(inc_raw) or is_nein(inc_raw):
                return ort, plz_list, is_ja(inc_raw), city_display
            print(f"{YELLOW}[!] Ungültige Eingabe.{RESET}")


# ---------------------------------------------------------------------------
# PLZ-Auswahl mit Erweiterungsmenü
# ---------------------------------------------------------------------------

def select_plz_for_city(gmaps, city):
    """Vollständige PLZ-Auswahl: offizielle PLZ → Erweiterung → Ausschluss → Bestätigung → Bekannte-Filter.
    Gibt (plz_liste, include_known, city_groups, radius_km) oder BACK_SENTINEL zurück.
    """
    country = "DE"
    print(f"\n{CYAN}[*] Suche offizielle PLZ für '{BOLD}{city}{RESET}{CYAN}' ...{RESET}")
    official_plz, center_lat, center_lon = get_official_city_plz(city, country)

    if not official_plz:
        print(f"{YELLOW}[!] Keine offiziellen PLZ für '{city}' gefunden. Geocodiere als Fallback ...{RESET}")
        try:
            result = gmaps.geocode(city)
            if not result:
                print(f"{RED}[!] Stadt '{city}' nicht gefunden.{RESET}")
                return BACK_SENTINEL
            loc = result[0]["geometry"]["location"]
            center_lat, center_lon = loc["lat"], loc["lng"]
        except Exception as e:
            print(f"{RED}[!] Geocoding-Fehler: {e}{RESET}")
            return BACK_SENTINEL
        official_plz = _get_plz_by_radius(center_lat, center_lon, 10.0, country)
        if not official_plz:
            print(f"{RED}[!] Keine PLZ gefunden.{RESET}")
            return BACK_SENTINEL
        print(f"{YELLOW}[!] Fallback: {len(official_plz)} PLZ im 10km Radius verwendet.{RESET}")
    else:
        print(f"{GREEN}[✓] Offizielle PLZ für {city} gefunden ({len(official_plz)} PLZ):{RESET}")
        print(f"     {', '.join(official_plz)}")

    # Gültige Koordinaten für Radius-Erweiterung sicherstellen
    if center_lat is None or center_lon is None:
        try:
            result = gmaps.geocode(city)
            if result:
                loc = result[0]["geometry"]["location"]
                center_lat, center_lon = loc["lat"], loc["lng"]
            else:
                # BUGFIX: stille Fallback-Zuweisung ohne Hinweis; Radius-Erweiterung wäre falsch berechnet worden
                print(f"{YELLOW}[!] Koordinaten für '{city}' nicht ermittelbar – Radius-Erweiterung deaktiviert{RESET}")
                center_lat, center_lon = 0.0, 0.0
        except Exception as e:
            # BUGFIX: stille Exception; Nutzer sah keinen Hinweis auf Geocoding-Fehler
            print(f"{YELLOW}[!] Geocoding-Fehler beim Ermitteln von Koordinaten: {e} – Radius-Erweiterung deaktiviert{RESET}")
            center_lat, center_lon = 0.0, 0.0

    sub_state        = "expansion"
    current_city_plz = list(official_plz)
    current_extra    = {}  # Stadtname -> [plz, ...] (Umkreis-basiert)
    extra_cities     = {}  # Stadtname -> [plz, ...] (via Option [2] hinzugefügt)
    selected_radius  = 0

    while True:
        # ── Erweiterungsmenü ─────────────────────────────────────────────
        if sub_state == "expansion":
            current_city_plz = list(official_plz)
            current_extra    = {}

            print(f"\nMöchtest du den Bereich erweitern?")
            print(f" [1]  Nein – nur {city} (empfohlen)")
            print(f" [2]  Weiteres Suchgebiet hinzufügen")
            print()
            print(f" Umkreis:")
            print(f" [3]  +5 km Umkreis")
            print(f" [4]  +10 km Umkreis")
            print(f" [5]  +15 km Umkreis")
            print(f" [6]  Eigene km-Zahl eingeben")
            print(f" (z = zurück)")
            choice = input("> ").strip().lower()

            if is_back(choice):
                return BACK_SENTINEL

            if choice == "1":
                sub_state = "confirmation"
                continue

            if choice == "2":
                v = input_with_back("Stadt oder Region")
                if v == BACK_SENTINEL or not v.strip():
                    continue
                add_city = v.strip()
                print(f"\n{CYAN}[*] Suche offizielle PLZ für '{add_city}' ...{RESET}")
                add_plz, _, _ = get_official_city_plz(add_city, country)
                if not add_plz:
                    try:
                        gc_result = gmaps.geocode(add_city)
                        if gc_result:
                            loc     = gc_result[0]["geometry"]["location"]
                            add_plz = _get_plz_by_radius(loc["lat"], loc["lng"], 10.0, country)
                            print(f"{YELLOW}[!] Fallback: {len(add_plz)} PLZ im 10km Radius für '{add_city}' verwendet.{RESET}")
                    except Exception as e:
                        print(f"{YELLOW}[!] Geocoding-Fehler: {e}{RESET}")
                if not add_plz:
                    print(f"{RED}[!] Keine PLZ für '{add_city}' gefunden.{RESET}")
                    continue
                extra_cities[add_city] = sorted(set(add_plz))
                print(f"{GREEN}[✓] {add_city}: {len(extra_cities[add_city])} PLZ hinzugefügt{RESET}")

                all_display    = {city: current_city_plz}
                all_display.update(extra_cities)
                total_combined = len(set(p for plz in all_display.values() for p in plz))
                max_len        = max(len(c) for c in all_display)
                print(f"\n{GREEN}[✓] Kombiniertes Suchgebiet:{RESET}")
                for cname, cplz in all_display.items():
                    preview = ', '.join(cplz[:5]) + (' ...' if len(cplz) > 5 else '')
                    print(f"     {cname:<{max_len}}  {preview}  ({len(cplz)} PLZ)")
                print(f"     Gesamt: {total_combined} PLZ")
                continue

            if choice in ("3", "4", "5", "6"):
                if   choice == "3": radius = 5.0
                elif choice == "4": radius = 10.0
                elif choice == "5": radius = 15.0
                else:
                    r_raw = input("Eigene km-Zahl: ").strip()
                    if is_back(r_raw):
                        continue
                    try:
                        radius = float(r_raw)
                    except ValueError:
                        print(f"{YELLOW}[!] Ungültige Eingabe.{RESET}")
                        continue

                selected_radius = radius
                print(f"\n{CYAN}[*] Suche zusätzliche PLZ im {radius:.0f}km Umkreis ...{RESET}")
                exclude_set   = set(current_city_plz) | set(p for plz in extra_cities.values() for p in plz)
                current_extra = get_additional_plz_in_radius(
                    gmaps, center_lat, center_lon, radius, exclude_set, country
                )

                if current_extra:
                    total_extra = sum(len(v) for v in current_extra.values())
                    total_all   = len(current_city_plz) + total_extra
                    max_len     = max(len(c) for c in current_extra)
                    print(f"\n{GREEN}[+] Zusätzliche PLZ im {radius:.0f}km Umkreis:{RESET}")
                    for cname, plz_list in sorted(current_extra.items()):
                        print(f"     {cname:<{max_len}}  →  {', '.join(plz_list)}")
                    print(f" Gesamt: {total_all} PLZ ({len(current_city_plz)} {city} + {total_extra} Umkreis)")
                else:
                    print(f"{YELLOW}[!] Keine zusätzlichen PLZ im Umkreis gefunden.{RESET}")

                sub_state = "exclusion"
                continue

            print(f"{YELLOW}[!] Ungültige Eingabe.{RESET}")

        # ── Ausschlussmenü ────────────────────────────────────────────────
        elif sub_state == "exclusion":
            print(f"\n{'━'*42}")
            print(f"Möchtest du etwas ausschließen?")
            print(f" [1]  Einzelne PLZ entfernen")
            print(f" [2]  Gesamten Ort/Bezirk entfernen")
            print(f" [3]  Nichts entfernen – weiter")
            print(f"{'━'*42}")
            ex_choice = input("> ").strip()

            if is_back(ex_choice):
                sub_state = "expansion"
                continue

            if ex_choice == "1":
                plz_raw = input("Welche PLZ entfernen? (kommagetrennt): ").strip()
                if is_back(plz_raw):
                    continue
                to_remove = set(p.strip().zfill(5) for p in plz_raw.split(",") if p.strip())
                current_city_plz = [p for p in current_city_plz if p not in to_remove]
                for cname in list(current_extra.keys()):
                    current_extra[cname] = [p for p in current_extra[cname] if p not in to_remove]
                    if not current_extra[cname]:
                        del current_extra[cname]
                print(f"{GREEN}[✓] Entfernt: {', '.join(sorted(to_remove))}{RESET}")

            elif ex_choice == "2":
                city_raw = input("Welchen Ort entfernen? (kommagetrennt, z.B. Teltow, Werder): ").strip()
                if is_back(city_raw):
                    continue
                to_remove_cities = [c.strip().lower() for c in city_raw.split(",") if c.strip()]
                removed_info = []
                for rem_city in to_remove_cities:
                    if rem_city == city.lower():
                        removed_info.append(f"{city} ({', '.join(current_city_plz)})")
                        current_city_plz = []
                    else:
                        for cname in list(current_extra.keys()):
                            if cname.lower() == rem_city:
                                removed_info.append(f"{cname} ({', '.join(current_extra[cname])})")
                                del current_extra[cname]
                                break
                if removed_info:
                    print(f"{GREEN}[✓] Entfernt: {', '.join(removed_info)}{RESET}")
                else:
                    print(f"{YELLOW}[!] Keine passenden Orte gefunden.{RESET}")

            elif ex_choice == "3":
                sub_state = "confirmation"

        # ── Finale Bestätigung ────────────────────────────────────────────
        elif sub_state == "confirmation":
            extra_flat        = [p for plz_list in current_extra.values() for p in plz_list]
            extra_cities_flat = [p for plz_list in extra_cities.values() for p in plz_list]
            all_plz           = sorted(set(current_city_plz + extra_cities_flat + extra_flat))
            total             = len(all_plz)

            # BUGFIX: leere PLZ-Liste wurde ohne Warnung bestätigt; führte zu stummem Abbruch später
            if total == 0:
                print(f"{YELLOW}[!] Keine PLZ ausgewählt – alle PLZ wurden ausgeschlossen. Bitte Auswahl neu starten.{RESET}")
                sub_state = "expansion"
                continue

            max_len = max(
                [len(city)] + [len(c) for c in extra_cities] + [len(c) for c in current_extra],
                default=len(city)
            )

            print(f"\n{GREEN}[✓] Finale PLZ-Liste ({total} PLZ):{RESET}")
            if current_city_plz:
                print(f"     {city:<{max_len}}  {', '.join(current_city_plz)}")
            for cname, plz_list in sorted(extra_cities.items()):
                print(f"     {cname:<{max_len}}  {', '.join(plz_list)}")
            for cname, plz_list in sorted(current_extra.items()):
                print(f"     {cname:<{max_len}}  {', '.join(plz_list)}")

            confirm = input("Suche starten? (j/n): ").strip()
            # BUGFIX: früher nur "n" akzeptiert; "nein" und Zurück-Token wurden ignoriert
            if is_back(confirm) or is_nein(confirm):
                sub_state = "expansion"
                continue
            # BUGFIX: früher nur "j" akzeptiert; "ja" wurde als ungültige Eingabe behandelt
            if is_ja(confirm):
                sub_state = "known"
                continue
            print(f"{YELLOW}[!] Ungültige Eingabe.{RESET}")

        # ── Bekannte Unternehmen Frage ────────────────────────────────────
        elif sub_state == "known":
            # BUGFIX: fehlende Validierungsschleife; ungültige Eingabe setzte include_known=False ohne Rückfrage
            while True:
                inc_raw = input(
                    "Sollen bereits bekannte Unternehmen ebenfalls angezeigt werden? (j/n): "
                ).strip()
                if is_back(inc_raw):
                    sub_state = "confirmation"
                    break
                if is_ja(inc_raw) or is_nein(inc_raw):
                    include_known = is_ja(inc_raw)
                    final_plz = sorted(set(
                        current_city_plz
                        + [p for plz_list in extra_cities.values() for p in plz_list]
                        + [p for plz_list in current_extra.values() for p in plz_list]
                    ))
                    built_groups = {}
                    if current_city_plz:
                        built_groups[city] = list(current_city_plz)
                    for cname, cplz in extra_cities.items():
                        built_groups[cname] = list(cplz)
                    for cname, cplz in current_extra.items():
                        built_groups[cname] = list(cplz)
                    return final_plz, include_known, built_groups, selected_radius
                print(f"{YELLOW}[!] Ungültige Eingabe. Bitte 'j' oder 'n' eingeben.{RESET}")


# ---------------------------------------------------------------------------
# Places-API-Suche
# ---------------------------------------------------------------------------

# Google Places Text Search liefert maximal 3 Seiten à 20 Ergebnisse pro Query.
MAX_SEITEN = 3

# Places API (New). Der Legacy-Endpunkt (maps.googleapis.com/.../textsearch/json) stellt
# inzwischen Tokens im neuen Format aus, die er selbst nicht mehr annimmt -> INVALID_REQUEST
# auf Seite 2. Deshalb laufen Seite 1 und alle Folgeseiten über diesen Endpunkt.
PLACES_SEARCHTEXT_URL = "https://places.googleapis.com/v1/places:searchText"

# Die neue API liefert nur die hier angeforderten Felder; nextPageToken steht top-level.
PLACES_FIELD_MASK = (
    "places.id,"
    "places.displayName,"
    "places.formattedAddress,"
    "places.nationalPhoneNumber,"
    "places.websiteUri,"
    "places.rating,"
    "places.userRatingCount,"
    "places.types,"
    "nextPageToken"
)


def _place_neu_zu_legacy(place):
    """Mappt einen Treffer der Places API (New) auf die Feldnamen des Legacy-Formats,
    damit die Auswertung in suche_leads unverändert bleiben kann."""
    return {
        "place_id":               place.get("id", ""),
        "name":                   (place.get("displayName") or {}).get("text", ""),
        "formatted_address":      place.get("formattedAddress", ""),
        "formatted_phone_number": place.get("nationalPhoneNumber", ""),
        "website":                place.get("websiteUri", ""),
        "rating":                 place.get("rating", ""),
        "user_ratings_total":     place.get("userRatingCount", 0),
        "types":                  place.get("types", []),
    }


def _places_textsearch(gmaps, query, page_token=None):
    """Ein Text-Search-Aufruf gegen die Places API (New).

    Gibt die Antwort im Legacy-Format zurück ({"results": [...], "next_page_token": ...}).
    Wirft bei HTTP- oder API-Fehlern eine Exception; die Behandlung erfolgt beim Aufrufer.
    """
    headers = {
        "Content-Type":     "application/json",
        "X-Goog-Api-Key":   gmaps.key,
        "X-Goog-FieldMask": PLACES_FIELD_MASK,
    }
    body = {"textQuery": query, "languageCode": "de", "maxResultCount": 20}
    if page_token:
        body["pageToken"] = page_token

    response = requests.post(PLACES_SEARCHTEXT_URL, headers=headers, json=body, timeout=30)

    # Die neue API meldet Fehler über den HTTP-Status plus ein "error"-Objekt im Body.
    if response.status_code != 200:
        try:
            meldung = (response.json().get("error") or {}).get("message", "")
        except ValueError:
            meldung = response.text[:200]
        raise RuntimeError(f"HTTP {response.status_code}: {meldung or '<keine Meldung>'}")

    daten = response.json()
    return {
        "results":         [_place_neu_zu_legacy(p) for p in daten.get("places", [])],
        "next_page_token": daten.get("nextPageToken"),
    }


def suche_leads(gmaps, branche, plz, city, max_results=60, fortschritt="", known_db=None):
    print(f"\n{CYAN}[*] Suche '{BOLD}{branche}{RESET}{CYAN}' in PLZ {plz} {fortschritt}{RESET}")
    if known_db is None:
        known_db = {}
    leads = []
    query = f"{plz} {city} {branche}"
    try:
        anfrage = _places_textsearch(gmaps, query)
    except Exception as e:  # deckt HTTP-/API-Fehler, Timeout und Netzwerkausfälle ab
        print(f"{RED}    [!] API-Fehler: {e}{RESET}")
        return leads

    seite = 1
    while anfrage and len(leads) < max_results:
        ergebnisse = anfrage.get("results", [])
        print(f"    Seite {seite}/{MAX_SEITEN}: {len(ergebnisse)} Einträge ...")

        for place in ergebnisse:
            if len(leads) >= max_results:
                break
            pid_quick = place.get("place_id", "")
            if pid_quick and pid_quick in known_db:
                entry   = known_db[pid_quick]
                address = entry.get("formatted_address", "")
                if "Deutschland" not in address and "Germany" not in address:
                    _non_de_zaehler[0] += 1
                else:
                    leads.append({
                        "Name":               entry.get("name", ""),
                        "Adresse":            address,
                        "Telefon":            entry.get("formatted_phone_number", "Keine Nummer"),
                        "Website":            entry.get("website", "KEINE WEBSITE"),
                        "Bewertung":          entry.get("rating", ""),
                        "Anzahl_Bewertungen": entry.get("user_ratings_total", 0),
                        "Oeffnungszeiten":    "",
                        "Unternehmenstyp":    "",
                        "PLZ":                plz,
                        "Branche":            branche,
                        "Place_ID":           pid_quick,
                        "Bekannt":            "Ja",
                    })
                continue
            try:
                details = gmaps.place(
                    place["place_id"],
                    fields=[
                        "name", "formatted_address", "formatted_phone_number",
                        "website", "rating", "user_ratings_total",
                        "opening_hours", "place_id", "type"  # BUGFIX: "types" ist kein gültiger field-Name im Legacy-API; führte zu Fehler bei jedem Detail-Abruf
                    ]
                )["result"]

                oeffnungszeiten = ""
                if "opening_hours" in details:
                    perioden = details["opening_hours"].get("weekday_text", [])
                    oeffnungszeiten = " | ".join(perioden)

                types_list = details.get("types", [])

                place_id = details.get("place_id", "")
                if not place_id:
                    continue  # leere Place_ID: Deduplizierung würde alle solchen Zeilen zusammenfassen, Daten verlieren

                # Deutschland-Filter: Einträge ohne deutschen Adresszusatz überspringen
                address = details.get("formatted_address", "")
                if "Deutschland" not in address and "Germany" not in address:
                    _non_de_zaehler[0] += 1
                    time.sleep(0.1)
                    continue

                leads.append({
                    "Name":               details.get("name", ""),
                    "Adresse":            address,
                    "Telefon":            details.get("formatted_phone_number", "Keine Nummer"),
                    "Website":            details.get("website", "KEINE WEBSITE"),
                    "Bewertung":          details.get("rating", ""),
                    "Anzahl_Bewertungen": details.get("user_ratings_total", 0),
                    "Oeffnungszeiten":    oeffnungszeiten,
                    "Unternehmenstyp":    ", ".join(types_list),
                    "PLZ":                plz,
                    "Branche":            branche,
                    "Place_ID":           place_id,
                })
                time.sleep(0.1)
            except Exception as e:
                print(f"{RED}    [!] Fehler bei Detail-Abruf: {e} – Basis-Daten werden gespeichert{RESET}")
                place_id = place.get("place_id", "")
                if not place_id:
                    continue
                # BUGFIX: gescheiterter Detail-Abruf führte zum kompletten Verwerfen des Eintrags
                # und falscher Inkrementierung des Deutschland-Zählers; jetzt: Basis-Daten aus
                # Such-Ergebnis speichern und nur überspringen wenn keine deutsche Adresse erkennbar
                fallback_address = place.get("formatted_address") or place.get("vicinity", "")
                if "Deutschland" not in fallback_address and "Germany" not in fallback_address:
                    # BUGFIX: ohne verifizierten deutschen Adresszusatz → Deutschland-Zähler erhöhen
                    _non_de_zaehler[0] += 1
                    time.sleep(0.1)  # BUGFIX: fehlender Sleep im Fehlerzweig konnte Rate-Limit auslösen
                    continue
                leads.append({
                    "Name":               place.get("name", ""),
                    "Adresse":            fallback_address,
                    "Telefon":            "Keine Nummer",
                    "Website":            "KEINE WEBSITE",
                    "Bewertung":          place.get("rating", ""),
                    "Anzahl_Bewertungen": place.get("user_ratings_total", 0),
                    "Oeffnungszeiten":    "",
                    "Unternehmenstyp":    ", ".join(place.get("types", [])),
                    "PLZ":                plz,
                    "Branche":            branche,
                    "Place_ID":           place_id,
                })
                time.sleep(0.1)  # BUGFIX: fehlender Sleep im Fehlerzweig konnte Rate-Limit auslösen
                continue

        next_token = anfrage.get("next_page_token")
        if next_token and len(leads) < max_results and seite < MAX_SEITEN:
            print(f"    Seite {seite + 1}/{MAX_SEITEN} wird geladen ...")
            time.sleep(5)  # Token ist direkt nach Ausstellung noch nicht gültig
            try:
                anfrage = _places_textsearch(gmaps, query, page_token=next_token)
            except Exception as e:
                print(f"{RED}    [!] Fehler beim Abruf der nächsten Seite: {e}"
                      f" – Paginierung gestoppt.{RESET}")
                break  # Paginierung stoppen, aber bisherige Leads behalten
            seite += 1
        else:
            break

    print()
    return leads


# ---------------------------------------------------------------------------
# Filter
# ---------------------------------------------------------------------------

def filter_anwenden(df, nur_ohne_website, nur_mit_website, min_bewertung, max_bewertung,
                    min_bewertungen, ausschluss_keywords, ausschluss_typen):
    if nur_ohne_website:
        df = df[df["Website"] == "KEINE WEBSITE"]
    if nur_mit_website:
        df = df[df["Website"] != "KEINE WEBSITE"]
    if min_bewertung is not None:
        df = df[pd.to_numeric(df["Bewertung"], errors="coerce") >= min_bewertung]
    if max_bewertung is not None:
        df = df[pd.to_numeric(df["Bewertung"], errors="coerce") <= max_bewertung]
    if min_bewertungen is not None:
        df = df[pd.to_numeric(df["Anzahl_Bewertungen"], errors="coerce") >= min_bewertungen]
    for kw in ausschluss_keywords:
        kw_lower = kw.strip().lower()
        if kw_lower:
            df = df[~df["Name"].str.lower().str.contains(kw_lower, na=False)]
    for typ in ausschluss_typen:
        typ_lower = typ.strip().lower()
        if typ_lower:
            df = df[~df["Name"].str.lower().str.contains(typ_lower, na=False)]
    return df


# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------

COLUMNS = [
    "Name", "Adresse", "Telefon", "Website", "Bewertung",
    "Anzahl_Bewertungen", "Oeffnungszeiten", "Unternehmenstyp",
    "PLZ", "Branche", "Place_ID",
]


def exportieren(df, format_wahl, branchen, ort, output_dir, cols=None):
    if cols is None:
        cols = COLUMNS
    branchen_str    = "_".join(safe_name(b) for b in branchen[:3])  # safe_name entfernt Windows-inkompatible Zeichen
    dateiname_basis = os.path.join(output_dir, f"leads_{branchen_str}_{safe_name(ort)}")

    if format_wahl == "1":
        dateiname = dateiname_basis + ".xlsx"
        os.makedirs(os.path.dirname(dateiname), exist_ok=True)
        with pd.ExcelWriter(dateiname, engine="openpyxl") as writer:
            # Ein Sheet pro PLZ
            for plz, gruppe in df.groupby("PLZ", sort=True):
                sheet_name = str(plz)[:31]
                gruppe[cols].to_excel(writer, sheet_name=sheet_name, index=False)
            # Kombiniertes ALL-Sheet (bereits dedupliziert)
            df[cols].to_excel(writer, sheet_name="ALL", index=False)
    else:
        dateiname = dateiname_basis + ".csv"
        os.makedirs(os.path.dirname(dateiname), exist_ok=True)
        df[cols].to_csv(dateiname, index=False, encoding="utf-8-sig")

    return dateiname


# ---------------------------------------------------------------------------
# PDF-Export
# ---------------------------------------------------------------------------

BASE_LEADS_DIR = None  # wird zur Laufzeit in main() gesetzt


def safe_name(s):
    """Leerzeichen → Unterstriche, Windows-inkompatible Zeichen entfernen."""
    s = s.replace(" ", "_")
    for ch in r'\/:*?"<>|':
        s = s.replace(ch, "")
    return s


def get_location_info(gmaps, city):
    """Extrahiert Land, Bundesland und Stadtname per Geocoding."""
    try:
        result = gmaps.geocode(city)
    except Exception as e:
        # BUGFIX: stille Exception; Nutzer sah keinen Hinweis warum Ordnerstruktur "Unknown" enthält
        print(f"{YELLOW}[!] Standortermittlung fehlgeschlagen: {e} – Ordner wird als 'Unknown' angelegt{RESET}")
        return {"country": "Unknown", "state": "Unknown", "locality": city}
    if not result:
        # BUGFIX: stille Rückgabe ohne Hinweis; Nutzer wunderte sich über "Unknown"-Ordner im Dateisystem
        print(f"{YELLOW}[!] Keine Geocoding-Ergebnisse für '{city}' – Ordner wird als 'Unknown' angelegt{RESET}")
        return {"country": "Unknown", "state": "Unknown", "locality": city}
    info = {"country": "Unknown", "state": "Unknown", "locality": city}
    for comp in result[0].get("address_components", []):
        types = comp.get("types", [])
        if "country" in types:
            info["country"] = comp["long_name"]
        elif "administrative_area_level_1" in types:
            info["state"] = comp["long_name"]
        elif "locality" in types:
            info["locality"] = comp["long_name"]
        elif "postal_town" in types and info["locality"] == city:
            info["locality"] = comp["long_name"]
    return info


def _datum_de(datum):
    """'2026-08-19' -> '19.08.2026'. Unbekannte Formate werden unveraendert zurueckgegeben."""
    try:
        return datetime.strptime(str(datum), "%Y-%m-%d").strftime("%d.%m.%Y")
    except ValueError:
        return str(datum)


def _url_mit_umbruch(url, max_breite, fontname="Helvetica", fontsize=9):
    """Bricht lange URLs an '/' oder '-' um und liefert XML-escapten Text mit <br/>.

    Reportlab trennt sonst nur an Leerzeichen; eine URL ist ein einziges 'Wort' und
    wuerde an beliebiger Stelle zerhackt oder ueber die Spalte hinauslaufen.
    """
    text = str(url or "")
    if not text:
        return ""

    # In Stuecke zerlegen, Trennzeichen bleibt jeweils am Ende des Stuecks
    teile, aktuell = [], ""
    for ch in text:
        aktuell += ch
        if ch in "/-":
            teile.append(aktuell)
            aktuell = ""
    if aktuell:
        teile.append(aktuell)

    zeilen, zeile = [], ""
    for teil in teile:
        if zeile and pdfmetrics.stringWidth(zeile + teil, fontname, fontsize) > max_breite:
            zeilen.append(zeile)
            zeile = teil
        else:
            zeile += teil
    if zeile:
        zeilen.append(zeile)

    return "<br/>".join(html.escape(z) for z in zeilen)


def _make_pdf(pdf_path, branchen_label, locality, datum, df_sorted,
              plz_liste=None, gesamt=None, filter_info=None):
    """Erstellt eine einzelne PDF-Datei mit allen Leads, sortiert nach PLZ.

    plz_liste / gesamt / filter_info speisen nur die Kopfzeile und sind optional,
    damit bestehende Aufrufer (exportieren_pdf, exportieren_pdf_pro_typ) unveraendert
    weiterlaufen. Fehlende Angaben werden aus df_sorted abgeleitet.
    """
    from reportlab.lib.pagesizes import landscape, A4
    from reportlab.lib import colors
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
    from reportlab.lib.styles import ParagraphStyle

    page_w, page_h = landscape(A4)

    # ── Kopfzeilen-Texte vorbereiten ───────────────────────────────────────
    datum_de = _datum_de(datum)

    if plz_liste is None and "PLZ" in df_sorted.columns:
        plz_liste = sorted({str(p) for p in df_sorted["PLZ"] if str(p).strip()})
    plz_liste = [str(p) for p in (plz_liste or [])]
    if len(plz_liste) > 12:
        plz_text = ", ".join(plz_liste[:12]) + f" (+{len(plz_liste) - 12} weitere)"
    else:
        plz_text = ", ".join(plz_liste) if plz_liste else "–"

    if gesamt is None:
        gesamt = len(df_sorted)

    filter_info   = filter_info or {}
    ohne_web_txt  = "Ja" if filter_info.get("nur_ohne_website") else "Nein"
    min_bew       = filter_info.get("min_bewertung")
    min_bew_txt   = str(min_bew) if min_bew is not None else "–"

    kopf_1 = f"Lead-Export – {branchen_label} | {locality}"
    kopf_2 = f"Suche vom {datum_de} | PLZ-Gebiet: {plz_text}"
    kopf_3 = (f"Gesamt: {gesamt} Leads | Filter: ohne Website = {ohne_web_txt}, "
              f"min. Bewertung = {min_bew_txt}")

    # topMargin so gewaehlt, dass die Tabelle unter der Trennlinie beginnt
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=landscape(A4),
        leftMargin=1*cm, rightMargin=1*cm,
        topMargin=2.35*cm, bottomMargin=1.2*cm,
    )

    def draw_page_deco(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica-Bold", 11)
        canvas.setFillColor(colors.black)
        canvas.drawString(1*cm, page_h - 1.00*cm, kopf_1)
        canvas.setFont("Helvetica", 9)
        canvas.setFillColor(colors.HexColor("#666666"))
        canvas.drawString(1*cm, page_h - 1.45*cm, kopf_2)
        canvas.drawString(1*cm, page_h - 1.80*cm, kopf_3)
        # Trennlinie unter dem Header
        canvas.setStrokeColor(colors.HexColor("#999999"))
        canvas.setLineWidth(0.6)
        canvas.line(1*cm, page_h - 2.05*cm, page_w - 1*cm, page_h - 2.05*cm)
        canvas.setFont("Helvetica", 8)
        canvas.drawRightString(page_w - 1*cm, 0.4*cm, f"Seite {doc.page}")
        canvas.restoreState()

    def _esc(v):
        # Zellwerte mit '<', '>' oder '&' würden einen XML-Parse-Fehler in Paragraph verursachen
        return html.escape(str(v)) if v is not None else ""

    cell_style  = ParagraphStyle("cell",  fontName="Helvetica",      fontSize=9,  leading=11)
    hcell_style = ParagraphStyle("hcell", fontName="Helvetica-Bold", fontSize=11, leading=13,
                                 textColor=colors.white)
    div_style   = ParagraphStyle("div",   fontName="Helvetica-Bold", fontSize=9,  leading=11,
                                 textColor=colors.black)

    # Breiten in Punkt; nutzbar sind bei A4 quer mit 1cm Raendern ca. 785pt
    col_headers = ["Name", "Adresse", "Telefon", "Website", "Bewertung"]
    col_widths  = [200, 180, 90, 180, 60]
    web_breite  = col_widths[3] - 8  # abzueglich LEFT-/RIGHTPADDING

    table_data = [[Paragraph(h, hcell_style) for h in col_headers]]
    style_cmds = [
        ("BACKGROUND",    (0, 0), (-1, 0),  colors.HexColor("#2d6a9f")),
        ("GRID",          (0, 0), (-1, -1), 0.4, colors.HexColor("#cccccc")),
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING",    (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING",   (0, 0), (-1, -1), 4),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 4),
    ]

    current_plz  = None
    data_row_idx = 0  # zählt nur Datenzeilen, steuert alternierende Zeilenfarbe

    for _, row in df_sorted.iterrows():
        plz = str(row.get("PLZ", ""))

        if plz != current_plz:
            # Trennzeile mit hellgrauem Hintergrund ueber alle Spalten
            div_i = len(table_data)
            table_data.append([Paragraph(f"PLZ {plz}", div_style), "", "", "", ""])
            style_cmds.append(("SPAN",       (0, div_i), (-1, div_i)))
            style_cmds.append(("BACKGROUND", (0, div_i), (-1, div_i), colors.HexColor("#e0e0e0")))
            current_plz  = plz
            data_row_idx = 0

        data_i = len(table_data)
        table_data.append([
            Paragraph(_esc(row.get("Name", "")),                          cell_style),
            Paragraph(_esc(row.get("Adresse", "")),                       cell_style),
            Paragraph(_esc(row.get("Telefon", "")),                       cell_style),
            Paragraph(_url_mit_umbruch(row.get("Website", ""), web_breite), cell_style),
            Paragraph(_esc(row.get("Bewertung", "")),                     cell_style),
        ])
        if data_row_idx % 2 == 1:
            style_cmds.append(("BACKGROUND", (0, data_i), (-1, data_i), colors.HexColor("#e8f4f8")))
        data_row_idx += 1

    table = Table(table_data, colWidths=col_widths, repeatRows=1)
    table.setStyle(TableStyle(style_cmds))
    doc.build([table], onFirstPage=draw_page_deco, onLaterPages=draw_page_deco)


def exportieren_pdf(df, branchen, location_info):
    """Erstellt eine einzelne PDF-Datei mit allen Leads im Zielordner."""
    land  = safe_name(location_info["country"])
    staat = safe_name(location_info["state"])
    stadt = safe_name(location_info["locality"])
    datum = pd.Timestamp.now().strftime("%Y-%m-%d")

    ordner = os.path.join(BASE_LEADS_DIR, land, staat, stadt)
    os.makedirs(ordner, exist_ok=True)

    # Deduplizierung nach Name+Adresse, dann nach PLZ aufsteigend sortieren
    df_pdf = df.drop_duplicates(subset=["Name", "Adresse"], keep="first")
    df_pdf = df_pdf.sort_values("PLZ", ascending=True)

    branchen_str   = "_".join(safe_name(b) for b in branchen[:3])
    branchen_label = " / ".join(branchen)
    pdf_name       = f"{branchen_str}_{stadt}.pdf"
    pdf_path       = os.path.join(ordner, pdf_name)

    print(f"\n{CYAN}[*] Erstelle PDF ...{RESET}")
    try:
        os.makedirs(os.path.dirname(pdf_path), exist_ok=True)
        _make_pdf(pdf_path, branchen_label, location_info["locality"], datum, df_pdf)
        rel = os.path.join(land, staat, stadt, pdf_name)
        print(f"{GREEN}[✓] PDF erstellt: {rel}{RESET}")
    except Exception as e:
        # PDF-Fehler abfangen damit Export-Fehler den Lauf nicht abbricht
        print(f"{RED}[!] PDF-Fehler ({pdf_name}): {e}{RESET}")


# ---------------------------------------------------------------------------
# Terminal-Vorschau
# ---------------------------------------------------------------------------

# (df-Spalte, Spaltentitel, max. Zeichen)
VORSCHAU_SPALTEN = [
    ("Name",    "Name",    24),
    ("Adresse", "Adresse", 22),
    ("Telefon", "Tel",     14),
    ("Website", "Web",     30),
]


# Rahmenzeichen 1:1 auf ASCII abbilden, falls die Konsole kein UTF-8 kann
# (z.B. cp1252 bei umgeleiteter Ausgabe). Gleiche Zeichenzahl -> Spalten bleiben buendig.
_ASCII_ERSATZ = str.maketrans({
    "┌": "+", "┐": "+", "└": "+", "┘": "+", "├": "+", "┤": "+",
    "┬": "+", "┴": "+", "┼": "+", "─": "-", "│": "|",
    "…": ".", "–": "-",
})


def _konsole_kann_unicode():
    enc = getattr(sys.stdout, "encoding", None) or "ascii"
    try:
        "┌┬┐├┼┤└┴┘│─…–".encode(enc)
        return True
    except (UnicodeEncodeError, LookupError):
        return False


def _drucke(text):
    """print() mit Rueckfall auf ASCII-Rahmen bei nicht-UTF-8-Konsolen."""
    print(text if _konsole_kann_unicode() else text.translate(_ASCII_ERSATZ))


def _kuerzen(text, laenge):
    """Kuerzt auf `laenge` Zeichen; gekuerzte Werte enden auf '…'."""
    text = str(text if text is not None else "")
    return text if len(text) <= laenge else text[:laenge - 1] + "…"


def _domain(url):
    """'https://www.beispiel.de/pfad' -> 'beispiel.de/pfad'."""
    text = str(url or "").strip()
    if not text or text == "KEINE WEBSITE":
        return "–"
    for prefix in ("https://", "http://"):
        if text.startswith(prefix):
            text = text[len(prefix):]
    if text.startswith("www."):
        text = text[4:]
    return text.rstrip("/") or "–"


def vorschau_ausgeben(df, branchen_label, stadt, datum_de, max_zeilen=5):
    """Gibt die ersten `max_zeilen` Leads als ANSI-Box im Terminal aus."""
    zellbreiten = [limit + 4 for _, _, limit in VORSCHAU_SPALTEN]  # 2 Zeichen Rand je Seite
    innen       = sum(zellbreiten) + len(zellbreiten) - 1          # Titelzeile ohne Spaltentrenner

    def baue_zeile(werte):
        felder = [
            "  " + _kuerzen(wert, limit).ljust(breite - 2)
            for (_, _, limit), breite, wert in zip(VORSCHAU_SPALTEN, zellbreiten, werte)
        ]
        return "│" + "│".join(felder) + "│"

    oben        = "┌" + "─" * innen + "┐"
    nach_titel  = "├" + "┬".join("─" * b for b in zellbreiten) + "┤"
    nach_kopf   = "├" + "┼".join("─" * b for b in zellbreiten) + "┤"
    unten       = "└" + "┴".join("─" * b for b in zellbreiten) + "┘"

    gezeigt = min(max_zeilen, len(df))
    titel   = f"  TOP {gezeigt} VORSCHAU – {branchen_label} | {stadt} | {datum_de}"

    print()
    _drucke(f"{CYAN}{oben}{RESET}")
    _drucke(f"{CYAN}│{RESET}{BOLD}{WHITE}{_kuerzen(titel, innen).ljust(innen)}{RESET}{CYAN}│{RESET}")
    _drucke(f"{CYAN}{nach_titel}{RESET}")
    _drucke(f"{GREEN}{BOLD}{baue_zeile([titel_kurz for _, titel_kurz, _ in VORSCHAU_SPALTEN])}{RESET}")
    _drucke(f"{CYAN}{nach_kopf}{RESET}")

    for _, row in df.head(max_zeilen).iterrows():
        werte = []
        for spalte, _, _ in VORSCHAU_SPALTEN:
            wert = row.get(spalte, "")
            werte.append(_domain(wert) if spalte == "Website" else wert)
        _drucke(baue_zeile(werte))

    _drucke(f"{CYAN}{unten}{RESET}")

    rest = len(df) - gezeigt
    if rest > 0:
        wort = "Lead" if rest == 1 else "Leads"
        _drucke(f"{CYAN}  Weitere {rest} {wort} im Export.{RESET}")


# ---------------------------------------------------------------------------
# Kategorie-Lookup und Pro-Typ-PDF-Export
# ---------------------------------------------------------------------------

def get_category_for_label(label, keywords_db):
    """Gibt die Top-Level-Kategorie aus keywords.json für einen Unternehmenstyp zurück.
    Gibt 'Sonstige' zurück wenn der Begriff nicht in der Datenbank gefunden wird (freie Eingabe).
    """
    for cat, types in keywords_db.items():
        if label in types:
            return cat
    return "Sonstige"


def exportieren_pdf_pro_typ(df_label, label, location_info, output_dir):
    """Erstellt eine PDF für einen einzelnen Unternehmenstyp im angegebenen Ordner."""
    stadt = safe_name(location_info["locality"])
    datum = pd.Timestamp.now().strftime("%Y-%m-%d")

    df_pdf   = df_label.drop_duplicates(subset=["Name", "Adresse"], keep="first")
    df_pdf   = df_pdf.sort_values("PLZ", ascending=True)
    pdf_name = f"{safe_name(label)}_{stadt}.pdf"
    pdf_path = os.path.join(output_dir, pdf_name)

    print(f"[*] Speichere nach: {pdf_path}")
    try:
        os.makedirs(os.path.dirname(pdf_path), exist_ok=True)
        _make_pdf(pdf_path, label, location_info["locality"], datum, df_pdf)
        print(f"[✓] Gespeichert: {pdf_path}")
    except Exception as e:
        print(f"[!] Fehler beim Speichern: {pdf_path} – {e}")


# ---------------------------------------------------------------------------
# Suchstrategie: Keyword-Textsuche + Deduplication
# ---------------------------------------------------------------------------

def search_one_term(gmaps, term, keywords, plz, city, max_results, fortschritt,
                    global_seen_ids, filter_params, zusammenfassung, known_db=None):
    """Sucht `term` in `plz` per Keyword-Textsuche. Dedupliziert global."""
    raw = []
    for kw in keywords:
        raw.extend(suche_leads(gmaps, kw, plz, city, max_results, fortschritt, known_db=known_db))

    for lead in raw:
        lead["Branche"] = term

    gefunden = len(raw)

    if raw:
        df_tmp = pd.DataFrame(raw)
        df_filtered = filter_anwenden(df_tmp, *filter_params)
        filtered_records = df_filtered.to_dict("records")
    else:
        filtered_records = []

    # Globale Deduplizierung: jede Place_ID darf nur einmal in der gesamten Session vorkommen
    new_leads = []
    for lead in filtered_records:
        pid = lead.get("Place_ID", "")
        if pid and pid not in global_seen_ids:
            global_seen_ids.add(pid)
            new_leads.append(lead)

    zusammenfassung.append({
        "Branche":     term,
        "PLZ":         plz,
        "Gefunden":    gefunden,
        "Nach_Filter": len(new_leads),
    })

    return new_leads


# ---------------------------------------------------------------------------
# Hauptprogramm
# ---------------------------------------------------------------------------

def run_search(gmaps, keywords_db):
    print(f"\n{CYAN}{BOLD}╔══════════════════════════════════════╗{RESET}")
    print(f"{CYAN}{BOLD}║         SUCHPARAMETER                ║{RESET}")
    print(f"{CYAN}{BOLD}╚══════════════════════════════════════╝{RESET}")

    _non_de_zaehler[0]  = 0   # Deutschland-Filter-Zähler für diesen Lauf zurücksetzen
    branchen_selected   = []  # Liste von (label, db_entry_or_none)
    ort                 = "Berlin"
    city_display        = "Berlin"  # Anzeigename (kann "Stadt1 + Stadt2" sein)
    plz_liste           = None
    include_known       = True
    search_history      = {}   # befüllt in State 8
    plz_per_branche     = {}   # befüllt in State 8; label -> gefilterte PLZ-Liste
    max_results         = 60
    nur_ohne_website    = False
    nur_mit_website     = False
    min_bewertung       = None
    max_bewertung       = None
    min_bewertungen     = None
    ausschluss_keywords = []
    ausschluss_typen    = []
    format_wahl         = "1"
    _via_saved_area     = False  # True wenn PLZ aus gespeichertem Gebiet, False nach PLZ-Auswahl

    # Zustandsmaschine für Zurück-Navigation durch die Suchparameter-Schritte.
    # 1 = Kategorie+Unternehmenstyp, 2 = Weitere hinzufügen, 3 = Stadt,
    # 4 = PLZ-Auswahl, 5 = Max-Ergebnisse, 6 = Filter, 7 = Export-Format,
    # 8 = Suchverlauf-Prüfung, 9 = Fertig
    state = 1
    while state < 9:
        if state == 1:
            entry = select_business_type(keywords_db)
            if entry == BACK_SENTINEL:
                print(f"{YELLOW}[!] Du bist bereits am Anfang.{RESET}")
                continue
            branchen_selected.append(entry)
            print(f"\n{GREEN}[✓] Hinzugefügt: {entry[0]}{RESET}")
            state = 2

        elif state == 2:
            print()
            more = input_with_back("Weiteren Unternehmenstyp hinzufügen? (j/n):")
            if more == BACK_SENTINEL:
                if branchen_selected:
                    branchen_selected.pop()
                print(f"{CYAN}[←] Zurück zum vorherigen Schritt...{RESET}")
                state = 1
                continue
            # BUGFIX: früher nur "y" akzeptiert; "j"/"ja" wurden als Nein behandelt
            if is_ja(more):
                state = 1
            else:
                state = 3

        elif state == 3:
            print()
            print(f"{'━'*42}")
            print(f" Suchgebiet wählen:")
            print(f" [1]  Gespeichertes Suchgebiet verwenden")
            print(f" [2]  Neues Suchgebiet eingeben")
            print(f"{'━'*42}")
            area_choice = input("(z = zurück)\n> ").strip()

            if is_back(area_choice):
                print(f"{CYAN}[←] Zurück zum vorherigen Schritt...{RESET}")
                state = 2
                continue

            if area_choice == "1":
                areas = load_search_areas()
                if not areas:
                    print(f"{YELLOW}[!] Keine gespeicherten Suchgebiete vorhanden.{RESET}")
                    continue
                area_result = _select_saved_area(areas)
                if area_result is None:
                    continue  # zurück zum Auswahlmenü
                ort, plz_liste, include_known, city_display = area_result
                _via_saved_area = True
                state = 5
            elif area_choice == "2":
                v = eingabe_mit_default_back("Stadt oder Region (z.B. Berlin, Potsdam)", "Berlin")
                if v == BACK_SENTINEL:
                    continue  # zurück zum Auswahlmenü
                ort = v
                state = 4
            else:
                print(f"{YELLOW}[!] Ungültige Eingabe.{RESET}")

        elif state == 4:
            print()
            combined_area_cities = {}  # city_name → [plz_list] (für kombiniertes Suchgebiet)
            current_ort          = ort
            include_known        = True
            _area_radius         = 0
            _area_city_groups    = {}

            while True:  # Schleife für mehrere Suchgebiete
                result = select_plz_for_city(gmaps, current_ort)
                if result == BACK_SENTINEL:
                    print(f"{CYAN}[←] Zurück zum vorherigen Schritt...{RESET}")
                    state = 3
                    break

                new_plz, new_include_known, _area_city_groups, new_radius = result
                if not combined_area_cities:
                    include_known = new_include_known
                    _area_radius  = new_radius

                combined_area_cities[current_ort] = sorted(set(new_plz))
                all_combined_plz = sorted(
                    set(p for plz in combined_area_cities.values() for p in plz)
                )

                if len(combined_area_cities) > 1:
                    print(f"\n{GREEN}[✓] Kombiniertes Suchgebiet:{RESET}")
                    max_len = max(len(c) for c in combined_area_cities)
                    for cname, cplz in combined_area_cities.items():
                        preview = ', '.join(cplz[:5]) + (' ...' if len(cplz) > 5 else '')
                        print(f"     {cname:<{max_len}}  {preview}  ({len(cplz)} PLZ)")
                    print(f"     Gesamt: {len(all_combined_plz)} PLZ")

                # "Weiteres Suchgebiet?" Schleife
                add_more = None
                while add_more is None:
                    add_raw = input("Weiteres Suchgebiet hinzufügen? (j/n): ").strip()
                    if is_ja(add_raw):
                        v = eingabe_mit_default_back("Stadt oder Region", "")
                        if v != BACK_SENTINEL and v.strip():
                            current_ort = v.strip()
                            add_more = True
                    elif is_nein(add_raw) or is_back(add_raw):
                        add_more = False
                    else:
                        print(f"{YELLOW}[!] Ungültige Eingabe.{RESET}")

                if add_more:
                    continue  # nächste Stadt abfragen

                # Finalisieren
                plz_liste    = all_combined_plz
                city_display = " + ".join(combined_area_cities.keys())
                ort          = list(combined_area_cities.keys())[0]  # erste Stadt für Geocoding

                while True:
                    save_raw = input("Dieses Suchgebiet speichern? (j/n): ").strip()
                    if is_ja(save_raw):
                        name_raw = input("Name für dieses Suchgebiet: ").strip()
                        if name_raw:
                            save_search_area(
                                name_raw, ort, _area_radius, plz_liste,
                                _area_city_groups,
                                cities=combined_area_cities if len(combined_area_cities) > 1 else None,
                            )
                            print(f"{GREEN}[✓] Suchgebiet '{name_raw}' gespeichert{RESET}")
                        break
                    if is_nein(save_raw) or is_back(save_raw):
                        break
                    print(f"{YELLOW}[!] Ungültige Eingabe.{RESET}")

                _via_saved_area = False
                state = 5
                break

        elif state == 5:
            print()
            v = eingabe_mit_default_back("Max. Ergebnisse pro PLZ pro Branche", "60")
            if v == BACK_SENTINEL:
                print(f"{CYAN}[←] Zurück zum vorherigen Schritt...{RESET}")
                state = 3 if _via_saved_area else 4
                continue
            try:
                max_results = int(v)
            except (TypeError, ValueError):
                max_results = 60
            state = 6

        elif state == 6:
            print(f"\n{BOLD}{WHITE}--- Filter (Enter = überspringen) ---{RESET}\n")
            back = False

            if not back:
                wf = input_with_back("Website-Filter? (1 = ohne Website, 2 = mit Website, Enter = kein Filter):")
                if wf == BACK_SENTINEL:
                    back = True
                else:
                    nur_ohne_website = (wf == "1")
                    nur_mit_website  = (wf == "2")

            if not back:
                mb = input_with_back("Mindest-Bewertung (z.B. 4.0, Enter = kein Filter):")
                if mb == BACK_SENTINEL:
                    back = True
                else:
                    try:
                        min_bewertung = float(mb) if mb else None
                    except ValueError:
                        print(f"  {YELLOW}[!] Ungültige Eingabe – Mindest-Bewertung wird ignoriert.{RESET}")
                        min_bewertung = None

            if not back:
                mxb = input_with_back("Maximal-Bewertung (z.B. 3.0, Enter = kein Filter):")
                if mxb == BACK_SENTINEL:
                    back = True
                else:
                    try:
                        max_bewertung = float(mxb) if mxb else None
                    except ValueError:
                        print(f"  {YELLOW}[!] Ungültige Eingabe – Maximal-Bewertung wird ignoriert.{RESET}")
                        max_bewertung = None

            if not back:
                ma = input_with_back("Mindestanzahl Bewertungen (z.B. 50, Enter = kein Filter):")
                if ma == BACK_SENTINEL:
                    back = True
                else:
                    try:
                        min_bewertungen = int(ma) if ma else None
                    except ValueError:
                        print(f"  {YELLOW}[!] Ungültige Eingabe – Mindestanzahl Bewertungen wird ignoriert.{RESET}")
                        min_bewertungen = None

            if not back:
                akw = input_with_back("Ausschluss-Keywords im Namen (z.B. Kette,GmbH,AG, Enter = kein Filter):")
                if akw == BACK_SENTINEL:
                    back = True
                else:
                    ausschluss_keywords = (
                        [k.strip() for k in akw.split(",") if k.strip()]
                        if akw else []
                    )

            if not back:
                atyp = input_with_back("Ausschluss-Unternehmenstypen (z.B. hotel,hostel, Enter = kein Filter):")
                if atyp == BACK_SENTINEL:
                    back = True
                else:
                    ausschluss_typen = (
                        [t.strip() for t in atyp.split(",") if t.strip()]
                        if atyp else []
                    )

            if back:
                print(f"{CYAN}[←] Zurück zum vorherigen Schritt...{RESET}")
                state = 5
                continue
            state = 7

        elif state == 7:
            print(f"\n{BOLD}{WHITE}--- Export-Format ---{RESET}\n")
            print("  1 = Excel (.xlsx)  ← empfohlen (ein Sheet pro PLZ + ALL)")
            print("  2 = CSV   (.csv)")
            v = eingabe_mit_default_back("Auswahl", "1")
            if v == BACK_SENTINEL:
                print(f"{CYAN}[←] Zurück zum vorherigen Schritt...{RESET}")
                state = 6
                continue
            format_wahl = v
            state = 8

        elif state == 8:
            # Suchverlauf pro Branche prüfen; PLZ-Listen ggf. filtern
            search_history  = load_search_history()
            plz_per_branche = {}
            back_triggered  = False
            for label, _ in branchen_selected:
                filtered = check_history_for_term(label, plz_liste, ort, search_history)
                if filtered == BACK_SENTINEL:
                    back_triggered = True
                    break
                plz_per_branche[label] = filtered
            if back_triggered:
                print(f"{CYAN}[←] Zurück zum vorherigen Schritt...{RESET}")
                state = 3 if _via_saved_area else 4
                continue
            state = 9

    branchen = [label for label, _ in branchen_selected]

    # Standort-Informationen aus der Geocoding-API ermitteln – Step 2
    # Land, Bundesland und Stadt werden für den Export-Pfad benötigt
    _loc           = get_location_info(gmaps, ort)
    export_country = _loc.get("country",  "Germany")
    export_state   = _loc.get("state",    "Unknown")
    export_city    = city_display  # kombinierter Anzeigename, z.B. "Spandau + Falkensee"

    # BUGFIX: sys.exit(1) beendet das gesamte Programm; return ermöglicht Neustart über main()-Schleife
    if not plz_liste:
        print(f"\n{RED}[!] Keine Postleitzahlen gefunden.{RESET}")
        return

    filter_params = (
        nur_ohne_website, nur_mit_website,
        min_bewertung, max_bewertung, min_bewertungen,
        ausschluss_keywords, ausschluss_typen,
    )

    # Suche: Keyword-Textsuche über alle Branchen und PLZs
    known_db        = load_known_businesses()
    alle_leads      = []
    zusammenfassung = []
    global_seen_ids = set()  # verfolgt Place_IDs über alle PLZs, Branchen und verwandte Erweiterungen

    print(f"\n{CYAN}{'='*70}{RESET}")

    for (label, db_entry) in branchen_selected:
        keywords       = db_entry.get("keywords", [label]) if db_entry else [label]  # BUGFIX: [] statt .get() warf KeyError bei fehlendem Schlüssel
        related        = db_entry.get("related",  [])     if db_entry else []       # BUGFIX: [] statt .get() warf KeyError bei fehlendem Schlüssel
        plz_fuer_label = plz_per_branche.get(label, plz_liste)

        if not plz_fuer_label:
            print(f"\n{YELLOW}[!] Alle PLZ für '{label}' übersprungen (Suchverlauf).{RESET}")
            continue

        print(f"\n{CYAN}[*] Unternehmenstyp: {BOLD}{label}{RESET}{CYAN} → Stichwörter: {', '.join(keywords)}{RESET}")

        # Suche über die (ggf. vom Verlauf gefilterten) PLZs für diesen Unternehmenstyp
        label_leads = []
        for i, plz in enumerate(plz_fuer_label):
            fortschritt = f"({i + 1}/{len(plz_fuer_label)})"
            new = search_one_term(
                gmaps, label, keywords, plz, ort, max_results, fortschritt,
                global_seen_ids, filter_params, zusammenfassung, known_db=known_db,
            )
            label_leads.extend(new)

        alle_leads.extend(label_leads)

        # Verwandte Branchen-Fallback: einmalig fragen wenn Ergebnisse spärlich sind
        if len(label_leads) <= 7 and related:
            rel_display = ", ".join(related[:2])
            print(f"\n{YELLOW}[!] Nur {len(label_leads)} Ergebnisse für '{label}' gefunden.{RESET}")
            print(f"{YELLOW}     Auch verwandte Typen durchsuchen: {rel_display}? (j/n): {RESET}", end="")
            # BUGFIX: früher nur "y" akzeptiert; "j"/"ja" wurden als Nein behandelt
            if is_ja(input().strip()):
                for rel_type in related[:2]:
                    rel_entry = None
                    for cat, types in keywords_db.items():
                        if rel_type in types:
                            rel_entry = types[rel_type]
                            break
                    rel_keywords = rel_entry.get("keywords", [rel_type]) if rel_entry else [rel_type]  # BUGFIX: [] statt .get() warf KeyError bei fehlendem Schlüssel

                    print(f"\n{CYAN}[*] Suche verwandten Typ: {BOLD}{rel_type}{RESET}")
                    for i, plz in enumerate(plz_fuer_label):
                        fortschritt = f"({i + 1}/{len(plz_fuer_label)})"
                        new = search_one_term(
                            gmaps, rel_type, rel_keywords, plz, ort, max_results, fortschritt,
                            global_seen_ids, filter_params, zusammenfassung, known_db=known_db,
                        )
                        alle_leads.extend(new)

    # Suchverlauf speichern
    save_counts = save_search_history(search_history, zusammenfassung)
    for term, count in save_counts.items():
        print(f"{GREEN}[✓] Suchverlauf aktualisiert: {count} PLZ für '{term}' gespeichert{RESET}")

    # Alle gefundenen Unternehmen in Known_businesses.json speichern (vor jeglichem Filtern)
    new_count = save_known_businesses(alle_leads, known_db)
    total_db  = len(known_db)
    print(f"{GREEN}[✓] {new_count} neue Unternehmen zur Datenbank hinzugefügt (Gesamt: {total_db}){RESET}")

    # Zusammenfassung ausgeben
    print(f"\n{CYAN}{'='*70}{RESET}")
    print(f"{CYAN}{BOLD}  ZUSAMMENFASSUNG{RESET}")
    print(f"{CYAN}{'='*70}{RESET}")
    print(f"  {BOLD}{WHITE}{'Branche':<25} {'PLZ':<7} {'Gefunden':>9} {'Nach Filter':>12}{RESET}")
    print(f"  {'-'*25} {'-'*7} {'-'*9} {'-'*12}")
    for z in zusammenfassung:
        row_color = GREEN if z["Nach_Filter"] > 0 else YELLOW
        print(f"{row_color}  {z['Branche']:<25} {z['PLZ']:<7} {z['Gefunden']:>9} {z['Nach_Filter']:>12}{RESET}")
    print(f"{CYAN}{'='*70}{RESET}")

    # BUGFIX: sys.exit(0) beendet das gesamte Programm; return ermöglicht Neustart über main()-Schleife
    if not alle_leads:
        print(f"\n{YELLOW}[!] Keine Ergebnisse gefunden.{RESET}")
        return

    df = pd.DataFrame(alle_leads)

    # Bekannt-Spalte immer hinzufügen
    df["Bekannt"] = df["Place_ID"].apply(
        lambda pid: "Ja" if pid in known_db else "Nein"
    )

    # Wenn include_known=False: nicht filtern, sondern nur markieren und informieren
    if not include_known:
        bekannte_anzahl = (df["Bekannt"] == "Ja").sum()
        if bekannte_anzahl:
            print(f"{CYAN}[*] {bekannte_anzahl} bekannte Unternehmen in den Ergebnissen (markiert){RESET}")

    # NEVER return here - always continue to export
    gesamt       = len(df)
    ohne_website = (df["Website"] == "KEINE WEBSITE").sum()

    if gesamt == 0:
        print(f"{YELLOW}[!] Keine Ergebnisse gefunden – Export übersprungen.{RESET}")
        return

    print(f"  Kombiniertes Ergebnis: {gesamt} Leads (davon {ohne_website} ohne Website)")
    print(f"{CYAN}{'='*70}{RESET}")

    # Vorschau Top 5
    vorschau_ausgeben(
        df,
        " / ".join(lbl for lbl, _ in branchen_selected),
        export_city,
        datetime.now().strftime("%d.%m.%Y"),
    )

    print(f"[*] Export wird gestartet...")

    # Pro Branche eine eigene PDF + Excel erstellen
    primary_labels = {lbl for lbl, _ in branchen_selected}
    for (label, db_entry) in branchen_selected:
        related_labels      = db_entry.get("related", [])[:2] if db_entry else []
        vorhandene_branchen = set(df["Branche"].unique())
        include_labels      = [label] + [
            r for r in related_labels
            if r in vorhandene_branchen and r not in primary_labels
        ]
        def branche_matches(branche_val, labels):
            branche_lower = str(branche_val).lower()
            return any(lbl.lower() in branche_lower or branche_lower in lbl.lower()
                       for lbl in labels)

        df_label = df[df["Branche"].apply(lambda x: branche_matches(x, include_labels))].copy()
        if df_label.empty:
            print(f"{YELLOW}[!] Keine Leads für '{label}' – übersprungen.{RESET}")
            if len(branchen_selected) == 1:
                df_label = df.copy()
                print(f"[*] Fallback: Exportiere alle {len(df_label)} Leads.")
            else:
                continue

        kategorie = get_category_for_label(label, keywords_db)

        try:
            export_leads(
                df=df_label,
                business_type=label,
                category=kategorie,
                city=export_city,
                state=export_state,
                country=export_country,
                base_dir=BASE_LEADS_DIR,
                filter_info={
                    "nur_ohne_website": nur_ohne_website,
                    "min_bewertung":    min_bewertung,
                },
            )
        except Exception as e:
            import traceback
            print(f"{RED}[!] Export-Fehler für '{label}': {e}{RESET}")
            traceback.print_exc()

    # Deutschland-Filter Zusammenfassung
    if _non_de_zaehler[0]:
        print(f"\n{CYAN}[*] {_non_de_zaehler[0]} Einträge außerhalb Deutschlands wurden entfernt{RESET}")

    print(f"\n{GREEN}{BOLD}Fertig! Viel Erfolg mit deinen Leads.{RESET}\n")


def main():
    global BASE_LEADS_DIR

    load_dotenv_file()
    api_key = os.environ.get("GOOGLE_MAPS_API_KEY", "").strip()

    if api_key:
        print(f"{GREEN}[✓] Google Maps API-Key aus GOOGLE_MAPS_API_KEY geladen{RESET}")
    else:
        api_key = get_api_key()

    # Speicherort beim Start abfragen
    default_leads = os.path.join(BASE_DIR, "leads")
    print(f"\n{CYAN}{'━'*60}{RESET}")
    print(f"{BOLD}  SPEICHERORT FÜR ERGEBNISSE{RESET}")
    print(f"{CYAN}{'━'*60}{RESET}")
    while True:
        print(f"\n  Aktueller Pfad: {BOLD}{default_leads}{RESET}")
        antwort = input("  Diesen Pfad verwenden? (j = ja / anderen Pfad eingeben): ").strip()
        if antwort.lower() in ("j", "ja", ""):
            gewaehlter_pfad = default_leads
        else:
            gewaehlter_pfad = antwort
        try:
            os.makedirs(gewaehlter_pfad, exist_ok=True)
            BASE_LEADS_DIR = gewaehlter_pfad
            print(f"{GREEN}[✓] Speicherort gesetzt: {BASE_LEADS_DIR}{RESET}")
            break
        except Exception as e:
            print(f"{RED}[!] Pfad ungültig: {e} – Bitte erneut eingeben.{RESET}")
            default_leads = gewaehlter_pfad

    print(f"\n{CYAN}{BOLD}╔══════════════════════════════════════╗{RESET}")
    print(f"{CYAN}{BOLD}║     GOOGLE MAPS LEAD GENERATOR       ║{RESET}")
    print(f"{CYAN}{BOLD}╚══════════════════════════════════════╝{RESET}")

    gmaps = googlemaps.Client(key=api_key)

    keywords_db = load_keywords_database()

    while True:
        run_search(gmaps, keywords_db)

        print(f"\n{'━'*43}")
        print(" Was möchtest du als nächstes tun?")
        print(" [1]  Neue Suche starten")
        print(" [2]  Programm beenden")
        print(f"{'━'*43}")
        choice = input("> ").strip()
        if choice == "1":
            continue
        else:
            print("Auf Wiedersehen!")
            break


if __name__ == "__main__":
    main()