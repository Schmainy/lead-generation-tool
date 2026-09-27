#!/usr/bin/env python3
"""
Erzeugt eine Test-PDF mit synthetischen Daten, um Kopfzeile, Spaltenbreiten,
Zeilenumbruch und PLZ-Trennzeilen visuell pruefen zu koennen.

Ruft _make_pdf() direkt auf - kein API-Call, kein Suchlauf, keine Kosten.
Start: python test_pdf.py
"""

import os
import sys

import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

from leads_generator import _make_pdf  # noqa: E402

ZIEL_ORDNER = os.path.join(BASE_DIR, "leads", "test_export")

BRANCHE = "Optiker"
STADT   = "Spandau"
DATUM   = "2026-08-19"

# 10 Datensaetze ueber 3 PLZ-Gruppen: lange Namen, lange Adressen, lange URLs,
# fehlende Websites, hohe und niedrige Bewertungen.
DATEN = [
    # ---- 13581 ----
    {"Name": "Musteroptik & Hoerakustik Beispiel Nachfolger GmbH & Co. KG",
     "Adresse": "Musterallee 93-95, Aufgang C, 2. Obergeschoss, 13581 Berlin-Spandau",
     "Telefon": "030 0000001",
     "Website": "https://www.musteroptik-und-hoerakustik-beispiel.example/filialen/musterallee/terminvereinbarung",
     "Bewertung": 4.8, "PLZ": "13581"},
    {"Name": "Brillen Mustermann",
     "Adresse": "Musterstraße 1, 13581 Berlin",
     "Telefon": "030 0000002",
     "Website": "KEINE WEBSITE",
     "Bewertung": 2.3, "PLZ": "13581"},
    {"Name": "Optik am Musterplatz",
     "Adresse": "Musterplatz 47, 13581 Berlin",
     "Telefon": "030 0000003",
     "Website": "https://optik-am-musterplatz.example",
     "Bewertung": 4.1, "PLZ": "13581"},

    # ---- 13583 ----
    {"Name": "Sehzentrum Musterstadt-Nord Kontaktlinsen-Institut",
     "Adresse": "Beispielweg 5-17, Haus 4, 13583 Berlin",
     "Telefon": "030 0000004",
     "Website": "http://www.sehzentrum-musterstadt-nord.example/leistungen/kontaktlinsen-anpassung",
     "Bewertung": 3.9, "PLZ": "13583"},
    {"Name": "Musterfirma GmbH",
     "Adresse": "Musterring 2, 13583 Berlin",
     "Telefon": "030 0000005",
     "Website": "https://www.musterfirma.example/filialen/berlin",
     "Bewertung": 4.4, "PLZ": "13583"},
    {"Name": "Optiker Beispiel & Söhne — Meisterbetrieb seit 1923",
     "Adresse": "Beispielstraße 14, 13583 Berlin-Spandau",
     "Telefon": "030 0000006",
     "Website": "KEINE WEBSITE",
     "Bewertung": 5.0, "PLZ": "13583"},
    {"Name": "Blickpunkt Musteroptik",
     "Adresse": "Musterstr. 8, 13583 Berlin",
     "Telefon": "030 0000007",
     "Website": "https://blickpunkt-musteroptik.example",
     "Bewertung": 1.8, "PLZ": "13583"},

    # ---- 13585 ----
    {"Name": "Die Brillenmacher am Musterdamm Inh. M. Mustermann e.K.",
     "Adresse": "Musterdamm 122a, Eingang Hofseite, 13585 Berlin",
     "Telefon": "030 0000008",
     "Website": "https://www.die-brillenmacher-am-musterdamm.example/service-und-beratung/index.html",
     "Bewertung": 4.6, "PLZ": "13585"},
    {"Name": "Optik Muster-Nord",
     "Adresse": "Beispielallee 90, 13585 Berlin",
     "Telefon": "030 0000009",
     "Website": "KEINE WEBSITE",
     "Bewertung": 3.2, "PLZ": "13585"},
    {"Name": "Kontaktlinsen-Studio Beispiel",
     "Adresse": "Musterweg 38, 13585 Berlin",
     "Telefon": "030 0000010",
     "Website": "https://www.kontaktlinsen-studio-beispiel.example/oeffnungszeiten-und-anfahrt",
     "Bewertung": 4.9, "PLZ": "13585"},
]


def main():
    df = pd.DataFrame(DATEN).sort_values("PLZ").reset_index(drop=True)

    os.makedirs(ZIEL_ORDNER, exist_ok=True)
    pdf_path = os.path.join(ZIEL_ORDNER, f"{BRANCHE}_{STADT}_{DATUM}.pdf")

    print("=" * 70)
    print("  TEST-PDF MIT SYNTHETISCHEN DATEN")
    print("=" * 70)
    print(f"  Datensaetze : {len(df)}")
    print(f"  PLZ-Gruppen : {', '.join(sorted(df['PLZ'].unique()))}")
    print(f"  Ohne Website: {(df['Website'] == 'KEINE WEBSITE').sum()}")
    print(f"  Ziel        : {pdf_path}")

    _make_pdf(
        pdf_path,
        BRANCHE,
        STADT,
        DATUM,
        df,
        gesamt=len(df),
        filter_info={"nur_ohne_website": False, "min_bewertung": 3.0},
    )

    groesse = os.path.getsize(pdf_path)
    print(f"\n[✓] PDF erstellt ({groesse} Bytes)")
    print("    Erwartet in der PDF:")
    print("      - Kopfzeile 'Lead-Export – Optiker | Spandau' auf jeder Seite")
    print("      - 'Suche vom 19.08.2026 | PLZ-Gebiet: 13581, 13583, 13585'")
    print("      - 'Gesamt: 10 Leads | Filter: ohne Website = Nein, min. Bewertung = 3.0'")
    print("      - Trennlinie unter dem Header, Tabelle beginnt darunter")
    print("      - 3 graue PLZ-Trennzeilen")
    print("      - Lange URLs umgebrochen an '/' bzw. '-'")

    print("\n[*] Oeffne PDF ...")
    try:
        os.startfile(pdf_path)
    except Exception as e:
        print(f"[!] Konnte PDF nicht oeffnen: {e}")
        print(f"    Bitte manuell oeffnen: {pdf_path}")


if __name__ == "__main__":
    main()
