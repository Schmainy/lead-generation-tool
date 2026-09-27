# Lead-Generation-Tool

Ein Python-Kommandozeilentool, das über die Google Maps Places API nach lokalen Unternehmen sucht – z. B. nach Branche und Berliner Bezirk oder Stadt plus Umkreis – und die Ergebnisse als **PDF** und **Excel** exportiert. Besonders nützlich, um Firmen ohne eigene Website zu finden.

## Features

- **Keyword-Datenbank** (`keywords.json`): Branchen und Unternehmenstypen nach Kategorien (Handwerk, Medizin, Gastronomie …), inkl. verwandter Berufsgruppen zur automatischen Suchausweitung
- **Berliner Bezirke**: fest hinterlegte PLZ-Zuordnung, alternativ Stadt + Suchradius (PLZ-Ermittlung über `pgeocode`)
- **Gespeicherte Suchgebiete** (`Search_areas.json`) zur Wiederverwendung
- **Duplikat-Erkennung**: bereits gefundene Unternehmen werden anhand der Place-ID in `Known_businesses.json` erkannt und im Export markiert
- **Suchverlauf** (`Search_history.json`): bereits durchsuchte PLZ je Branche werden erkannt, um doppelte API-Kosten zu vermeiden
- **Filter**: nur mit/ohne Website, Mindest-/Maximalbewertung, Mindestanzahl Bewertungen, Ausschluss-Keywords
- **Export**: PDF (nach PLZ gruppiert) und formatiertes Excel, abgelegt unter `leads/<Land>/<Bundesland>/<Stadt>/<Kategorie>/`
- **Farbige, deutschsprachige Terminal-UI** mit Schritt-für-Schritt-Dialog und „zurück“-Navigation

## Tech-Stack

Python 3 · [googlemaps](https://github.com/googlemaps/google-maps-services-python) · requests (Places API New) · pandas · numpy · pgeocode · openpyxl · reportlab

## Installation

```bash
pip install -r requirements.txt
```

## Konfiguration

1. In der [Google Cloud Console](https://console.cloud.google.com) ein Projekt anlegen und diese APIs aktivieren:
   - **Places API (New)** – Textsuche
   - **Places API** – Place Details (Telefon, Website, Öffnungszeiten)
   - **Geocoding API** – PLZ- und Stadterkennung
2. Einen API-Key erstellen (am besten auf diese APIs beschränken).
3. `.env.example` nach `.env` kopieren und den Key eintragen:

   ```
   GOOGLE_MAPS_API_KEY=dein_api_key
   ```

   Alternativ die Umgebungsvariable `GOOGLE_MAPS_API_KEY` setzen. Ist kein Key vorhanden, fragt das Tool beim Start danach.

Die Dateien `Known_businesses.json`, `Search_history.json` und `Search_areas.json` legt das Tool bei Bedarf selbst an. Wie sie aufgebaut sind, zeigen `Known_businesses_example.json`, `Search_history_example.json` und `Search_areas_example.json`.

## Nutzung

```bash
python leads_generator.py
```

oder unter Windows per Doppelklick auf `start.bat`. Eine ausführliche Schritt-für-Schritt-Anleitung steht in [`Anleitung.txt`](Anleitung.txt).

Ein Beispiel für den Excel-Export (erfundene Daten) liegt unter [`leads/sample_data/`](leads/sample_data/).

### PDF-Layout testen

`python test_pdf.py` erzeugt eine Test-PDF mit erfundenen Daten unter `leads/test_export/` – ohne API-Aufruf und ohne Kosten.

## Hinweise

- Die Nutzung der Google Maps Platform ist kostenpflichtig (mit monatlichem Freikontingent). Aktuelle Preise: <https://mapsplatform.google.com/pricing/>
- Beachte die [Nutzungsbedingungen der Google Maps Platform](https://cloud.google.com/maps-platform/terms) und die DSGVO, wenn du exportierte Daten speicherst oder für Kontaktaufnahmen nutzt.
- `.env`, `Known_businesses.json`, `Search_history.json`, `Search_areas.json` und der Ordner `leads/` (außer `sample_data/`) sind per `.gitignore` von der Versionierung ausgeschlossen.

## Lizenz

[MIT](LICENSE)
