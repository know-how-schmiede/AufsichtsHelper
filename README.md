# AufsichtsHelper

Flask-WebApp zum Upload einer Excel-Datei und Erzeugen von iCal-Einladungen fuer Aufsichten und Abloesungen.

## Features
- Upload von .xlsx-Dateien mit Pruefungsdaten
- Auswahl einer Aufsicht aus den gefundenen Namen
- Filter auf die Zeilen, in denen die Aufsicht in der Spalte "Aufsicht" steht
- Erzeugung von iCal/ICS-Terminen als ZIP-Download (nur fuer die ausgewaehlte Aufsicht, ohne Duplikate; SMTP-Versand spaeter aktivierbar)
- Personen-Stammdaten (Name, E-Mail, aktiv) + optionale Alias-Namen (optional, nur fuer Mailversand)
- Optionaler Kalendername beim Upload (Standard: "Prüfungsaufsicht_<Jahr>")
- Erstell-Log mit Schutz vor doppelten Paketen

## Voraussetzungen
- Python 3.11+ (64-bit empfohlen). Die App nutzt openpyxl statt pandas und laeuft damit auch unter Python 3.13.

## Setup (immer in venv ausfuehren)
1) Virtuelle Umgebung erstellen und aktivieren:
```bash
python -m venv .venv
```

Windows (PowerShell):
```bash
.\.venv\Scripts\activate
```

macOS/Linux:
```bash
source .venv/bin/activate
```

2) Abhaengigkeiten installieren:
```bash
python -m pip install -r requirements.txt
```

3) Konfiguration:
- Kopiere `.env.example` nach `.env` und passe die Werte an.
- SMTP-Werte sind nur fuer den spaeteren Mail-Versand relevant.

4) Datenbank initialisieren (im aktivierten venv):
```bash
python -m flask --app run.py db init  # falls migrations/ noch nicht existiert
python -m flask --app run.py db migrate -m "init"
python -m flask --app run.py db upgrade
```
Hinweis: Wenn `migrations/` bereits existiert, aber keine `env.py` enthaelt (z.B. leeres Verzeichnis),
bitte den Ordner loeschen und `db init` erneut ausfuehren.

5) App starten (im aktivierten venv):
```bash
python -m flask --app run.py run
```

## Debian 13 LXC Setup
Siehe `setup/README.md` fuer die Installation per Skript.
Bei root setzt das Skript automatisch `safe.directory` fuer das Repo.

## Excel-Format
Erwartete Felder (interne Namen):
- "Prüfungsname"
- "Datum"
- "Startzeit"
- "Dauer"
- "Prüfer"
- "Aufsicht"
- "Ablösung"
- "Raum"

Unterstuetzte Spaltennamen (Alias-Mapping):
- Prüfungsname: "Prüfungsname", "Fach", "Modul LV-Nr.", "Modul", "LV-Nr.", "EDV-Nr."
- Datum: "Datum", "Prüfungstag", "Tag"
- Startzeit: "Startzeit", "Uhrzeit"
- Dauer: "Dauer"
- Prüfer: "Prüfer" (mehrere Spalten werden zusammengefuehrt)
- Aufsicht: "Aufsicht"
- Ablösung: "Ablösung", "Ablösung/ Beisitzer", "Ablösung/Beisitzer"
- Raum: "Raum", "Räume", "Räume vorgezogen" (mehrere Spalten werden zusammengefuehrt)

Hinweise:
- Datum: Excel-Date, ISO-String oder deutsches Datum (dd.mm.yyyy) werden geparst.
- Startzeit: Excel-Time oder String (HH:MM / HH:MM:SS) wird geparst.
- Dauer: bevorzugt Minuten (int), alternativ "HH:MM".
- Fuer mehrere Namen in einer Zelle bitte Zeilenumbrueche, ";" oder "/" als Trennzeichen verwenden.
- Komma-Listen im Format "Nachname, Vorname, Nachname, Vorname" werden ebenfalls erkannt.
- Bei der Fehlermeldung "defektes XML" die Datei in Excel/LibreOffice oeffnen und erneut als .xlsx speichern.
- Bei Fehlern zur Datenvalidierung (z.B. "Value must be one of ...") ebenfalls neu speichern, damit das XLSX sauber ist.
- Die Kopfzeile wird automatisch in den ersten 10 Zeilen gesucht.

## SMTP-Konfiguration
Erforderliche Variablen in `.env` (nur bei Mail-Versand):
- SMTP_HOST
- SMTP_PORT
- SMTP_USER
- SMTP_PASS
- SMTP_USE_TLS (true/false)
- SMTP_FROM
- APP_BASE_URL (optional, fuer UID-Domain in ICS)

## Projektstruktur
```
aufsichtshelper/
  app/
    VERSION
    __init__.py
    extensions.py
    models.py
    excel.py
    ics.py
    mailer.py
    routes/
      main.py
      persons.py
    templates/
    static/
  migrations/  # wird durch flask db init erzeugt
  instance/
    exports/
    uploads/
  run.py
  config.py
  requirements.txt
  .env.example
  documentation/
    version.md
```

## Versionierung
- Die Versionsnummer liegt in `app/VERSION`.
- Aenderungen werden in `documentation/version.md` dokumentiert.
