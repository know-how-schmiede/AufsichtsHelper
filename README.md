# AufsichtsHelper

Flask-WebApp zum Upload einer Excel-Datei und Versand von iCal-Einladungen an Aufsichten und Ablosungen.

## Features
- Upload von .xlsx-Dateien mit Pruefungsdaten
- Auswahl einer Aufsicht aus den gefundenen Namen
- Filter auf die Zeilen, in denen die Aufsicht in der Spalte "Aufsicht" steht
- Versand von iCal/ICS-Terminen per SMTP an Aufsicht und ggf. Abloesung
- Personen-Stammdaten (Name, E-Mail, aktiv) + optionale Alias-Namen
- Versand-Log mit Schutz vor Doppelsendungen

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

2) Konfiguration:
- Kopiere `.env.example` nach `.env` und passe die SMTP-Werte an.

3) Datenbank initialisieren (im aktivierten venv):
```bash
python -m flask --app run.py db init  # nur falls migrations/ noch nicht existiert
python -m flask --app run.py db migrate -m "init"
python -m flask --app run.py db upgrade
```

4) App starten (im aktivierten venv):
```bash
python -m flask --app run.py run
```

## Excel-Format
Erwartete Spalten (exakt so benannt):
- "Prüfungsname"
- "Datum"
- "Startzeit"
- "Dauer"
- "Prüfer"
- "Aufsicht"
- "Ablösung"
- "Raum"

Hinweise:
- Datum: Excel-Date, ISO-String oder deutsches Datum (dd.mm.yyyy) werden geparst.
- Startzeit: Excel-Time oder String (HH:MM / HH:MM:SS) wird geparst.
- Dauer: bevorzugt Minuten (int), alternativ "HH:MM".
- Fuer mehrere Namen in einer Zelle bitte ";" oder "/" als Trennzeichen verwenden.

## SMTP-Konfiguration
Erforderliche Variablen in `.env`:
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
  migrations/
  instance/
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
