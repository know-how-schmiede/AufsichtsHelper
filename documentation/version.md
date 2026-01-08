# Version

Die aktuelle App-Version wird in `app/VERSION` gepflegt.

## Update-Workflow
1) Versionsnummer in `app/VERSION` aktualisieren.
2) Beschreibung der Aenderung in diesem Dokument ergaenzen.

## Historie

### Version 0.1.30

- Vorschau: Anzahl der Aufsichts-Namen im Label angezeigt

### Version 0.1.29

- Hilfe- und Datenschutzseiten hinzugefuegt

### Version 0.1.28

- Navigation: Personen/Aliases aus der Navbar entfernt

### Version 0.1.27

- Vorschau: Sortierpfeile im PrintFleet2-Stil

### Version 0.1.26

- Vorschau: Sortierbuttons als Pfeile statt Text

### Version 0.1.25

- Vorschau: Sortierbuttons und Trefferhinweis fuer gefilterte Liste

### Version 0.1.24

- Vorschau: Mehrzeilige Liste mit echten <br>-Zeilenumbruechen

### Version 0.1.23

- Vorschau: Pruefer/Aufsicht/Raum als mehrzeilige Liste

### Version 0.1.22

- Vorschau: ICS-Paket Button erst nach Filterung aktiv

### Version 0.1.21

- Setup: Git-Clone ohne Login (GIT_TERMINAL_PROMPT=0) in Debian-Skripten dokumentiert

### Version 0.1.20

- Setup: Debian 13 LXC Installationsskripte und Anleitung im setup/ Ordner

### Version 0.1.19

- Upload: Kalendername setzen (Default: Prüfungsaufsicht_<Jahr>)

### Version 0.1.18

- ICS-Export: nur fuer die ausgewaehlte Aufsicht erzeugen

### Version 0.1.17

- Vorschau: Ablösung ausgeblendet, Duplikate werden uebersprungen

### Version 0.1.16

- ICS-Export: Ablösungen werden nicht mehr erzeugt

### Version 0.1.15

- ICS: Dateiname und Summary an neues Schema angepasst

### Version 0.1.14

- ICS: Fallback auf pytz bei fehlender Europe/Berlin Zone

### Version 0.1.13

- Vorschau: Datum/Startzeit/Dauer werden aus Excel-Serials formatiert

### Version 0.1.12

- Download funktioniert auch ohne Personen-Stammdaten

### Version 0.1.11

- Fix: Listen-Join in Excel-Parser korrigiert

### Version 0.1.10

- Fix: split_names Syntaxfehler behoben

### Version 0.1.9

- Fix: Syntaxfehler im Excel-Parser behoben

### Version 0.1.8

- Mehrfachnamen in Aufsicht/Ablösung werden aufgeteilt

### Version 0.1.7

- Excel-Import: Kopfzeile wird automatisch in den ersten Zeilen erkannt

### Version 0.1.6

- Excel-Import: Fallback-Parser fuer problematische XLSX-Dateien

### Version 0.1.5

- Excel-Import: bessere Fehlerbehandlung bei defekten XLSX-Dateien

### Version 0.1.4

- Excel-Alias-Mapping fuer abweichende Spaltenueberschriften

### Version 0.1.3

- Setup: Hinweis zum migrations/ Ordner

### Version 0.1.2

- Excel-Import ohne pandas (kompatibel mit Python 3.13)

### Version 0.1.1

- iCal als ZIP-Download statt Mail-Versand

### Version 0.1.0

- Initiales Projektgeruest.
