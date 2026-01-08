# Installation (Debian 13 LXC)

Diese Anleitung richtet AufsichtsHelper in einem neuen Debian 13 LXC-Container ein.
Alle Schritte werden als root ausgefuehrt.

## 1) Repo holen

```bash
apt update -y
apt upgrade -y
apt install -y git
git clone https://github.com/know-how-schmiede/AufsichtsHelper.git /opt/aufsichtshelper
```

## 2) Scripts ausfuehrbar machen

```bash
cd /opt/aufsichtshelper/setup
chmod +x setupAufsichtsHelper setupAufsichtsHelperService updateAufsichtsHelperService
```

## 3) Grundinstallation

```bash
./setupAufsichtsHelper
```

Das Script fragt nach ein paar Parametern und nutzt Standardwerte, wenn nichts eingegeben wird.
Bei root setzt es automatisch `safe.directory` fuer das Repo.

Falls GitHub nach einem Passwort fragt, ist etwas mit der URL oder dem Terminal-Input schiefgelaufen.
Nimm die .git-URL und nutze optional:

```bash
GIT_TERMINAL_PROMPT=0 git clone https://github.com/know-how-schmiede/AufsichtsHelper.git /opt/aufsichtshelper
```

## 4) Testlauf (ohne Service)

```bash
cd /opt/aufsichtshelper
source .venv/bin/activate
flask --app run.py run --host 0.0.0.0 --port 5000
```

## 5) Service installieren

Nach einem erfolgreichen Testlauf:

```bash
cd /opt/aufsichtshelper/setup
./setupAufsichtsHelperService
```

## 6) Updates einspielen

```bash
cd /opt/aufsichtshelper/setup
./updateAufsichtsHelperService
```

## Hinweise

- Die Konfiguration liegt in `/opt/aufsichtshelper/.env`.
- Die Datenbank liegt standardmaessig in `/opt/aufsichtshelper/instance/app.db`.
