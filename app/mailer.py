import os
import smtplib
from email.message import EmailMessage


def parse_bool(value):
    return str(value).strip().lower() in {"1", "true", "yes", "on"}


def load_smtp_config():
    return {
        "host": os.environ.get("SMTP_HOST", ""),
        "port": int(os.environ.get("SMTP_PORT", "587")),
        "user": os.environ.get("SMTP_USER", ""),
        "password": os.environ.get("SMTP_PASS", ""),
        "use_tls": parse_bool(os.environ.get("SMTP_USE_TLS", "true")),
        "sender": os.environ.get("SMTP_FROM", ""),
    }


def build_email_body(event_data, role):
    role_label = "Ablösung" if role == "abloesung" else "Aufsicht"
    lines = [
        f"{role_label}-Termin für {event_data.get('pruefungsname', '')}",
        "",
        f"Prüfungsname: {event_data.get('pruefungsname', '')}",
        f"Datum: {event_data.get('datum').isoformat()}",
        f"Startzeit: {event_data.get('startzeit').strftime('%H:%M')}",
        f"Dauer (Minuten): {event_data.get('dauer_minuten')}",
        f"Raum: {event_data.get('raum', '')}",
        f"Prüfer: {event_data.get('pruefer', '')}",
        f"Aufsicht: {event_data.get('aufsicht', '')}",
        f"Ablösung: {event_data.get('abloesung', '')}",
    ]
    return "\n".join(lines)


def send_invite(recipient_email, subject, body, ics_bytes, ics_filename):
    config = load_smtp_config()
    if not config["host"]:
        raise RuntimeError("SMTP_HOST ist nicht gesetzt")
    if not config["sender"]:
        raise RuntimeError("SMTP_FROM ist nicht gesetzt")

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = config["sender"]
    msg["To"] = recipient_email
    msg.set_content(body)
    msg.add_attachment(
        ics_bytes,
        maintype="text",
        subtype="calendar",
        filename=ics_filename,
        params={"method": "REQUEST"},
    )

    with smtplib.SMTP(config["host"], config["port"]) as server:
        if config["use_tls"]:
            server.starttls()
        if config["user"]:
            server.login(config["user"], config["password"])
        server.send_message(msg)

