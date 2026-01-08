from datetime import datetime, timedelta
from urllib.parse import urlparse
from zoneinfo import ZoneInfo

from icalendar import Calendar, Event


def get_uid_domain(base_url):
    if not base_url:
        return "aufsichtshelper.local"
    parsed = urlparse(base_url)
    if parsed.hostname:
        return parsed.hostname
    return base_url


def build_event_uid(row_fingerprint, role, uid_domain=None):
    domain = uid_domain or "aufsichtshelper.local"
    return f"{row_fingerprint}-{role}@{domain}"


def build_summary(event_data, role):
    pruefungsname = event_data.get("pruefungsname", "")
    if role == "abloesung":
        return f"Ablösung: {pruefungsname}"
    return f"Aufsicht: {pruefungsname}"


def build_description(event_data):
    lines = [
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


def build_ics_event(event_data, role, uid_domain=None, event_uid=None):
    tz = ZoneInfo("Europe/Berlin")
    dtstart = datetime.combine(event_data["datum"], event_data["startzeit"], tzinfo=tz)
    dtend = dtstart + timedelta(minutes=event_data["dauer_minuten"])

    event_uid = event_uid or build_event_uid(
        event_data["row_fingerprint"], role, uid_domain
    )

    cal = Calendar()
    cal.add("prodid", "-//AufsichtsHelper//DE")
    cal.add("version", "2.0")
    cal.add("method", "REQUEST")

    event = Event()
    event.add("uid", event_uid)
    event.add("summary", build_summary(event_data, role))
    event.add("dtstart", dtstart)
    event.add("dtend", dtend)
    event.add("dtstamp", datetime.now(tz))
    event.add("location", event_data.get("raum", ""))
    event.add("description", build_description(event_data))

    cal.add_component(event)
    return cal.to_ical(), event_uid

