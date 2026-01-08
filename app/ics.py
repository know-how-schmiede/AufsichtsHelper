from datetime import datetime, timedelta, timezone
from urllib.parse import urlparse
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

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


def _get_timezone():
    try:
        return ZoneInfo("Europe/Berlin")
    except ZoneInfoNotFoundError:
        try:
            import pytz

            return pytz.timezone("Europe/Berlin")
        except Exception:
            return None


def build_summary(event_data, role):
    pruefungsname = event_data.get("pruefungsname", "")
    raum = event_data.get("raum", "")
    if role == "abloesung":
        return f"Abl\u00f6sung {pruefungsname} {raum}".strip()
    return f"Aufsicht {pruefungsname} {raum}".strip()


def build_description(event_data):
    lines = [
        f"Pr\u00fcfungsname: {event_data.get('pruefungsname', '')}",
        f"Datum: {event_data.get('datum').isoformat()}",
        f"Startzeit: {event_data.get('startzeit').strftime('%H:%M')}",
        f"Dauer (Minuten): {event_data.get('dauer_minuten')}",
        f"Raum: {event_data.get('raum', '')}",
        f"Pr\u00fcfer: {event_data.get('pruefer', '')}",
        f"Aufsicht: {event_data.get('aufsicht', '')}",
        f"Abl\u00f6sung: {event_data.get('abloesung', '')}",
    ]
    return "\n".join(lines)


def build_ics_event(event_data, role, uid_domain=None, event_uid=None, calendar_name=None):
    tz = _get_timezone()
    if tz:
        dtstart = datetime.combine(event_data["datum"], event_data["startzeit"], tzinfo=tz)
        dtend = dtstart + timedelta(minutes=event_data["dauer_minuten"])
        dtstamp = datetime.now(tz)
    else:
        dtstart = datetime.combine(event_data["datum"], event_data["startzeit"])
        dtend = dtstart + timedelta(minutes=event_data["dauer_minuten"])
        dtstamp = datetime.now(timezone.utc)

    event_uid = event_uid or build_event_uid(
        event_data["row_fingerprint"], role, uid_domain
    )

    cal = Calendar()
    cal.add("prodid", "-//AufsichtsHelper//DE")
    cal.add("version", "2.0")
    cal.add("method", "REQUEST")
    cal.add("X-WR-TIMEZONE", "Europe/Berlin")
    if calendar_name:
        cal.add("X-WR-CALNAME", calendar_name)
        cal.add("NAME", calendar_name)

    event = Event()
    event.add("uid", event_uid)
    event.add("summary", build_summary(event_data, role))
    event.add("dtstart", dtstart)
    event.add("dtend", dtend)
    event.add("dtstamp", dtstamp)
    event.add("location", event_data.get("raum", ""))
    event.add("description", build_description(event_data))

    cal.add_component(event)
    return cal.to_ical(), event_uid
