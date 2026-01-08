import hashlib
import re
from datetime import date, datetime, time, timedelta

import pandas as pd

EXPECTED_COLUMNS = [
    "Prüfungsname",
    "Datum",
    "Startzeit",
    "Dauer",
    "Prüfer",
    "Aufsicht",
    "Ablösung",
    "Raum",
]


def read_excel(file_path):
    df = pd.read_excel(file_path, engine="openpyxl", dtype=object)
    df = df.where(pd.notnull(df), None)
    missing = [col for col in EXPECTED_COLUMNS if col not in df.columns]
    if missing:
        raise ValueError(f"Missing columns: {', '.join(missing)}")
    return df


def normalize_name(value):
    if value is None:
        return ""
    return " ".join(str(value).strip().split()).casefold()


def name_candidates(value):
    if value is None:
        return []
    text = str(value).strip()
    if not text:
        return []

    if ";" in text or "/" in text:
        parts = re.split(r"[;/]", text)
    else:
        parts = [text]

    names = []
    for part in parts:
        norm = normalize_name(part)
        if norm:
            names.append(norm)

    full_norm = normalize_name(text)
    if full_norm and full_norm not in names:
        names.append(full_norm)

    return names


def matches_name(cell_value, selected_name):
    if not selected_name:
        return False
    target = normalize_name(selected_name)
    return target in name_candidates(cell_value)


def extract_aufsichten(df):
    seen = {}
    for value in df["Aufsicht"]:
        if value is None:
            continue
        text = str(value).strip()
        if not text:
            continue
        norm = normalize_name(text)
        if norm and norm not in seen:
            seen[norm] = text
    return [seen[key] for key in sorted(seen.keys())]


def filter_rows_by_aufsicht(df, selected_name):
    rows = df.to_dict(orient="records")
    return [row for row in rows if matches_name(row.get("Aufsicht"), selected_name)]


def display_value(value):
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, time):
        return value.strftime("%H:%M")
    return str(value)


def preview_rows(rows):
    formatted = []
    for row in rows:
        formatted.append({key: display_value(val) for key, val in row.items()})
    return formatted


def parse_date_value(value):
    if value is None:
        raise ValueError("Datum fehlt")
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        try:
            return (datetime(1899, 12, 30) + timedelta(days=float(value))).date()
        except Exception as exc:
            raise ValueError("Datum ungueltig") from exc

    text = str(value).strip()
    if not text:
        raise ValueError("Datum fehlt")

    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue

    try:
        return pd.to_datetime(text, dayfirst=True).date()
    except Exception as exc:
        raise ValueError("Datum ungueltig") from exc


def parse_time_value(value):
    if value is None:
        raise ValueError("Startzeit fehlt")
    if isinstance(value, datetime):
        return value.time()
    if isinstance(value, time):
        return value
    if isinstance(value, (int, float)):
        seconds = round(float(value) * 86400)
        return (datetime(1900, 1, 1) + timedelta(seconds=seconds)).time()

    text = str(value).strip()
    if not text:
        raise ValueError("Startzeit fehlt")

    for fmt in ("%H:%M", "%H:%M:%S"):
        try:
            return datetime.strptime(text, fmt).time()
        except ValueError:
            continue

    try:
        return pd.to_datetime(text).time()
    except Exception as exc:
        raise ValueError("Startzeit ungueltig") from exc


def parse_duration_minutes(value):
    if value is None:
        raise ValueError("Dauer fehlt")
    if isinstance(value, datetime):
        value = value.time()
    if isinstance(value, time):
        return value.hour * 60 + value.minute + round(value.second / 60)
    if isinstance(value, timedelta):
        return int(round(value.total_seconds() / 60))
    if isinstance(value, (int, float)):
        number = float(value)
        if 0 < number < 1:
            return int(round(number * 1440))
        return int(round(number))

    text = str(value).strip()
    if not text:
        raise ValueError("Dauer fehlt")

    if ":" in text:
        parts = text.split(":")
        if len(parts) >= 2:
            hours = int(parts[0])
            minutes = int(parts[1])
            return hours * 60 + minutes

    try:
        return int(round(float(text)))
    except Exception as exc:
        raise ValueError("Dauer ungueltig") from exc


def prepare_event_data(row):
    pruefungsname = display_value(row.get("Prüfungsname")).strip()
    pruefer = display_value(row.get("Prüfer")).strip()
    aufsicht = display_value(row.get("Aufsicht")).strip()
    abloesung = display_value(row.get("Ablösung")).strip()
    raum = display_value(row.get("Raum")).strip()

    datum = parse_date_value(row.get("Datum"))
    startzeit = parse_time_value(row.get("Startzeit"))
    dauer_minuten = parse_duration_minutes(row.get("Dauer"))

    return {
        "pruefungsname": pruefungsname,
        "pruefer": pruefer,
        "aufsicht": aufsicht,
        "abloesung": abloesung,
        "raum": raum,
        "datum": datum,
        "startzeit": startzeit,
        "dauer_minuten": dauer_minuten,
    }


def row_fingerprint(event_data):
    parts = [
        event_data.get("pruefungsname", ""),
        event_data.get("datum").isoformat(),
        event_data.get("startzeit").strftime("%H:%M"),
        str(event_data.get("dauer_minuten")),
        event_data.get("raum", ""),
        event_data.get("aufsicht", ""),
        event_data.get("abloesung", ""),
    ]
    joined = "|".join(part.strip() for part in parts)
    return hashlib.sha256(joined.encode("utf-8")).hexdigest()

