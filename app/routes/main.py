import os
import uuid
import zipfile
from datetime import datetime

from flask import (
    Blueprint,
    abort,
    current_app,
    flash,
    redirect,
    render_template,
    request,
    send_from_directory,
    session,
    url_for,
)
from werkzeug.utils import secure_filename

from ..excel import (
    EXPECTED_COLUMNS,
    extract_aufsichten,
    filter_rows_by_aufsicht,
    display_value,
    normalize_name,
    parse_date_value,
    parse_duration_minutes,
    parse_time_value,
    prepare_event_data,
    preview_rows,
    read_excel,
    row_fingerprint,
    split_display_names,
    split_display_rooms,
    split_names,
)
from ..extensions import db
from ..ics import build_event_uid, build_ics_event, get_uid_domain
from ..models import MailLog, Person, PersonAlias

bp = Blueprint("main", __name__)

ALLOWED_EXTENSIONS = {".xlsx"}


def allowed_file(filename):
    return os.path.splitext(filename.lower())[1] in ALLOWED_EXTENSIONS


def store_upload_path(filename):
    session["upload_path"] = os.path.join("uploads", filename)


def get_upload_path():
    rel_path = session.get("upload_path")
    if not rel_path:
        return None
    return os.path.join(current_app.instance_path, rel_path)


def build_person_index():
    index = {}
    for person in Person.query.all():
        if not person.active:
            continue
        key = normalize_name(person.name)
        if key:
            index.setdefault(key, person)

    for alias in PersonAlias.query.all():
        if not alias.person or not alias.person.active:
            continue
        key = normalize_name(alias.alias_name)
        if key:
            index.setdefault(key, alias.person)

    return index


def parse_missing_columns(error_message):
    if "Missing columns:" not in error_message:
        return []
    parts = error_message.split("Missing columns:", 1)[1]
    return [col.strip() for col in parts.split(",") if col.strip()]


def build_bundle_filename(aufsicht_name):
    timestamp = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
    base_name = f"ical_{aufsicht_name}_{timestamp}.zip"
    safe_name = secure_filename(base_name) or "ical_bundle.zip"
    unique_prefix = uuid.uuid4().hex
    return f"{unique_prefix}_{safe_name}"


def default_calendar_name(rows=None):
    year = datetime.now().year
    if rows:
        for row in rows:
            try:
                year = parse_date_value(row.get("Datum")).year
                break
            except ValueError:
                continue
    return f"Pr\u00fcfungsaufsicht_{year}"


def store_calendar_name(value):
    session["calendar_name"] = value


def get_calendar_name(rows=None):
    name = (session.get("calendar_name") or "").strip()
    if name:
        return name
    return default_calendar_name(rows)


@bp.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files.get("file")
        if not file or not file.filename:
            flash("Bitte eine .xlsx-Datei auswaehlen.")
            return render_template("index.html")
        if not allowed_file(file.filename):
            flash("Nur .xlsx-Dateien sind erlaubt.")
            return render_template("index.html")

        filename = secure_filename(file.filename)
        unique_name = f"{uuid.uuid4().hex}_{filename}"
        upload_path = os.path.join(current_app.config["UPLOAD_FOLDER"], unique_name)
        file.save(upload_path)
        store_upload_path(unique_name)

        try:
            rows = read_excel(upload_path)
        except ValueError as exc:
            missing = parse_missing_columns(str(exc))
            if missing:
                return render_template(
                    "error.html",
                    title="Fehlende Spalten",
                    message="Die hochgeladene Datei hat nicht alle benoetigten Spalten.",
                    missing=missing,
                )
            return render_template(
                "error.html",
                title="Excel-Fehler",
                message=str(exc),
                missing=[],
            )

        calendar_name = (request.form.get("calendar_name") or "").strip()
        if not calendar_name:
            calendar_name = default_calendar_name(rows)
        store_calendar_name(calendar_name)

        return redirect(url_for("main.preview"))

    return render_template(
        "index.html",
        calendar_name=session.get("calendar_name", ""),
        default_calendar_name=default_calendar_name(),
    )


@bp.route("/help", methods=["GET"])
def help():
    return render_template("help.html", title="Hilfe")


@bp.route("/privacy", methods=["GET"])
def privacy():
    return render_template("privacy.html", title="Datenschutz")


@bp.route("/preview", methods=["GET"])
def preview():
    upload_path = get_upload_path()
    if not upload_path or not os.path.exists(upload_path):
        flash("Bitte zuerst eine Excel-Datei hochladen.")
        return redirect(url_for("main.index"))

    try:
        df = read_excel(upload_path)
    except ValueError as exc:
        missing = parse_missing_columns(str(exc))
        return render_template(
            "error.html",
            title="Fehlende Spalten",
            message=str(exc),
            missing=missing,
        )

    aufsicht_names = extract_aufsichten(df)
    selected = request.args.get("aufsicht") or (
        aufsicht_names[0] if aufsicht_names else ""
    )
    filter_applied = bool(request.args.get("aufsicht"))
    sort_key = request.args.get("sort") or ""
    sort_dir = request.args.get("dir") or "asc"
    if sort_dir not in ("asc", "desc"):
        sort_dir = "asc"
    filtered_rows = filter_rows_by_aufsicht(df, selected) if selected else []

    preview_columns = [col for col in EXPECTED_COLUMNS if col != "Ablösung"]
    sortable_columns = [
        "Pr\u00fcfungsname",
        "Datum",
        "Startzeit",
        "Dauer",
        "Pr\u00fcfer",
        "Raum",
    ]

    if sort_key in sortable_columns:
        def sort_value(row):
            try:
                if sort_key == "Datum":
                    return (0, parse_date_value(row.get("Datum")))
                if sort_key == "Startzeit":
                    return (0, parse_time_value(row.get("Startzeit")))
                if sort_key == "Dauer":
                    return (0, parse_duration_minutes(row.get("Dauer")))
                if sort_key == "Pr\u00fcfer":
                    items = split_display_names(row.get("Pr\u00fcfer"))
                    return (0, " ".join(items).casefold())
                if sort_key == "Raum":
                    items = split_display_rooms(row.get("Raum"))
                    return (0, " ".join(items).casefold())
                value = display_value(row.get(sort_key)).casefold()
                return (0, value)
            except ValueError:
                return (1, "")

        filtered_rows = sorted(
            filtered_rows,
            key=sort_value,
            reverse=sort_dir == "desc",
        )

    formatted_rows = preview_rows(filtered_rows)
    multiline_columns = ["Pr\u00fcfer", "Aufsicht", "Raum"]
    for row in formatted_rows:
        row["Pr\u00fcfer"] = split_display_names(row.get("Pr\u00fcfer"))
        row["Aufsicht"] = split_display_names(row.get("Aufsicht"))
        row["Raum"] = split_display_rooms(row.get("Raum"))

    return render_template(
        "preview.html",
        expected_columns=preview_columns,
        total_rows=len(df),
        filtered_count=len(filtered_rows),
        aufsicht_names=aufsicht_names,
        selected_aufsicht=selected,
        preview_rows=formatted_rows,
        missing_contacts=[],
        filter_applied=filter_applied,
        multiline_columns=multiline_columns,
        sortable_columns=sortable_columns,
        sort_key=sort_key,
        sort_dir=sort_dir,
    )


@bp.route("/send", methods=["POST"])
def send():
    upload_path = get_upload_path()
    if not upload_path or not os.path.exists(upload_path):
        flash("Bitte zuerst eine Excel-Datei hochladen.")
        return redirect(url_for("main.index"))

    aufsicht_name = request.form.get("aufsicht")
    if not aufsicht_name:
        flash("Bitte eine Aufsicht auswaehlen.")
        return redirect(url_for("main.preview"))

    force_resend = request.form.get("force_resend") == "1"

    try:
        df = read_excel(upload_path)
    except ValueError as exc:
        missing = parse_missing_columns(str(exc))
        return render_template(
            "error.html",
            title="Fehlende Spalten",
            message=str(exc),
            missing=missing,
        )

    filtered_rows = filter_rows_by_aufsicht(df, aufsicht_name)
    person_index = build_person_index()

    results = {"generated": [], "skipped": [], "errors": []}
    uid_domain = get_uid_domain(current_app.config.get("APP_BASE_URL", ""))
    calendar_name = get_calendar_name(df)
    generated_files = []
    seen_exports = set()

    target_norm = normalize_name(aufsicht_name)

    for idx, row in enumerate(filtered_rows, start=1):
        try:
            event_data = prepare_event_data(row)
            event_data["row_fingerprint"] = row_fingerprint(event_data)
        except ValueError as exc:
            results["errors"].append(
                {
                    "row": idx,
                    "name": row.get("Prüfungsname", ""),
                    "role": "-",
                    "email": "-",
                    "reason": str(exc),
                }
            )
            continue

        row_fp = event_data["row_fingerprint"]
        label = (
            f"{event_data['pruefungsname']} ({event_data['datum'].isoformat()} "
            f"{event_data['startzeit'].strftime('%H:%M')})"
        )

        role = "aufsicht"
        names = split_names(row.get("Aufsicht"))
        names = [name for name in names if normalize_name(name) == target_norm]
        if not names:
            results["errors"].append(
                {
                    "row": idx,
                    "name": label,
                    "role": role,
                    "email": "-",
                    "reason": "Aufsicht fehlt",
                }
            )
            continue

        for name in names:
            name_key = normalize_name(name)
            export_key = (row_fp, role, name_key)
            if export_key in seen_exports:
                results["skipped"].append(
                    {
                        "row": idx,
                        "name": label,
                        "role": role,
                        "email": "-",
                        "reason": "Duplikat im Import",
                    }
                )
                continue
            seen_exports.add(export_key)
            person = person_index.get(normalize_name(name))
            recipient_email = ""
            if person and person.email:
                recipient_email = person.email

            existing = None
            if recipient_email:
                existing = (
                    MailLog.query.filter_by(
                        row_fingerprint=row_fp,
                        role=role,
                        recipient_email=recipient_email,
                    )
                    .filter(MailLog.status.in_(["sent", "generated"]))
                    .first()
                )
            if existing and not force_resend:
                results["skipped"].append(
                    {
                        "row": idx,
                        "name": label,
                        "role": role,
                        "email": recipient_email or "-",
                        "reason": "Bereits erstellt",
                    }
                )
                continue

            event_uid = build_event_uid(row_fp, role, uid_domain)
            event_data["row_fingerprint"] = row_fp
            try:
                ics_bytes, event_uid = build_ics_event(
                    event_data,
                    role,
                    uid_domain,
                    event_uid,
                    calendar_name=calendar_name,
                )
                prefix = "Aufsicht"
                date_str = event_data["datum"].strftime("%Y-%m-%d")
                time_str = event_data["startzeit"].strftime("%H-%M")
                base_name = (
                    f"{prefix}_{date_str}_{time_str}_"
                    f"{event_data['pruefungsname']}_{name}.ics"
                )
                ics_filename = secure_filename(base_name)
                if not ics_filename:
                    ics_filename = f"{prefix}_{row_fp[:12]}.ics"

                generated_files.append((ics_filename, ics_bytes))
                if not person:
                    reason = "ICS erzeugt (Stammdaten fehlen)"
                elif not recipient_email:
                    reason = "ICS erzeugt (E-Mail fehlt)"
                else:
                    reason = "ICS erzeugt"

                results["generated"].append(
                    {
                        "row": idx,
                        "name": label,
                        "role": role,
                        "email": recipient_email or "-",
                        "reason": reason,
                    }
                )

                if recipient_email:
                    db.session.add(
                        MailLog(
                            event_uid=event_uid,
                            role=role,
                            recipient_email=recipient_email,
                            row_fingerprint=row_fp,
                            sent_at=datetime.utcnow(),
                            status="generated",
                            error=None,
                        )
                    )
                    db.session.commit()
            except Exception as exc:
                if recipient_email:
                    db.session.add(
                        MailLog(
                            event_uid=event_uid,
                            role=role,
                            recipient_email=recipient_email,
                            row_fingerprint=row_fp,
                            sent_at=datetime.utcnow(),
                            status="error",
                            error=str(exc),
                        )
                    )
                    db.session.commit()
                results["errors"].append(
                    {
                        "row": idx,
                        "name": label,
                        "role": role,
                        "email": recipient_email or "-",
                        "reason": str(exc),
                    }
                )

    download_filename = None
    if generated_files:
        export_dir = current_app.config["EXPORT_FOLDER"]
        os.makedirs(export_dir, exist_ok=True)
        download_filename = build_bundle_filename(aufsicht_name)
        bundle_path = os.path.join(export_dir, download_filename)

        with zipfile.ZipFile(bundle_path, "w", zipfile.ZIP_DEFLATED) as bundle:
            for filename, payload in generated_files:
                bundle.writestr(filename, payload)

    return render_template(
        "send_result.html",
        selected_aufsicht=aufsicht_name,
        results=results,
        download_filename=download_filename,
    )


@bp.route("/download/<path:filename>")
def download(filename):
    safe_name = secure_filename(filename)
    if not safe_name:
        abort(404)

    export_dir = current_app.config["EXPORT_FOLDER"]
    bundle_path = os.path.join(export_dir, safe_name)
    if not os.path.exists(bundle_path):
        abort(404)

    return send_from_directory(export_dir, safe_name, as_attachment=True)
