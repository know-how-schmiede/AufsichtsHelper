import os
import uuid
from datetime import datetime

from flask import (
    Blueprint,
    current_app,
    flash,
    redirect,
    render_template,
    request,
    session,
    url_for,
)
from werkzeug.utils import secure_filename

from ..excel import (
    EXPECTED_COLUMNS,
    extract_aufsichten,
    filter_rows_by_aufsicht,
    normalize_name,
    prepare_event_data,
    preview_rows,
    read_excel,
    row_fingerprint,
)
from ..extensions import db
from ..ics import build_event_uid, build_ics_event, get_uid_domain
from ..mailer import build_email_body, send_invite
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
            read_excel(upload_path)
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

        return redirect(url_for("main.preview"))

    return render_template("index.html")


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
    selected = request.args.get("aufsicht") or (aufsicht_names[0] if aufsicht_names else "")
    filtered_rows = filter_rows_by_aufsicht(df, selected) if selected else []

    person_index = build_person_index()
    missing_contacts = []
    seen_missing = set()

    def register_missing(name, reason):
        if not name:
            return
        key = (normalize_name(name), reason)
        if key in seen_missing:
            return
        seen_missing.add(key)
        missing_contacts.append({"name": name, "reason": reason})

    for row in filtered_rows:
        for role, field in ("aufsicht", "Aufsicht"), ("abloesung", "Ablösung"):
            name = row.get(field)
            if not name:
                continue
            person = person_index.get(normalize_name(name))
            if not person:
                register_missing(name, "Person nicht gefunden")
                continue
            if not person.email:
                register_missing(name, "Keine E-Mail hinterlegt")

    return render_template(
        "preview.html",
        expected_columns=EXPECTED_COLUMNS,
        total_rows=len(df),
        filtered_count=len(filtered_rows),
        aufsicht_names=aufsicht_names,
        selected_aufsicht=selected,
        preview_rows=preview_rows(filtered_rows),
        missing_contacts=missing_contacts,
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

    results = {"sent": [], "skipped": [], "errors": []}
    uid_domain = get_uid_domain(current_app.config.get("APP_BASE_URL", ""))

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
        label = f"{event_data['pruefungsname']} ({event_data['datum'].isoformat()} {event_data['startzeit'].strftime('%H:%M')})"

        for role, field in ("aufsicht", "Aufsicht"), ("abloesung", "Ablösung"):
            name = row.get(field)
            if not name:
                if role == "abloesung":
                    continue
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

            person = person_index.get(normalize_name(name))
            if not person:
                results["errors"].append(
                    {
                        "row": idx,
                        "name": label,
                        "role": role,
                        "email": "-",
                        "reason": f"Person '{name}' nicht gefunden",
                    }
                )
                continue
            if not person.email:
                results["errors"].append(
                    {
                        "row": idx,
                        "name": label,
                        "role": role,
                        "email": "-",
                        "reason": f"Keine E-Mail fuer '{name}'",
                    }
                )
                continue

            recipient_email = person.email
            existing = MailLog.query.filter_by(
                row_fingerprint=row_fp,
                role=role,
                recipient_email=recipient_email,
                status="sent",
            ).first()
            if existing and not force_resend:
                results["skipped"].append(
                    {
                        "row": idx,
                        "name": label,
                        "role": role,
                        "email": recipient_email,
                        "reason": "Bereits gesendet",
                    }
                )
                continue

            event_uid = build_event_uid(row_fp, role, uid_domain)
            event_data["row_fingerprint"] = row_fp
            try:
                ics_bytes, event_uid = build_ics_event(
                    event_data, role, uid_domain, event_uid
                )
                subject = (
                    f"Ablösung: {event_data['pruefungsname']}"
                    if role == "abloesung"
                    else f"Aufsicht: {event_data['pruefungsname']}"
                )
                body = build_email_body(event_data, role)
                ics_filename = secure_filename(f"{role}_{row_fp[:12]}.ics")
                send_invite(recipient_email, subject, body, ics_bytes, ics_filename)

                db.session.add(
                    MailLog(
                        event_uid=event_uid,
                        role=role,
                        recipient_email=recipient_email,
                        row_fingerprint=row_fp,
                        sent_at=datetime.utcnow(),
                        status="sent",
                        error=None,
                    )
                )
                db.session.commit()

                results["sent"].append(
                    {
                        "row": idx,
                        "name": label,
                        "role": role,
                        "email": recipient_email,
                        "reason": "Gesendet",
                    }
                )
            except Exception as exc:
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
                        "email": recipient_email,
                        "reason": str(exc),
                    }
                )

    return render_template(
        "send_result.html",
        selected_aufsicht=aufsicht_name,
        results=results,
    )

