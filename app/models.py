from datetime import datetime

from .extensions import db


class Person(db.Model):
    __tablename__ = "persons"

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), unique=True, nullable=False)
    email = db.Column(db.String(200), nullable=True)
    active = db.Column(db.Boolean, default=True, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

    aliases = db.relationship(
        "PersonAlias", backref="person", cascade="all, delete-orphan", lazy=True
    )


class PersonAlias(db.Model):
    __tablename__ = "person_aliases"

    id = db.Column(db.Integer, primary_key=True)
    person_id = db.Column(db.Integer, db.ForeignKey("persons.id"), nullable=False)
    alias_name = db.Column(db.String(200), unique=True, nullable=False)


class MailLog(db.Model):
    __tablename__ = "mail_log"

    id = db.Column(db.Integer, primary_key=True)
    event_uid = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(20), nullable=False)
    recipient_email = db.Column(db.String(200), nullable=False)
    row_fingerprint = db.Column(db.String(64), nullable=False)
    sent_at = db.Column(db.DateTime, nullable=True)
    status = db.Column(db.String(20), nullable=False)
    error = db.Column(db.Text, nullable=True)


__all__ = ["Person", "PersonAlias", "MailLog"]

