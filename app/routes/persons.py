from flask import Blueprint, flash, redirect, render_template, request, url_for
from sqlalchemy.exc import IntegrityError

from ..extensions import db
from ..models import Person, PersonAlias

bp = Blueprint("persons", __name__)


@bp.route("/persons")
def list_persons():
    persons = Person.query.order_by(Person.name.asc()).all()
    return render_template("persons_list.html", persons=persons)


@bp.route("/persons/new", methods=["GET", "POST"])
def create_person():
    if request.method == "POST":
        name = (request.form.get("name") or "").strip()
        email = (request.form.get("email") or "").strip()
        active = request.form.get("active") == "1"

        if not name:
            flash("Name ist erforderlich.")
            return render_template("persons_form.html", person=None)

        person = Person(name=name, email=email or None, active=active)
        db.session.add(person)
        try:
            db.session.commit()
        except IntegrityError:
            db.session.rollback()
            flash("Name ist bereits vorhanden.")
            return render_template("persons_form.html", person=None)

        return redirect(url_for("persons.list_persons"))

    return render_template("persons_form.html", person=None)


@bp.route("/persons/<int:person_id>/edit", methods=["GET", "POST"])
def edit_person(person_id):
    person = Person.query.get_or_404(person_id)

    if request.method == "POST":
        name = (request.form.get("name") or "").strip()
        email = (request.form.get("email") or "").strip()
        active = request.form.get("active") == "1"

        if not name:
            flash("Name ist erforderlich.")
            return render_template("persons_form.html", person=person)

        person.name = name
        person.email = email or None
        person.active = active
        try:
            db.session.commit()
        except IntegrityError:
            db.session.rollback()
            flash("Name ist bereits vorhanden.")
            return render_template("persons_form.html", person=person)

        return redirect(url_for("persons.list_persons"))

    return render_template("persons_form.html", person=person)


@bp.route("/persons/<int:person_id>/delete", methods=["POST"])
def delete_person(person_id):
    person = Person.query.get_or_404(person_id)
    db.session.delete(person)
    db.session.commit()
    return redirect(url_for("persons.list_persons"))


@bp.route("/aliases", methods=["GET", "POST"])
def list_aliases():
    persons = Person.query.order_by(Person.name.asc()).all()
    aliases = PersonAlias.query.order_by(PersonAlias.alias_name.asc()).all()

    if request.method == "POST":
        person_id = request.form.get("person_id")
        alias_name = (request.form.get("alias_name") or "").strip()

        if not person_id or not alias_name:
            flash("Person und Alias sind erforderlich.")
            return render_template(
                "aliases_list.html", persons=persons, aliases=aliases
            )

        alias = PersonAlias(person_id=int(person_id), alias_name=alias_name)
        db.session.add(alias)
        try:
            db.session.commit()
        except IntegrityError:
            db.session.rollback()
            flash("Alias ist bereits vorhanden.")
        return redirect(url_for("persons.list_aliases"))

    return render_template("aliases_list.html", persons=persons, aliases=aliases)


@bp.route("/aliases/<int:alias_id>/delete", methods=["POST"])
def delete_alias(alias_id):
    alias = PersonAlias.query.get_or_404(alias_id)
    db.session.delete(alias)
    db.session.commit()
    return redirect(url_for("persons.list_aliases"))

