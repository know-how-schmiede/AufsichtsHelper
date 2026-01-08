import os

from flask import Flask

from config import Config
from .extensions import db, migrate
from .routes.main import bp as main_bp
from .routes.persons import bp as persons_bp


def create_app(config_class=Config):
    app = Flask(__name__, instance_relative_config=True)
    app.config.from_object(config_class)

    os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
    os.makedirs(app.config["EXPORT_FOLDER"], exist_ok=True)
    app.config["APP_VERSION"] = _load_app_version(app.root_path)

    db.init_app(app)
    migrate.init_app(app, db)

    app.register_blueprint(main_bp)
    app.register_blueprint(persons_bp)

    return app


def _load_app_version(root_path):
    version_path = os.path.join(root_path, "VERSION")
    try:
        with open(version_path, "r", encoding="utf-8") as handle:
            return handle.read().strip() or "0.0.0"
    except FileNotFoundError:
        return "0.0.0"

