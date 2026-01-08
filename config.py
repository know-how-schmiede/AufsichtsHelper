import os

from dotenv import load_dotenv

basedir = os.path.abspath(os.path.dirname(__file__))
load_dotenv()


class Config:
    SECRET_KEY = os.environ.get("SECRET_KEY", "dev-secret")
    SQLALCHEMY_DATABASE_URI = os.environ.get("DATABASE_URL") or (
        f"sqlite:///{os.path.join(basedir, 'instance', 'app.db')}"
    )
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    UPLOAD_FOLDER = os.path.join(basedir, "instance", "uploads")
    EXPORT_FOLDER = os.path.join(basedir, "instance", "exports")
    MAX_CONTENT_LENGTH = 10 * 1024 * 1024
    TIMEZONE = os.environ.get("APP_TIMEZONE", "Europe/Berlin")
    APP_BASE_URL = os.environ.get("APP_BASE_URL", "")

