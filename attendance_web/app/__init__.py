from pathlib import Path

from dotenv import load_dotenv
from flask import Flask

from .config import Config
from .database import db
from .routes import register_routes
from .services.attendance import ensure_default_data
from .services.backup import register_scheduler


def create_app():
    env_file = Path(__file__).resolve().parent.parent / ".env"
    load_dotenv(env_file)

    app = Flask(__name__)
    app.config.from_object(Config)

    Path(app.config["UPLOAD_FOLDER"]).mkdir(parents=True, exist_ok=True)
    Path(app.config["BACKUP_TARGET_DIR"]).mkdir(parents=True, exist_ok=True)

    db.init_app(app)

    with app.app_context():
        from . import models  # noqa: F401

        db.create_all()
        ensure_default_data(actor="system-seed")

    register_routes(app)
    register_scheduler(app)

    return app
