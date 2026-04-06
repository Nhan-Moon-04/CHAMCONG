from pathlib import Path

from dotenv import load_dotenv
from flask import Flask
from sqlalchemy import text

from .config import Config
from .database import db
from .routes import register_routes
from .services.attendance import ensure_default_data
from .services.backup import register_scheduler
from .services.users import ensure_default_admin_user


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
        db.session.execute(
            text(
                """
                ALTER TABLE advance_payments
                ADD COLUMN IF NOT EXISTS payment_method VARCHAR(32) NOT NULL DEFAULT 'cash'
                """
            )
        )
        db.session.commit()
        ensure_default_admin_user(
            app.config.get("LOGIN_USERNAME", "admin"),
            app.config.get("LOGIN_PASSWORD", "123456"),
            actor="system-seed",
        )
        ensure_default_data(actor="system-seed")

    register_routes(app)
    register_scheduler(app)

    return app
