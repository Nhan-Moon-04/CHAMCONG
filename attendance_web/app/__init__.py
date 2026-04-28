from pathlib import Path
import time

from dotenv import load_dotenv
from flask import Flask
from sqlalchemy import text
from sqlalchemy.exc import OperationalError

from .config import Config
from .database import db
from .routes import register_routes
from .union_routes import register_union_routes
from .services.attendance import ensure_default_data
from .services.backup import register_scheduler
from .services.salary_overview_export import register_salary_overview_export
from .services.users import ensure_default_admin_user


def _initialize_database(app):
    max_attempts = app.config.get("DB_CONNECT_RETRY_ATTEMPTS", 8)
    retry_delay_seconds = app.config.get("DB_CONNECT_RETRY_DELAY_SECONDS", 3)

    for attempt in range(1, max_attempts + 1):
        try:
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
            return
        except OperationalError:
            db.session.rollback()
            db.session.remove()

            if attempt >= max_attempts:
                raise

            app.logger.warning(
                "Chua ket noi duoc PostgreSQL (lan %s/%s). Thu lai sau %s giay.",
                attempt,
                max_attempts,
                retry_delay_seconds,
            )
            time.sleep(retry_delay_seconds)


def create_app():
    env_file = Path(__file__).resolve().parent.parent / ".env"
    load_dotenv(env_file)

    app = Flask(__name__)
    app.config.from_object(Config)

    Path(app.config["UPLOAD_FOLDER"]).mkdir(parents=True, exist_ok=True)
    Path(app.config["BACKUP_TARGET_DIR"]).mkdir(parents=True, exist_ok=True)

    db.init_app(app)

    with app.app_context():
        _initialize_database(app)

    register_routes(app)
    register_salary_overview_export(app)
    register_union_routes(app)
    register_scheduler(app)

    return app
