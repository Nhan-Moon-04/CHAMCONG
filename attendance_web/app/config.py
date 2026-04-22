import os
from pathlib import Path

from dotenv import load_dotenv
from sqlalchemy.engine import make_url

BASE_DIR = Path(__file__).resolve().parent.parent
load_dotenv(BASE_DIR / ".env")


def _get_int_env(name, default_value, minimum=None):
    raw_value = (os.getenv(name) or "").strip()
    if not raw_value:
        value = default_value
    else:
        try:
            value = int(raw_value)
        except ValueError:
            value = default_value

    if minimum is not None:
        value = max(minimum, value)

    return value


def _get_bool_env(name, default_value):
    raw_value = (os.getenv(name) or "").strip().lower()
    if not raw_value:
        return default_value
    return raw_value in {"1", "true", "yes", "on"}


def normalize_database_url(url):
    if not url:
        return "postgresql+psycopg://postgres:postgres@localhost:5432/attendance_db"
    if url.startswith("postgresql://"):
        return url.replace("postgresql://", "postgresql+psycopg://", 1)
    if url.startswith("postgres://"):
        return url.replace("postgres://", "postgresql+psycopg://", 1)
    return url


def _is_running_in_docker():
    if _get_bool_env("RUNNING_IN_DOCKER", False):
        return True
    return os.path.exists("/.dockerenv")


def _adapt_database_host_for_runtime(url):
    if not url:
        return url

    try:
        parsed_url = make_url(url)
    except Exception:
        return url

    if parsed_url.get_backend_name() == "sqlite":
        return url

    override_host = (os.getenv("DB_HOST_OVERRIDE") or "").strip()
    if override_host:
        return parsed_url.set(host=override_host).render_as_string(hide_password=False)

    if parsed_url.host != "db":
        return url

    if _is_running_in_docker():
        return url

    return parsed_url.set(host="localhost").render_as_string(hide_password=False)


def build_engine_options(database_url):
    if database_url.startswith("sqlite:"):
        return {}

    application_name = (os.getenv("DB_APPLICATION_NAME") or "attendance_web").strip() or "attendance_web"

    return {
        "pool_pre_ping": _get_bool_env("DB_POOL_PRE_PING", True),
        "pool_recycle": _get_int_env("DB_POOL_RECYCLE_SECONDS", 1800, minimum=1),
        "pool_timeout": _get_int_env("DB_POOL_TIMEOUT_SECONDS", 30, minimum=1),
        "pool_size": _get_int_env("DB_POOL_SIZE", 5, minimum=1),
        "max_overflow": _get_int_env("DB_MAX_OVERFLOW", 10, minimum=0),
        "pool_use_lifo": _get_bool_env("DB_POOL_USE_LIFO", True),
        "connect_args": {
            "connect_timeout": _get_int_env("DB_CONNECT_TIMEOUT_SECONDS", 10, minimum=1),
            "application_name": application_name,
        },
    }


DATABASE_URL = _adapt_database_host_for_runtime(
    normalize_database_url(
        os.getenv(
            "DATABASE_URL",
            "postgresql+psycopg://postgres:postgres@localhost:5432/attendance_db",
        )
    )
)


class Config:
    SECRET_KEY = os.getenv("SECRET_KEY", "dev-key-change")
    APP_NAME = os.getenv("APP_NAME", "HIEP LOI")
    LOGIN_USERNAME = os.getenv("LOGIN_USERNAME", "admin")
    LOGIN_PASSWORD = os.getenv("LOGIN_PASSWORD", "123456")
    SQLALCHEMY_DATABASE_URI = DATABASE_URL
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    SQLALCHEMY_ENGINE_OPTIONS = build_engine_options(DATABASE_URL)

    UPLOAD_FOLDER = str(BASE_DIR / "uploads")
    BACKUP_TARGET_DIR = os.getenv("BACKUP_TARGET_DIR", str(BASE_DIR / "backups"))
    BACKUP_RETENTION_DAYS = int(os.getenv("BACKUP_RETENTION_DAYS", "30"))
    PG_DUMP_PATH = os.getenv("PG_DUMP_PATH", "pg_dump")
    ENABLE_BACKUP_SCHEDULER = os.getenv("ENABLE_BACKUP_SCHEDULER", "1") == "1"
    TIMEZONE = os.getenv("APP_TIMEZONE", "Asia/Ho_Chi_Minh")
    DB_CONNECT_RETRY_ATTEMPTS = _get_int_env("DB_CONNECT_RETRY_ATTEMPTS", 8, minimum=1)
    DB_CONNECT_RETRY_DELAY_SECONDS = _get_int_env("DB_CONNECT_RETRY_DELAY_SECONDS", 3, minimum=1)
