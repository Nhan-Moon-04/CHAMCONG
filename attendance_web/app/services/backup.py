import gzip
import json
import os
import shutil
import subprocess
from datetime import date, datetime, time, timedelta
from decimal import Decimal
from pathlib import Path

from apscheduler.schedulers.background import BackgroundScheduler
from sqlalchemy import text
from sqlalchemy.sql.sqltypes import Boolean, Date, DateTime, Integer, Numeric, Time

from ..database import db
from ..models import (
    AdvancePayment,
    AppUser,
    AttendanceDaily,
    AttendanceDetail,
    AttendanceLog,
    AuditLog,
    Employee,
    Holiday,
    LeaveBalance,
    MonthlySalary,
    MonthlyWorkdayConfig,
    OvertimeEntry,
    ShiftTemplate,
    WorkSchedule,
)
from .audit import log_action


_scheduler = None
PORTABLE_BACKUP_SCHEMA_VERSION = 1

BACKUP_MODELS = [
    ShiftTemplate,
    AppUser,
    Employee,
    MonthlyWorkdayConfig,
    Holiday,
    WorkSchedule,
    OvertimeEntry,
    MonthlySalary,
    AdvancePayment,
    LeaveBalance,
    AttendanceLog,
    AttendanceDaily,
    AttendanceDetail,
    AuditLog,
]


def normalize_database_url(database_url):
    return database_url.replace("postgresql+psycopg://", "postgresql://")


def _resolve_pg_tool(binary_name, configured_path=None):
    executable_name = f"{binary_name}.exe" if os.name == "nt" else binary_name

    if configured_path:
        configured = Path(configured_path)
        if configured.is_file():
            return str(configured)
        if configured.is_dir():
            candidate = configured / executable_name
            if candidate.exists():
                return str(candidate)

    resolved = shutil.which(configured_path or binary_name)
    if resolved:
        return resolved

    if os.name == "nt":
        search_roots = [
            Path("C:/Program Files/PostgreSQL"),
            Path("C:/Program Files (x86)/PostgreSQL"),
        ]
        for root in search_roots:
            if not root.exists():
                continue
            for bin_dir in sorted(root.glob("*/bin"), reverse=True):
                candidate = bin_dir / executable_name
                if candidate.exists():
                    return str(candidate)

    raise FileNotFoundError(
        f"Khong tim thay {binary_name}. Cai PostgreSQL client hoac dat PG_DUMP_PATH."
    )


def cleanup_old_backups(target_dir, retention_days):
    removed_files = []
    cutoff = datetime.now() - timedelta(days=retention_days)

    candidates = []
    for pattern in ("attendance_*.dump", "attendance_full_*.json", "attendance_full_*.json.gz"):
        candidates.extend(Path(target_dir).glob(pattern))

    for file_path in sorted(set(candidates)):
        mtime = datetime.fromtimestamp(file_path.stat().st_mtime)
        if mtime < cutoff:
            file_path.unlink(missing_ok=True)
            removed_files.append(str(file_path))

    return removed_files


def run_pg_dump(database_url, target_dir, retention_days=30, pg_dump_path="pg_dump"):
    target_path = Path(target_dir)
    target_path.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file = target_path / f"attendance_{timestamp}.dump"

    normalized_url = normalize_database_url(database_url)
    resolved_pg_dump = _resolve_pg_tool("pg_dump", pg_dump_path)
    command = [
        resolved_pg_dump,
        "--format=custom",
        f"--file={backup_file}",
        normalized_url,
    ]

    result = subprocess.run(command, capture_output=True, text=True)
    if result.returncode != 0:
        stderr = result.stderr.strip() or "Khong xac dinh"
        raise RuntimeError(f"Backup that bai: {stderr}")

    removed_files = cleanup_old_backups(target_path, retention_days)
    return str(backup_file), removed_files


def _serialize_backup_value(value):
    if isinstance(value, (datetime, date, time)):
        return value.isoformat()
    if isinstance(value, Decimal):
        return float(value)
    return value


def _serialize_row_for_backup(model, row):
    payload = {}
    for column in model.__table__.columns:
        payload[column.name] = _serialize_backup_value(getattr(row, column.name))
    return payload


def _serialize_database_payload():
    payload = {
        "schema_version": PORTABLE_BACKUP_SCHEMA_VERSION,
        "created_at": datetime.utcnow().isoformat(),
        "tables": {},
        "row_counts": {},
    }

    for model in BACKUP_MODELS:
        query = db.session.query(model)
        if hasattr(model, "id"):
            query = query.order_by(model.id.asc())

        rows = [_serialize_row_for_backup(model, row) for row in query.all()]
        table_name = model.__tablename__
        payload["tables"][table_name] = rows
        payload["row_counts"][table_name] = len(rows)

    return payload


def run_portable_backup(target_dir, retention_days=30):
    target_path = Path(target_dir)
    target_path.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file = target_path / f"attendance_full_{timestamp}.json.gz"

    payload = _serialize_database_payload()
    with gzip.open(backup_file, "wt", encoding="utf-8") as output:
        json.dump(payload, output, ensure_ascii=False)

    removed_files = cleanup_old_backups(target_path, retention_days)
    summary = {
        "schema_version": payload["schema_version"],
        "created_at": payload["created_at"],
        "row_counts": payload["row_counts"],
        "total_rows": sum(payload["row_counts"].values()),
    }
    return str(backup_file), removed_files, summary


def run_database_backup(database_url, target_dir, retention_days=30, pg_dump_path="pg_dump"):
    try:
        backup_file, removed_files = run_pg_dump(
            database_url,
            target_dir,
            retention_days,
            pg_dump_path=pg_dump_path,
        )
        return backup_file, removed_files, "pg_dump", {}
    except FileNotFoundError as exc:
        backup_file, removed_files, summary = run_portable_backup(target_dir, retention_days)
        summary["fallback_reason"] = str(exc)
        return backup_file, removed_files, "portable_json", summary


def list_backup_files(target_dir):
    target_path = Path(target_dir)
    if not target_path.exists():
        return []

    file_map = {}
    for pattern in ("attendance_*.dump", "attendance_full_*.json", "attendance_full_*.json.gz"):
        for file_path in target_path.glob(pattern):
            if file_path.is_file():
                file_map[file_path.name] = file_path

    entries = []
    for file_path in file_map.values():
        suffixes = "".join(file_path.suffixes)
        backup_type = "portable_json" if suffixes.endswith(".json") or suffixes.endswith(".json.gz") else "pg_dump"
        entries.append(
            {
                "name": file_path.name,
                "path": str(file_path),
                "backup_type": backup_type,
                "size_bytes": file_path.stat().st_size,
                "modified_at": datetime.fromtimestamp(file_path.stat().st_mtime),
            }
        )

    entries.sort(key=lambda row: row["modified_at"], reverse=True)
    return entries


def _decode_backup_payload(raw_bytes, filename):
    if not raw_bytes:
        raise ValueError("File backup rong")

    lowered_name = (filename or "").lower()
    if lowered_name.endswith(".json.gz") or lowered_name.endswith(".gz"):
        raw_text = gzip.decompress(raw_bytes).decode("utf-8")
    else:
        raw_text = raw_bytes.decode("utf-8")

    payload = json.loads(raw_text)
    if not isinstance(payload, dict):
        raise ValueError("Noi dung backup khong hop le")

    tables = payload.get("tables")
    if not isinstance(tables, dict):
        raise ValueError("Backup khong co du lieu bang")

    return payload


def _coerce_column_value(column, value):
    if value is None:
        return None

    column_type = column.type

    if isinstance(column_type, DateTime):
        if isinstance(value, datetime):
            return value
        return datetime.fromisoformat(str(value))

    if isinstance(column_type, Date):
        if isinstance(value, date):
            return value
        return date.fromisoformat(str(value))

    if isinstance(column_type, Time):
        if isinstance(value, time):
            return value
        return time.fromisoformat(str(value))

    if isinstance(column_type, Numeric):
        return Decimal(str(value))

    if isinstance(column_type, Boolean):
        if isinstance(value, bool):
            return value
        if isinstance(value, (int, float)):
            return bool(value)
        return str(value).strip().lower() in {"1", "true", "t", "yes", "y"}

    if isinstance(column_type, Integer):
        return int(value)

    return value


def _coerce_row(model, row):
    prepared = {}
    for column in model.__table__.columns:
        if column.name not in row:
            continue
        prepared[column.name] = _coerce_column_value(column, row[column.name])
    return prepared


def _truncate_existing_data():
    dialect_name = db.session.get_bind().dialect.name

    if dialect_name == "postgresql":
        table_names = ", ".join(f'"{model.__tablename__}"' for model in BACKUP_MODELS)
        db.session.execute(text(f"TRUNCATE TABLE {table_names} RESTART IDENTITY CASCADE"))
        return

    for model in reversed(BACKUP_MODELS):
        db.session.query(model).delete()


def _reset_postgres_sequences():
    if db.session.get_bind().dialect.name != "postgresql":
        return

    for model in BACKUP_MODELS:
        primary_keys = list(model.__table__.primary_key.columns)
        if len(primary_keys) != 1:
            continue

        pk_column = primary_keys[0]
        if not isinstance(pk_column.type, Integer):
            continue

        sequence_name = db.session.execute(
            text("SELECT pg_get_serial_sequence(:table_name, :column_name)"),
            {"table_name": model.__tablename__, "column_name": pk_column.name},
        ).scalar()

        if not sequence_name:
            continue

        max_value = db.session.execute(
            text(f'SELECT COALESCE(MAX("{pk_column.name}"), 0) FROM "{model.__tablename__}"')
        ).scalar()

        if max_value and int(max_value) > 0:
            db.session.execute(
                text("SELECT setval(to_regclass(:sequence_name), :value, true)"),
                {"sequence_name": sequence_name, "value": int(max_value)},
            )
        else:
            db.session.execute(
                text("SELECT setval(to_regclass(:sequence_name), 1, false)"),
                {"sequence_name": sequence_name},
            )


def restore_portable_backup(file_stream, filename, fallback_user_password_hash=None):
    payload = _decode_backup_payload(file_stream.read(), filename)
    table_payload = payload.get("tables", {})

    _truncate_existing_data()

    inserted_counts = {}
    reset_user_password_count = 0
    for model in BACKUP_MODELS:
        table_name = model.__tablename__
        rows = table_payload.get(table_name, [])

        if rows is None:
            rows = []
        if not isinstance(rows, list):
            raise ValueError(f"Du lieu bang {table_name} khong hop le")

        prepared_rows = []
        for row in rows:
            if not isinstance(row, dict):
                raise ValueError(f"Dong du lieu trong bang {table_name} khong hop le")

            prepared_row = _coerce_row(model, row)
            if model is AppUser:
                password_hash = str(prepared_row.get("password_hash") or "").strip()
                if not password_hash:
                    if not fallback_user_password_hash:
                        raise ValueError(
                            "File backup thieu password_hash trong bang app_users. "
                            "Hay tao backup moi hoac cung cap fallback hash."
                        )
                    prepared_row["password_hash"] = fallback_user_password_hash
                    reset_user_password_count += 1

            prepared_rows.append(prepared_row)

        if prepared_rows:
            db.session.bulk_insert_mappings(model, prepared_rows)

        inserted_counts[table_name] = len(prepared_rows)

    db.session.flush()
    _reset_postgres_sequences()

    return {
        "schema_version": payload.get("schema_version"),
        "backup_created_at": payload.get("created_at"),
        "row_counts": inserted_counts,
        "reset_user_password_count": reset_user_password_count,
        "total_rows": sum(inserted_counts.values()),
    }


def run_backup_job(app):
    with app.app_context():
        try:
            backup_file, removed_files, backup_type, summary = run_database_backup(
                app.config["SQLALCHEMY_DATABASE_URI"],
                app.config["BACKUP_TARGET_DIR"],
                app.config["BACKUP_RETENTION_DAYS"],
                app.config.get("PG_DUMP_PATH", "pg_dump"),
            )
            backup_record_id = Path(backup_file).name
            log_action(
                "system_backup",
                backup_record_id,
                "BACKUP",
                changed_by="scheduler",
                after_data={
                    "backup_file": backup_file,
                    "removed_files": removed_files,
                    "backup_type": backup_type,
                    "summary": summary,
                },
                notes="Sao luu 17h hang ngay",
            )
            db.session.commit()
        except Exception:
            db.session.rollback()
            raise


def register_scheduler(app):
    global _scheduler

    if not app.config.get("ENABLE_BACKUP_SCHEDULER", True):
        return

    if _scheduler and _scheduler.running:
        return

    _scheduler = BackgroundScheduler(timezone=app.config["TIMEZONE"])
    _scheduler.add_job(
        run_backup_job,
        trigger="cron",
        hour=17,
        minute=0,
        args=[app],
        id="daily_backup",
        replace_existing=True,
    )
    _scheduler.start()
