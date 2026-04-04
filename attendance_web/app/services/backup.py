import subprocess
from datetime import datetime, timedelta
from pathlib import Path

from apscheduler.schedulers.background import BackgroundScheduler

from .audit import log_action


_scheduler = None


def normalize_database_url(database_url):
    return database_url.replace("postgresql+psycopg://", "postgresql://")


def cleanup_old_backups(target_dir, retention_days):
    removed_files = []
    cutoff = datetime.now() - timedelta(days=retention_days)

    for file_path in Path(target_dir).glob("attendance_*.dump"):
        mtime = datetime.fromtimestamp(file_path.stat().st_mtime)
        if mtime < cutoff:
            file_path.unlink(missing_ok=True)
            removed_files.append(str(file_path))

    return removed_files


def run_pg_dump(database_url, target_dir, retention_days=30):
    target_path = Path(target_dir)
    target_path.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file = target_path / f"attendance_{timestamp}.dump"

    normalized_url = normalize_database_url(database_url)
    command = [
        "pg_dump",
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


def run_backup_job(app):
    with app.app_context():
        backup_file, removed_files = run_pg_dump(
            app.config["SQLALCHEMY_DATABASE_URI"],
            app.config["BACKUP_TARGET_DIR"],
            app.config["BACKUP_RETENTION_DAYS"],
        )
        log_action(
            "system_backup",
            backup_file,
            "BACKUP",
            changed_by="scheduler",
            after_data={"backup_file": backup_file, "removed_files": removed_files},
            notes="Sao luu 17h hang ngay",
        )
        from ..database import db

        db.session.commit()


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
