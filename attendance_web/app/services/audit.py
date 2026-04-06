from datetime import date, datetime, time
from decimal import Decimal

from ..database import db
from ..models import AuditLog


def _normalize(value):
    if isinstance(value, (datetime, date, time)):
        return value.isoformat()
    if isinstance(value, Decimal):
        return float(value)
    if isinstance(value, dict):
        return {k: _normalize(v) for k, v in value.items()}
    if isinstance(value, list):
        return [_normalize(v) for v in value]
    return value


def _column_max_length(column_name):
    try:
        column = AuditLog.__table__.columns[column_name]
    except KeyError:
        return None
    return getattr(column.type, "length", None)


def _fit_text(value, column_name):
    if value is None:
        return None

    text_value = str(value)
    max_length = _column_max_length(column_name)

    if not max_length or len(text_value) <= max_length:
        return text_value

    if column_name == "record_id":
        file_name = text_value.replace("\\", "/").split("/")[-1]
        if file_name:
            if len(file_name) <= max_length:
                return file_name
            return file_name[-max_length:]

        return text_value[-max_length:]

    return text_value[:max_length]


def log_action(
    table_name,
    record_id,
    action,
    changed_by="system",
    before_data=None,
    after_data=None,
    notes=None,
):
    safe_table_name = _fit_text(table_name, "table_name") or "unknown"
    safe_record_id = _fit_text(record_id, "record_id") or "-"
    safe_action = _fit_text(action, "action") or "UNKNOWN"
    safe_changed_by = _fit_text(changed_by or "system", "changed_by") or "system"
    safe_notes = _fit_text(notes, "notes")

    entry = AuditLog(
        table_name=safe_table_name,
        record_id=safe_record_id,
        action=safe_action,
        changed_by=safe_changed_by,
        before_data=_normalize(before_data) if before_data is not None else None,
        after_data=_normalize(after_data) if after_data is not None else None,
        notes=safe_notes,
    )
    db.session.add(entry)
    return entry
