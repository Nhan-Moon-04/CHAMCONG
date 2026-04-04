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


def log_action(
    table_name,
    record_id,
    action,
    changed_by="system",
    before_data=None,
    after_data=None,
    notes=None,
):
    entry = AuditLog(
        table_name=table_name,
        record_id=str(record_id),
        action=action,
        changed_by=changed_by or "system",
        before_data=_normalize(before_data) if before_data is not None else None,
        after_data=_normalize(after_data) if after_data is not None else None,
        notes=notes,
    )
    db.session.add(entry)
    return entry
