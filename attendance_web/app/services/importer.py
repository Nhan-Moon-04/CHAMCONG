import os
import uuid
from datetime import datetime, timedelta

import pandas as pd

from ..database import db
from ..models import AttendanceDaily, AttendanceDetail, AttendanceLog, Employee
from .attendance import parse_month_key
from .audit import log_action


def _read_csv_with_fallback(file_path):
    encodings = ["utf-8-sig", "cp1258", "latin1"]
    last_error = None

    for encoding in encodings:
        try:
            return pd.read_csv(file_path, encoding=encoding, on_bad_lines="skip")
        except Exception as exc:
            last_error = exc

    raise ValueError(f"Khong doc duoc file CSV: {last_error}")


def _read_dataframe(file_path):
    extension = os.path.splitext(file_path)[1].lower()

    if extension == ".csv":
        frame = _read_csv_with_fallback(file_path)
    elif extension in {".xlsx", ".xls"}:
        frame = pd.read_excel(file_path)
    else:
        raise ValueError("Chi ho tro file .csv, .xlsx hoac .xls")

    if frame.shape[1] < 4:
        raise ValueError("File cham cong can it nhat 4 cot: Ma, Ten, Bo phan, Thoi gian")

    frame = frame.iloc[:, :4].copy()
    frame.columns = ["employee_code", "employee_name", "department", "event_time"]

    frame["employee_code"] = (
        frame["employee_code"].astype(str).str.replace("'", "", regex=False).str.strip()
    )
    frame["employee_name"] = frame["employee_name"].astype(str).str.strip()
    frame["department"] = frame["department"].astype(str).str.strip()
    frame["event_time"] = pd.to_datetime(frame["event_time"], dayfirst=True, errors="coerce")

    frame = frame.dropna(subset=["employee_code", "event_time"])
    frame = frame[frame["employee_code"] != ""]

    if frame.empty:
        raise ValueError("Khong tim thay ban ghi hop le trong file cham cong")

    frame = frame.sort_values(by=["employee_code", "event_time"]).reset_index(drop=True)
    return frame


def _purge_month_attendance(month_key):
    start_date, end_date = parse_month_key(month_key)
    end_exclusive = end_date + timedelta(days=1)

    start_dt = datetime.combine(start_date, datetime.min.time())
    end_dt = datetime.combine(end_exclusive, datetime.min.time())

    deleted_logs = AttendanceLog.query.filter(
        AttendanceLog.event_time >= start_dt,
        AttendanceLog.event_time < end_dt,
    ).delete(synchronize_session=False)

    deleted_daily = AttendanceDaily.query.filter(
        AttendanceDaily.work_date >= start_date,
        AttendanceDaily.work_date <= end_date,
    ).delete(synchronize_session=False)

    deleted_details = AttendanceDetail.query.filter_by(month_key=month_key).delete(
        synchronize_session=False
    )

    return {
        "month_key": month_key,
        "deleted_logs": int(deleted_logs),
        "deleted_daily": int(deleted_daily),
        "deleted_details": int(deleted_details),
    }


def import_attendance_file(
    file_path,
    source_name,
    actor="system",
    month_key=None,
    replace_existing=False,
):
    frame = _read_dataframe(file_path)

    if month_key:
        frame = frame[frame["event_time"].dt.strftime("%Y-%m") == month_key].copy()
        if frame.empty:
            raise ValueError(
                f"Khong co du lieu hop le cho thang {month_key} trong file upload"
            )

    batch_id = str(uuid.uuid4())

    touched_months = sorted(frame["event_time"].dt.strftime("%Y-%m").unique().tolist())
    replaced_months = []

    if replace_existing:
        for item in touched_months:
            purge_info = _purge_month_attendance(item)
            replaced_months.append(purge_info)

    employee_codes = frame["employee_code"].unique().tolist()
    existing_rows = Employee.query.filter(Employee.employee_code.in_(employee_codes)).all()
    employee_map = {row.employee_code: row for row in existing_rows}

    for row in frame.itertuples(index=False):
        if row.employee_code not in employee_map:
            employee = Employee(
                employee_code=row.employee_code,
                full_name=row.employee_name,
                default_shift_code="X",
            )
            db.session.add(employee)
            db.session.flush()
            employee_map[row.employee_code] = employee
            log_action(
                "employees",
                row.employee_code,
                "INSERT",
                changed_by=actor,
                after_data=employee.to_dict(),
                notes="Tu dong tao nhan vien tu file cham cong",
            )

    for row in frame.itertuples(index=False):
        log_row = AttendanceLog(
            employee_code=row.employee_code,
            employee_name=row.employee_name,
            department=row.department,
            event_time=row.event_time.to_pydatetime(),
            source_file=source_name,
            import_batch=batch_id,
        )
        db.session.add(log_row)

    grouped = (
        frame.groupby(["employee_code", frame["event_time"].dt.date])
        .agg(
            first_check_in=("event_time", "min"),
            last_check_out=("event_time", "max"),
            employee_name=("employee_name", "first"),
        )
        .reset_index()
    )

    for row in grouped.itertuples(index=False):
        employee = employee_map.get(row.employee_code)
        if not employee:
            continue

        work_date = row.event_time

        first_check_in = row.first_check_in.to_pydatetime()
        last_check_out = row.last_check_out.to_pydatetime()
        total_hours = max((last_check_out - first_check_in).total_seconds() / 3600, 0.0)

        daily = AttendanceDaily.query.filter_by(employee_id=employee.id, work_date=work_date).first()
        if not daily:
            daily = AttendanceDaily(
                employee_id=employee.id,
                work_date=work_date,
                first_check_in=first_check_in,
                last_check_out=last_check_out,
                total_hours=round(total_hours, 2),
                import_batch=batch_id,
            )
            db.session.add(daily)
        else:
            daily.first_check_in = first_check_in
            daily.last_check_out = last_check_out
            daily.total_hours = round(total_hours, 2)
            daily.import_batch = batch_id

    log_action(
        "attendance_import",
        batch_id,
        "IMPORT",
        changed_by=actor,
        after_data={
            "source_file": source_name,
            "rows": int(len(frame)),
            "grouped_days": int(len(grouped)),
            "months": touched_months,
            "replace_existing": bool(replace_existing),
            "replaced_months": replaced_months,
        },
    )

    db.session.commit()

    return {
        "batch_id": batch_id,
        "rows": int(len(frame)),
        "grouped_days": int(len(grouped)),
        "months": touched_months,
        "replace_existing": bool(replace_existing),
        "replaced_months": replaced_months,
    }
