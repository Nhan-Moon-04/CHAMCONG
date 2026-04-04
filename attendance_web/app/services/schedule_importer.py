import re
import uuid
from datetime import date, datetime

from openpyxl import load_workbook

from ..database import db
from ..models import AttendanceDetail, Employee, OvertimeEntry, ShiftTemplate, WorkSchedule
from .attendance import month_key_for_date, parse_month_key
from .audit import log_action


EMPLOYEE_HEADER_PATTERN = re.compile(r"^ID\s*'?\s*(\d+)\s*-\s*(.+)$", re.IGNORECASE)


def _parse_employee_header(value):
    if value is None:
        return None

    text = str(value).strip()
    if not text:
        return None

    matched = EMPLOYEE_HEADER_PATTERN.match(text)
    if not matched:
        return None

    employee_code = matched.group(1).strip()
    employee_name = matched.group(2).strip()
    return employee_code, employee_name


def _coerce_date(value):
    if value is None or value == "":
        return None

    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    text = str(value).strip()
    if not text:
        return None

    parsed = None
    for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y"]:
        try:
            parsed = datetime.strptime(text, fmt)
            break
        except ValueError:
            continue

    return parsed.date() if parsed else None


def _normalize_shift_code(value):
    if value is None:
        return None

    text = str(value).strip().upper()
    if text in {"", "-", "NONE", "NAN", "NULL"}:
        return None

    return text


def _purge_schedule_month(month_key):
    start_date, end_date = parse_month_key(month_key)

    schedule_ids = [
        row[0]
        for row in db.session.query(WorkSchedule.id)
        .filter(WorkSchedule.month_key == month_key)
        .all()
    ]

    deleted_overtime = 0
    if schedule_ids:
        deleted_overtime = OvertimeEntry.query.filter(
            OvertimeEntry.schedule_id.in_(schedule_ids)
        ).delete(synchronize_session=False)

    deleted_schedules = WorkSchedule.query.filter_by(month_key=month_key).delete(
        synchronize_session=False
    )
    deleted_details = AttendanceDetail.query.filter(
        AttendanceDetail.work_date >= start_date,
        AttendanceDetail.work_date <= end_date,
    ).delete(synchronize_session=False)

    return {
        "month_key": month_key,
        "deleted_schedules": int(deleted_schedules),
        "deleted_overtime": int(deleted_overtime),
        "deleted_details": int(deleted_details),
    }


def import_schedule_file(
    file_path,
    source_name,
    actor="system",
    target_month=None,
    replace_existing=True,
):
    workbook = load_workbook(file_path, data_only=True)
    worksheet = workbook[workbook.sheetnames[0]]

    employee_columns = []
    for column in range(1, worksheet.max_column + 1):
        parsed = _parse_employee_header(worksheet.cell(1, column).value)
        if not parsed:
            continue
        employee_columns.append((column, parsed[0], parsed[1]))

    if not employee_columns:
        raise ValueError("Khong tim thay cot nhan vien theo dinh dang: ID <ma> - <ten>")

    shift_map = {row.code.upper(): row for row in ShiftTemplate.query.all()}
    if not shift_map:
        raise ValueError("Chua co bang ma ca lam")

    employee_codes = [item[1] for item in employee_columns]
    existing_employees = Employee.query.filter(Employee.employee_code.in_(employee_codes)).all()
    employee_map = {row.employee_code: row for row in existing_employees}

    for _, employee_code, employee_name in employee_columns:
        if employee_code in employee_map:
            continue

        employee = Employee(
            employee_code=employee_code,
            full_name=employee_name,
            default_shift_code="X",
        )
        db.session.add(employee)
        db.session.flush()
        employee_map[employee_code] = employee
        log_action(
            "employees",
            employee_code,
            "INSERT",
            changed_by=actor,
            after_data=employee.to_dict(),
            notes="Tu dong tao nhan vien tu file lich lam",
        )

    entries = []
    touched_months = set()
    invalid_shift_rows = []

    for row in range(2, worksheet.max_row + 1):
        work_date = _coerce_date(worksheet.cell(row, 2).value)
        if not work_date:
            continue

        row_month = month_key_for_date(work_date)
        if target_month and row_month != target_month:
            continue

        touched_months.add(row_month)

        for column, employee_code, _ in employee_columns:
            shift_code = _normalize_shift_code(worksheet.cell(row, column).value)
            if not shift_code:
                continue

            if shift_code not in shift_map:
                invalid_shift_rows.append(
                    f"Ngay {work_date} - NV {employee_code}: ma ca '{shift_code}' khong ton tai"
                )
                continue

            entries.append((employee_code, work_date, shift_code))

    if invalid_shift_rows:
        preview = "; ".join(invalid_shift_rows[:15])
        raise ValueError(f"File lich co ma ca sai. Chi tiet: {preview}")

    touched_month_list = sorted(touched_months)
    if not touched_month_list:
        if target_month:
            raise ValueError(f"Khong tim thay du lieu lich cho thang {target_month}")
        raise ValueError("Khong tim thay du lieu lich hop le trong file")

    replaced_months = []
    if replace_existing:
        for month_key in touched_month_list:
            replaced_months.append(_purge_schedule_month(month_key))

    created_count = 0
    updated_count = 0

    for employee_code, work_date, shift_code in entries:
        employee = employee_map[employee_code]
        shift = shift_map[shift_code]
        month_key = month_key_for_date(work_date)

        row = WorkSchedule.query.filter_by(employee_id=employee.id, work_date=work_date).first()
        if row:
            row.shift_id = shift.id
            row.month_key = month_key
            updated_count += 1
        else:
            row = WorkSchedule(
                employee_id=employee.id,
                work_date=work_date,
                month_key=month_key,
                shift_id=shift.id,
                absence_hours=0,
                notes=None,
            )
            db.session.add(row)
            created_count += 1

    batch_id = str(uuid.uuid4())
    log_action(
        "work_schedule_import",
        batch_id,
        "IMPORT",
        changed_by=actor,
        after_data={
            "source_file": source_name,
            "target_month": target_month,
            "months": touched_month_list,
            "replace_existing": bool(replace_existing),
            "replaced_months": replaced_months,
            "rows_imported": len(entries),
            "created": created_count,
            "updated": updated_count,
        },
        notes="Import lich lam tu file xlsx",
    )

    db.session.commit()

    return {
        "batch_id": batch_id,
        "months": touched_month_list,
        "replace_existing": bool(replace_existing),
        "replaced_months": replaced_months,
        "rows_imported": len(entries),
        "created": created_count,
        "updated": updated_count,
    }
