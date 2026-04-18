from calendar import monthrange
from datetime import date, datetime, time, timedelta
from types import SimpleNamespace

from sqlalchemy.orm import joinedload

from ..database import db
from ..models import (
    AttendanceDaily,
    AttendanceDetail,
    Employee,
    Holiday,
    LeaveBalance,
    MonthlySalary,
    MonthlyWorkdayConfig,
    ShiftTemplate,
    WorkSchedule,
)
from .audit import log_action


def month_key_for_date(value):
    return value.strftime("%Y-%m")


def current_month_key():
    return datetime.now().strftime("%Y-%m")


def parse_month_key(month_key):
    year, month = [int(part) for part in month_key.split("-")]
    start = date(year, month, 1)
    end = date(year, month, monthrange(year, month)[1])
    return start, end


def _to_float(value):
    if value is None:
        return 0.0
    return float(value)


def _hours_between(start_at, end_at):
    if not start_at or not end_at:
        return 0.0
    value = (end_at - start_at).total_seconds() / 3600
    return max(value, 0.0)


def leave_deduction(status_code):
    if status_code == "P":
        return 1.0
    if status_code in {"S", "C"}:
        return 0.5
    return 0.0


STATUS_NOTE_LABELS = {
    "P": "Nghi phep",
    "S": "Nghi sang",
    "C": "Nghi chieu",
    "N": "Nghi khong phep",
    "OFF": "OFF",
    "L": "Ngay le",
}

MANUAL_WORK_OVERRIDE_NOTE = "Xac nhan co di lam do mat cham cong"


def has_manual_work_override(notes):
    note_text = str(notes or "").lower()
    if not note_text:
        return False
    return MANUAL_WORK_OVERRIDE_NOTE.lower() in note_text


def _append_note(note_list, text):
    value = (text or "").strip()
    if not value:
        return
    if value not in note_list:
        note_list.append(value)


def _format_hours_text(value):
    number = round(_to_float(value), 2)
    if number.is_integer():
        return str(int(number))
    return f"{number:.2f}".rstrip("0").rstrip(".")


def ensure_default_data(actor="system"):
    default_shifts = [
        {
            "code": "X",
            "name": "Ca sang cong nhan nam",
            "start_time": time(7, 0),
            "end_time": time(16, 0),
            "break_minutes": 60,
            "standard_hours": 8,
            "default_overtime_hours": 0,
            "meal_allowance": 25000,
            "is_leave_code": False,
            "is_paid_leave": False,
            "notes": "7h-16h, nghi 1h",
        },
        {
            "code": "XVP",
            "name": "Ca van phong",
            "start_time": time(8, 0),
            "end_time": time(17, 0),
            "break_minutes": 60,
            "standard_hours": 8,
            "default_overtime_hours": 0,
            "meal_allowance": 30000,
            "is_leave_code": False,
            "is_paid_leave": False,
            "notes": "8h-17h, nghi 1h",
        },
        {
            "code": "N4",
            "name": "Ca sang cong nhan nu",
            "start_time": time(6, 0),
            "end_time": time(18, 0),
            "break_minutes": 30,
            "standard_hours": 11.5,
            "default_overtime_hours": 0,
            "meal_allowance": 30000,
            "is_leave_code": False,
            "is_paid_leave": False,
            "notes": "6h-18h, nghi 30p",
        },
        {
            "code": "XT",
            "name": "Ca toi cong nhan nam",
            "start_time": time(18, 0),
            "end_time": time(6, 0),
            "break_minutes": 60,
            "standard_hours": 11,
            "default_overtime_hours": 0,
            "meal_allowance": 35000,
            "is_leave_code": False,
            "is_paid_leave": False,
            "notes": "18h-6h ngay hom sau, nghi 1h",
        },
        {
            "code": "X3",
            "name": "Ca nu toi (8h + 3h OT)",
            "start_time": time(18, 0),
            "end_time": time(6, 0),
            "break_minutes": 60,
            "standard_hours": 8,
            "default_overtime_hours": 3,
            "meal_allowance": 35000,
            "is_leave_code": False,
            "is_paid_leave": False,
            "notes": "Mac dinh 8h lam + 3h tang ca",
        },
        {
            "code": "X4",
            "name": "Ca dac biet (8h + 4h OT)",
            "start_time": time(7, 0),
            "end_time": time(19, 0),
            "break_minutes": 60,
            "standard_hours": 8,
            "default_overtime_hours": 4,
            "meal_allowance": 35000,
            "is_leave_code": False,
            "is_paid_leave": False,
            "notes": "Mac dinh 8h lam + 4h tang ca",
        },
        {
            "code": "N",
            "name": "Nghi khong phep",
            "start_time": None,
            "end_time": None,
            "break_minutes": 0,
            "standard_hours": 8,
            "default_overtime_hours": 0,
            "meal_allowance": 0,
            "is_leave_code": True,
            "is_paid_leave": False,
            "notes": "Khong huong luong",
        },
        {
            "code": "S",
            "name": "Nghi phep sang",
            "start_time": None,
            "end_time": None,
            "break_minutes": 0,
            "standard_hours": 8,
            "default_overtime_hours": 0,
            "meal_allowance": 0,
            "is_leave_code": True,
            "is_paid_leave": True,
            "notes": "Tru 0.5 ngay phep nam",
        },
        {
            "code": "C",
            "name": "Nghi phep chieu",
            "start_time": None,
            "end_time": None,
            "break_minutes": 0,
            "standard_hours": 8,
            "default_overtime_hours": 0,
            "meal_allowance": 0,
            "is_leave_code": True,
            "is_paid_leave": True,
            "notes": "Tru 0.5 ngay phep nam",
        },
        {
            "code": "P",
            "name": "Nghi phep ca ngay",
            "start_time": None,
            "end_time": None,
            "break_minutes": 0,
            "standard_hours": 8,
            "default_overtime_hours": 0,
            "meal_allowance": 0,
            "is_leave_code": True,
            "is_paid_leave": True,
            "notes": "Tru 1 ngay phep nam",
        },
        {
            "code": "L",
            "name": "Nghi le",
            "start_time": None,
            "end_time": None,
            "break_minutes": 0,
            "standard_hours": 8,
            "default_overtime_hours": 0,
            "meal_allowance": 0,
            "is_leave_code": True,
            "is_paid_leave": True,
            "notes": "Ngay le van tinh 1 ngay cong",
        },
    ]

    for data in default_shifts:
        existing = ShiftTemplate.query.filter_by(code=data["code"]).first()
        if existing:
            continue
        shift = ShiftTemplate(**data)
        db.session.add(shift)
        log_action("shift_templates", data["code"], "INSERT", changed_by=actor, after_data=data)

    sample_employees = [
        {
            "employee_code": "1",
            "full_name": "NGUYEN THI MO",
            "gender": "Nu",
            "hometown": "Vinh Phuc",
            "birth_year": 1997,
            "default_shift_code": "N4",
        },
        {
            "employee_code": "2",
            "full_name": "PHUNG VAN GHET",
            "gender": "Nam",
            "hometown": "Thai Binh",
            "birth_year": 1994,
            "default_shift_code": "X",
        },
        {
            "employee_code": "3",
            "full_name": "DANG VAN TINH",
            "gender": "Nam",
            "hometown": "Nam Dinh",
            "birth_year": 1992,
            "default_shift_code": "X",
        },
    ]

    for data in sample_employees:
        existing = Employee.query.filter_by(employee_code=data["employee_code"]).first()
        if existing:
            continue
        employee = Employee(**data)
        db.session.add(employee)
        log_action("employees", data["employee_code"], "INSERT", changed_by=actor, after_data=data)

    db.session.commit()


def rebuild_leave_balances(year):
    employees = Employee.query.filter_by(is_active=True).all()
    start_date = date(year, 1, 1)
    end_date = date(year, 12, 31)

    for employee in employees:
        details = AttendanceDetail.query.filter(
            AttendanceDetail.employee_id == employee.id,
            AttendanceDetail.work_date >= start_date,
            AttendanceDetail.work_date <= end_date,
        ).all()

        used_days = 0.0
        for detail in details:
            used_days += leave_deduction(detail.status_code)

        balance = LeaveBalance.query.filter_by(employee_id=employee.id, year=year).first()
        if not balance:
            balance = LeaveBalance(
                employee_id=employee.id,
                year=year,
                total_days=12,
                used_days=round(used_days, 2),
            )
            db.session.add(balance)
        else:
            balance.used_days = round(used_days, 2)


def _compute_month_detail_payloads(month_key, target_employee_id=None):
    start_date, end_date = parse_month_key(month_key)

    employee_query = Employee.query.filter_by(is_active=True).order_by(Employee.employee_code.asc())
    if target_employee_id is not None:
        employee_query = employee_query.filter(Employee.id == target_employee_id)

    employees = employee_query.all()
    if not employees:
        return []

    employee_ids = [row.id for row in employees]

    shift_rows = ShiftTemplate.query.all()
    shift_by_code = {row.code.upper(): row for row in shift_rows}

    holiday_rows = Holiday.query.filter(
        Holiday.holiday_date >= start_date,
        Holiday.holiday_date <= end_date,
    ).all()
    # is_paid is used by UI as a "tick nghi" flag: checked dates are treated as OFF.
    holiday_map = {row.holiday_date: row for row in holiday_rows}

    salary_rows = MonthlySalary.query.filter_by(month_key=month_key).all()
    salary_map = {row.employee_id: row for row in salary_rows}

    month_config = MonthlyWorkdayConfig.query.filter_by(month_key=month_key).first()
    company_work_days = _to_float(month_config.company_work_days if month_config else None)
    if company_work_days <= 0:
        legacy_coeff = 0.0
        for row in salary_rows:
            value = _to_float(row.salary_coefficient)
            if value >= 10:
                legacy_coeff = value
                break
        company_work_days = legacy_coeff if legacy_coeff > 0 else 26.0

    schedules = (
        WorkSchedule.query.options(
            joinedload(WorkSchedule.shift),
            joinedload(WorkSchedule.overtime),
        )
        .filter(
            WorkSchedule.employee_id.in_(employee_ids),
            WorkSchedule.work_date >= start_date,
            WorkSchedule.work_date <= end_date,
        )
        .all()
    )
    schedule_map = {(row.employee_id, row.work_date): row for row in schedules}

    daily_rows = AttendanceDaily.query.filter(
        AttendanceDaily.employee_id.in_(employee_ids),
        AttendanceDaily.work_date >= start_date,
        AttendanceDaily.work_date <= end_date + timedelta(days=1),
    ).all()
    attendance_map = {(row.employee_id, row.work_date): row for row in daily_rows}

    payload_rows = []

    for employee in employees:
        current = start_date
        while current <= end_date:
            schedule = schedule_map.get((employee.id, current))
            has_explicit_schedule = schedule is not None
            manual_work_override = bool(schedule and has_manual_work_override(schedule.notes))
            is_sunday = current.weekday() == 6
            holiday_row = holiday_map.get(current)
            is_holiday_off = bool(holiday_row and holiday_row.is_paid)
            is_sunday_off = is_sunday and (holiday_row.is_paid if holiday_row is not None else True)
            row_notes = []

            shift = None
            absence_hours = 0.0

            if schedule:
                shift = schedule.shift
                absence_hours = _to_float(schedule.absence_hours)
                if schedule.notes:
                    row_notes.append(schedule.notes)
            else:
                if is_sunday:
                    shift = None
                elif is_holiday_off:
                    shift = None
                else:
                    shift = shift_by_code.get((employee.default_shift_code or "X").upper())
                    if not shift:
                        shift = shift_by_code.get("X")

            planned_shift_code = "OFF" if shift is None else shift.code.upper()
            status_code = "OFF" if shift is None else shift.code.upper()

            attendance = attendance_map.get((employee.id, current))
            check_in = attendance.first_check_in if attendance else None
            check_out = attendance.last_check_out if attendance else None

            if shift and shift.start_time and shift.end_time and shift.end_time <= shift.start_time:
                next_day = attendance_map.get((employee.id, current + timedelta(days=1)))
                if next_day and next_day.first_check_in and next_day.first_check_in.hour < 12:
                    if check_out is None or next_day.first_check_in > check_out:
                        check_out = next_day.first_check_in

            has_scan = bool(check_in or check_out)

            total_span_hours = _hours_between(check_in, check_out)

            standard_hours = _to_float(shift.standard_hours if shift else 0)

            overtime_hours = 0.0
            if schedule and schedule.overtime is not None:
                overtime_hours = _to_float(schedule.overtime.hours)
            elif shift:
                overtime_hours = _to_float(shift.default_overtime_hours)

            adjusted_overtime = max(overtime_hours - absence_hours, 0.0)
            remaining_absence = max(absence_hours - overtime_hours, 0.0)

            if status_code == "N":
                base_paid_hours = 0.0
            elif status_code in {"S", "C"}:
                base_paid_hours = max((standard_hours / 2.0) - remaining_absence, 0.0)
            elif status_code == "OFF":
                base_paid_hours = 0.0
            else:
                base_paid_hours = max(standard_hours - remaining_absence, 0.0)

            should_have_attendance = bool(shift and not shift.is_leave_code)
            if should_have_attendance and (not check_in or not check_out) and not manual_work_override:
                status_code = "N"
                base_paid_hours = 0.0
                adjusted_overtime = 0.0

            actual_work_hours = 0.0
            if check_in and check_out:
                # Match VBA ModChamCong2: Gio Thuc = Gio Ra - Gio Vao.
                actual_work_hours = _hours_between(check_in, check_out)

            deviation_hours = actual_work_hours - standard_hours if standard_hours else actual_work_hours

            paid_hours = base_paid_hours + adjusted_overtime
            if status_code in {"P", "L"}:
                paid_hours = standard_hours
            if status_code in {"S", "C"} and paid_hours == 0 and standard_hours > 0:
                paid_hours = standard_hours / 2.0
            if status_code in {"N", "OFF"}:
                paid_hours = 0.0

            salary = salary_map.get(employee.id)
            monthly_wage = 0.0
            if salary:
                monthly_wage = _to_float(salary.base_daily_wage)

            base_daily = (monthly_wage / company_work_days) if company_work_days > 0 else 0.0

            if standard_hours > 0:
                daily_wage = base_daily * (paid_hours / standard_hours)
            else:
                daily_wage = 0.0

            meal_allowance = (
                _to_float(shift.meal_allowance)
                if shift and not shift.is_leave_code and status_code not in {"N"}
                else 0.0
            )

            context_note = None
            if has_scan and not has_explicit_schedule:
                if is_sunday_off:
                    context_note = "Khong co lich lam van cham cong (Chu nhat OFF)"
                elif is_sunday:
                    context_note = "Khong co lich lam van cham cong (Chu nhat, chi ai co ca moi lam)"
                elif is_holiday_off:
                    context_note = "Khong co lich lam van cham cong (Ngay le OFF)"
                else:
                    context_note = "Khong co lich"
            elif not has_scan and not has_explicit_schedule and is_holiday_off:
                context_note = "Ngay le OFF"

            if context_note:
                _append_note(row_notes, context_note)
            else:
                status_note = STATUS_NOTE_LABELS.get((status_code or "").upper())
                _append_note(row_notes, status_note)

            note_issue = None
            if not has_scan:
                if should_have_attendance and not manual_work_override:
                    if has_explicit_schedule:
                        note_issue = "Bo ca"
                    elif not (is_sunday or is_holiday_off):
                        note_issue = "Khong quet the"
            elif check_in and check_out and check_in == check_out:
                note_issue = "Quen checkout" if check_in.hour < 12 else "Quen checkin"
            elif check_in and not check_out:
                note_issue = "Quen checkout"
            elif check_out and not check_in:
                note_issue = "Quen checkin"

            _append_note(row_notes, note_issue)

            if absence_hours > 0:
                _append_note(row_notes, f"Nghi {_format_hours_text(absence_hours)} gio")

            payload_rows.append(
                {
                    "employee_id": employee.id,
                    "work_date": current,
                    "month_key": month_key,
                    "shift_code": planned_shift_code,
                    "shift_name": shift.name if shift else "Nghi",
                    "check_in": check_in,
                    "check_out": check_out,
                    "standard_hours": round(standard_hours, 2),
                    "actual_work_hours": round(actual_work_hours, 2),
                    "deviation_hours": round(deviation_hours, 2),
                    "overtime_hours": round(adjusted_overtime, 2),
                    "total_span_hours": round(total_span_hours, 2),
                    "status_code": status_code,
                    "paid_hours": round(paid_hours, 2),
                    "daily_wage": round(daily_wage, 2),
                    "notes": "; ".join(row_notes) if row_notes else None,
                    "meal_allowance_daily": round(meal_allowance, 2),
                    "_employee": employee,
                }
            )

            current += timedelta(days=1)

    def _employee_sort_key(employee):
        raw_code = (employee.employee_code or "").strip()
        code = raw_code.replace("'", "").strip()
        if code.isdigit():
            return (0, int(code))
        return (1, code.lower())

    payload_rows.sort(
        key=lambda row: (
            _employee_sort_key(row["_employee"]),
            row["work_date"],
            row["_employee"].id,
        )
    )
    return payload_rows


def build_live_month_details(month_key, employee_id=None):
    payload_rows = _compute_month_detail_payloads(month_key, target_employee_id=employee_id)
    live_rows = []

    for payload in payload_rows:
        row_data = dict(payload)
        employee = row_data.pop("_employee")
        live_row = SimpleNamespace(**row_data)
        live_row.employee = employee
        live_rows.append(live_row)

    return live_rows


def rebuild_month_details(month_key, actor="system", write_audit=True):
    start_date, _ = parse_month_key(month_key)

    payload_rows = _compute_month_detail_payloads(month_key)

    AttendanceDetail.query.filter_by(month_key=month_key).delete()

    created_count = 0
    for payload in payload_rows:
        row_data = {key: value for key, value in payload.items() if key != "_employee"}
        detail = AttendanceDetail(**row_data)
        db.session.add(detail)
        created_count += 1

    rebuild_leave_balances(start_date.year)

    if write_audit:
        log_action(
            "attendance_details",
            month_key,
            "REBUILD",
            changed_by=actor,
            after_data={"month_key": month_key, "records": created_count},
            notes="Tai tao bang chi tiet cham cong theo thang",
        )

    db.session.commit()
    return created_count
