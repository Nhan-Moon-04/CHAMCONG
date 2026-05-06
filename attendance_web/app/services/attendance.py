from calendar import monthrange
from datetime import date, datetime, time, timedelta
from types import SimpleNamespace

from sqlalchemy.orm import joinedload

from ..database import db
from ..models import (
    AttendanceDaily,
    AttendanceDetail,
    AttendanceLog,
    Employee,
    Holiday,
    LeaveBalance,
    MonthlySalary,
    MonthlyWorkdayConfig,
    ShiftTemplate,
    WorkSchedule,
)
from .audit import log_action
from .nu_shift import (
    NU_NIGHT_MODE,
    NU_STANDARD_HOURS,
    NU_STANDARD_HOURS_DEDUCTION_BY_CODE,
    build_nu_shift_day_results,
    is_nu_dynamic_shift_code,
)
from flask import current_app


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


def leave_deduction(status_code, shift_code=None):
    status = str(status_code or "").strip().upper()
    shift = str(shift_code or "").strip().upper()

    if status == "P":
        return 1.0
    if status in {"S", "C"}:
        return 0.5

    if status in {"N", "OFF", "O"}:
        return 0.0

    nu_shift_code = shift or status
    leave_hours = NU_STANDARD_HOURS_DEDUCTION_BY_CODE.get(nu_shift_code)
    if leave_hours and NU_STANDARD_HOURS > 0:
        return leave_hours / NU_STANDARD_HOURS

    return 0.0


STATUS_NOTE_LABELS = {
    "P": "Nghi phep",
    "S": "Nghi sang",
    "C": "Nghi chieu",
    "N": "Nghi khong phep",
    "OFF": "OFF",
    "O": "OFF khong cong",
    "L": "Ngay le",
}

MANUAL_WORK_OVERRIDE_NOTE = "Xac nhan co di lam do mat cham cong"
DRIVER_AUTO_OT_SHIFT_CODES = {"TX1", "TX2"}
PAID_OFF_SHIFT_CODE = "OFF"


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


def _compute_late_checkout_overtime(shift, work_date, check_out):
    if not shift or not shift.start_time or not shift.end_time or not check_out:
        return 0.0

    scheduled_end_date = work_date
    if shift.end_time <= shift.start_time:
        scheduled_end_date = work_date + timedelta(days=1)

    scheduled_end_at = datetime.combine(scheduled_end_date, shift.end_time)
    return max((check_out - scheduled_end_at).total_seconds() / 3600, 0.0)


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
            "meal_allowance": 35000,
            "is_leave_code": False,
            "is_paid_leave": False,
            "notes": "8h-17h, nghi 1h",
        },
        {
            "code": "TX1",
            "name": "Ca tai xe TX1",
            "start_time": time(7, 30),
            "end_time": time(16, 30),
            "break_minutes": 60,
            "standard_hours": 8,
            "default_overtime_hours": 0,
            "meal_allowance": 0,
            "is_leave_code": False,
            "is_paid_leave": False,
            "notes": "7h30-16h30. Tu dong tinh OT neu checkout tre hon 16h30",
        },
        {
            "code": "TX2",
            "name": "Ca tai xe TX2",
            "start_time": time(7, 0),
            "end_time": time(16, 0),
            "break_minutes": 60,
            "standard_hours": 8,
            "default_overtime_hours": 0,
            "meal_allowance": 0,
            "is_leave_code": False,
            "is_paid_leave": False,
            "notes": "7h-16h. Tu dong tinh OT neu checkout tre hon 16h",
        },
        {
            "code": "N4",
            "name": "Ca sang cong nhan nu",
            "start_time": time(6, 0),
            "end_time": time(18, 0),
            "break_minutes": 30,
            "standard_hours": 11.5,
            "default_overtime_hours": 0,
            "meal_allowance": 35000,
            "is_leave_code": False,
            "is_paid_leave": False,
            "notes": "6h-18h, nghi 30p",
        },
        {
            "code": "NU",
            "name": "Ca nu luan phien (tu dong sang/toi)",
            "start_time": time(6, 0),
            "end_time": time(18, 0),
            "break_minutes": 30,
            "standard_hours": 8,
            "default_overtime_hours": 0,
            "meal_allowance": 0,
            "is_leave_code": False,
            "is_paid_leave": False,
            "notes": "Tu dong nhan ca sang/toi theo cham cong. OT tinh theo gio vao/ra thuc te, khung 3h50-4h30 tinh 4h. Ca sang: OT >=3h cong them 35000 tien an.",
        },
        {
            "code": "NUT1",
            "name": "Ca nu luan phien +1h OT",
            "start_time": time(6, 0),
            "end_time": time(18, 0),
            "break_minutes": 30,
            "standard_hours": 8,
            "default_overtime_hours": 0,
            "meal_allowance": 35000,
            "is_leave_code": False,
            "is_paid_leave": False,
            "notes": "Ca NU +1h OT (cong them sau khi tinh OT thuc te theo gio vao/ra).",
        },
        {
            "code": "NUT2",
            "name": "Ca nu luan phien +2h OT",
            "start_time": time(6, 0),
            "end_time": time(18, 0),
            "break_minutes": 30,
            "standard_hours": 8,
            "default_overtime_hours": 0,
            "meal_allowance": 35000,
            "is_leave_code": False,
            "is_paid_leave": False,
            "notes": "Ca NU +2h OT (cong them sau khi tinh OT thuc te theo gio vao/ra).",
        },
        {
            "code": "NU1",
            "name": "Ca nu luan phien nghi 1h",
            "start_time": time(6, 0),
            "end_time": time(18, 0),
            "break_minutes": 30,
            "standard_hours": 7,
            "default_overtime_hours": 0,
            "meal_allowance": 35000,
            "is_leave_code": False,
            "is_paid_leave": False,
            "notes": "Ca NU tru 1h cong va tru 1/8 ngay phep nam.",
        },
        {
            "code": "NU2",
            "name": "Ca nu luan phien nghi 2h",
            "start_time": time(6, 0),
            "end_time": time(18, 0),
            "break_minutes": 30,
            "standard_hours": 6,
            "default_overtime_hours": 0,
            "meal_allowance": 35000,
            "is_leave_code": False,
            "is_paid_leave": False,
            "notes": "Ca NU tru 2h cong va tru 2/8 ngay phep nam.",
        },
        {
            "code": "NU3",
            "name": "Ca nu luan phien nghi 3h",
            "start_time": time(6, 0),
            "end_time": time(18, 0),
            "break_minutes": 30,
            "standard_hours": 5,
            "default_overtime_hours": 0,
            "meal_allowance": 35000,
            "is_leave_code": False,
            "is_paid_leave": False,
            "notes": "Ca NU tru 3h cong va tru 3/8 ngay phep nam.",
        },
        {
            "code": "NUN",
            "name": "Ca nu luan phien nghi nua buoi (4h)",
            "start_time": time(6, 0),
            "end_time": time(18, 0),
            "break_minutes": 30,
            "standard_hours": 4,
            "default_overtime_hours": 0,
            "meal_allowance": 35000,
            "is_leave_code": False,
            "is_paid_leave": False,
            "notes": "Ca NU tru 4h cong (nghi nua buoi) va tru 0.5 ngay phep nam.",
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
            "code": "OFF",
            "name": "Ca OFF huong luong (chu tich/giam doc)",
            "start_time": None,
            "end_time": None,
            "break_minutes": 0,
            "standard_hours": 8,
            "default_overtime_hours": 0,
            "meal_allowance": 0,
            "is_leave_code": True,
            "is_paid_leave": True,
            "notes": "Nghi OFF huong luong, khong tru phep nam, tinh du 1 ngay cong.",
        },
        {
            "code": "O",
            "name": "OFF khong cong",
            "start_time": None,
            "end_time": None,
            "break_minutes": 0,
            "standard_hours": 0,
            "default_overtime_hours": 0,
            "meal_allowance": 0,
            "is_leave_code": True,
            "is_paid_leave": False,
            "notes": "Nghi OFF khong cong, khong tru phep nam (giong ngay Chu nhat OFF).",
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

    shifts_with_30000_meal = ShiftTemplate.query.filter(ShiftTemplate.meal_allowance == 30000).all()
    for shift in shifts_with_30000_meal:
        before = shift.to_dict()
        shift.meal_allowance = 35000
        log_action(
            "shift_templates",
            shift.id,
            "UPDATE",
            changed_by=actor,
            before_data=before,
            after_data=shift.to_dict(),
            notes="Cap nhat tien an 30000 -> 35000 theo yeu cau",
        )

    nu_shift = ShiftTemplate.query.filter_by(code="NU").first()
    nu_dynamic_note = (
        "Tu dong nhan ca sang/toi theo cham cong. OT tinh theo gio vao/ra thuc te, "
        "khung 3h50-4h30 tinh 4h. Ca sang: OT >=3h cong them 35000 tien an."
    )
    if nu_shift and isinstance(nu_shift.notes, str):
        normalized_note = nu_shift.notes.lower()
        if "tien an 30000" in normalized_note or "ot 3.5" in normalized_note:
            before = nu_shift.to_dict()
            nu_shift.notes = nu_dynamic_note
            log_action(
                "shift_templates",
                nu_shift.id,
                "UPDATE",
                changed_by=actor,
                before_data=before,
                after_data=nu_shift.to_dict(),
                notes="Dong bo ghi chu ca NU theo quy tac OT dong",
            )

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
            used_days += leave_deduction(detail.status_code, detail.shift_code)

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

    employee_by_id = {row.id: row for row in employees}

    nu_shift_code_map = {}
    for employee in employees:
        current = start_date
        while current <= end_date:
            schedule = schedule_map.get((employee.id, current))
            is_sunday = current.weekday() == 6
            holiday_row = holiday_map.get(current)
            is_holiday_off = bool(holiday_row and holiday_row.is_paid)

            if schedule:
                planned_shift = schedule.shift
            else:
                default_shift_code = (employee.default_shift_code or "X").upper()

                if is_holiday_off:
                    planned_shift = None
                elif is_sunday and not is_nu_dynamic_shift_code(default_shift_code):
                    planned_shift = None
                else:
                    planned_shift = shift_by_code.get(default_shift_code)
                    if not planned_shift and not (is_sunday and is_nu_dynamic_shift_code(default_shift_code)):
                        planned_shift = shift_by_code.get("X")

            planned_shift_code = (planned_shift.code or "").upper() if planned_shift else ""
            if is_nu_dynamic_shift_code(planned_shift_code):
                nu_shift_code_map[(employee.id, current)] = planned_shift_code

            current += timedelta(days=1)

    nu_shift_day_map = {}
    if nu_shift_code_map:
        employee_id_by_code = {}
        candidate_employee_codes = set()

        for employee_id, _ in nu_shift_code_map.keys():
            employee = employee_by_id.get(employee_id)
            raw_code = str(employee.employee_code or "").strip() if employee else ""
            normalized_code = raw_code.replace("'", "").strip()

            if normalized_code:
                employee_id_by_code[normalized_code] = employee_id

            if raw_code:
                candidate_employee_codes.add(raw_code)
            if normalized_code:
                candidate_employee_codes.add(normalized_code)

        start_dt = datetime.combine(start_date, time.min)
        end_dt = datetime.combine(end_date + timedelta(days=2), time.min)

        log_rows = (
            AttendanceLog.query.filter(
                AttendanceLog.employee_code.in_(sorted(candidate_employee_codes)),
                AttendanceLog.event_time >= start_dt,
                AttendanceLog.event_time < end_dt,
            )
            .order_by(AttendanceLog.employee_code.asc(), AttendanceLog.event_time.asc())
            .all()
        )

        nu_shift_day_map = build_nu_shift_day_results(
            nu_shift_code_map=nu_shift_code_map,
            employee_id_by_code=employee_id_by_code,
            attendance_log_rows=log_rows,
        )

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
                default_shift_code = (employee.default_shift_code or "X").upper()

                if is_holiday_off:
                    shift = None
                elif is_sunday and not is_nu_dynamic_shift_code(default_shift_code):
                    shift = None
                else:
                    shift = shift_by_code.get(default_shift_code)
                    if not shift and not (is_sunday and is_nu_dynamic_shift_code(default_shift_code)):
                        shift = shift_by_code.get("X")

            planned_shift_code = "OFF" if shift is None else shift.code.upper()
            status_code = "OFF" if shift is None else shift.code.upper()
            is_paid_off_shift = bool(
                shift
                and (shift.code or "").upper() == PAID_OFF_SHIFT_CODE
                and shift.is_leave_code
                and shift.is_paid_leave
            )

            nu_shift_result = None
            if shift and is_nu_dynamic_shift_code((shift.code or "").upper()):
                nu_shift_result = nu_shift_day_map.get((employee.id, current))

            is_nu_night_week_sunday_off = bool(
                nu_shift_result and is_sunday and nu_shift_result.week_mode == NU_NIGHT_MODE
            )
            if nu_shift_result and is_sunday:
                is_effective_sunday_off = is_nu_night_week_sunday_off
            else:
                is_effective_sunday_off = is_sunday_off

            attendance = attendance_map.get((employee.id, current))
            check_in = attendance.first_check_in if attendance else None
            check_out = attendance.last_check_out if attendance else None

            if nu_shift_result:
                if nu_shift_result.check_in:
                    check_in = nu_shift_result.check_in
                if nu_shift_result.check_out:
                    check_out = nu_shift_result.check_out
                elif nu_shift_result.mode == NU_NIGHT_MODE:
                    next_day = attendance_map.get((employee.id, current + timedelta(days=1)))
                    if next_day and next_day.first_check_in and next_day.first_check_in.hour < 12:
                        check_out = next_day.first_check_in

            if is_nu_night_week_sunday_off and not manual_work_override:
                shift = None
                planned_shift_code = "OFF"
                status_code = "OFF"

                # Sunday OFF in NU night week should not accumulate any carried punches.
                check_in = None
                check_out = None

                _append_note(row_notes, "Chu nhat nghi theo tuan ca toi (NU)")

            if shift and shift.start_time and shift.end_time and shift.end_time <= shift.start_time:
                next_day = attendance_map.get((employee.id, current + timedelta(days=1)))
                if next_day and next_day.first_check_in and next_day.first_check_in.hour < 12:
                    if check_out is None or next_day.first_check_in > check_out:
                        check_out = next_day.first_check_in

            has_scan = bool(check_in or check_out)

            # OFF day from Holidays/Sunday takes precedence when there is no attendance scan.
            # This prevents OFF dates from being converted to "Nghi khong phep".
            if (
                (is_effective_sunday_off or is_holiday_off)
                and not has_scan
                and not manual_work_override
                and not is_paid_off_shift
            ):
                shift = None
                planned_shift_code = "OFF"
                status_code = "OFF"

            total_span_hours = _hours_between(check_in, check_out)

            standard_hours = _to_float(shift.standard_hours if shift else 0)
            if nu_shift_result and status_code != "OFF":
                standard_hours = _to_float(nu_shift_result.standard_hours)

            overtime_hours = 0.0
            if schedule and schedule.overtime is not None:
                overtime_hours = _to_float(schedule.overtime.hours)
            elif nu_shift_result and status_code != "OFF":
                overtime_hours = _to_float(nu_shift_result.default_overtime_hours)
            elif (
                shift
                and (shift.code or "").upper() in DRIVER_AUTO_OT_SHIFT_CODES
                and status_code != "OFF"
            ):
                # Existing driver OT: count late checkout OT
                overtime_hours = _compute_late_checkout_overtime(shift, current, check_out)
                # New rule: if employee checks in at least 1 hour earlier than scheduled start,
                # give 1 extra hour counted as OT (business request).
                try:
                    if check_in and shift.start_time:
                        scheduled_start_at = datetime.combine(current, shift.start_time)
                        early_hours = (scheduled_start_at - check_in).total_seconds() / 3600.0
                        if early_hours >= 1.0:
                            overtime_hours = max(overtime_hours + 1.0, 0.0)
                except Exception:
                    pass
            elif shift:
                overtime_hours = _to_float(shift.default_overtime_hours)

            adjusted_overtime = max(overtime_hours - absence_hours, 0.0)
            remaining_absence = max(absence_hours - overtime_hours, 0.0)

            if status_code == "N":
                base_paid_hours = 0.0
            elif status_code in {"S", "C"}:
                base_paid_hours = max((standard_hours / 2.0) - remaining_absence, 0.0)
            elif status_code in {"OFF", "O"}:
                base_paid_hours = max(standard_hours - remaining_absence, 0.0) if is_paid_off_shift else 0.0
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
            if status_code in {"N", "OFF", "O"} and not is_paid_off_shift:
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
                _to_float(nu_shift_result.meal_allowance)
                if nu_shift_result and shift and not shift.is_leave_code and status_code not in {"N", "OFF", "O"}
                else (
                    _to_float(shift.meal_allowance)
                    if shift and not shift.is_leave_code and status_code not in {"N", "O"}
                    else 0.0
                )
            )

            # Optional policy: if enabled in app config, count 1 meal when OT reaches/passes 18:00
            # Example: shift ends at 16:00, employee works until 18:00 (2h OT) -> count 1 meal.
            try:
                if (
                    not nu_shift_result
                    and current_app.config.get("ENABLE_OT_AFTER_6PM_MEAL", False)
                    and shift
                    and not shift.is_leave_code
                    and status_code not in {"N", "O", "OFF"}
                    and check_out
                ):
                    # check if the worker worked up to or beyond 18:00 on the work_date
                    if check_out.hour >= 18:
                        # Avoid double-counting if meal_allowance already set by NU logic above
                        if meal_allowance == 0.0:
                            meal_allowance = _to_float(shift.meal_allowance) if _to_float(shift.meal_allowance) > 0 else 35000.0
            except Exception:
                pass

            context_note = None
            if has_scan and not has_explicit_schedule:
                if is_effective_sunday_off:
                    context_note = "Khong co lich lam van cham cong (Chu nhat OFF)"
                elif is_sunday:
                    context_note = "Khong co lich lam van cham cong (Chu nhat, chi ai co ca moi lam)"
                elif is_holiday_off:
                    context_note = "Khong co lich lam van cham cong (Ngay le OFF)"
                else:
                    context_note = "Khong co lich"
            elif not has_scan and has_explicit_schedule and (is_effective_sunday_off or is_holiday_off):
                context_note = "Ngay OFF theo Holidays"
            elif not has_scan and not has_explicit_schedule and is_holiday_off:
                context_note = "Ngay le OFF"

            if context_note:
                _append_note(row_notes, context_note)
            else:
                if status_code == "OFF" and is_paid_off_shift:
                    status_note = "OFF huong luong"
                else:
                    status_note = STATUS_NOTE_LABELS.get((status_code or "").upper())
                _append_note(row_notes, status_note)

            note_issue = None
            if not has_scan:
                if should_have_attendance and not manual_work_override:
                    if has_explicit_schedule:
                        note_issue = "Bo ca"
                    elif not (is_effective_sunday_off or is_holiday_off):
                        note_issue = "Khong quet the"
            elif check_in and check_out and check_in == check_out:
                note_issue = "Quen checkout" if check_in.hour < 12 else "Quen checkin"
            elif check_in and not check_out:
                note_issue = "Quen checkout"
            elif check_out and not check_in:
                note_issue = "Quen checkin"

            _append_note(row_notes, note_issue)

            if nu_shift_result and nu_shift_result.warning_note:
                _append_note(row_notes, nu_shift_result.warning_note)

            if absence_hours > 0:
                _append_note(row_notes, f"Nghi {_format_hours_text(absence_hours)} gio")

            payload_rows.append(
                {
                    "employee_id": employee.id,
                    "work_date": current,
                    "month_key": month_key,
                    "shift_code": planned_shift_code,
                    "shift_name": (
                        nu_shift_result.shift_name
                        if nu_shift_result and status_code != "OFF"
                        else (shift.name if shift else "Nghi")
                    ),
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
