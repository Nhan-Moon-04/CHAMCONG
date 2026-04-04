import csv
import importlib
import io
import os
from datetime import date, datetime, timedelta
from pathlib import Path

from flask import (
    Response,
    current_app,
    flash,
    redirect,
    render_template,
    request,
    url_for,
)
from sqlalchemy import desc, func, or_
from sqlalchemy.orm import joinedload
from werkzeug.utils import secure_filename

from .database import db
from .models import (
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
from .services.attendance import (
    build_live_month_details,
    current_month_key,
    month_key_for_date,
    parse_month_key,
    rebuild_month_details,
)
from .services.audit import log_action
from .services.backup import run_pg_dump
from .services.importer import import_attendance_file
from .services.salary_importer import import_salary_file
from .services.schedule_importer import import_schedule_file


_holiday_lib = None
_holiday_lib_checked = False


def _safe_month_key(value):
    if not value:
        return current_month_key()
    try:
        datetime.strptime(value, "%Y-%m")
        return value
    except ValueError:
        return current_month_key()


def _parse_date(value):
    return datetime.strptime(value, "%Y-%m-%d").date()


def _parse_time(value):
    if not value:
        return None
    return datetime.strptime(value, "%H:%M").time()


def _to_float(value, default=0.0):
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def _iter_month_keys(start_date, end_date):
    current = date(start_date.year, start_date.month, 1)
    end_month = date(end_date.year, end_date.month, 1)

    while current <= end_month:
        yield current.strftime("%Y-%m")
        if current.month == 12:
            current = date(current.year + 1, 1, 1)
        else:
            current = date(current.year, current.month + 1, 1)


def _resolve_company_work_days(month_key):
    config = MonthlyWorkdayConfig.query.filter_by(month_key=month_key).first()
    config_value = _to_float(config.company_work_days if config else None, 0)
    if config_value > 0:
        return config_value, config

    legacy_salary = MonthlySalary.query.filter_by(month_key=month_key).order_by(MonthlySalary.id.asc()).first()
    legacy_value = _to_float(legacy_salary.salary_coefficient if legacy_salary else None, 0)
    if legacy_value >= 10:
        return legacy_value, config

    return 26.0, config


def _get_holiday_library():
    global _holiday_lib
    global _holiday_lib_checked

    if not _holiday_lib_checked:
        try:
            _holiday_lib = importlib.import_module("holidays")
        except ImportError:
            _holiday_lib = None
        _holiday_lib_checked = True

    return _holiday_lib


def _get_vietnam_holiday_map(start_date, end_date):
    holiday_map = {}
    years = list(range(start_date.year, end_date.year + 1))
    holiday_library = _get_holiday_library()

    if holiday_library:
        vn_calendar = holiday_library.country_holidays("VN", years=years)
        for holiday_date, holiday_name in vn_calendar.items():
            if start_date <= holiday_date <= end_date:
                holiday_map[holiday_date] = str(holiday_name)

    if not holiday_map:
        fixed_holidays = {
            (1, 1): "Tet Duong lich",
            (4, 30): "Ngay giai phong mien Nam",
            (5, 1): "Ngay Quoc te Lao dong",
            (9, 2): "Quoc khanh",
        }
        for year in years:
            for (month, day), holiday_name in fixed_holidays.items():
                holiday_date = date(year, month, day)
                if start_date <= holiday_date <= end_date:
                    holiday_map[holiday_date] = holiday_name

    return holiday_map


def register_routes(app):
    @app.route("/")
    def dashboard():
        month_key = _safe_month_key(request.args.get("month"))

        # Keep dashboard numbers in sync with source tables (schedule, shifts, attendance, salary).
        rebuild_month_details(month_key, actor="system-auto-sync", write_audit=False)

        total_employees = Employee.query.filter_by(is_active=True).count()
        detail_query = AttendanceDetail.query.filter_by(month_key=month_key)

        total_rows = detail_query.count()
        total_paid_hours = (
            db.session.query(func.coalesce(func.sum(AttendanceDetail.paid_hours), 0))
            .filter(AttendanceDetail.month_key == month_key)
            .scalar()
        )
        total_wage = (
            db.session.query(func.coalesce(func.sum(AttendanceDetail.daily_wage), 0))
            .filter(AttendanceDetail.month_key == month_key)
            .scalar()
        )
        absent_days = detail_query.filter(AttendanceDetail.status_code == "N").count()

        status_rows = (
            db.session.query(AttendanceDetail.status_code, func.count(AttendanceDetail.id))
            .filter(AttendanceDetail.month_key == month_key)
            .group_by(AttendanceDetail.status_code)
            .order_by(AttendanceDetail.status_code.asc())
            .all()
        )

        overtime_rows = (
            db.session.query(
                Employee.full_name,
                func.coalesce(func.sum(AttendanceDetail.overtime_hours), 0).label("ot_hours"),
            )
            .join(AttendanceDetail, AttendanceDetail.employee_id == Employee.id)
            .filter(AttendanceDetail.month_key == month_key)
            .group_by(Employee.full_name)
            .order_by(desc("ot_hours"))
            .limit(10)
            .all()
        )

        warning_rows = (
            db.session.query(AttendanceDetail, Employee)
            .join(Employee, AttendanceDetail.employee_id == Employee.id)
            .filter(AttendanceDetail.month_key == month_key)
            .filter(
                or_(
                    AttendanceDetail.status_code == "N",
                    AttendanceDetail.deviation_hours < 0,
                )
            )
            .order_by(AttendanceDetail.work_date.desc())
            .limit(20)
            .all()
        )

        return render_template(
            "dashboard.html",
            month_key=month_key,
            total_employees=total_employees,
            total_rows=total_rows,
            total_paid_hours=float(total_paid_hours or 0),
            total_wage=float(total_wage or 0),
            absent_days=absent_days,
            status_labels=[row[0] for row in status_rows],
            status_values=[int(row[1]) for row in status_rows],
            overtime_labels=[row[0] for row in overtime_rows],
            overtime_values=[float(row[1]) for row in overtime_rows],
            warning_rows=warning_rows,
        )

    @app.route("/employees", methods=["GET", "POST"])
    def employees():
        shift_codes = [row.code for row in ShiftTemplate.query.order_by(ShiftTemplate.code.asc()).all()]

        if request.method == "POST":
            actor = request.form.get("changed_by", "admin")
            employee_code = request.form.get("employee_code", "").strip()
            full_name = request.form.get("full_name", "").strip()
            gender = request.form.get("gender", "").strip() or None
            hometown = request.form.get("hometown", "").strip() or None
            birth_year = request.form.get("birth_year", "").strip()
            default_shift_code = request.form.get("default_shift_code", "X").strip().upper()

            if not employee_code or not full_name:
                flash("Can nhap Ma NV va Ho ten", "error")
                return redirect(url_for("employees"))

            if default_shift_code not in shift_codes:
                flash("Ma ca mac dinh khong hop le", "error")
                return redirect(url_for("employees"))

            existing = Employee.query.filter_by(employee_code=employee_code).first()
            if existing:
                flash("Ma nhan vien da ton tai", "error")
                return redirect(url_for("employees"))

            employee = Employee(
                employee_code=employee_code,
                full_name=full_name,
                gender=gender,
                hometown=hometown,
                birth_year=int(birth_year) if birth_year else None,
                default_shift_code=default_shift_code,
            )
            db.session.add(employee)
            db.session.flush()

            log_action(
                "employees",
                employee.id,
                "INSERT",
                changed_by=actor,
                after_data=employee.to_dict(),
            )
            db.session.commit()
            flash("Da them nhan vien", "success")
            return redirect(url_for("employees"))

        rows = Employee.query.order_by(Employee.employee_code.asc()).all()
        return render_template("employees.html", employees=rows, shift_codes=shift_codes)

    @app.route("/employees/<int:employee_id>")
    def employee_detail(employee_id):
        employee = Employee.query.get_or_404(employee_id)
        month_key = _safe_month_key(request.args.get("month"))

        details = build_live_month_details(month_key, employee_id=employee_id)
        salaries = (
            MonthlySalary.query.filter_by(employee_id=employee_id)
            .order_by(MonthlySalary.month_key.desc())
            .all()
        )
        balances = (
            LeaveBalance.query.filter_by(employee_id=employee_id)
            .order_by(LeaveBalance.year.desc())
            .all()
        )

        month_keys = [row.month_key for row in salaries]
        config_rows = (
            MonthlyWorkdayConfig.query.filter(MonthlyWorkdayConfig.month_key.in_(month_keys)).all()
            if month_keys
            else []
        )
        config_map = {row.month_key: _to_float(row.company_work_days, 0) for row in config_rows}

        salary_history_rows = []
        for row in salaries:
            company_work_days = config_map.get(row.month_key, 0)
            if company_work_days <= 0:
                legacy_value = _to_float(row.salary_coefficient, 0)
                company_work_days = legacy_value if legacy_value >= 10 else 26.0

            monthly_wage = _to_float(row.base_daily_wage)
            daily_rate = (monthly_wage / company_work_days) if company_work_days > 0 else 0.0
            salary_history_rows.append(
                {
                    "salary": row,
                    "monthly_wage": round(monthly_wage, 2),
                    "company_work_days": round(company_work_days, 2),
                    "daily_rate": round(daily_rate, 2),
                }
            )

        summary_paid_hours = sum(float(row.paid_hours or 0) for row in details)
        summary_wage = sum(float(row.daily_wage or 0) for row in details)

        return render_template(
            "employee_detail.html",
            employee=employee,
            month_key=month_key,
            details=details,
            salary_history_rows=salary_history_rows,
            balances=balances,
            summary_paid_hours=summary_paid_hours,
            summary_wage=summary_wage,
        )

    @app.route("/shifts", methods=["GET", "POST"])
    def shifts():
        if request.method == "POST":
            actor = request.form.get("changed_by", "admin").strip() or "admin"
            code = request.form.get("code", "").strip().upper()
            name = request.form.get("name", "").strip()
            payload = {
                "name": name,
                "start_time": _parse_time(request.form.get("start_time", "")),
                "end_time": _parse_time(request.form.get("end_time", "")),
                "break_minutes": int(request.form.get("break_minutes", 0) or 0),
                "standard_hours": _to_float(request.form.get("standard_hours"), 0),
                "default_overtime_hours": _to_float(
                    request.form.get("default_overtime_hours"), 0
                ),
                "meal_allowance": _to_float(request.form.get("meal_allowance"), 0),
                "is_leave_code": request.form.get("is_leave_code") == "on",
                "is_paid_leave": request.form.get("is_paid_leave") == "on",
                "notes": request.form.get("notes", "").strip() or None,
            }

            if not code or not name:
                flash("Can nhap ma ca va ten ca", "error")
                return redirect(url_for("shifts"))

            existing = ShiftTemplate.query.filter_by(code=code).first()
            if existing:
                flash("Ma ca da ton tai, vui long bam Sua tren dong can cap nhat", "error")
                return redirect(url_for("edit_shift", shift_id=existing.id))

            shift = ShiftTemplate(code=code, **payload)
            db.session.add(shift)
            db.session.flush()
            log_action(
                "shift_templates",
                shift.id,
                "INSERT",
                changed_by=actor,
                after_data=shift.to_dict(),
            )

            db.session.commit()
            flash("Da tao ca moi", "success")
            return redirect(url_for("shifts"))

        rows = ShiftTemplate.query.order_by(ShiftTemplate.code.asc()).all()
        return render_template("shifts.html", shifts=rows)

    @app.route("/shifts/<int:shift_id>/edit", methods=["GET", "POST"])
    def edit_shift(shift_id):
        shift = ShiftTemplate.query.get_or_404(shift_id)

        if request.method == "POST":
            actor = request.form.get("changed_by", "admin").strip() or "admin"
            name = request.form.get("name", "").strip()

            if not name:
                flash("Can nhap ten ca", "error")
                return redirect(url_for("edit_shift", shift_id=shift.id))

            payload = {
                "name": name,
                "start_time": _parse_time(request.form.get("start_time", "")),
                "end_time": _parse_time(request.form.get("end_time", "")),
                "break_minutes": int(request.form.get("break_minutes", 0) or 0),
                "standard_hours": _to_float(request.form.get("standard_hours"), 0),
                "default_overtime_hours": _to_float(
                    request.form.get("default_overtime_hours"), 0
                ),
                "meal_allowance": _to_float(request.form.get("meal_allowance"), 0),
                "is_leave_code": request.form.get("is_leave_code") == "on",
                "is_paid_leave": request.form.get("is_paid_leave") == "on",
                "notes": request.form.get("notes", "").strip() or None,
            }

            before = shift.to_dict()
            for key, value in payload.items():
                setattr(shift, key, value)

            log_action(
                "shift_templates",
                shift.id,
                "UPDATE",
                changed_by=actor,
                before_data=before,
                after_data=shift.to_dict(),
            )
            db.session.commit()
            flash("Da cap nhat ca", "success")
            return redirect(url_for("edit_shift", shift_id=shift.id))

        return render_template("shift_edit.html", shift=shift)

    @app.route("/shifts/delete/<int:shift_id>", methods=["POST"])
    def delete_shift(shift_id):
        actor = request.form.get("changed_by", "admin").strip() or "admin"

        shift = ShiftTemplate.query.get_or_404(shift_id)
        schedule_count = WorkSchedule.query.filter_by(shift_id=shift.id).count()
        employee_count = Employee.query.filter_by(default_shift_code=shift.code).count()

        if schedule_count > 0 or employee_count > 0:
            flash(
                "Khong the xoa ca nay vi dang duoc su dung "
                f"({schedule_count} lich lam, {employee_count} nhan vien mac dinh).",
                "error",
            )
            return redirect(url_for("edit_shift", shift_id=shift.id))

        before = shift.to_dict()
        db.session.delete(shift)
        log_action(
            "shift_templates",
            shift.id,
            "DELETE",
            changed_by=actor,
            before_data=before,
            notes="Xoa ma ca lam",
        )
        db.session.commit()
        flash("Da xoa ma ca", "success")
        return redirect(url_for("shifts"))

    @app.route("/salaries", methods=["GET", "POST"])
    def salaries():
        month_key = _safe_month_key(request.args.get("month"))

        if request.method == "POST":
            actor = request.form.get("changed_by", "admin").strip() or "admin"
            action = request.form.get("action", "save_employee_salary").strip()
            target_month = _safe_month_key(request.form.get("month_key") or month_key)

            if action == "save_month_workdays":
                company_work_days = _to_float(request.form.get("company_work_days"), 0)
                notes = request.form.get("notes", "").strip() or None

                if company_work_days <= 0:
                    flash("Cong chuan thang phai lon hon 0", "error")
                    return redirect(url_for("salaries", month=target_month))

                config = MonthlyWorkdayConfig.query.filter_by(month_key=target_month).first()
                if config:
                    before = config.to_dict()
                    config.company_work_days = company_work_days
                    config.notes = notes
                    log_action(
                        "monthly_workday_configs",
                        config.id,
                        "UPDATE",
                        changed_by=actor,
                        before_data=before,
                        after_data=config.to_dict(),
                    )
                else:
                    config = MonthlyWorkdayConfig(
                        month_key=target_month,
                        company_work_days=company_work_days,
                        notes=notes,
                    )
                    db.session.add(config)
                    db.session.flush()
                    log_action(
                        "monthly_workday_configs",
                        config.id,
                        "INSERT",
                        changed_by=actor,
                        after_data=config.to_dict(),
                    )

                month_salary_rows = MonthlySalary.query.filter_by(month_key=target_month).all()
                for salary_row in month_salary_rows:
                    salary_row.salary_coefficient = company_work_days

                db.session.commit()
                rebuild_month_details(target_month, actor)
                flash("Da luu cong chuan thang (ap dung toan bo nhan vien)", "success")
                return redirect(url_for("salaries", month=target_month))

            employee_id_raw = request.form.get("employee_id", "").strip()
            if not employee_id_raw.isdigit():
                flash("Can chon nhan vien", "error")
                return redirect(url_for("salaries", month=target_month))

            employee_id = int(employee_id_raw)
            base_monthly_wage = _to_float(
                request.form.get("base_monthly_wage") or request.form.get("base_daily_wage"),
                0,
            )
            pay_method = request.form.get("pay_method", "").strip() or None

            if base_monthly_wage < 0:
                flash("Luong thang khong hop le", "error")
                return redirect(url_for("salaries", month=target_month))

            company_work_days, _ = _resolve_company_work_days(target_month)

            row = MonthlySalary.query.filter_by(employee_id=employee_id, month_key=target_month).first()
            if row:
                before = row.to_dict()
                row.base_daily_wage = base_monthly_wage
                row.salary_coefficient = company_work_days
                row.pay_method = pay_method
                log_action(
                    "monthly_salaries",
                    row.id,
                    "UPDATE",
                    changed_by=actor,
                    before_data=before,
                    after_data=row.to_dict(),
                    notes="Luong thang theo nhan vien",
                )
            else:
                row = MonthlySalary(
                    employee_id=employee_id,
                    month_key=target_month,
                    base_daily_wage=base_monthly_wage,
                    salary_coefficient=company_work_days,
                    pay_method=pay_method,
                )
                db.session.add(row)
                db.session.flush()
                log_action(
                    "monthly_salaries",
                    row.id,
                    "INSERT",
                    changed_by=actor,
                    after_data=row.to_dict(),
                    notes="Luong thang theo nhan vien",
                )

            db.session.commit()
            rebuild_month_details(target_month, actor)
            flash("Da luu luong thang nhan vien", "success")
            return redirect(url_for("salaries", month=target_month))

        employees = Employee.query.order_by(Employee.employee_code.asc()).all()
        company_work_days, workday_config = _resolve_company_work_days(month_key)
        rows = (
            MonthlySalary.query.options(joinedload(MonthlySalary.employee))
            .filter(MonthlySalary.month_key == month_key)
            .order_by(MonthlySalary.employee_id.asc())
            .all()
        )

        salary_rows = []
        for row in rows:
            monthly_wage = _to_float(row.base_daily_wage)
            daily_rate = (monthly_wage / company_work_days) if company_work_days > 0 else 0.0
            salary_rows.append(
                {
                    "salary": row,
                    "monthly_wage": round(monthly_wage, 2),
                    "daily_rate": round(daily_rate, 2),
                }
            )

        return render_template(
            "salaries.html",
            month_key=month_key,
            employees=employees,
            salary_rows=salary_rows,
            company_work_days=round(company_work_days, 2),
            workday_config=workday_config,
        )

    @app.route("/salaries/import", methods=["POST"])
    def import_salaries():
        actor = request.form.get("changed_by", "admin").strip() or "admin"
        month_key = _safe_month_key(request.form.get("month_key") or request.args.get("month"))
        replace_existing_month = request.form.get("replace_existing_month") == "on"

        upload = request.files.get("salary_file")
        if not upload or upload.filename == "":
            flash("Can chon file he luong de import", "error")
            return redirect(url_for("salaries", month=month_key))

        extension = Path(upload.filename).suffix.lower()
        if extension not in {".xlsx", ".xlsm", ".xltx", ".xltm", ".xls", ".csv"}:
            flash("File he luong chi ho tro CSV/XLSX", "error")
            return redirect(url_for("salaries", month=month_key))

        upload_folder = Path(current_app.config["UPLOAD_FOLDER"])
        upload_folder.mkdir(parents=True, exist_ok=True)

        safe_name = secure_filename(upload.filename)
        temp_name = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{safe_name}"
        temp_path = upload_folder / temp_name
        upload.save(temp_path)

        try:
            default_work_days, _ = _resolve_company_work_days(month_key)
            result = import_salary_file(
                str(temp_path),
                upload.filename,
                actor=actor,
                target_month=month_key,
                default_company_work_days=default_work_days,
                replace_existing=replace_existing_month,
            )

            rebuild_month_details(month_key, actor)

            replaced_info = ""
            if result["replace_existing"]:
                replaced_info = f" Da xoa truoc {result['deleted_rows']} dong luong cu."

            unknown_info = ""
            if result["skipped_unknown"] > 0:
                unknown_preview = ", ".join(result["unknown_codes"][:10])
                unknown_info = (
                    f" Bo qua {result['skipped_unknown']} dong do khong tim thay Ma NV"
                    f" ({unknown_preview})."
                )

            flash(
                f"Import he luong thang {month_key} xong: tao moi {result['created']}, "
                f"cap nhat {result['updated']}, cong chuan {result['company_work_days']}.{replaced_info}{unknown_info}",
                "success",
            )
            return redirect(url_for("salaries", month=month_key))
        except Exception as exc:
            db.session.rollback()
            flash(f"Import he luong that bai: {exc}", "error")
            return redirect(url_for("salaries", month=month_key))
        finally:
            if temp_path.exists():
                os.remove(temp_path)

    @app.route("/holidays", methods=["GET", "POST"])
    def holidays():
        month_key = _safe_month_key(
            request.args.get("month") or request.form.get("month_key") or request.form.get("month")
        )
        start_date, end_date = parse_month_key(month_key)

        if request.method == "POST":
            actor = request.form.get("changed_by", "admin").strip() or "admin"
            action = request.form.get("action", "save_single").strip()

            if action == "generate_month":
                created_count = 0
                updated_count = 0
                sunday_total = 0
                sunday_created_count = 0
                vn_created_count = 0
                vn_holidays = _get_vietnam_holiday_map(start_date, end_date)

                current_day = start_date
                while current_day <= end_date:
                    if current_day.weekday() == 6:
                        sunday_total += 1
                        row = Holiday.query.filter_by(holiday_date=current_day).first()
                        if not row:
                            row = Holiday(
                                holiday_date=current_day,
                                name="Chu nhat",
                                is_paid=True,
                                notes="Tao tu dong theo thang",
                            )
                            db.session.add(row)
                            db.session.flush()
                            log_action(
                                "holidays",
                                row.id,
                                "INSERT",
                                changed_by=actor,
                                after_data=row.to_dict(),
                                notes="Tao nhanh chu nhat OFF theo thang",
                            )
                            created_count += 1
                            sunday_created_count += 1
                    current_day += timedelta(days=1)

                for holiday_date, holiday_name in sorted(vn_holidays.items()):
                    row = Holiday.query.filter_by(holiday_date=holiday_date).first()
                    if row:
                        before = row.to_dict()
                        changed = False

                        current_name = (row.name or "").strip()
                        if holiday_name and holiday_name.lower() not in current_name.lower():
                            row.name = f"{current_name} + {holiday_name}" if current_name else holiday_name
                            changed = True

                        if changed:
                            log_action(
                                "holidays",
                                row.id,
                                "UPDATE",
                                changed_by=actor,
                                before_data=before,
                                after_data=row.to_dict(),
                                notes="Bo sung ten ngay le Viet Nam tu dong",
                            )
                            updated_count += 1
                    else:
                        row = Holiday(
                            holiday_date=holiday_date,
                            name=holiday_name,
                            is_paid=True,
                            notes="Ngay le Viet Nam (tu dong theo thang)",
                        )
                        db.session.add(row)
                        db.session.flush()
                        log_action(
                            "holidays",
                            row.id,
                            "INSERT",
                            changed_by=actor,
                            after_data=row.to_dict(),
                            notes="Tao tu dong ngay le Viet Nam",
                        )
                        created_count += 1
                        vn_created_count += 1

                db.session.commit()
                rebuild_month_details(month_key, actor)
                library_note = ""
                if not _get_holiday_library():
                    library_note = " (Dang dung fallback ngay le co dinh do chua cai goi holidays)"

                flash(
                    f"Da tao moi {created_count} ngay OFF/le (Chu nhat moi: {sunday_created_count}/{sunday_total}, Le VN moi: {vn_created_count}), cap nhat {updated_count} ngay cho thang {month_key}{library_note}",
                    "success",
                )
                return redirect(url_for("holidays", month=month_key))

            if action == "update_row":
                row_id_raw = (request.form.get("holiday_id") or "").strip()
                if not row_id_raw.isdigit():
                    flash("Khong tim thay dong ngay OFF/le de cap nhat", "error")
                    return redirect(url_for("holidays", month=month_key))

                row = Holiday.query.get(int(row_id_raw))
                if not row:
                    flash("Dong ngay OFF/le khong ton tai", "error")
                    return redirect(url_for("holidays", month=month_key))

                name = request.form.get("name", "").strip()
                notes = request.form.get("notes", "").strip() or None
                is_paid = request.form.get("is_paid") == "on"

                if not name:
                    flash("Can nhap ten ngay OFF/le", "error")
                    return redirect(url_for("holidays", month=month_key_for_date(row.holiday_date)))

                before = row.to_dict()
                row.name = name
                row.is_paid = is_paid
                row.notes = notes

                log_action(
                    "holidays",
                    row.id,
                    "UPDATE",
                    changed_by=actor,
                    before_data=before,
                    after_data=row.to_dict(),
                )

                db.session.commit()
                rebuild_month_details(month_key_for_date(row.holiday_date), actor)
                flash("Da cap nhat tick nghi/ngay le", "success")
                return redirect(url_for("holidays", month=month_key_for_date(row.holiday_date)))

            try:
                holiday_date = _parse_date(request.form.get("holiday_date", "").strip())
            except (TypeError, ValueError):
                flash("Ngay OFF khong hop le", "error")
                return redirect(url_for("holidays", month=month_key))

            name = request.form.get("name", "").strip()
            is_paid = request.form.get("is_paid") == "on"
            notes = request.form.get("notes", "").strip() or None

            if not name:
                flash("Can nhap ten ngay OFF/le", "error")
                return redirect(url_for("holidays", month=month_key_for_date(holiday_date)))

            row = Holiday.query.filter_by(holiday_date=holiday_date).first()
            if row:
                before = row.to_dict()
                row.name = name
                row.is_paid = is_paid
                row.notes = notes
                log_action(
                    "holidays",
                    row.id,
                    "UPDATE",
                    changed_by=actor,
                    before_data=before,
                    after_data=row.to_dict(),
                )
            else:
                row = Holiday(holiday_date=holiday_date, name=name, is_paid=is_paid, notes=notes)
                db.session.add(row)
                db.session.flush()
                log_action(
                    "holidays",
                    row.id,
                    "INSERT",
                    changed_by=actor,
                    after_data=row.to_dict(),
                )

            db.session.commit()
            rebuild_month_details(month_key_for_date(holiday_date), actor)
            flash("Da luu ngay OFF/le", "success")
            return redirect(url_for("holidays", month=month_key_for_date(holiday_date)))

        rows = (
            Holiday.query.filter(
                Holiday.holiday_date >= start_date,
                Holiday.holiday_date <= end_date,
            )
            .order_by(Holiday.holiday_date.asc())
            .all()
        )
        return render_template("holidays.html", holidays=rows, month_key=month_key)

    @app.route("/schedules", methods=["GET", "POST"])
    def schedules():
        month_key = _safe_month_key(request.args.get("month"))

        if request.method == "POST":
            actor = request.form.get("changed_by", "admin")
            employee_id = int(request.form.get("employee_id"))
            work_date = _parse_date(request.form.get("work_date"))
            shift_code = request.form.get("shift_code", "").strip().upper()
            absence_hours = _to_float(request.form.get("absence_hours"), 0)
            overtime_hours_raw = request.form.get("overtime_hours", "").strip()
            overtime_reason = request.form.get("overtime_reason", "").strip() or None
            notes = request.form.get("notes", "").strip() or None

            shift = ShiftTemplate.query.filter_by(code=shift_code).first()
            if not shift:
                flash("Ma ca khong hop le", "error")
                return redirect(url_for("schedules", month=month_key_for_date(work_date)))

            row = WorkSchedule.query.filter_by(employee_id=employee_id, work_date=work_date).first()
            target_month = month_key_for_date(work_date)

            if row:
                before = row.to_dict()
                row.shift_id = shift.id
                row.month_key = target_month
                row.absence_hours = absence_hours
                row.notes = notes
                action = "UPDATE"
            else:
                row = WorkSchedule(
                    employee_id=employee_id,
                    work_date=work_date,
                    month_key=target_month,
                    shift_id=shift.id,
                    absence_hours=absence_hours,
                    notes=notes,
                )
                db.session.add(row)
                db.session.flush()
                before = None
                action = "INSERT"

            overtime_hours = (
                _to_float(overtime_hours_raw)
                if overtime_hours_raw
                else float(shift.default_overtime_hours or 0)
            )

            if overtime_hours > 0 or overtime_reason:
                if row.overtime:
                    row.overtime.hours = overtime_hours
                    row.overtime.reason = overtime_reason
                else:
                    overtime = OvertimeEntry(
                        schedule_id=row.id,
                        hours=overtime_hours,
                        reason=overtime_reason,
                    )
                    db.session.add(overtime)

            log_action(
                "work_schedules",
                row.id,
                action,
                changed_by=actor,
                before_data=before,
                after_data=row.to_dict(),
                notes="Lich lam + tang ca",
            )

            db.session.commit()
            rebuild_month_details(target_month, actor)
            flash("Da luu lich lam", "success")
            return redirect(url_for("schedules", month=target_month))

        employees = Employee.query.order_by(Employee.employee_code.asc()).all()
        shift_codes = ShiftTemplate.query.order_by(ShiftTemplate.code.asc()).all()

        rows = (
            WorkSchedule.query.options(
                joinedload(WorkSchedule.employee),
                joinedload(WorkSchedule.shift),
                joinedload(WorkSchedule.overtime),
            )
            .filter(WorkSchedule.month_key == month_key)
            .order_by(WorkSchedule.work_date.asc(), WorkSchedule.employee_id.asc())
            .all()
        )

        return render_template(
            "schedules.html",
            month_key=month_key,
            employees=employees,
            schedules=rows,
            shift_codes=shift_codes,
            valid_shift_codes=[row.code for row in shift_codes],
        )

    @app.route("/schedules/import", methods=["POST"])
    def import_schedules():
        actor = request.form.get("changed_by", "admin")
        month_key_input = request.form.get("month_key", "").strip()
        month_key = _safe_month_key(month_key_input) if month_key_input else None
        replace_existing_month = request.form.get("replace_existing_month") == "on"
        open_details_after_import = request.form.get("open_details_after_import") == "on"

        upload = request.files.get("schedule_file")
        if not upload or upload.filename == "":
            flash("Can chon file lich lam .xlsx", "error")
            return redirect(url_for("schedules", month=month_key or current_month_key()))

        extension = Path(upload.filename).suffix.lower()
        if extension not in {".xlsx", ".xlsm", ".xltx", ".xltm", ".xls"}:
            flash("File lich lam chi ho tro dinh dang Excel", "error")
            return redirect(url_for("schedules", month=month_key or current_month_key()))

        upload_folder = Path(current_app.config["UPLOAD_FOLDER"])
        upload_folder.mkdir(parents=True, exist_ok=True)

        safe_name = secure_filename(upload.filename)
        temp_name = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{safe_name}"
        temp_path = upload_folder / temp_name
        upload.save(temp_path)

        try:
            result = import_schedule_file(
                str(temp_path),
                upload.filename,
                actor=actor,
                target_month=month_key,
                replace_existing=replace_existing_month,
            )

            rebuilt = {}
            for item in result["months"]:
                rebuilt[item] = rebuild_month_details(item, actor)

            replaced_info = ""
            if result["replace_existing"] and result["replaced_months"]:
                replaced_info = (
                    f" Da xoa lich cu truoc khi import: {result['replaced_months']}."
                )

            flash(
                f"Import lich xong {result['rows_imported']} dong ca "
                f"(tao moi {result['created']}, cap nhat {result['updated']}). "
                f"Da tai tao chi tiet: {rebuilt}.{replaced_info}",
                "success",
            )

            redirect_month = month_key or (result["months"][0] if result["months"] else current_month_key())
            if open_details_after_import:
                return redirect(url_for("details", month=redirect_month))
            return redirect(url_for("schedules", month=redirect_month))
        except Exception as exc:
            db.session.rollback()
            flash(f"Import lich that bai: {exc}", "error")
            return redirect(url_for("schedules", month=month_key or current_month_key()))
        finally:
            if temp_path.exists():
                os.remove(temp_path)

    @app.route("/imports", methods=["GET", "POST"])
    def imports():
        if request.method == "POST":
            actor = request.form.get("changed_by", "admin")
            month_key_input = request.form.get("month_key", "").strip()
            month_key = _safe_month_key(month_key_input) if month_key_input else None
            replace_existing_month = request.form.get("replace_existing_month") == "on"

            upload = request.files.get("attendance_file")
            if not upload or upload.filename == "":
                flash("Can chon file CSV/XLSX", "error")
                return redirect(url_for("imports"))

            upload_folder = Path(current_app.config["UPLOAD_FOLDER"])
            upload_folder.mkdir(parents=True, exist_ok=True)

            safe_name = secure_filename(upload.filename)
            temp_name = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{safe_name}"
            temp_path = upload_folder / temp_name
            upload.save(temp_path)

            try:
                result = import_attendance_file(
                    str(temp_path),
                    upload.filename,
                    actor,
                    month_key=month_key,
                    replace_existing=replace_existing_month,
                )
                months = result["months"]
                rebuilt = {}
                for item in months:
                    rebuilt[item] = rebuild_month_details(item, actor)

                replaced_info = ""
                if result["replace_existing"] and result["replaced_months"]:
                    replaced_info = (
                        f" Da xoa du lieu cu truoc khi import: {result['replaced_months']}."
                    )

                flash(
                    f"Import xong {result['rows']} dong, tao {result['grouped_days']} ban ghi ngay. "
                    f"Da tai tao chi tiet: {rebuilt}.{replaced_info}",
                    "success",
                )
            except Exception as exc:
                db.session.rollback()
                flash(f"Import that bai: {exc}", "error")
            finally:
                if temp_path.exists():
                    os.remove(temp_path)

            return redirect(url_for("imports"))

        import_logs = (
            AuditLog.query.filter_by(table_name="attendance_import")
            .order_by(AuditLog.changed_at.desc())
            .limit(30)
            .all()
        )
        return render_template("imports.html", import_logs=import_logs)

    @app.route("/imports/delete/<batch_id>", methods=["POST"])
    def delete_import_batch(batch_id):
        actor = request.form.get("changed_by", "admin")

        log_rows = AttendanceLog.query.filter_by(import_batch=batch_id).all()
        if not log_rows:
            flash("Khong tim thay batch import de xoa", "error")
            return redirect(url_for("imports"))

        affected_months = sorted(
            {month_key_for_date(row.event_time.date()) for row in log_rows}
        )

        removed_logs = AttendanceLog.query.filter_by(import_batch=batch_id).delete(
            synchronize_session=False
        )
        removed_daily = AttendanceDaily.query.filter_by(import_batch=batch_id).delete(
            synchronize_session=False
        )

        log_action(
            "attendance_import",
            batch_id,
            "DELETE_BATCH",
            changed_by=actor,
            after_data={
                "removed_logs": int(removed_logs),
                "removed_daily": int(removed_daily),
                "affected_months": affected_months,
            },
            notes="Xoa du lieu cua mot lan import",
        )
        db.session.commit()

        rebuilt = {}
        for item in affected_months:
            rebuilt[item] = rebuild_month_details(item, actor)

        flash(
            f"Da xoa batch {batch_id}. Logs: {removed_logs}, Daily: {removed_daily}, Tai tao: {rebuilt}",
            "success",
        )
        return redirect(url_for("imports"))

    @app.route("/details")
    def details():
        month_key = _safe_month_key(request.args.get("month"))
        employee_id = request.args.get("employee_id", "").strip()
        selected_employee_id = int(employee_id) if employee_id else None
        start_date_raw = request.args.get("start_date", "").strip()
        end_date_raw = request.args.get("end_date", "").strip()

        parsed_start_date = None
        parsed_end_date = None
        has_date_parse_error = False

        if start_date_raw:
            try:
                parsed_start_date = _parse_date(start_date_raw)
            except (TypeError, ValueError):
                flash("Tu ngay khong hop le", "error")
                has_date_parse_error = True

        if end_date_raw:
            try:
                parsed_end_date = _parse_date(end_date_raw)
            except (TypeError, ValueError):
                flash("Den ngay khong hop le", "error")

                has_date_parse_error = True

        if has_date_parse_error:
            parsed_start_date = None
            parsed_end_date = None

        if parsed_start_date and not parsed_end_date:
            parsed_end_date = parsed_start_date
        if parsed_end_date and not parsed_start_date:
            parsed_start_date = parsed_end_date

        is_range_mode = bool(parsed_start_date and parsed_end_date)

        if is_range_mode and parsed_start_date > parsed_end_date:
            parsed_start_date, parsed_end_date = parsed_end_date, parsed_start_date

        if is_range_mode:
            rows = []
            for item_month in _iter_month_keys(parsed_start_date, parsed_end_date):
                month_rows = build_live_month_details(item_month, employee_id=selected_employee_id)
                rows.extend(
                    row for row in month_rows if parsed_start_date <= row.work_date <= parsed_end_date
                )

            def _sort_key(detail_row):
                raw_code = (detail_row.employee.employee_code or "").replace("'", "").strip()
                if raw_code.isdigit():
                    code_key = (0, int(raw_code))
                else:
                    code_key = (1, raw_code.lower())

                return code_key, detail_row.work_date, detail_row.employee.id

            rows.sort(key=_sort_key)
            period_label = f"{parsed_start_date} den {parsed_end_date}"
        else:
            rows = build_live_month_details(month_key, employee_id=selected_employee_id)
            period_label = month_key

        employees = Employee.query.order_by(Employee.employee_code.asc()).all()
        return render_template(
            "details.html",
            month_key=month_key,
            details=rows,
            employees=employees,
            selected_employee_id=selected_employee_id,
            start_date_value=parsed_start_date.isoformat() if parsed_start_date else "",
            end_date_value=parsed_end_date.isoformat() if parsed_end_date else "",
            is_range_mode=is_range_mode,
            period_label=period_label,
        )

    @app.route("/audit")
    def audit_logs():
        rows = AuditLog.query.order_by(AuditLog.changed_at.desc()).limit(300).all()
        return render_template("audit.html", audit_logs=rows)

    @app.route("/audit/export")
    def export_audit_csv():
        rows = AuditLog.query.order_by(AuditLog.changed_at.desc()).limit(5000).all()

        stream = io.StringIO()
        writer = csv.writer(stream)
        writer.writerow(
            [
                "ID",
                "Thoi gian",
                "Nguoi sua",
                "Bang",
                "Action",
                "Record",
                "Before",
                "After",
                "Ghi chu",
            ]
        )

        for row in rows:
            writer.writerow(
                [
                    row.id,
                    row.changed_at,
                    row.changed_by,
                    row.table_name,
                    row.action,
                    row.record_id,
                    row.before_data,
                    row.after_data,
                    row.notes,
                ]
            )

        output = stream.getvalue()
        stream.close()

        return Response(
            output,
            mimetype="text/csv",
            headers={
                "Content-Disposition": "attachment; filename=audit_logs.csv",
            },
        )

    @app.route("/backup/run", methods=["POST"])
    def run_backup_now():
        actor = request.form.get("changed_by", "admin")
        try:
            backup_file, removed = run_pg_dump(
                current_app.config["SQLALCHEMY_DATABASE_URI"],
                current_app.config["BACKUP_TARGET_DIR"],
                current_app.config["BACKUP_RETENTION_DAYS"],
            )
            log_action(
                "system_backup",
                backup_file,
                "BACKUP",
                changed_by=actor,
                after_data={"backup_file": backup_file, "removed_files": removed},
                notes="Backup thu cong",
            )
            db.session.commit()
            flash(f"Backup thanh cong: {backup_file}", "success")
        except Exception as exc:
            db.session.rollback()
            flash(f"Backup that bai: {exc}", "error")

        return redirect(url_for("audit_logs"))
