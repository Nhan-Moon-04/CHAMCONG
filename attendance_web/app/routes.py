import csv
import io
import os
from datetime import datetime
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
    OvertimeEntry,
    ShiftTemplate,
    WorkSchedule,
)
from .services.attendance import (
    build_live_month_details,
    current_month_key,
    month_key_for_date,
    rebuild_month_details,
)
from .services.audit import log_action
from .services.backup import run_pg_dump
from .services.importer import import_attendance_file
from .services.schedule_importer import import_schedule_file


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

        summary_paid_hours = sum(float(row.paid_hours or 0) for row in details)
        summary_wage = sum(float(row.daily_wage or 0) for row in details)

        return render_template(
            "employee_detail.html",
            employee=employee,
            month_key=month_key,
            details=details,
            salaries=salaries,
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
            actor = request.form.get("changed_by", "admin")
            employee_id = int(request.form.get("employee_id"))
            target_month = _safe_month_key(request.form.get("month_key"))
            base_daily_wage = _to_float(request.form.get("base_daily_wage"), 0)
            salary_coefficient = _to_float(request.form.get("salary_coefficient"), 1)
            pay_method = request.form.get("pay_method", "").strip() or None

            row = MonthlySalary.query.filter_by(employee_id=employee_id, month_key=target_month).first()
            if row:
                before = row.to_dict()
                row.base_daily_wage = base_daily_wage
                row.salary_coefficient = salary_coefficient
                row.pay_method = pay_method
                log_action(
                    "monthly_salaries",
                    row.id,
                    "UPDATE",
                    changed_by=actor,
                    before_data=before,
                    after_data=row.to_dict(),
                )
            else:
                row = MonthlySalary(
                    employee_id=employee_id,
                    month_key=target_month,
                    base_daily_wage=base_daily_wage,
                    salary_coefficient=salary_coefficient,
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
                )

            db.session.commit()
            rebuild_month_details(target_month, actor)
            flash("Da luu bang luong theo thang", "success")
            return redirect(url_for("salaries", month=target_month))

        employees = Employee.query.order_by(Employee.employee_code.asc()).all()
        rows = (
            MonthlySalary.query.options(joinedload(MonthlySalary.employee))
            .filter(MonthlySalary.month_key == month_key)
            .order_by(MonthlySalary.id.desc())
            .all()
        )
        return render_template(
            "salaries.html",
            month_key=month_key,
            employees=employees,
            salaries=rows,
        )

    @app.route("/holidays", methods=["GET", "POST"])
    def holidays():
        if request.method == "POST":
            actor = request.form.get("changed_by", "admin")
            holiday_date = _parse_date(request.form.get("holiday_date"))
            name = request.form.get("name", "").strip()
            is_paid = request.form.get("is_paid") == "on"
            notes = request.form.get("notes", "").strip() or None

            if not name:
                flash("Can nhap ten ngay le", "error")
                return redirect(url_for("holidays"))

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
            flash("Da luu ngay le", "success")
            return redirect(url_for("holidays"))

        rows = Holiday.query.order_by(Holiday.holiday_date.asc()).all()
        return render_template("holidays.html", holidays=rows)

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

        rows = build_live_month_details(month_key, employee_id=selected_employee_id)

        employees = Employee.query.order_by(Employee.employee_code.asc()).all()
        return render_template(
            "details.html",
            month_key=month_key,
            details=rows,
            employees=employees,
            selected_employee_id=selected_employee_id,
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
