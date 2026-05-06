from pathlib import Path
from types import SimpleNamespace

from datetime import datetime

from flask import current_app, flash, redirect, render_template, request, session, url_for
from flask import jsonify
from sqlalchemy import func
from .database import db
from .models import (
    AdvancePayment,
    AttendanceDetail,
    Employee,
    PayrollInsuranceContribution,
    PayrollLeaveSnapshot,
    PayrollSlip,
    PayrollTaxContribution,
    MonthlySalary,
    MonthlyWorkdayConfig,
)
from .services.attendance import current_month_key
from .services.audit import log_action
from .services.salary_importer import import_salary_detail_file


def _safe_month_key(value):
    if not value:
        return current_month_key()
    val = str(value).strip()
    # Accept several common month formats and normalize to YYYY-MM
    for fmt in ("%Y-%m", "%B %Y", "%b %Y", "%m/%Y", "%Y/%m"):
        try:
            dt = datetime.strptime(val, fmt)
            return dt.strftime("%Y-%m")
        except ValueError:
            continue
    return current_month_key()


def _to_float(value, default=0.0):
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def _require_admin():
    if session.get("is_admin"):
        return None
    flash("Ban khong co quyen truy cap chuc nang nay", "error")
    return redirect(url_for("dashboard"))


def _resolve_payroll_source_file():
    configured = (current_app.config.get("PAYROLL_DETAIL_SOURCE_FILE") or "").strip()
    if configured:
        candidate = Path(configured)
        if candidate.exists():
            return str(candidate)

    fallback = Path(current_app.root_path).parent.parent / "4 月份文房薪資 2026 (1).xls"
    if fallback.exists():
        return str(fallback)

    return None


def _ensure_payroll_month_data(month_key):
    existing_any = (
        PayrollLeaveSnapshot.query.filter_by(month_key=month_key).first()
        or PayrollSlip.query.filter_by(month_key=month_key).first()
        or PayrollInsuranceContribution.query.filter_by(month_key=month_key).first()
        or PayrollTaxContribution.query.filter_by(month_key=month_key).first()
    )
    if existing_any:
        return

    source_file = _resolve_payroll_source_file()
    if not source_file:
        return

    try:
        import_salary_detail_file(
            source_file,
            source_name=Path(source_file).name,
            actor="system",
            target_month=month_key,
            replace_existing=True,
        )
    except Exception as exc:
        current_app.logger.warning("Khong the nap du lieu payroll tu Excel: %s", exc)


def _month_has_salary_data(month_key):
    return bool(
        db.session.query(MonthlySalary.id)
        .filter(MonthlySalary.month_key == month_key)
        .first()
    )


def _build_slip_rows(month_key, search_query=""):
    employees = Employee.query.filter(Employee.is_active.is_(True)).order_by(Employee.employee_code.asc()).all()
    if not employees:
        return [], 0.0, 0.0

    summary_map = {
        employee.id: {
            "worked_days": 0.0,
            "paid_leave_days": 0.0,
            "unpaid_leave_days": 0.0,
            "overtime_hours": 0.0,
        }
        for employee in employees
    }

    detail_rows = (
        AttendanceDetail.query.filter(AttendanceDetail.month_key == month_key)
        .order_by(AttendanceDetail.employee_id.asc(), AttendanceDetail.work_date.asc())
        .all()
    )

    def _apply_status(summary_obj, detail_row):
        normalized = str(getattr(detail_row, "status_code", "") or "").upper()
        if normalized == "P":
            summary_obj["paid_leave_days"] += 1.0
        elif normalized in {"S", "C"}:
            summary_obj["paid_leave_days"] += 0.5
            summary_obj["worked_days"] += 0.5
        elif normalized == "N":
            summary_obj["unpaid_leave_days"] += 1.0
        elif normalized == "OFF":
            if str(getattr(detail_row, "paid_hours", 0) or 0) != "0":
                summary_obj["worked_days"] += 1.0
            return
        else:
            summary_obj["worked_days"] += 1.0

    for detail_row in detail_rows:
        summary = summary_map.get(detail_row.employee_id)
        if not summary:
            continue
        _apply_status(summary, detail_row)
        summary["overtime_hours"] += float(detail_row.overtime_hours or 0)

    salary_rows = MonthlySalary.query.filter(
        MonthlySalary.month_key == month_key,
        MonthlySalary.employee_id.in_([employee.id for employee in employees]),
    ).all()
    salary_by_employee = {row.employee_id: row for row in salary_rows}

    workday_config = MonthlyWorkdayConfig.query.filter_by(month_key=month_key).first()
    company_work_days = _to_float(workday_config.company_work_days if workday_config else None, 0)

    if company_work_days <= 0:
        for salary_row in salary_rows:
            if _to_float(salary_row.salary_coefficient, 0) >= 10:
                company_work_days = _to_float(salary_row.salary_coefficient, 0)
                break
    if company_work_days <= 0:
        company_work_days = 26.0

    advances = (
        db.session.query(AdvancePayment.employee_id, func.coalesce(func.sum(AdvancePayment.amount), 0))
        .filter(AdvancePayment.month_key == month_key, AdvancePayment.employee_id.in_([employee.id for employee in employees]))
        .group_by(AdvancePayment.employee_id)
        .all()
    )
    advance_by_employee = {employee_id: _to_float(total_amount) for employee_id, total_amount in advances}

    insurance_rows = PayrollInsuranceContribution.query.filter_by(month_key=month_key).all()
    insurance_by_employee = {row.employee_id: row for row in insurance_rows}

    tax_rows = PayrollTaxContribution.query.filter_by(month_key=month_key).all()
    tax_by_employee = {row.employee_id: row for row in tax_rows}

    slip_rows = PayrollSlip.query.filter_by(month_key=month_key).all()
    slip_by_employee = {row.employee_id: row for row in slip_rows}

    rows = []
    for employee in employees:
        summary = summary_map.get(employee.id, {})
        salary_row = salary_by_employee.get(employee.id)
        monthly_wage = _to_float(salary_row.base_daily_wage if salary_row else None, 0)

        overtime_hours = round(_to_float(summary.get("overtime_hours", 0)), 2)
        salary_day_units = (
            _to_float(summary.get("worked_days", 0), 0)
            + _to_float(summary.get("paid_leave_days", 0), 0)
            + (overtime_hours / 8.0)
        )
        daily_rate = (monthly_wage / company_work_days) if company_work_days > 0 else 0.0
        computed_salary_amount = round(daily_rate * salary_day_units, 2)

        imported_slip = slip_by_employee.get(employee.id)
        insurance = insurance_by_employee.get(employee.id)
        tax = tax_by_employee.get(employee.id)
        advance_amount = _to_float(advance_by_employee.get(employee.id, 0), 0)

        gross_total = _to_float(imported_slip.gross_total if imported_slip else computed_salary_amount, 0)
        # Determine social insurance (employee contribution). Prefer imported slip, then insurance row, else estimate.
        if imported_slip and _to_float(imported_slip.social_insurance_deduction, 0) > 0:
            social_insurance = _to_float(imported_slip.social_insurance_deduction, 0)
        elif insurance and _to_float(insurance.employee_total, 0) > 0:
            social_insurance = _to_float(insurance.employee_total, 0)
        else:
            social_insurance = None
        pit_tax = _to_float(imported_slip.pit_tax_deduction if imported_slip else (tax.pit_tax if tax else 0), 0)
        # If social_insurance still None, estimate using average rates from existing insurance rows
        if social_insurance is None:
            # build average rates from available insurance rows
            total_ratio = 0.0
            ratio_count = 0
            for r in insurance_rows:
                if _to_float(r.insured_salary, 0) > 0 and _to_float(r.employee_total, 0) > 0:
                    total_ratio += float(r.employee_total) / float(r.insured_salary)
                    ratio_count += 1
            default_employee_rate = (total_ratio / ratio_count) if ratio_count > 0 else 0.08
            social_insurance = round(gross_total * default_employee_rate, 2)

        net_income = _to_float(
            imported_slip.net_income
            if imported_slip
            else (gross_total - social_insurance - pit_tax - advance_amount),
            0,
        )

        slip_view = SimpleNamespace(
            attendance_days=round(_to_float(summary.get("worked_days", 0), 0), 2),
            leave_used_days=round(_to_float(summary.get("paid_leave_days", 0), 0), 2),
            leave_remaining_days=0.0,
            salary_by_attendance=round(computed_salary_amount, 2),
            overtime_weekday_hours=0.0,
            overtime_sunday_hours=overtime_hours,
            overtime_pay=0.0,
            role_allowance=_to_float(imported_slip.role_allowance if imported_slip else 0, 0),
            child_allowance=_to_float(imported_slip.child_allowance if imported_slip else 0, 0),
            transport_phone_allowance=_to_float(imported_slip.transport_phone_allowance if imported_slip else 0, 0),
            meal_allowance=_to_float(imported_slip.meal_allowance if imported_slip else 0, 0),
            attendance_allowance=_to_float(imported_slip.attendance_allowance if imported_slip else 0, 0),
            gross_total=gross_total,
            social_insurance_deduction=social_insurance,
            union_fee_deduction=_to_float(imported_slip.union_fee_deduction if imported_slip else 0, 0),
            pit_tax_deduction=pit_tax,
            advance_deduction=advance_amount,
            net_income=net_income,
            payroll_group=imported_slip.payroll_group if imported_slip else (salary_row.pay_method if salary_row else None),
        )

        if not search_query or search_query in (employee.employee_code or "").lower() or search_query in (employee.full_name or "").lower():
            rows.append((slip_view, employee))

    gross_total = sum(float(item[0].gross_total or 0) for item in rows)
    net_total = sum(float(item[0].net_income or 0) for item in rows)
    return rows, round(gross_total, 2), round(net_total, 2)


def _build_insurance_tax_rows(month_key, search_query=""):
    slip_rows, _, _ = _build_slip_rows(month_key, search_query)
    if not slip_rows:
        return [], 0.0, 0.0, 0.0

    tax_rows = PayrollTaxContribution.query.filter_by(month_key=month_key).all()
    tax_by_employee = {row.employee_id: row for row in tax_rows}

    rows = []
    total_employer = 0.0
    total_employee = 0.0
    total_tax = 0.0

    for slip_view, employee in slip_rows:
        insured_salary = _to_float(slip_view.gross_total, 0)
        if insured_salary <= 0:
            continue

        employer_total = round(insured_salary * 0.215, 0)
        employee_total = round(insured_salary * 0.105, 0)
        tax_row = tax_by_employee.get(employee.id)
        pit_tax = _to_float(tax_row.pit_tax if tax_row else 0, 0)

        insurance_view = SimpleNamespace(
            insured_salary=insured_salary,
            employer_total=employer_total,
            employee_total=employee_total,
            union_fund=0,
        )
        tax_view = SimpleNamespace(pit_tax=pit_tax)

        rows.append((insurance_view, tax_view, employee))
        total_employer += employer_total
        total_employee += employee_total
        total_tax += pit_tax

    return rows, round(total_employer, 2), round(total_employee, 2), round(total_tax, 2)


def register_payroll_routes(app):
    @app.route("/payroll/leave", methods=["GET"])
    def payroll_leave():
        month_key = _safe_month_key(request.args.get("month"))
        search_query = (request.args.get("q") or "").strip().lower()
        year = int(month_key[:4])

        # For each employee, pick the latest leave snapshot within the year
        # so data is visible even if imported under a different month_key
        latest_subq = (
            db.session.query(
                PayrollLeaveSnapshot.employee_id,
                func.max(PayrollLeaveSnapshot.month_key).label("latest_month"),
            )
            .filter(PayrollLeaveSnapshot.year == year)
            .group_by(PayrollLeaveSnapshot.employee_id)
            .subquery()
        )

        rows = (
            db.session.query(PayrollLeaveSnapshot, Employee)
            .join(Employee, Employee.id == PayrollLeaveSnapshot.employee_id)
            .join(
                latest_subq,
                db.and_(
                    PayrollLeaveSnapshot.employee_id == latest_subq.c.employee_id,
                    PayrollLeaveSnapshot.month_key == latest_subq.c.latest_month,
                ),
            )
            .order_by(Employee.employee_code.asc())
            .all()
        )

        if search_query:
            rows = [
                row
                for row in rows
                if search_query in (row[1].employee_code or "").lower()
                or search_query in (row[1].full_name or "").lower()
            ]

        total_used = sum(float(item[0].used_days or 0) for item in rows)
        total_remaining = sum(float(item[0].remaining_days or 0) for item in rows)
        total_entitled = sum(float(item[0].entitled_days or 0) for item in rows)

        return render_template(
            "payroll_leave.html",
            title="Phep nam",
            month_key=month_key,
            year=year,
            search_query=search_query,
            rows=rows,
            total_used=round(total_used, 2),
            total_remaining=round(total_remaining, 2),
            total_entitled=round(total_entitled, 2),
        )

    @app.route("/payroll/slips", methods=["GET"])
    def payroll_slips():
        month_key = _safe_month_key(request.args.get("month"))
        search_query = (request.args.get("q") or "").strip().lower()

        if not _month_has_salary_data(month_key):
            return render_template(
                "payroll_slips.html",
                title="Phieu luong",
                month_key=month_key,
                search_query=search_query,
                rows=[],
                gross_total=0,
                net_total=0,
            )

        rows, gross_total, net_total = _build_slip_rows(month_key, search_query)

        return render_template(
            "payroll_slips.html",
            title="Phieu luong",
            month_key=month_key,
            search_query=search_query,
            rows=rows,
            gross_total=round(gross_total, 2),
            net_total=round(net_total, 2),
        )


    @app.route("/payroll/slips/<int:employee_id>", methods=["GET"])
    def payroll_slip_detail(employee_id):
        month_key = _safe_month_key(request.args.get("month"))

        employee = Employee.query.filter_by(id=employee_id).first()
        if not employee:
            flash("Khong tim thay nhan vien", "error")
            return redirect(url_for("payroll_slips"))

        slip = PayrollSlip.query.filter_by(employee_id=employee_id, month_key=month_key).first()
        insurance = PayrollInsuranceContribution.query.filter_by(employee_id=employee_id, month_key=month_key).first()
        leave = PayrollLeaveSnapshot.query.filter_by(employee_id=employee_id, month_key=month_key).first()

        # attendance summary
        attendance_details = []
        try:
            from .models import AttendanceDetail

            attendance_details = (
                AttendanceDetail.query.filter_by(employee_id=employee_id, month_key=month_key)
                .order_by(AttendanceDetail.work_date.asc())
                .all()
            )
        except Exception:
            attendance_details = []

        # basic totals
        total_hours = sum(float(d.actual_work_hours or 0) for d in attendance_details)
        total_paid_hours = sum(float(d.paid_hours or 0) for d in attendance_details)

        # salary history (last 12 months)
        salary_history = (
            db.session.query(MonthlySalary)
            .filter(MonthlySalary.employee_id == employee_id)
            .order_by(MonthlySalary.month_key.desc())
            .limit(12)
            .all()
        )

        return render_template(
            "payroll_slip_detail.html",
            title=f"Phieu luong - {employee.full_name}",
            employee=employee,
            month_key=month_key,
            slip=slip,
            insurance=insurance,
            leave=leave,
            attendance_details=attendance_details,
            total_hours=round(total_hours, 2),
            total_paid_hours=round(total_paid_hours, 2),
            salary_history=salary_history,
        )

    @app.route("/payroll/insurance-tax", methods=["GET"])
    def payroll_insurance_tax():
        month_key = _safe_month_key(request.args.get("month"))
        search_query = (request.args.get("q") or "").strip().lower()

        if not _month_has_salary_data(month_key):
            return render_template(
                "payroll_insurance_tax.html",
                title="BHXH va Thue",
                month_key=month_key,
                search_query=search_query,
                rows=[],
                total_employer=0,
                total_employee=0,
                total_tax=0,
            )

        combined_rows, total_employer, total_employee, total_tax = _build_insurance_tax_rows(
            month_key,
            search_query,
        )

        return render_template(
            "payroll_insurance_tax.html",
            title="BHXH va Thue",
            month_key=month_key,
            search_query=search_query,
            rows=combined_rows,
            total_employer=round(total_employer, 2),
            total_employee=round(total_employee, 2),
            total_tax=round(total_tax, 2),
        )

    @app.route("/__debug/payroll-insurance", methods=["GET"])
    def _debug_payroll_insurance():
        # Temporary diagnostic endpoint: shows DB URI and counts of insurance/tax rows by month
        blocked = _require_admin()
        if blocked:
            return blocked

        ins = db.session.query(PayrollInsuranceContribution.month_key, db.func.count()).group_by(
            PayrollInsuranceContribution.month_key
        ).all()
        tax = db.session.query(PayrollTaxContribution.month_key, db.func.count()).group_by(
            PayrollTaxContribution.month_key
        ).all()

        return jsonify(
            db_uri=current_app.config.get("SQLALCHEMY_DATABASE_URI"),
            insurance_counts={k: int(v) for k, v in ins},
            tax_counts={k: int(v) for k, v in tax},
        )

    @app.route("/settings", methods=["GET"])
    def settings():
        blocked = _require_admin()
        if blocked:
            return blocked

        return render_template(
            "settings.html",
            title="Cai dat",
            enable_ot_after_6pm_meal=current_app.config.get("ENABLE_OT_AFTER_6PM_MEAL", False),
        )

    @app.route("/settings/toggle_ot_after_6pm", methods=["POST"])
    def settings_toggle_ot_after_6pm():
        blocked = _require_admin()
        if blocked:
            return blocked

        current_value = bool(current_app.config.get("ENABLE_OT_AFTER_6PM_MEAL", False))
        new_value = not current_value
        current_app.config["ENABLE_OT_AFTER_6PM_MEAL"] = new_value

        actor = session.get("username") or "admin"
        log_action(
            "app_config",
            "ENABLE_OT_AFTER_6PM_MEAL",
            "UPDATE",
            changed_by=actor,
            before_data={"enabled": current_value},
            after_data={"enabled": new_value},
            notes="Toggle OT-after-6pm meal from settings",
        )

        flash(f"Tinh tien an OT sau 18:00: {'Bat' if new_value else 'Tat'}", "success")
        return redirect(url_for("settings"))
