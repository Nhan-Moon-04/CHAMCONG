from datetime import date, datetime, time
from decimal import Decimal

from .database import db


class SerializableMixin:
    def to_dict(self):
        payload = {}
        for column in self.__table__.columns:
            value = getattr(self, column.name)
            if isinstance(value, (datetime, date, time)):
                payload[column.name] = value.isoformat()
            elif isinstance(value, Decimal):
                payload[column.name] = float(value)
            else:
                payload[column.name] = value
        return payload


class TimestampMixin:
    created_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    updated_at = db.Column(
        db.DateTime,
        nullable=False,
        default=datetime.utcnow,
        onupdate=datetime.utcnow,
    )


class AppUser(db.Model, TimestampMixin, SerializableMixin):
    __tablename__ = "app_users"

    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(64), nullable=False, unique=True, index=True)
    password_hash = db.Column(db.String(255), nullable=False)
    full_name = db.Column(db.String(120), nullable=True)
    is_admin = db.Column(db.Boolean, nullable=False, default=False)
    is_active = db.Column(db.Boolean, nullable=False, default=True)

    def to_dict(self):
        payload = super().to_dict()
        payload.pop("password_hash", None)
        return payload


class ShiftTemplate(db.Model, TimestampMixin, SerializableMixin):
    __tablename__ = "shift_templates"

    id = db.Column(db.Integer, primary_key=True)
    code = db.Column(db.String(16), nullable=False, unique=True, index=True)
    name = db.Column(db.String(120), nullable=False)
    start_time = db.Column(db.Time, nullable=True)
    end_time = db.Column(db.Time, nullable=True)
    break_minutes = db.Column(db.Integer, nullable=False, default=0)
    standard_hours = db.Column(db.Numeric(5, 2), nullable=False, default=8)
    default_overtime_hours = db.Column(db.Numeric(5, 2), nullable=False, default=0)
    meal_allowance = db.Column(db.Numeric(12, 2), nullable=False, default=0)
    is_leave_code = db.Column(db.Boolean, nullable=False, default=False)
    is_paid_leave = db.Column(db.Boolean, nullable=False, default=False)
    notes = db.Column(db.Text, nullable=True)


class Employee(db.Model, TimestampMixin, SerializableMixin):
    __tablename__ = "employees"

    id = db.Column(db.Integer, primary_key=True)
    employee_code = db.Column(db.String(32), nullable=False, unique=True, index=True)
    full_name = db.Column(db.String(120), nullable=False)
    gender = db.Column(db.String(16), nullable=True)
    hometown = db.Column(db.String(120), nullable=True)
    birth_year = db.Column(db.Integer, nullable=True)
    default_shift_code = db.Column(
        db.String(16), db.ForeignKey("shift_templates.code"), nullable=False, default="X"
    )
    is_active = db.Column(db.Boolean, nullable=False, default=True)

    default_shift = db.relationship(
        "ShiftTemplate",
        primaryjoin="Employee.default_shift_code == ShiftTemplate.code",
        foreign_keys=[default_shift_code],
    )


class MonthlySalary(db.Model, TimestampMixin, SerializableMixin):
    __tablename__ = "monthly_salaries"

    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey("employees.id"), nullable=False)
    month_key = db.Column(db.String(7), nullable=False, index=True)
    base_daily_wage = db.Column(db.Numeric(12, 2), nullable=False, default=0)
    pay_method = db.Column(db.String(32), nullable=True)
    salary_coefficient = db.Column(db.Numeric(10, 4), nullable=False, default=1)

    employee = db.relationship("Employee")

    __table_args__ = (
        db.UniqueConstraint("employee_id", "month_key", name="uq_salary_employee_month"),
    )


class AdvancePayment(db.Model, TimestampMixin, SerializableMixin):
    __tablename__ = "advance_payments"

    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey("employees.id"), nullable=False, index=True)
    advance_date = db.Column(db.Date, nullable=False, index=True)
    month_key = db.Column(db.String(7), nullable=False, index=True)
    amount = db.Column(db.Numeric(12, 2), nullable=False, default=0)
    input_mode = db.Column(db.String(16), nullable=False, default="amount")
    payment_method = db.Column(db.String(32), nullable=False, default="cash")
    advance_days = db.Column(db.Numeric(6, 2), nullable=True)
    notes = db.Column(db.String(255), nullable=True)

    employee = db.relationship("Employee")


class MonthlyWorkdayConfig(db.Model, TimestampMixin, SerializableMixin):
    __tablename__ = "monthly_workday_configs"

    id = db.Column(db.Integer, primary_key=True)
    month_key = db.Column(db.String(7), nullable=False, unique=True, index=True)
    company_work_days = db.Column(db.Numeric(6, 2), nullable=False, default=26)
    notes = db.Column(db.String(255), nullable=True)


class PayrollPaymentStatus(db.Model, TimestampMixin, SerializableMixin):
    __tablename__ = "payroll_payment_statuses"

    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey("employees.id"), nullable=False, index=True)
    month_key = db.Column(db.String(7), nullable=False, index=True)
    salary_received = db.Column(db.Boolean, nullable=False, default=False)
    meal_period_1_received = db.Column(db.Boolean, nullable=False, default=False)
    meal_period_2_received = db.Column(db.Boolean, nullable=False, default=False)
    updated_by = db.Column(db.String(64), nullable=True)

    employee = db.relationship("Employee")

    __table_args__ = (
        db.UniqueConstraint(
            "employee_id",
            "month_key",
            name="uq_payment_status_employee_month",
        ),
    )


class PayrollMonthLock(db.Model, TimestampMixin, SerializableMixin):
    __tablename__ = "payroll_month_locks"

    id = db.Column(db.Integer, primary_key=True)
    month_key = db.Column(db.String(7), nullable=False, unique=True, index=True)
    is_locked = db.Column(db.Boolean, nullable=False, default=False, index=True)
    locked_at = db.Column(db.DateTime, nullable=True)
    locked_by = db.Column(db.String(64), nullable=True)
    notes = db.Column(db.String(255), nullable=True)


class Holiday(db.Model, TimestampMixin, SerializableMixin):
    __tablename__ = "holidays"

    id = db.Column(db.Integer, primary_key=True)
    holiday_date = db.Column(db.Date, nullable=False, unique=True, index=True)
    name = db.Column(db.String(120), nullable=False)
    is_paid = db.Column(db.Boolean, nullable=False, default=True)
    notes = db.Column(db.String(255), nullable=True)


class WorkSchedule(db.Model, TimestampMixin, SerializableMixin):
    __tablename__ = "work_schedules"

    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey("employees.id"), nullable=False, index=True)
    work_date = db.Column(db.Date, nullable=False, index=True)
    month_key = db.Column(db.String(7), nullable=False, index=True)
    shift_id = db.Column(db.Integer, db.ForeignKey("shift_templates.id"), nullable=False)
    absence_hours = db.Column(db.Numeric(5, 2), nullable=False, default=0)
    notes = db.Column(db.String(255), nullable=True)

    employee = db.relationship("Employee")
    shift = db.relationship("ShiftTemplate")
    overtime = db.relationship(
        "OvertimeEntry",
        back_populates="schedule",
        uselist=False,
        cascade="all, delete-orphan",
    )

    __table_args__ = (
        db.UniqueConstraint("employee_id", "work_date", name="uq_schedule_employee_date"),
    )


class OvertimeEntry(db.Model, TimestampMixin, SerializableMixin):
    __tablename__ = "overtime_entries"

    id = db.Column(db.Integer, primary_key=True)
    schedule_id = db.Column(
        db.Integer,
        db.ForeignKey("work_schedules.id", ondelete="CASCADE"),
        nullable=False,
        unique=True,
    )
    hours = db.Column(db.Numeric(5, 2), nullable=False, default=0)
    reason = db.Column(db.String(255), nullable=True)

    schedule = db.relationship("WorkSchedule", back_populates="overtime")


class AttendanceLog(db.Model, TimestampMixin, SerializableMixin):
    __tablename__ = "attendance_logs"

    id = db.Column(db.Integer, primary_key=True)
    employee_code = db.Column(db.String(32), nullable=False, index=True)
    employee_name = db.Column(db.String(120), nullable=False)
    department = db.Column(db.String(120), nullable=True)
    event_time = db.Column(db.DateTime, nullable=False, index=True)
    source_file = db.Column(db.String(255), nullable=False)
    import_batch = db.Column(db.String(36), nullable=False, index=True)


class AttendanceDaily(db.Model, TimestampMixin, SerializableMixin):
    __tablename__ = "attendance_daily"

    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey("employees.id"), nullable=False, index=True)
    work_date = db.Column(db.Date, nullable=False, index=True)
    first_check_in = db.Column(db.DateTime, nullable=True)
    last_check_out = db.Column(db.DateTime, nullable=True)
    total_hours = db.Column(db.Numeric(6, 2), nullable=False, default=0)
    import_batch = db.Column(db.String(36), nullable=False, index=True)

    employee = db.relationship("Employee")

    __table_args__ = (
        db.UniqueConstraint("employee_id", "work_date", name="uq_daily_employee_date"),
    )


class AttendanceDetail(db.Model, TimestampMixin, SerializableMixin):
    __tablename__ = "attendance_details"

    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey("employees.id"), nullable=False, index=True)
    work_date = db.Column(db.Date, nullable=False, index=True)
    month_key = db.Column(db.String(7), nullable=False, index=True)
    shift_code = db.Column(db.String(16), nullable=True)
    shift_name = db.Column(db.String(120), nullable=True)
    check_in = db.Column(db.DateTime, nullable=True)
    check_out = db.Column(db.DateTime, nullable=True)
    standard_hours = db.Column(db.Numeric(6, 2), nullable=False, default=0)
    actual_work_hours = db.Column(db.Numeric(6, 2), nullable=False, default=0)
    deviation_hours = db.Column(db.Numeric(6, 2), nullable=False, default=0)
    overtime_hours = db.Column(db.Numeric(6, 2), nullable=False, default=0)
    total_span_hours = db.Column(db.Numeric(6, 2), nullable=False, default=0)
    status_code = db.Column(db.String(16), nullable=False, index=True)
    paid_hours = db.Column(db.Numeric(6, 2), nullable=False, default=0)
    daily_wage = db.Column(db.Numeric(12, 2), nullable=False, default=0)
    notes = db.Column(db.String(255), nullable=True)
    meal_allowance_daily = db.Column(db.Numeric(12, 2), nullable=False, default=0)

    employee = db.relationship("Employee")

    __table_args__ = (
        db.UniqueConstraint("employee_id", "work_date", name="uq_detail_employee_date"),
    )


class LeaveBalance(db.Model, TimestampMixin, SerializableMixin):
    __tablename__ = "leave_balances"

    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey("employees.id"), nullable=False, index=True)
    year = db.Column(db.Integer, nullable=False, index=True)
    total_days = db.Column(db.Numeric(5, 2), nullable=False, default=12)
    used_days = db.Column(db.Numeric(5, 2), nullable=False, default=0)

    employee = db.relationship("Employee")

    __table_args__ = (
        db.UniqueConstraint("employee_id", "year", name="uq_leave_employee_year"),
    )


class AuditLog(db.Model, SerializableMixin):
    __tablename__ = "audit_logs"

    id = db.Column(db.Integer, primary_key=True)
    table_name = db.Column(db.String(64), nullable=False, index=True)
    record_id = db.Column(db.String(64), nullable=False)
    action = db.Column(db.String(32), nullable=False, index=True)
    changed_by = db.Column(db.String(64), nullable=False, default="system")
    changed_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow, index=True)
    before_data = db.Column(db.JSON, nullable=True)
    after_data = db.Column(db.JSON, nullable=True)
    notes = db.Column(db.String(255), nullable=True)
