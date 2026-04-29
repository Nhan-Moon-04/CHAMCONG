"""
Salary calculator service for detailed employee wage breakdown.

Rules:
1. Regular days (Mon-Sat): salary = (base_wage / salary_coefficient / 8) * paid_hours
2. OT on regular days: OT_wage = (base_wage / salary_coefficient / 8) * 1.5 * OT_hours
3. Sunday work: Sunday_wage = (base_wage / salary_coefficient / 8) * 2 * Sunday_hours
"""

from datetime import date
from ..models import AttendanceDetail, MonthlySalary


def _to_float(value, default=0.0):
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def get_salary_month_details(employee_id, month_key):
    """
    Calculate detailed salary breakdown for an employee in a given month.
    
    Returns dict with:
    - monthly_salary: MonthlySalary record for the month
    - base_wage: base_daily_wage from record
    - salary_coefficient: salary_coefficient from record (hệ số lương)
    - hourly_rate: base_wage / coefficient / 8
    - regular_days_breakdown: list of daily breakdowns (Mon-Sat)
    - sunday_breakdown: list of Sunday breakdowns
    - summary: totals for regular, OT, Sunday
    """
    
    monthly_salary = MonthlySalary.query.filter_by(
        employee_id=employee_id,
        month_key=month_key
    ).first()
    
    if not monthly_salary:
        return {
            "monthly_salary": None,
            "base_wage": 0.0,
            "salary_coefficient": 1.0,
            "hourly_rate": 0.0,
            "ot_hourly_rate": 0.0,
            "sunday_hourly_rate": 0.0,
            "regular_days_breakdown": [],
            "sunday_breakdown": [],
            "summary": {
                "regular_wage": 0.0,
                "regular_hours": 0.0,
                "ot_wage": 0.0,
                "ot_hours": 0.0,
                "sunday_wage": 0.0,
                "sunday_hours": 0.0,
                "total_wage": 0.0,
            }
        }
    
    base_wage = _to_float(monthly_salary.base_daily_wage, 0.0)
    salary_coefficient = _to_float(monthly_salary.salary_coefficient, 1.0)
    
    if salary_coefficient <= 0:
        salary_coefficient = 1.0
    
    # Hourly rate = (base_wage / coefficient / 8)
    hourly_rate = (base_wage / salary_coefficient / 8) if base_wage > 0 else 0.0
    ot_hourly_rate = hourly_rate * 1.5
    sunday_hourly_rate = hourly_rate * 2.0
    
    # Fetch attendance details for the month
    from .attendance import parse_month_key
    start_date, end_date = parse_month_key(month_key)
    
    details = (
        AttendanceDetail.query.filter(
            AttendanceDetail.employee_id == employee_id,
            AttendanceDetail.month_key == month_key,
            AttendanceDetail.work_date >= start_date,
            AttendanceDetail.work_date <= end_date,
        )
        .order_by(AttendanceDetail.work_date.asc())
        .all()
    )
    
    regular_days_breakdown = []
    sunday_breakdown = []
    summary = {
        "regular_wage": 0.0,
        "regular_hours": 0.0,
        "ot_wage": 0.0,
        "ot_hours": 0.0,
        "sunday_wage": 0.0,
        "sunday_hours": 0.0,
        "total_wage": 0.0,
    }
    
    for detail in details:
        work_date = detail.work_date
        weekday = work_date.weekday()  # 0=Mon, 6=Sun
        is_sunday = (weekday == 6)
        
        paid_hours = _to_float(detail.paid_hours, 0.0)
        overtime_hours = _to_float(detail.overtime_hours, 0.0)
        
        if is_sunday:
            # Sunday: all hours at 2x rate
            sunday_wage = sunday_hourly_rate * paid_hours
            sunday_breakdown.append({
                "work_date": work_date,
                "shift_code": detail.shift_code,
                "shift_name": detail.shift_name,
                "paid_hours": paid_hours,
                "overtime_hours": overtime_hours,
                "hourly_rate": sunday_hourly_rate,
                "wage": round(sunday_wage, 2),
            })
            summary["sunday_wage"] += sunday_wage
            summary["sunday_hours"] += paid_hours
        else:
            # Regular day (Mon-Sat)
            regular_wage = hourly_rate * paid_hours
            ot_wage = ot_hourly_rate * overtime_hours
            day_total_wage = regular_wage + ot_wage
            
            regular_days_breakdown.append({
                "work_date": work_date,
                "shift_code": detail.shift_code,
                "shift_name": detail.shift_name,
                "paid_hours": paid_hours,
                "regular_hourly_rate": hourly_rate,
                "regular_wage": round(regular_wage, 2),
                "overtime_hours": overtime_hours,
                "ot_hourly_rate": ot_hourly_rate,
                "ot_wage": round(ot_wage, 2),
                "day_total": round(day_total_wage, 2),
            })
            summary["regular_wage"] += regular_wage
            summary["regular_hours"] += paid_hours
            summary["ot_wage"] += ot_wage
            summary["ot_hours"] += overtime_hours
    
    summary["regular_wage"] = round(summary["regular_wage"], 2)
    summary["ot_wage"] = round(summary["ot_wage"], 2)
    summary["sunday_wage"] = round(summary["sunday_wage"], 2)
    summary["total_wage"] = round(
        summary["regular_wage"] + summary["ot_wage"] + summary["sunday_wage"],
        2
    )
    
    return {
        "monthly_salary": monthly_salary,
        "base_wage": base_wage,
        "salary_coefficient": salary_coefficient,
        "hourly_rate": round(hourly_rate, 2),
        "ot_hourly_rate": round(ot_hourly_rate, 2),
        "sunday_hourly_rate": round(sunday_hourly_rate, 2),
        "regular_days_breakdown": regular_days_breakdown,
        "sunday_breakdown": sunday_breakdown,
        "summary": summary,
    }
