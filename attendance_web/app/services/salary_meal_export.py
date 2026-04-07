import io
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Font
from sqlalchemy.orm import joinedload

from ..models import AttendanceDetail, Employee
from .attendance import parse_month_key


def _to_float(value, default=0.0):
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def _employee_code_sort_key(employee_code):
    raw_code = str(employee_code or "").replace("'", "").strip()
    if raw_code.isdigit():
        return (0, int(raw_code))
    return (1, raw_code.lower())


def _normalize_period(period):
    return 2 if str(period or "").strip() == "2" else 1


def collect_salary_meal_overview_data(month_key, period, search_query=""):
    period = _normalize_period(period)
    search_query = (search_query or "").strip()

    start_date, end_date = parse_month_key(month_key)
    period_1_end = date(start_date.year, start_date.month, 15)
    period_2_start = date(start_date.year, start_date.month, 16)

    if period == 2:
        period_start = period_2_start
        period_end = end_date
        period_label = f"{period_start.strftime('%d/%m')} - {period_end.strftime('%d/%m')}"
        period_title = "Tien an dot 2"
    else:
        period_start = start_date
        period_end = period_1_end
        period_label = f"{period_start.strftime('%d/%m')} - {period_end.strftime('%d/%m')}"
        period_title = "Tien an dot 1"

    employees = Employee.query.filter(Employee.is_active.is_(True)).order_by(Employee.employee_code.asc()).all()
    meal_summary_map = {
        row.id: {
            "employee": row,
            "worked_days": 0.0,
            "paid_leave_days": 0.0,
            "unpaid_leave_days": 0.0,
            "meal_amount": 0.0,
        }
        for row in employees
    }

    detail_rows = (
        AttendanceDetail.query.options(joinedload(AttendanceDetail.employee))
        .filter(
            AttendanceDetail.month_key == month_key,
            AttendanceDetail.work_date >= period_start,
            AttendanceDetail.work_date <= period_end,
        )
        .order_by(AttendanceDetail.employee_id.asc(), AttendanceDetail.work_date.asc())
        .all()
    )

    def _apply_status(summary_obj, status_code):
        normalized = str(status_code or "").upper()
        if normalized == "P":
            summary_obj["paid_leave_days"] += 1.0
        elif normalized in {"S", "C"}:
            summary_obj["paid_leave_days"] += 0.5
            summary_obj["worked_days"] += 0.5
        elif normalized == "N":
            summary_obj["unpaid_leave_days"] += 1.0
        elif normalized == "OFF":
            return
        else:
            summary_obj["worked_days"] += 1.0

    for detail_row in detail_rows:
        if not detail_row.employee:
            continue

        meal_summary = meal_summary_map.get(detail_row.employee_id)
        if not meal_summary:
            meal_summary = {
                "employee": detail_row.employee,
                "worked_days": 0.0,
                "paid_leave_days": 0.0,
                "unpaid_leave_days": 0.0,
                "meal_amount": 0.0,
            }
            meal_summary_map[detail_row.employee_id] = meal_summary

        _apply_status(meal_summary, detail_row.status_code)
        meal_summary["meal_amount"] += _to_float(detail_row.meal_allowance_daily)

    meal_rows = []
    for employee_id, meal_summary in meal_summary_map.items():
        meal_rows.append(
            {
                "employee_id": employee_id,
                "employee": meal_summary["employee"],
                "worked_days": round(meal_summary["worked_days"], 2),
                "paid_leave_days": round(meal_summary["paid_leave_days"], 2),
                "unpaid_leave_days": round(meal_summary["unpaid_leave_days"], 2),
                "meal_amount": round(meal_summary["meal_amount"], 2),
            }
        )

    meal_rows.sort(
        key=lambda item: _employee_code_sort_key(item["employee"].employee_code)
    )

    if search_query:
        search_text = search_query.lower()

        def _match_meal_row(item):
            values = [
                item["employee"].employee_code,
                item["employee"].full_name,
                item["worked_days"],
                item["paid_leave_days"],
                item["unpaid_leave_days"],
                item["meal_amount"],
            ]
            return any(
                search_text in str(value).lower()
                for value in values
                if value is not None
            )

        meal_rows = [item for item in meal_rows if _match_meal_row(item)]

    return {
        "month_key": month_key,
        "period": period,
        "period_title": period_title,
        "period_label": period_label,
        "search_query": search_query,
        "meal_rows": meal_rows,
    }


def build_salary_meal_export_excel(meal_data):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = f"Tien an dot {meal_data['period']}"

    headers = [
        "STT",
        "Ho ten",
        "So ngay lam",
        "So ngay nghi co phep",
        "So ngay nghi khong phep",
        "So tien an",
    ]
    sheet.append(headers)

    header_font = Font(bold=True)
    amount_font = Font(bold=True)
    for cell in sheet[1]:
        cell.font = header_font

    for index, row in enumerate(meal_data["meal_rows"], start=1):
        sheet.append(
            [
                index,
                row["employee"].full_name,
                float(row["worked_days"]),
                float(row["paid_leave_days"]),
                float(row["unpaid_leave_days"]),
                float(row["meal_amount"]),
            ]
        )

        amount_cell = sheet.cell(row=sheet.max_row, column=6)
        amount_cell.number_format = "#,##0"
        amount_cell.font = amount_font

    sheet.freeze_panes = "A2"

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)

    month_label = str(meal_data["month_key"]).replace("-", "")
    filename = f"tien_an_dot_{meal_data['period']}_{month_label}.xlsx"
    return output, filename
