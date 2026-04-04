import os
import re
import uuid
import unicodedata
from collections import Counter

import pandas as pd

from ..database import db
from ..models import Employee, MonthlySalary, MonthlyWorkdayConfig
from .audit import log_action


CSV_ENCODINGS = ["utf-8-sig", "cp1258", "latin1"]


def _read_csv_with_fallback(file_path):
    last_error = None
    for encoding in CSV_ENCODINGS:
        try:
            return pd.read_csv(file_path, encoding=encoding, on_bad_lines="skip")
        except Exception as exc:
            last_error = exc

    raise ValueError(f"Khong doc duoc file CSV: {last_error}")


def _normalize_text(value):
    text = str(value or "").strip().lower()
    text = unicodedata.normalize("NFD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    text = re.sub(r"[^a-z0-9]+", "", text)
    return text


def _normalize_employee_code(value):
    text = str(value or "").replace("'", "").strip()
    if re.fullmatch(r"\d+\.0+", text):
        text = text.split(".", 1)[0]
    return text


def _to_float(value):
    if value is None:
        return None

    text = str(value).strip()
    if text == "":
        return None

    lowered = text.lower()
    if lowered in {"nan", "none", "null", "-"}:
        return None

    text = text.replace(" ", "")

    if text.count(",") > 1 and "." not in text:
        text = text.replace(",", "")
    elif text.count(".") > 1 and "," not in text:
        text = text.replace(".", "")
    else:
        text = text.replace(",", "")

    try:
        return float(text)
    except ValueError:
        return None


def _read_frame(file_path):
    extension = os.path.splitext(file_path)[1].lower()

    if extension == ".csv":
        frame = _read_csv_with_fallback(file_path)
    elif extension in {".xlsx", ".xls", ".xlsm", ".xltx", ".xltm"}:
        frame = pd.read_excel(file_path)
    else:
        raise ValueError("Chi ho tro file luong dinh dang .csv hoac .xlsx/.xls")

    frame = frame.dropna(axis=0, how="all")
    frame = frame.dropna(axis=1, how="all")

    if frame.empty:
        raise ValueError("File he luong khong co du lieu")

    return frame


def _find_columns(frame):
    employee_code_col = None
    monthly_wage_col = None
    workday_coeff_col = None
    pay_method_col = None

    for column in frame.columns:
        norm = _normalize_text(column)

        if employee_code_col is None and (
            norm in {"manv", "manhanvien", "idnv"}
            or ("ma" in norm and "nv" in norm)
            or (norm.startswith("id") and "nv" in norm)
        ):
            employee_code_col = column
            continue

        if monthly_wage_col is None and (
            "luongthang" in norm
            or ("muc" in norm and "luong" in norm and "thang" in norm)
            or norm in {"luong", "mucluong"}
        ):
            monthly_wage_col = column
            continue

        if workday_coeff_col is None and (
            "hesochialuong" in norm
            or ("heso" in norm and ("ngay" in norm or "cong" in norm or "chia" in norm))
            or "congchuan" in norm
        ):
            workday_coeff_col = column
            continue

        if pay_method_col is None and (
            "hinhthucnhantien" in norm
            or ("hinhthuc" in norm and "tien" in norm)
            or "paymethod" in norm
        ):
            pay_method_col = column

    if employee_code_col is None:
        raise ValueError("Khong tim thay cot Ma NV trong file he luong")

    if monthly_wage_col is None:
        raise ValueError("Khong tim thay cot Muc Luong Thang trong file he luong")

    return {
        "employee_code": employee_code_col,
        "monthly_wage": monthly_wage_col,
        "workday_coeff": workday_coeff_col,
        "pay_method": pay_method_col,
    }


def import_salary_file(
    file_path,
    source_name,
    actor="system",
    target_month=None,
    default_company_work_days=26.0,
    replace_existing=False,
):
    if not target_month:
        raise ValueError("Can chon thang de import he luong")

    frame = _read_frame(file_path)
    columns = _find_columns(frame)

    entries = []
    workday_candidates = []

    for _, row in frame.iterrows():
        employee_code = _normalize_employee_code(row.get(columns["employee_code"]))
        monthly_wage = _to_float(row.get(columns["monthly_wage"]))

        if columns["workday_coeff"] is not None:
            workday_value = _to_float(row.get(columns["workday_coeff"]))
            if workday_value and workday_value > 0:
                workday_candidates.append(float(workday_value))

        if not employee_code:
            continue

        if monthly_wage is None:
            continue

        pay_method = None
        if columns["pay_method"] is not None:
            pay_method_raw = str(row.get(columns["pay_method"]) or "").strip()
            if pay_method_raw and pay_method_raw.lower() not in {"nan", "none", "null", "-"}:
                pay_method = pay_method_raw

        entries.append((employee_code, float(monthly_wage), pay_method))

    if not entries:
        raise ValueError("Khong tim thay dong luong hop le trong file")

    company_work_days = float(default_company_work_days or 26.0)
    if workday_candidates:
        rounded_values = [round(value, 2) for value in workday_candidates if value > 0]
        if rounded_values:
            company_work_days = Counter(rounded_values).most_common(1)[0][0]

    if company_work_days <= 0:
        company_work_days = 26.0

    deleted_rows = 0
    if replace_existing:
        deleted_rows = MonthlySalary.query.filter_by(month_key=target_month).delete(
            synchronize_session=False
        )

    unique_codes = sorted({item[0] for item in entries})
    employee_rows = Employee.query.filter(Employee.employee_code.in_(unique_codes)).all()
    employee_map = {row.employee_code: row for row in employee_rows}

    employee_ids = [row.id for row in employee_rows]
    existing_map = {}
    if employee_ids and not replace_existing:
        existing_rows = MonthlySalary.query.filter(
            MonthlySalary.month_key == target_month,
            MonthlySalary.employee_id.in_(employee_ids),
        ).all()
        existing_map = {row.employee_id: row for row in existing_rows}

    created = 0
    updated = 0
    skipped_unknown = 0
    unknown_codes = set()

    for employee_code, monthly_wage, pay_method in entries:
        employee = employee_map.get(employee_code)
        if not employee:
            skipped_unknown += 1
            unknown_codes.add(employee_code)
            continue

        row = existing_map.get(employee.id)
        if row:
            row.base_daily_wage = monthly_wage
            row.salary_coefficient = company_work_days
            if pay_method is not None:
                row.pay_method = pay_method
            updated += 1
        else:
            row = MonthlySalary(
                employee_id=employee.id,
                month_key=target_month,
                base_daily_wage=monthly_wage,
                salary_coefficient=company_work_days,
                pay_method=pay_method,
            )
            db.session.add(row)
            existing_map[employee.id] = row
            created += 1

    config = MonthlyWorkdayConfig.query.filter_by(month_key=target_month).first()
    if config:
        config.company_work_days = company_work_days
        config.notes = "Import he luong"
    else:
        config = MonthlyWorkdayConfig(
            month_key=target_month,
            company_work_days=company_work_days,
            notes="Import he luong",
        )
        db.session.add(config)

    month_salary_rows = MonthlySalary.query.filter_by(month_key=target_month).all()
    for row in month_salary_rows:
        row.salary_coefficient = company_work_days

    batch_id = str(uuid.uuid4())
    log_action(
        "salary_import",
        batch_id,
        "IMPORT",
        changed_by=actor,
        after_data={
            "source_file": source_name,
            "month_key": target_month,
            "replace_existing": bool(replace_existing),
            "deleted_rows": int(deleted_rows),
            "rows_in_file": len(entries),
            "created": created,
            "updated": updated,
            "skipped_unknown": skipped_unknown,
            "unknown_codes": sorted(list(unknown_codes))[:50],
            "company_work_days": company_work_days,
        },
        notes="Import he luong theo thang",
    )

    db.session.commit()

    return {
        "batch_id": batch_id,
        "month_key": target_month,
        "replace_existing": bool(replace_existing),
        "deleted_rows": int(deleted_rows),
        "rows_in_file": len(entries),
        "created": created,
        "updated": updated,
        "skipped_unknown": skipped_unknown,
        "unknown_codes": sorted(list(unknown_codes)),
        "company_work_days": company_work_days,
    }
