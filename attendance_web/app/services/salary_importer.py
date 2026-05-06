import os
import re
import uuid
import unicodedata
from collections import Counter
from datetime import datetime

import pandas as pd

from ..database import db
from ..models import (
    Employee,
    MonthlySalary,
    MonthlyWorkdayConfig,
    PayrollInsuranceContribution,
    PayrollLeaveSnapshot,
    PayrollSlip,
    PayrollTaxContribution,
)
from .audit import log_action


CSV_ENCODINGS = ["utf-8-sig", "cp1258", "latin1"]
PREFERRED_SALARY_SHEET_NAME = "bangluong"


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


def _normalize_person_name(value):
    text = str(value or "").strip().lower()
    text = unicodedata.normalize("NFD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    text = re.sub(r"[^a-z0-9]+", "", text)
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


def _safe_number(value, default=0.0):
    parsed = _to_float(value)
    return float(parsed if parsed is not None else default)


def _extract_number_candidates(value):
    if value is None:
        return []

    if isinstance(value, (int, float)):
        number = float(value)
        return [number] if number > 0 else []

    text = str(value).strip()
    if not text:
        return []

    direct = _to_float(text)
    if direct is not None and direct > 0:
        return [float(direct)]

    numbers = []
    for matched in re.findall(r"[-+]?\d[\d.,]*", text):
        parsed = _to_float(matched)
        if parsed is not None and parsed > 0:
            numbers.append(float(parsed))

    return numbers


def _is_employee_code_header(norm):
    return (
        norm in {"manv", "manhanvien", "idnv"}
        or ("ma" in norm and "nv" in norm)
        or (norm.startswith("id") and "nv" in norm)
    )


def _is_monthly_wage_header(norm):
    if not norm:
        return False

    if "heso" in norm:
        return False

    return (
        "luongthang" in norm
        or "tongluong" in norm
        or ("muc" in norm and "luong" in norm and "thang" in norm)
        or norm in {"luong", "mucluong", "salary", "monthlysalary"}
    )


def _is_base_wage_header(norm):
    if not norm:
        return False

    if "heso" in norm:
        return False

    return (
        "luongcoban" in norm
        or ("luong" in norm and "coban" in norm)
        or "basewage" in norm
        or "basicwage" in norm
    )


def _is_allowance_header(norm):
    return "phucap" in norm or "allowance" in norm


def _is_workday_coeff_header(norm):
    return (
        "hesochialuong" in norm
        or "hesoluong" in norm
        or ("heso" in norm and ("ngay" in norm or "cong" in norm or "chia" in norm))
        or "congchuan" in norm
    )


def _is_pay_method_header(norm):
    return (
        "hinhthucnhantien" in norm
        or ("hinhthuc" in norm and "tien" in norm)
        or "paymethod" in norm
        or "paymentmethod" in norm
    )


def _make_unique_headers(raw_headers):
    unique_headers = []
    seen = {}

    for index, value in enumerate(raw_headers):
        base = str(value).strip() if value is not None else ""
        if not base:
            base = f"__col_{index}"

        count = seen.get(base, 0)
        seen[base] = count + 1

        if count == 0:
            unique_headers.append(base)
        else:
            unique_headers.append(f"{base}__{count}")

    return unique_headers


def _detect_salary_header_row(raw_frame):
    scan_limit = min(20, len(raw_frame.index))

    for row_index in range(scan_limit):
        row_values = raw_frame.iloc[row_index].tolist()
        normalized_values = [_normalize_text(value) for value in row_values]

        has_employee = any(_is_employee_code_header(norm) for norm in normalized_values)
        has_salary = any(
            _is_monthly_wage_header(norm) or _is_base_wage_header(norm) or _is_allowance_header(norm)
            for norm in normalized_values
        )

        if has_employee and has_salary:
            return row_index

    return None


def _extract_workday_candidates_from_raw(raw_frame):
    candidates = []
    scan_rows = min(12, raw_frame.shape[0])
    scan_columns = raw_frame.shape[1]

    for row_index in range(scan_rows):
        for column_index in range(scan_columns):
            value = raw_frame.iat[row_index, column_index]
            norm = _normalize_text(value)
            if not norm or not _is_workday_coeff_header(norm):
                continue

            candidates.extend(_extract_number_candidates(value))

            for offset in (1, 2, 3):
                next_col = column_index + offset
                if next_col >= scan_columns:
                    break
                candidates.extend(_extract_number_candidates(raw_frame.iat[row_index, next_col]))

    return [value for value in candidates if value > 0]


def _build_frame_from_raw(raw_frame, header_row):
    header_values = [raw_frame.iat[header_row, index] for index in range(raw_frame.shape[1])]
    frame = raw_frame.iloc[header_row + 1 :].copy()
    frame.columns = _make_unique_headers(header_values)
    frame = frame.dropna(axis=0, how="all")

    # Keep declared header columns even when current rows are empty.
    # This allows format detection on template-like files where salary cells are not filled yet.
    keep_columns = []
    for column in frame.columns:
        norm = _normalize_text(column)
        has_data = frame[column].notna().any()
        if has_data or (norm and not str(column).startswith("__col_")):
            keep_columns.append(column)

    if keep_columns:
        frame = frame[keep_columns]

    return frame


def _read_excel_frame(file_path):
    workbook = pd.ExcelFile(file_path)
    sheet_names = workbook.sheet_names

    ordered_sheet_names = sorted(
        sheet_names,
        key=lambda name: (_normalize_text(name) != PREFERRED_SALARY_SHEET_NAME, name.lower()),
    )

    for sheet_name in ordered_sheet_names:
        raw_frame = workbook.parse(sheet_name=sheet_name, header=None)
        raw_frame = raw_frame.dropna(axis=0, how="all")
        raw_frame = raw_frame.dropna(axis=1, how="all")

        if raw_frame.empty:
            continue

        header_row = _detect_salary_header_row(raw_frame)
        if header_row is None:
            continue

        frame = _build_frame_from_raw(raw_frame, header_row)
        if frame.empty:
            continue

        return frame, {
            "sheet_name": sheet_name,
            "header_row": int(header_row + 1),
            "workday_candidates": _extract_workday_candidates_from_raw(raw_frame),
        }

    fallback_frame = pd.read_excel(file_path)
    fallback_frame = fallback_frame.dropna(axis=0, how="all")
    fallback_frame = fallback_frame.dropna(axis=1, how="all")

    if fallback_frame.empty:
        raise ValueError("File he luong khong co du lieu")

    return fallback_frame, {
        "sheet_name": (sheet_names[0] if sheet_names else None),
        "header_row": 1,
        "workday_candidates": [],
    }


def _read_frame(file_path):
    extension = os.path.splitext(file_path)[1].lower()

    if extension == ".csv":
        frame = _read_csv_with_fallback(file_path)
        source_info = {
            "sheet_name": None,
            "header_row": 1,
            "workday_candidates": [],
        }
    elif extension in {".xlsx", ".xls", ".xlsm", ".xltx", ".xltm"}:
        frame, source_info = _read_excel_frame(file_path)
    else:
        raise ValueError("Chi ho tro file luong dinh dang .csv hoac .xlsx/.xls")

    frame = frame.dropna(axis=0, how="all")

    if frame.empty:
        raise ValueError("File he luong khong co du lieu")

    return frame, source_info


def _find_columns(frame):
    employee_code_col = None
    monthly_wage_col = None
    base_wage_col = None
    allowance_col = None
    workday_coeff_col = None
    pay_method_col = None

    for column in frame.columns:
        norm = _normalize_text(column)

        if employee_code_col is None and _is_employee_code_header(norm):
            employee_code_col = column
            continue

        if monthly_wage_col is None and _is_monthly_wage_header(norm):
            monthly_wage_col = column
            continue

        if base_wage_col is None and _is_base_wage_header(norm):
            base_wage_col = column
            continue

        if allowance_col is None and _is_allowance_header(norm):
            allowance_col = column
            continue

        if workday_coeff_col is None and _is_workday_coeff_header(norm):
            workday_coeff_col = column
            continue

        if pay_method_col is None and _is_pay_method_header(norm):
            pay_method_col = column

    if employee_code_col is None:
        raise ValueError("Khong tim thay cot Ma NV trong file he luong")

    if monthly_wage_col is None and base_wage_col is None:
        raise ValueError(
            "Khong tim thay cot luong hop le trong file he luong (Muc Luong Thang hoac Luong Co Ban)"
        )

    return {
        "employee_code": employee_code_col,
        "monthly_wage": monthly_wage_col,
        "base_wage": base_wage_col,
        "allowance": allowance_col,
        "workday_coeff": workday_coeff_col,
        "pay_method": pay_method_col,
    }


def _find_sheet_name(workbook, target_name):
    target_norm = _normalize_text(target_name)
    for name in workbook.sheet_names:
        if _normalize_text(name) == target_norm:
            return name
    return None


def _build_employee_maps():
    employee_rows = Employee.query.all()
    by_code = {}
    by_name = {}

    for row in employee_rows:
        normalized_code = _normalize_employee_code(row.employee_code)
        if normalized_code:
            by_code[normalized_code] = row

        normalized_name = _normalize_person_name(row.full_name)
        if normalized_name and normalized_name not in by_name:
            by_name[normalized_name] = row

    return by_code, by_name


def _resolve_employee(by_code, by_name, code_value=None, name_value=None):
    code = _normalize_employee_code(code_value)
    if code and code in by_code:
        return by_code[code]

    name_key = _normalize_person_name(name_value)
    if name_key and name_key in by_name:
        return by_name[name_key]

    return None


def _upsert_leave_snapshots(workbook, month_key, source_name, by_code, by_name):
    sheet_name = _find_sheet_name(workbook, "phep nam")
    if not sheet_name:
        return {"sheet": None, "created": 0, "updated": 0, "skipped": 0}

    frame = workbook.parse(sheet_name=sheet_name, header=None).fillna("")
    year = int(str(month_key).split("-")[0])

    created = 0
    updated = 0
    skipped = 0

    for row_index in range(3, len(frame.index)):
        row = frame.iloc[row_index]
        name_value = row.iloc[1] if len(row) > 1 else None

        employee = _resolve_employee(by_code, by_name, name_value=name_value)
        if not employee:
            skipped += 1
            continue

        monthly_breakdown = {
            f"t{month}": _safe_number(row.iloc[9 + month]) if len(row) > (9 + month) else 0.0
            for month in range(1, 13)
        }

        values = {
            "year": year,
            "entitled_days": _safe_number(row.iloc[9] if len(row) > 9 else 0),
            "bonus_entitled_days": _safe_number(row.iloc[22] if len(row) > 22 else 0),
            "used_days": _safe_number(row.iloc[23] if len(row) > 23 else 0),
            "remaining_days": _safe_number(row.iloc[24] if len(row) > 24 else 0),
            "work_days": _safe_number(row.iloc[25] if len(row) > 25 else 0),
            "sick_leave_days": _safe_number(row.iloc[26] if len(row) > 26 else 0),
            "monthly_breakdown": monthly_breakdown,
            "source_file": source_name,
        }

        existing = PayrollLeaveSnapshot.query.filter_by(
            employee_id=employee.id,
            month_key=month_key,
        ).first()
        if existing:
            for key, value in values.items():
                setattr(existing, key, value)
            updated += 1
        else:
            db.session.add(
                PayrollLeaveSnapshot(
                    employee_id=employee.id,
                    month_key=month_key,
                    **values,
                )
            )
            created += 1

    return {"sheet": sheet_name, "created": created, "updated": updated, "skipped": skipped}


def _upsert_insurance_rows(workbook, month_key, source_name, by_code, by_name):
    sheet_name = _find_sheet_name(workbook, "BAO HIEM")
    if not sheet_name:
        return {"sheet": None, "created": 0, "updated": 0, "skipped": 0}

    frame = workbook.parse(sheet_name=sheet_name, header=None).fillna("")

    created = 0
    updated = 0
    skipped = 0

    for row_index in range(5, len(frame.index)):
        row = frame.iloc[row_index]
        name_value = row.iloc[4] if len(row) > 4 else None
        employee = _resolve_employee(by_code, by_name, name_value=name_value)
        if not employee:
            skipped += 1
            continue

        values = {
            "insured_salary": _safe_number(row.iloc[12] if len(row) > 12 else 0),
            "employer_bhxh": _safe_number(row.iloc[13] if len(row) > 13 else 0),
            "employee_bhxh": _safe_number(row.iloc[14] if len(row) > 14 else 0),
            "employer_bhyt": _safe_number(row.iloc[15] if len(row) > 15 else 0),
            "employee_bhyt": _safe_number(row.iloc[16] if len(row) > 16 else 0),
            "employer_bhtn": _safe_number(row.iloc[17] if len(row) > 17 else 0),
            "employee_bhtn": _safe_number(row.iloc[18] if len(row) > 18 else 0),
            "employer_accident": _safe_number(row.iloc[19] if len(row) > 19 else 0),
            "employer_total": _safe_number(row.iloc[20] if len(row) > 20 else 0),
            "employee_total": _safe_number(row.iloc[21] if len(row) > 21 else 0),
            "union_fund": _safe_number(row.iloc[23] if len(row) > 23 else 0),
            "source_file": source_name,
        }

        existing = PayrollInsuranceContribution.query.filter_by(
            employee_id=employee.id,
            month_key=month_key,
        ).first()
        if existing:
            for key, value in values.items():
                setattr(existing, key, value)
            updated += 1
        else:
            db.session.add(
                PayrollInsuranceContribution(
                    employee_id=employee.id,
                    month_key=month_key,
                    **values,
                )
            )
            created += 1

    return {"sheet": sheet_name, "created": created, "updated": updated, "skipped": skipped}


def _parse_slip_block(frame, start_col):
    max_col = start_col + 2
    if frame.shape[1] <= max_col:
        return None

    name_value = frame.iat[1, start_col + 1] if frame.shape[0] > 1 else None
    if not str(name_value or "").strip():
        return None

    raw_values = {}
    payroll_group = frame.iat[3, start_col + 2] if frame.shape[0] > 3 else None

    for row_index in range(5, min(frame.shape[0], 40)):
        label = frame.iat[row_index, start_col + 1]
        value = frame.iat[row_index, start_col + 2]
        norm_label = _normalize_text(label)
        if norm_label:
            raw_values[norm_label] = value

    return {
        "name": str(name_value).strip(),
        "payroll_group": str(payroll_group).strip() if payroll_group is not None else None,
        "values": raw_values,
    }


def _upsert_slip_rows(workbook, month_key, source_name, by_code, by_name):
    sheet_name = _find_sheet_name(workbook, "PHIEU LUONG")
    if not sheet_name:
        return {"sheet": None, "created": 0, "updated": 0, "skipped": 0, "tax_updates": 0}

    frame = workbook.parse(sheet_name=sheet_name, header=None).fillna("")

    blocks = []
    for start_col in (0, 4, 8, 12):
        block = _parse_slip_block(frame, start_col)
        if block:
            blocks.append(block)

    created = 0
    updated = 0
    skipped = 0
    tax_updates = 0

    for block in blocks:
        employee = _resolve_employee(by_code, by_name, name_value=block["name"])
        if not employee:
            skipped += 1
            continue

        values_map = block["values"]

        def _clamp_small(value, max_abs=9999.99):
            try:
                v = float(value) if value is not None else 0.0
            except Exception:
                v = 0.0
            if abs(v) > max_abs:
                return max_abs if v >= 0 else -max_abs
            return v

        values = {
            "payroll_group": block["payroll_group"],
            "attendance_days": _clamp_small(_safe_number(values_map.get("ngaydilam"))),
            "leave_used_days": _clamp_small(_safe_number(values_map.get("songayphepsudungthang"))),
            "leave_remaining_days": _clamp_small(_safe_number(values_map.get("songayphepnamton"))),
            "salary_by_attendance": _safe_number(values_map.get("luongngaycongthang")),
            "overtime_weekday_hours": _clamp_small(_safe_number(values_map.get("tonggiotangcathuong"))),
            "overtime_sunday_hours": _clamp_small(_safe_number(values_map.get("tonggiotangcachunhat"))),
            "overtime_pay": _safe_number(values_map.get("tongtientangca")),
            "role_allowance": _safe_number(values_map.get("boiduongchucvutrachnhiem")),
            "child_allowance": _safe_number(values_map.get("boiduongphucapnuoiconnho6tuoi")),
            "transport_phone_allowance": _safe_number(values_map.get("tienxangdienthoaithuong")),
            "meal_allowance": _safe_number(values_map.get("tiencom")),
            "attendance_allowance": _safe_number(values_map.get("tienchuyencan")),
            "gross_total": _safe_number(values_map.get("tongluong")),
            "social_insurance_deduction": _safe_number(values_map.get("trubhxh105")),
            "union_fee_deduction": _safe_number(values_map.get("tiencongdoan")),
            "pit_tax_deduction": _safe_number(values_map.get("thuetncn")),
            "advance_deduction": _safe_number(values_map.get("trutientamung")),
            "net_income": _safe_number(values_map.get("luongthuclanh")),
            "source_file": source_name,
            "extra_data": values_map,
        }

        existing = PayrollSlip.query.filter_by(employee_id=employee.id, month_key=month_key).first()
        if existing:
            for key, value in values.items():
                setattr(existing, key, value)
            updated += 1
        else:
            db.session.add(PayrollSlip(employee_id=employee.id, month_key=month_key, **values))
            created += 1

        tax = PayrollTaxContribution.query.filter_by(employee_id=employee.id, month_key=month_key).first()
        tax_values = {
            "pit_tax": values["pit_tax_deduction"],
            "source_file": source_name,
            "notes": "Auto from PHIEU LUONG",
        }
        if tax:
            for key, value in tax_values.items():
                setattr(tax, key, value)
        else:
            db.session.add(
                PayrollTaxContribution(
                    employee_id=employee.id,
                    month_key=month_key,
                    **tax_values,
                )
            )
        tax_updates += 1

    return {
        "sheet": sheet_name,
        "created": created,
        "updated": updated,
        "skipped": skipped,
        "tax_updates": tax_updates,
    }


def import_salary_detail_file(
    file_path,
    source_name,
    actor="system",
    target_month=None,
    replace_existing=False,
):
    if not target_month:
        raise ValueError("Can chon thang de import du lieu luong chi tiet")

    extension = os.path.splitext(file_path)[1].lower()
    if extension not in {".xlsx", ".xls", ".xlsm", ".xltx", ".xltm"}:
        return {"detected": False, "reason": "not_excel"}

    workbook = pd.ExcelFile(file_path)
    normalized_sheet_names = {_normalize_text(name) for name in workbook.sheet_names}
    has_detail_sheets = any(
        key in normalized_sheet_names
        for key in {
            _normalize_text("phep nam"),
            _normalize_text("PHIEU LUONG"),
            _normalize_text("BAO HIEM"),
        }
    )

    if not has_detail_sheets:
        return {"detected": False, "reason": "detail_sheets_not_found"}

    by_code, by_name = _build_employee_maps()

    if replace_existing:
        PayrollLeaveSnapshot.query.filter_by(month_key=target_month).delete(synchronize_session=False)
        PayrollSlip.query.filter_by(month_key=target_month).delete(synchronize_session=False)
        PayrollInsuranceContribution.query.filter_by(month_key=target_month).delete(
            synchronize_session=False
        )
        PayrollTaxContribution.query.filter_by(month_key=target_month).delete(synchronize_session=False)

    leave_result = _upsert_leave_snapshots(workbook, target_month, source_name, by_code, by_name)
    slip_result = _upsert_slip_rows(workbook, target_month, source_name, by_code, by_name)
    insurance_result = _upsert_insurance_rows(workbook, target_month, source_name, by_code, by_name)

    batch_id = str(uuid.uuid4())
    log_action(
        "salary_detail_import",
        batch_id,
        "IMPORT",
        changed_by=actor,
        after_data={
            "month_key": target_month,
            "source_file": source_name,
            "leave": leave_result,
            "slips": slip_result,
            "insurance": insurance_result,
        },
        notes="Import luong chi tiet: phep nam, phieu luong, BHXH, thue",
    )

    return {
        "detected": True,
        "batch_id": batch_id,
        "month_key": target_month,
        "leave": leave_result,
        "slips": slip_result,
        "insurance": insurance_result,
    }


def import_salary_file(
    file_path,
    source_name,
    actor="system",
    target_month=None,
    default_company_work_days=26.0,
    replace_existing=False,
    blocked_month_keys=None,
):
    if not target_month:
        raise ValueError("Can chon thang de import he luong")

    blocked_month_keys = {item for item in (blocked_month_keys or []) if item}
    if target_month in blocked_month_keys:
        raise ValueError(f"Khong the import he luong vi thang {target_month} da chot so")

    frame, source_info = _read_frame(file_path)
    columns = _find_columns(frame)

    entries = []
    workday_candidates = list(source_info.get("workday_candidates") or [])

    for _, row in frame.iterrows():
        employee_code = _normalize_employee_code(row.get(columns["employee_code"]))
        monthly_wage = None

        if columns["monthly_wage"] is not None:
            monthly_wage = _to_float(row.get(columns["monthly_wage"]))

        if monthly_wage is None and columns["base_wage"] is not None:
            base_wage = _to_float(row.get(columns["base_wage"]))
            allowance = None
            if columns["allowance"] is not None:
                allowance = _to_float(row.get(columns["allowance"]))

            if base_wage is not None or allowance is not None:
                monthly_wage = float(base_wage or 0) + float(allowance or 0)

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
            "sheet_name": source_info.get("sheet_name"),
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
