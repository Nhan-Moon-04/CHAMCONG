from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Optional


NU_SHIFT_CODE = "NU"
NU_MORNING_MODE = "morning"
NU_NIGHT_MODE = "night"

NU_DYNAMIC_SHIFT_CODES = {"NU", "NUT1", "NUT2", "NU1", "NU2", "NU3", "NUN"}

NU_STANDARD_HOURS = 8.0
NU_MORNING_DEFAULT_OT_HOURS = 3.5
NU_NIGHT_DEFAULT_OT_HOURS = 4.0
NU_MORNING_MEAL_ALLOWANCE = 35000.0
NU_MORNING_MEAL_ALLOWANCE_OT_BONUS = 35000.0
NU_NIGHT_MEAL_ALLOWANCE = 135000.0

NU_OT_HALF_HOUR_FROM_MINUTES = 20.0
NU_OT_FULL_HOUR_FROM_MINUTES = 40.0


NU_EXTRA_OT_BY_CODE = {
    "NUT1": 1.0,
    "NUT2": 2.0,
}

NU_STANDARD_HOURS_DEDUCTION_BY_CODE = {
    "NU1": 1.0,
    "NU2": 2.0,
    "NU3": 3.0,
    "NUN": 4.0,
}

NU_WARNING_NOTE_PREFIX = "Canh bao NU:"


@dataclass
class NuShiftDayResult:
    mode: str
    week_mode: str
    shift_code: str
    has_midday_check: bool
    warning_note: Optional[str]
    check_in: Optional[datetime]
    check_out: Optional[datetime]
    standard_hours: float
    default_overtime_hours: float
    meal_allowance: float
    shift_name: str


def _normalize_employee_code(value):
    return str(value or "").replace("'", "").strip()


def is_nu_dynamic_shift_code(code):
    return str(code or "").strip().upper() in NU_DYNAMIC_SHIFT_CODES


def _is_midday_check(event_time):
    return 10 <= event_time.hour <= 13


def _is_evening_check(event_time):
    return event_time.hour >= 17


def _is_morning_check(event_time):
    return event_time.hour <= 8


def _detect_daily_mode(today_events, next_day_events):
    has_midday = any(_is_midday_check(item) for item in today_events)
    if has_midday:
        return NU_MORNING_MODE, has_midday

    has_evening = any(_is_evening_check(item) for item in today_events)
    has_next_day_morning = any(_is_morning_check(item) for item in next_day_events)
    if has_evening and has_next_day_morning:
        return NU_NIGHT_MODE, has_midday

    has_morning = any(_is_morning_check(item) for item in today_events)
    if has_morning and has_evening:
        return NU_MORNING_MODE, has_midday

    return None, has_midday


def _fallback_mode(today_events):
    if not today_events:
        return NU_MORNING_MODE

    first_event = today_events[0]
    if first_event.hour <= 10:
        return NU_MORNING_MODE
    if first_event.hour >= 15:
        return NU_NIGHT_MODE

    if any(_is_evening_check(item) for item in today_events):
        return NU_NIGHT_MODE

    return NU_MORNING_MODE


def _pick_check_times(mode, today_events, next_day_events):
    if mode == NU_MORNING_MODE:
        morning_candidates = [item for item in today_events if item.hour < 14]
        evening_candidates = [item for item in today_events if item.hour >= 14]

        check_in = morning_candidates[0] if morning_candidates else (today_events[0] if today_events else None)
        check_out = (
            evening_candidates[-1]
            if evening_candidates
            else (today_events[-1] if today_events else None)
        )
        return check_in, check_out

    # Night mode: check-in belongs to the current date evening, check-out belongs to next day.
    evening_candidates = [item for item in today_events if item.hour >= 15]
    next_day_candidates = [item for item in next_day_events if item.hour <= 12]

    check_in = evening_candidates[0] if evening_candidates else (today_events[-1] if today_events else None)
    check_out = next_day_candidates[0] if next_day_candidates else None

    if check_in and check_out and check_out <= check_in:
        return check_in, None

    return check_in, check_out


def _build_shift_name(mode, shift_code):
    mode_label = "sang" if mode == NU_MORNING_MODE else "toi"
    code = (shift_code or NU_SHIFT_CODE).upper()

    if code == "NUT1":
        return f"Ca nu {mode_label} +1h OT (NUT1)"
    if code == "NUT2":
        return f"Ca nu {mode_label} +2h OT (NUT2)"
    if code == "NU1":
        return f"Ca nu {mode_label} tru 1h cong (NU1)"
    if code == "NU2":
        return f"Ca nu {mode_label} tru 2h cong (NU2)"
    if code == "NU3":
        return f"Ca nu {mode_label} tru 3h cong (NU3)"
    if code == "NUN":
        return f"Ca nu {mode_label} tru 4h cong (NUN)"

    return f"Ca nu {mode_label} (NU)"


def _hours_between(start_at, end_at):
    if not start_at or not end_at or end_at <= start_at:
        return 0.0
    return (end_at - start_at).total_seconds() / 3600.0


def _normalize_nu_overtime_hours(raw_overtime_hours):
    if raw_overtime_hours <= 0:
        return 0.0

    total_minutes = raw_overtime_hours * 60.0
    whole_hours = int(total_minutes // 60.0)
    remainder_minutes = round(total_minutes - (whole_hours * 60.0), 4)

    # OT is paid only on .0 or .5 boundaries.
    # Examples: 1h19 -> 1h, 1h20 -> 1h30, 14h35 -> 14h30.
    if remainder_minutes < NU_OT_HALF_HOUR_FROM_MINUTES:
        return float(whole_hours)

    if remainder_minutes < NU_OT_FULL_HOUR_FROM_MINUTES:
        return whole_hours + 0.5

    return whole_hours + 1.0


def _compute_dynamic_nu_overtime_hours(check_in, check_out):
    worked_hours = _hours_between(check_in, check_out)
    if worked_hours <= 0:
        return None

    raw_overtime_hours = max(worked_hours - NU_STANDARD_HOURS, 0.0)
    return _normalize_nu_overtime_hours(raw_overtime_hours)


def _build_result(mode, week_mode, shift_code, has_midday_check, warning_note, check_in, check_out):
    code = (shift_code or NU_SHIFT_CODE).upper()
    standard_hours = max(
        NU_STANDARD_HOURS - NU_STANDARD_HOURS_DEDUCTION_BY_CODE.get(code, 0.0),
        0.0,
    )
    base_overtime = (
        NU_NIGHT_DEFAULT_OT_HOURS if mode == NU_NIGHT_MODE else NU_MORNING_DEFAULT_OT_HOURS
    )
    dynamic_overtime = _compute_dynamic_nu_overtime_hours(check_in, check_out)
    overtime_hours = (
        dynamic_overtime if dynamic_overtime is not None else base_overtime
    )
    overtime_hours += NU_EXTRA_OT_BY_CODE.get(code, 0.0)

    if mode == NU_NIGHT_MODE:
        meal_allowance = NU_NIGHT_MEAL_ALLOWANCE
    else:
        meal_allowance = NU_MORNING_MEAL_ALLOWANCE
        if overtime_hours >= 3.0:
            meal_allowance += NU_MORNING_MEAL_ALLOWANCE_OT_BONUS

    return NuShiftDayResult(
        mode=mode,
        week_mode=week_mode,
        shift_code=code,
        has_midday_check=has_midday_check,
        warning_note=warning_note,
        check_in=check_in,
        check_out=check_out,
        standard_hours=standard_hours,
        default_overtime_hours=overtime_hours,
        meal_allowance=meal_allowance,
        shift_name=_build_shift_name(mode, code),
    )


def is_nu_warning_note(notes):
    text_value = str(notes or "").lower()
    return NU_WARNING_NOTE_PREFIX.lower() in text_value


def build_nu_shift_day_results(
    nu_shift_code_map,
    employee_id_by_code,
    attendance_log_rows,
):
    events_by_employee_date = defaultdict(list)

    for row in attendance_log_rows:
        employee_code = _normalize_employee_code(getattr(row, "employee_code", ""))
        employee_id = employee_id_by_code.get(employee_code)
        event_time = getattr(row, "event_time", None)

        if employee_id is None or not isinstance(event_time, datetime):
            continue

        events_by_employee_date[(employee_id, event_time.date())].append(event_time)

    for event_list in events_by_employee_date.values():
        event_list.sort()

    work_dates_by_employee = defaultdict(set)
    for key, shift_code in nu_shift_code_map.items():
        employee_id, work_date = key
        if is_nu_dynamic_shift_code(shift_code):
            work_dates_by_employee[employee_id].add(work_date)

    results = {}

    for employee_id, work_dates in work_dates_by_employee.items():
        if not work_dates:
            continue

        sorted_dates = sorted(work_dates)
        day_mode_candidates = {}
        week_to_modes = defaultdict(list)

        for work_date in sorted_dates:
            today_events = events_by_employee_date.get((employee_id, work_date), [])
            next_day_events = events_by_employee_date.get((employee_id, work_date + timedelta(days=1)), [])

            detected_mode, has_midday = _detect_daily_mode(today_events, next_day_events)
            fallback_mode = _fallback_mode(today_events)
            day_mode_candidates[work_date] = {
                "detected_mode": detected_mode,
                "fallback_mode": fallback_mode,
                "has_midday": has_midday,
                "today_events": today_events,
                "next_day_events": next_day_events,
            }

            week_key = (work_date.isocalendar().year, work_date.isocalendar().week)
            if detected_mode:
                week_to_modes[week_key].append(detected_mode)

        week_mode_map = {}
        for work_date in sorted_dates:
            week_key = (work_date.isocalendar().year, work_date.isocalendar().week)
            if week_key in week_mode_map:
                continue

            detected_modes = week_to_modes.get(week_key, [])
            if detected_modes:
                week_mode_map[week_key] = Counter(detected_modes).most_common(1)[0][0]
                continue

            week_days = [
                item
                for item in sorted_dates
                if (item.isocalendar().year, item.isocalendar().week) == week_key
            ]
            fallback_modes = [day_mode_candidates[item]["fallback_mode"] for item in week_days]
            week_mode_map[week_key] = Counter(fallback_modes).most_common(1)[0][0]

        for work_date in sorted_dates:
            shift_code = str(nu_shift_code_map.get((employee_id, work_date), NU_SHIFT_CODE)).upper()
            data = day_mode_candidates[work_date]
            detected_mode = data["detected_mode"]
            week_key = (work_date.isocalendar().year, work_date.isocalendar().week)
            week_mode = week_mode_map.get(week_key) or data["fallback_mode"]
            effective_mode = week_mode
            if work_date.weekday() == 6 and week_mode == NU_MORNING_MODE:
                # Sunday is the transition point from morning week to night mode.
                effective_mode = NU_NIGHT_MODE

            today_events = data["today_events"]
            next_day_events = data["next_day_events"]
            has_midday = data["has_midday"]

            warning_parts = []
            if detected_mode and detected_mode != effective_mode:
                warning_parts.append("Lech mode tuan")

            if (
                week_mode == NU_MORNING_MODE
                and work_date.weekday() != 6
                and today_events
                and not has_midday
            ):
                warning_parts.append("Tuan sang thieu check giua ca (10h-13h)")

            warning_note = None
            if warning_parts:
                warning_note = f"{NU_WARNING_NOTE_PREFIX} {'; '.join(warning_parts)}"

            check_in, check_out = _pick_check_times(effective_mode, today_events, next_day_events)

            results[(employee_id, work_date)] = _build_result(
                effective_mode,
                week_mode,
                shift_code,
                has_midday_check=has_midday,
                warning_note=warning_note,
                check_in=check_in,
                check_out=check_out,
            )

    return results
