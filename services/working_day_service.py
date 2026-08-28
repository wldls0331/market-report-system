from __future__ import annotations

import datetime as dt

import holidays

_SG_HOLIDAYS: dict[int, object] = {}


class WorkingDayError(ValueError):
    pass


def _singapore_holidays(year: int):
    cached = _SG_HOLIDAYS.get(year)
    if cached is None:
        cached = holidays.country_holidays("SG", years=year)
        _SG_HOLIDAYS[year] = cached
    return cached


def singapore_holiday_name(date: dt.date) -> str | None:
    return _singapore_holidays(date.year).get(date)


def is_singapore_working_day(date: dt.date) -> bool:
    if date.weekday() >= 5:
        return False
    return date not in _singapore_holidays(date.year)


def validate_report_date(date: dt.date) -> None:
    if date.weekday() >= 5:
        weekday = "Saturday" if date.weekday() == 5 else "Sunday"
        raise WorkingDayError(
            f"Report Date must be a Singapore working day. {date.isoformat()} is a {weekday}."
        )
    name = singapore_holiday_name(date)
    if name:
        raise WorkingDayError(
            f"Report Date must be a Singapore working day. "
            f"{date.isoformat()} is a Singapore public holiday ({name})."
        )


def previous_singapore_working_day(date: dt.date) -> dt.date:
    cursor = date - dt.timedelta(days=1)
    while not is_singapore_working_day(cursor):
        cursor -= dt.timedelta(days=1)
    return cursor


def week_starting_monday(date: dt.date) -> dt.date:
    return date - dt.timedelta(days=date.weekday())


def previous_week_last_working_day(report_date: dt.date) -> dt.date:
    """Last Singapore working day in the calendar week before `report_date`.

    Weeks start Monday. Start from the Sunday of the previous week and walk
    backward until a Singapore working day (Mon–Fri, not a public holiday).
    """
    previous_week_end = week_starting_monday(report_date) - dt.timedelta(days=1)
    cursor = previous_week_end
    while not is_singapore_working_day(cursor):
        cursor -= dt.timedelta(days=1)
    return cursor


def format_data_reference_display(date: dt.date) -> str:
    return f"{date.day} {date.strftime('%b %Y')}"


def format_sheet_name(date: dt.date) -> str:
    return date.strftime("%y.%m.%d") + " 보고자료"


def format_output_stem(date: dt.date) -> str:
    return date.strftime("%Y.%m.%d") + "_Market Report"


def format_report_title(date: dt.date) -> str:
    return date.strftime("%y.%m.%d") + "_WEEKLY BUNKERING REPORT"


def format_pdf_filename(date: dt.date) -> str:
    return f"Weekly Report_Bunkering_{date.strftime('%y.%m.%d')}.pdf"


def friday_of_week(date: dt.date) -> dt.date:
    return week_starting_monday(date) + dt.timedelta(days=4)


def last_working_day_of_week(date: dt.date) -> dt.date:
    """Friday of that week, or the last Singapore working day if Friday is off."""
    friday = friday_of_week(date)
    monday = week_starting_monday(date)
    cursor = friday
    while cursor >= monday:
        if is_singapore_working_day(cursor):
            return cursor
        cursor -= dt.timedelta(days=1)
    return previous_singapore_working_day(monday)


def weekly_fridays(report_date: dt.date) -> tuple[dt.date, dt.date, dt.date]:
    """This week, previous week, and two weeks ago (Singapore Friday working days)."""
    this_week = last_working_day_of_week(report_date)
    previous_week = last_working_day_of_week(report_date - dt.timedelta(days=7))
    two_weeks_ago = last_working_day_of_week(report_date - dt.timedelta(days=14))
    return this_week, previous_week, two_weeks_ago
