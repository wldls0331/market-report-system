"""Editable email copy for the Weekly Bunkering Market Report."""

from __future__ import annotations

import datetime as dt

SUBJECT_TEMPLATE = "Weekly Bunkering Market Report - {report_long}"

BODY_TEMPLATE = """Dear All,

Please find attached the Weekly Bunkering Market Report as of {report_long}.

Best regards,
"""

_MONTH_ABBR = (
    "",
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
)


def format_report_long(report_date: dt.date) -> str:
    return f"{report_date.day} {_MONTH_ABBR[report_date.month]} {report_date.year}"


def default_subject(report_date: dt.date) -> str:
    return SUBJECT_TEMPLATE.format(report_long=format_report_long(report_date))


def default_body(report_date: dt.date) -> str:
    return BODY_TEMPLATE.format(report_long=format_report_long(report_date))
