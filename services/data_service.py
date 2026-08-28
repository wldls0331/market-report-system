"""Persistence facade. Market data lives in Google Sheets (see google_sheets_service)."""

from __future__ import annotations

from services.google_sheets_service import (  # noqa: F401
    DataServiceError,
    ReportData,
    format_sgt_full,
    format_sgt_short,
    load_report_data,
    now_sgt,
    report_data_exists,
    sheets_status,
)

format_sgt = format_sgt_short
