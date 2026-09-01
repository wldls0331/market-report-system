"""HTML report preview payload from INPUT + dated YY.MM.DD 보고자료."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from config.cell_mapping import (
    COMMENT_FIELDS,
    KOREA_COMMENT_SOURCE_KEY,
    KOREA_WORLDWIDE_COMMENT_KEY,
    STRATEGY_FIELDS,
)
from config.email_config import default_body, default_subject
from config.google_config import google_sheet_url
from config.input_sheet import AUTO_PREV_WEEK, AUTO_THIS_WEEK, AUTO_TWO_WEEKS
from config.paths import OUTPUT_DIR, session_output_dir
from services.chart_service import korea_premium_chart, spread_trend_chart, worldwide_vlsfo_chart
from services.excel_service import GenerateError, generate_market_report, output_excel_path
from services.gmail_service import gmail_status
from services.google_sheets_service import (
    DataServiceError,
    ReportData,
    _grid_value,
    _typed_cell,
    active_recipient_emails,
    find_report_worksheet,
    latest_report_sheet,
    load_email_recipients_result,
    load_report_data,
    open_spreadsheet,
    sheets_status,
    spreadsheet_title,
)
from services.input_sheet_service import (
    create_or_update_report_from_input,
    dated_report_exists,
    ensure_input_sheet,
    load_input_data,
)
from services.pricing_month_service import YearMonth, default_pricing_month
from services.working_day_service import (
    format_data_reference_display,
    format_pdf_filename,
    format_report_title,
    format_sheet_name,
    singapore_holiday_name,
    validate_report_date,
)

REPORT_GRID_RANGE = "A1:AI80"


def _safe_output_file(path: Path) -> Path | None:
    try:
        resolved = path.resolve()
        resolved.relative_to(OUTPUT_DIR.resolve())
    except Exception:
        return None
    return resolved if resolved.is_file() else None


def output_paths_for_date(report_date) -> dict[str, Path | None]:
    excel = _safe_output_file(output_excel_path(report_date))
    pdf = _safe_output_file(session_output_dir() / format_pdf_filename(report_date))
    return {"excel": excel, "pdf": pdf}


def resolve_output_file(kind: str) -> Path | None:
    """Resolve PDF/Excel from Google Sheets INPUT report date. No process-global cache."""
    if kind not in {"pdf", "excel"}:
        return None
    record = load_input_data()
    files = output_paths_for_date(record.report_date)
    return files.get(kind)


def _file_status(report_date) -> dict[str, Any]:
    files = output_paths_for_date(report_date)
    excel = files.get("excel")
    pdf = files.get("pdf")
    return {
        "excel": excel.name if excel else None,
        "pdf": pdf.name if pdf else None,
        "has_excel": excel is not None,
        "has_pdf": pdf is not None,
    }


def _tbn(value: Any) -> str:
    text = "" if value is None else str(value).strip()
    return text if text else "TBN"


def _comment_from_inputs(inputs: dict[str, Any], key: str) -> str:
    return _tbn(inputs.get(key))


def _pricing_label(record: ReportData) -> str:
    if record.pricing_month:
        try:
            return YearMonth.parse(record.pricing_month).label()
        except ValueError:
            pass
    return default_pricing_month(record.report_date).label()


def _pricing_month(record: ReportData) -> YearMonth:
    if record.pricing_month:
        try:
            return YearMonth.parse(record.pricing_month)
        except ValueError:
            pass
    return default_pricing_month(record.report_date)


def _date_warning(report_date) -> str | None:
    try:
        validate_report_date(report_date)
        return None
    except Exception as exc:
        if report_date.weekday() >= 5:
            return str(exc)
        if singapore_holiday_name(report_date):
            return str(exc)
        return str(exc)


def _grid_getter(grid: list[list[Any]]):
    def get_value(address: str) -> Any:
        return _typed_cell(address, _grid_value(grid, address))

    return get_value


def _empty_charts() -> dict[str, Any]:
    return {
        "korea_premium": {"title": "Korea Major 4 Refiners - VLSFO Premium Trends", "labels": [], "series": []},
        "worldwide_vlsfo": {"title": "Worldwide Ports - VLSFO Bunker Price Trend", "labels": [], "series": []},
        "spread": {
            "title": "SPREAD TREND (BUNKER WIRE - MOPS SINGAPORE 0.5%)",
            "labels": [],
            "series": [],
        },
    }


def _load_report_grid(report_date):
    spreadsheet = open_spreadsheet()
    worksheet = find_report_worksheet(spreadsheet, report_date)
    if worksheet is None:
        return None, None
    grid = worksheet.get(REPORT_GRID_RANGE) or []
    return worksheet.title, grid


def connection_status() -> dict[str, Any]:
    sheets_ok, sheets_label = sheets_status()
    gmail_ok, gmail_label = gmail_status()
    title = ""
    if sheets_ok:
        try:
            title = spreadsheet_title()
        except Exception:
            title = ""
    return {
        "sheets_ok": sheets_ok,
        "sheets": sheets_label,
        "gmail_ok": gmail_ok,
        "gmail": gmail_label,
        "sheet_url": google_sheet_url(),
        "spreadsheet": title,
    }


def build_preview(*, sync_sheet: bool = False) -> dict[str, Any]:
    status = connection_status()
    payload: dict[str, Any] = {
        "status": status,
        "meta": None,
        "comments": {},
        "charts": _empty_charts(),
        "email": {"to": "", "cc": "", "subject": "", "body": "", "attachment": None},
        "files": {"excel": None, "pdf": None},
        "error": None,
        "warning": None,
        "synced": False,
        "is_update": None,
    }
    if not status["sheets_ok"]:
        payload["error"] = "Google Sheets is not connected."
        return payload

    try:
        if sync_sheet:
            ensure_input_sheet(refresh_auto=True)
            record = load_input_data()
        else:
            record = load_input_data()
    except DataServiceError as exc:
        payload["error"] = str(exc)
        return payload
    except Exception as exc:
        payload["error"] = f"Could not load INPUT sheet: {exc}"
        return payload

    is_update = None
    if sync_sheet:
        try:
            _sheet_name, is_update = create_or_update_report_from_input(record)
            payload["synced"] = True
            payload["is_update"] = is_update
        except DataServiceError as exc:
            payload["warning"] = str(exc)
        except Exception as exc:
            payload["warning"] = f"Could not update the dated report sheet: {exc}"

    report_date = record.report_date
    extras = record.extra_cells or {}
    exists = False
    try:
        exists = dated_report_exists(report_date)
    except Exception:
        exists = False

    inputs = dict(record.inputs)
    worldwide = str(inputs.get(KOREA_WORLDWIDE_COMMENT_KEY, "") or "").strip()
    korea = _comment_from_inputs(inputs, KOREA_COMMENT_SOURCE_KEY)
    comments = {
        "korea": korea,
        "korea_worldwide": _tbn(worldwide or korea),
        "singapore": _comment_from_inputs(inputs, "comment_singapore"),
        "china": _comment_from_inputs(inputs, "comment_china"),
        "japan": _comment_from_inputs(inputs, "comment_japan"),
        "strategy": [
            _tbn(inputs.get(item.key)) if str(inputs.get(item.key) or "").strip() else "TBN"
            for item in STRATEGY_FIELDS
        ],
    }

    grid_date = report_date
    sheet_title = format_sheet_name(report_date)
    title, grid = (None, None)
    try:
        title, grid = _load_report_grid(report_date)
    except Exception:
        title, grid = None, None
    if grid is None:
        latest = None
        try:
            latest = latest_report_sheet()
        except Exception:
            latest = None
        if latest is not None:
            grid_date = latest[0]
            try:
                title, grid = _load_report_grid(grid_date)
            except Exception:
                title, grid = None, None
            if title:
                sheet_title = title
    elif title:
        sheet_title = title

    if grid:
        getter = _grid_getter(grid)
        payload["charts"] = {
            "korea_premium": korea_premium_chart(getter),
            "worldwide_vlsfo": worldwide_vlsfo_chart(getter),
            "spread": spread_trend_chart(getter),
        }
        if exists or sync_sheet:
            for item in COMMENT_FIELDS:
                text = "\n".join(
                    str(_grid_value(grid, cell) or "").strip()
                    for cell in item.cells
                    if str(_grid_value(grid, cell) or "").strip()
                )
                if item.key == "comment_korea":
                    comments["korea"] = _tbn(text)
                elif item.key == "comment_korea_worldwide":
                    comments["korea_worldwide"] = _tbn(text or comments["korea"])
                elif item.key == "comment_singapore":
                    comments["singapore"] = _tbn(text)
                elif item.key == "comment_china":
                    comments["china"] = _tbn(text)
                elif item.key == "comment_japan":
                    comments["japan"] = _tbn(text)
            comments["strategy"] = [
                _tbn(_grid_value(grid, item.cell)) for item in STRATEGY_FIELDS
            ]

    to_list, cc_list = [], []
    recipients_warning = None
    try:
        _recipients, recipients_warning = load_email_recipients_result()
        to_list, cc_list = active_recipient_emails()
    except DataServiceError as exc:
        recipients_warning = str(exc)
    if recipients_warning and not payload["warning"]:
        payload["warning"] = recipients_warning

    payload["meta"] = {
        "report_date": report_date.isoformat(),
        "report_title": format_report_title(report_date),
        "sheet_name": sheet_title,
        "pricing_month": _pricing_label(record),
        "data_reference_date": format_data_reference_display(record.data_reference_date),
        "this_week_friday": extras.get(AUTO_THIS_WEEK, ""),
        "previous_week_friday": extras.get(AUTO_PREV_WEEK, record.data_reference_date.isoformat()),
        "two_weeks_ago_friday": extras.get(AUTO_TWO_WEEKS, ""),
        "report_exists": exists or bool(sync_sheet and payload["synced"]),
        "date_ok": _date_warning(report_date) is None,
        "date_warning": _date_warning(report_date),
    }
    payload["comments"] = comments
    payload["email"] = {
        "to": "\n".join(to_list),
        "cc": "\n".join(cc_list),
        "subject": default_subject(report_date),
        "body": default_body(report_date),
        "attachment": None,
    }
    payload["files"] = _file_status(report_date)
    if payload["files"].get("has_pdf"):
        payload["email"]["attachment"] = payload["files"].get("pdf")
    payload["_record"] = record
    return payload


def public_preview(payload: dict[str, Any]) -> dict[str, Any]:
    """JSON-safe preview. Drops in-memory objects and absolute disk paths."""
    payload = dict(payload)
    payload.pop("_record", None)
    files = dict(payload.get("files") or {})
    excel_name = Path(str(files.get("excel") or "")).name or None
    pdf_name = Path(str(files.get("pdf") or "")).name or None
    has_excel = bool(files.get("has_excel"))
    has_pdf = bool(files.get("has_pdf"))
    report_date = (payload.get("meta") or {}).get("report_date")
    if report_date and (not has_excel or not has_pdf):
        try:
            import datetime as dt

            live = _file_status(dt.date.fromisoformat(str(report_date)))
            excel_name = excel_name or live.get("excel")
            pdf_name = pdf_name or live.get("pdf")
            has_excel = has_excel or bool(live.get("has_excel"))
            has_pdf = has_pdf or bool(live.get("has_pdf"))
        except Exception:
            pass
    payload["files"] = {
        "excel": excel_name,
        "pdf": pdf_name,
        "has_excel": has_excel,
        "has_pdf": has_pdf,
    }
    if has_pdf and payload.get("email") is not None:
        payload["email"]["attachment"] = pdf_name
    return payload


def create_report_files() -> dict[str, Any]:
    """Re-read Google Sheets, write the dated tab, then generate PDF."""
    payload = build_preview(sync_sheet=True)
    if payload.get("error"):
        return payload
    record: ReportData | None = payload.pop("_record", None)
    if record is None:
        payload["error"] = "INPUT data is not available."
        return payload
    warning = payload.get("meta", {}).get("date_warning")
    if warning:
        payload["error"] = warning
        return payload

    extra_cells = None
    try:
        sheet_record = load_report_data(record.report_date)
        if sheet_record is not None:
            extra_cells = sheet_record.extra_cells
    except Exception:
        extra_cells = None

    try:
        result = generate_market_report(
            record.report_date,
            record.inputs,
            export_pdf=True,
            pricing_month=_pricing_month(record),
            extra_cells=extra_cells,
            data_reference_date=record.data_reference_date,
        )
    except GenerateError as exc:
        payload["error"] = str(exc)
        return payload
    except Exception as exc:
        payload["error"] = f"Report generation failed: {exc}"
        return payload

    pdf_name = format_pdf_filename(record.report_date)
    pdf_path = str(result.pdf_path) if result.pdf_path else None
    if pdf_path and Path(pdf_path).name != pdf_name:
        pdf_path = None
    payload["files"] = _file_status(record.report_date)
    if pdf_path and Path(pdf_path).exists():
        payload["files"]["pdf"] = Path(pdf_path).name
        payload["files"]["has_pdf"] = True
    if result.excel_path and Path(result.excel_path).exists():
        payload["files"]["excel"] = Path(result.excel_path).name
        payload["files"]["has_excel"] = True
    payload["email"]["attachment"] = payload["files"].get("pdf")
    payload["warnings"] = list(result.warnings or [])
    payload["pdf_page_count"] = result.pdf_page_count
    return payload
