"""HTML report preview payload from INPUT + dated YY.MM.DD 보고자료."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from config.cell_mapping import (
    COMMENT_FIELDS,
    KOREA_WORLDWIDE_COMMENT_KEY,
    STRATEGY_FIELDS,
    WORLDWIDE_CHART_INPUT_KEYS,
)
from config.email_config import default_body, default_subject
from config.google_config import google_sheet_url
from config.input_sheet import AUTO_PREV_WEEK, AUTO_THIS_WEEK, AUTO_TWO_WEEKS, INPUT_NUMBER_LABELS, PREMIUM_INPUT_KEYS
from config.paths import OUTPUT_DIR, PROJECT_ROOT, session_output_dir
from services.chart_service import korea_premium_chart, worldwide_vlsfo_chart
from services.excel_service import GenerateError, generate_market_report, output_excel_path, parse_optional_number
from services.supplier_premium_service import format_premium_signed, parse_supplier_premium
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
        if path.is_file() and path.stat().st_size > 0:
            return path
    except OSError:
        return None
    return None


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


def _supplier_premiums(inputs: dict[str, Any]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for key in PREMIUM_INPUT_KEYS:
        raw = str(inputs.get(key) or "").strip()
        parsed = parse_supplier_premium(raw)
        rows.append(
            {
                "key": key,
                "label": INPUT_NUMBER_LABELS.get(key, key),
                "input": raw or "TBN",
                "vlsfo": format_premium_signed(parsed["vlsfo"]),
                "hsfo": format_premium_signed(parsed["hsfo"]),
                "vlsfo_value": parsed["vlsfo"],
                "hsfo_value": parsed["hsfo"],
            }
        )
    return rows


def _overlay_this_week_worldwide(chart: dict[str, Any], inputs: dict[str, Any]) -> dict[str, Any]:
    series = list(chart.get("series") or [])
    for index, key in enumerate(WORLDWIDE_CHART_INPUT_KEYS):
        if index >= len(series):
            break
        data = list(series[index].get("data") or [])
        if not data:
            continue
        try:
            data[-1] = parse_optional_number(inputs.get(key))
        except Exception:
            data[-1] = None
        series[index] = {**series[index], "data": data}
    chart["series"] = series
    return chart


def _overlay_this_week_premium(chart: dict[str, Any], inputs: dict[str, Any]) -> dict[str, Any]:
    series = list(chart.get("series") or [])
    for index, key in enumerate(PREMIUM_INPUT_KEYS):
        if index >= len(series):
            break
        data = list(series[index].get("data") or [])
        if not data:
            continue
        data[-1] = parse_supplier_premium(inputs.get(key))["vlsfo"]
        series[index] = {**series[index], "data": data}
    chart["series"] = series
    return chart


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
    }


def _load_report_grid(report_date):
    from services.sheets_cache import GRID_TTL_SECONDS, cache_get, cache_set

    key = f"report_grid:{report_date.isoformat()}"
    cached = cache_get(key)
    if cached is not None:
        return cached
    spreadsheet = open_spreadsheet()
    worksheet = find_report_worksheet(spreadsheet, report_date)
    if worksheet is None:
        result = (None, None)
        cache_set(key, result, GRID_TTL_SECONDS)
        return result
    grid = worksheet.get(REPORT_GRID_RANGE) or []
    result = (worksheet.title, grid)
    cache_set(key, result, GRID_TTL_SECONDS)
    return result


def connection_status() -> dict[str, Any]:
    from services.perf import timed

    with timed("gmail status"):
        gmail_ok, gmail_label = gmail_status()
    with timed("google sheets status"):
        sheets_ok, sheets_label = sheets_status()
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


def build_preview(*, sync_sheet: bool = False, include_status: bool = True) -> dict[str, Any]:
    from services.perf import timed
    from services.sheets_cache import cache_clear

    if sync_sheet:
        cache_clear()

    status = None
    if include_status:
        with timed("connection status"):
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
    if include_status and status and not status["sheets_ok"]:
        payload["error"] = "Google Sheets is not connected."
        return payload

    try:
        if sync_sheet:
            with timed("input sheet sync"):
                ensure_input_sheet(refresh_auto=True)
                record = load_input_data(refresh_auto=True, bypass_cache=True)
        else:
            record = load_input_data(refresh_auto=False)
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

    inputs = dict(record.inputs)
    korea_world = str(inputs.get(KOREA_WORLDWIDE_COMMENT_KEY) or inputs.get("comment_korea") or "").strip()
    comments = {
        "korea_worldwide": _tbn(korea_world),
        "singapore": _comment_from_inputs(inputs, "comment_singapore"),
        "china": _comment_from_inputs(inputs, "comment_china"),
        "japan": _comment_from_inputs(inputs, "comment_japan"),
        "strategy": [
            _tbn(inputs.get(item.key)) if str(inputs.get(item.key) or "").strip() else "TBN"
            for item in STRATEGY_FIELDS
        ],
    }
    payload["supplier_premiums"] = _supplier_premiums(inputs)

    grid_date = report_date
    sheet_title = format_sheet_name(report_date)
    title, grid = (None, None)
    try:
        with timed("chart data build"):
            title, grid = _load_report_grid(report_date)
    except Exception:
        title, grid = None, None
    exists = grid is not None
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
            "korea_premium": _overlay_this_week_premium(korea_premium_chart(getter), inputs),
            "worldwide_vlsfo": _overlay_this_week_worldwide(worldwide_vlsfo_chart(getter), inputs),
        }
        if exists or sync_sheet:
            comments["strategy"] = [
                _tbn(_grid_value(grid, item.cell)) for item in STRATEGY_FIELDS
            ]

    korea_chart = payload["charts"].get("korea_premium") or _empty_charts()["korea_premium"]
    if not korea_chart.get("series"):
        korea_chart = {
            "title": "Korea Major 4 Refiners - VLSFO Premium Trends",
            "labels": ["This week"],
            "series": [
                {"name": INPUT_NUMBER_LABELS.get(key, key), "data": [None]}
                for key in PREMIUM_INPUT_KEYS
            ],
        }
    payload["charts"]["korea_premium"] = _overlay_this_week_premium(korea_chart, inputs)
    world_chart = payload["charts"].get("worldwide_vlsfo") or _empty_charts()["worldwide_vlsfo"]
    if not world_chart.get("series"):
        world_chart = {
            "title": "Worldwide Ports - VLSFO Bunker Price Trend",
            "labels": ["This week"],
            "series": [
                {"name": name, "data": [None]}
                for name in ("Korea-South", "Singapore", "Nagoya", "Zhoushan")
            ],
        }
    payload["charts"]["worldwide_vlsfo"] = _overlay_this_week_worldwide(world_chart, inputs)

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
        "report_date_display": format_data_reference_display(report_date),
        "report_title": "WEEKLY BUNKERING REPORT",
        "report_title_full": format_report_title(report_date),
        "updated_display": format_data_reference_display(record.data_reference_date),
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
    payload["supplier_premiums"] = payload.get("supplier_premiums") or _supplier_premiums(inputs)
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
    """Re-read Google Sheets, write the dated tab, then generate Excel (no Excel-COM PDF)."""
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
            export_pdf=False,
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

    payload["files"] = _file_status(record.report_date)
    if result.excel_path and Path(result.excel_path).exists():
        payload["files"]["excel"] = Path(result.excel_path).name
        payload["files"]["has_excel"] = True
    payload["email"]["attachment"] = payload["files"].get("pdf")
    payload["warnings"] = list(result.warnings or [])
    payload["pdf_page_count"] = None
    return payload


def ensure_excel_file() -> Path:
    """Return the dated Excel workbook, generating it if it is missing."""
    record = load_input_data()
    existing = _safe_output_file(output_excel_path(record.report_date))
    if existing is not None:
        return existing

    extra_cells = None
    try:
        sheet_record = load_report_data(record.report_date)
        extra_cells = sheet_record.extra_cells if sheet_record else None
    except Exception:
        extra_cells = None

    result = generate_market_report(
        record.report_date,
        record.inputs,
        export_pdf=False,
        pricing_month=_pricing_month(record),
        extra_cells=extra_cells,
        data_reference_date=record.data_reference_date,
    )
    path = Path(result.excel_path) if result.excel_path else output_excel_path(record.report_date)
    saved = _safe_output_file(path)
    if saved is None:
        raise GenerateError(
            "Excel workbook was not saved. "
            f"Expected path: {path} (exists={path.exists()})"
        )
    return saved


def build_print_html(preview: dict[str, Any] | None = None) -> str:
    """Render print.html without Flask so Streamlit can generate the same HTML PDF."""
    import json

    from jinja2 import Environment, FileSystemLoader
    from markupsafe import Markup

    if preview is None:
        preview = public_preview(build_preview(sync_sheet=False, include_status=False))

    env = Environment(
        loader=FileSystemLoader(str(PROJECT_ROOT / "web" / "templates")),
        autoescape=True,
    )
    env.filters["tojson"] = lambda value: Markup(
        json.dumps(value).replace("<", "\\u003c")
    )
    env.globals["url_for"] = lambda endpoint, filename="", **_kwargs: f"/static/{filename}"
    return env.get_template("print.html").render(preview=preview)


def render_html_preview_pdf(print_url: str | None = None, html: str | None = None) -> Path:
    """Render the live Report Preview HTML to the dated PDF filename."""
    from services.html_pdf_service import preview_pdf_path, render_preview_pdf

    record = load_input_data()
    return render_preview_pdf(
        output_path=preview_pdf_path(record.report_date),
        print_url=print_url or "http://127.0.0.1:8502/print",
        html=html if html is not None else build_print_html(),
        static_root=PROJECT_ROOT / "web" / "static",
    )
