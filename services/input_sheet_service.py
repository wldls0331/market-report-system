"""Read/write the INPUT tab and copy values onto dated report sheets."""

from __future__ import annotations

import datetime as dt
from typing import Any

from config.cell_mapping import (
    COMMENT_FIELDS,
    DATE_CELLS,
    FORMULA_CELLS,
    KOREA_BUNKER_MARKET_CELLS,
    KOREA_COMMENT_SOURCE_KEY,
    KOREA_WORLDWIDE_COMMENT_KEY,
    NUMBER_FIELDS,
    SALES_ASSUMPTION_HEADER_CELL,
    SALES_ASSUMPTION_HEADER_TEMPLATE,
    STRATEGY_FIELDS,
    TITLE_CELL,
    WORLDWIDE_NEW_POINT_CELLS,
    WORLDWIDE_THIS_WEEK_CELLS,
)
from config.input_sheet import (
    AUTO_FORMULAS,
    AUTO_PREV_WEEK,
    AUTO_PRICING_MONTH,
    AUTO_REPORT_DATE,
    AUTO_THIS_WEEK,
    AUTO_TWO_WEEKS,
    INPUT_SHEET_NAME,
    OBSOLETE_INPUT_KEYS,
    PREMIUM_INPUT_HINT,
    PREMIUM_INPUT_KEYS,
    PREV_WEEK_FORMULA,
    PRICING_MONTH_FORMULA,
    THIS_WEEK_FORMULA,
    TWO_WEEKS_FORMULA,
    InputRow,
    input_layout,
)
from services.chart_service import shift_weekly_chart_window
from services.comment_service import TBN, compose_comment_lines, compose_strategy_text
from services.supplier_premium_service import parse_supplier_premium
from services.pricing_month_service import YearMonth, default_pricing_month
from services.working_day_service import (
    format_report_title,
    format_sheet_name,
    singapore_public_holiday_dates,
    weekly_fridays,
)
from services.google_sheets_service import (
    EMAIL_RECIPIENTS_SHEET,
    DataServiceError,
    ReportData,
    _grid_value,
    _join_comment,
    _parse_date_value,
    _stringify,
    find_report_worksheet,
    latest_report_sheet,
)


def _layout_rows() -> tuple[InputRow, ...]:
    return input_layout()


def _spreadsheet(*, require_write: bool = False):
    """Lazy import so Streamlit reloads cannot hit a half-loaded google_sheets_service."""
    from services.google_sheets_service import open_spreadsheet as connect_spreadsheet

    return connect_spreadsheet(require_write=require_write)


def _parse_input_date(value: Any) -> dt.date | None:
    parsed = _parse_date_value(value)
    if parsed is not None:
        return parsed
    text = _stringify(value)
    if not text:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%y.%m.%d"):
        try:
            return dt.datetime.strptime(text[:10], fmt).date()
        except ValueError:
            continue
    return None


def _row_map(worksheet) -> dict[str, int]:
    """Map INPUT field key -> 1-based row number using column D."""
    values = worksheet.get("A1:D80") or []
    found: dict[str, int] = {}
    for index, row in enumerate(values, start=1):
        key = _stringify(row[3] if len(row) > 3 else "")
        if key:
            found[key] = index
    return found


def _cell_b(worksheet, row: int) -> Any:
    values = worksheet.get(f"B{row}:B{row}") or []
    if not values or not values[0]:
        return None
    return values[0][0]


def _previous_hints(spreadsheet) -> dict[str, str]:
    latest = None
    try:
        latest = latest_report_sheet()
    except Exception:
        latest = None
    if latest is None:
        return {}
    worksheet = find_report_worksheet(spreadsheet, latest[0])
    if worksheet is None:
        return {}
    grid = worksheet.get("A1:AI80") or []
    hints: dict[str, str] = {}
    for item in NUMBER_FIELDS:
        text = _stringify(_grid_value(grid, item.cell))
        if text:
            hints[item.key] = f"Previous: {text}"
    for item in COMMENT_FIELDS:
        text = _join_comment(grid, item.cells)
        if text:
            hints[item.key] = f"Previous: {text[:80]}"
    for item in STRATEGY_FIELDS:
        text = _stringify(_grid_value(grid, item.cell))
        if text:
            hints[item.key] = f"Previous: {text[:80]}"
    return hints


def _format_input_sheet(spreadsheet, worksheet) -> None:
    sheet_id = worksheet.id
    requests = [
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 0,
                    "endRowIndex": 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": 3,
                },
                "cell": {
                    "userEnteredFormat": {
                        "textFormat": {"bold": True, "fontSize": 14},
                    }
                },
                "fields": "userEnteredFormat.textFormat",
            }
        },
        {
            "updateSheetProperties": {
                "properties": {
                    "sheetId": sheet_id,
                    "gridProperties": {"frozenRowCount": 1},
                },
                "fields": "gridProperties.frozenRowCount",
            }
        },
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": 0,
                    "endIndex": 1,
                },
                "properties": {"pixelSize": 220},
                "fields": "pixelSize",
            }
        },
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": 1,
                    "endIndex": 2,
                },
                "properties": {"pixelSize": 180},
                "fields": "pixelSize",
            }
        },
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": 2,
                    "endIndex": 3,
                },
                "properties": {"pixelSize": 420},
                "fields": "pixelSize",
            }
        },
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": 3,
                    "endIndex": 4,
                },
                "properties": {"hiddenByUser": True},
                "fields": "hiddenByUser",
            }
        },
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": 5,
                    "endIndex": 6,
                },
                "properties": {"hiddenByUser": True},
                "fields": "hiddenByUser",
            }
        },
    ]
    try:
        spreadsheet.batch_update({"requests": requests})
    except Exception:
        pass


def _build_input_grid(previous_hints: dict[str, str], existing: dict[str, Any]) -> list[list[Any]]:
    rows: list[list[Any]] = []
    for spec in _layout_rows():
        if spec.kind == "spacer":
            rows.append(["", "", "", ""])
            continue
        if spec.kind == "title":
            rows.append([spec.label, "", spec.hint, ""])
            continue
        if spec.kind == "section":
            rows.append([spec.label, "", "", ""])
            continue
        value = existing.get(spec.key, "")
        if spec.kind == "auto":
            value = AUTO_FORMULAS.get(spec.key, value)
        hint = spec.hint
        if spec.editable and spec.key in previous_hints:
            hint = f"{spec.hint}  {previous_hints[spec.key]}"
        rows.append([spec.label, value, hint, spec.key])
    return rows


def _parse_optional_number(raw: str | None) -> float | None:
    if raw is None:
        return None
    text = str(raw).strip()
    if not text or text.upper() == TBN:
        return None
    text = text.replace(",", "")
    try:
        return float(text)
    except ValueError:
        return None


def _required_keys() -> set[str]:
    return {row.key for row in _layout_rows() if row.key}


def ensure_input_sheet(*, refresh_auto: bool = True):
    spreadsheet = _spreadsheet(require_write=True)
    try:
        worksheet = spreadsheet.worksheet(INPUT_SHEET_NAME)
        created = False
    except Exception:
        worksheet = spreadsheet.add_worksheet(title=INPUT_SHEET_NAME, rows=250, cols=6)
        created = True

    mapping = {} if created else _row_map(worksheet)
    missing_keys = _required_keys() - set(mapping)
    if created or missing_keys or (OBSOLETE_INPUT_KEYS & set(mapping)):
        existing_values: dict[str, Any] = {}
        if mapping:
            values = worksheet.get("A1:D80") or []
            for key, row_number in mapping.items():
                row = values[row_number - 1] if row_number <= len(values) else []
                existing_values[key] = row[1] if len(row) > 1 else ""
        previous_hints = _previous_hints(spreadsheet)
        if existing_values.get("comment_korea") and not existing_values.get("comment_korea_worldwide"):
            existing_values["comment_korea_worldwide"] = existing_values["comment_korea"]
        grid = _build_input_grid(previous_hints, existing_values)
        worksheet.update("A1:D80", grid, value_input_option="USER_ENTERED")
        _format_input_sheet(spreadsheet, worksheet)

    worksheets = list(spreadsheet.worksheets())
    if worksheets and worksheets[0].title != INPUT_SHEET_NAME:
        others = [item for item in worksheets if item.title != INPUT_SHEET_NAME]
        try:
            spreadsheet.reorder_worksheets([worksheet] + others)
        except Exception:
            pass

    if refresh_auto:
        refresh_input_auto_fields(spreadsheet, worksheet)
    return spreadsheet, worksheet


def _write_holiday_list(worksheet) -> None:
    try:
        worksheet.resize(rows=max(worksheet.row_count, 250), cols=max(worksheet.col_count, 6))
    except Exception:
        pass
    years = list(range(2024, 2033))
    dates = singapore_public_holiday_dates(*years)
    rows: list[list[str]] = [["sg_holidays"]]
    rows.extend([[item.isoformat()] for item in dates])
    worksheet.update(range_name="F1:F" + str(len(rows)), values=rows, value_input_option="USER_ENTERED")


def refresh_input_auto_fields(spreadsheet=None, worksheet=None) -> None:
    if spreadsheet is None or worksheet is None:
        spreadsheet = _spreadsheet(require_write=True)
        try:
            worksheet = spreadsheet.worksheet(INPUT_SHEET_NAME)
        except Exception as exc:
            raise DataServiceError("INPUT sheet not found.") from exc
    mapping = _row_map(worksheet)
    report_row = mapping.get(AUTO_REPORT_DATE)
    raw_date = _cell_b(worksheet, report_row) if report_row else None
    report_date = _parse_input_date(raw_date)
    if report_date is None:
        latest = latest_report_sheet()
        if latest and report_row:
            worksheet.update(
                range_name=f"B{report_row}",
                values=[[latest[0].isoformat()]],
                value_input_option="USER_ENTERED",
            )
    _write_holiday_list(worksheet)
    updates = []
    for key, formula in (
        (AUTO_PRICING_MONTH, PRICING_MONTH_FORMULA),
        (AUTO_THIS_WEEK, THIS_WEEK_FORMULA),
        (AUTO_PREV_WEEK, PREV_WEEK_FORMULA),
        (AUTO_TWO_WEEKS, TWO_WEEKS_FORMULA),
    ):
        row = mapping.get(key)
        if row:
            updates.append({"range": f"B{row}", "values": [[formula]]})
    if updates:
        worksheet.batch_update(updates, value_input_option="USER_ENTERED")
    hint_updates = []
    for key in PREMIUM_INPUT_KEYS:
        row = mapping.get(key)
        if row:
            hint_updates.append({"range": f"C{row}", "values": [[PREMIUM_INPUT_HINT]]})
    if hint_updates:
        try:
            worksheet.batch_update(hint_updates, value_input_option="USER_ENTERED")
        except Exception:
            pass


def _read_input_values(worksheet) -> dict[str, str]:
    mapping = _row_map(worksheet)
    values = worksheet.get("A1:D80") or []
    result: dict[str, str] = {}
    for key, row_number in mapping.items():
        row = values[row_number - 1] if row_number <= len(values) else []
        result[key] = _stringify(row[1] if len(row) > 1 else "")
    return result


def load_input_data(*, refresh_auto: bool = False, bypass_cache: bool = False) -> ReportData:
    from services.google_sheets_service import _assert_google_ready_if_configured
    from services.perf import timed
    from services.sheets_cache import INPUT_TTL_SECONDS, cache_clear, cache_get, cache_set

    if bypass_cache or refresh_auto:
        cache_clear("input_data")
    elif not refresh_auto:
        cached = cache_get("input_data")
        if cached is not None:
            return cached

    _assert_google_ready_if_configured()
    try:
        with timed("input sheet read"):
            spreadsheet = _spreadsheet()
            try:
                worksheet = spreadsheet.worksheet(INPUT_SHEET_NAME)
            except Exception:
                spreadsheet, worksheet = ensure_input_sheet(refresh_auto=True)
            else:
                if refresh_auto:
                    try:
                        write_book = _spreadsheet(require_write=True)
                        write_ws = write_book.worksheet(INPUT_SHEET_NAME)
                        refresh_input_auto_fields(write_book, write_ws)
                        worksheet = write_ws
                    except Exception:
                        pass
            raw = _read_input_values(worksheet)
    except DataServiceError:
        raise
    except Exception as exc:
        detail = str(exc)
        if "insufficient" in detail.lower() or "403" in detail:
            raise DataServiceError(
                "Google Sheets write permission is required to manage the INPUT sheet. "
                "Authorize Google Sheets locally, then copy the new sheets_token.json "
                "refresh_token into GOOGLE_REFRESH_TOKEN."
            ) from exc
        raise DataServiceError(f"Could not load INPUT sheet: {exc}") from exc

    report_date = _parse_input_date(raw.get(AUTO_REPORT_DATE))
    if report_date is None:
        raise DataServiceError("INPUT Report Date is empty or invalid. Enter YYYY-MM-DD in column B.")
    this_week, prev_week, two_weeks = weekly_fridays(report_date)
    month = default_pricing_month(report_date)
    skip_keys = {
        AUTO_REPORT_DATE,
        AUTO_PRICING_MONTH,
        AUTO_THIS_WEEK,
        AUTO_PREV_WEEK,
        AUTO_TWO_WEEKS,
    }
    inputs = {key: value for key, value in raw.items() if key not in skip_keys}
    if not str(inputs.get(KOREA_WORLDWIDE_COMMENT_KEY) or "").strip() and str(inputs.get(KOREA_COMMENT_SOURCE_KEY) or "").strip():
        inputs[KOREA_WORLDWIDE_COMMENT_KEY] = inputs[KOREA_COMMENT_SOURCE_KEY]
    record = ReportData(
        report_date=report_date,
        data_reference_date=prev_week,
        pricing_month=month.iso_key(),
        inputs=inputs,
        sheet_name=format_sheet_name(report_date),
        extra_cells={
            AUTO_THIS_WEEK: this_week.isoformat(),
            AUTO_PREV_WEEK: prev_week.isoformat(),
            AUTO_TWO_WEEKS: two_weeks.isoformat(),
        },
    )
    cache_set("input_data", record, INPUT_TTL_SECONDS)
    return record


def dated_report_exists(report_date: dt.date) -> bool:
    spreadsheet = _spreadsheet()
    return find_report_worksheet(spreadsheet, report_date) is not None


def _format_paper_label(value: float | None) -> str:
    if value is None:
        return "TBN"
    if value == int(value):
        return f"{int(value):.2f}"
    return f"{value:.2f}"


def _apply_inputs_to_report_worksheet(
    worksheet,
    report_date: dt.date,
    inputs: dict[str, Any],
    pricing_month: YearMonth,
    data_reference_date: dt.date,
) -> None:
    grid = worksheet.get("A1:AI80") or []
    cache: dict[str, Any] = {}

    def get_value(address: str) -> Any:
        if address in cache:
            return cache[address]
        return _grid_value(grid, address)

    def set_value(address: str, value: Any) -> None:
        if address in FORMULA_CELLS:
            return
        cache[address] = value

    shift_weekly_chart_window(get_value, set_value, data_reference_date)

    set_value(DATE_CELLS["paper_date"], data_reference_date.isoformat())
    set_value(DATE_CELLS["paper_table_date"], data_reference_date.isoformat())
    set_value(DATE_CELLS["market_table_date"], data_reference_date.isoformat())
    set_value(DATE_CELLS["premium_table_date"], data_reference_date.isoformat())
    set_value(TITLE_CELL, format_report_title(report_date))
    set_value(
        SALES_ASSUMPTION_HEADER_CELL,
        SALES_ASSUMPTION_HEADER_TEMPLATE.format(month=pricing_month.full_name),
    )

    for item in NUMBER_FIELDS:
        if item.section == "premium":
            parsed = parse_supplier_premium(inputs.get(item.key))
            set_value(item.cell, parsed["vlsfo"])
            continue
        number = _parse_optional_number(inputs.get(item.key))
        set_value(item.cell, number)
        if item.label_cell and item.label_template:
            set_value(
                item.label_cell,
                item.label_template.format(
                    month=pricing_month.abbr,
                    value=_format_paper_label(number),
                ),
            )

    for address, key in WORLDWIDE_NEW_POINT_CELLS.items():
        set_value(address, _parse_optional_number(inputs.get(key)))
    for address, key in WORLDWIDE_THIS_WEEK_CELLS.items():
        set_value(address, _parse_optional_number(inputs.get(key)))

    comment_inputs = dict(inputs)
    if not str(comment_inputs.get(KOREA_WORLDWIDE_COMMENT_KEY) or "").strip():
        comment_inputs[KOREA_WORLDWIDE_COMMENT_KEY] = str(inputs.get(KOREA_COMMENT_SOURCE_KEY) or "")
    for item in COMMENT_FIELDS:
        if item.key == KOREA_COMMENT_SOURCE_KEY:
            continue
        lines = compose_comment_lines(comment_inputs.get(item.key, ""), len(item.cells))
        for address, line in zip(item.cells, lines):
            set_value(address, line or None)

    for address in KOREA_BUNKER_MARKET_CELLS:
        set_value(address, None)

    for item in STRATEGY_FIELDS:
        set_value(item.cell, compose_strategy_text(item.prefix, inputs.get(item.key, "")))

    updates = []
    for address, value in cache.items():
        if address in FORMULA_CELLS:
            continue
        updates.append({"range": address, "values": [["" if value is None else value]]})
    if updates:
        worksheet.batch_update(updates, value_input_option="USER_ENTERED")


def create_or_update_report_from_input(record: ReportData) -> tuple[str, bool]:
    """Copy latest dated sheet if needed, then write INPUT values. Returns (sheet_name, is_update)."""
    spreadsheet = _spreadsheet(require_write=True)
    report_date = record.report_date
    target_name = format_sheet_name(report_date)
    existing = find_report_worksheet(spreadsheet, report_date)
    is_update = existing is not None
    if existing is None:
        latest = latest_report_sheet()
        if latest is None:
            raise DataServiceError("No existing dated report sheet to copy.")
        source = find_report_worksheet(spreadsheet, latest[0])
        if source is None:
            raise DataServiceError(f"Could not open source sheet `{latest[1]}`.")
        try:
            worksheet = spreadsheet.duplicate_sheet(
                source.id,
                insert_sheet_index=1,
                new_sheet_name=target_name,
            )
        except Exception as exc:
            raise DataServiceError(f"Could not copy `{latest[1]}` to `{target_name}`: {exc}") from exc
    else:
        worksheet = existing

    if worksheet.title in {INPUT_SHEET_NAME, EMAIL_RECIPIENTS_SHEET}:
        raise DataServiceError("Refusing to write INPUT values onto INPUT or Email Recipients.")

    month = YearMonth.parse(record.pricing_month) if record.pricing_month else default_pricing_month(report_date)
    try:
        _apply_inputs_to_report_worksheet(
            worksheet,
            report_date,
            record.inputs,
            month,
            record.data_reference_date,
        )
    except Exception as exc:
        raise DataServiceError(f"Could not write INPUT values to `{worksheet.title}`: {exc}") from exc
    return worksheet.title, is_update
