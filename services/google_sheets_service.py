"""Google Spreadsheet is the Market Report master: INPUT + dated report sheets + Email Recipients.

The INPUT tab is the live source. Dated YY.MM.DD 보고자료 sheets are history.
"""

from __future__ import annotations

import datetime as dt
import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from openpyxl.utils.cell import column_index_from_string, coordinate_from_string

from config.cell_mapping import (
    COMMENT_FIELDS,
    DATE_CELLS,
    FORMULA_CELLS,
    NUMBER_FIELDS,
    REPORT_SHEET_CELLS,
    SHEET_NAME_PATTERN,
    STRATEGY_FIELDS,
)
from config.input_sheet import INPUT_SHEET_NAME
from config.google_config import (
    credentials_from_refresh_token,
    google_credentials_file,
    google_oauth_secrets,
    google_service_account_info,
    google_sheet_id,
    has_google_secret_oauth,
    has_google_sheets_secrets,
    invalid_google_oauth_secret_keys,
    missing_google_sheets_secret_keys,
    persist_oauth_token,
    sheets_token_file,
)
from config.paths import DATA_DIR
from config.runtime import is_hosted, supports_browser_oauth
from services.working_day_service import format_sheet_name, previous_week_last_working_day

try:
    from zoneinfo import ZoneInfo

    SGT = ZoneInfo("Asia/Singapore")
except Exception:
    SGT = dt.timezone(dt.timedelta(hours=8), name="SGT")

EMAIL_RECIPIENTS_SHEET = "Email Recipients"
RECIPIENT_HEADERS: tuple[str, ...] = ("Type", "Name", "Email", "Active")
SHEETS_READONLY_SCOPE = "https://www.googleapis.com/auth/spreadsheets.readonly"
SHEETS_FULL_SCOPE = "https://www.googleapis.com/auth/spreadsheets"
SHEETS_SCOPES = (SHEETS_FULL_SCOPE,)
LOCAL_RECIPIENTS_PATH = DATA_DIR / "local_email_recipients.json"
REPORT_GRID_RANGE = "A1:AI80"
SHEET_RE = re.compile(SHEET_NAME_PATTERN)


class DataServiceError(RuntimeError):
    pass


@dataclass(frozen=True)
class ReportData:
    report_date: dt.date
    data_reference_date: dt.date
    pricing_month: str
    inputs: dict[str, str]
    sheet_name: str = ""
    extra_cells: dict[str, Any] = field(default_factory=dict)
    saved_at: dt.datetime | None = None
    last_updated_at: dt.datetime | None = None
    updated_by: str = ""

    def last_saved_label(self) -> str:
        if self.last_updated_at is None:
            return ""
        return format_sgt_short(self.last_updated_at)


@dataclass(frozen=True)
class EmailRecipient:
    kind: str
    name: str
    email: str
    active: bool


def now_sgt() -> dt.datetime:
    return dt.datetime.now(SGT)


def format_sgt_short(moment: dt.datetime) -> str:
    if moment.tzinfo is None:
        moment = moment.replace(tzinfo=SGT)
    return moment.astimezone(SGT).strftime("%Y-%m-%d %H:%M SGT")


def format_sgt_full(moment: dt.datetime) -> str:
    if moment.tzinfo is None:
        moment = moment.replace(tzinfo=SGT)
    return moment.astimezone(SGT).strftime("%Y-%m-%d %H:%M:%S SGT")


def _parse_sheet_date(name: str) -> dt.date | None:
    match = SHEET_RE.match(name.strip())
    if not match:
        return None
    year, month, day = (int(part) for part in match.groups())
    year += 2000
    try:
        return dt.date(year, month, day)
    except ValueError:
        return None


def _stringify(value: Any) -> str:
    if value is None or value == "":
        return ""
    if isinstance(value, dt.datetime):
        return value.date().isoformat()
    if isinstance(value, dt.date):
        return value.isoformat()
    if isinstance(value, float):
        if abs(value - round(value)) < 1e-9:
            return str(int(round(value)))
        return f"{value:.4f}".rstrip("0").rstrip(".")
    return str(value).strip()


def _join_comment(grid: list[list[Any]], addresses: tuple[str, ...]) -> str:
    lines: list[str] = []
    for address in addresses:
        text = _stringify(_grid_value(grid, address))
        if text:
            lines.append(text)
    return "\n".join(lines)


def _grid_value(grid: list[list[Any]], address: str) -> Any:
    column_letter, row = coordinate_from_string(address)
    column = column_index_from_string(column_letter)
    if row < 1 or row > len(grid):
        return None
    line = grid[row - 1]
    if column < 1 or column > len(line):
        return None
    return line[column - 1]


def _parse_date_value(value: Any) -> dt.date | None:
    if value is None or value == "":
        return None
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if isinstance(value, (int, float)) and value > 20000:
        try:
            return (dt.datetime(1899, 12, 30) + dt.timedelta(days=int(value))).date()
        except Exception:
            return None
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%b-%Y"):
        try:
            return dt.datetime.strptime(text[:20], fmt).date()
        except ValueError:
            continue
    match = re.match(r"^(\d{4})-(\d{2})-(\d{2})", text)
    if match:
        try:
            return dt.date(int(match.group(1)), int(match.group(2)), int(match.group(3)))
        except ValueError:
            return None
    return None


def _typed_cell(address: str, value: Any) -> Any:
    if value is None or value == "":
        return None
    if address in FORMULA_CELLS:
        return value
    if address in DATE_CELLS.values() or address in {"C22", "D22", "E22", "P22", "Q22", "R22"}:
        parsed = _parse_date_value(value)
        return parsed if parsed is not None else value
    number_cells = {item.cell for item in NUMBER_FIELDS}
    chart_numbers = {
        f"{col}{row}"
        for row in range(23, 27)
        for col in ("C", "D", "E", "F", "P", "Q", "R", "S")
    }
    if address in number_cells or address in chart_numbers:
        if isinstance(value, (int, float)):
            return value
        text = str(value).strip().replace(",", "")
        if text in {"-", "TBN", ""}:
            return None
        try:
            return float(text)
        except ValueError:
            return value
    return _stringify(value) or None


class _LocalRecipients:
    def __init__(self, path: Path | None = None) -> None:
        self.path = path or LOCAL_RECIPIENTS_PATH

    def rows(self) -> list[list[str]]:
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        if not self.path.exists():
            return [list(RECIPIENT_HEADERS)]
        try:
            payload = json.loads(self.path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            return [list(RECIPIENT_HEADERS)]
        rows = payload.get("recipients") or [list(RECIPIENT_HEADERS)]
        return [[str(cell) for cell in row] for row in rows]


def _read_credentials_payload() -> dict[str, Any]:
    creds_path = google_credentials_file()
    if not creds_path.exists():
        return {}
    try:
        return json.loads(creds_path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError:
        return {}


def _is_service_account(payload: dict[str, Any]) -> bool:
    return payload.get("type") == "service_account"


def _is_oauth_client(payload: dict[str, Any]) -> bool:
    return "installed" in payload or "web" in payload


def _session_get(key: str, default: Any = None) -> Any:
    try:
        import streamlit as st

        return st.session_state.get(key, default)
    except Exception:
        return default


def _session_set(key: str, value: Any) -> None:
    try:
        import streamlit as st

        st.session_state[key] = value
    except Exception:
        pass


def sheets_status() -> tuple[bool, str]:
    from services.sheets_cache import STATUS_TTL_SECONDS, cache_get, cache_set

    cached = cache_get("sheets_status")
    if isinstance(cached, tuple) and len(cached) == 2:
        return cached[0], cached[1]
    session_cached = _session_get("_sheets_status_result")
    if isinstance(session_cached, tuple) and len(session_cached) == 2:
        return session_cached[0], session_cached[1]
    if not google_sheet_id():
        result = (False, "Not Connected")
        cache_set("sheets_status", result, STATUS_TTL_SECONDS)
        _session_set("_sheets_status_result", result)
        return result
    try:
        from concurrent.futures import ThreadPoolExecutor, TimeoutError as FuturesTimeout

        from services.perf import timed

        with timed("google sheets auth"):
            with ThreadPoolExecutor(max_workers=1) as pool:
                future = pool.submit(_google_client)
                spreadsheet = future.result(timeout=12)
        title = str(spreadsheet.title or "").strip()
        _session_set("_spreadsheet_title_live", title)
        _session_set("_sheets_error", "")
        cache_set("spreadsheet_title", title, STATUS_TTL_SECONDS)
        result = (True, "Connected")
    except FuturesTimeout:
        _session_set("_spreadsheet_title_live", "")
        _session_set("_sheets_error", "Google Sheets timed out.")
        result = (False, "Not Connected")
    except DataServiceError as exc:
        _session_set("_spreadsheet_title_live", "")
        _session_set("_sheets_error", str(exc))
        if supports_browser_oauth() and "authorization required" in str(exc).lower():
            result = (False, "Authorization Required")
        else:
            result = (False, "Not Connected")
    except Exception as exc:
        _session_set("_spreadsheet_title_live", "")
        _session_set("_sheets_error", str(exc))
        result = (False, "Not Connected")
    cache_set("sheets_status", result, STATUS_TTL_SECONDS)
    _session_set("_sheets_status_result", result)
    return result


def spreadsheet_title() -> str:
    from services.sheets_cache import cache_get

    cached = cache_get("spreadsheet_title")
    if isinstance(cached, str) and cached:
        return cached
    cached = _session_get("_spreadsheet_title_live")
    if cached:
        return str(cached)
    if not sheets_status()[0]:
        return ""
    try:
        title = str(_google_client().title or "").strip()
        _session_set("_spreadsheet_title_live", title)
        return title
    except Exception:
        return ""


def _assert_google_ready_if_configured() -> None:
    connected, label = sheets_status()
    if connected:
        return
    missing = missing_google_sheets_secret_keys()
    if missing:
        raise DataServiceError("Google Sheets environment variables missing: " + ", ".join(missing))
    invalid = invalid_google_oauth_secret_keys()
    if invalid:
        raise DataServiceError("Google Sheets OAuth secrets invalid: " + " ".join(invalid))
    if label == "Authorization Required":
        raise DataServiceError(
            "Google Sheets authorization required. Click Authorize Google Sheets in the sidebar."
        )
    err = _session_get("_sheets_error")
    if err:
        raise DataServiceError(str(err))
    if not google_sheet_id():
        return
    raise DataServiceError("Google Sheets is not connected.")


def _token_can_read_sheets(credentials) -> bool:
    scopes = set(credentials.scopes or [])
    if not scopes:
        return bool(getattr(credentials, "valid", False))
    return SHEETS_READONLY_SCOPE in scopes or SHEETS_FULL_SCOPE in scopes


def _token_can_write_sheets(credentials) -> bool:
    scopes = set(credentials.scopes or [])
    if not scopes:
        return bool(getattr(credentials, "valid", False))
    return SHEETS_FULL_SCOPE in scopes


def _sheets_user_credentials(*, allow_browser: bool, require_write: bool = False):
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials as UserCredentials
    from google_auth_oauthlib.flow import InstalledAppFlow

    creds_path = google_credentials_file()
    token_path = sheets_token_file()
    credentials = None
    if token_path.exists():
        credentials = UserCredentials.from_authorized_user_file(str(token_path))
        if credentials and not _token_can_read_sheets(credentials):
            credentials = None
        if require_write and credentials and not _token_can_write_sheets(credentials):
            credentials = None
    if credentials and credentials.expired and credentials.refresh_token:
        credentials.refresh(Request())
        persist_oauth_token(token_path, credentials.to_json())
    if credentials and credentials.valid:
        return credentials
    if not allow_browser or not supports_browser_oauth():
        raise DataServiceError(
            "Google Sheets authorization required. "
            "Set GOOGLE_SHEET_ID, GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET, and "
            "GOOGLE_REFRESH_TOKEN, or authorize locally with credentials.json."
        )
    flow = InstalledAppFlow.from_client_secrets_file(str(creds_path), list(SHEETS_SCOPES))
    credentials = flow.run_local_server(
        port=0,
        access_type="offline",
        prompt="consent",
        open_browser=True,
    )
    persist_oauth_token(token_path, credentials.to_json())
    return credentials


def authorize_sheets() -> None:
    """OAuth for Sheets only. Does not touch Gmail token.json."""
    if has_google_sheets_secrets():
        _google_client()
        return
    creds_path = google_credentials_file()
    if not creds_path.exists():
        raise DataServiceError(f"Google credentials file not found: {creds_path}")
    payload = _read_credentials_payload()
    if _is_service_account(payload):
        if not google_sheet_id():
            raise DataServiceError("Set GOOGLE_SHEET_ID before connecting a service account.")
        _google_client()
        return
    if not _is_oauth_client(payload):
        raise DataServiceError(
            "credentials.json is not an OAuth Desktop Client file. "
            "Do not use a service account key as credentials.json for this setup."
        )
    _sheets_user_credentials(allow_browser=True, require_write=True)
    if google_sheet_id():
        _google_client(require_write=True)


def _google_credentials(*, require_write: bool = False):
    client_id, client_secret, refresh_token = google_oauth_secrets()
    if client_id and client_secret and refresh_token:
        invalid = invalid_google_oauth_secret_keys()
        if invalid:
            raise DataServiceError("Google Sheets OAuth secrets invalid: " + " ".join(invalid))
        try:
            return credentials_from_refresh_token(
                client_id=client_id,
                client_secret=client_secret,
                refresh_token=refresh_token,
                scopes=list(SHEETS_SCOPES),
            )
        except DataServiceError:
            raise
        except Exception as exc:
            detail = str(exc)
            if "invalid_client" in detail.lower():
                raise DataServiceError(
                    "Google Sheets OAuth client_id was not recognized. "
                    "Copy GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET from credentials.json "
                    "installed.client_id and installed.client_secret, "
                    "and copy GOOGLE_REFRESH_TOKEN from sheets_token.json."
                ) from exc
            raise DataServiceError(f"Google Sheets secret refresh failed: {exc}") from exc

    if is_hosted() or not google_credentials_file().exists():
        missing = missing_google_sheets_secret_keys()
        if missing:
            raise DataServiceError("Google Sheets environment variables missing: " + ", ".join(missing))
        raise DataServiceError(
            "Google Sheets could not authenticate. "
            "Set GOOGLE_SHEET_ID, GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET, and GOOGLE_REFRESH_TOKEN."
        )

    service_account = google_service_account_info()
    if service_account:
        try:
            from google.oauth2.service_account import Credentials as ServiceCredentials
        except ImportError as exc:
            raise DataServiceError(
                "Google Sheets libraries are missing. Install gspread and google-auth."
            ) from exc
        return ServiceCredentials.from_service_account_info(
            service_account, scopes=list(SHEETS_SCOPES)
        )

    creds_path = google_credentials_file()
    if not creds_path.exists():
        raise DataServiceError(f"Google credentials file not found: {creds_path}")
    try:
        from google.oauth2.service_account import Credentials as ServiceCredentials
    except ImportError as exc:
        raise DataServiceError(
            "Google Sheets libraries are missing. Install gspread and google-auth."
        ) from exc
    payload = json.loads(creds_path.read_text(encoding="utf-8-sig"))
    if _is_service_account(payload):
        return ServiceCredentials.from_service_account_file(str(creds_path), scopes=list(SHEETS_SCOPES))
    if _is_oauth_client(payload):
        return _sheets_user_credentials(allow_browser=False, require_write=require_write)
    raise DataServiceError(
        "credentials.json is not a service account key or an OAuth Desktop/Web client file."
    )


def _google_client(*, require_write: bool = False):
    sheet_id = google_sheet_id()
    if not sheet_id:
        raise DataServiceError("GOOGLE_SHEET_ID is not configured.")
    try:
        import gspread
    except ImportError as exc:
        raise DataServiceError(
            "Google Sheets libraries are missing. Install gspread and google-auth."
        ) from exc
    credentials = _google_credentials(require_write=require_write)
    client = gspread.authorize(credentials)
    return client.open_by_key(sheet_id)


def open_spreadsheet(*, require_write: bool = False):
    """Public Sheets connection used by INPUT read/write. Wraps `_google_client`."""
    return _google_client(require_write=require_write)


def _recipients_worksheet(spreadsheet):
    try:
        return spreadsheet.worksheet(EMAIL_RECIPIENTS_SHEET)
    except Exception:
        return None


def find_report_worksheet(spreadsheet, report_date: dt.date):
    expected = format_sheet_name(report_date)
    worksheets = list(spreadsheet.worksheets())
    for worksheet in worksheets:
        if worksheet.title == expected:
            return worksheet
    for worksheet in worksheets:
        if worksheet.title == EMAIL_RECIPIENTS_SHEET or worksheet.title == INPUT_SHEET_NAME:
            continue
        parsed = _parse_sheet_date(worksheet.title)
        if parsed == report_date:
            return worksheet
    return None


def latest_report_sheet() -> tuple[dt.date, str] | None:
    titles = list_worksheet_titles()
    latest: tuple[dt.date, str] | None = None
    for title in titles:
        if title in {EMAIL_RECIPIENTS_SHEET, INPUT_SHEET_NAME}:
            continue
        parsed = _parse_sheet_date(title)
        if parsed is None:
            continue
        if latest is None or parsed > latest[0]:
            latest = (parsed, title)
    return latest


def _record_from_grid(report_date: dt.date, sheet_name: str, grid: list[list[Any]]) -> ReportData:
    inputs: dict[str, str] = {}
    for item in NUMBER_FIELDS:
        inputs[item.key] = _stringify(_grid_value(grid, item.cell))
    for item in COMMENT_FIELDS:
        inputs[item.key] = _join_comment(grid, item.cells)
    for item in STRATEGY_FIELDS:
        inputs[item.key] = _stringify(_grid_value(grid, item.cell))

    extra_cells: dict[str, Any] = {}
    for address in REPORT_SHEET_CELLS:
        extra_cells[address] = _typed_cell(address, _grid_value(grid, address))

    af2 = _parse_date_value(_grid_value(grid, DATE_CELLS["paper_date"]))
    data_reference_date = af2 or previous_week_last_working_day(report_date)
    return ReportData(
        report_date=report_date,
        data_reference_date=data_reference_date,
        pricing_month="",
        inputs=inputs,
        sheet_name=sheet_name,
        extra_cells=extra_cells,
    )


def load_report_data(report_date: dt.date) -> ReportData | None:
    _assert_google_ready_if_configured()
    connected, _label = sheets_status()
    if not connected:
        return None
    try:
        spreadsheet = _google_client()
        worksheet = find_report_worksheet(spreadsheet, report_date)
        if worksheet is None:
            return None
        grid = worksheet.get(REPORT_GRID_RANGE) or []
        return _record_from_grid(report_date, worksheet.title, grid)
    except DataServiceError:
        raise
    except Exception as exc:
        raise DataServiceError(f"Could not load Google Sheets report: {exc}") from exc


def report_data_exists(report_date: dt.date) -> bool:
    return load_report_data(report_date) is not None


def list_worksheet_titles() -> list[str]:
    _assert_google_ready_if_configured()
    connected, _label = sheets_status()
    if not connected:
        return []
    try:
        return [worksheet.title for worksheet in _google_client().worksheets()]
    except DataServiceError:
        raise
    except Exception as exc:
        raise DataServiceError(f"Could not list Google Sheets tabs: {exc}") from exc


def _parse_recipient_rows(rows: list[list[str]]) -> list[EmailRecipient]:
    recipients: list[EmailRecipient] = []
    for row in rows[1:]:
        if len(row) < 3:
            continue
        kind = str(row[0]).strip().upper()
        name = str(row[1]).strip() if len(row) > 1 else ""
        email = str(row[2]).strip() if len(row) > 2 else ""
        active_raw = str(row[3]).strip().upper() if len(row) > 3 else "TRUE"
        active = active_raw in {"TRUE", "YES", "1", "Y"}
        if kind in {"TO", "CC"} and email:
            recipients.append(EmailRecipient(kind=kind, name=name, email=email, active=active))
    return recipients


def load_email_recipients_result() -> tuple[list[EmailRecipient], str | None]:
    from services.sheets_cache import RECIPIENTS_TTL_SECONDS, cache_get, cache_set

    cached = cache_get("email_recipients")
    if cached is not None:
        return cached
    _assert_google_ready_if_configured()
    connected, _label = sheets_status()
    if connected:
        try:
            spreadsheet = _google_client()
            worksheet = _recipients_worksheet(spreadsheet)
            if worksheet is None:
                result = ([], "Email Recipients sheet not found")
                cache_set("email_recipients", result, RECIPIENTS_TTL_SECONDS)
                return result
            result = (_parse_recipient_rows(worksheet.get_all_values()), None)
            cache_set("email_recipients", result, RECIPIENTS_TTL_SECONDS)
            return result
        except DataServiceError:
            raise
        except Exception as exc:
            raise DataServiceError(f"Could not load email recipients: {exc}") from exc
    return _parse_recipient_rows(_LocalRecipients().rows()), None


def load_email_recipients() -> list[EmailRecipient]:
    recipients, _warning = load_email_recipients_result()
    return recipients


def active_recipient_emails() -> tuple[list[str], list[str]]:
    to_list: list[str] = []
    cc_list: list[str] = []
    for item in load_email_recipients():
        if not item.active:
            continue
        if item.kind == "TO":
            to_list.append(item.email)
        elif item.kind == "CC":
            cc_list.append(item.email)
    return to_list, cc_list
