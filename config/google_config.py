"""Google API settings from Streamlit secrets, .env, or local files. No hardcoded IDs."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any

from config.paths import PROJECT_ROOT
from config.runtime import is_streamlit_cloud

_DOTENV_LOADED = False
TOKEN_URI = "https://oauth2.googleapis.com/token"


def _load_dotenv() -> None:
    global _DOTENV_LOADED
    if _DOTENV_LOADED:
        return
    _DOTENV_LOADED = True
    path = PROJECT_ROOT / ".env"
    if not path.exists():
        return
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


def _st_secrets():
    try:
        import streamlit as st

        return st.secrets
    except Exception:
        return None


def _as_plain_dict(block: Any) -> dict[str, Any]:
    if block is None:
        return {}
    if isinstance(block, str):
        return {}
    try:
        return {str(key): value for key, value in block.items()}
    except Exception:
        try:
            return dict(block)
        except Exception:
            return {}


def _scalar(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, dict):
        return ""
    text = str(value).strip()
    if (len(text) >= 2) and text[0] == text[-1] and text[0] in {'"', "'"}:
        text = text[1:-1].strip()
    return text.replace("\ufeff", "").strip()


def _secret(name: str, default: str = "") -> str:
    _load_dotenv()
    secrets = _st_secrets()
    if secrets is not None:
        try:
            value = secrets[name]
            text = _scalar(value)
            if text:
                return text
        except Exception:
            pass
    return str(os.environ.get(name, default) or default).strip()


def _section(name: str) -> dict[str, str]:
    secrets = _st_secrets()
    if secrets is None:
        return {}
    try:
        block = secrets.get(name) if hasattr(secrets, "get") else secrets[name]
    except Exception:
        try:
            block = secrets[name]
        except Exception:
            return {}
    result: dict[str, str] = {}
    for key, value in _as_plain_dict(block).items():
        text = _scalar(value)
        if text:
            result[str(key)] = text
    return result


def _google_section_value(key: str) -> str:
    """Read st.secrets['google'][key] for Cloud OAuth. Does not use local files."""
    secrets = _st_secrets()
    if secrets is None:
        return ""
    try:
        google = secrets["google"]
        try:
            text = _scalar(google[key])
            if text:
                return text
        except Exception:
            pass
        try:
            text = _scalar(getattr(google, key))
            if text:
                return text
        except Exception:
            pass
    except Exception:
        pass
    try:
        text = _scalar(secrets[f"google.{key}"])
        if text:
            return text
    except Exception:
        pass
    return ""


def google_sheet_id() -> str:
    return _google_section_value("sheet_id") or _secret("GOOGLE_SHEET_ID")


def google_sheet_url() -> str:
    explicit = _google_section_value("sheet_url") or _secret("GOOGLE_SHEET_URL")
    if explicit:
        return explicit
    sheet_id = google_sheet_id()
    if sheet_id:
        return f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit"
    return ""


def google_credentials_file() -> Path:
    raw = _secret("GOOGLE_CREDENTIALS_FILE") or _secret("GOOGLE_SHEETS_CREDENTIALS_FILE") or "credentials.json"
    path = Path(raw)
    if not path.is_absolute():
        path = PROJECT_ROOT / path
    return path


def gmail_credentials_file() -> Path:
    raw = _secret("GMAIL_CREDENTIALS_FILE") or _secret("GOOGLE_CREDENTIALS_FILE") or "credentials.json"
    path = Path(raw)
    if not path.is_absolute():
        path = PROJECT_ROOT / path
    return path


def gmail_token_file() -> Path:
    raw = _secret("GMAIL_TOKEN_FILE") or "token.json"
    path = Path(raw)
    if not path.is_absolute():
        path = PROJECT_ROOT / path
    return path


def sheets_token_file() -> Path:
    raw = _secret("GOOGLE_SHEETS_TOKEN_FILE") or "sheets_token.json"
    path = Path(raw)
    if not path.is_absolute():
        path = PROJECT_ROOT / path
    return path


def _oauth_triplet(section_name: str) -> tuple[str, str, str]:
    section = _section(section_name)
    fallback = _section("google") if section_name != "google" else {}
    client_id = section.get("client_id") or fallback.get("client_id") or ""
    client_secret = section.get("client_secret") or fallback.get("client_secret") or ""
    refresh_token = section.get("refresh_token") or ""
    return client_id, client_secret, refresh_token


def google_oauth_secrets() -> tuple[str, str, str]:
    return (
        _google_section_value("client_id"),
        _google_section_value("client_secret"),
        _google_section_value("refresh_token"),
    )


def gmail_oauth_secrets() -> tuple[str, str, str]:
    return _oauth_triplet("gmail")


def has_google_secret_oauth() -> bool:
    client_id, client_secret, refresh_token = google_oauth_secrets()
    return bool(client_id and client_secret and refresh_token)


def google_service_account_info() -> dict[str, Any] | None:
    secrets = _st_secrets()
    if secrets is None:
        return None
    candidates: list[Any] = []
    for key in ("gcp_service_account", "google_service_account"):
        try:
            candidates.append(secrets[key])
        except Exception:
            continue
    try:
        raw_google = secrets["google"]
        nested = _as_plain_dict(raw_google).get("service_account")
        if nested is not None:
            candidates.append(nested)
    except Exception:
        pass
    for raw in candidates:
        info = _as_plain_dict(raw)
        email = str(info.get("client_email") or "").strip()
        private_key = str(info.get("private_key") or "").replace("\\n", "\n").strip()
        if not email or not private_key:
            continue
        info["client_email"] = email
        info["private_key"] = private_key
        info.setdefault("type", "service_account")
        return info
    return None


def has_google_service_account_secret() -> bool:
    return google_service_account_info() is not None


def has_google_sheets_secrets() -> bool:
    return bool(google_sheet_id()) and (
        has_google_secret_oauth() or has_google_service_account_secret()
    )


def missing_google_sheets_secret_keys() -> list[str]:
    missing: list[str] = []
    if not google_sheet_id():
        missing.append("[google].sheet_id")
    client_id, client_secret, refresh_token = google_oauth_secrets()
    if not client_id:
        missing.append("[google].client_id")
    if not client_secret:
        missing.append("[google].client_secret")
    if not refresh_token:
        missing.append("[google].refresh_token")
    return missing


def has_gmail_secret_oauth() -> bool:
    client_id, client_secret, refresh_token = gmail_oauth_secrets()
    return bool(client_id and client_secret and refresh_token)


def persist_oauth_token(path: Path, payload: str) -> None:
    """Write local token files only. Streamlit Cloud filesystem is ephemeral."""
    if is_streamlit_cloud():
        return
    try:
        path.write_text(payload, encoding="utf-8")
    except OSError:
        return


def credentials_from_refresh_token(
    *,
    client_id: str,
    client_secret: str,
    refresh_token: str,
    scopes: list[str],
) -> Any:
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials

    credentials = Credentials(
        token=None,
        refresh_token=refresh_token,
        token_uri=TOKEN_URI,
        client_id=client_id,
        client_secret=client_secret,
        scopes=scopes,
    )
    credentials.refresh(Request())
    return credentials
