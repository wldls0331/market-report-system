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

        return getattr(st, "secrets", None)
    except Exception:
        return None


def _secret(name: str, default: str = "") -> str:
    _load_dotenv()
    secrets = _st_secrets()
    if secrets is not None:
        try:
            if name in secrets:
                value = secrets[name]
                if value is not None and str(value).strip() and not hasattr(value, "items"):
                    return str(value).strip()
        except Exception:
            pass
    return str(os.environ.get(name, default) or default).strip()


def _section(name: str) -> dict[str, str]:
    secrets = _st_secrets()
    if secrets is None:
        return {}
    try:
        block = secrets[name]
    except Exception:
        return {}
    try:
        items = dict(block)
    except Exception:
        return {}
    result: dict[str, str] = {}
    for key, value in items.items():
        if value is None:
            continue
        if hasattr(value, "items"):
            continue
        text = str(value).strip()
        if text:
            result[str(key)] = text
    return result


def google_sheet_id() -> str:
    section = _section("google")
    if section.get("sheet_id"):
        return section["sheet_id"]
    return _secret("GOOGLE_SHEET_ID")


def google_sheet_url() -> str:
    section = _section("google")
    if section.get("sheet_url"):
        return section["sheet_url"]
    explicit = _secret("GOOGLE_SHEET_URL")
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
    return _oauth_triplet("google")


def gmail_oauth_secrets() -> tuple[str, str, str]:
    return _oauth_triplet("gmail")


def has_google_secret_oauth() -> bool:
    client_id, client_secret, refresh_token = google_oauth_secrets()
    return bool(client_id and client_secret and refresh_token)


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
