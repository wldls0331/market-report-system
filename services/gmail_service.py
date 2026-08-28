"""Gmail API send via OAuth 2.0. Requests gmail.send only."""

from __future__ import annotations

import base64
import json
from email.message import EmailMessage
from pathlib import Path

from config.google_config import (
    credentials_from_refresh_token,
    gmail_credentials_file,
    gmail_oauth_secrets,
    gmail_token_file,
    has_gmail_secret_oauth,
    persist_oauth_token,
)
from config.runtime import supports_browser_oauth

GMAIL_SCOPES = ("https://www.googleapis.com/auth/gmail.send",)


class GmailServiceError(RuntimeError):
    pass


def gmail_status() -> tuple[bool, str]:
    if has_gmail_secret_oauth():
        return True, "Connected"
    token = gmail_token_file()
    if token.exists():
        return True, "Connected"
    if gmail_credentials_file().exists() and supports_browser_oauth():
        return False, "Authorization Required"
    return False, "Not Connected"


def _require_oauth_desktop_client(creds_path: Path) -> None:
    try:
        payload = json.loads(creds_path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as exc:
        raise GmailServiceError(f"credentials.json is not valid JSON: {creds_path}") from exc
    if payload.get("type") == "service_account":
        raise GmailServiceError(
            "Gmail requires an OAuth client, not a service account key."
        )
    if "installed" not in payload and "web" not in payload:
        raise GmailServiceError(
            "credentials.json is not an OAuth client file. Use a Desktop app OAuth client JSON."
        )


def _load_credentials():
    try:
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
    except ImportError as exc:
        raise GmailServiceError(
            "Gmail libraries are missing. Install google-api-python-client and google-auth-oauthlib."
        ) from exc

    if has_gmail_secret_oauth():
        client_id, client_secret, refresh_token = gmail_oauth_secrets()
        try:
            return credentials_from_refresh_token(
                client_id=client_id,
                client_secret=client_secret,
                refresh_token=refresh_token,
                scopes=list(GMAIL_SCOPES),
            )
        except Exception as exc:
            raise GmailServiceError(f"Gmail secret refresh failed: {exc}") from exc

    creds_path = gmail_credentials_file()
    token_path = gmail_token_file()
    credentials = None
    if token_path.exists():
        credentials = Credentials.from_authorized_user_file(str(token_path), list(GMAIL_SCOPES))
    if credentials and credentials.expired and credentials.refresh_token:
        credentials.refresh(Request())
        persist_oauth_token(token_path, credentials.to_json())
    if credentials and credentials.valid:
        return credentials

    if not supports_browser_oauth():
        raise GmailServiceError(
            "Gmail is not connected. Set [gmail] client_id, client_secret, and refresh_token in Streamlit Secrets."
        )
    if not creds_path.exists():
        raise GmailServiceError(
            f"Gmail OAuth client file not found: {creds_path}. Add credentials.json to authorize Gmail."
        )
    _require_oauth_desktop_client(creds_path)
    flow = InstalledAppFlow.from_client_secrets_file(str(creds_path), list(GMAIL_SCOPES))
    credentials = flow.run_local_server(
        port=0,
        access_type="offline",
        prompt="consent",
        open_browser=True,
    )
    persist_oauth_token(token_path, credentials.to_json())
    return credentials


def authorize_gmail() -> None:
    _load_credentials()


def _gmail_service(credentials):
    try:
        from googleapiclient.discovery import build
    except ImportError as exc:
        raise GmailServiceError("google-api-python-client is required to send email.") from exc
    return build("gmail", "v1", credentials=credentials, cache_discovery=False)


def _split_addresses(raw: str) -> list[str]:
    text = (raw or "").replace(";", ",")
    parts: list[str] = []
    for chunk in text.replace("\n", ",").split(","):
        email = chunk.strip()
        if email:
            parts.append(email)
    return parts


def build_message(
    *,
    to_text: str,
    cc_text: str,
    subject: str,
    body: str,
    attachment_path: Path | None,
) -> EmailMessage:
    to_list = _split_addresses(to_text)
    if not to_list:
        raise GmailServiceError("At least one TO recipient is required.")
    message = EmailMessage()
    message["To"] = ", ".join(to_list)
    cc_list = _split_addresses(cc_text)
    if cc_list:
        message["Cc"] = ", ".join(cc_list)
    message["Subject"] = (subject or "").strip() or "(no subject)"
    message.set_content(body or "")
    if attachment_path is not None:
        path = Path(attachment_path)
        if not path.exists():
            raise GmailServiceError(f"Attachment not found: {path.name}")
        data = path.read_bytes()
        message.add_attachment(
            data,
            maintype="application",
            subtype="pdf",
            filename=path.name,
        )
    return message


def send_email(
    *,
    to_text: str,
    cc_text: str,
    subject: str,
    body: str,
    attachment_path: Path | None,
) -> None:
    credentials = _load_credentials()
    if credentials is None or not getattr(credentials, "valid", False):
        raise GmailServiceError("Gmail is not authorized. Complete Google authorization first.")
    message = build_message(
        to_text=to_text,
        cc_text=cc_text,
        subject=subject,
        body=body,
        attachment_path=attachment_path,
    )
    encoded = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
    try:
        service = _gmail_service(credentials)
        service.users().messages().send(userId="me", body={"raw": encoded}).execute()
    except GmailServiceError:
        raise
    except Exception as exc:
        raise GmailServiceError(f"Gmail send failed: {exc}") from exc
