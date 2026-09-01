"""Streamlit Community Cloud entry. Local HTML UI stays on Flask server.py:8502.

Google Sheets is the source of truth. session_state only holds UI artifacts
(email draft, last PDF path, last preview payload for widget reruns).
"""

from __future__ import annotations

import html
from pathlib import Path

import plotly.graph_objects as go
import streamlit as st

from config.google_config import (
    google_sheet_url,
    invalid_google_oauth_secret_keys,
    missing_google_sheets_secret_keys,
)
from config.runtime import is_streamlit_cloud, supports_browser_oauth
from services.gmail_service import GmailServiceError, authorize_gmail, send_email
from services.google_sheets_service import DataServiceError, authorize_sheets
from services.preview_service import (
    build_preview,
    connection_status,
    create_report_files,
    public_preview,
)

PALETTE = ["#1b365d", "#6b7280", "#c4a35a", "#334155", "#94a3b8"]

st.set_page_config(page_title="Market Report", layout="wide", initial_sidebar_state="collapsed")

st.markdown(
    """
    <style>
    [data-testid="stToolbar"], [data-testid="stDecoration"],
    [data-testid="stStatusWidget"], .stDeployButton,
    footer { display: none !important; }
    .block-container { padding-top: 1.2rem; max-width: 1280px; }
    h1 { font-weight: 650; letter-spacing: -0.02em; color: #1b365d; }
    .status-ok { color: #166534; font-weight: 600; }
    .status-bad { color: #9f1239; font-weight: 600; }
    .report-banner {
        background: #1b365d; color: #fff; padding: 0.85rem 1rem;
        border-bottom: 3px solid #c4a35a; border-radius: 8px 8px 0 0;
        font-weight: 650; letter-spacing: 0.04em;
    }
    .comment-box {
        white-space: pre-wrap; line-height: 1.55; font-size: 0.95rem;
        margin: 0 0 0.8rem 0;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def _comment_html(text: str) -> str:
    return f'<p class="comment-box">{html.escape(text or "TBN").replace(chr(10), "<br>")}</p>'


def _status_html(label: str, value: str, ok: bool) -> str:
    klass = "status-ok" if ok else "status-bad"
    return f'{label}: <span class="{klass}">{value}</span>'


def _line_fig(payload: dict) -> go.Figure:
    fig = go.Figure()
    labels = payload.get("labels") or []
    for index, series in enumerate(payload.get("series") or []):
        fig.add_trace(
            go.Scatter(
                x=labels,
                y=series.get("data") or [],
                name=series.get("name") or f"Series {index + 1}",
                mode="lines+markers",
                line={"color": PALETTE[index % len(PALETTE)], "width": 2},
                connectgaps=True,
            )
        )
    fig.update_layout(
        height=340,
        margin={"l": 40, "r": 16, "t": 8, "b": 40},
        paper_bgcolor="white",
        plot_bgcolor="white",
        legend={"orientation": "h", "y": -0.2},
        xaxis={"showgrid": False},
        yaxis={"gridcolor": "#eef2f7"},
    )
    return fig


def _bar_fig(payload: dict) -> go.Figure:
    fig = go.Figure()
    labels = payload.get("labels") or []
    for index, series in enumerate(payload.get("series") or []):
        fig.add_trace(
            go.Bar(
                x=labels,
                y=series.get("data") or [],
                name=series.get("name") or f"Series {index + 1}",
                marker_color=PALETTE[index % len(PALETTE)],
            )
        )
    fig.update_layout(
        barmode="group",
        height=380,
        margin={"l": 40, "r": 16, "t": 8, "b": 40},
        paper_bgcolor="white",
        plot_bgcolor="white",
        legend={"orientation": "h", "y": -0.2},
        xaxis={"showgrid": False},
        yaxis={"gridcolor": "#eef2f7"},
    )
    return fig


def _hydrate_email(preview: dict) -> None:
    email = preview.get("email") or {}
    st.session_state["email_to"] = email.get("to") or ""
    st.session_state["email_cc"] = email.get("cc") or ""
    st.session_state["email_subject"] = email.get("subject") or ""
    st.session_state["email_body"] = email.get("body") or ""


def _store_preview(payload: dict, *, hydrate_email: bool) -> dict:
    preview = public_preview(payload)
    files = preview.get("files") or {}
    if files.get("pdf_path"):
        st.session_state["pdf_path"] = files["pdf_path"]
    if files.get("excel_path"):
        st.session_state["excel_path"] = files["excel_path"]
    st.session_state["preview"] = preview
    if hydrate_email:
        _hydrate_email(preview)
    return preview


def _load_from_sheets(*, sync_sheet: bool) -> dict:
    st.session_state.pop("_sheets_status_result", None)
    st.session_state.pop("_spreadsheet_title_live", None)
    st.session_state.pop("_sheets_error", None)
    return _store_preview(build_preview(sync_sheet=sync_sheet), hydrate_email=True)


def _create_from_sheets() -> dict:
    st.session_state.pop("_sheets_status_result", None)
    st.session_state.pop("_spreadsheet_title_live", None)
    st.session_state.pop("_sheets_error", None)
    return _store_preview(create_report_files(), hydrate_email=True)


def _file_bytes(path_text: str | None) -> bytes | None:
    if not path_text:
        return None
    path = Path(path_text)
    if not path.exists():
        return None
    return path.read_bytes()


def main() -> None:
    st.title("Market Report")
    st.caption("Google Sheets INPUT is the source. This page shows the finished weekly report.")

    if "preview" not in st.session_state:
        try:
            _load_from_sheets(sync_sheet=False)
        except DataServiceError as exc:
            st.session_state["preview"] = {
                "error": str(exc),
                "status": connection_status(),
                "meta": None,
                "comments": {},
                "charts": {},
                "email": {},
                "files": {},
            }

    preview = st.session_state.get("preview") or {}
    status = preview.get("status") or connection_status()
    sheets_ok = bool(status.get("sheets_ok"))
    gmail_ok = bool(status.get("gmail_ok"))

    top = st.columns([2.2, 2.2, 1.4, 1.4, 1.6])
    with top[0]:
        st.markdown(
            _status_html("Google Sheets", status.get("sheets") or "Not Connected", sheets_ok),
            unsafe_allow_html=True,
        )
        if status.get("spreadsheet"):
            st.caption(status["spreadsheet"])
    with top[1]:
        st.markdown(
            _status_html("Gmail", status.get("gmail") or "Not Connected", gmail_ok),
            unsafe_allow_html=True,
        )
    with top[2]:
        refresh = st.button("Refresh Data", width="stretch")
    with top[3]:
        sheet_url = status.get("sheet_url") or google_sheet_url()
        if sheet_url:
            st.link_button("Open Google Sheet", sheet_url, width="stretch")
    with top[4]:
        create = st.button("Create / Update Report", type="primary", width="stretch")

    if not sheets_ok:
        missing = missing_google_sheets_secret_keys()
        invalid = invalid_google_oauth_secret_keys()
        if is_streamlit_cloud():
            if missing:
                st.error("Streamlit Secrets missing: " + ", ".join(missing))
            if invalid:
                st.error("Streamlit Secrets invalid: " + " ".join(invalid))
            if not missing and not invalid:
                st.error(preview.get("error") or "Google Sheets is not connected.")
        elif status.get("sheets") == "Authorization Required" and supports_browser_oauth():
            if st.button("Authorize Google Sheets"):
                try:
                    authorize_sheets()
                    st.rerun()
                except Exception as exc:
                    st.error(str(exc))
        elif preview.get("error"):
            st.error(preview["error"])

    if not gmail_ok and supports_browser_oauth() and not is_streamlit_cloud():
        if st.button("Authorize Gmail"):
            try:
                authorize_gmail()
                st.rerun()
            except Exception as exc:
                st.error(str(exc))
    elif not gmail_ok and is_streamlit_cloud():
        st.caption("Add [gmail] client_id, client_secret, and refresh_token in Streamlit Secrets.")

    if refresh:
        with st.spinner("Reading the latest Google Sheets INPUT and report tabs…"):
            try:
                preview = _load_from_sheets(sync_sheet=False)
            except DataServiceError as exc:
                st.error(str(exc))
                st.stop()
        st.rerun()

    if create:
        with st.spinner("Re-reading Google Sheets, updating the dated tab, and generating PDF…"):
            try:
                preview = _create_from_sheets()
            except DataServiceError as exc:
                st.error(str(exc))
                st.stop()
        st.rerun()

    preview = st.session_state.get("preview") or {}
    if preview.get("error"):
        st.error(preview["error"])
        st.stop()
    if preview.get("warning"):
        st.warning(preview["warning"])

    meta = preview.get("meta") or {}
    if meta.get("date_warning"):
        st.warning(meta["date_warning"])

    info = st.columns(5)
    info[0].metric("Report Date", meta.get("report_date") or "—")
    info[1].metric("Pricing Month", meta.get("pricing_month") or "—")
    info[2].metric("This Week Friday", meta.get("this_week_friday") or "—")
    info[3].metric("Previous Week Friday", meta.get("previous_week_friday") or "—")
    info[4].metric("Two Weeks Ago Friday", meta.get("two_weeks_ago_friday") or "—")

    comments = preview.get("comments") or {}
    charts = preview.get("charts") or {}

    st.markdown(
        f'<div class="report-banner">{meta.get("report_title") or "WEEKLY BUNKERING REPORT"}</div>',
        unsafe_allow_html=True,
    )
    page1 = st.columns([1.15, 0.85])
    with page1[0]:
        st.subheader("Korea Major 4 Refiners - VLSFO Premium Trends")
        st.plotly_chart(_line_fig(charts.get("korea_premium") or {}), width="stretch")
    with page1[1]:
        st.subheader("Korea Bunker Market")
        st.markdown(_comment_html(comments.get("korea") or "TBN"), unsafe_allow_html=True)
        strategy = comments.get("strategy") or []
        if strategy:
            st.caption("\n".join(strategy))

    st.markdown('<div class="report-banner">Worldwide Market</div>', unsafe_allow_html=True)
    page2 = st.columns([1.15, 0.85])
    with page2[0]:
        st.subheader("Worldwide Ports - VLSFO Bunker Price Trend")
        st.plotly_chart(_line_fig(charts.get("worldwide_vlsfo") or {}), width="stretch")
    with page2[1]:
        st.subheader("Worldwide Market for this week")
        st.markdown("**South Korea**")
        st.markdown(_comment_html(comments.get("korea_worldwide") or comments.get("korea") or "TBN"), unsafe_allow_html=True)
        st.markdown("**Singapore**")
        st.markdown(_comment_html(comments.get("singapore") or "TBN"), unsafe_allow_html=True)
        st.markdown("**China / Zhoushan**")
        st.markdown(_comment_html(comments.get("china") or "TBN"), unsafe_allow_html=True)
        st.markdown("**Japan**")
        st.markdown(_comment_html(comments.get("japan") or "TBN"), unsafe_allow_html=True)

    st.markdown('<div class="report-banner">SPREAD TREND · BUNKER WIRE - MOPS SINGAPORE 0.5%</div>', unsafe_allow_html=True)
    st.plotly_chart(_bar_fig(charts.get("spread") or {}), width="stretch")

    files = preview.get("files") or {}
    pdf_path = st.session_state.get("pdf_path") or files.get("pdf_path")
    excel_path = st.session_state.get("excel_path") or files.get("excel_path")
    pdf_bytes = _file_bytes(pdf_path)
    excel_bytes = _file_bytes(excel_path)

    st.markdown("### Report files")
    st.caption("Create / Update Report writes the dated Google sheet and generates the PDF. It does not send email.")
    downs = st.columns(2)
    with downs[0]:
        st.download_button(
            "Download PDF",
            data=pdf_bytes or b"",
            file_name=files.get("pdf") or "report.pdf",
            mime="application/pdf",
            disabled=pdf_bytes is None,
            width="stretch",
        )
    with downs[1]:
        st.download_button(
            "Download Excel",
            data=excel_bytes or b"",
            file_name=files.get("excel") or "report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            disabled=excel_bytes is None,
            width="stretch",
        )

    st.markdown("### Email Preview")
    st.caption("Review recipients, then send. Create Report never sends mail.")
    if "email_to" not in st.session_state:
        _hydrate_email(preview)
    st.text_area("To", key="email_to", height=80)
    st.text_area("CC", key="email_cc", height=70)
    st.text_input("Subject", key="email_subject")
    st.text_area("Body", key="email_body", height=160)
    st.caption(f"Attachment: {Path(pdf_path).name if pdf_bytes and pdf_path else 'none'}")

    if st.button("Send Email", type="primary"):
        if not gmail_ok:
            st.error("Gmail is not connected.")
        elif not (st.session_state.get("email_to") or "").strip():
            st.error("At least one TO recipient is required.")
        elif not pdf_bytes or not pdf_path:
            st.error("Create / Update Report first so a PDF attachment exists.")
        else:
            try:
                send_email(
                    to_text=st.session_state.get("email_to", ""),
                    cc_text=st.session_state.get("email_cc", ""),
                    subject=st.session_state.get("email_subject", ""),
                    body=st.session_state.get("email_body", ""),
                    attachment_path=Path(pdf_path),
                )
                st.success("Email sent successfully.")
            except GmailServiceError as exc:
                st.error(str(exc))
            except Exception as exc:
                st.error(f"Email send failed: {exc}")


if __name__ == "__main__":
    main()
