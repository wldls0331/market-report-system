"""Streamlit UI for the Market Report. Same preview, news, HTML PDF, Excel, and Gmail as Flask."""

from __future__ import annotations

import base64
import html
from pathlib import Path

import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components

from config.google_config import (
    google_sheet_url,
    invalid_google_oauth_secret_keys,
    missing_google_sheets_secret_keys,
)
from config.runtime import is_streamlit_cloud, supports_browser_oauth
from services.excel_service import GenerateError
from services.gmail_service import GmailServiceError, authorize_gmail, send_email
from services.google_sheets_service import DataServiceError, authorize_sheets
from services.html_pdf_service import HtmlPdfError
from services.news_service import get_news_payload, refresh_market_news_if_stale
from services.preview_service import (
    build_preview,
    connection_status,
    create_report_files,
    ensure_excel_file,
    public_preview,
    render_html_preview_pdf,
)

PALETTE = ["#1b365d", "#6b7280", "#c4a35a", "#334155", "#94a3b8"]

st.set_page_config(page_title="Market Report", layout="wide", initial_sidebar_state="collapsed")

st.markdown(
    """
    <style>
    [data-testid="stToolbar"], [data-testid="stDecoration"],
    [data-testid="stStatusWidget"], .stDeployButton,
    footer { display: none !important; }
    .block-container { padding-top: 1.1rem; max-width: 1280px; }
    h1 { font-weight: 650; letter-spacing: -0.02em; color: #1b365d; }
    .status-ok { color: #166534; font-weight: 600; }
    .status-bad { color: #9f1239; font-weight: 600; }
    .report-banner {
        background: #1b365d; color: #fff; padding: 0.85rem 1rem;
        border-bottom: 3px solid #c4a35a; border-radius: 8px 8px 0 0;
        font-weight: 650; letter-spacing: 0.04em;
    }
    .report-banner .sub {
        display: block; margin-top: 4px; font-size: 0.8rem;
        font-weight: 500; letter-spacing: 0; opacity: 0.9;
    }
    .chart-meta { color: #6b7280; font-size: 0.78rem; margin: 0 0 0.4rem; }
    .comment-box {
        white-space: pre-wrap; line-height: 1.55; font-size: 0.95rem;
        margin: 0 0 0.8rem 0;
    }
    .premium-item {
        border: 1px solid #e5e7eb; border-radius: 8px;
        padding: 10px 12px; background: #f8fafc; margin-bottom: 10px;
    }
    .premium-item strong { display: block; color: #1b365d; margin-bottom: 4px; }
    .premium-item span { display: block; font-size: 0.86rem; line-height: 1.45; }
    .news-takeaway {
        margin: 0 0 1rem; padding: 12px 14px; background: #f8fafc;
        border: 1px solid #e5e7eb; border-left: 3px solid #c4a35a; border-radius: 6px;
    }
    .news-item {
        border: 1px solid #e5e7eb; border-radius: 8px;
        padding: 12px 14px; margin-bottom: 10px; background: #fff;
    }
    .news-kicker { color: #c4a35a; font-size: 0.75rem; font-weight: 650; }
    .news-item h4 { margin: 4px 0 6px; color: #1b365d; font-size: 0.98rem; }
    .news-item p { margin: 0; line-height: 1.5; }
    .news-meta { color: #6b7280; font-size: 0.8rem; margin-top: 6px !important; }
    .hint { color: #6b7280; font-size: 0.85rem; }
    </style>
    """,
    unsafe_allow_html=True,
)


def _comment_html(text: str) -> str:
    return f'<p class="comment-box">{html.escape(text or "TBN").replace(chr(10), "<br>")}</p>'


def _status_html(label: str, value: str, ok: bool) -> str:
    klass = "status-ok" if ok else "status-bad"
    return f'{label}: <span class="{klass}">{html.escape(value or "—")}</span>'


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
                marker={"size": 6},
                connectgaps=True,
            )
        )
    fig.update_layout(
        height=340,
        margin={"l": 40, "r": 16, "t": 8, "b": 40},
        paper_bgcolor="white",
        plot_bgcolor="white",
        legend={"orientation": "h", "y": -0.22, "font": {"size": 11}},
        xaxis={"showgrid": False, "tickfont": {"color": "#6b7280", "size": 11}},
        yaxis={"gridcolor": "#eef2f7", "tickfont": {"color": "#6b7280", "size": 11}},
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
    st.session_state["preview"] = preview
    if hydrate_email:
        _hydrate_email(preview)
    return preview


def _load_from_sheets(*, sync_sheet: bool) -> dict:
    return _store_preview(build_preview(sync_sheet=sync_sheet), hydrate_email=True)


def _create_from_sheets() -> dict:
    try:
        refresh_market_news_if_stale()
    except Exception:
        pass
    st.session_state.pop("news", None)
    return _store_preview(create_report_files(), hydrate_email=True)


def _trigger_download(data: bytes, filename: str, mime: str) -> None:
    b64 = base64.b64encode(data).decode("ascii")
    safe_name = html.escape(filename, quote=True)
    components.html(
        (
            f'<a id="auto-dl" href="data:{mime};base64,{b64}" download="{safe_name}"></a>'
            '<script>document.getElementById("auto-dl").click();</script>'
        ),
        height=0,
    )


def _latest_pdf_path() -> Path | None:
    path_text = st.session_state.get("pdf_path")
    if not path_text:
        return None
    path = Path(path_text)
    return path if path.exists() else None


def _render_pdf() -> Path:
    path = render_html_preview_pdf()
    st.session_state["pdf_path"] = str(path)
    st.session_state["pdf_name"] = path.name
    st.session_state["pdf_bytes"] = path.read_bytes()
    preview = st.session_state.get("preview") or {}
    files = dict(preview.get("files") or {})
    files["pdf"] = path.name
    files["has_pdf"] = True
    preview["files"] = files
    email = dict(preview.get("email") or {})
    email["attachment"] = path.name
    preview["email"] = email
    st.session_state["preview"] = preview
    return path


@st.fragment
def news_panel() -> None:
    head_l, head_r = st.columns([5, 1.2])
    with head_l:
        st.markdown("### Market News Summary")
        window = ((st.session_state.get("news") or {}).get("window") or {}).get("label") or "Last 7 days"
        st.caption(window)
    with head_r:
        refresh_news = st.button("Refresh News", width="stretch")

    if refresh_news or "news" not in st.session_state:
        label = "Refreshing last 7 days of market news…" if refresh_news else "Loading recent market news..."
        with st.spinner(label):
            st.session_state["news"] = get_news_payload(force=refresh_news)

    news = st.session_state.get("news") or {}
    items = news.get("items") or []
    if news.get("error") and not items:
        st.caption(news["error"])
        return

    if news.get("takeaway"):
        st.markdown(
            f'<div class="news-takeaway"><h4>Weekly Market Takeaway</h4>'
            f'<p>{html.escape(news["takeaway"])}</p></div>',
            unsafe_allow_html=True,
        )

    for item in items:
        st.markdown(
            (
                f'<article class="news-item">'
                f'<span class="news-kicker">{html.escape(item.get("category_label") or "")}</span>'
                f'<h4>{html.escape(item.get("headline") or "")}</h4>'
                f'<p>{html.escape(item.get("summary") or "")}</p>'
                f'<p class="news-meta">{html.escape(item.get("source") or "Unknown")} | '
                f'{html.escape(item.get("published_date") or "")}</p>'
                f"</article>"
            ),
            unsafe_allow_html=True,
        )

    if news.get("stale"):
        st.caption("Showing the last saved briefing. Refresh News for the latest 7 days.")
    elif not items:
        st.caption("No bunker-relevant headlines in the last 7 days.")
    if refresh_news and not (news.get("error") and not items):
        st.success("Market News Summary updated.")


def main() -> None:
    st.title("Market Report")

    if "preview" not in st.session_state:
        with st.spinner("Loading Market Report..."):
            try:
                _load_from_sheets(sync_sheet=False)
            except DataServiceError as exc:
                st.session_state["preview"] = {
                    "error": str(exc),
                    "status": connection_status(),
                    "meta": None,
                    "comments": {},
                    "charts": {},
                    "supplier_premiums": [],
                    "email": {},
                    "files": {},
                }

    preview = st.session_state.get("preview") or {}
    status = preview.get("status") or connection_status()
    sheets_ok = bool(status.get("sheets_ok"))
    gmail_ok = bool(status.get("gmail_ok"))

    top = st.columns([2.2, 2.2, 1.4, 1.6, 1.6])
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
        if not sheets_ok and status.get("sheets") == "Authorization Required" and supports_browser_oauth():
            if st.button("Authorize Sheets", width="stretch"):
                try:
                    authorize_sheets()
                    st.rerun()
                except Exception as exc:
                    st.error(str(exc))
        elif not gmail_ok and supports_browser_oauth() and not is_streamlit_cloud():
            if st.button("Authorize Gmail", width="stretch"):
                try:
                    authorize_gmail()
                    st.rerun()
                except Exception as exc:
                    st.error(str(exc))

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
        elif preview.get("error"):
            st.error(preview["error"])

    if not gmail_ok and is_streamlit_cloud():
        st.caption("Add [gmail] client_id, client_secret, and refresh_token in Streamlit Secrets.")

    flash = st.session_state.pop("flash", None)
    if flash:
        kind, text = flash
        if kind == "error":
            st.error(text)
        elif kind == "warning":
            st.warning(text)
        else:
            st.success(text)

    if refresh:
        with st.spinner("Reading the latest Google Sheets INPUT and report tabs…"):
            try:
                _load_from_sheets(sync_sheet=True)
            except DataServiceError as exc:
                st.error(str(exc))
                st.stop()
        st.session_state["flash"] = ("ok", "Updated just now")
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
    premiums = preview.get("supplier_premiums") or []
    updated = meta.get("updated_display") or meta.get("data_reference_date") or "—"
    report_date_display = meta.get("report_date_display") or meta.get("report_date") or "—"

    st.markdown(
        f'<div class="report-banner">WEEKLY BUNKERING REPORT'
        f'<span class="sub">Report Date {html.escape(str(report_date_display))}</span></div>',
        unsafe_allow_html=True,
    )
    page1 = st.columns([1.15, 0.85])
    with page1[0]:
        st.markdown("#### Korea Major 4 Refiners")
        st.caption("VLSFO Premium Trend")
        st.markdown(
            f'<p class="chart-meta">Updated: {html.escape(str(updated))} &nbsp; Unit: USD/MT</p>',
            unsafe_allow_html=True,
        )
        st.plotly_chart(_line_fig(charts.get("korea_premium") or {}), width="stretch")
    with page1[1]:
        st.markdown("#### Korean Supplier Premium")
        if premiums:
            cards = "".join(
                (
                    f'<div class="premium-item"><strong>{html.escape(row.get("label") or "")}</strong>'
                    f'<span>VLSFO: {html.escape(row.get("vlsfo") or "TBN")}</span>'
                    f'<span>HSFO: {html.escape(row.get("hsfo") or "TBN")}</span></div>'
                )
                for row in premiums
            )
            st.markdown(cards, unsafe_allow_html=True)
        else:
            st.caption("TBN")

    st.markdown(
        '<div class="report-banner">Worldwide Ports'
        '<span class="sub">VLSFO Bunker Price Trend</span></div>',
        unsafe_allow_html=True,
    )
    page2 = st.columns([1.15, 0.85])
    with page2[0]:
        st.markdown(
            f'<p class="chart-meta">Updated: {html.escape(str(updated))} &nbsp; Unit: USD/MT</p>',
            unsafe_allow_html=True,
        )
        st.plotly_chart(_line_fig(charts.get("worldwide_vlsfo") or {}), width="stretch")
    with page2[1]:
        st.markdown("#### Worldwide Market for this week")
        st.markdown("**South Korea – Southern**")
        st.markdown(_comment_html(comments.get("korea_worldwide") or "TBN"), unsafe_allow_html=True)
        st.markdown("**Singapore**")
        st.markdown(_comment_html(comments.get("singapore") or "TBN"), unsafe_allow_html=True)
        st.markdown("**China**")
        st.markdown(_comment_html(comments.get("china") or "TBN"), unsafe_allow_html=True)
        st.markdown("**Japan**")
        st.markdown(_comment_html(comments.get("japan") or "TBN"), unsafe_allow_html=True)

    news_panel()

    st.markdown("### Report")
    st.caption(
        "Create / Update Report writes INPUT into the dated Google sheet and generates Excel. "
        "Download PDF saves the Report Preview as a 2-page A4 PDF. Neither action sends email."
    )
    actions = st.columns(3)
    with actions[0]:
        create = st.button(
            "Create / Update Report",
            type="primary",
            width="stretch",
            disabled=meta.get("date_ok") is False,
        )
    with actions[1]:
        download_pdf = st.button("Download PDF", width="stretch")
    with actions[2]:
        download_excel = st.button("Download Excel", width="stretch")

    if create:
        with st.spinner("Updating the dated sheet and Excel file…"):
            try:
                preview = _create_from_sheets()
            except DataServiceError as exc:
                st.error(str(exc))
                st.stop()
        if preview.get("error"):
            st.error(preview["error"])
        else:
            st.session_state["flash"] = (
                "ok",
                "Market Report generated. Excel is ready. Download PDF from the current Preview.",
            )
            st.rerun()

    if download_pdf:
        with st.spinner("Rendering Preview to PDF…"):
            try:
                path = _render_pdf()
            except (HtmlPdfError, DataServiceError) as exc:
                st.error(str(exc))
            except Exception as exc:
                st.error(f"PDF generation failed: {exc}")
            else:
                _trigger_download(
                    st.session_state["pdf_bytes"],
                    path.name,
                    "application/pdf",
                )
                st.success(f"{path.name} downloaded.")

    if download_excel:
        with st.spinner("Preparing Excel…"):
            try:
                path = ensure_excel_file()
            except (GenerateError, DataServiceError) as exc:
                st.error(str(exc))
            except Exception as exc:
                st.error(f"Excel download failed: {exc}")
            else:
                data = path.read_bytes()
                st.session_state["excel_path"] = str(path)
                st.session_state["excel_name"] = path.name
                _trigger_download(
                    data,
                    path.name,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                st.success(f"{path.name} downloaded.")

    files = (st.session_state.get("preview") or {}).get("files") or {}

    st.markdown("### Email Preview")
    st.caption("Review recipients and copy, then send. Create Report never sends mail.")
    if "email_to" not in st.session_state:
        _hydrate_email(preview)
    st.text_area("To", key="email_to", height=80)
    st.text_area("CC", key="email_cc", height=70)
    st.text_input("Subject", key="email_subject")
    st.text_area("Body", key="email_body", height=160)
    attachment_name = (
        Path(st.session_state["pdf_path"]).name
        if _latest_pdf_path()
        else files.get("pdf")
    )
    st.caption(f"Attachment: {attachment_name or 'none'}")

    if st.button("Send Email", type="primary"):
        if not gmail_ok:
            st.error("Gmail is not connected.")
        elif not (st.session_state.get("email_to") or "").strip():
            st.error("At least one TO recipient is required.")
        else:
            with st.spinner("Sending email…"):
                try:
                    pdf_path = _latest_pdf_path() or _render_pdf()
                    send_email(
                        to_text=st.session_state.get("email_to", ""),
                        cc_text=st.session_state.get("email_cc", ""),
                        subject=st.session_state.get("email_subject", ""),
                        body=st.session_state.get("email_body", ""),
                        attachment_path=pdf_path,
                    )
                    st.success("Email sent successfully.")
                except (HtmlPdfError, DataServiceError, GmailServiceError) as exc:
                    st.error(str(exc))
                except Exception as exc:
                    st.error(f"Email send failed: {exc}")


if __name__ == "__main__":
    main()
