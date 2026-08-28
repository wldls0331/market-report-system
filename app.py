from __future__ import annotations

import base64
import datetime as dt
import html
from collections import defaultdict
from pathlib import Path

import streamlit as st

from config.cell_mapping import NUMBER_FIELDS, SHEET_NAME_SUFFIX, STRATEGY_FIELDS
from config.email_config import default_body, default_subject
from config.google_config import (
    google_sheet_url,
    invalid_google_oauth_secret_keys,
    missing_google_sheets_secret_keys,
)
from config.paths import find_template
from config.runtime import is_streamlit_cloud, supports_browser_oauth
from services.data_service import DataServiceError
from services.excel_service import GenerateError, generate_market_report, report_exists_for_date
from services.gmail_service import GmailServiceError, authorize_gmail, gmail_status, send_email
from services.google_sheets_service import (
    ReportData,
    active_recipient_emails,
    authorize_sheets,
    latest_report_sheet,
    load_email_recipients_result,
    load_report_data,
    sheets_status,
    spreadsheet_title,
)
from services.pricing_month_service import YearMonth, default_pricing_month
from services.working_day_service import (
    format_data_reference_display,
    format_pdf_filename,
    format_sheet_name,
    previous_week_last_working_day,
    singapore_holiday_name,
    validate_report_date,
)

st.set_page_config(page_title="Market Report", layout="wide", initial_sidebar_state="expanded")

st.markdown(
    """
    <style>
    .block-container { padding-top: 1.4rem; max-width: 1180px; }
    h1 { font-weight: 650; letter-spacing: -0.02em; }
    .hint { color: #6b7280; font-size: 0.86rem; margin-bottom: 0.55rem; }
    .draft-value { white-space: pre-wrap; margin: 0 0 0.7rem 0; }
    .stButton>button {
        width: 100%;
        height: 3rem;
        font-size: 1.05rem;
        font-weight: 600;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

PAPER_DRAFT_LABELS = {
    "paper_fo": "SING 0.5",
    "paper_hsfo": "SING 380",
    "paper_go": "GASOIL 10PPM",
}
PREMIUM_DRAFT_LABELS = {
    "hdo_premium": "HDO",
    "sk_premium": "SK",
    "soil_premium": "S-OIL",
    "gs_premium": "GS",
}


def _today() -> dt.date:
    return dt.date.today()


def _ui_sheet(name: str | None) -> str:
    if not name:
        return ""
    text = str(name)
    suffix = f" {SHEET_NAME_SUFFIX}"
    return text[: -len(suffix)] if text.endswith(suffix) else text


def _tbn(value) -> str:
    text = "" if value is None else str(value).strip()
    return text if text else "TBN"


def _show_field(label: str, value) -> None:
    st.markdown(f"**{label}**")
    rendered = html.escape(_tbn(value)).replace("\n", "<br>")
    st.markdown(f'<div class="draft-value">{rendered}</div>', unsafe_allow_html=True)


def _pricing_month_for(record: ReportData | None, data_reference_date: dt.date) -> YearMonth:
    if record and record.pricing_month:
        try:
            return YearMonth.parse(record.pricing_month)
        except ValueError:
            pass
    return default_pricing_month(data_reference_date)


def _show_pdf(path: Path) -> None:
    data = path.read_bytes()
    pdf_widget = getattr(st, "pdf", None)
    if callable(pdf_widget):
        pdf_widget(data, height=720)
        return
    encoded = base64.b64encode(data).decode("ascii")
    st.components.v1.html(
        f'<iframe src="data:application/pdf;base64,{encoded}" width="100%" height="720"></iframe>',
        height=740,
    )


def _render_connection_status() -> None:
    sheets_ok, sheets_label = sheets_status()
    gmail_ok, gmail_label = gmail_status()
    sheet_url = google_sheet_url()
    book_title = ""
    if sheets_ok:
        try:
            book_title = spreadsheet_title()
        except Exception:
            book_title = ""
    with st.sidebar:
        st.subheader("Connection Status")
        st.write(f"Google Sheets: **{sheets_label}**")
        if sheets_ok:
            st.write(f"Spreadsheet: **{book_title or '—'}**")
        st.write(f"Gmail: **{gmail_label}**")
        if sheet_url:
            st.link_button("Open Google Sheet", sheet_url)
        if not gmail_ok and supports_browser_oauth():
            if st.button("Authorize Gmail", key="auth_gmail"):
                try:
                    authorize_gmail()
                    st.success("Gmail authorized.")
                    st.rerun()
                except GmailServiceError as exc:
                    st.error(str(exc))
                except Exception as exc:
                    st.error(f"Gmail authorization failed: {exc}")
        elif not gmail_ok and is_streamlit_cloud():
            st.caption("Add [gmail] client_id, client_secret, and refresh_token in Streamlit Secrets.")
        if not sheets_ok:
            if sheets_label == "Authorization Required" and supports_browser_oauth():
                if st.button("Authorize Google Sheets", key="auth_sheets"):
                    try:
                        authorize_sheets()
                        st.success("Google Sheets authorized.")
                        st.rerun()
                    except DataServiceError as exc:
                        st.error(str(exc))
                    except Exception as exc:
                        st.error(f"Google Sheets authorization failed: {exc}")
            else:
                missing = missing_google_sheets_secret_keys()
                if missing:
                    st.caption("Streamlit Secrets missing: " + ", ".join(missing))
                invalid = invalid_google_oauth_secret_keys()
                if invalid:
                    st.caption("Streamlit Secrets invalid: " + " ".join(invalid))
                err = st.session_state.get("_sheets_error")
                if err:
                    st.caption(str(err))
                elif not missing and not invalid and is_streamlit_cloud():
                    st.caption("Add [google] sheet_id, client_id, client_secret, and refresh_token in Streamlit Secrets.")
                elif not missing and not invalid:
                    st.caption("Set GOOGLE_SHEET_ID in .env or Streamlit Secrets, then authorize Google Sheets.")


def _render_draft(record: ReportData | None, report_date: dt.date, data_reference_date: dt.date) -> None:
    st.markdown("### Draft Preview")
    st.caption("Market data is managed in Google Sheets.")
    sheet_url = google_sheet_url()
    if sheet_url:
        st.link_button("Open Google Sheet", sheet_url)
    inputs = record.inputs if record else {}
    month = _pricing_month_for(record, data_reference_date)

    st.markdown("#### Report Information")
    info_cols = st.columns(3)
    with info_cols[0]:
        _show_field("Report Date", report_date.isoformat())
        _show_field("Data Reference Date", format_data_reference_display(data_reference_date))
    with info_cols[1]:
        _show_field("Pricing Month", month.label())
        if record and record.sheet_name:
            _show_field("Google Sheet", record.sheet_name)
    with info_cols[2]:
        _show_field("Last Updated", record.last_saved_label() if record else "")
        _show_field("Updated By", record.updated_by if record else "")

    st.markdown("#### Paper / MOPS")
    paper_fields = [item for item in NUMBER_FIELDS if item.section == "paper"]
    paper_cols = st.columns(len(paper_fields))
    for col, item in zip(paper_cols, paper_fields):
        with col:
            _show_field(PAPER_DRAFT_LABELS.get(item.key, item.label), inputs.get(item.key))

    st.markdown("#### Bunker Market Price")
    bunker_fields = [item for item in NUMBER_FIELDS if item.section == "bunker"]
    bunker_groups: dict[str, list] = defaultdict(list)
    for item in bunker_fields:
        bunker_groups[item.group].append(item)
    group_cols = st.columns(len(bunker_groups))
    for col, (group, fields) in zip(group_cols, bunker_groups.items()):
        with col:
            st.markdown(f"**{group}**")
            for item in fields:
                _show_field(item.label, inputs.get(item.key))

    st.markdown("#### Korea Refinery Premium")
    premium_fields = [item for item in NUMBER_FIELDS if item.section == "premium"]
    premium_cols = st.columns(len(premium_fields))
    for col, item in zip(premium_cols, premium_fields):
        with col:
            _show_field(PREMIUM_DRAFT_LABELS.get(item.key, item.label), inputs.get(item.key))

    st.markdown("#### Korea Market")
    _show_field("Commentary", inputs.get("comment_korea"))

    st.markdown("#### Worldwide Market")
    world_cols = st.columns(2)
    with world_cols[0]:
        _show_field("South Korea", inputs.get("comment_korea_worldwide") or inputs.get("comment_korea"))
        _show_field("Singapore", inputs.get("comment_singapore"))
    with world_cols[1]:
        _show_field("China", inputs.get("comment_china"))
        _show_field("Japan", inputs.get("comment_japan"))

    st.markdown("#### Strategy")
    strategy_cols = st.columns(3)
    for col, item in zip(strategy_cols, STRATEGY_FIELDS):
        with col:
            _show_field(item.label, inputs.get(item.key))


def main() -> None:
    st.session_state.pop("_sheets_status_result", None)
    st.session_state.pop("_spreadsheet_title_live", None)
    st.session_state.pop("_sheets_error", None)
    st.title("Market Report")
    _sheets_ok, sheets_label = sheets_status()
    gmail_ok, gmail_label = gmail_status()
    st.caption(
        "Review Google Sheets data, generate Excel/PDF, then send email. "
        "Market data is not edited in this app."
    )
    st.caption(f"Google Sheets: {sheets_label} · Gmail: {gmail_label}")

    _render_connection_status()

    try:
        template = find_template()
        st.caption(f"Template: `{template.name}`")
    except FileNotFoundError as exc:
        st.error(str(exc))
        st.stop()

    default_report_date = _today()
    if _sheets_ok:
        try:
            latest_available = latest_report_sheet()
            if latest_available:
                default_report_date = latest_available[0]
        except Exception:
            pass

    top_left, top_mid = st.columns([1.2, 1])
    with top_left:
        report_date = st.date_input("Report Date", value=default_report_date, format="YYYY-MM-DD")
        data_reference_date = previous_week_last_working_day(report_date)
        st.markdown("**Data Reference Date**")
        st.caption(
            f"{format_data_reference_display(data_reference_date)} "
            "(Previous Week's Last Working Day)"
        )

    date_ok = True
    try:
        validate_report_date(report_date)
    except Exception as exc:
        date_ok = False
        st.warning(str(exc))
        if report_date.weekday() >= 5:
            st.info("Weekends are not Singapore working days. A report cannot be generated.")
        elif singapore_holiday_name(report_date):
            st.info("Singapore public holidays cannot be used as Report Date.")

    record = None
    load_error = None
    try:
        record = load_report_data(report_date)
    except DataServiceError as exc:
        load_error = str(exc)
        st.error(load_error)

    pricing_month = _pricing_month_for(
        record, record.data_reference_date if record else data_reference_date
    )
    with top_mid:
        st.markdown("**Pricing Month**")
        st.caption(pricing_month.label())

    refresh_col, sheet_col = st.columns([1, 1])
    with refresh_col:
        if st.button("Refresh Data"):
            st.rerun()
    sheet_url = google_sheet_url()
    with sheet_col:
        if sheet_url:
            st.link_button("Open Google Sheet", sheet_url)

    report_exists = False
    if date_ok:
        try:
            report_exists = report_exists_for_date(report_date)
        except Exception:
            report_exists = False
        target_sheet = format_sheet_name(report_date)
        if report_exists:
            st.caption(f"Existing Excel sheet `{_ui_sheet(target_sheet)}` will be updated.")

    if load_error is None and record is None:
        latest = None
        try:
            latest = latest_report_sheet()
        except DataServiceError:
            latest = None
        if latest:
            st.info(
                "No report sheet for this Report Date in Google Sheets. "
                f"Latest available sheet: `{latest[1]}`"
            )
        else:
            st.info("No report sheet for this Report Date in Google Sheets.")

    recipients_warning = None
    if _sheets_ok:
        try:
            _recipients, recipients_warning = load_email_recipients_result()
        except DataServiceError as exc:
            recipients_warning = str(exc)
    if recipients_warning:
        st.warning(recipients_warning)

    _render_draft(
        record,
        report_date,
        record.data_reference_date if record else data_reference_date,
    )

    st.markdown("---")
    st.markdown("### Report")
    create_label = "Update Report" if report_exists else "Create Report"
    create_clicked = st.button(create_label, type="primary", disabled=not date_ok)
    if create_clicked:
        if not date_ok:
            st.error("Report Date must be a Singapore working day.")
        else:
            latest = None
            try:
                latest = load_report_data(report_date)
            except DataServiceError as exc:
                st.error(str(exc))
            if latest is None:
                st.error("No report sheet for this Report Date in Google Sheets.")
            else:
                spinner_msg = (
                    "Updating the existing sheet and regenerating the PDF..."
                    if report_exists
                    else "Copying the Excel sheet and generating the PDF..."
                )
                with st.spinner(spinner_msg):
                    try:
                        generate_month = _pricing_month_for(latest, latest.data_reference_date)
                        result = generate_market_report(
                            report_date,
                            latest.inputs,
                            export_pdf=True,
                            pricing_month=generate_month,
                            extra_cells=latest.extra_cells,
                            data_reference_date=latest.data_reference_date,
                        )
                        expected_pdf = format_pdf_filename(report_date)
                        pdf_path = str(result.pdf_path) if result.pdf_path else None
                        if pdf_path and Path(pdf_path).name != expected_pdf:
                            pdf_path = None
                        to_list, cc_list = [], []
                        try:
                            to_list, cc_list = active_recipient_emails()
                        except DataServiceError as exc:
                            st.warning(str(exc))
                        st.session_state["result"] = {
                            "excel_path": str(result.excel_path),
                            "pdf_path": pdf_path,
                            "sheet_name": result.sheet_name,
                            "previous_sheet": result.previous_sheet,
                            "warnings": result.warnings,
                            "pdf_page_count": result.pdf_page_count,
                            "pricing_month": result.pricing_month,
                            "is_update": result.is_update,
                            "used_saved_at": latest.last_saved_label(),
                            "data_reference_date": latest.data_reference_date.isoformat(),
                            "report_date": report_date.isoformat(),
                        }
                        st.session_state["email_to"] = "\n".join(to_list)
                        st.session_state["email_cc"] = "\n".join(cc_list)
                        st.session_state["email_subject"] = default_subject(report_date)
                        st.session_state["email_body"] = default_body(report_date)
                        st.rerun()
                    except GenerateError as exc:
                        st.session_state.pop("result", None)
                        st.error(str(exc))
                    except Exception as exc:
                        st.session_state.pop("result", None)
                        st.error(f"Report generation failed: {exc}")

    result = st.session_state.get("result")
    matching_result = bool(result) and result.get("report_date") == report_date.isoformat()
    if matching_result:
        st.success("Market Report generated successfully.")
        extras = []
        if result.get("data_reference_date"):
            extras.append(
                "Data Reference Date "
                + format_data_reference_display(dt.date.fromisoformat(result["data_reference_date"]))
            )
        if result.get("pdf_page_count"):
            extras.append(f"PDF {result['pdf_page_count']} page(s)")
        extra_note = (" · " + " · ".join(extras)) if extras else ""
        if result.get("is_update"):
            st.caption(f"Updated sheet `{_ui_sheet(result['sheet_name'])}`{extra_note}")
        else:
            st.caption(
                f"Sheet `{_ui_sheet(result['previous_sheet'])}` → `{_ui_sheet(result['sheet_name'])}`{extra_note}"
            )
        if result.get("used_saved_at"):
            st.caption(f"Generated from Google Sheets data last updated {result['used_saved_at']}")
        for warning in result.get("warnings") or []:
            st.warning(warning)

        excel_path = result["excel_path"]
        pdf_path = result.get("pdf_path")
        expected_name = format_pdf_filename(report_date)
        pdf_ok = bool(pdf_path) and Path(str(pdf_path)).exists() and Path(str(pdf_path)).name == expected_name

        st.markdown("#### PDF Preview")
        if pdf_ok:
            st.caption(f"Generated PDF: `{expected_name}`")
            _show_pdf(Path(str(pdf_path)))
        else:
            st.warning("PDF was not generated for this Report Date.")

        down_left, down_right = st.columns(2)
        with down_left:
            if pdf_ok:
                with open(str(pdf_path), "rb") as handle:
                    st.download_button(
                        "Download PDF",
                        data=handle.read(),
                        file_name=expected_name,
                        mime="application/pdf",
                    )
        with down_right:
            with open(excel_path, "rb") as handle:
                st.download_button(
                    "Download Excel",
                    data=handle.read(),
                    file_name=Path(excel_path).name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

        st.markdown("---")
        st.markdown("### Email Preview")
        st.caption("Create / Update Report never sends mail. Review, then click Send Email.")
        st.text_area("To", key="email_to", height=80)
        st.text_area("CC", key="email_cc", height=70)
        st.text_input("Subject", key="email_subject")
        st.text_area("Body", key="email_body", height=160)
        if pdf_ok:
            st.caption(f"Attachment: `{expected_name}`")
        else:
            st.warning("No matching PDF attachment for this Report Date.")

        if st.button("Send Email", type="primary"):
            if not gmail_ok:
                st.error("Gmail is not connected. Authorize Gmail first.")
            elif not (st.session_state.get("email_to") or "").strip():
                st.error("At least one TO recipient is required.")
            elif not pdf_ok:
                st.error("PDF attachment for this Report Date was not found.")
            elif not (st.session_state.get("email_subject") or "").strip():
                st.error("Subject is required.")
            elif not (st.session_state.get("email_body") or "").strip():
                st.error("Body is required.")
            else:
                try:
                    send_email(
                        to_text=st.session_state.get("email_to", ""),
                        cc_text=st.session_state.get("email_cc", ""),
                        subject=st.session_state.get("email_subject", ""),
                        body=st.session_state.get("email_body", ""),
                        attachment_path=Path(str(pdf_path)),
                    )
                    st.success("Email sent successfully.")
                except GmailServiceError as exc:
                    st.error(str(exc))
                except Exception as exc:
                    st.error(f"Email send failed: {exc}")


if __name__ == "__main__":
    main()
