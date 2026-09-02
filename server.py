"""Market Report HTML UI. Business logic stays in services/."""

from __future__ import annotations

import io
import os
import time
import traceback
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file

from config.paths import PROJECT_ROOT
from services.data_service import DataServiceError
from services.excel_service import GenerateError
from services.gmail_service import GmailServiceError, authorize_gmail, send_email
from services.google_sheets_service import authorize_sheets
from services.html_pdf_service import HtmlPdfError
from services.news_service import get_news_payload, refresh_market_news_if_stale
from services.perf import timed
from services.preview_service import (
    build_preview,
    connection_status,
    create_report_files,
    ensure_excel_file,
    public_preview,
    render_html_preview_pdf,
)


app = Flask(
    __name__,
    template_folder=str(PROJECT_ROOT / "web" / "templates"),
    static_folder=str(PROJECT_ROOT / "web" / "static"),
)


def _bind_host_port() -> tuple[str, int]:
    """Local Flask: 127.0.0.1:8502. Render uses PORT and gunicorn binds itself."""
    port = int(os.environ.get("PORT") or "8502")
    if os.environ.get("RENDER") or os.environ.get("RENDER_SERVICE_ID"):
        return "0.0.0.0", port
    return "127.0.0.1", 8502 if "PORT" not in os.environ else port


def _print_url() -> str:
    return request.url_root.rstrip("/") + "/print"


def _error_payload(exc: BaseException) -> dict[str, str]:
    return {"error": f"{exc}\n\n{traceback.format_exc()}"}


def _read_file_bytes(path: Path) -> bytes:
    last: OSError | None = None
    for _ in range(6):
        try:
            data = path.read_bytes()
            if data:
                return data
            last = OSError(f"{path.name} is empty")
        except OSError as exc:
            last = exc
        time.sleep(0.35)
    raise OSError(f"Could not read {path} (file may be locked or empty): {last}") from last


def _send_download(path: Path, mime: str):
    payload = _read_file_bytes(path)
    return send_file(
        io.BytesIO(payload),
        as_attachment=True,
        download_name=path.name,
        mimetype=mime,
        max_age=0,
    )


def _preview_html() -> str:
    payload = public_preview(build_preview(sync_sheet=False))
    return render_template("print.html", preview=payload)


@app.get("/")
def index():
    with timed("render index"):
        return render_template("index.html")


@app.get("/health")
def health():
    return jsonify({"ok": True}), 200


@app.get("/print")
def print_view():
    try:
        payload = public_preview(build_preview(sync_sheet=False))
    except DataServiceError as exc:
        payload = {
            "error": str(exc),
            "meta": {},
            "comments": {},
            "charts": {},
            "supplier_premiums": [],
        }
    return render_template("print.html", preview=payload)


@app.get("/api/status")
def api_status():
    with timed("api status"):
        return jsonify(connection_status())


@app.get("/api/preview")
def api_preview():
    try:
        with timed("api preview"):
            return jsonify(public_preview(build_preview(sync_sheet=False)))
    except DataServiceError as exc:
        return jsonify({"error": str(exc), "status": connection_status()}), 400


@app.get("/api/report-data")
def api_report_data():
    try:
        with timed("api report-data"):
            return jsonify(public_preview(build_preview(sync_sheet=False, include_status=False)))
    except DataServiceError as exc:
        return jsonify({"error": str(exc)}), 400


@app.get("/api/chart-data")
def api_chart_data():
    try:
        with timed("api chart-data"):
            payload = public_preview(build_preview(sync_sheet=False, include_status=False))
        return jsonify({
            "charts": payload.get("charts") or {},
            "error": payload.get("error"),
        })
    except DataServiceError as exc:
        return jsonify({"error": str(exc), "charts": {}}), 400


@app.post("/api/refresh")
def api_refresh():
    try:
        with timed("api refresh"):
            return jsonify(public_preview(build_preview(sync_sheet=True)))
    except DataServiceError as exc:
        return jsonify({"error": str(exc), "status": connection_status()}), 400


@app.post("/api/create-report")
def api_create_report():
    try:
        refresh_market_news_if_stale()
    except Exception:
        pass
    payload = create_report_files()
    status = 400 if payload.get("error") else 200
    return jsonify(public_preview(payload)), status


@app.get("/api/news")
def api_news():
    return jsonify(get_news_payload(force=False))


@app.post("/api/news/refresh")
def api_news_refresh():
    return jsonify(get_news_payload(force=True))


@app.get("/api/download/<kind>")
def api_download(kind: str):
    if kind not in {"pdf", "excel"}:
        return jsonify({"error": "Unknown file type."}), 404
    try:
        if kind == "pdf":
            path = render_html_preview_pdf(_print_url(), html=_preview_html())
            return _send_download(path, "application/pdf")
        path = ensure_excel_file()
        return _send_download(
            path,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except (HtmlPdfError, GenerateError, DataServiceError) as exc:
        return jsonify(_error_payload(exc)), 400
    except Exception as exc:
        return jsonify(_error_payload(exc)), 500


@app.post("/api/send-email")
def api_send_email():
    body = request.get_json(silent=True) or {}
    try:
        pdf_path = render_html_preview_pdf(_print_url(), html=_preview_html())
    except DataServiceError as exc:
        return jsonify({"error": str(exc)}), 400
    except HtmlPdfError as exc:
        return jsonify(_error_payload(exc)), 400
    except Exception as exc:
        return jsonify(_error_payload(exc)), 500
    try:
        send_email(
            to_text=str(body.get("to") or ""),
            cc_text=str(body.get("cc") or ""),
            subject=str(body.get("subject") or ""),
            body=str(body.get("body") or ""),
            attachment_path=pdf_path,
        )
    except GmailServiceError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        return jsonify({"error": f"Email send failed: {exc}"}), 400
    return jsonify({"ok": True, "message": "Email sent successfully."})


@app.post("/api/authorize/sheets")
def api_authorize_sheets():
    try:
        authorize_sheets()
        return jsonify({"ok": True, **connection_status()})
    except DataServiceError as exc:
        return jsonify({"ok": False, "error": str(exc), **connection_status()}), 400
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc), **connection_status()}), 400


@app.post("/api/authorize/gmail")
def api_authorize_gmail():
    try:
        authorize_gmail()
        return jsonify({"ok": True, **connection_status()})
    except GmailServiceError as exc:
        return jsonify({"ok": False, "error": str(exc), **connection_status()}), 400
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc), **connection_status()}), 400


if __name__ == "__main__":
    host, port = _bind_host_port()
    app.run(host=host, port=port, debug=False, threaded=True)
