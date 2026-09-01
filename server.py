"""Market Report HTML UI. Business logic stays in services/."""

from __future__ import annotations

import os

from flask import Flask, jsonify, render_template, request, send_file

from config.paths import PROJECT_ROOT
from services.data_service import DataServiceError
from services.gmail_service import GmailServiceError, authorize_gmail, send_email
from services.google_sheets_service import authorize_sheets
from services.preview_service import (
    build_preview,
    connection_status,
    create_report_files,
    public_preview,
    resolve_output_file,
)

app = Flask(
    __name__,
    template_folder=str(PROJECT_ROOT / "web" / "templates"),
    static_folder=str(PROJECT_ROOT / "web" / "static"),
)


def _bind_host_port() -> tuple[str, int]:
    """Local Flask: 127.0.0.1:8502. Render sets RENDER and PORT; gunicorn binds itself."""
    if os.environ.get("RENDER") or os.environ.get("RENDER_SERVICE_ID"):
        return "0.0.0.0", int(os.environ.get("PORT") or "8502")
    return "127.0.0.1", 8502


@app.get("/")
def index():
    return render_template("index.html")


@app.get("/api/status")
def api_status():
    return jsonify(connection_status())


@app.get("/api/preview")
def api_preview():
    try:
        return jsonify(public_preview(build_preview(sync_sheet=False)))
    except DataServiceError as exc:
        return jsonify({"error": str(exc), "status": connection_status()}), 400


@app.post("/api/refresh")
def api_refresh():
    try:
        return jsonify(public_preview(build_preview(sync_sheet=True)))
    except DataServiceError as exc:
        return jsonify({"error": str(exc), "status": connection_status()}), 400


@app.post("/api/create-report")
def api_create_report():
    payload = create_report_files()
    status = 400 if payload.get("error") else 200
    return jsonify(public_preview(payload)), status


@app.get("/api/download/<kind>")
def api_download(kind: str):
    if kind not in {"pdf", "excel"}:
        return jsonify({"error": "Unknown file type."}), 404
    try:
        path = resolve_output_file(kind)
    except DataServiceError as exc:
        return jsonify({"error": str(exc)}), 400
    if path is None:
        return jsonify({"error": f"No {kind.upper()} has been generated yet. Create / Update Report first."}), 404
    mime = (
        "application/pdf"
        if kind == "pdf"
        else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    return send_file(path, as_attachment=True, download_name=path.name, mimetype=mime)


@app.post("/api/send-email")
def api_send_email():
    body = request.get_json(silent=True) or {}
    try:
        pdf_path = resolve_output_file("pdf")
    except DataServiceError as exc:
        return jsonify({"error": str(exc)}), 400
    if pdf_path is None:
        return jsonify({"error": "Create / Update Report first so a PDF attachment exists."}), 400
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
