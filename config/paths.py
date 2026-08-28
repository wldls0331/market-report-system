from __future__ import annotations

from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
TEMPLATE_DIR = PROJECT_ROOT / "templates"
OUTPUT_DIR = PROJECT_ROOT / "output"
DATA_DIR = PROJECT_ROOT / "data"


def session_output_dir() -> Path:
    """Per-browser-session output on Streamlit Cloud; shared output/ locally."""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    try:
        from config.runtime import is_streamlit_cloud

        if not is_streamlit_cloud():
            return OUTPUT_DIR
        from streamlit.runtime.scriptrunner import get_script_run_ctx

        ctx = get_script_run_ctx()
        session_id = getattr(ctx, "session_id", "") if ctx is not None else ""
        if session_id:
            path = OUTPUT_DIR / "sessions" / str(session_id)
            path.mkdir(parents=True, exist_ok=True)
            return path
    except Exception:
        pass
    return OUTPUT_DIR


def find_template() -> Path:
    if not TEMPLATE_DIR.exists():
        raise FileNotFoundError(f"Templates folder not found: {TEMPLATE_DIR}")

    files = [
        path
        for path in TEMPLATE_DIR.iterdir()
        if path.is_file()
        and path.suffix.lower() in {".xlsx", ".xlsm"}
        and not path.name.startswith("~$")
    ]
    if not files:
        raise FileNotFoundError("No Excel file found in the templates folder.")

    preferred = [
        path
        for path in files
        if "bunkering" in path.name.lower() or "weekly" in path.name.lower()
    ]
    return sorted(preferred or files, key=lambda p: p.stat().st_mtime, reverse=True)[0]
