"""Runtime helpers for local Windows vs hosted Linux (Render / Streamlit Cloud)."""

from __future__ import annotations

import os
import platform
import sys
from pathlib import Path


def is_windows() -> bool:
    return platform.system() == "Windows" or sys.platform == "win32"


def is_streamlit_cloud() -> bool:
    if str(os.environ.get("STREAMLIT_CLOUD") or "").strip().lower() in {"1", "true"}:
        return True
    runtime_env = str(os.environ.get("STREAMLIT_RUNTIME_ENV") or "").strip().lower()
    if runtime_env in {"cloud", "streamlit-cloud"}:
        return True
    hostname = str(os.environ.get("HOSTNAME") or "")
    if hostname.startswith("streamlit"):
        return True
    if Path("/mount/src").exists():
        return True
    if Path("/home/appuser").exists() and not is_windows():
        return True
    cwd = str(Path.cwd()).replace("\\", "/")
    if cwd.startswith("/mount/src"):
        return True
    return False


def is_render() -> bool:
    return bool(
        str(os.environ.get("RENDER") or "").strip()
        or str(os.environ.get("RENDER_SERVICE_ID") or "").strip()
    )


def is_hosted() -> bool:
    """True on Render, Streamlit Cloud, and other public web hosts."""
    return is_render() or is_streamlit_cloud()


def supports_browser_oauth() -> bool:
    """Desktop OAuth localhost callback is not available on hosted platforms."""
    return not is_hosted()
