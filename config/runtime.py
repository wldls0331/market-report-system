"""Runtime helpers for local Windows vs Streamlit Cloud."""

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
    return False


def supports_browser_oauth() -> bool:
    """Desktop OAuth localhost callback is not available on Streamlit Cloud."""
    return not is_streamlit_cloud()
