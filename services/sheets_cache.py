"""Short TTL cache so one INPUT read can serve status, report, and charts."""

from __future__ import annotations

import threading
import time
from typing import Any

INPUT_TTL_SECONDS = 10.0
STATUS_TTL_SECONDS = 15.0
TITLES_TTL_SECONDS = 15.0
RECIPIENTS_TTL_SECONDS = 20.0
GRID_TTL_SECONDS = 10.0

_lock = threading.Lock()
_store: dict[str, tuple[float, Any]] = {}


def cache_get(key: str) -> Any:
    with _lock:
        row = _store.get(key)
        if row is None:
            return None
        expires, value = row
        if time.monotonic() > expires:
            _store.pop(key, None)
            return None
        return value


def cache_set(key: str, value: Any, ttl: float) -> Any:
    with _lock:
        _store[key] = (time.monotonic() + ttl, value)
    return value


def cache_clear(*keys: str) -> None:
    with _lock:
        if not keys:
            _store.clear()
            return
        for key in keys:
            _store.pop(key, None)
