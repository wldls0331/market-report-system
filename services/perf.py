"""Lightweight [PERF] timing for localhost request stages."""

from __future__ import annotations

import time
from contextlib import contextmanager
from typing import Iterator


@contextmanager
def timed(label: str) -> Iterator[None]:
    start = time.perf_counter()
    try:
        yield
    finally:
        elapsed = time.perf_counter() - start
        print(f"[PERF] {label}: {elapsed:.2f}s", flush=True)
