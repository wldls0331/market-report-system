"""Parse Korean Supplier Premium INPUT values as VLSFO/HSFO pairs."""

from __future__ import annotations

from typing import Any

TBN = "TBN"


def _parse_part(raw: str) -> float | None:
    text = (raw or "").strip().replace(",", "")
    if not text or text.upper() == TBN:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def parse_supplier_premium(raw: Any) -> dict[str, float | None]:
    """Split a supplier premium cell into VLSFO (left) and HSFO (right).

    Examples:
        "80/120" -> {"vlsfo": 80.0, "hsfo": 120.0}
        "80 / 120" -> {"vlsfo": 80.0, "hsfo": 120.0}
        "80.5/120.5" -> {"vlsfo": 80.5, "hsfo": 120.5}
        "80" or "80/" -> {"vlsfo": 80.0, "hsfo": None}
        "/120" -> {"vlsfo": None, "hsfo": 120.0}
        "" or "TBN" -> {"vlsfo": None, "hsfo": None}
    """
    if raw is None:
        return {"vlsfo": None, "hsfo": None}
    if isinstance(raw, bool):
        return {"vlsfo": None, "hsfo": None}
    if isinstance(raw, (int, float)):
        return {"vlsfo": float(raw), "hsfo": None}

    text = str(raw).strip()
    if not text or text.upper() == TBN:
        return {"vlsfo": None, "hsfo": None}

    if "/" not in text:
        return {"vlsfo": _parse_part(text), "hsfo": None}

    left, right, *rest = text.split("/")
    del rest
    return {"vlsfo": _parse_part(left), "hsfo": _parse_part(right)}


def premium_chart_value(raw: Any) -> float | None:
    """VLSFO only. HSFO is never used on the VLSFO Premium Trend chart."""
    return parse_supplier_premium(raw)["vlsfo"]


def format_premium_signed(value: float | None) -> str:
    if value is None:
        return TBN
    if value == int(value):
        return f"+{int(value)}"
    return f"+{value}"
