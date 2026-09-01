"""Weekly chart window helpers for the Bunkering Report.

Korea Major 4 Refiners VLSFO Premium:
  categories C22:E22 historical dates, F22 = 'This week assumption'
  values C23:F26 (HyunDai / SK / SOIL / GS)

Worldwide Ports VLSFO Bunker Price:
  categories P22:R22 historical dates, S22 = '금주 예상'
  values P23:S26 (Korea-South / Singapore / Nagoya / Zhoushan)

On each new weekly sheet, shift C/D/E (and P/Q/R) left and land the new
historical date — Data Reference Date — in E22 / R22. Do not convert the
series to daily points. Skip the shift when that date is already in E22.
"""

from __future__ import annotations

import datetime as dt
from collections.abc import Callable
from typing import Any

from config.cell_mapping import (
    KOREA_CHART_DATE_CELLS,
    KOREA_CHART_VALUE_COLS,
    KOREA_CHART_VALUE_ROWS,
    WORLDWIDE_CHART_DATE_CELLS,
    WORLDWIDE_CHART_VALUE_COLS,
    WORLDWIDE_CHART_VALUE_ROWS,
)

GetValue = Callable[[str], Any]
SetValue = Callable[[str, Any], None]


def excel_value_as_date(value: Any) -> dt.date | None:
    if value is None or value == "":
        return None
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if isinstance(value, (int, float)):
        try:
            return (dt.datetime(1899, 12, 30) + dt.timedelta(days=int(value))).date()
        except Exception:
            return None
    return None


def needs_weekly_shift(end_date_value: Any, data_reference_date: dt.date) -> bool:
    current = excel_value_as_date(end_date_value)
    if current is None:
        return True
    return current != data_reference_date


def _shift_three_dates(get_value: GetValue, set_value: SetValue, cells: tuple[str, str, str], new_date: dt.date) -> None:
    left, mid, right = cells
    set_value(left, get_value(mid))
    set_value(mid, get_value(right))
    set_value(right, dt.datetime.combine(new_date, dt.time()))


def _shift_value_window(
    get_value: GetValue,
    set_value: SetValue,
    cols: tuple[str, ...],
    rows: tuple[int, ...],
) -> None:
    for row in rows:
        values = [get_value(f"{col}{row}") for col in cols]
        for index, col in enumerate(cols[:-1]):
            set_value(f"{col}{row}", values[index + 1])


def shift_weekly_chart_window(
    get_value: GetValue,
    set_value: SetValue,
    data_reference_date: dt.date,
) -> bool:
    """Shift historical weekly columns left. Return True if a shift ran."""
    if not needs_weekly_shift(get_value(KOREA_CHART_DATE_CELLS[2]), data_reference_date):
        return False
    _shift_three_dates(get_value, set_value, KOREA_CHART_DATE_CELLS, data_reference_date)
    _shift_value_window(get_value, set_value, KOREA_CHART_VALUE_COLS, KOREA_CHART_VALUE_ROWS)
    _shift_three_dates(get_value, set_value, WORLDWIDE_CHART_DATE_CELLS, data_reference_date)
    _shift_value_window(get_value, set_value, WORLDWIDE_CHART_VALUE_COLS, WORLDWIDE_CHART_VALUE_ROWS)
    return True


def preserve_weekly_chart_ranges() -> None:
    """Kept for call-site compatibility. Shift is applied in excel_service."""
    return None


def _chart_axis_label(value: Any) -> str:
    parsed = excel_value_as_date(value)
    if parsed is not None:
        return f"{parsed.month}/{parsed.day}"
    text = str(value or "").strip()
    if not text:
        return ""
    lowered = text.lower()
    if "assumption" in lowered or "금주" in text:
        return "This week"
    return text


def _chart_number(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text or text.upper() in {"TBN", "-"}:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def korea_premium_chart(get_value: GetValue) -> dict[str, Any]:
    """Korea Major 4 Refiners VLSFO Premium. Source C22:F26, series labels B23:B26."""
    labels = [_chart_axis_label(get_value(address)) for address in ("C22", "D22", "E22", "F22")]
    series = []
    for row in KOREA_CHART_VALUE_ROWS:
        series.append(
            {
                "name": str(get_value(f"B{row}") or "").strip() or f"Series {row}",
                "data": [_chart_number(get_value(f"{col}{row}")) for col in KOREA_CHART_VALUE_COLS],
            }
        )
    return {
        "title": "Korea Major 4 Refiners - VLSFO Premium Trends",
        "labels": labels,
        "series": series,
    }


def worldwide_vlsfo_chart(get_value: GetValue) -> dict[str, Any]:
    """Worldwide Ports VLSFO. Excel chart source P22:R26; S22:S26 is this-week assumption."""
    labels = [_chart_axis_label(get_value(address)) for address in ("P22", "Q22", "R22", "S22")]
    series = []
    for row in WORLDWIDE_CHART_VALUE_ROWS:
        series.append(
            {
                "name": str(get_value(f"O{row}") or "").strip() or f"Series {row}",
                "data": [_chart_number(get_value(f"{col}{row}")) for col in WORLDWIDE_CHART_VALUE_COLS],
            }
        )
    return {
        "title": "Worldwide Ports - VLSFO Bunker Price Trend",
        "labels": labels,
        "series": series,
    }


def spread_trend_chart(get_value: GetValue) -> dict[str, Any]:
    """Current-week A−B spreads from AF:AI formula cells. Not the helper table itself."""
    labels = ["SG vs Paper", "ZS vs Paper", "KR vs Paper", "KR − SG", "KR − ZS"]
    rows = (24, 25, 26, 27, 28)
    return {
        "title": "SPREAD TREND (BUNKER WIRE - MOPS SINGAPORE 0.5%)",
        "labels": labels,
        "series": [
            {"name": "FO", "data": [_chart_number(get_value(f"AG{row}")) for row in rows]},
            {"name": "GO", "data": [_chart_number(get_value(f"AH{row}")) for row in rows]},
            {"name": "HSFO", "data": [_chart_number(get_value(f"AI{row}")) for row in rows]},
        ],
    }
