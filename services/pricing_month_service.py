"""Pricing month (Paper / MOPS WMA) derived from Data Reference Date."""

from __future__ import annotations

import calendar
import datetime as dt
from dataclasses import dataclass

MONTH_ABBR = (
    "",
    "JAN",
    "FEB",
    "MAR",
    "APR",
    "MAY",
    "JUN",
    "JUL",
    "AUG",
    "SEP",
    "OCT",
    "NOV",
    "DEC",
)

ROLLOVER_DAY = 25


@dataclass(frozen=True)
class YearMonth:
    year: int
    month: int

    @property
    def abbr(self) -> str:
        return MONTH_ABBR[self.month]

    @property
    def full_name(self) -> str:
        return calendar.month_name[self.month]

    def label(self) -> str:
        return f"{self.abbr} {self.year}"

    def iso_key(self) -> str:
        return f"{self.year:04d}-{self.month:02d}"

    def next_month(self) -> YearMonth:
        if self.month == 12:
            return YearMonth(self.year + 1, 1)
        return YearMonth(self.year, self.month + 1)

    @classmethod
    def from_date(cls, date: dt.date) -> YearMonth:
        return cls(date.year, date.month)

    @classmethod
    def parse_label(cls, text: str) -> YearMonth:
        abbr, year_text = text.strip().split()
        month = MONTH_ABBR.index(abbr.upper())
        return cls(int(year_text), month)

    @classmethod
    def parse(cls, text: str) -> YearMonth:
        raw = (text or "").strip()
        if not raw:
            raise ValueError("Empty pricing month")
        if raw[0].isdigit() and "-" in raw:
            year_text, month_text = raw.split("-", 1)
            return cls(int(year_text), int(month_text))
        return cls.parse_label(raw)


def default_pricing_month(as_of_date: dt.date) -> YearMonth:
    current = YearMonth.from_date(as_of_date)
    if as_of_date.day >= ROLLOVER_DAY:
        return current.next_month()
    return current


def pricing_month_choices(as_of_date: dt.date) -> tuple[YearMonth, ...]:
    current = YearMonth.from_date(as_of_date)
    nxt = current.next_month()
    return (current, nxt)


def paper_ui_label(month: YearMonth, product: str) -> str:
    return f"{month.abbr} {product}"
