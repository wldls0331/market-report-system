"""Linux / Streamlit Cloud PDF renderer. Does not use Excel COM.

Mirrors the local Excel print layout as three A4 portrait pages:
1. Korea (title, paper/MOPS, bunker, premiums, Korea commentary, strategy)
2. Worldwide market commentary
3. Premium / spread table
"""

from __future__ import annotations

import datetime as dt
from pathlib import Path
from typing import Any

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfgen import canvas

from config.cell_mapping import EXPECTED_PDF_PAGES, NUMBER_FIELDS, STRATEGY_FIELDS
from services.working_day_service import format_data_reference_display, format_report_title

NAVY = (0.106, 0.212, 0.365)
GOLD = (0.769, 0.639, 0.353)
GRAY = (0.957, 0.965, 0.973)
LINE = (0.843, 0.863, 0.890)
TEXT = (0.122, 0.161, 0.216)
WHITE = (1, 1, 1)

KOREAN_FONT = "HYGothic-Medium"
_FONT_REGISTERED = False
PAGE_W, PAGE_H = A4


def _ensure_fonts() -> None:
    global _FONT_REGISTERED
    if _FONT_REGISTERED:
        return
    pdfmetrics.registerFont(UnicodeCIDFont(KOREAN_FONT))
    _FONT_REGISTERED = True


def _tbn(value: Any) -> str:
    text = "" if value is None else str(value).strip()
    return text if text else "TBN"


def _wrap(text: str, max_chars: int) -> list[str]:
    raw = _tbn(text).replace("\r", "")
    lines: list[str] = []
    for paragraph in raw.split("\n"):
        chunk = paragraph.strip()
        if not chunk:
            continue
        while len(chunk) > max_chars:
            lines.append(chunk[:max_chars])
            chunk = chunk[max_chars:]
        lines.append(chunk)
    return lines or ["TBN"]


class _Pdf:
    def __init__(self, path: Path) -> None:
        _ensure_fonts()
        self.c = canvas.Canvas(str(path), pagesize=A4)
        self.page = 0

    def new_page(self) -> None:
        if self.page:
            self.c.showPage()
        self.page += 1
        self._footer()

    def _footer(self) -> None:
        c = self.c
        c.setFillColorRGB(*NAVY)
        c.rect(0, 0, PAGE_W, 12 * mm, fill=1, stroke=0)
        c.setFillColorRGB(*WHITE)
        c.setFont(KOREAN_FONT, 8)
        c.drawString(16 * mm, 4.5 * mm, "Weekly Bunkering Market Report")
        c.drawRightString(PAGE_W - 16 * mm, 4.5 * mm, f"{self.page} / {EXPECTED_PDF_PAGES}")

    def banner(self, title: str) -> float:
        c = self.c
        top = PAGE_H - 10 * mm
        c.setFillColorRGB(*NAVY)
        c.rect(12 * mm, top - 16 * mm, PAGE_W - 24 * mm, 16 * mm, fill=1, stroke=0)
        c.setFillColorRGB(*GOLD)
        c.rect(12 * mm, top - 16 * mm, PAGE_W - 24 * mm, 1.6 * mm, fill=1, stroke=0)
        c.setFillColorRGB(*WHITE)
        c.setFont(KOREAN_FONT, 13)
        c.drawString(16 * mm, top - 11 * mm, title)
        return top - 22 * mm

    def heading(self, y: float, text: str) -> float:
        c = self.c
        c.setFillColorRGB(*NAVY)
        c.setFont(KOREAN_FONT, 11)
        c.drawString(16 * mm, y, text)
        c.setStrokeColorRGB(*GOLD)
        c.setLineWidth(1)
        c.line(16 * mm, y - 2 * mm, PAGE_W - 16 * mm, y - 2 * mm)
        return y - 8 * mm

    def table(self, y: float, headers: list[str], rows: list[list[str]], col_w: list[float]) -> float:
        c = self.c
        x = 16 * mm
        row_h = 7 * mm
        header_h = 7.5 * mm
        c.setFillColorRGB(*NAVY)
        c.rect(x, y - header_h, sum(col_w), header_h, fill=1, stroke=0)
        c.setFillColorRGB(*WHITE)
        c.setFont(KOREAN_FONT, 8)
        cx = x
        for header, width in zip(headers, col_w):
            c.drawString(cx + 2 * mm, y - 5.2 * mm, header)
            cx += width
        y -= header_h
        for index, row in enumerate(rows):
            if index % 2 == 1:
                c.setFillColorRGB(*GRAY)
                c.rect(x, y - row_h, sum(col_w), row_h, fill=1, stroke=0)
            c.setStrokeColorRGB(*LINE)
            c.setLineWidth(0.4)
            c.rect(x, y - row_h, sum(col_w), row_h, fill=0, stroke=1)
            c.setFillColorRGB(*TEXT)
            c.setFont(KOREAN_FONT, 8)
            cx = x
            for cell, width in zip(row, col_w):
                c.drawString(cx + 2 * mm, y - 4.8 * mm, _tbn(cell)[:42])
                cx += width
            y -= row_h
        return y - 4 * mm

    def box(self, y: float, text: str, max_lines: int = 6) -> float:
        c = self.c
        lines = _wrap(text, 92)[:max_lines]
        height = max(len(lines), 1) * 4.4 * mm + 4 * mm
        x = 16 * mm
        width = PAGE_W - 32 * mm
        c.setStrokeColorRGB(*LINE)
        c.setFillColorRGB(*WHITE)
        c.setLineWidth(0.6)
        c.roundRect(x, y - height, width, height, 2, fill=1, stroke=1)
        c.setFillColorRGB(*TEXT)
        c.setFont(KOREAN_FONT, 8)
        ty = y - 5.5 * mm
        for line in lines:
            c.drawString(x + 3 * mm, ty, line)
            ty -= 4.4 * mm
        return y - height - 3 * mm


def render_market_report_pdf(
    *,
    pdf_path: Path,
    report_date: dt.date,
    data_reference_date: dt.date,
    pricing_month: str,
    sheet_name: str,
    inputs: dict[str, Any],
    extra_cells: dict[str, Any] | None = None,
) -> Path:
    extra_cells = extra_cells or {}
    pdf_path = Path(pdf_path)
    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    if pdf_path.exists():
        pdf_path.unlink()

    paper = {item.key: inputs.get(item.key) for item in NUMBER_FIELDS if item.section == "paper"}
    bunker = [item for item in NUMBER_FIELDS if item.section == "bunker"]
    premium = [item for item in NUMBER_FIELDS if item.section == "premium"]
    groups: dict[str, dict[str, str]] = {}
    for item in bunker:
        groups.setdefault(item.group, {})[item.label] = _tbn(inputs.get(item.key))
    bunker_rows = [
        [
            group,
            values.get("VLSFO (FO)", "TBN"),
            values.get("LSMGO/MGO (GO)", "TBN"),
            values.get("HSFO", "TBN"),
        ]
        for group, values in groups.items()
    ]

    pdf = _Pdf(pdf_path)

    pdf.new_page()
    y = pdf.banner(format_report_title(report_date))
    y = pdf.heading(y, "Report Information")
    y = pdf.table(
        y,
        ["Field", "Value"],
        [
            ["Report Date", report_date.isoformat()],
            ["Data Reference Date", format_data_reference_display(data_reference_date)],
            ["Pricing Month", pricing_month],
            ["Google Sheet", sheet_name or "TBN"],
        ],
        [50 * mm, 128 * mm],
    )
    y = pdf.heading(y, "Paper / MOPS")
    y = pdf.table(
        y,
        ["SING 0.5", "GASOIL 10PPM", "SING 380"],
        [[_tbn(paper.get("paper_fo")), _tbn(paper.get("paper_go")), _tbn(paper.get("paper_hsfo"))]],
        [59.3 * mm, 59.3 * mm, 59.4 * mm],
    )
    y = pdf.heading(y, "Bunker Market Price")
    y = pdf.table(
        y,
        ["Port", "VLSFO (FO)", "LSMGO/MGO (GO)", "HSFO"],
        bunker_rows,
        [44.5 * mm, 44.5 * mm, 44.5 * mm, 44.5 * mm],
    )
    y = pdf.heading(y, "Korea Refinery Premium")
    y = pdf.table(
        y,
        [item.label for item in premium],
        [[_tbn(inputs.get(item.key)) for item in premium]],
        [44.5 * mm] * len(premium),
    )
    y = pdf.heading(y, "Korea Market")
    y = pdf.box(y, inputs.get("comment_korea"), max_lines=7)
    y = pdf.heading(y, "Strategy")
    for item in STRATEGY_FIELDS:
        y = pdf.box(y, f"{item.label}: {_tbn(inputs.get(item.key))}", max_lines=3)

    pdf.new_page()
    y = pdf.banner("Worldwide Market")
    y = pdf.heading(y, "South Korea")
    y = pdf.box(y, inputs.get("comment_korea_worldwide") or inputs.get("comment_korea"), max_lines=8)
    y = pdf.heading(y, "Singapore")
    y = pdf.box(y, inputs.get("comment_singapore"), max_lines=8)
    y = pdf.heading(y, "China / Zhoushan")
    y = pdf.box(y, inputs.get("comment_china"), max_lines=8)
    y = pdf.heading(y, "Japan")
    y = pdf.box(y, inputs.get("comment_japan"), max_lines=8)

    pdf.new_page()
    y = pdf.banner("Premium / Spread")
    y = pdf.heading(y, "A − B Premium and Korea Spreads")
    y = pdf.table(
        y,
        ["", "FO", "GO", "HSFO"],
        [
            ["SG vs Paper", extra_cells.get("AG24"), extra_cells.get("AH24"), extra_cells.get("AI24")],
            ["ZS vs Paper", extra_cells.get("AG25"), extra_cells.get("AH25"), extra_cells.get("AI25")],
            ["KR vs Paper", extra_cells.get("AG26"), extra_cells.get("AH26"), extra_cells.get("AI26")],
            ["KR − SG", extra_cells.get("AG27"), extra_cells.get("AH27"), extra_cells.get("AI27")],
            ["KR − ZS", extra_cells.get("AG28"), extra_cells.get("AH28"), extra_cells.get("AI28")],
        ],
        [44.5 * mm, 44.5 * mm, 44.5 * mm, 44.5 * mm],
    )
    y = pdf.heading(y, "Bunker recap")
    y = pdf.table(
        y,
        ["Port", "VLSFO (FO)", "LSMGO/MGO (GO)", "HSFO"],
        bunker_rows,
        [44.5 * mm, 44.5 * mm, 44.5 * mm, 44.5 * mm],
    )

    pdf.c.save()
    if pdf.page != EXPECTED_PDF_PAGES:
        raise RuntimeError(f"Cloud PDF has {pdf.page} page(s); expected {EXPECTED_PDF_PAGES}.")
    return pdf_path
