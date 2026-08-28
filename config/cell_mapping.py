"""Excel cell mapping derived from templates/Weekly Bunkering Report.xlsx.xlsx.

Analyzed sheets: `26.08.11 보고자료` (latest) vs `26.08.04 보고자료`.

Sheet naming: `YY.MM.DD 보고자료` (newest sheet first).
Formulas live only in AF:AI premium/spread scratch cells and must not be overwritten.
Weekly chart ranges C22:F26 and P22:R26 are preserved on daily copy.

Google Sheets reads the same addresses via snapshot_cells() / chart_window_cells().
There is no separate Market Data table mapping.
"""

from __future__ import annotations

from dataclasses import dataclass

SHEET_NAME_PATTERN = r"^(\d{2})\.(\d{2})\.(\d{2})\s+보고자료"
SHEET_NAME_SUFFIX = "보고자료"

# Workbook cells have no "WEEKLY BUNKERING REPORT" title (headers empty, B1 empty).
# Reference PDF `Weekly Report_Bunkering_26.06.09.pdf` prints the title at the top of page 1.
# B1 is the top-left visible cell of the printed Korea page.
TITLE_CELL = "B1"
TITLE_TEMPLATE = "{date}_WEEKLY BUNKERING REPORT"

# PDF is exported as three separate 1-page ranges then merged.
# Excel coalesces "$B$1:$L$47,$B$48:$L$76" into "$B$1:$L$76", so a single
# PrintArea string cannot keep Korea and Spread on different pages.
# Rows 22-27 are chart source tables plus a leftover bordered spacer row.
# Column N is a wide spacer; the old B:X FitWide=2 range put it on page 2.
# AF:AI orange scratch pad is outside these ranges.
PDF_PAGE_RANGES: tuple[str, ...] = (
    "$B$1:$L$47",
    "$O$1:$X$47",
    "$B$48:$L$76",
)
PDF_PRINT_AREA = ",".join(PDF_PAGE_RANGES)
PDF_ORIENTATION_PORTRAIT = 1
PDF_ORDER_OVER_THEN_DOWN = 2
PDF_FIT_TO_PAGES_WIDE = 1
PDF_FIT_TO_PAGES_TALL = 1
PDF_HELPER_ROWS = (22, 27)
PDF_HELPER_COLUMNS = ("N",)
EXPECTED_PDF_PAGES = 3


@dataclass(frozen=True)
class NumberField:
    key: str
    label: str
    cell: str
    section: str
    group: str
    previous: bool = True
    blank_policy: str = "blank"  # numeric source: never write "TBN"
    label_cell: str | None = None
    label_template: str | None = None


@dataclass(frozen=True)
class CommentField:
    key: str
    label: str
    cells: tuple[str, ...]
    section: str
    previous: bool = True
    blank_policy: str = "tbn"


@dataclass(frozen=True)
class StrategyField:
    key: str
    label: str
    cell: str
    prefix: str
    previous: bool = True
    blank_policy: str = "tbn"


# Excel scratch pad on latest report sheets (26.08.11 보고자료):
#   B = Paper / MOPS source   AG12 FO, AH12 GO, AI12 HSFO   header AH10='페이퍼'
#   A = Bunker market source  AG17:AI20 by port             header AH15='시장가'
#   Premium/Spread formulas (do not overwrite):
#     AG24=AG17-$AG$12  AH24=AH17/7.45-$AH$12  AI24=AI17-$AI$12  (A - B)
NUMBER_FIELDS: tuple[NumberField, ...] = (
    NumberField(
        key="paper_fo",
        label="SING 0.5",
        cell="AG12",
        section="paper",
        group="Paper / MOPS",
        label_cell="AF3",
        label_template="{month} SING 0.5 {value}",
    ),
    NumberField(
        key="paper_hsfo",
        label="SING 380",
        cell="AI12",
        section="paper",
        group="Paper / MOPS",
        label_cell="AF4",
        label_template="{month} SING 380 {value}",
    ),
    NumberField(
        key="paper_go",
        label="SING GASOIL 10PPM",
        cell="AH12",
        section="paper",
        group="Paper / MOPS",
        label_cell="AF5",
        label_template="{month} SING GASOIL 10PPM {value}",
    ),
    NumberField(key="sg_vlsfo", label="VLSFO (FO)", cell="AG17", section="bunker", group="Singapore"),
    NumberField(key="sg_lsmgo", label="LSMGO/MGO (GO)", cell="AH17", section="bunker", group="Singapore"),
    NumberField(key="sg_hsfo", label="HSFO", cell="AI17", section="bunker", group="Singapore"),
    NumberField(key="zs_vlsfo", label="VLSFO (FO)", cell="AG18", section="bunker", group="Zhoushan"),
    NumberField(key="zs_lsmgo", label="LSMGO/MGO (GO)", cell="AH18", section="bunker", group="Zhoushan"),
    NumberField(key="zs_hsfo", label="HSFO", cell="AI18", section="bunker", group="Zhoushan"),
    NumberField(key="kr_vlsfo", label="VLSFO (FO)", cell="AG19", section="bunker", group="Korea"),
    NumberField(key="kr_lsmgo", label="LSMGO/MGO (GO)", cell="AH19", section="bunker", group="Korea"),
    NumberField(key="kr_hsfo", label="HSFO", cell="AI19", section="bunker", group="Korea"),
    NumberField(key="jp_vlsfo", label="VLSFO (FO)", cell="AG20", section="bunker", group="Japan"),
    NumberField(key="hdo_premium", label="HyunDai", cell="F23", section="premium", group="Korea Refinery"),
    NumberField(key="sk_premium", label="SK", cell="F24", section="premium", group="Korea Refinery"),
    NumberField(key="soil_premium", label="SOIL", cell="F25", section="premium", group="Korea Refinery"),
    NumberField(key="gs_premium", label="GS", cell="F26", section="premium", group="Korea Refinery"),
)

COMMENT_FIELDS: tuple[CommentField, ...] = (
    CommentField(
        key="comment_korea",
        label="Korea",
        cells=("B29", "B30", "B31", "B32"),
        section="comment",
    ),
    CommentField(
        key="comment_korea_worldwide",
        label="Korea (Worldwide)",
        cells=("O31", "O32", "O33"),
        section="comment_internal",
    ),
    CommentField(
        key="comment_singapore",
        label="Singapore",
        cells=("O36", "O37"),
        section="comment",
    ),
    CommentField(
        key="comment_china",
        label="China / Zhoushan",
        cells=("O39", "O40"),
        section="comment",
    ),
    CommentField(
        key="comment_japan",
        label="Japan",
        cells=("O43", "O44"),
        section="comment",
    ),
)

STRATEGY_FIELDS: tuple[StrategyField, ...] = (
    StrategyField(key="strategy_vlsfo", label="VLSFO", cell="B44", prefix="-VLSFO: "),
    StrategyField(key="strategy_lsmgo", label="LSMGO", cell="B45", prefix="-LSMGO: "),
    StrategyField(key="strategy_hsfo", label="HSFO", cell="B46", prefix="-HSFO: "),
)

# Report Date stamps the title. Helper table dates use Data Reference Date.
DATE_CELLS: dict[str, str] = {
    "paper_date": "AF2",
    "paper_table_date": "AF10",
    "market_table_date": "AF15",
    "premium_table_date": "AF22",
}

KOREA_CHART_DATE_CELLS = ("C22", "D22", "E22")
KOREA_CHART_ASSUMPTION_DATE_CELL = "F22"
KOREA_CHART_VALUE_ROWS = (23, 24, 25, 26)
KOREA_CHART_VALUE_COLS = ("C", "D", "E", "F")

WORLDWIDE_CHART_DATE_CELLS = ("P22", "Q22", "R22")
WORLDWIDE_CHART_ASSUMPTION_DATE_CELL = "S22"
WORLDWIDE_CHART_VALUE_ROWS = (23, 24, 25, 26)
WORLDWIDE_CHART_VALUE_COLS = ("P", "Q", "R", "S")
# New historical point (column R) VLSFO as of Data Reference Date.
WORLDWIDE_NEW_POINT_CELLS: dict[str, str] = {
    "R23": "kr_vlsfo",
    "R24": "sg_vlsfo",
    "R25": "jp_vlsfo",
    "R26": "zs_vlsfo",
}

# Copied with the sheet. Never overwrite these cells.
FORMULA_CELLS: tuple[str, ...] = (
    "AG24",
    "AH24",
    "AI24",
    "AG25",
    "AH25",
    "AI25",
    "AG26",
    "AH26",
    "AI26",
    "AG27",
    "AH27",
    "AI27",
    "AG28",
    "AH28",
    "AI28",
)

# Structural labels used to verify the copied sheet still matches mapping.
REQUIRED_LABELS: dict[str, str] = {
    "B23": "HyunDai",
    "B24": "SK",
    "B25": "SOIL",
    "B26": "GS",
    "AF17": "SG",
    "AF18": "ZS",
    "AF19": "KR",
    "AF20": "JP",
    "AH10": "페이퍼",
    "AH15": "시장가",
}

# Korea Worldwide section is filled from the same Korea comment input.
KOREA_WORLDWIDE_COMMENT_KEY = "comment_korea_worldwide"
KOREA_COMMENT_SOURCE_KEY = "comment_korea"

SALES_ASSUMPTION_HEADER_CELL = "B35"
SALES_ASSUMPTION_HEADER_TEMPLATE = "※ Korea Refineries-VLSFO sales assumption for {month}"


def chart_window_cells() -> tuple[str, ...]:
    cells: list[str] = []
    for row in range(22, 27):
        for col in KOREA_CHART_VALUE_COLS:
            cells.append(f"{col}{row}")
    for row in range(22, 27):
        for col in WORLDWIDE_CHART_VALUE_COLS:
            cells.append(f"{col}{row}")
    cells.extend(("O23", "O24", "O25", "O26", "F22", "S22"))
    return tuple(dict.fromkeys(cells))


def snapshot_cells() -> tuple[str, ...]:
    cells: list[str] = []
    for field in NUMBER_FIELDS:
        cells.append(field.cell)
        if field.label_cell:
            cells.append(field.label_cell)
    for comment in COMMENT_FIELDS:
        cells.extend(comment.cells)
    cells.extend(item.cell for item in STRATEGY_FIELDS)
    cells.extend(DATE_CELLS.values())
    cells.append(TITLE_CELL)
    cells.append(SALES_ASSUMPTION_HEADER_CELL)
    cells.extend(chart_window_cells())
    return tuple(dict.fromkeys(cells))


# Google Sheets reads these Excel addresses from `YY.MM.DD 보고자료`.
# Formula cells in AF:AI are excluded.
REPORT_SHEET_CELLS: tuple[str, ...] = snapshot_cells()


def all_mapped_cells() -> list[str]:
    cells: list[str] = []
    cells.extend(field.cell for field in NUMBER_FIELDS)
    cells.extend(DATE_CELLS.values())
    cells.extend(FORMULA_CELLS)
    for comment in COMMENT_FIELDS:
        cells.extend(comment.cells)
    cells.extend(item.cell for item in STRATEGY_FIELDS)
    cells.extend(REQUIRED_LABELS.keys())
    return cells
