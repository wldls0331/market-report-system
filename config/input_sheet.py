"""INPUT sheet layout. Target cells come only from config.cell_mapping."""

from __future__ import annotations

from dataclasses import dataclass

from config.cell_mapping import COMMENT_FIELDS, NUMBER_FIELDS, STRATEGY_FIELDS

INPUT_SHEET_NAME = "INPUT"
INPUT_KEY_COLUMN = 4  # D, hidden machine keys
INPUT_HEADER = "MARKET REPORT INPUT"
INPUT_NOTE = (
    "Enter values in column B only. Leave a cell blank to show TBN. "
    "Do not edit dated report sheets. Create / Update Report copies this INPUT into YY.MM.DD 보고자료."
)

# User-facing labels. Keys and target cells stay on NUMBER_FIELDS / COMMENT_FIELDS / STRATEGY_FIELDS.
INPUT_NUMBER_LABELS: dict[str, str] = {
    "paper_fo": "SING 0.5",
    "paper_hsfo": "SING 380",
    "paper_go": "SING Gasoil 10ppm",
    "sg_vlsfo": "Singapore VLSFO",
    "sg_hsfo": "Singapore HSFO",
    "sg_lsmgo": "Singapore LSMGO",
    "zs_vlsfo": "Zhoushan VLSFO",
    "zs_hsfo": "Zhoushan HSFO",
    "zs_lsmgo": "Zhoushan LSMGO",
    "kr_vlsfo": "Korea VLSFO",
    "kr_hsfo": "Korea HSFO",
    "kr_lsmgo": "Korea LSMGO",
    "jp_vlsfo": "Japan VLSFO",
    "hdo_premium": "Hyundai Oilbank",
    "sk_premium": "SK",
    "soil_premium": "S-OIL",
    "gs_premium": "GS",
}

BUNKER_INPUT_KEYS: tuple[str, ...] = (
    "sg_vlsfo",
    "sg_hsfo",
    "sg_lsmgo",
    "zs_vlsfo",
    "zs_hsfo",
    "zs_lsmgo",
    "kr_vlsfo",
    "kr_hsfo",
    "kr_lsmgo",
    "jp_vlsfo",
)
PREMIUM_INPUT_KEYS: tuple[str, ...] = (
    "hdo_premium",
    "sk_premium",
    "soil_premium",
    "gs_premium",
)
COMMENT_INPUT_KEYS: tuple[str, ...] = (
    "comment_korea_worldwide",
    "comment_singapore",
    "comment_china",
    "comment_japan",
)
INPUT_COMMENT_LABELS: dict[str, str] = {
    "comment_korea_worldwide": "South Korea – Southern",
    "comment_singapore": "Singapore",
    "comment_china": "China",
    "comment_japan": "Japan",
}
OBSOLETE_INPUT_KEYS: frozenset[str] = frozenset({"comment_korea"})
PREMIUM_INPUT_HINT = "VLSFO/HSFO e.g. 80/120. Chart uses the first number only."

AUTO_REPORT_DATE = "report_date"
AUTO_PRICING_MONTH = "pricing_month"
AUTO_THIS_WEEK = "this_week_friday"
AUTO_PREV_WEEK = "previous_week_friday"
AUTO_TWO_WEEKS = "two_weeks_ago_friday"

# Fixed INPUT addresses. All auto dates are computed from B4, never TODAY().
REPORT_DATE_CELL = "B4"
PRICING_MONTH_CELL = "B5"
THIS_WEEK_CELL = "B8"
PREV_WEEK_CELL = "B9"
TWO_WEEKS_CELL = "B10"
HOLIDAY_RANGE = "$F$2:$F$500"

PRICING_MONTH_FORMULA = (
    f'=IF({REPORT_DATE_CELL}="","",'
    f'INDEX({{"JAN";"FEB";"MAR";"APR";"MAY";"JUN";"JUL";"AUG";"SEP";"OCT";"NOV";"DEC"}},'
    f'MONTH(DATE(YEAR({REPORT_DATE_CELL}),MONTH({REPORT_DATE_CELL})+'
    f'IF(DAY({REPORT_DATE_CELL})>=27,1,0),1)))&" "&'
    f'YEAR(DATE(YEAR({REPORT_DATE_CELL}),MONTH({REPORT_DATE_CELL})+'
    f'IF(DAY({REPORT_DATE_CELL})>=27,1,0),1)))'
)
THIS_WEEK_FORMULA = (
    f'=IF({REPORT_DATE_CELL}="","",'
    f'WORKDAY.INTL({REPORT_DATE_CELL}-WEEKDAY({REPORT_DATE_CELL},2)+6,-1,"0000011",{HOLIDAY_RANGE}))'
)
PREV_WEEK_FORMULA = (
    f'=IF({REPORT_DATE_CELL}="","",'
    f'WORKDAY.INTL({REPORT_DATE_CELL}-WEEKDAY({REPORT_DATE_CELL},2)-1,-1,"0000011",{HOLIDAY_RANGE}))'
)
TWO_WEEKS_FORMULA = (
    f'=IF({REPORT_DATE_CELL}="","",'
    f'WORKDAY.INTL({REPORT_DATE_CELL}-WEEKDAY({REPORT_DATE_CELL},2)-8,-1,"0000011",{HOLIDAY_RANGE}))'
)
AUTO_FORMULAS: dict[str, str] = {
    AUTO_PRICING_MONTH: PRICING_MONTH_FORMULA,
    AUTO_THIS_WEEK: THIS_WEEK_FORMULA,
    AUTO_PREV_WEEK: PREV_WEEK_FORMULA,
    AUTO_TWO_WEEKS: TWO_WEEKS_FORMULA,
}


@dataclass(frozen=True)
class InputRow:
    kind: str  # title, note, spacer, section, field, auto
    label: str = ""
    key: str = ""
    hint: str = ""
    editable: bool = False


def _number_by_key() -> dict[str, object]:
    return {item.key: item for item in NUMBER_FIELDS}


def _comment_by_key() -> dict[str, object]:
    return {item.key: item for item in COMMENT_FIELDS}


def _hint_for_number(key: str) -> str:
    item = _number_by_key()[key]
    if key in PREMIUM_INPUT_KEYS:
        return f"{PREMIUM_INPUT_HINT} This Week cell {item.cell}"
    return f"Report cell {item.cell}"


def _hint_for_comment(key: str) -> str:
    item = _comment_by_key()[key]
    cells = ", ".join(item.cells)
    return f"Report cells {cells}."


def input_layout() -> tuple[InputRow, ...]:
    numbers = _number_by_key()
    paper_keys = [item.key for item in NUMBER_FIELDS if item.section == "paper"]
    rows: list[InputRow] = [
        InputRow(kind="title", label=INPUT_HEADER, hint=INPUT_NOTE),
        InputRow(kind="spacer"),
        InputRow(kind="section", label="REPORT INFORMATION"),
        InputRow(
            kind="field",
            label="Report Date",
            key=AUTO_REPORT_DATE,
            hint="YYYY-MM-DD. Sheet name becomes YY.MM.DD 보고자료",
            editable=True,
        ),
        InputRow(
            kind="auto",
            label="Pricing Month",
            key=AUTO_PRICING_MONTH,
            hint="Auto from B4. Day 1-26 same month, day 27-end next month",
        ),
        InputRow(kind="spacer"),
        InputRow(kind="section", label="WEEKLY DATES"),
        InputRow(
            kind="auto",
            label="This Week Friday",
            key=AUTO_THIS_WEEK,
            hint="Friday of the calendar week that contains B4. Holiday → previous SG working day",
        ),
        InputRow(
            kind="auto",
            label="Previous Week Friday",
            key=AUTO_PREV_WEEK,
            hint="This Week Friday minus 7 days, holiday-adjusted",
        ),
        InputRow(
            kind="auto",
            label="Two Weeks Ago Friday",
            key=AUTO_TWO_WEEKS,
            hint="This Week Friday minus 14 days, holiday-adjusted",
        ),
        InputRow(kind="spacer"),
        InputRow(kind="section", label="PAPER / MOPS"),
    ]
    for key in paper_keys:
        rows.append(
            InputRow(
                kind="field",
                label=INPUT_NUMBER_LABELS[key],
                key=key,
                hint=_hint_for_number(key),
                editable=True,
            )
        )
    rows.append(InputRow(kind="spacer"))
    rows.append(InputRow(kind="section", label="BUNKER MARKET PRICE"))
    for key in BUNKER_INPUT_KEYS:
        if key not in numbers:
            continue
        rows.append(
            InputRow(
                kind="field",
                label=INPUT_NUMBER_LABELS[key],
                key=key,
                hint=_hint_for_number(key),
                editable=True,
            )
        )
    rows.append(InputRow(kind="spacer"))
    rows.append(InputRow(kind="section", label="KOREA REFINERY PREMIUM"))
    for key in PREMIUM_INPUT_KEYS:
        rows.append(
            InputRow(
                kind="field",
                label=INPUT_NUMBER_LABELS[key],
                key=key,
                hint=_hint_for_number(key),
                editable=True,
            )
        )
    rows.append(InputRow(kind="spacer"))
    rows.append(InputRow(kind="section", label="MARKET COMMENT"))
    comments = _comment_by_key()
    for key in COMMENT_INPUT_KEYS:
        item = comments[key]
        rows.append(
            InputRow(
                kind="field",
                label=INPUT_COMMENT_LABELS.get(key, item.label),
                key=key,
                hint=_hint_for_comment(key),
                editable=True,
            )
        )
    rows.append(InputRow(kind="spacer"))
    rows.append(InputRow(kind="section", label="STRATEGY"))
    for item in STRATEGY_FIELDS:
        rows.append(
            InputRow(
                kind="field",
                label=item.label,
                key=item.key,
                hint=f"Report cell {item.cell}",
                editable=True,
            )
        )
    return tuple(rows)


def input_field_keys() -> tuple[str, ...]:
    return tuple(row.key for row in input_layout() if row.kind in {"field", "auto"} and row.key)
