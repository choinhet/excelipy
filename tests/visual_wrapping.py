"""
Generates a workbook for checking auto row heights by eye.

    uv run python -m tests.visual_wrapping [output.xlsx]

Every sheet is one case. Each row carries what excelipy decided it needs, so
open the file in Excel and check the two things a row height can get wrong:

    too tall   the text ends before the row does, leaving an empty band
    too short  the text is clipped, or spills out of its row

A row marked "1 line" must show one line of text in a row of default height.
"""

import io
import sys
from pathlib import Path
from typing import NamedTuple

import pandas as pd
import xlsxwriter

import excelipy as ep
from excelipy.writers.table import (
    COL_CACHE_NAME,
    DEFAULT_FONT_FAMILY,
    DEFAULT_FONT_SIZE,
    DEFAULT_LINE_SPACING,
    ROW_CACHE_NAME,
    _display_text,
    _excel_to_px,
    _load_font,
    _max_digit_px,
    get_text_px,
    write_table,
)

PLACEHOLDER = "? lines"


def measure(
    table: ep.Table, style: ep.Style
) -> tuple[dict[int, int], dict[int, float]]:
    """What excelipy would give this table, without writing a file."""
    workbook = xlsxwriter.Workbook(io.BytesIO())
    worksheet = workbook.add_worksheet()
    write_table(workbook, worksheet, table, style)
    widths = dict(getattr(worksheet, COL_CACHE_NAME, {}))
    heights = dict(getattr(worksheet, ROW_CACHE_NAME, {}))
    workbook.close()
    return widths, heights


def metrics(
    frame: pd.DataFrame,
    width: int,
    font_sizes: dict[int, int],
    font_family: str | None = None,
) -> str:
    """
    The numbers behind the decision, for the first column of the table.

    These are what a row height is argued from, and they depend on the fonts
    installed here, so they are written into the sheet: a picture of the sheet
    is then enough to see why a row came out the way it did.
    """
    capacity = _excel_to_px(width)
    shown = [
        f"{get_text_px(_display_text(value, ep.Style()), font_sizes.get(idx), font_family):.0f}"
        for idx, value in enumerate(frame.iloc[:, 0])
    ]
    named = f" in {font_family}" if font_family else ""
    return (
        f"[{width} units = {capacity:.0f}px of room; "
        f"text{named} measures {', '.join(shown)}px]"
    )


def expected(height: float | None, font_size: int | None) -> str:
    """The row height read back as a line count, which is what you can see."""
    if height is None:
        lines = 1
    else:
        size = font_size or DEFAULT_FONT_SIZE
        lines = max(round(height / (size * DEFAULT_LINE_SPACING)), 1)
    return "1 line" if lines == 1 else f"{lines} lines"


class Case(NamedTuple):
    sheet: ep.Sheet
    expectations: list[str]


def case(
    name: str,
    look_for: str,
    data: dict[str, list] | pd.DataFrame,
    style: ep.Style | None = None,
    font_sizes: dict[int, int] | None = None,
    font_family: str | None = None,
    **table_args,
) -> Case:
    """One case: a note on what to look for, then the table it describes."""
    sheet_style = style or ep.Style(valign="vcenter")
    table_style = ep.Style(font_family=font_family) if font_family else ep.Style()
    frame = data if isinstance(data, pd.DataFrame) else pd.DataFrame(data)
    rows = len(frame)
    font_sizes = font_sizes or {}
    row_style = {idx: ep.Style(font_size=size) for idx, size in font_sizes.items()}

    with_expect = frame.copy()
    with_expect["excelipy says"] = [PLACEHOLDER] * rows
    widths, heights = measure(
        ep.Table(
            data=with_expect,
            row_style=row_style,
            style=table_style,
            **table_args,
        ),
        sheet_style,
    )
    with_expect["excelipy says"] = [
        expected(heights.get(idx + 1), font_sizes.get(idx)) for idx in range(rows)
    ]
    sheet = ep.Sheet(
        name=name,
        components=[
            ep.Text(
                text=(
                    f"{look_for}  "
                    f"{metrics(frame, widths.get(0, 0), font_sizes, font_family)}"
                ),
                style=ep.Style(bold=True),
            ),
            ep.Table(
                data=with_expect,
                row_style=row_style,
                style=table_style.merge(ep.Style(padding_top=1)),
                **table_args,
            ),
        ],
        style=sheet_style,
    )
    body = list(with_expect["excelipy says"])
    header = expected(heights.get(0), None)
    return Case(sheet, [f"header {header}", *body])


SHORT = "a short value"
LONG = "a sentence long enough that no sensible column can hold it on one line"
LONGER = " ".join([LONG, "and then it keeps going for a good while after that"])
WORD = "one-unbroken-token-far-too-long-for-any-of-these-columns-to-hold-it"
FONT_SENSITIVE = "wrapping depends on the font"


def merged_header_frame() -> pd.DataFrame:
    """Two columns under one repeated header, which the writer merges."""
    header = "a repeated header long enough that it has to wrap"
    frame = pd.DataFrame({header: [1, 2], "second": [3, 4]})
    return frame.rename(columns={"second": header})


def build() -> list[Case]:
    return [
        case(
            "01 short text",
            "Every row is one line: the columns are sized to their own text.",
            {"text": [SHORT, "tiny", "a slightly longer value"]},
            wrap_header=True,
        ),
        case(
            "02 long text unclamped",
            "One line each: without a clamp the column grows to fit the text.",
            {"text": [LONG, LONGER, SHORT]},
            wrap_header=True,
        ),
        case(
            "03 clamped long text",
            "Wrapped over several lines, with no empty band under the last one.",
            {"text": [LONG, LONGER, SHORT]},
            wrap_header=True,
            max_col_size=30,
        ),
        case(
            "04 percentages",
            "One line each. This is the bug: 7.14% must not be sized as 0.0714285714.",
            {"ratio": [0.0714285714285714, 0.021818181818, 0.0559, 0.000344]},
            wrap_header=True,
            max_col_size=18,
            column_style={"ratio": ep.Style(numeric_format=".2%")},
        ),
        case(
            "05 numbers and dates",
            "One line each: formatted values are measured as they are drawn.",
            {
                "count": [1234567, 22, 9981],
                "when": pd.to_datetime(["2026-01-23", "2026-04-20", "2026-06-11"]),
            },
            wrap_header=True,
            max_col_size=20,
            column_style={
                "count": ep.Style(numeric_format=",d"),
                "when": ep.Style(numeric_format="%m/%d/%Y"),
            },
        ),
        case(
            "06 fixed width that fits",
            "One line: the text measures wider than 14 units but Excel still fits it.",
            {"text": ["some short text", "also fits here", "tiny"]},
            wrap_header=True,
            column_width={"text": 14},
        ),
        case(
            "07 fixed width too narrow",
            "Wrapped: eight units cannot hold any of these on one line.",
            {"text": [SHORT, LONG, "two words"]},
            wrap_header=True,
            column_width={"text": 8},
        ),
        case(
            "08 explicit newlines",
            "Two and three lines: the newlines in the text are the breaks.",
            {"text": ["first\nsecond", "one\ntwo\nthree", SHORT]},
            wrap_header=True,
            max_col_size=40,
        ),
        case(
            "09 unbreakable word",
            "Broken mid-word, because the token is wider than the column.",
            {"text": [WORD, SHORT]},
            wrap_header=True,
            max_col_size=20,
        ),
        case(
            "10 mixed font sizes",
            "Same text, three sizes: each row is as tall as its own font needs.",
            {"text": [LONG, LONG, LONG]},
            wrap_header=True,
            max_col_size=30,
            font_sizes={0: 8, 1: 11, 2: 18},
        ),
        case(
            "11 long headers",
            "The header row wraps; the body rows stay on one line.",
            {
                "a header long enough that it has to wrap over the column": [1, 2],
                "another header of roughly the same generous length here": [3, 4],
            },
            wrap_header=True,
            max_col_size=20,
        ),
        case(
            "12 merged headers",
            "One wrapped header merged across two columns, body rows on one line.",
            merged_header_frame(),
            wrap_header=True,
            max_col_size=18,
        ),
        case(
            "13 min column size",
            "One line each: min_col_size leaves far more room than the text needs.",
            {"text": ["tiny", "small", SHORT]},
            wrap_header=True,
            min_col_size=40,
        ),
        case(
            "14 accents and symbols",
            "Wrapped like any other text, with the accents measured properly.",
            {"text": ["ação coração pão à ré ñ ü ß", LONG]},
            wrap_header=True,
            max_col_size=22,
        ),
        case(
            "15 wrapping turned off",
            "One line each, overflowing the column: nothing here wraps, so no row grows.",
            {"text": [LONG, LONGER]},
            max_col_size=20,
        ),
        case(
            "16 arial body",
            "Arial is wider than Calibri, so this wraps where Calibri would not.",
            {"text": [FONT_SENSITIVE, LONG, "tiny"]},
            font_family="Arial",
            wrap_header=True,
            column_width={"text": 25},
        ),
        case(
            "17 arial clamped",
            "Arial in a clamped column: a line more than the same text in Calibri.",
            {"text": [LONG, LONGER]},
            font_family="Arial",
            wrap_header=True,
            max_col_size=20,
        ),
        case(
            "18 calibri for comparison",
            "The same text and widths as sheet 17, in Calibri: a line less.",
            {"text": [LONG, LONGER]},
            wrap_header=True,
            max_col_size=20,
        ),
    ]


def main(out: Path) -> None:
    cases = build()
    ep.save(ep.Excel(path=out, sheets=[c.sheet for c in cases]))
    path = getattr(_load_font(DEFAULT_FONT_FAMILY, DEFAULT_FONT_SIZE), "path", None)
    using = path if isinstance(path, str) else f"a stand-in for {DEFAULT_FONT_FAMILY}"
    print(f"wrote {out}")
    print(f"measured with {using}, one column unit = {_max_digit_px():.1f}px")
    print("open it and check every row against what excelipy says it needs:\n")
    for cur in cases:
        print(f"  {cur.sheet.name:26} {', '.join(cur.expectations)}")
    print("\ntoo tall  -> the text stops short of the bottom of the row")
    print("too short -> the text is clipped or spills out of the row")


if __name__ == "__main__":
    main(Path(sys.argv[1] if len(sys.argv) > 1 else "wrapping_check.xlsx"))
