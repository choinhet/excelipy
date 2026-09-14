import io
import math
import random
import string

import pandas as pd
import pytest
import xlsxwriter

import excelipy as ep
from excelipy.writers.table import (
    DEFAULT_FONT_SIZE,
    DEFAULT_LINE_SPACING,
    FIT_TOLERANCE_PX,
    PADDING_DEFAULT,
    _break_chunks,
    _excel_to_px,
    count_lines,
    get_row_height,
    get_text_px,
    get_text_size,
    write_table,
)


def lines(heights: dict[int, float], row: int, font_size: int | None = None) -> int:
    """The line count a written row height stands for."""
    height = heights.get(row)
    if height is None:
        return 1
    return round(height / ((font_size or DEFAULT_FONT_SIZE) * DEFAULT_LINE_SPACING))


def write(table: ep.Table, style: ep.Style | None = None) -> tuple[dict, dict]:
    """Write a table to a throwaway workbook, returning its column widths and row heights."""
    workbook = xlsxwriter.Workbook(io.BytesIO())
    worksheet = workbook.add_worksheet()
    heights: dict[int, float] = {}
    original = worksheet.set_row

    def set_row(row, height, *args, **kwargs):
        heights[row] = height
        return original(row, height, *args, **kwargs)

    worksheet.set_row = set_row
    write_table(workbook, worksheet, table, style or ep.Style())
    widths = dict(getattr(worksheet, "_excelipy_col_sizes"))
    workbook.close()
    return widths, heights


@pytest.fixture
def percent_df() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "ratio": [0.0714285714285714, 0.021818181818, 0.0559, 0.000344],
        }
    )


def test_row_height_follows_the_formatted_value(percent_df: pd.DataFrame):
    """A ratio shown as ``7.14%`` must not be sized as ``0.0714285714285714``."""
    widths, heights = write(
        ep.Table(
            data=percent_df,
            wrap_header=True,
            max_col_size=18,
            column_style={"ratio": ep.Style(numeric_format=".2%")},
        )
    )
    # Sized for "7.14%", not for the 0.0714285714285714 behind it
    assert widths[0] == get_text_size("7.14%")
    assert all(lines(heights, row) == 1 for row in range(1, 5))


def test_long_text_still_wraps():
    text = "a sentence long enough that no sensible column can hold it in one line"
    _, heights = write(
        ep.Table(
            data=pd.DataFrame({"col": [text]}),
            wrap_header=True,
            max_col_size=20,
        )
    )
    lines = count_lines(text, _excel_to_px(20))
    assert lines > 1
    assert heights[1] == pytest.approx(get_row_height(lines, None))


def test_auto_sized_column_never_wraps_its_own_text():
    """The invariant behind auto sizing: a column fits the text it was measured from."""
    random.seed(7)
    alphabet = string.ascii_letters + string.digits + " .-()%"
    for _ in range(200):
        text = "".join(random.choice(alphabet) for _ in range(random.randint(1, 80)))
        assert count_lines(text, _excel_to_px(get_text_size(text))) == 1


def test_column_narrower_than_the_padded_measure_stays_on_one_line():
    """
    Auto width adds breathing room on top of what the text needs, so a column
    trimmed back by that padding still holds the text on one line.
    """
    text = "some short text"
    width = get_text_size(text) - PADDING_DEFAULT
    assert count_lines(text, _excel_to_px(width)) == 1
    _, heights = write(
        ep.Table(
            data=pd.DataFrame({"col": [text]}),
            wrap_header=True,
            column_width={"col": width},
        )
    )
    assert lines(heights, 1) == 1


def test_embedded_newlines_grow_the_row_even_when_the_text_fits():
    """A cell that fits across is still two lines tall if it carries a break."""
    _, heights = write(
        ep.Table(
            data=pd.DataFrame({"col": ["first\nsecond"]}),
            wrap_header=True,
            max_col_size=40,
        )
    )
    assert heights[1] == pytest.approx(get_row_height(2, None))


def test_multiline_text_is_sized_by_its_widest_line():
    widths, _ = write(
        ep.Table(
            data=pd.DataFrame({"col": ["a much wider first line\nshort"]}),
            wrap_header=True,
        )
    )
    assert widths[0] == get_text_size("a much wider first line")


def test_bigger_fonts_need_more_lines_in_the_same_column():
    """A column is a fixed width, so a larger font fits less of it per line."""
    text = "a sentence long enough that no sensible column can hold it on one line"
    per_size = {}
    for size in (8, 11, 18):
        _, heights = write(
            ep.Table(
                data=pd.DataFrame({"col": [text]}),
                wrap_header=True,
                max_col_size=30,
                body_style=ep.Style(font_size=size),
            )
        )
        per_size[size] = lines(heights, 1, size)
    assert per_size[8] <= per_size[11] < per_size[18]


def test_hyphenated_token_breaks_after_its_dashes():
    """
    Excel ends a line on a dash rather than cutting a word anywhere, which
    leaves the tail of each line empty and costs a line the width alone allows.
    """
    token = "one-unbroken-token-far-too-long-for-any-of-these-columns-to-hold-it"
    room = _excel_to_px(20)

    packed, current = 1, ""
    for chunk in _break_chunks(token):
        if get_text_px(current + chunk) <= room + FIT_TOLERANCE_PX:
            current += chunk
        else:
            packed += 1
            current = chunk

    assert count_lines(token, room) == packed
    assert packed > math.ceil(get_text_px(token) / room)


def test_cells_that_cannot_wrap_do_not_grow_rows():
    long_text = "a very long piece of text that overflows its column by a lot"
    _, heights = write(
        ep.Table(
            data=pd.DataFrame({"col": [long_text]}),
            max_col_size=10,
            body_style=ep.Style(text_wrap=False),
        )
    )
    assert heights == {}  # nothing wraps, so no row is given a height at all


def test_row_heights_only_grow():
    """A second table sharing a row cannot shrink a height the first one needed."""
    workbook = xlsxwriter.Workbook(io.BytesIO())
    worksheet = workbook.add_worksheet()
    heights: dict[int, float] = {}
    original = worksheet.set_row
    worksheet.set_row = lambda row, height, *a, **k: (
        heights.__setitem__(row, height),
        original(row, height, *a, **k),
    )[1]

    tall = ep.Table(
        data=pd.DataFrame({"col": ["word " * 30]}),
        wrap_header=True,
        max_col_size=12,
    )
    short = ep.Table(
        data=pd.DataFrame({"other": ["tiny"]}),
        wrap_header=True,
        max_col_size=12,
    )
    write_table(workbook, worksheet, tall, ep.Style())
    grown = heights[1]
    write_table(workbook, worksheet, short, ep.Style(), origin=(1, 0))
    workbook.close()

    assert heights[1] == grown


def test_count_lines_wraps_on_words_and_newlines():
    width = _excel_to_px(get_text_size("hello world"))
    assert count_lines("hello world", width) == 1
    assert count_lines("hello world hello world", width) == 2
    assert count_lines("hello\nworld", width) == 2
    # A single word wider than the line is broken mid-word, as Excel does.
    # Whole characters have to land on one line or the next, so greedy packing
    # can cost one line over what the raw width would allow, but never more.
    narrow = _excel_to_px(get_text_size("x" * 20))
    ideal = math.ceil(get_text_px("x" * 200) / narrow)
    assert ideal <= count_lines("x" * 200, narrow) <= ideal + 1


def test_dates_and_missing_values_are_measured_as_shown():
    df = pd.DataFrame(
        {
            "when": pd.to_datetime(["2026-01-23", "2026-05-29"]),
            "what": ["ok", "also ok"],
        }
    )
    widths, heights = write(
        ep.Table(
            data=df,
            wrap_header=True,
            column_style={"when": ep.Style(numeric_format="%d - %B")},
        )
    )
    assert widths[0] <= get_text_size("23 - January")
    assert all(lines(heights, row) == 1 for row in (1, 2))


if __name__ == "__main__":
    pytest.main([__file__])
