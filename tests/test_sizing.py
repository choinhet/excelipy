import random
import string

import pandas as pd
import pytest
import xlsxwriter

import excelipy as ep
from excelipy.writers.table import (
    _excel_to_px,
    count_lines,
    get_row_height,
    get_text_size,
    write_table,
)


def write(table: ep.Table, style: ep.Style | None = None) -> tuple[dict, dict]:
    """Write a table to a throwaway workbook, returning its column widths and row heights."""
    workbook = xlsxwriter.Workbook("/dev/null")
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
    assert heights == {}


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


def test_cells_that_cannot_wrap_do_not_grow_rows():
    long_text = "a very long piece of text that overflows its column by a lot"
    _, heights = write(
        ep.Table(
            data=pd.DataFrame({"col": [long_text]}),
            max_col_size=10,
            body_style=ep.Style(text_wrap=False),
        )
    )
    assert heights == {}


def test_row_heights_only_grow():
    """A second table sharing a row cannot shrink a height the first one needed."""
    workbook = xlsxwriter.Workbook("/dev/null")
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
    # A single word wider than the line is broken mid-word, as Excel does
    assert count_lines("x" * 200, _excel_to_px(get_text_size("x" * 20))) >= 10


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
    assert heights == {}


if __name__ == "__main__":
    pytest.main([__file__])
