import logging
import math
from collections import defaultdict
from functools import lru_cache, wraps
from typing import cast

import numpy as np
import pandas as pd
from PIL import ImageFont
from xlsxwriter.workbook import Workbook, Worksheet

from excelipy.models import Link, Style, StyleFunc, Table
from excelipy.style import merge_styles, process_style
from excelipy.styles.table import DEFAULT_BODY_STYLE, DEFAULT_HEADER_STYLE

log = logging.getLogger("excelipy")

DEFAULT_FONT_SIZE = 11
DEFAULT_LINE_SPACING = 1.4
DEFAULT_ROW_HEIGHT = 15.0
DEFAULT_FONT_FAMILY = "Calibri"

TUNING_DEFAULT = 5
PADDING_DEFAULT = 2

ROW_WISE_ARG = "_ep_row_wise"
COL_CACHE_NAME = "_excelipy_col_sizes"


def row_wise(func):
    """
    Marks a StyleFunc to receive all columns instead of only the current column values
    """

    @wraps(func)
    def wrapper(*args, **kwargs):
        return func(*args, **kwargs)

    setattr(wrapper, ROW_WISE_ARG, True)
    return wrapper


def _static_col_style(component: Table, col_name: str, col_idx: int) -> Style:
    idx_style = component.idx_column_style.get(col_idx)
    col_style = component.column_style.get(col_name)
    maybe = idx_style or col_style
    return Style() if callable(maybe) or maybe is None else maybe


@lru_cache
def _load_font(
    font_family: str,
    font_size: int,
) -> ImageFont.ImageFont | ImageFont.FreeTypeFont:
    try:
        return ImageFont.truetype(f"{font_family.lower()}.ttf", font_size)
    except Exception as e:
        log.debug(
            f"Could not load custom font {font_family}, using default.\nException: {e}"
        )
        return ImageFont.load_default()


def _px_to_excel(px: float) -> int:
    return int(px // TUNING_DEFAULT + PADDING_DEFAULT)


@lru_cache
def get_char_size(
    char: str,
    font_size: int,
    font_family: str,
) -> int | float:
    return _load_font(font_family, font_size).getlength(char)


def get_text_size(
    text: str,
    font_size: int | None = None,
    font_family: str | None = None,
) -> int:
    cur_font_size = font_size or DEFAULT_FONT_SIZE
    cur_font_family = font_family or DEFAULT_FONT_FAMILY
    total_size = 0
    for char in str(text):
        total_size += get_char_size(char, cur_font_size, cur_font_family)
    return _px_to_excel(total_size)


def get_row_height(lines: int, font_size: int | None) -> float:
    return max(
        DEFAULT_ROW_HEIGHT,
        (font_size or DEFAULT_FONT_SIZE) * DEFAULT_LINE_SPACING * lines,
    )


def _maybe_format(text: float | int | str, num_format: str | None) -> str:
    """
    Examples:
        >>> _maybe_format(1.2321, None)
        '1.2321'
        >>> _maybe_format(1.2321, ",.2f")
        '1.23'
        >>> _maybe_format(1.2321, ",d")
        '1'
        >>> _maybe_format(20000, ",d")
        '20,000'
        >>> _maybe_format("text", ".2f")
        'text'
    """
    if num_format is None:
        return str(text)
    clz = int
    if "." in str(text) or "f" in num_format:
        clz = float
    if "d" in num_format:
        clz = int
    try:
        return format(clz(text), num_format)
    except Exception:
        return str(text)


def write_table(
    workbook: Workbook,
    worksheet: Worksheet,
    component: Table,
    default_style: Style,
    origin: tuple[int, int] = (0, 0),
) -> tuple[int, int]:
    """
    Examples:
        >>> n = 30_000
        >>> data = pd.DataFrame({"A": [1, 2, 3] * n, "B": [4, "ha" * 50, 6] * n, "C": [4, 5, 6] * n})
        >>> long_text = "This is an avocado toast" * 3
        >>> data.rename(columns={"A": long_text, "B": long_text}, inplace=True)
        >>> origin = (0, 0)
        >>> default_style = Style(align="center", valign="vcenter")
        >>> row_style = {1: Style(font_size=14)}
        >>> component = Table(data=data, row_style=row_style, min_col_size=10, max_col_size=20, wrap_header=True)
        >>> import xlsxwriter
        >>> workbook = xlsxwriter.Workbook("output.xlsx")
        >>> worksheet = workbook.add_worksheet()
        >>> _ = write_table(workbook, worksheet, component, default_style)
        >>> workbook.close()
    """
    x_size = component.data.shape[1]
    y_size = component.data.shape[0] + 1

    df_columns = list(component.data.columns)
    df_rows = component.data.values.tolist()

    header_size_cache: dict[int, tuple[int, int | None]] = {}
    body_size_cache: dict[int, dict[int, tuple[int, int | None]]] = defaultdict(dict)
    biggest_body: dict[int, int] = defaultdict(lambda: 0)

    base_column_range = [(idx, idx) for idx in range(len(df_columns))]
    column_ranges = list(base_column_range)
    # =============================== Write headers ================================
    prev = None
    prev_format = None
    min_idx = 0
    for col_idx, cur_col in enumerate(df_columns):
        header_style = merge_styles(
            DEFAULT_HEADER_STYLE if component.default_style else None,
            default_style,
            component.style,
            component.header_style.get(cur_col),
            Style(text_wrap=True) if component.wrap_header else None,
        )
        header_format = process_style(workbook, [header_style])
        worksheet.write(origin[1], origin[0] + col_idx, cur_col, header_format)
        if component.merge_equal_headers:
            if prev is not None and prev != cur_col and col_idx - min_idx > 1:
                worksheet.merge_range(
                    first_row=origin[1],
                    first_col=origin[0] + min_idx,
                    last_row=origin[1],
                    last_col=origin[0] + col_idx - 1,
                    data=prev,
                    cell_format=prev_format,
                )
            if prev is not None and prev != cur_col:
                for _idx in range(min_idx, col_idx):
                    column_ranges.remove((_idx, _idx))
                column_ranges.append((min_idx, col_idx - 1))
                column_ranges.sort(key=lambda x: x[0])
                min_idx = col_idx
            prev = cur_col
            prev_format = header_format
        if component.auto_size:
            header_size_cache[col_idx] = (
                get_text_size(
                    cur_col,
                    header_style.font_size,
                    header_style.font_family,
                ),
                header_style.font_size,
            )

    # =============================== Header filters ===============================
    actually_merged = set(base_column_range) != set(column_ranges)
    if component.header_filters and not actually_merged:
        worksheet.autofilter(
            origin[1],
            origin[0],
            origin[1],
            origin[0] + len(list(component.data.columns)) - 1,
        )

    # ================================= Write body =================================
    for col_idx, col in enumerate(df_columns):
        base_style = merge_styles(
            DEFAULT_BODY_STYLE if component.default_style else None,
            default_style,
            component.style,
            component.body_style,
            _static_col_style(component, col, col_idx),
            Style(text_wrap=True) if component.wrap_header else None,
        )
        _maybe = Style | StyleFunc | None
        maybe_func_col_style: _maybe = component.column_style.get(col)
        maybe_func_idx_col_style: _maybe = component.idx_column_style.get(col_idx)
        maybe_func_style = maybe_func_idx_col_style or maybe_func_col_style
        style_func: StyleFunc | None = None
        if callable(maybe_func_style):
            style_func: StyleFunc = cast(StyleFunc, maybe_func_style)
        for row_idx, row in enumerate(df_rows):
            cell = row[col_idx]
            row_style = component.row_style.get(row_idx)
            merged_style = (
                base_style.merge(row_style) if row_style is not None else base_style
            )
            url = None
            if isinstance(cell, Link):
                url = cell.url
                cell = cell.text
            if merged_style.fill_na is not None and pd.isna(cell):
                cell = merged_style.fill_na
                merged_style = merged_style.model_copy(update=dict(numeric_format=None))
            if merged_style.fill_zero is not None and cell == 0:
                cell = merged_style.fill_zero
                merged_style = merged_style.model_copy(update=dict(numeric_format=None))
            if merged_style.fill_inf is not None and cell in (np.inf, -np.inf):
                cell = merged_style.fill_inf
                merged_style = merged_style.model_copy(update=dict(numeric_format=None))
            if style_func:
                dyn_style = (
                    style_func(row)
                    if getattr(style_func, ROW_WISE_ARG, False)
                    else style_func(cell)
                )
                merged_style = merged_style.merge(dyn_style)
            if (row_style := component.row_style.get(row_idx)) is not None:
                merged_style = merged_style.merge(row_style)
            current_format = process_style(workbook, [merged_style])

            if component.auto_size:
                cur_txt_size = get_text_size(
                    str(cell),
                    merged_style.font_size,
                    merged_style.font_family,
                )
                body_size_cache[col_idx][row_idx] = (
                    cur_txt_size,
                    merged_style.font_size,
                )
                biggest_body[col_idx] = max(cur_txt_size, biggest_body[col_idx])

            if url is None:
                worksheet.write(
                    origin[1] + row_idx + 1,
                    origin[0] + col_idx,
                    cell,
                    current_format,
                )
            else:
                worksheet.write_url(
                    origin[1] + row_idx + 1,
                    origin[0] + col_idx,
                    url,
                    current_format,
                    cell,
                )

    # =============================== Auto Set Width ===============================
    if component.auto_size:
        col_sizes = getattr(worksheet, COL_CACHE_NAME, None) or defaultdict(lambda: 0)
        # Compare cache to body
        for col_idx, text_size in biggest_body.items():
            col_sizes[origin[0] + col_idx] = max(
                text_size,
                col_sizes[origin[0] + col_idx],
            )
        # Compare cache to header (considering merged spans)
        for beg, end in column_ranges:
            text_size = header_size_cache[beg][0]
            cur_body_sizes = [
                col_sizes[origin[0] + col_idx] for col_idx in range(beg, end + 1)
            ]
            num_cols = end - beg + 1
            total_size = sum(cur_body_sizes)
            diff = text_size - total_size
            if diff > 0:
                to_increase = diff // num_cols
                for col_idx in range(beg, end + 1):
                    col_sizes[origin[0] + col_idx] += to_increase
        # Hard set sizes
        for col, width in component.column_width.items():
            idxs = [i for i, c in enumerate(df_columns) if col == c]
            for idx in idxs:
                col_sizes[origin[0] + idx] = width
        # apply constraints
        for sheet_idx, text_size in col_sizes.items():
            if component.min_col_size and text_size < component.min_col_size:
                text_size = component.min_col_size
            if component.max_col_size and text_size > component.max_col_size:
                text_size = component.max_col_size
            col_sizes[sheet_idx] = text_size
            worksheet.set_column(sheet_idx, sheet_idx, col_sizes[sheet_idx])
        setattr(worksheet, COL_CACHE_NAME, col_sizes)
        if component.wrap_header:
            # row wrap headers
            for beg, end in column_ranges:
                text_size, text_font = header_size_cache[beg]
                cur_body_sizes = [
                    col_sizes[origin[0] + col_idx] for col_idx in range(beg, end + 1)
                ]
                line_size = sum(cur_body_sizes)
                diff = text_size - line_size
                if diff > 0:
                    lines_needed = math.ceil(text_size / line_size)
                    row_height = get_row_height(lines_needed, text_font)
                    worksheet.set_row(origin[1], row_height)
            # row wrap body
            for row_idx in range(len(df_rows)):
                biggest_diff = 0
                row_size = 0
                row_font = None
                biggest_col_size = 0
                for col, rows in body_size_cache.items():
                    col_size = col_sizes[origin[0] + col]
                    cur_row, cur_font = rows[row_idx]
                    diff = cur_row - col_size
                    if diff > biggest_diff:
                        biggest_diff = diff
                        row_size = cur_row
                        biggest_col_size = col_size
                        row_font = cur_font
                if biggest_diff > 0:
                    lines_needed = math.ceil(row_size / biggest_col_size)
                    row_height = get_row_height(lines_needed, row_font)
                    worksheet.set_row(origin[1] + row_idx + 1, row_height)

    return x_size, y_size
