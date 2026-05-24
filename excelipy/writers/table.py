import logging
import math
from collections import defaultdict
from functools import lru_cache, wraps
from typing import TypedDict, cast

import numpy as np
import pandas as pd
from PIL import ImageFont
from xlsxwriter.format import Format
from xlsxwriter.workbook import Workbook, Worksheet

from excelipy.models import Link, Style, StyleFunc, Table
from excelipy.style import merge_styles, process_style
from excelipy.styles.table import DEFAULT_BODY_STYLE, DEFAULT_HEADER_STYLE

log = logging.getLogger("excelipy")

DEFAULT_FONT_SIZE = 11
DEFAULT_FONT_FAMILY = "Calibri"

DEFAULT_ROW_HEIGHT = 15.0
TUNING_DEFAULT = 5
PADDING_DEFAULT = 2
ROW_WISE_ARG = "_ep_row_wise"


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


@lru_cache(maxsize=32)
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


def get_text_size(
    text: str,
    font_size: int | None = None,
    font_family: str | None = None,
) -> int:
    text = str(text)
    cur_font_size = font_size or DEFAULT_FONT_SIZE
    cur_font_family = font_family or "Calibri"
    cur_font = _load_font(cur_font_family, cur_font_size)
    size_px = cur_font.getlength(text)
    return _px_to_excel(size_px)


def trunc(num: float, precision: int = 1) -> float:
    x = 10**precision
    return int(num * x) / x


def _count_lines(text: str, text_px: float, col_px: float) -> int:
    """How many lines does text_px need when the column is col_px wide."""
    ratio = trunc(text_px / max(col_px, 1), 1)
    result = math.ceil(ratio)
    return result


def _row_height_for_lines(lines: int, font_size: int | None) -> float:
    fs = font_size or DEFAULT_FONT_SIZE
    return max(DEFAULT_ROW_HEIGHT, fs * 1.3 * lines)


def _maybe_format(text: float | int | str, num_format: str | None) -> str:
    """
    Examples:
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
        >>> data = pd.DataFrame({"A": [1, 2, 3] * n, "B": [4, 5, 6] * n, "C": [4, 5, 6] * n})
        >>> long_text = "This is an avocado toast" * 3
        >>> data.rename(columns={"A": long_text, "B": long_text}, inplace=True)
        >>> default_style = Style(align="center", valign="vcenter")
        >>> row_style = {1: Style(font_size=14)}
        >>> component = Table(data=data, row_style=row_style, min_col_size=10, max_col_size=20)
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

    # =============================== Write headers ================================
    class SheetCache(TypedDict):
        content: str
        style: Style
        format: Format

    header_cache: dict[int, SheetCache] = {}
    body_cache: dict[int, dict[int, SheetCache]] = defaultdict(dict)
    for col_idx, cur_col in enumerate(df_columns):
        header_style = merge_styles(
            DEFAULT_HEADER_STYLE if component.default_style else None,
            default_style,
            component.style,
            component.header_style.get(cur_col),
            Style(text_wrap=True) if component.wrap_header else None,
        )
        header_format = process_style(workbook, [header_style])
        header_cache[col_idx] = {
            "format": header_format,
            "style": header_style,
            "content": str(cur_col),
        }
        worksheet.write(origin[1], origin[0] + col_idx, cur_col, header_format)

    # =============================== Merge Headers ================================
    column_ranges = [(idx, idx) for idx in range(len(df_columns))]
    if component.merge_equal_headers:
        column_ranges = []
        prev = None
        min_idx = 0
        for idx, col in enumerate(df_columns):
            if prev is not None and prev != col and idx - min_idx > 1:
                worksheet.merge_range(
                    first_row=origin[1],
                    first_col=origin[0] + min_idx,
                    last_row=origin[1],
                    last_col=origin[0] + idx - 1,
                    data=prev,
                    cell_format=header_cache[min_idx]["format"],
                )
            if prev is not None and prev != col:
                column_ranges.append((min_idx, idx - 1))
                min_idx = idx
            prev = col
        column_ranges.append((min_idx, min_idx))

    # =============================== Header filters ===============================
    if component.header_filters and not component.merge_equal_headers:
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
        maybe_func_idx_col_style: _maybe = component.idx_column_style.get(col)
        maybe_func_style = maybe_func_col_style or maybe_func_idx_col_style
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
            body_cache[col_idx][row_idx] = {
                "format": current_format,
                "style": merged_style,
                "content": str(cell),
            }
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
    if component.auto_width:
        # ================================ Body Maximum ================================
        biggest_body: dict[int, tuple] = defaultdict(lambda: (0, "", 0))
        for col_idx, rows in body_cache.items():
            for _, row in rows.items():
                cell = row["content"]
                font_size = row["style"].font_size or DEFAULT_FONT_SIZE
                font_family = row["style"].font_family or DEFAULT_FONT_FAMILY
                num_format = row["style"].numeric_format
                formatted_content = _maybe_format(cell, num_format)
                # the product is an approx. so we don't need to calculate for the whole dataframe
                cur_size = len(formatted_content) * font_size
                if cur_size > biggest_body[col_idx][0]:
                    biggest_body[col_idx] = (
                        cur_size,
                        formatted_content,
                        font_size,
                        font_family,
                    )
        cache_name = "_excelipy_col_sizes"
        col_sizes = getattr(worksheet, cache_name, None) or defaultdict(lambda: 0)
        for col_idx, (_, content, font_size, font_family) in biggest_body.items():
            txt_size = get_text_size(content, font_size, font_family)
            col_sizes[origin[0] + col_idx] = max(
                txt_size, col_sizes[origin[0] + col_idx]
            )

        # =============================== Header Maximum ===============================
        for beg, end in column_ranges:
            cache = header_cache[beg]
            style = cache["style"]
            font_size = style.font_size or DEFAULT_FONT_SIZE
            font_family = style.font_family or DEFAULT_FONT_FAMILY
            content = cache["content"]
            txt_size = get_text_size(content, font_size, font_family)
            cur_body_sizes = [
                col_sizes[origin[0] + col_idx] for col_idx in range(beg, end + 1)
            ]
            num_cols = end - beg + 1
            total_size = sum(cur_body_sizes)
            diff = txt_size - total_size
            if diff > 0:
                to_increase = diff // num_cols
                for col_idx in range(beg, end + 1):
                    col_sizes[origin[0] + col_idx] += to_increase

        # ================================ Actual Sizes ================================
        for sheet_idx, txt_size in col_sizes.items():
            if component.min_col_size and txt_size < component.min_col_size:
                txt_size = component.min_col_size
            if component.max_col_size and txt_size > component.max_col_size:
                txt_size = component.max_col_size
            col_sizes[sheet_idx] = txt_size
            worksheet.set_column(sheet_idx, sheet_idx, col_sizes[sheet_idx])
        setattr(worksheet, cache_name, col_sizes)

    return x_size, y_size
