import logging
from collections import defaultdict
from functools import reduce, wraps
from typing import Tuple, Dict, Optional, Set

import numpy as np
import pandas as pd
from PIL import ImageFont
from xlsxwriter.workbook import Workbook, Worksheet

from excelipy.models import Style, Table, Link
from excelipy.style import process_style
from excelipy.styles.table import DEFAULT_HEADER_STYLE, DEFAULT_BODY_STYLE

log = logging.getLogger("excelipy")

DEFAULT_FONT_SIZE = 11
DEFAULT_ROW_HEIGHT = 15.0
TUNING_DEFAULT = 5
PADDING_DEFAULT = 2
ROW_WISE_ARG = "_ep_row_wise"


def row_wise(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        return func(*args, **kwargs)

    setattr(wrapper, ROW_WISE_ARG, True)
    return wrapper


# ---------------------------------------------------------------------------
# Style helpers
# ---------------------------------------------------------------------------


def _static_col_style(component: Table, col_name: str, col_idx: int) -> Style:
    maybe = component.idx_column_style.get(col_idx) or component.column_style.get(
        col_name
    )
    return Style() if callable(maybe) or maybe is None else maybe


def _col_style_chain(
    component: Table,
    col_name: str,
    col_idx: int,
    default_style: Style,
) -> Tuple[Style, ...]:
    return (
        default_style,
        component.style,
        component.body_style,
        _static_col_style(component, col_name, col_idx),
    )


def get_style_font_family(*styles: Style) -> Optional[str]:
    cur_font = None
    for s in filter(None, styles):
        cur_font = s.font_family or cur_font
    return cur_font


def get_style_font_size(*styles: Style) -> Optional[int]:
    cur_font = None
    for s in filter(None, styles):
        cur_font = s.font_size or cur_font
    return cur_font


# ---------------------------------------------------------------------------
# Text measurement
# ---------------------------------------------------------------------------


def get_text_size(
    text: str,
    font_size: Optional[int] = None,
    font_family: Optional[str] = None,
) -> float:
    text = str(text)
    cur_font_size = font_size or DEFAULT_FONT_SIZE
    cur_font_family = font_family or "Arial"
    try:
        cur_font = ImageFont.truetype(f"{cur_font_family}.ttf".lower(), cur_font_size)
    except Exception as e:
        cur_font = ImageFont.load_default()
        log.debug(
            f"Could not load custom font {cur_font_family}, using default. Exception: {e}"
        )
    return cur_font.getlength(text)


def _px_to_excel(px: float, tuning: int, padding: int) -> float:
    return px // tuning + padding


def _excel_to_px(excel_units: float, tuning: int, padding: int) -> float:
    return (excel_units - padding) * tuning


def _header_font(cur_col: str, component: Table) -> Tuple[Optional[int], Optional[str]]:
    font_size = get_style_font_size(
        DEFAULT_HEADER_STYLE, component.style, component.header_style.get(cur_col)
    )
    font_family = get_style_font_family(
        DEFAULT_HEADER_STYLE, component.style, component.header_style.get(cur_col)
    )
    return font_size, font_family


def _header_excel_width(
    cur_col: str, component: Table, tuning: int, padding: int
) -> float:
    font_size, font_family = _header_font(cur_col, component)
    return _px_to_excel(get_text_size(cur_col, font_size, font_family), tuning, padding)


def _count_lines(text_px: float, col_px: float) -> int:
    """How many lines does text_px need when the column is col_px wide."""
    return max(1, int(text_px / max(col_px, 1)) + 1)


def _row_height_for_lines(lines: int, font_size: Optional[int]) -> float:
    fs = font_size or DEFAULT_FONT_SIZE
    return max(DEFAULT_ROW_HEIGHT, fs * 1.3 * lines)


# ---------------------------------------------------------------------------
# Column width cache
# ---------------------------------------------------------------------------


def _get_sheet_cache(workbook: Workbook, worksheet: Worksheet) -> Dict[int, float]:
    cache: Dict[str, Dict[int, float]] = getattr(
        workbook, "_excelipy_col_size_cache", {}
    )
    return cache.setdefault(worksheet.name, {})


def _set_col_width(
    workbook: Workbook,
    worksheet: Worksheet,
    abs_col_idx: int,
    width: float,
    min_col_size: Optional[float],
    max_col_size: Optional[float],
) -> float:
    """
    Clamp width to [min_col_size, max_col_size], then persist the maximum
    value seen so far for this column and call set_column.
    Always returns the value actually written.
    """
    if min_col_size is not None:
        width = max(width, min_col_size)
    if max_col_size is not None:
        width = min(width, max_col_size)

    sheet_cache = _get_sheet_cache(workbook, worksheet)
    final = max(sheet_cache.get(abs_col_idx, 0), width)
    sheet_cache[abs_col_idx] = final

    cache = getattr(workbook, "_excelipy_col_size_cache", {})
    cache[worksheet.name] = sheet_cache
    setattr(workbook, "_excelipy_col_size_cache", cache)

    worksheet.set_column(abs_col_idx, abs_col_idx, int(final))
    return final


# ---------------------------------------------------------------------------
# Auto-width calculation
# ---------------------------------------------------------------------------


def get_auto_width(
    cur_col: str,
    col_idx: int,
    data: pd.Series,
    component: Table,
    default_style: Style,
    is_merged_header: bool = False,
) -> float:
    """
    Compute ideal column width in Excel character units.

    Non-merged  → max(header_px, body_px): fits both with no overflow or waste.
    Merged      → body_px only: the post-pass grows the span to fit the header.

    min/max_col_size are applied by _set_col_width, not here, so this returns
    the content-driven width before clamping.
    """
    tuning = component.auto_width_tuning or TUNING_DEFAULT
    padding = component.auto_width_padding or PADDING_DEFAULT

    chain = _col_style_chain(component, cur_col, col_idx, default_style)
    col_font_size = get_style_font_size(*chain)
    col_font_family = get_style_font_family(*chain)

    body_px = (
        data.apply(str)
        .apply(lambda it: get_text_size(it, col_font_size, col_font_family))
        .max()
    )

    if not is_merged_header:
        font_size, font_family = _header_font(cur_col, component)
        header_px = get_text_size(cur_col, font_size, font_family)
        max_px = max(header_px, body_px)
    else:
        # Merged: body drives individual column width.
        # _fix_merged_header_widths will grow the span to fit the header.
        max_px = body_px

    return _px_to_excel(max_px, tuning, padding)


# ---------------------------------------------------------------------------
# Merged header post-pass
# ---------------------------------------------------------------------------


def _fix_merged_header_widths(
    workbook: Workbook,
    worksheet: Worksheet,
    component: Table,
    idx_by_header: Dict[str, list],
    origin: Tuple[int, int],
    this_table_widths: Dict[int, float],
) -> None:
    """
    Ensure each merged span is wide enough to contain its header text.
    Distributes any deficit proportionally across the span's columns.
    Respects min/max_col_size.
    """
    tuning = component.auto_width_tuning or TUNING_DEFAULT
    padding = component.auto_width_padding or PADDING_DEFAULT

    for cur_col, indices in idx_by_header.items():
        total_span = sum(this_table_widths.get(i, 0) for i in indices)
        header_len = _header_excel_width(cur_col, component, tuning, padding)

        # If a max is set, the header can wrap — span only needs to reach
        # max_col_size * n_cols at most.
        if component.max_col_size is not None:
            header_len = min(header_len, component.max_col_size * len(indices))

        if total_span >= header_len:
            continue

        deficit = header_len - total_span
        per_col = deficit / len(indices)
        for i in indices:
            new_width = this_table_widths.get(i, 0) + per_col
            this_table_widths[i] = new_width
            _set_col_width(
                workbook=workbook,
                worksheet=worksheet,
                abs_col_idx=origin[0] + i,
                width=new_width,
                min_col_size=component.min_col_size,
                max_col_size=component.max_col_size,
            )


# ---------------------------------------------------------------------------
# Row height calculation
# ---------------------------------------------------------------------------


def _calc_header_height(
    component: Table,
    idx_by_header: Dict[str, list],
    this_table_widths: Dict[int, float],
) -> float:
    """
    Estimate the header row height based on how many lines each header cell
    needs given the final column widths. Covers both merged and non-merged headers.
    """
    tuning = component.auto_width_tuning or TUNING_DEFAULT
    padding = component.auto_width_padding or PADDING_DEFAULT
    max_lines = 1
    last_font_size = None

    # Merged headers: span width is the sum of all columns in the group.
    merged_col_indices: Set[int] = set()
    for cur_col, indices in idx_by_header.items():
        font_size, font_family = _header_font(cur_col, component)
        last_font_size = font_size
        text_px = get_text_size(cur_col, font_size, font_family)
        total_span = sum(this_table_widths.get(i, 0) for i in indices)
        col_px = _excel_to_px(total_span, tuning, padding)
        max_lines = max(max_lines, _count_lines(text_px, col_px))
        merged_col_indices.update(indices)

    # Non-merged headers: each column stands alone.
    for col_idx, cur_col in enumerate(component.data.columns):
        if col_idx in merged_col_indices:
            continue
        font_size, font_family = _header_font(cur_col, component)
        last_font_size = font_size
        text_px = get_text_size(cur_col, font_size, font_family)
        col_width = this_table_widths.get(col_idx, 10)
        col_px = _excel_to_px(col_width, tuning, padding)
        max_lines = max(max_lines, _count_lines(text_px, col_px))

    return _row_height_for_lines(max_lines, last_font_size)


def _calc_body_row_height(
    row: pd.Series,
    col_widths: Dict[int, float],
    component: Table,
    default_style: Style,
) -> float:
    """
    Estimate the body row height based on the widest content in each cell.
    """
    tuning = component.auto_width_tuning or TUNING_DEFAULT
    padding = component.auto_width_padding or PADDING_DEFAULT
    max_lines = 1

    for col_idx, cell in enumerate(row):
        col_width = col_widths.get(col_idx, 10)
        chain = _col_style_chain(component, row.index[col_idx], col_idx, default_style)
        font_size = get_style_font_size(*chain) or DEFAULT_FONT_SIZE
        font_family = get_style_font_family(*chain) or "Arial"
        text_px = get_text_size(str(cell), font_size, font_family)
        col_px = _excel_to_px(col_width, tuning, padding)
        max_lines = max(max_lines, _count_lines(text_px, col_px))

    font_size = get_style_font_size(default_style) or DEFAULT_FONT_SIZE
    return _row_height_for_lines(max_lines, font_size)


# ---------------------------------------------------------------------------
# Main write function
# ---------------------------------------------------------------------------


def write_table(
    workbook: Workbook,
    worksheet: Worksheet,
    component: Table,
    default_style: Style,
    origin: Tuple[int, int] = (0, 0),
) -> Tuple[int, int]:
    x_size = component.data.shape[1]
    y_size = component.data.shape[0] + 1  # +1 for the header row

    # Build merged-header index: consecutive columns sharing the same name.
    idx_by_header: Dict[str, list] = defaultdict(list)
    if component.merge_equal_headers:
        for idx, cur_col in enumerate(component.data.columns):
            existing = idx_by_header[cur_col]
            if not existing or idx == existing[-1] + 1:
                existing.append(idx)
        idx_by_header = {k: v for k, v in idx_by_header.items() if len(v) >= 2}

    merged_col_indices: Set[int] = {
        i for indices in idx_by_header.values() for i in indices
    }

    wrap_style = Style(text_wrap=True) if component.wrap_header else None

    # Per-table column widths (relative col_idx → excel units).
    # this_table_widths: pre-clamp content-driven widths, used by the post-pass
    #   so merged header deficit calculations are isolated from other tables.
    # final_table_widths: post-clamp widths actually set on the sheet, used by
    #   row height calculations so line-wrap estimates match what Excel renders.
    this_table_widths: Dict[int, float] = {}
    final_table_widths: Dict[int, float] = {}

    # ------------------------------------------------------------------ headers
    for col_idx, cur_col in enumerate(component.data.columns):
        current_header_style = component.header_style.get(cur_col, component.style)
        header_styles = [default_style, current_header_style]
        if component.default_style:
            header_styles = [DEFAULT_HEADER_STYLE] + header_styles
        if wrap_style:
            header_styles = header_styles + [wrap_style]

        header_format = process_style(workbook, header_styles)
        header_write_skip = idx_by_header.get(cur_col, [])
        is_first = False
        cur_skip = col_idx in header_write_skip and not (
            is_first := col_idx == header_write_skip[0]
        )

        if is_first:
            worksheet.merge_range(
                origin[1],
                origin[0],
                origin[1],
                origin[0] + len(header_write_skip) - 1,
                cur_col,
                header_format,
            )
        elif not cur_skip:
            worksheet.write(origin[1], origin[0] + col_idx, cur_col, header_format)
        else:
            worksheet.write(origin[1], origin[0] + col_idx, "", header_format)

        # Column width
        set_width = component.idx_column_width.get(
            col_idx
        ) or component.column_width.get(cur_col)
        if set_width:
            estimated_width = float(set_width)
        else:
            estimated_width = get_auto_width(
                cur_col,
                col_idx,
                component.data.iloc[:, col_idx],
                component,
                default_style,
                is_merged_header=col_idx in merged_col_indices,
            )

        this_table_widths[col_idx] = estimated_width
        final_width = _set_col_width(
            workbook=workbook,
            worksheet=worksheet,
            abs_col_idx=origin[0] + col_idx,
            width=estimated_width,
            min_col_size=component.min_col_size,
            max_col_size=component.max_col_size,
        )
        final_table_widths[col_idx] = final_width
        log.debug(
            f"Estimated width for {cur_col}: {final_width} [Sheet: {worksheet.name}]"
        )

    # Post-pass: grow merged spans to fit their header text.
    if component.merge_equal_headers and idx_by_header:
        _fix_merged_header_widths(
            workbook, worksheet, component, idx_by_header, origin, this_table_widths
        )

    # Header row height — must run after the post-pass since _fix_merged_header_widths
    # updates this_table_widths in-place with the final expanded span widths.
    if component.wrap_header:
        header_height = _calc_header_height(component, idx_by_header, this_table_widths)
        log.debug(f"Header height: {header_height} [Sheet: {worksheet.name}]")
        if header_height > DEFAULT_ROW_HEIGHT:
            worksheet.set_row(origin[1], header_height)

    if component.header_filters:
        worksheet.autofilter(
            origin[1],
            origin[0],
            origin[1],
            origin[0] + len(list(component.data.columns)) - 1,
        )

    # -------------------------------------------------------------------- body
    for col_idx, col in enumerate(component.data.columns):
        col_style = _static_col_style(component, col, col_idx)
        for row_idx, (_, row) in enumerate(component.data.iterrows()):
            row_style = component.row_style.get(row_idx)
            body_style = [
                default_style,
                component.style,
                component.body_style,
                col_style,
                row_style,
            ]
            if component.default_style:
                body_style = [DEFAULT_BODY_STYLE] + body_style
            if wrap_style:
                body_style = body_style + [wrap_style]

            merged_style: Style = reduce(
                lambda acc, s: acc.merge(s),
                filter(None, body_style),
                Style(),
            )

            cell = row.iloc[col_idx]
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

            maybe_callable = component.idx_column_style.get(
                col_idx
            ) or component.column_style.get(col)
            if callable(maybe_callable):
                dyn_style = (
                    maybe_callable(row)
                    if getattr(maybe_callable, ROW_WISE_ARG, False)
                    else maybe_callable(cell)
                )
                if dyn_style is not None:
                    merged_style = merged_style.merge(dyn_style)

            if row_style is not None:
                merged_style = merged_style.merge(row_style)

            current_format = process_style(workbook, [merged_style])
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

    # Body row heights — computed per-row after all cells are written.
    # Only runs when wrap_header=True since that's when cells can overflow vertically.
    if component.wrap_header:
        for row_idx, (_, row) in enumerate(component.data.iterrows()):
            height = _calc_body_row_height(
                row, final_table_widths, component, default_style
            )
            if height > DEFAULT_ROW_HEIGHT:
                worksheet.set_row(origin[1] + row_idx + 1, height)

    return x_size, y_size
