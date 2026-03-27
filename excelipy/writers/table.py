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
TUNING_DEFAULT = 5
PADDING_DEFAULT = 2
ROW_WISE_ARG = "_ep_row_wise"


def row_wise(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        return func(*args, **kwargs)

    setattr(wrapper, ROW_WISE_ARG, True)
    return wrapper


def _static_col_style(component: Table, col_name: str, col_idx: int) -> Style:
    maybe = component.idx_column_style.get(col_idx) or component.column_style.get(
        col_name
    )
    return Style() if callable(maybe) or maybe is None else maybe


def get_text_size(
    text: str,
    font_size: Optional[int] = None,
    font_family: Optional[str] = None,
) -> float:
    text = str(text)
    cur_font_size = font_size or DEFAULT_FONT_SIZE
    cur_font_family = font_family or "Arial"

    try:
        cur_font = ImageFont.truetype(
            f"{cur_font_family}.ttf".lower(),
            cur_font_size,
        )
    except Exception as e:
        cur_font = ImageFont.load_default()
        log.debug(
            f"Could not load custom font {cur_font_family}, using default. Exception: {e}",
        )

    return cur_font.getlength(text)


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
) -> float:
    """
    Update the sheet-level column width cache to max(existing, width),
    then apply it to the worksheet. Always returns the value actually set.
    """
    sheet_cache = _get_sheet_cache(workbook, worksheet)
    final = max(sheet_cache.get(abs_col_idx, 0), width)
    sheet_cache[abs_col_idx] = final
    # Ensure the parent cache dict is attached (setdefault already does this, but be explicit).
    cache = getattr(workbook, "_excelipy_col_size_cache", {})
    cache[worksheet.name] = sheet_cache
    setattr(workbook, "_excelipy_col_size_cache", cache)
    worksheet.set_column(abs_col_idx, abs_col_idx, int(final))
    return final


def _px_to_excel(px: float, tuning: int, padding: int) -> float:
    return px // tuning + padding


def _header_excel_width(
    cur_col: str, component: Table, tuning: int, padding: int
) -> float:
    font_size = get_style_font_size(
        DEFAULT_HEADER_STYLE,
        component.style,
        component.header_style.get(cur_col),
    )
    font_family = get_style_font_family(
        DEFAULT_HEADER_STYLE,
        component.style,
        component.header_style.get(cur_col),
    )
    return _px_to_excel(get_text_size(cur_col, font_size, font_family), tuning, padding)


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

    - Non-merged: max(header_px, body_px) — no overflow, no wasted space.
    - Merged: body only — post-pass grows the span to fit the header.

    wrap_header=True caps at max_col_size; cells get wrap_text.
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
        header_px = get_text_size(
            cur_col,
            get_style_font_size(
                DEFAULT_HEADER_STYLE,
                component.style,
                component.header_style.get(cur_col),
            ),
            get_style_font_family(
                DEFAULT_HEADER_STYLE,
                component.style,
                component.header_style.get(cur_col),
            ),
        )
        max_px = max(header_px, body_px)
    else:
        max_px = body_px

    result = _px_to_excel(max_px, tuning, padding)

    if component.wrap_header and component.max_col_size is not None:
        result = min(result, component.max_col_size)

    return result


def _fix_merged_header_widths(
    workbook: Workbook,
    worksheet: Worksheet,
    component: Table,
    idx_by_header: Dict[str, list],
    origin: Tuple[int, int],
    this_table_widths: Dict[int, float],
) -> None:
    """
    Post-pass: grow each merged span so its total width fits the header text.

    Uses this_table_widths (relative col_idx → excel units) so comparisons are
    isolated from other tables on the same sheet.

    _set_col_width ensures the cross-table cache is also updated, so a later
    narrower table cannot overwrite a width that was expanded here.
    """
    tuning = component.auto_width_tuning or TUNING_DEFAULT
    padding = component.auto_width_padding or PADDING_DEFAULT

    for cur_col, indices in idx_by_header.items():
        total_span = sum(this_table_widths.get(i, 0) for i in indices)
        header_len = _header_excel_width(cur_col, component, tuning, padding)

        if component.wrap_header and component.max_col_size is not None:
            header_len = min(header_len, component.max_col_size * len(indices))

        if total_span < header_len:
            deficit = header_len - total_span
            per_col = deficit / len(indices)
            for i in indices:
                new_width = this_table_widths.get(i, 0) + per_col
                if component.wrap_header and component.max_col_size is not None:
                    new_width = min(new_width, component.max_col_size)
                this_table_widths[i] = new_width
                # _set_col_width takes max with cache and calls set_column.
                _set_col_width(workbook, worksheet, origin[0] + i, new_width)


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

    # Width computed for each column in THIS table (relative col_idx → excel units).
    # Isolated from the cross-table cache so the post-pass compares correctly.
    this_table_widths: Dict[int, float] = {}

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
            merge_size = len(header_write_skip)
            worksheet.merge_range(
                origin[1],
                origin[0],
                origin[1],
                origin[0] + merge_size - 1,
                cur_col,
                header_format,
            )
        elif not cur_skip:
            worksheet.write(origin[1], origin[0] + col_idx, cur_col, header_format)
        else:
            worksheet.write(origin[1], origin[0] + col_idx, "", header_format)

        # --------------------------------------------------------- column width
        set_width = component.idx_column_width.get(
            col_idx
        ) or component.column_width.get(cur_col)
        if set_width:
            estimated_width = float(set_width)
        else:
            data = component.data.iloc[:, col_idx]
            estimated_width = get_auto_width(
                cur_col,
                col_idx,
                data,
                component,
                default_style,
                is_merged_header=col_idx in merged_col_indices,
            )

        this_table_widths[col_idx] = estimated_width

        # _set_col_width takes max(cache, estimated) and calls set_column.
        final_width = _set_col_width(
            workbook, worksheet, origin[0] + col_idx, estimated_width
        )

        log.debug(
            f"Estimated width for {cur_col}: {final_width} [Sheet: {worksheet.name}]"
        )

    # Post-pass: grow merged spans to fit their header text.
    # Runs before any later table can overwrite set_column — and _set_col_width
    # in the post-pass also updates the cache, so later tables cannot narrow it back.
    if component.merge_equal_headers and idx_by_header:
        _fix_merged_header_widths(
            workbook, worksheet, component, idx_by_header, origin, this_table_widths
        )

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

    return x_size, y_size
