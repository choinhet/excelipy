import logging
from typing import Tuple

from xlsxwriter.workbook import Workbook, Worksheet

from excelipy.models import Style, Table
from excelipy.style import process_style

log = logging.getLogger("excelipy")

DEFAULT_FONT_SIZE = 11
SCALING_FACTOR = 1
BASE_PADDING = 3


def get_auto_width(
        header: str,
        component: Table,
        default_style: Style,
) -> float:
    header_len = len(header)
    col_len = component.data[header].apply(str).apply(len).max()
    max_len = max(header_len, col_len)
    header_font_size = (
            component.header_style.font_size
            or default_style.font_size
            or DEFAULT_FONT_SIZE
    )
    column_font_size = (
            component.column_style.get(header, Style()).font_size
            or component.body_style.font_size
            or default_style.font_size
            or DEFAULT_FONT_SIZE
    )
    max_row_font_size = max(
        list(
            s.font_size
            or component.body_style.font_size
            or default_style.font_size
            or DEFAULT_FONT_SIZE
            for s in component.row_style.values()
        ) or [DEFAULT_FONT_SIZE]
    )

    max_font_size = max(
        header_font_size,
        column_font_size,
        max_row_font_size,
    )
    font_factor = max_font_size / DEFAULT_FONT_SIZE
    return SCALING_FACTOR * font_factor * max_len + BASE_PADDING


def write_table(
        workbook: Workbook,
        worksheet: Worksheet,
        component: Table,
        default_style: Style,
        origin: Tuple[int, int] = (0, 0),
) -> Tuple[int, int]:
    x_size = component.data.shape[1]
    y_size = component.data.shape[0] + 1  # +1 for header row

    header_format = process_style(
        workbook,
        [default_style, component.header_style]
        )
    for col_idx, header in enumerate(component.data.columns):
        worksheet.write(
            origin[1],
            origin[0] + col_idx,
            header,
            header_format,
        )
        set_width = component.column_width.get(header)
        if set_width:
            estimated_width = set_width
        else:
            estimated_width = get_auto_width(header, component, default_style)
        worksheet.set_column(
            origin[1],
            origin[0] + col_idx,
            int(estimated_width)
        )

    if component.header_filters:
        worksheet.autofilter(
            origin[1],
            origin[0],
            origin[1],
            origin[0] + len(list(component.data.columns)) - 1,
        )

    for col_idx, col in enumerate(component.data.columns):
        col_style = component.column_style.get(col)
        for row_idx, (_, row) in enumerate(component.data.iterrows()):
            row_style = component.row_style.get(row_idx)
            non_none = filter(
                None,
                [
                    default_style,
                    component.body_style,
                    col_style,
                    row_style,
                ],
            )
            current_format = process_style(workbook, list(non_none))
            cell = row[col]
            worksheet.write(
                origin[1] + row_idx + 1,
                origin[0] + col_idx,
                cell,
                current_format,
            )

    return x_size, y_size
