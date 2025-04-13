import logging
from typing import Tuple

from xlsxwriter.workbook import Workbook, Worksheet

from excelipy.models import Style, Table
from excelipy.style import process_style

log = logging.getLogger("excelipy")


def write_table(
    workbook: Workbook,
    worksheet: Worksheet,
    component: Table,
    default_style: Style,
    origin: Tuple[int, int] = (0, 0),
) -> Tuple[int, int]:
    headers = list(component.data.columns)
    data = component.data.values

    x_size = len(headers)
    y_size = len(data)

    col_formats = {}
    for idx, col in enumerate(component.data.columns):
        max_content_size = component.data[col].apply(str).apply(len).max()
        max_size = max(len(col), max_content_size)
        cur_col = origin[0] + idx

        all_styles = [
            default_style,
            component.body_style,
            component.column_style.get(col),
        ]
        filtered_styles = [s for s in all_styles if s is not None]
        col_format = process_style(
            workbook,
            filtered_styles,
        )
        col_formats[col] = col_format

        col_formats[col] = col_format
        f = col_format.font_size
        cur_size = ((f * 10 // 4) + (max_size * 18)) // 10

        cur_size = min(
            cur_size,
            component.max_col_width or cur_size,
        )

        worksheet.set_column(cur_col, cur_col, cur_size)

    header_format = process_style(
        workbook,
        [
            default_style,
            component.header_style,
        ],
    )

    table_options = {
        "data": data,
        "columns": [
            {
                "header": header,
                "header_format": header_format,
                "format": col_formats[header],
            }
            for header in headers
        ],
        "style": component.predefined_style,
    }

    log.debug(f"Writing table at {origin}")
    worksheet.add_table(
        origin[1],
        origin[0],
        y_size + origin[1],
        x_size + origin[0] - 1,
        table_options,
    )
    return x_size, y_size + 1
