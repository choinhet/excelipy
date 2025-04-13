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

    header_format = process_style(
        workbook,
        [
            default_style,
            component.header_style,
        ],
    )

    col_formats = {}
    for header in headers:
        all_styles = [
            default_style,
            component.body_style,
            component.column_style.get(header),
        ]
        filtered_styles = [s for s in all_styles if s is not None]
        col_format = process_style(
            workbook,
            filtered_styles,
        )
        col_formats[header] = col_format

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
