import logging
from typing import Tuple

from xlsxwriter.workbook import Workbook, Worksheet

from excelipy.models import Style, Link
from excelipy.style import process_style
from excelipy.styles.link import DEFAULT_LINK_STYLE

log = logging.getLogger("excelipy")


def write_link(
        workbook: Workbook,
        worksheet: Worksheet,
        component: Link,
        default_style: Style,
        origin: Tuple[int, int] = (0, 0),
) -> Tuple[int, int]:
    log.debug(f"Writing link at {origin}")

    processed_style = process_style(
        workbook,
        [
            DEFAULT_LINK_STYLE,
            default_style,
            component.style,
        ],
    )

    col0, row0 = origin
    width = component.width
    height = component.height

    if component.merged and (width > 1 or height > 1):
        worksheet.merge_range(row0, col0, row0 + height - 1, col0 + width - 1, component.text, processed_style)
        worksheet.write_url(row0, col0, component.url, processed_style, component.text)
    else:
        for dy in range(height):
            for dx in range(width):
                row = row0 + dy
                col = col0 + dx
                if dy == 0 and dx == 0:
                    worksheet.write_url(row, col, component.url, processed_style, component.text)
                else:
                    worksheet.write_blank(row, col, "", processed_style)

    return width, height
