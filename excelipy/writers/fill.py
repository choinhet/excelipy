import logging
from typing import Tuple

from xlsxwriter.workbook import Workbook, Worksheet

from excelipy.models import Fill, Style
from excelipy.style import process_style

log = logging.getLogger("excelipy")


def write_fill(
        workbook: Workbook,
        worksheet: Worksheet,
        component: Fill,
        default_style: Style,
        origin: Tuple[int, int] = (0, 0),
) -> Tuple[int, int]:
    log.debug(f"Writing fill at {origin}")

    style = process_style(workbook, [default_style, component.style])
    col0, row0 = origin
    width = component.width
    height = component.height

    if component.merged and (width > 1 or height > 1):
        worksheet.merge_range(row0, col0, row0 + height - 1, col0 + width - 1, "", style)
    else:
        for dy in range(height):
            for dx in range(width):
                row = row0 + dy
                col = col0 + dx
                worksheet.write_blank(row, col, "", style)

    return width, height
