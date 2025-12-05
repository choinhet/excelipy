import logging
from typing import Tuple

from xlsxwriter.workbook import Workbook, Worksheet

from excelipy.models import Style, Text
from excelipy.style import process_style
from excelipy.styles.text import DEFAULT_TEXT_STYLE

log = logging.getLogger("excelipy")


def write_text(
        workbook: Workbook,
        worksheet: Worksheet,
        component: Text,
        default_style: Style,
        origin: Tuple[int, int] = (0, 0),
) -> Tuple[int, int]:
    log.debug(f"Writing text at {origin}")

    processed_style = process_style(
        workbook,
        [
            DEFAULT_TEXT_STYLE,
            default_style,
            component.style,
        ],
    )

    col0, row0 = origin
    width = component.width
    height = component.height

    if component.merged and (width > 1 or height > 1):
        worksheet.merge_range(
            row0,
            col0,
            row0 + height - 1,
            col0 + width - 1,
            component.text,
            processed_style,
        )
    else:
        for dy in range(height):
            for dx in range(width):
                row = row0 + dy
                col = col0 + dx
                if dy == 0 and dx == 0:
                    worksheet.write(row, col, component.text, processed_style)
                else:
                    worksheet.write_blank(row, col, "", processed_style)

    return width, height
