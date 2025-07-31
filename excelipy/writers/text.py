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

    if component.width > 1 or component.height > 1:
        worksheet.merge_range(
            origin[1],
            origin[0],
            origin[1] + component.height - 1,
            origin[0] + component.width - 1,
            "",
            processed_style,
        )

    worksheet.write(
        origin[1],
        origin[0],
        component.text,
        processed_style,
    )
    return component.width, component.height
