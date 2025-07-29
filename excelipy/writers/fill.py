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
    if component.width > 1 or component.height > 1:
        worksheet.merge_range(
            origin[1],
            origin[0],
            origin[1] + component.height - 1,
            origin[0] + component.width - 1,
            "",
            process_style(
                workbook,
                [
                    default_style,
                    component.style,
                ],
            ),
        )
    else:
        worksheet.write_blank(
            origin[1],
            origin[0],
            "",
            process_style(
                workbook,
                [
                    default_style,
                    component.style,
                ],
            ),
            )

    return component.width, component.height
