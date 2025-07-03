import logging
from typing import Tuple

from PIL import Image as PILImage
from xlsxwriter.workbook import Workbook, Worksheet

from excelipy.models import Image, Style
from excelipy.style import process_style

log = logging.getLogger("excelipy")

DEFAULT_COLUMN_WIDTH = 64
DEFAULT_ROW_HEIGHT = 20


def write_image(
        workbook: Workbook,
        worksheet: Worksheet,
        component: Image,
        default_style: Style,
        origin: Tuple[int, int] = (0, 0),
) -> Tuple[int, int]:
    log.debug(f"Writing image at {origin}")
    img = PILImage.open(component.path)
    img_w, img_h = img.size
    log.debug(f"img_w={img_w}, img_h={img_h}")
    expected_width = DEFAULT_COLUMN_WIDTH * component.width
    expected_height = DEFAULT_ROW_HEIGHT * component.height
    log.debug(
        f"expected_width={expected_width}, expected_height={expected_height}"
    )
    scale_width = expected_width / img_w
    scale_height = expected_height / img_h
    log.debug(f"scale_width={scale_width}, scale_height={scale_height}")

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

    worksheet.insert_image(
        origin[1],
        origin[0],
        component.path.as_posix(),
        {
            'x_scale': scale_width,
            'y_scale': scale_height,
            'object_position': 1,  # Move and size with cells
        }
    )
    return component.width, component.height
