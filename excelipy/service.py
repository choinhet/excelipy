import logging
from typing import Tuple, Dict, Callable

import xlsxwriter
from xlsxwriter.workbook import Workbook, Worksheet

from excelipy.models import Component, Excel, Fill, Style, Table, Text, Image, Link
from excelipy.writers import (
    write_fill,
    write_table,
    write_text,
    write_image, write_link,
)

log = logging.getLogger("excelipy")


def write_component(
        workbook: Workbook,
        worksheet: Worksheet,
        component: Component,
        default_style: Style,
        origin: Tuple[int, int] = (0, 0),
) -> Tuple[int, int]:
    writing_map: Dict[
        Callable[..., Component],
        Callable[..., Tuple[int, int]]
    ] = {
        Table: write_table,
        Text: write_text,
        Link: write_link,
        Fill: write_fill,
        Image: write_image,
    }

    render_func = writing_map.get(type(component))

    return render_func(
        workbook,
        worksheet,
        component,
        default_style,
        origin,
    )


def save(excel: Excel):
    with xlsxwriter.Workbook(excel.path, {
        "nan_inf_to_errors": excel.nan_inf_to_errors,
    }) as workbook:
        for sheet in excel.sheets:
            origin = (
                sheet.style.pl(),
                sheet.style.pt(),
            )
            worksheet = workbook.add_worksheet(sheet.name)
            if not sheet.grid_lines:
                worksheet.hide_gridlines(2)

            for component in sheet.components:
                cur_origin = (
                    origin[0] + component.style.pl(),
                    origin[1] + component.style.pt(),
                )
                x, y = write_component(
                    workbook,
                    worksheet,
                    component,
                    sheet.style,
                    cur_origin,
                )
                origin = origin[0] + component.style.pr(), origin[1] + y + component.style.pb()
