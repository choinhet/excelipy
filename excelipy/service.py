import logging
from typing import Tuple

import pandas as pd
import xlsxwriter
from xlsxwriter.workbook import Workbook, Worksheet

from excelipy.models import Component, Excel, Fill, Table, Text

log = logging.getLogger("excelipy")


def write_table(
    workbook: Workbook,
    worksheet: Worksheet,
    component: Table,
    origin: Tuple[int, int] = (0, 0),
) -> Tuple[int, int]:
    headers = list(component.data.columns)
    data = component.data.values
    y_size = len(data)
    x_size = len(headers)
    table_options = {
        "data": data,
        "columns": [{"header": header} for header in headers],
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


def write_text(
    workbook: Workbook,
    worksheet: Worksheet,
    component: Text,
    origin: Tuple[int, int] = (0, 0),
) -> Tuple[int, int]:
    log.debug(f"Writing text at {origin}")
    worksheet.write(
        origin[1],
        origin[0],
        component.text,
    )
    return 1, 1


def write_fill(
    workbook: Workbook,
    worksheet: Worksheet,
    component: Fill,
    origin: Tuple[int, int] = (0, 0),
) -> Tuple[int, int]:
    log.debug(f"Writing fill at {origin}")
    style_dict = {"bg_color": "#303030"}
    worksheet.merge_range(
        origin[1],
        origin[0],
        origin[1] + component.height - 1,
        origin[0] + component.width - 1,
        "",
        workbook.add_format(style_dict),
    )
    return component.width, component.height


def write_component(
    workbook: Workbook,
    worksheet: Worksheet,
    component: Component,
    origin: Tuple[int, int] = (0, 0),
) -> Tuple[int, int]:
    writing_map = {
        Table: write_table,
        Text: write_text,
        Fill: write_fill,
    }
    render_func = writing_map.get(type(component))
    if render_func is None:
        return 0, 0
    return render_func(
        workbook,
        worksheet,
        component,
        origin,
    )


def save(excel: Excel):
    workbook = xlsxwriter.Workbook(excel.path)
    log.debug("Workbook opened")
    origin = (0, 0)
    for sheet in excel.sheets:
        worksheet = workbook.add_worksheet(sheet.name)
        for component in sheet.components:
            x, y = write_component(
                workbook,
                worksheet,
                component,
                origin,
            )
            origin = origin[0], origin[1] + y

    workbook.close()
    log.debug("Workbook closed")
