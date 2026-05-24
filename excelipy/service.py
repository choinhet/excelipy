import logging
from collections.abc import Callable, Sequence

import xlsxwriter

from excelipy.models import (
    Component,
    Excel,
    Fill,
    Group,
    Image,
    Link,
    Table,
    Text,
)
from excelipy.writers import (
    write_fill,
    write_image,
    write_link,
    write_table,
    write_text,
)

log = logging.getLogger("excelipy")

WRITING_MAP: dict[type[Component], Callable[..., tuple[int, int]]] = {
    Table: write_table,
    Text: write_text,
    Link: write_link,
    Fill: write_fill,
    Image: write_image,
}


def remove_groups(comp: Component) -> list[Component]:
    if not isinstance(comp, Group):
        return [comp]
    flattened_comps: list[Component] = []
    for c in comp.components:
        flattened_comps.extend(remove_groups(c))
    return flattened_comps


def unnest_components(components: Sequence[Component]) -> list[Component]:
    """
    Removes hierarchical groupings and flattens nested components into a single list.

    Args:
        components: A sequence of `Component` objects to be unnested.

    Returns:
        A flat list of all `Component` objects after removing groups and unnesting.
    """
    nested_comps = [remove_groups(c) for c in components]
    unnested_comps = [c for comps in nested_comps for c in comps]
    return unnested_comps


def save(excel: Excel):
    workbook_args = {
        "nan_inf_to_errors": excel.nan_inf_to_errors,
    }
    with xlsxwriter.Workbook(excel.path, workbook_args) as workbook:
        for sheet in excel.sheets:
            origin = (sheet.style.pl(), sheet.style.pt())
            worksheet = workbook.add_worksheet(sheet.name)

            if not sheet.grid_lines:
                worksheet.hide_gridlines(2)

            for component in unnest_components(sheet.components):
                cur_origin = (
                    origin[0] + component.style.pl(),
                    origin[1] + component.style.pt(),
                )
                _, y = WRITING_MAP[type(component)](
                    workbook,
                    worksheet,
                    component,
                    sheet.style,
                    cur_origin,
                )
                origin = (
                    origin[0] + component.style.pr(),
                    origin[1] + y + component.style.pb(),
                )
