from typing import Sequence

from xlsxwriter.workbook import Format, Workbook

from excelipy.const import PROP_MAP, PRE_PROCESS_MAP
from excelipy.models import Style


def _process_single(workbook: Workbook, style: Style) -> Format:
    style_dict = style.model_dump(exclude_none=True)
    style_map = {}
    for prop, value in style_dict.items():
        if (mapped_prop := PROP_MAP.get(prop)) is not None:
            if prop in PRE_PROCESS_MAP:
                value = PRE_PROCESS_MAP[prop](value)
            style_map[mapped_prop] = value
    return workbook.add_format(style_map)


def process_style(
        workbook: Workbook,
        styles: Sequence[Style],
) -> Format:
    styles = list(filter(None, styles))
    cur_style = Style()
    for style in styles:
        cur_style = cur_style.merge(style)

    cached_styles = getattr(workbook, "_excelipy_format_cache", None)
    if cached_styles is None:
        cached_styles = {}
        setattr(workbook, "_excelipy_format_cache", cached_styles)

    if cur_style in cached_styles:
        return cached_styles[cur_style]

    result = _process_single(workbook, cur_style)
    cached_styles[cur_style] = result
    return result
