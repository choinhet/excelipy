from typing import Dict, Sequence

from xlsxwriter.workbook import Format, Workbook

from excelipy.const import PROP_MAP
from excelipy.models import Style

cached_styles: Dict[Style, Format] = {}


def _process_single(workbook: Workbook, style: Style) -> Format:
    style_dict = style.model_dump(exclude_none=True)
    style_map = {
        mapped_prop: value
        for property, value in style_dict.items()
        if (mapped_prop := PROP_MAP.get(property)) is not None
    }
    format = workbook.add_format(style_map)
    cached_styles[style] = format
    return format


def process_style(
    workbook: Workbook,
    styles: Sequence[Style],
) -> Format:
    cur_style = Style()
    for style in styles:
        cur_style = cur_style.merge(style)
    if cur_style in cached_styles:
        return cached_styles[cur_style]
    format = _process_single(workbook, cur_style)
    return format
