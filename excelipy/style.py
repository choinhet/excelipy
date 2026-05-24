from collections.abc import Collection, Sequence
from functools import lru_cache

from xlsxwriter.workbook import Format, Workbook

from excelipy.const import PRE_PROCESS_MAP, PROP_MAP
from excelipy.models import Style


def convert_style_to_format(workbook: Workbook, style: Style) -> Format:
    style_dict = style.model_dump(exclude_none=True)
    style_map = {}
    for prop, value in style_dict.items():
        if (mapped_prop := PROP_MAP.get(prop)) is not None:
            if prop in PRE_PROCESS_MAP:
                value = PRE_PROCESS_MAP[prop](value)
            style_map[mapped_prop] = value
    return workbook.add_format(style_map)


@lru_cache
def merge_styles(*styles: Style | None) -> Style:
    """
    Merge multiple styles into one prioritizing the last style provided.

    Args:
        *styles: Styles to be merged

    Returns:
        Merged style

    >>> result = merge_styles(None, Style(font_size=12), None, Style(font_size=11), None)
    >>> result.font_size
    11
    """
    _styles = list(filter(None, styles))
    cur_style = Style()
    for style in _styles:
        cur_style = cur_style.merge(style)
    return cur_style


def process_style(
    workbook: Workbook,
    styles: Collection[Style | None],
) -> Format:
    cur_style = merge_styles(*styles)
    cached_formats = getattr(workbook, "_excelipy_format_cache", None)
    if cached_formats is None:
        cached_formats = {}
        setattr(workbook, "_excelipy_format_cache", cached_formats)

    if cur_style in cached_formats:
        return cached_formats[cur_style]

    result = convert_style_to_format(workbook, cur_style)
    cached_formats[cur_style] = result
    return result
