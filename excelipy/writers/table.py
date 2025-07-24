import logging
from typing import Tuple, Dict, Optional

from PIL import ImageFont
from xlsxwriter.workbook import Workbook, Worksheet

from excelipy.models import Style, Table
from excelipy.style import process_style
from excelipy.styles.table import DEFAULT_HEADER_STYLE, DEFAULT_BODY_STYLE

log = logging.getLogger("excelipy")

DEFAULT_FONT_SIZE = 11


def get_text_size(
        text: str,
        font_size: Optional[int] = None,
        font_family: Optional[str] = None,
):
    cur_font_size = font_size or DEFAULT_FONT_SIZE
    cur_font_family = font_family or "Arial"

    try:
        cur_font = ImageFont.truetype(
            f"{cur_font_family}.ttf".lower(),
            cur_font_size
        )
    except Exception as e:
        cur_font = ImageFont.load_default()
        log.debug(
            f"Could not load custom font {cur_font_family}, using default. Exception: {e}"
        )

    return cur_font.getlength(text)


def get_style_font_family(*styles: Style) -> Optional[str]:
    cur_font = None
    for s in filter(None, styles):
        cur_font = s.font_family or cur_font
    return cur_font


def get_style_font_size(*styles: Style) -> Optional[int]:
    cur_font = None
    for s in filter(None, styles):
        cur_font = s.font_size or cur_font
    return cur_font


def get_auto_width(
        header: str,
        component: Table,
        default_style: Style,
) -> float:
    header_font_size = get_style_font_size(
        DEFAULT_HEADER_STYLE,
        component.style,
        component.header_style.get(header),
    )
    header_font_family = get_style_font_family(
        DEFAULT_HEADER_STYLE,
        component.style,
        component.header_style.get(header),
    )
    header_len = get_text_size(
        header,
        header_font_size,
        header_font_family,
    )
    col_font_size = get_style_font_size(
        default_style,
        component.style,
        component.body_style,
        component.column_style.get(header),
    )
    col_font_family = get_style_font_family(
        default_style,
        component.style,
        component.body_style,
        component.column_style.get(header),
    )
    all_col_len = component.data[header].apply(str).apply(
        lambda it: get_text_size(
            it,
            col_font_size,
            col_font_family,
        )
    )

    col_len = all_col_len.max()
    max_len = max(header_len, col_len)
    result = max_len // component.auto_width_tuning + component.auto_width_padding
    return result


col_size_cache_by_sheet: Dict[str, Dict[int, int]] = dict()


def write_table(
        workbook: Workbook,
        worksheet: Worksheet,
        component: Table,
        default_style: Style,
        origin: Tuple[int, int] = (0, 0),
) -> Tuple[int, int]:
    x_size = component.data.shape[1]
    y_size = component.data.shape[0] + 1  # +1 for header row

    global col_size_cache_by_sheet

    for col_idx, header in enumerate(component.data.columns):
        current_header_style = component.header_style.get(
            header,
            component.style
        )
        header_styles = [default_style, current_header_style]

        if component.default_style:
            header_styles = [DEFAULT_HEADER_STYLE] + header_styles

        header_format = process_style(workbook, header_styles)
        worksheet.write(
            origin[1],
            origin[0] + col_idx,
            header,
            header_format,
        )
        set_width = component.column_width.get(header)
        if set_width:
            estimated_width = set_width
        else:
            estimated_width = get_auto_width(
                header,
                component,
                default_style
            )
            col_size_cache = col_size_cache_by_sheet.get(worksheet.name, {})
            cur_cached = col_size_cache.get(origin[0] + col_idx, 0)
            estimated_width = max((cur_cached, estimated_width))
            col_size_cache[origin[0] + col_idx] = estimated_width
            col_size_cache_by_sheet[worksheet.name] = col_size_cache

        log.debug(
            f"Estimated width for {header}: {estimated_width} [Sheet: {worksheet.name}]"
        )
        worksheet.set_column(
            origin[0] + col_idx,
            origin[0] + col_idx,
            int(estimated_width)
        )

    if component.header_filters:
        worksheet.autofilter(
            origin[1],
            origin[0],
            origin[1],
            origin[0] + len(list(component.data.columns)) - 1,
        )

    for col_idx, col in enumerate(component.data.columns):
        col_style = component.column_style.get(col)
        for row_idx, (_, row) in enumerate(component.data.iterrows()):
            row_style = component.row_style.get(row_idx)
            body_style = [
                default_style,
                component.style,
                component.body_style,
                col_style,
                row_style,
            ]

            if component.default_style:
                body_style = [DEFAULT_BODY_STYLE] + body_style

            current_format = process_style(workbook, body_style)
            cell = row[col]
            worksheet.write(
                origin[1] + row_idx + 1,
                origin[0] + col_idx,
                cell,
                current_format,
            )

    return x_size, y_size
