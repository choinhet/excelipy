import datetime
import logging
import math
from collections import defaultdict
from functools import lru_cache
from typing import Any, NamedTuple, cast

import numpy as np
import pandas as pd
from PIL import ImageFont
from xlsxwriter.workbook import Workbook, Worksheet

from excelipy.models import Link, Style, StyleFunc, Table
from excelipy.style import merge_styles, process_style
from excelipy.styles.table import DEFAULT_BODY_STYLE, DEFAULT_HEADER_STYLE

log = logging.getLogger("excelipy")

DEFAULT_FONT_SIZE = 11
DEFAULT_LINE_SPACING = 1.4
DEFAULT_ROW_HEIGHT = 15.0
DEFAULT_FONT_FAMILY = "Calibri"

PADDING_DEFAULT = 2

# Excel draws at 96 dpi, where a point is 96/72 of a pixel. Pillow sizes a font
# in pixels, so a font asked for at its point size comes out a quarter too
# small and every string measures short. The size is left fractional: rounding
# 11pt up to a 15px em reads every string 2% wide.
PX_PER_POINT = 96 / 72

# Fonts are measured at this multiple of their real size and scaled back down.
# A hinted font rounds its advances to whole pixels at text sizes, by a
# different amount on every platform and in every font; measured large, the
# rounding is a rounding error and what comes back is the advance the glyph was
# drawn with - which is what Excel lays text out on.
MEASURE_SCALE = 10

# Excel sizes columns in units of the font's widest digit and keeps 5 pixels of
# every cell for padding and the gridline, plus room for a filter button in a
# header that carries one.
EXCEL_PADDING_PX = 5
FILTER_BUTTON_PX = 16

# Excel lays text out on whole pixels, so a line that overruns by less than one
# is drawn, not wrapped.
FIT_TOLERANCE_PX = 1.0

# Characters a line may end on, dash kept on the line it ends
BREAK_AFTER_CHARS = "-\u2013\u2014"

ROW_WISE_ARG = "_excelipy_row_wise"
COL_CACHE_NAME = "_excelipy_col_sizes"
ROW_CACHE_NAME = "_excelipy_row_heights"


def row_wise(func):
    """
    Marks a StyleFunc (Callable[[...], ep.Style]) to receive all columns instead of only the current row value

    Callable[[Any (data type)], ep.Style] -> Callable[[pd.Series], ep.Style]
    """
    setattr(func, ROW_WISE_ARG, True)
    return func


def _static_col_style(component: Table, col_name: str, col_idx: int) -> Style:
    idx_style = component.idx_column_style.get(col_idx)
    col_style = component.column_style.get(col_name)
    maybe = idx_style or col_style
    return Style() if callable(maybe) or maybe is None else maybe


# Files a font is installed as where that is not simply its name: Windows
# abbreviates its own, and a machine without the font itself is likely to carry
# the metric-compatible substitute, which measures the same by design.
FONT_FILE_ALIASES = {
    "arial": ("arial.ttf", "LiberationSans-Regular.ttf"),
    "book antiqua": ("bkant.ttf",),
    "calibri": ("calibri.ttf", "Carlito-Regular.ttf"),
    "cambria": ("cambria.ttf", "Caladea-Regular.ttf"),
    "century gothic": ("gothic.ttf",),
    "comic sans ms": ("comic.ttf",),
    "consolas": ("consola.ttf",),
    "courier new": ("cour.ttf", "LiberationMono-Regular.ttf"),
    "garamond": ("gara.ttf",),
    "lucida console": ("lucon.ttf",),
    "palatino linotype": ("pala.ttf",),
    "segoe ui": ("segoeui.ttf",),
    "times new roman": ("times.ttf", "LiberationSerif-Regular.ttf"),
    "trebuchet ms": ("trebuc.ttf",),
}


def _font_candidates(font_family: str) -> tuple[str, ...]:
    """
    File names a font family is likely installed under, best guess first.

    Examples:
        >>> _font_candidates("Verdana")
        ('verdana.ttf', 'Verdana.ttf')
        >>> "times.ttf" in _font_candidates("Times New Roman")
        True
    """
    lower = font_family.lower()
    packed = font_family.replace(" ", "")
    names = (lower, lower.replace(" ", ""), font_family, packed)
    seen = {f"{name}.ttf": None for name in names}
    for alias in FONT_FILE_ALIASES.get(lower, ()):
        seen[alias] = None
    return tuple(seen)


@lru_cache
def _load_font(
    font_family: str,
    font_size: int,
) -> ImageFont.ImageFont | ImageFont.FreeTypeFont:
    size_px = font_size * PX_PER_POINT * MEASURE_SCALE
    for candidate in _font_candidates(font_family):
        try:
            return ImageFont.truetype(candidate, size_px)
        except Exception as e:
            log.debug(f"Could not load font file {candidate}.\nException: {e}")
    log.debug(f"Could not load custom font {font_family}, using default.")
    try:
        # Sized, so metrics stay proportional to the font: the unsized default
        # font would measure every size like the default one.
        return ImageFont.load_default(size=size_px)
    except TypeError:  # pragma: no cover - Pillow < 10.1
        return ImageFont.load_default()


def _max_digit_px() -> int:
    """
    Width of the digit zero in the workbook's default font, in whole pixels.

    This is the unit Excel measures columns in, so it comes from the default
    font whatever font a cell itself uses, and it is a whole number of pixels:
    Excel's column arithmetic is done in pixels of a rendered digit. For the
    Calibri 11 this writes by default, that is the documented 7.

    The fraction is dropped rather than rounded, because that is what a digit
    renders to: Calibri 11 measures anywhere from 7.4 to 7.9 depending on the
    font file and the machine, and Excel draws all of them 7 wide. Rounding
    one of those up to 8 hands every column a seventh more room than it has,
    and the text Excel wraps into it then fits.

    Examples:
        >>> _max_digit_px() >= 1
        True
    """
    return max(
        math.trunc(get_char_size("0", DEFAULT_FONT_SIZE, DEFAULT_FONT_FAMILY)), 1
    )


def _column_px(size: float) -> float:
    """
    The pixels Excel gives a column of ``size`` units, by its own conversion.

    Examples:
        >>> _column_px(0)
        0.0
    """
    digit = _max_digit_px()
    return float(math.trunc(((256 * size + math.trunc(128 / digit)) / 256) * digit))


def _px_to_excel(px: float) -> int:
    """
    Column units needed to show ``px`` pixels of text, Excel's own formula
    plus :data:`PADDING_DEFAULT` units of breathing room.
    """
    return math.ceil((px + EXCEL_PADDING_PX) / _max_digit_px()) + PADDING_DEFAULT


def _raw_px(text: str, font_size: int, font_family: str) -> float:
    """Width of ``text`` as laid out at :data:`MEASURE_SCALE` times its size."""
    return _load_font(font_family, font_size).getlength(text)


@lru_cache
def get_char_size(
    char: str,
    font_size: int,
    font_family: str,
) -> int | float:
    return _raw_px(char, font_size, font_family) / MEASURE_SCALE


@lru_cache(maxsize=1 << 16)
def _kern_px(pair: str, font_size: int, font_family: str) -> float:
    """
    What laying out two characters together takes off their two advances.

    A line is exactly its characters' advances plus the kerning between each
    neighbouring pair, so caching the pairs - of which text uses few - measures
    a line exactly without laying one out, which costs several times as much.
    """
    joined = _raw_px(pair, font_size, font_family)
    apart = _raw_px(pair[0], font_size, font_family) + _raw_px(
        pair[1], font_size, font_family
    )
    return (joined - apart) / MEASURE_SCALE


@lru_cache(maxsize=1 << 16)
def _line_px(line: str, font_size: int, font_family: str) -> float:
    """
    Width of one line with no breaks in it.

    Cached because the same text is measured over and over: the words wrapping
    walks through, and the repeated values a column of a table tends to hold.
    """
    total = 0.0
    previous = ""
    for char in line:
        total += get_char_size(char, font_size, font_family)
        if previous:
            total += _kern_px(previous + char, font_size, font_family)
        previous = char
    return total


def _append_px(
    tail: str,
    addition: str,
    font_size: int,
    font_family: str,
) -> float:
    """
    What ``addition`` adds to a line ending in ``tail``.

    Adding a piece at a time and keeping the running width is what makes
    wrapping cost one pass over the text rather than one over every line.
    """
    if not addition:
        return 0.0
    kerned = _kern_px(tail[-1] + addition[0], font_size, font_family) if tail else 0.0
    return kerned + _line_px(addition, font_size, font_family)


def get_text_px(
    text: str,
    font_size: int | None = None,
    font_family: str | None = None,
) -> float:
    """
    Width of the widest line of ``text``, in pixels.

    A line is its characters' advances plus the kerning between neighbouring
    pairs, both cached: dropping the kerning reads several pixels over a line
    of text - enough to wrap one that fits - and laying the line out instead
    costs several times as much for the same answer.

    Examples:
        >>> get_text_px("wide line\\nshort") == get_text_px("wide line")
        True
        >>> get_text_px("")
        0.0
    """
    family = font_family or DEFAULT_FONT_FAMILY
    size = font_size or DEFAULT_FONT_SIZE
    return max(_line_px(line, size, family) for line in str(text).split("\n"))


def get_text_size(
    text: str,
    font_size: int | None = None,
    font_family: str | None = None,
) -> int:
    return _px_to_excel(get_text_px(text, font_size, font_family))


def _excel_to_px(size: float, filtered: bool = False) -> float:
    """
    How many pixels of text a column (or merged span) of ``size`` fits.

    A column is a fixed width in pixels - ``size`` digits of the default font,
    less the padding Excel keeps and the filter button if the cell carries one.
    A cell in a larger font fits less of it, which is why text is measured in
    its own font and the column in the default one.

    The pixels come from Excel's own conversion, which works in whole digits.
    Reading a column back through a fractional digit makes it some 6% wider
    than Excel draws it, and text Excel wraps then fits.

    Examples:
        >>> _excel_to_px(_px_to_excel(83.0)) >= 83.0
        True
        >>> _excel_to_px(20, filtered=True) < _excel_to_px(20)
        True
    """
    taken = EXCEL_PADDING_PX + (FILTER_BUTTON_PX if filtered else 0)
    return max(_column_px(size) - taken, 0.0)


def _break_chunks(word: str) -> list[str]:
    """
    The pieces a word can be broken into, each keeping the dash it ends on.

    Excel treats a dash as a place a line may end, so a hyphenated token is
    laid out a chunk at a time rather than cut wherever the room runs out.

    Examples:
        >>> _break_chunks("one-unbroken-token")
        ['one-', 'unbroken-', 'token']
        >>> _break_chunks("plain")
        ['plain']
    """
    chunks: list[str] = []
    current = ""
    for char in word:
        current += char
        if char in BREAK_AFTER_CHARS:
            chunks.append(current)
            current = ""
    if current:
        chunks.append(current)
    return chunks or [""]


def _longest_prefix(
    word: str,
    available_px: float,
    font_size: int,
    font_family: str,
) -> int:
    """
    How much of ``word`` fits on an empty line, at least one character.

    Examples:
        >>> _longest_prefix("xxxxxxxx", 0, 11, "Calibri")
        1
        >>> _longest_prefix("fits", 10_000, 11, "Calibri")
        4
    """
    used = 0.0
    tail = ""
    for taken, char in enumerate(word):
        used += _append_px(tail, char, font_size, font_family)
        if taken and used > available_px + FIT_TOLERANCE_PX:
            return taken
        tail = char
    return max(len(word), 1)


def _count_paragraph_lines(
    paragraph: str,
    available_px: float,
    font_size: int,
    font_family: str,
) -> int:
    """Lines one run of text without newlines takes, wrapped Excel's way."""
    lines = 1
    used = 0.0
    tail = ""
    for word in paragraph.split(" "):
        for idx, chunk in enumerate(_break_chunks(word)):
            # A word starts after a space; a chunk of one carries straight on
            piece = f" {chunk}" if used and idx == 0 else chunk
            width = _append_px(tail, piece, font_size, font_family)
            if used + width <= available_px + FIT_TOLERANCE_PX:
                used += width
                tail = piece[-1:] or tail
                continue
            if used:
                lines += 1
                used, tail = 0.0, ""
                width = _line_px(chunk, font_size, font_family)
            if width <= available_px + FIT_TOLERANCE_PX:
                used, tail = width, chunk[-1:]
                continue
            # Wider than a whole line on its own: Excel breaks it mid-chunk
            rest = chunk
            while True:
                width = _line_px(rest, font_size, font_family)
                if width <= available_px + FIT_TOLERANCE_PX:
                    break
                rest = rest[
                    _longest_prefix(rest, available_px, font_size, font_family) :
                ]
                lines += 1
            used, tail = width, rest[-1:]
    return lines


def count_lines(
    text: str,
    available_px: float,
    font_size: int | None = None,
    font_family: str | None = None,
) -> int:
    """
    Lines Excel needs to draw ``text`` wrapped inside ``available_px`` pixels.

    Mirrors how Excel wraps: explicit newlines always break, words break on
    spaces, and a word only breaks mid-word when it is wider than the whole
    line on its own. Each candidate line is measured whole, since a line is
    narrower than its characters measured one by one.

    Examples:
        >>> count_lines("a b", 1000)
        1
        >>> count_lines("one two three", 0)
        1
        >>> count_lines("hello\\nworld", 1000)
        2
    """
    text = str(text)
    if available_px <= 0:
        return 1
    family = font_family or DEFAULT_FONT_FAMILY
    size = font_size or DEFAULT_FONT_SIZE
    return sum(
        _count_paragraph_lines(paragraph, available_px, size, family)
        for paragraph in text.split("\n")
    )


def get_row_height(lines: int, font_size: int | None) -> float:
    return max(
        DEFAULT_ROW_HEIGHT,
        (font_size or DEFAULT_FONT_SIZE) * DEFAULT_LINE_SPACING * lines,
    )


def _maybe_format(text: float | int | str, num_format: str | None) -> str:
    """
    Examples:
        >>> _maybe_format(1.2321, None)
        '1.2321'
        >>> _maybe_format(0.1 + 0.2, None)
        '0.3'
        >>> _maybe_format(1.2321, ",.2f")
        '1.23'
        >>> _maybe_format(1.2321, ",d")
        '1'
        >>> _maybe_format(20000, ",d")
        '20,000'
        >>> _maybe_format("text", ".2f")
        'text'
    """
    if num_format is None:
        if isinstance(text, float):
            # Excel's "General" shows ~11 significant digits, not float repr
            return f"{text:.11g}"
        return str(text)
    clz = int
    if "." in str(text) or "f" in num_format:
        clz = float
    if "d" in num_format:
        clz = int
    try:
        return format(clz(text), num_format)
    except Exception:
        return str(text)


def _display_text(value: Any, style: Style) -> str:
    """
    The text Excel actually renders for ``value``, used for every measurement.

    Measuring the raw value instead of the rendered one is what makes a
    formatted column look like it overflows: a ratio written as
    ``0.0714285714285714`` is three times wider than the ``7.14%`` the sheet
    shows, so the column is sized (and the row wrapped) for text nobody sees.

    Examples:
        >>> _display_text(0.0714285714285714, Style(numeric_format=".2%"))
        '7.14%'
        >>> _display_text(None, Style())
        ''
        >>> _display_text(float("nan"), Style())
        ''
        >>> _display_text(pd.NaT, Style(numeric_format="%Y"))
        ''
        >>> _display_text(True, Style())
        'TRUE'
        >>> _display_text(datetime.date(2026, 1, 23), Style(numeric_format="%d - %B"))
        '23 - January'
        >>> _display_text("plain", Style())
        'plain'
    """
    if value is None or value is pd.NaT or value is pd.NA:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    if isinstance(value, (datetime.datetime, datetime.date, pd.Timestamp)):
        return _display_date(value, style.numeric_format)
    if isinstance(value, bool):
        return str(value).upper()
    return _maybe_format(value, style.numeric_format)


def _display_date(
    value: datetime.datetime | datetime.date,
    num_format: str | None,
) -> str:
    """
    Rendered width of a date cell.

    Python formats are applied directly. An Excel pattern is used as its own
    proxy, since ``dd/mm/yyyy`` is as wide as the date it renders.

    Examples:
        >>> _display_date(datetime.date(2026, 1, 23), None)
        '2026-01-23'
        >>> _display_date(datetime.date(2026, 1, 23), "dd/mm/yyyy")
        'dd/mm/yyyy'
    """
    if num_format:
        if "%" in num_format:
            try:
                return value.strftime(num_format)
            except (ValueError, TypeError):  # pragma: no cover - platform specific
                return str(value)
        return num_format
    return (
        value.isoformat(sep=" ")
        if isinstance(value, datetime.datetime)
        else value.isoformat()
    )


class _Measure(NamedTuple):
    """A measured cell: what it shows, how wide that is, and how it is drawn."""

    text: str
    size: int
    font_size: int | None
    font_family: str | None
    wraps: bool


def _fit_row(
    worksheet: Worksheet,
    row: int,
    measure: _Measure,
    available_size: int,
    filtered: bool = False,
) -> None:
    """
    Set the height a wrapped cell needs, growing the row but never shrinking it.

    Every wrapped row is given an explicit height, one line included. Left
    without one, Excel picks the height itself, and on text that nearly fills
    its column it reserves a second line it then draws nothing on - the empty
    band under a single line of text.

    Heights are kept per worksheet, so a later table (or a narrower column
    further along the row) cannot take away a height another cell needed.
    """
    if not measure.wraps:
        # Drawn on one line whatever it holds, so the row is not ours to size
        return
    lines = 1
    if "\n" in measure.text or measure.size > available_size:
        lines = count_lines(
            measure.text,
            _excel_to_px(available_size, filtered),
            measure.font_size,
            measure.font_family,
        )
    height = get_row_height(lines, measure.font_size)
    heights = getattr(worksheet, ROW_CACHE_NAME, None)
    if heights is None:
        heights = {}
        setattr(worksheet, ROW_CACHE_NAME, heights)
    if height > heights.get(row, 0.0):
        heights[row] = height
        worksheet.set_row(row, height)


def write_table(
    workbook: Workbook,
    worksheet: Worksheet,
    component: Table,
    default_style: Style,
    origin: tuple[int, int] = (0, 0),
) -> tuple[int, int]:
    """
    Examples:
        >>> n = 30_000
        >>> data = pd.DataFrame({"A": [1, 2, 3] * n, "B": [4, "ha" * 50, 6] * n, "C": [4, 5, 6] * n})
        >>> long_text = "This is an avocado toast" * 3
        >>> data.rename(columns={"A": long_text, "B": long_text}, inplace=True)
        >>> origin = (0, 0)
        >>> default_style = Style(align="center", valign="vcenter")
        >>> row_style = {1: Style(font_size=14)}
        >>> component = Table(data=data, row_style=row_style, min_col_size=10, max_col_size=20, wrap_header=True)
        >>> import xlsxwriter
        >>> workbook = xlsxwriter.Workbook("output.xlsx")
        >>> worksheet = workbook.add_worksheet()
        >>> _ = write_table(workbook, worksheet, component, default_style)
        >>> workbook.close()
    """
    x_size = component.data.shape[1]
    y_size = component.data.shape[0] + 1

    df_columns = list(component.data.columns)
    df_rows = component.data.values.tolist()

    header_size_cache: dict[int, _Measure] = {}
    body_size_cache: dict[int, dict[int, _Measure]] = defaultdict(dict)
    biggest_body: dict[int, int] = defaultdict(lambda: 0)

    base_column_range = [(idx, idx) for idx in range(len(df_columns))]
    column_ranges = list(base_column_range)
    # =============================== Write headers ================================
    prev = None
    prev_format = None
    min_idx = 0
    for col_idx, cur_col in enumerate(df_columns):
        header_style = merge_styles(
            DEFAULT_HEADER_STYLE if component.default_style else None,
            default_style,
            component.style,
            component.header_style.get(cur_col),
            Style(text_wrap=True) if component.wrap_header else None,
        )
        header_format = process_style(workbook, [header_style])
        worksheet.write(origin[1], origin[0] + col_idx, cur_col, header_format)
        if component.merge_equal_headers:
            last_idx = len(df_columns) - 1
            group_ends = prev is not None and (cur_col != prev or col_idx == last_idx)
            end_idx = (
                col_idx if (col_idx == last_idx and cur_col == prev) else col_idx - 1
            )
            if group_ends:
                if end_idx > min_idx:
                    worksheet.merge_range(
                        first_row=origin[1],
                        first_col=origin[0] + min_idx,
                        last_row=origin[1],
                        last_col=origin[0] + end_idx,
                        data=prev,
                        cell_format=prev_format,
                    )
                    for _idx in range(min_idx, end_idx + 1):
                        column_ranges.remove((_idx, _idx))
                    column_ranges.append((min_idx, end_idx))
                    column_ranges.sort(key=lambda x: x[0])
                min_idx = col_idx
            prev = cur_col
            prev_format = header_format
        if component.auto_size:
            header_size_cache[col_idx] = _Measure(
                text=str(cur_col),
                size=get_text_size(
                    cur_col,
                    header_style.font_size,
                    header_style.font_family,
                ),
                font_size=header_style.font_size,
                font_family=header_style.font_family,
                wraps=bool(header_style.text_wrap),
            )

    # =============================== Header filters ===============================
    actually_merged = set(base_column_range) != set(column_ranges)
    if component.header_filters and not actually_merged:
        worksheet.autofilter(
            origin[1],
            origin[0],
            origin[1],
            origin[0] + len(list(component.data.columns)) - 1,
        )

    # ================================= Write body =================================
    for col_idx, col in enumerate(df_columns):
        base_style = merge_styles(
            DEFAULT_BODY_STYLE if component.default_style else None,
            default_style,
            component.style,
            component.body_style,
            _static_col_style(component, col, col_idx),
            Style(text_wrap=True) if component.wrap_header else None,
        )
        _maybe = Style | StyleFunc | None
        maybe_func_col_style: _maybe = component.column_style.get(col)
        maybe_func_idx_col_style: _maybe = component.idx_column_style.get(col_idx)
        maybe_func_style = maybe_func_idx_col_style or maybe_func_col_style
        style_func: StyleFunc | None = None
        if callable(maybe_func_style):
            style_func: StyleFunc = cast(StyleFunc, maybe_func_style)
        for row_idx, row in enumerate(df_rows):
            cell = row[col_idx]
            row_style = component.row_style.get(row_idx)
            merged_style = (
                base_style.merge(row_style) if row_style is not None else base_style
            )
            url = None
            if isinstance(cell, Link):
                url = cell.url
                cell = cell.text
            if merged_style.fill_na is not None and pd.isna(cell):
                cell = merged_style.fill_na
                merged_style = merged_style.model_copy(update=dict(numeric_format=None))
            if merged_style.fill_zero is not None and cell == 0:
                cell = merged_style.fill_zero
                merged_style = merged_style.model_copy(update=dict(numeric_format=None))
            if merged_style.fill_inf is not None and cell in (np.inf, -np.inf):
                cell = merged_style.fill_inf
                merged_style = merged_style.model_copy(update=dict(numeric_format=None))
            if style_func:
                dyn_style = (
                    style_func(row)
                    if getattr(style_func, ROW_WISE_ARG, False)
                    else style_func(cell)
                )
                merged_style = merged_style.merge(dyn_style)
            if (row_style := component.row_style.get(row_idx)) is not None:
                merged_style = merged_style.merge(row_style)
            current_format = process_style(workbook, [merged_style])

            if component.auto_size:
                shown = _display_text(cell, merged_style)
                cur_txt_size = get_text_size(
                    shown,
                    merged_style.font_size,
                    merged_style.font_family,
                )
                biggest_body[col_idx] = max(cur_txt_size, biggest_body[col_idx])
                if merged_style.text_wrap:
                    # Only wrapped cells can grow a row
                    body_size_cache[col_idx][row_idx] = _Measure(
                        text=shown,
                        size=cur_txt_size,
                        font_size=merged_style.font_size,
                        font_family=merged_style.font_family,
                        wraps=True,
                    )

            if url is None:
                worksheet.write(
                    origin[1] + row_idx + 1,
                    origin[0] + col_idx,
                    cell,
                    current_format,
                )
            else:
                worksheet.write_url(
                    origin[1] + row_idx + 1,
                    origin[0] + col_idx,
                    url,
                    current_format,
                    cell,
                )

    # =============================== Auto Set Width ===============================
    if component.auto_size:
        col_sizes = getattr(worksheet, COL_CACHE_NAME, None) or defaultdict(lambda: 0)
        # Compare cache to body
        for col_idx, text_size in biggest_body.items():
            col_sizes[origin[0] + col_idx] = max(
                text_size,
                col_sizes[origin[0] + col_idx],
            )
        # Compare cache to header (considering merged spans)
        for beg, end in column_ranges:
            text_size = header_size_cache[beg].size
            cur_body_sizes = [
                col_sizes[origin[0] + col_idx] for col_idx in range(beg, end + 1)
            ]
            num_cols = end - beg + 1
            total_size = sum(cur_body_sizes)
            diff = text_size - total_size
            if diff > 0:
                to_increase = diff // num_cols
                for col_idx in range(beg, end + 1):
                    col_sizes[origin[0] + col_idx] += to_increase
        # Hard set sizes
        for col, width in component.column_width.items():
            idxs = [i for i, c in enumerate(df_columns) if col == c]
            for idx in idxs:
                col_sizes[origin[0] + idx] = width
        # apply constraints
        for sheet_idx, text_size in col_sizes.items():
            if component.min_col_size and text_size < component.min_col_size:
                text_size = component.min_col_size
            if component.max_col_size and text_size > component.max_col_size:
                text_size = component.max_col_size
            col_sizes[sheet_idx] = text_size
            worksheet.set_column(sheet_idx, sheet_idx, col_sizes[sheet_idx])
        setattr(worksheet, COL_CACHE_NAME, col_sizes)
        # row wrap headers
        for beg, end in column_ranges:
            measure = header_size_cache[beg]
            span_size = sum(
                col_sizes[origin[0] + col_idx] for col_idx in range(beg, end + 1)
            )
            _fit_row(
                worksheet,
                origin[1],
                measure,
                span_size,
                filtered=component.header_filters and not actually_merged,
            )
        # row wrap body
        for col, rows in body_size_cache.items():
            col_size = col_sizes[origin[0] + col]
            for row_idx, measure in rows.items():
                _fit_row(worksheet, origin[1] + row_idx + 1, measure, col_size)

    return x_size, y_size
