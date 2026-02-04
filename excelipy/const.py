import re

DATE_CONVERSION = {
    "%Y": "yyyy",
    "%y": "yy",
    "%m": "mm",
    "%d": "dd",
    "%H": "hh",
    "%M": "mm",
    "%S": "ss",
    "%b": "mmm",
    "%B": "mmmm",
}


def python_to_excel_fmt(fmt: str) -> str:
    fmt = (fmt or "General").strip()
    if not fmt or fmt == "General":
        return "General"

    # 1. Identify Python-specific patterns
    is_python_date = "%" in fmt and any(x in fmt for x in DATE_CONVERSION)

    # Python formats: .2f, ,.1f, .0%, .1%, %, f, d
    # Excel formats start with 0 or #: 0.0%, 0%, #,##0.00
    is_python_num = (
        bool(re.search(r"\.(\d+)[f%]|^[fd]$|^%$|^,\.\d+f$", fmt)) and not fmt[0] in "0#"
    )

    # 2. Handle Python Dates
    if is_python_date:
        for py, xl in DATE_CONVERSION.items():
            fmt = fmt.replace(py, xl)
        return fmt

    # 3. Handle Python Numerics
    if is_python_num:
        precision = re.search(r"\.(\d+)", fmt)
        decimals = int(precision.group(1)) if precision else 0
        decimal_part = f".{'0' * decimals}" if decimals > 0 else ""

        if "%" in fmt:
            return f"0{decimal_part}%"
        if "," in fmt:
            return f"#,##0{decimal_part}"
        return f"0{decimal_part}"  # Handles 'd' and 'f'

    # 4. If it's not a Python format, it's either already Excel or General
    return fmt


PROP_MAP = dict(
    align="align",
    valign="valign",
    font_size="font_size",
    font_color="font_color",
    font_family="font_name",
    bold="bold",
    border="border",
    border_left="left",
    border_right="right",
    border_top="top",
    border_bottom="bottom",
    border_color="border_color",
    background="bg_color",
    numeric_format="num_format",
    underline="underline",
)

PRE_PROCESS_MAP = dict(
    numeric_format=python_to_excel_fmt,
)
