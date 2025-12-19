def python_to_excel_fmt(fmt: str) -> str:
    fmt = fmt.strip()
    if fmt.endswith("%"):
        decimals = int(fmt.strip("%").lstrip(".") or 0)
        return f"0.{decimals * '0'}%" if decimals > 0 else "0%"
    if "," in fmt and "f" in fmt:
        decimals = int(fmt.split(".")[-1].rstrip("f") or 0)
        return f"#,##0.{decimals * '0'}" if decimals > 0 else "#,##0"
    if "f" in fmt:
        decimals = int(fmt.split(".")[-1].rstrip("f") or 0)
        return f"0.{decimals * '0'}" if decimals > 0 else "0"
    return "General"


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
