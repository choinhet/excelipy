import logging
from pathlib import Path

import numpy as np
import pandas as pd

import excelipy as ep


def simple_df() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "testing": [1, 2, 3],
            "tested": ["Yay", "Thanks", "Bud"],
        }
    )


def numeric_df() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "integers": [0, 2, 3],
            "invalid": [1, 2, 3],
            "floats": [1.2, 2.3, 3.1],
            "big_numbers": [100000000, 2001230, np.inf],
            "percents": [0.2129, np.nan, 1.11],
        }
    )


def df2() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "testing": [1, 2, 3],
            "tested": [
                "Yayyyyyyyyyyyyyyyyyyyyyyyyy this is a long phrase",
                "Thanks a lot",
                "Bud",
            ],
        }
    )


def duplicated_col_df() -> pd.DataFrame:
    title = "this is a long long long long long title"
    df = pd.DataFrame(
        {
            title: [
                "Yayyyyyyyyyyyyyyyyyyyyyyyyy this is a long phrase",
                "Thanks a lot",
                "Bud",
            ],
            # "bogus": [1, 2, 3],
            "testing": [1, 2, 3],
            "testing2": [10, 20, 30],
        }
    )
    return df.rename(columns={"testing": title, "testing2": title})


def simple_example():
    sheets = [
        ep.Sheet(
            name="Hello!",
            components=[
                ep.Text(text="Hello world!", width=2),
                ep.Fill(width=2, style=ep.Style(background="#33c481")),
                ep.Table(data=simple_df()),
            ],
            style=ep.Style(padding=1),
            grid_lines=False,
        ),
    ]

    excel = ep.Excel(
        path=Path("filename.xlsx"),
        sheets=sheets,
    )

    ep.save(excel)


def one_table():
    sheets = [
        ep.Sheet(
            name="Hello!",
            components=[
                ep.Table(data=simple_df())
            ],
        ),
    ]

    excel = ep.Excel(
        path=Path("filename.xlsx"),
        sheets=sheets,
    )

    ep.save(excel)


def two_tables():
    sheets = [
        ep.Sheet(
            name="Hello!",
            components=[
                ep.Table(
                    data=df2(),
                    style=ep.Style(padding_bottom=1, font_size=20)
                ),
                ep.Table(data=simple_df()),
            ],
        ),
        ep.Sheet(
            name="Hello again!",
            components=[
                ep.Table(data=simple_df(), style=ep.Style(padding_bottom=1)),
                ep.Table(data=simple_df()),
            ],
        ),
    ]

    excel = ep.Excel(
        path=Path("filename.xlsx"),
        sheets=sheets,
    )

    ep.save(excel)


def simple_image():
    sheets = [
        ep.Sheet(
            name="Hello!",
            components=[
                ep.Image(
                    path=Path("resources/img.png"),
                    width=2,
                    height=5,
                    style=ep.Style(border=2),
                ),
            ],
        ),
    ]

    excel = ep.Excel(
        path=Path("filename.xlsx"),
        sheets=sheets,
    )

    ep.save(excel)


def one_table_no_grid():
    sheets = [
        ep.Sheet(
            name="Hello!",
            components=[
                ep.Table(data=simple_df())
            ],
            grid_lines=False,
            style=ep.Style(padding=1),
        ),
    ]

    excel = ep.Excel(
        path=Path("filename.xlsx"),
        sheets=sheets,
    )

    ep.save(excel)


def default_text_style():
    sheets = [
        ep.Sheet(
            name="Hello!",
            components=[
                ep.Text(
                    text="Hello world! This text should be bigger than the table",
                    width=2,
                ),
                ep.Table(data=simple_df())
            ],
            grid_lines=False,
            style=ep.Style(padding=1),
        ),
    ]

    excel = ep.Excel(
        path=Path("filename.xlsx"),
        sheets=sheets,
    )

    ep.save(excel)


def merged_cols():
    df = duplicated_col_df()
    centered_style = {
        col: ep.Style(
            align="center",
            valign="vcenter"
        ) for col in df.columns
    }
    sheets = [
        ep.Sheet(
            name="Hello!",
            components=[
                ep.Table(
                    data=df,
                    header_style=centered_style,
                    column_style=centered_style,
                    header_filters=False,
                    # merge_equal_headers=False,
                )
            ],
            grid_lines=False,
            style=ep.Style(padding=1),
        ),
    ]

    excel = ep.Excel(
        path=Path("filename.xlsx"),
        sheets=sheets,
    )

    ep.save(excel)


def dataframe_formatting():
    df = numeric_df()
    formats = {
        "integers": ".0f",
        "floats": ".2f",
        "big_numbers": ",.1f",
        "percents": ".1%",
    }
    # for col, f in formats.items():
    #     df[col] = df[col].apply(lambda x: format(x, f))

    sheets = [
        ep.Sheet(
            name="Hello!",
            components=[
                ep.Table(
                    data=df,
                    default_style=False,
                    header_filters=False,
                    column_style={
                        col: ep.Style(
                            numeric_format=formats.get(col),
                            align="center",
                            fill_inf="-",
                            fill_na="-",
                            fill_zero="-",
                        )
                        for col in df.columns
                    }
                ),
            ],
        ),
    ]

    excel = ep.Excel(
        path=Path("filename.xlsx"),
        sheets=sheets,
    )

    ep.save(excel)


def merged_cols_formatting():
    df = duplicated_col_df()

    formatting = {
        1: ".2%",
        2: ".2f",
    }

    centered_style = {
        col: ep.Style(
            align="center",
            valign="vcenter",
            numeric_format=formatting.get(idx),
        ) for idx, col in enumerate(df.columns)
    }
    col_style = {
        idx: ep.Style(
            align="center",
            valign="vcenter",
            numeric_format=formatting.get(idx),
        ) for idx, col in enumerate(df.columns)
    }

    sheets = [
        ep.Sheet(
            name="Hello!",
            components=[
                ep.Table(
                    data=df,
                    header_style=centered_style,
                    idx_column_style=col_style,
                    header_filters=False,
                )
            ],
            grid_lines=False,
            style=ep.Style(padding=1),
        ),
    ]

    excel = ep.Excel(
        path=Path("filename.xlsx"),
        sheets=sheets,
    )

    ep.save(excel)


def large_width_no_merge():
    style = ep.Style(background="#33c481")
    sheets = [
        ep.Sheet(
            name="Hello!",
            components=[
                ep.Text(text="Hello", width=10, style=style),
                ep.Text(text="Hello", width=10, style=style, merged=False),
                ep.Fill(width=10, style=style),
                ep.Fill(width=10, style=style, merged=False),
            ],
            grid_lines=False,
        ),
    ]

    excel = ep.Excel(
        path=Path("filename.xlsx"),
        sheets=sheets,
    )

    ep.save(excel)


if __name__ == "__main__":
    logging.basicConfig(level=logging.DEBUG)
    large_width_no_merge()
