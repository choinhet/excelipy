import logging
from pathlib import Path

import pandas as pd

import excelipy as ep


def main():
    df = pd.DataFrame(
        {
            "testing": [1, 2, 3],
            "tested": ["Yay", "Thanks", "Bud"],
        }
    )

    sheets = [
        ep.Sheet(
            name="Hello!",
            components=[
                ep.Text(
                    text="This is my table",
                    style=ep.Style(bold=True),
                    width=4,
                ),
                ep.Fill(
                    width=4,
                    style=ep.Style(background="#D0D0D0"),
                ),
                ep.Table(
                    data=df,
                    header_style=ep.Style(
                        bold=True,
                        border=5,
                        border_color="#F02932",
                    ),
                    body_style=ep.Style(font_size=18),
                    column_style={
                        "testing": ep.Style(
                            font_size=10,
                            align="center",
                        ),
                    },
                    column_width={
                        "tested": 20,
                    },
                    row_style={
                        1: ep.Style(
                            border=2,
                            border_color="#F02932",
                        )
                    },
                    style=ep.Style(padding=1),
                ).with_stripes(pattern="even"),
            ],
            style=ep.Style(
                font_size=14,
                font_family="Times New Roman",
                padding=1,
            ),
        ),
    ]

    excel = ep.Excel(
        path=Path("filename.xlsx"),
        sheets=sheets,
    )

    ep.save(excel)


if __name__ == "__main__":
    logging.basicConfig(level=logging.DEBUG)
    main()
