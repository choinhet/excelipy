import logging
from pathlib import Path

import pandas as pd

import excelipy as ep


def df() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "testing": [1, 2, 3],
            "tested": ["Yay", "Thanks", "Bud"],
        }
    )


def one_table():
    sheets = [
        ep.Sheet(
            name="Hello!",
            components=[
                ep.Table(data=df())
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
                ep.Table(data=df(), style=ep.Style(padding_bottom=1)),
                ep.Table(data=df()),
            ],
        ),
    ]

    excel = ep.Excel(
        path=Path("filename.xlsx"),
        sheets=sheets,
    )

    ep.save(excel)

if __name__ == "__main__":
    logging.basicConfig(level=logging.DEBUG)
    two_tables()
