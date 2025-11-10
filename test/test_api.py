import importlib.resources as pkg_resources
from pathlib import Path

import pandas as pd
import pytest

import excelipy as ep
from test import resources


@pytest.fixture
def resources_path() -> Path:
    return Path(str(pkg_resources.files(resources)))


@pytest.fixture
def img_path(resources_path: Path) -> Path:
    return resources_path / "img.png"


@pytest.fixture
def sample_df() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "testing": [1, 2, 3],
            "tested": ["Yay", "Thanks", "Bud"],
        }
    )


@pytest.fixture
def numeric_df() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "integers": [1, 2, 3],
            "invalid": [1, 2, 3],
            "floats": [1.2, 2.3, 3.1],
            "big_numbers": [100000000, 2001230, 120392222],
            "percents": [0.2129, 0.522, 1.11],
        }
    )


@pytest.fixture
def empty_df() -> pd.DataFrame:
    num_rows = 2
    num_cols = 9
    cols = [" "] * num_cols
    return pd.DataFrame(" ", index=range(num_rows), columns=cols)


def test_api(
        sample_df: pd.DataFrame,
        empty_df: pd.DataFrame,
        img_path: Path,
        numeric_df: pd.DataFrame,
):
    numeric_formats = {
        "integers": ".0f",
        "floats": ".2f",
        "big_numbers": ",.1f",
        "percents": ".1%",
        "invalid": "invalid",
    }
    sheets = [
        ep.Sheet(
            name="Hello!",
            components=[
                ep.Text(
                    text="This is my table",
                    style=ep.Style(bold=True),
                    width=4,
                ),
                ep.Fill(),
                ep.Text(text="Monowidth"),
                ep.Table(data=empty_df),
                ep.Fill(
                    width=4,
                    style=ep.Style(background="#D0D0D0"),
                ),
                ep.Table(
                    data=sample_df,
                    header_style={
                        col: ep.Style(
                            bold=True,
                            border=5,
                            border_color="#F02932",
                        ) for col in sample_df.columns
                    },
                    body_style=ep.Style(font_size=18),
                    column_style={
                        "testing": ep.Style(
                            font_size=10,
                            align="center",
                            numeric_format=".0f",
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
                ep.Image(
                    path=img_path,
                    width=2,
                    height=5,
                    style=ep.Style(border=2),
                ),
                ep.Table(
                    data=numeric_df,
                    default_style=False,
                    header_filters=False,
                    column_style={
                        col: ep.Style(numeric_format=numeric_formats.get(col))
                        for col in numeric_df.columns
                    }
                ),
            ],
            style=ep.Style(
                font_size=14,
                font_family="Times New Roman",
                padding=1,
            ),
            grid_lines=False,
        ),
    ]

    temp_dir = Path("temp")
    temp_dir.mkdir(exist_ok=True, parents=True)
    temp_path = temp_dir / "filename.xlsx"

    excel = ep.Excel(
        path=temp_path,
        sheets=sheets,
    )

    ep.save(excel)

    assert temp_path.exists(), "Excel file was not created"
    assert temp_path.is_file(), "Path is not a file"
    assert temp_path.stat().st_size > 0, "Excel file is empty"
    temp_path.unlink(missing_ok=True)
    temp_dir.rmdir()


if __name__ == "__main__":
    pytest.main([__file__])
