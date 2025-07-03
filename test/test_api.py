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


def test_api(sample_df: pd.DataFrame, img_path: Path):
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
                    data=sample_df,
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
                ep.Image(
                    path=img_path,
                    width=2,
                    height=5,
                    style=ep.Style(border=2),
                ),
            ],
            style=ep.Style(
                font_size=14,
                font_family="Times New Roman",
                padding=1,
            ),
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
