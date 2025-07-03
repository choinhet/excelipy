# Excelipy

[![codecov](https://codecov.io/gh/choinhet/excelipy/graph/badge.svg?token=${CODECOV_TOKEN})](https://codecov.io/gh/choinhet/excelipy)

## Installation

You can install the package using pip:

```bash
pip install excelipy
```

## Usage

The idea for this package is for it to be a declarative way of using the
xlsxwritter.
It allows you to define Excel files using Python objects, which can be more
intuitive and easier to manage than writing raw Excel files.

## Simple Example

```python
import excelipy as ep

sheets = [
    ep.Sheet(
        name="Hello!",
        components=[
            ep.Text(text="Hello world!", width=2),
            ep.Fill(width=2, style=ep.Style(background="#33c481")),
            ep.Table(data=df),
        ],
        style=ep.Style(padding=1)
    ),
]

excel = ep.Excel(
    path=Path("filename.xlsx"),
    sheets=sheets,
)

ep.save(excel)
```

Result:

![simple_example.png](static/simple_example.png)

## Working with images

You can also add images to your Excel sheets.
Auto-

```python
import excelipy as ep

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
```

Result:

![image_example.png](static/image_example.png)

## Advanced Example

```python
import excelipy as ep

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
```

This is an exaggerated example, to show the capabilities of the package. You can
see that for a table, you can define the header style, body style, column
styles, row styles, and even the column widths. You can also add text and fill
components to the sheet.

This is the result:

![advanced_example.png](static/advanced_example.png)