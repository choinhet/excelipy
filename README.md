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
It also has auto-width detection that uses PIL under the hood :)

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
        style=ep.Style(padding=1),
        grid_lines=False,
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
Auto-scale, based on PIL image size.

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

## Extra

There are a lot of options regarding how to style your sheets.
The tables come with a default style, which you can deactivate through the `default_style` flag.

The majority of the other styles will be added through the Style class. 
Everything should be well typed, so the autocomplete should help with the options available.