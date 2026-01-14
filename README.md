# Excelipy

[![codecov](https://codecov.io/gh/choinhet/excelipy/graph/badge.svg?token=${CODECOV_TOKEN})](https://codecov.io/gh/choinhet/excelipy)

## Installation

You can install the package using pip:

```bash
pip install excelipy
```

Or add it as a dependency in your `pyproject.toml`:

```bash
uv add excelipy
```

## Appeal

This was created with less than 100 lines of code, in a declarative, clean, and easy to understand style:

![demo_example.png](static/demo_example.png)

## Usage

### Detailed Model Overview

Excelipy is built on top of Pydantic models, making it easy to define your Excel structure programmatically.

#### `Excel`

The top-level container for your workbook.

- `path`: Path to the output file (string, `Path`, or `BytesIO`).
- `sheets`: A list of `Sheet` objects.
- `nan_inf_to_errors`: If `True` (default), converts `NaN` and `Inf` values to Excel errors.

#### `Sheet`

Represents a single worksheet.

- `name`: The name of the sheet.
- `components`: A list of components (Tables, Text, etc.) to be rendered sequentially.
- `grid_lines`: Boolean to show or hide grid lines.
- `style`: A `Style` object to apply to the entire sheet (e.g., padding).

#### `Table`

The main component for rendering data.

- `data`: A `pandas.DataFrame`.
- `header_style`: A dictionary mapping column names to `Style` objects.
- `body_style`: A default `Style` for the table body.
- `column_style`: Styles for specific columns. Can be a `Style` object or a function that returns a `Style` based on
  cell value.
- `row_style`: Styles for specific rows.
- `header_filters`: Boolean to enable/disable Excel header filters.
- `merge_equal_headers`: Boolean. If `True`, adjacent columns with the same name will have their headers merged.

#### `Style`

Defines how cells look. Supports most common Excel formatting:

- **Font**: `bold`, `font_color`, `font_family`, `font_size`, `underline`.
- **Alignment**: `align` (left, center, right, etc.), `valign` (top, vcenter, bottom), `text_wrap`.
- **Border**: `border` (sets all sides) or individual `border_top`, `border_bottom`, etc. `border_color`.
- **Background**: `background` (hex color).
- **Format**: `numeric_format` (e.g., `"0.00%"`, `"#,##0.00"`).

#### Other Components

- `Text`: A simple text cell or merged range. Supports `width`, `height`, and `style`.
- `Link`: A clickable hyperlink (This can be nested inside a `Table` component).
- `Image`: Embeds an image from a `path`.
- `Fill`: Empty space component to push other components down or across.

### Examples

### Displaying a table

The most basic usage is to display a `pandas.DataFrame` as a table.

```python
import excelipy as ep

df = ...  # your pandas DataFrame

ep.save(
    excel=ep.Excel(
        path="output.xlsx",
        sheets=[ep.Sheet(name="Sheet1", components=[ep.Table(data=df)])],
    )
)
```

By default, `ep.Table` renders a `pandas.DataFrame` with standard formatting, including header filters and automatic
column width adjustment.

![displaying_a_table.png](test/resources/output/image_output/displaying_a_table.png)

### Basic column formatting

You can apply styles to headers and specific columns. Here, we use `header_style` to bold and center the headers, and
`column_style` to format the "Value" column with two decimal places.

```python
ep.save(
    excel=ep.Excel(
        path="output.xlsx",
        sheets=[
            ep.Sheet(
                name="Sheet1",
                components=[
                    ep.Table(
                        data=df,
                        header_style={
                            col: ep.Style(
                                bold=True, align="center", valign="vcenter"
                            )
                            for col in df.columns
                        },
                        column_style={"Value": ep.Style(numeric_format=",.2f")},
                    )
                ],
            )
        ],
    )
)
```

The output shows bold, centered headers and a formatted numeric column.

![basic_column_formatting.png](test/resources/output/image_output/basic_column_formatting.png)

### Adding a title

You can add text components above or around your tables. In this example, an `ep.Text` component is added before the
`ep.Table`. We set its `width` to match the number of columns in the table and apply a background color and centered
alignment to create a unified title bar.

```python
num_cols = len(df.columns)
ep.save(
    excel=ep.Excel(
        path="output.xlsx",
        sheets=[
            ep.Sheet(
                name="Sheet1",
                components=[
                    ep.Text(
                        text="Sales by Product",
                        width=num_cols,
                        style=ep.Style(
                            bold=True,
                            background="#ecedef",
                            align="center",
                            valign="vcenter",
                        ),
                    ),
                    ep.Table(
                        data=df,
                        header_style={
                            col: ep.Style(
                                bold=True, align="center", valign="vcenter"
                            )
                            for col in df.columns
                        },
                        column_style={"Value": ep.Style(numeric_format=",.2f")},
                    ),
                ],
            )
        ],
    )
)
```

The resulting Excel sheet features a stylized title spanning across the top of the table.

![adding_a_title.png](test/resources/output/image_output/adding_a_title.png)

### Category coloring

Styles can be applied dynamically based on the cell content. By passing a function to `column_style`, we can change the
background and font color of the "Store" column depending on its value, making different categories easy to distinguish.

```python
def get_store_color(store: str) -> ep.Style:
    return ep.Style(
        background=store_colors[store],
        font_color=choose_font_color(store_colors[store]),
        bold=True,
    )


ep.save(
    excel=ep.Excel(
        path="output.xlsx",
        sheets=[
            ep.Sheet(
                name="Sheet1",
                components=[
                    ep.Table(
                        data=df,
                        column_style={
                            "Value": ep.Style(numeric_format=",.2f"),
                            "Store": get_store_color,
                        },
                    ),
                ],
            )
        ],
    )
)
```

Each store now has its own unique color coding in the "Store" column.

![category_coloring.png](test/resources/output/image_output/category_coloring.png)

### Merging columns

By giving multiple columns the same name and setting `merge_equal_headers=True` (which is the default), they will be
merged in the header. In this example, all columns are renamed to the same title, creating a single merged header over
the entire table. We also use `idx_column_style` to apply styles by column index instead of name.

```python
unified = "Sales by Product by Store"
df = df.rename(columns={col: unified for col in df.columns})

ep.save(
    excel=ep.Excel(
        path="output.xlsx",
        sheets=[
            ep.Sheet(
                name="Sheet1",
                components=[
                    ep.Table(
                        data=df,
                        header_style={
                            col: ep.Style(
                                bold=True, align="center", valign="vcenter"
                            )
                            for col in df.columns
                        },
                        body_style=ep.Style(align="center", valign="vcenter"),
                        idx_column_style={
                            0: get_store_color,
                            2: ep.Style(numeric_format=",.2f"),
                        },
                        header_filters=False,
                    )
                ],
            )
        ],
    )
)
```

The header columns are merged into one, and the body cells are centered.

![merging_columns.png](test/resources/output/image_output/merging_columns.png)

### Conditional formatting

You can also apply row-wise conditional formatting. Using the `@ep.row_wise` decorator on a styling function allows you
to access the entire row's data. Here, we highlight "Value" in red if it falls below the average for that product. We
also hide `grid_lines` and add `padding` to the sheet for a cleaner look.

```python
@ep.row_wise
def get_value_style(row) -> ep.Style:
    store, product, value = row
    prod_avg = avg_by_product[product]
    if value < prod_avg:
        return ep.Style(font_color="#ff0014", numeric_format=",.2f", bold=True)
    return ep.Style(numeric_format=",.2f", bold=True)


ep.save(
    excel=ep.Excel(
        path="output.xlsx",
        sheets=[
            ep.Sheet(
                name="Sheet1",
                components=[
                    ep.Table(
                        data=df,
                        idx_column_style={
                            0: get_store_color,
                            2: get_value_style,
                        },
                        header_filters=False,
                    ),
                    ep.Fill(width=num_cols),
                    ep.Text(
                        text="Products that sold below average are highlighted in red",
                        style=ep.Style(
                            bold=True,
                            valign="vcenter",
                            align="center",
                            border=3,
                            border_color="#ff0014",
                        ),
                        width=num_cols,
                        height=3,
                    ),
                ],
                grid_lines=False,
                style=ep.Style(padding=2),
            )
        ],
    )
)
```

The final output features conditional red text for below-average sales, a descriptive footer box, and no visible grid
lines.

![conditional_formatting.png](test/resources/output/image_output/conditional_formatting.png)


