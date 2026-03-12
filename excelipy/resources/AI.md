# excelipy — AI Agent Guide

This guide instructs AI agents on how to generate valid Excel files via `excelipy` using **structured JSON output**
deserialized into Pydantic models.

---

## Choosing a component — decision rules

Before defining any component, apply these rules in order:

1. Displaying a title, label, or note? → **`Text`**
2. Displaying row/column data? → **`Table`**
3. Displaying a clickable URL? → **`Link`**
4. Need blank vertical space between two other components? → **`Fill`** *(only valid use)*
5. Unsure? → Use `Text` or `Table`. **Never default to `Fill`.**

> **Omission beats `Fill`**: if content is unknown or not yet available, omit the component entirely.
> Do not insert `Fill` as a placeholder.

---

## ❌ Anti-patterns — never do these

| Wrong                                    | Right                                 |
|------------------------------------------|---------------------------------------|
| Use `Fill` when content is unknown       | Omit the component entirely           |
| Use `Fill` as a title or label cell      | Use `Text`                            |
| Use `Fill` for decorative colored blocks | Use `Text` with a background style    |
| Render `Table` with `data: []`           | Omit the `Table` if there are no rows |
| Use `Text` for row data                  | Use `Table`                           |

---

## Top-level structure

```json
{
  "name": "Sheet1",
  "components": [],
  "grid_lines": true,
  "style": {}
}
```

| Field        | Type    | Default  | Description                                                                                                                                   |
|--------------|---------|----------|-----------------------------------------------------------------------------------------------------------------------------------------------|
| `name`       | string  | required | Sheet tab name                                                                                                                                |
| `components` | array   | `[]`     | Ordered list of components (top to bottom). Every component must carry real content — do not insert `Fill` unless explicit spacing is needed. |
| `grid_lines` | boolean | `true`   | Show/hide grid lines                                                                                                                          |
| `style`      | Style   | `{}`     | Default style applied to the whole sheet                                                                                                      |

---

## Component types

### Text

A single cell (or merged row of cells) with a text label.

**Use for:** titles, section headers, footnotes, annotations, disclaimers.

```json
{
  "text": "Sales Report",
  "width": 3,
  "height": 1,
  "merged": true,
  "style": {
    "bold": true,
    "align": "center",
    "background": "#ecedef"
  }
}
```

| Field    | Type   | Default  | Notes                       |
|----------|--------|----------|-----------------------------|
| `text`   | string | required | The label to display        |
| `width`  | int    | `1`      | Number of columns to span   |
| `height` | int    | `1`      | Number of rows to span      |
| `merged` | bool   | `true`   | Merges cells across `width` |
| `style`  | Style  | `{}`     | See Style reference below   |

---

### Link

A clickable hyperlink cell.

**Use for:** URLs, dashboards, external references.

```json
{
  "text": "Open Dashboard",
  "url": "https://example.com",
  "width": 2,
  "merged": true,
  "style": {}
}
```

---

### Table

The primary data component. `data` is an **array of records** (list of objects). Each object is one row; keys become
column headers.

**Use for:** any structured, row-based data.

> ⚠️ If `data` would be empty (`[]`), **omit the Table entirely**.

```json
{
  "data": [
    {
      "Product": "Apple",
      "Value": 1200.5
    },
    {
      "Product": "Banana",
      "Value": 800.0
    }
  ],
  "header_style": {
    "Product": {
      "bold": true,
      "align": "center"
    },
    "Value": {
      "bold": true,
      "align": "center"
    }
  },
  "body_style": {
    "valign": "vcenter"
  },
  "column_style": {
    "Value": {
      "numeric_format": ",.2f",
      "align": "right"
    }
  },
  "column_width": {
    "Product": 20
  },
  "row_style": {
    "0": {
      "background": "#f0f0f0"
    }
  },
  "header_filters": true,
  "default_style": true,
  "merge_equal_headers": true
}
```

| Field                 | Type             | Default  | Description                                                   |
|-----------------------|------------------|----------|---------------------------------------------------------------|
| `data`                | array of objects | required | Rows of data — **must be non-empty**                          |
| `header_style`        | `{ col: Style }` | `{}`     | Style per column header cell                                  |
| `body_style`          | Style            | `{}`     | Applied to all body cells                                     |
| `column_style`        | `{ col: Style }` | `{}`     | Static style per column *(no callables in structured output)* |
| `idx_column_style`    | `{ int: Style }` | `{}`     | Style by column index (0-based)                               |
| `column_width`        | `{ col: int }`   | `{}`     | Fixed width per column                                        |
| `idx_column_width`    | `{ int: int }`   | `{}`     | Fixed width by column index                                   |
| `row_style`           | `{ int: Style }` | `{}`     | Style by row index (0-based)                                  |
| `header_filters`      | bool             | `true`   | Show Excel autofilter dropdowns                               |
| `default_style`       | bool             | `true`   | Apply excelipy default table styling                          |
| `max_col_width`       | int              | `null`   | Cap auto-detected column width                                |
| `merge_equal_headers` | bool             | `true`   | Merge adjacent headers with the same name                     |

> **Note:** `column_style` and `idx_column_style` only support static `Style` objects in structured output.
> Callable (conditional) styles require Python code.

---

### Fill ⚠️ — spacer only

An empty cell used **exclusively** to add blank vertical space between two other components.

**Valid use:** inserting whitespace between a title and a table, or between two tables.  
**Invalid use:** placeholder for unknown content, decorative color blocks, labels, or any data.

```json
{
  "height": 1
}
```

| Field    | Type  | Default | Notes                              |
|----------|-------|---------|------------------------------------|
| `width`  | int   | `1`     | Columns to span                    |
| `height` | int   | `1`     | Rows of blank space                |
| `style`  | Style | `{}`    | Avoid styling — it signals content |

> If you're tempted to add a `background` or `text` to a `Fill`, use `Text` instead.

---

## Style object

All fields are optional. Omit a field to inherit from the parent or sheet default.

```json
{
  "align": "left | center | right | fill | justify | center_across | distributed",
  "valign": "top | vcenter | bottom | vjustify",
  "background": "#RRGGBB",
  "font_color": "#RRGGBB",
  "font_family": "Arial",
  "font_size": 12,
  "bold": true,
  "italic": false,
  "text_wrap": false,
  "underline": 1,
  "border": 1,
  "border_color": "#RRGGBB",
  "border_top": 1,
  "border_bottom": 1,
  "border_left": 1,
  "border_right": 1,
  "numeric_format": ",.2f",
  "padding": 1,
  "padding_top": 1,
  "padding_bottom": 1,
  "padding_left": 1,
  "padding_right": 1,
  "fill_na": "-",
  "fill_inf": "-",
  "fill_zero": "-"
}
```

### `numeric_format` reference

| Format string | Example output |
|---------------|----------------|
| `".0f"`       | `1234`         |
| `".2f"`       | `1234.56`      |
| `",.2f"`      | `1,234.56`     |
| `",.1f"`      | `1,234.6`      |
| `".1%"`       | `12.3%`        |
| `"%d - %B"`   | `01 - January` |

### `underline` values

| Value | Meaning           |
|-------|-------------------|
| `1`   | Single            |
| `2`   | Double            |
| `33`  | Single accounting |
| `34`  | Double accounting |

### `border` / `border_*` values

Use xlsxwriter border style index: `1` = thin, `2` = medium, `5` = thick. Range: 1–13.

---

## Complete example

```json
{
  "name": "Sales",
  "grid_lines": false,
  "style": {
    "font_size": 11,
    "font_family": "Calibri",
    "padding": 1
  },
  "components": [
    {
      "text": "Sales by Product",
      "width": 2,
      "style": {
        "bold": true,
        "background": "#ecedef",
        "align": "center",
        "valign": "vcenter"
      }
    },
    {
      "data": [
        {
          "Product": "Apple",
          "Value": 1200.50
        },
        {
          "Product": "Banana",
          "Value": 800.00
        },
        {
          "Product": "Cherry",
          "Value": 350.75
        }
      ],
      "header_style": {
        "Product": {
          "bold": true,
          "align": "center"
        },
        "Value": {
          "bold": true,
          "align": "center"
        }
      },
      "column_style": {
        "Value": {
          "numeric_format": ",.2f",
          "align": "right"
        }
      },
      "column_width": {
        "Product": 18
      }
    },
    {
      "height": 1
    },
    {
      "text": "All values in USD",
      "width": 2,
      "style": {
        "italic": true,
        "font_color": "#888888",
        "align": "right"
      }
    }
  ]
}
```

> Note: the `Fill` spacer above (`{ "height": 1 }`) is the correct, minimal form — no background, no style, no content.
