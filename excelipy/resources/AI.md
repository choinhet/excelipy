# excelipy — AI Guide

This guide is for AI assistants generating Excel files via `excelipy` using **structured output** (JSON deserialized
into Pydantic models).

## Top-level structure

```json
{
  "name": "Sheet1",
  "components": [
    ...
  ],
  "grid_lines": true,
  "style": {}
}
```

| Field        | Type    | Default  | Description                                |
|--------------|---------|----------|--------------------------------------------|
| `name`       | string  | required | Sheet tab name                             |
| `components` | array   | `[]`     | Ordered list of components (top to bottom) |
| `grid_lines` | boolean | `true`   | Show/hide grid lines                       |
| `style`      | Style   | `{}`     | Default style applied to the whole sheet   |

---

## Component types

Each component is one of: `Text`, `Link`, `Fill`, `Image`, `Table`.

### Text

A single cell (or merged row of cells) with a text label.

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

| Field    | Type   | Default                              |
|----------|--------|--------------------------------------|
| `text`   | string | required                             |
| `width`  | int    | `1`                                  |
| `height` | int    | `1`                                  |
| `merged` | bool   | `true` — merges cells across `width` |
| `style`  | Style  | `{}`                                 |

### Link

A hyperlink cell.

```json
{
  "text": "Open Dashboard",
  "url": "https://example.com",
  "width": 2,
  "merged": true,
  "style": {}
}
```

### Fill

An empty spacer cell (used for visual padding between components).

```json
{
  "width": 3,
  "height": 1,
  "style": {
    "background": "#D0D0D0"
  }
}
```

### Table

The primary component. `data` is an **array of records** (list of objects).

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

| Field                 | Type             | Default  | Description                                                     |
|-----------------------|------------------|----------|-----------------------------------------------------------------|
| `data`                | array of objects | required | Each object is a row; keys become column headers                |
| `header_style`        | `{ col: Style }` | `{}`     | Style per column header cell                                    |
| `body_style`          | Style            | `{}`     | Applied to all body cells                                       |
| `column_style`        | `{ col: Style }` | `{}`     | Static style per column (**no callables in structured output**) |
| `idx_column_style`    | `{ int: Style }` | `{}`     | Style by column index (0-based)                                 |
| `column_width`        | `{ col: int }`   | `{}`     | Fixed width per column                                          |
| `idx_column_width`    | `{ int: int }`   | `{}`     | Fixed width by column index                                     |
| `row_style`           | `{ int: Style }` | `{}`     | Style by row index (0-based)                                    |
| `header_filters`      | bool             | `true`   | Show Excel autofilter dropdowns                                 |
| `default_style`       | bool             | `true`   | Apply excelipy default table styling                            |
| `max_col_width`       | int              | `null`   | Cap auto-detected column width                                  |
| `merge_equal_headers` | bool             | `true`   | Merge adjacent headers with the same name                       |

> **Note:** `column_style` and `idx_column_style` only support static `Style` objects in structured output. Callable (
> conditional) styles require Python code.

---

## Style object

All style fields are optional. Omit a field to leave it unset (inherits from parent or default).

```json
{
  "align": "left | center | right | fill | justify | center_across | distributed",
  "valign": "top | vcenter | bottom | vjustify",
  "background": "#RRGGBB",
  "font_color": "#RRGGBB",
  "font_family": "Arial",
  "font_size": 12,
  "bold": true,
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

**`numeric_format` examples:**

| Format string | Result                 |
|---------------|------------------------|
| `".0f"`       | `1234`                 |
| `".2f"`       | `1234.56`              |
| `",.2f"`      | `1,234.56`             |
| `",.1f"`      | `1,234.6`              |
| `".1%"`       | `12.3%`                |
| `"%d - %B"`   | `01 - January` (dates) |

**`underline` values:** `1` = single, `2` = double, `33` = single accounting, `34` = double accounting.

**`border` / `border_*` values:** 1–13 (xlsxwriter border style index). Use `1` for thin, `2` for medium, `5` for thick.

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
          "Value": 1200.5
        },
        {
          "Product": "Banana",
          "Value": 800.0
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
      "width": 2
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
