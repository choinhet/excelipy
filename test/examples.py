import logging
from importlib.resources import files
from pathlib import Path

import duckdb
import excel2img
import pandas as pd
from matplotlib import pyplot as plt
from matplotlib.colors import rgb2hex, to_rgb

import excelipy as ep
from test import resources

RESOURCES = Path(str(files(resources)))
log = logging.getLogger("excelipy")
EXAMPLES = []


def example(func):
    EXAMPLES.append(func)
    return func


def query(sql: str) -> pd.DataFrame:
    con = duckdb.connect()
    return con.query(sql).df()


def choose_font_color(background_color: str, threshold: int = 0.5) -> str:
    def get_luminance(rgb_color):
        r, g, b = rgb_color
        luminance = 0.299 * r + 0.587 * g + 0.114 * b
        return luminance

    rgb_tuple = to_rgb(background_color)
    luminance = get_luminance(rgb_tuple)
    if luminance < threshold:
        return "#ffffff"
    return "#000000"


@example
def displaying_a_table(out_path: Path):
    df = query(
        """
        select
            ProductName as Product,
            sum(Value) as Value
        from 'test/resources/samples/sales.parquet'
        group by 1
        order by 2 desc
    """
    )

    ep.save(
        excel=ep.Excel(
            path=out_path,
            sheets=[ep.Sheet(name="Sheet1", components=[ep.Table(data=df)])],
        )
    )


@example
def basic_column_formatting(out_path: Path):
    df = query(
        """
                      select
                          ProductName as Product,
                          sum(Value) as Value
                      from 'test/resources/samples/sales.parquet'
                      group by 1
                      order by 2 desc
                      """
    )

    ep.save(
        excel=ep.Excel(
            path=out_path,
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


@example
def adding_a_title(out_path: Path):
    df = query(
        """
               select
                   ProductName as Product,
                   sum(Value) as Value
               from 'test/resources/samples/sales.parquet'
               group by 1
               order by 2 desc
               """
    )

    num_cols = len(df.columns)
    ep.save(
        excel=ep.Excel(
            path=out_path,
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


@example
def category_coloring(out_path: Path):
    df = query(
        """
               select StoreName as Store,
                      ProductName as Product,
                      sum(Value) as Value
               from 'test/resources/samples/sales.parquet'
               group by 1, 2
               order by 3 desc
               """
    )

    unique_stores = list(df["Store"].unique())
    base_cmap = plt.get_cmap("tab20")
    cmap = [base_cmap(i % base_cmap.N) for i in range(len(unique_stores))]
    store_colors = dict(zip(unique_stores, [rgb2hex(rgba) for rgba in cmap]))

    def get_store_color(store: str) -> ep.Style:
        return ep.Style(
            background=store_colors[store],
            font_color=choose_font_color(store_colors[store]),
            bold=True,
        )

    num_cols = len(df.columns)
    ep.save(
        excel=ep.Excel(
            path=out_path,
            sheets=[
                ep.Sheet(
                    name="Sheet1",
                    components=[
                        ep.Text(
                            text="Sales by Product by Store",
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


@example
def merging_columns(out_path: Path):
    df = query(
        """
               select StoreName as Store,
                      ProductName as Product,
                      sum(Value) as Value
               from 'test/resources/samples/sales.parquet'
               group by 1, 2
               order by 3 desc
               """
    )

    unique_stores = list(df["Store"].unique())
    base_cmap = plt.get_cmap("tab20")
    cmap = [base_cmap(i % base_cmap.N) for i in range(len(unique_stores))]
    store_colors = dict(zip(unique_stores, [rgb2hex(rgba) for rgba in cmap]))

    def get_store_color(store: str) -> ep.Style:
        return ep.Style(
            background=store_colors[store],
            font_color=choose_font_color(store_colors[store]),
            bold=True,
        )

    unified = "Sales by Product by Store"
    df = df.rename(columns={col: unified for col in df.columns})

    ep.save(
        excel=ep.Excel(
            path=out_path,
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


@example
def conditional_formatting(out_path: Path):
    df = query(
        """
        select StoreName as Store,
               ProductName as Product,
               sum(Value) as Value
        from 'test/resources/samples/sales.parquet'
        group by 1, 2
        order by 3 desc
        """
    )
    avg_by_product = df.groupby("Product").agg({"Value": "mean"})["Value"].to_dict()

    unique_stores = list(df["Store"].unique())
    base_cmap = plt.get_cmap("tab20")
    cmap = [base_cmap(i % base_cmap.N) for i in range(len(unique_stores))]
    store_colors = dict(zip(unique_stores, [rgb2hex(rgba) for rgba in cmap]))

    def get_store_color(store: str) -> ep.Style:
        return ep.Style(
            background=store_colors[store],
            font_color=choose_font_color(store_colors[store]),
            bold=True,
        )

    @ep.row_wise
    def get_value_style(row) -> ep.Style:
        store, product, value = row
        prod_avg = avg_by_product[product]
        if value < prod_avg:
            return ep.Style(font_color="#ff0014", numeric_format=",.2f", bold=True)
        return ep.Style(numeric_format=",.2f", bold=True)

    unified = "Sales by Product by Store"
    num_cols = len(df.columns)
    df = df.rename(columns={col: unified for col in df.columns})

    ep.save(
        excel=ep.Excel(
            path=out_path,
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


def run_all(
        skip_existing: bool = True,
        out_path: Path = RESOURCES / "output",
):
    xlsx_out = out_path / "xlsx_output"
    img_out_path = out_path / "image_output"

    xlsx_out.mkdir(exist_ok=True, parents=True)
    img_out_path.mkdir(exist_ok=True, parents=True)

    for func in EXAMPLES:
        name = func.__name__

        cur_path = xlsx_out / f"{name}.xlsx"
        cur_img_path = img_out_path / f"{name}.png"

        if not skip_existing or not cur_path.exists():
            log.info(f"Running {name}")
            func(cur_path)

        if not skip_existing or not cur_img_path.exists():
            log.info(f"Taking screenshot of {name}")
            excel2img.export_img(cur_path, cur_img_path.as_posix())

def run_all(
    skip_existing: bool = True,
    out_path: Path = RESOURCES / "output",
):
    xlsx_out = out_path / "xlsx_output"
    img_out_path = out_path / "image_output"

    xlsx_out.mkdir(exist_ok=True, parents=True)
    img_out_path.mkdir(exist_ok=True, parents=True)

    for func in EXAMPLES:
        name = func.__name__

        cur_path = xlsx_out / f"{name}.xlsx"
        cur_img_path = img_out_path / f"{name}.png"

        if not skip_existing or not cur_path.exists():
            log.info(f"Running {name}")
            func(cur_path)

        if not skip_existing or not cur_img_path.exists():
            log.info(f"Taking screenshot of {name}")
            excel2img.export_img(cur_path, cur_img_path.as_posix())


if __name__ == "__main__":
    logging.basicConfig(level=logging.DEBUG)
    run_all()
