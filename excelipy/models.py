import io
from collections.abc import Callable, Sequence
from pathlib import Path
from typing import Annotated, Any, Literal

import pandas as pd
from pydantic import (
    BaseModel,
    ConfigDict,
    Field,
    GetCoreSchemaHandler,
    GetJsonSchemaHandler,
    model_validator,
)
from pydantic.json_schema import JsonSchemaValue
from pydantic_core import core_schema
from typing_extensions import Self

AlignOptions = Literal[
    "left",
    "center",
    "right",
    "fill",
    "justify",
    "center_across",
    "distributed",
]
VAlignOptions = Literal[
    "top",
    "vcenter",
    "bottom",
    "vcenter",
    "bottom",
    "vjustify",
]


class Style(BaseModel):
    align: AlignOptions | None = Field(default=None)
    background: str | None = Field(default=None)
    bold: bool | None = Field(default=None)
    border: int | None = Field(default=None)
    border_bottom: int | None = Field(default=None)
    border_color: str | None = Field(default=None)
    border_left: int | None = Field(default=None)
    border_right: int | None = Field(default=None)
    border_top: int | None = Field(default=None)
    fill_inf: str | int | float | None = Field(default=None)
    fill_na: str | int | float | None = Field(default=None)
    fill_zero: str | None = Field(default=None)
    font_color: str | None = Field(default=None)
    font_family: str | None = Field(default=None)
    font_size: int | None = Field(default=None)
    numeric_format: str | None = Field(default=None)
    padding: int | None = Field(default=None)
    padding_bottom: int | None = Field(default=None)
    padding_left: int | None = Field(default=None)
    padding_right: int | None = Field(default=None)
    padding_top: int | None = Field(default=None)
    text_wrap: bool | None = Field(default=None)
    underline: Literal[1, 2, 33, 34] | None = Field(default=None)
    valign: VAlignOptions | None = Field(default=None)

    model_config = ConfigDict(frozen=True)

    def __str__(self) -> str:
        """
        Returns:
            Examples:
            >>> str(Style(font_size=14))
            "{'font_size': 14}"
        """
        return str(self.model_dump(exclude_defaults=True))

    def merge(self, other: Self) -> Self:
        self_dict = self.model_dump(exclude_none=True)
        other_dict = other.model_dump(exclude_none=True)
        self_dict.update(other_dict)
        return self.model_validate(self_dict)

    def pl(self) -> int:
        return self.padding_left or self.padding or 0

    def pt(self) -> int:
        return self.padding_top or self.padding or 0

    def pr(self) -> int:
        return self.padding_right or self.padding or 0

    def pb(self) -> int:
        return self.padding_bottom or self.padding or 0


class BaseComponent(BaseModel):
    type: Literal["base"] = Field(default="base")
    style: Style = Field(default_factory=Style)
    model_config = ConfigDict(arbitrary_types_allowed=True, frozen=True)
    name: str = Field(default="")

    @model_validator(mode="before")
    @classmethod
    def auto_set_name(cls, d: dict[str, Any]) -> dict[str, Any]:
        d["name"] = d.get("name", cls.__name__.lower())
        return d


class Text(BaseComponent):
    type: Literal["text"] = Field(default="text")
    text: str
    width: int = Field(default=1)
    height: int = Field(default=1)
    merged: bool = Field(default=True)


class Link(BaseComponent):
    type: Literal["link"] = Field(default="link")
    text: str
    url: str
    width: int = Field(default=1)
    height: int = Field(default=1)
    merged: bool = Field(default=True)

    def __str__(self):
        """
        Examples:
            >>> str(Link(text="example", url="https://example.com"))
            'example'
        """
        return self.text


class Fill(BaseComponent):
    type: Literal["fill"] = "fill"
    width: int = Field(default=1)
    height: int = Field(default=1)
    merged: bool = Field(default=True)


class Image(BaseComponent):
    type: Literal["image"] = Field(default="image")
    path: Path
    width: int = Field(default=1)
    height: int = Field(default=1)


class DataFrameAsJsonLines(pd.DataFrame):
    """A pandas DataFrame subclass that Pydantic can serialize/deserialize as JSON Lines."""

    @classmethod
    def _validate(cls, value: Any) -> pd.DataFrame:
        if isinstance(value, pd.DataFrame):
            return value
        if isinstance(value, str):
            return pd.read_json(io.StringIO(value), lines=True)
        if isinstance(value, list):
            return pd.DataFrame(value)
        raise ValueError(f"Cannot convert {type(value)} to DataFrame")

    @classmethod
    def _serialize(cls, df: pd.DataFrame) -> list:
        return df.to_dict(orient="records")

    @classmethod
    def __get_pydantic_core_schema__(
        cls,
        source_type: Any,
        handler: GetCoreSchemaHandler,
    ) -> core_schema.CoreSchema:
        return core_schema.no_info_plain_validator_function(
            cls._validate,
            serialization=core_schema.plain_serializer_function_ser_schema(
                cls._serialize,
                info_arg=False,
                return_schema=core_schema.list_schema(core_schema.dict_schema()),
            ),
        )

    @classmethod
    def __get_pydantic_json_schema__(
        cls,
        schema: core_schema.CoreSchema,
        handler: GetJsonSchemaHandler,
    ) -> JsonSchemaValue:
        return {
            "type": "array",
            "items": {"type": "object"},
            "description": "DataFrame as array of records",
        }


StyleFunc = Callable[[Any], Style]


class Table(BaseComponent):
    type: Literal["table"] = Field(default="table")
    data: Annotated[pd.DataFrame, DataFrameAsJsonLines]
    auto_size: bool = Field(default=True)
    auto_width_padding: int | None = Field(default=None)
    auto_width_tuning: int | None = Field(default=None)
    body_style: Style = Field(default_factory=Style)
    column_style: dict[str, Style | StyleFunc] = Field(default_factory=dict)
    column_width: dict[str, int] = Field(default_factory=dict)
    default_style: bool = Field(default=True)
    header_filters: bool = Field(default=True)
    header_style: dict[str, Style] = Field(default_factory=dict)
    idx_column_style: dict[int, Style | StyleFunc] = Field(default_factory=dict)
    idx_column_width: dict[int, int] = Field(default_factory=dict)
    max_col_size: int | None = Field(default=None)
    max_col_width: int | None = Field(default=None)
    merge_equal_headers: bool = Field(default=True)
    min_col_size: int | None = Field(default=None)
    row_style: dict[int, Style] = Field(default_factory=dict)
    wrap_header: bool = Field(default=False)

    def with_stripes(
        self,
        color: str = "#D0D0D0",
        pattern: Literal["even", "odd"] = "odd",
    ) -> Self:
        return self.model_copy(
            update=dict(
                row_style={
                    idx: (
                        self.row_style.get(idx, Style()).merge(Style(background=color))
                        if (pattern == "odd" and idx % 2 != 0)
                        or (pattern == "even" and idx % 2 == 0)
                        else self.row_style.get(idx, Style())
                    )
                    for idx in range(self.data.shape[0])
                }
            )
        )


class Group(BaseComponent):
    type: Literal["group"] = Field(default="group")
    components: Sequence["Component"] = Field(default_factory=list)

    model_config = ConfigDict(arbitrary_types_allowed=True)


Component = Annotated[
    Text | Link | Fill | Image | Table | Group,
    Field(discriminator="type"),
]


class Sheet(BaseModel):
    name: str
    components: list[Component] = Field(default_factory=list)
    grid_lines: bool = Field(default=True)
    style: Style = Field(default_factory=Style)


class Excel(BaseModel):
    path: Path | io.BytesIO
    sheets: list[Sheet] = Field(default_factory=list)
    nan_inf_to_errors: bool = Field(default=True)
    model_config = ConfigDict(arbitrary_types_allowed=True)
