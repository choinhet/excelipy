from __future__ import annotations

import io
from pathlib import Path
from typing import Any, Annotated, List
from typing import Dict, Optional, Sequence, Literal, Union, Callable

import pandas as pd
from pydantic import BaseModel, GetCoreSchemaHandler
from pydantic import Field, ConfigDict
from pydantic import GetJsonSchemaHandler
from pydantic.json_schema import JsonSchemaValue
from pydantic_core import core_schema


class Style(BaseModel):
    class Config:
        frozen = True

    align: Optional[
        Literal["left", "center", "right", "fill", "justify", "center_across", "distributed"]
    ] = Field(default=None)
    background: Optional[str] = Field(default=None)
    bold: Optional[bool] = Field(default=None)
    border: Optional[int] = Field(default=None)
    border_bottom: Optional[int] = Field(default=None)
    border_color: Optional[str] = Field(default=None)
    border_left: Optional[int] = Field(default=None)
    border_right: Optional[int] = Field(default=None)
    border_top: Optional[int] = Field(default=None)
    fill_inf: Optional[Union[str, int, float]] = Field(default=None)
    fill_na: Optional[Union[str, int, float]] = Field(default=None)
    fill_zero: Optional[str] = Field(default=None)
    font_color: Optional[str] = Field(default=None)
    font_family: Optional[str] = Field(default=None)
    font_size: Optional[int] = Field(default=None)
    numeric_format: Optional[str] = Field(default=None)
    padding: Optional[int] = Field(default=None)
    padding_bottom: Optional[int] = Field(default=None)
    padding_left: Optional[int] = Field(default=None)
    padding_right: Optional[int] = Field(default=None)
    padding_top: Optional[int] = Field(default=None)
    text_wrap: Optional[bool] = Field(default=None)
    underline: Optional[Literal[1, 2, 33, 34]] = Field(default=None)
    valign: Optional[
        Literal["top", "vcenter", "bottom", "vcenter", "bottom", "vjustify"]
    ] = Field(default=None)

    def merge(self, other: "Style") -> "Style":
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
    model_config = ConfigDict(arbitrary_types_allowed=True)

    @property
    def name(self) -> str:
        return self.type


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
            cls, source_type: Any, handler: GetCoreSchemaHandler
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
            cls, schema: core_schema.CoreSchema, handler: GetJsonSchemaHandler
    ) -> JsonSchemaValue:
        return {
            "type": "array",
            "items": {"type": "object"},
            "description": "DataFrame as array of records",
        }


class Table(BaseComponent):
    type: Literal["table"] = Field(default="table")
    data: Annotated[pd.DataFrame, DataFrameAsJsonLines]
    header_style: Dict[str, Style] = Field(default_factory=dict)
    body_style: Style = Field(default_factory=Style)
    column_style: Dict[str, Union[Style, Callable[[Any], Style]]] = Field(default_factory=dict)
    idx_column_style: Dict[int, Union[Style, Callable[[Any], Style]]] = Field(default_factory=dict)
    column_width: Dict[str, int] = Field(default_factory=dict)
    idx_column_width: Dict[int, int] = Field(default_factory=dict)
    row_style: Dict[int, Style] = Field(default_factory=dict)
    max_col_width: Optional[int] = Field(default=None)
    header_filters: bool = Field(default=True)
    default_style: bool = Field(default=True)
    auto_width_tuning: Optional[int] = Field(default=None)
    auto_width_padding: Optional[int] = Field(default=None)
    merge_equal_headers: bool = Field(default=True)
    wrap_header: bool = Field(default=False)
    max_col_size: Optional[int] = Field(default=None)

    def with_stripes(
            self,
            color: str = "#D0D0D0",
            pattern: Literal["even", "odd"] = "odd",
    ) -> "Table":
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


class Group(BaseModel):
    type: Literal["group"] = Field(default="group")
    name: str = Field(default="")
    components: Sequence["Component"] = Field(default_factory=list)
    model_config = ConfigDict(arbitrary_types_allowed=True)


Component = Annotated[
    Union[Text, Link, Fill, Image, Table, Group],
    Field(discriminator="type"),
]


class Sheet(BaseModel):
    name: str
    components: List[Component] = Field(default_factory=list)
    grid_lines: bool = Field(default=True)
    style: Style = Field(default_factory=Style)


class Excel(BaseModel):
    path: Union[Path, io.BytesIO]
    sheets: List[Sheet] = Field(default_factory=list)
    nan_inf_to_errors: bool = Field(default=True)
    model_config = ConfigDict(arbitrary_types_allowed=True)
