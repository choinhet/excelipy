import io
from pathlib import Path
from typing import Dict, Optional, Sequence, Literal, Any, Union, Callable

import pandas as pd
from pydantic import BaseModel, Field


class Style(BaseModel):
    class Config:
        frozen = True

    align: Optional[Literal["left", "center", "right", "fill", "justify", "center_across", "distributed"]] = Field(default=None)
    background: Optional[str] = Field(default=None)
    bold: Optional[bool] = Field(default=None)
    border: Optional[int] = Field(default=None)
    border_bottom: Optional[int] = Field(default=None)
    border_color: Optional[str] = Field(default=None)
    border_left: Optional[int] = Field(default=None)
    border_right: Optional[int] = Field(default=None)
    border_top: Optional[int] = Field(default=None)
    fill_inf: Optional[Any] = Field(default=None)
    fill_na: Optional[Any] = Field(default=None)
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
    valign: Optional[Literal["top", "vcenter", "bottom", "vcenter", "bottom", "vjustify"]] = Field(default=None)

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


class Component(BaseModel):
    style: Style = Field(default_factory=Style)

    class Config:
        arbitrary_types_allowed = True


class Text(Component):
    text: str
    width: int = Field(default=1)
    height: int = Field(default=1)
    merged: bool = Field(default=True)


class Link(Component):
    text: str
    url: str
    width: int = Field(default=1)
    height: int = Field(default=1)
    merged: bool = Field(default=True)


class Fill(Component):
    width: int = Field(default=1)
    height: int = Field(default=1)
    merged: bool = Field(default=True)


class Image(Component):
    path: Path
    width: int = Field(default=1)
    height: int = Field(default=1)


class Table(Component):
    data: pd.DataFrame
    header_style: Dict[str, Style] = Field(default_factory=dict)
    body_style: Style = Field(default_factory=Style)
    column_style: Dict[str, Union[Style, Callable[[Any], Style]]] = Field(default_factory=dict)
    idx_column_style: Dict[int, Style] = Field(default_factory=dict)
    column_width: Dict[str, int] = Field(default_factory=dict)
    idx_column_width: Dict[int, int] = Field(default_factory=dict)
    row_style: Dict[int, Style] = Field(default_factory=dict)
    max_col_width: Optional[int] = Field(default=None)
    header_filters: bool = Field(default=True)
    default_style: bool = Field(default=True)
    auto_width_tuning: int = Field(default=5)
    auto_width_padding: int = Field(default=5)
    merge_equal_headers: bool = Field(default=True)

    def with_stripes(
            self,
            color: str = "#D0D0D0",
            pattern: Literal["even", "odd"] = "odd",
    ) -> "Table":
        return self.model_copy(
            update=dict(
                row_style={
                    idx: self.row_style.get(idx, Style()).merge(
                        Style(background=color)
                    )
                    if (pattern == "odd" and idx % 2 != 0)
                       or (pattern == "even" and idx % 2 == 0)
                    else self.row_style.get(idx, Style())
                    for idx in range(self.data.shape[0])
                }
            )
        )


class Sheet(BaseModel):
    name: str
    components: Sequence[Component] = Field(default_factory=list)
    grid_lines: bool = Field(default=True)
    style: Style = Field(default_factory=Style)


class Excel(BaseModel):
    class Config:
        arbitrary_types_allowed = True

    path: Union[Path, io.BytesIO]
    sheets: Sequence[Sheet] = Field(default_factory=list)
    nan_inf_to_errors: bool = Field(default=True)
