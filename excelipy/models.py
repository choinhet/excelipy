from pathlib import Path
from typing import Dict, Optional, Sequence, Literal

import pandas as pd
from pydantic import BaseModel, Field


class Style(BaseModel):
    class Config:
        frozen = True

    align: Optional[
        Literal[
            "left",
            "center",
            "right",
            "fill",
            "justify",
            "center_across",
            "distributed",
        ]
    ] = Field(default=None)
    valign: Optional[
        Literal[
            "top",
            "vcenter",
            "bottom",
            "vcenter",
            "bottom",
            "vjustify",
        ]
    ] = Field(default=None)
    padding: Optional[int] = Field(default=None)
    padding_left: Optional[int] = Field(default=None)
    padding_right: Optional[int] = Field(default=None)
    padding_top: Optional[int] = Field(default=None)
    padding_bottom: Optional[int] = Field(default=None)
    font_size: Optional[int] = Field(default=None)
    font_color: Optional[str] = Field(default=None)
    font_family: Optional[str] = Field(default=None)
    bold: Optional[bool] = Field(default=None)
    border: Optional[int] = Field(default=None)
    border_left: Optional[int] = Field(default=None)
    border_right: Optional[int] = Field(default=None)
    border_top: Optional[int] = Field(default=None)
    border_bottom: Optional[int] = Field(default=None)
    border_color: Optional[str] = Field(default=None)
    background: Optional[str] = Field(default=None)

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


class Fill(Component):
    width: int = Field(default=1)
    height: int = Field(default=1)


class Image(Component):
    path: Path
    width: int = Field(default=1)
    height: int = Field(default=1)


class Table(Component):
    data: pd.DataFrame
    header_style: Style = Field(default_factory=Style)
    body_style: Style = Field(default_factory=Style)
    column_style: Dict[str, Style] = Field(default_factory=dict)
    column_width: Dict[str, int] = Field(default_factory=dict)
    row_style: Dict[int, Style] = Field(default_factory=dict)
    max_col_width: Optional[int] = Field(default=None)
    header_filters: bool = Field(default=True)

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
    style: Style = Field(default_factory=Style)


class Excel(BaseModel):
    path: Path
    sheets: Sequence[Sheet] = Field(default_factory=list)
