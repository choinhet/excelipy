from pathlib import Path
from typing import Dict, Optional, Sequence

import pandas as pd
from pydantic import BaseModel, Field


class Style(BaseModel):
    class Config:
        frozen = True

    padding: Optional[int] = Field(default=None)
    padding_left: Optional[int] = Field(default=None)
    padding_right: Optional[int] = Field(default=None)
    padding_top: Optional[int] = Field(default=None)
    padding_bottom: Optional[int] = Field(default=None)
    margin: Optional[int] = Field(default=None)
    margin_left: Optional[int] = Field(default=None)
    margin_right: Optional[int] = Field(default=None)
    margin_top: Optional[int] = Field(default=None)
    margin_bottom: Optional[int] = Field(default=None)
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


class Table(Component):
    data: pd.DataFrame
    header_style: Style = Field(default_factory=Style)
    body_style: Style = Field(default_factory=Style)
    column_style: Dict[str, Style] = Field(default_factory=dict)
    predefined_style: Optional[str] = Field(default="Table Style Light 1")
    max_col_width: Optional[int] = Field(default=None)


class Sheet(BaseModel):
    name: str
    components: Sequence[Component] = Field(default_factory=list)
    style: Style = Field(default_factory=Style)


class Excel(BaseModel):
    path: Path
    sheets: Sequence[Sheet] = Field(default_factory=list)
