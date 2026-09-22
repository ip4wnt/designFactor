"""Контракты данных дизайн-манифеста шаблона (выход Парсера, app/pipeline/parser.py)."""
from __future__ import annotations

from enum import Enum

from pydantic import BaseModel, Field


class LayoutRoleType(str, Enum):
    TITLE = "title"
    BODY = "body"
    CHART = "chart"
    TABLE = "table"
    PICTURE = "picture"


class TypographyStyle(BaseModel):
    font: str
    size: int  # пункты (pt)
    bold: bool = False


class LayoutRole(BaseModel):
    role: LayoutRoleType
    placeholder_idx: int
    max_items: int | None = None
    max_words_per_item: int | None = None


class Layout(BaseModel):
    layout_id: str
    source_layout_name: str
    roles: list[LayoutRole] = Field(default_factory=list)
    supports_chart: bool = False
    supports_table: bool = False
    supports_image: bool = False


class Typography(BaseModel):
    title: TypographyStyle
    body: TypographyStyle


class DesignManifest(BaseModel):
    template_id: str
    palette: dict[str, str]  # роль темы (dk1/lt1/accent1..6/...) -> hex-цвет
    typography: Typography
    layouts: list[Layout]
    slide_width_emu: int
    slide_height_emu: int
