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


class TableStyleTokens(BaseModel):
    """Стиль таблицы, извлечённый из реальной таблицы на слайде-образце
    шаблона (если такая в шаблоне есть) — приоритетный источник стиля нативных
    таблиц в генераторе (app/pipeline/styling.py) перед выводимым из палитры/
    типографики. None в любом поле — в исходной таблице это не было однозначно
    определимо (наследовано из темы, а не задано явно), тогда генератор берёт
    значение из палитры для этого поля.
    """

    header_fill: str | None = None
    header_text_color: str | None = None
    header_bold: bool | None = None
    row_odd_fill: str | None = None
    row_even_fill: str | None = None
    body_font: str | None = None
    body_size: int | None = None


class ChartStyleTokens(BaseModel):
    """Стиль графика, извлечённый из реального графика на слайде-образце
    шаблона, аналогично TableStyleTokens.
    """

    series_colors: list[str] = Field(default_factory=list)
    font: str | None = None
    font_size: int | None = None
    has_legend: bool | None = None


class DesignManifest(BaseModel):
    template_id: str
    palette: dict[str, str]  # роль темы (dk1/lt1/accent1..6/...) -> hex-цвет
    typography: Typography
    layouts: list[Layout]
    slide_width_emu: int
    slide_height_emu: int
    # None если в слайдах-образцах шаблона не нашлось ни одной реальной
    # таблицы/графика — тогда styling.py использует только palette/typography,
    # как и раньше.
    table_style: TableStyleTokens | None = None
    chart_style: ChartStyleTokens | None = None
