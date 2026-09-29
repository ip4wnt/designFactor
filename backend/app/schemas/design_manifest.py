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


class SlideElementType(str, Enum):
    TEXT = "text"
    PICTURE = "picture"
    TABLE = "table"
    CHART = "chart"
    OTHER = "other"


class SlideElement(BaseModel):
    """Одна фигура на реальном слайде-образце — минимум, нужный для
    отрисовки HTML-макетика в UI (позиционированный div на холсте слайда),
    без рендера через LibreOffice.

    x/y/width/height — EMU, та же единица что и slide_width_emu/
    slide_height_emu на DesignManifest, поэтому фронтенд может отмасштабировать
    напрямую делением на размеры слайда.
    """

    element_type: SlideElementType
    x_emu: int
    y_emu: int
    width_emu: int
    height_emu: int
    text: str | None = None  # только для TEXT — реальный текст фигуры, обрезанный
    fill_hex: str | None = None  # явный fill фигуры, если задан прямо на слайде
    text_color_hex: str | None = None  # явный цвет текста первого run с текстом
    font_size: int | None = None  # pt, первого run с текстом
    bold: bool | None = None
    is_title_role: bool = False  # фигура — плейсхолдер типа title/centerTitle
    # Поля ниже — из доработки парсера коллеги (pptx_template_parser_generator):
    # выравнивание текста внутри фигуры и метаданные картинки, нужные для более
    # точной раскладки мокапов в UI и генератора.
    horizontal_align: str | None = None  # left | center | right | justify | ...
    vertical_align: str | None = None  # top | middle | bottom
    margin_left_emu: int | None = None
    margin_right_emu: int | None = None
    margin_top_emu: int | None = None
    margin_bottom_emu: int | None = None
    picture_description: str | None = None  # alt-текст картинки (cNvPr/@descr), если задан автором шаблона


class LayoutExample(BaseModel):
    """Один реальный слайд-образец из презентации-шаблона, использующий
    данный layout_id — не пустая заготовка макета, а то, как его фактически
    заполнили автором шаблона (позиции, текст, цвета фигур). Обходятся ВСЕ
    слайды-образцы, попавшие в этот layout — в шаблоне может быть много
    примеров на один и тот же (особенно 'свободный дизайн') макет.
    """

    slide_index: int  # 1-based позиция в исходной презентации, для стабильной сортировки
    slide_name: str | None = None
    elements: list[SlideElement] = Field(default_factory=list)


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
    # layout_id -> список слайдов-образцов, использующих этот макет. Пусто
    # для layout_id, на который в презентации не нашлось ни одного слайда-
    # примера (только пустая заготовка макета в самом шаблоне).
    layout_examples: dict[str, list[LayoutExample]] = Field(default_factory=dict)
