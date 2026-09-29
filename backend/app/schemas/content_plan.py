"""Контракты данных контент-плана (выход Контент-агента, app/pipeline/content_agent.py)."""
from __future__ import annotations

from enum import Enum
from typing import Self

from pydantic import Field, model_validator

from app.schemas.strict_model import StrictModel


class ContentBlockType(str, Enum):
    BULLETS = "bullets"
    CHART = "chart"
    TABLE = "table"
    TEXT = "text"


class ChartType(str, Enum):
    BAR = "bar"
    LINE = "line"
    PIE = "pie"


class ChartSeries(StrictModel):
    name: str
    values: list[float]


class ChartData(StrictModel):
    chart_type: ChartType
    categories: list[str]
    series: list[ChartSeries]

    @model_validator(mode="after")
    def validate_series(self) -> Self:
        names = [series.name for series in self.series]
        if len(names) != len(set(names)):
            raise ValueError("Названия серий должны быть уникальными")

        for series in self.series:
            if len(series.values) != len(self.categories):
                raise ValueError(
                    f"Серия «{series.name}»: число значений должно совпадать "
                    "с числом категорий"
                )
        return self


class TableData(StrictModel):
    headers: list[str]
    rows: list[list[str]]


class ContentBlock(StrictModel):
    """Заполнено только поле, соответствующее type; остальные поля равны null."""

    type: ContentBlockType
    bullets: list[str] | None = None
    text: str | None = None
    chart: ChartData | None = None
    table: TableData | None = None

    @model_validator(mode="after")
    def validate_payload(self) -> Self:
        populated = [
            field for field in ("bullets", "text", "chart", "table")
            if getattr(self, field) is not None
        ]
        if populated != [self.type.value]:
            raise ValueError(
                f"Для type={self.type.value} заполни только {self.type.value}; "
                "остальные поля должны быть null"
            )
        if not getattr(self, self.type.value):
            raise ValueError(f"Поле {self.type.value} не должно быть пустым")
        return self


class SlideSpec(StrictModel):
    slide_id: str = Field(min_length=1)
    purpose: str
    title: str
    content_blocks: list[ContentBlock] = Field(
        default_factory=list,
        description="Обычно 1–2 разных блока. Не повторяй одинаковые блоки на слайде.",
    )


class ContentPlan(StrictModel):
    brief: str = Field(description="Краткое резюме задания, одно предложение.")
    purpose: str = Field(description="Цель презентации, одно предложение.")
    slides: list[SlideSpec]

    @model_validator(mode="after")
    def validate_slide_ids(self) -> Self:
        ids = [slide.slide_id for slide in self.slides]
        if len(ids) != len(set(ids)):
            raise ValueError("slide_id должны быть уникальными")
        return self
