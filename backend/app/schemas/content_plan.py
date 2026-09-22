"""Контракты данных контент-плана (выход Контент-агента, app/pipeline/content_agent.py)."""
from __future__ import annotations

from enum import Enum

from pydantic import BaseModel, Field


class ContentBlockType(str, Enum):
    BULLETS = "bullets"
    CHART = "chart"
    TABLE = "table"
    TEXT = "text"


class ChartType(str, Enum):
    BAR = "bar"
    LINE = "line"
    PIE = "pie"


class ChartData(BaseModel):
    chart_type: ChartType
    categories: list[str]
    series: dict[str, list[float]]


class TableData(BaseModel):
    headers: list[str]
    rows: list[list[str]]


class ContentBlock(BaseModel):
    type: ContentBlockType
    bullets: list[str] | None = None
    text: str | None = None
    chart: ChartData | None = None
    table: TableData | None = None


class SlideSpec(BaseModel):
    slide_id: str
    purpose: str
    title: str
    content_blocks: list[ContentBlock] = Field(default_factory=list)


class ContentPlan(BaseModel):
    brief: str
    purpose: str
    slides: list[SlideSpec]
