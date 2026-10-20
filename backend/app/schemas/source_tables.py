"""Server-owned material tables. Models select IDs, never supply source cells."""
from typing import Any, Literal
from pydantic import BaseModel, Field

from app.schemas.content_plan import ChartData, TableData
from app.schemas.strict_model import StrictModel


class TableColumn(StrictModel):
    key: str = Field(min_length=1, description="Точное рабочее имя колонки результата SQL/DataFrame.")
    label: str = Field(min_length=1, max_length=160, description="Читаемое название для аудитории на языке презентации, без SQL-имен и сокращений вроде aov.")
    unit: str = Field(description="Единица в текущем масштабе значений, например тыс. руб., шт., %. Пустая строка для категории или когда единица неизвестна; не выдумывай ее.")


class SourceTable(BaseModel):
    table_id: str
    source_name: str
    ordinal: int
    summary: str
    available: bool
    limitation: str = ""
    data: TableData | None = None
    parent_table_ids: list[str] = Field(default_factory=list)
    operations: list[dict[str, Any]] = Field(default_factory=list)
    column_metadata: list[TableColumn] = Field(default_factory=list)
    sheet_name: str = ""
    cell_range: str = ""
    number_formats: dict[str, list[str]] = Field(default_factory=dict)
    page_number: int | None = None
    bounding_box: list[float] = Field(default_factory=list)
    recommended_presentation: Literal["table", "chart"] = "table"
    presentation_reason: str = ""
    chart: ChartData | None = None


class TableAnalysisReport(BaseModel):
    key: str = ""
    status: Literal["completed", "skipped", "disabled", "timeout", "failed", "partial"]
    artifacts: list[SourceTable] = Field(default_factory=list)
    notes: list[str] = Field(default_factory=list)
    elapsed_ms: int = 0
    cache_hit: bool = False
    model_calls: int = 0
    tool_calls: int = 0


class DocumentMaterial(BaseModel):
    format_version: int = 2
    material_id: str
    filename: str
    text: str
    tables: list[SourceTable] = Field(default_factory=list)
    note: str = ""
