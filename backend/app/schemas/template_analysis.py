"""Контракт данных для отдельной стадии «анализ шаблона» — вне пайплайна Job.

Позволяет проверить/отредактировать дизайн-систему, извлечённую из .pptx,
ДО запуска генерации презентации. TemplateAnalysis хранит как «сырой»
DesignManifest (контракт парсера, см. design_manifest.py), так и путь к
исходному файлу шаблона — чтобы Job мог переиспользовать уже проверенный
(и, возможно, вручную отредактированный) манифест вместо повторного парсинга.
"""
from __future__ import annotations

from enum import Enum

from pydantic import BaseModel

from app.schemas.design_manifest import DesignManifest


class TemplateAnalysisStatus(str, Enum):
    ANALYZING = "analyzing"
    READY = "ready"
    FAILED = "failed"


class TemplateAnalysis(BaseModel):
    template_analysis_id: str
    status: TemplateAnalysisStatus = TemplateAnalysisStatus.ANALYZING
    original_filename: str
    template_path: str
    manifest: DesignManifest | None = None
    # True после ручной правки через PUT .../manifest — отличает
    # отредактированный вручную манифест от чистого вывода парсера
    # (полезно для UI и для будущей телеметрии качества авто-разбора).
    manually_edited: bool = False
    error: str | None = None
