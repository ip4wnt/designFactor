"""Контракт данных задачи генерации (Job) — состояние, которое ведёт Оркестратор."""
from __future__ import annotations

from enum import Enum

from pydantic import BaseModel, Field

from app.schemas.content_plan import ContentPlan
from app.schemas.design_manifest import DesignManifest


class JobStatus(str, Enum):
    QUEUED = "queued"
    PARSING = "parsing"
    PLANNING = "planning"
    ASSEMBLING = "assembling"
    AUDITING = "auditing"
    DONE = "done"
    FAILED = "failed"


class AuditCategory(str, Enum):
    LAYOUT = "layout"
    TEMPLATE = "template"
    DENSITY = "density"
    INTEGRITY = "integrity"
    CONTENT = "content"


class AuditIssue(BaseModel):
    issue_id: str
    slide_id: str
    category: AuditCategory
    deterministic: bool
    description: str
    auto_fixable: bool = False


class Job(BaseModel):
    job_id: str
    status: JobStatus = JobStatus.QUEUED
    brief: str
    purpose: str
    template_path: str
    design_manifest: DesignManifest | None = None
    content_plan: ContentPlan | None = None
    # Ключ — variant_a/variant_b/variant_c: у каждого варианта вёрстки свой
    # набор проблем аудита, т.к. геометрия и, соответственно, наложения/
    # переполнения у них разные (см. app/pipeline/layout_engine.py).
    audit_issues: dict[str, list[AuditIssue]] = Field(default_factory=dict)
    variant_paths: dict[str, str] = Field(default_factory=dict)
    export_paths: dict[str, str] = Field(default_factory=dict)
    error: str | None = None
