"""Контракт данных задачи генерации (Job) — состояние, которое ведёт Оркестратор."""
from __future__ import annotations

from enum import Enum

from pydantic import BaseModel, Field

from app.schemas.content_plan import ContentPlan
from app.schemas.design_manifest import DesignManifest
from app.schemas.source_tables import SourceTable, TableAnalysisReport


class JobStatus(str, Enum):
    QUEUED = "queued"
    PARSING = "parsing"
    PLANNING = "planning"
    # Три варианта плана выступления готовы — ждём выбора содержания
    # (POST /jobs/{id}/select-plan), после чего собираются варианты дизайна.
    AWAITING_PLAN_CHOICE = "awaiting_plan_choice"
    ASSEMBLING = "assembling"
    # Три варианта собраны, аудит ещё не запущен — ждём выбора варианта
    # от пользователя (POST /jobs/{id}/select-variant) перед тем как предложить аудит/экспорт.
    AWAITING_VARIANT_CHOICE = "awaiting_variant_choice"
    # Вариант выбран — ждём решения пользователя: запустить детерминированный
    # аудит (POST /jobs/{id}/audit) или пропустить его (POST /jobs/{id}/skip-audit).
    AWAITING_AUDIT_CHOICE = "awaiting_audit_choice"
    AUDITING = "auditing"
    DONE = "done"
    FAILED = "failed"


class AuditCategory(str, Enum):
    LAYOUT = "layout"
    TEMPLATE = "template"
    DENSITY = "density"
    INTEGRITY = "integrity"
    CONTENT = "content"


class AuditBBox(BaseModel):
    """Рамка проблемной фигуры, нормализованная к размеру слайда (0..1) —
    чтобы подсветить место поверх PNG-превью слайда любого разрешения без
    передачи EMU-размеров слайда отдельно (см. app/pipeline/audit_overlay.py)."""

    x: float
    y: float
    width: float
    height: float


class AuditIssue(BaseModel):
    issue_id: str
    slide_id: str
    category: AuditCategory
    deterministic: bool
    description: str
    auto_fixable: bool = False
    # None — проблема на уровне всего слайда (пустой слайд, дубликат,
    # заполненность, VLM-вопросы о смысле) и обводить нечего, это ожидаемо.
    bbox: AuditBBox | None = None
    # Машиночитаемый тип проблемы (например, bullet_too_long, font_size_off_scale)
    # и фигура на слайде, к которой она относится — для точечных исправлений
    # (app/pipeline/audit_fix.py). None — проблема уровня слайда/колоды.
    code: str | None = None
    shape_id: int | None = None
    shape_name: str | None = None
    fixed: bool = False


class PlanVariant(BaseModel):
    """Один из трёх вариантов плана выступления (одна логика рассказа)."""

    key: str  # A | B | C
    name: str
    description: str
    story: str = ""
    plan: ContentPlan


class MaterialInfo(BaseModel):
    """Источник содержания, который пользователь добавил на втором шаге."""

    name: str
    kind: str  # file | link | note
    chars: int = 0
    status: str = "ok"  # ok | error
    note: str = ""


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
    # issue_id, к которым применялось точечное исправление (POST .../fix-issues);
    # после повторного аудита замечание либо исчезает, либо остаётся с той же формулировкой
    fixed_issue_ids: list[str] = Field(default_factory=list)
    variant_paths: dict[str, str] = Field(default_factory=dict)
    export_paths: dict[str, str] = Field(default_factory=dict)
    error: str | None = None
    # Вариант, выбранный пользователем на экране сравнения трёх вариантов вёрстки.
    # None до тех пор, пока пользователь не выбрал (статус AWAITING_VARIANT_CHOICE).
    selected_variant: str | None = None
    # Запущен ли детерминированный аудит (или пропущен пользователем) для выбранного варианта.
    audit_skipped: bool = False
    # Путь к Template JSON движка pptx_template_parser (см. app/pipeline/template_engine.py):
    # сам JSON (~0,5 МБ) в ответ API не попадает, только путь на диске.
    template_json_path: str | None = None
    # План планировщика по макетам шаблона (variant + content на слайд) для
    # каждого из трёх вариантов вёрстки — чтобы было видно, какие макеты выбраны.
    variant_plans: dict[str, list[str]] = Field(default_factory=dict)
    # Человекочитаемое описание оси различия каждого варианта (для карточек UI).
    variant_descriptions: dict[str, dict[str, str]] = Field(default_factory=dict)
    # Вид композиции каждого слайда варианта (cover/section/cards/chart/…) —
    # чтобы UI показывал, чем варианты отличаются на каждом слайде.
    variant_kinds: dict[str, list[str]] = Field(default_factory=dict)
    # Самовосстановление сборки: что сервис исправил сам в каждом варианте
    # («слайд 4: композиция не поместилась — подобрана другая») и почему
    # вариант не собрался совсем (остальные варианты при этом отдаются).
    variant_notes: dict[str, list[str]] = Field(default_factory=dict)
    variant_errors: dict[str, str] = Field(default_factory=dict)
    variant_sources: dict[str, list[int]] = Field(default_factory=dict, description="Для каждого варианта — индекс слайда плана содержания для каждого слайда варианта (варианты без разделителей короче плана)")
    # Аудит выбранного варианта проводился, но после него слайды дорабатывались
    # (другая компоновка / тексты / слайд из другого варианта) — результаты устарели.
    audit_stale: bool = False
    # История ручных «другая компоновка» по вариантам и индексам слайдов:
    # variant → {"3": [id композиции, …]} — чтобы каждый раз давать новую.
    slide_edits: dict[str, dict[str, list[str]]] = Field(default_factory=dict)
    # --- Брифование (4 шага на главной) -----------------------------------------
    idea: str = ""
    audience: str = ""
    goal: str = ""
    tone: str = ""
    must_have: str = ""
    materials_text: str = Field(default="", exclude=True)
    source_tables: list[SourceTable] = Field(default_factory=list, exclude=True)
    table_analysis: TableAnalysisReport | None = None
    materials: list[MaterialInfo] = Field(default_factory=list)
    # Три варианта плана выступления и выбранный пользователем.
    plan_variants: dict[str, PlanVariant] = Field(default_factory=dict)
    selected_plan: str | None = None
    # Результат аудита смысла (связь с главной идеей, целью, призывом в финале).
    alignment: dict | None = None
    # Какие операции с планом выполняются прямо сейчас — для индикаторов в интерфейсе.
    busy_note: str = ""
    slide_count: int = 10
    slide_count_max: int | None = None
    # План изменился после сборки дизайна — дизайн нужно собрать заново.
    designs_stale: bool = False
