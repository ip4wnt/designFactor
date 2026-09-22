"""Оркестратор: последовательно прогоняет Job через стадии пайплайна.

Каждая стадия выполняется синхронно внутри одной asyncio-задачи процесса
backend (app.api.routes создаёт её через asyncio.create_task) — без очередей,
воркеров и Celery/Redis: это осознанное упрощение MVP-монолита (см. README).
Ошибка на любой стадии останавливает пайплайн и переводит job в failed.

Стадия ASSEMBLING+AUDITING прогоняется по ТРЁМ вариантам вёрстки
(variant_a/b/c, см. app/pipeline/layout_engine.py) — на выходе три .pptx-файла
и три независимых набора проблем аудита, а не один: ТЗ прямо требует "три
файла презентации" на выходе одного job.
"""
from __future__ import annotations

import logging

from app.config import get_settings
from app.pipeline import assembly, audit, content_agent, parser
from app.pipeline.assembly import VALID_VARIANTS
from app.schemas.content_plan import ContentPlan
from app.schemas.job import Job, JobStatus

logger = logging.getLogger(__name__)

_DEFAULT_SLIDE_COUNT = 10


async def run_pipeline(job: Job, content_plan: ContentPlan | None = None) -> None:
    """queued -> parsing -> [planning] -> assembling -> auditing -> done/failed.

    `content_plan`, когда передан, — это сценарий «модель или пользователь
    уже подготовили ContentPlan (включая ChartData/TableData) и передают его
    прямо парсеру/компоновщику», в обход LLM-контент-агента. Стадия PLANNING в этом
    случае не вызывает generate_content_plan, а лишь принимает готовый план —
    остальной пайплайн (сборка, аудит) идентичен обоим сценариям.
    """
    settings = get_settings()
    try:
        job.status = JobStatus.PARSING
        job.design_manifest = parser.parse_template(job.template_path)

        job.status = JobStatus.PLANNING
        if content_plan is not None:
            job.content_plan = content_plan
        else:
            job.content_plan = await content_agent.generate_content_plan(
                brief=job.brief, purpose=job.purpose, slide_count=_DEFAULT_SLIDE_COUNT
            )

        job.status = JobStatus.ASSEMBLING
        _assemble_and_audit_all_variants(job, settings)

        job.status = JobStatus.DONE
    except Exception as exc:  # noqa: BLE001 — любая ошибка стадии должна перевести job в failed
        logger.exception("Пайплайн упал для job %s", job.job_id)
        job.status = JobStatus.FAILED
        job.error = str(exc)


async def rerun_assembly_and_audit(job: Job, variant: str | None = None) -> None:
    """Перезапуск Компоновщика+Аудита для POST /jobs/{id}/fix.

    `variant=None` (по умолчанию) пересобирает ВСЕ три варианта — это путь
    для правок ContentPlan (импорт Excel, добавление блока), которые должны
    попасть во все варианты одинаково. `variant="variant_b"` и т.п. пересобирает
    только один — для точечной пересборки после локальной правки конкретного
    варианта (зарезервировано на будущее, сейчас UI всегда правит план целиком).

    Ограничение текущей версии: список issue_id из запроса пока не
    применяется точечно (нет привязки AuditIssue к конкретному
    content-блоку) — просто пересобирает колоду(и) и гоняет аудит заново.
    Честно описано в README/docs/AUDIT.md как TODO.
    """
    settings = get_settings()
    try:
        job.status = JobStatus.ASSEMBLING
        if variant is None:
            _assemble_and_audit_all_variants(job, settings)
        else:
            _assemble_and_audit_one_variant(job, settings, variant)

        job.status = JobStatus.DONE
    except Exception as exc:  # noqa: BLE001
        logger.exception("Пере-сборка упала для job %s", job.job_id)
        job.status = JobStatus.FAILED
        job.error = str(exc)


def _assemble_and_audit_all_variants(job: Job, settings) -> None:
    for variant in VALID_VARIANTS:
        _assemble_and_audit_one_variant(job, settings, variant)
    job.status = JobStatus.AUDITING


def _assemble_and_audit_one_variant(job: Job, settings, variant: str) -> None:
    output_path = settings.outputs_dir / job.job_id / f"presentation_{variant}.pptx"
    assembly.assemble_presentation(
        template_path=job.template_path,
        manifest=job.design_manifest,
        content_plan=job.content_plan,
        output_path=output_path,
        variant=variant,
    )
    job.variant_paths[variant] = str(output_path)
    job.export_paths["pptx"] = str(output_path)  # последний собранный вариант — путь для /export по умолчанию

    job.status = JobStatus.AUDITING
    job.audit_issues[variant] = audit.audit_presentation(output_path)
