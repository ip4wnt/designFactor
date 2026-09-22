"""Оркестратор: последовательно прогоняет Job через стадии пайплайна.

Каждая стадия выполняется синхронно внутри одной asyncio-задачи процесса
backend (app.api.routes создаёт её через asyncio.create_task) — без очередей,
воркеров и Celery/Redis: это осознанное упрощение MVP-монолита (см. README).
Ошибка на любой стадии останавливает пайплайн и переводит job в failed.
"""
from __future__ import annotations

import logging

from app.config import get_settings
from app.pipeline import assembly, audit, content_agent, parser
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
        output_path = settings.outputs_dir / job.job_id / "presentation.pptx"
        assembly.assemble_presentation(
            template_path=job.template_path,
            manifest=job.design_manifest,
            content_plan=job.content_plan,
            output_path=output_path,
        )
        job.variant_paths["variant_a"] = str(output_path)
        job.export_paths["pptx"] = str(output_path)

        job.status = JobStatus.AUDITING
        job.audit_issues = audit.audit_presentation(output_path)

        job.status = JobStatus.DONE
    except Exception as exc:  # noqa: BLE001 — любая ошибка стадии должна перевести job в failed
        logger.exception("Пайплайн упал для job %s", job.job_id)
        job.status = JobStatus.FAILED
        job.error = str(exc)


async def rerun_assembly_and_audit(job: Job, variant: str = "variant_a") -> None:
    """Перезапуск Компоновщика+Аудита для POST /jobs/{id}/fix.

    Ограничение текущей версии: список issue_id из запроса пока не
    применяется точечно (нет привязки AuditIssue к конкретному
    content-блоку) — просто пересобирает колоду и гоняет аудит заново.
    Честно описано в README/docs/AUDIT.md как TODO.
    """
    settings = get_settings()
    try:
        job.status = JobStatus.ASSEMBLING
        output_path = settings.outputs_dir / job.job_id / "presentation.pptx"
        assembly.assemble_presentation(
            template_path=job.template_path,
            manifest=job.design_manifest,
            content_plan=job.content_plan,
            output_path=output_path,
            variant=variant,
        )
        job.variant_paths[variant] = str(output_path)
        job.export_paths["pptx"] = str(output_path)

        job.status = JobStatus.AUDITING
        job.audit_issues = audit.audit_presentation(output_path)

        job.status = JobStatus.DONE
    except Exception as exc:  # noqa: BLE001
        logger.exception("Пере-сборка упала для job %s", job.job_id)
        job.status = JobStatus.FAILED
        job.error = str(exc)
