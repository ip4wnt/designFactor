"""HTTP API поверх Job-оркестратора."""
from __future__ import annotations

import asyncio
import tempfile
import uuid
from pathlib import Path

from fastapi import APIRouter, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse
from pydantic import BaseModel

from app.config import get_settings
from app.job_store import get as get_job
from app.job_store import save as save_job
from app.orchestrator import rerun_assembly_and_audit, run_pipeline
from app.pipeline.excel_import import ExcelImportError, import_excel_block, list_sheet_names
from app.pipeline.export import ExportError, export_presentation
from app.schemas.content_plan import ContentBlock, ContentBlockType, ContentPlan
from app.schemas.job import Job, JobStatus

router = APIRouter()

_EXPORT_MEDIA_TYPES = {
    "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "pdf": "application/pdf",
    "html": "text/html",
}


class FixRequest(BaseModel):
    issue_ids: list[str]


class AddContentBlockRequest(BaseModel):
    """Второй вариант ввода данных таблиц/графиков — готовый JSON вручную
    (без загрузки Excel), для тест-фронта и будущей модели-генератора контента."""

    block: ContentBlock


@router.post("/jobs")
async def create_job(
    template: UploadFile = File(...),
    brief: str = Form(...),
    purpose: str = Form(...),
    content_plan: str | None = Form(
        None,
        description=(
            "Опционально: готовый ContentPlan в виде JSON-строки (включая "
            "content_blocks с ChartData/TableData). Когда передан, планирование "
            "через LLM-контент-агента пропускается — это путь для сценария "
            "«модель/пользователь уже подготовили данные»."
        ),
    ),
) -> dict[str, str]:
    settings = get_settings()
    job_id = str(uuid.uuid4())

    template_path = settings.templates_dir / f"{job_id}_{template.filename}"
    template_path.write_bytes(await template.read())

    parsed_plan: ContentPlan | None = None
    if content_plan is not None:
        try:
            parsed_plan = ContentPlan.model_validate_json(content_plan)
        except Exception as exc:  # noqa: BLE001 — невалидный JSON от клиента, не наша ошибка
            raise HTTPException(status_code=400, detail=f"content_plan невалиден: {exc}") from exc

    job = Job(job_id=job_id, brief=brief, purpose=purpose, template_path=str(template_path))
    save_job(job)

    # Не дожидаемся генерации — задача летит в фоне того же процесса
    # (asyncio-таск, без Celery/Redis), эндпоинт сразу возвращает job_id.
    asyncio.create_task(run_pipeline(job, content_plan=parsed_plan))

    return {"job_id": job_id}


@router.get("/jobs/{job_id}")
async def get_job_status(job_id: str) -> Job:
    job = get_job(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="Job не найден")
    return job


@router.post("/jobs/{job_id}/fix")
async def fix_job(job_id: str, payload: FixRequest) -> dict[str, str]:
    job = get_job(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="Job не найден")
    if job.status not in (JobStatus.DONE, JobStatus.FAILED):
        raise HTTPException(status_code=409, detail="Job ещё выполняется")
    if job.content_plan is None or job.design_manifest is None:
        raise HTTPException(status_code=409, detail="Job не дошёл до стадии сборки")

    known_ids = {
        issue.issue_id for issues in job.audit_issues.values() for issue in issues
    }
    unknown = set(payload.issue_ids) - known_ids
    if unknown:
        raise HTTPException(status_code=400, detail=f"Неизвестные issue_id: {sorted(unknown)}")

    asyncio.create_task(rerun_assembly_and_audit(job))
    return {"job_id": job_id, "status": job.status.value}


@router.post("/jobs/{job_id}/slides/{slide_id}/import-excel")
async def import_excel_into_slide(
    job_id: str,
    slide_id: str,
    file: UploadFile = File(...),
    sheet_name: str | None = Form(None),
    kind: str | None = Form(None),
) -> dict[str, object]:
    """Читает лист Excel, конвертирует в ChartData/TableData и добавляет
    получившийся content_block к указанному слайду ContentPlan'а.

    Не пересобирает .pptx сама — после успешного импорта нужно вызвать
    POST /jobs/{job_id}/fix, чтобы Компоновщик учёл новый content_block.
    Это то же самое разделение, что уже применяется для audit fix: сначала
    правится план, потом одна явная пересборка.
    """
    job = get_job(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="Job не найден")
    if job.content_plan is None:
        raise HTTPException(status_code=409, detail="Job ещё не дошёл до стадии планирования")
    if kind not in (None, "table", "chart"):
        raise HTTPException(status_code=400, detail="kind должен быть 'table', 'chart' или не задан")

    slide_spec = next((s for s in job.content_plan.slides if s.slide_id == slide_id), None)
    if slide_spec is None:
        raise HTTPException(status_code=404, detail=f"Слайд {slide_id} не найден в ContentPlan")

    suffix = Path(file.filename or "upload.xlsx").suffix or ".xlsx"
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        tmp.write(await file.read())
        tmp_path = Path(tmp.name)

    try:
        resolved_kind, data = import_excel_block(tmp_path, sheet_name=sheet_name, kind=kind)
    except ExcelImportError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    finally:
        tmp_path.unlink(missing_ok=True)

    if resolved_kind == "table":
        block = ContentBlock(type=ContentBlockType.TABLE, table=data)
    else:
        block = ContentBlock(type=ContentBlockType.CHART, chart=data)

    slide_spec.content_blocks.append(block)
    save_job(job)

    return {"job_id": job_id, "slide_id": slide_id, "kind": resolved_kind, "block": block.model_dump()}


@router.post("/jobs/{job_id}/slides/{slide_id}/content-block")
async def add_content_block(job_id: str, slide_id: str, payload: AddContentBlockRequest) -> dict[str, object]:
    """Добавляет готовый ContentBlock (ChartData/TableData в виде JSON) к слайду.

    Симметричен import-excel: там лист Excel конвертируется в тот же
    ContentBlock и добавляется так же — оба пути сходятся в один контракт.
    Аналогично требует последующего POST /jobs/{job_id}/fix для пересборки.
    """
    job = get_job(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="Job не найден")
    if job.content_plan is None:
        raise HTTPException(status_code=409, detail="Job ещё не дошёл до стадии планирования")

    slide_spec = next((s for s in job.content_plan.slides if s.slide_id == slide_id), None)
    if slide_spec is None:
        raise HTTPException(status_code=404, detail=f"Слайд {slide_id} не найден в ContentPlan")

    slide_spec.content_blocks.append(payload.block)
    save_job(job)

    return {"job_id": job_id, "slide_id": slide_id, "block": payload.block.model_dump()}


@router.post("/excel/sheets")
async def preview_excel_sheets(file: UploadFile = File(...)) -> dict[str, list[str]]:
    """Возвращает имена листов загруженного Excel-файла для выбора перед импортом."""
    suffix = Path(file.filename or "upload.xlsx").suffix or ".xlsx"
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        tmp.write(await file.read())
        tmp_path = Path(tmp.name)
    try:
        return {"sheets": list_sheet_names(tmp_path)}
    except Exception as exc:  # noqa: BLE001 — файл может быть не .xlsx вовсе
        raise HTTPException(status_code=400, detail=f"Не удалось прочитать файл: {exc}") from exc
    finally:
        tmp_path.unlink(missing_ok=True)


@router.get("/jobs/{job_id}/export")
async def export_job(job_id: str, format: str = "pptx", variant: str = "variant_a") -> FileResponse:
    job = get_job(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="Job не найден")
    if job.status != JobStatus.DONE:
        raise HTTPException(status_code=409, detail="Job ещё не готов")

    pptx_path = job.variant_paths.get(variant)
    if pptx_path is None:
        raise HTTPException(status_code=404, detail=f"У job нет собранного варианта '{variant}'")

    settings = get_settings()
    output_dir = settings.outputs_dir / job.job_id / variant
    try:
        result_path = await export_presentation(pptx_path, output_dir, format)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except ExportError as exc:
        raise HTTPException(status_code=502, detail=str(exc)) from exc
    job.export_paths[f"{variant}_{format}"] = str(result_path)

    media_type = _EXPORT_MEDIA_TYPES.get(format, "application/octet-stream")
    return FileResponse(result_path, media_type=media_type, filename=result_path.name)
