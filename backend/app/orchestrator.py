"""Оркестратор: последовательно прогоняет Job через стадии пайплайна.

Каждая стадия выполняется синхронно внутри одной asyncio-задачи процесса
backend (app.api.routes создаёт её через asyncio.create_task) — без очередей,
воркеров и Celery/Redis: это осознанное упрощение MVP-монолита (см. README).
Ошибка на любой стадии останавливает пайплайн и переводит job в failed.

Стадия PARSING строит два представления шаблона: DesignManifest (наш парсер —
цвета/шрифты/стили таблиц и графиков, нужен стилизации объектов и аудиту) и
Template JSON движка pptx_template_parser (макеты шаблона с ролями слотов и
контрактами вместимости — источник истины для вёрстки). Стадия PLANNING
параллельно получает от модели план по макетам (планировщик коллеги, промпты
в pptx_template_parser_generator/prompts/) и ContentPlan с данными для
графиков/таблиц (наш контент-агент). Стадия ASSEMBLING строит из одного плана
ТРИ варианта вёрстки (variant_a/b/c — ось различия задаёт
template_engine.VARIANT_STRATEGIES) и рендерит их в оригинальных макетах
шаблона: ТЗ прямо требует "три варианта дизайна" на выходе одного job. После сборки всего job
останавливается в AWAITING_VARIANT_CHOICE и ждёт выбора пользователя (POST
/jobs/{id}/select-variant). Аудит больше НЕ запускается автоматически для всех
вариантов — после выбора варианта job переходит в AWAITING_AUDIT_CHOICE, и
пользователь явно решает, запустить детерминированный аудит (POST
/jobs/{id}/audit) или сразу перейти к экспорту (POST /jobs/{id}/skip-audit).
"""
from __future__ import annotations

import asyncio
import json
import logging
import uuid
from pathlib import Path

from app.config import get_settings
from app.pipeline import audit, audit_fix, content_agent, dramaturgy, parser, template_engine
from app.pipeline.template_engine import VALID_VARIANTS, VARIANT_STRATEGIES
from app.schemas.content_plan import ContentPlan
from app.schemas.job import AuditCategory, AuditIssue, Job, JobStatus

logger = logging.getLogger(__name__)

_DEFAULT_SLIDE_COUNT = 10


async def run_pipeline(
    job: Job,
    content_plan: ContentPlan | None = None,
    case_prompt_addition: str = "",
    slide_count: int | None = None,
    slide_count_max: int | None = None,
) -> None:
    """queued -> parsing -> [planning] -> assembling -> awaiting_variant_choice -> ... -> done/failed.

    `content_plan`, когда передан, — это сценарий «модель или пользователь
    уже подготовили ContentPlan (включая ChartData/TableData) и передают его
    прямо парсеру/компоновщику», в обход LLM-контент-агента. Стадия PLANNING в этом
    случае не вызывает generate_content_plan, а лишь принимает готовый план —
    остальной пайплайн (сборка) идентичен обоим сценариям.

    `case_prompt_addition` — доп. указания контент-агенту от выбранного кейса-пресета
    (см. app/pipeline/cases.py), `slide_count` — целевое число слайдов из того же
    пресета (иначе используется _DEFAULT_SLIDE_COUNT); `slide_count_max` — верхняя
    граница диапазона, если пользователь задал «от–до» (модель выбирает сама).
    """
    settings = get_settings()
    try:
        job.status = JobStatus.PARSING
        # design_manifest может быть уже заполнен, если job создан через
        # template_analysis_id (см. POST /jobs в api/routes.py) — тогда используется
        # уже проанализированный (возможно, вручную отредактированный) манифест
        # как есть, без повторного разбора .pptx.
        if job.design_manifest is None:
            job.design_manifest = parser.parse_template(job.template_path)
        template = await asyncio.to_thread(template_engine.build_template_json, job.template_path)
        job.template_json_path = str(_persist_template_json(job, settings, template))

        job.status = JobStatus.PLANNING
        target_slides = slide_count or _DEFAULT_SLIDE_COUNT
        model = template_engine.BackendModel(settings.LLM_BASE_URL, settings.LLM_MODEL_NAME, settings.LLM_API_KEY)
        # Сначала контент-агент (структура и содержание слайдов), затем
        # планировщик макетов получает эту структуру как бриф — иначе два
        # независимых прохода LLM дают две разные истории, и графики/таблицы
        # из ContentPlan попадают не на «свои» слайды.
        if content_plan is None:
            content_plan = await content_agent.generate_content_plan(
                brief=job.brief,
                purpose=job.purpose,
                slide_count=target_slides,
                case_prompt_addition=case_prompt_addition,
                slide_count_max=slide_count_max,
            )
        job.content_plan = content_plan
        # Драматургия: структурированные тексты (один вызов модели на колоду,
        # с детерминированным запасным вариантом) → композиции выбираются
        # правилами в _assemble_one_variant. План по макетам старого
        # планировщика строится лениво, только если драматургия не справилась.
        structured = await asyncio.to_thread(dramaturgy.structure_content, content_plan, job.brief, model)
        _persist_json(settings.outputs_dir / job.job_id / "structured.json", structured)
        layout_plan = None

        job.status = JobStatus.ASSEMBLING
        await _assemble_all_variants(job, settings, template=template, layout_plan=layout_plan)

        job.status = JobStatus.AWAITING_VARIANT_CHOICE
    except Exception as exc:  # noqa: BLE001 — любая ошибка стадии должна перевести job в failed
        logger.exception("Пайплайн упал для job %s", job.job_id)
        job.status = JobStatus.FAILED
        job.error = str(exc)


async def select_variant(job: Job, variant: str) -> None:
    """Пользователь выбрал один из трёх вариантов вёрстки на экране сравнения.

    Просто фиксирует выбор и переводит job в AWAITING_AUDIT_CHOICE — сама
    сборка уже выполнена на стадии ASSEMBLING для всех трёх вариантов сразу,
    здесь ничего пересчитывать не нужно.
    """
    job.selected_variant = variant
    job.status = JobStatus.AWAITING_AUDIT_CHOICE


async def run_audit_for_selected_variant(job: Job) -> None:
    """Запуск детерминированного аудита (POST /jobs/{id}/audit) для уже
    выбранного пользователем варианта. Ничего не делает, если вариант не
    выбран — это должно быть отсечено в API-слое до вызова."""
    settings = get_settings()
    variant = job.selected_variant
    if variant is None:
        return
    try:
        job.status = JobStatus.AUDITING
        await _run_audit_one_variant(job, settings, variant)
        job.audit_skipped = False
        job.audit_stale = False
        job.status = JobStatus.DONE
    except Exception as exc:  # noqa: BLE001
        logger.exception("Аудит упал для job %s", job.job_id)
        job.status = JobStatus.FAILED
        job.error = str(exc)


async def fix_audit_issues(job: Job, issue_ids: list[str]) -> list[str]:
    """Точечное исправление замечаний аудита в .pptx выбранного варианта
    (app/pipeline/audit_fix.py) и повторный аудит этого варианта.

    В отличие от POST /jobs/{id}/fix (полная пересборка всех вариантов),
    здесь правятся только конкретные фигуры, к которым привязаны замечания;
    выбор варианта сохраняется, job возвращается в DONE. Возвращает список
    issue_id, для которых исправление реально применилось."""
    settings = get_settings()
    variant = job.selected_variant
    if variant is None or variant not in job.variant_paths:
        return []
    wanted = set(issue_ids)
    issues = [issue for issue in job.audit_issues.get(variant, []) if issue.issue_id in wanted]
    if not issues:
        return []
    try:
        job.status = JobStatus.AUDITING
        template = _load_job_template(job, settings)
        applied = await audit_fix.apply_fixes(
            job.variant_paths[variant],
            issues,
            job.design_manifest,
            allowed_fonts=template_engine.font_families(template),
            allowed_sizes=template_engine.font_sizes(template),
        )
        job.fixed_issue_ids.extend(applied)
        await _run_audit_one_variant(job, settings, variant)
        job.status = JobStatus.DONE
        return applied
    except Exception as exc:  # noqa: BLE001
        logger.exception("Исправление замечаний упало для job %s", job.job_id)
        job.status = JobStatus.FAILED
        job.error = str(exc)
        return []


def skip_audit(job: Job) -> None:
    """Пользователь явно пропустил аудит и переходит сразу к экспорту."""
    job.audit_skipped = True
    job.status = JobStatus.DONE


async def rerun_assembly(job: Job, variant: str | None = None) -> None:
    """Перезапуск Компоновщика для POST /jobs/{id}/fix (правка ContentPlan) или
    для точечной пересборки конкретного варианта после правки дизайн-манифеста.

    `variant=None` (по умолчанию) пересобирает ВСЕ три варианта и возвращает job в
    AWAITING_VARIANT_CHOICE заново — пользовательский выбор варианта и решение
    об аудите нужно повторить заново. `variant="variant_b"` и т.п. пересобирает
    только один вариант (при активной fix-петле после аудита выбранного
    варианта) и оставляет job в AWAITING_AUDIT_CHOICE, т.к. выбор уже сделан.

    Ограничение текущей версии: список issue_id из запроса пока не
    применяется точечно (нет привязки AuditIssue к конкретному
    content-блоку) — просто пересобирает колоду(и) заново.
    Честно описано в README/AUDIT.md как TODO.
    """
    settings = get_settings()
    try:
        job.status = JobStatus.ASSEMBLING
        if variant is None:
            await _assemble_all_variants(job, settings)
            job.selected_variant = None
            job.audit_issues = {}
            job.audit_skipped = False
            job.status = JobStatus.AWAITING_VARIANT_CHOICE
        else:
            await asyncio.to_thread(_assemble_one_variant, job, settings, variant)
            job.audit_issues.pop(variant, None)
            job.status = JobStatus.AWAITING_AUDIT_CHOICE
    except Exception as exc:  # noqa: BLE001
        logger.exception("Пере-сборка упала для job %s", job.job_id)
        job.status = JobStatus.FAILED
        job.error = str(exc)


def _persist_json(path: Path, payload) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=1), encoding="utf-8")
    return path


def _persist_template_json(job: Job, settings, template) -> Path:
    return _persist_json(settings.outputs_dir / job.job_id / "template.json", template)


def _load_job_template(job: Job, settings):
    if job.template_json_path and Path(job.template_json_path).exists():
        return json.loads(Path(job.template_json_path).read_text(encoding="utf-8"))
    template = template_engine.build_template_json(job.template_path)
    job.template_json_path = str(_persist_template_json(job, settings, template))
    return template


def _load_layout_plan(job: Job, settings, template=None):
    """План по макетам старого планировщика; строится лениво, если его ещё нет."""
    path = settings.outputs_dir / job.job_id / "layout_plan.json"
    if path.exists():
        return json.loads(path.read_text(encoding="utf-8"))
    if job.content_plan is None:
        raise RuntimeError("План по макетам шаблона не найден — запустите генерацию заново")
    template = template or _load_job_template(job, settings)
    model = template_engine.BackendModel(settings.LLM_BASE_URL, settings.LLM_MODEL_NAME, settings.LLM_API_KEY)
    brief_for_planner = template_engine.planner_brief(job.brief, job.content_plan)
    _outline, layout_plan = template_engine.plan_presentation(template, brief_for_planner, job.purpose, len(job.content_plan.slides), model)
    _persist_json(path, layout_plan)
    return layout_plan


def _load_structured(job: Job, settings):
    path = settings.outputs_dir / job.job_id / "structured.json"
    if path.exists():
        return json.loads(path.read_text(encoding="utf-8"))
    if job.content_plan is None:
        return None
    structured = [dramaturgy.structure_fallback(i, spec, len(job.content_plan.slides)) for i, spec in enumerate(job.content_plan.slides)]
    _persist_json(path, structured)
    return structured


async def _assemble_all_variants(job: Job, settings, template=None, layout_plan=None) -> None:
    template = template or _load_job_template(job, settings)
    for variant in VALID_VARIANTS:
        await asyncio.to_thread(_assemble_one_variant, job, settings, variant, template, layout_plan)


def _assemble_one_variant(job: Job, settings, variant: str, template=None, layout_plan=None) -> None:
    """Один вариант вёрстки: драматургия (композиции по правилам + плотность
    варианта) → рендер. Если драматургия не справилась — старый путь: план
    модели по макетам → стратегия варианта → рендер."""
    template = template or _load_job_template(job, settings)
    job_dir = settings.outputs_dir / job.job_id
    output_path = job_dir / f"presentation_{variant}.pptx"

    plan = None
    description = None
    structured = _load_structured(job, settings) if job.content_plan is not None else None
    if structured is not None:
        try:
            plan = dramaturgy.build_variant_plan(template, job.content_plan, structured, variant, job.brief, template_engine._shorten)
            profile = dramaturgy.DENSITY_PROFILES[variant]
            description = {"label": profile["label"], "description": profile["description"]}
            template_engine.fill_image_slots(
                template, plan, job.template_path, job_dir / "assets", brief=job.brief, fill_optional=bool(profile["images"])
            )
            _persist_json(job_dir / f"plan_{variant}.json", plan)
            template_engine.render_variant(
                template, plan, job.template_path, output_path, job.design_manifest, job.content_plan, variant
            )
        except Exception:  # noqa: BLE001 — откатываемся на старый планировщик
            logger.exception("Драматургия не собрала вариант %s — используем план по макетам", variant)
            plan = None
    if plan is None:
        layout_plan = layout_plan or _load_layout_plan(job, settings, template)
        data_slides = template_engine.data_slide_indices(job.content_plan, layout_plan) if job.content_plan else set()
        plan = template_engine.derive_variant_plan(template, layout_plan, variant, data_slides)
        template_engine.supplement_plan_text(template, plan, job.content_plan)
        template_engine.fill_image_slots(
            template, plan, job.template_path, job_dir / "assets", brief=job.brief, fill_optional=(variant == "variant_c")
        )
        _persist_json(job_dir / f"plan_{variant}.json", plan)
        template_engine.render_variant(
            template, plan, job.template_path, output_path, job.design_manifest, job.content_plan, variant
        )
        description = {"label": VARIANT_STRATEGIES[variant]["label"], "description": VARIANT_STRATEGIES[variant]["description"]}
    _register_variant(job, variant, plan, description, output_path)


def _register_variant(job: Job, variant: str, plan: dict, description, output_path: Path) -> None:
    job.variant_plans[variant] = list(template_engine.plan_signature(plan))
    job.variant_kinds[variant] = [str(s.get("composition_kind") or "") for s in plan.get("slides") or []]
    job.variant_sources[variant] = [int(s.get("source_index", i)) for i, s in enumerate(plan.get("slides") or [])]
    if description is not None:
        job.variant_descriptions[variant] = description
    job.variant_paths[variant] = str(output_path)
    job.export_paths["pptx"] = str(output_path)  # последний собранный вариант — путь для /export по умолчанию


SLIDE_EDIT_ACTIONS = {"recompose", "take_from", "rewrite"}


async def edit_slide(job: Job, variant: str, index: int, action: str, source: str | None = None) -> dict:
    """Ручная доработка одного слайда уже собранного варианта.

    - ``recompose`` — другая композиция для слайда: план варианта считается
      заново с жёстким запретом уже показанных композиций на этом индексе, и
      из него берётся только этот слайд (остальные не трогаем);
    - ``take_from`` — слайд из другого варианта (``source``) как есть;
    - ``rewrite`` — модель переписывает тексты слайда, слайд собирается заново.

    Вариант рендерится в тот же файл; замечания аудита по нему сбрасываются,
    а если аудит уже проводился — помечается устаревшим (``audit_stale``).
    Статус задачи на время операции — ASSEMBLING, затем возвращается прежний.
    """
    settings = get_settings()
    job_dir = settings.outputs_dir / job.job_id
    plan_path = job_dir / f"plan_{variant}.json"
    if not plan_path.exists() or job.content_plan is None:
        raise ValueError("План варианта не найден — вариант собран старым планировщиком или задача устарела")
    plan = json.loads(plan_path.read_text(encoding="utf-8"))
    if not (0 <= index < len(plan["slides"])):
        raise ValueError(f"Нет слайда с индексом {index}")
    if plan.get("density") not in dramaturgy.DENSITY_PROFILES:
        raise ValueError("Слайд собран по макетам старого планировщика — доработка недоступна")
    previous_status = job.status
    job.status = JobStatus.ASSEMBLING
    try:
        template = _load_job_template(job, settings)
        structured = _load_structured(job, settings) or []
        current_id = plan["slides"][index]["variant"]
        src = int(plan["slides"][index].get("source_index", index))  # индекс в плане содержания
        result: dict = {"action": action, "index": index, "before": current_id}

        def _by_source(slides: list[dict]) -> dict | None:
            return next((sl for sl in slides if int(sl.get("source_index", -1)) == src), None)
        if action == "take_from":
            if not source or source == variant:
                raise ValueError("Укажите другой вариант-источник")
            source_path = job_dir / f"plan_{source}.json"
            if not source_path.exists():
                raise ValueError(f"План варианта {source} не найден")
            source_plan = json.loads(source_path.read_text(encoding="utf-8"))
            source_slide = _by_source(source_plan["slides"])
            if source_slide is None:
                raise ValueError("В варианте-источнике нет этого слайда (он собран без разделителей)")
            plan["slides"][index] = json.loads(json.dumps(source_slide))
        else:
            if action == "rewrite":
                model = template_engine.BackendModel(settings.LLM_BASE_URL, settings.LLM_MODEL_NAME, settings.LLM_API_KEY)
                previous = structured[src] if src < len(structured) else None
                new_struct = await asyncio.to_thread(dramaturgy.restructure_slide, job.content_plan, job.brief, model, src, previous)
                while len(structured) <= src:
                    structured.append(dramaturgy.structure_fallback(len(structured), job.content_plan.slides[len(structured)], len(job.content_plan.slides)))
                structured[src] = new_struct
                _persist_json(job_dir / "structured.json", structured)
                exclude = None
            else:
                history = job.slide_edits.setdefault(variant, {}).setdefault(str(index), [])
                if current_id not in history:
                    history.append(current_id)
                exclude = {src: set(history)}
            fresh = await asyncio.to_thread(
                dramaturgy.build_variant_plan, template, job.content_plan, structured, plan["density"], job.brief, template_engine._shorten, exclude
            )
            new_slide = _by_source(fresh["slides"]) or fresh["slides"][min(index, len(fresh["slides"]) - 1)]
            if action == "recompose" and new_slide["variant"] == current_id:
                # Все композиции уже показаны — начинаем круг заново, исключая только текущую
                job.slide_edits[variant][str(index)] = [current_id]
                fresh = await asyncio.to_thread(
                    dramaturgy.build_variant_plan, template, job.content_plan, structured, plan["density"], job.brief, template_engine._shorten, {src: {current_id}}
                )
                new_slide = _by_source(fresh["slides"]) or fresh["slides"][min(index, len(fresh["slides"]) - 1)]
            if action == "recompose":
                job.slide_edits[variant][str(index)].append(new_slide["variant"])
            plan["slides"][index] = new_slide
        profile = dramaturgy.DENSITY_PROFILES[plan["density"]]
        await asyncio.to_thread(
            template_engine.fill_image_slots, template, plan, job.template_path, job_dir / "assets", job.brief, bool(profile["images"])
        )
        _persist_json(plan_path, plan)
        output_path = job_dir / f"presentation_{variant}.pptx"
        await asyncio.to_thread(
            template_engine.render_variant, template, plan, job.template_path, output_path, job.design_manifest, job.content_plan, variant
        )
        _register_variant(job, variant, plan, None, output_path)
        if job.audit_issues.pop(variant, None) is not None and previous_status == JobStatus.DONE and not job.audit_skipped:
            job.audit_stale = True
        result["after"] = plan["slides"][index]["variant"]
        result["kind"] = plan["slides"][index].get("composition_kind")
        return result
    finally:
        job.status = previous_status


async def _run_audit_one_variant(job: Job, settings, variant: str) -> None:
    output_path = Path(job.variant_paths[variant])
    template = _load_job_template(job, settings)
    issues = audit.audit_presentation(
        output_path,
        manifest=job.design_manifest,
        allowed_fonts=template_engine.font_families(template),
        allowed_sizes=template_engine.font_sizes(template),
    )
    issues.extend(await _run_vlm_audit_safely(output_path, settings))
    job.audit_issues[variant] = issues


async def _run_vlm_audit_safely(output_path, settings) -> list[AuditIssue]:
    """VLM-аудит контента (Приложение 1, "Валидация контента") — недетерминированный,
    требует живого VLM-эндпоинта. Деградируем мягко: если выключен в настройках
    или эндпоинт недоступен/падает, пайплайн НЕ должен падать целиком —
    остальные (детерминированные) проверки остаются полезны сами по себе."""
    if not settings.VLM_AUDIT_ENABLED:
        return []
    try:
        return await audit.audit_content_with_vlm(output_path)
    except Exception as exc:  # noqa: BLE001 — сбой VLM не должен ронять сборку/детерминированный аудит
        logger.warning("VLM-аудит контента недоступен для %s: %s", output_path, exc)
        return [
            AuditIssue(
                issue_id=str(uuid.uuid4()),
                slide_id="deck",
                category=AuditCategory.CONTENT,
                deterministic=False,
                description=f"VLM-аудит контента пропущен (эндпоинт недоступен): {exc}",
            )
        ]
