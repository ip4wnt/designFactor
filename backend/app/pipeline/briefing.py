"""Брифование: главная идея, три варианта плана выступления, перегенерация
слайда плана и аудит смысла.

Все тексты промптов лежат в skills/*.yaml (main_idea, plan_variants, slide_regen,
goal_alignment), здесь — только сборка запросов и разбор ответов.
"""
from __future__ import annotations

import asyncio
import json
import logging
import re
from typing import Any, Literal

from pydantic import Field, ValidationError, create_model, model_validator

from app.config import get_settings
from app.llm_client import LLMClient
from app.pipeline.docx_tables import bind_source_tables, table_card
from app.pipeline.plan_repair import repair_loop
from app.pipeline.table_analysis import analysis_prompt, prepare_table_analysis
from app.schemas.content_plan import ContentPlan, ContentBlock, ContentBlockType, SlideSpec
from app.schemas.job import Job, PlanVariant
from app.schemas.strict_model import StrictModel
from app.skills import load_skill, skill_text

logger = logging.getLogger(__name__)

MATERIALS_IN_PROMPT = 9000  # знаков материалов, которые уходят в один запрос
PLAN_KEYS = ("A", "B", "C")


# --- схемы ответов модели -------------------------------------------------------

class Clarification(StrictModel):
    question: str
    options: list[str]


class IdeaResult(StrictModel):
    idea: str
    clarification: Clarification | None


class PlanOut(StrictModel):
    brief: str
    purpose: str
    story: str
    slides: list[SlideSpec]


class _PlanningTextBlock(ContentBlock):
    type: Literal[ContentBlockType.BULLETS, ContentBlockType.TEXT]


class _PlanningSlide(SlideSpec):
    @model_validator(mode="before")
    @classmethod
    def source_cells_belong_to_server(cls, value):
        """A source reference owns the cells, including an empty model placeholder.

        Do this only in LLM planning responses, before ContentBlock validation.
        The normal SlideSpec and ContentBlock contracts remain strict.
        """
        if isinstance(value, dict) and value.get("source_table_id") and isinstance(value.get("content_blocks"), list):
            return {**value, "content_blocks": [block for block in value["content_blocks"]
                if not (isinstance(block, dict) and block.get("type") in {"table", "chart"})]}
        return value


def _planning_slide_model(tables):
    available = tuple(dict.fromkeys(table.table_id for table in tables if table.available and table.data is not None))
    # The provider and the local parser both reject invented or oversized references.
    selection = Literal[(*available, None)]
    fields = {"source_table_id": (selection, Field(default=None,
        description="Только ID доступного результата из каталога либо null. Таблицу или рекомендованную диаграмму вставляет сервер."))}
    if tables:
        # Uploaded data must come from the registry/tools, never hand-written LLM cells.
        fields["content_blocks"] = (list[_PlanningTextBlock], Field(default_factory=list,
            description="Только текст и буллеты. Таблицу/диаграмму по source_table_id добавляет сервер; не сочиняй числа."))
    return create_model("PlanningSlide", __base__=_PlanningSlide, **fields)


def _plan_response_model(tables):
    slide_model = _planning_slide_model(tables)
    return create_model("PlanOut", __base__=PlanOut, slides=(list[slide_model], ...))


class AlignmentCheck(StrictModel):
    ok: bool
    comment: str


class SlideRemark(StrictModel):
    slide: int
    remark: str


class AlignmentResult(StrictModel):
    idea_through_story: AlignmentCheck
    final_action: AlignmentCheck
    slide_remarks: list[SlideRemark]


# --- общие помощники --------------------------------------------------------------

def deyo(text: str) -> str:
    return text.replace("ё", "е").replace("Ё", "Е")


def _deyo_model(model):
    return type(model).model_validate_json(deyo(model.model_dump_json()), strict=True)


def snippet(materials: str, limit: int = MATERIALS_IN_PROMPT) -> str:
    materials = (materials or "").strip()
    if not materials:
        return "(материалов нет — опирайся только на описание автора)"
    if len(materials) <= limit:
        return materials
    return materials[:limit].rsplit("\n", 1)[0] + "\n…(материалы сокращены)"


def _client(timeout: float | None = None) -> LLMClient:
    settings = get_settings()
    return LLMClient(
        base_url=settings.LLM_BASE_URL,
        model_name=settings.LLM_MODEL_NAME,
        api_key=settings.LLM_API_KEY,
        max_retries=0,
        total_timeout=timeout or settings.LLM_PLANNING_TIMEOUT_S,
    )


async def _structured(system: str, user: str, model, *, max_tokens: int, temperature: float = 0.3, attempts: int = 2, timeout: float | None = None, validate=None):
    last: Exception | None = None
    async with _client(timeout) as client:
        for attempt in range(attempts):
            try:
                result = await client.chat_structured(
                    system_prompt=system,
                    user_input=user,
                    output_model=model,
                    reasoning_effort="none",
                    temperature=temperature,
                    max_completion_tokens=max_tokens,
                )
                result = _deyo_model(result)
                if validate is not None:
                    validate(result)
                return result
            except (ValidationError, ValueError) as exc:
                last = exc
                logger.warning("Ответ модели не разобран (%s, попытка %d/%d): %s", model.__name__, attempt + 1, attempts, str(exc)[:300])
                user += "\n\nИсправь ответ по схеме. Ошибка проверки: " + str(exc)[:500]
    raise RuntimeError("Модель вернула ответ в неверном формате") from last


def _text(value: str | None, fallback: str = "не указано") -> str:
    value = (value or "").strip()
    return value or fallback


# --- главная идея -----------------------------------------------------------------

async def generate_main_idea(
    topic: str,
    materials: str,
    audience: str,
    goal: str,
    clarification: str = "",
    previous_idea: str = "",
) -> IdeaResult:
    skill = load_skill("main_idea")
    clar = f"\nОтвет на уточнение: {clarification.strip()}\n" if clarification.strip() else ""
    variation = (
        f"\nПредыдущая формулировка: {previous_idea.strip()}\nПредложи другую формулировку той же сути или с другим акцентом; уточняющий вопрос не задавай.\n"
        if previous_idea.strip() else ""
    )
    user = skill["user_prompt_template"].format(
        topic=_text(topic), materials=snippet(materials, 6000), audience=_text(audience), goal=_text(goal),
        clarification=clar, variation=variation,
    )
    result = await _structured(skill["system_prompt"], user, IdeaResult, max_tokens=700, temperature=0.5 if previous_idea else 0.3, timeout=60)
    if clarification.strip() or previous_idea.strip():
        result = IdeaResult(idea=result.idea, clarification=None)
    if result.clarification is not None:
        opts = [o.strip() for o in result.clarification.options if o.strip()][:3]
        if not result.clarification.question.strip() or len(opts) < 2:
            result = IdeaResult(idea=result.idea, clarification=None)
        else:
            result = IdeaResult(idea=result.idea, clarification=Clarification(question=result.clarification.question.strip(), options=opts))
    return result


# --- три варианта плана -----------------------------------------------------------

def _plan_inputs(job: Job, slide_count: int, slide_count_max: int | None) -> dict[str, Any]:
    if slide_count_max is not None and slide_count_max > slide_count:
        count_text = f"от {slide_count} до {slide_count_max} (выбери число внутри диапазона по объему материала)"
    else:
        count_text = str(slide_count)
    return {
        "idea": _text(job.idea),
        "audience": _text(job.audience),
        "goal": _text(job.goal),
        "tone": _text(job.tone, "деловой и понятный"),
        "must": _text(job.must_have, "ничего особенного"),
        "topic": _text(job.brief),
        # Registry cards stay outside the legacy leading-text limit, even for late tables.
        "materials": snippet(job.materials_text) + "\n\nРазрешенные source_table_id для вставки: " +
                     json.dumps([t.table_id for t in job.source_tables if t.available and t.data is not None], ensure_ascii=False) +
                     ". Если список пуст, source_table_id=null на всех слайдах. "
                     "Недоступные таблицы описаны только для контекста; они не являются готовыми результатами анализа.\n\nКаталог исходных таблиц:\n" +
                     ("\n\n".join(table_card(table) for table in job.source_tables) or "(нет)") + analysis_prompt(job.table_analysis),
        "slide_count": count_text,
    }


def _required_data_kinds(job: Job) -> set[str]:
    if not job.source_tables:
        return set()
    return {"table" if m[2].startswith(("таблиц", "table")) else "chart"
            for m in re.finditer(r"(без\s+|не\s+(?:нужны|нужна|используй|использовать|добавляй|добавлять)\s+)?"
                                 r"(таблиц\w*|график\w*|диаграмм\w*|charts?|tables?)", job.must_have.casefold()) if not m[1]}


def _validate_plan_data(out: PlanOut, job: Job) -> None:
    bind_source_tables(out.slides, job.source_tables)
    present = {block.type.value for slide in out.slides for block in slide.content_blocks}
    missing = _required_data_kinds(job) - present - _unavailable_required_kinds(job)
    if missing:
        raise ValueError("Автор явно запросил " + ", ".join(sorted(missing)) +
                         "; выбери подходящие source_table_id из каталога для внутренних слайдов.")


def _unavailable_required_kinds(job: Job) -> set[str]:
    available = {"chart" if t.chart is not None else "table" for t in job.source_tables if t.available and t.data is not None}
    return _required_data_kinds(job) - available


async def _generate_variant(job: Job, key: str, inputs: dict[str, Any], slide_count: int, slide_count_max: int | None, nudge: str = "") -> PlanVariant:
    unavailable = _unavailable_required_kinds(job)
    if unavailable:
        # Анализ не успел или не смог подготовить запрошенные графики/таблицы: план все равно строим,
        # а нехватку честно показываем в data_gap и в заметках анализа, вместо отказа целиком.
        labels = ", ".join(sorted("графики" if k == "chart" else "таблицы" for k in unavailable))
        note = (f"Запрошены {labels}, но анализ данных не подготовил готовые результаты; "
                "план построен с пометками о недостающих данных. Нажмите «Перегенерировать», чтобы повторить анализ.")
        if job.table_analysis is not None and note not in job.table_analysis.notes:
            job.table_analysis.notes.append(note)
        nudge = (nudge + "\n" if nudge else "") + (
            f"Автор просил {labels}, но готовых данных для них нет. Не выдумывай числа и не вставляй "
            "source_table_id без каталога: на слайдах, где нужны эти данные, заполни data_gap "
            "описанием, какой расчет нужен.")
    angles = load_skill("plan_variants")["angles"]
    angle = angles[key]
    base = load_skill("content_planning")
    variants_skill = load_skill("plan_variants")
    system = base["system_prompt"].rstrip() + "\n\n" + variants_skill["extra_rules"].strip()
    user = variants_skill["user_prompt_template"].format(
        angle_instruction="Вариант плана: «" + angle["name"] + "».\n" + angle["instruction"].strip() + (("\n" + nudge) if nudge else ""),
        extra="",
        **inputs,
    )
    top = max(slide_count, slide_count_max or 0)
    out: PlanOut = await _structured(
        system, user, _plan_response_model(job.source_tables), max_tokens=min(12000, 1500 + 700 * top), temperature=0.55, attempts=2,
        timeout=get_settings().LLM_PLANNING_TIMEOUT_S * 1.5,
        validate=lambda out: _validate_plan_data(out, job),
    )

    async def request_fix(feedback: str) -> PlanOut:
        # Температура ниже, чем у генерации: правка должна точечно убрать нарушения, а не переписать рассказ варианта.
        return await _structured(
            system, user + feedback, PlanOut, max_tokens=min(12000, 1500 + 700 * top), temperature=0.2, attempts=1,
            timeout=get_settings().LLM_PLANNING_TIMEOUT_S * 1.5,
        )

    out = await repair_loop(
        out,
        as_plan=lambda candidate: ContentPlan(brief=candidate.brief, purpose=candidate.purpose, slides=candidate.slides),
        request_fix=request_fix,
        slide_count=slide_count,
        slide_count_max=slide_count_max,
        rounds=get_settings().PLAN_REPAIR_ROUNDS,
        timeout=get_settings().PLAN_REPAIR_TIMEOUT_S,
        label=f"План {key}",
    )
    plan = ContentPlan(brief=out.brief, purpose=out.purpose, slides=out.slides)  # проверяет уникальность slide_id
    return PlanVariant(key=key, name=angle["name"], description=angle["description"], story=out.story.strip(), plan=plan)


async def generate_plan_variants(job: Job, slide_count: int, slide_count_max: int | None) -> dict[str, PlanVariant]:
    inputs = _plan_inputs(job, slide_count, slide_count_max)
    results = await asyncio.gather(
        *[_generate_variant(job, key, inputs, slide_count, slide_count_max) for key in PLAN_KEYS],
        return_exceptions=True,
    )
    variants: dict[str, PlanVariant] = {}
    errors: list[str] = []
    for key, result in zip(PLAN_KEYS, results):
        if isinstance(result, Exception):
            logger.warning("Вариант плана %s не получился: %s", key, result)
            errors.append(f"{key}: {result}")
        else:
            variants[key] = result
    if not variants:
        raise RuntimeError("Не удалось построить план выступления. " + "; ".join(errors)[:300])
    return variants


async def regenerate_plan_variant(job: Job, key: str, slide_count: int, slide_count_max: int | None) -> PlanVariant:
    if any(t.data is not None and not t.parent_table_ids for t in job.source_tables):
        await prepare_table_analysis(job, get_settings())
    old = job.plan_variants.get(key)
    nudge = ""
    if old is not None:
        titles = "; ".join(s.title for s in old.plan.slides)
        nudge = "Это повторная генерация этого варианта. Предложи другие заголовки, другой порядок аргументов и другие примеры. Прежние заголовки (так больше не писать): " + titles
    return await _generate_variant(job, key, _plan_inputs(job, slide_count, slide_count_max), slide_count, slide_count_max, nudge)


# --- один слайд плана -------------------------------------------------------------

def slide_as_text(index: int, spec: SlideSpec) -> str:
    lines = [f"Слайд {index + 1}. {spec.title} — {spec.purpose}"]
    for block in spec.content_blocks:
        if block.type == ContentBlockType.BULLETS:
            lines += [f"  • {item}" for item in block.bullets or []]
        elif block.type == ContentBlockType.TEXT:
            lines.append(f"  {block.text}")
        elif block.type == ContentBlockType.CHART:
            lines.append("  [график]")
        elif block.type == ContentBlockType.TABLE:
            lines.append("  [таблица]")
    return "\n".join(lines)


async def regenerate_slide(job: Job, plan: ContentPlan, index: int, instruction: str = "") -> SlideSpec:
    skill = load_skill("slide_regen")
    spec = plan.slides[index]
    outline = "\n".join(f"{i + 1}. {s.title}" for i, s in enumerate(plan.slides))
    user = skill["user_prompt_template"].format(
        idea=_text(job.idea), goal=_text(job.goal), audience=_text(job.audience), tone=_text(job.tone, "деловой и понятный"),
        outline=outline, number=index + 1, slide=slide_as_text(index, spec),
        instruction=(f"Инструкция автора: {instruction.strip()}\n" if instruction.strip() else ""),
    )
    user += "\nКаталог исходных таблиц:\n" + "\n\n".join(table_card(t) for t in job.source_tables)
    user += analysis_prompt(job.table_analysis)
    def validate_slide(slide):
        slide.source_table_id = slide.source_table_id or spec.source_table_id
        if slide.source_table_id and index in {0, len(plan.slides) - 1}:
            raise ValueError("Обложка и финал не могут содержать исходную таблицу")
        bind_source_tables([slide], job.source_tables, single_slide=True)
    response_model = _planning_slide_model([] if index in {0, len(plan.slides) - 1} else job.source_tables)
    new: SlideSpec = await _structured(skill["system_prompt"], user, response_model, max_tokens=1800, temperature=0.8, timeout=90,
                                       validate=validate_slide)
    keep_data = [b for b in spec.content_blocks if b.type in (ContentBlockType.CHART, ContentBlockType.TABLE)]
    if keep_data and not any(b.type in (ContentBlockType.CHART, ContentBlockType.TABLE) for b in new.content_blocks):
        new = new.model_copy(update={"content_blocks": list(new.content_blocks) + keep_data})
    return new.model_copy(update={"slide_id": spec.slide_id, "purpose": new.purpose or spec.purpose})


# --- аудит смысла -----------------------------------------------------------------

async def audit_alignment(job: Job) -> dict[str, Any] | None:
    """Связь колоды с главной идеей и целью. Возвращает None, если модель недоступна —
    это не должно ломать остальной аудит."""
    if job.content_plan is None or not (job.idea or job.goal):
        return None
    skill = load_skill("goal_alignment")
    slides = "\n".join(slide_as_text(i, s) for i, s in enumerate(job.content_plan.slides))
    user = skill["user_prompt_template"].format(idea=_text(job.idea), goal=_text(job.goal), audience=_text(job.audience), slides=slides)
    try:
        result: AlignmentResult = await _structured(skill["system_prompt"], user, AlignmentResult, max_tokens=900, temperature=0.2, timeout=60)
    except Exception as exc:  # noqa: BLE001
        logger.warning("Аудит смысла недоступен: %s", exc)
        return None
    total = len(job.content_plan.slides)
    return {
        "idea_through_story": result.idea_through_story.model_dump(),
        "final_action": result.final_action.model_dump(),
        "slide_remarks": [r.model_dump() for r in result.slide_remarks if 1 <= r.slide <= total][:4],
    }
