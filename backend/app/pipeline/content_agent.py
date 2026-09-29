"""Контент-агент: бриф + назначение -> ContentPlan.

Не знает о макетах шаблона — работает только со смыслом. Промпт живёт
отдельным yaml-файлом (backend/skills/content_planning.yaml), а не в коде,
чтобы его можно было версионировать и подбирать модель под конкретную
версию промпта без релиза backend.
"""
from __future__ import annotations

from pathlib import Path

import logging
import yaml
from pydantic import ValidationError

from app.config import get_settings
from app.llm_client import LLMClient
from app.schemas.content_plan import ContentPlan

logger = logging.getLogger(__name__)


class ContentPlanningSkill:
    """system_prompt + user_prompt_template, загруженные из yaml-скилла."""

    def __init__(self, path: Path) -> None:
        data = yaml.safe_load(path.read_text(encoding="utf-8"))
        self.version: str = data["version"]
        self.model_name: str | None = data.get("model_name")
        self.system_prompt: str = data["system_prompt"]
        self.user_prompt_template: str = data["user_prompt_template"]

    def render_user_prompt(
        self,
        brief: str,
        purpose: str,
        slide_count: int,
        case_prompt_addition: str = "",
        slide_count_max: int | None = None,
    ) -> str:
        # Диапазон «от–до»: модель сама выбирает число слайдов по объёму
        # материала, но в заданных границах (ползунки на главной).
        if slide_count_max is not None and slide_count_max > slide_count:
            count_text = f"от {slide_count} до {slide_count_max} (выбери число внутри диапазона по объёму материала)"
        else:
            count_text = str(slide_count)
        base = self.user_prompt_template.format(brief=brief, purpose=purpose, slide_count=count_text)
        if case_prompt_addition:
            # Добавка от выбранного пользователем кейса (см. app/pipeline/cases.py) —
            # влияет на тон, структуру и приоритеты слайдов, но не меняет саму
            # схему ContentPlan — она остаётся частью system_prompt.
            base += f"\n\nДополнительный ориентир от выбранного кейса использования:\n{case_prompt_addition}"
        return base


def _load_skill() -> ContentPlanningSkill:
    settings = get_settings()
    skill_path = settings.SKILLS_DIR / "content_planning.yaml"
    return ContentPlanningSkill(skill_path)


async def generate_content_plan(
    brief: str,
    purpose: str,
    slide_count: int,
    case_prompt_addition: str = "",
    slide_count_max: int | None = None,
) -> ContentPlan:
    """Вызывает LLM по скиллу content_planning и валидирует ответ как ContentPlan.

    `case_prompt_addition` — текстовая добавка к user-промпту от выбранного на
    фронте кейса-пресета («pitch», «отчёт по проекту» и т.д.) — см.
    app/pipeline/cases.py. Пустая строка, если кейс не выбран.
    """
    skill = _load_skill()
    settings = get_settings()
    async with LLMClient(
        base_url=settings.LLM_BASE_URL,
        model_name=skill.model_name or settings.LLM_MODEL_NAME,
        api_key=settings.LLM_API_KEY,
        max_retries=0,
        total_timeout=settings.LLM_PLANNING_TIMEOUT_S,
    ) as client:
        user_prompt = skill.render_user_prompt(
            brief=brief,
            purpose=purpose,
            slide_count=slide_count,
            case_prompt_addition=case_prompt_addition,
            slide_count_max=slide_count_max,
        )
        attempts = 2
        last_error: Exception | None = None
        for attempt in range(attempts):
            try:
                plan = await client.chat_structured(
                    system_prompt=skill.system_prompt,
                    user_input=user_prompt,
                    output_model=ContentPlan,
                    reasoning_effort="none",
                    max_completion_tokens=min(8192, 1024 + 512 * max(slide_count, slide_count_max or 0)),
                )
                break
            except ValidationError as exc:
                # Модель иногда отвечает не JSON, а фразой-отказом («Не могу
                # выполнить…») или обрывком — это не ошибка данных пользователя,
                # повторяем запрос с напоминанием формата (в пределах дедлайна)
                last_error = exc
                logger.warning("Контент-план не разобран (попытка %d/%d): %s", attempt + 1, attempts, str(exc)[:300])
                user_prompt = (
                    user_prompt
                    + "\n\nВАЖНО: это деловая презентация, ответ обязателен. Верни только JSON по схеме, без пояснений и отказов."
                )
        else:
            raise RuntimeError("Модель не вернула план содержания в формате JSON") from last_error

    return plan
