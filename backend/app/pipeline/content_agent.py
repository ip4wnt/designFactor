"""Контент-агент: бриф + назначение -> ContentPlan.

Не знает о макетах шаблона — работает только со смыслом. Промпт живёт
отдельным yaml-файлом (backend/skills/content_planning.yaml), а не в коде,
чтобы его можно было версионировать и подбирать модель под конкретную
версию промпта без релиза backend.
"""
from __future__ import annotations

from pathlib import Path

import yaml

from app.config import get_settings
from app.llm_client import LLMClient
from app.schemas.content_plan import ContentPlan


class ContentPlanningSkill:
    """system_prompt + user_prompt_template, загруженные из yaml-скилла."""

    def __init__(self, path: Path) -> None:
        data = yaml.safe_load(path.read_text(encoding="utf-8"))
        self.version: str = data["version"]
        self.model_name: str | None = data.get("model_name")
        self.system_prompt: str = data["system_prompt"]
        self.user_prompt_template: str = data["user_prompt_template"]

    def render_user_prompt(self, brief: str, purpose: str, slide_count: int) -> str:
        return self.user_prompt_template.format(brief=brief, purpose=purpose, slide_count=slide_count)


def _load_skill() -> ContentPlanningSkill:
    settings = get_settings()
    skill_path = settings.SKILLS_DIR / "content_planning.yaml"
    return ContentPlanningSkill(skill_path)


async def generate_content_plan(brief: str, purpose: str, slide_count: int) -> ContentPlan:
    """Вызывает LLM по скиллу content_planning и валидирует ответ как ContentPlan."""
    skill = _load_skill()
    settings = get_settings()
    client = LLMClient(
        base_url=settings.LLM_BASE_URL,
        model_name=skill.model_name or settings.LLM_MODEL_NAME,
    )

    user_prompt = skill.render_user_prompt(brief=brief, purpose=purpose, slide_count=slide_count)
    raw_plan = await client.chat_json(skill.system_prompt, user_prompt)

    # LLM может не продублировать brief/purpose в ответе — подстрахуемся,
    # чтобы ContentPlan всегда был валиден и трассируем к исходному job.
    raw_plan.setdefault("brief", brief)
    raw_plan.setdefault("purpose", purpose)

    return ContentPlan.model_validate(raw_plan)
