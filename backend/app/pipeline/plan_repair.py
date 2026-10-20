"""Цикл «проверить план → попросить модель исправить → принять, только если лучше».

Общий для всех мест, где LLM строит ContentPlan (content_agent — обычный план,
briefing — три варианта плана выступления): они различаются промптами и
моделями ответа, поэтому сам запрос к модели передаётся колбэком, а здесь —
проверка (plan_validator), решение «принять или оставить» и защита от сбоев.
"""
from __future__ import annotations

import asyncio
import logging
from collections.abc import Awaitable, Callable
from typing import TypeVar

from pydantic import BaseModel

from app.pipeline.plan_validator import PlanIssue, format_issues, validate_plan
from app.schemas.content_plan import ContentPlan

logger = logging.getLogger(__name__)

T = TypeVar("T", bound=BaseModel)


def count_errors(issues: list[PlanIssue]) -> int:
    return sum(1 for issue in issues if issue.severity == "error")


# Замечания, исправление которых меняет число слайдов. Модели они не передаются: на живом
# прогоне из-за «нужно 2–3 разделителя» правка добавила два слайда и вышла за диапазон,
# заданный пользователем (ошибка хуже исходной).
_SLIDE_COUNT_SENSITIVE = {"divider_count", "divider_in_short"}


def repair_feedback(current: BaseModel, plan: ContentPlan, issues: list[PlanIssue]) -> str:
    """Текст, который дописывается к исходному запросу: прежний ответ целиком
    (в формате ответа модели — со всеми её полями) и список нарушений."""
    shown = [issue for issue in issues if issue.code not in _SLIDE_COUNT_SENSITIVE]
    return (
        f"\n\nПредыдущий вариант плана:\n{current.model_dump_json()}\n\n"
        f"В нём нарушены правила:\n{format_issues(shown, plan)}\n\n"
        "Исправь план: устрани перечисленное, всё остальное оставь без изменений. "
        "Верни полный план JSON по схеме."
    )


async def repair_loop(
    current: T,
    *,
    as_plan: Callable[[T], ContentPlan],
    request_fix: Callable[[str], Awaitable[T]],
    slide_count: int | None,
    slide_count_max: int | None,
    rounds: int,
    timeout: float | None = 60.0,
    label: str = "План",
) -> T:
    """Возвращает `current` или его исправленную версию.

    `as_plan` приводит ответ модели к ContentPlan для проверки; `request_fix(feedback)`
    повторяет исходный запрос с добавленным `feedback` и возвращает новый ответ.
    Правка запрашивается только при ошибках (предупреждения сами повода не дают,
    но передаются модели вместе с ошибками). Сбой или превышение `timeout` секунд на
    повторный запрос и результат, не лучше исходного, оставляют `current`: правка
    качества не должна ронять job и растягивать ожидание пользователя.
    """
    for round_number in range(1, rounds + 1):
        try:
            plan = as_plan(current)
        except Exception:  # noqa: BLE001 — невалидный по схеме план проверять нечем, ошибка уйдёт вызывающему коду как раньше
            return current
        issues = validate_plan(plan, slide_count, slide_count_max)
        errors = count_errors(issues)
        if errors == 0:
            return current
        logger.info(
            "%s: %d замечаний (%d ошибок), правка %d/%d:\n%s",
            label, len(issues), errors, round_number, rounds, format_issues(issues, plan),
        )

        try:
            async with asyncio.timeout(timeout):
                candidate = await request_fix(repair_feedback(current, plan, issues))
            new_issues = validate_plan(as_plan(candidate), slide_count, slide_count_max)
        except Exception as exc:  # noqa: BLE001 — сеть, дедлайн, невалидный JSON: остаёмся на исходном плане
            logger.warning("%s: правка не удалась, оставляю исходный: %s", label, str(exc)[:300])
            return current

        new_errors = count_errors(new_issues)
        introduced = {i.code for i in new_issues if i.severity == "error"} - {i.code for i in issues if i.severity == "error"}
        if introduced or (new_errors, len(new_issues)) >= (errors, len(issues)):
            reason = f"внесла новые ошибки {sorted(introduced)}" if introduced else "не улучшила результат"
            logger.info("%s: правка %s (ошибок %d, замечаний %d), оставляю исходный", label, reason, new_errors, len(new_issues))
            return current
        logger.info("%s: правка принята: ошибок %d → %d, замечаний %d → %d", label, errors, new_errors, len(issues), len(new_issues))
        current = candidate
    return current
