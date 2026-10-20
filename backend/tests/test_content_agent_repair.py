import asyncio
from pathlib import Path
from types import SimpleNamespace

import pytest
from plan_factory import good_plan, to_plan

from app.pipeline import content_agent
from app.schemas.content_plan import ContentPlan

SKILLS_DIR = Path(__file__).resolve().parents[1] / "skills"


def valid_plan(content_slides: int = 3) -> ContentPlan:
    plan = good_plan(content_slides)
    for i, s in enumerate(plan["slides"][1:-1], start=1):
        s["content_blocks"] = [
            {"type": "bullets", "bullets": [f"Первый тезис {i} про рынок", f"Второй тезис {i} про клиентов", f"Третий тезис {i} про продукт"]}
        ]
    return to_plan(plan)


def broken_plan(extra_errors: int = 0) -> ContentPlan:
    """Заголовок длиннее 60 знаков (+ слайды без содержимого для усиления)."""
    plan = valid_plan().model_dump(mode="json")
    plan["slides"][1]["title"] = "Очень длинный заголовок, который явно не помещается в лимит в шестьдесят знаков"
    for i in range(extra_errors):
        plan["slides"][2 + i]["content_blocks"] = []
    return to_plan(plan)


@pytest.fixture
def agent(monkeypatch):
    """content_agent с подставным LLM-клиентом: ответы берутся из очереди, вызовы пишутся в calls."""

    class Env:
        responses: list = []
        calls: list = []

    def settings(rounds: int = 1):
        return SimpleNamespace(
            LLM_BASE_URL="http://llm",
            LLM_MODEL_NAME="m",
            LLM_API_KEY=None,
            LLM_PLANNING_TIMEOUT_S=10.0,
            PLAN_REPAIR_ROUNDS=rounds,
            PLAN_REPAIR_TIMEOUT_S=10.0,
            SKILLS_DIR=SKILLS_DIR,
        )

    class FakeClient:
        def __init__(self, **kwargs):
            pass

        async def __aenter__(self):
            return self

        async def __aexit__(self, *exc):
            return None

        async def chat_structured(self, **kwargs):
            Env.calls.append(kwargs)
            item = Env.responses.pop(0)
            if isinstance(item, Exception):
                raise item
            return item

    monkeypatch.setattr(content_agent, "LLMClient", FakeClient)
    Env.responses, Env.calls = [], []

    def run(responses, rounds: int = 1):
        Env.responses, Env.calls = list(responses), []
        monkeypatch.setattr(content_agent, "get_settings", lambda: settings(rounds))
        plan = asyncio.run(content_agent.generate_content_plan("бриф", "цель", 5, slide_count_max=5))
        return plan, Env.calls

    return run


def test_valid_plan_needs_single_call(agent):
    plan, calls = agent([valid_plan()])
    assert len(calls) == 1
    assert plan.slides[1].title.startswith("Вывод номер 1")


def test_invalid_plan_is_repaired_with_issue_list(agent):
    plan, calls = agent([broken_plan(), valid_plan()])
    assert len(calls) == 2
    repair_input = calls[1]["user_input"]
    assert "Предыдущий вариант плана" in repair_input
    assert "Заголовок" in repair_input and "максимум 60" in repair_input
    assert len(plan.slides[1].title) <= 60


def test_rounds_zero_disables_repair(agent):
    plan, calls = agent([broken_plan()], rounds=0)
    assert len(calls) == 1
    assert len(plan.slides[1].title) > 60


def test_failed_repair_keeps_original_plan(agent):
    original = broken_plan()
    plan, calls = agent([original, TimeoutError("дедлайн")])
    assert len(calls) == 2
    assert plan == original


def test_worse_repair_is_discarded(agent):
    original = broken_plan()
    worse = broken_plan(extra_errors=2)
    plan, _ = agent([original, worse])
    assert plan == original


def test_repair_stops_after_configured_rounds(agent):
    plan, calls = agent([broken_plan(), broken_plan(), broken_plan()], rounds=1)
    assert len(calls) == 2


def test_repair_that_introduces_new_error_is_discarded(agent):
    # правка исправила обложку, но раздула колоду до 7 слайдов при запрошенных 5 — это новая ошибка
    original = broken_plan()
    bloated = valid_plan(content_slides=5)
    plan, calls = agent([original, bloated])
    assert len(calls) == 2
    assert plan == original
