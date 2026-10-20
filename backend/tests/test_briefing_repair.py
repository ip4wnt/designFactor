import asyncio
import json
from pathlib import Path
from types import SimpleNamespace

import pytest
from plan_factory import good_plan, to_plan

from app import skills
from app.pipeline import briefing
from app.pipeline.plan_repair import repair_feedback, repair_loop
from app.pipeline.plan_validator import validate_plan
from app.schemas.job import Job

SKILLS_DIR = Path(__file__).resolve().parents[1] / "skills"
LONG_COVER = "Масштабирование корпоративного мессенджера"  # 42 знака при лимите 30


def plan_out(cover_title: str = "Продажи 2026", story: str = "история варианта") -> briefing.PlanOut:
    plan = good_plan()
    for i, s in enumerate(plan["slides"][1:-1], start=1):
        s["content_blocks"] = [
            {"type": "bullets", "bullets": [f"Первый тезис {i} про рынок", f"Второй тезис {i} про клиентов", f"Третий тезис {i} про продукт"]}
        ]
    plan["slides"][0]["title"] = cover_title
    data = {"brief": "b", "purpose": "p", "story": story, "slides": plan["slides"]}
    return briefing.PlanOut.model_validate_json(json.dumps(data, ensure_ascii=False))


@pytest.fixture
def variant(monkeypatch):
    """_generate_variant с подставным _structured: ответы из очереди, вызовы пишутся в calls."""
    monkeypatch.setattr(skills, "skills_dir", lambda: SKILLS_DIR)
    skills.load_skill.cache_clear()

    def run(responses, rounds: int = 1):
        queue, calls = list(responses), []

        async def fake_structured(system, user, model, **kwargs):
            calls.append({"system": system, "user": user, "model": model, **kwargs})
            item = queue.pop(0)
            if isinstance(item, Exception):
                raise item
            return item

        monkeypatch.setattr(briefing, "_structured", fake_structured)
        monkeypatch.setattr(briefing, "get_settings", lambda: SimpleNamespace(LLM_PLANNING_TIMEOUT_S=10.0, PLAN_REPAIR_ROUNDS=rounds, PLAN_REPAIR_TIMEOUT_S=10.0))
        job = Job(job_id="j", brief="бриф", purpose="цель", template_path="t.pptx")
        inputs = briefing._plan_inputs(job, 5, 5)
        result = asyncio.run(briefing._generate_variant(job, "A", inputs, 5, 5))
        return result, calls

    return run


def test_variant_with_error_is_repaired_and_keeps_story(variant):
    result, calls = variant([plan_out(LONG_COVER), plan_out("Продажи 2026", story="история варианта")])
    assert len(calls) == 2
    assert "Предыдущий вариант плана" in calls[1]["user"] and "максимум 30" in calls[1]["user"]
    assert calls[1]["model"] is briefing.PlanOut
    assert calls[1]["temperature"] < calls[0]["temperature"]
    assert result.plan.slides[0].title == "Продажи 2026"
    assert result.story == "история варианта"


def test_clean_variant_makes_single_call(variant):
    result, calls = variant([plan_out()])
    assert len(calls) == 1
    assert result.key == "A"


def test_repair_disabled_by_setting(variant):
    result, calls = variant([plan_out(LONG_COVER)], rounds=0)
    assert len(calls) == 1
    assert result.plan.slides[0].title == LONG_COVER


def test_failed_repair_keeps_original_variant(variant):
    result, calls = variant([plan_out(LONG_COVER), TimeoutError("дедлайн")])
    assert len(calls) == 2
    assert result.plan.slides[0].title == LONG_COVER


def test_repair_loop_skips_plan_that_cannot_be_checked():
    called = []

    async def request_fix(feedback):
        called.append(feedback)
        return plan_out()

    def broken_as_plan(_):
        raise ValueError("не план")

    original = plan_out(LONG_COVER)
    result = asyncio.run(
        repair_loop(original, as_plan=broken_as_plan, request_fix=request_fix, slide_count=5, slide_count_max=5, rounds=1)
    )
    assert result is original and called == []


def test_feedback_hides_divider_warnings_that_would_change_slide_count():
    plan = to_plan(good_plan(content_slides=8))  # 10 слайдов без разделителей -> предупреждение divider_count
    issues = validate_plan(plan, 10, 10)
    assert any(i.code == "divider_count" for i in issues)
    plan_dict = good_plan(content_slides=8)
    plan_dict["slides"][0]["title"] = LONG_COVER
    plan = to_plan(plan_dict)
    feedback = repair_feedback(plan, plan, validate_plan(plan, 10, 10))
    assert "разделител" not in feedback and "максимум 30" in feedback


def test_slow_repair_is_cut_off_by_timeout():
    async def slow_fix(feedback):
        await asyncio.sleep(5)
        return plan_out()

    original = plan_out(LONG_COVER)
    result = asyncio.run(
        repair_loop(
            original,
            as_plan=lambda o: briefing.ContentPlan(brief=o.brief, purpose=o.purpose, slides=o.slides),
            request_fix=slow_fix,
            slide_count=5,
            slide_count_max=5,
            rounds=1,
            timeout=0.05,
        )
    )
    assert result is original
