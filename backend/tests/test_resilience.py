"""Самовосстановление сборки: сломанный слайд заменяется, сломанный вариант не роняет остальные."""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest import mock

from app.pipeline import resilience


def _plan(*ids: str) -> dict:
    return {
        "density": "variant_a",
        "slides": [{"variant": v, "composition_kind": "cards", "source_index": i, "content": {}} for i, v in enumerate(ids)],
    }


class RenderWithRepairTest(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.out = str(Path(self._tmp.name) / "out.pptx")

    def tearDown(self):
        self._tmp.cleanup()

    def _fake_render(self, broken: set[str]):
        def render(template, plan, template_path, output_path, manifest, content_plan, variant_key):
            bad = [s["variant"] for s in plan["slides"] if s["variant"] in broken]
            if bad:
                raise ValueError(f"Текстовый слот выходит за границы слайда:\n\nSlot: text\n{{'variant': {bad[0]!r}}}")
            Path(output_path).write_bytes(b"ok")
        return render

    def test_renders_without_changes_when_everything_is_fine(self):
        plan = _plan("cover", "cards", "final")
        with mock.patch.object(resilience.template_engine, "render_variant", self._fake_render(set())):
            notes = resilience.render_with_repair({}, plan, "t.pptx", self.out, None, None, "variant_a", None)
        self.assertEqual(notes, [])
        self.assertEqual([s["variant"] for s in plan["slides"]], ["cover", "cards", "final"])

    def test_broken_slide_gets_alternative_composition(self):
        plan = _plan("cover", "broken_layout", "final")
        calls: list[tuple[int, int, set[str]]] = []

        def rebuild(index, source_index, tried):
            calls.append((index, source_index, set(tried)))
            # отдаём следующую композицию, которой ещё не было
            for candidate in ("also_broken", "good_cards"):
                if candidate not in tried:
                    return {"variant": candidate, "composition_kind": "cards", "source_index": source_index, "content": {}}
            return None

        with mock.patch.object(resilience.template_engine, "render_variant", self._fake_render({"broken_layout", "also_broken"})):
            notes = resilience.render_with_repair({}, plan, "t.pptx", self.out, None, None, "variant_a", rebuild)
        self.assertEqual([s["variant"] for s in plan["slides"]], ["cover", "good_cards", "final"])
        self.assertEqual(len(notes), 1)
        self.assertIn("Слайд 2", notes[0])
        self.assertIn("good_cards", notes[0])
        self.assertTrue(all(c[1] == 1 for c in calls))

    def test_slide_without_alternatives_is_dropped_with_note(self):
        plan = _plan("cover", "broken_layout", "final")
        with mock.patch.object(resilience.template_engine, "render_variant", self._fake_render({"broken_layout"})):
            notes = resilience.render_with_repair({}, plan, "t.pptx", self.out, None, None, "variant_a", None)
        self.assertEqual([s["variant"] for s in plan["slides"]], ["cover", "final"])
        self.assertEqual(len(notes), 1)
        self.assertIn("убран", notes[0])

    def test_error_not_reproducible_per_slide_is_reraised(self):
        plan = _plan("cover", "cards")

        def render(template, plan, template_path, output_path, manifest, content_plan, variant_key):
            if len(plan["slides"]) > 1:
                raise RuntimeError("ошибка склейки колоды")
            Path(output_path).write_bytes(b"ok")

        with mock.patch.object(resilience.template_engine, "render_variant", render):
            with self.assertRaises(RuntimeError):
                resilience.render_with_repair({}, plan, "t.pptx", self.out, None, None, "variant_a", None)

    def test_friendly_error_strips_technical_dump(self):
        exc = ValueError("Текстовый слот выходит за границы слайда:\n\nSlot: text\n{'x': 29.23}")
        self.assertEqual(resilience.friendly_error(exc), "Текстовый слот выходит за границы слайда")


class AssembleAllVariantsTest(unittest.IsolatedAsyncioTestCase):
    async def test_broken_variant_does_not_cancel_others(self):
        from app import orchestrator
        from app.schemas.job import Job

        job = Job(job_id="j1", brief="тест", purpose="", template_path="t.pptx")

        def fake_one(job_, settings, variant, template, layout_plan):
            if variant == "variant_b":
                raise ValueError("Текстовый слот выходит за границы слайда:\n\nSlot: text\n{...}")
            job_.variant_paths[variant] = f"/tmp/{variant}.pptx"

        with mock.patch.object(orchestrator, "_assemble_one_variant", fake_one), \
             mock.patch.object(orchestrator, "_load_job_template", lambda *_: {}):
            await orchestrator._assemble_all_variants(job, settings=None)
        self.assertEqual(sorted(job.variant_paths), ["variant_a", "variant_c"])
        self.assertEqual(job.variant_errors, {"variant_b": "Текстовый слот выходит за границы слайда"})

    async def test_all_variants_broken_fails_with_readable_message(self):
        from app import orchestrator
        from app.schemas.job import Job

        job = Job(job_id="j2", brief="тест", purpose="", template_path="t.pptx")

        def fake_one(job_, settings, variant, template, layout_plan):
            raise RuntimeError("макет не найден")

        with mock.patch.object(orchestrator, "_assemble_one_variant", fake_one), \
             mock.patch.object(orchestrator, "_load_job_template", lambda *_: {}):
            with self.assertRaises(RuntimeError) as ctx:
                await orchestrator._assemble_all_variants(job, settings=None)
        self.assertIn("Ни один вариант дизайна не собрался", str(ctx.exception))
        self.assertNotIn("Traceback", str(ctx.exception))
