import unittest
import tempfile
import json
import os
import subprocess
import sys
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.util import Cm

from pptx_template_parser.template.normalizer import normalize_variant, normalize_images
from scripts.generate import add_background, create_presentation, next_output_pair
from pptx_template_parser import build_template, validate_generation_plan
from pptx_template_parser.layout.variants import parse_variants
from scripts.export_json import main as save_template_json
from scripts.validate_plan import main as validate_plan_cli
from scripts.render_plan import main as render_plan_cli
from scripts.full_test import PROFILES, _build_plan
from pptx_template_parser.model import ModelError, parse_json_response
from pptx_template_parser.planning import generate_plan


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SAMPLE_TEMPLATES = PROJECT_ROOT / "sample_templates"


class BackgroundTests(unittest.TestCase):
    def test_model_json_parser_accepts_fenced_json(self):
        self.assertEqual(parse_json_response('```json\n{"ok": true}\n```'), {"ok": True})
        with self.assertRaises(ModelError):
            parse_json_response("not json")

    def test_two_stage_planner_repairs_invalid_text(self):
        template = {
            "source_mode": "layouts",
            "selection_contract": {},
            "variant_catalog": [{"variant": "v1"}],
            "model_contract": {"rules": []},
            "variants": [{
                "id": "v1",
                "generation_contract": {
                    "variant": "v1",
                    "text_slots": [{
                        "slot": "title", "role": "slide_title",
                        "constraints": {"recommended_characters": 5,
                                        "maximum_characters": 8,
                                        "recommended_lines": 1,
                                        "maximum_lines": 1},
                    }],
                    "image_slots": [],
                    "required_slots": ["title"],
                    "optional_slots": [],
                },
            }],
        }

        class FakeModel:
            def __init__(self):
                self.responses = [
                    {"slides": [{"variant": "v1", "intent": "Тема"}]},
                    {"slides": [{"variant": "v1", "content": {"title": "слишком длинно"}}]},
                    {"slides": [{"variant": "v1", "content": {"title": "Тема"}}]},
                    {"slides": [{"index": 0, "title": "Тема", "subtitle": "",
                                 "body": "", "bullets": [], "metrics": [],
                                 "footer": ""}]},
                ]

            def json(self, *args, **kwargs):
                return self.responses.pop(0)

        outline, plan = generate_plan(FakeModel(), template, "brief", slide_count=1)
        self.assertEqual(outline["slides"][0]["variant"], "v1")
        self.assertEqual(plan["slides"][0]["content"]["title"], "Тема")

    def test_full_acceptance_plans_cover_all_templates(self):
        templates = [
            path for path in SAMPLE_TEMPLATES.glob("*.pptx")
            if not path.name.startswith("~$")
        ]
        self.assertGreaterEqual(len(templates), 4)
        for source in templates:
            template = build_template(source)
            for profile in PROFILES:
                plan = _build_plan(template, profile, 10, Path("test-image.png"))
                self.assertEqual(len(plan["slides"]), 10)
                self.assertEqual(validate_generation_plan(template, plan), [])

    def test_default_master_uses_authored_slides_as_empty_templates(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "source.pptx"
            output = Path(directory) / "result.pptx"
            presentation = Presentation()
            slide = presentation.slides.add_slide(presentation.slide_layouts[1])
            slide.shapes.title.text = "Настоящий заголовок"
            slide.placeholders[1].text = "Настоящий текст"
            presentation.save(source)
            template = build_template(source)
            self.assertEqual(template["source_mode"], "slides")
            self.assertEqual(len(template["variants"]), 1)
            self.assertNotIn("Настоящий текст", json.dumps(template, ensure_ascii=False))
            create_presentation(template, [{"variant": "slide_001", "content": {}}],
                                output, None, source_file=source)
            generated = Presentation(output)
            self.assertEqual(generated.slides[0].shapes.title.text, "title")
            self.assertNotIn("Настоящий текст", " ".join(
                shape.text for shape in generated.slides[0].shapes if shape.has_text_frame))

    def test_custom_master_background_keeps_layout_mode(self):
        from pptx.dml.color import RGBColor
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "custom.pptx"
            presentation = Presentation()
            presentation.slides.add_slide(presentation.slide_layouts[0])
            presentation.slide_master.background.fill.solid()
            presentation.slide_master.background.fill.fore_color.rgb = RGBColor(20, 30, 40)
            presentation.save(source)
            self.assertEqual(build_template(source)["source_mode"], "layouts")

    def test_supplied_plain_slide_deck_uses_slides(self):
        matches = list(SAMPLE_TEMPLATES.glob("Шаблон коты без образцов слайдов.pptx"))
        if not matches:
            self.skipTest("Пример с обычными слайдами отсутствует")
        template = build_template(matches[0])
        self.assertEqual(template["source_mode"], "slides")
        self.assertEqual(len(template["variants"]), len(Presentation(matches[0]).slides))
        self.assertTrue(any(v["image_slots"] for v in template["variants"]))
        self.assertTrue(any(v["constraints"]["image_regions_cm"]
                            for v in template["variants"]))
        self.assertTrue(all(v["background"] and v["background"].get("target")
                            for v in template["variants"]))
        self.assertEqual(template["variants"][0]["background"]["shape_index"], 0)
        template_before_generation = json.dumps(template, sort_keys=True)
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "output.pptx"
            create_presentation(template, [{"variant": v["id"], "content": {}}
                                           for v in template["variants"]],
                                output, None, source_file=matches[0])
            self.assertEqual(json.dumps(template, sort_keys=True), template_before_generation)
            generated = Presentation(output)
            self.assertEqual(len(generated.slides), len(Presentation(matches[0]).slides))
            for slide in generated.slides:
                self.assertEqual(sum(shape.shape_type == 13 for shape in slide.shapes), 1)
                background = slide.shapes[0]
                self.assertEqual((background.left, background.top, background.width,
                                  background.height),
                                 (0, 0, generated.slide_width, generated.slide_height))
            image_slot = next(v["image_slots"] for v in template["variants"]
                              if v["image_slots"])
            self.assertTrue(image_slot)
            placeholder = next(shape for shape in generated.slides[0].shapes
                               if shape.is_placeholder and shape.placeholder_format.type ==
                               PP_PLACEHOLDER.PICTURE)
            self.assertGreater(placeholder.width, 0)
            from io import BytesIO
            original_picture = next(shape for index, shape in enumerate(
                Presentation(matches[0]).slides[0].shapes)
                if index > 0 and shape.shape_type == 13)
            placeholder.insert_picture(BytesIO(original_picture.image.blob))

    def test_same_image_at_distinct_positions_is_not_lost(self):
        images = [
            {"target": "ppt/media/logo.png", "x": x, "y": 1,
             "w": 2, "h": 2, "name": "logo"}
            for x in (1, 8, 1)
        ]
        result = normalize_images(images, 25, 14)
        self.assertEqual([image["x"] for image in result["images"]], [1, 8])

    def test_generator_creates_new_pptx_and_json_on_each_run(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "template.pptx"
            output = Path(directory) / "result.pptx"
            Presentation().save(source)
            command = [
                sys.executable,
                "-m", "scripts.generate",
                str(source), "--output", str(output),
            ]
            for _ in range(2):
                environment = os.environ.copy()
                environment["PYTHONIOENCODING"] = "utf-8"
                subprocess.run(
                    command, check=True, capture_output=True, text=True,
                    encoding="utf-8", errors="strict", cwd=PROJECT_ROOT,
                    env=environment,
                )

            for name in ("result", "result_2"):
                generated = output.with_name(f"{name}.pptx")
                metadata = output.with_name(f"{name}.json")
                self.assertEqual(len(Presentation(generated).slides), 11)
                self.assertEqual(len(json.loads(metadata.read_text(encoding="utf-8"))
                                     ["variants"]), 11)
            self.assertEqual(next_output_pair(output)[0].name, "result_3.pptx")

    def test_generator_does_not_overwrite_existing_output(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "template.pptx"
            output = Path(directory) / "existing.pptx"
            Presentation().save(source)
            output.write_bytes(b"existing user data")
            template = build_template(source)
            with self.assertRaises(FileExistsError):
                create_presentation(template, [], output, None, source_file=source)
            self.assertEqual(output.read_bytes(), b"existing user data")

    def test_json_contains_actual_layout_types_and_constraints(self):
        presentation = Presentation()
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "input.pptx"
            output = Path(directory) / "rules.json"
            presentation.save(source)
            save_template_json([str(source), "--output", str(output)])
            data = json.loads(output.read_text(encoding="utf-8"))

        self.assertEqual(data["type"], "presentation_template")
        self.assertEqual(len(data["variants"]), len(presentation.slide_layouts))
        title_layout = data["variants"][0]
        self.assertEqual(title_layout["type"], "title_text")
        self.assertIn("title", title_layout["constraints"]["text_regions_cm"])
        self.assertEqual(title_layout["constraints"]["slide_bounds_cm"]["w"],
                         data["slide_size"]["width"])

    def test_json_contains_backward_compatible_semantic_contract(self):
        presentation = Presentation()
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "input.pptx"
            presentation.save(source)
            data = build_template(source)

        self.assertEqual(data["schema_version"], "3.0")
        self.assertEqual(data["schema_id"], "pptx-template-semantics-3.0")
        schema_path = PROJECT_ROOT / data["schema_file"]
        self.assertEqual(json.loads(schema_path.read_text(encoding="utf-8"))["$id"],
                         data["schema_id"])
        self.assertEqual(data["schema"]["canonical_layers"].keys(),
                         {"observed", "inferred", "generation_contract"})
        self.assertIn("generator_payload", data["model_contract"])
        self.assertIn("generator_capabilities", data["model_contract"])
        variant = data["variants"][0]
        # Legacy renderer fields remain the source of truth.
        self.assertIn("slots", variant)
        self.assertIn("image_slots", variant)
        self.assertIn("constraints", variant)
        self.assertIn("layout_role", variant["semantic"])
        self.assertIn("content_budget", variant["semantic"])
        for slot in variant["slots"].values():
            self.assertIn("editable", slot)
            self.assertIn("semantic_role", slot)
            if slot["editable"]:
                self.assertEqual(slot["generator_field"], "content")
                self.assertGreater(slot["capacity"]["maximum_characters"], 0)

    def test_v3_separates_observations_inferences_and_generation_contract(self):
        presentation = Presentation()
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "input.pptx"
            presentation.save(source)
            variant = build_template(source)["variants"][0]

        self.assertIn("observed", variant)
        self.assertIn("inferred", variant)
        self.assertIn("generation_contract", variant)
        observed_title = variant["observed"]["text_slots"]["title"]
        self.assertIn("placeholder_type", observed_title)
        self.assertNotIn("semantic_role", observed_title)
        self.assertNotIn("confidence", observed_title)

        inferred_title = variant["inferred"]["text_slots"]["title"]
        self.assertEqual(inferred_title["role"]["primary"], "presentation_title")
        self.assertIn("confidence", inferred_title["role"])
        self.assertIn("evidence", inferred_title["role"])

        contract = variant["generation_contract"]
        self.assertEqual(contract["variant"], variant["id"])
        self.assertIn("title", contract["payload"]["content"])
        self.assertIn("title", contract["required_slots"])

    def test_observed_layer_does_not_leak_heuristic_fields(self):
        variant = normalize_variant(
            {"name": "unknown", "slots_raw": [{
                "name": "Text Box 1", "shape_id": 2, "shape_index": 0,
                "is_placeholder": False, "placeholder_idx": None,
                "placeholder_type": None, "placeholder_type_id": None,
                "text": "Example", "font": {"size": 18}, "margins": {},
                "x": 1, "y": 1, "w": 8, "h": 3,
            }], "images_raw": [], "image_placeholders_raw": []},
            25.4, 14.29,
        )
        observed = variant["observed"]["text_slots"]["text"]
        forbidden = {"editable", "decorative", "semantic_role", "role_confidence",
                     "content_kind", "required", "instructions", "capacity"}
        self.assertTrue(forbidden.isdisjoint(observed))
        self.assertEqual(
            variant["inferred"]["text_slots"]["text"]["role"]["primary"],
            "body",
        )

    def test_slide_examples_are_associated_with_their_layout_without_text_leak(self):
        presentation = Presentation()
        slide = presentation.slides.add_slide(presentation.slide_layouts[1])
        slide.shapes.title.text = "Revenue grew 42%"
        slide.placeholders[1].text = "First point\nSecond point"
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "examples.pptx"
            presentation.save(source)
            raw = parse_variants(source)
            title_layout = next(item for item in raw if item["layout_name"] ==
                                presentation.slide_layouts[1].name)
            self.assertEqual(len(title_layout["usage_examples_raw"]), 1)
            variant = build_template(source)["variants"][0]

        serialized = json.dumps(variant, ensure_ascii=False)
        self.assertNotIn("Revenue grew", serialized)
        example = variant["observed"]["usage_examples"][0]
        signals = {signal for shape in example["text_shapes"]
                   for signal in shape["content_signals"]}
        self.assertIn("contains_number", signals)
        self.assertIn("multiline", signals)
        self.assertEqual(variant["inferred"]["usage_summary"]["example_count"], 1)

    def test_geometry_is_compacted_to_composition_and_semantic_relations(self):
        raw_slots = []
        for index, (x, y) in enumerate(((1, 2), (1, 5), (13, 2), (13, 5)), start=1):
            raw_slots.append({
                "name": f"Box {index}", "shape_id": index,
                "shape_index": index - 1, "is_placeholder": False,
                "placeholder_idx": None, "placeholder_type": None,
                "placeholder_type_id": None, "text": "Example",
                "font": {"size": 18}, "margins": {},
                "x": x, "y": y, "w": 8, "h": 2,
            })
        variant = normalize_variant(
            {"name": "unknown", "slots_raw": raw_slots, "images_raw": [],
             "image_placeholders_raw": []},
            25.4, 14.29,
        )
        self.assertNotIn("spatial_graph", variant["observed"])
        self.assertNotIn("geometry_patterns", variant["inferred"])
        composition = variant["inferred"]["composition"]
        exported_types = {composition["primary"]} | {
            item["type"] for item in composition["alternatives"]
        }
        self.assertIn("multi_column", exported_types)
        self.assertIn("repeated_blocks", exported_types)
        self.assertTrue(composition["groups"])

    def test_image_caption_exports_semantic_relation_not_spatial_edges(self):
        paths = [p for p in SAMPLE_TEMPLATES.glob("*3.pptx")
                 if not p.name.startswith("~$")]
        if not paths:
            self.skipTest("Третий шаблон не найден")
        template = build_template(paths[0])
        variant = next(v for v in template["variants"]
                       if v["semantic"]["layout_role"] == "photo_with_caption")
        relation_types = {item["type"] for item in variant["inferred"]["relations"]}
        self.assertIn("caption_of", relation_types)
        self.assertTrue(relation_types.isdisjoint({
            "near", "aligned_left", "same_row", "same_column",
        }))

    def test_flexible_slot_is_not_reported_as_unknown(self):
        variant = normalize_variant(
            {"name": "neutral", "slots_raw": [{
                "name": "Text Box", "shape_id": 2, "shape_index": 0,
                "is_placeholder": False, "placeholder_idx": None,
                "placeholder_type": None, "placeholder_type_id": None,
                "text": "Example", "font": {"size": 18}, "margins": {},
                "x": 2, "y": 2, "w": 12, "h": 5,
            }], "images_raw": [], "image_placeholders_raw": []},
            25.4, 14.29,
        )
        slot = variant["inferred"]["text_slots"]["text"]
        self.assertEqual(slot["semantic_scope"], "flexible")
        self.assertIn("paragraph", slot["allowed_content_kinds"])
        self.assertIn("bullets", slot["allowed_content_kinds"])
        self.assertEqual(variant["inferred"]["ambiguities"], [])
        self.assertTrue(variant["inferred"]["downstream_choice"]["allowed"])
        self.assertFalse(variant["inferred"]["review"]["recommended"])
        self.assertEqual(variant["generation_contract"]["decision_owner"],
                         "downstream_model")

    def test_empty_unknown_layout_still_requires_review(self):
        variant = normalize_variant(
            {"name": "neutral", "slots_raw": [], "images_raw": [],
             "image_placeholders_raw": []},
            25.4, 14.29,
        )
        self.assertEqual(variant["inferred"]["layout"]["semantic_scope"], "unknown")
        self.assertTrue(variant["inferred"]["review"]["recommended"])
        self.assertFalse(variant["inferred"]["downstream_choice"]["allowed"])

    def test_specific_role_and_allowed_content_kind_are_separate(self):
        presentation = Presentation()
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "input.pptx"
            presentation.save(source)
            variant = build_template(source)["variants"][0]
        title = variant["inferred"]["text_slots"]["title"]
        self.assertEqual(title["role"]["primary"], "presentation_title")
        self.assertEqual(title["allowed_content_kinds"], ["short_heading"])
        self.assertNotIn("presentation_title", title["allowed_content_kinds"])

    def test_exact_geometry_matching_produces_honest_usage_evidence(self):
        presentation = Presentation()
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        box = slide.shapes.add_textbox(Cm(2), Cm(3), Cm(10), Cm(2))
        box.text = "42% growth"
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "geometry.pptx"
            presentation.save(source)
            variant = build_template(source)["variants"][0]
        example = variant["observed"]["usage_examples"][0]["text_shapes"][0]
        self.assertEqual(example["slot"], "text")
        self.assertEqual(example["match_method"], "exact_geometry")
        evidence = variant["inferred"]["text_slots"]["text"]["usage_evidence"]
        self.assertEqual(evidence["sample_count"], 1)
        self.assertEqual(evidence["evidence_strength"], "low")
        self.assertIn("contains_number", evidence["observed_once"]["content_signals"])
        self.assertNotIn("summary", evidence)
        self.assertNotIn("usage_evidence",
                         variant["generation_contract"]["text_slots"][0])

    def test_variant_catalog_is_compact_and_references_real_contracts(self):
        path = SAMPLE_TEMPLATES / "VK Tech шаблон 1.pptx"
        if not path.exists():
            self.skipTest("Шаблон не найден")
        template = build_template(path)
        self.assertEqual(
            [item["id"] for item in template["variant_catalog"]],
            [item["id"] for item in template["variants"]],
        )
        full_size = len(json.dumps(template["variants"], ensure_ascii=False))
        catalog_size = len(json.dumps(template["variant_catalog"], ensure_ascii=False))
        self.assertLess(catalog_size, full_size * 0.35)
        serialized = json.dumps(template["variant_catalog"], ensure_ascii=False)
        self.assertNotIn("placeholder_idx", serialized)
        self.assertNotIn("template_text", serialized)
        self.assertNotIn("usage_evidence", serialized)

    def test_catalog_only_export(self):
        presentation = Presentation()
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "input.pptx"
            output = Path(directory) / "catalog.json"
            presentation.save(source)
            save_template_json([
                str(source), "--catalog-only", "--output", str(output),
            ])
            data = json.loads(output.read_text(encoding="utf-8"))
        self.assertEqual(data["type"], "presentation_template_selection_catalog")
        self.assertIn("selection_contract", data)
        self.assertEqual(len(data["variants"]), len(presentation.slide_layouts))
        self.assertNotIn("model_contract", data)

    def test_catalog_integrity_for_all_supplied_templates(self):
        for path in SAMPLE_TEMPLATES.glob("*.pptx"):
            if path.name.startswith("~$"):
                continue
            with self.subTest(template=path.name):
                template = build_template(path)
                variants = {item["id"]: item for item in template["variants"]}
                catalog = template["variant_catalog"]
                self.assertEqual(len(catalog), len(variants))
                self.assertEqual(len({item["id"] for item in catalog}), len(catalog))
                for item in catalog:
                    variant = variants[item["id"]]
                    contract = variant["generation_contract"]
                    self.assertEqual(item["required_slots"], contract["required_slots"])
                    self.assertEqual(item["decision_owner"], contract["decision_owner"])
                    self.assertEqual(item["review_required"],
                                     variant["inferred"]["review"]["recommended"])
                    catalog_slots = {slot["slot"] for slot in
                                     item["text_slots"] + item["image_slots"]}
                    contract_slots = {slot["slot"] for slot in
                                      contract["text_slots"] + contract["image_slots"]}
                    self.assertEqual(catalog_slots, contract_slots)

    def test_family_members_have_the_same_selection_signature(self):
        path = SAMPLE_TEMPLATES / "VK Tech шаблон 1.pptx"
        if not path.exists():
            self.skipTest("Шаблон не найден")
        catalog = build_template(path)["variant_catalog"]
        families = {}
        for item in catalog:
            signature = (
                item["narrative_role"], item["content_pattern"], item["composition"],
                tuple((slot["scope"], slot["role"],
                       tuple(slot["allowed_content_kinds"]), slot["required"])
                      for slot in item["text_slots"]),
                tuple((slot["scope"], slot["role"],
                       tuple(slot["allowed_content_kinds"]), slot["required"])
                      for slot in item["image_slots"]),
            )
            previous = families.setdefault(item["family"], signature)
            self.assertEqual(previous, signature)

    def test_generation_plan_validator_rejects_contract_violations(self):
        presentation = Presentation()
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "input.pptx"
            presentation.save(source)
            template = build_template(source)
        variant = template["variants"][0]
        variant_id = variant["id"]

        missing = validate_generation_plan(template, {
            "slides": [{"variant": variant_id, "content": {}}]
        })
        self.assertIn("missing_required_slot", {item["code"] for item in missing})

        unknown = validate_generation_plan(template, {
            "slides": [{"variant": variant_id,
                        "content": {"title": "OK", "not_a_slot": "bad"}}]
        })
        self.assertIn("unknown_text_slot", {item["code"] for item in unknown})

        title = variant["generation_contract"]["text_slots"][0]
        too_long = "x" * (title["constraints"]["maximum_characters"] + 1)
        overflow = validate_generation_plan(template, {
            "slides": [{"variant": variant_id, "content": {"title": too_long}}]
        })
        self.assertIn("text_too_long", {item["code"] for item in overflow})

        recommended = title["constraints"]["recommended_characters"]
        maximum = title["constraints"]["maximum_characters"]
        if recommended < maximum:
            unsafe = "x" * (recommended + 1)
            errors = validate_generation_plan(template, {
                "slides": [{"variant": variant_id, "content": {"title": unsafe}}]
            })
            self.assertIn("text_exceeds_recommended_capacity",
                          {item["code"] for item in errors})

    def test_catalog_to_contract_to_pptx_end_to_end_for_all_templates(self):
        for source in SAMPLE_TEMPLATES.glob("*.pptx"):
            if source.name.startswith("~$"):
                continue
            with self.subTest(template=source.name), tempfile.TemporaryDirectory() as directory:
                template = build_template(source)
                catalog_item = next(
                    item for item in template["variant_catalog"]
                    if not item["review_required"] and not any(
                        slot["required"] for slot in item["image_slots"]
                    )
                )
                variant = next(item for item in template["variants"]
                               if item["id"] == catalog_item["id"])
                content = {
                    slot["slot"]: "Проверка"
                    for slot in variant["generation_contract"]["text_slots"]
                    if slot["required"]
                }
                plan = {"slides": [{"variant": catalog_item["id"],
                                     "content": content}]}
                self.assertEqual(validate_generation_plan(template, plan), [])
                output = Path(directory) / "result.pptx"
                create_presentation(template, plan["slides"], output, None,
                                    source_file=source)
                generated = Presentation(output)
                self.assertEqual(len(generated.slides), 1)

    def test_slide_mode_keeps_selected_variants_in_plan_order(self):
        source = SAMPLE_TEMPLATES / "Шаблон коты без образцов слайдов.pptx"
        if not source.exists():
            self.skipTest("Шаблон с authored slides не найден")
        template = build_template(source)
        selected = [template["variants"][2], template["variants"][0]]
        plan = {"slides": [{"variant": item["id"], "content": {}}
                           for item in selected]}
        self.assertEqual(validate_generation_plan(template, plan), [])
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "selected.pptx"
            create_presentation(template, plan["slides"], output, None,
                                source_file=source)
            generated = Presentation(output)
            self.assertEqual(len(generated.slides), 2)
            original = Presentation(source)
            for generated_slide, source_index in zip(generated.slides, (2, 0)):
                generated_background = next(
                    shape for shape in generated_slide.shapes
                    if shape.shape_type == 13
                )
                original_background = next(
                    shape for shape in original.slides[source_index].shapes
                    if shape.shape_type == 13
                )
                self.assertEqual(generated_background.image.blob,
                                 original_background.image.blob)

    def test_slide_mode_rejects_duplicate_authored_variant(self):
        source = SAMPLE_TEMPLATES / "Шаблон коты без образцов слайдов.pptx"
        if not source.exists():
            self.skipTest("Шаблон с authored slides не найден")
        template = build_template(source)
        variant_id = template["variants"][0]["id"]
        errors = validate_generation_plan(template, {
            "slides": [{"variant": variant_id}, {"variant": variant_id}]
        })
        self.assertIn("duplicate_authored_slide_variant",
                      {item["code"] for item in errors})

    def test_validate_plan_cli_reports_valid_and_invalid_plans(self):
        presentation = Presentation()
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "input.pptx"
            valid_path = Path(directory) / "valid.json"
            invalid_path = Path(directory) / "invalid.json"
            presentation.save(source)
            template = build_template(source)
            variant = template["variants"][0]
            valid_path.write_text(json.dumps({
                "slides": [{"variant": variant["id"],
                            "content": {"title": "Проверка"}}]
            }, ensure_ascii=False), encoding="utf-8")
            invalid_path.write_text(json.dumps({
                "slides": [{"variant": "missing", "content": {}}]
            }), encoding="utf-8")
            self.assertEqual(validate_plan_cli([str(source), str(valid_path)]), 0)
            self.assertEqual(validate_plan_cli([str(source), str(invalid_path)]), 1)

    def test_render_plan_cli_creates_openable_pptx_and_rejects_invalid_plan(self):
        presentation = Presentation()
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "input.pptx"
            valid_path = Path(directory) / "valid.json"
            invalid_path = Path(directory) / "invalid.json"
            output = Path(directory) / "result.pptx"
            rejected_output = Path(directory) / "rejected.pptx"
            presentation.save(source)
            template = build_template(source)
            variant_id = template["variants"][0]["id"]
            valid_path.write_text(json.dumps({
                "slides": [{"variant": variant_id,
                            "content": {"title": "Проверка"}}]
            }, ensure_ascii=False), encoding="utf-8")
            invalid_path.write_text(json.dumps({
                "slides": [{"variant": "missing"}]
            }), encoding="utf-8")
            self.assertEqual(render_plan_cli([
                str(source), str(valid_path), "--output", str(output),
            ]), 0)
            self.assertEqual(len(Presentation(output).slides), 1)
            self.assertEqual(render_plan_cli([
                str(source), str(invalid_path), "--output", str(rejected_output),
            ]), 1)
            self.assertFalse(rejected_output.exists())

    def test_subtitle_placeholder_is_not_misclassified_as_title(self):
        presentation = Presentation()
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "input.pptx"
            presentation.save(source)
            variant = build_template(source)["variants"][0]
        self.assertIn("title", variant["slots"])
        self.assertIn("subtitle", variant["slots"])
        self.assertEqual(variant["slots"]["subtitle"]["semantic_role"],
                         "presentation_subtitle")

    def test_all_masters_are_parsed_without_slides(self):
        presentation = Presentation()
        presentation.slides.add_slide(presentation.slide_layouts[0])
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "source.pptx"
            presentation.save(path)
            variants = parse_variants(path)
        self.assertEqual(len(variants), len(presentation.slide_master.slide_layouts))
        self.assertNotEqual(len(variants), len(presentation.slides))

    def test_supplied_template_includes_second_master(self):
        path = SAMPLE_TEMPLATES / "VK Tech шаблон 1.pptx"
        if not path.exists():
            self.skipTest("Шаблон не найден")
        presentation = Presentation(path)
        variants = parse_variants(path)
        self.assertEqual(
            len(variants),
            sum(len(master.slide_layouts) for master in presentation.slide_masters),
        )
        self.assertEqual(len({variant["name"] for variant in variants}), len(variants))

    def test_generated_layout_preserves_grouped_logo_and_text(self):
        path = SAMPLE_TEMPLATES / "VK Tech шаблон 1.pptx"
        if not path.exists():
            self.skipTest("Шаблон не найден")
        template = build_template(path)
        variant = template["variants"][3]
        source = Presentation(path)
        layout = source.slide_masters[1].slide_layouts[0]
        self.assertTrue(any(shape.shape_type == 6 for shape in layout.shapes))
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "generated.pptx"
            create_presentation(
                template,
                [{"variant": variant["id"], "content": {"title": "Проверка"}}],
                output,
                None,
                source_file=path,
            )
            generated = Presentation(output)
            self.assertEqual(len(generated.slides), 1)
            self.assertEqual(generated.slides[0].slide_layout.name, layout.name)
            self.assertEqual(generated.slides[0].shapes.title.text, "Проверка")
            self.assertTrue(any(
                shape.shape_type == 6
                for shape in generated.slides[0].slide_layout.shapes
            ))

    def test_workspace_template_uses_its_own_layouts(self):
        path = SAMPLE_TEMPLATES / "VK Tech шаблон 2.pptx"
        if not path.exists():
            path = SAMPLE_TEMPLATES / (
                "VK_WorkSpace_Клиентская_конференция_Шаблон_03.pptx"
            )
        if not path.exists():
            self.skipTest("Второй шаблон не найден")
        template = build_template(path)
        self.assertEqual(
            len(template["variants"]),
            sum(len(master.slide_layouts) for master in Presentation(path).slide_masters),
        )
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "generated.pptx"
            slides = [
                {"variant": variant["id"],
                 "content": {name: name for name, slot in variant["slots"].items()
                             if slot["placeholder_idx"] is not None}}
                for variant in template["variants"]
            ]
            create_presentation(template, slides, output, None, source_file=path)
            generated = Presentation(output)
            self.assertEqual(len(generated.slides), len(slides))
            self.assertEqual(
                [slide.slide_layout.name for slide in generated.slides],
                [variant["id"] for variant in template["variants"]],
            )

    def test_top_footer_is_recreated_as_an_editable_slot(self):
        candidates = [
            path for path in SAMPLE_TEMPLATES.glob("*3.pptx")
            if not path.name.startswith("~$")
        ]
        if not candidates:
            self.skipTest("Третий шаблон не найден")
        path = candidates[0]
        template = build_template(path)
        presentation = Presentation(path)
        self.assertEqual(len(template["variants"]), sum(
            len(master.slide_layouts) for master in presentation.slide_masters
        ))
        footer_variants = [variant for variant in template["variants"]
                           if "footer" in variant["slots"]]
        self.assertTrue(footer_variants)
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "footer.pptx"
            create_presentation(
                template,
                [{"variant": variant["id"], "content": {"footer": "footer"}}
                 for variant in footer_variants],
                output, None, source_file=path,
            )
            generated = Presentation(output)
            for variant, slide in zip(footer_variants, generated.slides):
                slot = variant["slots"]["footer"]
                footer = slide.placeholders[slot["placeholder_idx"]]
                self.assertEqual(footer.placeholder_format.type, PP_PLACEHOLDER.FOOTER)
                self.assertEqual(footer.text, "footer")
                self.assertAlmostEqual(footer.top / 360000, slot["y"], places=2)

    def test_picture_placeholder_is_not_replaced_with_text(self):
        paths = [p for p in SAMPLE_TEMPLATES.glob("*3.pptx")
                 if not p.name.startswith("~$")]
        if not paths:
            self.skipTest("Третий шаблон не найден")
        source = paths[0]
        template = build_template(source)
        variant = next(v for v in template["variants"] if v["image_slots"])
        self.assertNotIn("image", variant["slots"])
        image_idx = variant["image_slots"]["image"]["placeholder_idx"]
        self.assertNotIn(image_idx, [s["placeholder_idx"] for s in
                                     variant["slots"].values()])

        text = {name: name for name, slot in variant["slots"].items()
                if slot["placeholder_idx"] is not None}
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "picture-placeholder.pptx"
            create_presentation(template, [{"variant": variant["id"],
                                            "content": text}], output,
                                None, source_file=source)
            picture = Presentation(output).slides[0].placeholders[image_idx]
            self.assertEqual(picture.placeholder_format.type, PP_PLACEHOLDER.PICTURE)
            self.assertEqual(picture.text, "")

            photo = PROJECT_ROOT / "assets" / "vktech" / "logo.png"
            if photo.is_file():
                picture_output = Path(directory) / "with-picture.pptx"
                create_presentation(template, [{"variant": variant["id"],
                                                "content": text,
                                                "image_content": {"image": photo}}],
                                    picture_output, None, source_file=source)
                self.assertGreater(
                    len(Presentation(picture_output).slides[0]
                        .placeholders[image_idx].image.blob),
                    0,
                )

    def test_third_template_semantics_cover_tricky_layouts(self):
        paths = [p for p in SAMPLE_TEMPLATES.glob("*3.pptx")
                 if not p.name.startswith("~$")]
        if not paths:
            self.skipTest("Третий шаблон не найден")
        template = build_template(paths[0])

        speaker = next(v for v in template["variants"]
                       if v["semantic"]["layout_role"] == "speaker")
        self.assertEqual(speaker["slots"]["title"]["semantic_role"], "speaker_name")
        self.assertEqual(speaker["image_slots"]["image"]["semantic_role"],
                         "speaker_portrait")
        self.assertTrue(speaker["image_slots"]["image"]["required"])

        quote = next(v for v in template["variants"]
                     if v["semantic"]["layout_role"] == "quote")
        quote_roles = {slot["semantic_role"] for slot in quote["slots"].values()}
        self.assertIn("quote_author", quote_roles)
        punctuation = next(slot for slot in quote["slots"].values()
                           if slot.get("template_text") == "«")
        self.assertFalse(punctuation["editable"])
        self.assertTrue(punctuation["decorative"])

        caption = next(v for v in template["variants"]
                       if v["semantic"]["layout_role"] == "photo_with_caption")
        self.assertIn("image_caption",
                      {slot["semantic_role"] for slot in caption["slots"].values()})
        self.assertEqual(caption["image_slots"]["image"]["group_id"], "image_1")

        closing = next(v for v in template["variants"]
                       if v["semantic"]["layout_role"] == "closing_qr")
        self.assertEqual(closing["image_slots"]["image"]["semantic_role"], "qr_code")

        cover = next(v for v in template["variants"]
                     if v["semantic"]["layout_role"] == "cover")
        tiny = [slot for slot in cover["slots"].values() if slot["w"] <= 0.35]
        self.assertTrue(tiny)
        self.assertTrue(all(not slot["editable"] for slot in tiny))

    def test_master_color_survives_without_full_slide_image(self):
        variant = normalize_variant(
            {"name": "test", "slots_raw": [], "images_raw": [],
             "background_raw": {"color": "#123456"}},
            25.4, 14.29,
        )
        self.assertEqual(variant["background"], {"color": "#123456"})

        presentation = Presentation()
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        add_background(slide, variant["background"], "")
        self.assertEqual(str(slide.background.fill.fore_color.rgb), "123456")

    def test_master_image_becomes_full_slide_background(self):
        variant = normalize_variant(
            {"name": "test", "slots_raw": [], "images_raw": [],
             "background_raw": {"target": "ppt/media/bg.png"}},
            25.4, 14.29,
        )
        self.assertEqual(variant["background"]["w"], 25.4)
        self.assertEqual(variant["background"]["target"], "ppt/media/bg.png")


if __name__ == "__main__":
    unittest.main()
