"""End-to-end generation: template -> model plan -> validated PPTX -> audit."""

import argparse
import json
import os
import time
import zipfile
from pathlib import Path

from pptx_template_parser import build_template, validate_generation_plan
from pptx_template_parser.model import ModelError, OpenAICompatibleModel
from pptx_template_parser.planning import generate_plan
from scripts.full_test import _audit_pptx, _build_plan, _slot_coverage
from scripts.generate import create_presentation


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def _save_json(path, value):
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")


def _extract_fallback_image(source, output_dir):
    with zipfile.ZipFile(source) as archive:
        candidates = [
            name for name in archive.namelist()
            if name.startswith("ppt/media/")
            and Path(name).suffix.lower() in {".png", ".jpg", ".jpeg"}
        ]
        if not candidates:
            return None
        name = candidates[0]
        target = output_dir / ("template_image" + Path(name).suffix.lower())
        target.write_bytes(archive.read(name))
        return target.resolve()


def _fill_required_images(template, plan, image_path):
    variants = {variant["id"]: variant for variant in template["variants"]}
    missing = []
    for index, slide in enumerate(plan["slides"]):
        contract = variants[slide["variant"]]["generation_contract"]
        images = slide.setdefault("image_content", {})
        image_slots = {item["slot"] for item in contract["image_slots"]}
        for slot in image_slots:
            current = images.get(slot)
            current_exists = (
                isinstance(current, str) and Path(current).expanduser().is_file()
            )
            if not current_exists:
                if image_path is not None:
                    images[slot] = str(image_path)
                else:
                    images.pop(slot, None)
        for slot in contract["required_slots"]:
            if slot in image_slots and not images.get(slot):
                if image_path is None:
                    missing.append(f"slide[{index}].image_content.{slot}")
                else:
                    images[slot] = str(image_path)
    if missing:
        raise ValueError(
            "В шаблоне обязательны изображения, но подходящий файл не найден: "
            + ", ".join(missing)
        )


def run(source, brief, purpose, slide_count, output_dir, model=None,
        mock=False, image=None, mock_profile="01_balanced"):
    started = time.perf_counter()
    output_dir.mkdir(parents=True, exist_ok=True)
    pptx_path = output_dir / "presentation.pptx"
    if pptx_path.exists():
        raise FileExistsError(f"Результат уже существует: {pptx_path}")

    template = build_template(source)
    _save_json(output_dir / "catalog.json", {
        "schema_id": template.get("schema_id"),
        "source_mode": template.get("source_mode"),
        "selection_contract": template.get("selection_contract"),
        "variant_catalog": template.get("variant_catalog"),
    })
    if mock:
        profile = mock_profile
        outline = {"slides": [
            {"variant": item["variant"], "intent": f"Тестовый слайд {index + 1}"}
            for index, item in enumerate(
                _build_plan(template, profile, slide_count, Path("placeholder.png"))["slides"]
            )
        ]}
        plan = _build_plan(template, profile, slide_count, None)
    else:
        outline, plan = generate_plan(
            model, template, brief, purpose=purpose, slide_count=slide_count
        )

    fallback = Path(image).resolve() if image else _extract_fallback_image(source, output_dir)
    if fallback is not None and not fallback.is_file():
        raise FileNotFoundError(f"Изображение не найдено: {fallback}")
    _fill_required_images(template, plan, fallback)
    errors = validate_generation_plan(template, plan, check_image_files=True)
    if errors:
        _save_json(output_dir / "validation_errors.json", errors)
        raise ValueError("Финальный план не прошёл строгую валидацию.")
    stale_errors = output_dir / "validation_errors.json"
    if stale_errors.exists():
        stale_errors.unlink()

    _save_json(output_dir / "outline.json", outline)
    _save_json(output_dir / "plan.json", plan)
    create_presentation(template, plan["slides"], pptx_path, None, source_file=source)
    audit = _audit_pptx(pptx_path, slide_count)
    audit["slot_coverage"] = _slot_coverage(template, plan)
    audit["seconds"] = round(time.perf_counter() - started, 3)
    audit["model"] = "mock" if mock else model.model
    audit["passed"] = (
        not audit["findings"]
        and not audit["slot_coverage"]["missing_editable_slots"]
        and audit["slide_count"] == slide_count
    )
    _save_json(output_dir / "audit.json", audit)
    return pptx_path, audit


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Создать презентацию из брифа через OpenAI-compatible модель"
    )
    parser.add_argument("source", type=Path)
    parser.add_argument("--brief", required=True)
    parser.add_argument("--purpose", default="business")
    parser.add_argument("--slides", type=int, default=10)
    parser.add_argument("--output-dir", type=Path, default=PROJECT_ROOT / "output" / "ai_run")
    parser.add_argument("--base-url", default=os.getenv("PPTX_LLM_BASE_URL"))
    parser.add_argument("--model", default=os.getenv("PPTX_LLM_MODEL"))
    parser.add_argument("--api-key", default=os.getenv("PPTX_LLM_API_KEY"))
    parser.add_argument("--image", type=Path)
    parser.add_argument("--mock", action="store_true",
                        help="Проверить весь пайплайн без обращения к модели")
    args = parser.parse_args(argv)
    if not args.source.is_file():
        parser.error(f"Шаблон не найден: {args.source}")
    if not 10 <= args.slides <= 15:
        parser.error("Для приемочного запуска укажи --slides от 10 до 15")
    model = None if args.mock else OpenAICompatibleModel(
        base_url=args.base_url, model=args.model, api_key=args.api_key
    )
    try:
        path, audit = run(
            args.source, args.brief, args.purpose, args.slides,
            args.output_dir, model=model, mock=args.mock, image=args.image,
        )
    except (OSError, ValueError, ModelError) as exc:
        parser.error(str(exc))
    print(json.dumps({
        "presentation": str(path.resolve()),
        "audit": audit,
    }, ensure_ascii=False, indent=2))
    return 0 if audit["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
