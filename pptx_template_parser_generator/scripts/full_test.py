"""Generate and audit the complete acceptance-test matrix.

The command intentionally uses only the public parser/validator/generator pipeline.
It creates three different 10-slide plans for every PPTX in sample_templates.
"""

import argparse
import json
import re
import textwrap
import time
import zipfile
from pathlib import Path

from pptx import Presentation

from pptx_template_parser import build_template, validate_generation_plan
from scripts.generate import create_presentation


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_TEMPLATES = PROJECT_ROOT / "sample_templates"
DEFAULT_OUTPUT = PROJECT_ROOT / "output" / "full_test"

PROFILES = {
    "01_short": 0.35,
    "02_medium": 0.60,
    "03_full": 0.90,
}


def _make_test_text(target_length, seed=0):
    """
    Генерирует нейтральный тестовый текст примерно нужной длины.
    Никаких знаний о назначении слота не требуется.
    """
    if target_length <= 0:
        return ""

    words = [
        "данные",
        "система",
        "анализ",
        "результат",
        "процесс",
        "проект",
        "информация",
        "решение",
        "проверка",
        "пример",
    ]

    result = []
    length = 0
    index = seed

    while length < target_length:
        word = words[index % len(words)]

        candidate = (
            " ".join(result + [word])
            if result
            else word
        )

        if len(candidate) > target_length:
            break

        result.append(word)
        length = len(candidate)
        index += 1

    text = " ".join(result)

    # Если слот настолько маленький, что даже слово не влезает
    if not text:
        text = "Т" * target_length

    return text


def _get_target_length(constraints, density):
    """
    Определяет объём текста исключительно из ограничений,
    которые нашёл парсер.
    """
    recommended = constraints.get("recommended_characters")
    maximum = constraints.get("maximum_characters")

    if recommended:
        base = recommended
    elif maximum:
        base = maximum
    else:
        # Контракт не сообщил вместимость.
        # Используем небольшой технический текст.
        base = 40

    target = max(1, round(base * density))

    if maximum:
        target = min(target, maximum)

    return target


def _fit_to_slot(value, constraints, slot=None):
    """
    Ограничивает тестовый текст геометрией и контрактом слота.
    Не знает ничего о semantic role.
    """
    maximum = constraints.get("maximum_characters")

    if maximum:
        value = value[:maximum]

    recommended_lines = constraints.get("recommended_lines")

    if not slot or not recommended_lines:
        return value

    margins = slot.get("margins") or {}

    width_cm = max(
        0.1,
        slot.get("w", 1)
        - margins.get("left", 0)
        - margins.get("right", 0),
    )

    font_size = constraints.get("font_size_pt") or 18

    chars_per_line = max(
        1,
        int(
            width_cm
            * 28.346
            / (font_size * 0.68)
        ),
    )

    wrapped = textwrap.wrap(
        value,
        width=chars_per_line,
        break_long_words=False,
        break_on_hyphens=False,
    )

    wrapped = wrapped[:recommended_lines]

    return "\n".join(wrapped)


def _variant_is_renderable(template, variant):
    """
    Проверяем только техническую корректность геометрии.
    """
    width = template["slide_size"]["width"]
    height = template["slide_size"]["height"]

    slots = list(
        (variant.get("slots") or {}).values()
    )

    slots += list(
        (variant.get("image_slots") or {}).values()
    )

    return all(
        slot.get("x", 0) >= 0
        and slot.get("y", 0) >= 0
        and slot.get("w", 0) >= 0
        and slot.get("h", 0) >= 0
        and slot.get("x", 0)
        + slot.get("w", 0)
        <= width + 0.02
        and slot.get("y", 0)
        + slot.get("h", 0)
        <= height + 0.02
        for slot in slots
    )


def _get_renderable_variants(template):
    """
    Получаем все варианты, которые:
    1. имеют хотя бы один редактируемый слот;
    2. технически помещаются на слайде.

    Никакой семантики.
    """
    result = []

    for variant in template["variants"]:
        contract = variant.get(
            "generation_contract"
        ) or {}

        text_slots = contract.get(
            "text_slots"
        ) or []

        image_slots = contract.get(
            "image_slots"
        ) or []

        if not text_slots and not image_slots:
            continue

        if not _variant_is_renderable(
            template,
            variant,
        ):
            continue

        result.append(variant)

    if not result:
        raise ValueError(
            "В шаблоне нет заполняемых вариантов."
        )

    return result


def _choose_variants(template, count, profile_index):
    """
    Выбирает варианты без знания их назначения.

    Разные профили просто начинают обход массива
    с разных позиций.
    """
    variants = _get_renderable_variants(
        template
    )

    if (
        template.get("source_mode") == "slides"
        and len(variants) >= count
    ):
        offset = (
            profile_index * count
        ) % len(variants)

        return [
            variants[
                (offset + index)
                % len(variants)
            ]
            for index in range(count)
        ]

    return [
        variants[index % len(variants)]
        for index in range(count)
    ]


def _build_plan(
    template,
    profile_name,
    slide_count,
    image_path,
):
    """
    Строит план исключительно из generation_contract.

    Скрипт не знает:
    - какие существуют типы слайдов;
    - какие существуют semantic roles;
    - что такое title / metric / quote;
    - какую информацию "положено" писать в слот.
    """

    density = PROFILES[profile_name]

    profile_names = list(PROFILES)

    profile_index = profile_names.index(
        profile_name
    )

    variants = _choose_variants(
        template,
        slide_count,
        profile_index,
    )

    slides = []

    for slide_index, variant in enumerate(
        variants
    ):
        contract = variant.get(
            "generation_contract"
        ) or {}

        text_specs = contract.get(
            "text_slots"
        ) or []

        image_specs = contract.get(
            "image_slots"
        ) or []

        content = {}

        for slot_index, spec in enumerate(
            text_specs
        ):
            constraints = (
                spec.get("constraints")
                or {}
            )

            target_length = (
                _get_target_length(
                    constraints,
                    density,
                )
            )

            value = _make_test_text(
                target_length,
                seed=(
                    slide_index
                    + slot_index
                ),
            )

            slot = (
                variant.get("slots")
                or {}
            ).get(spec["slot"])

            value = _fit_to_slot(
                value,
                constraints,
                slot,
            )

            content[
                spec["slot"]
            ] = value

        images = {}

        if image_path:
            for spec in image_specs:
                images[
                    spec["slot"]
                ] = str(image_path)

        item = {
            "variant": variant["id"],
            "content": content,
        }

        if images:
            item["image_content"] = images

        slides.append(item)

    return {
        "profile": profile_name,
        "purpose": (
            "Автоматический приемочный "
            "прогон генератора"
        ),
        "slides": slides,
    }


def _slug(value):
    value = re.sub(r"[^0-9A-Za-zА-Яа-я]+", "_", value).strip("_")
    return value[:60] or "template"




def _fit(value, constraints, density, slot=None, fallback="План"):
    limit = constraints.get("recommended_characters")
    if limit is None:
        limit = constraints.get("maximum_characters")
    if limit is None:
        limit = 160
    target = max(1, min(limit, round(limit * density)))
    lines = max(1, constraints.get("recommended_lines") or 1)
    if slot:
        margins = slot.get("margins") or {}
        width_cm = max(
            0.1,
            slot.get("w", 1)
            - margins.get("left", 0)
            - margins.get("right", 0),
        )
        font_size = constraints.get("font_size_pt") or 18
        # Conservative Cyrillic estimate. Real fitting remains the renderer's job.
        chars_per_line = max(1, int(width_cm * 28.346 / (font_size * 0.68)))
        target = min(target, chars_per_line * lines)
    else:
        chars_per_line = target
    if chars_per_line <= 2:
        return str((len(value) % 9) + 1)[:chars_per_line]
    if len(value) <= target:
        shortened = value
    else:
        words = value.split()
        kept = []
        for word in words:
            candidate = " ".join(kept + [word])
            if len(candidate) > target:
                break
            kept.append(word)
        shortened = " ".join(kept)
        if not shortened:
            shortened = fallback if len(fallback) <= target else str((target % 9) + 1)
    wrapped = textwrap.wrap(
        shortened,
        width=chars_per_line,
        break_long_words=False,
        break_on_hyphens=False,
    )[:lines]
    return "\n".join(wrapped) or value[:1]






def _extract_test_image(source, asset_dir):
    asset_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    candidates = []

    with zipfile.ZipFile(source) as archive:
        for name in archive.namelist():
            if not name.startswith(
                "ppt/media/"
            ):
                continue

            suffix = Path(
                name
            ).suffix.lower()

            if suffix not in {
                ".png",
                ".jpg",
                ".jpeg",
            }:
                continue

            data = archive.read(name)

            candidates.append(
                (len(data), name, data)
            )

    if not candidates:
        return None

    _, name, data = max(
        candidates,
        key=lambda item: item[0],
    )

    target = (
        asset_dir
        / (
            source.stem
            + Path(name).suffix.lower()
        )
    )

    target.write_bytes(data)

    return target.resolve()


def _audit_pptx(path, expected_slides):
    deck = Presentation(path)
    findings = []
    if len(deck.slides) != expected_slides:
        findings.append(f"slide_count:{len(deck.slides)}")
    marker_pattern = re.compile(
        r"lorem ipsum|\bXXX\b|\bTODO\b|вставьте текст|\btext(?:_\d+|\s+\d+)?\b",
        re.I,
    )
    for number, slide in enumerate(deck.slides, 1):
        text_shapes = 0
        picture_shapes = 0
        for shape in slide.shapes:
            geometry = (shape.left, shape.top, shape.width, shape.height)
            if all(value is not None for value in geometry):
                if shape.left < 0 or shape.top < 0:
                    findings.append(f"slide_{number}:negative_position")
                if (shape.left + shape.width > deck.slide_width
                        or shape.top + shape.height > deck.slide_height):
                    findings.append(f"slide_{number}:outside_slide")
            if getattr(shape, "has_text_frame", False):
                text = shape.text.strip()
                if text:
                    text_shapes += 1
                    if marker_pattern.search(text):
                        findings.append(f"slide_{number}:placeholder_text")
            if shape.shape_type == 13:  # MSO_SHAPE_TYPE.PICTURE
                picture_shapes += 1
        if not text_shapes and not picture_shapes:
            findings.append(f"slide_{number}:empty")
        if picture_shapes == 1 and len(slide.shapes) == 1:
            findings.append(f"slide_{number}:single_raster")
    return {
        "openable": True,
        "slide_count": len(deck.slides),
        "native_objects": not any("single_raster" in item for item in findings),
        "findings": sorted(set(findings)),
    }


def _slot_coverage(template, plan):
    variants = {item["id"]: item for item in template["variants"]}
    totals = {"editable_text": 0, "filled_text": 0,
              "editable_images": 0, "filled_images": 0}
    missing = []
    for slide_number, item in enumerate(plan["slides"], 1):
        contract = variants[item["variant"]]["generation_contract"]
        content = item.get("content") or {}
        images = item.get("image_content") or {}
        totals["editable_text"] += len(contract["text_slots"])
        totals["editable_images"] += len(contract["image_slots"])
        for spec in contract["text_slots"]:
            if content.get(spec["slot"]) not in (None, ""):
                totals["filled_text"] += 1
            else:
                missing.append(f"slide_{slide_number}:text:{spec['slot']}")
        for spec in contract["image_slots"]:
            if images.get(spec["slot"]) not in (None, ""):
                totals["filled_images"] += 1
            else:
                missing.append(f"slide_{slide_number}:image:{spec['slot']}")
    totals["missing_editable_slots"] = missing
    return totals


def run(templates_dir, output_dir, slide_count):
    if not 10 <= slide_count <= 15:
        raise ValueError("По ТЗ тестовый объем должен быть от 10 до 15 слайдов.")
    sources = sorted(templates_dir.glob("*.pptx"))
    if not sources:
        raise ValueError(f"В {templates_dir} нет PPTX-шаблонов.")
    output_dir.mkdir(parents=True, exist_ok=True)
    plan_dir = output_dir / "plans"
    deck_dir = output_dir / "presentations"
    asset_dir = output_dir / "assets"
    plan_dir.mkdir(exist_ok=True)
    deck_dir.mkdir(exist_ok=True)
    records = []
    for source in sources:
        template = build_template(source)
        image_path = _extract_test_image(source, asset_dir)
        for profile_name in PROFILES:
            started = time.perf_counter()
            plan = _build_plan(template, profile_name, slide_count, image_path)
            errors = validate_generation_plan(template, plan, check_image_files=True)
            stem = f"{_slug(source.stem)}__{profile_name}"
            plan_path = plan_dir / f"{stem}.json"
            deck_path = deck_dir / f"{stem}.pptx"
            plan_path.write_text(
                json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8"
            )
            if deck_path.exists():
                deck_path.unlink()
            if not errors:
                create_presentation(
                    template, plan["slides"], deck_path, None, source_file=source
                )
                audit = _audit_pptx(deck_path, slide_count)
            else:
                audit = {"openable": False, "findings": ["plan_validation_failed"]}
            records.append({
                "template": source.name,
                "source_mode": template.get("source_mode"),
                "profile": profile_name,
                "plan": str(plan_path.resolve()),
                "presentation": str(deck_path.resolve()),
                "validation_errors": errors,
                "slot_coverage": _slot_coverage(template, plan),
                "audit": audit,
                "seconds": round(time.perf_counter() - started, 3),
            })
    report = {
        "spec": {
            "target_slides": "10-15",
            "generated_slides_per_deck": slide_count,
            "variants_per_template": 3,
            "time_limit_seconds_per_deck": 300,
        },
        "summary": {
            "templates": len(sources),
            "presentations": len(records),
            "passed": sum(
                not r["validation_errors"] and not r["audit"].get("findings")
                and not r["slot_coverage"]["missing_editable_slots"]
                and r["seconds"] <= 300 for r in records
            ),
        },
        "results": records,
    }
    report_path = output_dir / "report.json"
    report_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return report_path, report


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Сгенерировать 3 варианта по 10-15 слайдов для каждого шаблона"
    )
    parser.add_argument("--templates", type=Path, default=DEFAULT_TEMPLATES)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--slides", type=int, default=10)
    args = parser.parse_args(argv)
    try:
        report_path, report = run(args.templates, args.output, args.slides)
    except (OSError, ValueError) as exc:
        parser.error(str(exc))
    print(json.dumps({
        "report": str(report_path.resolve()),
        **report["summary"],
    }, ensure_ascii=False, indent=2))
    return 0 if report["summary"]["passed"] == report["summary"]["presentations"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
