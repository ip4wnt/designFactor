"""Two-stage semantic planning against Template JSON 3.0."""

import json
from pathlib import Path

from .model import ModelError
from .template.validation import validate_generation_plan


PROMPTS_DIR = Path(__file__).resolve().parents[1] / "prompts"


def _prompt(name):
    return (PROMPTS_DIR / name).read_text(encoding="utf-8")


def generate_plan(model, template, brief, purpose="business", slide_count=10,
                  max_repairs=2):
    if not 1 <= slide_count <= 30:
        raise ValueError("slide_count должен быть от 1 до 30.")
    outline = model.json(
        _prompt("outline_system.txt"),
        json.dumps({
            "brief": brief,
            "purpose": purpose,
            "slide_count": slide_count,
            "source_mode": template.get("source_mode"),
            "selection_contract": template.get("selection_contract"),
            "variant_catalog": template.get("variant_catalog"),
        }, ensure_ascii=False),
    )
    for attempt in range(max_repairs + 1):
        outline_errors = _validate_outline(template, outline, slide_count)
        if not outline_errors:
            break
        if attempt == max_repairs:
            outline = _normalize_outline(template, outline, slide_count, brief)
            outline_errors = _validate_outline(template, outline, slide_count)
            if outline_errors:
                raise ValueError(
                    "Модель вернула невалидную структуру: "
                    + "; ".join(outline_errors)
                )
            break
        outline = model.json(
            _prompt("outline_repair_system.txt"),
            json.dumps({
                "invalid_outline": outline,
                "errors": outline_errors,
                "slide_count": slide_count,
                "source_mode": template.get("source_mode"),
                "variant_catalog": template.get("variant_catalog"),
            }, ensure_ascii=False),
            temperature=0,
        )

    selected = {item["variant"] for item in outline["slides"]}
    contracts = {
        variant["id"]: variant["generation_contract"]
        for variant in template["variants"] if variant["id"] in selected
    }
    request = {
        "brief": brief,
        "purpose": purpose,
        "outline": outline,
        "contracts": contracts,
        "rules": template.get("model_contract", {}).get("rules", []),
    }
    plan = model.json(
        _prompt("fill_system.txt"),
        json.dumps(request, ensure_ascii=False),
    )
    for attempt in range(max_repairs + 1):
        errors = [
            item for item in validate_generation_plan(
                template, plan, check_image_files=False
            )
            if not (
                item.get("code") == "missing_required_slot"
                and item.get("field") == "image_content"
            )
        ]
        errors.extend(_validate_plan_shape(plan, outline))
        if not errors:
            plan = _normalize_plan(plan, outline, contracts)
            return outline, _enrich_weak_plan(
                model, plan, outline, contracts, brief, purpose
            )
        if attempt == max_repairs:
            plan = _normalize_plan(plan, outline, contracts)
            final_errors = [
                item for item in validate_generation_plan(
                    template, plan, check_image_files=False
                )
                if not (
                    item.get("code") == "missing_required_slot"
                    and item.get("field") == "image_content"
                )
            ]
            final_errors.extend(_validate_plan_shape(plan, outline))
            if not final_errors:
                return outline, _enrich_weak_plan(
                    model, plan, outline, contracts, brief, purpose
                )
            raise ValueError(
                "План модели не прошёл валидацию:\n" +
                json.dumps(final_errors, ensure_ascii=False, indent=2)
            )
        plan = model.json(
            _prompt("repair_system.txt"),
            json.dumps({
                "invalid_plan": plan,
                "errors": errors,
                "contracts": contracts,
            }, ensure_ascii=False),
            temperature=0,
        )
    raise AssertionError("unreachable")


def _validate_outline(template, outline, slide_count):
    slides = outline.get("slides") if isinstance(outline, dict) else None
    if not isinstance(slides, list) or len(slides) != slide_count:
        return [f"slides должен содержать ровно {slide_count} элементов"]
    ids = {variant["id"] for variant in template["variants"]}
    errors = []
    seen = set()
    for index, slide in enumerate(slides):
        if not isinstance(slide, dict):
            errors.append(f"slides[{index}] не объект")
            continue
        variant = slide.get("variant")
        if variant not in ids:
            errors.append(f"slides[{index}].variant неизвестен: {variant}")
        if not isinstance(slide.get("intent"), str) or not slide["intent"].strip():
            errors.append(f"slides[{index}].intent пуст")
        if template.get("source_mode") == "slides" and variant in seen:
            errors.append(f"authored variant повторён: {variant}")
        seen.add(variant)
    return errors


def _validate_plan_shape(plan, outline):
    slides = plan.get("slides") if isinstance(plan, dict) else None
    if not isinstance(slides, list):
        return []  # основной валидатор уже вернёт ошибку
    errors = []
    if len(slides) != len(outline["slides"]):
        errors.append({
            "code": "wrong_slide_count",
            "message": f"Нужно {len(outline['slides'])} слайдов, получено {len(slides)}.",
        })
    return errors


def _normalize_outline(template, outline, slide_count, brief):
    """Keep model intents, but replace unusable layout choices deterministically."""
    catalog = template.get("variant_catalog") or []
    if len(catalog) < slide_count and template.get("source_mode") == "slides":
        return outline
    supplied = outline.get("slides", []) if isinstance(outline, dict) else []
    chosen = []
    used = set()
    ids = {item["id"] for item in catalog}

    def candidates(index):
        if index == 0:
            preferred = {"cover", "introduction"}
        elif index == slide_count - 1:
            preferred = {"conclusion", "closing", "call_to_action", "action"}
        else:
            preferred = {
                "content", "explanation", "comparison", "process", "evidence",
                "section", "analysis", "overview", "universal",
            }
        ranked = sorted(
            catalog,
            key=lambda item: (
                item.get("narrative_role") not in preferred,
                item.get("review_required", False),
                item.get("confidence", 0) * -1,
            ),
        )
        return [item["id"] for item in ranked]

    for index in range(slide_count):
        source = supplied[index] if index < len(supplied) and isinstance(
            supplied[index], dict
        ) else {}
        variant = source.get("variant")
        # Prefer visual variety whenever the template has enough alternatives.
        # For authored slides uniqueness is mandatory; for layouts it is a
        # deterministic quality guard against a weak model choosing one layout
        # for the whole deck.
        unique_required = (
            template.get("source_mode") == "slides"
            or len(catalog) >= slide_count
        )
        if variant not in ids or (unique_required and variant in used):
            variant = next(
                item for item in candidates(index)
                if not unique_required or item not in used
            )
        used.add(variant)
        intent = source.get("intent")
        if not isinstance(intent, str) or not intent.strip():
            intent = _fallback_intent(brief, index, slide_count)
        kinds = source.get("content_kinds")
        if not isinstance(kinds, list) or not kinds:
            kinds = ["text"]
        chosen.append({"variant": variant, "intent": intent, "content_kinds": kinds})
    return {"slides": chosen}


def _fallback_intent(brief, index, slide_count):
    topic = " ".join(str(brief).split()).split(".", 1)[0].strip()
    topic = _shorten(topic, 72) or "Предлагаемое решение"
    if index == 0:
        return topic
    if index == slide_count - 1:
        return "Следующий шаг и контакты"
    stages = [
        "Проблема",
        "Аудитория",
        "Решение",
        "Механика",
        "Выгоды",
        "Внедрение",
        "Результат",
        "Экономика",
        "Риски",
    ]
    return stages[min(index - 1, len(stages) - 1)]


def _normalize_plan(plan, outline, contracts):
    """Force the model result into the selected contracts and safe text budgets."""
    supplied = plan.get("slides", []) if isinstance(plan, dict) else []
    normalized = []
    for index, outline_slide in enumerate(outline["slides"]):
        variant = outline_slide["variant"]
        contract = contracts[variant]
        source = supplied[index] if index < len(supplied) and isinstance(
            supplied[index], dict
        ) else {}
        source_content = source.get("content", {})
        if not isinstance(source_content, dict):
            source_content = {}
        content = {}
        for slot in contract.get("text_slots", []):
            name = slot["slot"]
            value = source_content.get(name)
            if not isinstance(value, str) or not value.strip():
                value = _fallback_text(slot, outline_slide["intent"], index)
            limit = slot.get("constraints", {}).get("recommended_characters")
            content[name] = _shorten(value.strip(), limit)
        images = source.get("image_content", {})
        if not isinstance(images, dict):
            images = {}
        allowed_images = {item["slot"] for item in contract.get("image_slots", [])}
        normalized.append({
            "variant": variant,
            "content": content,
            "image_content": {
                key: value for key, value in images.items() if key in allowed_images
            },
        })
    return {"slides": normalized}


def _shorten(text, limit):
    if not limit or len(text) <= limit:
        return text
    clipped = text[:limit + 1]
    boundary = clipped.rfind(" ")
    if boundary >= max(4, limit // 2):
        clipped = clipped[:boundary]
    else:
        # Never leave a visibly chopped word in a generated presentation.
        clipped = "Суть" if limit >= 4 else "—"
    return clipped.rstrip(" ,;:.—-")


def _fallback_text(slot, intent, index):
    role = slot.get("role", "")
    if role in {"page_number", "slide_number"}:
        return str(index + 1)
    if role in {"footer", "eyebrow", "label", "section_label"}:
        return "Ключевой вывод"
    return intent


def _enrich_weak_plan(model, plan, outline, contracts, brief, purpose):
    """Use a simple copy schema when the contract-shaped model answer is empty."""
    values = [
        value.strip().lower()
        for slide in plan["slides"]
        for value in slide.get("content", {}).values()
        if isinstance(value, str) and value.strip()
    ]
    substantial = [value for value in values if len(value) >= 35]
    if substantial and len(set(values)) >= max(6, len(outline["slides"])):
        return plan
    try:
        copy = model.json(
            _prompt("copy_system.txt"),
            json.dumps({
                "brief": brief,
                "purpose": purpose,
                "slides": [
                    {"index": index, "intent": slide["intent"]}
                    for index, slide in enumerate(outline["slides"])
                ],
            }, ensure_ascii=False),
            temperature=0.35,
        )
    except ModelError:
        return plan
    copy_slides = copy.get("slides") if isinstance(copy, dict) else None
    if not isinstance(copy_slides, list):
        return plan
    by_index = {
        item.get("index"): item for item in copy_slides
        if isinstance(item, dict) and isinstance(item.get("index"), int)
    }
    enriched = []
    for index, base in enumerate(plan["slides"]):
        item = by_index.get(index)
        if not item:
            enriched.append(base)
            continue
        contract = contracts[base["variant"]]
        content = {}
        sequence = _copy_sequence(item)
        for slot_index, slot in enumerate(contract.get("text_slots", [])):
            value = _copy_for_slot(item, slot, slot_index, sequence)
            if not value:
                value = base.get("content", {}).get(slot["slot"], "")
            limit = slot.get("constraints", {}).get("recommended_characters")
            content[slot["slot"]] = _shorten(str(value).strip(), limit)
        enriched.append({
            "variant": base["variant"],
            "content": content,
            "image_content": base.get("image_content", {}),
        })
    return {"slides": enriched}


def _copy_sequence(item):
    values = []
    for key in ("body", "subtitle"):
        if isinstance(item.get(key), str) and item[key].strip():
            values.append(item[key].strip())
    for key in ("bullets", "metrics"):
        if isinstance(item.get(key), list):
            values.extend(str(value).strip() for value in item[key] if str(value).strip())
    return values


def _copy_for_slot(item, slot, slot_index, sequence):
    role = slot.get("role", "")
    if "title" in role or role in {"heading", "section_heading"}:
        return item.get("title")
    if role in {"subtitle", "presentation_subtitle", "section_subtitle"}:
        return item.get("subtitle") or item.get("body")
    if role in {"footer", "eyebrow", "label", "section_label"}:
        return item.get("footer") or item.get("subtitle")
    if any(token in role for token in ("metric", "fact", "stat", "number")):
        metrics = item.get("metrics")
        if isinstance(metrics, list) and metrics:
            return metrics[slot_index % len(metrics)]
    if role == "quote" and isinstance(item.get("body"), str):
        return item["body"]
    return sequence[slot_index % len(sequence)] if sequence else item.get("body")
