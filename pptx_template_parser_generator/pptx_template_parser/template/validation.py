"""Strict validation for plans produced from generation_contract."""

from pathlib import Path


def validate_generation_plan(template, plan, check_image_files=False):
    """Return structured errors; an empty list means the plan is renderable."""
    slides = plan.get("slides") if isinstance(plan, dict) else plan
    if not isinstance(slides, list):
        return [_error(None, None, "plan", None, "invalid_plan",
                       "План должен быть списком или объектом с массивом slides.")]
    variants = {variant["id"]: variant for variant in template.get("variants", [])}
    errors = []
    seen_variants = set()
    for index, slide in enumerate(slides):
        if not isinstance(slide, dict):
            errors.append(_error(index, None, "slide", None, "invalid_slide",
                                 "Слайд должен быть объектом."))
            continue
        variant_id = slide.get("variant")
        if template.get("source_mode") == "slides" and variant_id in seen_variants:
            errors.append(_error(
                index, variant_id, "variant", None, "duplicate_authored_slide_variant",
                "В режиме source_mode=slides authored-slide variant нельзя повторять.",
            ))
        seen_variants.add(variant_id)
        variant = variants.get(variant_id)
        if variant is None:
            errors.append(_error(index, variant_id, "variant", None, "unknown_variant",
                                 "variant отсутствует в шаблоне."))
            continue
        contract = variant["generation_contract"]
        content = slide.get("content", {})
        images = slide.get("image_content", {})
        if not isinstance(content, dict):
            errors.append(_error(index, variant_id, "content", None, "invalid_content",
                                 "content должен быть объектом."))
            content = {}
        if not isinstance(images, dict):
            errors.append(_error(index, variant_id, "image_content", None,
                                 "invalid_image_content",
                                 "image_content должен быть объектом."))
            images = {}

        text_contract = {item["slot"]: item for item in contract["text_slots"]}
        image_contract = {item["slot"]: item for item in contract["image_slots"]}
        errors.extend(_validate_text(index, variant_id, content, text_contract))
        errors.extend(_validate_images(index, variant_id, images, image_contract,
                                       check_image_files))

        for name in contract["required_slots"]:
            is_image = name in image_contract
            values = images if is_image else content
            if values.get(name) in (None, ""):
                field = "image_content" if is_image else "content"
                errors.append(_error(index, variant_id, field, name,
                                     "missing_required_slot",
                                     "Обязательный слот не заполнен."))
    return errors


def assert_valid_generation_plan(template, plan, check_image_files=False):
    errors = validate_generation_plan(template, plan, check_image_files)
    if errors:
        summary = "\n".join(
            f"slide[{item['slide_index']}] {item['code']}: {item['message']}"
            for item in errors
        )
        raise ValueError(f"План презентации не прошёл валидацию:\n{summary}")
    return plan


def _validate_text(index, variant_id, values, contract):
    errors = []
    for name, value in values.items():
        spec = contract.get(name)
        if spec is None:
            errors.append(_error(index, variant_id, "content", name, "unknown_text_slot",
                                 "Слот отсутствует среди editable text slots."))
            continue
        if value is None:
            continue
        if not isinstance(value, str):
            errors.append(_error(index, variant_id, "content", name, "invalid_text_value",
                                 "Значение текстового слота должно быть строкой."))
            continue
        constraints = spec.get("constraints") or {}
        recommended_characters = constraints.get("recommended_characters")
        maximum_characters = constraints.get("maximum_characters")
        recommended_lines = constraints.get("recommended_lines")
        maximum_lines = constraints.get("maximum_lines")
        if maximum_characters is not None and len(value) > maximum_characters:
            errors.append(_error(index, variant_id, "content", name, "text_too_long",
                                 f"Текст длиннее {maximum_characters} символов."))
        elif recommended_characters is not None and len(value) > recommended_characters:
            errors.append(_error(
                index, variant_id, "content", name,
                "text_exceeds_recommended_capacity",
                f"Текст длиннее безопасного бюджета {recommended_characters} символов.",
            ))
        line_count = len(value.splitlines() or [""])
        if maximum_lines is not None and line_count > maximum_lines:
            errors.append(_error(index, variant_id, "content", name, "too_many_lines",
                                 f"Текст содержит больше {maximum_lines} строк."))
        elif recommended_lines is not None and line_count > recommended_lines:
            errors.append(_error(
                index, variant_id, "content", name,
                "text_exceeds_recommended_lines",
                f"Текст содержит больше безопасного бюджета {recommended_lines} строк.",
            ))
    return errors


def _validate_images(index, variant_id, values, contract, check_files):
    errors = []
    for name, value in values.items():
        if name not in contract:
            errors.append(_error(index, variant_id, "image_content", name,
                                 "unknown_image_slot",
                                 "Слот отсутствует среди editable image slots."))
            continue
        if value is None:
            continue
        if not isinstance(value, (str, Path)):
            errors.append(_error(index, variant_id, "image_content", name,
                                 "invalid_image_value",
                                 "Изображение должно быть путём к файлу."))
        elif check_files and not Path(value).is_file():
            errors.append(_error(index, variant_id, "image_content", name,
                                 "image_file_not_found", "Файл изображения не найден."))
    return errors


def _error(slide_index, variant, field, slot, code, message):
    return {
        "slide_index": slide_index,
        "variant": variant,
        "field": field,
        "slot": slot,
        "code": code,
        "message": message,
    }
