"""Deterministic semantic enrichment for machine-readable PPTX templates.

The existing geometry remains the source of truth for rendering. This module
adds conservative hints for an LLM or another planner and never requires those
hints to generate a presentation.
"""

import math
import re
import hashlib
from collections import Counter

from pptx_template_parser.template.geometry import (
    build_spatial_graph,
    detect_geometry_patterns,
)


ROLE_GUIDANCE = {
    "presentation_title": "Название всей презентации; короткая формулировка без точки.",
    "presentation_subtitle": "Подзаголовок об аудитории, событии, авторе или дате.",
    "section_title": "Название нового смыслового раздела.",
    "slide_title": "Главная мысль или тема конкретного слайда.",
    "subtitle": "Короткое уточнение к заголовку.",
    "body": "Основное объяснение, абзац или маркированный список.",
    "column_body": "Содержимое одной колонки; сопоставляй с соседней колонкой.",
    "metric_value": "Короткое число, процент или иной крупный показатель.",
    "metric_label": "Короткая расшифровка показателя рядом с metric_value.",
    "quote": "Цитата без имени автора.",
    "quote_author": "Имя автора и при необходимости должность.",
    "speaker_name": "Имя спикера.",
    "speaker_bio": "Должность, компания или короткое описание спикера.",
    "image_caption": "Подпись, непосредственно описывающая изображение.",
    "image_source": "Источник, автор или дополнительная строка под изображением.",
    "contact_text": "Контакт, ссылка, призыв или подпись финального слайда.",
    "footer": "Служебный колонтитул; заполняй только при наличии данных.",
    "generic_text": "Универсальный текстовый блок; используй только если роль не уточнена.",
    "photo": "Содержательная фотография, соответствующая тезису слайда.",
    "speaker_portrait": "Портрет спикера без текста и посторонних логотипов.",
    "qr_code": "QR-код, ведущий на явно указанную ссылку.",
    "illustration": "Содержательная иллюстрация или схема.",
    "generic_image": "Изображение подходящих пропорций, если точная роль неизвестна.",
    "decorative": "Не заполняй: это постоянная часть оформления.",
    "closing_title": "Короткая финальная фраза, благодарность или призыв к действию.",
}


LAYOUT_HINTS = {
    "cover": ["обложка", "название презентации", "автор и дата"],
    "section_divider": ["начало раздела", "переход между темами"],
    "speaker": ["представление спикера", "эксперт или участник"],
    "metrics": ["ключевые показатели", "KPI", "факты и числа"],
    "quote": ["цитата", "отзыв", "прямая речь"],
    "comparison": ["сравнение", "до и после", "два варианта"],
    "two_columns": ["две параллельные группы текста", "сопоставление списков"],
    "photo_with_caption": ["фотография с подписью", "кейс с визуальным примером"],
    "content_with_image": ["тезис с фотографией или иллюстрацией"],
    "closing_qr": ["финальный слайд с QR-кодом", "контакты и ссылка"],
    "closing": ["финальный вывод", "благодарность", "контакты"],
    "agenda": ["содержание", "план выступления"],
    "title_only": ["один короткий тезис", "заголовок без основного текста"],
    "content": ["обычный содержательный слайд"],
    "generic": ["универсальное содержание"],
}


def schema_contract():
    """Describe the canonical v3 layers without embedding a JSON Schema file."""
    return {
        "name": "pptx-template-semantics",
        "version": "3.0",
        "canonical_layers": {
            "observed": "Факты, непосредственно извлечённые из PPTX; без смысловых догадок.",
            "inferred": "Вероятностные смысловые выводы с confidence, alternatives и evidence.",
            "generation_contract": "Разрешённый интерфейс выбора макета и заполнения слотов.",
        },
        "compatibility": {
            "legacy_fields_retained": [
                "variants[].slots", "variants[].image_slots", "variants[].semantic",
                "model_contract",
            ],
            "canonical_for_reasoning": [
                "variants[].observed", "variants[].inferred",
                "variants[].generation_contract",
            ],
            "renderer_payload_unchanged": True,
        },
        "confidence": {
            "range": [0.0, 1.0],
            "low_threshold": 0.65,
            "rule": "При низкой уверенности используй alternatives и ограничения, а не узкую роль.",
        },
    }


def selection_contract():
    return {
        "purpose": "Выбрать подходящий variant до заполнения конкретных слотов.",
        "input": "Краткий план одного слайда и variants из variant_catalog.",
        "output": {
            "variant": "точное значение variant_catalog[].id",
            "reason": "краткое обоснование соответствия содержания возможностям макета",
        },
        "rules": [
            "Сначала отфильтруй варианты по review_required=false и достаточному числу слотов.",
            "Для specific учитывай content_pattern и semantic roles.",
            "Для flexible выбирай по allowed_content_kinds, вместимости и наличию изображений.",
            "Варианты одной family взаимозаменяемы по смыслу; visual_variant можно выбрать для разнообразия.",
            "После выбора используй только generation_contract варианта с выбранным id.",
        ],
    }


def build_variant_catalog(variants):
    return [_catalog_entry(variant) for variant in variants]


def _catalog_entry(variant):
    inferred = variant["inferred"]
    layout = inferred["layout"]
    contract = variant["generation_contract"]
    composition = inferred["composition"]
    text_slots = [_catalog_slot(item) for item in contract["text_slots"]]
    image_slots = [_catalog_slot(item) for item in contract["image_slots"]]
    slot_signature = "|".join(
        f"{kind}:{item['scope']}:{item['role']}:{','.join(item['allowed_content_kinds'])}:{int(item['required'])}"
        for kind, items in (("t", text_slots), ("i", image_slots))
        for item in items
    )
    family_digest = hashlib.sha1(slot_signature.encode("utf-8")).hexdigest()[:10]
    family = ".".join((
        layout["narrative_role"], layout["content_pattern"], composition["primary"],
        family_digest,
    ))
    return {
        "id": variant["id"],
        "family": family,
        "visual_variant": variant["id"],
        "semantic_scope": layout["semantic_scope"],
        "narrative_role": layout["narrative_role"],
        "content_pattern": layout["content_pattern"],
        "composition": composition["primary"],
        "confidence": layout["confidence"],
        "best_for": contract["selection"]["best_for"],
        "avoid_when": contract["selection"]["avoid_when"],
        "text_slots": text_slots,
        "image_slots": image_slots,
        "required_slots": contract["required_slots"],
        "content_budget": contract["content_budget"],
        "review_required": inferred["review"]["recommended"],
        "decision_owner": contract["decision_owner"],
    }


def _catalog_slot(slot):
    constraints = slot.get("constraints") or {}
    result = {
        "slot": slot["slot"],
        "scope": slot["semantic_scope"],
        "role": slot["role"],
        "allowed_content_kinds": slot["allowed_content_kinds"],
        "required": slot["required"],
    }
    if slot["generator_field"] == "content":
        result["recommended_characters"] = constraints.get("recommended_characters")
        result["maximum_lines"] = constraints.get("maximum_lines")
    else:
        result["orientation"] = constraints.get("orientation")
        result["aspect_ratio"] = constraints.get("aspect_ratio")
    return result


def model_contract():
    """Instructions embedded once in JSON for any downstream model."""
    return {
        "deprecated_name": True,
        "canonical_replacement": "generation_contract",
        "purpose": "Выбор макета и заполнение только редактируемых слотов.",
        "planning_sequence": [
            "Раздели исходный материал на одну главную мысль на слайд.",
            "Определи нужный layout_role и наличие обязательного изображения.",
            "Отбери варианты по selection_hints.best_for и avoid_when.",
            "Проверь наличие данных для required_slots и соответствие slot_groups.",
            "Проверь каждый текст по capacity.recommended_* и при необходимости смени макет или раздели слайд.",
            "Сформируй generator_payload только из точных id и имён editable-слотов.",
        ],
        "rules": [
            "Используй variant.semantic.layout_role и selection_hints при выборе макета.",
            "Заполняй только слоты с editable=true; decorative=true никогда не заполняй.",
            "Соблюдай semantic_role, required и content_kind каждого слота.",
            "Не изменяй background и images: это постоянное оформление шаблона.",
            "Не превышай capacity.recommended_*; при превышении выбери другой макет, сократи или раздели слайд.",
            "Учитывай reading_order и group_id при создании колонок, метрик и подписей.",
            "При confidence ниже 0.65 используй безопасное универсальное содержание и не делай узких предположений.",
            "Не придумывай изображение для optional-слота, если оно не помогает содержанию.",
            "Ключи content и image_content должны в точности совпадать с именами слотов выбранного variant.",
            "Не помещай пути к изображениям в content: для них предназначен image_content.",
            "template_text является подсказкой или постоянным текстом источника, но не готовым содержанием нового слайда.",
            "Если semantic.ambiguous_slots не пуст, опирайся на content_kind, геометрию и placeholder_type; не выдумывай узкую роль.",
        ],
        "generator_payload": {
            "format": {
                "slides": [{
                    "variant": "точное значение variant.id",
                    "content": {"имя_текстового_слота": "строка"},
                    "image_content": {"имя_слота_изображения": "путь_к_файлу"},
                }]
            },
            "validation": [
                "variant обязан существовать в variants[].id.",
                "content принимает только editable=true слоты из variant.slots.",
                "image_content принимает только editable=true слоты из variant.image_slots.",
                "required=true означает: предоставь значение либо выбери другой вариант.",
                "Пустой optional-слот разрешено не включать в payload.",
            ],
        },
        "generator_capabilities": {
            "text": "Поддерживается через slides[].content.",
            "raster_images": "Поддерживаются через slides[].image_content при наличии image_slots.",
            "tables_charts_video": "Не генерируются текущим рендерером; не выдавай их как готовый payload.",
            "template_artwork": "Фоны, логотипы, группы и декоративные изображения сохраняются из исходного PPTX.",
        },
        "fallback_policy": {
            "unknown_text_role": "generic_text",
            "unknown_image_role": "generic_image",
            "low_confidence_threshold": 0.65,
            "overflow": ["choose_roomier_layout", "shorten", "split_slide"],
            "missing_required_content": "choose_another_layout",
            "ambiguous_slot": "use_content_kind_and_geometry_without_role_specific_assumptions",
        },
        "role_vocabulary": ROLE_GUIDANCE,
    }


def enrich_variant(variant, slide_width, slide_height):
    """Add semantic information while preserving all legacy fields."""
    layout_role, layout_confidence, evidence = _infer_layout_role(variant)

    for name, slot in variant.get("slots", {}).items():
        _enrich_text_slot(name, slot, variant, layout_role, slide_width, slide_height)

    if layout_role == "metrics":
        _assign_metric_groups(variant)
    elif layout_role == "photo_with_caption":
        _assign_caption_roles(variant)
    elif layout_role == "quote":
        _assign_quote_roles(variant)

    for name, slot in variant.get("image_slots", {}).items():
        _enrich_image_slot(name, slot, variant, layout_role)

    # Specialized group inference can change roles; derive scope and accepted
    # kinds only after all role assignments are final.
    for slot in variant.get("slots", {}).values():
        slot["semantic_scope"] = _slot_scope(slot, "text")
        slot["allowed_content_kinds"] = _allowed_content_kinds(slot, "text")
    for slot in variant.get("image_slots", {}).values():
        slot["semantic_scope"] = _slot_scope(slot, "image")
        slot["allowed_content_kinds"] = _allowed_content_kinds(slot, "image")

    reading_order = _reading_order(variant)
    spatial_graph = build_spatial_graph(variant, slide_width, slide_height)
    geometry_patterns = detect_geometry_patterns(spatial_graph)
    required = [
        name for name, slot in variant.get("slots", {}).items()
        if slot.get("required") and slot.get("editable")
    ] + [
        name for name, slot in variant.get("image_slots", {}).items()
        if slot.get("required") and slot.get("editable")
    ]
    optional = [
        name for name, slot in variant.get("slots", {}).items()
        if slot.get("editable") and not slot.get("required")
    ] + [
        name for name, slot in variant.get("image_slots", {}).items()
        if slot.get("editable") and not slot.get("required")
    ]
    ambiguities = [
        name for name, slot in {
            **variant.get("slots", {}), **variant.get("image_slots", {})
        }.items()
        if slot.get("editable") and slot.get("semantic_scope") == "unknown"
    ]
    flexible_slots = [
        name for name, slot in {
            **variant.get("slots", {}), **variant.get("image_slots", {})
        }.items()
        if slot.get("editable") and slot.get("semantic_scope") == "flexible"
    ]
    layout_scope = _layout_scope(layout_role, layout_confidence, evidence, flexible_slots)

    variant["semantic"] = {
        "layout_role": layout_role,
        "confidence": layout_confidence,
        "evidence": evidence,
        "reading_order": reading_order,
        "selection_hints": {
            "best_for": LAYOUT_HINTS.get(layout_role, LAYOUT_HINTS["generic"]),
            "required_slots": required,
            "optional_slots": optional,
            "avoid_when": _avoid_when(variant, layout_role),
        },
        "ambiguous_slots": ambiguities,
        "flexible_slots": flexible_slots,
        "semantic_scope": layout_scope,
        "slot_groups": _slot_groups(variant),
        "content_budget": _content_budget(variant),
        "review": _review_status(layout_scope, ambiguities),
        "downstream_choice": _downstream_choice(layout_scope, flexible_slots),
    }
    variant["observed"] = _observed_variant(variant)
    variant["inferred"] = _inferred_variant(
        variant, layout_role, layout_confidence, evidence, ambiguities,
        geometry_patterns,
    )
    variant["generation_contract"] = _generation_contract(variant)
    variant.pop("usage_examples_raw", None)
    return variant


def _observed_slot(slot, kind):
    """Copy only source facts. Never expose a heuristic as an observation."""
    keys = (
        "shape_id", "shape_index", "shape_name", "is_placeholder",
        "placeholder_idx", "placeholder_type", "placeholder_type_id",
        "source_kind", "x", "y", "w", "h", "font", "margins",
        "horizontal_align", "vertical_align", "template_text",
    )
    result = {key: slot.get(key) for key in keys if key in slot}
    result["object_kind"] = kind
    return result


def _observed_variant(variant):
    return {
        "source": {
            "layout_name": variant.get("layout_name"),
            "layout_file": variant.get("layout_file"),
            "master_index": variant.get("master_index"),
            "slide_index": variant.get("slide_index"),
        },
        "geometry": variant.get("constraints", {}).get("slide_bounds_cm", {}),
        "text_slots": {
            name: _observed_slot(slot, "text")
            for name, slot in variant.get("slots", {}).items()
        },
        "image_slots": {
            name: _observed_slot(slot, "image")
            for name, slot in variant.get("image_slots", {}).items()
        },
        "fixed_artwork": {
            "background": variant.get("background"),
            "images": variant.get("images", []),
        },
        "usage_examples": _usage_examples(variant),
    }


def _usage_examples(variant):
    examples = []
    for raw in variant.get("usage_examples_raw", []):
        text_shapes = []
        for shape in raw.get("text_shapes", []):
            text = str(shape.get("text") or "").strip()
            if not text:
                continue
            match = _match_usage_shape(shape, variant.get("slots", {}))
            text_shapes.append({
                **match,
                "placeholder_type": shape.get("placeholder_type"),
                "characters": len(text),
                "words": len(re.findall(r"\w+", text, flags=re.UNICODE)),
                "paragraphs": len([line for line in text.splitlines() if line.strip()]),
                "content_signals": _content_signals(text),
            })
        image_shapes = [
            _match_usage_shape(shape, variant.get("image_slots", {}))
            for shape in raw.get("image_shapes", [])
        ]
        examples.append({
            "slide_index": raw.get("slide_index"),
            "text_shapes": text_shapes,
            "image_shapes": image_shapes,
            "image_count": raw.get("image_count", 0),
        })
    return examples


def _match_usage_shape(shape, slots):
    placeholder_idx = shape.get("placeholder_idx")
    if placeholder_idx is not None:
        matches = [name for name, slot in slots.items()
                   if slot.get("placeholder_idx") == placeholder_idx]
        if len(matches) == 1:
            return {"slot": matches[0], "match_method": "placeholder_idx",
                    "match_confidence": 1.0}

    scores = sorted(
        ((_geometry_iou(shape, slot), name) for name, slot in slots.items()),
        reverse=True,
    )
    if scores and scores[0][0] >= 0.98:
        tied = len(scores) > 1 and scores[1][0] >= 0.98
        if not tied:
            return {"slot": scores[0][1], "match_method": "exact_geometry",
                    "match_confidence": round(scores[0][0], 3)}
    if scores and scores[0][0] >= 0.80:
        tied = len(scores) > 1 and scores[1][0] >= scores[0][0] - 0.02
        if not tied:
            return {"slot": scores[0][1], "match_method": "geometry_iou",
                    "match_confidence": round(scores[0][0], 3)}
    return {"slot": None, "match_method": "unmatched", "match_confidence": 0.0}


def _geometry_iou(first, second):
    values = [first.get(key) for key in ("x", "y", "w", "h")]
    other = [second.get(key) for key in ("x", "y", "w", "h")]
    if any(value is None for value in values + other):
        return 0.0
    ax, ay, aw, ah = map(float, values)
    bx, by, bw, bh = map(float, other)
    overlap_w = max(0.0, min(ax + aw, bx + bw) - max(ax, bx))
    overlap_h = max(0.0, min(ay + ah, by + bh) - max(ay, by))
    intersection = overlap_w * overlap_h
    union = aw * ah + bw * bh - intersection
    return intersection / union if union > 0 else 0.0


def _content_signals(text):
    compact = text.strip()
    signals = []
    if re.search(r"\d", compact):
        signals.append("contains_number")
    if re.search(r"\d\s*%", compact):
        signals.append("percentage")
    if re.search(r"https?://|www\.", compact, re.I):
        signals.append("url")
    if re.search(r"\S+@\S+\.\S+", compact):
        signals.append("email")
    if len(compact) <= 60:
        signals.append("short_text")
    if len(compact.splitlines()) > 1:
        signals.append("multiline")
    if compact[:1] in {"•", "-", "–"}:
        signals.append("list_like")
    return signals or ["plain_text"]


def _role_alternatives(role, confidence, kind, editable):
    if not editable or role == "decorative":
        return []
    fallback = "generic_image" if kind == "image" else "generic_text"
    if role == fallback:
        return []
    remaining = round(max(0.0, 1.0 - confidence), 2)
    return [{"role": fallback, "confidence": remaining}]


def _slot_evidence(slot):
    evidence = []
    if slot.get("placeholder_type"):
        evidence.append({"source": "placeholder_type", "value": slot["placeholder_type"]})
    if slot.get("shape_name"):
        evidence.append({"source": "shape_name", "value": slot["shape_name"]})
    if slot.get("template_text"):
        evidence.append({"source": "template_text", "value": slot["template_text"]})
    evidence.append({"source": "geometry", "value": slot.get("relative_position")})
    return evidence


def _inferred_slot(slot, kind):
    role = slot.get("semantic_role", "generic_image" if kind == "image" else "generic_text")
    confidence = slot.get("role_confidence", 0.0)
    return {
        "role": {
            "primary": role,
            "confidence": confidence,
            "alternatives": _role_alternatives(
                role, confidence, kind, slot.get("editable", False),
            ),
            "evidence": _slot_evidence(slot),
        },
        "content_kind": slot.get("content_kind"),
        "semantic_scope": slot.get("semantic_scope"),
        "allowed_content_kinds": slot.get("allowed_content_kinds", []),
        "editable": slot.get("editable", False),
        "decorative": slot.get("decorative", False),
        "relative_position": slot.get("relative_position"),
        "group_id": slot.get("group_id"),
    }


def _layout_dimensions(layout_role):
    mapping = {
        "cover": ("cover", "single_focus", "introduction"),
        "section_divider": ("section", "single_focus", "section_transition"),
        "agenda": ("agenda", "list", "navigation"),
        "closing": ("closing", "single_focus", "call_to_action"),
        "closing_qr": ("closing", "text_with_media", "call_to_action"),
        "speaker": ("content", "text_with_media", "person_profile"),
        "metrics": ("evidence", "metric_cards", "metrics"),
        "quote": ("evidence", "single_focus", "quote"),
        "comparison": ("content", "two_columns", "comparison"),
        "two_columns": ("content", "two_columns", "parallel_content"),
        "photo_with_caption": ("content", "image_with_caption", "visual_example"),
        "content_with_image": ("content", "text_with_media", "explanation"),
        "title_only": ("content", "single_focus", "statement"),
        "content": ("content", "title_body", "explanation"),
        "generic": ("unknown", "unknown", "generic"),
    }
    narrative, composition, pattern = mapping.get(layout_role, mapping["generic"])
    return {
        "narrative_role": narrative,
        "composition": composition,
        "content_pattern": pattern,
        "legacy_layout_role": layout_role,
    }


def _inferred_variant(variant, layout_role, confidence, evidence, ambiguities,
                      geometry_patterns):
    dimensions = _layout_dimensions(layout_role)
    usage_evidence = _slot_usage_evidence(variant)
    return {
        "layout": {
            **dimensions,
            "semantic_scope": variant["semantic"]["semantic_scope"],
            "confidence": confidence,
            "alternatives": ([{"layout_role": "generic", "confidence": round(1 - confidence, 2)}]
                             if layout_role != "generic" else []),
            "evidence": [{"source": item} for item in evidence],
        },
        "text_slots": {
            name: {**_inferred_slot(slot, "text"),
                   "usage_evidence": usage_evidence["text"].get(name,
                                                                  _empty_usage_evidence("text"))}
            for name, slot in variant.get("slots", {}).items()
        },
        "image_slots": {
            name: {**_inferred_slot(slot, "image"),
                   "usage_evidence": usage_evidence["image"].get(name,
                                                                   _empty_usage_evidence("image"))}
            for name, slot in variant.get("image_slots", {}).items()
        },
        "reading_order": variant["semantic"]["reading_order"],
        "composition": _composition(variant, geometry_patterns),
        "relations": _relations(variant, geometry_patterns),
        "usage_summary": _usage_summary(variant),
        "ambiguities": ambiguities,
        "flexible_slots": variant["semantic"]["flexible_slots"],
        "review": variant["semantic"]["review"],
        "downstream_choice": variant["semantic"]["downstream_choice"],
    }


def _usage_summary(variant):
    examples = _usage_examples(variant)
    signals = {}
    for example in examples:
        for shape in example["text_shapes"]:
            for signal in shape["content_signals"]:
                signals[signal] = signals.get(signal, 0) + 1
    return {
        "example_count": len(examples),
        "slides_with_images": sum(example["image_count"] > 0 for example in examples),
        "content_signal_counts": signals,
    }


def _slot_usage_evidence(variant):
    examples = _usage_examples(variant)
    text_values = {}
    image_values = {}
    for example in examples:
        for shape in example["text_shapes"]:
            if not shape["slot"]:
                continue
            item = text_values.setdefault(shape["slot"], {
                "characters": [], "words": [], "paragraphs": [], "signals": Counter(),
            })
            item["characters"].append(shape["characters"])
            item["words"].append(shape["words"])
            item["paragraphs"].append(shape["paragraphs"])
            item["signals"].update(shape["content_signals"])
        for shape in example["image_shapes"]:
            if shape["slot"]:
                image_values[shape["slot"]] = image_values.get(shape["slot"], 0) + 1

    text = {}
    for name, item in text_values.items():
        count = len(item["characters"])
        evidence = {
            "sample_count": count,
            "evidence_strength": _sample_strength(count),
        }
        if count == 1:
            evidence["observed_once"] = {
                "characters": item["characters"][0],
                "words": item["words"][0],
                "paragraphs": item["paragraphs"][0],
                "content_signals": sorted(item["signals"]),
            }
        else:
            evidence["summary"] = {
                "median_characters": _median(item["characters"]),
                "maximum_characters_observed": max(item["characters"]),
                "median_words": _median(item["words"]),
                "median_paragraphs": _median(item["paragraphs"]),
                "numeric_rate": round(item["signals"]["contains_number"] / count, 3),
                "percentage_rate": round(item["signals"]["percentage"] / count, 3),
                "multiline_rate": round(item["signals"]["multiline"] / count, 3),
            }
        text[name] = evidence
    image = {
        name: {
            "sample_count": count,
            "evidence_strength": _sample_strength(count),
            **({"observed_once": {"filled": True}} if count == 1 else {
                "summary": {"fill_rate": round(count / max(1, len(examples)), 3)}
            }),
        }
        for name, count in image_values.items()
    }
    return {"text": text, "image": image}


def _empty_usage_evidence(kind):
    return {"sample_count": 0, "evidence_strength": "none"}


def _median(values):
    ordered = sorted(values)
    middle = len(ordered) // 2
    if len(ordered) % 2:
        return ordered[middle]
    return round((ordered[middle - 1] + ordered[middle]) / 2, 2)


def _sample_strength(count):
    if count >= 5:
        return "high"
    if count >= 2:
        return "medium"
    return "low"


def _contract_slot(name, slot, kind):
    result = {
        "slot": name,
        "generator_field": "image_content" if kind == "image" else "content",
        "accepted_value": slot.get("accepted_value"),
        "purpose": slot.get("instructions"),
        "content_kind": slot.get("content_kind"),
        "required": slot.get("required", False),
        "role": slot.get("semantic_role"),
        "role_confidence": slot.get("role_confidence"),
        "semantic_scope": slot.get("semantic_scope"),
        "allowed_content_kinds": slot.get("allowed_content_kinds", []),
    }
    if kind == "text":
        result["constraints"] = slot.get("capacity")
    else:
        result["constraints"] = {
            "aspect_ratio": slot.get("aspect_ratio"),
            "orientation": slot.get("orientation"),
            "crop_mode": slot.get("crop_mode"),
        }
    if slot.get("group_id"):
        result["group_id"] = slot["group_id"]
    return result


def _generation_contract(variant):
    text = [
        _contract_slot(name, slot, "text")
        for name, slot in variant.get("slots", {}).items() if slot.get("editable")
    ]
    images = [
        _contract_slot(name, slot, "image")
        for name, slot in variant.get("image_slots", {}).items() if slot.get("editable")
    ]
    return {
        "variant": variant["id"],
        "payload": {
            "variant": variant["id"],
            "content": {item["slot"]: "string" for item in text},
            "image_content": {item["slot"]: "existing_file_path" for item in images},
        },
        "text_slots": text,
        "image_slots": images,
        "required_slots": [item["slot"] for item in text + images if item["required"]],
        "optional_slots": [item["slot"] for item in text + images if not item["required"]],
        "selection": variant["semantic"]["selection_hints"],
        "semantic_scope": variant["semantic"]["semantic_scope"],
        "decision_owner": ("downstream_model"
                           if variant["semantic"]["downstream_choice"]["allowed"]
                           else "template_contract"),
        "content_budget": variant["semantic"]["content_budget"],
        "overflow_policy": ["choose_roomier_layout", "shorten", "split_slide"],
    }


def _composition(variant, patterns):
    priority = {
        "image_caption_pair": 4,
        "multi_column": 3,
        "repeated_blocks": 2,
        "horizontal_sequence": 1,
    }
    ordered = sorted(patterns, key=lambda item: priority.get(item["type"], 0),
                     reverse=True)
    legacy = variant["semantic"]["layout_role"]
    fallback = _layout_dimensions(legacy)["composition"]
    primary = ordered[0]["type"] if ordered else fallback
    confidence = ordered[0]["confidence"] if ordered else variant["semantic"]["confidence"]
    groups = []
    for index, pattern in enumerate(ordered, start=1):
        if pattern["type"] not in {"multi_column", "repeated_blocks", "horizontal_sequence"}:
            continue
        groups.append({
            "id": f"composition_{index}",
            "type": pattern["type"],
            "members": [_slot_reference(member) for member in pattern["members"]],
            "confidence": pattern["confidence"],
        })
    return {
        "primary": primary,
        "confidence": confidence,
        "alternatives": [
            {"type": item["type"], "confidence": item["confidence"]}
            for item in ordered[1:]
        ],
        "groups": groups,
        "evidence": [item["evidence"][0] for item in ordered],
    }


def _slot_reference(value):
    kind, name = value.split(":", 1)
    return {"kind": kind, "slot": name}


def _relations(variant, geometry_patterns):
    result = []
    for group in _slot_groups(variant):
        members = group["members"]
        roles = {
            (item["kind"], item["slot"]): _member_role(variant, item)
            for item in members
        }
        values = [item for item in members if roles[(item["kind"], item["slot"])] == "metric_value"]
        labels = [item for item in members if roles[(item["kind"], item["slot"])] == "metric_label"]
        for label in labels:
            for value in values:
                result.append(_semantic_relation("label_of", label, value, group["id"], 0.9))
        quotes = [item for item in members if roles[(item["kind"], item["slot"])] == "quote"]
        authors = [item for item in members if roles[(item["kind"], item["slot"])] == "quote_author"]
        for author in authors:
            for quote in quotes:
                result.append(_semantic_relation("author_of", author, quote, group["id"], 0.9))
        images = [item for item in members if item["kind"] == "image"]
        captions = [item for item in members
                    if roles[(item["kind"], item["slot"])] == "image_caption"]
        sources = [item for item in members
                   if roles[(item["kind"], item["slot"])] == "image_source"]
        for image in images:
            for caption in captions:
                result.append(_semantic_relation("caption_of", caption, image, group["id"], 0.92))
            for source in sources:
                result.append(_semantic_relation("source_of", source, image, group["id"], 0.88))

    for pattern in geometry_patterns:
        if pattern["type"] != "image_caption_pair":
            continue
        image = next((_slot_reference(item) for item in pattern["members"]
                      if item.startswith("image:")), None)
        text = next((_slot_reference(item) for item in pattern["members"]
                     if item.startswith("text:")), None)
        if image and text and not any(
            item["type"] == "caption_of" and item["source"] == text and item["target"] == image
            for item in result
        ):
            result.append(_semantic_relation("caption_of", text, image, None,
                                             pattern["confidence"]))
    return result


def _member_role(variant, member):
    collection = variant.get("image_slots" if member["kind"] == "image" else "slots", {})
    return collection.get(member["slot"], {}).get("semantic_role")


def _semantic_relation(kind, source, target, group_id, confidence):
    relation = {
        "type": kind,
        "source": source,
        "target": target,
        "confidence": confidence,
    }
    if group_id:
        relation["group_id"] = group_id
    return relation


def _normalized_name(variant):
    value = " ".join(filter(None, [variant.get("id"), variant.get("layout_name")]))
    return re.sub(r"[_\-]+", " ", value.casefold())


def _contains(text, *needles):
    return any(needle in text for needle in needles)


def _infer_layout_role(variant):
    name = _normalized_name(variant)
    text_slots = variant.get("slots", {})
    image_slots = variant.get("image_slots", {})
    editable_text = [slot for slot in text_slots.values() if not _is_decorative(slot)]
    keyword_rules = [
        ("closing_qr", ("qr", "qr-код", "куар")),
        ("speaker", ("спикер", "speaker", "докладчик")),
        ("section_divider", ("раздел", "section", "divider")),
        ("closing", ("финал", "final", "closing", "спасибо", "thank")),
        ("metrics", ("фактоид", "метрик", "metric", "kpi", "показател")),
        ("quote", ("цитат", "quote", "отзыв")),
        ("agenda", ("содержание", "agenda", "план выступ")),
        ("comparison", ("сравнен", "comparison", "до и после", "versus")),
        ("photo_with_caption", ("фото + подпись", "photo with caption", "picture with caption")),
        ("content_with_image", ("фото", "photo", "picture", "изображен")),
        ("two_columns", ("2 объект", "два объект", "2 колон", "две колон", "two content", "two column")),
        ("cover", ("титульн", "title slide", "cover")),
    ]
    for role, keywords in keyword_rules:
        if _contains(name, *keywords):
            return role, 0.94, ["layout_name"]
    if image_slots and len(editable_text) >= 2:
        return "content_with_image", 0.78, ["slot_topology"]
    if len(editable_text) == 1 and "title" in text_slots:
        return "title_only", 0.76, ["slot_topology"]
    if len(editable_text) >= 3:
        body_slots = [name for name in text_slots if name not in {"title", "subtitle", "footer"}]
        if len(body_slots) == 2:
            return "two_columns", 0.70, ["slot_topology"]
    if editable_text:
        return "content", 0.68, ["slot_topology"]
    return "generic", 0.45, ["fallback"]


def _is_decorative(slot):
    if slot.get("is_placeholder"):
        return False
    template_text = str(slot.get("template_text") or "").strip()
    if template_text and not any(character.isalnum() for character in template_text):
        return True
    area = max(slot.get("w", 0), 0) * max(slot.get("h", 0), 0)
    return slot.get("w", 0) <= 0.35 or slot.get("h", 0) <= 0.2 or area <= 0.25


def _enrich_text_slot(name, slot, variant, layout_role, slide_width, slide_height):
    decorative = _is_decorative(slot)
    role, confidence = _text_role(name, slot, variant, layout_role)
    if decorative:
        role, confidence = "decorative", 0.99
    slot["editable"] = not decorative
    slot["decorative"] = decorative
    slot["semantic_role"] = role
    slot["role_confidence"] = confidence
    slot["content_kind"] = _text_content_kind(role)
    slot["generator_field"] = "content"
    slot["accepted_value"] = "string"
    slot["required"] = not decorative and role in {
        "presentation_title", "section_title", "slide_title", "speaker_name",
        "closing_title", "metric_value",
    }
    slot["instructions"] = ROLE_GUIDANCE.get(role, "Не заполняй этот декоративный объект.")
    slot["capacity"] = _estimate_capacity(slot, role) if not decorative else None
    slot["relative_position"] = _relative_position(slot, slide_width, slide_height)
    slot["semantic_scope"] = _slot_scope(slot, "text")
    slot["allowed_content_kinds"] = _allowed_content_kinds(slot, "text")


def _text_role(name, slot, variant, layout_role):
    native = str(slot.get("placeholder_type") or "").upper()
    if name == "footer" or "FOOTER" in native:
        return "footer", 0.99
    if "SUBTITLE" in native or name == "subtitle":
        return ("presentation_subtitle" if layout_role == "cover" else "subtitle"), 0.98
    if name == "title" or "TITLE" in native:
        mapping = {
            "cover": "presentation_title",
            "section_divider": "section_title",
            "speaker": "speaker_name",
            "closing": "closing_title",
            "closing_qr": "closing_title",
        }
        return mapping.get(layout_role, "slide_title"), 0.98
    role_by_layout = {
        "cover": "presentation_subtitle",
        "speaker": "speaker_bio",
        "quote": "quote",
        "closing": "contact_text",
        "closing_qr": "contact_text",
        "two_columns": "column_body",
        "comparison": "column_body",
    }
    if layout_role in role_by_layout:
        return role_by_layout[layout_role], 0.82
    return "body", 0.72 if slot.get("is_placeholder") else 0.58


def _text_content_kind(role):
    if role == "metric_value":
        return "number_or_short_value"
    if role in {"body", "column_body", "speaker_bio", "quote"}:
        return "paragraph_or_bullets"
    if role in {
        "footer", "image_source", "image_caption", "metric_label", "subtitle",
        "presentation_subtitle", "contact_text", "quote_author",
    }:
        return "short_text"
    if role == "decorative":
        return "none"
    return "short_heading"


def _slot_scope(slot, kind):
    if slot.get("decorative") or not slot.get("editable"):
        return "decorative"
    role = slot.get("semantic_role")
    confidence = slot.get("role_confidence", 0)
    generic = "generic_image" if kind == "image" else "generic_text"
    if role == generic:
        return "flexible"
    if confidence >= 0.65:
        return "specific"
    if kind == "image" or slot.get("capacity"):
        return "flexible"
    return "unknown"


def _allowed_content_kinds(slot, kind):
    scope = slot.get("semantic_scope")
    if scope == "decorative":
        return []
    if scope == "unknown":
        return ["generic_image" if kind == "image" else "generic_text"]
    role = slot.get("semantic_role")
    if scope == "specific":
        if kind == "image":
            return [{
                "speaker_portrait": "portrait_photo",
                "qr_code": "qr_code",
                "photo": "photo",
                "illustration": "illustration",
            }.get(role, "image")]
        content_kind = slot.get("content_kind")
        if content_kind == "paragraph_or_bullets":
            return ["paragraph", "bullets"]
        return [content_kind]
    if kind == "image":
        return ["photo", "illustration", "screenshot", "diagram"]
    capacity = slot.get("capacity") or {}
    max_lines = capacity.get("maximum_lines", 1)
    max_characters = capacity.get("maximum_characters", 1)
    result = ["number_or_short_value", "short_text"]
    if max_lines <= 2:
        result.append("short_heading")
    if max_characters <= 200:
        result.append("caption")
    if max_lines >= 2:
        result.extend(["paragraph", "bullets"])
    return list(dict.fromkeys(result))


def _layout_scope(layout_role, confidence, evidence, flexible_slots):
    if evidence == ["layout_name"] and layout_role != "generic":
        return "specific"
    if flexible_slots or layout_role in {
        "content", "content_with_image", "two_columns", "title_only",
    }:
        return "flexible"
    if confidence < 0.65:
        return "unknown"
    return "specific"


def _estimate_capacity(slot, role):
    font = slot.get("font") or {}
    explicit_size = font.get("size")
    fallback = {
        "presentation_title": 44, "section_title": 38, "slide_title": 30,
        "speaker_name": 36, "closing_title": 40, "metric_value": 40,
        "metric_label": 18, "footer": 12, "image_source": 12,
        "image_caption": 16, "body": 18, "column_body": 18,
    }.get(role, 18)
    font_pt = float(explicit_size or fallback)
    margins = slot.get("margins") or {}
    usable_w = max(0.1, slot.get("w", 0) - (margins.get("left") or 0) - (margins.get("right") or 0))
    usable_h = max(0.1, slot.get("h", 0) - (margins.get("top") or 0) - (margins.get("bottom") or 0))
    pt_cm = 0.0352778
    line_height = max(0.1, font_pt * pt_cm * 1.20)
    average_char_width = max(0.05, font_pt * pt_cm * 0.52)
    max_lines = max(1, math.floor(usable_h / line_height))
    chars_per_line = max(1, math.floor(usable_w / average_char_width))
    maximum = max_lines * chars_per_line
    return {
        "font_size_pt": round(font_pt, 2),
        "font_size_source": "explicit" if explicit_size else "role_fallback",
        "recommended_lines": max(1, math.floor(max_lines * 0.75)),
        "maximum_lines": max_lines,
        "recommended_characters": max(1, math.floor(maximum * 0.70)),
        "maximum_characters": maximum,
        "estimate_only": True,
    }


def _relative_position(slot, slide_width, slide_height):
    cx = (slot.get("x", 0) + slot.get("w", 0) / 2) / max(slide_width, 0.1)
    cy = (slot.get("y", 0) + slot.get("h", 0) / 2) / max(slide_height, 0.1)
    horizontal = "left" if cx < 0.4 else "right" if cx > 0.6 else "center"
    vertical = "top" if cy < 0.35 else "bottom" if cy > 0.65 else "middle"
    return f"{vertical}_{horizontal}"


def _assign_metric_groups(variant):
    slots = variant.get("slots", {})
    candidates = [
        (name, slot) for name, slot in slots.items()
        if slot.get("editable") and slot.get("semantic_role") not in {"slide_title", "footer"}
    ]
    if not candidates:
        return
    areas = {name: slot.get("w", 0) * slot.get("h", 0) for name, slot in candidates}
    largest_name = max(areas, key=areas.get)
    if areas[largest_name] > 40:
        slot = slots[largest_name]
        slot.update(
            semantic_role="body", role_confidence=0.76,
            content_kind="paragraph_or_bullets", required=False,
            instructions=ROLE_GUIDANCE["body"],
        )
        candidates = [(name, slot) for name, slot in candidates if name != largest_name]
    columns = []
    for name, slot in sorted(candidates, key=lambda item: (item[1].get("x", 0), item[1].get("y", 0))):
        center = slot.get("x", 0) + slot.get("w", 0) / 2
        group = next((item for item in columns if abs(item["center"] - center) <= max(slot.get("w", 0), 1) * 0.55), None)
        if group is None:
            group = {"center": center, "items": []}
            columns.append(group)
        group["items"].append((name, slot))
    for index, group in enumerate(columns, start=1):
        items = sorted(group["items"], key=lambda item: item[1].get("y", 0))
        for item_index, (_, slot) in enumerate(items):
            role = "metric_value" if item_index == 0 else "metric_label"
            slot.update(
                semantic_role=role,
                role_confidence=0.80,
                content_kind=_text_content_kind(role),
                required=role == "metric_value",
                instructions=ROLE_GUIDANCE[role],
                group_id=f"metric_{index}",
            )
            slot["capacity"] = _estimate_capacity(slot, role)


def _assign_caption_roles(variant):
    images = list(variant.get("image_slots", {}).values())
    if not images:
        return
    image = images[0]
    below = []
    for name, slot in variant.get("slots", {}).items():
        if not slot.get("editable") or slot.get("semantic_role") == "slide_title":
            continue
        overlaps_x = min(slot["x"] + slot["w"], image["x"] + image["w"]) > max(slot["x"], image["x"])
        if overlaps_x and slot["y"] >= image["y"] + image["h"] * 0.75:
            below.append((name, slot))
    for index, (_, slot) in enumerate(sorted(below, key=lambda item: item[1]["y"])):
        role = "image_caption" if index == 0 else "image_source"
        slot.update(
            semantic_role=role,
            role_confidence=0.86,
            content_kind="short_text",
            required=False,
            instructions=ROLE_GUIDANCE[role],
            group_id="image_1",
        )
        slot["capacity"] = _estimate_capacity(slot, role)


def _assign_quote_roles(variant):
    """Separate the quotation itself from a compact author/source line."""
    candidates = [
        (name, slot) for name, slot in variant.get("slots", {}).items()
        if slot.get("editable") and slot.get("semantic_role") == "quote"
    ]
    if len(candidates) < 2:
        return
    author_name, author = min(
        candidates,
        key=lambda item: (item[1].get("h", 0), -item[1].get("y", 0)),
    )
    other_heights = [slot.get("h", 0) for name, slot in candidates if name != author_name]
    if other_heights and author.get("h", 0) <= min(other_heights) * 0.6:
        author.update(
            semantic_role="quote_author",
            role_confidence=0.84,
            content_kind="short_text",
            required=False,
            instructions=ROLE_GUIDANCE["quote_author"],
            group_id="quote_1",
        )
        author["capacity"] = _estimate_capacity(author, "quote_author")
        for name, slot in candidates:
            if name != author_name:
                slot["group_id"] = "quote_1"


def _enrich_image_slot(name, slot, variant, layout_role):
    mapping = {
        "speaker": ("speaker_portrait", 0.96),
        "closing_qr": ("qr_code", 0.98),
        "photo_with_caption": ("photo", 0.94),
        "content_with_image": ("photo", 0.82),
    }
    role, confidence = mapping.get(layout_role, ("generic_image", 0.58))
    width, height = slot.get("w", 0), slot.get("h", 0)
    ratio = round(width / height, 3) if height else None
    orientation = "square" if ratio and 0.9 <= ratio <= 1.1 else "landscape" if ratio and ratio > 1.1 else "portrait"
    slot.update({
        "editable": True,
        "decorative": False,
        "semantic_role": role,
        "role_confidence": confidence,
        "content_kind": "image",
        "generator_field": "image_content",
        "accepted_value": "existing_file_path",
        "required": layout_role in {"speaker", "photo_with_caption", "closing_qr"},
        "instructions": ROLE_GUIDANCE[role],
        "aspect_ratio": ratio,
        "orientation": orientation,
        "crop_mode": "cover",
    })
    if layout_role == "photo_with_caption":
        slot["group_id"] = "image_1"
    slot["semantic_scope"] = _slot_scope(slot, "image")
    slot["allowed_content_kinds"] = _allowed_content_kinds(slot, "image")


def _reading_order(variant):
    items = []
    for name, slot in variant.get("slots", {}).items():
        if slot.get("editable"):
            role = slot.get("semantic_role")
            priority = 0 if role in {
                "presentation_title", "section_title", "slide_title", "speaker_name", "closing_title"
            } else 3 if role in {"footer", "image_source"} else 2 if role == "image_caption" else 1
            items.append((priority, slot.get("y", 0), slot.get("x", 0), "text", name))
    for name, slot in variant.get("image_slots", {}).items():
        if slot.get("editable"):
            items.append((1, slot.get("y", 0), slot.get("x", 0), "image", name))
    return [
        {"kind": kind, "slot": name}
        for _, _, _, kind, name in sorted(items)
    ]


def _slot_groups(variant):
    grouped = {}
    for kind, slots in (("text", variant.get("slots", {})), ("image", variant.get("image_slots", {}))):
        for name, slot in slots.items():
            group_id = slot.get("group_id")
            if group_id:
                grouped.setdefault(group_id, []).append({"kind": kind, "slot": name})
    return [
        {"id": group_id, "members": members}
        for group_id, members in sorted(grouped.items())
    ]


def _content_budget(variant):
    capacities = [
        slot["capacity"] for slot in variant.get("slots", {}).values()
        if slot.get("editable") and slot.get("capacity")
    ]
    return {
        "recommended_characters_total": sum(item["recommended_characters"] for item in capacities),
        "maximum_characters_total": sum(item["maximum_characters"] for item in capacities),
        "editable_text_slots": sum(bool(slot.get("editable")) for slot in variant.get("slots", {}).values()),
        "editable_image_slots": sum(bool(slot.get("editable")) for slot in variant.get("image_slots", {}).values()),
        "estimate_only": True,
    }


def _review_status(layout_scope, ambiguities):
    reasons = []
    if layout_scope == "unknown":
        reasons.append("unknown_layout_purpose")
    if ambiguities:
        reasons.append("unknown_slot_roles")
    return {
        "recommended": bool(reasons),
        "reasons": reasons,
        "can_use_without_review": not reasons,
    }


def _downstream_choice(layout_scope, flexible_slots):
    allowed = layout_scope == "flexible" or bool(flexible_slots)
    return {
        "allowed": allowed,
        "reason": ("template_supports_multiple_valid_content_interpretations"
                   if allowed else None),
        "flexible_slots": flexible_slots,
    }


def _avoid_when(variant, layout_role):
    result = []
    if not variant.get("image_slots"):
        result.append("content_requires_replaceable_image")
    if layout_role == "title_only":
        result.append("content_needs_body_text")
    if layout_role in {"closing", "closing_qr"}:
        result.append("ordinary_content_slide")
    if any(
        slot.get("capacity") and slot["capacity"]["maximum_lines"] <= 2
        for slot in variant.get("slots", {}).values()
        if slot.get("semantic_role") in {"body", "column_body"}
    ):
        result.append("long_body_text")
    return result
