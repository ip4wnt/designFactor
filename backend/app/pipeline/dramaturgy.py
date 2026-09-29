"""Драматургия презентации: выбор композиции для каждого слайда и три варианта
по оси плотности (Лаконичный / Сбалансированный / Насыщенный).

Вход — ContentPlan (структура и тезисы) и Template JSON, в котором наряду с
макетами (`source: layout`) могут быть композиции слайдов-образцов
(`source: sample`, см. sample_slides.py) и синтетические композиции из
собственной библиотеки (`source: library`, см. library.py). Выход — план в
формате движка `{"slides": [{"variant", "content", "image_content"}]}`,
который дальше рендерит template_engine.render_variant.

Правила детерминированы (без LLM): обложка → финал, разделители, метрики для
«Цифр», таймлайн для дат/этапов, карточки для 3–8 тезисов, таблица/график для
соответствующих блоков; ритм — не более двух одинаковых композиций подряд.
Тексты для слотов берутся из «структурированного контента» (structure_content:
один вызов модели на колоду с детерминированным запасным вариантом).
"""
from __future__ import annotations

import datetime as _dt
import logging
import math
import re
from typing import Any

from app.schemas.content_plan import ContentBlockType, ContentPlan
from app.skills import load_skill, skill_text

logger = logging.getLogger(__name__)

DENSITY_PROFILES: dict[str, dict[str, Any]] = {
    "variant_a": {
        "label": "Лаконичный",
        "description": (
            "Меньше слов на слайде: 3–4 ключевых блока, только заголовки тезисов, "
            "короткие подписи, без слайдов-разделителей — самая короткая колода. "
            "Подходит для выступления, где детали проговаривает спикер."
        ),
        "sections": False,
        "max_units": 4,
        "unit_text": False,
        "text_chars": 0,
        "body_sentences": 1,
        "note": False,
        "images": False,
        # Характер варианта: крупная типографика, воздух, один тезис на слайд
        "prefer": {"statement": -0.9, "section": -0.4, "metrics": -0.3, "image_text": -0.4, "columns": -0.2},
        "avoid": {"cards": 0.3, "table": 0.2},
        "rotation": 0,
    },
    "variant_b": {
        "label": "Сбалансированный",
        "description": (
            "Заголовок и короткое пояснение к каждому тезису, до 6 блоков на слайде, "
            "без разделителей. Универсальный вариант для отправки и для показа."
        ),
        "sections": False,
        "max_units": 6,
        "unit_text": True,
        "text_chars": 110,
        "body_sentences": 2,
        "note": False,
        "images": False,
        # Характер варианта: структура — карточки, колонки, таймлайны
        "prefer": {"cards": -0.4, "columns": -0.4, "timeline": -0.3, "problems": -0.3},
        "avoid": {"statement": 0.4},
        "rotation": 1,
    },
    "variant_c": {
        "label": "Насыщенный",
        "description": (
            "Все тезисы с развёрнутыми пояснениями, выводы-заметки, таблицы, графики "
            "и изображения там, где макет их предусматривает; смысловые части открывают "
            "слайды-разделители — колода длиннее. Для чтения без спикера."
        ),
        "sections": True,
        "max_units": 8,
        "unit_text": True,
        "text_chars": 170,
        "body_sentences": 3,
        "note": True,
        "images": True,
        # Характер варианта: насыщенность — фото, текст с пояснениями, таблицы
        "prefer": {"image_text": -0.7, "text": -0.3, "table": -0.3, "cards": -0.2},
        "avoid": {"statement": 0.5, "title_only": 0.5},
        "rotation": 2,
    },
}

_METRIC_RE = re.compile(r"(?<![\w.])([+\-−×x]?\s?\d[\d\s\u00a0]*(?:[.,]\d+)?\s*(?:%|млн|млрд|тыс\.?|руб\.?|₽|\$|€|x|×|раз|человек|чел\.|дней|часов|лет|месяцев|пунктов|п\.п\.)?)", re.I)
_DATE_RE = re.compile(r"\b(?:20\d\d|19\d\d|Q[1-4]|[1-4]\s*кв\.?|январ|феврал|март|апрел|ма[йя]|июн|июл|август|сентябр|октябр|ноябр|декабр|этап|шаг|неделя|недел|день|месяц|квартал|полугоди|фаза|спринт)", re.I)
_PROBLEM_RE = re.compile(r"проблем|вызов|риск|барьер|боль|сложност|ограничен", re.I)
_AGENDA_RE = re.compile(r"содержан|оглавлен|план презентац|о чём|повестк|структура презентац", re.I)
_INSIST_RE = re.compile(r"(?:ровно|обязательно|строго|именно|не меньше|минимум|все)\s+(\d{1,2})\s*(?:тезис|пункт|блок|карточ|шаг|этап)|(\d{1,2})\s*(?:тезис|пункт|блок|карточ)\w*\s+(?:обязательно|ровно|строго|на слайде|на одном слайде)", re.I)

_MONTHS = ["января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", "сентября", "октября", "ноября", "декабря"]


# ---------------------------------------------------------------------------
# Каталог композиций
# ---------------------------------------------------------------------------


def _layout_kind(variant: dict[str, Any], cat: dict[str, Any]) -> tuple[str, int]:
    """Вид композиции и число блоков для макета из парсера (source=layout)."""
    slots = variant.get("slots") or {}
    roles = [s.get("semantic_role") for s in slots.values()]
    role = cat.get("narrative_role")
    pattern = cat.get("content_pattern")
    comp = cat.get("composition")
    metrics = sum(1 for r in roles if r == "metric_value")
    columns = sum(1 for r in roles if r == "column_body")
    has_body = any(r in {"body", "column_body"} for r in roles)
    if role == "cover":
        return "cover", 0
    if role == "closing":
        return "closing", 0
    if role == "section":
        return "section", 0
    name = (variant.get("id") or "").lower()
    if role == "speaker" or pattern == "person_profile" or re.search(r"визитк|спикер|команд|контакт|team|speaker", name):
        return "speaker", 0
    if re.search(r"содержан|оглавлен|agenda|contents", name):
        return "agenda", 0
    if pattern == "metrics" or metrics >= 2:
        return "metrics", max(metrics, 1)
    if pattern == "quote":
        return "statement", 0
    if columns >= 2:
        return "columns", columns
    if cat.get("has_image") and has_body:
        return "image_text", 0
    if has_body:
        return "text", 0
    if comp == "single_focus" and not has_body:
        return "title_only", 0
    return "text" if has_body else "title_only", 0


def catalog_compositions(template: dict[str, Any]) -> list[dict[str, Any]]:
    """Единый список композиций (образцы, макеты, библиотека) с общими полями."""
    catalog = {item["id"]: item for item in template.get("variant_catalog") or []}
    result: list[dict[str, Any]] = []
    sw = float((template.get("slide_size") or {}).get("width") or 0)
    sh = float((template.get("slide_size") or {}).get("height") or 0)
    for variant in template.get("variants") or []:
        source = variant.get("source") or "layout"
        slots = variant.get("slots") or {}
        if source == "layout" and sw and sh and any(
            float(s.get("x") or 0) < -0.2 or float(s.get("y") or 0) < -0.2
            or float(s.get("x") or 0) + float(s.get("w") or 0) > sw + 0.2
            or float(s.get("y") or 0) + float(s.get("h") or 0) > sh + 0.2
            for s in slots.values() if not s.get("virtual")
        ):
            continue  # макет с плейсхолдером за границей слайда — рендер его отвергнет
        if source == "layout" and template.get("source_mode") == "slides" and template.get("design_source") == "library":
            # Авторские слайды без слота заголовка: контент с заголовком на них
            # не собрать, а движок не умеет их повторять — колода целиком из
            # библиотеки на шрифтах и цветах шаблона
            continue
        roles = {s.get("semantic_role") for s in slots.values()}
        entry: dict[str, Any] = {
            "id": variant["id"],
            "source": source,
            "has_body": bool(roles & {"body", "column_body"}),
            "has_note": "note" in roles,
            "has_subtitle": bool(roles & {"subtitle", "subheading", "presentation_subtitle"}),
            "image_slots": [n for n, s in (variant.get("image_slots") or {}).items() if s.get("semantic_role") != "qr_code"],
            "qr_slots": [n for n, s in (variant.get("image_slots") or {}).items() if s.get("semantic_role") == "qr_code"],
        }
        if source in {"sample", "library"}:
            comp = variant.get("composition") or {}
            entry.update(
                kind=comp.get("kind") or "text",
                units=int(comp.get("units") or 0),
                min_units=int(comp.get("min_units") or 0),
                max_units=int(comp.get("max_units") or comp.get("units") or 0),
                grid=comp.get("grid") or {},
                fixed_units=bool(comp.get("fixed_units")),
                unit_roles=sorted({slots[n]["semantic_role"] for names in (comp.get("unit_slots") or []) for n in names if n in slots}),
                metrics=sum(1 for r in (s.get("semantic_role") for s in slots.values()) if r == "metric_value"),
                hero_metrics=sum(1 for s in slots.values() if s.get("semantic_role") == "metric_value" and s.get("unit") is None and not s.get("virtual")),
                table=bool(comp.get("table_shape_id")),
                chart=bool(comp.get("chart_pictures")),
                charts=len(comp.get("chart_pictures") or []),
            )
        else:
            kind, units = _layout_kind(variant, catalog.get(variant["id"], {}))
            entry.update(kind=kind, units=units, min_units=units, max_units=units, unit_roles=[], metrics=units if kind == "metrics" else 0, table=False, chart=False, charts=0)
        result.append(entry)
    return result


# ---------------------------------------------------------------------------
# Намерение слайда
# ---------------------------------------------------------------------------


def _bullets(spec) -> list[str]:
    items: list[str] = []
    for block in spec.content_blocks:
        if block.type == ContentBlockType.BULLETS and block.bullets:
            items.extend(b.strip() for b in block.bullets if b and b.strip())
    return items


def _texts(spec) -> list[str]:
    return [block.text.strip() for block in spec.content_blocks if block.type == ContentBlockType.TEXT and block.text and block.text.strip()]


def _has(spec, kind: ContentBlockType) -> bool:
    return any(block.type == kind for block in spec.content_blocks)


_KICKER_RE = re.compile(r"^\s*(часть|раздел|глава|блок|этап|шаг)\s*(\d{1,2}|[ivx]{1,4})\s*[:.\-–—]\s*(.+)$", re.I)


def split_kicker(title: str) -> tuple[str, str]:
    """«Часть 1: Физическое здоровье» → («Часть 1», «Физическое здоровье»)."""
    m = _KICKER_RE.match(title or "")
    if not m or len(m.group(3).strip()) < 3:
        return "", (title or "").strip()
    return f"{m.group(1).capitalize()} {m.group(2).upper() if m.group(2).isalpha() else m.group(2)}", m.group(3).strip()


def is_section_slide(index: int, spec, total: int) -> bool:
    if index == 0 or index == total - 1:
        return False
    purpose = (spec.purpose or "").strip().lower()
    return purpose.startswith("раздел") or (not _bullets(spec) and not _texts(spec) and not _has(spec, ContentBlockType.TABLE) and not _has(spec, ContentBlockType.CHART))


def section_map(content_plan: ContentPlan) -> dict[int, int]:
    """Индекс слайда-разделителя → число контентных слайдов внутри раздела."""
    total = len(content_plan.slides)
    result: dict[int, int] = {}
    for index, spec in enumerate(content_plan.slides):
        if not is_section_slide(index, spec, total):
            continue
        inside = 0
        for nxt_index, nxt in enumerate(content_plan.slides[index + 1:-1], start=index + 1):
            if is_section_slide(nxt_index, nxt, total):
                break
            inside += 1
        result[index] = inside
    return result


def dropped_slides(content_plan: ContentPlan, density_key: str) -> set[int]:
    """Слайды плана, которых в варианте не будет: разделители — в вариантах без
    разделителей все, в варианте с разделителями — открывающие меньше двух
    слайдов (такой разделитель только повторяет заголовок соседа)."""
    keep_sections = bool(DENSITY_PROFILES[density_key].get("sections"))
    return {index for index, inside in section_map(content_plan).items() if not keep_sections or inside < 2}


def slide_intent(index: int, spec, total: int, content_plan: ContentPlan) -> dict[str, Any]:
    purpose = (spec.purpose or "").strip().lower()
    title = (spec.title or "").strip()
    bullets = _bullets(spec)
    texts = _texts(spec)
    intent: dict[str, Any] = {"items": len(bullets), "has_text": bool(texts), "kinds": []}
    if index == 0:
        intent["kinds"] = ["cover"]
        return intent
    if index == total - 1:
        intent["kinds"] = ["closing", "cta", "qa", "cover"]
        return intent
    if _has(spec, ContentBlockType.TABLE):
        intent["kinds"] = ["table", "text", "columns", "cards"]
        return intent
    if _has(spec, ContentBlockType.CHART):
        intent["kinds"] = ["chart", "image_text", "text", "columns", "cards"]
        return intent
    if purpose.startswith("раздел") or (not bullets and not texts):
        intent["kinds"] = ["section", "statement", "cards", "title_only"]
        # Разделитель без содержания выглядит пустым: подсказываем, что внутри —
        # заголовки следующих слайдов до следующего раздела (≤4)
        inside: list[str] = []
        for nxt in content_plan.slides[index + 1:-1]:
            if (nxt.purpose or "").lower().startswith("раздел"):
                break
            if nxt.title:
                inside.append(nxt.title.strip())
            if len(inside) >= 4:
                break
        intent["section_inside"] = inside
        return intent
    if index == 1 and (_AGENDA_RE.search(purpose) or _AGENDA_RE.search(title)):
        sections = [s.title for s in content_plan.slides[2:-1] if (s.purpose or "").lower().startswith("раздел")]
        intent["agenda"] = sections if len(sections) >= 3 else (bullets or sections)
        intent["items"] = len(intent["agenda"])
        intent["kinds"] = ["agenda", "cards", "text"]
        return intent
    if len(bullets) >= 3 and sum(1 for b in bullets if _DATE_RE.match(b.strip())) >= max(2, math.ceil(len(bullets) * 0.6)):
        intent["timeline"] = True
        intent["kinds"] = ["timeline", "cards", "text"]
        return intent
    metric_like = sum(1 for b in bullets if _looks_metric(b))
    if purpose.startswith("цифры") or (bullets and metric_like >= max(2, math.ceil(len(bullets) * 0.6))):
        intent["metrics"] = True
        intent["kinds"] = ["metrics", "cards", "columns", "text"]
        return intent
    if _PROBLEM_RE.search(purpose) or _PROBLEM_RE.search(title):
        intent["kinds"] = ["problems", "cards", "columns", "text"]
        return intent
    if len(bullets) >= 3:
        intent["kinds"] = ["cards", "problems", "columns", "text", "image_text"]
    elif len(bullets) == 2:
        intent["kinds"] = ["columns", "cards", "text", "image_text"]
    elif bullets:
        intent["kinds"] = ["text", "image_text", "statement", "cards"]
    else:
        # Только текст: тезис/абзац; если в шаблоне нет текстовых композиций —
        # предложения абзаца станут карточками
        intent["sentences"] = True
        intent["kinds"] = ["text", "statement", "image_text", "cards"]
    return intent


_YEAR_RE = re.compile(r"^(?:19|20)\d\d\s*(?:г\.?|год[а-я]*)?$")


def _slot_reading_order(slots: dict[str, Any], names: list[str]):
    """Ключ сортировки слотов одной роли «как читают»: столбик (одинаковый x) —
    сверху вниз, иначе — слева направо. Так главная цифра попадает в верхний /
    левый слот, а подпись — парой к своему значению."""
    xs = [float(slots[n].get("x") or 0) for n in names]
    stacked = bool(xs) and (max(xs) - min(xs)) < 1.0

    def key(name: str):
        sl = slots.get(name) or {}
        x, y = float(sl.get("x") or 0), float(sl.get("y") or 0)
        return (round(y, 1), x) if stacked else (round(x, 1), y)

    return key


def _metrics_from_bullets(bullets: list[str], items: list[dict[str, Any]], exclude: set[str] | None = None) -> list[dict[str, Any]]:
    """Цифры из тезисов плана (и текстов блоков) как метрики {value, label}.
    Годы и номера кварталов не считаются; повторы значений из `exclude` пропускаем."""
    exclude = set(exclude or ())
    out: list[dict[str, Any]] = []
    sources = list(bullets) + [f"{i.get('head', '')} — {i.get('text', '')}" for i in items]
    for text in sources:
        if not text or not _looks_metric(text):
            continue
        match = _METRIC_RE.search(text)
        if not match:
            continue
        value = re.sub(r"\s+", " ", match.group(1)).strip()
        if value in exclude:
            continue
        label = (text[: match.start()] + text[match.end():]).strip(" ,.:—–-")
        exclude.add(value)
        out.append({"value": value, "label": _head_of(label, 5) if label else ""})
    return out


def _looks_metric(text: str) -> bool:
    match = _METRIC_RE.search(text)
    if not match or not re.search(r"\d", match.group(1)):
        return False
    value = match.group(1).strip()
    if _YEAR_RE.match(value) or re.match(r"^Q[1-4]", text.strip(), re.I):
        return False
    has_unit = bool(re.search(r"%|млн|млрд|тыс|руб|₽|\$|€|×|x\b|раз|человек|чел|дней|часов|лет|месяц|пункт|п\.п", value, re.I))
    at_start = match.start() <= 2 or text.strip().startswith(("×", "x"))
    return has_unit or at_start


# ---------------------------------------------------------------------------
# Структурированный контент (один вызов модели на колоду + запасной вариант)
# ---------------------------------------------------------------------------

# Промпт структурирования лежит в skills/content_structuring.yaml (не в коде).
_SKILL = "content_structuring"


def _structure_system() -> str:
    return skill_text(_SKILL, "system_prompt")


def _rewrite_hint() -> str:
    return "\n" + skill_text(_SKILL, "rewrite_hint")


def _skill_param(key: str, default):
    value = load_skill(_SKILL).get(key)
    return default if value is None else value


def structure_content(content_plan: ContentPlan, brief: str, model) -> list[dict[str, Any]]:
    """Структурированные тексты по слайдам. Модель вызывается порциями по 8
    слайдов; при любой ошибке используется детерминированный запасной вариант."""
    fallback = [structure_fallback(index, spec, len(content_plan.slides)) for index, spec in enumerate(content_plan.slides)]
    if model is None:
        return fallback
    result = list(fallback)
    slides = list(content_plan.slides)
    chunk = int(_skill_param("chunk_slides", 8))
    for start in range(0, len(slides), chunk):
        lines = [f"БРИФ: {brief.strip()[:1200]}", "", "СЛАЙДЫ:"]
        for offset, spec in enumerate(slides[start:start + chunk]):
            index = start + offset
            lines.append(f"Слайд {index + 1}. {spec.title.strip()} — {spec.purpose.strip()}")
            for item in _bullets(spec):
                lines.append(f"  • {item}")
            for text in _texts(spec):
                lines.append(f"  {text}")
            if _has(spec, ContentBlockType.CHART):
                lines.append("  [на слайде будет график]")
            if _has(spec, ContentBlockType.TABLE):
                lines.append("  [на слайде будет таблица]")
        try:
            response = model.json(_structure_system(), "\n".join(lines), temperature=float(_skill_param("temperature", 0.2)))
            for item in response.get("slides") or []:
                index = int(item.get("index") or 0) - 1
                if not (start <= index < min(start + chunk, len(slides))):
                    continue
                merged = _merge_structured(fallback[index], item, slides[index])
                result[index] = merged
        except Exception as exc:  # noqa: BLE001 — запасной вариант уже есть
            logger.warning("Структурирование контента (слайды %d–%d) не удалось: %s", start + 1, start + chunk, exc)
    return result


def restructure_slide(content_plan: ContentPlan, brief: str, model, index: int, previous: dict[str, Any] | None = None) -> dict[str, Any]:
    """Переписать тексты одного слайда (ручная доработка пользователя).
    Возвращает новую структуру слайда; при ошибке модели — детерминированный
    запасной вариант."""
    total = len(content_plan.slides)
    spec = content_plan.slides[index]
    fallback = structure_fallback(index, spec, total)
    if model is None:
        return fallback
    lines = [f"БРИФ: {brief.strip()[:1200]}", "", "СЛАЙДЫ:", f"Слайд {index + 1}. {spec.title.strip()} — {spec.purpose.strip()}"]
    for item in _bullets(spec):
        lines.append(f"  • {item}")
    for text in _texts(spec):
        lines.append(f"  {text}")
    if _has(spec, ContentBlockType.CHART):
        lines.append("  [на слайде будет график]")
    if _has(spec, ContentBlockType.TABLE):
        lines.append("  [на слайде будет таблица]")
    if previous:
        lines += ["", "ПРЕДЫДУЩАЯ ВЕРСИЯ (так больше не писать):", f"  title: {previous.get('title', '')}"]
        for item in previous.get("items") or []:
            lines.append(f"  • {item.get('head', '')} — {item.get('text', '')}")
        if previous.get("body"):
            lines.append(f"  body: {previous['body']}")
    try:
        response = model.json(_structure_system() + _rewrite_hint(), "\n".join(lines), temperature=float(_skill_param("rewrite_temperature", 0.8)))
        for item in response.get("slides") or []:
            if int(item.get("index") or 0) - 1 == index:
                return _merge_structured(fallback, item, spec)
    except Exception as exc:  # noqa: BLE001
        logger.warning("Переписать слайд %d не удалось: %s", index + 1, exc)
    return fallback


def _merge_structured(base: dict[str, Any], item: dict[str, Any], spec) -> dict[str, Any]:
    out = dict(base)
    title = _clean(item.get("title"))
    if title:
        out["title"] = title
    out["subtitle"] = _clean(item.get("subtitle")) or base.get("subtitle", "")
    out["body"] = _clean(item.get("body")) or base.get("body", "")
    out["note"] = _clean(item.get("note"))
    items = []
    for raw in item.get("items") or []:
        if not isinstance(raw, dict):
            continue
        head = _clean(raw.get("head"))
        text = _clean(raw.get("text"))
        if not head and not text:
            continue
        items.append({"head": head or _head_of(text), "text": text, "label": _clean(raw.get("label"))})
    bullets = _bullets(spec)
    if items and (not bullets or abs(len(items) - len(bullets)) <= max(1, len(bullets) // 3)):
        out["items"] = items
    metrics = []
    for raw in item.get("metrics") or []:
        if isinstance(raw, dict) and _clean(raw.get("value")):
            metrics.append({"value": _clean(raw.get("value")), "label": _clean(raw.get("label"))})
    if metrics:
        out["metrics"] = metrics
    return out


def _clean(value) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text in {"—", "-", "–", "...", "…", "null", "None"}:
        return ""
    return re.sub(r"\s+", " ", text)


def _head_of(text: str, max_words: int = 4) -> str:
    text = text.strip().rstrip(".")
    for sep in (":", " — ", " – ", ". "):
        if sep in text:
            head = text.split(sep, 1)[0].strip()
            if 1 <= len(head.split()) <= 6:
                return head
    words = text.split()
    cut = words[:max_words]
    # Обрезанный заголовок не должен кончаться предлогом/союзом («Документы и задачи в»)
    while len(cut) > 1 and _HANGING_WORD_RE.match(cut[-1]):
        cut.pop()
    head = " ".join(cut)
    return head.rstrip(",;:—-")


_HANGING_WORD_RE = re.compile(r"^(и|а|но|в|во|на|с|со|к|ко|о|об|обо|от|до|за|из|у|по|под|при|про|для|без|над|через|или|что|как|не|ни|же|бы|ли)$", re.I)


def _tail_of(text: str) -> str:
    text = text.strip()
    for sep in (":", " — ", " – "):
        if sep in text:
            tail = text.split(sep, 1)[1].strip()
            if tail:
                return tail[0].upper() + tail[1:]
    return text


def structure_fallback(index: int, spec, total: int) -> dict[str, Any]:
    bullets = _bullets(spec)
    texts = _texts(spec)
    items = []
    for bullet in bullets:
        label = ""
        match = re.match(r"^\s*((?:Q[1-4]\s*20\d\d)|(?:20\d\d(?:\s*г\.?)?)|(?:этап|шаг|фаза|спринт)\s*\d+|[А-ЯЁ][а-яё]+\s+20\d\d)\s*[:—–-]\s*(.+)$", bullet, re.I)
        if match:
            label, bullet = match.group(1).strip(), match.group(2).strip()
        head, tail = _head_of(bullet), _tail_of(bullet)
        if tail.lower().startswith(head.lower()) and len(tail) - len(head) < 25:
            tail = ""  # текст дословно повторяет заголовок — не дублируем
        items.append({"head": head, "text": tail, "label": label})
    metrics = []
    for bullet in bullets:
        match = _METRIC_RE.search(bullet)
        if match and re.search(r"\d", match.group(1)):
            value = re.sub(r"\s+", " ", match.group(1)).strip()
            label = (bullet[: match.start()] + bullet[match.end():]).strip(" ,.:—–-")
            metrics.append({"value": value, "label": _head_of(label, 5) if label else ""})
    body = " ".join(texts)[:400] if texts else ""
    if not bullets and texts and index not in {0, total - 1}:
        sentences = [x.strip() for x in re.split(r"(?<=[.!?])\s+", body) if len(x.strip()) > 12]
        if len(sentences) >= 2:
            items = [{"head": _head_of(x, 3), "text": x.rstrip("."), "label": ""} for x in sentences[:4]]
    subtitle = ""
    if index == 0:
        purpose = (spec.purpose or "").strip()
        # назначение вида «Обложка»/«Титульный слайд» — не подзаголовок
        generic = not purpose or re.match(r"^(обложка|титул|cover|title)", purpose, re.I)
        subtitle = body or ("" if generic else purpose)
    if index == total - 1:
        subtitle = body or (bullets[0] if bullets else "")
    return {"title": (spec.title or "").strip(), "subtitle": subtitle, "items": items, "metrics": metrics, "body": body, "note": ""}


# ---------------------------------------------------------------------------
# Выбор композиций
# ---------------------------------------------------------------------------

_COMPATIBLE: dict[str, list[str]] = {
    "cover": ["cover"],
    "closing": ["closing", "cta", "qa", "cover"],
    "section": ["section", "title_only", "statement", "cover"],
    "agenda": ["agenda", "cards", "timeline", "text"],
    "metrics": ["metrics", "cards", "columns", "text"],
    "timeline": ["timeline", "cards", "columns", "text"],
    "problems": ["problems", "cards", "columns", "text"],
    "cards": ["cards", "problems", "columns", "timeline", "text", "image_text"],
    "columns": ["columns", "cards", "text", "image_text"],
    "text": ["text", "image_text", "statement", "cards", "columns", "title_only"],
    "table": ["table", "text", "columns", "cards", "title_only"],
    "chart": ["chart", "image_text", "text", "columns", "cards", "title_only"],
}


def insisted_count(brief: str) -> int | None:
    match = _INSIST_RE.search(brief or "")
    if not match:
        return None
    value = match.group(1) or match.group(2)
    try:
        number = int(value)
    except (TypeError, ValueError):
        return None
    return number if 2 <= number <= 12 else None


def select_compositions(
    template: dict[str, Any],
    content_plan: ContentPlan,
    structured: list[dict[str, Any]],
    density_key: str,
    brief: str = "",
    avoid: dict[int, set[str]] | None = None,
    exclude: dict[int, set[str]] | None = None,
) -> list[dict[str, Any]]:
    """Для каждого слайда — композиция и число блоков. Детерминированно.

    Варианты различаются не только плотностью: B знает, что выбрал A, C — что
    выбрали A и B (`avoid`, считается детерминированно заново), и на том же
    слайде предпочитает другую композицию, если она есть.

    `exclude` — жёсткий запрет композиций на слайде (ручная «другая
    компоновка» пользователя): такие композиции берутся, только если больше
    нечего взять.
    """
    profile = DENSITY_PROFILES[density_key]
    exclude = exclude or {}
    if avoid is None:
        avoid = {}
        order = list(DENSITY_PROFILES)
        for previous_key in order[: order.index(density_key)]:
            previous = select_compositions(template, content_plan, structured, previous_key, brief, avoid=dict(avoid))
            skipped = dropped_slides(content_plan, previous_key)
            for idx, sel in enumerate(previous):
                if idx in skipped:
                    continue  # этого слайда в предыдущем варианте нет — избегать нечего
                avoid.setdefault(idx, set()).add(sel["variant"])
    comps = catalog_compositions(template)
    by_kind: dict[str, list[dict[str, Any]]] = {}
    for comp in comps:
        by_kind.setdefault(comp["kind"], []).append(comp)
    insist = insisted_count(brief)
    total = len(content_plan.slides)
    usage: dict[str, int] = {}
    history: list[str] = []
    selections: list[dict[str, Any]] = []
    sample_kinds = {c["kind"] for c in comps if c["source"] == "sample"}
    for index, spec in enumerate(content_plan.slides):
        intent = slide_intent(index, spec, total, content_plan)
        # У шаблона есть образец, совместимый с назначением слайда, — библиотечная
        # компоновка тогда лишь запасной вариант (см. _score)
        intent["_sample_fits"] = bool(sample_kinds & set(_COMPATIBLE.get(intent["kinds"][0], [intent["kinds"][0]])))
        intent["_sample_kinds"] = sample_kinds
        struct = structured[index] if index < len(structured) else structure_fallback(index, spec, total)
        wanted = intent["kinds"][0]
        n_items = _items_for(intent, struct)
        limit = profile["max_units"]
        if insist and n_items >= insist:
            limit = max(limit, insist)
        n_target = min(n_items, limit) if n_items else 0
        candidate_kinds = list(dict.fromkeys(intent["kinds"] + _COMPATIBLE.get(wanted, []) + ["text", "cards", "title_only"]))
        chosen = None
        scored: list[tuple[float, int, dict[str, Any]]] = []
        for rank, kind in enumerate(candidate_kinds):
            for comp in by_kind.get(kind) or []:
                score = _score(comp, kind, n_target, intent, profile, usage, history)
                if score is None:
                    continue
                # Приоритет вида композиции: каждый следующий вид дороже на 1.0
                if comp["id"] in (avoid.get(index) or set()):
                    # эту композицию на этом слайде уже взял другой вариант; если она —
                    # единственная нужного вида (одни «Ключевые цифры» на весь шаблон),
                    # повтор лучше, чем цифры мелким текстом в карточках
                    score += 0.4 if (kind == wanted and len(by_kind.get(kind) or []) == 1) else 1.2
                if comp["id"] in (exclude.get(index) or set()):
                    score += 50.0  # пользователь попросил другую компоновку
                # Обложка и финал: чужой вид (текстовый образец вместо титула) — только если титулов нет
                rank_weight = 3.0 if wanted in {"cover", "closing"} else 1.0
                scored.append((score + rank * rank_weight, rank, comp))
        if scored:
            scored.sort(key=lambda item: (item[0], item[1]))
            chosen = scored[0][2]
        if chosen is None:
            chosen = comps[0]
        units = 0
        if chosen.get("max_units"):
            # «Герой»-метрика вне блоков забирает первую цифру: блоков нужно меньше
            hero = chosen.get("hero_metrics") or 0 if intent.get("metrics") else 0
            wanted_units = max(n_target - hero, 1) if n_target else 0
            units = min(max(wanted_units, chosen.get("min_units") or 1), chosen["max_units"]) if n_target else 0
            if units and insist and n_items >= insist:
                units = min(n_items, chosen["max_units"])
            elif units == (chosen.get("units") or 0) + 1 and (chosen.get("units") or 0) >= 6 and chosen["source"] == "sample":
                # 7 тезисов на сетку из 6: достраивать один блок — сложно и некрасиво
                # (одинокий ряд); без настойчивого требования пользователя урезаем до 6
                grid = chosen.get("grid") or {}
                if (grid.get("cols") or 0) >= 3 and (grid.get("rows") or 0) >= 2:
                    units = chosen["units"]
            if chosen["kind"] in {"closing", "cta", "qa"}:
                units = 1  # один блок контактов
            if n_target == 0 and chosen["kind"] in {"cards", "problems", "columns", "timeline"}:
                units = 0
        usage[chosen["id"]] = usage.get(chosen["id"], 0) + 1
        history.append(chosen["id"])
        selections.append({"variant": chosen["id"], "kind": chosen["kind"], "units": units, "intent": intent, "comp": chosen})
    return selections


def _items_for(intent: dict[str, Any], struct: dict[str, Any]) -> int:
    if intent.get("agenda") is not None:
        return len(intent["agenda"])
    if intent.get("section_inside") and not struct.get("items"):
        return len(intent["section_inside"])
    if intent.get("metrics"):
        return len(struct.get("metrics") or []) or len(struct.get("items") or [])
    return len(struct.get("items") or [])


def _stable_hash(text: str) -> int:
    value = 0
    for ch in text:
        value = (value * 31 + ord(ch)) % 1_000_003
    return value


def _score(comp, kind, n_target, intent, profile, usage, history) -> float | None:
    """Меньше — лучше. None — композиция не подходит."""
    score = 0.0
    units = comp.get("units") or 0
    max_units = comp.get("max_units") or units
    min_units = comp.get("min_units") or 0
    if kind in {"cards", "problems", "timeline", "agenda", "columns"}:
        if n_target == 0:
            return None
        if comp.get("fixed_units") and n_target != units:
            return None  # ярлыки на декоре: ни убрать, ни добавить блок нельзя
        if n_target > max_units:
            score += 3.0 + (n_target - max_units)  # придётся урезать
        elif n_target < max(min_units, 1):
            score += 2.0
        elif n_target > units:
            score += 0.6 * (n_target - units)  # достраивать блоки
        else:
            score += 0.25 * (units - n_target)  # убирать блоки
        if comp["source"] == "layout" and units and units != n_target:
            # Плашки макета нельзя убрать со слайда: 3 тезиса в 6 карточках —
            # три пустых плашки. Лучше текстовый макет
            score += 3.0 + 1.0 * abs(units - n_target)
    if kind == "metrics":
        metrics = comp.get("metrics") or 0
        if n_target and metrics:
            # Потерять метрику (3 цифры в композицию на 2) — хуже, чем показать их карточками
            # Лишние слоты цифр убираются (removable_when_empty / библиотека строит
            # ровно n блоков) — штраф за них небольшой
            score += (0.2 if metrics > n_target else 0.5) * abs(metrics - n_target) + (2.5 * (n_target - metrics) if metrics < n_target else 0.0)
    if kind == "table" and not comp.get("table") and comp["source"] != "layout":
        return None
    if kind == "chart" and not comp.get("chart") and comp["source"] != "layout":
        return None
    # Слайду нужны таблица/график: образец без них их потеряет — только макет с телом
    # (старый путь внедрения объектов) или образец/библиотека с таким объектом
    primary = (intent.get("kinds") or [kind])[0]
    if primary in {"table", "chart"} and kind != primary:
        if comp["source"] == "layout" and (not comp.get("has_body") or kind in {"speaker", "agenda"}):
            return None
        if comp["source"] != "layout":
            return None
    if kind == "chart" and comp.get("charts"):
        score += 0.7 * abs(comp["charts"] - 1)
    if kind in {"text", "image_text", "columns"} and comp["source"] == "layout" and not comp.get("has_body"):
        return None
    # Плотность: заметки/подзаголовки/картинки — для насыщенного
    if comp.get("has_note") and not profile["note"]:
        score += 0.2
    if comp.get("image_slots"):
        score += -0.4 if profile["images"] else 0.3
    if comp["source"] == "layout" and kind not in {"section", "title_only"}:
        score += 0.8  # образцы и библиотека выразительнее «голых» макетов
    if comp["source"] == "library" and intent.get("_sample_fits"):
        # Библиотека — заполнение пробелов: образцы шаблона, если они подходят по
        # смыслу, важнее. Исключение — библиотечная композиция ровно нужного вида
        # без замены (график/таблица/цифры/разделитель), которого среди образцов
        # нет: она и добавлена ради этого
        gap_kind = kind == primary and primary in {"chart", "table", "metrics", "section"} and primary not in (intent.get("_sample_kinds") or set())
        if not gap_kind:
            score += 1.5
    # Разделитель с лидом/пунктами раздела содержательнее «голого» заголовка
    if primary == "section" and intent.get("section_inside"):
        if kind in {"section", "statement", "title_only"} and not (comp.get("has_body") or comp.get("has_subtitle")):
            score += 2.0  # разделитель без единого слова, кроме заголовка, выглядит пустым
        elif kind in {"section", "statement"}:
            score -= 0.3
    # Тезис без тела для слайда с текстом — потеря содержания
    if kind == "statement" and primary not in {"section", "statement"} and not comp.get("has_body") and not comp.get("has_subtitle"):
        score += 1.5
    # Цифры карточками — заголовки блоков мельче «геройских» значений
    if primary == "metrics" and kind == "cards":
        score += 0.3
    # «Только заголовок» для слайда с тезисами/текстом — всё содержание пропадёт;
    # допустимо лишь для разделителей/обложек или когда больше ничего нет
    if kind == "title_only" and primary not in {"section", "cover", "closing", "statement", "title_only"} and (n_target or intent.get("has_text")):
        score += 4.0
    # Обложка/финал с блоками-карточками, которые нечем заполнить, — пустые рамки
    if kind in {"cover", "closing"} and units and not n_target and not intent.get("metrics"):
        score += 1.0
    # Характер варианта: каждому — свои любимые и нелюбимые композиции
    score += (profile.get("prefer") or {}).get(kind, 0.0) + (profile.get("avoid") or {}).get(kind, 0.0)
    # Ротация: равные по смыслу композиции в разных вариантах выбираются по-разному
    rotation = profile.get("rotation") or 0
    score += 0.45 * ((_stable_hash(comp["id"]) + rotation) % 3) / 2
    # Разделители в колоде должны быть одинаковыми, а не «карточками» ради разнообразия
    if primary == "section":
        if kind not in {"section", "statement", "title_only"}:
            score += 2.5
        elif kind == "section":
            score -= 0.5 * usage.get(comp["id"], 0)  # повтор разделителя — норма
    # Ритм: не повторять одну композицию подряд, равномерно чередовать
    score += 0.5 * usage.get(comp["id"], 0)
    if history and history[-1] == comp["id"]:
        score += 1.0
    if len(history) >= 2 and history[-2] == comp["id"]:
        score += 0.4
    return score


# ---------------------------------------------------------------------------
# Заполнение слотов
# ---------------------------------------------------------------------------


def _today_ru() -> str:
    today = _dt.date.today()
    return f"{today.day} {_MONTHS[today.month - 1]} {today.year}"


def _contacts_from_brief(brief: str) -> str:
    parts: list[str] = []
    for regex in (r"[\w.+-]+@[\w-]+\.[\w.]+", r"https?://[^\s)\]>«»\"']+", r"\+?\d[\d\s()-]{9,}\d"):
        found = re.findall(regex, brief or "")
        if found:
            parts.append(found[0].strip())
    return " · ".join(dict.fromkeys(parts))


def fill_slots(
    variant: dict[str, Any],
    selection: dict[str, Any],
    struct: dict[str, Any],
    spec,
    index: int,
    total: int,
    density_key: str,
    brief: str,
    content_plan: ContentPlan,
    shorten,
) -> dict[str, str]:
    """Контент слота ← структурированный контент, по семантическим ролям."""
    profile = DENSITY_PROFILES[density_key]
    slots = variant.get("slots") or {}
    contract = {item["slot"]: item for item in (variant.get("generation_contract") or {}).get("text_slots") or []}
    comp = variant.get("composition") or {}
    kind = selection["kind"]
    intent = selection["intent"]
    units = selection.get("units") or 0
    content: dict[str, str] = {}

    section_free = bool(intent.get("section_inside")) and kind in {"section", "statement", "title_only", "text", "cover"}

    def put(name: str, text: str, is_title: bool = False) -> None:
        text = (text or "").strip()
        if not text:
            return
        cs = contract.get(name) or {}
        constraints = cs.get("constraints") or {}
        limit = constraints.get("maximum_characters") or constraints.get("recommended_characters")
        lines = constraints.get("maximum_lines") or constraints.get("recommended_lines")
        if is_title and limit and slots.get(name, {}).get("semantic_role") in {"slide_title", "presentation_title", "section_title", "closing_title"}:
            limit = int(limit * 3.0)  # заголовок не режем — рендер уменьшит кегль (валидатор допускает 3.2×)
            lines = None
        elif is_title and limit and slots.get(name, {}).get("semantic_role") == "unit_title":
            limit = int(limit * 1.8)  # заголовок блока: кегль уменьшится до места под ним (валидатор допускает 2×)
            lines = None
        role = slots.get(name, {}).get("semantic_role")
        if role in {"contact_text", "speaker_name", "url", "date"} and limit and " " not in text.split(" · ")[0]:
            content[name] = text  # e-mail/ссылку не режем — рендер уменьшит кегль
            return
        if role in {"metric_value", "unit_number"} and len(text) <= 24:
            content[name] = text  # «150 млн» → «150 мл» недопустимо: значение не режем, рендер уменьшит кегль
            return
        if section_free and role in {"body", "subtitle", "subheading"} and "\n" in text:
            # лид + пункты раздела: слот разделителя однострочный, рамка вырастет вниз
            # (валидатор допускает ≤ 6 строк и ≤ 420 знаков)
            content[name] = shorten(text, 420, 6, is_title=False)
            return
        content[name] = shorten(text, limit, lines if "\n" in text or is_title else None, is_title=is_title)

    title = struct.get("title") or spec.title
    if index == total - 1 and kind in {"closing", "qa", "cta"}:
        title = struct.get("title") or "Спасибо за внимание"
    subtitle = struct.get("subtitle") or ""
    contacts = _contacts_from_brief(brief)
    items = list(struct.get("items") or [])
    if intent.get("agenda") is not None:
        items = [{"head": t, "text": "", "label": ""} for t in intent["agenda"]]
    metrics = list(struct.get("metrics") or [])
    if intent.get("metrics") and not metrics:
        metrics = [{"value": _head_of(i["head"], 2), "label": i["text"] or i["head"]} for i in items]
    # Композиции с несколькими «геройскими» цифрами (2–3 metric_value вне блоков):
    # если модель выделила меньше метрик, чем слотов, добираем цифры из тезисов
    # плана («10 млн — рост от новых клиентов» → 10 млн / рост от новых клиентов»),
    # иначе обязательный слот останется пустым и вариант не пройдёт валидацию
    need_metrics = sum(1 for sl in slots.values() if sl.get("semantic_role") == "metric_value" and not sl.get("virtual"))
    if need_metrics > len(metrics):
        metrics = metrics + _metrics_from_bullets(_bullets(spec), items, exclude={m["value"] for m in metrics})
    metric_cards = bool(intent.get("metrics") and kind in {"cards", "problems", "columns", "timeline"} and metrics)
    if metric_cards:
        items = [{"head": m["value"], "text": m["label"], "label": ""} for m in metrics]
    # Контакты внутри блока «спикер» (две строки: имя / должность): только контакты,
    # по одному на строку; подзаголовок туда не кладём
    unit_contact_lines = [c.strip() for c in contacts.split(" · ") if c.strip()] if contacts else ([subtitle] if index == 0 and subtitle else [])

    # --- слоты блоков
    unit_slots: list[list[str]] = comp.get("unit_slots") or []
    section_inside = list(intent.get("section_inside") or [])
    if len(section_inside) < 2:
        section_inside = []  # один пункт — это заголовок следующего слайда, не повторяем его
    library_section = comp.get("library_key") == "section"
    if intent.get("section_inside") is not None and kind in {"section", "statement", "title_only", "text", "cover", "cards"}:
        lead = struct.get("subtitle") or struct.get("body") or ""
        if re.search(r"разделител|смыслов\w+ част|переход к|раздел о\b|раздел,? посвящ", lead.lower()):
            lead = ""  # служебная фраза модели («разделитель…») — не содержание
        lead = _first_sentences(lead, 2) if lead else ""
        inside_text = "\n".join(f"• {t}" for t in section_inside)
        if library_section:
            # библиотечный разделитель рисует лид и пункты сам: без маркеров, раздельно
            struct = dict(struct, body="\n".join(section_inside), subtitle=lead)
            subtitle = lead
        elif not section_inside:
            # лид один раз: в тело, если оно есть у композиции, иначе в подзаголовок
            has_body_slot = any(sl.get("semantic_role") == "body" and not sl.get("virtual") for sl in slots.values())
            struct = dict(struct, body=lead if has_body_slot else "", subtitle="" if has_body_slot else lead)
            subtitle = "" if has_body_slot else lead
        elif kind == "cards":
            # пункты раздела — в карточках; в теле/подзаголовке только лид
            struct = dict(struct, body=lead)
            subtitle = subtitle or lead
        else:
            # лид + пункты раздела всегда вместе (модель иногда кладёт лид в body)
            struct = dict(struct, body=f"{lead}\n{inside_text}" if lead else inside_text, subtitle=lead)
            subtitle = lead or section_inside[0]
        if kind in {"section", "statement", "cover", "cards"} and not items:
            items = [{"head": t, "text": "", "label": ""} for t in section_inside]
    hero_metrics = sum(1 for s in slots.values() if s.get("semantic_role") == "metric_value" and s.get("unit") is None and not s.get("virtual"))
    split_hero = bool(hero_metrics) and len(metrics) > hero_metrics
    unit_metrics = metrics[hero_metrics:] if split_hero else metrics
    hero_pool = metrics[:hero_metrics] if split_hero else metrics
    # «Геройские» цифры раздаём по положению слота (сверху вниз, слева направо),
    # а не по порядку имён в макете: главная метрика должна стоять первой
    hero_names = sorted(
        (n for n, sl in slots.items() if sl.get("semantic_role") == "metric_value" and sl.get("unit") is None and not sl.get("virtual")),
        key=_slot_reading_order(slots, [n for n, sl in slots.items() if sl.get("semantic_role") == "metric_value" and sl.get("unit") is None and not sl.get("virtual")]),
    )
    hero_by_name = {n: hero_pool[k] for k, n in enumerate(hero_names) if k < len(hero_pool)}
    for k, names in enumerate(unit_slots[:units]):
        item = items[k] if k < len(items) else None
        metric = unit_metrics[k] if k < len(unit_metrics) else None
        seen_roles: set[str] = set()
        for name in names:
            role = slots.get(name, {}).get("semantic_role")
            if role in {"unit_title", "unit_text", "unit_label", "metric_value", "metric_label", "agenda_item"}:
                if role in seen_roles:
                    continue  # второй текст той же роли в блоке — остаётся пустым (очистится)
                seen_roles.add(role)
            if role == "unit_title" and item:
                put(name, item["head"], is_title=True)
            elif role == "unit_text" and item:
                has_title = any(slots.get(n, {}).get("semantic_role") == "unit_title" for n in names)
                text = ""
                if has_title:
                    # под заголовком блока — только настоящее пояснение, не его копия;
                    # у метрики подпись обязательна в любой плотности («20%» само по себе ни о чём)
                    if (profile["unit_text"] or metric_cards) and item["text"]:
                        text = item["text"]
                elif profile["unit_text"] and item["text"]:
                    text = item["text"] if item["text"].lower().startswith(item["head"].lower()) else f"{item['head']}. {item['text']}"
                else:
                    text = item["head"]
                if text:
                    if profile["text_chars"]:
                        text = shorten(text, profile["text_chars"], None)
                    put(name, text)
            elif role == "unit_label" and item:
                put(name, item.get("label") or item["head"], is_title=True)
            elif role == "agenda_item" and item:
                put(name, item["head"], is_title=True)
            elif role == "metric_value" and metric:
                put(name, metric["value"])
            elif role == "metric_label" and metric:
                put(name, metric["label"] or (item["head"] if item else ""))
            elif role == "contact_text":
                put(name, _fit_contacts(unit_contact_lines.pop(0) if unit_contact_lines else "", contract.get(name)))
            # unit_number — автонумерация в рендере

    if library_section:
        kicker, bare_title = split_kicker(title)
        if kicker:
            title = bare_title
            content["label"] = kicker
    # --- одиночные слоты
    body_lines = [f"• {i['head']}" + (f" — {i['text']}" if profile["unit_text"] and i["text"] else "") for i in items]
    body_seen = 0
    for name, slot in slots.items():
        if slot.get("virtual") or slot.get("unit") is not None or name in content:
            continue
        role = slot.get("semantic_role")
        if role in {"slide_title", "presentation_title", "section_title", "closing_title"}:
            put(name, title, is_title=True)
        elif role in {"subtitle", "presentation_subtitle", "subheading"}:
            if index == 0 or kind in {"closing", "cta", "qa", "section"}:
                subtitle_seen = content.get("__subtitle_seen", 0)
                content["__subtitle_seen"] = subtitle_seen + 1
                if subtitle_seen == 0:
                    put(name, subtitle)
                elif subtitle_seen == 1 and contacts:
                    put(name, _fit_contacts(contacts, contract.get(name)))  # «Имя Фамилия / Должность» → контакты из брифа
                # третий и далее подзаголовок обложки («Должность») — пусто, плейсхолдер уберётся
            elif profile["note"] or not struct.get("body"):
                put(name, struct.get("note") or "")
        elif role == "body":
            body_index = body_seen
            body_seen += 1
            if intent.get("section_inside") and struct.get("body") and body_index == 0:
                put(name, struct["body"])  # лид раздела и его пункты — как есть
                continue
            if body_index >= 1:
                # несколько текстовых полей: раскладываем тезисы по одному, лишние — пустые
                if items and body_index < len(items) and units == 0:
                    it = items[body_index]
                    put(name, f"{it['head']}. {it['text']}" if it.get("text") and not it["text"].lower().startswith(it["head"].lower()) else (it.get("text") or it["head"]))
                continue
            if struct.get("body") and (intent.get("sentences") or kind in {"text", "statement", "image_text"} or not items):
                put(name, _first_sentences(struct["body"], profile["body_sentences"]))
            elif units == 0 and items and kind not in {"metrics"}:
                put(name, "\n".join(body_lines))
            elif struct.get("body"):
                put(name, _first_sentences(struct["body"], profile["body_sentences"]))
            elif items and units < len(items) and kind not in {"metrics", "agenda"}:
                put(name, "\n".join(body_lines[units:]))
        elif role == "column_body":
            pass  # раскладываем ниже
        elif role == "note":
            if profile["note"]:
                put(name, struct.get("note") or "")
        elif role == "metric_value":
            metric = hero_by_name.get(name) or _pop_first([m for m in hero_pool if m not in hero_by_name.values()])
            if metric:
                put(name, metric["value"])
                content[f"__metric_label_for_{name}"] = metric["label"]
        elif role == "metric_label":
            pass  # ниже — парой к значению
        elif role == "contact_text" or role == "speaker_name":
            # Несколько контактных строк («Имя Фамилия» / «Должность»): по одному контакту
            # из брифа на строку, без повторов; подзаголовок — только если у композиции
            # нет собственного слота подзаголовка и контактов нет
            contact_lines = [c.strip() for c in contacts.split(" · ") if c.strip()] if contacts else []
            seen = content.get("__contact_seen", 0)
            content["__contact_seen"] = seen + 1
            has_sub_slot = any(sl.get("semantic_role") in {"subtitle", "presentation_subtitle"} and not sl.get("virtual") for sl in slots.values())
            if seen < len(contact_lines):
                put(name, _fit_contacts(contact_lines[seen], contract.get(name)))
            elif seen == 0 and not contact_lines and subtitle and not has_sub_slot:
                put(name, _fit_contacts(subtitle, contract.get(name)))
        elif role == "date":
            put(name, _today_ru())
        elif role == "tag":
            if re.search(r"кейс", (spec.purpose or "") + " " + (spec.title or ""), re.I):
                put(name, "Кейс")
        elif role == "presentation_title":
            put(name, title, is_title=True)

    # metric_label парой к ближайшему metric_value (по порядку)
    value_names = [n for n, s in slots.items() if s.get("semantic_role") == "metric_value" and s.get("unit") is None and n in content]
    label_names = [n for n, s in slots.items() if s.get("semantic_role") == "metric_label" and s.get("unit") is None]
    value_names.sort(key=_slot_reading_order(slots, value_names))
    label_names.sort(key=_slot_reading_order(slots, label_names))
    for v_name, l_name in zip(value_names, label_names):
        label = content.pop(f"__metric_label_for_{v_name}", "")
        if label:
            put(l_name, label)
    for key in [k for k in content if k.startswith("__")]:
        content.pop(key)

    # колонки макетов парсера
    column_names = [n for n, s in slots.items() if s.get("semantic_role") == "column_body"]
    if column_names and items:
        buckets: list[list[str]] = [[] for _ in column_names]
        for k, line in enumerate(body_lines):
            buckets[k % len(column_names)].append(line)
        for name, bucket in zip(column_names, buckets):
            if bucket:
                put(name, "\n".join(bucket))
    return content


def _fit_contacts(text: str, contract_slot: dict[str, Any] | None) -> str:
    """Контакты «email · сайт · телефон»: оставляем столько частей, сколько помещается."""
    if not text or " · " not in text:
        return text
    limit = ((contract_slot or {}).get("constraints") or {}).get("maximum_characters")
    if not limit:
        return text
    parts = text.split(" · ")
    out: list[str] = []
    for part in parts:
        candidate = " · ".join(out + [part])
        if len(candidate) <= limit:
            out.append(part)
    return " · ".join(out) if out else parts[0]


def _first_sentences(text: str, count: int) -> str:
    sentences = re.split(r"(?<=[.!?])\s+", text.strip())
    return " ".join(sentences[:max(count, 1)]).strip()


def _pop_first(items: list[dict[str, Any]]) -> dict[str, Any] | None:
    return items.pop(0) if items else None


# ---------------------------------------------------------------------------
# План варианта
# ---------------------------------------------------------------------------


def build_variant_plan(
    template: dict[str, Any],
    content_plan: ContentPlan,
    structured: list[dict[str, Any]],
    density_key: str,
    brief: str,
    shorten,
    exclude: dict[int, set[str]] | None = None,
) -> dict[str, Any]:
    variants = {v["id"]: v for v in template.get("variants") or []}
    selections = select_compositions(template, content_plan, structured, density_key, brief, exclude=exclude)
    total = len(content_plan.slides)
    dropped = dropped_slides(content_plan, density_key)
    sections_total = sum(1 for i in section_map(content_plan) if i not in dropped)
    slides = []
    section_ordinal = 0
    for index, (spec, selection) in enumerate(zip(content_plan.slides, selections)):
        if index in dropped:
            continue
        variant = variants[selection["variant"]]
        struct = structured[index] if index < len(structured) else structure_fallback(index, spec, total)
        content = fill_slots(variant, selection, struct, spec, index, total, density_key, brief, content_plan, shorten)
        slide: dict[str, Any] = {
            "variant": selection["variant"],
            "content": content,
            "image_content": {},
            "composition_kind": selection["kind"],
            "units": selection.get("units") or 0,
            "source_index": index,
        }
        if selection["intent"].get("metrics") and selection["kind"] in {"cards", "problems", "columns", "timeline"}:
            slide["metric_units"] = True  # значения метрик в заголовках карточек — рендер укрупнит кегль
        if index in section_map(content_plan):
            section_ordinal += 1
            slide["section_ordinal"] = section_ordinal
            slide["sections_total"] = sections_total
        slides.append(slide)
    return {"slides": slides, "density": density_key}
