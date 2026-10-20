"""Самовосстановление сборки дизайна.

Сборка не должна падать целиком из-за одного слайда или одного варианта:

- сломанный слайд → ищем, какой именно слайд не рендерится (пробная сборка
  по одному слайду), и подставляем другую композицию для него; если ни одна
  композиция не собирается — слайд выпадает из этого варианта с пометкой;
- сломанный вариант → остальные варианты все равно отдаются пользователю,
  а причина записывается в `job.variant_errors`.

Все замены фиксируются человеко-читаемыми заметками (`job.variant_notes`),
чтобы интерфейс честно показывал, что сервис исправил сам.
"""

from __future__ import annotations

import copy
import logging
import tempfile
from pathlib import Path
from typing import Any, Callable

from app.pipeline import template_engine

logger = logging.getLogger(__name__)

# Сколько альтернативных композиций пробуем для одного сломанного слайда
MAX_ALTERNATIVES = 6

# (индекс слайда в плане варианта, индекс слайда плана содержания, уже
# опробованные композиции) → новый слайд плана или None, если альтернатив нет
RebuildSlide = Callable[[int, int, set[str]], dict[str, Any] | None]


def friendly_error(exc: BaseException) -> str:
    """Короткое сообщение для пользователя: первая строка без технических дампов."""
    text = str(exc).strip() or exc.__class__.__name__
    first = text.split("\n", 1)[0].strip().rstrip(":").strip()
    if len(first) > 240:
        first = first[:237].rstrip() + "…"
    return first


def _probe(template: dict[str, Any], slide: dict[str, Any], density: str | None, template_path: str | Path) -> BaseException | None:
    """Собирает один слайд в отдельный файл. None — слайд собирается, иначе ошибка."""
    mini = {"slides": [copy.deepcopy(slide)]}
    if density:
        mini["density"] = density
    with tempfile.TemporaryDirectory(prefix="df-probe-") as tmp:
        try:
            template_engine.render_variant(template, mini, template_path, Path(tmp) / "probe.pptx", None, None, "probe")
        except Exception as exc:  # noqa: BLE001 — любая ошибка рендера = слайд сломан
            return exc
    return None


def _broken_slides(template: dict[str, Any], plan: dict[str, Any], template_path: str | Path) -> dict[int, BaseException]:
    broken: dict[int, BaseException] = {}
    for index, slide in enumerate(plan.get("slides") or []):
        exc = _probe(template, slide, plan.get("density"), template_path)
        if exc is not None:
            broken[index] = exc
    return broken


def _slide_label(slide: dict[str, Any]) -> str:
    kind = str(slide.get("composition_kind") or "").strip()
    return f"{slide.get('variant')}" + (f" ({kind})" if kind else "")


def render_with_repair(
    template: dict[str, Any],
    plan: dict[str, Any],
    template_path: str | Path,
    output_path: str | Path,
    manifest,
    content_plan,
    variant_key: str,
    rebuild_slide: RebuildSlide | None,
) -> list[str]:
    """Рендер варианта с автоматическим исправлением сломанных слайдов.

    Меняет `plan` на месте (заменяет или убирает слайды). Возвращает список
    заметок о том, что было исправлено. Если собрать не удалось даже после
    исправлений — пробрасывает исходную ошибку.
    """
    output_path = Path(output_path)
    try:
        template_engine.render_variant(template, plan, template_path, output_path, manifest, content_plan, variant_key)
        return []
    except Exception as first:  # noqa: BLE001 — дальше разбираемся, какой слайд виноват
        first_error = first
    logger.warning("Сборка %s упала (%s) — ищем проблемные слайды", variant_key, friendly_error(first_error))

    broken = _broken_slides(template, plan, template_path)
    if not broken:
        # Ошибка не воспроизводится на отдельных слайдах (например, в графике
        # или при склейке колоды) — чинить нечего, пусть решает вызывающий код
        raise first_error

    notes: list[str] = []
    slides = plan["slides"]
    for index, exc in broken.items():
        original = slides[index]
        tried = {str(original.get("variant"))}
        logger.warning("Слайд %d варианта %s не собирается: %s", index + 1, variant_key, friendly_error(exc))
        replacement = None
        if rebuild_slide is not None:
            source_index = int(original.get("source_index", index))
            for _ in range(MAX_ALTERNATIVES):
                candidate = rebuild_slide(index, source_index, set(tried))
                if candidate is None or str(candidate.get("variant")) in tried:
                    break
                tried.add(str(candidate.get("variant")))
                if _probe(template, candidate, plan.get("density"), template_path) is None:
                    replacement = candidate
                    break
        if replacement is not None:
            slides[index] = replacement
            notes.append(
                f"Слайд {index + 1}: композиция «{_slide_label(original)}» не поместилась в шаблон, "
                f"подобрана другая — «{_slide_label(replacement)}»"
            )
        else:
            slides[index] = None
            notes.append(f"Слайд {index + 1} убран из этого варианта: ни одна композиция шаблона для него не собралась")

    plan["slides"] = [s for s in slides if s is not None]
    if not plan["slides"]:
        raise first_error
    template_engine.render_variant(template, plan, template_path, output_path, manifest, content_plan, variant_key)
    logger.info("Вариант %s собран после исправлений: %s", variant_key, "; ".join(notes))
    return notes
