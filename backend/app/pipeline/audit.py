"""Пайплайн проверки готовой презентации.

Детерминированные проверки (audit_presentation) работают напрямую с
объектами python-pptx и не вызывают никаких моделей — только поэтому
аудит всегда отрабатывает быстро и одинаково для одного и того же .pptx.

audit_content_with_vlm — отдельная async-заглушка под недетерминированные
проверки смысла слайда (11 да/нет вопросов, см. docs/AUDIT.md).
"""
from __future__ import annotations

import re
import uuid
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor

from app.schemas.job import AuditCategory, AuditIssue

_MAX_BULLETS_PER_SLIDE = 6
_MAX_WORDS_PER_BULLET = 15
_MIN_CONTRAST_RATIO = 3.0

_PLACEHOLDER_TEXT_PATTERNS = (
    re.compile(r"lorem ipsum", re.IGNORECASE),
    re.compile(r"\bXXX\b"),
    re.compile(r"\bTODO\b", re.IGNORECASE),
    re.compile(r"вставьте текст", re.IGNORECASE),
)


def audit_presentation(pptx_path: str | Path) -> list[AuditIssue]:
    prs = Presentation(str(pptx_path))
    issues: list[AuditIssue] = []

    for slide_idx, slide in enumerate(prs.slides):
        slide_id = f"slide-{slide_idx + 1}"
        issues.extend(_check_bounds(slide, slide_id, prs.slide_width, prs.slide_height))
        issues.extend(_check_overlap(slide, slide_id))
        issues.extend(_check_density(slide, slide_id))
        issues.extend(_check_placeholder_text(slide, slide_id))
        issues.extend(_check_contrast(slide, slide_id))

    return issues


def _new_issue(
    slide_id: str,
    category: AuditCategory,
    description: str,
    *,
    deterministic: bool = True,
    auto_fixable: bool = False,
) -> AuditIssue:
    return AuditIssue(
        issue_id=str(uuid.uuid4()),
        slide_id=slide_id,
        category=category,
        deterministic=deterministic,
        description=description,
        auto_fixable=auto_fixable,
    )


def _check_bounds(slide, slide_id: str, slide_width: int, slide_height: int) -> list[AuditIssue]:
    issues = []
    for shape in slide.shapes:
        if None in (shape.left, shape.top, shape.width, shape.height):
            continue
        out_of_bounds = (
            shape.left < 0
            or shape.top < 0
            or shape.left + shape.width > slide_width
            or shape.top + shape.height > slide_height
        )
        if out_of_bounds:
            issues.append(
                _new_issue(
                    slide_id,
                    AuditCategory.LAYOUT,
                    f"Фигура «{shape.name}» выходит за границы слайда",
                )
            )
    return issues


def _check_overlap(slide, slide_id: str) -> list[AuditIssue]:
    issues = []
    shapes = [s for s in slide.shapes if None not in (s.left, s.top, s.width, s.height)]
    for i in range(len(shapes)):
        for j in range(i + 1, len(shapes)):
            if _rects_overlap(shapes[i], shapes[j]):
                issues.append(
                    _new_issue(
                        slide_id,
                        AuditCategory.LAYOUT,
                        f"Фигуры «{shapes[i].name}» и «{shapes[j].name}» перекрываются",
                    )
                )
    return issues


def _rects_overlap(a, b) -> bool:
    ax0, ay0, ax1, ay1 = a.left, a.top, a.left + a.width, a.top + a.height
    bx0, by0, bx1, by1 = b.left, b.top, b.left + b.width, b.top + b.height
    return ax0 < bx1 and bx0 < ax1 and ay0 < by1 and by0 < ay1


def _check_density(slide, slide_id: str) -> list[AuditIssue]:
    issues = []
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        paragraphs = [p for p in shape.text_frame.paragraphs if p.text.strip()]
        if len(paragraphs) > _MAX_BULLETS_PER_SLIDE:
            issues.append(
                _new_issue(
                    slide_id,
                    AuditCategory.DENSITY,
                    f"Больше {_MAX_BULLETS_PER_SLIDE} буллетов в блоке «{shape.name}» ({len(paragraphs)})",
                    auto_fixable=True,
                )
            )
        for p in paragraphs:
            word_count = len(p.text.split())
            if word_count > _MAX_WORDS_PER_BULLET:
                issues.append(
                    _new_issue(
                        slide_id,
                        AuditCategory.DENSITY,
                        f"Буллет длиннее {_MAX_WORDS_PER_BULLET} слов ({word_count}): «{p.text[:60]}»",
                        auto_fixable=True,
                    )
                )
    return issues


def _check_placeholder_text(slide, slide_id: str) -> list[AuditIssue]:
    issues = []
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        text = shape.text_frame.text
        for pattern in _PLACEHOLDER_TEXT_PATTERNS:
            if pattern.search(text):
                issues.append(
                    _new_issue(
                        slide_id,
                        AuditCategory.INTEGRITY,
                        f"Найден текст-заглушка в «{shape.name}»: «{text[:60]}»",
                    )
                )
                break
    return issues


def _relative_luminance(rgb: RGBColor) -> float:
    def channel(value: int) -> float:
        c = value / 255
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4

    r, g, b = channel(rgb[0]), channel(rgb[1]), channel(rgb[2])
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def _contrast_ratio(rgb_a: RGBColor, rgb_b: RGBColor) -> float:
    l1, l2 = _relative_luminance(rgb_a), _relative_luminance(rgb_b)
    lighter, darker = max(l1, l2), min(l1, l2)
    return (lighter + 0.05) / (darker + 0.05)


def _background_rgb(slide) -> RGBColor | None:
    try:
        fill = slide.background.fill
        if fill.type is None:
            return None
        return fill.fore_color.rgb
    except (AttributeError, TypeError, KeyError):
        return None


def _check_contrast(slide, slide_id: str) -> list[AuditIssue]:
    """Проверяет контраст ЯВНО заданного цвета текста относительно фона.

    Ограничение: считается, только если и текст, и фон имеют явный srgbClr
    (не унаследованный из темы) — иначе достоверно определить итоговый цвет
    без полного разрешения темы нельзя, и слайд молча пропускается, чтобы
    не плодить ложные срабатывания.
    """
    background_rgb = _background_rgb(slide)
    if background_rgb is None:
        return []

    issues = []
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                try:
                    if run.font.color.type is None:
                        continue
                    rgb = run.font.color.rgb
                except (AttributeError, TypeError, KeyError):
                    continue
                if rgb is None or not run.text.strip():
                    continue
                ratio = _contrast_ratio(rgb, background_rgb)
                if ratio < _MIN_CONTRAST_RATIO:
                    issues.append(
                        _new_issue(
                            slide_id,
                            AuditCategory.LAYOUT,
                            f"Низкий контраст текста «{run.text[:30]}»: {ratio:.1f}:1 (минимум {_MIN_CONTRAST_RATIO}:1)",
                        )
                    )
    return issues


async def audit_content_with_vlm(pptx_path: str | Path) -> list[AuditIssue]:
    """TODO: недетерминированные проверки смысла слайда через VLM.

    Должна отрендерить каждый слайд в изображение и задать VLM 11 да/нет
    вопросов про смысловое качество (см. docs/AUDIT.md), возвращая
    AuditIssue(deterministic=False, ...) на каждый ответ "нет".
    """
    raise NotImplementedError("VLM-аудит контента ещё не реализован")
