"""Рисует рамки вокруг проблемных мест аудита поверх PNG-превью слайда.

Не трогает исходный PNG из кэша app/pipeline/preview.py (тот же кэш нужен и
экрану сравнения вариантов без рамок) — сохраняет отдельный файл рядом.
Issue без bbox (проблемы уровня всего слайда: пустой слайд, дубликат,
заполненность, VLM-вопросы о смысле) не рисуются — им нечего обводить, и
это ожидаемо, а не баг.
"""
from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw

from app.schemas.job import AuditIssue

_CATEGORY_COLORS: dict[str, tuple[int, int, int]] = {
    "layout": (239, 68, 68),  # красный
    "template": (245, 158, 11),  # оранжевый
    "density": (168, 85, 247),  # фиолетовый
    "integrity": (220, 38, 38),  # тёмно-красный
    "content": (59, 130, 246),  # синий (VLM, обычно без bbox)
}
_DEFAULT_COLOR = (239, 68, 68)
_STROKE_WIDTH = 4


def draw_issue_overlay(png_path: str | Path, issues: list[AuditIssue], output_path: str | Path) -> Path:
    """Копирует png_path в output_path, обводя рамкой каждый issue.bbox."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with Image.open(png_path) as source:
        canvas = source.convert("RGB")

    width, height = canvas.size
    draw = ImageDraw.Draw(canvas)

    for issue in issues:
        if issue.bbox is None:
            continue
        x0 = issue.bbox.x * width
        y0 = issue.bbox.y * height
        x1 = x0 + issue.bbox.width * width
        y1 = y0 + issue.bbox.height * height
        color = _CATEGORY_COLORS.get(issue.category.value, _DEFAULT_COLOR)
        draw.rectangle([x0, y0, x1, y1], outline=color, width=_STROKE_WIDTH)

    canvas.save(output_path, "PNG")
    return output_path
