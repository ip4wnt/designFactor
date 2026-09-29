"""Рендер .pptx в PNG-превью по слайдам — для экрана сравнения вариантов.

Используется до того, как job доходит до статуса DONE: на этапе
AWAITING_VARIANT_CHOICE у нас уже есть все три собранных .pptx
(job.variant_paths), и нужно показать пользователю реальные картинки
слайдов, а не текстовую заглушку "Предпросмотр {label}".

Конвейер: .pptx -> .pdf (headless LibreOffice через общий сериализованный
запуск app/pipeline/soffice.py — три варианта рендерятся по очереди, а не
параллельно) -> PNG на слайд (pdf2image/poppler). Результат кэшируется на диске по mtime исходного
.pptx, чтобы повторный опрос экрана сравнения не пересчитывал рендер заново.
"""
from __future__ import annotations

import asyncio
import shutil
from pathlib import Path

from pdf2image import convert_from_path

from app.pipeline import soffice

# 80 dpi достаточно для карточек сравнения (~1070 px по ширине слайда) и в
# ~2 раза дешевле по памяти/времени, чем 110 dpi, на слабом сервере.
_DPI = 80


class PreviewError(RuntimeError):
    pass


async def render_variant_previews(pptx_path: str | Path, cache_dir: str | Path) -> list[Path]:
    """Возвращает пути к PNG-файлам (по одному на слайд) для указанного .pptx.

    Кэш валиден, пока mtime .pptx не изменился (после fix/rerun_assembly
    вариант пересобирается в тот же путь — тогда кэш нужно перегенерировать).
    """
    pptx_path = Path(pptx_path)
    cache_dir = Path(cache_dir)
    marker = cache_dir / ".source_mtime"
    expected_mtime = str(pptx_path.stat().st_mtime)

    if cache_dir.exists() and marker.exists() and marker.read_text().strip() == expected_mtime:
        existing = sorted(cache_dir.glob("slide_*.png"))
        if existing:
            return existing

    if cache_dir.exists():
        shutil.rmtree(cache_dir)
    cache_dir.mkdir(parents=True, exist_ok=True)

    try:
        pdf_path = await soffice.convert(pptx_path, cache_dir, "pdf")
    except soffice.SofficeError as exc:
        raise PreviewError(str(exc)) from exc

    images = await asyncio.to_thread(convert_from_path, str(pdf_path), dpi=_DPI)
    paths: list[Path] = []
    for idx, image in enumerate(images, start=1):
        out_path = cache_dir / f"slide_{idx:02d}.png"
        image.save(out_path, "PNG")
        paths.append(out_path)

    pdf_path.unlink(missing_ok=True)
    marker.write_text(expected_mtime)
    return paths
