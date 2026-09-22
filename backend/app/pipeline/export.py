"""Экспорт готового .pptx в .pdf/.html через headless LibreOffice.

.pptx уже полностью готов на выходе Компоновщика (app/pipeline/assembly.py) —
здесь только конвертация форматов, без какой-либо генерации контента.
"""
from __future__ import annotations

import asyncio
import shutil
from pathlib import Path

_SUPPORTED_FORMATS = {"pptx", "pdf", "html"}


class ExportError(RuntimeError):
    pass


async def export_presentation(pptx_path: str | Path, output_dir: str | Path, fmt: str) -> Path:
    if fmt not in _SUPPORTED_FORMATS:
        raise ValueError(f"Неподдерживаемый формат экспорта: {fmt}")

    pptx_path = Path(pptx_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    if fmt == "pptx":
        destination = output_dir / pptx_path.name
        if destination.resolve() != pptx_path.resolve():
            shutil.copyfile(pptx_path, destination)
        return destination

    return await _convert_with_libreoffice(pptx_path, output_dir, fmt)


async def _convert_with_libreoffice(pptx_path: Path, output_dir: Path, fmt: str) -> Path:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice is None:
        raise ExportError(
            "LibreOffice (soffice) не найден в PATH — установите LibreOffice, "
            "чтобы экспортировать в pdf/html (см. README, «Ограничения текущей версии»)"
        )

    process = await asyncio.create_subprocess_exec(
        soffice,
        "--headless",
        "--convert-to",
        fmt,
        "--outdir",
        str(output_dir),
        str(pptx_path),
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE,
    )
    stdout, stderr = await process.communicate()

    if process.returncode != 0:
        raise ExportError(
            f"soffice завершился с кодом {process.returncode}: {stderr.decode(errors='ignore')}"
        )

    result_path = output_dir / f"{pptx_path.stem}.{fmt}"
    if not result_path.exists():
        raise ExportError(
            f"soffice отработал без ошибки, но файл {result_path} не появился: "
            f"{stdout.decode(errors='ignore')}"
        )
    return result_path
