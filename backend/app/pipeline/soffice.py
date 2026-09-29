"""Единая точка запуска headless LibreOffice для превью и экспорта.

Почему отдельный модуль. На production-сервере (2 CPU / 2 ГБ RAM) экран
сравнения запрашивал превью трёх вариантов одновременно → три параллельных
процесса soffice делили один профиль пользователя и упирались в память; один
из них падал с «soffice завершился с кодом 1» без текста ошибки. Здесь:

* `asyncio.Lock` — конвертации выполняются строго по очереди;
* у каждого запуска свой `-env:UserInstallation` во временной папке, чтобы
  параллельные/зависшие инстансы не блокировали друг друга lock-файлом
  профиля;
* жёсткий таймаут — зависший soffice убивается, а не держит очередь вечно;
* один повтор при сбое (LibreOffice иногда падает на первом холодном старте).
"""
from __future__ import annotations

import asyncio
import logging
import shutil
import tempfile
import uuid
from pathlib import Path

logger = logging.getLogger(__name__)

_SOFFICE_LOCK = asyncio.Lock()
_DEFAULT_TIMEOUT_S = 240.0


class SofficeError(RuntimeError):
    pass


def soffice_binary() -> str | None:
    return shutil.which("soffice") or shutil.which("libreoffice")


async def convert(pptx_path: str | Path, output_dir: str | Path, fmt: str, timeout: float = _DEFAULT_TIMEOUT_S) -> Path:
    """.pptx → `output_dir/<stem>.<fmt>` (pdf, html, …). Бросает SofficeError."""
    binary = soffice_binary()
    if binary is None:
        raise SofficeError(
            "LibreOffice (soffice) не найден в PATH — установите LibreOffice, "
            "чтобы строить превью и экспортировать pdf/html (см. README)"
        )
    pptx_path = Path(pptx_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    result_path = output_dir / f"{pptx_path.stem}.{fmt.split(':')[0]}"

    last_error: str = ""
    async with _SOFFICE_LOCK:
        for attempt in (1, 2):
            profile_dir = Path(tempfile.gettempdir()) / f"lo_profile_{uuid.uuid4().hex}"
            try:
                process = await asyncio.create_subprocess_exec(
                    binary,
                    f"-env:UserInstallation=file://{profile_dir}",
                    "--headless",
                    "--norestore",
                    "--nologo",
                    "--convert-to",
                    fmt,
                    "--outdir",
                    str(output_dir),
                    str(pptx_path),
                    stdout=asyncio.subprocess.PIPE,
                    stderr=asyncio.subprocess.PIPE,
                )
                try:
                    stdout, stderr = await asyncio.wait_for(process.communicate(), timeout=timeout)
                except asyncio.TimeoutError:
                    process.kill()
                    await process.wait()
                    last_error = f"soffice не уложился в {int(timeout)} с"
                    logger.warning("%s (%s, попытка %d)", last_error, pptx_path.name, attempt)
                    continue
                if process.returncode != 0:
                    last_error = (
                        f"soffice завершился с кодом {process.returncode}: "
                        f"{stderr.decode(errors='ignore').strip() or stdout.decode(errors='ignore').strip() or 'без сообщения'}"
                    )
                    logger.warning("%s (%s, попытка %d)", last_error, pptx_path.name, attempt)
                    continue
                if not result_path.exists():
                    last_error = (
                        f"soffice отработал без ошибки, но файл {result_path.name} не появился: "
                        f"{stdout.decode(errors='ignore').strip()}"
                    )
                    logger.warning("%s (попытка %d)", last_error, attempt)
                    continue
                return result_path
            finally:
                shutil.rmtree(profile_dir, ignore_errors=True)
    raise SofficeError(last_error or "soffice: неизвестная ошибка")
