"""One cached analytical run shared by all plan variants, with a hard process deadline."""
from __future__ import annotations

import asyncio
import hashlib
import json
import os
from pathlib import Path
import sys
from time import perf_counter
from weakref import WeakValueDictionary

from app.schemas.source_tables import TableAnalysisReport
from app.skills import load_skill

_locks: WeakValueDictionary[str, asyncio.Lock] = WeakValueDictionary()


def analysis_request(job, settings) -> dict:
    tables = [t.model_dump() for t in job.source_tables if t.data is not None and not t.parent_table_ids]
    instructions = load_skill("table_analysis")["system_prompt"]
    request = {"brief": job.brief, "idea": job.idea, "audience": job.audience, "goal": job.goal,
               "must_have": job.must_have, "materials": job.materials_text[:9000],
               "tables": tables, "instructions": instructions,
               "timeout": settings.TABLE_ANALYSIS_TIMEOUT_S, "max_turns": settings.TABLE_ANALYSIS_MAX_TURNS,
               "max_tool_calls": settings.TABLE_ANALYSIS_MAX_TOOL_CALLS,
               "max_results": settings.TABLE_ANALYSIS_MAX_RESULTS}
    code_dir = Path(__file__).resolve().parents[1] / "analysis"
    versions = [hashlib.sha256((code_dir / name).read_bytes()).hexdigest() for name in ("worker.py", "dataframe_tools.py")]
    identity = {"request": request, "model": settings.LLM_MODEL_NAME,
                "base_url": settings.LLM_BASE_URL, "reasoning_effort": "low", "code": versions}
    request["key"] = hashlib.sha256(json.dumps(identity, ensure_ascii=False, sort_keys=True).encode()).hexdigest()
    return request


async def _run_worker(request, directory, settings) -> TableAnalysisReport:
    directory.mkdir(parents=True, exist_ok=True)
    request_path, progress = directory / "request.json", directory / "progress.json"
    progress.unlink(missing_ok=True)
    request_path.write_text(json.dumps(request, ensure_ascii=False), encoding="utf-8")
    environment = dict(os.environ)
    # Credentials are passed only in the process environment, never request/cache files.
    environment.update(LLM_BASE_URL=settings.LLM_BASE_URL, LLM_MODEL_NAME=settings.LLM_MODEL_NAME,
                       LLM_API_KEY=settings.LLM_API_KEY or "", SKILLS_DIR=str(settings.SKILLS_DIR.resolve()),
                       STORAGE_DIR=str(settings.STORAGE_DIR.resolve()), OPENAI_AGENTS_DISABLE_TRACING="1")
    started = perf_counter()
    process = None
    timed_out = False
    with (directory / "worker.log").open("wb") as log:
        try:
            process = await asyncio.create_subprocess_exec(sys.executable, "-m", "app.analysis.worker",
                str(request_path.resolve()), str(progress.resolve()), env=environment,
                cwd=str(Path(__file__).resolve().parents[2]), stdout=log, stderr=log)
            await asyncio.wait_for(process.wait(), timeout=request["timeout"])
        except TimeoutError:
            timed_out = True
        finally:
            if process is not None and process.returncode is None:
                process.kill()
                await process.wait()
    try:
        report = TableAnalysisReport.model_validate_json(progress.read_text(encoding="utf-8"))
    except (FileNotFoundError, ValueError):
        report = TableAnalysisReport(key=request["key"], status="failed", notes=["Анализ не вернул сохранённых результатов."])
    if timed_out:
        report.status = "partial" if report.artifacts else "timeout"
        report.notes.append("Достигнут лимит времени анализа; используются только завершённые расчёты.")
    report.elapsed_ms = round((perf_counter() - started) * 1000)
    return report


async def prepare_table_analysis(job, settings) -> TableAnalysisReport:
    if not settings.TABLE_ANALYSIS_ENABLED:
        report = TableAnalysisReport(status="disabled")
    elif not any(t.data is not None and not t.parent_table_ids for t in job.source_tables):
        report = TableAnalysisReport(status="skipped")
    else:
        request = analysis_request(job, settings)
        key = request["key"]
        directory = settings.STORAGE_DIR.resolve() / "table_analysis" / key
        cache = directory / "report.json"
        async with _locks.setdefault(key, asyncio.Lock()):
            report = None
            if cache.exists():
                try:
                    report = TableAnalysisReport.model_validate_json(cache.read_text(encoding="utf-8"))
                    # A transient timeout/failure must be retried on a new request.
                    # Keep the report on disk for diagnostics, not as a permanent answer.
                    if report.key != key or report.status != "completed":
                        report = None
                    else:
                        report.cache_hit = True
                except ValueError:
                    pass  # An interrupted cache write should never break planning.
            if report is None:
                job.busy_note = "Анализируем таблицы для всех вариантов плана"
                try:
                    report = await _run_worker(request, directory, settings)
                except Exception as error:
                    report = TableAnalysisReport(key=key, status="failed", notes=["Анализ недоступен: " + type(error).__name__])
                directory.mkdir(parents=True, exist_ok=True)
                temporary = cache.with_suffix(".tmp")
                temporary.write_text(report.model_dump_json(indent=2), encoding="utf-8")
                temporary.replace(cache)
    job.table_analysis = report
    job.source_tables = [t for t in job.source_tables if not t.parent_table_ids] + report.artifacts
    job.busy_note = ""
    return report


def analysis_prompt(report: TableAnalysisReport | None) -> str:
    if report is None or report.status in {"skipped", "disabled"}:
        return ""
    payload = {"status": report.status, "limitations": report.notes,
               "results": [{"table_id": t.table_id, "summary": t.summary, "source_table_ids": t.parent_table_ids,
                            "operations": t.operations, "data": t.data.model_dump(), "notes": t.limitation,
                            "recommended_presentation": t.recommended_presentation, "presentation_reason": t.presentation_reason,
                            "chart_type": t.chart.chart_type.value if t.chart else None}
                           for t in report.artifacts]}
    return "\n\nОбщий анализ таблиц (значения рассчитаны инструментами):\n" + json.dumps(payload, ensure_ascii=False)
