"""Конфигурация приложения из .env (pydantic-settings)."""
from __future__ import annotations

from functools import lru_cache
from pathlib import Path

from pydantic_settings import BaseSettings, SettingsConfigDict
from pydantic import Field


class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file=".env", env_file_encoding="utf-8", extra="ignore")

    LLM_BASE_URL: str = "http://localhost:11434/v1"
    LLM_MODEL_NAME: str = "gpt-4o-mini"
    # None -> заголовок Authorization не отправляется (локальные Ollama/vLLM
    # без авторизации). Для облачного OpenAI-совместимого API впиши ключ сюда.
    LLM_API_KEY: str | None = None
    # Общий дедлайн одного вызова LLM на этапе планирования (ContentPlan),
    # секунды, включая повторы. Провайдеры на длинных бриф-планах иногда
    # отвечают дольше минуты — при «Превышен общий дедлайн LLM» поднимай здесь.
    LLM_PLANNING_TIMEOUT_S: float = 120.0
    # Сколько раз просить модель исправить план, нарушающий правила промпта
    # (app/pipeline/plan_validator.py). Каждая правка — ещё один вызов LLM
    # (не дольше PLAN_REPAIR_TIMEOUT_S); 0 отключает проверку-правку.
    PLAN_REPAIR_ROUNDS: int = 1
    # Сколько секунд ждать ответ на запрос правки; не уложилась — остаётся исходный план.
    PLAN_REPAIR_TIMEOUT_S: float = 60.0
    # Дедлайн LLM-правок по результатам аудита (audit_fix), секунды.
    LLM_AUDIT_FIX_TIMEOUT_S: float = 45.0
    TABLE_SUMMARY_TIMEOUT_S: float = 20.0
    TABLE_ANALYSIS_ENABLED: bool = True
    TABLE_ANALYSIS_TIMEOUT_S: float = Field(default=150.0, gt=0, le=300)
    TABLE_ANALYSIS_MAX_TURNS: int = Field(default=14, ge=1, le=20)
    TABLE_ANALYSIS_MAX_TOOL_CALLS: int = Field(default=20, ge=1, le=40)
    TABLE_ANALYSIS_MAX_RESULTS: int = Field(default=3, ge=1, le=10)
    # Сколько тяжёлых стадий (разбор шаблона + сборка трёх вариантов, пересборка,
    # аудит с рендером) выполняется одновременно. Остальные задачи ждут в
    # статусе queued. Ориентир — ~3 ГБ памяти на задачу: на VM с 2 ГБ без
    # лимита три параллельных генерации с LibreOffice довели машину до
    # исчерпания памяти и зависания до аппаратной перезагрузки.
    JOB_CONCURRENCY: int = 3

    VLM_BASE_URL: str = "http://localhost:11434/v1"
    VLM_MODEL_NAME: str = "gpt-4o-mini"
    VLM_API_KEY: str | None = None
    # Выключатель VLM-аудита содержимого (Приложение 1, "Валидация контента").
    # По умолчанию включён, но оркестратор всё равно деградирует мягко (см.
    # orchestrator.py) — если эндпоинт недоступен, пайплайн не падает, просто
    # в audit_issues[variant] попадает один issue с описанием сбоя, а не 11.
    VLM_AUDIT_ENABLED: bool = True

    STORAGE_DIR: Path = Path("storage")
    SKILLS_DIR: Path = Path("skills")

    @property
    def templates_dir(self) -> Path:
        return self.STORAGE_DIR / "templates"

    @property
    def outputs_dir(self) -> Path:
        return self.STORAGE_DIR / "outputs"


@lru_cache
def get_settings() -> Settings:
    settings = Settings()
    settings.templates_dir.mkdir(parents=True, exist_ok=True)
    settings.outputs_dir.mkdir(parents=True, exist_ok=True)
    return settings
