"""Конфигурация приложения из .env (pydantic-settings)."""
from __future__ import annotations

from functools import lru_cache
from pathlib import Path

from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file=".env", env_file_encoding="utf-8", extra="ignore")

    LLM_BASE_URL: str = "http://localhost:11434/v1"
    LLM_MODEL_NAME: str = "gpt-4o-mini"
    # None -> заголовок Authorization не отправляется (локальные Ollama/vLLM
    # без авторизации). Для облачного OpenAI-совместимого API впиши ключ сюда.
    LLM_API_KEY: str | None = None

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
