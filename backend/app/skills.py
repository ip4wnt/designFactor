"""Загрузка скиллов — промптов и конфигов агентов, которые лежат в репозитории
отдельными yaml-файлами (`backend/skills/*.yaml`), а не зашиты в код.

Каждый скилл — словарь с полями `name`, `version`, `model_name` и текстами
промптов (`system_prompt`, `rewrite_hint`, `questions`, …). Файлы читаются один
раз и кэшируются на процесс; путь к каталогу — `SKILLS_DIR` из `.env`.
"""

from __future__ import annotations

from functools import lru_cache
from pathlib import Path
from typing import Any

import yaml

from app.config import get_settings


class SkillNotFound(RuntimeError):
    """Файл скилла отсутствует или не разбирается — сервис не должен стартовать
    с «половиной» промптов, поэтому ошибка не глотается."""


def skills_dir() -> Path:
    return Path(get_settings().SKILLS_DIR)


@lru_cache(maxsize=None)
def load_skill(name: str) -> dict[str, Any]:
    """Возвращает содержимое `skills/<name>.yaml` как словарь."""
    path = skills_dir() / f"{name}.yaml"
    if not path.exists():
        raise SkillNotFound(f"Не найден файл скилла {path}")
    with path.open("r", encoding="utf-8") as handle:
        data = yaml.safe_load(handle) or {}
    if not isinstance(data, dict):
        raise SkillNotFound(f"Скилл {path} должен быть yaml-объектом")
    data.setdefault("name", name)
    return data


def skill_text(name: str, key: str) -> str:
    """Текстовое поле скилла без хвостовых переводов строки."""
    value = load_skill(name).get(key)
    if not isinstance(value, str) or not value.strip():
        raise SkillNotFound(f"В скилле {name} нет текстового поля {key!r}")
    return value.strip()


def list_skills() -> list[dict[str, Any]]:
    """Список скиллов для документации/диагностики: имя, версия, файл."""
    result = []
    for path in sorted(skills_dir().glob("*.yaml")):
        data = load_skill(path.stem)
        result.append({"name": data.get("name", path.stem), "version": data.get("version"), "file": str(path)})
    return result
