"""Тонкий httpx-клиент к OpenAI-совместимому LLM endpoint.

Никакой бизнес-логики (промптов, парсинга под ContentPlan) здесь нет —
это отдельная зона ответственности app/pipeline/content_agent.py.
"""
from __future__ import annotations

import json
from typing import Any

import httpx

from app.config import get_settings


class LLMClient:
    def __init__(
        self,
        base_url: str | None = None,
        model_name: str | None = None,
        api_key: str | None = None,
    ) -> None:
        settings = get_settings()
        self._base_url = (base_url or settings.LLM_BASE_URL).rstrip("/")
        self._model_name = model_name or settings.LLM_MODEL_NAME
        # api_key может быть не передан явно (None) — тогда берём из настроек;
        # итоговое отсутствие ключа (локальный Ollama/vLLM) не ошибка.
        self._api_key = api_key if api_key is not None else settings.LLM_API_KEY

    async def chat_json(self, system_prompt: str, user_prompt: str, *, temperature: float = 0.3) -> dict[str, Any]:
        """Вызывает /chat/completions и парсит content ответа как JSON-объект."""
        payload = {
            "model": self._model_name,
            "temperature": temperature,
            "response_format": {"type": "json_object"},
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
        }
        headers = {"Authorization": f"Bearer {self._api_key}"} if self._api_key else {}

        async with httpx.AsyncClient(timeout=120.0) as client:
            response = await client.post(f"{self._base_url}/chat/completions", json=payload, headers=headers)
            response.raise_for_status()
            data = response.json()

        content = data["choices"][0]["message"]["content"]
        return json.loads(content)
