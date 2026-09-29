"""Асинхронный клиент: текст/фото -> текст или ответ по Pydantic-схеме.

Настройки передаются явно. Клиент поддерживает async with; общий клиент
приложения нужно закрыть через aclose() при остановке сервера.
"""
from __future__ import annotations

import asyncio
import base64
import json
import logging
import math
from dataclasses import dataclass
from io import BytesIO
from json import JSONDecodeError
from pathlib import Path
from time import perf_counter
from typing import Any, Sequence, TypeVar

from openai import AsyncOpenAI, omit
from PIL import Image
from pydantic import BaseModel

T = TypeVar("T", bound=BaseModel)
logger = logging.getLogger(__name__)


class LLMError(RuntimeError):
    """Отказ модели, незавершённый или неожиданный ответ API."""


def _strip_markdown_fence(raw: str) -> str:
    """Некоторые модели игнорируют `response_format: json_schema` и всё равно оборачивают
    ответ в ограду ```json ... ``` — снимаем её перед валидацией Pydantic-моделью,
    чтобы не падать пайплайн из-за чисто косметического отличия ответа модели."""
    text = raw.strip()
    if text.startswith("```"):
        text = text[3:]
        if text.startswith("json"):
            text = text[4:]
        if text.endswith("```"):
            text = text[:-3]
        text = text.strip()
    return text


class LLMTimeoutError(LLMError, TimeoutError):
    """Истёк общий дедлайн вызова, включая повторы SDK."""


@dataclass(frozen=True)
class ImageInput:
    data_url: str

    @classmethod
    def from_bytes(cls, data: bytes) -> ImageInput:
        """Проверяет PNG/JPEG/WEBP и кодирует без изменения изображения.

        Лимиты размера и количества загрузок задаются обработчиком API.
        """
        with Image.open(BytesIO(data)) as image:
            mime = {
                "PNG": "image/png",
                "JPEG": "image/jpeg",
                "WEBP": "image/webp",
            }.get(image.format)
            if mime is None:
                raise ValueError("Поддерживаются PNG, JPEG и WEBP")
            image.verify()

        encoded = base64.b64encode(data).decode("ascii")
        return cls(data_url=f"data:{mime};base64,{encoded}")

    @classmethod
    def from_path(cls, path: str | Path) -> ImageInput:
        return cls.from_bytes(Path(path).read_bytes())


class LLMClient:
    """Обёртка над Chat Completions SDK с проверкой завершения и Pydantic.

    Повторы выполняет SDK, включая повторы сетевых ошибок и таймаутов.
    ``max_retries=0`` отключает их. ``timeout`` ограничивает сетевые операции,
    ``total_timeout`` — весь вызов API вместе с повторами и ожиданием между ними.
    Общий дедлайн всей генерации презентации задаётся вызывающим кодом.
    """

    def __init__(
        self,
        *,
        base_url: str,
        model_name: str,
        api_key: str | None,
        timeout: float = 180.0,
        max_retries: int = 2,
        total_timeout: float = 90.0,
    ) -> None:
        if max_retries < 0:
            raise ValueError("max_retries должен быть неотрицательным")
        for name, value in (("timeout", timeout), ("total_timeout", total_timeout)):
            if not math.isfinite(value) or value <= 0:
                raise ValueError(f"{name} должен быть положительным конечным числом")
        self._model = model_name
        self._total_timeout = total_timeout
        self._request_headers = {
            "Authorization": f"Bearer {api_key}" if api_key else omit,
        }
        self._client = AsyncOpenAI(
            base_url=base_url,
            api_key=api_key or "unused",
            timeout=timeout,
            max_retries=max_retries,
        )

    async def __aenter__(self) -> LLMClient:
        return self

    async def __aexit__(self, exc_type, exc, tb) -> None:
        await self.aclose()

    async def aclose(self) -> None:
        await self._client.close()

    async def chat(
        self,
        *,
        system_prompt: str,
        user_input: str | BaseModel,
        images: Sequence[ImageInput] = (),
        max_completion_tokens: int = 8192,
        temperature: float = 0.1,
        reasoning_effort: str | None = "none",
        extra_body: dict[str, Any] | None = None,
    ) -> str:
        """Возвращает текст без response_format и проверки структуры ответа.

        ``reasoning_effort=None`` не отправляет настройку reasoning;
        ``"none"`` просит отключить его, ``"low"``/``"medium"``/``"high"``
        задают глубину. Допустимые значения зависят от модели и провайдера.

        ``extra_body=None`` не добавляет полей и не меняет provider policy.
        Поля extra_body добавляются SDK на верхний уровень JSON. Например,
        чтобы явно выбрать провайдера AllTokens (имя взять из /providers)::

            extra_body={
                "metadata": {
                    "provider_policy": {
                        "only": ["ИМЯ_ПРОВАЙДЕРА"],
                        "allow_fallbacks": False,
                    }
                }
            }

        Через extra_body нельзя переопределять model, messages, temperature,
        max_completion_tokens, max_tokens, response_format, reasoning_effort,
        stream или n: ими управляет клиент. Ошибки HTTP/сети передаются как
        исключения SDK openai; незавершённый/пустой ответ или отказ — LLMError.
        """
        return await self._chat(
            system_prompt=system_prompt,
            user_input=user_input,
            images=images,
            max_completion_tokens=max_completion_tokens,
            temperature=temperature,
            reasoning_effort=reasoning_effort,
            extra_body=extra_body,
        )

    async def chat_structured(
        self,
        *,
        system_prompt: str,
        user_input: str | BaseModel,
        output_model: type[T],
        schema_in_prompt: bool = True,
        images: Sequence[ImageInput] = (),
        max_completion_tokens: int = 8192,
        temperature: float = 0.1,
        reasoning_effort: str | None = "none",
        extra_body: dict[str, Any] | None = None,
    ) -> T:
        """Отправляет запрос со strict JSON Schema и возвращает output_model.

        Схема генерируется непосредственно Pydantic, без дополнительного
        преобразования. Её поддержку проверяет провайдер. Ошибки Pydantic,
        HTTP и сети (исключения SDK openai) передаются вызывающему коду;
        отказ/неполный ответ — LLMError.

        По умолчанию схема также включается в system_prompt: не все совместимые
        серверы одинаково передают модели структуру из response_format. Явное
        описание помогает ей учитывать назначение полей. schema_in_prompt=False
        отключает дублирование для провайдеров, где оно не требуется.

        reasoning_effort="none" просит отключить reasoning. extra_body=None
        оставляет маршрутизацию провайдеру; пример provider_policy — в chat().
        extra_body не может подменить схему response_format.
        """
        schema = output_model.model_json_schema(mode="serialization")
        if schema_in_prompt:
            system_prompt += (
                "\n\nВерни только JSON по следующей схеме. "
                "Учитывай назначения полей и не дублируй содержание.\nJSON Schema:\n"
                + json.dumps(schema, ensure_ascii=False, separators=(",", ":"))
            )
        raw = await self._chat(
            system_prompt=system_prompt,
            user_input=user_input,
            images=images,
            max_completion_tokens=max_completion_tokens,
            temperature=temperature,
            reasoning_effort=reasoning_effort,
            extra_body=extra_body,
            response_format={
                "type": "json_schema",
                "json_schema": {
                    "name": output_model.__name__,
                    "strict": True,
                    "schema": schema,
                },
            },
        )
        return output_model.model_validate_json(_strip_markdown_fence(raw), strict=True)

    async def _chat(
        self,
        *,
        system_prompt: str,
        user_input: str | BaseModel,
        images: Sequence[ImageInput],
        max_completion_tokens: int,
        temperature: float,
        reasoning_effort: str | None,
        extra_body: dict[str, Any] | None,
        response_format: dict[str, Any] | None = None,
    ) -> str:
        if max_completion_tokens < 1:
            raise ValueError("max_completion_tokens должен быть положительным")
        reserved_fields = {
            "model", "messages", "temperature", "max_completion_tokens",
            "max_tokens", "response_format", "reasoning_effort", "stream", "n",
        }
        if extra_body is not None:
            conflicts = reserved_fields.intersection(extra_body)
            if conflicts:
                raise ValueError(
                    "extra_body не может переопределять поля: "
                    + ", ".join(sorted(conflicts))
                )

        if isinstance(user_input, BaseModel):
            # Проверяем также изменения списков после создания входной модели.
            validated = type(user_input).model_validate_json(
                user_input.model_dump_json(), strict=True,
            )
            text = validated.model_dump_json()
        elif isinstance(user_input, str):
            text = user_input
        else:
            raise TypeError("user_input должен быть строкой или Pydantic-моделью")

        content: list[dict[str, Any]] = [{"type": "text", "text": text}]
        content.extend(
            {"type": "image_url", "image_url": {"url": image.data_url}}
            for image in images
        )

        payload: dict[str, Any] = {
            "model": self._model,
            "temperature": temperature,
            "max_completion_tokens": max_completion_tokens,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": content if images else text},
            ],
        }
        if response_format is not None:
            payload["response_format"] = response_format
        if reasoning_effort is not None:
            payload["reasoning_effort"] = reasoning_effort

        started = perf_counter()
        try:
            async with asyncio.timeout(self._total_timeout):
                completion = await self._client.chat.completions.create(
                    **payload, extra_body=extra_body, extra_headers=self._request_headers,
                )
        except TimeoutError as exc:
            logger.warning("LLM deadline model=%s elapsed=%.2fs", self._model, perf_counter() - started)
            raise LLMTimeoutError(
                f"Превышен общий дедлайн LLM: {self._total_timeout:g} с (включая повторы)"
            ) from exc
        except JSONDecodeError as exc:
            raise LLMError("API вернул невалидный JSON") from exc
        except Exception as exc:
            logger.warning(
                "LLM failed model=%s elapsed=%.2fs error=%s",
                self._model, perf_counter() - started, type(exc).__name__,
            )
            raise
        usage = getattr(completion, "usage", None)
        details = getattr(usage, "completion_tokens_details", None)
        logger.info(
            "LLM completed model=%s provider=%s id=%s elapsed=%.2fs "
            "prompt_tokens=%s completion_tokens=%s reasoning_tokens=%s",
            self._model, (getattr(completion, "model_extra", None) or {}).get("provider"),
            getattr(completion, "id", None), perf_counter() - started,
            getattr(usage, "prompt_tokens", None), getattr(usage, "completion_tokens", None),
            getattr(details, "reasoning_tokens", None),
        )
        try:
            choice = completion.choices[0]
            message = choice.message
            refusal = message.refusal
            finish_reason = choice.finish_reason
            raw = message.content
        except (AttributeError, IndexError, TypeError) as exc:
            raise LLMError("Неожиданный формат ответа API") from exc
        if refusal:
            raise LLMError("Модель отказалась выполнять запрос")

        if finish_reason != "stop":
            raise LLMError(f"Ответ не завершён: finish_reason={finish_reason}")

        if not isinstance(raw, str) or not raw.strip():
            raise LLMError("Модель вернула пустой ответ")
        return raw
