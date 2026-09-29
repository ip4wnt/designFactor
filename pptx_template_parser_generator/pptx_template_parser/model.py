"""OpenAI-compatible JSON client used by the presentation planner."""

import json
import os
import re
import urllib.error
import urllib.request


class ModelError(RuntimeError):
    pass


class OpenAICompatibleModel:
    def __init__(self, base_url=None, model=None, api_key=None, timeout=180):
        self.base_url = (base_url or os.getenv(
            "PPTX_LLM_BASE_URL", "http://localhost:11434/v1"
        )).rstrip("/")
        self.model = model or os.getenv("PPTX_LLM_MODEL", "qwen3:8b")
        self.api_key = api_key if api_key is not None else os.getenv("PPTX_LLM_API_KEY")
        self.timeout = timeout

    def json(self, system, user, temperature=0.2):
        messages = [
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ]
        is_ollama = "11434" in self.base_url and self.base_url.endswith("/v1")
        if is_ollama:
            url = f"{self.base_url[:-3]}/api/chat"
            request_body = {
                "model": self.model,
                "messages": messages,
                "stream": False,
                "format": "json",
                "think": False,
                "options": {"temperature": temperature, "num_predict": 8192},
            }
        else:
            url = f"{self.base_url}/chat/completions"
            request_body = {
            "model": self.model,
            "temperature": temperature,
            "max_tokens": 8192,
            "response_format": {"type": "json_object"},
            "messages": messages,
            }
        payload = json.dumps(request_body, ensure_ascii=False).encode("utf-8")
        headers = {"Content-Type": "application/json"}
        if self.api_key:
            headers["Authorization"] = f"Bearer {self.api_key}"
        request = urllib.request.Request(
            url,
            data=payload,
            headers=headers,
            method="POST",
        )
        try:
            with urllib.request.urlopen(request, timeout=self.timeout) as response:
                data = json.loads(response.read().decode("utf-8"))
        except (OSError, urllib.error.HTTPError, json.JSONDecodeError) as exc:
            raise ModelError(f"Не удалось вызвать модель {self.model}: {exc}") from exc
        try:
            content = (
                data["message"]["content"] if is_ollama
                else data["choices"][0]["message"]["content"]
            )
        except (KeyError, IndexError, TypeError) as exc:
            raise ModelError("Ответ модели не содержит choices[0].message.content") from exc
        # Some Ollama builds expose Qwen3-VL's final structured answer in the
        # `thinking` field even when thinking is explicitly disabled.
        if not content and is_ollama:
            content = data.get("message", {}).get("thinking")
        if not content:
            raise ModelError(
                "Модель вернула пустой content. Отключите thinking/reasoning "
                "или используйте текстовую instruct-модель."
            )
        return parse_json_response(content)


def parse_json_response(content):
    if isinstance(content, dict):
        return content
    text = str(content).strip()
    fenced = re.fullmatch(r"```(?:json)?\s*(.*?)\s*```", text, re.S | re.I)
    if fenced:
        text = fenced.group(1)
    try:
        value = json.loads(text)
    except json.JSONDecodeError as exc:
        raise ModelError(f"Модель вернула невалидный JSON: {exc}") from exc
    if not isinstance(value, dict):
        raise ModelError("Корневое значение ответа модели должно быть объектом.")
    return value
