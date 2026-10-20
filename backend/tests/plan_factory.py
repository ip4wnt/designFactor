def bullets(*items: str) -> dict:
    return {"type": "bullets", "bullets": list(items)}


def text(value: str) -> dict:
    return {"type": "text", "text": value}


def slide(slide_id: str, title: str, blocks: list[dict], purpose: str = "суть") -> dict:
    return {"slide_id": slide_id, "purpose": purpose, "title": title, "content_blocks": blocks}


THREE_BULLETS = bullets(
    "Выручка выросла за счёт новых клиентов",
    "Средний чек увеличился после обновления тарифов",
    "Отток остался на уровне прошлого года",
)


def good_plan(content_slides: int = 3) -> dict:
    slides = [slide("s0", "Продажи 2026", [text("Итоги года для руководства")], "обложка")]
    for i in range(1, content_slides + 1):
        slides.append(slide(f"s{i}", f"Вывод номер {i}: рост на {i * 10}%", [THREE_BULLETS]))
    slides.append(slide("sx", "Спасибо за внимание", [text("Вопросы и контакты")], "финал"))
    return {"brief": "b", "purpose": "p", "slides": slides}


def to_plan(data: dict):
    """Разбор через JSON, как у реального ответа LLM: StrictModel не принимает строки вместо enum из dict."""
    import json

    from app.schemas.content_plan import ContentPlan

    return ContentPlan.model_validate_json(json.dumps(data, ensure_ascii=False))
