"""Проверка ContentPlan по правилам промпта skills/content_planning.yaml.

Схема ContentPlan гарантирует только форму ответа; сами правила («заголовок до
60 знаков», «не больше 6 буллетов», «разделители только от 9 слайдов») модель
иногда нарушает. Здесь они проверяются детерминированно, а content_agent при
нарушениях просит модель исправить именно перечисленные места.

Пороги плотности совпадают с аудитом (app/pipeline/audit.py): иначе план,
прошедший проверку, всё равно получил бы замечания аудита после вёрстки.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Literal

from app.schemas.content_plan import ContentBlockType, ContentPlan, SlideSpec

MAX_BULLETS_PER_SLIDE = 6
MAX_WORDS_PER_BULLET = 15
MAX_TABLE_ROWS = 7
MAX_TABLE_COLS = 5
MAX_CHART_SERIES = 5

MAX_COVER_TITLE = 30
MAX_TITLE = 60
# Вёрстка режет заголовок по границе слова до вместимости слота (template_engine._shorten):
# превышение на пару знаков почти безболезненно, а на 40% — калечит смысл. Поэтому
# до +25% от лимита это предупреждение, и только дальше — ошибка, требующая правки
# (каждая правка — лишний вызов LLM на 15–30 секунд).
TITLE_SOFT_OVERSHOOT = 1.25
MIN_SLIDES_FOR_DIVIDERS = 9

# Слово в начале заголовка, по которому он почти наверняка называет тему, а не
# делает вывод. Список намеренно короткий: ложные срабатывания хуже пропуска.
_TOPIC_TITLE_FIRST_WORDS = {
    "результаты", "итоги", "обзор", "введение", "заключение", "анализ", "описание",
    "общая", "общие", "основные", "ключевые", "цели", "задачи", "структура", "повестка", "агенда",
}

_JUNK_PATTERNS = (
    re.compile(r"lorem ipsum", re.IGNORECASE),
    re.compile(r"\bXXX\b"),
    re.compile(r"\bTODO\b", re.IGNORECASE),
    re.compile(r"вставьте текст", re.IGNORECASE),
    re.compile(r"как (ии|ассистент|языковая модель)", re.IGNORECASE),
    re.compile(r"\bне могу (выполнить|помочь)", re.IGNORECASE),
)

Severity = Literal["error", "warning"]


@dataclass(frozen=True)
class PlanIssue:
    """slide_id=None — замечание ко всей презентации, а не к одному слайду."""

    slide_id: str | None
    code: str
    message: str
    severity: Severity = "error"


def validate_plan(
    plan: ContentPlan,
    slide_count: int | None = None,
    slide_count_max: int | None = None,
) -> list[PlanIssue]:
    """Все замечания к плану. `slide_count`/`slide_count_max` — границы числа слайдов
    из запроса пользователя; без них число слайдов не проверяется."""
    issues: list[PlanIssue] = []
    slides = plan.slides
    total = len(slides)

    issues.extend(_check_slide_count(total, slide_count, slide_count_max))

    divider_count = 0
    for index, slide in enumerate(slides):
        role = _slide_role(slide, index, total)
        if role == "divider":
            divider_count += 1
        issues.extend(_check_slide(slide, role, total))

    if total >= MIN_SLIDES_FOR_DIVIDERS and not 2 <= divider_count <= 3:
        issues.append(
            PlanIssue(None, "divider_count", f"При {total} слайдах нужно 2–3 слайда-разделителя, сейчас {divider_count}", "warning")
        )

    issues.extend(_check_duplicates(slides))
    return issues


def format_issues(issues: list[PlanIssue], plan: ContentPlan) -> str:
    """Список замечаний для повторного запроса к модели."""
    titles = {slide.slide_id: slide.title for slide in plan.slides}
    lines = []
    for issue in issues:
        if issue.slide_id is None:
            lines.append(f"- вся презентация: {issue.message}")
        else:
            lines.append(f"- слайд {issue.slide_id} («{titles.get(issue.slide_id, '')[:50]}»): {issue.message}")
    return "\n".join(lines)


def _check_slide_count(total: int, slide_count: int | None, slide_count_max: int | None) -> list[PlanIssue]:
    if slide_count is None:
        return []
    if slide_count_max is not None and slide_count_max > slide_count:
        low, high = slide_count, slide_count_max
    else:
        # Точное число модель нередко берёт с промахом в один слайд — это не повод перегенерировать план.
        low, high = max(slide_count - 1, 1), slide_count + 1
    if low <= total <= high:
        return []
    wanted = f"от {slide_count} до {slide_count_max}" if high == slide_count_max else str(slide_count)
    return [PlanIssue(None, "slide_count", f"Слайдов {total}, а нужно {wanted}")]


def _slide_role(slide: SlideSpec, index: int, total: int) -> str:
    purpose = slide.purpose.strip().casefold()
    if index == 0:
        return "cover"
    if index == total - 1:
        return "final"
    if purpose.startswith("раздел"):
        return "divider"
    if purpose.startswith("цифры"):
        return "facts"
    return "content"


def _check_slide(slide: SlideSpec, role: str, total: int) -> list[PlanIssue]:
    sid = slide.slide_id
    issues: list[PlanIssue] = []
    title = slide.title.strip()
    blocks = slide.content_blocks

    limit = MAX_COVER_TITLE if role == "cover" else MAX_TITLE
    if not title:
        issues.append(PlanIssue(sid, "title_empty", "Пустой заголовок"))
    elif len(title) > limit * TITLE_SOFT_OVERSHOOT:
        issues.append(PlanIssue(sid, "title_too_long", f"Заголовок {len(title)} знаков, максимум {limit}"))
    elif len(title) > limit:
        issues.append(
            PlanIssue(sid, "title_slightly_long", f"Заголовок {len(title)} знаков при лимите {limit} — сократи, иначе вёрстка обрежет", "warning")
        )

    if role in ("content", "facts") and _looks_like_topic(title):
        issues.append(PlanIssue(sid, "title_is_topic", "Заголовок называет тему, а не делает вывод", "warning"))

    for text in _all_texts(slide):
        if any(pattern.search(text) for pattern in _JUNK_PATTERNS):
            issues.append(PlanIssue(sid, "junk_text", f"Служебный текст или заглушка: «{text[:60]}»"))
            break

    if role in ("cover", "final", "divider"):
        if len(blocks) != 1 or blocks[0].type != ContentBlockType.TEXT:
            issues.append(PlanIssue(sid, "single_text_block", "Здесь нужен ровно один короткий блок text", "warning"))
        if role == "divider" and total < MIN_SLIDES_FOR_DIVIDERS:
            issues.append(PlanIssue(sid, "divider_in_short", f"В презентации из {total} слайдов разделители не нужны", "warning"))
        return issues

    if not blocks:
        issues.append(PlanIssue(sid, "no_content", "Слайд без содержимого"))
        return issues

    issues.extend(_check_bullets(slide, role))
    issues.extend(_check_data_blocks(slide))
    return issues


def _check_bullets(slide: SlideSpec, role: str) -> list[PlanIssue]:
    sid = slide.slide_id
    issues: list[PlanIssue] = []
    bullets = [b for block in slide.content_blocks if block.bullets for b in block.bullets]

    if len(bullets) > MAX_BULLETS_PER_SLIDE:
        issues.append(PlanIssue(sid, "too_many_bullets", f"{len(bullets)} буллетов, максимум {MAX_BULLETS_PER_SLIDE}"))

    long_ones = [(len(b.split()), b) for b in bullets if len(b.split()) > MAX_WORDS_PER_BULLET]
    if long_ones:
        example = "; ".join(f"{words} слов: «{text[:40]}…»" for words, text in long_ones[:2])
        issues.append(PlanIssue(sid, "bullet_too_long", f"Буллетов длиннее {MAX_WORDS_PER_BULLET} слов: {len(long_ones)} ({example})"))

    types = {block.type for block in slide.content_blocks}
    only_bullets = types == {ContentBlockType.BULLETS}
    if role == "content" and only_bullets and len(bullets) < 3:
        issues.append(PlanIssue(sid, "thin_slide", f"Всего {len(bullets)} буллетов — слайд слишком тонкий (нужно 3–5)", "warning"))

    if role == "facts" and not 2 <= len(bullets) <= 3:
        issues.append(PlanIssue(sid, "facts_count", f"На слайде цифр {len(bullets)} буллетов, нужно 2–3", "warning"))
    if role == "facts" and any(not re.match(r"\s*\d", b) for b in bullets):
        issues.append(PlanIssue(sid, "facts_start", "Каждый буллет слайда цифр должен начинаться с числа", "warning"))

    has_data = bool(types & {ContentBlockType.CHART, ContentBlockType.TABLE})
    if has_data and not bullets and ContentBlockType.TEXT not in types:
        issues.append(PlanIssue(sid, "data_without_takeaways", "К графику или таблице нужны 2–3 буллета с выводами", "warning"))
    return issues


def _check_data_blocks(slide: SlideSpec) -> list[PlanIssue]:
    sid = slide.slide_id
    issues: list[PlanIssue] = []
    for block in slide.content_blocks:
        if block.table is not None:
            table = block.table
            cols = len(table.headers)
            if any(len(row) != cols for row in table.rows):
                issues.append(PlanIssue(sid, "table_ragged", f"В таблице есть строки, где число ячеек не равно числу колонок ({cols})"))
            if len(table.rows) > MAX_TABLE_ROWS:
                issues.append(PlanIssue(sid, "table_rows", f"В таблице {len(table.rows)} строк, максимум {MAX_TABLE_ROWS}"))
            if cols > MAX_TABLE_COLS:
                issues.append(PlanIssue(sid, "table_cols", f"В таблице {cols} колонок, максимум {MAX_TABLE_COLS}"))
        if block.chart is not None:
            chart = block.chart
            if len(chart.series) > MAX_CHART_SERIES:
                issues.append(PlanIssue(sid, "chart_series", f"В диаграмме {len(chart.series)} серий, максимум {MAX_CHART_SERIES}"))
            if len(chart.categories) < 2:
                issues.append(PlanIssue(sid, "chart_categories", "В диаграмме меньше двух категорий — сравнивать нечего", "warning"))
    return issues


def _check_duplicates(slides: list[SlideSpec]) -> list[PlanIssue]:
    issues: list[PlanIssue] = []
    seen_titles: dict[str, str] = {}
    seen_bullets: dict[str, str] = {}
    for slide in slides:
        key = _normalize(slide.title)
        if key and key in seen_titles:
            issues.append(PlanIssue(slide.slide_id, "duplicate_title", f"Заголовок повторяет слайд {seen_titles[key]}", "warning"))
        elif key:
            seen_titles[key] = slide.slide_id
        for block in slide.content_blocks:
            for bullet in block.bullets or []:
                bkey = _normalize(bullet)
                if bkey in seen_bullets and seen_bullets[bkey] != slide.slide_id:
                    issues.append(PlanIssue(slide.slide_id, "duplicate_bullet", f"Буллет повторяет слайд {seen_bullets[bkey]}: «{bullet[:40]}»", "warning"))
                elif bkey:
                    seen_bullets.setdefault(bkey, slide.slide_id)
    return issues


def _looks_like_topic(title: str) -> bool:
    words = title.split()
    if not words or len(words) > 4 or re.search(r"\d", title):
        return False
    return words[0].casefold().strip(":,.") in _TOPIC_TITLE_FIRST_WORDS


def _all_texts(slide: SlideSpec) -> list[str]:
    texts = [slide.title]
    for block in slide.content_blocks:
        texts.extend(block.bullets or [])
        if block.text:
            texts.append(block.text)
    return texts


def _normalize(text: str) -> str:
    return " ".join(text.casefold().split())
