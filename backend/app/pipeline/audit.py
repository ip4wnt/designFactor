"""Пайплайн проверки готовой презентации.

Детерминированные проверки (audit_presentation) работают напрямую с
объектами python-pptx (+ опционально DesignManifest шаблона) и не вызывают
никаких моделей — только поэтому аудит всегда отрабатывает быстро и
одинаково для одного и того же .pptx. Полный список проверок и их область
покрытия задокументированы в docs/AUDIT.md (требование ТЗ, раздел 4) —
чек-лист происходит из Приложения 1 официального ТЗ ("Критерии качества и
аудита"), сокращённого/расширенного с обоснованием там же.

audit_content_with_vlm — недетерминированные проверки смысла слайда через
VLM (11 да/нет вопросов из Приложения 1, "Валидация контента"). Рендерит
каждый слайд в PNG (LibreOffice -> pdf -> pdftoppm) и просит VLM ответить
на все 11 вопросов сразу одним JSON-объектом на слайд — так один слайд стоит
один VLM-вызов, а не 11.
"""
from __future__ import annotations

import asyncio
import base64
import logging
import re
import shutil
import uuid
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Emu

from app.config import get_settings
from app.llm_client import LLMClient
from app.pipeline.export import ExportError, export_presentation
from app.schemas.design_manifest import DesignManifest
from app.schemas.job import AuditCategory, AuditIssue

logger = logging.getLogger(__name__)

# --- Плотность (Приложение 1, "Плотность") -----------------------------
_MAX_BULLETS_PER_SLIDE = 6
_MAX_WORDS_PER_BULLET = 15
_MAX_TABLE_ROWS = 7
_MAX_TABLE_COLS = 5
_MAX_CHART_SERIES = 5
_MIN_FILL_RATIO = 0.25
# 0.75 из черновика ТЗ оказался слишком агрессивным на реальных шаблонах: полноширинные
# шапки/подложки-фоны занимают большую часть площади bbox по дизайну, не значит
# перегруженности контентом (плотность буллетов/слов проверяется отдельно в _check_density).
# Поднят до 0.92, чтобы ловить реально забитые слайды (плотно заполненный
# текст+таблица+картинка на одном слайде), а не обычный дизайн шаблона.
_MAX_FILL_RATIO = 0.92

# --- Шаблон (Приложение 1, "Шаблон") ------------------------------------
_MIN_CONTRAST_RATIO = 4.5  # ТЗ требует именно 4.5:1, не WCAG-минимум 3.0
_FONT_SIZE_TOLERANCE_PT = 0.5
_MAX_DISTINCT_FONTS = 2

# --- Верстка (Приложение 1, "Вёрстка") ----------------------------------
_EDGE_MARGIN_RATIO = 0.02  # "контент заходит в поля у краёв" — 2% от размера слайда
_CHARS_PER_LINE_AT_10PT = 18.0  # эвристика: символов на 1 см строки при кегле 10pt

_PLACEHOLDER_TEXT_PATTERNS = (
    re.compile(r"lorem ipsum", re.IGNORECASE),
    re.compile(r"\bXXX\b"),
    re.compile(r"\bTODO\b", re.IGNORECASE),
    re.compile(r"вставьте текст", re.IGNORECASE),
    re.compile(r"как (ии|ассистент|языковая модель)", re.IGNORECASE),  # служебный мусор из промпта
)


def audit_presentation(pptx_path: str | Path, manifest: DesignManifest | None = None) -> list[AuditIssue]:
    """Все детерминированные проверки. `manifest` включает проверки категории
    TEMPLATE (шрифт/кегль/цвет из шаблона) — без него они пропускаются, а не
    падают, чтобы вызов оставался обратно совместим."""
    path = Path(pptx_path)
    if not path.exists():
        return [
            _new_issue(
                "file",
                AuditCategory.INTEGRITY,
                f"Файл {path.name} не найден на диске",
            )
        ]

    try:
        prs = Presentation(str(path))
    except Exception as exc:  # noqa: BLE001 — "файл не открывается" из Приложения 1
        return [
            _new_issue(
                "file",
                AuditCategory.INTEGRITY,
                f"Файл {path.name} не открывается: {exc}",
            )
        ]

    issues: list[AuditIssue] = []
    slide_texts: list[str] = []

    for slide_idx, slide in enumerate(prs.slides):
        slide_id = f"slide-{slide_idx + 1}"
        issues.extend(_check_bounds(slide, slide_id, prs.slide_width, prs.slide_height))
        issues.extend(_check_edge_margins(slide, slide_id, prs.slide_width, prs.slide_height))
        issues.extend(_check_overlap(slide, slide_id))
        issues.extend(_check_text_overflow(slide, slide_id))
        issues.extend(_check_density(slide, slide_id))
        issues.extend(_check_table_size(slide, slide_id))
        issues.extend(_check_chart_series(slide, slide_id))
        issues.extend(_check_chart_annotations(slide, slide_id))
        issues.extend(_check_placeholder_text(slide, slide_id))
        issues.extend(_check_contrast(slide, slide_id))
        issues.extend(_check_image_aspect_ratio(slide, slide_id))
        issues.extend(_check_empty_slide(slide, slide_id))
        issues.extend(_check_fill_ratio(slide, slide_id, prs.slide_width, prs.slide_height))
        issues.extend(_check_rasterized_slide(slide, slide_id))
        if manifest is not None:
            issues.extend(_check_template_conformance(slide, slide_id, manifest))
        slide_texts.append(_slide_text_signature(slide))

    issues.extend(_check_duplicate_slides(slide_texts))

    return issues


def _new_issue(
    slide_id: str,
    category: AuditCategory,
    description: str,
    *,
    deterministic: bool = True,
    auto_fixable: bool = False,
) -> AuditIssue:
    return AuditIssue(
        issue_id=str(uuid.uuid4()),
        slide_id=slide_id,
        category=category,
        deterministic=deterministic,
        description=description,
        auto_fixable=auto_fixable,
    )


# --------------------------------------------------------------------------
# Вёрстка
# --------------------------------------------------------------------------


def _check_bounds(slide, slide_id: str, slide_width: int, slide_height: int) -> list[AuditIssue]:
    """"Элемент вышел за границы слайда" + "текст обрезан краем слайда"
    (тот же геометрический признак — фигура с текстом, торчащая за край,
    обрезает свой текст этим же краем)."""
    issues = []
    for shape in slide.shapes:
        if None in (shape.left, shape.top, shape.width, shape.height):
            continue
        out_of_bounds = (
            shape.left < 0
            or shape.top < 0
            or shape.left + shape.width > slide_width
            or shape.top + shape.height > slide_height
        )
        if out_of_bounds:
            has_text = getattr(shape, "has_text_frame", False) and shape.text_frame.text.strip()
            description = (
                f"Текст в «{shape.name}» обрезан краем слайда (фигура выходит за границы)"
                if has_text
                else f"Фигура «{shape.name}» выходит за границы слайда"
            )
            issues.append(_new_issue(slide_id, AuditCategory.LAYOUT, description))
    return issues


def _check_edge_margins(slide, slide_id: str, slide_width: int, slide_height: int) -> list[AuditIssue]:
    """"Контент заходит в поля у краёв" — эвристика по размеру фигуры.

    У большинства шаблонов крупные структурные блоки (шапки на всю ширину,
    колонки с графиком/таблицей, фоновые подложки) намеренно примыкают вплотную
    к одному или нескольким краям слайда по дизайну — это осознанная композиция,
    а не ошибка. Случайно съехавший элемент, наоборот, почти всегда небольшой
    (иконка, подпись, мелкий текстовый блок) — крупный контентный блок (>15%
    площади слайда) у края почти никогда не является случайностью, поэтому такие
    блоки исключаются из проверки целиком; репортим только небольшие/средние
    фигуры, зашедшие в поле у края.
    """
    margin_x = int(slide_width * _EDGE_MARGIN_RATIO)
    margin_y = int(slide_height * _EDGE_MARGIN_RATIO)
    slide_area = slide_width * slide_height
    full_width_threshold = slide_width * 0.95
    full_height_threshold = slide_height * 0.95
    issues = []
    for shape in slide.shapes:
        if None in (shape.left, shape.top, shape.width, shape.height):
            continue
        if getattr(shape, "has_text_frame", False) and not shape.text_frame.text.strip():
            continue
        # Уже отмечена как "вышла за границы" в _check_bounds — не дублируем.
        if (
            shape.left < 0
            or shape.top < 0
            or shape.left + shape.width > slide_width
            or shape.top + shape.height > slide_height
        ):
            continue

        touches_left = shape.left < margin_x
        touches_right = (slide_width - (shape.left + shape.width)) < margin_x
        touches_top = shape.top < margin_y
        touches_bottom = (slide_height - (shape.top + shape.height)) < margin_y
        edge_touch_count = sum([touches_left, touches_right, touches_top, touches_bottom])
        if edge_touch_count == 0:
            continue

        shape_area_ratio = (shape.width * shape.height) / slide_area if slide_area > 0 else 0
        is_full_span = shape.width >= full_width_threshold or shape.height >= full_height_threshold
        is_large_structural_block = shape_area_ratio > 0.15 or is_full_span

        if not is_large_structural_block:
            issues.append(
                _new_issue(
                    slide_id,
                    AuditCategory.LAYOUT,
                    f"«{shape.name}» заходит в поле у края слайда (меньше {_EDGE_MARGIN_RATIO:.0%} отступа)",
                )
            )
    return issues


def _check_overlap(slide, slide_id: str) -> list[AuditIssue]:
    issues = []
    shapes = [s for s in slide.shapes if None not in (s.left, s.top, s.width, s.height)]
    for i in range(len(shapes)):
        for j in range(i + 1, len(shapes)):
            if _rects_overlap(shapes[i], shapes[j]):
                issues.append(
                    _new_issue(
                        slide_id,
                        AuditCategory.LAYOUT,
                        f"Фигуры «{shapes[i].name}» и «{shapes[j].name}» перекрываются",
                    )
                )
    return issues


def _rects_overlap(a, b) -> bool:
    ax0, ay0, ax1, ay1 = a.left, a.top, a.left + a.width, a.top + a.height
    bx0, by0, bx1, by1 = b.left, b.top, b.left + b.width, b.top + b.height
    return ax0 < bx1 and bx0 < ax1 and ay0 < by1 and by0 < ay1


def _check_text_overflow(slide, slide_id: str) -> list[AuditIssue]:
    """"Текст не поместился в свою рамку" — эвристика без реального рендера
    шрифтов: считаем, сколько строк текст займёт при word-wrap на ширине
    фигуры (символов на строку убывает пропорционально кеглю), сравниваем с
    тем, сколько строк входит по высоте фигуры при этом же кегле. Заведомо
    приблизительно — авто-уменьшение шрифта (`normAutofit`) и специфика
    конкретного шрифта не моделируются, поэтому берём консервативный запас
    20% прежде чем считать это проблемой, чтобы не плодить ложные срабатывания.
    """
    issues = []
    for shape in slide.shapes:
        if not getattr(shape, "has_text_frame", False):
            continue
        tf = shape.text_frame
        if not tf.text.strip():
            continue
        if None in (shape.width, shape.height):
            continue
        # auto_size=TEXT_TO_FIT_SHAPE/SHAPE_TO_FIT_TEXT — PowerPoint сам не даст
        # тексту перелиться, поэтому такие фигуры не проверяем.
        if tf.auto_size is not None and tf.auto_size != 0:
            continue

        width_cm = Emu(shape.width).cm
        height_cm = Emu(shape.height).cm
        if width_cm <= 0 or height_cm <= 0:
            continue

        needed_lines = 0
        for paragraph in tf.paragraphs:
            text = paragraph.text
            if not text.strip():
                needed_lines += 1
                continue
            size_pt = _paragraph_font_size_pt(paragraph)
            chars_per_line = max(_CHARS_PER_LINE_AT_10PT * (10.0 / size_pt) * width_cm, 1.0)
            needed_lines += max(1, -(-len(text) // int(chars_per_line)))  # ceil div

        avg_size_pt = _paragraph_font_size_pt(tf.paragraphs[0]) if tf.paragraphs else 18.0
        line_height_cm = (avg_size_pt / 72) * 2.54 * 1.2
        available_lines = height_cm / line_height_cm if line_height_cm > 0 else 0

        if needed_lines > available_lines * 1.2:
            issues.append(
                _new_issue(
                    slide_id,
                    AuditCategory.LAYOUT,
                    f"Текст в «{shape.name}» вероятно не помещается в рамку "
                    f"(нужно ~{needed_lines} строк, влезает ~{available_lines:.1f})",
                )
            )
    return issues


def _paragraph_font_size_pt(paragraph) -> float:
    for run in paragraph.runs:
        if run.font.size is not None:
            return run.font.size.pt
    if paragraph.font.size is not None:
        return paragraph.font.size.pt
    return 18.0


def _check_image_aspect_ratio(slide, slide_id: str) -> list[AuditIssue]:
    """"Картинка растянута, пропорции нарушены" — сравнивает соотношение
    сторон итоговой рамки (shape.width/height) с исходным соотношением
    сторон файла картинки (image.size, в пикселях исходного файла)."""
    issues = []
    for shape in slide.shapes:
        if shape.shape_type != 13:  # MSO_SHAPE_TYPE.PICTURE
            continue
        if None in (shape.width, shape.height):
            continue
        try:
            native_w, native_h = shape.image.size
        except (AttributeError, ValueError):
            continue
        if native_w <= 0 or native_h <= 0 or shape.height == 0:
            continue
        native_ratio = native_w / native_h
        placed_ratio = shape.width / shape.height
        distortion = abs(placed_ratio - native_ratio) / native_ratio
        if distortion > 0.05:  # >5% отклонения от исходных пропорций
            issues.append(
                _new_issue(
                    slide_id,
                    AuditCategory.LAYOUT,
                    f"Изображение «{shape.name}» растянуто — пропорции искажены на {distortion:.0%}",
                )
            )
    return issues


# --------------------------------------------------------------------------
# Шаблон
# --------------------------------------------------------------------------


def _check_contrast(slide, slide_id: str) -> list[AuditIssue]:
    """Проверяет контраст ЯВНО заданного цвета текста относительно фона.

    Ограничение: считается, только если и текст, и фон имеют явный srgbClr
    (не унаследованный из темы) — иначе достоверно определить итоговый цвет
    без полного разрешения темы нельзя, и слайд молча пропускается, чтобы
    не плодить ложные срабатывания. Порог 4.5:1 — из Приложения 1 ТЗ.
    """
    background_rgb = _background_rgb(slide)
    if background_rgb is None:
        return []

    issues = []
    for shape in slide.shapes:
        if not getattr(shape, "has_text_frame", False):
            continue
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                try:
                    if run.font.color.type is None:
                        continue
                    rgb = run.font.color.rgb
                except (AttributeError, TypeError, KeyError):
                    continue
                if rgb is None or not run.text.strip():
                    continue
                ratio = _contrast_ratio(rgb, background_rgb)
                if ratio < _MIN_CONTRAST_RATIO:
                    issues.append(
                        _new_issue(
                            slide_id,
                            AuditCategory.TEMPLATE,
                            f"Низкий контраст текста «{run.text[:30]}»: {ratio:.1f}:1 (минимум {_MIN_CONTRAST_RATIO}:1)",
                        )
                    )
    return issues


def _relative_luminance(rgb: RGBColor) -> float:
    def channel(value: int) -> float:
        c = value / 255
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4

    r, g, b = channel(rgb[0]), channel(rgb[1]), channel(rgb[2])
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def _contrast_ratio(rgb_a: RGBColor, rgb_b: RGBColor) -> float:
    l1, l2 = _relative_luminance(rgb_a), _relative_luminance(rgb_b)
    lighter, darker = max(l1, l2), min(l1, l2)
    return (lighter + 0.05) / (darker + 0.05)


def _background_rgb(slide) -> RGBColor | None:
    try:
        fill = slide.background.fill
        if fill.type is None:
            return None
        return fill.fore_color.rgb
    except (AttributeError, TypeError, KeyError):
        return None


def _check_template_conformance(slide, slide_id: str, manifest: DesignManifest) -> list[AuditIssue]:
    """"Шрифт не из шаблона или гарнитур больше двух", "кегль не из
    типографической шкалы шаблона", "цвет не из палитры шаблона" —
    сравнивает явно заданные (не унаследованные из темы) атрибуты runs
    против DesignManifest, извлечённого Парсером из этого же шаблона.
    """
    issues: list[AuditIssue] = []

    allowed_fonts = {manifest.typography.title.font, manifest.typography.body.font}
    allowed_sizes = {float(manifest.typography.title.size), float(manifest.typography.body.size)}
    allowed_colors = {c.upper().lstrip("#") for c in manifest.palette.values()}

    seen_fonts: set[str] = set()
    reported_font_mismatch = False
    reported_size_mismatch = False
    reported_color_mismatch = False

    for shape in slide.shapes:
        if not getattr(shape, "has_text_frame", False):
            continue
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                if not run.text.strip():
                    continue

                if run.font.name:
                    seen_fonts.add(run.font.name)
                    if run.font.name not in allowed_fonts and not reported_font_mismatch:
                        issues.append(
                            _new_issue(
                                slide_id,
                                AuditCategory.TEMPLATE,
                                f"Шрифт «{run.font.name}» не входит в типографику шаблона ({', '.join(allowed_fonts)})",
                            )
                        )
                        reported_font_mismatch = True

                if run.font.size is not None and not reported_size_mismatch:
                    size_pt = run.font.size.pt
                    if not any(abs(size_pt - allowed) <= _FONT_SIZE_TOLERANCE_PT for allowed in allowed_sizes):
                        issues.append(
                            _new_issue(
                                slide_id,
                                AuditCategory.TEMPLATE,
                                f"Кегль {size_pt:.0f}pt не входит в типографическую шкалу шаблона "
                                f"({', '.join(f'{s:.0f}pt' for s in sorted(allowed_sizes))})",
                            )
                        )
                        reported_size_mismatch = True

                try:
                    if run.font.color.type is not None and run.font.color.rgb is not None and not reported_color_mismatch:
                        color_hex = str(run.font.color.rgb).upper()
                        if color_hex not in allowed_colors:
                            issues.append(
                                _new_issue(
                                    slide_id,
                                    AuditCategory.TEMPLATE,
                                    f"Цвет текста #{color_hex} не входит в палитру шаблона",
                                )
                            )
                            reported_color_mismatch = True
                except (AttributeError, TypeError, KeyError):
                    pass

    if len(seen_fonts) > _MAX_DISTINCT_FONTS:
        issues.append(
            _new_issue(
                slide_id,
                AuditCategory.TEMPLATE,
                f"На слайде использовано {len(seen_fonts)} разных гарнитур (максимум {_MAX_DISTINCT_FONTS})",
            )
        )

    return issues


# --------------------------------------------------------------------------
# Плотность
# --------------------------------------------------------------------------


def _check_density(slide, slide_id: str) -> list[AuditIssue]:
    issues = []
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        paragraphs = [p for p in shape.text_frame.paragraphs if p.text.strip()]
        if len(paragraphs) > _MAX_BULLETS_PER_SLIDE:
            issues.append(
                _new_issue(
                    slide_id,
                    AuditCategory.DENSITY,
                    f"Больше {_MAX_BULLETS_PER_SLIDE} буллетов в блоке «{shape.name}» ({len(paragraphs)})",
                    auto_fixable=True,
                )
            )
        for p in paragraphs:
            word_count = len(p.text.split())
            if word_count > _MAX_WORDS_PER_BULLET:
                issues.append(
                    _new_issue(
                        slide_id,
                        AuditCategory.DENSITY,
                        f"Буллет длиннее {_MAX_WORDS_PER_BULLET} слов ({word_count}): «{p.text[:60]}»",
                        auto_fixable=True,
                    )
                )
    return issues


def _check_table_size(slide, slide_id: str) -> list[AuditIssue]:
    issues = []
    for shape in slide.shapes:
        if not shape.has_table:
            continue
        table = shape.table
        n_rows, n_cols = len(table.rows), len(table.columns)
        if n_rows > _MAX_TABLE_ROWS:
            issues.append(
                _new_issue(
                    slide_id,
                    AuditCategory.DENSITY,
                    f"Таблица «{shape.name}» содержит {n_rows} строк (максимум {_MAX_TABLE_ROWS})",
                )
            )
        if n_cols > _MAX_TABLE_COLS:
            issues.append(
                _new_issue(
                    slide_id,
                    AuditCategory.DENSITY,
                    f"Таблица «{shape.name}» содержит {n_cols} колонок (максимум {_MAX_TABLE_COLS})",
                )
            )
    return issues


def _check_chart_series(slide, slide_id: str) -> list[AuditIssue]:
    issues = []
    for shape in slide.shapes:
        if not shape.has_chart:
            continue
        try:
            n_series = len(shape.chart.plots[0].series)
        except (AttributeError, IndexError):
            continue
        if n_series > _MAX_CHART_SERIES:
            issues.append(
                _new_issue(
                    slide_id,
                    AuditCategory.DENSITY,
                    f"Диаграмма «{shape.name}» содержит {n_series} серий (максимум {_MAX_CHART_SERIES})",
                )
            )
    return issues


def _check_fill_ratio(slide, slide_id: str, slide_width: int, slide_height: int) -> list[AuditIssue]:
    """"Слайд заполнен меньше чем на четверть или больше чем на три четверти"
    — доля площади слайда, покрытая объединением bbox всех фигур (учитывает
    перекрытия грубо, суммируя площади без пересечений — достаточно для
    ориентировочной оценки плотности, не требует точной геометрии полигонов)."""
    slide_area = slide_width * slide_height
    if slide_area <= 0:
        return []

    covered = 0
    for shape in slide.shapes:
        if None in (shape.left, shape.top, shape.width, shape.height):
            continue
        covered += max(shape.width, 0) * max(shape.height, 0)

    ratio = min(covered / slide_area, 1.0)
    if ratio < _MIN_FILL_RATIO:
        return [
            _new_issue(
                slide_id,
                AuditCategory.DENSITY,
                f"Слайд заполнен только на {ratio:.0%} площади (меньше {_MIN_FILL_RATIO:.0%})",
            )
        ]
    if ratio > _MAX_FILL_RATIO:
        return [
            _new_issue(
                slide_id,
                AuditCategory.DENSITY,
                f"Слайд заполнен на {ratio:.0%} площади (больше {_MAX_FILL_RATIO:.0%})",
            )
        ]
    return []


# --------------------------------------------------------------------------
# Целостность
# --------------------------------------------------------------------------


def _check_placeholder_text(slide, slide_id: str) -> list[AuditIssue]:
    issues = []
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        text = shape.text_frame.text
        for pattern in _PLACEHOLDER_TEXT_PATTERNS:
            if pattern.search(text):
                issues.append(
                    _new_issue(
                        slide_id,
                        AuditCategory.INTEGRITY,
                        f"Найден текст-заглушка в «{shape.name}»: «{text[:60]}»",
                    )
                )
                break
    return issues


def _check_empty_slide(slide, slide_id: str) -> list[AuditIssue]:
    """"Пустой слайд или слайд с одним заголовком" — считаем непустыми
    текстовые фигуры, таблицы, графики и картинки; заголовок — первая
    текстовая фигура с самым большим кеглем на слайде (эвристика без
    привязки к роли placeholder'а, т.к. в этом генераторе заголовок —
    обычный textbox, не placeholder)."""
    content_shapes = 0
    text_shapes = []
    for shape in slide.shapes:
        if shape.has_table or shape.has_chart or shape.shape_type == 13:
            content_shapes += 1
            continue
        if getattr(shape, "has_text_frame", False) and shape.text_frame.text.strip():
            text_shapes.append(shape)

    if content_shapes > 0:
        return []

    if len(text_shapes) == 0:
        return [_new_issue(slide_id, AuditCategory.INTEGRITY, "Слайд полностью пуст")]
    if len(text_shapes) == 1:
        return [
            _new_issue(
                slide_id,
                AuditCategory.INTEGRITY,
                f"На слайде только заголовок «{text_shapes[0].text_frame.text[:40]}», нет содержания",
            )
        ]
    return []


def _check_chart_annotations(slide, slide_id: str) -> list[AuditIssue]:
    """"У диаграммы нет подписей осей, единиц или легенды"."""
    issues = []
    for shape in slide.shapes:
        if not shape.has_chart:
            continue
        chart = shape.chart
        has_legend = False
        try:
            has_legend = bool(chart.has_legend)
        except AttributeError:
            pass
        has_axis_titles = False
        try:
            has_axis_titles = bool(
                (chart.category_axis.has_title and chart.category_axis.axis_title.text_frame.text.strip())
                or (chart.value_axis.has_title and chart.value_axis.axis_title.text_frame.text.strip())
            )
        except (AttributeError, ValueError):
            pass
        has_data_labels = False
        try:
            has_data_labels = bool(chart.plots[0].has_data_labels)
        except (AttributeError, IndexError):
            pass

        if not (has_legend or has_axis_titles or has_data_labels):
            issues.append(
                _new_issue(
                    slide_id,
                    AuditCategory.INTEGRITY,
                    f"У диаграммы «{shape.name}» нет ни легенды, ни подписей осей, ни подписей данных",
                )
            )
    return issues


def _check_rasterized_slide(slide, slide_id: str) -> list[AuditIssue]:
    """"Слайд оказался картинкой, а не редактируемыми объектами" — слайд,
    целиком состоящий из одной картинки, покрывающей почти всю площадь, и
    больше ничего (кроме опционально пустых фигур)."""
    picture_shapes = [s for s in slide.shapes if s.shape_type == 13]
    other_meaningful = [
        s
        for s in slide.shapes
        if s.shape_type != 13
        and (
            (getattr(s, "has_text_frame", False) and s.text_frame.text.strip())
            or getattr(s, "has_table", False)
            or getattr(s, "has_chart", False)
        )
    ]
    if len(picture_shapes) == 1 and not other_meaningful:
        return [
            _new_issue(
                slide_id,
                AuditCategory.INTEGRITY,
                "Слайд состоит из единственной картинки без редактируемого текста/таблиц/графиков",
            )
        ]
    return []


def _slide_text_signature(slide) -> str:
    """Конкатенация всего текста слайда в порядке фигур — используется для
    поиска дублирующихся слайдов. Не учитывает геометрию/стиль намеренно:
    два слайда с одинаковым текстом, но разной раскладкой, всё равно
    дублируют друг друга по смыслу."""
    parts = []
    for shape in slide.shapes:
        if getattr(shape, "has_text_frame", False):
            text = shape.text_frame.text.strip()
            if text:
                parts.append(text)
    return "\n".join(parts)


def _check_duplicate_slides(slide_texts: list[str]) -> list[AuditIssue]:
    """"Два слайда дублируют друг друга" — сравнение по точному совпадению
    текстовой подписи слайда (после нормализации пробелов). Слайды без
    текста (например, только с картинкой) не сравниваются — иначе два пустых
    слайда ложно считались бы дубликатами."""
    issues = []
    seen: dict[str, int] = {}
    for idx, text in enumerate(slide_texts):
        normalized = " ".join(text.split())
        if not normalized:
            continue
        if normalized in seen:
            first_idx = seen[normalized]
            issues.append(
                _new_issue(
                    f"slide-{idx + 1}",
                    AuditCategory.INTEGRITY,
                    f"Слайд дублирует slide-{first_idx + 1} (идентичный текст)",
                )
            )
        else:
            seen[normalized] = idx
    return issues


# --------------------------------------------------------------------------
# Валидация контента через VLM (недетерминированная, Приложение 1)
# --------------------------------------------------------------------------

_VLM_QUESTIONS = [
    "Заголовок содержит вывод, а не просто называет тему?",
    "Содержимое слайда соответствует заголовку?",
    "Слайд пересказывается одним предложением?",
    "Все цифры и факты со слайда выглядят взаимно согласованными (не противоречат друг другу)?",
    "На слайде есть содержание, а не только заголовок?",
    "Картинки и иконки (если есть) относятся к теме слайда?",
    "Нет служебного мусора: реплик спикера, кусков промпта, технических артефактов?",
    "Текст без опечаток?",
    "Текст на слайде на одном языке (без случайного смешения языков)?",
    "Все строки таблицы и элементы легенды (если есть) работают на мысль слайда, а не случайны?",
    "Слайд логично мог бы стоять в деловой презентации на заявленную тему?",
]

_VLM_QUESTION_KEYS = [f"q{i}" for i in range(1, len(_VLM_QUESTIONS) + 1)]


async def render_slides_to_png(pptx_path: str | Path, output_dir: str | Path) -> list[Path]:
    """Рендерит каждый слайд .pptx в отдельный .png через LibreOffice (pptx->pdf)
    + pdftoppm (pdf->png, одна страница = один слайд). Возвращает пути к PNG
    в порядке слайдов (1-indexed имена файлов, отсортированные лексикографически
    после дозаполнения нулями — pdftoppm сам поддерживает -r/-png с нумерацией)."""
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    pdftoppm = shutil.which("pdftoppm")
    if pdftoppm is None:
        raise ExportError("pdftoppm не найден в PATH — установите poppler-utils для VLM-аудита")

    pdf_path = await export_presentation(pptx_path, output_dir, "pdf")

    prefix = output_dir / "slide"
    process = await asyncio.create_subprocess_exec(
        pdftoppm,
        "-png",
        "-r",
        "110",
        str(pdf_path),
        str(prefix),
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE,
    )
    stdout, stderr = await process.communicate()
    if process.returncode != 0:
        raise ExportError(f"pdftoppm завершился с кодом {process.returncode}: {stderr.decode(errors='ignore')}")

    png_files = sorted(output_dir.glob("slide-*.png"))
    if not png_files:
        # На некоторых сборках poppler имя без ведущих нулей: slide-1.png vs slide-01.png.
        png_files = sorted(output_dir.glob("slide*.png"))

    def _page_num(path: Path) -> int:
        match = re.search(r"(\d+)$", path.stem)
        return int(match.group(1)) if match else 0

    return sorted(png_files, key=_page_num)


async def audit_content_with_vlm(
    pptx_path: str | Path,
    *,
    workdir: str | Path | None = None,
    concurrency: int = 3,
) -> list[AuditIssue]:
    """Недетерминированные проверки смысла слайда через VLM — 11 да/нет
    вопросов из Приложения 1 ("Валидация контента"), один VLM-вызов на
    слайд (не 11), картинка слайда передаётся как data:image/png;base64.

    Каждый ответ "нет" -> один AuditIssue(deterministic=False, category=CONTENT).
    Слайды обрабатываются с ограниченным параллелизмом (`concurrency`), чтобы
    не упереться в rate-limit VLM-провайдера на колоде из 10-15 слайдов.
    """
    settings = get_settings()
    pptx_path = Path(pptx_path)
    workdir = Path(workdir) if workdir is not None else pptx_path.parent / f"{pptx_path.stem}_vlm_render"

    png_paths = await render_slides_to_png(pptx_path, workdir)
    if not png_paths:
        return []

    client = LLMClient(
        base_url=settings.VLM_BASE_URL,
        model_name=settings.VLM_MODEL_NAME,
        api_key=settings.VLM_API_KEY,
    )

    semaphore = asyncio.Semaphore(max(1, concurrency))

    async def _audit_one(idx: int, png_path: Path) -> list[AuditIssue]:
        slide_id = f"slide-{idx + 1}"
        async with semaphore:
            try:
                answers = await _ask_vlm_about_slide(client, png_path)
            except Exception as exc:  # noqa: BLE001 — сбой VLM не должен ронять весь аудит
                logger.warning("VLM-аудит слайда %s упал: %s", slide_id, exc)
                return [
                    _new_issue(
                        slide_id,
                        AuditCategory.CONTENT,
                        f"VLM-проверка контента не выполнена: {exc}",
                        deterministic=False,
                    )
                ]

        slide_issues = []
        for question, key in zip(_VLM_QUESTIONS, _VLM_QUESTION_KEYS):
            verdict = answers.get(key)
            if verdict is False:
                reason = answers.get(f"{key}_reason", "")
                description = question
                if reason:
                    description += f" — нет: {reason}"
                slide_issues.append(
                    _new_issue(slide_id, AuditCategory.CONTENT, description, deterministic=False)
                )
        return slide_issues

    results = await asyncio.gather(*(_audit_one(i, p) for i, p in enumerate(png_paths)))
    issues: list[AuditIssue] = []
    for slide_issues in results:
        issues.extend(slide_issues)
    return issues


async def _ask_vlm_about_slide(client: LLMClient, png_path: Path) -> dict:
    image_b64 = base64.b64encode(png_path.read_bytes()).decode("ascii")
    data_url = f"data:image/png;base64,{image_b64}"

    questions_block = "\n".join(f"{key}: {question}" for key, question in zip(_VLM_QUESTION_KEYS, _VLM_QUESTIONS))
    system_prompt = (
        "Ты — контроль качества деловых презентаций. На вход — картинка одного слайда. "
        "Ответь на каждый из перечисленных вопросов da/net (true/false в JSON), строго по тому, "
        "что видно на картинке — не придумывай контекст, которого там нет. "
        "Для каждого вопроса, где ответ false, добавь короткое поле <key>_reason с объяснением "
        "в одно предложение. Ответ — СТРОГО JSON-объект с ключами q1..q11 (boolean) и опциональными "
        "*_reason (string), без markdown-обёртки."
    )
    user_content = [
        {"type": "text", "text": f"Вопросы:\n{questions_block}"},
        {"type": "image_url", "image_url": {"url": data_url}},
    ]

    return await client.chat_json_multimodal(system_prompt, user_content)
