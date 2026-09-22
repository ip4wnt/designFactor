"""Компоновщик слайдов: DesignManifest + ContentPlan -> .pptx.

Вёрстка — синтетическая, через библиотеку геометрических паттернов
(app/pipeline/layout_engine.py + app/config/layout_patterns.yaml), а НЕ через
заполнение plaeholder-ов исходных slide_layouts шаблона. Так решается
фундаментальное ограничение первой версии: шаблон может не иметь ни одного
макета с ролью BODY/CHART/TABLE (только "Обложка"/TITLE-only макеты — типичный
случай для реальных корпоративных шаблонов), и тогда placeholder-матчинг
не может предложить вариативность вообще.

Фирменный стиль результата обеспечивают дизайн-токены DesignManifest
(палитра/типографика темы, см. styling.py), а не переиспользование самих
plaeholder-ов — они читаются один раз из XML темы Парсером и применяются к
любой геометрии.

Три варианта вёрстки (variant_a/b/c) — три разные стратегии ВЫБОРА паттерна
для одного и того же content_plan (см. layout_engine.select_pattern), не три
разных рендерера: рендерер один, разница только в том, какой паттерн он
получает для каждого слайда.

Диаграммы и таблицы — нативные объекты python-pptx (add_chart/add_table),
растровые изображения слайдов результатом не считаются.
"""
from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt

from app.pipeline.capacity import fit_chart_data, fit_table_data
from app.pipeline.layout_engine import (
    Pattern,
    Zone,
    ZoneAssignment,
    assign_zones,
    select_pattern,
    title_zone,
    zone_bbox_emu,
)
from app.pipeline.styling import style_body_textbox, style_chart, style_table, style_title_textbox
from app.schemas.content_plan import ContentBlock, ContentBlockType, ContentPlan, SlideSpec
from app.schemas.design_manifest import DesignManifest

_CHART_TYPE_MAP = {
    "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "line": XL_CHART_TYPE.LINE,
    "pie": XL_CHART_TYPE.PIE,
}

VALID_VARIANTS = ("variant_a", "variant_b", "variant_c")


def assemble_presentation(
    template_path: str | Path,
    manifest: DesignManifest,
    content_plan: ContentPlan,
    output_path: str | Path,
    variant: str = "variant_a",
) -> Path:
    if variant not in VALID_VARIANTS:
        raise ValueError(f"Неизвестный вариант вёрстки: {variant}")

    prs = Presentation(str(template_path))

    # Presentation(template_path) открывает ИСХОДНЫЙ файл целиком, вместе со
    # всеми слайдами-образцами, которые в нём уже есть (это те самые слайды,
    # которые Парсер разбирал, чтобы понять дизайн-систему). Нам нужны только
    # master/theme из этого файла — сами слайды-образцы в выходную колоду
    # попадать не должны, иначе результат = шаблон + новые слайды.
    _strip_existing_slides(prs)

    # Пустой (blank) макет темы — на нём нет никаких plaeholder-ов, поэтому
    # синтетические зоны паттерна не конфликтуют с чужой геометрией. Тема
    # (цвета/шрифты по умолчанию) всё равно наследуется на уровне presentation,
    # так что новые textbox/chart/table получают её через явные токены стиля.
    blank_layout = _find_blank_layout(prs)

    for slide_spec in content_plan.slides:
        pattern = select_pattern(slide_spec, variant)
        slide = prs.slides.add_slide(blank_layout)
        _render_slide(slide, pattern, slide_spec, manifest)

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))
    return output_path


def _find_blank_layout(prs: Presentation):
    """Ищет макет без plaeholder-ов вообще (обычно называется "Пустой слайд"/
    "Blank"). Если такого нет — берёт макет с наименьшим числом plaeholder-ов
    и просто удаляет их со слайда после add_slide (см. _render_slide), т.к.
    вёрстка полностью синтетическая и старые plaeholder-ы никогда не используются.
    """
    layouts = list(prs.slide_masters[0].slide_layouts)
    no_placeholder = [layout for layout in layouts if len(list(layout.placeholders)) == 0]
    if no_placeholder:
        return no_placeholder[0]
    return min(layouts, key=lambda layout: len(list(layout.placeholders)))


def _strip_existing_slides(prs: Presentation) -> None:
    """Удаляет все слайды-образцы, уже присутствующие в загруженном шаблоне.

    python-pptx не даёт slides.remove() — удаляем каждый <p:sldId> из
    sldIdLst и дропаем связанную slide part из package — иначе мёртвые
    части остаются в .pptx и раздувают файл без надобности.
    """
    sldIdLst = prs.slides._sldIdLst
    for sldId in list(sldIdLst):
        rId = sldId.rId
        prs.part.drop_rel(rId)
        sldIdLst.remove(sldId)


def _strip_all_placeholders(slide) -> None:
    """Удаляет унаследованные от макета plaeholder-ы со слайда — вёрстка
    целиком синтетическая (см. модульный docstring), унаследованные
    plaeholder-ы дают пустые декоративные рамки в аудите, если их не убрать.
    """
    for shape in list(slide.shapes):
        if shape.is_placeholder:
            shape._element.getparent().remove(shape._element)


def _render_slide(slide, pattern: Pattern, slide_spec: SlideSpec, manifest: DesignManifest) -> None:
    _strip_all_placeholders(slide)
    slide_width = manifest.slide_width_emu
    slide_height = manifest.slide_height_emu

    title_z = title_zone(pattern)
    if title_z is not None:
        _render_title(slide, title_z, slide_spec.title, manifest, slide_width, slide_height)

    assignments = assign_zones(pattern, slide_spec)
    for assignment in assignments:
        _render_zone(slide, assignment, manifest, slide_width, slide_height)


def _render_title(slide, zone: Zone, title_text: str, manifest: DesignManifest, sw: int, sh: int) -> None:
    left, top, width, height = zone_bbox_emu(zone, sw, sh)
    textbox = slide.shapes.add_textbox(left, top, width, height)
    text_frame = textbox.text_frame
    text_frame.word_wrap = True
    text_frame.text = title_text
    text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
    style_title_textbox(text_frame, manifest)


def _render_zone(slide, assignment: ZoneAssignment, manifest: DesignManifest, sw: int, sh: int) -> None:
    block = assignment.block
    if block is None:
        return
    bbox = zone_bbox_emu(assignment.zone, sw, sh)

    if block.type in (ContentBlockType.BULLETS, ContentBlockType.TEXT):
        _render_text_block(slide, bbox, block, manifest)
    elif block.type == ContentBlockType.CHART:
        _add_native_chart(slide, bbox, block, manifest)
    elif block.type == ContentBlockType.TABLE:
        _add_native_table(slide, bbox, block, manifest)


def _render_text_block(slide, bbox: tuple[int, int, int, int], block: ContentBlock, manifest: DesignManifest) -> None:
    left, top, width, height = bbox
    textbox = slide.shapes.add_textbox(left, top, width, height)
    text_frame = textbox.text_frame
    text_frame.word_wrap = True

    if block.type == ContentBlockType.BULLETS:
        items = block.bullets or [""]
        text_frame.text = items[0]
        for extra in items[1:]:
            p = text_frame.add_paragraph()
            p.text = extra
        for paragraph in text_frame.paragraphs:
            paragraph.level = 0
    else:
        text_frame.text = block.text or ""

    style_body_textbox(text_frame, manifest)
    # Буллеты получают явный маркер — python-pptx не рисует его сам без
    # <a:buChar>/<a:buAutoNum> в pPr, а голый текстовый фрейм по умолчанию
    # выводит параграфы без каких-либо маркеров.
    if block.type == ContentBlockType.BULLETS:
        for paragraph in text_frame.paragraphs:
            _set_bullet_char(paragraph)


def _set_bullet_char(paragraph) -> None:
    from pptx.oxml.ns import qn
    from pptx.oxml import parse_xml

    pPr = paragraph._p.get_or_add_pPr()
    for tag in ("a:buChar", "a:buAutoNum", "a:buNone"):
        existing = pPr.find(qn(tag))
        if existing is not None:
            pPr.remove(existing)
    bu_char = parse_xml(
        '<a:buChar xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" char="•"/>'
    )
    pPr.append(bu_char)


def _add_native_chart(slide, bbox: tuple[int, int, int, int], block: ContentBlock, manifest: DesignManifest) -> None:
    if block.chart is None:
        return
    left, top, width, height = bbox

    categories, series, _truncated = fit_chart_data(
        block.chart.categories, block.chart.series, width
    )

    chart_data = CategoryChartData()
    chart_data.categories = categories
    for series_name, values in series.items():
        chart_data.add_series(series_name, values)

    xl_chart_type = _CHART_TYPE_MAP[block.chart.chart_type.value]
    graphic_frame = slide.shapes.add_chart(xl_chart_type, left, top, width, height, chart_data)
    style_chart(graphic_frame.chart, manifest)


def _add_native_table(slide, bbox: tuple[int, int, int, int], block: ContentBlock, manifest: DesignManifest) -> None:
    if block.table is None:
        return
    left, top, width, height = bbox

    headers, data_rows, _truncated = fit_table_data(
        block.table.headers,
        block.table.rows,
        height,
        width,
        manifest.typography.body.size,
    )

    rows = len(data_rows) + 1
    cols = max(len(headers), 1)
    graphic_frame = slide.shapes.add_table(rows, cols, left, top, width, height)
    table = graphic_frame.table

    for col_idx, header in enumerate(headers):
        table.cell(0, col_idx).text = header
    for row_idx, row in enumerate(data_rows, start=1):
        for col_idx, value in enumerate(row):
            table.cell(row_idx, col_idx).text = str(value)

    style_table(table, manifest)
