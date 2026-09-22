"""Применение стиля темы шаблона к нативным таблицам и графикам.

Главный источник — DesignManifest.table_style/chart_style, если парсер нашёл реальную
таблицу/график на одном из слайдов-образцов шаблона (app/pipeline/parser.py,
_extract_sample_styles) — это точный фирменный стиль, а не выведенный. Для любого
поля, где такого токена нет (None — в образце его не было, или в шаблоне вовсе не
нашлось таблиц/графика), фоллбек выводит значение из palette/typography, как
раньше:

Правила фоллбека:
- Заголовок таблицы красится в accent1 (сплошная заливка), текст заголовка —
  в контрастный цвет (lt1 на тёмном accent, dk1 на светлом accent).
  Обычные строки чередуются: bg1 / лёгкий тон accent1 (полосатая заливка)
  для читаемости.
- Столбцы/серии графика получают цвета по кругу accent1..accent6 в этом
  порядке — так PowerPoint обычно значит "фирменная палитра".
- Шрифт заголовков и ячеек — typography.body.font; вес — typography.body.bold.
"""
from __future__ import annotations

from pptx.dml.color import RGBColor
from pptx.util import Emu, Pt

from app.schemas.design_manifest import DesignManifest

_ACCENT_ORDER = ("accent1", "accent2", "accent3", "accent4", "accent5", "accent6")


def _hex_to_rgbcolor(hex_value: str | None, fallback: str = "808080") -> RGBColor:
    value = (hex_value or fallback).lstrip("#")
    if len(value) != 6:
        value = fallback
    return RGBColor.from_string(value.upper())


def _relative_luminance(rgb: RGBColor) -> float:
    def channel(v: int) -> float:
        c = v / 255
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4

    r, g, b = channel(rgb[0]), channel(rgb[1]), channel(rgb[2])
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def _contrasting_text_color(background: RGBColor, manifest: DesignManifest) -> RGBColor:
    """Выбирает dk1 или lt1 из темы — какой из них контрастнее на фоне background."""
    dk1 = _hex_to_rgbcolor(manifest.palette.get("dk1"), "000000")
    lt1 = _hex_to_rgbcolor(manifest.palette.get("lt1"), "FFFFFF")
    bg_luminance = _relative_luminance(background)
    # Тёмный фон -> светлый текст, светлый фон -> тёмный текст.
    return lt1 if bg_luminance < 0.5 else dk1


def _tint(rgb: RGBColor, amount: float) -> RGBColor:
    """Осветляет цвет к белому на `amount` (0..1) — для чередующихся строк таблицы."""
    r = int(rgb[0] + (255 - rgb[0]) * amount)
    g = int(rgb[1] + (255 - rgb[1]) * amount)
    b = int(rgb[2] + (255 - rgb[2]) * amount)
    return RGBColor(r, g, b)


def accent_colors(manifest: DesignManifest) -> list[RGBColor]:
    """Палитра для серий графика: цвета, реально используемые в шаблоне
    (manifest.chart_style.series_colors), если такие были найдены парсером, иначе —
    accent1..accent6 из темы в фиксированном порядке.
    """
    if manifest.chart_style is not None and manifest.chart_style.series_colors:
        return [_hex_to_rgbcolor(c) for c in manifest.chart_style.series_colors]
    colors = [
        _hex_to_rgbcolor(manifest.palette.get(name), fallback)
        for name, fallback in zip(
            _ACCENT_ORDER,
            ("4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47"),
        )
    ]
    return colors


def style_table(table, manifest: DesignManifest) -> None:
    """Красит нативную python-pptx таблицу в стиле темы: шапка + чередование строк.

    Где есть manifest.table_style (токены, извлечённые из реальной таблицы в шаблоне),
    они перебивают соответствующие значения, выведенные из palette/typography.
    """
    tokens = manifest.table_style
    body_font = manifest.typography.body

    header_bg = _hex_to_rgbcolor(
        tokens.header_fill if tokens else None,
        (manifest.palette.get("accent1") or "4472C4").lstrip("#"),
    )
    header_text = (
        _hex_to_rgbcolor(tokens.header_text_color)
        if tokens and tokens.header_text_color
        else _contrasting_text_color(header_bg, manifest)
    )
    header_bold = tokens.header_bold if tokens and tokens.header_bold is not None else True
    row_bg_even = _hex_to_rgbcolor(
        tokens.row_odd_fill if tokens else None,
        (manifest.palette.get("lt1") or "FFFFFF").lstrip("#"),
    )
    row_bg_odd = (
        _hex_to_rgbcolor(tokens.row_even_fill)
        if tokens and tokens.row_even_fill
        else _tint(header_bg, 0.85)
    )
    cell_font_name = (tokens.body_font if tokens and tokens.body_font else None) or body_font.font
    cell_font_size = (tokens.body_size if tokens and tokens.body_size else None) or body_font.size

    # Отключаем встроенный стиль таблицы PowerPoint, чтобы наши явные заливки
    # не перебивались темой таблицы (banded rows/first row) из макета.
    table.first_row = False
    table.horz_banding = False

    n_rows = len(table.rows)
    n_cols = len(table.columns)

    # python-pptx делит заданную высоту поровну между строками при создании
    # таблицы, но PowerPoint/LibreOffice авто-увеличивают строку, если текст
    # с текущим кеглем в неё не влезает по их собственному расчёту отступов —
    # тогда таблица визуально "выезжает" за исходный bbox. Задаём явную
    # минимальную высоту строки от кегля темы, чтобы её итоговый размер был
    # осознанным решением, а не слишком узкой полосой после деления поровну.
    min_row_height = Emu(int(cell_font_size * 1.6 * 12700))
    for row in table.rows:
        if row.height < min_row_height:
            row.height = min_row_height

    for col_idx in range(n_cols):
        header_cell = table.cell(0, col_idx)
        header_cell.fill.solid()
        header_cell.fill.fore_color.rgb = header_bg
        _style_cell_text_raw(header_cell, header_text, cell_font_name, cell_font_size, bold=header_bold)

    for row_idx in range(1, n_rows):
        row_color = row_bg_even if row_idx % 2 == 1 else row_bg_odd
        row_text_color = _contrasting_text_color(row_color, manifest)
        for col_idx in range(n_cols):
            cell = table.cell(row_idx, col_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = row_color
            _style_cell_text_raw(cell, row_text_color, cell_font_name, cell_font_size, bold=False)


def _style_cell_text(cell, color: RGBColor, font_style, bold: bool) -> None:
    _style_cell_text_raw(cell, color, font_style.font, font_style.size, bold)


def _style_cell_text_raw(cell, color: RGBColor, font_name: str, size_pt: int, bold: bool) -> None:
    for paragraph in cell.text_frame.paragraphs:
        if not paragraph.runs:
            # Пустая ячейка/заголовок без runs — python-pptx создаёт run лениво
            # только при первом обращении к paragraph.font, что достаточно
            # для окраски заголовка "по умолчанию" без текста.
            paragraph.font.name = font_name
            paragraph.font.size = Pt(size_pt)
            paragraph.font.bold = bold
            paragraph.font.color.rgb = color
            continue
        for run in paragraph.runs:
            run.font.name = font_name
            run.font.size = Pt(size_pt)
            run.font.bold = bold
            run.font.color.rgb = color


def style_chart(chart, manifest: DesignManifest) -> None:
    """Красит серии нативного python-pptx графика в порядок accent1..accent6 темы."""
    colors = accent_colors(manifest)
    plot = chart.plots[0]

    for series_idx, series in enumerate(plot.series):
        color = colors[series_idx % len(colors)]
        try:
            series.format.fill.solid()
            series.format.fill.fore_color.rgb = color
        except (AttributeError, TypeError):
            # Line-графики красят линию, а не заливку.
            try:
                series.format.line.color.rgb = color
                series.format.line.width = Pt(2.25)
            except (AttributeError, TypeError):
                pass

    chart_font_name = (
        (manifest.chart_style.font if manifest.chart_style else None) or manifest.typography.body.font
    )
    chart_font_size = (
        (manifest.chart_style.font_size if manifest.chart_style else None) or manifest.typography.body.size
    )
    try:
        chart.font.name = chart_font_name
        chart.font.size = Pt(max(chart_font_size - 2, 8))
    except AttributeError:
        pass

    if chart.has_legend:
        chart.legend.font.name = chart_font_name
        chart.legend.font.size = Pt(max(chart_font_size - 2, 8))


def style_title_textbox(text_frame, manifest: DesignManifest) -> None:
    """Красит текстовый фрейм заголовка в стиль темы (typography.title)."""
    style = manifest.typography.title
    color = _hex_to_rgbcolor(manifest.palette.get("dk1"), "000000")
    for paragraph in text_frame.paragraphs:
        _apply_run_style(paragraph, style.font, style.size, style.bold, color)


def style_body_textbox(text_frame, manifest: DesignManifest, bold: bool | None = None) -> None:
    """Красит текстовый фрейм тела (буллеты/текст) в стиль темы (typography.body)."""
    style = manifest.typography.body
    color = _hex_to_rgbcolor(manifest.palette.get("dk1"), "000000")
    for paragraph in text_frame.paragraphs:
        _apply_run_style(paragraph, style.font, style.size, bold if bold is not None else style.bold, color)


def _apply_run_style(paragraph, font_name: str, size_pt: int, bold: bool, color: RGBColor) -> None:
    if not paragraph.runs:
        paragraph.font.name = font_name
        paragraph.font.size = Pt(size_pt)
        paragraph.font.bold = bold
        paragraph.font.color.rgb = color
        return
    for run in paragraph.runs:
        run.font.name = font_name
        run.font.size = Pt(size_pt)
        run.font.bold = bold
        run.font.color.rgb = color
