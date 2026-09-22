"""Применение стиля темы шаблона к нативным таблицам и графикам.

Источник стиля — только DesignManifest.palette/typography (уже извлечённые
Парсером из a:clrScheme и p:txStyles темы), без анализа реальных
таблиц/графиков на слайдах шаблона: в этом проекте образцы могут вообще не
содержать готовых таблиц/графиков, а тема есть всегда.

Правила, которыми руководствуемся:
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
    """Палитра accent1..accent6 из темы, в фиксированном порядке для серий графика."""
    colors = [
        _hex_to_rgbcolor(manifest.palette.get(name), fallback)
        for name, fallback in zip(
            _ACCENT_ORDER,
            ("4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47"),
        )
    ]
    return colors


def style_table(table, manifest: DesignManifest) -> None:
    """Красит нативную python-pptx таблицу в стиле темы: шапка + чередование строк."""
    header_bg = _hex_to_rgbcolor(manifest.palette.get("accent1"), "4472C4")
    header_text = _contrasting_text_color(header_bg, manifest)
    body_font = manifest.typography.body
    row_bg_even = _hex_to_rgbcolor(manifest.palette.get("lt1"), "FFFFFF")
    row_bg_odd = _tint(header_bg, 0.85)

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
    min_row_height = Emu(int(body_font.size * 1.6 * 12700))
    for row in table.rows:
        if row.height < min_row_height:
            row.height = min_row_height

    for col_idx in range(n_cols):
        header_cell = table.cell(0, col_idx)
        header_cell.fill.solid()
        header_cell.fill.fore_color.rgb = header_bg
        _style_cell_text(header_cell, header_text, body_font, bold=True)

    for row_idx in range(1, n_rows):
        row_color = row_bg_even if row_idx % 2 == 1 else row_bg_odd
        row_text_color = _contrasting_text_color(row_color, manifest)
        for col_idx in range(n_cols):
            cell = table.cell(row_idx, col_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = row_color
            _style_cell_text(cell, row_text_color, body_font, bold=False)


def _style_cell_text(cell, color: RGBColor, font_style, bold: bool) -> None:
    for paragraph in cell.text_frame.paragraphs:
        if not paragraph.runs:
            # Пустая ячейка/заголовок без runs — python-pptx создаёт run лениво
            # только при первом обращении к paragraph.font, что достаточно
            # для окраски заголовка "по умолчанию" без текста.
            paragraph.font.name = font_style.font
            paragraph.font.size = Pt(font_style.size)
            paragraph.font.bold = bold
            paragraph.font.color.rgb = color
            continue
        for run in paragraph.runs:
            run.font.name = font_style.font
            run.font.size = Pt(font_style.size)
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

    try:
        chart.font.name = manifest.typography.body.font
        chart.font.size = Pt(max(manifest.typography.body.size - 2, 8))
    except AttributeError:
        pass

    if chart.has_legend:
        chart.legend.font.name = manifest.typography.body.font
        chart.legend.font.size = Pt(max(manifest.typography.body.size - 2, 8))
