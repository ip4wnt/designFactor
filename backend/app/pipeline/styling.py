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


_A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
_P_NS = "{http://schemas.openxmlformats.org/presentationml/2006/main}"


def _scheme_hex(prs, name: str) -> str | None:
    """HEX цвета схемы темы (bg1/tx1 → lt1/dk1) первого мастера."""
    aliases = {"bg1": "lt1", "tx1": "dk1", "bg2": "lt2", "tx2": "dk2"}
    name = aliases.get(name, name)
    try:
        theme_part = prs.slide_masters[0].part.part_related_by(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
        )
        from lxml import etree

        root = etree.fromstring(theme_part.blob)
        node = root.find(f".//{_A_NS}clrScheme/{_A_NS}{name}")
        if node is None:
            return None
        srgb = node.find(f"{_A_NS}srgbClr")
        if srgb is not None:
            return srgb.get("val")
        sys_ = node.find(f"{_A_NS}sysClr")
        if sys_ is not None:
            return sys_.get("lastClr")
    except Exception:  # noqa: BLE001
        return None
    return None


def _fill_luminance(fill_parent, prs) -> float | None:
    """Светимость сплошной заливки (a:solidFill) внутри элемента, если она есть."""
    solid = fill_parent.find(f"{_A_NS}solidFill")
    if solid is None:
        return None
    srgb = solid.find(f"{_A_NS}srgbClr")
    if srgb is not None and srgb.get("val"):
        return _relative_luminance(_hex_to_rgbcolor(srgb.get("val")))
    scheme = solid.find(f"{_A_NS}schemeClr")
    if scheme is not None:
        hex_value = _scheme_hex(prs, scheme.get("val") or "")
        if hex_value:
            lum = _relative_luminance(_hex_to_rgbcolor(hex_value))
            mod = scheme.find(f"{_A_NS}lumMod")
            off = scheme.find(f"{_A_NS}lumOff")
            if mod is not None or off is not None:
                lum = lum * (int(mod.get("val")) / 100000 if mod is not None else 1) + (int(off.get("val")) / 100000 if off is not None else 0)
            return max(0.0, min(1.0, lum))
    return None


def _picture_luminance(part, blob_rel_id: str) -> float | None:
    try:
        from io import BytesIO

        from PIL import Image

        image_part = part.related_part(blob_rel_id)
        image = Image.open(BytesIO(image_part.blob)).convert("L")
        image.thumbnail((64, 64))
        pixels = list(image.getdata())
        return (sum(pixels) / len(pixels)) / 255 if pixels else None
    except Exception:  # noqa: BLE001
        return None


def background_luminance(slide) -> float:
    """Оценка светимости фона слайда (0 — чёрный, 1 — белый).

    Порядок: заливка фона слайда → макета → мастера; картинка почти на весь
    слайд (на слайде/макете/мастере) → средняя яркость картинки; иначе lt1 темы.
    """
    prs = slide.part.package.presentation_part.presentation
    slide_area = int(prs.slide_width) * int(prs.slide_height)
    chain = [slide, slide.slide_layout, slide.slide_layout.slide_master]
    for holder in chain:
        bg = holder._element.find(f"{_P_NS}cSld/{_P_NS}bg")
        if bg is not None:
            bg_pr = bg.find(f"{_P_NS}bgPr")
            if bg_pr is not None:
                lum = _fill_luminance(bg_pr, prs)
                if lum is not None:
                    return lum
                blip = bg_pr.find(f".//{_A_NS}blip")
                if blip is not None and blip.get(f"{{http://schemas.openxmlformats.org/officeDocument/2006/relationships}}embed"):
                    lum = _picture_luminance(holder.part, blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"))
                    if lum is not None:
                        return lum
            bg_ref = bg.find(f"{_P_NS}bgRef")
            if bg_ref is not None:
                scheme = bg_ref.find(f"{_A_NS}schemeClr")
                if scheme is not None:
                    hex_value = _scheme_hex(prs, scheme.get("val") or "")
                    if hex_value:
                        return _relative_luminance(_hex_to_rgbcolor(hex_value))
        # Картинка/плашка почти на весь слайд
        for shape in holder.shapes:
            try:
                area = int(shape.width) * int(shape.height)
            except (TypeError, ValueError):
                continue
            if area < 0.75 * slide_area:
                continue
            if shape.shape_type is not None and "PICTURE" in str(shape.shape_type):
                blip = shape._element.find(f".//{_A_NS}blip")
                rid = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed") if blip is not None else None
                if rid:
                    lum = _picture_luminance(holder.part, rid)
                    if lum is not None:
                        return lum
            sppr = shape._element.find(f"{_P_NS}spPr")
            if sppr is not None:
                lum = _fill_luminance(sppr, prs)
                if lum is not None:
                    return lum
    hex_value = _scheme_hex(prs, "lt1")
    return _relative_luminance(_hex_to_rgbcolor(hex_value)) if hex_value else 1.0


def slide_is_dark(slide) -> bool:
    return background_luminance(slide) < 0.35


def style_chart(chart, manifest: DesignManifest, dark: bool | None = None, slide=None) -> None:
    """Красит серии нативного python-pptx графика в порядок accent1..accent6 темы.

    Текст графика (оси, легенда, подписи данных) — контрастный к фону слайда:
    на тёмном фоне светлый, на светлом — тёмный. Подписи данных включены,
    жирные и чуть крупнее легенды; легенда снизу.
    """
    colors = accent_colors(manifest)
    plot = chart.plots[0]
    if dark is None and slide is not None:
        try:
            dark = slide_is_dark(slide)
        except Exception:  # noqa: BLE001
            dark = False
    dark = bool(dark)
    text_color = _hex_to_rgbcolor(manifest.palette.get("lt1"), "FFFFFF") if dark else _hex_to_rgbcolor(manifest.palette.get("dk1"), "1B1B1F")
    if dark and _relative_luminance(text_color) < 0.5:
        text_color = RGBColor(0xFF, 0xFF, 0xFF)
    if not dark and _relative_luminance(text_color) > 0.5:
        text_color = RGBColor(0x1B, 0x1B, 0x1F)
    grid_color = RGBColor(0x5A, 0x5F, 0x6B) if dark else RGBColor(0xD0, 0xD4, 0xDC)

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
    base_size = max(chart_font_size - 2, 8)
    try:
        chart.font.name = chart_font_name
        chart.font.size = Pt(base_size)
        chart.font.color.rgb = text_color
    except AttributeError:
        pass

    # Легенда: снизу, контрастная; для одной серии не нужна
    try:
        if len(list(plot.series)) <= 1:
            chart.has_legend = False
        else:
            from pptx.enum.chart import XL_LEGEND_POSITION

            chart.has_legend = True
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            chart.legend.include_in_layout = False
            chart.legend.font.name = chart_font_name
            chart.legend.font.size = Pt(base_size)
            chart.legend.font.color.rgb = text_color
    except Exception:  # noqa: BLE001
        pass

    # Оси: подписи контрастные, сетка приглушённая
    for axis_name in ("category_axis", "value_axis"):
        try:
            axis = getattr(chart, axis_name)
        except Exception:  # noqa: BLE001
            continue
        try:
            axis.tick_labels.font.name = chart_font_name
            axis.tick_labels.font.size = Pt(base_size)
            axis.tick_labels.font.color.rgb = text_color
            axis.format.line.color.rgb = grid_color
            if axis.has_major_gridlines:
                axis.major_gridlines.format.line.color.rgb = grid_color
        except Exception:  # noqa: BLE001
            pass

    # Подписи данных: значения на узлах/столбцах, жирные, чуть крупнее легенды
    try:
        from pptx.enum.chart import XL_LABEL_POSITION

        plot.has_data_labels = True
        labels = plot.data_labels
        labels.show_value = True
        labels.font.name = chart_font_name
        labels.font.size = Pt(base_size + 2)
        labels.font.bold = True
        labels.font.color.rgb = text_color
        labels.number_format = "#,##0.##"
        labels.number_format_is_linked = False
        chart_type = str(getattr(chart, "chart_type", ""))
        if "LINE" in chart_type:
            labels.position = XL_LABEL_POSITION.ABOVE
        elif "PIE" in chart_type or "DOUGHNUT" in chart_type:
            labels.position = XL_LABEL_POSITION.OUTSIDE_END
        elif "COLUMN" in chart_type or "BAR" in chart_type:
            labels.position = XL_LABEL_POSITION.OUTSIDE_END
    except Exception:  # noqa: BLE001
        pass
    # Линейный график: точки-маркеры на узлах
    try:
        if "LINE" in str(getattr(chart, "chart_type", "")):
            from pptx.enum.chart import XL_MARKER_STYLE

            for series_idx, series in enumerate(plot.series):
                series.smooth = False
                series.marker.style = XL_MARKER_STYLE.CIRCLE
                series.marker.size = 7
                series.marker.format.fill.solid()
                series.marker.format.fill.fore_color.rgb = colors[series_idx % len(colors)]
                series.marker.format.line.color.rgb = colors[series_idx % len(colors)]
    except Exception:  # noqa: BLE001
        pass


def style_title_textbox(text_frame, manifest: DesignManifest, box_width_emu: int | None = None, box_height_emu: int | None = None) -> None:
    """Красит текстовый фрейм заголовка в стиль темы (typography.title).

    Если переданы размеры бокса — заранее уменьшает кегль так, чтобы обёрнутый
    текст гарантированно поместился по высоте (см. _fit_font_size_pt). Это
    страхует от переполнения в рендерерах (LibreOffice/аудит-превью), которые
    не всегда пересчитывают <a:normAutofit> динамически, как это делает сам
    PowerPoint при открытии файла.
    """
    style = manifest.typography.title
    color = _hex_to_rgbcolor(manifest.palette.get("dk1"), "000000")
    size_pt = style.size
    if box_width_emu and box_height_emu:
        text = text_frame.text
        size_pt = _fit_font_size_pt(text, style.size, box_width_emu, box_height_emu, bold=style.bold)
    for paragraph in text_frame.paragraphs:
        _apply_run_style(paragraph, style.font, size_pt, style.bold, color)
    _enable_shrink_to_fit(text_frame)


def style_body_textbox(text_frame, manifest: DesignManifest, bold: bool | None = None, box_width_emu: int | None = None, box_height_emu: int | None = None) -> None:
    """Красит текстовый фрейм тела (буллеты/текст) в стиль темы (typography.body)."""
    style = manifest.typography.body
    color = _hex_to_rgbcolor(manifest.palette.get("dk1"), "000000")
    size_pt = style.size
    if box_width_emu and box_height_emu:
        text = "\n".join(p.text for p in text_frame.paragraphs)
        size_pt = _fit_font_size_pt(text, style.size, box_width_emu, box_height_emu, bold=bool(bold))
    for paragraph in text_frame.paragraphs:
        _apply_run_style(paragraph, style.font, size_pt, bold if bold is not None else style.bold, color)
    _enable_shrink_to_fit(text_frame)


def _enable_shrink_to_fit(text_frame) -> None:
    """Включает <a:normAutofit> вместо <a:spAutoFit> — PowerPoint будет сам
    уменьшать кегль, если текст всё равно не влезет (например, после ручного
    редактирования), вместо того чтобы расширять рамку и наезжать на соседние
    зоны."""
    from pptx.oxml.ns import qn

    body_pr = text_frame._txBody.find(qn("a:bodyPr"))
    if body_pr is None:
        return
    for tag in ("a:spAutoFit", "a:noAutofit", "a:normAutofit"):
        existing = body_pr.find(qn(tag))
        if existing is not None:
            body_pr.remove(existing)
    from pptx.oxml import parse_xml

    norm_autofit = parse_xml(
        '<a:normAutofit xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>'
    )
    body_pr.append(norm_autofit)


def _fit_font_size_pt(text: str, base_size_pt: int, box_width_emu: int, box_height_emu: int, bold: bool = False, min_size_pt: int = 12) -> int:
    """Оценивает, сколько строк займёт `text` при переносе по ширине бокса на
    кегле `base_size_pt`, и если по высоте не влезает — пропорционально
    уменьшает кегль (не ниже `min_size_pt`).

    Оценка ширины символа — грубая эвристика (Ariel/Calibri-подобные шрифты:
    средний символ ~0.52 от кегля в ширину), этого достаточно, чтобы поймать
    явные переполнения в 2-3 строки, не считая точную метрику каждого глифа.
    """
    if not text:
        return base_size_pt

    EMU_PER_PT = 12700
    box_width_pt = box_width_emu / EMU_PER_PT
    box_height_pt = box_height_emu / EMU_PER_PT
    avg_char_width_factor = 0.56 if bold else 0.52
    line_height_factor = 1.25

    def lines_needed(size_pt: float) -> int:
        chars_per_line = max(1, int(box_width_pt / (size_pt * avg_char_width_factor)))
        total_lines = 0
        for raw_line in text.split("\n"):
            raw_line = raw_line or " "
            words = raw_line.split(" ")
            cur_len = 0
            line_count = 1
            for word in words:
                w_len = len(word) + 1
                if cur_len + w_len > chars_per_line and cur_len > 0:
                    line_count += 1
                    cur_len = w_len
                else:
                    cur_len += w_len
            total_lines += line_count
        return total_lines

    size_pt = base_size_pt
    while size_pt > min_size_pt:
        n_lines = lines_needed(size_pt)
        needed_height_pt = n_lines * size_pt * line_height_factor
        if needed_height_pt <= box_height_pt:
            break
        size_pt -= 2
    return max(size_pt, min_size_pt)


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
