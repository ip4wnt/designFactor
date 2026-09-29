"""Парсер .pptx-шаблона в DesignManifest.

Детерминированный путь (parse_template) достаёт из XML только факты:
палитру темы (a:clrScheme), типографику title/body из p:txStyles мастера,
и список макетов с ролями плейсхолдеров. Никаких моделей здесь не вызывается.

classify_layouts_with_vlm — отдельная async-заглушка под будущую
классификацию макетов через VLM (например, по скриншоту layout'а), она
не блокирует и не подменяет детерминированный путь.
"""
from __future__ import annotations

import uuid
from pathlib import Path

from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from pptx.util import Emu

from app.schemas.design_manifest import (
    ChartStyleTokens,
    DesignManifest,
    Layout,
    LayoutExample,
    LayoutRole,
    LayoutRoleType,
    SlideElement,
    SlideElementType,
    TableStyleTokens,
    Typography,
    TypographyStyle,
)

# PP_PLACEHOLDER.OBJECT — универсальный контент-плейсхолдер PowerPoint: в него
# можно вставить текст, таблицу, диаграмму или картинку, поэтому он размечен
# сразу и как "body"-роль, и во все supports_* флаги ниже.
_TITLE_TYPES = {"TITLE", "CENTER_TITLE"}
_BODY_TYPES = {"BODY", "SUBTITLE", "OBJECT"}
_CHART_TYPES = {"CHART", "OBJECT"}
_TABLE_TYPES = {"TABLE", "OBJECT"}
_PICTURE_TYPES = {"PICTURE", "OBJECT"}

_THEME_COLOR_TAGS = (
    "dk1", "lt1", "dk2", "lt2",
    "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "hlink", "folHlink",
)


def parse_template(pptx_path: str | Path) -> DesignManifest:
    """Синхронный детерминированный разбор .pptx в DesignManifest.

    Шаблон PowerPoint может содержать НЕСКОЛЬКО slide master (обычно так
    бывает, когда шаблон собирали из нескольких Google Slides/PowerPoint
    тем) — каждый со своим набором макетов и потенциально своей темой.
    Раньше здесь бралось только prs.slide_masters[0], из-за чего терялись
    все макеты остальных мастеров (частый случай: 3 макета в первом
    мастере против 36 во втором). Теперь обходятся все мастера, палитра и
    типография берутся из первого мастера с непустой темой (на практике
    у всех мастеров одного шаблона палитра почти всегда идентична).
    """
    prs = Presentation(str(pptx_path))
    masters = list(prs.slide_masters)

    palette: dict[str, str] = {}
    typography: Typography | None = None
    layouts: list[Layout] = []
    # part.partname (уникальный путь в .pptx, например /ppt/slideLayouts/slideLayout23.xml)
    # — надёжный ключ для матчинга slide -> layout_id ниже. Имя макета
    # (source_layout_name) не годится: одинаковые имена повторяются между разными
    # мастерами (например два разных layout с именем "1_Титульный слайд").
    layout_id_by_partname: dict[str, str] = {}
    for master in masters:
        if not palette:
            palette = _extract_palette(master)
        if typography is None:
            typography = _extract_typography(master)
        for layout in master.slide_layouts:
            extracted = _extract_layout(layout)
            layouts.append(extracted)
            layout_id_by_partname[str(layout.part.partname)] = extracted.layout_id

    if typography is None:
        typography = Typography(
            title=TypographyStyle(font="Calibri", size=44, bold=True),
            body=TypographyStyle(font="Calibri", size=18, bold=False),
        )
    typography = _refine_typography_from_layouts(typography, masters)

    # Слайды-образцы шаблона могут уже содержать готовые таблицы/графики —
    # если так есть, это более точный источник фирменного стиля, чем вывод из
    # palette/typography вслепую (автор шаблона мог сознательно выбрать другой accent
    # для таблиц, отличный от accent1, или свой порядок цветов для серий графика).
    table_style, chart_style = _extract_sample_styles(prs)

    # Слайды-образцы — это не только источник стиля таблиц/графиков, но и
    # готовые примеры того, как автор шаблона реально скомпоновал контент на
    # каждом макете (композиция, реальный текст, цвета фигур) — то, что
    # никогда не считывается с пустой заготовки layout. Группируем по
    # layout_id макета, на который ссылается каждый слайд (по partname, см. выше).
    layout_examples = _extract_layout_examples(prs, layout_id_by_partname)

    return DesignManifest(
        template_id=str(uuid.uuid4()),
        palette=palette,
        typography=typography,
        layouts=layouts,
        slide_width_emu=int(Emu(prs.slide_width)),
        slide_height_emu=int(Emu(prs.slide_height)),
        table_style=table_style,
        chart_style=chart_style,
        layout_examples=layout_examples,
    )


def _extract_sample_styles(prs: Presentation) -> tuple[TableStyleTokens | None, ChartStyleTokens | None]:
    """Обходит все слайды-образцы шаблона в поисках первой реальной таблицы
    и первого реального графика — вызывается до удаления этих слайдов в assembly.py,
    так что видит точно то, что загрузил пользователь в шаблоне.
    """
    table_style: TableStyleTokens | None = None
    chart_style: ChartStyleTokens | None = None

    for slide in prs.slides:
        for shape in slide.shapes:
            if table_style is None and shape.has_table:
                try:
                    table_style = _extract_table_style(shape.table)
                except Exception:  # noqa: BLE001 — повреждённая/нестандартная таблица не должна ронять весь анализ
                    table_style = None
            if chart_style is None and shape.has_chart:
                try:
                    chart_style = _extract_chart_style(shape.chart)
                except Exception:  # noqa: BLE001 — python-pptx падает на некоторых типах встроенных/связанных графиков (chart_part — обычный Part, не ChartPart)
                    chart_style = None
        if table_style is not None and chart_style is not None:
            break

    return table_style, chart_style


_MAX_ELEMENT_TEXT_LEN = 200  # обрезаем длинный текст фигуры для компактности манифеста


def _extract_layout_examples(
    prs: Presentation, layout_id_by_partname: dict[str, str]
) -> dict[str, list[LayoutExample]]:
    """Обходит слайды-образцы шаблона и группирует их по layout_id макета,
    на который они ссылаются (через partname макета — см. parse_template).

    Слайд без сопоставимого layout_id (крайне редкий случай повреждённого
    файла) тихо пропускается — это не должно ронять весь анализ шаблона.
    """
    examples_by_layout: dict[str, list[LayoutExample]] = {}

    for slide_index, slide in enumerate(prs.slides, start=1):
        slide_layout = slide.slide_layout
        if slide_layout is None:
            continue
        partname = str(slide_layout.part.partname)
        layout_id = layout_id_by_partname.get(partname)
        if layout_id is None:
            continue

        elements = [_extract_slide_element(shape) for shape in slide.shapes]
        elements = [el for el in elements if el is not None]
        if not elements:
            continue

        example = LayoutExample(
            slide_index=slide_index,
            slide_name=_slide_title_text(slide),
            elements=elements,
        )
        examples_by_layout.setdefault(layout_id, []).append(example)

    return examples_by_layout


def _slide_title_text(slide) -> str | None:
    """Текст тайтл-плейсхолдера слайда, если он есть и заполнен — для
    подписи под мини-макетиком в UI, чтобы было видно, какой именно образец
    это был."""
    try:
        title_shape = slide.shapes.title
    except (AttributeError, ValueError):
        return None
    if title_shape is None or not getattr(title_shape, "has_text_frame", False):
        return None
    text = title_shape.text_frame.text.strip()
    return text[:80] if text else None


def _extract_slide_element(shape) -> SlideElement | None:
    """Строит SlideElement из фигуры реального слайда — минимум для HTML-
    отрисовки макетика (позиция/размер/тип/текст/цвета), без рендера.

    Фигуры без геометрии (group-only контейнеры без top/left в некоторых
    крайних случаях python-pptx) пропускаются — None.
    """
    try:
        x, y, w, h = shape.left, shape.top, shape.width, shape.height
    except AttributeError:
        return None
    if x is None or y is None or w is None or h is None:
        return None

    element_type = SlideElementType.OTHER
    text: str | None = None
    fill_hex: str | None = None
    text_color_hex: str | None = None
    font_size: int | None = None
    bold: bool | None = None
    is_title_role = False

    if getattr(shape, "has_table", False):
        element_type = SlideElementType.TABLE
    elif getattr(shape, "has_chart", False):
        element_type = SlideElementType.CHART
    elif shape.shape_type is not None and shape.shape_type.name in ("PICTURE",):
        element_type = SlideElementType.PICTURE
    elif getattr(shape, "has_text_frame", False) and shape.text_frame.text.strip():
        element_type = SlideElementType.TEXT
        text = shape.text_frame.text.strip()[:_MAX_ELEMENT_TEXT_LEN]
        text_color_hex, bold, _font_name, font_size = _first_run_style(shape.text_frame)
        try:
            is_title_role = (
                shape.placeholder_format is not None
                and shape.placeholder_format.type is not None
                and shape.placeholder_format.type.name in _TITLE_TYPES
            )
        except (AttributeError, ValueError):
            is_title_role = False

    fill_hex = _shape_fill_hex(shape)
    horizontal_align, vertical_align = _text_alignment(shape) if element_type == SlideElementType.TEXT else (None, None)
    margins = _text_margins_emu(shape) if element_type == SlideElementType.TEXT else None
    picture_description = _picture_description(shape) if element_type == SlideElementType.PICTURE else None

    return SlideElement(
        element_type=element_type,
        x_emu=int(x),
        y_emu=int(y),
        width_emu=int(w),
        height_emu=int(h),
        text=text,
        fill_hex=fill_hex,
        text_color_hex=text_color_hex,
        font_size=font_size,
        bold=bold,
        is_title_role=is_title_role,
        horizontal_align=horizontal_align,
        vertical_align=vertical_align,
        margin_left_emu=margins[0] if margins else None,
        margin_right_emu=margins[1] if margins else None,
        margin_top_emu=margins[2] if margins else None,
        margin_bottom_emu=margins[3] if margins else None,
        picture_description=picture_description,
    )


_HORIZONTAL_ALIGN_MAP = {
    "l": "left",
    "ctr": "center",
    "r": "right",
    "just": "justify",
    "dist": "distributed",
    "thaiDist": "thai_distributed",
}


def _text_alignment(shape) -> tuple[str | None, str | None]:
    """Горизонтальное выравнивание — из явного algn первого абзаца с текстом,
    с фолбэком на python-pptx paragraph.alignment. Вертикальное — из vertical_anchor
    текстовой рамки целиком (один на всю фигуру, не на абзац).
    Паттерн взят из доработки коллеги (pptx_template_parser/layout/shapes.py).
    """
    horizontal: str | None = None
    vertical: str | None = None
    try:
        text_frame = shape.text_frame
    except (AttributeError, ValueError):
        return None, None

    for paragraph in text_frame.paragraphs:
        p_pr = paragraph._pPr
        algn = p_pr.get("algn") if p_pr is not None else None
        if algn:
            horizontal = _HORIZONTAL_ALIGN_MAP.get(algn, algn)
            break
        if paragraph.alignment is not None:
            horizontal = str(paragraph.alignment).split(".")[-1].lower()
            break

    try:
        anchor = text_frame.vertical_anchor
        if anchor is not None:
            vertical = str(anchor).split(".")[-1].lower()
    except (AttributeError, ValueError):
        pass

    return horizontal, vertical


def _text_margins_emu(shape) -> tuple[int, int, int, int] | None:
    try:
        tf = shape.text_frame
        return (
            int(tf.margin_left) if tf.margin_left is not None else 0,
            int(tf.margin_right) if tf.margin_right is not None else 0,
            int(tf.margin_top) if tf.margin_top is not None else 0,
            int(tf.margin_bottom) if tf.margin_bottom is not None else 0,
        )
    except (AttributeError, ValueError):
        return None


def _picture_description(shape) -> str | None:
    """alt-текст картинки из cNvPr/@descr — автор шаблона иногда описывает
    там назначение изображения (например, «placeholder для фото спикера») —
    полезный контекст для LLM-контент-агента, когда он решает, что подставить в этот слот.
    """
    try:
        c_nv_pr = shape._element.find(".//" + qn("p:cNvPr"))
        if c_nv_pr is None:
            return None
        descr = c_nv_pr.get("descr")
        return descr.strip() if descr and descr.strip() else None
    except (AttributeError, ValueError):
        return None


def _shape_fill_hex(shape) -> str | None:
    try:
        fill = shape.fill
        if fill.type is None:
            return None
        return f"#{fill.fore_color.rgb}"
    except (AttributeError, TypeError, KeyError, ValueError):
        return None


def _first_run_style(text_frame) -> tuple[str | None, bool | None, str | None, int | None]:
    """Читает цвет/жирность/шрифт/кегль первого run с непустым текстом —
    аналог _cell_text_style, но для обычной (не табличной) текстовой рамки."""
    for paragraph in text_frame.paragraphs:
        for run in paragraph.runs:
            if not run.text.strip():
                continue
            color = None
            try:
                if run.font.color.type is not None:
                    color = f"#{run.font.color.rgb}"
            except (AttributeError, TypeError, KeyError, ValueError):
                pass
            bold = run.font.bold
            font_name = run.font.name
            size = int(run.font.size.pt) if run.font.size is not None else None
            return color, bold, font_name, size
    return None, None, None, None


def _extract_table_style(table) -> TableStyleTokens | None:
    if len(table.rows) < 1:
        return None

    header_cell = table.cell(0, 0)
    header_fill = _cell_fill_hex(header_cell)
    header_text_color, header_bold, body_font, body_size = _cell_text_style(header_cell)

    row_odd_fill = None
    row_even_fill = None
    if len(table.rows) > 1:
        row_odd_fill = _cell_fill_hex(table.cell(1, 0))
    if len(table.rows) > 2:
        row_even_fill = _cell_fill_hex(table.cell(2, 0))

    return TableStyleTokens(
        header_fill=header_fill,
        header_text_color=header_text_color,
        header_bold=header_bold,
        row_odd_fill=row_odd_fill,
        row_even_fill=row_even_fill,
        body_font=body_font,
        body_size=body_size,
    )


def _cell_fill_hex(cell) -> str | None:
    try:
        if cell.fill.type is None:
            return None
        return f"#{cell.fill.fore_color.rgb}"
    except (AttributeError, TypeError, KeyError):
        return None


def _cell_text_style(cell) -> tuple[str | None, bool | None, str | None, int | None]:
    """Читает цвет/жирность/шрифт/кегль первого run с текстом в ячейке."""
    for paragraph in cell.text_frame.paragraphs:
        for run in paragraph.runs:
            color = None
            try:
                if run.font.color.type is not None:
                    color = f"#{run.font.color.rgb}"
            except (AttributeError, TypeError, KeyError):
                pass
            bold = run.font.bold
            font_name = run.font.name
            size = int(run.font.size.pt) if run.font.size is not None else None
            return color, bold, font_name, size
    return None, None, None, None


def _extract_chart_style(chart) -> ChartStyleTokens | None:
    series_colors: list[str] = []
    try:
        plot = chart.plots[0]
        for series in plot.series:
            hex_color = None
            try:
                if series.format.fill.type is not None:
                    hex_color = f"#{series.format.fill.fore_color.rgb}"
            except (AttributeError, TypeError, KeyError):
                pass
            if hex_color is None:
                try:
                    hex_color = f"#{series.format.line.color.rgb}"
                except (AttributeError, TypeError, KeyError):
                    pass
            if hex_color is not None:
                series_colors.append(hex_color)
    except (AttributeError, IndexError):
        pass

    font = None
    font_size = None
    try:
        font = chart.font.name
        font_size = int(chart.font.size.pt) if chart.font.size is not None else None
    except AttributeError:
        pass

    has_legend = None
    try:
        has_legend = bool(chart.has_legend)
    except AttributeError:
        pass

    if not series_colors and font is None and has_legend is None:
        return None

    return ChartStyleTokens(
        series_colors=series_colors,
        font=font,
        font_size=font_size,
        has_legend=has_legend,
    )


def _extract_palette(master) -> dict[str, str]:
    """Читает a:clrScheme прямо из XML части темы, связанной с мастером.

    ThemePart в python-pptx — это обычный opc-Part без объектной модели
    (в отличие от SlideMasterPart), поэтому его XML разбирается вручную
    через pptx.oxml.parse_xml(blob), а не через атрибут .element.
    """
    theme_part = master.part.part_related_by(RT.THEME)
    theme_element = parse_xml(theme_part.blob)
    clr_scheme = theme_element.find(qn("a:themeElements") + "/" + qn("a:clrScheme"))

    palette: dict[str, str] = {}
    if clr_scheme is None:
        return palette

    for tag in _THEME_COLOR_TAGS:
        color_node = clr_scheme.find(qn(f"a:{tag}"))
        if color_node is None:
            continue
        palette[tag] = _read_color(color_node)
    return palette


def _read_color(color_node) -> str:
    """Читает srgbClr или sysClr внутри цветового узла темы."""
    srgb = color_node.find(qn("a:srgbClr"))
    if srgb is not None:
        return f"#{srgb.get('val')}"
    sys_clr = color_node.find(qn("a:sysClr"))
    if sys_clr is not None:
        return f"#{sys_clr.get('lastClr', sys_clr.get('val', '000000'))}"
    return "#000000"


def _extract_typography(master) -> Typography:
    """Читает дефолтные стили title/body из p:txStyles мастера слайдов.

    Это фоновый дефолт уровня мастера — почти всегда переопределяется прямым
    форматированием текста на конкретных макетах/слайдах (например,
    master может декларировать sz=1400, а реальный заголовок на layout
    набран 4800). Используется только как fallback, когда ни в одном
    макете этого мастера не нашлось явного размера — см.
    _refine_typography_from_layouts, вызываемую в parse_template после
    того, как собраны все layouts.
    """
    tx_styles = master.element.find(qn("p:txStyles"))
    title_style = _read_style_level(tx_styles, "p:titleStyle") if tx_styles is not None else None
    body_style = _read_style_level(tx_styles, "p:bodyStyle") if tx_styles is not None else None

    return Typography(
        title=title_style or TypographyStyle(font="Calibri", size=44, bold=True),
        body=body_style or TypographyStyle(font="Calibri", size=18, bold=False),
    )


def _refine_typography_from_layouts(typography: Typography, masters: list) -> Typography:
    """Уточняет typography реальными размерами, найденными прямо в
    плейсхолдерах макетов (title/body), а не дефолтом уровня мастера.

    Берёт первый найденный явный sz/жирность/шрифт для роли TITLE и
    первый — для роли BODY, обходя макеты всех мастеров по порядку.
    Если ни один макет не даёт явного значения для роли, остаётся
    исходный typography (дефолт мастера / глобальный fallback).
    """
    title_style: TypographyStyle | None = None
    body_style: TypographyStyle | None = None

    for master in masters:
        for layout in master.slide_layouts:
            for placeholder in layout.placeholders:
                if title_style is not None and body_style is not None:
                    break
                ph_type = placeholder.placeholder_format.type
                if ph_type is None:
                    continue
                type_name = ph_type.name
                role = _role_for_type(type_name)
                style = _read_placeholder_text_style(placeholder)
                if style is None:
                    continue
                if role == LayoutRoleType.TITLE and title_style is None:
                    title_style = style
                elif role == LayoutRoleType.BODY and body_style is None:
                    body_style = style

    return Typography(
        title=title_style or typography.title,
        body=body_style or typography.body,
    )


def _read_placeholder_text_style(placeholder) -> TypographyStyle | None:
    """Читает явный размер/жирность/шрифт заголовка/текста плейсхолдера
    на уровне САМОГО МАКЕТА (layout), а не дефолта мастера.

    Плейсхолдеры на layout (в отличие от плейсхолдеров реального слайда) почти
    никогда не содержат готовых runs с текстом — автор шаблона задаёт
    форматирование через <a:lstStyle><a:lvl1pPr><a:defRPr sz=".."/> внутри
    самого плейсхолдера. Именно оттуда и нужно брать реальный размер (например,
    layout может объявлять title sz=4800, пока master.txStyles говорит 1400).
    Если у плейсхолдера есть реальный run с явным sz — он приоритетнее lstStyle
    (более специфичное форматирование).
    """
    try:
        text_frame = placeholder.text_frame
    except (AttributeError, ValueError):
        return None

    for paragraph in text_frame.paragraphs:
        for run in paragraph.runs:
            size = run.font.size
            if size is None:
                continue
            font_name = run.font.name or "Calibri"
            bold = bool(run.font.bold)
            return TypographyStyle(font=font_name, size=int(size.pt), bold=bold)

    # Нет явного run — пробуем lstStyle/lvl1pPr/defRPr самого плейсхолдера.
    txbody = placeholder._element.find(qn("p:txBody"))
    if txbody is None:
        return None
    lst_style = txbody.find(qn("a:lstStyle"))
    if lst_style is None:
        return None
    lvl1 = lst_style.find(qn("a:lvl1pPr"))
    if lvl1 is None:
        return None
    def_rpr = lvl1.find(qn("a:defRPr"))
    if def_rpr is None:
        return None
    size_hundredths = def_rpr.get("sz")
    if size_hundredths is None:
        return None
    size_pt = int(size_hundredths) // 100
    bold = def_rpr.get("b") == "1"
    latin = def_rpr.find(qn("a:latin"))
    font_name = latin.get("typeface") if latin is not None and latin.get("typeface") else "Calibri"
    return TypographyStyle(font=font_name, size=size_pt, bold=bold)


def _read_style_level(tx_styles, style_tag: str) -> TypographyStyle | None:
    style_node = tx_styles.find(qn(style_tag))
    if style_node is None:
        return None
    lvl1 = style_node.find(qn("a:lvl1pPr"))
    if lvl1 is None:
        return None
    def_rpr = lvl1.find(qn("a:defRPr"))
    if def_rpr is None:
        return None

    size_hundredths = def_rpr.get("sz")
    size = int(size_hundredths) // 100 if size_hundredths else 18
    bold = def_rpr.get("b") == "1"

    latin = def_rpr.find(qn("a:latin"))
    font = latin.get("typeface") if latin is not None and latin.get("typeface") else "Calibri"

    return TypographyStyle(font=font, size=size, bold=bold)


def _extract_layout(layout) -> Layout:
    roles: list[LayoutRole] = []
    supports_chart = False
    supports_table = False
    supports_image = False

    for placeholder in layout.placeholders:
        ph_type = placeholder.placeholder_format.type
        if ph_type is None:
            continue
        type_name = ph_type.name
        idx = placeholder.placeholder_format.idx

        role = _role_for_type(type_name)
        if role is not None:
            roles.append(LayoutRole(role=role, placeholder_idx=idx))

        if type_name in _CHART_TYPES:
            supports_chart = True
        if type_name in _TABLE_TYPES:
            supports_table = True
        if type_name in _PICTURE_TYPES:
            supports_image = True

    return Layout(
        layout_id=str(uuid.uuid4()),
        source_layout_name=layout.name,
        roles=roles,
        supports_chart=supports_chart,
        supports_table=supports_table,
        supports_image=supports_image,
    )


def _role_for_type(type_name: str) -> LayoutRoleType | None:
    if type_name in _TITLE_TYPES:
        return LayoutRoleType.TITLE
    if type_name == "CHART":
        return LayoutRoleType.CHART
    if type_name == "TABLE":
        return LayoutRoleType.TABLE
    if type_name == "PICTURE":
        return LayoutRoleType.PICTURE
    if type_name in _BODY_TYPES:
        return LayoutRoleType.BODY
    return None


async def classify_layouts_with_vlm(pptx_path: str | Path, manifest: DesignManifest) -> DesignManifest:
    """TODO: классификация макетов через VLM (например, по скриншоту layout'а).

    Должна лишь дополнять уже построенный детерминированным parse_template()
    manifest (уточнять роли/стили), а не заменять его — вызывается
    Оркестратором опционально и не должна блокировать основной путь.
    """
    raise NotImplementedError("Классификация макетов через VLM ещё не реализована")
