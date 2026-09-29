"""Обогащение сырого DesignManifest (выход парсера) в «дизайн-систему» —
человекочитаемое представление с семантическими ролями, готовое для
визуализации в UI и для использования как явный контракт правил генерации.

DesignManifest — плоский набор фактов (dk1/accent1.../roles/placeholder_idx),
корректный, но не самообъясняющийся: непонятно без знания темы PowerPoint,
что dk1 — это обычно цвет текста, а accent1..6 — акценты, не глядя в
спецификацию OOXML. build_design_system() не меняет и не заменяет
DesignManifest (тот остаётся источником истины для сборки, см.
app/pipeline/styling.py) — это read-only проекция для отображения и ручной
правки поверх того же самого manifest.model_dump(), к которому вернутся при
сохранении правок обратно в DesignManifest.
"""
from __future__ import annotations

from pydantic import BaseModel, Field

from app.schemas.design_manifest import DesignManifest, LayoutRoleType, SlideElementType

# Роль каждого слота темы OOXML — по спецификации, а не по эвристике:
# dk1/dk2 — «тёмные» цвета (обычно текст), lt1/lt2 — «светлые» (обычно фон),
# accent1..6 — акцентные, hlink/folHlink — цвета гиперссылок.
_COLOR_ROLE_LABELS: dict[str, str] = {
    "dk1": "Основной текст",
    "dk2": "Вторичный тёмный",
    "lt1": "Фон (светлый)",
    "lt2": "Вторичный светлый / фон карточек",
    "accent1": "Акцент 1 — основной (CTA, заголовки-выделения)",
    "accent2": "Акцент 2",
    "accent3": "Акцент 3",
    "accent4": "Акцент 4",
    "accent5": "Акцент 5",
    "accent6": "Акцент 6",
    "hlink": "Гиперссылка",
    "folHlink": "Посещённая гиперссылка",
}

_COLOR_GROUP: dict[str, str] = {
    "dk1": "text",
    "dk2": "text",
    "lt1": "background",
    "lt2": "background",
    "accent1": "accent",
    "accent2": "accent",
    "accent3": "accent",
    "accent4": "accent",
    "accent5": "accent",
    "accent6": "accent",
    "hlink": "link",
    "folHlink": "link",
}

# Реальные названия шрифтов для служебных тем-плейсхолдеров OOXML —
# +mn-lt/+mj-lt ссылаются на minor/major latin font темы, а не на
# конкретное имя; без этой подсказки в UI будет непонятный символ.
_THEME_FONT_LABELS: dict[str, str] = {
    "+mn-lt": "Шрифт темы — основной (minor latin)",
    "+mj-lt": "Шрифт темы — заголовочный (major latin)",
    "+mn-ea": "Шрифт темы — основной (East Asian)",
    "+mj-ea": "Шрифт темы — заголовочный (East Asian)",
}

_ROLE_LABELS: dict[LayoutRoleType, str] = {
    LayoutRoleType.TITLE: "Заголовок",
    LayoutRoleType.BODY: "Текст / контент",
    LayoutRoleType.CHART: "График",
    LayoutRoleType.TABLE: "Таблица",
    LayoutRoleType.PICTURE: "Изображение",
}


class ColorToken(BaseModel):
    key: str
    hex: str
    label: str
    group: str  # text | background | accent | link


class FontToken(BaseModel):
    """Единый слот шрифтовой шкалы (заголовок/тело), с раскрытым
    настоящим именем шрифта вместо служебного +mn-lt/+mj-lt темы."""

    role: str  # title | body
    label: str
    raw_font: str
    resolved_font_label: str | None = None  # None если raw_font уже конкретное имя
    size_pt: int
    bold: bool


class LayoutRoleSummary(BaseModel):
    role: LayoutRoleType
    role_label: str
    placeholder_idx: int


class ExampleElementCard(BaseModel):
    """Один элемент внутри мини-макетика — процентные координаты/размер
    относительно габарита слайда — готовы для absolute-positioning в HTML без
    дополнительных вычислений на фронте."""

    element_type: SlideElementType
    left_pct: float
    top_pct: float
    width_pct: float
    height_pct: float
    text: str | None = None
    fill_hex: str | None = None
    text_color_hex: str | None = None
    font_size: int | None = None
    bold: bool | None = None
    is_title_role: bool = False
    horizontal_align: str | None = None
    vertical_align: str | None = None
    picture_description: str | None = None


class LayoutExampleCard(BaseModel):
    """Один реальный слайд-образец из презентации шаблона, отрисованный как
    мини-макет с реальным текстом и цветами этого конкретного слайда."""

    slide_index: int
    slide_name: str | None = None
    elements: list[ExampleElementCard]


class LayoutCard(BaseModel):
    layout_id: str
    name: str
    category: str  # cover | content | section | blank
    roles: list[LayoutRoleSummary]
    capabilities: list[str]  # человекочитаемый список: "Текст", "Таблица", ...
    aspect_ratio: float
    examples: list[LayoutExampleCard] = Field(default_factory=list)
    examples_count: int = 0


class TableStyleCard(BaseModel):
    found: bool  # False = в шаблоне не было ни одной реальной таблицы — сборщик выведет стиль из палитры сам
    header_fill: str | None = None
    header_text_color: str | None = None
    header_bold: bool | None = None
    row_odd_fill: str | None = None
    row_even_fill: str | None = None
    body_font_label: str | None = None
    body_size: int | None = None
    source: str  # "sample" (взято из реальной таблицы шаблона) | "not_found" (в шаблоне не найдено ни одной таблицы)


class ChartStyleCard(BaseModel):
    found: bool  # False = в шаблоне не было ни одного реального графика — сборщик выведет стиль из палитры сам
    series_colors: list[str] = Field(default_factory=list)
    font_label: str | None = None
    font_size: int | None = None
    has_legend: bool | None = None
    source: str  # "sample" | "not_found"


class TypeScaleToken(BaseModel):
    """Строка полной типографической шкалы шаблона: одна семантическая роль
    текста (заголовок слайда, тезис, фактоид, подпись…) с гарнитурой, всеми
    встретившимися кеглями и цветами и числом макетов, где она используется."""

    role: str
    label: str
    family: str
    sizes_pt: list[float]
    bold: bool
    estimated: bool  # кегль не задан явно в макете — оценка движка по высоте слота
    colors: list[str]
    layouts_count: int
    layouts: list[str]


class ThemeFonts(BaseModel):
    major: str | None = None  # шрифт заголовков темы (+mj-lt)
    minor: str | None = None  # шрифт текста темы (+mn-lt)
    families: list[str] = Field(default_factory=list)  # все гарнитуры, встречающиеся в макетах


class TemplateElementCard(BaseModel):
    """Элемент макета шаблона в процентах от размеров слайда — для миниатюр."""

    kind: str  # text | image | decor_image | decor
    role: str
    name: str | None = None
    left_pct: float
    top_pct: float
    width_pct: float
    height_pct: float
    font_size_pt: float | None = None
    color: str | None = None
    align: str | None = None


class TemplateLayoutCard(BaseModel):
    """Макет шаблона (slideLayout) из Template JSON движка: именно из этих
    макетов собираются слайды, поэтому карточка показывает роль в нарративе,
    композицию, число текстовых и картиночных слотов и их геометрию."""

    layout_id: str
    name: str
    family: str | None = None
    narrative_role: str | None = None
    narrative_label: str = ""
    layout_role: str | None = None
    composition: str | None = None
    content_pattern: str | None = None
    best_for: list[str] = Field(default_factory=list)
    text_slots: int
    image_slots: int
    background_hex: str | None = None
    elements: list[TemplateElementCard]


class CompositionCard(BaseModel):
    kind: str
    label: str
    source: str  # sample | layout | library
    count: int
    examples: list[str] = Field(default_factory=list)  # id композиций (до 4)


_DESIGN_SOURCE_LABELS = {
    "layouts": "макеты шаблона",
    "slides": "слайды-образцы",
    "mixed": "макеты и слайды-образцы",
    "library": "собственная библиотека компоновок",
}


def composition_cards(template_json: dict) -> list[CompositionCard]:
    """Сводка распознанных композиций по видам и источникам (для панели дизайн-системы)."""
    from app.pipeline import dramaturgy
    from app.pipeline.sample_slides import _KIND_LABELS

    groups: dict[tuple[str, str], list[str]] = {}
    try:
        comps = dramaturgy.catalog_compositions(template_json)
    except Exception:  # noqa: BLE001 — сводка не должна ломать панель
        return []
    for comp in comps:
        key = (comp.get("source") or "layout", comp.get("kind") or "text")
        groups.setdefault(key, []).append(comp["id"])
    order = {"sample": 0, "layout": 1, "library": 2}
    cards = [
        CompositionCard(kind=kind, label=_KIND_LABELS.get(kind, kind), source=source, count=len(ids), examples=ids[:4])
        for (source, kind), ids in groups.items()
    ]
    cards.sort(key=lambda c: (order.get(c.source, 9), -c.count, c.label))
    return cards


class DesignSystem(BaseModel):
    """Обогащённая, самообъясняющаяся проекция DesignManifest для UI.

    Не самостоятельный источник истины — редактирование происходит через
    правку исходного DesignManifest (PUT .../manifest), эта модель только
    для отображения. Поле `manifest` включено как есть, чтобы вкладка
    "JSON" в UI показывала именно то, что реально пойдёт в сборку.
    """

    template_analysis_id: str
    colors: list[ColorToken]
    fonts: list[FontToken]
    layouts: list[LayoutCard]
    table_style: TableStyleCard  # всегда возвращается, даже когда в шаблоне не найдено таблиц (found=False)
    chart_style: ChartStyleCard  # всегда возвращается, даже когда в шаблоне не найдено графиков (found=False)
    slide_width_emu: int
    slide_height_emu: int
    slide_aspect_ratio: float
    manifest: DesignManifest
    # Полная типографика и каталог макетов из Template JSON движка
    # pptx_template_parser (см. app/pipeline/template_engine.py). Пустые,
    # если Template JSON для шаблона не удалось построить.
    theme_fonts: ThemeFonts = Field(default_factory=ThemeFonts)
    typography_scale: list[TypeScaleToken] = Field(default_factory=list)
    template_layouts: list[TemplateLayoutCard] = Field(default_factory=list)
    template_layouts_count: int = 0
    # Откуда генератор берёт дизайн: layouts / slides / mixed / library
    # (см. sample_slides.detect_design_source) и что он распознал: сколько
    # слайдов-образцов разобрано на композиции и каких видов, какие
    # компоновки добавлены из собственной библиотеки.
    design_source: str | None = None
    design_source_info: dict = Field(default_factory=dict)
    compositions: list[CompositionCard] = Field(default_factory=list)
    template_source_mode: str | None = None  # layouts | slides


def _resolve_font_label(raw_font: str) -> str | None:
    return _THEME_FONT_LABELS.get(raw_font)


def _layout_category(name: str, roles: list[LayoutRoleSummary]) -> str:
    lowered = name.lower()
    if "обложка" in lowered or "titul" in lowered or "cover" in lowered:
        return "cover"
    if "раздел" in lowered or "section" in lowered:
        return "section"
    if not roles:
        return "blank"
    return "content"


def _capabilities(roles: list[LayoutRoleSummary], supports_chart: bool, supports_table: bool, supports_image: bool) -> list[str]:
    caps: list[str] = []
    role_types = {r.role for r in roles}
    if LayoutRoleType.TITLE in role_types:
        caps.append("Заголовок")
    if LayoutRoleType.BODY in role_types:
        caps.append("Текст")
    if supports_table:
        caps.append("Таблица")
    if supports_chart:
        caps.append("График")
    if supports_image:
        caps.append("Изображение")
    if not caps:
        caps.append("Пустой (декоративный/служебный)")
    return caps


def _build_example_card(example, slide_w: int, slide_h: int) -> LayoutExampleCard:
    """Переводит абсолютные EMU-координаты элемента в проценты от габарита
    слайда, чтобы фронт мог просто поставить left/top/width/height в % без знания
    EMU/аспекта слайда."""
    elements = [
        ExampleElementCard(
            element_type=el.element_type,
            left_pct=round(el.x_emu / slide_w * 100, 3),
            top_pct=round(el.y_emu / slide_h * 100, 3),
            width_pct=round(el.width_emu / slide_w * 100, 3),
            height_pct=round(el.height_emu / slide_h * 100, 3),
            text=el.text,
            fill_hex=el.fill_hex,
            text_color_hex=el.text_color_hex,
            font_size=el.font_size,
            bold=el.bold,
            is_title_role=el.is_title_role,
            horizontal_align=el.horizontal_align,
            vertical_align=el.vertical_align,
            picture_description=el.picture_description,
        )
        for el in example.elements
    ]
    return LayoutExampleCard(
        slide_index=example.slide_index,
        slide_name=example.slide_name,
        elements=elements,
    )


def build_design_system(
    template_analysis_id: str,
    manifest: DesignManifest,
    template_json: dict | None = None,
) -> DesignSystem:
    """Строит обогащённую проекцию манифеста. Чистая функция без побочных
    эффектов — вызывается и сразу после парсинга, и после каждой ручной
    правки манифеста, поэтому не должна ничего писать в стор сама.

    `template_json` — Template JSON движка: из него берутся полная
    типографическая шкала по ролям и каталог макетов с геометрией слотов."""

    colors = [
        ColorToken(
            key=key,
            hex=hex_value,
            label=_COLOR_ROLE_LABELS.get(key, key),
            group=_COLOR_GROUP.get(key, "accent"),
        )
        for key, hex_value in manifest.palette.items()
    ]

    fonts = [
        FontToken(
            role="title",
            label="Заголовок",
            raw_font=manifest.typography.title.font,
            resolved_font_label=_resolve_font_label(manifest.typography.title.font),
            size_pt=manifest.typography.title.size,
            bold=manifest.typography.title.bold,
        ),
        FontToken(
            role="body",
            label="Текст / контент",
            raw_font=manifest.typography.body.font,
            resolved_font_label=_resolve_font_label(manifest.typography.body.font),
            size_pt=manifest.typography.body.size,
            bold=manifest.typography.body.bold,
        ),
    ]

    layouts: list[LayoutCard] = []
    aspect = (manifest.slide_width_emu / manifest.slide_height_emu) if manifest.slide_height_emu else 16 / 9
    slide_w = manifest.slide_width_emu or 1
    slide_h = manifest.slide_height_emu or 1
    for layout in manifest.layouts:
        role_summaries = [
            LayoutRoleSummary(role=r.role, role_label=_ROLE_LABELS.get(r.role, r.role.value), placeholder_idx=r.placeholder_idx)
            for r in layout.roles
        ]
        raw_examples = manifest.layout_examples.get(layout.layout_id, [])
        example_cards = [_build_example_card(ex, slide_w, slide_h) for ex in raw_examples]
        layouts.append(
            LayoutCard(
                layout_id=layout.layout_id,
                name=layout.source_layout_name,
                category=_layout_category(layout.source_layout_name, role_summaries),
                roles=role_summaries,
                capabilities=_capabilities(role_summaries, layout.supports_chart, layout.supports_table, layout.supports_image),
                aspect_ratio=round(aspect, 4),
                examples=example_cards,
                examples_count=len(example_cards),
            )
        )

    if manifest.table_style is not None:
        ts = manifest.table_style
        table_style = TableStyleCard(
            found=True,
            header_fill=ts.header_fill,
            header_text_color=ts.header_text_color,
            header_bold=ts.header_bold,
            row_odd_fill=ts.row_odd_fill,
            row_even_fill=ts.row_even_fill,
            body_font_label=_resolve_font_label(ts.body_font) if ts.body_font else ts.body_font,
            body_size=ts.body_size,
            source="sample",
        )
    else:
        # В слайдах-образцах шаблона не найдено ни одной реальной таблицы —
        # явно сообщаем об этом в UI, а не молча опускаем поле.
        table_style = TableStyleCard(found=False, source="not_found")

    if manifest.chart_style is not None:
        cs = manifest.chart_style
        chart_style = ChartStyleCard(
            found=True,
            series_colors=cs.series_colors,
            font_label=_resolve_font_label(cs.font) if cs.font else cs.font,
            font_size=cs.font_size,
            has_legend=cs.has_legend,
            source="sample",
        )
    else:
        chart_style = ChartStyleCard(found=False, source="not_found")

    theme_fonts = ThemeFonts()
    typography_scale: list[TypeScaleToken] = []
    template_layouts: list[TemplateLayoutCard] = []
    source_mode: str | None = None
    if template_json:
        from app.pipeline import template_engine

        theme = template_json.get("theme_fonts") or {}
        theme_fonts = ThemeFonts(
            major=theme.get("major"),
            minor=theme.get("minor"),
            families=template_engine.font_families(template_json),
        )
        typography_scale = [TypeScaleToken(**row) for row in template_engine.typography_scale(template_json)]
        template_layouts = [
            TemplateLayoutCard(**{k: v for k, v in card.items() if k != "source_mode"})
            for card in template_engine.layout_cards(template_json)
        ]
        source_mode = template_json.get("source_mode")
        design_source = template_json.get("design_source")
        info = dict(template_json.get("design_source_info") or {})
        info.pop("samples", None)
        info["label"] = _DESIGN_SOURCE_LABELS.get(design_source or "", design_source or "")
        design_source_info = info
        compositions = composition_cards(template_json)
    else:
        design_source, design_source_info, compositions = None, {}, []

    return DesignSystem(
        template_analysis_id=template_analysis_id,
        colors=colors,
        fonts=fonts,
        layouts=layouts,
        table_style=table_style,
        chart_style=chart_style,
        slide_width_emu=manifest.slide_width_emu,
        slide_height_emu=manifest.slide_height_emu,
        slide_aspect_ratio=round(aspect, 4),
        manifest=manifest,
        theme_fonts=theme_fonts,
        typography_scale=typography_scale,
        template_layouts=template_layouts,
        template_layouts_count=len(template_layouts),
        template_source_mode=source_mode,
        design_source=design_source,
        design_source_info=design_source_info,
        compositions=compositions,
    )
