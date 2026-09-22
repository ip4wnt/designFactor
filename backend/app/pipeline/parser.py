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
    DesignManifest,
    Layout,
    LayoutRole,
    LayoutRoleType,
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
    """Синхронный детерминированный разбор .pptx в DesignManifest."""
    prs = Presentation(str(pptx_path))
    master = prs.slide_masters[0]

    palette = _extract_palette(master)
    typography = _extract_typography(master)
    layouts = [_extract_layout(layout) for layout in master.slide_layouts]

    return DesignManifest(
        template_id=str(uuid.uuid4()),
        palette=palette,
        typography=typography,
        layouts=layouts,
        slide_width_emu=int(Emu(prs.slide_width)),
        slide_height_emu=int(Emu(prs.slide_height)),
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
    """Читает дефолтные стили title/body из p:txStyles мастера слайдов."""
    tx_styles = master.element.find(qn("p:txStyles"))
    title_style = _read_style_level(tx_styles, "p:titleStyle") if tx_styles is not None else None
    body_style = _read_style_level(tx_styles, "p:bodyStyle") if tx_styles is not None else None

    return Typography(
        title=title_style or TypographyStyle(font="Calibri", size=44, bold=True),
        body=body_style or TypographyStyle(font="Calibri", size=18, bold=False),
    )


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
