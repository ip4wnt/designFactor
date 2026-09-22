"""Движок свободной вёрстки: подбирает геометрический паттерн под контент
слайда и раскладывает content_blocks по его зонам.

Отвязан от plaeholder-ов конкретного шаблона (в отличие от старого пути в
assembly.py, который заполнял только plaeholder-ы существующего slide_layout).
Вместо этого паттерны — универсальная библиотека геометрических схем
(app/config/layout_patterns.yaml), применимая к любому шаблону; фирменный вид
даёт не сам макет шаблона, а токены стиля (палитра/типографика из
DesignManifest, см. styling.py) плюс сам факт, что слайд строится "с нуля"
внутри prs той же презентации (тема/шрифты темы наследуются автоматически).

Три варианта вёрстки (variant_a/b/c) — три разные СТРАТЕГИИ ВЫБОРА паттерна
для одного и того же content_plan, а не три разных рендерера:
  - variant_a ("сбалансированный"): из подходящих паттернов берёт с тегом
    balanced/universal при равенстве баллов, наименее "крайний" вариант.
  - variant_b ("плотный"): предпочитает паттерны density=compact — больше
    контент-зон на слайде, режим для насыщенных контент-планов.
  - variant_c ("просторный"): предпочитает density=airy — один крупный
    акцент на слайде, большие поля, крупная типографика.
"""
from __future__ import annotations

from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path

import yaml

from app.schemas.content_plan import ContentBlock, ContentBlockType, SlideSpec

_PATTERNS_PATH = Path(__file__).resolve().parent.parent / "config" / "layout_patterns.yaml"

_BLOCK_TO_ROLE = {
    ContentBlockType.CHART: "chart",
    ContentBlockType.TABLE: "table",
    ContentBlockType.BULLETS: "any",
    ContentBlockType.TEXT: "any",
}

# Роли зон, которые ОБЯЗАТЕЛЬНО должны совпасть по типу (нельзя положить
# буллеты в зону role=chart) — "any" принимает всё, chart/table принимают
# только свой тип ИЛИ any-блоки не претендуют на них первыми (см. _assign_zones).
_STRICT_ROLES = {"chart", "table", "image"}

_DENSITY_PRIORITY = {
    "variant_a": ("balanced", "compact", "airy"),
    "variant_b": ("compact", "balanced", "airy"),
    "variant_c": ("airy", "balanced", "compact"),
}


@dataclass(frozen=True)
class Zone:
    name: str
    role: str
    x: float
    y: float
    w: float
    h: float


@dataclass(frozen=True)
class Pattern:
    id: str
    description: str
    tags: tuple[str, ...]
    slot_count: int
    density: str
    zones: tuple[Zone, ...]


@dataclass
class ZoneAssignment:
    zone: Zone
    block: ContentBlock | None  # None для зоны title (текст берётся из slide_spec.title)


@lru_cache(maxsize=1)
def load_patterns() -> list[Pattern]:
    with open(_PATTERNS_PATH, encoding="utf-8") as f:
        raw = yaml.safe_load(f)
    patterns = []
    for entry in raw["patterns"]:
        zones = tuple(Zone(**z) for z in entry["zones"])
        patterns.append(
            Pattern(
                id=entry["id"],
                description=entry["description"],
                tags=tuple(entry.get("tags", [])),
                slot_count=entry["slot_count"],
                density=entry["density"],
                zones=zones,
            )
        )
    return patterns


def select_pattern(slide_spec: SlideSpec, variant: str) -> Pattern:
    """Выбирает паттерн, чьи content-зоны лучше всего покрывают content_blocks
    слайда — по количеству зон и по совпадению строгих ролей (chart/table),
    при равенстве баллов — по приоритету density для данного варианта.
    """
    patterns = load_patterns()
    blocks = slide_spec.content_blocks
    needed_slots = len(blocks)
    chart_needed = sum(1 for b in blocks if b.type == ContentBlockType.CHART)
    table_needed = sum(1 for b in blocks if b.type == ContentBlockType.TABLE)

    density_order = _DENSITY_PRIORITY.get(variant, _DENSITY_PRIORITY["variant_a"])

    def score(pattern: Pattern) -> tuple[float, int]:
        chart_slots = sum(1 for z in pattern.zones if z.role == "chart")
        table_slots = sum(1 for z in pattern.zones if z.role == "table")

        s = 0.0
        # Совпадение общего числа зон с числом блоков — самый важный сигнал:
        # не хотим ни терять контент (слотов меньше, чем блоков), ни оставлять
        # пустые декоративные зоны (слотов намного больше, чем блоков).
        s -= abs(pattern.slot_count - needed_slots) * 3.0
        # Достаточно строгих слотов под чарт/таблицу — важно, чтобы не пришлось
        # запихивать график в зону, где ему не хватит места на легенду/оси.
        s -= max(chart_needed - chart_slots, 0) * 4.0
        s -= max(table_needed - table_slots, 0) * 4.0
        # Небольшой бонус, если строгих слотов не намного больше, чем нужно
        # (избыток слотов = пустая декоративная зона в аудите).
        s -= max(chart_slots - chart_needed, 0) * 0.5
        s -= max(table_slots - table_needed, 0) * 0.5

        density_rank = density_order.index(pattern.density) if pattern.density in density_order else 99
        return (s, -density_rank)

    best = max(patterns, key=score)
    return best


def assign_zones(pattern: Pattern, slide_spec: SlideSpec) -> list[ZoneAssignment]:
    """Раскладывает content_blocks слайда по зонам паттерна.

    chart/table-блоки идут первыми в зоны с role совпадающей строго; остаток
    (bullets/text) идёт по возрастанию имени зоны в оставшиеся any-зоны.
    Если блоков больше, чем зон — лишние блоки отбрасываются (сигнал для
    аудита: content_plan не подходит под выбранный паттерн, такое возможно
    только при существенном рассинхроне и должно быть очень редким при
    корректном select_pattern).
    """
    blocks = list(slide_spec.content_blocks)
    zones = [z for z in pattern.zones if z.role != "title"]
    assignments: list[ZoneAssignment] = []

    remaining_blocks = list(blocks)
    remaining_zones = list(zones)

    # 1. Строгие зоны (chart/table) — забирают блок своего типа первым.
    for role, block_type in (("chart", ContentBlockType.CHART), ("table", ContentBlockType.TABLE)):
        strict_zones = [z for z in remaining_zones if z.role == role]
        for zone in strict_zones:
            match_idx = next((i for i, b in enumerate(remaining_blocks) if b.type == block_type), None)
            if match_idx is not None:
                assignments.append(ZoneAssignment(zone=zone, block=remaining_blocks.pop(match_idx)))
                remaining_zones.remove(zone)

    # 2. Оставшиеся any-зоны (и строгие зоны, для которых не нашлось точного
    # типа блока — на практике select_pattern такое избегает) — по порядку.
    any_zones = sorted(remaining_zones, key=lambda z: z.name)
    for zone in any_zones:
        if not remaining_blocks:
            break
        assignments.append(ZoneAssignment(zone=zone, block=remaining_blocks.pop(0)))

    return assignments


def zone_bbox_emu(zone: Zone, slide_width_emu: int, slide_height_emu: int) -> tuple[int, int, int, int]:
    """Переводит зону паттерна (доли слайда) в абсолютные EMU-координаты."""
    left = int(zone.x * slide_width_emu)
    top = int(zone.y * slide_height_emu)
    width = int(zone.w * slide_width_emu)
    height = int(zone.h * slide_height_emu)
    return left, top, width, height


def title_zone(pattern: Pattern) -> Zone | None:
    return next((z for z in pattern.zones if z.role == "title"), None)
