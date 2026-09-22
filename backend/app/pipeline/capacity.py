"""Проверка вместимости area под таблицу/график перед вставкой.

Не подгоняет размер шрифта и не масштабирует контент — только предупреждает
и урезает совсем нереалистичные случаи (10 строк текста в блоке высотой
2 см), чтобы не отдавать пользователю визуально сломанный слайд без единого
сигнала о проблеме. Полноценный auto-fit — за рамками этой задачи.
"""
from __future__ import annotations

from pptx.util import Emu

# Грубые пороги на основе типографики тела текста в pt, переведённые в EMU
# построчной высоты (line-height ~= 1.3 * font size). Достаточно точно для
# предупреждения о переполнении, не для точной вёрстки.
_EMU_PER_PT = 12700
# Совпадает с коэффициентом в styling.style_table (min_row_height) — обе оценки
# должны сходиться, иначе fit_table_data решит, что строки влезли,
# а style_table затем растянет их сильнее, чем реально было в area.
_LINE_HEIGHT_FACTOR = 1.6
_MIN_ROW_HEIGHT_EMU = Emu(int(0.6 * 360000))  # 0.6 см — нижняя граница читаемой строки таблицы
_MIN_COL_WIDTH_EMU = Emu(int(1.2 * 360000))  # 1.2 см — нижняя граница читаемого столбца


def table_capacity(height_emu: int, width_emu: int, font_size_pt: int) -> tuple[int, int]:
    """Сколько строк/столбцов помещается в area без визуального переполнения.

    Возвращает (max_rows, max_cols) — не то, сколько ЕСТЬ в данных, а то,
    сколько area физически вмещает при текущем шрифте темы.
    """
    line_height_emu = int(font_size_pt * _LINE_HEIGHT_FACTOR * _EMU_PER_PT)
    row_height = max(line_height_emu, _MIN_ROW_HEIGHT_EMU)
    max_rows = max(int(height_emu // row_height), 1)
    max_cols = max(int(width_emu // _MIN_COL_WIDTH_EMU), 1)
    return max_rows, max_cols


def fit_table_data(headers: list[str], rows: list[list[str]], height_emu: int, width_emu: int, font_size_pt: int) -> tuple[list[str], list[list[str]], bool]:
    """Урезает headers/rows так, чтобы таблица (шапка + строки) не переполняла area.

    Возвращает (headers, rows, truncated) — truncated=True сигнализирует, что
    исходные данные были урезаны и это стоит показать пользователю/в аудите.
    """
    max_rows, max_cols = table_capacity(height_emu, width_emu, font_size_pt)
    # Шапка занимает одну строку вместимости.
    max_data_rows = max(max_rows - 1, 1)

    truncated = False
    if len(headers) > max_cols:
        headers = headers[:max_cols]
        truncated = True
    if len(rows) > max_data_rows:
        rows = rows[:max_data_rows]
        truncated = True

    fitted_rows = []
    for row in rows:
        if len(row) > len(headers):
            fitted_rows.append(row[: len(headers)])
            truncated = True
        else:
            fitted_rows.append(row)

    return headers, fitted_rows, truncated


def fit_chart_data(categories: list[str], series: dict[str, list[float]], width_emu: int) -> tuple[list[str], dict[str, list[float]], bool]:
    """Урезает число категорий графика, если area слишком узкая для читаемых подписей.

    Грубая эвристика: категория требует ~1.1 см ширины для читаемой подписи
    на столбчатой/линейной диаграмме.
    """
    min_category_width_emu = Emu(int(1.1 * 360000))
    max_categories = max(int(width_emu // min_category_width_emu), 2)

    truncated = False
    if len(categories) > max_categories:
        categories = categories[:max_categories]
        truncated = True

    fitted_series = {}
    for name, values in series.items():
        if len(values) > len(categories):
            fitted_series[name] = values[: len(categories)]
            truncated = True
        else:
            fitted_series[name] = values

    return categories, fitted_series, truncated
