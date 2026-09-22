"""Конвертация загруженного Excel-листа в ChartData/TableData ContentPlan'а.

Контекст: помимо брифа, который Контент-агент (LLM) превращает в ContentPlan
сам, пользователь может приложить Excel с реальными данными для конкретного
слайда. Этот модуль не решает, ГДЕ на слайде окажется эта таблица/график —
он только приводит лист Excel к тому же контракту (ChartData/TableData),
которым уже оперирует ассемблер (app/pipeline/assembly.py), чтобы дальше оба
источника (LLM и Excel) были неотличимы для остального пайплайна.

Конвенция определения формы данных на листе (без ML/эвристик по семантике):
- Первая строка листа — заголовки столбцов.
- Первый столбец — подписи строк (`row_label` / категория графика).
- Если имя листа начинается с "chart" (без учёта регистра) ИЛИ явно передан
  `kind="chart"` — лист превращается в ChartData: первый столбец → categories,
  остальные столбцы → серии (по имени заголовка), значения приводятся к float.
- Иначе (или явный `kind="table"`) — лист превращается в TableData: первая
  строка → headers, остальные строки → rows (значения как строки).
- Если ни kind не передан, ни имя листа не подсказывает — используется
  TableData как более безопасный (не роняющий данные при неверном типе)
  дефолт; ChartData требует, чтобы все "данные" столбцы были числовыми,
  иначе автоматически откатывается на TableData.
"""
from __future__ import annotations

import datetime
from pathlib import Path
from typing import Any, Literal

import openpyxl

from app.schemas.content_plan import ChartData, ChartType, TableData

MAX_ROWS = 500
MAX_COLS = 50


class ExcelImportError(Exception):
    """Лист Excel не удалось разобрать или он пуст."""


def _cell_to_str(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        return str(int(value)) if value == int(value) else str(round(value, 6))
    if isinstance(value, (datetime.datetime, datetime.date)):
        return value.isoformat()
    return str(value)


def _cell_to_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).replace(",", "."))
    except ValueError:
        return None


def read_sheet_rows(path: str | Path, sheet_name: str | None = None) -> tuple[str, list[list[Any]]]:
    """Читает сырые значения ячеек одного листа (по умолчанию — первый лист)."""
    workbook = openpyxl.load_workbook(str(path), data_only=True, read_only=True)
    worksheet = workbook[sheet_name] if sheet_name else workbook.worksheets[0]

    rows: list[list[Any]] = []
    for row in worksheet.iter_rows(max_row=MAX_ROWS, max_col=MAX_COLS):
        rows.append([cell.value for cell in row])

    while rows and all(v is None or v == "" for v in rows[-1]):
        rows.pop()

    max_col_used = 0
    for row in rows:
        for i in range(len(row) - 1, -1, -1):
            if row[i] not in (None, ""):
                max_col_used = max(max_col_used, i + 1)
                break
    rows = [row[:max_col_used] for row in rows]

    if not rows or max_col_used == 0:
        raise ExcelImportError(f"Лист «{worksheet.title}» пуст")

    return worksheet.title, rows


def list_sheet_names(path: str | Path) -> list[str]:
    workbook = openpyxl.load_workbook(str(path), data_only=True, read_only=True)
    return list(workbook.sheetnames)


def _looks_like_chart_sheet(sheet_name: str) -> bool:
    return sheet_name.strip().lower().startswith("chart")


def sheet_to_table_data(sheet_name: str, rows: list[list[Any]]) -> TableData:
    header_row, *data_rows = rows
    headers = [_cell_to_str(v) or f"col_{i + 1}" for i, v in enumerate(header_row)]
    table_rows = [[_cell_to_str(v) for v in row] + [""] * (len(headers) - len(row)) for row in data_rows]
    return TableData(headers=headers, rows=table_rows)


def sheet_to_chart_data(sheet_name: str, rows: list[list[Any]], chart_type: ChartType = ChartType.BAR) -> ChartData | None:
    """Возвращает None (вместо бросания исключения) если данные не числовые —
    вызывающая сторона тогда откатывается на TableData."""
    header_row, *data_rows = rows
    series_names = [_cell_to_str(v) or f"series_{i + 1}" for i, v in enumerate(header_row[1:], start=1)]

    categories: list[str] = []
    series_values: dict[str, list[float]] = {name: [] for name in series_names}

    for row in data_rows:
        category = _cell_to_str(row[0]) if row else ""
        categories.append(category)
        for col_idx, series_name in enumerate(series_names, start=1):
            raw = row[col_idx] if col_idx < len(row) else None
            value = _cell_to_float(raw)
            if value is None:
                return None  # непохоже на числовой ряд -> не график
            series_values[series_name].append(value)

    if not categories or not series_names:
        return None

    return ChartData(chart_type=chart_type, categories=categories, series=series_values)


def import_excel_block(
    path: str | Path,
    sheet_name: str | None = None,
    kind: Literal["table", "chart"] | None = None,
    chart_type: ChartType = ChartType.BAR,
) -> tuple[Literal["table", "chart"], TableData | ChartData]:
    """Точка входа: Excel-файл (+ опционально имя листа/явный kind) -> (kind, data).

    `kind=None` включает автоопределение: имя листа, начинающееся на "chart",
    или успешное построение ChartData (все значения кроме первого столбца
    числовые) приводит к графику; иначе — таблица.
    """
    resolved_name, rows = read_sheet_rows(path, sheet_name)

    if kind == "table":
        return "table", sheet_to_table_data(resolved_name, rows)
    if kind == "chart":
        chart = sheet_to_chart_data(resolved_name, rows, chart_type)
        if chart is None:
            raise ExcelImportError(
                f"Лист «{resolved_name}» запрошен как график, но содержит нечисловые значения"
            )
        return "chart", chart

    if _looks_like_chart_sheet(resolved_name):
        chart = sheet_to_chart_data(resolved_name, rows, chart_type)
        if chart is not None:
            return "chart", chart

    chart = sheet_to_chart_data(resolved_name, rows, chart_type)
    if chart is not None and _looks_like_chart_sheet(resolved_name):
        return "chart", chart

    return "table", sheet_to_table_data(resolved_name, rows)
