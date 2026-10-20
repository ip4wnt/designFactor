"""CSV, табличный JSON и markdown-таблицы в TXT/MD -> серверные таблицы (как листы XLSX).

Требование второй ступени: табличные данные приходят не только в Excel, но и в CSV, JSON
и внутри текстовых отчетов. Все они приводятся к одному представлению SourceTable, чтобы
анализ данных и план работали одинаково независимо от формата файла.
"""
from __future__ import annotations

import csv
import io
import json
import re
from uuid import uuid4

from app.pipeline.docx_tables import MAX_COLUMNS, MAX_DATA_ROWS, MAX_TABLES
from app.pipeline.materials import MaterialError, _decode
from app.schemas.content_plan import TableData
from app.schemas.source_tables import DocumentMaterial, SourceTable

MAX_ROWS = 100_000
MAX_ANALYSIS_COLUMNS = 100
TABULAR_SUFFIXES = {".csv", ".json", ".txt", ".md"}


def _cell(value) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, (dict, list)):
        return json.dumps(value, ensure_ascii=False)
    return str(value).strip()


def _finish(table: SourceTable, rows: list[list[str]]) -> SourceTable:
    """Те же правила пригодности, что у листов XLSX: заголовки, лимиты, размер для слайда."""
    rows = [row for row in rows if any(c.strip() for c in row)]
    if not rows:
        table.limitation = "Таблица пустая"
        return table
    width = max(len(r) for r in rows)
    if width > MAX_ANALYSIS_COLUMNS or len(rows) - 1 > MAX_ROWS:
        table.limitation = "Превышен лимит анализа: 100000 строк данных или 100 столбцов"
        return table
    rows = [(r + [""] * (width - len(r)))[:width] for r in rows]
    headers = rows[0]
    table.summary = table.summary + ". Колонки: " + "; ".join(headers)
    if len(rows) < 2 or any(not h.strip() for h in headers) or len(set(headers)) != len(headers):
        table.limitation = "Первая строка должна содержать непустые уникальные заголовки и хотя бы одну строку данных"
        return table
    table.data = TableData(headers=headers, rows=rows[1:])
    if len(rows) - 1 > MAX_DATA_ROWS or width > MAX_COLUMNS:
        table.limitation = f"Для слайда допустимо до {MAX_DATA_ROWS} строк данных и {MAX_COLUMNS} столбцов; таблица не обрезана"
    elif any(len(c) > 160 for row in rows for c in row) or sum(len(c) for row in rows for c in row) > 2000:
        table.limitation = "Слишком длинный текст в ячейках для одного слайда; таблица не сокращена"
    else:
        table.available = True
    return table


def _new_table(material_id: str, filename: str, ordinal: int, origin: str) -> SourceTable:
    return SourceTable(table_id=f"{material_id}.table_{ordinal}", source_name=filename,
                       ordinal=ordinal, sheet_name=origin, summary=origin, available=False)


# ---------- CSV ----------

def _csv_rows(text: str) -> list[list[str]]:
    lines = [ln for ln in text[:20_000].splitlines() if ln.strip()][:30]
    best, best_score = ",", (-1, 0)
    for candidate in (";", "\t", "|", ","):
        counts = [ln.count(candidate) for ln in lines]
        if not counts or max(counts) == 0:
            continue
        # Одинаковое число разделителей в каждой строке важнее общего количества:
        # так «12,5» с десятичной запятой не ломает CSV с «;».
        consistent = 1 if min(counts) == max(counts) else 0
        score = (consistent, sum(counts))
        if score > best_score:
            best, best_score = candidate, score
    delimiter = best
    reader = csv.reader(io.StringIO(text), delimiter=delimiter)
    return [[c.strip() for c in row] for row in reader]


def extract_csv(filename: str, content: bytes) -> DocumentMaterial:
    text = _decode(content)
    if not text.strip():
        raise MaterialError("CSV-файл пустой")
    material_id = "material_" + uuid4().hex
    table = _finish(_new_table(material_id, filename, 1, "Файл CSV"), _csv_rows(text))
    return DocumentMaterial(material_id=material_id, filename=filename,
                            text=f"Файл CSV\n[[{table.table_id}]]", tables=[table])


# ---------- JSON ----------

def _records_to_rows(records: list) -> list[list[str]] | None:
    if not records or not all(isinstance(r, dict) for r in records):
        return None
    headers: list[str] = []
    for r in records:
        for k in r:
            if k not in headers:
                headers.append(str(k))
    return [headers] + [[_cell(r.get(h)) for h in headers] for r in records]


def _json_tables(value, path: str, found: list[tuple[str, list[list[str]]]]) -> None:
    if len(found) >= MAX_TABLES:
        return
    if isinstance(value, list):
        rows = _records_to_rows(value)
        if rows is None and value and all(isinstance(r, list) for r in value):
            rows = [[_cell(c) for c in r] for r in value]
        if rows is not None:
            found.append((path or "корень", rows))
            return
        for i, item in enumerate(value):
            _json_tables(item, f"{path}[{i}]", found)
    elif isinstance(value, dict):
        columns = value.get("columns") or value.get("headers")
        data = value.get("rows") or value.get("data")
        if isinstance(columns, list) and isinstance(data, list) and data and all(isinstance(r, list) for r in data):
            found.append((path or "корень", [[_cell(c) for c in columns]] + [[_cell(c) for c in r] for r in data]))
            return
        for k, v in value.items():
            _json_tables(v, f"{path}.{k}" if path else str(k), found)


def extract_json(filename: str, content: bytes) -> DocumentMaterial:
    try:
        payload = json.loads(_decode(content))
    except ValueError as exc:
        raise MaterialError("Файл .json не разбирается") from exc
    found: list[tuple[str, list[list[str]]]] = []
    _json_tables(payload, "", found)
    if not found:
        raise MaterialError("В JSON не найдено табличных данных (список объектов или columns/rows)")
    material_id = "material_" + uuid4().hex
    tables, text = [], []
    for ordinal, (path, rows) in enumerate(found, 1):
        origin = f"JSON, узел «{path}»"
        table = _finish(_new_table(material_id, filename, ordinal, origin), rows)
        tables.append(table)
        text.extend([origin, f"[[{table.table_id}]]"])
    return DocumentMaterial(material_id=material_id, filename=filename, text="\n\n".join(text), tables=tables)


# ---------- markdown / ascii tables inside text ----------

_PIPE_ROW = re.compile(r"^\s*\|?.*\|.*\|?\s*$")
_SEPARATOR = re.compile(r"^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?\s*$")


def _split_pipe(line: str) -> list[str]:
    line = line.strip()
    if line.startswith("|"):
        line = line[1:]
    if line.endswith("|"):
        line = line[:-1]
    return [c.strip() for c in line.split("|")]


def extract_text_tables(filename: str, content: bytes) -> DocumentMaterial | None:
    """Возвращает материал с таблицами, если в тексте есть markdown-таблицы; иначе None."""
    text = _decode(content)
    lines = text.splitlines()
    material_id = "material_" + uuid4().hex
    tables: list[SourceTable] = []
    out: list[str] = []
    i = 0
    while i < len(lines):
        if _PIPE_ROW.match(lines[i]) and lines[i].count("|") >= 2:
            j = i
            block = []
            while j < len(lines) and _PIPE_ROW.match(lines[j]) and lines[j].count("|") >= 2:
                if not _SEPARATOR.match(lines[j]):
                    block.append(_split_pipe(lines[j]))
                j += 1
            if len(block) >= 2 and len(tables) < MAX_TABLES:
                caption = ""
                for back in range(len(out) - 1, -1, -1):
                    if out[back].strip():
                        caption = out[back].strip()[:120]
                        break
                origin = f"Таблица {len(tables) + 1} в тексте" + (f" («{caption}»)" if caption else "")
                table = _finish(_new_table(material_id, filename, len(tables) + 1, origin), block)
                tables.append(table)
                out.append(f"[[{table.table_id}]]")
                i = j
                continue
        out.append(lines[i])
        i += 1
    if not tables:
        return None
    return DocumentMaterial(material_id=material_id, filename=filename, text="\n".join(out), tables=tables)


def extract_tabular(filename: str, content: bytes, suffix: str) -> DocumentMaterial | None:
    if suffix == ".csv":
        return extract_csv(filename, content)
    if suffix == ".json":
        return extract_json(filename, content)
    return extract_text_tables(filename, content)
