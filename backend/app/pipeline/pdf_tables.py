"""Text PDF -> ordered prose/table references and persistent, exact extracted cells."""
from __future__ import annotations

from io import BytesIO
from uuid import uuid4

import pdfplumber
from pdfplumber.utils.exceptions import PdfminerException
from pdfminer.pdfdocument import PDFPasswordIncorrect
from pdfminer.pdfparser import PDFSyntaxError

from app.pipeline.docx_tables import MAX_COLUMNS, MAX_DATA_ROWS, MAX_TABLES
from app.pipeline.materials import MaterialError
from app.schemas.content_plan import TableData
from app.schemas.source_tables import DocumentMaterial, SourceTable

MAX_PAGES = 100
MAX_ANALYSIS_ROWS = 100_000
MAX_ANALYSIS_COLUMNS = 100
MAX_PAGE_OBJECTS = 200_000


def _source_table(found, material_id, filename, ordinal, page_number):
    raw = found.extract(x_tolerance=3, y_tolerance=3)
    width = max((len(row) for row in raw), default=0)
    rows = [["" if cell is None else str(cell).strip() for cell in row] + [""] * (width - len(row))
            for row in raw]
    # Filled PDF rectangles can create a phantom gutter. Remove only wholly empty
    # detected columns, never empty cells or rows containing source values.
    columns = [i for i in range(width) if any(row[i] for row in rows)]
    rows = [[row[i] for i in columns] for row in rows]
    limitation = ""
    if len(rows) < 2 or not columns:
        limitation = "Нужны заголовки и хотя бы одна строка данных"
    elif any(not header for header in rows[0]) or len(set(rows[0])) != len(rows[0]):
        limitation = "Первая строка должна содержать непустые уникальные заголовки; составные заголовки пока не поддерживаются"
    elif len(rows) - 1 > MAX_ANALYSIS_ROWS or len(columns) > MAX_ANALYSIS_COLUMNS:
        limitation = "Превышен лимит анализа: 100000 строк данных или 100 столбцов"
    data = TableData(headers=rows[0], rows=rows[1:]) if not limitation else None
    if data is not None:
        if len(rows) - 1 > MAX_DATA_ROWS or len(columns) > MAX_COLUMNS:
            limitation = f"Для слайда допустимо до {MAX_DATA_ROWS} строк данных и {MAX_COLUMNS} столбцов; таблица не обрезана"
        elif any(len(cell) > 160 for row in rows for cell in row) or sum(len(cell) for row in rows for cell in row) > 2000:
            limitation = "Слишком длинный текст в ячейках для одного слайда; таблица не сокращена"
    return SourceTable(table_id=f"{material_id}.table_{ordinal}", source_name=filename, ordinal=ordinal,
                       page_number=page_number, bounding_box=list(found.bbox), data=data,
                       available=not limitation, limitation=limitation,
                       summary=("Колонки: " + "; ".join(rows[0]))[:600] if rows else "Пустая таблица")


def _page_text(page, tables):
    boxes = [table.bounding_box for table in tables]

    def outside_tables(obj):
        if obj.get("object_type") != "char":
            return True
        x = (obj["x0"] + obj["x1"]) / 2
        y = (obj["top"] + obj["bottom"]) / 2
        return not any(left <= x <= right and top <= y <= bottom for left, top, right, bottom in boxes)

    if not tables:
        return page.extract_text(x_tolerance=3, y_tolerance=3) or ""
    words = page.filter(outside_tables).extract_words(x_tolerance=3, y_tolerance=3)
    lines = []
    for word in sorted(words, key=lambda word: (word["top"], word["x0"])):
        if not lines or abs(word["top"] - lines[-1][0]) > 3:
            lines.append((word["top"], [word]))
        else:
            lines[-1][1].append(word)
    events = [(top, min(word["x0"] for word in line),
               " ".join(word["text"] for word in sorted(line, key=lambda word: word["x0"])))
              for top, line in lines]
    events.extend((table.bounding_box[1], table.bounding_box[0], f"[[{table.table_id}]]") for table in tables)
    return "\n\n".join(text for _, _, text in sorted(events))


def extract_pdf(filename: str, content: bytes) -> DocumentMaterial:
    """Only explicit grid boundaries define tables; never infer cells from prose."""
    material_id = "material_" + uuid4().hex
    tables, sections, unreadable_pages = [], [], []
    try:
        with pdfplumber.open(BytesIO(content)) as pdf:
            if len(pdf.pages) > MAX_PAGES:
                raise MaterialError(f"В PDF больше {MAX_PAGES} страниц; разделите документ")
            for original_page in pdf.pages:
                if len(original_page.chars) + len(original_page.edges) > MAX_PAGE_OBJECTS:
                    raise MaterialError(f"Страница {original_page.page_number} слишком сложная для извлечения таблиц")
                page = original_page.dedupe_chars()
                if not page.chars:
                    if page.images:
                        unreadable_pages.append(page.page_number)
                    original_page.close()
                    continue
                page_tables = []
                for found in sorted(page.find_tables(), key=lambda table: (table.bbox[1], table.bbox[0])):
                    # Single cells and rows are often decorative boxes, not data tables.
                    if len(found.rows) < 2 or len(found.columns) < 2:
                        continue
                    if len(tables) >= MAX_TABLES:
                        raise MaterialError(f"В PDF больше {MAX_TABLES} таблиц; разделите документ")
                    table = _source_table(found, material_id, filename, len(tables) + 1, page.page_number)
                    tables.append(table)
                    page_tables.append(table)
                text = _page_text(page, page_tables)
                if text.strip():
                    sections.append(f"Страница {page.page_number}\n\n" + text)
                original_page.close()
        if not sections:
            raise MaterialError("В PDF нет читаемого текстового слоя; для скана сначала требуется OCR")
        note = ("Нет текстового слоя на страницах: " + ", ".join(map(str, unreadable_pages)) +
                ". Их содержимое не извлечено; для сканов требуется OCR.") if unreadable_pages else ""
        text = "\n\n".join(sections)
        if note:
            text = "Ограничение извлечения PDF: " + note + "\n\n" + text
        return DocumentMaterial(material_id=material_id, filename=filename, text=text, tables=tables, note=note)
    except MaterialError:
        raise
    except PDFPasswordIncorrect as exc:
        raise MaterialError("PDF защищен паролем; загрузите копию без пароля") from exc
    except PdfminerException as exc:
        if exc.args and isinstance(exc.args[0], PDFPasswordIncorrect):
            raise MaterialError("PDF защищен паролем; загрузите копию без пароля") from exc
        raise MaterialError("Файл .pdf поврежден или не является документом PDF") from exc
    except (PDFSyntaxError, ValueError, TypeError, KeyError, OSError, EOFError) as exc:
        raise MaterialError("Файл .pdf поврежден или не является документом PDF") from exc
