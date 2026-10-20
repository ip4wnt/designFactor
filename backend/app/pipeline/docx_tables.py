"""DOCX -> ordered text/cards + persistent, unmodified table cells."""
from __future__ import annotations

import json
import logging
import re
from io import BytesIO
from pathlib import Path
from uuid import uuid4
from zipfile import BadZipFile, ZipFile

from lxml import etree
from pydantic import Field

from app.llm_client import LLMClient
from app.pipeline.materials import MaterialError
from app.schemas.content_plan import ContentBlock, ContentBlockType, SlideSpec, TableData
from app.schemas.source_tables import DocumentMaterial, SourceTable
from app.schemas.strict_model import StrictModel
from app.skills import load_skill

logger = logging.getLogger(__name__)
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
MAX_DATA_ROWS = 6  # + header = audit's 7 rows; fits every table renderer.
MAX_COLUMNS = 5
MAX_TABLES = 10
MAX_XML_BYTES = 20 * 1024 * 1024
MATERIAL_ID = re.compile(r"material_[0-9a-f]{32}\Z")


def _paragraph(element) -> str:
    parts = []
    for node in element.iter():
        if node.tag == W + "t":
            parts.append(node.text or "")
        elif node.tag == W + "tab":
            parts.append("\t")
        elif node.tag in {W + "br", W + "cr"}:
            parts.append("\n")
    return "".join(parts).strip()


def _blocks(parent):
    for child in parent:
        if child.tag in {W + "p", W + "tbl"}:
            yield child
        elif child.tag == W + "sdt":
            content = child.find(W + "sdtContent")
            if content is not None:
                yield from _blocks(content)


def extract_docx(filename: str, content: bytes) -> DocumentMaterial:
    try:
        with ZipFile(BytesIO(content)) as archive:
            if archive.getinfo("word/document.xml").file_size > MAX_XML_BYTES:
                raise MaterialError("Содержимое DOCX больше 20 МБ")
            xml = archive.read("word/document.xml")
        root = etree.fromstring(xml, etree.XMLParser(resolve_entities=False, no_network=True))
        if root.getroottree().docinfo.doctype:
            raise MaterialError("DOCX с объявлениями DTD не поддерживается")
        body = root.find(W + "body")
        if body is None:
            raise MaterialError("В DOCX не найдено содержимое документа")
    except (BadZipFile, KeyError, etree.XMLSyntaxError, RuntimeError) as exc:
        raise MaterialError("Файл .docx поврежден или не является документом Word") from exc
    material_id = "material_" + uuid4().hex
    tables, text = [], []
    for block in _blocks(body):
        if block.tag == W + "p":
            paragraph = _paragraph(block)
            if paragraph:
                text.append(paragraph)
            continue
        if len(tables) >= MAX_TABLES:
            raise MaterialError(f"В DOCX больше {MAX_TABLES} таблиц; разделите документ")
        rows = [["\n".join(_paragraph(p) for p in cell.findall(W + "p")).strip()
                 for cell in row.findall(W + "tc")] for row in block.findall(W + "tr")]
        limitation = ""
        merged = any(node.tag in {W + "vMerge", W + "hMerge"} or
                     (node.tag == W + "gridSpan" and node.get(W + "val", "1") != "1")
                     for node in block.iter())
        offset_cells = any(node.tag in {W + "gridBefore", W + "gridAfter"} and node.get(W + "val", "0") != "0"
                           for node in block.iter())
        if merged or any(node is not block for node in block.iter(W + "tbl")):
            limitation = "Объединенные ячейки и вложенные таблицы пока не поддерживаются"
        elif offset_cells or len(rows) < 2 or not rows[0] or any(len(row) != len(rows[0]) for row in rows):
            limitation = "Нужны прямоугольная таблица, строка заголовков и хотя бы одна строка данных"
        elif any(not header for header in rows[0]) or len(set(rows[0])) != len(rows[0]):
            limitation = "Первая строка должна содержать непустые уникальные заголовки"
        structural_error = bool(limitation)
        if not structural_error and (len(rows) - 1 > 100_000 or len(rows[0]) > 100):
            limitation = "Превышен лимит анализа: 100000 строк данных или 100 столбцов"
            structural_error = True
        elif not structural_error and (len(rows) - 1 > MAX_DATA_ROWS or len(rows[0]) > MAX_COLUMNS):
            limitation = f"Для слайда допустимо до {MAX_DATA_ROWS} строк данных и {MAX_COLUMNS} столбцов; таблица не обрезана"
        elif not structural_error and (any(len(cell) > 160 for row in rows for cell in row) or sum(len(cell) for row in rows for cell in row) > 2000):
            limitation = "Слишком длинный текст в ячейках для одного слайда; таблица не сокращена"
        ordinal = len(tables) + 1
        data = TableData(headers=rows[0], rows=rows[1:]) if not structural_error else None
        table = SourceTable(table_id=f"{material_id}.table_{ordinal}", source_name=filename,
                            ordinal=ordinal, available=not limitation, limitation=limitation,
                            summary=("Колонки: " + "; ".join(rows[0]))[:600] if rows else "Пустая таблица",
                            data=data)
        tables.append(table)
        # Replace only this element, keeping surrounding paragraphs in source order.
        text.append(f"[[{table.table_id}]]")
    if not text:
        raise MaterialError("В файле не нашлось текста или таблиц")
    return DocumentMaterial(material_id=material_id, filename=filename, text="\n\n".join(text), tables=tables)


class TableSummary(StrictModel):
    table_id: str
    summary: str = Field(min_length=1, max_length=500)


class TableSummaries(StrictModel):
    tables: list[TableSummary]


async def describe_tables(material: DocumentMaterial, settings) -> None:
    available = [table for table in material.tables if table.data is not None]
    if not available:
        return
    skill = load_skill("table_summary")
    try:
        async with LLMClient(base_url=settings.LLM_BASE_URL, model_name=settings.LLM_MODEL_NAME,
                             api_key=settings.LLM_API_KEY, max_retries=0,
                             total_timeout=settings.TABLE_SUMMARY_TIMEOUT_S) as client:
            response = await client.chat_structured(
                system_prompt=skill["system_prompt"],
                user_input=json.dumps({"tables": [{"table_id": t.table_id,
                                                   "sheet_name": t.sheet_name,
                                                   "cell_range": t.cell_range,
                                                   "number_formats": t.number_formats,
                                                   "page_number": t.page_number,
                                                   "row_count": len(t.data.rows),
                                                   "preview_only": not t.available,
                                                   "data": {"headers": t.data.headers,
                                                            "rows": t.data.rows if t.available else t.data.rows[:5]}}
                                                   for t in available]}, ensure_ascii=False),
                output_model=TableSummaries, reasoning_effort="none", temperature=0.1,
                max_completion_tokens=300 + 250 * len(available))
        summaries = {table.table_id: table.summary for table in response.tables}
        if len(summaries) != len(response.tables) or set(summaries) != {t.table_id for t in available}:
            raise ValueError("Модель вернула неверные ID описаний таблиц")
        for table in available:
            table.summary = summaries[table.table_id]
    except Exception as exc:
        # Parsing and exact cells remain useful even when the configured LLM is offline.
        logger.warning("Material table descriptions unavailable (%s)", type(exc).__name__)
        material.note = " ".join(filter(None, [material.note,
            "LLM-описания недоступны; показаны описания по заголовкам. Ячейки сохранены."]))


def table_card(table: SourceTable) -> str:
    location = f'Лист «{table.sheet_name}», диапазон {table.cell_range}. ' if table.sheet_name else ""
    if table.page_number is not None:
        location += f"Страница {table.page_number}. "
    if not table.available:
        analysis = " Полностью сохранена для анализа." if table.data is not None else ""
        return f"[Таблица {table.table_id}] {location}{table.summary}. Недоступна целиком для слайда: {table.limitation}{analysis}"
    return (f"[Таблица {table.table_id}] {location}{table.summary}\n"
            f"Колонки: {'; '.join(table.data.headers)}. Строк данных: {len(table.data.rows)}. "
            "Полные ячейки сохранены на сервере; выбирай source_table_id, не переписывай значения." +
            (f"\nРекомендация аналитика: {table.recommended_presentation}. {table.presentation_reason}" if table.parent_table_ids else ""))


def public_table(table: SourceTable) -> dict:
    """Large original tables stay server-side; do not flood the browser or job polling."""
    result = table.model_dump(exclude={"data"})
    result.update(data=table.data.model_dump() if table.available and table.data is not None else None,
                  row_count=len(table.data.rows) if table.data is not None else None,
                  column_count=len(table.data.headers) if table.data is not None else None,
                  analysis_available=table.data is not None)
    return result


def save_material(material: DocumentMaterial, original: bytes, settings) -> None:
    directory = settings.STORAGE_DIR / "materials" / material.material_id
    directory.mkdir(parents=True, exist_ok=False)
    suffix = Path(material.filename).suffix.lower()
    if suffix not in {".docx", ".xlsx", ".pdf", ".csv", ".json", ".txt", ".md"}:
        raise MaterialError("Неподдерживаемый формат сохраняемого материала")
    (directory / ("source" + suffix)).write_bytes(original)
    # Text is finalized once, with the same ID and summary shown to the planner/UI.
    for table in material.tables:
        material.text = material.text.replace(f"[[{table.table_id}]]", table_card(table))
    (directory / "manifest.json").write_text(material.model_dump_json(indent=2), encoding="utf-8")


def load_materials(ids: list[str], settings) -> list[DocumentMaterial]:
    if len(ids) > 30 or len(set(ids)) != len(ids):
        raise MaterialError("Передайте не более 30 разных материалов")
    result = []
    for material_id in ids:
        if not isinstance(material_id, str) or not MATERIAL_ID.fullmatch(material_id):
            raise MaterialError("Неверный ID материала")
        try:
            directory = settings.STORAGE_DIR / "materials" / material_id
            payload = json.loads((directory / "manifest.json").read_text(encoding="utf-8"))
            record = DocumentMaterial.model_validate(payload)
            if payload.get("format_version", 1) < 2:
                # Re-extract full cells from the retained original without another LLM call.
                recovered = extract_docx(record.filename, (directory / "source.docx").read_bytes())
                for old, new in zip(record.tables, recovered.tables):
                    new.table_id, new.summary = old.table_id, old.summary
                record.tables = recovered.tables
                record.format_version = 2
                temporary = directory / "manifest.tmp"
                temporary.write_text(record.model_dump_json(indent=2), encoding="utf-8")
                temporary.replace(directory / "manifest.json")
        except FileNotFoundError as exc:
            raise MaterialError("Материал не найден; загрузите файл заново") from exc
        result.append(record)
    return result


def bind_source_tables(slides: list[SlideSpec], tables: list[SourceTable], *, single_slide: bool = False) -> None:
    registry = {table.table_id: table for table in tables}
    for index, slide in enumerate(slides):
        if not slide.source_table_id:
            continue
        table = registry.get(slide.source_table_id)
        if table is None or not table.available or table.data is None:
            raise ValueError("source_table_id должен ссылаться на доступную таблицу из каталога")
        if not single_slide and index in {0, len(slides) - 1}:
            raise ValueError("Исходную таблицу можно выбрать только для внутреннего слайда, не обложки или финала")
        text_blocks = [b for b in slide.content_blocks if b.type not in {ContentBlockType.TABLE, ContentBlockType.CHART}]
        block = (ContentBlock(type=ContentBlockType.CHART, chart=table.chart.model_copy(deep=True)) if table.chart is not None
                 else ContentBlock(type=ContentBlockType.TABLE, table=table.data.model_copy(deep=True)))
        slide.content_blocks = text_blocks + [block]
        def source_label(source):
            if source.page_number is not None:
                return source.source_name + f", страница {source.page_number}, таблица {source.ordinal}"
            return source.source_name + (f', лист «{source.sheet_name}», {source.cell_range}'
                                         if source.sheet_name else ", таблица " + str(source.ordinal))
        sources = [source_label(registry[parent])
                   for parent in table.parent_table_ids if parent in registry]
        slide.data_note = ("Результат анализа. Источники: " + "; ".join(sources) + ". " + table.summary
                           if table.parent_table_ids else "Исходная таблица: " + source_label(table))


def verify_rendered_tables(presentation, content_plan, layout_plan) -> None:
    """A successful export must contain every selected source table, without edits."""
    if content_plan is None:
        return
    expected_indices = {i for i, slide in enumerate(content_plan.slides) if slide.source_table_id}
    rendered_indices = set()
    for index, layout in enumerate(layout_plan["slides"]):
        source_index = int(layout.get("source_index", index))
        if source_index not in expected_indices:
            continue
        spec = content_plan.slides[source_index]
        expected_chart = next((b.chart for b in spec.content_blocks if b.chart is not None), None)
        if expected_chart is not None:
            expected = (expected_chart.categories, [(s.name, tuple(s.values)) for s in expected_chart.series])
            actual = [(list(c.label for c in shape.chart.plots[0].categories),
                       [(s.name, tuple(s.values)) for s in shape.chart.series])
                      for shape in presentation.slides[index].shapes if shape.has_chart]
            if expected not in actual:
                raise ValueError(f"Диаграмма {spec.source_table_id} не перенесена с точными значениями")
            rendered_indices.add(source_index)
            continue
        expected = next(b.table for b in spec.content_blocks if b.table is not None)
        expected_cells = [expected.headers, *expected.rows]
        actual_tables = []
        for shape in presentation.slides[index].shapes:
            if shape.has_table:
                actual_tables.append([[cell.text for cell in row.cells] for row in shape.table.rows])
        if expected_cells not in actual_tables:
            raise ValueError(f"Исходная таблица {spec.source_table_id} не перенесена целиком; выберите другую компоновку")
        rendered_indices.add(source_index)
    if rendered_indices != expected_indices:
        raise ValueError("При сборке пропущен слайд с исходной таблицей")
