"""Bounded subprocess hosting the existing analytical agent and its tools."""
from __future__ import annotations

import asyncio
from copy import deepcopy
from dataclasses import replace
import json
import logging
import math
from pathlib import Path
import re
import sys
import unicodedata

from agents import Agent, FunctionTool, ModelBehaviorError, ModelSettings, OpenAIChatCompletionsModel, Runner, RunConfig, ToolsToFinalOutputResult, function_tool
from openai import AsyncOpenAI, omit
import pandas as pd

from app.analysis.dataframe_tools import DATAFRAME_TOOLS, DATAFRAME_AGENT_INSTRUCTIONS, DataContext
from app.config import get_settings
from app.schemas.content_plan import ChartData, ChartSeries, ChartType, TableData
from app.schemas.source_tables import SourceTable, TableAnalysisReport, TableColumn

logger = logging.getLogger(__name__)

def sql_column_names(headers):
    letters = dict(zip('абвгдеёжзийклмнопрстуфхцчшщъыьэюя',
        ['a','b','v','g','d','e','yo','zh','z','i','y','k','l','m','n','o','p','r','s','t','u','f','kh','ts','ch','sh','shch','','y','','e','yu','ya']))
    result = []
    for index, header in enumerate(headers, 1):
        readable = ''.join(letters.get(char, char) for char in header.casefold())
        readable = unicodedata.normalize('NFKD', readable).encode('ascii', 'ignore').decode()
        readable = re.sub(r'[^a-z0-9]+', '_', readable).strip('_')[:70] or 'value'
        result.append(f'col_{index:02d}_{readable}')
    return result


_MONTHS = ("янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек")
_ISO_DATETIME = re.compile(r"^(\d{4})-(\d{2})-(\d{2})(?:[T ]\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?)?$")


def display_value(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float):
        if not math.isfinite(value):
            raise ValueError("Nonfinite result")
        if value.is_integer():
            return str(int(value))
    text = str(value)
    # Даты/метки времени из SQL (2023-01-01T00:00:00.000) на слайде читаются как «янв 2023»
    # или «15 янв 2023»; сами значения расчета не меняются.
    match = _ISO_DATETIME.match(text)
    if match:
        year, month, day = match.group(1), int(match.group(2)), int(match.group(3))
        if 1 <= month <= 12:
            return f"{_MONTHS[month - 1]} {year}" if day == 1 else f"{day} {_MONTHS[month - 1]} {year}"
    return text


def publication_columns(keys, metadata):
    if not isinstance(metadata, list) or not metadata:
        raise ValueError("Передай presentation_columns: название и известную единицу для каждой колонки.")
    columns = [item if isinstance(item, TableColumn) else TableColumn.model_validate(item, strict=True) for item in metadata]
    by_key = {column.key: column for column in columns}
    if len(by_key) != len(columns) or set(by_key) != set(keys):
        raise ValueError("presentation_columns должны точно покрывать все колонки результата, без повторов и лишних имен.")
    ordered, headers = [], []
    for key in keys:
        column = by_key[key].model_copy(update={"label": by_key[key].label.strip(), "unit": by_key[key].unit.strip()})
        if not column.label or '_' in column.label or re.fullmatch(r'[a-z][a-z0-9]*', column.label):
            raise ValueError(f"Для {key} нужно читаемое название вместо рабочего имени: например «Годовая выручка».")
        header = column.label
        if column.unit and not header.casefold().rstrip(') ').endswith(column.unit.casefold()):
            header += ", " + column.unit
        if len(header) > 160:
            raise ValueError("Заголовок с единицей должен быть не длиннее 160 символов.")
        ordered.append(column)
        headers.append(header)
    if len(set(headers)) != len(headers):
        raise ValueError("Читаемые заголовки должны быть уникальными.")
    return ordered, headers


def compile_artifacts(slides, request) -> list[SourceTable]:
    artifacts = []
    for index, slide in enumerate(slides):
        kind = slide["kind"]
        if kind not in {"table", "bar", "column", "line", "pie"}:
            raise ValueError("Поддерживаются table/bar/column/line/pie")
        keys = slide["columns"]
        metadata, headers = publication_columns(keys, slide.get("presentation_columns"))
        rows = [[display_value(row.get(column)) for column in keys] for row in slide["data"]]
        if (kind == "table" and len(rows) > 6) or len(headers) > 5 or any(len(c) > 160 for row in [headers, *rows] for c in row):
            raise ValueError("Result exceeds 6 data rows, 5 columns or 160 characters per cell")
        if kind == "table" and sum(len(c) for row in [headers, *rows] for c in row) > 2000:
            raise ValueError("Result exceeds slide text budget")
        chart = None
        if kind != "table":
            x, y = slide["x"], slide["y"]
            columns = {column.key: column for column in metadata}
            if len({columns[key].unit.casefold() for key in y}) != 1:
                raise ValueError("Нельзя совмещать разные единицы на одной оси; подготовь таблицу или отдельные диаграммы")
            chart = ChartData(chart_type=ChartType("bar" if kind == "column" else kind),
                categories=[display_value(row[x]) for row in slide["data"]],
                series=[ChartSeries(name=columns[key].label, values=[row[key] for row in slide["data"]]) for key in y],
                category_axis_title=columns[x].label, value_axis_title=columns[y[0]].label if len(y) == 1 else "Показатели",
                units=columns[y[0]].unit, orientation="horizontal" if kind == "bar" else "vertical" if kind == "column" else "auto",
                source="; ".join(t["source_name"] for t in request.get("tables", []) if t["table_id"] in slide["parent_table_ids"])[:300])
        reason = slide["reason"]
        notes = list(slide.get("notes", []))
        # Честность выборки: маленький источник (до 10 строк) показан не целиком —
        # говорим об этом прямо, чтобы «3 из 5 регионов» не выдавалось за всю картину.
        parent_rows = [len(t["data"]["rows"]) for t in request.get("tables", [])
                       if t["table_id"] in slide["parent_table_ids"] and t.get("data")]
        if len(parent_rows) == 1 and len(rows) < parent_rows[0] <= 10:
            coverage = f"Показано {len(rows)} из {parent_rows[0]} строк источника"
            if coverage not in reason:
                reason = reason + " " + coverage + "."
            notes.append(coverage)
        artifacts.append(SourceTable(
            table_id=f"analysis_{request['key']}.table_{index + 1}", source_name="Результат анализа",
            ordinal=index + 1, summary=slide["title"] + ". " + reason, available=True,
            data=TableData(headers=headers, rows=rows), parent_table_ids=slide["parent_table_ids"],
            operations=slide["operations"], column_metadata=metadata, limitation="; ".join(notes),
            recommended_presentation="chart" if chart else "table", presentation_reason=reason, chart=chart))
    return artifacts


async def run_agent(request, output_path: Path, settings):
    sources = request["tables"]
    frames = {f"table_{i+1}": pd.DataFrame(t["data"]["rows"], columns=t["data"]["headers"], dtype="string")
              for i, t in enumerate(sources)}
    context = DataContext(df=next(iter(frames.values())))
    context.views.update(frames)
    parents = {alias: [table["table_id"]] for alias, table in zip(frames, sources)}
    parents["original"] = [sources[0]["table_id"]]
    sql_aliases = {}
    for alias, frame in frames.items():
        # Stable SQL identifiers avoid spelling/escaping multiline or Cyrillic headers.
        # Original frames and source cell values stay unchanged.
        sql_alias = alias + "_sql"
        sql_frame = frame.copy(deep=False)
        sql_frame.columns = sql_column_names(frame.columns)
        context.views[sql_alias] = sql_frame
        parents[sql_alias] = list(parents[alias])
        sql_aliases[alias] = sql_alias
    protected_inputs = set(frames) | set(sql_aliases.values())
    header_to_sql = {alias: dict(zip(frames[alias].columns, context.views[sql_alias].columns))
                     for alias, sql_alias in sql_aliases.items()}

    def normalize_sql(sql: str) -> str:
        """Модель часто пишет исходные заголовки («Месяц») вместо SQL-имен (col_01_mesyats):
        подменяем их для таблиц, упомянутых в запросе, чтобы не тратить бюджет на Binder Error."""
        mentioned = set(re.findall(r"[A-Za-z_][A-Za-z0-9_]*", sql))
        for alias, mapping in header_to_sql.items():
            # Только для *_sql-снимков: обычные table_N сохраняют исходные заголовки и работают как есть.
            if sql_aliases[alias] in mentioned:
                for header, column in sorted(mapping.items(), key=lambda kv: -len(str(kv[0]))):
                    if header and header != column:
                        sql = sql.replace(f'"{header}"', f'"{column}"')
        return sql
    history = {alias: [] for alias in parents}
    counts = {"model": 0, "tool": 0}

    def publish(status="partial", notes=None):
        report = TableAnalysisReport(key=request["key"], status=status,
            artifacts=compile_artifacts(context.slides, request), notes=notes or [],
            model_calls=counts["model"], tool_calls=counts["tool"])
        temporary = output_path.with_suffix(".tmp")
        temporary.write_text(report.model_dump_json(indent=2), encoding="utf-8")
        temporary.replace(output_path)
        return report

    tool_lock = asyncio.Lock()

    def instrument(tool):
        async def invoke(ctx, arguments):
            # Модель может прислать несколько вызовов за один ход; SDK выполняет их параллельно.
            # Без сериализации prepare_slide_data разных результатов перепутывают друг друга
            # (context.slides[-1] оказывается чужим) и все отбраковываются при публикации.
            async with tool_lock:
                return await invoke_serialized(ctx, arguments)

        async def invoke_serialized(ctx, arguments):
            if counts["tool"] >= request["max_tool_calls"]:
                logger.info("tool %s rejected: budget exhausted", tool.name)
                return json.dumps({"error": "Tool budget exhausted; finish_analysis now"})
            counts["tool"] += 1
            args = json.loads(arguments)
            source = args.get("source", "original" if tool.name != "prepare_slide_data" else "result")
            if tool.name == "drop_view" and args.get("name") in protected_inputs:
                return json.dumps({"error": "Input table snapshots cannot be dropped"})
            if tool.name == "prepare_slide_data" and (args.get("kind") not in {"table", "bar", "column", "line", "pie"} or len(context.slides) >= request["max_results"]):
                return json.dumps({"error": "Publish table/bar/column/line/pie within the result budget"})
            metadata = None
            if tool.name == "prepare_slide_data":
                supplied = args.pop("presentation_columns", None)
                frame = context.df if source == "original" else context.result if source == "result" else context.views.get(source)
                try:
                    if frame is None:
                        raise ValueError("Сначала подготовь результат расчета в указанном source.")
                    keys = list(frame.columns) if args["kind"] == "table" else [args.get("x"), *(args.get("y") or [])]
                    metadata, _ = publication_columns(keys, supplied)
                except ValueError as error:
                    return json.dumps({"error": str(error)}, ensure_ascii=False)
            # SQL may join any registered input/view; provenance is conservatively inclusive.
            if tool.name == "create_view" and isinstance(args.get("sql"), str):
                fixed = normalize_sql(args["sql"])
                if fixed != args["sql"]:
                    logger.info("SQL нормализован: заголовки заменены на SQL-имена колонок")
                    args["sql"] = fixed
                    arguments = json.dumps(args, ensure_ascii=False)
            if tool.name == "create_view":
                # Источник подписывается по таблицам, которые реально упомянуты в SQL
                # (identifiers вроде table_2, table_2_sql, result, названия представлений);
                # если разобрать не удалось — консервативно берем все.
                mentioned = set(re.findall(r"[A-Za-z_][A-Za-z0-9_]*", str(args.get("sql", ""))))
                used = [alias for alias in parents if alias in mentioned]
                if not used:
                    used = list(parents)
                parent_ids = sorted({p for alias in used for p in parents.get(alias, [])})
                previous = [op for alias in used for op in history.get(alias, [])]
                previous = list({json.dumps(op, sort_keys=True, ensure_ascii=False): op for op in previous}.values())
            else:
                parent_ids, previous = parents.get(source, []), history.get(source, [])
            before = len(context.slides)
            forwarded = json.dumps({**args, "kind": "donut" if args["kind"] == "pie" else args["kind"]}, ensure_ascii=False) if tool.name == "prepare_slide_data" else arguments
            result = await tool.on_invoke_tool(ctx, forwarded)
            logger.info("tool %s args=%s -> %s", tool.name, json.dumps(args, ensure_ascii=False)[:600],
                        str(result)[:400].replace("\n", " "))
            if isinstance(result, str) and "error" in json.loads(result):
                return result
            if tool.name in {"create_view", "select_dataframe", "aggregate_dataframe", "pivot_dataframe"}:
                operation = {"tool": tool.name, "arguments": args}
                parents["result"], history["result"] = list(parent_ids), [*previous, operation]
                if tool.name == "create_view":
                    parents[args["name"]], history[args["name"]] = list(parent_ids), [*previous, operation]
            if tool.name == "prepare_slide_data" and len(context.slides) > before:
                produced = json.loads(result).get("slide_index") if isinstance(result, str) else None
                item = context.slides[produced] if isinstance(produced, int) and 0 <= produced < len(context.slides) else context.slides[-1]
                item["kind"] = args["kind"]
                item.update(parent_table_ids=list(parent_ids), operations=[*previous])
                item["presentation_columns"] = [column.model_dump() for column in metadata]
                try:
                    publish()
                except ValueError as error:
                    context.slides.remove(item)
                    logger.info("prepare_slide_data отклонен при публикации: %s (сохранено слайдов: %s)", error, len(context.slides))
                    return json.dumps({"error": str(error)})
                logger.info("prepare_slide_data принят; сохранено слайдов: %s", len(context.slides))
                payload = json.loads(result)
                payload["kind"] = item["kind"]
                payload["display_headers"] = publication_columns(item["columns"], metadata)[1]
                result = json.dumps(payload, ensure_ascii=False)
            return result
        schema = deepcopy(tool.params_json_schema)
        if tool.name == "prepare_slide_data":
            schema["properties"]["kind"]["enum"] = ["table", "bar", "column", "line", "pie"]
            schema["properties"]["presentation_columns"] = {
                "type": "array", "items": TableColumn.model_json_schema(mode="serialization"),
                "description": "Для каждой рабочей колонки укажи key, читаемый label и известную unit. Значения и масштаб не меняй."}
            schema["required"] = [*schema["required"], "presentation_columns"]
        description = ("Сохрани результат и рекомендуй table, bar (полосы), column (столбцы), line или pie. "
                       "В reason объясни выбор; x/y задают данные, presentation_columns - подписи и единицы."
                       if tool.name == "prepare_slide_data" else tool.description)
        return replace(tool, description=description, params_json_schema=schema, on_invoke_tool=invoke)

    @function_tool
    def finish_analysis() -> str:
        """Finish the common analysis; prepared results have already been saved."""
        publish("completed")
        return "Analysis complete"

    def complete_at_budget(ctx, tool_results):
        if len(context.slides) >= request["max_results"] or any(result.tool.name == "finish_analysis" for result in tool_results):
            publish("completed")
            return ToolsToFinalOutputResult(is_final_output=True, final_output="Analysis complete")
        return ToolsToFinalOutputResult(is_final_output=False)

    class CountingModel(OpenAIChatCompletionsModel):
        retried = False

        async def get_response(self, *args, **kwargs):
            counts["model"] += 1
            publish()
            try:
                return await super().get_response(*args, **kwargs)
            except ModelBehaviorError:
                if self.retried:
                    raise
                self.retried = True
                logger.warning("Invalid analytical model response; retrying once", exc_info=True)
                counts["model"] += 1
                publish()
                return await super().get_response(*args, **kwargs)

    # Uses this project's backend/.env in local runs and the same variables in Docker.
    async with AsyncOpenAI(base_url=settings.LLM_BASE_URL, api_key=settings.LLM_API_KEY or "unused",
                           max_retries=0, timeout=request["timeout"],
                           default_headers={"Authorization": f"Bearer {settings.LLM_API_KEY}" if settings.LLM_API_KEY else omit}) as client:
        agent = Agent(name="Аналитик данных презентации",
            instructions=DATAFRAME_AGENT_INSTRUCTIONS + "\n\n" + request["instructions"],
            model=CountingModel(model=settings.LLM_MODEL_NAME, openai_client=client),
            model_settings=ModelSettings(parallel_tool_calls=False, temperature=0.1, max_tokens=6000,
                                         extra_body={"reasoning_effort": "low"}),
            tools=[instrument(tool) for tool in DATAFRAME_TOOLS] + [finish_analysis],
            tool_use_behavior=complete_at_budget)
        user = {"brief": request["brief"], "idea": request["idea"], "audience": request["audience"],
                "goal": request["goal"], "must_have": request["must_have"], "materials": request["materials"],
                "budget": {"model_turns": request["max_turns"], "tool_calls": request["max_tool_calls"],
                           "max_results": request["max_results"],
                           "instruction": "Сохрани первый полезный расчет через prepare_slide_data сразу; оставь шаг на finish_analysis."},
                "tables": [{"source": alias, "table_id": t["table_id"], "description": t["summary"],
                            "source_name": t["source_name"], "sheet_name": t.get("sheet_name", ""),
                            "cell_range": t.get("cell_range", ""), "number_formats": t.get("number_formats", {}),
                            "page_number": t.get("page_number"), "bounding_box": t.get("bounding_box", []),
                            "sql_source": sql_aliases[alias],
                            "sql_columns": [{"column": column, "header": name}
                                            for column, name in zip(context.views[sql_aliases[alias]].columns,t["data"]["headers"])],
                            "columns": t["data"]["headers"], "rows": len(t["data"]["rows"])}
                           for alias, t in zip(frames, sources)]}
        try:
            await Runner.run(agent, json.dumps(user, ensure_ascii=False), context=context,
                max_turns=request["max_turns"], run_config=RunConfig(tracing_disabled=True))
            return publish("completed")
        except Exception as error:
            logger.exception("Table analysis failed after %s model and %s tool calls", counts["model"], counts["tool"])
            return publish("partial" if context.slides else "failed",
                           ["Анализ не завершён: " + type(error).__name__ + ". Используются только сохранённые результаты."])


if __name__ == "__main__":
    # Журнал шагов агента уходит в worker.log рядом с отчетом: без него пустой результат не разобрать.
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s", stream=sys.stderr)
    request_path, output = map(Path, sys.argv[1:3])
    request = json.loads(request_path.read_text(encoding="utf-8"))
    try:
        asyncio.run(run_agent(request, output, get_settings()))
    except Exception as error:
        output.write_text(TableAnalysisReport(key=request["key"], status="failed",
            notes=["Не удалось запустить аналитика: " + type(error).__name__]).model_dump_json(), encoding="utf-8")
