"""Shared analytical pass, actual SQL tools, cache and exact PPTX values; no API key."""
import asyncio
from copy import deepcopy
import json
import re
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import AsyncMock, patch

from agents.tool_context import ToolContext

from app.analysis import worker
from app.config import Settings
from app.pipeline import briefing, docx_tables, table_analysis
from app.schemas.content_plan import TableData
from app.schemas.job import Job
from app.schemas.source_tables import SourceTable, TableAnalysisReport
from test_docx_tables import docx_bytes, slide
import test_docx_tables as fixtures


def source():
    rows = [[region, str(number)] for region in ["A", "B", "C"] for number in [1, 2, 3]]
    return docx_tables.extract_docx("sales.docx", docx_bytes([[["Region", "Value"], *rows]])).tables[0]


def make_job(table=None):
    return Job(job_id="test-analysis", brief="Медиана по регионам", purpose="Сравнить",
               idea="Сравнить распределение", template_path="unused", source_tables=[table or source()])


def settings_at(directory):
    return Settings(_env_file=None, STORAGE_DIR=Path(directory), LLM_BASE_URL="http://model.invalid/v1",
                    LLM_MODEL_NAME="test", SKILLS_DIR=Path(__file__).resolve().parents[1] / "skills")


class AnalysisTests(unittest.IsolatedAsyncioTestCase):
    async def test_invalid_model_response_retries_only_once_then_reports_failure(self):
        from agents import ModelBehaviorError
        with TemporaryDirectory() as tmp:
            settings=settings_at(tmp)
            with patch('app.skills.get_settings',return_value=settings):
                request=table_analysis.analysis_request(make_job(),settings)
            async def fake_run(agent,*args,**kwargs):
                await agent.model.get_response()
            response=AsyncMock(side_effect=ModelBehaviorError('empty truncated response'))
            with patch.object(worker.Runner,'run',new=fake_run), \
                 patch.object(worker.OpenAIChatCompletionsModel,'get_response',response):
                report=await worker.run_agent(request,Path(tmp)/'progress.json',settings)
            self.assertEqual(response.await_count,2)
            self.assertEqual(report.status,'failed')
            self.assertEqual(report.artifacts,[])

    async def test_sql_aliases_handle_multiline_headers_without_changing_sources(self):
        with TemporaryDirectory() as tmp:
            settings=settings_at(tmp)
            settings.TABLE_ANALYSIS_MAX_RESULTS=1
            original=source()
            original.data.headers=['Город\nтеста','Выручка\nруб.']
            job=make_job(original)
            before=deepcopy(original)
            with patch('app.skills.get_settings',return_value=settings):
                request=table_analysis.analysis_request(job,settings)
            async def fake_run(agent,user,*,context,**kwargs):
                self.assertFalse(agent.tool_use_behavior(None,[]).is_final_output)
                metadata=json.loads(user)['tables'][0]
                self.assertEqual(metadata['sql_source'],'table_1_sql')
                self.assertEqual(metadata['sql_columns'][1],{'column':'col_02_vyruchka_rub','header':'Выручка\nруб.'})
                self.assertEqual(list(context.views['table_1'].columns),original.data.headers)
                self.assertEqual(list(context.views['table_1_sql'].columns),['col_01_gorod_testa','col_02_vyruchka_rub'])
                tools={t.name:t for t in agent.tools}
                async def call(name,args):
                    ctx=ToolContext(context=context,tool_name=name,tool_call_id='test',tool_arguments=json.dumps(args))
                    result=await tools[name].on_invoke_tool(ctx,json.dumps(args))
                    return json.loads(result)
                dropped=await call('drop_view',{'name':'table_1_sql'})
                self.assertIn('error',dropped)
                result=await call('create_view',{'name':'sums','sql':
                    'SELECT col_01_gorod_testa AS "Город", SUM(TRY_CAST(col_02_vyruchka_rub AS DOUBLE)) AS "Выручка" '
                    'FROM table_1_sql GROUP BY col_01_gorod_testa ORDER BY col_01_gorod_testa'})
                self.assertNotIn('error',result)
                self.assertFalse(agent.tool_use_behavior(None,[]).is_final_output)
                result=await call('prepare_slide_data',{'title':'Выручка по городам','kind':'table',
                    'reason':'Сумма всех наблюдений по городу','x':None,'y':None,'source':'sums',
                    'presentation_columns':[{'key':'Город','label':'Город','unit':''},{'key':'Выручка','label':'Выручка','unit':'руб.'}]})
                self.assertNotIn('error',result)
                self.assertTrue(agent.tool_use_behavior(None,[]).is_final_output)
            with patch.object(worker.Runner,'run',new=fake_run):
                report=await worker.run_agent(request,Path(tmp)/'progress.json',settings)
            self.assertEqual(report.status,'completed')
            self.assertEqual(report.artifacts[0].data.rows,[['A','6'],['B','6'],['C','6']])
            self.assertEqual(report.artifacts[0].data.headers,['Город','Выручка, руб.'])
            self.assertEqual(report.artifacts[0].parent_table_ids,[original.table_id])
            self.assertEqual(original,before)

    def test_sql_identifiers_preserve_quarter_meaning_and_are_unique(self):
        names=worker.sql_column_names(['Город','I кв.\nтыс. руб.','II кв.\nтыс. руб.','IV кв.\nтыс. руб.','收益','收益'])
        self.assertIn('_i_kv_',names[1])
        self.assertIn('_ii_kv_',names[2])
        self.assertIn('_iv_kv_',names[3])
        self.assertEqual(len(set(names)),6)
        self.assertTrue(all(re.fullmatch(r'[a-z][a-z0-9_]*',name) for name in names))

    async def test_existing_agent_executes_sql_join_median_then_exact_plan_and_pptx(self):
        from app.orchestrator import apply_manual_text
        with TemporaryDirectory() as tmp:
            settings = settings_at(tmp)
            job = make_job()
            other = SourceTable(table_id="lookup.table_1", source_name="lookup.docx", ordinal=1,
                summary="Region labels", available=True,
                data=TableData(headers=["Region", "Label"], rows=[["A", "Alpha"], ["B", "Beta"], ["C", "Gamma"]]))
            job.source_tables.append(other)
            original = deepcopy(job.source_tables)
            with patch("app.skills.get_settings", return_value=settings):
                request = table_analysis.analysis_request(job, settings)
            output = Path(tmp) / "progress.json"

            async def fake_run(agent, user, *, context, **kwargs):
                self.assertEqual(agent.model_settings.extra_body["reasoning_effort"], "low")
                self.assertFalse(agent.model_settings.parallel_tool_calls)
                self.assertTrue(kwargs["run_config"].tracing_disabled)
                self.assertEqual(len(context.df), 9)
                tools = {t.name: t for t in agent.tools}
                async def call(name, args):
                    ctx = ToolContext(context=context, tool_name=name, tool_call_id="test", tool_arguments=json.dumps(args))
                    result = await tools[name].on_invoke_tool(ctx, json.dumps(args))
                    self.assertNotIn("error", json.loads(result), result)
                await call("describe_dataframe", {"source": "table_1"})
                await call("create_view", {"name": "medians", "sql":
                    'SELECT b.Label, median(CAST(a.Value AS DOUBLE)) AS Median, count(*) AS Count '
                    'FROM table_1 a JOIN table_2 b ON a.Region=b.Region GROUP BY b.Label ORDER BY b.Label'})
                await call("prepare_slide_data", {"title": "Медиана равна 2", "kind": "table",
                    "reason": "Медиана по всем наблюдениям каждого региона после соединения со справочником.",
                    "x": None, "y": None, "source": "medians", "presentation_columns":[
                        {"key":"Label","label":"Регион","unit":""},
                        {"key":"Median","label":"Медиана","unit":""},
                        {"key":"Count","label":"Наблюдения","unit":"шт."}]})
                self.assertTrue(output.exists(), "Completed computations survive a subsequent timeout")
                await tools["finish_analysis"].on_invoke_tool(ToolContext(context=context, tool_name="finish_analysis", tool_call_id="finish", tool_arguments="{}"), "{}")

            with patch.object(worker.Runner, "run", new=fake_run):
                report = await worker.run_agent(request, output, settings)
            self.assertEqual(report.status, "completed", report.notes)
            self.assertEqual(report.artifacts[0].data.rows, [["Alpha", "2", "3"], ["Beta", "2", "3"], ["Gamma", "2", "3"]])
            self.assertEqual(len(report.artifacts[0].parent_table_ids), 2)
            self.assertEqual(report.artifacts[0].operations[-1]["tool"], "create_view")
            self.assertEqual(job.source_tables, original)
            job.table_analysis = report
            job.source_tables += report.artifacts
            selected = report.artifacts[0]
            client = AsyncMock()
            client.__aenter__.return_value = client
            client.chat_structured.side_effect = lambda **kw: briefing.PlanOut(brief="Отчет", purpose="Сравнить", story="Сравнение",
                slides=[slide("cover"), slide("data", source_table_id=selected.table_id), slide("end")])
            with patch.object(briefing, "_client", return_value=client), patch("app.skills.get_settings", return_value=settings):
                variants = await briefing.generate_plan_variants(job, 3, None)
            self.assertEqual(len(variants), 3)
            for variant in variants.values():
                spec = variant.plan.slides[1]
                self.assertEqual(spec.content_blocks[-1].table, selected.data)
                self.assertIn("sales.docx", spec.data_note)
                self.assertEqual(apply_manual_text(spec, "Новый заголовок", "Текст", None).data_note, spec.data_note)
            for call in client.chat_structured.call_args_list:
                self.assertIn('"Alpha", "2", "3"', call.kwargs["user_input"])
            # Reuse the actual three-design roundtrip with computed cells.
            with patch("test_docx_tables.docx_tables.extract_docx", return_value=type("Material", (), {"tables": [selected]})()):
                fixtures.RenderTests("test_all_three_designs_keep_exact_editable_source_tables").test_all_three_designs_keep_exact_editable_source_tables()

    async def test_shared_cache_concurrent_jobs_and_invalidation(self):
        with TemporaryDirectory() as tmp:
            settings = settings_at(tmp)
            first = make_job()
            second = first.model_copy(deep=True)
            second.tone = "Другой тон"
            async def calculate(request, directory, config):
                await asyncio.sleep(0)
                return TableAnalysisReport(key=request["key"], status="completed", model_calls=4)
            mock = AsyncMock(side_effect=calculate)
            with patch("app.skills.get_settings", return_value=settings), patch.object(table_analysis, "_run_worker", mock):
                await asyncio.gather(table_analysis.prepare_table_analysis(first, settings), table_analysis.prepare_table_analysis(second, settings))
                self.assertEqual(mock.await_count, 1)
                self.assertTrue(second.table_analysis.cache_hit)
                second.goal = "Другой вопрос к данным"
                await table_analysis.prepare_table_analysis(second, settings)
                self.assertEqual(mock.await_count, 2)
                second.source_tables[0].data.rows[0][1] = "100"
                await table_analysis.prepare_table_analysis(second, settings)
                self.assertEqual(mock.await_count, 3)
            self.assertNotIn("LLM_API_KEY", (Path(tmp) / "table_analysis" / first.table_analysis.key / "report.json").read_text())

    async def test_no_tables_and_disabled_never_start_agent(self):
        with TemporaryDirectory() as tmp:
            settings = settings_at(tmp)
            job = make_job()
            job.source_tables = []
            with patch.object(table_analysis, "_run_worker", new=AsyncMock()) as run:
                self.assertEqual((await table_analysis.prepare_table_analysis(job, settings)).status, "skipped")
                job.source_tables = [source()]
                settings.TABLE_ANALYSIS_ENABLED = False
                self.assertEqual((await table_analysis.prepare_table_analysis(job, settings)).status, "disabled")
                run.assert_not_called()

    async def test_worker_error_does_not_break_planning_and_can_be_retried(self):
        with TemporaryDirectory() as tmp:
            settings = settings_at(tmp)
            job = make_job()
            with patch("app.skills.get_settings", return_value=settings), \
                 patch.object(table_analysis, "_run_worker", new=AsyncMock(side_effect=RuntimeError("unavailable"))) as run:
                report = await table_analysis.prepare_table_analysis(job, settings)
                self.assertEqual(report.status, "failed")
                self.assertIn("failed", table_analysis.analysis_prompt(report))
                await table_analysis.prepare_table_analysis(job, settings)
                self.assertEqual(run.await_count, 2)
                self.assertEqual(len(job.source_tables[0].data.rows), 9)

    async def test_deadline_kills_process_and_retains_only_completed_results(self):
        with TemporaryDirectory() as tmp:
            settings = settings_at(tmp)
            settings.TABLE_ANALYSIS_TIMEOUT_S = 0.01
            job = make_job()
            with patch("app.skills.get_settings", return_value=settings):
                request = table_analysis.analysis_request(job, settings)
            directory = Path(tmp) / "run"
            directory.mkdir()
            artifact = SourceTable(table_id="computed", source_name="Analysis", ordinal=1, summary="Median", available=True,
                data=TableData(headers=["Median"], rows=[["2"]]), parent_table_ids=[job.source_tables[0].table_id])
            done = asyncio.Event()
            class Process:
                returncode = None
                killed = False
                async def wait(self):
                    await done.wait()
                def kill(self):
                    self.killed = True
                    self.returncode = -9
                    done.set()
            process = Process()
            async def launch(*args, **kwargs):
                self.assertEqual(kwargs["env"]["LLM_BASE_URL"], settings.LLM_BASE_URL)
                (directory / "progress.json").write_text(TableAnalysisReport(key=request["key"], status="partial", artifacts=[artifact]).model_dump_json())
                return process
            with patch.object(asyncio, "create_subprocess_exec", new=launch):
                report = await table_analysis._run_worker(request, directory, settings)
            self.assertTrue(process.killed)
            self.assertEqual(report.status, "partial")
            self.assertEqual(report.artifacts[0].data.rows, [["2"]])


class PersistenceTests(unittest.TestCase):
    def test_old_registry_recovers_large_table_from_original_with_same_id(self):
        with TemporaryDirectory() as tmp:
            settings = settings_at(tmp)
            original = docx_bytes([[["Region", "Value"], *[["A", str(i)] for i in range(12)]]])
            document = docx_tables.extract_docx("old.docx", original)
            docx_tables.save_material(document, original, settings)
            path = Path(tmp) / "materials" / document.material_id / "manifest.json"
            payload = json.loads(path.read_text(encoding="utf-8"))
            payload.pop("format_version")
            payload["tables"][0]["data"] = None
            path.write_text(json.dumps(payload), encoding="utf-8")
            loaded = docx_tables.load_materials([document.material_id], settings)[0]
            self.assertEqual(loaded.tables[0].table_id, document.tables[0].table_id)
            self.assertEqual(loaded.tables[0].data, document.tables[0].data)
            self.assertFalse(loaded.tables[0].available)
            self.assertEqual(json.loads(path.read_text(encoding="utf-8"))["format_version"], 2)


if __name__ == "__main__":
    unittest.main()
