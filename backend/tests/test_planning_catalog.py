"""Planning schema follows insertable tables; source cells remain server-owned."""
import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import AsyncMock, patch

from pydantic import ValidationError

from app.config import Settings
from app.pipeline import briefing, docx_tables, table_analysis
from app.schemas.content_plan import ContentBlock, TableData, ChartData, ChartSeries, ChartType
from app.schemas.job import Job
from app.schemas.source_tables import SourceTable, TableAnalysisReport


def source(identifier="available",available=True):
    return SourceTable(table_id=identifier,source_name="test.docx",ordinal=1,
        available=available,summary="Продажи",limitation="Слишком большая таблица" if not available else "",
        data=TableData(headers=["Город","Выручка"],rows=[["Орёл","007"]]))


def payload(identifier=None):
    def slide(key,table_id=None):
        return dict(slide_id=key,title="Вывод",purpose="Сравнить",source_table_id=table_id,
            content_blocks=[dict(type="text",text="Отчет",bullets=None,chart=None,table=None)])
    return dict(brief="Отчет",purpose="Сравнить",story="Результаты и решение",
                slides=[slide("cover"),slide("data",identifier),slide("end")])


class PlanningCatalogTests(unittest.TestCase):
    def test_schema_allows_only_insertable_ids_and_none(self):
        model=briefing._plan_response_model([source(),source("blocked",False)])
        schema=model.model_json_schema(mode="serialization")
        field=schema["$defs"]["PlanningSlide"]["properties"]["source_table_id"]
        self.assertEqual(field["enum"],["available",None])
        for invalid in ("blocked","invented"):
            with self.subTest(invalid=invalid),self.assertRaises(ValidationError):
                model.model_validate_json(json.dumps(payload(invalid)),strict=True)
        self.assertEqual(model.model_validate_json(json.dumps(payload("available")),strict=True).slides[1].source_table_id,"available")

    def test_no_results_after_timeout_accepts_narrative_plan(self):
        model=briefing._plan_response_model([source("blocked",False)])
        out=model.model_validate_json(json.dumps(payload()),strict=True)
        self.assertTrue(all(s.source_table_id is None for s in out.slides))
        with self.assertRaises(ValidationError):
            model.model_validate_json(json.dumps(payload("blocked")),strict=True)

    def test_selected_source_accepts_placeholder_and_binds_exact_cells(self):
        table=source()
        raw=payload(table.table_id)
        raw["slides"][1]["content_blocks"].append(dict(type="table",table=None,chart=None,text=None,bullets=None))
        out=briefing._plan_response_model([table]).model_validate_json(json.dumps(raw),strict=True)
        self.assertEqual(len(out.slides[1].content_blocks),1)
        docx_tables.bind_source_tables(out.slides,[table])
        self.assertEqual(out.slides[1].content_blocks[-1].table,table.data)
        self.assertEqual(out.slides[1].content_blocks[-1].table.rows[0],["Орёл","007"])
        self.assertEqual(raw["slides"][1]["content_blocks"][-1]["type"],"table")

    def test_table_without_source_still_requires_cells(self):
        raw=payload()
        raw["slides"][1]["content_blocks"]=[dict(type="table",table=None,chart=None,text=None,bullets=None)]
        with self.assertRaises(ValidationError):
            briefing._plan_response_model([]).model_validate_json(json.dumps(raw),strict=True)
        with self.assertRaises(ValidationError):
            ContentBlock.model_validate_json(json.dumps(raw["slides"][1]["content_blocks"][0]),strict=True)

    def test_computed_table_remains_selectable_while_its_parent_is_blocked(self):
        parent=source("parent",False)
        computed=source("computed")
        computed.parent_table_ids=[parent.table_id]
        out=briefing._plan_response_model([parent,computed]).model_validate_json(json.dumps(payload("computed")),strict=True)
        docx_tables.bind_source_tables(out.slides,[parent,computed])
        self.assertEqual(out.slides[1].content_blocks[-1].table,computed.data)

    def test_catalog_rejects_handwritten_table_even_when_numbers_are_schema_valid(self):
        raw=payload()
        raw['slides'][1]['content_blocks']=[dict(type='table',table=dict(headers=['Город','Выручка'],
            rows=[['Москва','1250000']]),text=None,bullets=None,chart=None)]
        for available in (True,False):
            with self.subTest(available=available),self.assertRaises(ValidationError):
                briefing._plan_response_model([source(available=available)]).model_validate_json(json.dumps(raw),strict=True)

    def test_catalog_rejects_handwritten_chart_but_plain_text_workflow_keeps_chart_support(self):
        raw=payload()
        raw['slides'][1]['content_blocks']=[dict(type='chart',chart=dict(chart_type='bar',
            categories=['Москва'],series=[dict(name='Выручка',values=[1250000])]),text=None,bullets=None,table=None)]
        with self.assertRaises(ValidationError):
            briefing._plan_response_model([source()]).model_validate_json(json.dumps(raw),strict=True)
        self.assertEqual(briefing._plan_response_model([]).model_validate_json(json.dumps(raw),strict=True).slides[1].content_blocks[0].type.value,'chart')


class PlanningRetryTests(unittest.IsolatedAsyncioTestCase):
    async def test_plan_regeneration_retries_failed_analysis_before_planning(self):
        job=Job(job_id="test",brief="Отчет",purpose="Сравнить",template_path="unused",source_tables=[source()],
                table_analysis=TableAnalysisReport(status="failed"))
        with patch.object(briefing,"prepare_table_analysis",new=AsyncMock()) as analyse, \
             patch.object(briefing,"_generate_variant",new=AsyncMock()) as generate:
            await briefing.regenerate_plan_variant(job,"A",3,None)
        analyse.assert_awaited_once()
        generate.assert_awaited_once()

    async def test_explicit_chart_request_retries_text_only_response_in_all_plans(self):
        with TemporaryDirectory() as tmp:
            settings=Settings(_env_file=None,STORAGE_DIR=Path(tmp),SKILLS_DIR=Path(__file__).resolve().parents[1]/"skills")
            table=source()
            table.chart=ChartData(chart_type=ChartType.BAR,categories=["Орёл"],series=[ChartSeries(name="Выручка",values=[7])])
            job=Job(job_id="test",brief="Отчет",purpose="Сравнить",template_path="unused",
                    must_have="Графики и диаграммы",source_tables=[table])
            client=AsyncMock(); client.__aenter__.return_value=client
            def reply(**kwargs):
                selected=table.table_id if "Исправь ответ" in kwargs["user_input"] else None
                return kwargs["output_model"].model_validate_json(json.dumps(payload(selected)),strict=True)
            client.chat_structured.side_effect=reply
            with patch.object(briefing,"_client",return_value=client),patch("app.skills.get_settings",return_value=settings):
                variants=await briefing.generate_plan_variants(job,3,None)
            self.assertEqual(set(variants),{"A","B","C"})
            self.assertEqual(client.chat_structured.await_count,6)
            self.assertTrue(all(v.plan.slides[1].content_blocks[-1].chart==table.chart for v in variants.values()))

    async def test_failed_analysis_builds_plan_with_data_gap_instead_of_refusing(self):
        # Анализ не подготовил запрошенные графики: план все равно строится (модель вызывается),
        # а нехватка данных честно попадает в заметки анализа и в указание модели не выдумывать числа.
        job=Job(job_id="test",brief="Отчет",purpose="Сравнить",template_path="unused",
                must_have="Графики и диаграммы",source_tables=[source("blocked",False)],
                table_analysis=TableAnalysisReport(status="failed"))
        self.assertEqual(briefing._unavailable_required_kinds(job),{"chart"})
        with patch.object(briefing,"_client") as client:
            client.side_effect=RuntimeError("модель недоступна в тесте")
            with self.assertRaisesRegex(RuntimeError,"модель недоступна в тесте"):
                await briefing.generate_plan_variants(job,3,None)
            self.assertTrue(client.called)
        self.assertTrue(any("не подготовил готовые результаты" in note for note in job.table_analysis.notes))
        self.assertEqual(sum("не подготовил готовые результаты" in note for note in job.table_analysis.notes),1)

    def test_explicit_table_and_chart_are_both_required_but_negative_request_is_not(self):
        job=Job(job_id="test",brief="Отчет",purpose="Сравнить",template_path="unused",source_tables=[source()],
                must_have="Таблица и график")
        self.assertEqual(briefing._required_data_kinds(job),{"table","chart"})
        job.must_have="Графики без таблиц"
        self.assertEqual(briefing._required_data_kinds(job),{"chart"})

    async def test_three_plans_after_timeout_have_no_unavailable_references(self):
        with TemporaryDirectory() as tmp:
            settings=Settings(_env_file=None,STORAGE_DIR=Path(tmp),SKILLS_DIR=Path(__file__).resolve().parents[1]/"skills")
            job=Job(job_id="test",brief="Отчет",purpose="Сравнить",template_path="unused",
                    source_tables=[source("blocked",False)],table_analysis=TableAnalysisReport(status="timeout"))
            client=AsyncMock()
            client.__aenter__.return_value=client
            def reply(**kwargs):
                self.assertIn('Разрешенные source_table_id для вставки: []',kwargs["user_input"])
                return kwargs["output_model"].model_validate_json(json.dumps(payload()),strict=True)
            client.chat_structured.side_effect=reply
            with patch.object(briefing,"_client",return_value=client),patch("app.skills.get_settings",return_value=settings):
                variants=await briefing.generate_plan_variants(job,3,None)
            self.assertEqual(set(variants),{"A","B","C"})

    async def test_timed_out_cache_is_retried_then_completed_result_is_reused(self):
        with TemporaryDirectory() as tmp:
            settings=Settings(_env_file=None,STORAGE_DIR=Path(tmp),SKILLS_DIR=Path(__file__).resolve().parents[1]/"skills")
            job=Job(job_id="test",brief="Отчет",purpose="Сравнить",template_path="unused",source_tables=[source("blocked",False)])
            def result(request,directory,settings):
                return TableAnalysisReport(key=request["key"],status="timeout" if run.await_count==1 else "completed")
            run=AsyncMock(side_effect=result)
            with patch("app.skills.get_settings",return_value=settings),patch.object(table_analysis,"_run_worker",run):
                self.assertEqual((await table_analysis.prepare_table_analysis(job,settings)).status,"timeout")
                self.assertEqual((await table_analysis.prepare_table_analysis(job,settings)).status,"completed")
                self.assertTrue((await table_analysis.prepare_table_analysis(job,settings)).cache_hit)
            self.assertEqual(run.await_count,2)
