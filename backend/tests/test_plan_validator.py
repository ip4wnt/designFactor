from plan_factory import THREE_BULLETS, bullets, good_plan, text, to_plan

from app.pipeline.plan_validator import format_issues, validate_plan


def validate(plan: dict, **kwargs):
    return validate_plan(to_plan(plan), **kwargs)


def codes(issues, severity=None) -> set[str]:
    return {i.code for i in issues if severity is None or i.severity == severity}


def test_good_plan_has_no_issues():
    # 5 слайдов: обложка + 3 контентных + финал; одинаковые буллеты в разных слайдах — это warning,
    # поэтому у каждого слайда свои.
    plan = good_plan()
    for i, s in enumerate(plan["slides"][1:-1], start=1):
        s["content_blocks"] = [bullets(f"Первый тезис {i} про рынок", f"Второй тезис {i} про клиентов", f"Третий тезис {i} про продукт")]
    assert validate(plan, slide_count=5, slide_count_max=5) == []


def test_title_too_long_is_error():
    plan = good_plan()
    plan["slides"][1]["title"] = "Очень длинный заголовок, который явно не помещается в лимит в шестьдесят знаков"
    assert "title_too_long" in codes(validate(plan), "error")


def test_cover_title_limit_is_30():
    plan = good_plan()
    plan["slides"][0]["title"] = "Название презентации длиннее тридцати знаков"
    assert "title_too_long" in codes(validate(plan), "error")


def test_slight_title_overshoot_is_only_a_warning():
    plan = good_plan()
    plan["slides"][0]["title"] = "Отчёт по продукту: рост и планы"  # 31 знак при лимите 30
    found = validate(plan)
    assert "title_slightly_long" in codes(found, "warning")
    assert codes(found, "error") == set()

    plan["slides"][0]["title"] = "Единый мессенджер для всей компании"  # 35 знаков — всё ещё в пределах +25%
    assert codes(validate(plan), "error") == set()

    plan["slides"][0]["title"] = "Масштабирование корпоративного мессенджера"  # 42 знака — уже ошибка
    assert "title_too_long" in codes(validate(plan), "error")


def test_content_title_soft_overshoot():
    plan = good_plan()
    plan["slides"][1]["title"] = "Выручка выросла на 23% за счёт новых клиентов и обновления тарифной сетки"  # 70 знаков при 60
    found = validate(plan)
    assert "title_slightly_long" in codes(found, "warning") and codes(found, "error") == set()


def test_bullet_limits():
    plan = good_plan()
    fifteen = " ".join(["слово"] * 15)
    sixteen = " ".join(["слово"] * 16)
    plan["slides"][1]["content_blocks"] = [bullets(fifteen, "раз", "два")]
    assert "bullet_too_long" not in codes(validate(plan))
    plan["slides"][1]["content_blocks"] = [bullets(sixteen, "раз", "два")]
    assert "bullet_too_long" in codes(validate(plan), "error")
    plan["slides"][1]["content_blocks"] = [bullets(*[f"тезис {i}" for i in range(7)])]
    assert "too_many_bullets" in codes(validate(plan), "error")


def test_table_and_chart_limits():
    plan = good_plan()
    plan["slides"][1]["content_blocks"] = [
        THREE_BULLETS,
        {"type": "table", "table": {"headers": ["a", "b"], "rows": [["1", "2"], ["3"]]}},
    ]
    assert "table_ragged" in codes(validate(plan), "error")

    rows = [[str(i), str(i)] for i in range(8)]
    plan["slides"][1]["content_blocks"] = [THREE_BULLETS, {"type": "table", "table": {"headers": ["a", "b"], "rows": rows}}]
    assert "table_rows" in codes(validate(plan), "error")

    one_category = {
        "type": "chart",
        "chart": {"chart_type": "bar", "categories": ["Янв"], "series": [{"name": "Продажи", "values": [1.0]}]},
    }
    plan["slides"][1]["content_blocks"] = [THREE_BULLETS, one_category]
    assert "chart_categories" in codes(validate(plan), "warning")


def test_chart_without_takeaways_is_warning():
    plan = good_plan()
    chart = {
        "type": "chart",
        "chart": {"chart_type": "line", "categories": ["Янв", "Фев"], "series": [{"name": "S", "values": [1.0, 2.0]}]},
    }
    plan["slides"][1]["content_blocks"] = [chart]
    assert "data_without_takeaways" in codes(validate(plan), "warning")


def test_slide_count_range_and_tolerance():
    plan = good_plan()  # 5 слайдов
    assert "slide_count" not in codes(validate(plan, slide_count=5, slide_count_max=8))
    assert "slide_count" in codes(validate(plan, slide_count=7, slide_count_max=10), "error")
    # точное число допускает промах на один слайд
    assert "slide_count" not in codes(validate(plan, slide_count=6, slide_count_max=6))
    assert "slide_count" in codes(validate(plan, slide_count=8, slide_count_max=8), "error")
    assert "slide_count" not in codes(validate(plan))


def test_dividers():
    plan = good_plan()
    plan["slides"][2]["purpose"] = "Раздел: рынок"
    plan["slides"][2]["content_blocks"] = [text("О чём эта часть")]
    assert "divider_in_short" in codes(validate(plan), "warning")

    long_plan = good_plan(content_slides=8)  # 10 слайдов, разделителей нет
    assert "divider_count" in codes(validate(long_plan), "warning")


def test_junk_text_is_error():
    plan = good_plan()
    plan["slides"][1]["content_blocks"] = [bullets("Lorem ipsum dolor sit", "два", "три")]
    assert "junk_text" in codes(validate(plan), "error")
    plan["slides"][1]["content_blocks"] = [bullets("Как ИИ-ассистент, я не могу", "два", "три")]
    assert "junk_text" in codes(validate(plan), "error")


def test_topic_title_and_duplicates_are_warnings():
    plan = good_plan()
    plan["slides"][1]["title"] = "Результаты продаж"
    assert "title_is_topic" in codes(validate(plan), "warning")

    plan = good_plan()
    plan["slides"][2]["title"] = plan["slides"][1]["title"]
    assert "duplicate_title" in codes(validate(plan), "warning")
    # одинаковые буллеты в разных слайдах
    assert "duplicate_bullet" in codes(validate(plan), "warning")


def test_conclusion_title_with_number_is_not_topic():
    plan = good_plan()
    plan["slides"][1]["title"] = "Итоги: выручка выросла на 23%"
    assert "title_is_topic" not in codes(validate(plan))


def test_empty_slide_is_error():
    plan = good_plan()
    plan["slides"][1]["content_blocks"] = []
    assert "no_content" in codes(validate(plan), "error")


def test_format_issues_names_slide_and_title():
    plan_dict = good_plan()
    plan_dict["slides"][1]["content_blocks"] = []
    plan = to_plan(plan_dict)
    rendered = format_issues(validate_plan(plan), plan)
    assert "s1" in rendered and "Вывод номер 1" in rendered
