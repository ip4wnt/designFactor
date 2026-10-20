"""На слайд не должны попадать обрубки фраз и служебные подписи планировщика."""
import unittest

from app.pipeline import dramaturgy
from app.pipeline.template_engine import _shorten, _title_overflow_allowed


class HeadOfTests(unittest.TestCase):
    def test_short_sentence_is_kept_whole(self):
        self.assertEqual(dramaturgy._head_of("Усиление слабых каналов может дать рост", 3),
                         "Усиление слабых каналов может дать рост")
        self.assertEqual(dramaturgy._head_of("Отсутствие тренда на рост.", 3), "Отсутствие тренда на рост")

    def test_long_sentence_cut_without_hanging_words(self):
        text = "Рост подтверждает эффективность текущей стратегии развития компании на всех рынках присутствия"
        head = dramaturgy._head_of(text, 4)
        self.assertEqual(head, "Рост подтверждает эффективность")
        text = "Усиление слабых каналов может обеспечить дополнительные 15 процентов выручки уже в следующем году"
        self.assertEqual(dramaturgy._head_of(text, 4), "Усиление слабых каналов")

    def test_separator_part_preferred(self):
        self.assertEqual(dramaturgy._head_of("Конверсия: стабильно низкие значения на протяжении всего квартала без тренда", 3),
                         "Конверсия")
        self.assertEqual(dramaturgy._head_of("Почта и соцсети отвечают медленно, поэтому клиенты уходят к конкурентам сразу после первого обращения", 4),
                         "Почта и соцсети отвечают медленно")

    def test_strip_hanging_keeps_nouns(self):
        self.assertEqual(dramaturgy.strip_hanging("Выручка компаний"), "Выручка компаний")
        self.assertEqual(dramaturgy.strip_hanging("Анализ роста, региональных разрывов и"), "Анализ роста, региональных разрывов")
        self.assertEqual(dramaturgy.strip_hanging("Достижение целей"), "Достижение целей")
        self.assertEqual(dramaturgy.strip_hanging("Динамика выручки и конверсия за 12"), "Динамика выручки и конверсия")
        self.assertEqual(dramaturgy.strip_hanging("Рост на 12%"), "Рост на 12%")


class ShortenTests(unittest.TestCase):
    def test_whole_sentences_preferred(self):
        text = "Рост выручки подтвержден. Без усиления слабых каналов потенциал не будет реализован. Просим утвердить бюджет."
        self.assertEqual(_shorten(text, 60), "Рост выручки подтвержден.")

    def test_clause_boundary_then_no_hanging_tail(self):
        text = "Анализ роста, региональных разрывов и эффективности каналов продаж за 2023–2025 годы"
        out = _shorten(text, 40)
        self.assertFalse(out.endswith(" и"))
        self.assertIn(out, {"Анализ роста, региональных разрывов", "Анализ роста"})
        out = _shorten("Усиление слабых каналов может дать плюс пятнадцать процентов", 30)
        self.assertEqual(out, "Усиление слабых каналов")

    def test_within_limit_untouched(self):
        self.assertEqual(_shorten("Короткая фраза", 40), "Короткая фраза")


class ServicePhraseTests(unittest.TestCase):
    def test_role_descriptions_detected(self):
        for phrase in ("Финал и призыв к действию", "Заключительный слайд", "Титульный слайд презентации", "Разделитель смысловой части"):
            self.assertTrue(dramaturgy._is_service_phrase(phrase), phrase)

    def test_real_content_kept(self):
        for phrase in ("Согласуйте бюджет до конца квартала", "Выручка выросла на 12% за год",
                       "Анализ роста, региональных разрывов и эффективности каналов продаж"):
            self.assertFalse(dramaturgy._is_service_phrase(phrase), phrase)


class OverflowAllowanceTests(unittest.TestCase):
    def test_subtitle_may_exceed_budget_up_to_factor(self):
        template = {"variants": [{"id": "v", "generation_contract": {"text_slots": [
            {"slot": "subtitle", "role": "subtitle", "constraints": {"maximum_characters": 40}}]}}]}
        plan = {"slides": [{"variant": "v", "content": {"subtitle": "x" * 60}}]}
        error = {"slide_index": 0, "slot": "subtitle"}
        self.assertTrue(_title_overflow_allowed(template, plan, error))
        plan["slides"][0]["content"]["subtitle"] = "x" * 80
        self.assertFalse(_title_overflow_allowed(template, plan, error))


if __name__ == "__main__":
    unittest.main()
