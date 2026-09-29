"""In-memory стор для TemplateAnalysis — то же осознанное упрощение MVP,
что и app/job_store.py (без БД, см. README)."""
from __future__ import annotations

from app.schemas.template_analysis import TemplateAnalysis

_analyses: dict[str, TemplateAnalysis] = {}


def save(analysis: TemplateAnalysis) -> None:
    _analyses[analysis.template_analysis_id] = analysis


def get(template_analysis_id: str) -> TemplateAnalysis | None:
    return _analyses.get(template_analysis_id)
