"""In-memory job-стор — осознанное упрощение MVP, без БД (см. README)."""
from __future__ import annotations

from app.schemas.job import Job

_jobs: dict[str, Job] = {}


def save(job: Job) -> None:
    _jobs[job.job_id] = job


def get(job_id: str) -> Job | None:
    return _jobs.get(job_id)
