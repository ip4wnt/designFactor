"""FastAPI-приложение: монолитный backend «Цифрового дизайнера презентаций»."""
from __future__ import annotations

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from app.api.routes import router
from app.config import get_settings

app = FastAPI(title="Цифровой дизайнер презентаций", version="0.1.0")

# CORS открыт на все origin для MVP/хакатона — сузить перед реальным деплоем
# (см. README, «Ограничения текущей версии»). allow_credentials=False,
# т.к. wildcard-origin с credentials браузеры всё равно отвергают.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(router)


@app.on_event("startup")
async def on_startup() -> None:
    get_settings()  # инициализирует storage-директории (templates/, outputs/)
