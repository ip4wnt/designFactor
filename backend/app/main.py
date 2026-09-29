"""FastAPI-приложение: монолитный backend «Цифрового дизайнера презентаций»."""
from __future__ import annotations

from pathlib import Path

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

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

# Интерфейс — статические файлы из ../frontend с того же origin: локальный
# запуск одной командой (uvicorn) без nginx. API-маршруты объявлены выше и
# имеют приоритет; в проде эту роль выполняет nginx (см. README).
_FRONTEND_DIR = Path(__file__).resolve().parents[2] / "frontend"
if _FRONTEND_DIR.is_dir():
    app.mount("/", StaticFiles(directory=str(_FRONTEND_DIR), html=True), name="frontend")


@app.on_event("startup")
async def on_startup() -> None:
    get_settings()  # инициализирует storage-директории (templates/, outputs/)
