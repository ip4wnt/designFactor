FROM python:3.14-slim-bookworm

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    SAL_USE_VCLPLUGIN=svp \
    LLM_BASE_URL=http://host.docker.internal:11434/v1 \
    VLM_BASE_URL=http://host.docker.internal:11434/v1

RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-impress poppler-utils \
    fonts-dejavu fonts-liberation fonts-crosextra-carlito fonts-noto-core \
    libglib2.0-0 libgl1 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app/backend
COPY backend/requirements.txt ./requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

COPY backend/app ./app
COPY backend/skills ./skills
COPY frontend /app/frontend
COPY pptx_template_parser_generator/pptx_template_parser /app/pptx_template_parser_generator/pptx_template_parser
COPY pptx_template_parser_generator/prompts /app/pptx_template_parser_generator/prompts
COPY pptx_template_parser_generator/scripts /app/pptx_template_parser_generator/scripts

RUN useradd --create-home --uid 10001 app \
    && mkdir -p /app/backend/storage \
    && chown app:app /app/backend/storage

USER app

EXPOSE 8000
HEALTHCHECK --interval=15s --timeout=5s --start-period=30s --retries=3 \
    CMD python -c "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8000/openapi.json', timeout=3).close()"

CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000", "--workers", "1"]
