# Модели

Требование ТЗ: open-weight модели до 35B параметров, OpenAI-совместимый API,
промпты и конфигурация — файлами в репозитории. Ниже — какие модели
используются, где именно в пайплайне, что нужно для их запуска и где лежат
промпты (скиллы).


## Используемые модели

| Роль | Модель на проде | Переменные `.env` | Где в пайплайне |
| --- | --- | --- | --- |
| LLM (текст → JSON) | **Qwen3.8-27B** (`qwen/qwen3.8-27b`) — плотная open-weight модель на 27B параметров, Apache-2.0, контекст 262K | `LLM_BASE_URL`, `LLM_MODEL_NAME`, `LLM_API_KEY`, `LLM_PLANNING_TIMEOUT_S`, `LLM_AUDIT_FIX_TIMEOUT_S` | план содержания (`content_planning`); структурирование тезисов и переписывание слайда (`content_structuring`); планировщик макетов движка (outline / copy / fill / repair); сокращение тезисов по аудиту (`audit_text_fix`) |
| VLM (картинка → JSON) | та же **Qwen3.8-27B** — нативная vision-language модель, отдельный экземпляр не нужен | `VLM_BASE_URL`, `VLM_MODEL_NAME`, `VLM_API_KEY`, `VLM_AUDIT_ENABLED` | 11 вопросов «Валидации контента» по PNG каждого слайда (`visual_audit`) |

Hugging Face:

- веса bf16 — <https://huggingface.co/Qwen/Qwen3.8-27B>;
- FP8 — <https://huggingface.co/Qwen/Qwen3.8-27B-FP8>;
- GGUF для llama.cpp / Ollama — <https://huggingface.co/unsloth/Qwen3.8-27B-GGUF>,
  <https://huggingface.co/bartowski/Qwen3.8-27B-GGUF>.

На проде модель подключена через OpenAI-совместимый шлюз
(`LLM_BASE_URL=https://api.alltokens.ru/api/v1`, идентификатор
`qwen/qwen3.8-27b`); тот же шлюз и модель используются как VLM. Любой другой
OpenAI-совместимый сервер (vLLM, Ollama, llama.cpp server, OpenRouter)
подключается заменой трёх переменных; локальный сервер без авторизации —
`LLM_API_KEY` пустой. Клиент — `backend/app/llm_client.py` (`openai`/`httpx`):
структурированный JSON запрашивается через `response_format`, с запасным
разбором «JSON в markdown-обёртке».

Генерация изображений не используется: картинки берутся из медиатеки
загруженного шаблона, QR-код строится детерминированно (`segno`). Все
остальные шаги — парсинг, выбор композиций, вёрстка, детерминированный аудит,
экспорт — работают без моделей и воспроизводимы.


## Области применения по шагам

| Шаг | Скилл / промпт | Модель | Вызовов на задачу | Что если модель недоступна |
| --- | --- | --- | --- | --- |
| План содержания (`content_agent`) | `backend/skills/content_planning.yaml` | LLM | 1 (повторы в пределах `LLM_PLANNING_TIMEOUT_S`) | задача завершается ошибкой `failed` — без плана собрать нечего |
| Структурирование тезисов (`dramaturgy.structure_content`) | `backend/skills/content_structuring.yaml` | LLM | ⌈слайдов / 8⌉ | детерминированный запасной вариант: тезисы плана делятся на head/text по знакам препинания |
| Планировщик макетов движка (`template_engine.plan_presentation`) | `pptx_template_parser_generator/prompts/*.txt` | LLM | 2–4 (outline, copy, при ошибках валидатора — repair) | классический путь по макетам не собирается; используется драматургия |
| Переписать слайд (`restructure_slide`, кнопка «Переписать») | `content_structuring.yaml` → `rewrite_hint` | LLM | 1 на нажатие | слайд остаётся прежним, пользователю показывается ошибка |
| Сокращение тезисов по аудиту (`audit_fix`) | `backend/skills/audit_text_fix.yaml` | LLM | 1 на замечание | обрезка по ближайшему знаку препинания / первые 6 пунктов |
| Валидация контента (`audit.audit_content_with_vlm`) | `backend/skills/visual_audit.yaml` | VLM | 1 на слайд (до 3 параллельно) | одно замечание «VLM-проверка не выполнена» на слайд; детерминированный аудит работает |


## Параметры генерации

| Параметр | Значение | Где задаётся |
| --- | --- | --- |
| temperature контент-агента | 0.1 | `skills/content_planning.yaml` |
| temperature структурирования / переписывания | 0.2 / 0.8 | `skills/content_structuring.yaml` |
| temperature планировщика макетов | 0.2 | `template_engine.BackendModel.json` |
| порция слайдов на один вызов структурирования | 8 | `skills/content_structuring.yaml` → `chunk_slides` |
| дедлайн планирования | 120 с (прод), включая повторы | `.env` → `LLM_PLANNING_TIMEOUT_S` |
| дедлайн LLM-правок аудита | 45 с | `.env` → `LLM_AUDIT_FIX_TIMEOUT_S` |
| параллельных VLM-вызовов | 3 | `audit.py` (`asyncio.Semaphore`) |
| таймаут LibreOffice | 240 с, 2 попытки | `pipeline/soffice.py` |
| dpi превью | 80 (1067×600) | `pipeline/preview.py` |
| целевой объём | 10–15 слайдов, ≤ 5 мин на три варианта | `POST /jobs` (`slide_count`, `slide_count_max`) |


## Системные требования

### Сервис (backend + frontend)

- Ubuntu 22.04+/Debian 12+ (проверено), macOS для разработки; Python 3.11+
  (проверено на 3.12).
- LibreOffice (`soffice`) и poppler (`pdftoppm`) — превью слайдов, экспорт
  pdf/html, PNG для VLM-аудита; шрифты `fonts-liberation`,
  `fonts-crosextra-carlito`, `fonts-noto-core` для метрически совместимых
  замен Arial/Calibri.
- Прод работает на 2 vCPU / 2 GB RAM без GPU: парсинг шаблона ≈ 5 с,
  планирование ≈ 60–120 с (время ответа LLM), сборка трёх вариантов ≈ 40–60 с,
  превью одного варианта через LibreOffice ≈ 10–20 с. Пик памяти —
  LibreOffice на шаблонах с крупными фото (до ≈ 1 GB), поэтому запуски
  `soffice` сериализованы.
- Сеть: доступ к OpenAI-совместимому эндпоинту LLM/VLM. Хранилище задач — в
  памяти процесса, артефакты — на диске в `STORAGE_DIR` (≈ 20–60 MB на задачу
  в зависимости от шаблона).

### Модель (если разворачивать Qwen3.8-27B у себя)

| Вариант весов | Память | Ориентир по железу |
| --- | --- | --- |
| bf16 (`Qwen/Qwen3.8-27B`) | ≈ 54 GB весов + KV-кэш | A100/H100 80 GB, или 2 × 48 GB (vLLM, tensor parallel) |
| FP8 (`Qwen/Qwen3.8-27B-FP8`) | ≈ 27 GB | одна карта 40–48 GB (L40S, A6000) под vLLM |
| GGUF Q4_K_M (`unsloth/…-GGUF`, `bartowski/…-GGUF`) | ≈ 17 GB | одна карта 24 GB (RTX 3090/4090) через llama.cpp / Ollama; медленнее, но достаточно для 10–15 слайдов |

Для VLM-аудита сервер должен принимать изображения в сообщениях
(`image_url` в OpenAI-формате) — vLLM и llama.cpp server это поддерживают;
если VLM недоступна, выставьте `VLM_AUDIT_ENABLED=false`.

Контекст задачи небольшой: самый длинный запрос — планировщик макетов с
каталогом шаблона (10–25 тысяч токенов), поэтому 32K контекста на сервере
достаточно.


## Скиллы (промпты как файлы)

Все промпты и параметры агентов лежат в `backend/skills/*.yaml` и
`pptx_template_parser_generator/prompts/*.txt`; загрузчик — `backend/app/skills.py`.
В коде нет текстов промптов — только имена скиллов и полей.

### `content_planning.yaml` (v1.2.0) — контент-стратег

Вход: бриф, назначение (кейс), число слайдов, дополнительные пожелания из
уточняющих вопросов. Выход — `ContentPlan` (`backend/app/schemas/content_plan.py`):
слайды с назначением, заголовком, тезисами, текстом, графиками (`chart` с
сериями) и таблицами (`table`). Промпт задаёт структуру: слайд 1 — обложка,
последний — финал; при 9+ слайдах 2–3 разделителя перед частями из 2+
слайдов с лидом в одно предложение; слайды ключевых цифр (буллеты начинаются
с числа); 3–5 буллетов по 8–15 слов или абзац плюс буллеты; к графику/таблице
— 2–3 вывода.

### `content_structuring.yaml` (v1.0.0) — редактор тезисов

Один вызов на 8 слайдов: каждый тезис → `head` (2–4 слова) + `text` (одно
предложение до 140 знаков) + `label` (дата/этап), метрики → `value` + `label`,
подзаголовок обложки, лид раздела, `body`, `note`. Поле `rewrite_hint`
дописывается к системному промпту при «Переписать»: другой угол подачи при том
же смысле и порядке тезисов, `rewrite_temperature: 0.8`.

### `audit_text_fix.yaml` (v1.0.0) — сокращение тезисов

`rewrite_system_prompt` — сократить каждый тезис до ≤ 14 слов, сохранив цифры
и термины; `condense_system_prompt` — объединить список до ≤ 6 пунктов.
Лимиты (`max_words_per_bullet`, `max_bullets`) совпадают с порогами аудита.

### `visual_audit.yaml` (v1.0.0) — валидация контента

Системный промпт VLM и 11 вопросов Приложения 1 (`questions[]`: `key`, `text`),
`skip_for_structural` — вопросы, которые не задаются обложке, разделителю и
финалу (q1, q3, q5, q11). Ответ — JSON `{q1..q11: bool, *_reason: str}`.
Переформулировки вопросов 4, 9 и 11 под однослайдовый вызов объяснены в
[AUDIT.md](AUDIT.md).

### Планировщик макетов — `pptx_template_parser_generator/prompts/*.txt`

| Файл | Шаг | Что делает |
| --- | --- | --- |
| `outline_system.txt` | outline | раскладывает бриф на слайды и выбирает макет из `variant_catalog` по `narrative_role` / `content_pattern` / `best_for` |
| `outline_repair_system.txt` | outline-repair | правит структуру по замечаниям валидатора (нет обложки, лишние слайды, недопустимый макет) |
| `copy_system.txt` | copy | пишет текст в слоты с учётом `recommended_characters` / `maximum_lines` |
| `fill_system.txt` | fill | дозаполняет пропущенные обязательные слоты |
| `repair_system.txt` | repair | укорачивает/переписывает текст по ошибкам `validate_generation_plan` |

Backend вызывает планировщик через `template_engine.plan_presentation`
(адаптер `BackendModel` даёт синхронный `model.json(system, user, temperature)`),
передавая relaxed-копию Template JSON (рекомендуемый бюджет символов поднят до
максимального) и пожелания пользователя. Бриф для планировщика
(`template_engine.planner_brief`) — это не промпт, а данные: исходный бриф плюс
пронумерованная структура `ContentPlan` с пометками `[ОБЛОЖКА]`,
`[РАЗДЕЛИТЕЛЬ]`, `[КЛЮЧЕВЫЕ ЦИФРЫ]`, `[ФИНАЛ]`.

### Детерминированные шаги (без моделей)

`dramaturgy.select_compositions` (выбор композиций под роль слайда и профиль
плотности), `slide_clone.adapt_units` (подгонка числа блоков), `supplement_plan_text`,
`_fit_slide_text` (shrink-to-fit не ниже 55 %), `fill_image_slots`, стилизация
графиков под фон (`styling.py`), все проверки аудита кроме VLM — подробнее в
[ARCHITECTURE.md](ARCHITECTURE.md) и [AUDIT.md](AUDIT.md).
