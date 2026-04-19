# ПАСПОРТ.ЦИФ

Веб-приложение для автономного извлечения значимых данных из PDF/сканов паспортов оборудования, проверки качества извлечения и экспорта в Excel для последующей загрузки в учетные системы (включая подготовку данных для 1С).

Проект ориентирован на задачу хакатона: главная цель — корректно извлечь данные из готовых PDF и сформировать таблицу. Интеграция с 1С напрямую в этом репозитории не реализована, но данные и формат выгрузки для этого подготовлены.

## Что делает система

1. Принимает PDF/изображения паспортов.
2. Формирует предпросмотр всех страниц.
3. Извлекает ключевые поля через OCR + правила, опционально усиливает результат локальной LLM (Ollama).
4. Позволяет оператору вручную исправить/дополнить поля.
5. Оценивает качество извлечения (`quality_score`, `needs_review`, чек-лист).
6. Сохраняет реестр между перезапусками.
7. Экспортирует реестр в многолистовый Excel.
8. Поддерживает отдельный режим анализа перечня шкафа и сверки с реестром паспортов.

## Ключевые возможности UI

### Режим `Паспорта`

- Загрузка одного или нескольких файлов (через кнопку и drag&drop).
- Очередь файлов слева:
  - выбрать конкретный файл для предпросмотра;
  - запустить обработку одного выбранного файла;
  - выбрать несколько файлов в очереди (чекбоксы);
  - `Запустить выбранные`;
  - `Удалить из очереди` (только отмеченные);
  - `Запустить очередь` (все);
  - `Очистить очередь` (все).
- Предпросмотр документа:
  - все страницы;
  - миниатюры страниц;
  - зум `- / Fit / +`;
  - поворот (хранится по страницам);
  - поддержка `Ctrl + колесо`, `Ctrl + +/-`, `Ctrl + 0`.
- Правая панель редактирования:
  - редактирование всех ключевых полей;
  - ручное добавление/удаление серийных номеров;
  - автогенерация штрихкода Code128 по основному значению;
  - индикация качества и пробелов в данных;
  - блок `Источники и риски` (по полям: страница/фрагмент + риск low/medium/high).
- Реестр:
  - сохранение записи;
  - удаление одной записи;
  - множественное удаление отмеченных;
  - удаление всех;
  - сохранение/восстановление сессии;
  - очистка сессии.

### Режим `Шкаф`

- Загрузка документа перечня шкафа.
- Выделение позиций (наименование/серийник/обозначение).
- Сверка позиций со сформированным реестром (`OK/MISS`).
- Экспорт шкафа в отдельный лист Excel вместе с реестром.

### Дополнительно

- Светлая/тёмная тема.
- Кнопка `Контроль` для контрольного прогона на эталонных примерах.
- Кнопка `Экспорт 1С JSON` (машиночитаемый payload для загрузки/валидации).
- Встроенная форма обратной связи (`/api/feedback`) с сохранением на сервере.
- Индикатор активной LLM-модели/провайдера в статус-баре.
- Контакт обратной связи в футере.

## Извлекаемые поля

- `document_type`: `single_passport | group_passport | cabinet_list | unknown`
- `tip_pasporta`: `individual | group | no_serial`
- `naimenovanie`
- `kod_dokumenta`
- `kod_zakaza`
- `data_vypuska`
- `data_priemki`
- `zavodskie_nomera[]`
- `proizvoditel`
- `adres`
- `kontakty`
- `garantia`
- `srok_sluzhby`
- `sertifikat`
- `normativnye_dok[]`
- `komplektnost[]`

Сервис также формирует служебные поля:
- `_meta` (данные о пайплайне и LLM/OCR);
- `_checklist`;
- `_evidence` (источники значений по полям);
- `quality_score`, `missing_fields`, `needs_review`.

## Архитектура обработки

### 1) Preview layer

- Для PDF рендерятся страницы (`PyMuPDF`) в PNG.
- Для изображений используется исходный файл.
- Возвращается `images[]` + `page_count`.

### 2) OCR layer

- Мультивариантная предобработка изображения.
- Multi-pass Tesseract по нескольким PSM.
- Параллельная OCR-обработка страниц (`OCR_WORKERS`).
- Эвристики оценки качества OCR, ранний выход при хорошем результате.

### 3) Rules/Regex layer

- Извлечение по лейблам и регулярным выражениям.
- Нормализация и фильтрация серийных номеров.
- Нормализация/фильтрация дат, гарантий, сроков, сертификатов.
- Отдельные правила для групповых документов и шкафов.

### 4) Optional LLM layer (локально)

- Провайдер: Ollama (`/api/generate`), без выхода в облачные API.
- Автовыбор модели из кандидатов (с кэшем и health-check).
- Ограничение по времени и failover между моделями.
- Fallback на OCR+rules при таймауте/ошибке.

### 5) Post-processing & scoring

- Слияние правил/LLM и финальная санитизация payload.
- Расчет качества записи.
- Формирование чек-листа и evidence.

## Реестр и сохранение состояния

Состояние хранится двумя способами:

1. Backend-файл: `data/registry_state.json` (главный источник).
2. Browser localStorage (fallback).

Особенности:
- Восстановление не открывает запись автоматически: пользователь выбирает вручную.
- Многостраничные документы сохраняются полностью (`_images`) в пределах лимитов.
- Сохраняются параметры просмотра (страница/зум/повороты).

## Excel-экспорт

`POST /api/export/excel` формирует файл с листами:

- `Паспорта`
- `Серийные номера`
- `Нормативы`
- `1C_Импорт`
- `1C_Ошибки` (если найдены проблемы валидации 1C-строк)
- `Чек-лист`
- `Источники`
- `Шкаф` (если есть данные шкафа)

В UI файл скачивается как `Паспорта_реестр.xlsx`.
На backend имя вложения `Registry.xlsx`.
Дополнительно доступен `POST /api/export/1c_json` для выгрузки структурированного JSON + списка ошибок валидации.

## API

### `GET /`
UI приложения.

### `POST /api/preview`
Подготовка предпросмотра файла.

- Request: `multipart/form-data`, поле `file`.
- Response:

```json
{
  "images": ["base64_png", "..."],
  "image": "base64_png_first",
  "page_count": 3
}
```

### `POST /api/extract`
Основное извлечение данных из файла.

- Request: `multipart/form-data`, поле `file`.
- Response: JSON записи (поля + служебные `_meta/_checklist/_evidence`).

### `POST /api/parse_cabinet`
Разбор перечня шкафа.

- Request: `multipart/form-data`, поле `file`.
- Response:

```json
{
  "shkaf_naim": "...",
  "shkaf_kod": "...",
  "shkaf_zav_nomer": "...",
  "pozicii": [
    {
      "nomer": 1,
      "naimenovanie": "...",
      "zavodskoy_nomer": "...",
      "oboznachenie_dok": "..."
    }
  ]
}
```

### `POST /api/barcode`
Генерация штрихкода Code128.

- Request JSON:

```json
{ "value": "TREI.421457.001" }
```

- Response JSON:

```json
{ "barcode": "base64_png", "encoded_value": "TREI-421457-001" }
```

### `GET /api/meta`
Диагностика текущей конфигурации OCR/LLM.

### `GET /api/registry/load`
Загрузка реестра из `data/registry_state.json`.

### `POST /api/registry/save`
Сохранение реестра.

- Request: объект state (или `{ state: ... }`).
- Response: `{ ok, updated_at, records_count, state }`.

### `POST /api/registry/clear`
Очистка сохраненного реестра.

### `POST /api/evaluate/control`
Контрольный прогон по переданному набору.

- Request JSON:

```json
{
  "fields": ["naimenovanie", "kod_dokumenta"],
  "samples": [
    {
      "filename": "Приложение 2 ...pdf",
      "expected": {
        "naimenovanie": "...",
        "kod_dokumenta": "..."
      }
    }
  ]
}
```

Опциональные флаги запроса:

- `fast` (`true|false`): быстрый режим (`true`), отключает LLM и тяжелый доскан даты по печати.
- `cache` (`true|false`): использовать in-memory кэш результатов extraction для повторных прогонов.

### `GET /api/evaluate/default`
Контрольный прогон по файлу по умолчанию.

Query-параметры:

- `fast` (`1|0`, default из `EVAL_FAST_DEFAULT`) — быстрый/полный режим.
- `cache` (`1|0`, default из `EVAL_USE_CACHE_DEFAULT`) — использовать кэш extraction.

В ответе дополнительно возвращаются:

- `elapsed_sec` — длительность прогона;
- `run_mode` — `fast` или `full`;
- `cache` — статистика `{ enabled, hits, misses }`.

Поиск файла идет в порядке:
1. `CONTROL_SAMPLES_FILE`
2. `samples/control_samples.json`
3. `control_samples.json`
4. `control_samples.example.json`

### `POST /api/export/excel`
Экспорт реестра в Excel.

- Request JSON:

```json
{
  "records": [ ... ],
  "cabinet": { ... }
}
```

### `POST /api/export/1c_json`
Экспорт 1С-ориентированного JSON и отчета валидации.

- Request JSON:

```json
{
  "records": [ ... ]
}
```

- Response JSON:

```json
{
  "generated_at": "2026-04-19T10:00:00Z",
  "rows_count": 128,
  "errors_count": 4,
  "rows": [ ... ],
  "validation_errors": [ ... ]
}
```

### `POST /api/feedback`
Сохранение пользовательской обратной связи в серверный лог.

- Request JSON:

```json
{
  "topic": "OCR",
  "contact": "user@example.com",
  "message": "Описание проблемы",
  "context": {
    "mode": "passport"
  }
}
```

## Формат `samples/control_samples.json`

Поддерживаются два варианта:

1) Объект:

```json
{
  "fields": ["naimenovanie", "kod_dokumenta", "zavodskie_nomera"],
  "samples": [
    {
      "filename": "Приложение 2 к задаче 1 Паспорт с одним заводским номером.pdf",
      "expected": {
        "naimenovanie": "Мастер-модуль М1201Е",
        "kod_dokumenta": "TREI.421457.001",
        "zavodskie_nomera": ["G4M0821"]
      }
    }
  ]
}
```

2) Массив `samples` без оболочки (поля будут взяты по умолчанию).

Принцип использования:
- `expected` заполняется вручную как эталон;
- контрольный прогон сравнивает только непустые поля из `expected`;
- метрика отражает качество на контрольной выборке, а не «глобальную» точность для любых документов.

## Запуск проекта

## Вариант 1 (рекомендуется): Docker

Требования:
- Docker Desktop;
- интернет на первом запуске (скачивание образов/моделей).

### Быстрый старт (prod)

- Windows: `scripts\start.bat`
- Linux/macOS: `./scripts/start.sh`

Что делает скрипт:
1. Проверяет Docker.
2. Запускает `infra/docker-compose.yml`.
3. Если LLM включен (`ENABLE_LLM != 0/false/no`) — ждет Ollama и выполняет `ollama pull` модели из `OLLAMA_MODEL`.
4. Печатает URL `http://localhost:5000`.

### Быстрый старт (dev hot-reload)

- Windows: `scripts\start-dev.bat`
- Linux/macOS: `./scripts/start-dev.sh`

Dev-режим:
- подключает `infra/docker-compose.dev.yml`;
- монтирует проект в контейнер (`../:/app`);
- запускает Flask с `--debug` и авто-перезапуском при изменении кода.

### Остановка

- prod: `scripts\stop.bat` или `./scripts/stop.sh`
- dev: `scripts\stop-dev.bat` или `./scripts/stop-dev.sh`

### Ручные команды Docker

```bash
docker compose -f infra/docker-compose.yml up -d --build
docker compose -f infra/docker-compose.yml down
docker compose -f infra/docker-compose.yml down -v
```

## Вариант 2: локально без Docker

Требования:
- Python 3.10+;
- Tesseract OCR (`rus+eng`);
- зависимости из `requirements.txt`.

Запуск:

```bash
python -m venv .venv
# Windows PowerShell
.\.venv\Scripts\Activate.ps1
# Windows CMD
# .\.venv\Scripts\activate.bat
# Linux/macOS
# source .venv/bin/activate

pip install -r requirements.txt
copy .env.example .env  # Windows
# cp .env.example .env  # Linux/macOS
python app/app.py
```

Открыть: `http://localhost:5000`

## Переменные окружения (`.env`)

Ниже полный перечень переменных, используемых приложением.

### LLM / Ollama

- `ENABLE_LLM` (default в коде: `0`): включение локальной LLM.
- `OLLAMA_BASE_URL` (default: `http://localhost:11434`): URL Ollama API.
- `OLLAMA_MODEL` (default: `qwen2.5vl:32b`): основная модель.
- `OLLAMA_MODEL_CANDIDATES` (default: `qwen2.5vl:72b,qwen2.5vl:32b,qwen2.5vl:7b,llama3.2-vision`): кандидаты failover.
- `MAX_LLM_IMAGES` (default: `4`): максимум страниц для LLM.
- `LLM_PRIMARY_MAX_IMAGES` (default: `1`): страницы в первой попытке.
- `LLM_IMAGE_MAX_SIDE` (default: `1600`): ограничение размера стороны изображения для LLM.
- `LLM_TEXT_LIMIT` (default: `14000`): лимит OCR-текста для первичного LLM-запроса.
- `LLM_RETRY_TEXT_LIMIT` (default: `7000`): лимит текста для ретрая.
- `LLM_FALLBACK_TEXT_LIMIT` (default: `3200`): лимит текста для fallback-попытки.
- `LLM_REQUEST_TIMEOUT` (default: `55`): таймаут запроса.
- `LLM_CONNECT_TIMEOUT` (default: `5`): таймаут соединения.
- `LLM_MODEL_FAILOVER_TRIES` (default: `1`): количество переходов между моделями.
- `LLM_MAX_TOTAL_SEC` (default: `90`): общий бюджет времени на LLM-этап.
- `OLLAMA_TAGS_TIMEOUT_SEC` (default: `3`): таймаут запроса тегов моделей.
- `OLLAMA_MODEL_CACHE_TTL_SEC` (default: `60`): TTL кэша списка/выбора модели.
- `OLLAMA_HEALTH_FAIL_TTL_SEC` (default: `25`): TTL cache для состояния “Ollama недоступна”.

### OCR / Render

- `TESSERACT_PATH` (Windows path по умолчанию): путь к tesseract.exe.
- `OCR_LANG` (default: `rus+eng`): языки OCR.
- `OCR_PSMS` (default: `3,6,11`): режимы Tesseract.
- `OCR_WORKERS` (default: `4`): параллельные OCR-воркеры.
- `OCR_IF_EMBEDDED_CHARS` (default: `120`): порог, ниже которого OCR выполняется даже при embedded text.
- `PDF_RENDER_SCALE` (default: `2.0`): масштаб рендера PDF для extraction.
- `PREVIEW_RENDER_SCALE` (default: `2.2`): масштаб рендера PDF для preview.
- `OCR_MAX_VARIANTS` (default: `3`): количество вариантов предобработки страницы.
- `OCR_QUALITY_SHORTCIRCUIT` (default: `320`): порог раннего выхода OCR.
- `OCR_TESSERACT_TIMEOUT_SEC` (default: `8`): timeout OCR-прохода.
- `OCR_DATE_TESSERACT_TIMEOUT_SEC` (default: `3`): timeout OCR для дат/печатей.
- `OCR_RELEASE_SCAN_BUDGET_SEC` (default: `5`): бюджет на доизвлечение даты выпуска по печатям.

### Registry / Evaluation

- `REGISTRY_STATE_FILE` (default: `data/registry_state.json`): файл реестра.
- `MAX_REGISTRY_RECORDS` (default: `2000`): максимум записей в реестре.
- `REGISTRY_B64_MAX` (default: `12000000`): лимит base64 payload.
- `REGISTRY_MAX_IMAGES` (default: `1000`): максимум сохраняемых страниц на запись.
- `CONTROL_SAMPLES_FILE` (default: `samples/control_samples.json`): путь к контрольным примерам.
- `EVAL_FAST_DEFAULT` (default: `1`): быстрый режим для `/api/evaluate/default` по умолчанию.
- `EVAL_USE_CACHE_DEFAULT` (default: `1`): кэш extraction для `/api/evaluate/default` по умолчанию.
- `EVAL_MAX_WORKERS` (default: `2`): reserved под дальнейшую распараллелку оценки.
- `FEEDBACK_FILE` (default: `data/feedback_log.jsonl`): файл журналирования обратной связи.

### Flask

- `FLASK_HOST` (default: `0.0.0.0`)
- `FLASK_PORT` (default: `5000`)
- `FLASK_DEBUG` (default: `0`)

Примечание: `.env.example` содержит рекомендованные значения для текущей демонстрации, они могут отличаться от дефолтов, зашитых в коде.

## Docker-архитектура

`infra/docker-compose.yml` поднимает:

- `app`:
  - build из `infra/Dockerfile`;
  - порт `5000:5000`;
  - volume `registry_data:/app/data`.
- `ollama`:
  - image `ollama/ollama:latest`;
  - порт `11434:11434`;
  - volume `ollama_data:/root/.ollama`.

Это обеспечивает:
- сохранение реестра между рестартами контейнеров;
- сохранение скачанных LLM-моделей.

## Перенос на другой компьютер

### Вариант A (предпочтительный)

1. Клонировать/скопировать проект.
2. Создать `.env` из `.env.example`.
3. Запустить `scripts/start.bat` (Windows) или `./scripts/start.sh` (Linux/macOS).

### Важный момент про данные

- Если переносите только папку проекта, Docker volumes (`registry_data`, `ollama_data`) не перенесутся автоматически.
- Для полного переноса истории реестра/моделей нужен отдельный экспорт volume или повторная обработка документов на новом ПК.

## Производительность и качество

Практические рекомендации:

- Для CPU-only окружения:
  - `ENABLE_LLM=0`;
  - `OCR_WORKERS` в диапазоне 2-6 (зависит от CPU);
  - `OCR_MAX_VARIANTS=2`.
- Для GPU/мощных станций с Ollama:
  - `ENABLE_LLM=1`;
  - увеличить `MAX_LLM_IMAGES` при сложных многостраничных документах.
- Если часты таймауты LLM:
  - уменьшить `MAX_LLM_IMAGES` и `LLM_TEXT_LIMIT`;
  - проверить модель в Ollama (`ollama list`, `ollama run ...`).

## Ограничения

- Очень шумные/нечитаемые сканы требуют ручной валидации.
- Документы нестандартного формата могут попадать в `unknown`.
- Штрихкод строится из ограниченного набора символов (`A-Z0-9-`) и обрезается до 20 символов при генерации.

## Troubleshooting

### `LLM extraction failed ... Read timed out`

Причина: медленная модель или перегрузка Ollama.
Что делать:
- уменьшить `MAX_LLM_IMAGES`;
- уменьшить `LLM_TEXT_LIMIT`;
- увеличить `LLM_REQUEST_TIMEOUT`/`LLM_MAX_TOTAL_SEC`;
- временно отключить LLM (`ENABLE_LLM=0`) и работать OCR-only.

### `401 Unauthorized` к OpenAI

Проект сейчас рассчитан на локальную Ollama. Для автономного режима оставьте `ENABLE_LLM=1` + `OLLAMA_BASE_URL`, без облачных ключей.

### После изменения кода в Docker не видно обновлений

Используйте dev-скрипт (`start-dev`) — там включен hot-reload и bind mount проекта.

### Не работает OCR локально без Docker

Проверьте установку Tesseract и `TESSERACT_PATH` (Windows).

## Структура проекта

- `app/app.py` — backend/API/OCR/LLM/excel/evaluate
- `app/templates/index.html` — frontend (single-page UI)
- `infra/Dockerfile` — образ приложения
- `infra/docker-compose.yml` — prod стек app + ollama
- `infra/docker-compose.dev.yml` — dev hot-reload override
- `scripts/` — start/stop скрипты для Windows/Linux
- `samples/control_samples.json` — эталонный набор для контроля
- `data/.gitkeep` — директория runtime-данных
- `приложения/` — локальные примеры документов
- `.env.example` — шаблон конфигурации
- `requirements.txt` — Python-зависимости

## Минимальная проверка после изменений

```bash
python -m py_compile app/app.py
```

Дополнительно (если запущен сервис):

```bash
curl http://localhost:5000/api/meta
```

## Обратная связь

- `su8618@mail.ru`
