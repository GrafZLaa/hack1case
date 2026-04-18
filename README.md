# ПАСПОРТ.ЦИФ

Веб-приложение для извлечения значимых данных из PDF/сканов паспортов оборудования, проверки качества и выгрузки в Excel для последующей загрузки в учетные системы.

Проект рассчитан на автономную работу:
- OCR + правила работают полностью локально.
- LLM-слой опциональный и также локальный (через Ollama, без облачных API).

## Ключевые возможности

### 1) Работа с документами
- Загрузка `PDF` и изображений (`png/jpg/...`).
- Очередь файлов: можно закинуть сразу несколько документов, они отображаются в перечне слева.
- Два режима запуска обработки:
  - индивидуально по выбранному файлу из очереди;
  - пакетно по всей очереди (`Запустить очередь`).
- Отложенный старт: анализ не начинается автоматически при загрузке.
- Предпросмотр всех страниц документа:
  - миниатюры;
  - переход по страницам;
  - `- / Fit / +`;
  - поворот (сохраняется по страницам).

### 2) Извлечение значимых данных
Извлекаются поля:
- наименование изделия;
- код документа (паспорт);
- код заказа;
- дата выпуска;
- дата приемки ОТК;
- гарантия;
- срок службы;
- сертификат;
- серийные номера;
- производитель;
- адрес;
- контакты;
- нормативные документы.

Классификация документа:
- `single_passport`;
- `group_passport`;
- `cabinet_list`;
- `unknown`.

### 3) UI и ручная валидация
- Режимы: `Паспорта` / `Шкаф`.
- Ручное редактирование всех ключевых полей.
- Добавление/удаление серийных номеров вручную.
- Автооценка качества (`quality_score`, `needs_review`, список пробелов).
- Генерация штрихкода `Code128`.
- Светлая/темная тема.
- Подсказки по найденным кодам (`ОКПД2`, `ОКВЭД2`, `ТН ВЭД`, `ГОСТ`, `ТУ`, `СТО`, `ТР ТС/ЕАЭС`).

### 4) Реестр и сохранение состояния
- Сохранение/восстановление состояния:
  - на backend: `data/registry_state.json`;
  - в браузере (localStorage, как fallback).
- В левой панели одновременно доступны:
  - очередь на обработку;
  - уже обработанные записи реестра.
- После перезагрузки страницы реестр восстанавливается, но запись не открывается автоматически (пользователь выбирает вручную).
- Удаление записей:
  - по одной;
  - выбранные (multi-select);
  - все.
- Сохранение многостраничных документов в реестре (`_images`) без потери страниц (настроечные лимиты в `.env`).

### 5) Выгрузка и контроль качества
- Экспорт в Excel (несколько листов).
- Контрольный прогон по эталонному набору (`samples/control_samples.json`) из UI и API.
- Метрики: `accuracy_pct`, `error_rate_pct`, `per_field`, `mismatches`.

## Архитектура обработки

1. `Preview`:
- PDF рендерится в набор PNG-страниц.

2. OCR:
- многошаговая предобработка изображения;
- multi-pass Tesseract (`OCR_PSMS`);
- эвристики и регулярные выражения по полям.

3. LLM (опционально):
- локальная Ollama-модель (vision), если `ENABLE_LLM=1`;
- при ошибке/тайм-ауте автоматический fallback на OCR+rules.

4. Постобработка:
- санитизация значений;
- фильтрация шумовых кандидатов;
- нормализация серийных номеров;
- расчет качества и чек-листа.

## Технологии

- `Flask`
- `PyMuPDF` (рендер PDF)
- `pytesseract` + `Pillow` (OCR)
- `openpyxl` (Excel)
- `python-barcode` (Code128)
- `requests` (Ollama API)

## Требования

- Python 3.10+
- Tesseract OCR (желательно `rus+eng`)
- Docker Desktop (для Docker-запуска)
- (Опционально) Ollama в контейнере, если нужен LLM-слой

## Быстрый локальный запуск (без Docker)

```bash
python -m venv .venv
# Windows
.\.venv\Scripts\activate
# Linux/macOS
# source .venv/bin/activate

pip install -r requirements.txt
cp .env.example .env  # для Windows можно создать вручную
python app/app.py
```

Открыть: `http://localhost:5000`

## Запуск через Docker

### One-click скрипты

- Windows (prod): `scripts\start.bat`
- Windows (dev hot-reload): `scripts\start-dev.bat`
- Linux/macOS (prod): `./scripts/start.sh`
- Linux/macOS (dev): `./scripts/start-dev.sh`

Остановка:
- Windows: `scripts\stop.bat`, `scripts\stop-dev.bat`
- Linux/macOS: `./scripts/stop.sh`, `./scripts/stop-dev.sh`

### Ручной запуск

```bash
docker compose -f infra/docker-compose.yml up -d --build
```

Проверка:
```bash
curl http://localhost:5000/api/meta
```

Открыть UI:
- `http://localhost:5000`

Остановка:
```bash
docker compose -f infra/docker-compose.yml down
```

Полная очистка (включая volume):
```bash
docker compose -f infra/docker-compose.yml down -v
```

## Перенос на другой компьютер

### Вариант A (рекомендуемый)

1. Скопировать папку проекта целиком.
2. Установить Docker Desktop.
3. Создать/проверить `.env` (можно из `.env.example`).
4. Запустить `scripts/start.bat` (Windows) или `./scripts/start.sh` (Linux/macOS).

Что важно:
- На первом запуске нужен интернет для загрузки Docker-образов и Ollama-модели.
- Дальше можно работать офлайн (если все уже скачано).

### Вариант B (без Docker)

- Установить Python + Tesseract,
- поставить зависимости из `requirements.txt`,
- запустить `python app/app.py`.

## Постоянство данных

В Docker-режиме:
- реестр хранится в volume `registry_data` (`/app/data` в контейнере),
- Ollama-модели — в volume `ollama_data`.

Это значит:
- простое копирование папки проекта не переносит содержимое volume;
- для полного переноса состояния/моделей нужен перенос Docker volume отдельно.

## Конфигурация `.env`

Базовые параметры:
- `ENABLE_LLM=1` — включить локальный LLM-слой.
- `OLLAMA_BASE_URL=http://localhost:11434`
- `OLLAMA_MODEL=llama3.2-vision`
- `OLLAMA_MODEL_CANDIDATES=...`

OCR:
- `OCR_LANG=rus+eng`
- `OCR_PSMS=3,6,11`
- `OCR_WORKERS=4`
- `OCR_MAX_VARIANTS=2`

Реестр:
- `REGISTRY_STATE_FILE=data/registry_state.json`
- `MAX_REGISTRY_RECORDS=2000`
- `REGISTRY_B64_MAX=12000000`
- `REGISTRY_MAX_IMAGES=1000`

Контрольный набор:
- `CONTROL_SAMPLES_FILE=samples/control_samples.json`

Если нужен режим полностью без LLM:
```env
ENABLE_LLM=0
```

## API

- `GET /` — UI
- `POST /api/preview` — подготовка предпросмотра (все страницы)
- `POST /api/extract` — извлечение данных из документа
- `POST /api/parse_cabinet` — разбор перечня шкафа
- `POST /api/barcode` — генерация Code128
- `GET /api/meta` — активная конфигурация/статус
- `GET /api/registry/load` — загрузка сохраненного реестра
- `POST /api/registry/save` — сохранение реестра
- `POST /api/registry/clear` — очистка сохраненного реестра
- `POST /api/evaluate/control` — оценка на пользовательском наборе
- `GET /api/evaluate/default` — оценка на наборе по умолчанию
- `POST /api/export/excel` — экспорт в Excel

## Структура Excel-экспорта

Формируются листы:
- `Паспорта`
- `Серийные номера`
- `Нормативы`
- `1C_Импорт`
- `Чек-лист`
- `Источники`
- `Шкаф` (если есть данные шкафа)

## Контроль качества

`/api/evaluate/default` читает `samples/control_samples.json` и возвращает:
- общую точность;
- процент ошибок;
- детализацию по полям;
- список расхождений по каждому файлу.

`/api/evaluate/control` позволяет передать собственный набор `samples` в JSON.

## Структура проекта

- `app/app.py` — backend, OCR/LLM/excel/API
- `app/templates/index.html` — frontend
- `infra/` — Dockerfile и compose
- `scripts/` — старт/стоп (prod/dev)
- `samples/control_samples.json` — эталонный набор
- `приложения/` — локальные PDF-примеры
- `data/` — runtime-данные (реестр)
- `.env.example` — шаблон конфигурации

## Ограничения

- Качество OCR зависит от качества скана.
- На очень шумных/слабо читаемых печатях часть полей требует ручной проверки.
- Для первого Docker-запуска нужен интернет.

## Обратная связь

- `su8618@mail.ru`
