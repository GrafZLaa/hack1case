# ПАСПОРТ.ЦИФ

Веб-приложение для извлечения значимых данных из PDF/сканов паспортов оборудования и экспорта в Excel.
Проект ориентирован на автономную работу (без внешних облачных API): OCR + правила + локальная Ollama (опционально).

## Что умеет

- Загружать PDF и изображения.
- Показывать предпросмотр всех страниц, масштаб, `Fit`, поворот.
- Извлекать поля паспорта:
  - наименование;
  - код документа;
  - код заказа;
  - даты;
  - серийные номера;
  - производитель, адрес, контакты;
  - гарантия/срок службы/сертификат.
- Различать типы документов:
  - `single_passport`;
  - `group_passport`;
  - `cabinet_list`;
  - `unknown`.
- Работать в fallback-режиме без LLM, если Ollama недоступна.
- Формировать источники распознавания по полям в API/Excel (страница + OCR-фрагмент).
- Генерировать Code128.
- Экспортировать реестр в Excel (включая `1C_Импорт`, `Чек-лист`, `Источники`).
- Сохранять/восстанавливать реестр между перезапусками страницы:
  - локально в браузере;
  - на backend в `data/registry_state.json`.
- Запускать контрольный прогон качества по эталонному набору (`samples/control_samples.json`) из UI и API.
- Работать полностью офлайн в интерфейсе (без внешних CDN-шрифтов).

## Стек

- `Flask`
- `PyMuPDF` (рендер PDF)
- `pytesseract` + `Pillow` (OCR и предобработка)
- `openpyxl` (Excel)
- `python-barcode` (штрихкод)
- `requests` (Ollama API)

## Требования

- Python 3.10+
- Tesseract OCR (желательно `rus+eng`)
- (Опционально) локальная Ollama для улучшения извлечения на сложных сканах

## Быстрый запуск

```bash
python -m venv .venv
# Windows
.\.venv\Scripts\activate
# Linux/macOS
# source .venv/bin/activate

pip install -r requirements.txt
python app/app.py
```

Если `.env` отсутствует, создайте его из шаблона:

```bash
cp .env.example .env
```

Открыть в браузере:

- `http://localhost:5000`

## Запуск в Docker

Самый простой вариант (one-click):

- Windows: запустить `scripts\start.bat`
- Linux/macOS:
  1. `chmod +x scripts/start.sh scripts/stop.sh`
  2. `./scripts/start.sh`

Остановка:

- Windows: `scripts\stop.bat`
- Linux/macOS: `./scripts/stop.sh`

DEV-режим (hot-reload, без постоянной пересборки):

- Windows: `scripts\start-dev.bat`
- Linux/macOS:
  1. `chmod +x scripts/start-dev.sh scripts/stop-dev.sh`
  2. `./scripts/start-dev.sh`

Остановка DEV:

- Windows: `scripts\stop-dev.bat`
- Linux/macOS: `./scripts/stop-dev.sh`

В DEV-режиме изменения в `app/app.py` и `app/templates/index.html` подхватываются автоматически.

1. Собрать и поднять сервисы:

```bash
docker compose -f infra/docker-compose.yml up -d --build
```

2. Один раз скачать модель в Ollama (по умолчанию из `.env` это `llama3.2-vision`):

```bash
docker compose -f infra/docker-compose.yml exec ollama ollama pull llama3.2-vision
```

3. Проверить, что backend отвечает:

```bash
curl http://localhost:5000/api/meta
```

4. Открыть UI:

- `http://localhost:5000`

Остановить:

```bash
docker compose -f infra/docker-compose.yml down
```

Полная очистка (включая кэш моделей Ollama):

```bash
docker compose -f infra/docker-compose.yml down -v
```

Примечание: в Docker реестр сохраняется в volume `infra_registry_data` (`/app/data` внутри контейнера).

Режим только OCR (без LLM):

- в `.env` поставить `ENABLE_LLM=0`;
- запустить снова `docker compose -f infra/docker-compose.yml up -d --build`.

## Настройки (.env)

Базовые (уже есть в проекте):

- `ENABLE_LLM=1` — включить попытки LLM через Ollama.
- `OLLAMA_BASE_URL=http://localhost:11434`
- `OLLAMA_MODEL=llama3.2-vision`
- `OLLAMA_MODEL_CANDIDATES=llama3.2-vision,moondream,llama3.2`

OCR:

- `OCR_LANG=rus+eng`
- `OCR_PSMS=3,6,11`
- `OCR_WORKERS=4`
- `OCR_MAX_VARIANTS=2`

Если нужно полностью автономно без LLM:

```env
ENABLE_LLM=0
```

Настройки хранения состояния и контрольного набора:

- `REGISTRY_STATE_FILE=data/registry_state.json`
- `CONTROL_SAMPLES_FILE=samples/control_samples.json`
- `MAX_REGISTRY_RECORDS=2000`

## API

- `POST /api/preview` — страницы для предпросмотра.
- `POST /api/extract` — извлечение значимых данных.
- `POST /api/parse_cabinet` — разбор перечня шкафа.
- `POST /api/barcode` — генерация штрихкода.
- `POST /api/export/excel` — экспорт реестра.
- `POST /api/evaluate/control` — расчет метрик точности/ошибок по контрольному набору.
- `GET /api/evaluate/default` — быстрый контрольный прогон по `samples/control_samples.json`.
- `GET /api/meta` — активная конфигурация/статус.
- `GET /api/registry/load` — загрузка сохраненного реестра.
- `POST /api/registry/save` — сохранение реестра.
- `POST /api/registry/clear` — очистка сохраненного реестра.

Пример запроса для `POST /api/evaluate/control`:

```json
{
  "samples": [
    {
      "filename": "Приложение 2 к задаче 1 Паспорт с одним заводским номером.pdf",
      "expected": {
        "document_type": "single_passport",
        "kod_dokumenta": "TREI.421457.001 ПС",
        "naimenovanie": "Мастер-модуль М1201Е",
        "zavodskie_nomera": ["G4M0821"]
      }
    }
  ]
}
```

Быстрый запуск оценки на примере:

```bash
curl -X POST http://localhost:5000/api/evaluate/control ^
  -H "Content-Type: application/json" ^
  --data "{\"samples\":[{\"filename\":\"Приложение 2 к задаче 1 Паспорт с одним заводским номером.pdf\",\"expected\":{\"document_type\":\"single_passport\",\"zavodskie_nomera\":[\"G4M0821\"]}}]}"
```

Быстрый запуск контрольного набора по умолчанию:

```bash
curl http://localhost:5000/api/evaluate/default
```

В ответе будут:

- `accuracy_pct` — точность по заданным контрольным полям.
- `error_rate_pct` — процент ошибок.
- `per_field` — детализация по каждому полю.
- `samples[].mismatches` — точные расхождения.

`/api/evaluate/control` ищет файлы по имени сначала в корне проекта, затем в папке `приложения/`.

## Критерии оценки и покрытие

Ниже свели критерии из задания и текущее покрытие в проекте.

- Точность распознавания текстовой и цифровой информации (`0-4`):
  реализовано OCR + предобработка + правила + fallback без LLM.
- Полнота/точность извлечения и визуализация источников (`0-2`):
  реализовано заполнение ключевых полей, предпросмотр страниц, ручная корректировка и показ источников распознавания по полям.
- Удобство и интерактивность интерфейса (`0-2`):
  реализовано drag-and-drop, multi-page preview, zoom/fit/rotate, ручное редактирование и сохранение в реестр.
- Процент ошибок при распознавании (`0-2`):
  реализованы санитизация, эвристики проверки, индикатор качества и endpoint оценки `POST /api/evaluate/control` с расчетом `accuracy_pct/error_rate_pct`.

Бонусные критерии:

- Возможность работы в среде 1С (`0-1`):
  прямая интеграция не включена, но реализована подготовка данных для загрузки через Excel и структурированные поля.
- Поддержка пользовательской обратной связи (`0-1`):
  добавлен канал обратной связи: `su8618@mail.ru` (в интерфейсе и документации).
- Генерация чек-листов/выгрузка для дальнейшей интеграции (`0-1`):
  реализованы Excel-выгрузка, листы `Чек-лист` и `1C_Импорт`, разбор перечня шкафа и сопоставление позиций.

## Структура проекта

- `app/app.py` — backend.
- `app/templates/index.html` — интерфейс.
- `infra/` — Docker/Compose.
- `scripts/` — скрипты запуска/остановки (prod/dev).
- `приложения/` — локальные PDF-примеры для тестов.
- `samples/control_samples.json` — эталонный контрольный набор для оценки точности.
- `data/registry_state.json` — сохраненное состояние реестра (runtime).
- `.env` — параметры работы.
- `requirements.txt` — зависимости.

## Трудности и как решали

- Низкое качество сканов: добавили несколько вариантов предобработки изображения и multi-pass OCR (`OCR_PSMS=3,6,11`).
- Нестабильность/недоступность LLM: сделали автоматический fallback на OCR+rules без остановки обработки.
- Ошибки классификации документов: усилили правила определения `single/group/cabinet/unknown` и фильтры шумовых документов.
- Ложные значения в полях: добавили санитизацию результатов (плейсхолдеры, мусорные серийники, шум в адресах/сертификатах).
- Медленная обработка: ограничили тяжелые этапы по тайм-аутам и включили ранние остановки в OCR.
- Неудобная проверка пользователем: улучшили предпросмотр страниц, зум и поворот для ручной валидации.

## Примечания

- Если Ollama недоступна, приложение автоматически продолжает обработку в OCR+rules режиме.
- Для очень плохих сканов точность зависит от качества изображения и языка OCR.
- При перезагрузке страницы реестр можно восстановить кнопкой `Восстановить` (backend + browser fallback).

## Обратная связь

- `su8618@mail.ru`
