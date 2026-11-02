# doqa_ai_toolkit

Утилита генерации тест-кейсов с помощью LLM на основе требований и существующих кейсов из **doqa**.

## Пайплайн

1. **Экспорт из doqa** — `POST /api/cases/export` по `spaceId` (все фильтры API доступны через CLI).
2. **Сбор требований** — рекурсивно читаются все файлы из папки (`md, txt, rst, adoc, html, docx, pdf`).
3. **Запросы к OpenAI-модели**:
   - **3.1** по требованиям строится **чек-лист проверок** → сохраняется в `checklist_<ts>.json` и `checklist_<ts>.md`;
   - **3.2** по требованиям + чек-листу + экспортированным кейсам генерируются **тест-кейсы**; шаги берутся по аналогии со схожими кейсами из выгрузки, формат выхода повторяет колонки экспорта → `generated_cases_<ts>.xlsx` (+ `.json`, `.csv`).

Фактические промпты, отправленные модели, сохраняются в `output/prompt_3_1_*.txt` и `output/prompt_3_2_*.txt` — их можно подправлять и переиспользовать.

## Установка

```bash
pip install -r requirements.txt
cp .env.example .env   # и заполнить значения
```

## Конфигурация (.env)

| Переменная | Описание |
|---|---|
| `DOQA_BASE_URL` | URL инстанса doqa (обязательно) |
| `DOQA_API_KEY` | Bearer-ключ API doqa |
| `DOQA_USERNAME` / `DOQA_PASSWORD` | альтернатива — basic-auth |
| `HTTP_TIMEOUT`, `VERIFY_SSL` | HTTP-настройки |
| `OPENAI_API_KEY` | ключ OpenAI (обязательно) |
| `OPENAI_BASE_URL` | опционально — совместимый endpoint (Azure и т.п.) |
| `OPENAI_MODEL` | модель (по умолчанию `gpt-4o`) |
| `OPENAI_TEMPERATURE`, `OPENAI_MAX_TOKENS` | параметры генерации |
| `EXPORT_FORMAT` | `xlsx` \| `pdf` |
| `OUTPUT_DIR` | каталог результатов |

## Запуск

```bash
python doqa_ai_toolkit.py --space-id 123 --requirements-dir ./requirements

# с фильтрами экспорта doqa:
python doqa_ai_toolkit.py -s 123 -r ./docs \
    --statuses ready,draft --priorities high,medium \
    --folder-ids 10,11 --sort createdAt --direction desc \
    --columns title,preconditions,step,expectedResult,priority,status \
    --sample-limit 80 --log-level DEBUG
```

Все параметры `/api/cases/export` (`folderIds, itemIds, tagIds, attributeIds, statuses,
priorities, creatorIds, search, changedAt, createdAt, sort, direction, automationStatuses,
creationSources, columns, format`) маппятся на одноимённые флаги CLI; enum-значения валидируются.

## Промпты

### Системный промпт (оба запроса)
> Ты — senior QA-инженер и аналитик тестовой документации. Работаешь строго по требованиям заказчика. Отвечаешь только на русском языке. Никогда не выдумывай функциональность, которой нет в требованиях... Возвращай строго валидный JSON без markdown-обёрток.

### 3.1 — Чек-лист (`CHECKLIST_PROMPT_TEMPLATE`)
Вход: полный текст требований (с метками источников `### ФАЙЛ: ...`).
Задание: разбить на группы функций, сформулировать атомарные проверки (positive/negative/boundary/security/validation/state), привязать каждую к фрагменту требования.
Выход: JSON `{groups: [{name, source, items: [{id, check, type, requirement_ref}]}]}`.

### 3.2 — Тест-кейсы (`TESTCASES_PROMPT_TEMPLATE`)
Вход: требования + чек-лист (JSON) + образцы экспортированных кейсов (JSON) + схема колонок выгрузки.
Задание: покрыть все пункты чек-листа кейсами; **шаги и ожидаемые результаты конструировать по аналогии со схожими кейсами из образцов** (терминология, навигация, нумерация); поля и стиль — как в выгрузке; `id=null`, `status=draft`, ссылка на пункт чек-листа в `description` («Покрытие: CL-007»).
Выход: JSON `{cases: [{...полная схема колонок...}]}`.

## Результат в `output/`

```
doqa_export_<space>_<ts>.xlsx|.pdf   # сырая выгрузка doqa
doqa_export_<space>_<ts>.json        # та же выгрузка, распарсенная в структурированный JSON
checklist_<ts>.json / .md            # шаг 3.1
generated_cases_<ts>.xlsx/.json/.csv # шаг 3.2 — новые кейсы в формате экспорта
prompt_3_1_checklist_<ts>.txt        # фактический промпт 3.1
prompt_3_2_testcases_<ts>.txt        # фактический промпт 3.2
```

## Ограничения / примечания

- Для очень больших объёмов требований/кейсов текст ускается до лимитов (`--max-req-chars`, `--max-samples-chars`); при необходимости запускайте по подпапкам/папкам doqa отдельно.
- PDF-выгрузка парсится как plain text (структура кейсов теряется) — для точного соответствия формата используйте `xlsx`.
- Сгенерированные `.xlsx/.json` готовы для последующего импорта; импорт через API в этой утилите не выполняется (можно добавить `POST /api/cases/import` по аналогии с экспортом).
