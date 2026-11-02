#!/usr/bin/env python3
"""
doqa_ai_toolkit — утилита генерации тест-кейсов с помощью LLM.

Пайплайн:
  1) Экспорт тест-кейсов из doqa по ID пространства (POST /api/cases/export).
  2) Сбор требований из папки (рекурсивно, все файлы и подпапки).
  3) Запросы к OpenAI-модели:
     3.1 По требованиям -> чек-лист проверок (сохраняется в файл).
     3.2 По требованиям + чек-листу + экспортированным кейсам -> тест-кейсы
         в формате, аналогичном экспортированным (сохраняются в .xlsx/.json).

Данные подключения вынесены в .env (см. .env.example).

Использование:
    python doqa_ai_toolkit.py --space-id 123 --requirements-dir ./requirements
    python doqa_ai_toolkit.py -s 123 -r ./docs --statuses ready,draft --log-level DEBUG
"""

from __future__ import annotations

import argparse
import csv
import json
import logging
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any

import requests
from dotenv import load_dotenv

try:
    from openai import OpenAI
except ImportError:  # pragma: no cover
    OpenAI = None

log = logging.getLogger("doqa-toolkit")

# ---------------------------------------------------------------------------
# Константы: допустимые значения параметров /api/cases/export (по документации)
# ---------------------------------------------------------------------------
EXPORT_FORMATS = ("xlsx", "pdf")
EXPORT_COLUMNS = (
    "number", "id", "path", "title", "creator", "creationDate", "priority",
    "status", "keys", "description", "parameters", "preconditions",
    "postconditions", "attachments", "step", "expectedResult", "testData",
    "automationStatus", "autotest", "attributes",
)
EXPORT_STATUSES = ("ready", "draft", "archived", "broken")
EXPORT_PRIORITIES = ("low", "medium", "high")
EXPORT_SORT_FIELDS = ("id", "title", "creatorId", "createdAt", "changedAt",
                      "priority", "status", "tag")
EXPORT_DIRECTIONS = ("asc", "desc")
EXPORT_AUTOMATION_STATUSES = ("manual", "planned", "automated")
EXPORT_CREATION_SOURCES = ("manual", "autotest")

TEXT_REQ_SUFFIXES = {".md", ".txt", ".rst", ".adoc", ".html", ".htm"}
DOCX_SUFFIXES = {".docx"}
PDF_SUFFIXES = {".pdf"}


# ---------------------------------------------------------------------------
# Конфигурация (из .env / переменных окружения)
# ---------------------------------------------------------------------------
class Config:
    """Настройки утилиты. Всё, что связано с подключением, берётся из .env."""

    def __init__(self, env_file: str | None = None) -> None:
        if env_file:
            load_dotenv(env_file)
        else:
            load_dotenv()  # .env в текущей директории / окружении

        # --- doqa ---
        self.doqa_base_url = os.getenv("DOQA_BASE_URL", "").rstrip("/")
        self.doqa_api_key = os.getenv("DOQA_API_KEY", "")
        self.doqa_username = os.getenv("DOQA_USERNAME", "")
        self.doqa_password = os.getenv("DOQA_PASSWORD", "")
        self.http_timeout = int(os.getenv("HTTP_TIMEOUT", "120"))
        self.verify_ssl = os.getenv("VERIFY_SSL", "true").lower() not in ("0", "false", "no")

        # --- OpenAI ---
        self.openai_api_key = os.getenv("OPENAI_API_KEY", "")
        self.openai_base_url = os.getenv("OPENAI_BASE_URL") or None  # совместимые API
        self.openai_model = os.getenv("OPENAI_MODEL", "gpt-4o")
        self.openai_temperature = float(os.getenv("OPENAI_TEMPERATURE", "0.2"))
        self.openai_max_tokens = int(os.getenv("OPENAI_MAX_TOKENS", "8000"))

        # --- Утилита ---
        self.export_format = os.getenv("EXPORT_FORMAT", "xlsx")
        self.output_dir = os.getenv("OUTPUT_DIR", "./output")

        missing = [name for name, val in (
            ("DOQA_BASE_URL", self.doqa_base_url),
            ("OPENAI_API_KEY", self.openai_api_key),
        ) if not val]
        if missing:
            raise SystemExit(
                f"Не заданы обязательные переменные окружения: {', '.join(missing)}. "
                "Скопируйте .env.example в .env и заполните значения."
            )


# ---------------------------------------------------------------------------
# Промты для запросов к LLM
# ---------------------------------------------------------------------------
SYSTEM_PROMPT = (
    "Ты — senior QA-инженер и аналитик тестовой документации. "
    "Работаешь строго по требованиям заказчика. Отвечаешь только на русском языке. "
    "Никогда не выдумывай функциональность, которой нет в требованиях; при нехватке "
    "информации формулируй нейтральные проверки общего вида. Возвращай строго валидный JSON "
    "без markdown-обёрток и пояснений вне JSON."
)

CHECKLIST_PROMPT_TEMPLATE = """\
Ниже приведены требования к продукту (несколько документов, объединённых в один текст).

ЗАДАЧА:
Подготовь чек-лист проверок (verification checklist) по этим требованиям.

ТРЕБОВАНИЯ:
===НАЧАЛО ТРЕБОВАНИЙ===
{requirements}
===КОНЕЦ ТРЕБОВАНИЙ===

ИНСТРУКЦИИ:
1. Разбей требования на логические группы/функции.
2. Для каждой группы перечисли атомарные проверяемые пункты (один пункт = одна проверка).
3. Учитывай позитивные и негативные проверки, граничные значения, проверки прав доступа,
   валидации данных и состояний, если они явно или неявно следуют из требований.
4. Каждый пункт привяжи к источнику требования (имя файла/раздел, если он указан в тексте).

ФОРМАТ ОТВЕТА — строго JSON:
{{
  "groups": [
    {{
      "name": "Название группы проверок",
      "source": "Файл/раздел требования-источника",
      "items": [
        {{
          "id": "CL-001",
          "check": "Формулировка проверки (что именно проверяем)",
          "type": "positive|negative|boundary|security|validation|state",
          "requirement_ref": "Ссылка на фрагмент требования"
        }}
      ]
    }}
  ]
}}

Ответ должен содержать ТОЛЬКО JSON.
"""

TESTCASES_PROMPT_TEMPLATE = """\
Ты — senior QA-инженер. Нужно написать новые тест-кейсы по чек-листу проверок,
опираясь на требования и на образцы ранее созданных тест-кейсов проекта.

ТРЕБОВАНИЯ:
===НАЧАЛО ТРЕБОВАНИЙ===
{requirements}
===КОНЕЦ ТРЕБОВАНИЙ===

ЧЕК-ЛИСТ ПРОВЕРОК (JSON):
===НАЧАЛО ЧЕК-ЛИСТА===
{checklist}
===КОНЕЦ ЧЕК-ЛИСТА===

ОБРАЗЦЫ СУЩЕСТВУЮЩИХ ТЕСТ-КЕЙСОВ ПРОЕКТА (выгрузка из doqa, JSON-массив объектов):
===НАЧАЛО ОБРАЗЦОВ===
{cases_sample}
===КОНЕЦ ОБРАЗЦОВ===

СХЕМА ПОЛЕЙ КЕЙСА (ключи объектов из выгрузки):
{columns}

ИНСТРУКЦИИ:
1. Создай тест-кейсы, покрывающие ВСЕ пункты чек-листа. Один кейс покрывает один пункт
   (или несколько closely related пунктов). В поле description указывай id пункта чек-листа,
   например "Покрытие: CL-007".
2. Формат каждого кейса ДОЛЖЕН полностью совпадать со схемой полей образцов: те же ключи,
   тот же стиль формулировок, тот же язык.
3. ШАГИ (step) и ожидаемые результаты (expectedResult) конструируй по аналогии с похожими
   кейсами из образцов: используй их терминологию, навигацию, названия разделов/кнопок,
   формат нумерации шагов ({steps_hint}).
4. Если похожего кейса в образцах нет — опирайся на требования, но сохраняй стиль выгрузки.
5. Заполняй title, path, description, preconditions, step, expectedResult, priority, status.
   Остальные поля — как в образцах (можно пустые строки/null).
6. Поле "id" НЕ заполняй (проставь null) — id назначит doqa. Поле "number" — сквозная
   нумерация новых кейсов начиная с 1.
7. status: "draft"; automationStatus: "manual"; приоритеты бери из диапазона образцов
   (low/medium/high).

ФОРМАТ ОТВЕТА — строго JSON:
{{
  "cases": [
    {{ ...объект кейса со всеми ключами из СХЕМЫ ПОЛЕЙ... }}
  ]
}}

Ответ должен содержать ТОЛЬКО JSON.
"""


# ---------------------------------------------------------------------------
# Клиент doqa
# ---------------------------------------------------------------------------
class DoqaClient:
    """Минимальный клиент doqa API (экспорт тест-кейсов)."""

    def __init__(self, cfg: Config) -> None:
        self.cfg = cfg
        self.session = requests.Session()
        if cfg.doqa_api_key:
            self.session.headers["Authorization"] = f"Bearer {cfg.doqa_api_key}"
        elif cfg.doqa_username and cfg.doqa_password:
            self.session.auth = (cfg.doqa_username, cfg.doqa_password)

    def export_cases(self, payload: dict[str, Any]) -> bytes:
        """POST /api/cases/export -> байты файла выгрузки."""
        url = f"{self.cfg.doqa_base_url}/api/cases/export"
        log.info("Экспорт кейсов из doqa: POST %s", url)
        log.debug("Payload экспорта: %s", json.dumps(payload, ensure_ascii=False))

        resp = self.session.post(url, json=payload, timeout=self.cfg.http_timeout,
                                 verify=self.cfg.verify_ssl)
        if resp.status_code >= 400:
            raise RuntimeError(
                f"Ошибка экспорта doqa: HTTP {resp.status_code}\n{resp.text[:2000]}"
            )
        ctype = resp.headers.get("Content-Type", "")
        if "application/json" in ctype:
            # Некоторые конфигурации возвращают ошибку в теле 200 OK
            try:
                err = resp.json()
                raise RuntimeError(f"doqa вернула ошибку в JSON-ответе: {err}")
            except ValueError:
                pass
        log.info("Получен файл экспорта: %d байт (%s)", len(resp.content), ctype)
        return resp.content


# ---------------------------------------------------------------------------
# Парсинг выгрузки doqa (xlsx / pdf) в список кейсов-словарей
# ---------------------------------------------------------------------------
def parse_export(file_path: Path, fmt: str) -> list[dict[str, Any]]:
    if fmt == "xlsx":
        return _parse_xlsx(file_path)
    if fmt == "pdf":
        text = _extract_pdf_text(file_path)
        return [{"raw_text": text}]
    raise ValueError(f"Неизвестный формат экспорта: {fmt}")


def _parse_xlsx(file_path: Path) -> list[dict[str, Any]]:
    from openpyxl import load_workbook  # локальный импорт — зависимость опциональна

    wb = load_workbook(file_path, read_only=True, data_only=True)
    ws = wb.active
    rows = ws.iter_rows(values_only=True)
    header = next(rows, None)
    if not header:
        return []
    header = [str(h).strip() if h is not None else "" for h in header]
    cases: list[dict[str, Any]] = []
    for row in rows:
        if row is None or all(c is None or str(c).strip() == "" for c in row):
            continue
        case = {}
        for i, col in enumerate(header):
            val = row[i] if i < len(row) else None
            case[col] = val if val is not None else ""
        cases.append(case)
    wb.close()
    log.info("Разобрано xlsx: %d кейсов, колонки: %s", len(cases), header)
    return cases


def _extract_pdf_text(file_path: Path) -> str:
    try:
        from pypdf import PdfReader
    except ImportError:
        log.warning("pypdf не установлен — PDF-выгрузка будет передана как пустой текст")
        return ""
    reader = PdfReader(str(file_path))
    return "\n".join((page.extract_text() or "") for page in reader.pages)


# ---------------------------------------------------------------------------
# Сбор требований из папки (рекурсивно)
# ---------------------------------------------------------------------------
def collect_requirements(root: Path) -> tuple[str, int]:
    """Возвращает (объединённый текст требований, число файлов)."""
    parts: list[str] = []
    count = 0
    for path in sorted(root.rglob("*")):
        if not path.is_file() or path.name.startswith("."):
            continue
        suffix = path.suffix.lower()
        text = ""
        if suffix in TEXT_REQ_SUFFIXES:
            text = path.read_text(encoding="utf-8", errors="replace")
        elif suffix in DOCX_SUFFIXES:
            text = _extract_docx_text(path)
        elif suffix in PDF_SUFFIXES:
            text = _extract_pdf_text(path)
        else:
            log.debug("Пропущен файл неизвестного формата: %s", path)
            continue
        if not text.strip():
            continue
        count += 1
        parts.append(f"\n\n### ФАЙЛ: {path.relative_to(root)} ###\n{text.strip()}")
    return "".join(parts).strip(), count


def _extract_docx_text(path: Path) -> str:
    try:
        import docx  # python-docx
    except ImportError:
        log.warning("python-docx не установлен — пропускаю %s", path)
        return ""
    d = docx.Document(str(path))
    return "\n".join(p.text for p in d.paragraphs if p.text.strip())


# ---------------------------------------------------------------------------
# Работа с LLM
# ---------------------------------------------------------------------------
class LLMClient:
    def __init__(self, cfg: Config) -> None:
        if OpenAI is None:
            raise SystemExit("Не найден пакет 'openai'. Установите: pip install openai")
        kwargs: dict[str, Any] = {"api_key": cfg.openai_api_key}
        if cfg.openai_base_url:
            kwargs["base_url"] = cfg.openai_base_url
        self.client = OpenAI(**kwargs)
        self.model = cfg.openai_model
        self.temperature = cfg.openai_temperature
        self.max_tokens = cfg.openai_max_tokens

    def ask(self, prompt: str, max_retries: int = 3) -> str:
        messages = [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": prompt},
        ]
        last_err: Exception | None = None
        for attempt in range(1, max_retries + 1):
            try:
                log.info("Запрос к модели '%s' (попытка %d/%d, ~%d символов промпта)",
                         self.model, attempt, max_retries, len(prompt))
                resp = self.client.chat.completions.create(
                    model=self.model,
                    messages=messages,
                    temperature=self.temperature,
                    max_tokens=self.max_tokens,
                )
                content = (resp.choices[0].message.content or "").strip()
                if content:
                    return content
                last_err = RuntimeError("Пустой ответ модели")
            except Exception as e:  # noqa: BLE001
                last_err = e
                log.warning("Ошибка LLM-запроса: %s", e)
                time.sleep(min(2 ** attempt, 10))
        raise RuntimeError(f"LLM-запрос не удался после {max_retries} попыток: {last_err}")


def extract_json(text: str) -> Any:
    """Достаёт JSON из ответа модели (срезает ```json ... ``` и мусор вокруг)."""
    text = text.strip()
    fence = re.search(r"```(?:json)?\s*(.+?)\s*```", text, re.DOTALL)
    if fence:
        text = fence.group(1).strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass
    start = min((i for i in (text.find("{"), text.find("[")) if i != -1), default=-1)
    end = max(text.rfind("}"), text.rfind("]"))
    if start != -1 and end > start:
        return json.loads(text[start:end + 1])
    raise ValueError("Не удалось извлечь JSON из ответа модели:\n" + text[:1000])


def truncate(text: str, limit: int) -> str:
    if len(text) <= limit:
        return text
    log.warning("Текст усечён до %d символов (было %d) — увеличьте контекст модели "
                "или разбейте запуски", limit, len(text))
    return text[:limit] + "\n...[УСЕЧЕНО]..."


# ---------------------------------------------------------------------------
# Сохранение результатов
# ---------------------------------------------------------------------------
def save_checklist_md(checklist: dict[str, Any], path: Path) -> None:
    lines = ["# Чек-лист проверок",
             f"_Сгенерирован {datetime.now():%Y-%m-%d %H:%M}_", ""]
    for g in checklist.get("groups", []):
        lines.append(f"## {g.get('name', 'Группа')}")
        if g.get("source"):
            lines.append(f"_Источник: {g['source']}_")
        for item in g.get("items", []):
            lines.append(f"- [ ] **{item.get('id', '?')}** {item.get('check', '')} "
                         f"({item.get('type', '')}) — {item.get('requirement_ref', '')}")
        lines.append("")
    path.write_text("\n".join(lines), encoding="utf-8")


def save_cases_xlsx(cases: list[dict[str, Any]], columns: list[str], path: Path) -> None:
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Test Cases"
    ws.append(columns)
    for c in cases:
        ws.append([c.get(col, "") for col in columns])
    wb.save(path)


def save_cases_json(cases: list[dict[str, Any]], path: Path) -> None:
    path.write_text(json.dumps(cases, ensure_ascii=False, indent=2), encoding="utf-8")


def save_cases_csv(cases: list[dict[str, Any]], columns: list[str], path: Path) -> None:
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=columns, extrasaction="ignore")
        writer.writeheader()
        for c in cases:
            writer.writerow({k: c.get(k, "") for k in columns})


# ---------------------------------------------------------------------------
# Утилиты CLI
# ---------------------------------------------------------------------------
def parse_enum_arg(name: str, allowed: tuple[str, ...], values: str | None) -> list[str] | None:
    if not values:
        return None
    items = [v.strip() for v in values.split(",") if v.strip()]
    bad = [v for v in items if v not in allowed]
    if bad:
        raise argparse.ArgumentTypeError(
            f"--{name}: недопустимые значения {bad}. Допустимо: {list(allowed)}"
        )
    return items


def build_export_payload(args: argparse.Namespace) -> dict[str, Any]:
    payload: dict[str, Any] = {"spaceId": args.space_id, "format": args.format}
    optional = {
        "columns": args.columns, "folderIds": args.folder_ids, "itemIds": args.item_ids,
        "tagIds": args.tag_ids, "attributeIds": args.attribute_ids, "statuses": args.statuses,
        "priorities": args.priorities, "creatorIds": args.creator_ids, "search": args.search,
        "sort": args.sort, "direction": args.direction,
        "automationStatuses": args.automation_statuses,
        "creationSources": args.creation_sources,
    }
    for key, val in optional.items():
        if val:
            payload[key] = val
    if args.created_from or args.created_to:
        payload["createdAt"] = {}
        if args.created_from:
            payload["createdAt"]["from"] = args.created_from
        if args.created_to:
            payload["createdAt"]["to"] = args.created_to
    if args.changed_from or args.changed_to:
        payload["changedAt"] = {}
        if args.changed_from:
            payload["changedAt"]["from"] = args.changed_from
        if args.changed_to:
            payload["changedAt"]["to"] = args.changed_to
    return payload


def to_int_list(values: str | None) -> list[int] | None:
    if not values:
        return None
    return [int(v.strip()) for v in values.split(",") if v.strip()]


# ---------------------------------------------------------------------------
# Основной пайплайн
# ---------------------------------------------------------------------------
def run(args: argparse.Namespace) -> None:
    cfg = Config(args.env_file)
    out_dir = Path(args.output_dir or cfg.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")

    doqa = DoqaClient(cfg)
    llm = LLMClient(cfg)

    # --- Шаг 1: экспорт кейсов из doqa -------------------------------------
    fmt = args.format
    export_bytes = doqa.export_cases(build_export_payload(args))
    exported_raw = out_dir / f"doqa_export_{args.space_id}_{ts}.{fmt}"
    exported_raw.write_bytes(export_bytes)
    log.info("Выгрузка doqa сохранена: %s", exported_raw)

    cases = parse_export(exported_raw, fmt)
    if not cases:
        log.warning("В выгрузке не найдено кейсов — образцы для стиля будут пустыми")
    cases_json = out_dir / f"doqa_export_{args.space_id}_{ts}.json"
    save_cases_json(cases, cases_json)

    # Колонки для схемы (из фактической выгрузки либо из запрошенных)
    columns = args.columns_list or (list(cases[0].keys()) if cases else list(EXPORT_COLUMNS))

    # --- Шаг 2: сбор требований ---------------------------------------------
    req_root = Path(args.requirements_dir)
    if not req_root.is_dir():
        raise SystemExit(f"Папка требований не найдена: {req_root}")
    requirements_text, n_files = collect_requirements(req_root)
    if not requirements_text:
        raise SystemExit(f"В папке {req_root} не найдено читаемых файлов требований")
    log.info("Собраны требования из %d файлов (%d символов)", n_files, len(requirements_text))

    # --- Шаг 3.1: чек-лист ----------------------------------------------------
    checklist_prompt = CHECKLIST_PROMPT_TEMPLATE.format(
        requirements=truncate(requirements_text, args.max_req_chars))
    (out_dir / f"prompt_3_1_checklist_{ts}.txt").write_text(checklist_prompt, encoding="utf-8")

    checklist_raw = llm.ask(checklist_prompt)
    checklist = extract_json(checklist_raw)
    checklist_json = out_dir / f"checklist_{ts}.json"
    checklist_json.write_text(json.dumps(checklist, ensure_ascii=False, indent=2),
                              encoding="utf-8")
    checklist_md = out_dir / f"checklist_{ts}.md"
    save_checklist_md(checklist, checklist_md)
    n_items = sum(len(g.get("items", [])) for g in checklist.get("groups", []))
    log.info("Чек-лист готов: %d групп, %d пунктов -> %s",
             len(checklist.get("groups", [])), n_items, checklist_md)

    # --- Шаг 3.2: тест-кейсы ---------------------------------------------------
    samples = cases[:args.sample_limit]
    steps_hint = ("обычный текст шагов, по одному шагу на строку, нумерация '1.', '2.' ..."
                  if fmt == "pdf" else
                  "используй ровно тот же формат полей step/expectedResult, что в образцах "
                  "(переносы строк '\\n', та же нумерация)")
    tc_prompt = TESTCASES_PROMPT_TEMPLATE.format(
        requirements=truncate(requirements_text, args.max_req_chars),
        checklist=truncate(json.dumps(checklist, ensure_ascii=False), args.max_checklist_chars),
        cases_sample=truncate(json.dumps(samples, ensure_ascii=False, default=str),
                              args.max_samples_chars),
        columns=json.dumps(columns, ensure_ascii=False),
        steps_hint=steps_hint,
    )
    (out_dir / f"prompt_3_2_testcases_{ts}.txt").write_text(tc_prompt, encoding="utf-8")

    tc_raw = llm.ask(tc_prompt)
    parsed = extract_json(tc_raw)
    new_cases = parsed.get("cases", parsed if isinstance(parsed, list) else [])
    if not isinstance(new_cases, list) or not new_cases:
        raise RuntimeError("Модель вернула пустой список кейсов")

    # Нормализация: гарантируем наличие всех колонок и сброс id
    normalized: list[dict[str, Any]] = []
    for i, c in enumerate(new_cases, start=1):
        row = {col: c.get(col, "") for col in columns}
        if "number" in row:
            row["number"] = i
        if "id" in row:
            row["id"] = None
        normalized.append(row)

    gen_xlsx = out_dir / f"generated_cases_{ts}.xlsx"
    gen_json = out_dir / f"generated_cases_{ts}.json"
    gen_csv = out_dir / f"generated_cases_{ts}.csv"
    save_cases_xlsx(normalized, columns, gen_xlsx)
    save_cases_json(normalized, gen_json)
    save_cases_csv(normalized, columns, gen_csv)
    log.info("Сгенерировано %d тест-кейсов -> %s (+ .json/.csv)", len(normalized), gen_xlsx)

    print("\n=== Готово ===")
    print(f"Выгрузка doqa:        {exported_raw}")
    print(f"Чек-лист:             {checklist_md} / {checklist_json}")
    print(f"Новые тест-кейсы:     {gen_xlsx}")
    print(f"Промты LLM сохранены: {out_dir}/prompt_3_1_*.txt, prompt_3_2_*.txt")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="doqa_ai_toolkit",
        description="Экспорт кейсов из doqa + генерация чек-листа и тест-кейсов через OpenAI.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    p.add_argument("-s", "--space-id", type=int, required=True, help="ID пространства doqa")
    p.add_argument("-r", "--requirements-dir", required=True,
                   help="Папка с требованиями (рекурсивный обход)")
    p.add_argument("--format", choices=EXPORT_FORMATS, default=None,
                   help="Формат экспорта doqa (по умолчанию EXPORT_FORMAT из .env)")
    p.add_argument("--env-file", default=None, help="Путь к .env (по умолчанию ./.env)")
    p.add_argument("-o", "--output-dir", default=None, help="Каталог результатов")

    # Фильтры экспорта (см. описание /api/cases/export)
    g = p.add_argument_group("фильтры экспорта doqa")
    g.add_argument("--columns", help=f"Через запятую: {', '.join(EXPORT_COLUMNS)}")
    g.add_argument("--folder-ids", help="ID папок через запятую")
    g.add_argument("--item-ids", help="ID кейсов через запятую")
    g.add_argument("--tag-ids", help="ID тегов через запятую")
    g.add_argument("--attribute-ids", help="ID атрибутов через запятую")
    g.add_argument("--statuses", help=f"Через запятую: {', '.join(EXPORT_STATUSES)}")
    g.add_argument("--priorities", help=f"Через запятую: {', '.join(EXPORT_PRIORITIES)}")
    g.add_argument("--creator-ids", help="ID авторов через запятую")
    g.add_argument("--search", help="Строка полнотекстового поиска")
    g.add_argument("--created-from", help="createdAt.from (ISO 8601)")
    g.add_argument("--created-to", help="createdAt.to (ISO 8601)")
    g.add_argument("--changed-from", help="changedAt.from (ISO 8601)")
    g.add_argument("--changed-to", help="changedAt.to (ISO 8601)")
    g.add_argument("--sort", choices=EXPORT_SORT_FIELDS)
    g.add_argument("--direction", choices=EXPORT_DIRECTIONS)
    g.add_argument("--automation-statuses", help=f"Через запятую: {', '.join(EXPORT_AUTOMATION_STATUSES)}")
    g.add_argument("--creation-sources", help=f"Через запятую: {', '.join(EXPORT_CREATION_SOURCES)}")

    # Управление LLM-промптами
    t = p.add_argument_group("настройки LLM")
    t.add_argument("--sample-limit", type=int, default=60,
                   help="Сколько экспортированных кейсов передать модели как образцы")
    t.add_argument("--max-req-chars", type=int, default=60_000,
                   help="Максимум символов текста требований в промпте")
    t.add_argument("--max-checklist-chars", type=int, default=30_000)
    t.add_argument("--max-samples-chars", type=int, default=60_000)

    p.add_argument("--log-level", default="INFO",
                   choices=["DEBUG", "INFO", "WARNING", "ERROR"])
    return p


def main(argv: list[str] | None = None) -> None:
    parser = build_parser()
    args = parser.parse_args(argv)

    logging.basicConfig(
        level=getattr(logging, args.log_level),
        format="%(asctime)s %(levelname)-7s %(name)s: %(message)s",
    )

    # Валидация enum-параметров и приведение списков
    try:
        args.columns_list = parse_enum_arg("columns", EXPORT_COLUMNS, args.columns)
        args.statuses_list = parse_enum_arg("statuses", EXPORT_STATUSES, args.statuses)
        args.priorities_list = parse_enum_arg("priorities", EXPORT_PRIORITIES, args.priorities)
        args.automation_list = parse_enum_arg("automation-statuses",
                                              EXPORT_AUTOMATION_STATUSES, args.automation_statuses)
        args.sources_list = parse_enum_arg("creation-sources",
                                           EXPORT_CREATION_SOURCES, args.creation_sources)
    except argparse.ArgumentTypeError as e:
        parser.error(str(e))

    # Для payload используем уже разобранные значения
    args.folder_ids = to_int_list(args.folder_ids)
    args.item_ids = to_int_list(args.item_ids)
    args.tag_ids = to_int_list(args.tag_ids)
    args.attribute_ids = to_int_list(args.attribute_ids)
    args.creator_ids = to_int_list(args.creator_ids)
    args.columns_list = args.columns_list
    args.statuses = args.statuses_list
    args.priorities = args.priorities_list
    args.automation_statuses = args.automation_list
    args.creation_sources = args.sources_list

    cfg_probe_env = args.env_file
    if not args.format:
        load_dotenv(cfg_probe_env) if cfg_probe_env else load_dotenv()
        args.format = os.getenv("EXPORT_FORMAT", "xlsx")

    try:
        run(args)
    except KeyboardInterrupt:
        log.error("Прервано пользователем")
        sys.exit(130)
    except Exception as e:  # noqa: BLE001
        log.error("Ошибка: %s", e, exc_info=args.log_level == "DEBUG")
        sys.exit(1)


if __name__ == "__main__":
    main()
