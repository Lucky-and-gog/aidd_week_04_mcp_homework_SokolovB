#!/usr/bin/env python3
"""Генерация Excel-отчёта по задачам из API http://127.0.0.1:5000/api/tasks.

Отчёт содержит три листа:
- "Сводка": общее число задач, выполненные (шт. и %), просроченные,
  разбивка по приоритетам High/Normal/Low.
- "Просроченные": невыполненные задачи с дедлайном строго раньше сегодня,
  сортировка по дате.
- "Ближайшие": невыполненные задачи с дедлайном от сегодня до +7 дней,
  сортировка по дате.

Использует только openpyxl и стандартную библиотеку (urllib).
"""

import argparse
import json
import sys
from datetime import date, timedelta
from urllib.request import urlopen
from urllib.error import URLError, HTTPError

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

API_URL = "http://127.0.0.1:5000/api/tasks"
PRIORITIES = ["High", "Normal", "Low"]
HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF")


def fetch_tasks(url: str) -> list[dict]:
    """Загружает список задач в виде JSON из API."""
    try:
        with urlopen(url, timeout=10) as resp:
            raw = resp.read().decode("utf-8")
    except HTTPError as exc:
        sys.exit(f"Ошибка запроса к {url}: HTTP {exc.code}")
    except URLError as exc:
        sys.exit(f"Не удалось подключиться к {url}: {exc.reason}")
    except ValueError:
        sys.exit(f"Ответ от {url} не является валидным JSON")

    try:
        tasks = json.loads(raw)
    except json.JSONDecodeError:
        sys.exit(f"Некорректный JSON от {url}")
    if not isinstance(tasks, list):
        sys.exit(f"Ожидался массив задач, получен {type(tasks).__name__}")
    return tasks


def to_iso(value) -> str:
    """Приводит дату к строке YYYY-MM-DD; для None возвращает пустую строку."""
    if value is None:
        return ""
    if isinstance(value, date):
        return value.isoformat()
    return str(value)


def parse_due(raw_due) -> date | None:
    """Парсит поле due в date; пустые и некорректные значения -> None."""
    if not raw_due:
        return None
    try:
        return date.fromisoformat(str(raw_due).split("T")[0])
    except ValueError:
        return None


def main() -> None:
    parser = argparse.ArgumentParser(description="Генерация Excel-отчёта по задачам API")
    parser.add_argument("--url", default=API_URL, help="URL API задач (по умолчанию: %s)" % API_URL)
    parser.add_argument("--out", default="report_skill.xlsx", help="Путь к выходному xlsx (по умолчанию: report_skill.xlsx)")
    args = parser.parse_args()

    tasks = fetch_tasks(args.url)

    today = date.today()
    upcoming_end = today + timedelta(days=7)

    # Разделение задач по состоянию и дедлайну
    done_tasks = [t for t in tasks if t.get("done")]
    overdue = [
        t for t in tasks
        if not t.get("done") and (d := parse_due(t.get("due"))) is not None and d < today
    ]
    upcoming = [
        t for t in tasks
        if not t.get("done") and (d := parse_due(t.get("due"))) is not None and today <= d <= upcoming_end
    ]

    overdue.sort(key=lambda t: parse_due(t.get("due")))
    upcoming.sort(key=lambda t: parse_due(t.get("due")))

    # --- Лист "Сводка" ---
    wb = Workbook()
    wb.remove(wb.active)

    summary = wb.create_sheet("Сводка")
    summary.append(["Показатель", "Значение"])
    total = len(tasks)
    done_count = len(done_tasks)
    overdue_count = len(overdue)
    done_pct = (done_count / total * 100) if total else 0.0

    summary.append(["Всего задач", total])
    summary.append(["Выполнено", f"{done_count} ({done_pct:.1f}%)"])
    summary.append(["Не выполнено", total - done_count])
    summary.append(["Просрочено", overdue_count])
    summary.append([])
    summary.append(["Приоритет", "Всего", "Выполнено", "%"])

    prio_total = {p: 0 for p in PRIORITIES}
    prio_done = {p: 0 for p in PRIORITIES}
    for t in tasks:
        p = str(t.get("priority", "")).strip().capitalize()
        if p in prio_total:
            prio_total[p] += 1
            if t in done_tasks:
                prio_done[p] += 1

    for p in PRIORITIES:
        pct = (prio_done[p] / prio_total[p] * 100) if prio_total[p] else 0.0
        summary.append([p, prio_total[p], prio_done[p], f"{pct:.1f}%"])

    # --- Лист "Просроченные" ---
    overdue_sheet = wb.create_sheet("Просроченные")
    overdue_sheet.append(["ID", "Название", "Приоритет", "Дедлайн"])
    for t in overdue:
        overdue_sheet.append([
            t.get("id"),
            t.get("title"),
            str(t.get("priority", "")).strip().capitalize(),
            to_iso(parse_due(t.get("due"))),
        ])

    # --- Лист "Ближайшие" ---
    upcoming_sheet = wb.create_sheet("Ближайшие")
    upcoming_sheet.append(["ID", "Название", "Приоритет", "Дедлайн"])
    for t in upcoming:
        upcoming_sheet.append([
            t.get("id"),
            t.get("title"),
            str(t.get("priority", "")).strip().capitalize(),
            to_iso(parse_due(t.get("due"))),
        ])

    # Единый стиль заголовков для таблиц
    for sheet in (overdue_sheet, upcoming_sheet):
        for cell in sheet[1]:
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT

    wb.save(args.out)
    print(f"Отчёт сохранён: {args.out}")
    print(f"Всего задач: {total}, выполнено: {done_count}, просрочено: {overdue_count}")


if __name__ == "__main__":
    main()