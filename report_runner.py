#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Патч report_runner.py под понедельничный закрытый отчёт.

Что исправляет:
1) REPORT_AS_OF_DATE задаёт дату отчёта, например 2026-05-31.
2) REPORT_REQUIRE_DATA_DATE запрещает отправку отчёта, если последняя дата в данных меньше нужной.
3) REPORT_CLOSED_WEEK_AS_PREVIOUS=1 делает блок "Прошлая полная неделя" равным закрытой неделе REPORT_AS_OF_DATE.

Пример для понедельника 2026-06-01:
    REPORT_AS_OF_DATE=2026-05-31
    REPORT_REQUIRE_DATA_DATE=2026-05-31
    REPORT_CLOSED_WEEK_AS_PREVIOUS=1

Результат:
    Текущая неделя: 25.05-31.05.2026
    Прошлая полная неделя: 25.05-31.05.2026
    Сравнение для прошлой полной недели: 18.05-24.05.2026

Запускать из корня репозитория:
    python patch_report_runner_closed_week.py
"""
from __future__ import annotations

import py_compile
import re
from pathlib import Path

RUNNER_FILE = Path.cwd() / "report_runner.py"

DATE_BLOCK = '''    if pd.isna(latest):
        latest = pd.Timestamp.today().normalize()

    # data_latest = последняя дата, реально найденная в исходных данных.
    # latest может быть переопределён REPORT_AS_OF_DATE, но для проверки полноты
    # отчёта нужна именно фактическая последняя дата из данных.
    data_latest = pd.Timestamp(latest).normalize()

    forced_latest = os.getenv("REPORT_AS_OF_DATE", "").strip()
    if forced_latest:
        forced_ts = pd.to_datetime(forced_latest, errors="coerce")
        if pd.isna(forced_ts):
            raise ValueError(f"REPORT_AS_OF_DATE должен быть датой YYYY-MM-DD, получено: {forced_latest}")
        latest = forced_ts
        log(f"REPORT_AS_OF_DATE override: latest={pd.Timestamp(latest).normalize().date()}")

    latest = pd.Timestamp(latest).normalize()

    required_report_date_raw = os.getenv("REPORT_REQUIRE_DATA_DATE", "").strip()
    if required_report_date_raw:
        required_report_date = pd.to_datetime(required_report_date_raw, errors="coerce")
        if pd.isna(required_report_date):
            raise ValueError(
                f"REPORT_REQUIRE_DATA_DATE должен быть датой YYYY-MM-DD, получено: {required_report_date_raw}"
            )
        required_report_date = pd.Timestamp(required_report_date).normalize()
        if data_latest < required_report_date:
            raise RuntimeError(
                f"Нет данных за нужную дату {required_report_date:%d.%m.%Y}. "
                f"Последняя дата в данных: {data_latest:%d.%m.%Y}. "
                f"Отчёт не отправляю, чтобы не уйти в Telegram за старый день."
            )
        log(
            f"REPORT_REQUIRE_DATA_DATE OK: required={required_report_date.date()}, "
            f"data_latest={data_latest.date()}"
        )

    cur_start = latest - pd.Timedelta(days=int(latest.weekday()))
    cur_end = cur_start + pd.Timedelta(days=6)
    cur_actual_end = latest

    # Обычная логика:
    #   Текущая неделя = неделя latest.
    #   Прошлая полная неделя = неделя до текущей.
    #
    # Понедельничный закрытый отчёт:
    #   REPORT_AS_OF_DATE = последнее воскресенье.
    #   REPORT_CLOSED_WEEK_AS_PREVIOUS=1.
    #   Тогда "Прошлая полная неделя" должна быть этой закрытой неделей,
    #   например 25.05-31.05, а сравнение — 18.05-24.05.
    closed_week_as_previous = os.getenv("REPORT_CLOSED_WEEK_AS_PREVIOUS", "").strip().lower() in {
        "1", "true", "yes", "y", "да",
    }
    if closed_week_as_previous:
        if int(latest.weekday()) != 6:
            raise RuntimeError(
                "REPORT_CLOSED_WEEK_AS_PREVIOUS=1 можно включать только когда REPORT_AS_OF_DATE — воскресенье. "
                f"Сейчас REPORT_AS_OF_DATE/latest={latest:%d.%m.%Y}, weekday={int(latest.weekday())}."
            )
        prev_start = cur_start
        prev_end = cur_end
        prev2_start = cur_start - pd.Timedelta(days=7)
        prev2_end = cur_start - pd.Timedelta(days=1)
        log(
            f"Closed-week mode: previous full week set to {prev_start:%d.%m.%Y}-{prev_end:%d.%m.%Y}; "
            f"comparison week {prev2_start:%d.%m.%Y}-{prev2_end:%d.%m.%Y}"
        )
    else:
        prev_start = cur_start - pd.Timedelta(days=7)
        prev_end = cur_start - pd.Timedelta(days=1)
        prev2_start = cur_start - pd.Timedelta(days=14)
        prev2_end = cur_start - pd.Timedelta(days=8)

    closed_end = cur_start.replace(day=1) - pd.Timedelta(days=1)
    closed_start = closed_end.replace(day=1)
    closed_prev_end = closed_start - pd.Timedelta(days=1)
    closed_prev_start = closed_prev_end.replace(day=1)
'''

# Ищем стандартный блок дат. Регекс допускает уже добавленный ранее REPORT_AS_OF_DATE,
# но без closed-week логики.
ORIGINAL_DATE_RE = re.compile(
    r"    if pd\.isna\(latest\):\n"
    r"        latest = pd\.Timestamp\.today\(\)\.normalize\(\)\n"
    r"(?:\n?"
    r"    forced_latest = os\.getenv\(\"REPORT_AS_OF_DATE\", \"\"\)\.strip\(\)\n"
    r"    if forced_latest:\n"
    r"        forced_ts = pd\.to_datetime\(forced_latest, errors=\"coerce\"\)\n"
    r"        if pd\.isna\(forced_ts\):\n"
    r"            raise ValueError\(f\"REPORT_AS_OF_DATE должен быть датой YYYY-MM-DD, получено: \{forced_latest\}\"\)\n"
    r"        latest = forced_ts\n"
    r"(?:        log\(f\"REPORT_AS_OF_DATE override: latest=\{pd\.Timestamp\(latest\)\.normalize\(\)\.date\(\)\}\"\)\n)?"
    r")?"
    r"\n?    latest = pd\.Timestamp\(latest\)\.normalize\(\)\n"
    r"    cur_start = latest - pd\.Timedelta\(days=int\(latest\.weekday\(\)\)\)\n"
    r"    cur_end = cur_start \+ pd\.Timedelta\(days=6\)\n"
    r"    cur_actual_end = latest\n"
    r"    prev_start = cur_start - pd\.Timedelta\(days=7\)\n"
    r"    prev_end = cur_start - pd\.Timedelta\(days=1\)\n"
    r"    prev2_start = cur_start - pd\.Timedelta\(days=14\)\n"
    r"    prev2_end = cur_start - pd\.Timedelta\(days=8\)\n"
    r"    closed_end = cur_start\.replace\(day=1\) - pd\.Timedelta\(days=1\)\n"
    r"    closed_start = closed_end\.replace\(day=1\)\n"
    r"    closed_prev_end = closed_start - pd\.Timedelta\(days=1\)\n"
    r"    closed_prev_start = closed_prev_end\.replace\(day=1\)\n"
)


def main() -> None:
    if not RUNNER_FILE.exists():
        raise FileNotFoundError(f"Не найден {RUNNER_FILE}. Запусти патчер из корня репозитория.")

    text = RUNNER_FILE.read_text(encoding="utf-8")
    if "REPORT_CLOSED_WEEK_AS_PREVIOUS" in text and "REPORT_REQUIRE_DATA_DATE" in text:
        print("OK: report_runner.py уже содержит нужную closed-week логику")
        py_compile.compile(str(RUNNER_FILE), doraise=True)
        return

    new_text, n = ORIGINAL_DATE_RE.subn(DATE_BLOCK, text, count=1)
    if n != 1:
        raise RuntimeError(
            "Не смог найти блок дат в report_runner.py. "
            "Нужен участок от `if pd.isna(latest):` до `closed_prev_start = ...`."
        )

    backup = RUNNER_FILE.with_suffix(RUNNER_FILE.suffix + ".bak_closed_week")
    if not backup.exists():
        backup.write_text(text, encoding="utf-8")

    RUNNER_FILE.write_text(new_text, encoding="utf-8")
    py_compile.compile(str(RUNNER_FILE), doraise=True)
    print(f"OK: report_runner.py пропатчен. Бэкап: {backup}")
    print("OK: python compile check прошёл")
    print("Теперь в full_refresh при REPORT_CLOSED_WEEK_AS_PREVIOUS=1 блок 'Прошлая полная неделя' будет закрытой неделей REPORT_AS_OF_DATE.")


if __name__ == "__main__":
    main()
