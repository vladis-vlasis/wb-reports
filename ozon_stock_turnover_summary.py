#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Отдельный сводный отчёт Ozon FBO: остатки, оборачиваемость, активные поставки и прогноз OOS.

Это САМОСТОЯТЕЛЬНЫЙ скрипт. Он не запускает рекламу, финансы, отзывы и прочие
модули большого сборщика.

Главный результат в Object Storage:
  Сводные отчёты/<STORE>/Остатки_и_оборачиваемость_YYYY-MM-DD.xlsx
  Сводные отчёты/<STORE>/Последний.xlsx

Что считает:
- среднесуточные продажи за 7 дней по артикулу и по кластеру;
- доступный остаток FBO сейчас;
- товары в пути внутри Ozon и возвраты покупателей на склад;
- активные FBO-заявки продавца через рабочую связку /v3/supply-order/list -> /v3/supply-order/get;
- состав активных поставок через /v1/supply-order/bundle;
- дату отгрузки из timeslot поставки;
- плановую дату прибытия: сначала из API, если она есть, иначе дата отгрузки + норматив кросс-дока;
- самый свежий процент выкупа по артикулу: последние 50 завершённых единиц (delivered/cancelled), ClientReturn учитывается как невыкуп;
- запас в днях сейчас и с учётом пути/активных заявок;
- прогноз дефицита по кластерам.

Для FINICK расчёт кросс-дока использует точку «Москва — Кавказский» и верхнюю
границу сроков из переданного тарифного файла Ozon от 03.08.2026. Нормативы
встроены в код, поэтому внешний сайт/Яндекс Диск для расчёта не требуются.
"""
from __future__ import annotations

import argparse
import io
import json
import logging
import math
import os
import re
import sys
import time
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from difflib import SequenceMatcher
from typing import Any, Dict, Iterable, List, Mapping, Optional, Sequence, Tuple
from zoneinfo import ZoneInfo

import boto3
import pandas as pd
import requests
from botocore.client import Config
from botocore.exceptions import ClientError
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

SCRIPT_VERSION = "OZON_STOCK_TURNOVER_SUMMARY_STANDALONE_V1_8_20260807"
MOSCOW_TZ = ZoneInfo("Europe/Moscow")
OZON_API_BASE = "https://api-seller.ozon.ru"
DEFAULT_BUCKET = "ozon-assist"
SELLER_MIN_INTERVAL_SECONDS = 0.55

# FINICK отгружается из Москвы. Для текущего рабочего контура используем
# точку Москва — Кавказский, как и в предыдущем расчёте.
CROSSDOCK_ORIGIN = "Москва — Кавказский"
CROSSDOCK_TARIFF_AS_OF = "2026-08-03"
CROSSDOCK_TARIFF_SOURCE = "Тарифы Ozon cross-dock, файл от 03.08.2026, Москва — Кавказский"
PREPARATION_DAYS = 4

# Выкуп: считаем по самым свежим завершённым единицам товара.
# Берём до 50 последних delivered/cancelled. Если за окно истории набралось меньше,
# считаем по фактически доступной базе и явно показываем её размер в отчёте.
BUYOUT_TARGET_UNITS = 50
BUYOUT_LOOKBACK_DAYS = 60

# Верхняя граница планового срока доставки из переданного тарифного XLSX.
# Сначала пытаемся сопоставить конкретный конечный склад, а если API отдаёт
# только макролокальный кластер — используем консервативный максимум по кластеру.
CROSSDOCK_WAREHOUSE_DAYS: Dict[str, int] = {
    "Адыгейск РФЦ": 6,
    "Алматы 2 РФЦ": 19,
    "Астана РФЦ": 17,
    "Волгоград МРФЦ": 6,
    "Воронеж МРФЦ": 6,
    "Воронеж Негабарит РФЦ": 5,
    "Воронеж-2 РФЦ": 5,
    "Гривно РФЦ": 3,
    "Гривно РФЦ Негабарит": 3,
    "Давыдовское РФЦ Негабарит": 3,
    "Домодедово РФЦ": 5,
    "Екатеринбург РФЦ": 8,
    "Екатеринбург РФЦ Негабарит": 8,
    "Жуковский РФЦ": 3,
    "Казань РФЦ": 6,
    "Казань Столбище РФЦ Негабарит": 6,
    "Калининград МРФЦ": 32,
    "Красноярск МРФЦ": 24,
    "Махачкала РФЦ": 14,
    "Минск МПСЦ": 11,
    "Минск МРФЦ Негабарит": 11,
    "Невинномысск РФЦ": 14,
    "Нижний Новгород РФЦ": 5,
    "Нижний Новгород-2 РФЦ": 5,
    "Новороссийск МРФЦ": 6,
    "Новосибирск РФЦ": 24,
    "Ногинск РФЦ": 5,
    "Ногинск РФЦ Негабарит": 5,
    "Омск РФЦ": 24,
    "Оренбург РФЦ": 6,
    "Павловская Слобода РФЦ Негабарит": 3,
    "Пермь РФЦ": 8,
    "Петровское РФЦ": 3,
    "Пушкино-1 РФЦ": 3,
    "Пушкино-2 РФЦ": 3,
    "Радумля РФЦ Негабарит": 3,
    "Ростов-На-Дону РФЦ": 6,
    "СПБ Бугры РФЦ": 5,
    "СПБ Волхонка РФЦ Негабарит": 5,
    "СПБ Колпино РФЦ": 5,
    "СПБ Шушары РФЦ": 5,
    "СПБ Шушары РФЦ Негабарит": 5,
    "Самара РФЦ": 6,
    "Самара РФЦ Негабарит": 6,
    "Санкт-Петербург РФЦ": 5,
    "Саратов РФЦ": 6,
    "Софьино РФЦ": 3,
    "Тверь РФЦ": 5,
    "Тюмень РФЦ": 8,
    "Уфа РФЦ": 6,
    "Хабаровск-2 РФЦ": 44,
    "Хоругвино РФЦ": 3,
    "Хоругвино РФЦ Негабарит": 5,
    "Южный Обход РФЦ Негабарит": 13,
    "Ярославль РФЦ": 5,
}

CROSSDOCK_CLUSTER_DAYS: Dict[str, int] = {
    "Алматы": 19,
    "Астана": 17,
    "Беларусь": 11,
    "Воронеж": 6,
    "Дальний Восток": 44,
    "Екатеринбург": 8,
    "Казань": 6,
    "Калининград": 32,
    "Краснодар": 13,
    "Красноярск": 24,
    "Махачкала": 14,
    "Москва, МО и Дальние регионы": 5,
    "Невинномысск": 14,
    "Новосибирск": 24,
    "Омск": 24,
    "Оренбург": 6,
    "Пермь": 8,
    "Ростов": 6,
    "Самара": 6,
    "Санкт-Петербург и СЗО": 5,
    "Саратов": 6,
    "Тверь": 5,
    "Тюмень": 8,
    "Уфа": 6,
    "Ярославль": 5,
}

TERMINAL_SUPPLY_STATES = {
    "COMPLETED", "CANCELLED", "CANCELED", "OVERDUE", "REJECTED", "ARCHIVED",
    "REJECTED_AT_SUPPLY_WAREHOUSE", "REPORTS_CONFIRMATION_AWAITING", "REPORT_REJECTED"
}

# Резервное сопоставление физического FBO-склада с кластером. Оно применяется
# только когда API не вернул кластер. Название склада НИКОГДА не становится
# отдельным «кластером» само по себе.
STATIC_WAREHOUSE_CLUSTER_PATTERNS = [
    (("ПЕТРОВСК", "ПУШКИНО", "ДОМОДЕД", "ХОРУГВ", "ЖУКОВСК", "СОФЬИНО", "ИСТРА", "ГРИВНО", "ПАВЛОВСК"), "Москва, МО и Дальние регионы"),
    (("СПБ", "САНКТ", "КОЛПИНО", "ШУШАР", "БУГРЫ"), "Санкт-Петербург и СЗО"),
    (("РОСТОВ",), "Ростов"),
    (("КРАСНОДАР", "НОВОРОССИЙСК", "АДЫГЕЙСК"), "Краснодар"),
    (("КРАСНОЯРСК",), "Красноярск"),
    (("НЕВИННОМЫССК",), "Невинномысск"),
    (("КАЛИНИНГРАД",), "Калининград"),
    (("ПЕРМ",), "Пермь"),
    (("САМАР",), "Самара"),
    (("ЕКАТЕРИНБУРГ",), "Екатеринбург"),
    (("КАЗАН",), "Казань"),
    (("МАХАЧКАЛ",), "Махачкала"),
    (("НОВОСИБИР",), "Новосибирск"),
    (("САРАТОВ",), "Саратов"),
    (("ХАБАРОВСК", "ВЛАДИВОСТОК"), "Дальний Восток"),
    (("ОМСК",), "Омск"),
    (("ТВЕР",), "Тверь"),
    (("ТЮМЕН",), "Тюмень"),
    (("УФА",), "Уфа"),
    (("ЯРОСЛАВ",), "Ярославль"),
    (("ВОЛГОГРАД",), "Волгоград"),
    (("ВОРОНЕЖ",), "Воронеж"),
    (("НИЖНИЙ", "НОВГОРОД"), "Нижний Новгород"),
    (("ОРЕНБУРГ",), "Оренбург"),
    (("АСТАН",), "Казахстан"),
    (("ЕРЕВАН",), "Армения"),
]


ACTIVE_STATES = {
    "DATA_FILLING",
    "READY_TO_SUPPLY",
    "ACCEPTED_AT_SUPPLY_WAREHOUSE",
    "IN_TRANSIT",
    "ACCEPTANCE_AT_STORAGE_WAREHOUSE",
}
PHYSICAL_STATES = {
    "ACCEPTED_AT_SUPPLY_WAREHOUSE",
    "IN_TRANSIT",
    "ACCEPTANCE_AT_STORAGE_WAREHOUSE",
}
PRE_HANDOFF_STATES = {"DATA_FILLING", "READY_TO_SUPPLY"}
STATUS_RU = {
    "DATA_FILLING": "Заполняются данные",
    "READY_TO_SUPPLY": "Готова к отгрузке",
    "ACCEPTED_AT_SUPPLY_WAREHOUSE": "Принята в точке отгрузки",
    "IN_TRANSIT": "В пути",
    "ACCEPTANCE_AT_STORAGE_WAREHOUSE": "Приёмка на складе назначения",
    "REPORTS_CONFIRMATION_AWAITING": "Ожидается подтверждение актов",
    "REPORT_REJECTED": "Спор по акту",
    "COMPLETED": "Завершена",
    "CANCELLED": "Отменена",
    "OVERDUE": "Просрочена",
}


def load_report_env() -> None:
    raw = os.getenv("REPORT_ENV", "") or ""
    for line in raw.splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        if line.startswith("export "):
            line = line[7:].strip()
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


def parse_date(value: str) -> date:
    return datetime.strptime(value.strip(), "%Y-%m-%d").date()


def resolve_target_date(value: str = "") -> date:
    forced = (value or os.getenv("OZON_TARGET_DATE", "") or "").strip()
    if forced:
        return parse_date(forced)
    return datetime.now(MOSCOW_TZ).date() - timedelta(days=1)


def to_ozon_datetime(d: date, end: bool = False) -> str:
    return f"{d.isoformat()}T{'23:59:59.999Z' if end else '00:00:00.000Z'}"


def norm_id(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    s = str(value).strip()
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s


def num(value: Any, default: float = 0.0) -> float:
    if value is None:
        return default
    try:
        if isinstance(value, str):
            value = value.replace("\xa0", "").replace(" ", "").replace(",", ".")
        x = float(value)
        return x if math.isfinite(x) else default
    except Exception:
        return default


def int_qty(value: Any) -> int:
    return max(0, int(round(num(value))))


def extract_items(data: Any, paths: Sequence[Sequence[str]]) -> List[Dict[str, Any]]:
    for path in paths:
        cur = data
        ok = True
        for key in path:
            if not isinstance(cur, Mapping) or key not in cur:
                ok = False
                break
            cur = cur[key]
        if ok and isinstance(cur, list):
            return [dict(x) for x in cur if isinstance(x, Mapping)]
    if isinstance(data, list):
        return [dict(x) for x in data if isinstance(x, Mapping)]
    return []


def extract_cursor(data: Any) -> str:
    for path in (("result", "cursor"), ("cursor",), ("result", "last_id"), ("last_id",),
                 ("result", "next_cursor"), ("next_cursor",)):
        cur = data
        ok = True
        for key in path:
            if not isinstance(cur, Mapping) or key not in cur:
                ok = False
                break
            cur = cur[key]
        if ok and cur not in (None, ""):
            return str(cur)
    return ""


def extract_total(data: Any) -> Optional[int]:
    for path in (("result", "total"), ("total",), ("result", "total_count"), ("total_count",)):
        cur = data
        ok = True
        for key in path:
            if not isinstance(cur, Mapping) or key not in cur:
                ok = False
                break
            cur = cur[key]
        if ok:
            try:
                return int(cur)
            except Exception:
                pass
    return None


def walk_dicts(obj: Any) -> Iterable[Mapping[str, Any]]:
    if isinstance(obj, Mapping):
        yield obj
        for v in obj.values():
            yield from walk_dicts(v)
    elif isinstance(obj, list):
        for v in obj:
            yield from walk_dicts(v)


def deep_first(obj: Any, keys: Sequence[str]) -> Any:
    keyset = {k.lower() for k in keys}
    for d in walk_dicts(obj):
        for k, v in d.items():
            if str(k).lower() in keyset and v not in (None, ""):
                return v
    return None


def parse_any_date(value: Any) -> Optional[date]:
    if value in (None, ""):
        return None
    text = str(value).strip()
    if not text:
        return None
    for candidate in (text[:10], text):
        try:
            return datetime.fromisoformat(candidate.replace("Z", "+00:00")).date()
        except Exception:
            pass
        for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y"):
            try:
                return datetime.strptime(candidate, fmt).date()
            except Exception:
                pass
    return None


def norm_name(value: Any) -> str:
    s = str(value or "").lower().replace("ё", "е")
    s = re.sub(r"[^a-zа-я0-9]+", " ", s)
    stop = {"кластер", "макролокальный", "макрорегион", "регион", "склад", "озон", "ozon"}
    return " ".join(x for x in s.split() if x not in stop).strip()


def similarity(a: str, b: str) -> float:
    a, b = norm_name(a), norm_name(b)
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    if a in b or b in a:
        return 0.92
    ta, tb = set(a.split()), set(b.split())
    jaccard = len(ta & tb) / max(1, len(ta | tb))
    return max(jaccard, SequenceMatcher(None, a, b).ratio())


def conservative_days(value: Any) -> Optional[int]:
    """Берём верхнюю границу срока: '2–4' -> 4, 'до 7 дней' -> 7."""
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return None
    if isinstance(value, (int, float)):
        x = int(math.ceil(float(value)))
        return x if 0 <= x <= 60 else None
    text = str(value).lower().replace("−", "-").replace("–", "-").replace("—", "-")
    nums = [int(x) for x in re.findall(r"(?<!\d)(\d{1,2})(?!\d)", text)]
    nums = [x for x in nums if 0 <= x <= 60]
    return max(nums) if nums else None


def static_cluster_from_warehouse(value: Any) -> str:
    raw = str(value or "").upper().replace("Ё", "Е")
    raw = re.sub(r"[^A-ZА-Я0-9]+", "_", raw)
    for tokens, cluster in STATIC_WAREHOUSE_CLUSTER_PATTERNS:
        # Для Нижнего Новгорода требуем оба слова, для остальных достаточно одного маркера.
        if tokens == ("НИЖНИЙ", "НОВГОРОД"):
            if all(token in raw for token in tokens):
                return cluster
        elif any(token in raw for token in tokens):
            return cluster
    return ""


def add_workdays(start: date, workdays: int) -> date:
    result = start
    added = 0
    while added < max(0, int(workdays)):
        result += timedelta(days=1)
        if result.weekday() < 5:
            added += 1
    return result


class OzonApiError(RuntimeError):
    def __init__(self, path: str, status: int, message: str):
        super().__init__(f"POST {path}: HTTP {status}: {message}")
        self.path = path
        self.status = status


class OzonClient:
    def __init__(self, client_id: str, api_key: str):
        self.session = requests.Session()
        self.session.headers.update({
            "Client-Id": str(client_id),
            "Api-Key": str(api_key),
            "Content-Type": "application/json",
            "Accept": "application/json",
            "User-Agent": f"ozon-assist/{SCRIPT_VERSION}",
        })
        self.last_at = 0.0

    def _throttle(self) -> None:
        wait = SELLER_MIN_INTERVAL_SECONDS - (time.monotonic() - self.last_at)
        if wait > 0:
            time.sleep(wait)
        self.last_at = time.monotonic()

    def post(self, path: str, payload: Mapping[str, Any], retries: int = 6) -> Dict[str, Any]:
        url = OZON_API_BASE + path
        for attempt in range(1, retries + 1):
            try:
                self._throttle()
                r = self.session.post(url, json=dict(payload), timeout=180)
            except requests.RequestException as exc:
                if attempt == retries:
                    raise OzonApiError(path, 0, str(exc)) from exc
                time.sleep(min(60, 2 ** attempt))
                continue
            if r.status_code == 200:
                try:
                    return r.json()
                except Exception as exc:
                    raise OzonApiError(path, 200, "Ответ не JSON") from exc
            if r.status_code in {429, 500, 502, 503, 504} and attempt < retries:
                try:
                    sec = int(float(r.headers.get("Retry-After", "")))
                except Exception:
                    sec = min(60, 2 ** attempt)
                logging.warning("%s: HTTP %s, повтор через %s сек", path, r.status_code, max(1, sec))
                time.sleep(max(1, sec))
                continue
            try:
                body = r.json()
                msg = body.get("message") or body.get("error") or body.get("code") or r.text[:1000]
            except Exception:
                msg = r.text[:1000]
            raise OzonApiError(path, r.status_code, str(msg))
        raise OzonApiError(path, 0, "Неизвестная ошибка")

    def offset_pages(self, path: str, payload: Dict[str, Any], paths: Sequence[Sequence[str]],
                     limit: int = 1000, max_pages: int = 200) -> List[Dict[str, Any]]:
        rows: List[Dict[str, Any]] = []
        offset = int(payload.get("offset", 0) or 0)
        for _ in range(max_pages):
            body = dict(payload)
            body["limit"] = min(limit, int(body.get("limit", limit)))
            body["offset"] = offset
            data = self.post(path, body)
            got = extract_items(data, paths)
            rows.extend(got)
            total = extract_total(data)
            if not got or len(got) < body["limit"] or (total is not None and len(rows) >= total):
                break
            offset += len(got)
        return rows

    def cursor_pages(self, path: str, payload: Dict[str, Any], paths: Sequence[Sequence[str]],
                     cursor_field: str = "cursor", limit: int = 1000, max_pages: int = 200) -> List[Dict[str, Any]]:
        rows: List[Dict[str, Any]] = []
        cursor = str(payload.get(cursor_field, "") or "")
        for _ in range(max_pages):
            body = dict(payload)
            body["limit"] = min(limit, int(body.get("limit", limit)))
            body[cursor_field] = cursor
            data = self.post(path, body)
            got = extract_items(data, paths)
            rows.extend(got)
            nxt = extract_cursor(data)
            if not got or not nxt or nxt == cursor:
                break
            cursor = nxt
        return rows


class Storage:
    def __init__(self, access_key: str, secret_key: str, bucket: str, endpoint: str):
        self.bucket = bucket
        self.client = boto3.client(
            "s3", endpoint_url=endpoint,
            aws_access_key_id=access_key, aws_secret_access_key=secret_key,
            region_name="ru-central1",
            config=Config(signature_version="s3v4", read_timeout=300, connect_timeout=60,
                          retries={"max_attempts": 7, "mode": "standard"}),
        )

    def exists(self, key: str) -> bool:
        try:
            self.client.head_object(Bucket=self.bucket, Key=key)
            return True
        except ClientError as exc:
            if str(exc.response.get("Error", {}).get("Code", "")) in {"404", "NoSuchKey", "NotFound"}:
                return False
            raise

    def read_json(self, key: str) -> Dict[str, Any]:
        if not self.exists(key):
            return {}
        try:
            raw = self.client.get_object(Bucket=self.bucket, Key=key)["Body"].read()
            return json.loads(raw.decode("utf-8"))
        except Exception:
            return {}

    def upload(self, key: str, data: bytes, content_type: str) -> None:
        self.client.put_object(Bucket=self.bucket, Key=key, Body=data, ContentType=content_type)

    def upload_json(self, key: str, value: Any) -> None:
        self.upload(key, json.dumps(value, ensure_ascii=False, indent=2, default=str).encode("utf-8"), "application/json")


@dataclass
class ClusterMaps:
    cluster_names: List[str]
    warehouse_to_cluster: Dict[str, str]
    warehouse_name_to_cluster: Dict[str, str]
    macro_to_cluster: Dict[str, str]
    raw_endpoint: str

    def canonical(self, name: Any) -> str:
        text = str(name or "").strip()
        if not text:
            return ""
        best_name, best_score = text, 0.0
        for candidate in self.cluster_names:
            score = similarity(text, candidate)
            if score > best_score:
                best_name, best_score = candidate, score
        return best_name if best_score >= 0.62 else text


class CrossdockTariffProvider:
    """Встроенные сроки кросс-дока для FINICK из тарифного XLSX от 03.08.2026."""

    def __init__(self, store: str):
        self.store = store.upper()
        self.source = CROSSDOCK_TARIFF_SOURCE
        self.updated_at = CROSSDOCK_TARIFF_AS_OF
        self.rows: Dict[str, int] = dict(CROSSDOCK_CLUSTER_DAYS)

    @staticmethod
    def _best_match(value: str, mapping: Mapping[str, int], threshold: float) -> Tuple[Optional[int], str]:
        value = str(value or "").strip()
        if not value:
            return None, ""
        target = norm_name(value)
        for key, days in mapping.items():
            if norm_name(key) == target:
                return int(days), key
        best_key, best_score = "", 0.0
        for key in mapping:
            score = similarity(value, key)
            if score > best_score:
                best_key, best_score = key, score
        if best_key and best_score >= threshold:
            return int(mapping[best_key]), best_key
        return None, ""

    def lookup(self, cluster: str, warehouse: str = "") -> Tuple[Optional[int], str]:
        if self.store != "FINICK":
            return None, f"Тарифный маршрут пока настроен только для FINICK ({CROSSDOCK_ORIGIN})"
        # Точный конечный склад сильнее агрегата по кластеру.
        days, matched = self._best_match(warehouse, CROSSDOCK_WAREHOUSE_DAYS, 0.72)
        if days is not None:
            return days, f"{self.source}; склад={matched}"
        days, matched = self._best_match(cluster, CROSSDOCK_CLUSTER_DAYS, 0.60)
        if days is not None:
            return days, f"{self.source}; кластер={matched}, верхняя граница"
        return None, self.source

    def as_dataframe(self) -> pd.DataFrame:
        rows: List[Dict[str, Any]] = []
        for warehouse, days in sorted(CROSSDOCK_WAREHOUSE_DAYS.items()):
            rows.append({
                "Точка отправления": CROSSDOCK_ORIGIN,
                "Уровень норматива": "Конечный склад",
                "Кластер / конечный склад": warehouse,
                "Плановый срок доставки, дней (верхняя граница)": days,
                "Источник": self.source,
                "Актуально на": self.updated_at,
            })
        for cluster, days in sorted(CROSSDOCK_CLUSTER_DAYS.items()):
            rows.append({
                "Точка отправления": CROSSDOCK_ORIGIN,
                "Уровень норматива": "Кластер — резерв",
                "Кластер / конечный склад": cluster,
                "Плановый срок доставки, дней (верхняя граница)": days,
                "Источник": self.source,
                "Актуально на": self.updated_at,
            })
        return pd.DataFrame(rows)


class ReportBuilder:
    def __init__(self, store: str, target_date: date, client: OzonClient, storage: Storage):
        self.store = store.upper()
        self.target_date = target_date
        self.client = client
        self.storage = storage
        self.offer_by_sku: Dict[str, str] = {}
        self.name_by_sku: Dict[str, str] = {}
        self.sku_ids: List[int] = []
        self.cluster_maps = ClusterMaps([], {}, {}, {}, "")
        self.sla = CrossdockTariffProvider(self.store)
        # FBO-отправления за 7 дней загружаем один раз для расчёта продаж.
        self._postings_7d_cache: Optional[List[Dict[str, Any]]] = None
        # Для свежего выкупа используем историю до BUYOUT_LOOKBACK_DAYS,
        # при этом последние 7 дней повторно не скачиваем.
        self._postings_buyout_cache: Optional[List[Dict[str, Any]]] = None
        self._buyout_diagnostics: Dict[str, Any] = {}
        self._posting_qty_scale: int = 1
        self._posting_qty_detection: Dict[str, Any] = {}
        self._unknown_warehouses: set[str] = set()
        self._supply_diagnostics: Dict[str, Any] = {}

    # ---------- справочники ----------
    def fetch_products(self) -> None:
        logging.info("1/8 Справочник товаров")
        items = self.client.cursor_pages(
            "/v3/product/list", {"filter": {"visibility": "ALL"}, "limit": 1000, "last_id": ""},
            [("result", "items"), ("items",)], cursor_field="last_id")
        offers = [str(x.get("offer_id")) for x in items if x.get("offer_id") not in (None, "")]
        for i in range(0, len(offers), 100):
            batch = offers[i:i+100]
            data = self.client.post("/v3/product/info/list", {"offer_id": batch, "product_id": [], "sku": []})
            for item in extract_items(data, [("items",), ("result", "items")]):
                offer = str(item.get("offer_id") or item.get("offerId") or "").strip()
                name = str(item.get("name") or item.get("product_name") or "").strip()
                for key in ("sku", "fbo_sku"):
                    sku = norm_id(item.get(key))
                    if sku:
                        if sku.isdigit():
                            self.sku_ids.append(int(sku))
                        if offer:
                            self.offer_by_sku[sku] = offer
                        if name:
                            self.name_by_sku[sku] = name
        self.sku_ids = sorted(set(self.sku_ids))
        logging.info("Товаров SKU в справочнике: %s", len(self.sku_ids))

    def fetch_clusters(self) -> None:
        logging.info("2/8 Кластеры FBO")
        names: List[str] = []
        wh_id: Dict[str, str] = {}
        wh_name: Dict[str, str] = {}
        macro: Dict[str, str] = {}

        def add_name(value: Any) -> str:
            name = str(value or "").strip()
            if name and name not in names:
                names.append(name)
            return name

        def parse_cluster_tree(node: Any, inherited: str = "", parent_key: str = "") -> None:
            if isinstance(node, list):
                for item in node:
                    parse_cluster_tree(item, inherited, parent_key)
                return
            if not isinstance(node, Mapping):
                return

            explicit = str(node.get("macrolocal_cluster_name") or node.get("cluster_name") or "").strip()
            if not explicit and parent_key in {"clusters", "macrolocal_clusters", "logistic_clusters"}:
                explicit = str(node.get("name") or "").strip()
            if not explicit and any(k in node for k in ("warehouses", "logistic_clusters")):
                explicit = str(node.get("name") or "").strip()
            cname = add_name(explicit or inherited) if (explicit or inherited) else ""

            if cname:
                for key in ("macrolocal_cluster_id", "macro_cluster_id", "cluster_id", "id"):
                    cid = norm_id(node.get(key))
                    if cid and parent_key not in {"warehouses", "warehouse"}:
                        macro.setdefault(cid, cname)

            # Если узел явно складской — связываем его с наследованным кластером.
            is_warehouse = parent_key in {"warehouses", "warehouse"} or "warehouse_id" in node or "warehouse_name" in node
            if is_warehouse and cname:
                wid = norm_id(node.get("warehouse_id") or node.get("id"))
                wname = str(node.get("warehouse_name") or node.get("name") or "").strip()
                if wid:
                    wh_id[wid] = cname
                if wname:
                    wh_name[norm_name(wname)] = cname

            for key, value in node.items():
                if isinstance(value, (Mapping, list)):
                    parse_cluster_tree(value, cname or inherited, str(key))

        response: Any = {}
        try:
            response = self.client.post("/v2/cluster/list", {})
            parse_cluster_tree(response)
        except OzonApiError as exc:
            logging.warning("/v2/cluster/list недоступен: %s", exc)

        # Дополнительный актуальный метод складов Ozon. Не делаем его обязательным:
        # структура бета/основной версии могла меняться.
        try:
            wresp = self.client.post("/v1/warehouse/ozon/list", {})
            parse_cluster_tree(wresp)
        except OzonApiError as exc:
            if exc.status not in {400, 403, 404, 405, 422}:
                logging.warning("/v1/warehouse/ozon/list: %s", exc)

        # Добавляем известные названия кластеров, чтобы canonical() мог нормализовать
        # статический warehouse fallback даже если /v2/cluster/list вернул новую схему.
        for _tokens, cluster in STATIC_WAREHOUSE_CLUSTER_PATTERNS:
            add_name(cluster)

        self.cluster_maps = ClusterMaps(names, wh_id, wh_name, macro, "/v2/cluster/list + /v1/warehouse/ozon/list")
        logging.info("Кластеров/нормализованных названий: %s, складов с API-привязкой: %s", len(names), len(wh_id) + len(wh_name))
        if isinstance(response, Mapping) and not wh_id and not wh_name:
            result = response.get("result")
            if isinstance(result, Mapping):
                logging.warning("Кластерный API не дал привязку складов. Поля result: %s", list(result.keys())[:30])
            elif isinstance(result, list):
                first_keys = list(result[0].keys()) if result and isinstance(result[0], Mapping) else []
                logging.warning("Кластерный API: result=list[%s], поля первого элемента: %s", len(result), first_keys)

    # ---------- остатки ----------
    def fetch_stock_warehouses(self) -> pd.DataFrame:
        logging.info("3/8 Остатки FBO по складам")
        rows = self.client.offset_pages(
            "/v2/analytics/stock_on_warehouses",
            {"limit": 1000, "offset": 0, "warehouse_type": "ALL"},
            [("result", "rows"), ("rows",), ("result", "items")], limit=1000)
        out: List[Dict[str, Any]] = []
        for r in rows:
            sku = norm_id(r.get("sku"))
            if not sku:
                continue
            offer = str(r.get("item_code") or r.get("offer_id") or self.offer_by_sku.get(sku, "")).strip()
            name = str(r.get("item_name") or r.get("name") or self.name_by_sku.get(sku, "")).strip()
            wh_name = str(r.get("warehouse_name") or "").strip()
            wh_id = norm_id(r.get("warehouse_id"))
            cluster = str(r.get("cluster_name") or r.get("macrolocal_cluster_name") or r.get("cluster") or "").strip()
            source = "API остатков" if cluster else ""
            if not cluster and wh_id:
                cluster = self.cluster_maps.warehouse_to_cluster.get(wh_id, "")
                if cluster:
                    source = "Справочник API"
            if not cluster and wh_name:
                cluster = self.cluster_maps.warehouse_name_to_cluster.get(norm_name(wh_name), "")
                if cluster:
                    source = "Справочник API"
            if not cluster and wh_name:
                cluster = static_cluster_from_warehouse(wh_name)
                if cluster:
                    source = "Резерв по названию склада"
            if cluster:
                cluster = self.cluster_maps.canonical(cluster)
            else:
                cluster = "Кластер не определён"
                if wh_name:
                    self._unknown_warehouses.add(wh_name)
            out.append({
                "SKU Ozon": sku,
                "Артикул": offer or f"SKU {sku}",
                "Название": name,
                "Кластер": cluster,
                "Склад Ozon": wh_name,
                "Источник кластера": source or "Не определён",
                "Доступно сейчас, шт.": int_qty(r.get("free_to_sell_amount")),
                "Подтверждено Ozon в поставках, шт.": int_qty(r.get("promised_amount")),
                "Зарезервировано Ozon, шт.": int_qty(r.get("reserved_amount")),
            })
            if offer:
                self.offer_by_sku[sku] = offer
            if name:
                self.name_by_sku[sku] = name
        logging.info("Строк остатков: %s; складов без кластера: %s", len(out), len(self._unknown_warehouses))
        return pd.DataFrame(out)

    def fetch_transit_stocks(self) -> pd.DataFrame:
        logging.info("4/8 Возвраты покупателей и перемещения внутри Ozon")
        if not self.sku_ids:
            return pd.DataFrame()
        out: List[Dict[str, Any]] = []
        for i in range(0, len(self.sku_ids), 100):
            batch = self.sku_ids[i:i+100]
            try:
                rows = self.client.offset_pages(
                    "/v1/analytics/stocks", {"skus": batch, "limit": 1000, "offset": 0},
                    [("result", "rows"), ("rows",), ("result", "items"), ("items",)], limit=1000)
            except OzonApiError:
                rows = self.client.offset_pages(
                    "/v1/analytics/stocks", {"skus": [str(x) for x in batch], "limit": 1000, "offset": 0},
                    [("result", "rows"), ("rows",), ("result", "items"), ("items",)], limit=1000)
            for r in rows:
                sku = norm_id(r.get("sku"))
                if not sku:
                    continue
                cluster = str(r.get("cluster_name") or r.get("cluster") or "").strip()
                wh_name = str(r.get("warehouse_name") or "").strip()
                wh_id = norm_id(r.get("warehouse_id"))
                if not cluster and wh_id:
                    cluster = self.cluster_maps.warehouse_to_cluster.get(wh_id, "")
                if not cluster and wh_name:
                    cluster = self.cluster_maps.warehouse_name_to_cluster.get(norm_name(wh_name), "")
                out.append({
                    "SKU Ozon": sku,
                    "Артикул": self.offer_by_sku.get(sku, f"SKU {sku}"),
                    "Кластер": self.cluster_maps.canonical(cluster) if cluster else "",
                    "Возврат покупателя в пути, шт.": int_qty(r.get("return_from_customer_stock_count")),
                    "Перемещение внутри Ozon, шт.": int_qty(r.get("transit_stock_count")),
                })
        return pd.DataFrame(out)

    # ---------- продажи ----------
    def _fetch_postings_v3(self, start: date, end: date) -> List[Dict[str, Any]]:
        rows: List[Dict[str, Any]] = []
        seen_postings: set[str] = set()
        seen_cursors: set[str] = set()
        cursor = ""
        for page in range(1, 501):
            payload: Dict[str, Any] = {
                "filter": {"since": to_ozon_datetime(start), "to": to_ozon_datetime(end, True)},
                "limit": 100,
                "sort_dir": "ASC",
                "with": {"analytics_data": True, "financial_data": True, "legal_info": False},
            }
            if cursor:
                payload["cursor"] = cursor
            data = self.client.post("/v3/posting/fbo/list", payload)
            postings = extract_items(data, [("result", "postings"), ("postings",), ("result", "items"), ("items",)])
            added = 0
            for p in postings:
                key = str(p.get("posting_number") or p.get("order_id") or p.get("order_number") or "").strip()
                if not key:
                    key = json.dumps(p, ensure_ascii=False, sort_keys=True, default=str)[:500]
                if key in seen_postings:
                    continue
                seen_postings.add(key)
                rows.append(p)
                added += 1
            root = data.get("result", {}) if isinstance(data, Mapping) and isinstance(data.get("result"), Mapping) else {}
            next_cursor = str(
                (data.get("cursor") if isinstance(data, Mapping) else "") or
                root.get("cursor") or root.get("next_cursor") or
                deep_first(data, ["cursor", "next_cursor"]) or ""
            ).strip()
            has_next_val = (data.get("has_next") if isinstance(data, Mapping) else None)
            if has_next_val is None:
                has_next_val = root.get("has_next") if isinstance(root, Mapping) else None
            has_next = bool(has_next_val) if has_next_val is not None else bool(next_cursor and len(postings) >= 100)
            logging.info("FBO v3: страница %s, получено %s, новых %s, итого %s", page, len(postings), added, len(rows))
            if not postings or added == 0:
                break
            if not has_next:
                # Если страница полная, но Ozon не дал признака/курсора, считаем это
                # подозрительным и переключаемся на v2 fallback, пока v2 ещё доступен.
                if len(postings) >= 100 and not next_cursor:
                    raise RuntimeError("FBO v3: полная страница без cursor/has_next")
                break
            if not next_cursor:
                raise RuntimeError("FBO v3 сообщил продолжение, но не вернул cursor")
            if next_cursor == cursor or next_cursor in seen_cursors:
                raise RuntimeError("FBO v3 вернул повторяющийся cursor")
            seen_cursors.add(next_cursor)
            cursor = next_cursor
        else:
            raise RuntimeError("FBO v3: превышен защитный лимит 500 страниц")
        return rows

    def _fetch_postings_v2_fallback(self, start: date, end: date) -> List[Dict[str, Any]]:
        """Временный fallback до отключения v2 31.08.2026. Нужен только если v3
        изменил схему cursor. Защита от повторяющейся страницы обязательна."""
        rows: List[Dict[str, Any]] = []
        seen: set[str] = set()
        offset = 0
        for page in range(1, 501):
            payload = {
                "dir": "ASC",
                "filter": {"since": to_ozon_datetime(start), "to": to_ozon_datetime(end, True), "status": ""},
                "limit": 100,
                "offset": offset,
                "translit": False,
                "with": {"analytics_data": True, "financial_data": True},
            }
            data = self.client.post("/v2/posting/fbo/list", payload)
            postings = extract_items(data, [("result",), ("result", "postings"), ("postings",)])
            # /v2 обычно result=list; extract_items умеет list только на верхнем уровне,
            # поэтому обрабатываем этот вариант отдельно.
            if isinstance(data, Mapping) and isinstance(data.get("result"), list):
                postings = [dict(x) for x in data["result"] if isinstance(x, Mapping)]
            added = 0
            for p in postings:
                key = str(p.get("posting_number") or p.get("order_id") or p.get("order_number") or "").strip()
                if not key:
                    key = json.dumps(p, ensure_ascii=False, sort_keys=True, default=str)[:500]
                if key in seen:
                    continue
                seen.add(key)
                rows.append(p)
                added += 1
            logging.info("FBO v2 fallback: страница %s, получено %s, новых %s, итого %s", page, len(postings), added, len(rows))
            if not postings or added == 0 or len(postings) < 100:
                break
            offset += len(postings)
        return rows

    def fetch_postings(self, start: date, end: date) -> List[Dict[str, Any]]:
        try:
            rows = self._fetch_postings_v3(start, end)
            logging.info("FBO postings: основной источник /v3/posting/fbo/list, строк %s", len(rows))
            return rows
        except Exception as exc:
            logging.warning("FBO v3 не удалось надёжно прочитать (%s). Пробуем временный v2 fallback.", exc)
            try:
                rows = self._fetch_postings_v2_fallback(start, end)
                logging.info("FBO postings: fallback /v2/posting/fbo/list, строк %s", len(rows))
                return rows
            except Exception as exc2:
                raise RuntimeError(f"Не удалось получить FBO-отправления ни v3, ни v2: v3={exc}; v2={exc2}") from exc2

    def _detect_posting_quantity_scale(self, postings: Sequence[Mapping[str, Any]]) -> None:
        vals: List[float] = []
        for p in postings:
            products = p.get("products") if isinstance(p.get("products"), list) else []
            for product in products:
                if not isinstance(product, Mapping):
                    continue
                raw = num(product.get("quantity"), 0.0)
                if raw > 0:
                    vals.append(raw)
                if len(vals) >= 5000:
                    break
            if len(vals) >= 5000:
                break
        if not vals:
            self._posting_qty_scale = 1
            self._posting_qty_detection = {"sample": 0, "scale": 1, "reason": "нет quantity"}
            return
        multiples = sum(1 for x in vals if x >= 1000 and abs(x / 1000 - round(x / 1000)) < 1e-9)
        ratio = multiples / len(vals)
        sorted_vals = sorted(vals)
        median = sorted_vals[len(sorted_vals)//2]
        scale = 1000 if len(vals) >= 5 and ratio >= 0.80 and median >= 1000 else 1
        self._posting_qty_scale = scale
        self._posting_qty_detection = {
            "sample": len(vals), "scale": scale, "ratio_multiples_1000": round(ratio, 3),
            "median_raw": median, "min_raw": min(vals), "max_raw": max(vals),
        }
        logging.info("FBO quantity: масштаб=%s, sample=%s, доля кратных 1000=%.1f%%, медиана raw=%s", scale, len(vals), ratio*100, median)

    def _posting_product_qty(self, product: Mapping[str, Any]) -> int:
        raw = num(product.get("quantity"), 1.0)
        if self._posting_qty_scale == 1000:
            raw = raw / 1000.0
        q = int(round(raw))
        return max(1, q)

    def postings_7d(self) -> List[Dict[str, Any]]:
        if self._postings_7d_cache is None:
            start = self.target_date - timedelta(days=6)
            logging.info("Загрузка FBO-отправлений только за последние 7 дней")
            self._postings_7d_cache = self.fetch_postings(start, self.target_date)
            self._detect_posting_quantity_scale(self._postings_7d_cache)
            logging.info("FBO-отправлений за 7 дней получено: %s", len(self._postings_7d_cache))
        return self._postings_7d_cache

    @staticmethod
    def _posting_date(posting: Mapping[str, Any]) -> Optional[date]:
        return parse_any_date(
            posting.get("created_at") or
            posting.get("in_process_at") or
            posting.get("shipment_date")
        )

    def sales_7d_by_cluster(self) -> Tuple[pd.DataFrame, List[Dict[str, Any]]]:
        logging.info("5/8 Продажи за 7 дней по кластерам")
        start = self.target_date - timedelta(days=6)
        all_postings = self.postings_7d()
        postings = [
            p for p in all_postings
            if (self._posting_date(p) is not None and start <= self._posting_date(p) <= self.target_date)
        ]
        logging.info("FBO-отправлений в окне последних 7 дней: %s", len(postings))

        rows: List[Dict[str, Any]] = []
        for p in postings:
            status = str(p.get("status") or "").lower()
            if status == "cancelled":
                continue
            financial = p.get("financial_data") if isinstance(p.get("financial_data"), Mapping) else {}
            analytics = p.get("analytics_data") if isinstance(p.get("analytics_data"), Mapping) else {}
            cluster = str(
                financial.get("cluster_to") or analytics.get("cluster_to") or
                p.get("cluster_to") or deep_first(p, ["cluster_to"]) or "Кластер не определён"
            ).strip()
            cluster = self.cluster_maps.canonical(cluster)
            products = p.get("products") if isinstance(p.get("products"), list) else []
            for product in products:
                if not isinstance(product, Mapping):
                    continue
                sku = norm_id(product.get("sku") or product.get("product_id"))
                if not sku:
                    continue
                offer = str(product.get("offer_id") or self.offer_by_sku.get(sku, f"SKU {sku}")).strip()
                name = str(product.get("name") or self.name_by_sku.get(sku, "")).strip()
                qty = self._posting_product_qty(product)
                rows.append({"Артикул": offer, "SKU Ozon": sku, "Название": name, "Кластер": cluster, "Заказано, шт.": qty})
        df = pd.DataFrame(rows)
        if df.empty:
            return pd.DataFrame(columns=["Артикул", "Кластер", "Заказано за 7 дней, шт.", "Среднесуточные продажи за 7 дней, шт."]), postings
        agg = df.groupby(["Артикул", "Кластер"], as_index=False).agg(
            **{"Заказано за 7 дней, шт.": ("Заказано, шт.", "sum")})
        agg["Среднесуточные продажи за 7 дней, шт."] = (agg["Заказано за 7 дней, шт."] / 7).round(1)
        return agg, postings

    # ---------- выкуп: самые свежие завершённые единицы ----------
    def fetch_returns(self, start: date, end: date) -> List[Dict[str, Any]]:
        """Возвраты Ozon за период. Для выкупа используем только type=ClientReturn.

        Cancellation из returns/list НЕ считаем отдельно: отменённая единица уже
        учитывается по status=cancelled у FBO posting, иначе получится двойной невыкуп.
        """
        rows: List[Dict[str, Any]] = []
        last_id: Any = 0
        seen_last_ids: set[str] = set()
        for page in range(1, 101):
            payload = {
                "filter": {
                    "logistic_return_date": {
                        "time_from": to_ozon_datetime(start),
                        "time_to": to_ozon_datetime(end, True),
                    }
                },
                "limit": 500,
                "last_id": last_id,
            }
            data = self.client.post("/v1/returns/list", payload)
            got = extract_items(data, [("returns",), ("result", "returns"), ("result", "items"), ("items",)])

            # Дополнительный локальный контроль диапазона: API иногда содержит несколько
            # дат возврата, поэтому не даём событию позже даты отчёта попасть в расчёт.
            accepted: List[Dict[str, Any]] = []
            for item in got:
                rdate = parse_any_date(deep_first(item, ["logistic_return_date", "return_date", "final_moment", "created_at"]))
                if rdate is not None and not (start <= rdate <= end):
                    continue
                accepted.append(item)
            rows.extend(accepted)

            nxt = deep_first(data, ["last_id"])
            logging.info("Возвраты для выкупа: страница %s, получено %s, в диапазоне %s, итого %s", page, len(got), len(accepted), len(rows))
            if not got or nxt in (None, "", 0, last_id):
                break
            nxt_s = str(nxt)
            if nxt_s in seen_last_ids:
                logging.warning("Возвраты для выкупа: повторился last_id, пагинация остановлена")
                break
            seen_last_ids.add(nxt_s)
            last_id = nxt
        else:
            logging.warning("Возвраты для выкупа: достигнут защитный лимит 100 страниц")
        return rows

    @staticmethod
    def _return_posting_number(item: Mapping[str, Any]) -> str:
        return str(deep_first(item, ["posting_number", "postingNumber"]) or "").strip()

    def postings_for_buyout(self) -> List[Dict[str, Any]]:
        """История FBO для последних 50 завершённых единиц без повторной загрузки 7 дней."""
        if self._postings_buyout_cache is not None:
            return self._postings_buyout_cache

        recent = list(self.postings_7d())
        start = self.target_date - timedelta(days=BUYOUT_LOOKBACK_DAYS - 1)
        older_end = self.target_date - timedelta(days=7)
        older: List[Dict[str, Any]] = []
        if start <= older_end:
            logging.info(
                "7/8 Выкуп: загружаем FBO-историю %s дней; последние 7 дней уже в кэше",
                BUYOUT_LOOKBACK_DAYS,
            )
            # FINICK имеет большой поток заказов. Один длинный диапазон может упереться
            # в защитный лимит страниц FBO list, поэтому читаем историю кусками по 14 дней.
            chunk_start = start
            while chunk_start <= older_end:
                chunk_end = min(older_end, chunk_start + timedelta(days=13))
                logging.info("Выкуп: FBO-история %s..%s", chunk_start, chunk_end)
                older.extend(self.fetch_postings(chunk_start, chunk_end))
                chunk_start = chunk_end + timedelta(days=1)

        merged: List[Dict[str, Any]] = []
        seen: set[str] = set()
        for p in older + recent:
            key = str(p.get("posting_number") or p.get("order_id") or p.get("order_number") or "").strip()
            if not key:
                key = json.dumps(p, ensure_ascii=False, sort_keys=True, default=str)[:500]
            if key in seen:
                continue
            seen.add(key)
            pdate = self._posting_date(p)
            if pdate is not None and start <= pdate <= self.target_date:
                merged.append(p)

        self._postings_buyout_cache = merged
        logging.info("FBO-отправлений в истории выкупа: %s", len(merged))
        return merged

    def buyout_latest_by_article(self) -> pd.DataFrame:
        """Самый свежий выкуп по артикулу: до 50 последних завершённых единиц.

        completed = delivered + cancelled. Незавершённые статусы не участвуют.
        delivered + ClientReturn по ключу posting_number + SKU считается возвратом.
        Cancellation из returns/list повторно не вычитается.
        """
        postings = self.postings_for_buyout()
        start = self.target_date - timedelta(days=BUYOUT_LOOKBACK_DAYS - 1)
        returns = self.fetch_returns(start, self.target_date)

        # Сколько ClientReturn приходится на конкретный posting + SKU.
        client_returns: Dict[Tuple[str, str], int] = {}
        client_return_rows = 0
        for r in returns:
            rtype = str(r.get("type") or r.get("return_type") or "").strip().lower().replace("_", "")
            if rtype != "clientreturn":
                continue
            product = r.get("product") if isinstance(r.get("product"), Mapping) else {}
            sku = norm_id(product.get("sku") or r.get("sku") or deep_first(r, ["sku"]))
            posting_number = self._return_posting_number(r)
            if not sku or not posting_number:
                continue
            qty = max(1, int_qty(product.get("quantity") or r.get("quantity") or 1))
            key = (posting_number, sku)
            client_returns[key] = client_returns.get(key, 0) + qty
            client_return_rows += 1

        # На один артикул храним события завершённых единиц. Сортировка по дате
        # идёт от свежих к старым; затем отрезаем первые BUYOUT_TARGET_UNITS.
        # tuple: (date, posting_number, bought_qty, cancelled_qty, returned_qty)
        events: Dict[str, List[Tuple[date, str, int, int, int]]] = {}
        for p in postings:
            status = str(p.get("status") or "").strip().lower()
            if status not in {"delivered", "cancelled"}:
                continue
            pdate = self._posting_date(p)
            if pdate is None:
                continue
            posting_number = str(p.get("posting_number") or p.get("order_number") or p.get("order_id") or "").strip()
            products = p.get("products") if isinstance(p.get("products"), list) else []
            for product in products:
                if not isinstance(product, Mapping):
                    continue
                sku = norm_id(product.get("sku") or product.get("product_id"))
                if not sku:
                    continue
                offer = str(product.get("offer_id") or self.offer_by_sku.get(sku, f"SKU {sku}")).strip()
                if not offer:
                    continue
                qty = self._posting_product_qty(product)
                if status == "cancelled":
                    bought_qty, cancelled_qty, returned_qty = 0, qty, 0
                else:
                    returned_qty = min(qty, client_returns.get((posting_number, sku), 0))
                    bought_qty = max(0, qty - returned_qty)
                    cancelled_qty = 0
                events.setdefault(offer, []).append((pdate, posting_number, bought_qty, cancelled_qty, returned_qty))

        result: List[Dict[str, Any]] = []
        bases: List[int] = []
        full50 = 0
        for offer, article_events in events.items():
            article_events.sort(key=lambda x: (x[0], x[1]), reverse=True)
            remaining = BUYOUT_TARGET_UNITS
            bought = cancelled = returned = 0
            latest_date: Optional[date] = article_events[0][0] if article_events else None
            oldest_used: Optional[date] = None

            for event_date, _posting, b, c, r in article_events:
                if remaining <= 0:
                    break
                total = b + c + r
                if total <= 0:
                    continue
                take = min(remaining, total)

                # Обычно quantity=1. Если граница 50 попала внутрь многоколичественного
                # posting, сначала берём возвраты/отмены, затем выкуп — консервативно.
                take_r = min(r, take)
                left = take - take_r
                take_c = min(c, left)
                left -= take_c
                take_b = min(b, left)

                returned += take_r
                cancelled += take_c
                bought += take_b
                remaining -= take_r + take_c + take_b
                oldest_used = event_date

            base = bought + cancelled + returned
            if base <= 0:
                continue
            pct = round(bought / base * 100, 1)
            bases.append(base)
            if base >= BUYOUT_TARGET_UNITS:
                full50 += 1
            result.append({
                "Артикул": offer,
                "Выкуп, % (последние завершённые)": pct,
                "База выкупа, шт.": base,
                "Выкуплено в базе, шт.": bought,
                "Отменено в базе, шт.": cancelled,
                "ClientReturn в базе, шт.": returned,
                "Дата самого свежего завершённого заказа": latest_date,
                "Дата самого старого заказа в базе": oldest_used,
                "Целевая база, шт.": BUYOUT_TARGET_UNITS,
            })

        self._buyout_diagnostics = {
            "lookback_days": BUYOUT_LOOKBACK_DAYS,
            "target_units": BUYOUT_TARGET_UNITS,
            "postings_history": len(postings),
            "returns_rows": len(returns),
            "client_return_rows_with_key": client_return_rows,
            "articles_with_buyout": len(result),
            "articles_with_full_50": full50,
            "min_base": min(bases) if bases else 0,
            "max_base": max(bases) if bases else 0,
        }
        logging.info(
            "Выкуп рассчитан: артикулов=%s, полная база 50=%s, ClientReturn с posting+SKU=%s",
            len(result), full50, client_return_rows,
        )
        return pd.DataFrame(result)

    # ---------- поставки ----------
    @staticmethod
    def _extract_timeslot(obj: Any) -> Tuple[Optional[date], Optional[date]]:
        if not isinstance(obj, Mapping):
            return None, None
        timeslot = obj.get("timeslot") if isinstance(obj.get("timeslot"), Mapping) else {}
        nested = timeslot.get("timeslot") if isinstance(timeslot.get("timeslot"), Mapping) else {}
        start = parse_any_date(nested.get("from") or timeslot.get("from"))
        end = parse_any_date(nested.get("to") or timeslot.get("to"))
        if start or end:
            return start or end, end or start

        # Резерв на старую схему ответа, где from/to могли лежать глубже.
        pairs: List[Tuple[date, date]] = []
        for d in walk_dicts(obj):
            if d.get("from") and d.get("to"):
                a, b = parse_any_date(d.get("from")), parse_any_date(d.get("to"))
                if a and b:
                    pairs.append((a, b))
        if not pairs:
            return None, None
        pairs.sort(key=lambda x: x[0])
        return pairs[0]

    def _bundle_items(self, bundle_id: str) -> List[Dict[str, Any]]:
        rows: List[Dict[str, Any]] = []
        last_id = ""
        for _ in range(100):
            data = None
            last_exc: Optional[Exception] = None
            for payload in (
                {"bundle_ids": [bundle_id], "limit": 1000, "last_id": last_id},
                {"bundle_ids": [bundle_id], "limit": 1000, "last_id": last_id, "is_asc": True, "sort_field": "SKU", "query": ""},
            ):
                try:
                    data = self.client.post("/v1/supply-order/bundle", payload)
                    break
                except OzonApiError as exc:
                    last_exc = exc
                    if exc.status in {400, 404, 422}:
                        continue
                    raise
            if data is None:
                if last_exc:
                    raise last_exc
                break
            got = extract_items(data, [("items",), ("result", "items")])
            rows.extend(got)
            has_next = bool(data.get("has_next")) if isinstance(data, Mapping) else False
            if isinstance(data, Mapping) and isinstance(data.get("result"), Mapping):
                has_next = has_next or bool(data["result"].get("has_next"))
            nxt = extract_cursor(data)
            if not got or not has_next or not nxt or nxt == last_id:
                break
            last_id = nxt
        return rows

    def _supply_list_v3(self) -> List[Dict[str, Any]]:
        """Получаем order_id по каждому активному статусу.

        Это та же схема, которая уже работает в пользовательском контуре поставок:
        /v3/supply-order/list -> order_ids -> /v3/supply-order/get.
        """
        result: List[Dict[str, Any]] = []
        seen: set[str] = set()
        per_state: Dict[str, int] = {}

        for state in sorted(ACTIVE_STATES):
            last_id = ""
            state_ids: List[str] = []
            for page in range(1, 101):
                payload: Dict[str, Any] = {
                    "filter": {"states": [state]},
                    "limit": 100,
                    "sort_by": "ORDER_CREATION",
                    "sort_dir": "DESC",
                }
                if last_id:
                    payload["last_id"] = last_id
                data = self.client.post("/v3/supply-order/list", payload)
                raw_ids: Any = data.get("order_ids") if isinstance(data, Mapping) else []
                if not isinstance(raw_ids, list) and isinstance(data, Mapping) and isinstance(data.get("result"), Mapping):
                    raw_ids = data["result"].get("order_ids")
                ids = [norm_id(x) for x in (raw_ids or []) if norm_id(x)]
                logging.info("Supply list v3: статус=%s, страница=%s, order_ids=%s", state, page, len(ids))
                state_ids.extend(ids)

                nxt = str((data.get("last_id") if isinstance(data, Mapping) else "") or "").strip()
                if not nxt and isinstance(data, Mapping) and isinstance(data.get("result"), Mapping):
                    nxt = str(data["result"].get("last_id") or "").strip()
                if not nxt or len(ids) < 100 or nxt == last_id:
                    break
                last_id = nxt

            unique_state_ids = list(dict.fromkeys(state_ids))
            per_state[state] = len(unique_state_ids)
            for oid in unique_state_ids:
                if oid in seen:
                    continue
                seen.add(oid)
                result.append({"order_id": oid, "state": state})

        self._supply_diagnostics["order_ids_by_state"] = per_state
        self._supply_diagnostics["list_source"] = "/v3/supply-order/list -> order_ids"
        return result

    def _legacy_supply_items(self, oid: str) -> List[Dict[str, Any]]:
        rows: List[Dict[str, Any]] = []
        page = 1
        for _ in range(100):
            try:
                data = self.client.post("/v1/supply-order/items", {
                    "supply_order_id": int(oid) if str(oid).isdigit() else oid,
                    "page": page,
                    "page_size": 100,
                })
            except OzonApiError:
                return rows
            got = extract_items(data, [("items",), ("result", "items")])
            rows.extend(got)
            has_next = bool(data.get("has_next")) if isinstance(data, Mapping) else False
            if not got or not has_next:
                break
            page += 1
        return rows

    @staticmethod
    def _supply_state(obj: Mapping[str, Any]) -> str:
        return str(obj.get("supply_state") or obj.get("state") or obj.get("status") or "").upper().strip()

    @staticmethod
    def _is_active_supply_state(state: str) -> bool:
        if not state:
            return True  # неизвестный новый статус не теряем раньше времени
        return not any(token in state for token in TERMINAL_SUPPLY_STATES)

    def _supply_detail(self, oid: str) -> Mapping[str, Any]:
        payload_id: Any = int(oid) if str(oid).isdigit() else oid
        data = self.client.post("/v3/supply-order/get", {"order_ids": [payload_id]})
        orders = data.get("orders") if isinstance(data, Mapping) else None
        if not isinstance(orders, list) and isinstance(data, Mapping) and isinstance(data.get("result"), Mapping):
            orders = data["result"].get("orders")
        if not isinstance(orders, list) or not orders:
            raise RuntimeError(f"/v3/supply-order/get не вернул orders для order_id={oid}")
        for order in orders:
            if isinstance(order, Mapping) and norm_id(order.get("order_id")) == norm_id(oid):
                return order
        first = orders[0]
        return first if isinstance(first, Mapping) else {}

    def supply_orders(self) -> pd.DataFrame:
        logging.info("6/8 Активные FBO-заявки и поставки")
        list_items = self._supply_list_v3()
        self._supply_diagnostics["list_total"] = len(list_items)
        logging.info("Уникальных активных заявок supply-order/list: %s", len(list_items))
        if not list_items:
            logging.info("Активных заявок нет")
            return pd.DataFrame()

        status_counts: Dict[str, int] = {}
        for item in list_items:
            state = self._supply_state(item) or "НЕИЗВЕСТНО"
            status_counts[state] = status_counts.get(state, 0) + 1
        self._supply_diagnostics["status_counts"] = status_counts

        seen: set[str] = set()
        rows: List[Dict[str, Any]] = []
        detail_errors = 0

        for idx, short in enumerate(list_items, 1):
            oid = norm_id(short.get("order_id") or short.get("supply_order_id") or short.get("id"))
            if not oid or oid in seen:
                continue
            seen.add(oid)
            try:
                order_root = self._supply_detail(oid)
            except Exception as exc:
                detail_errors += 1
                logging.warning("Заявка %s: детали недоступны: %s", oid, exc)
                continue

            order_state = self._supply_state(order_root) or self._supply_state(short)
            if not self._is_active_supply_state(order_state):
                continue

            order_slot_from, _ = self._extract_timeslot(order_root)
            order_dropoff = order_root.get("drop_off_warehouse") if isinstance(order_root.get("drop_off_warehouse"), Mapping) else {}
            order_dropoff_id = norm_id(order_dropoff.get("warehouse_id") or order_dropoff.get("id"))
            order_dropoff_name = str(order_dropoff.get("name") or "").strip()

            supplied = order_root.get("supplies") if isinstance(order_root, Mapping) else None
            supply_nodes = [x for x in (supplied or []) if isinstance(x, Mapping)] if isinstance(supplied, list) else []
            if not supply_nodes:
                # Резерв на случай очередной смены вложенности ответа Ozon.
                for node in walk_dicts(order_root):
                    if node.get("bundle_id") not in (None, ""):
                        supply_nodes.append(node)
            if not supply_nodes:
                logging.warning("Заявка %s: в /v3/supply-order/get не найдено supplies/bundle_id", oid)
                continue

            unique_nodes: List[Mapping[str, Any]] = []
            seen_nodes: set[str] = set()
            for node in supply_nodes:
                key = norm_id(node.get("supply_id") or node.get("id")) or norm_id(node.get("bundle_id")) or str(id(node))
                if key in seen_nodes:
                    continue
                seen_nodes.add(key)
                unique_nodes.append(node)

            for supply in unique_nodes:
                state = self._supply_state(supply) or order_state
                if not self._is_active_supply_state(state):
                    continue

                supply_slot_from, _ = self._extract_timeslot(supply)
                shipment_date = supply_slot_from or order_slot_from
                bundle_id = norm_id(supply.get("bundle_id"))

                storage_wh = supply.get("storage_warehouse") if isinstance(supply.get("storage_warehouse"), Mapping) else {}
                supply_dropoff = supply.get("drop_off_warehouse") if isinstance(supply.get("drop_off_warehouse"), Mapping) else {}
                storage_id = norm_id(supply.get("storage_warehouse_id") or storage_wh.get("warehouse_id") or storage_wh.get("id"))
                storage_name = str(storage_wh.get("name") or supply.get("storage_warehouse_name") or deep_first(supply, ["storage_warehouse_name"]) or "").strip()
                dropoff_id = norm_id(
                    supply.get("dropoff_warehouse_id") or supply.get("drop_off_warehouse_id") or
                    supply_dropoff.get("warehouse_id") or supply_dropoff.get("id") or order_dropoff_id
                )
                dropoff_name = str(supply_dropoff.get("name") or order_dropoff_name or "").strip()

                macro_id = norm_id(
                    supply.get("macrolocal_cluster_id") or
                    deep_first(supply, ["macrolocal_cluster_id", "cluster_id"]) or
                    deep_first(order_root, ["macrolocal_cluster_id", "cluster_id"])
                )
                cluster = self.cluster_maps.macro_to_cluster.get(macro_id, "")
                if not cluster and storage_id:
                    cluster = self.cluster_maps.warehouse_to_cluster.get(storage_id, "")
                if not cluster and storage_name:
                    cluster = self.cluster_maps.warehouse_name_to_cluster.get(norm_name(storage_name), "") or static_cluster_from_warehouse(storage_name)
                if not cluster:
                    cluster = str(
                        deep_first(supply, ["macrolocal_cluster_name", "cluster_name"]) or
                        deep_first(order_root, ["macrolocal_cluster_name", "cluster_name"]) or ""
                    ).strip()
                cluster = self.cluster_maps.canonical(cluster) if cluster else "Кластер не определён"

                # FINICK всегда сдаёт поставки в Москве. Тарифный норматив текущей версии
                # привязан к Москва — Кавказский. Фактическое имя из API показываем отдельно.
                route = "Кросс-док"
                tariff_origin = CROSSDOCK_ORIGIN
                if self.store == "FINICK" and dropoff_name and "моск" not in norm_name(dropoff_name):
                    logging.warning("FINICK: API вернул немосковскую точку отгрузки '%s'; тариф всё равно считается от %s", dropoff_name, tariff_origin)

                arrival_api = parse_any_date(
                    storage_wh.get("arrival_date") or supply.get("arrival_date") or
                    deep_first(supply, ["planned_arrival_date", "estimated_arrival_date", "arrival_date"])
                )
                planned_arrival, eta_basis, sla_days = self._estimate_supply_arrival(
                    cluster=cluster,
                    storage_warehouse=storage_name,
                    shipment_date=shipment_date,
                    arrival_api=arrival_api,
                )

                products: List[Dict[str, Any]] = []
                if bundle_id:
                    try:
                        products = self._bundle_items(bundle_id)
                    except Exception as exc:
                        logging.warning("Состав поставки bundle=%s: %s", bundle_id, exc)
                if not products:
                    for key in ("products", "items"):
                        val = supply.get(key)
                        if isinstance(val, list):
                            products = [dict(x) for x in val if isinstance(x, Mapping)]
                            if products:
                                break
                if not products:
                    products = self._legacy_supply_items(oid)
                if not products:
                    logging.warning("Заявка %s bundle=%s: состав товаров не получен", oid, bundle_id)
                    continue

                for product in products:
                    sku = norm_id(product.get("sku") or product.get("product_id"))
                    if not sku:
                        continue
                    offer = str(
                        product.get("contractor_item_code") or product.get("offer_id") or
                        self.offer_by_sku.get(sku, f"SKU {sku}")
                    ).strip()
                    qty = int_qty(product.get("quantity"))
                    rows.append({
                        "Артикул": offer,
                        "SKU Ozon": sku,
                        "Название": str(product.get("name") or self.name_by_sku.get(sku, "")),
                        "Кластер назначения": cluster,
                        "Количество, шт.": qty,
                        "Статус": STATUS_RU.get(state, state or "Неизвестно"),
                        "Статус API": state,
                        "Уже передано Ozon": "Да" if state in PHYSICAL_STATES else "Нет",
                        "Тип поставки": route,
                        "Точка отгрузки из API": dropoff_name,
                        "Точка для тарифа": tariff_origin,
                        "Склад назначения": storage_name,
                        "Дата отгрузки": shipment_date,
                        "Плановая дата прибытия от API": arrival_api,
                        "Плановая дата прибытия": planned_arrival,
                        "Источник плановой даты прибытия": eta_basis,
                        "Срок кросс-дока, дней": sla_days,
                        "ID заявки": oid,
                        "ID поставки": norm_id(supply.get("supply_id") or supply.get("id")),
                        "Bundle ID": bundle_id,
                        "ID точки отгрузки": dropoff_id,
                    })

            if idx % 20 == 0:
                logging.info("Обработано активных заявок: %s/%s", idx, len(list_items))

        self._supply_diagnostics["detail_errors"] = detail_errors
        self._supply_diagnostics["product_rows"] = len(rows)
        logging.info("Строк товаров в активных поставках: %s", len(rows))
        return pd.DataFrame(rows)

    def _estimate_supply_arrival(self, cluster: str, storage_warehouse: str,
                                 shipment_date: Optional[date], arrival_api: Optional[date]) -> Tuple[Optional[date], str, Optional[int]]:
        # 1. Если Ozon сам дал плановую дату прибытия — не пересчитываем её.
        if arrival_api:
            return arrival_api, "Плановая дата прибытия получена из API Ozon", None

        # 2. Иначе нужен реальный/плановый слот отгрузки и срок маршрута.
        if not shipment_date:
            return None, "API не вернул дату отгрузки; рассчитать прибытие нельзя", None

        sla_days, source = self.sla.lookup(cluster, storage_warehouse)
        if sla_days is None:
            return None, f"Нет тарифа кросс-дока для склада/кластера ({source})", None

        arrival = shipment_date + timedelta(days=sla_days)
        return arrival, f"Дата отгрузки + {sla_days} дн. кросс-дока ({source})", sla_days

    # ---------- сборка ----------
    def build(self) -> Dict[str, pd.DataFrame]:
        self.fetch_products()
        self.fetch_clusters()
        logging.info("Встроенных нормативов кросс-дока: складов=%s, кластеров=%s", len(CROSSDOCK_WAREHOUSE_DAYS), len(CROSSDOCK_CLUSTER_DAYS))
        stocks = self.fetch_stock_warehouses()
        transit = self.fetch_transit_stocks()
        sales_cluster, _ = self.sales_7d_by_cluster()
        supplies = self.supply_orders()
        buyout = self.buyout_latest_by_article()
        logging.info("8/8 Формирование сводки")

        by_article = self._build_by_article(stocks, transit, sales_cluster, supplies, buyout)
        by_cluster = self._build_by_cluster(stocks, transit, sales_cluster, supplies)
        method = self._methodology()
        cluster_ref = self._cluster_reference()
        diag = self._diagnostics(stocks, supplies)
        return {
            "По артикулам": by_article,
            "По кластерам": by_cluster,
            "Поставки и прогноз": supplies,
            "Выкуп - последние 50": buyout,
            "Нормативы кросс-дока": self.sla.as_dataframe(),
            "Диагностика": diag,
            "Справочник кластеров": cluster_ref,
            "Методика": method,
        }

    def _diagnostics(self, stocks: pd.DataFrame, supplies: pd.DataFrame) -> pd.DataFrame:
        rows: List[Dict[str, Any]] = [
            {"Проверка": "Версия кода", "Значение": SCRIPT_VERSION},
            {"Проверка": "Масштаб quantity в FBO postings", "Значение": self._posting_qty_scale},
            {"Проверка": "Диагностика quantity", "Значение": json.dumps(self._posting_qty_detection, ensure_ascii=False)},
            {"Проверка": "Строк остатков", "Значение": 0 if stocks is None else len(stocks)},
            {"Проверка": "Складов без определённого кластера", "Значение": len(self._unknown_warehouses)},
            {"Проверка": "Строк товаров в активных поставках", "Значение": 0 if supplies is None else len(supplies)},
            {"Проверка": "Диагностика поставок", "Значение": json.dumps(self._supply_diagnostics, ensure_ascii=False, default=str)},
            {"Проверка": "Диагностика выкупа", "Значение": json.dumps(self._buyout_diagnostics, ensure_ascii=False, default=str)},
            {"Проверка": "Источник сроков кросс-дока", "Значение": self.sla.source},
            {"Проверка": "Тариф актуален на", "Значение": self.sla.updated_at},
            {"Проверка": "Нормативов по конечным складам", "Значение": len(CROSSDOCK_WAREHOUSE_DAYS)},
            {"Проверка": "Нормативов по кластерам", "Значение": len(CROSSDOCK_CLUSTER_DAYS)},
        ]
        for wh in sorted(self._unknown_warehouses):
            rows.append({"Проверка": "Склад без кластера", "Значение": wh})
        return pd.DataFrame(rows)

    def _build_by_article(self, stocks: pd.DataFrame, transit: pd.DataFrame, sales_cluster: pd.DataFrame,
                          supplies: pd.DataFrame, buyout: pd.DataFrame) -> pd.DataFrame:
        articles = set()
        for df in (stocks, transit, sales_cluster, supplies):
            if df is not None and not df.empty and "Артикул" in df.columns:
                articles |= set(str(x) for x in df["Артикул"].dropna() if str(x).strip())
        rows: List[Dict[str, Any]] = []
        for article in sorted(articles):
            s = stocks[stocks["Артикул"] == article] if not stocks.empty else pd.DataFrame()
            t = transit[transit["Артикул"] == article] if not transit.empty else pd.DataFrame()
            sc = sales_cluster[sales_cluster["Артикул"] == article] if not sales_cluster.empty else pd.DataFrame()
            sp = supplies[supplies["Артикул"] == article] if not supplies.empty else pd.DataFrame()
            bo = buyout[buyout["Артикул"] == article] if buyout is not None and not buyout.empty else pd.DataFrame()
            buyout_pct = float(bo.iloc[0]["Выкуп, % (последние завершённые)"]) if not bo.empty else None
            buyout_base = int(bo.iloc[0]["База выкупа, шт."]) if not bo.empty else 0
            current = int(s["Доступно сейчас, шт."].sum()) if not s.empty else 0
            ret = int(t["Возврат покупателя в пути, шт."].sum()) if not t.empty else 0
            internal = int(t["Перемещение внутри Ozon, шт."].sum()) if not t.empty else 0
            physical = int(sp.loc[sp["Уже передано Ozon"] == "Да", "Количество, шт."].sum()) if not sp.empty else 0
            requests = int(sp.loc[sp["Уже передано Ozon"] == "Нет", "Количество, шт."].sum()) if not sp.empty else 0
            supply_api_qty = int(sp["Количество, шт."].sum()) if not sp.empty else 0
            promised = int(s["Подтверждено Ozon в поставках, шт."].sum()) if (not s.empty and "Подтверждено Ozon в поставках, шт." in s.columns) else 0
            # promised_amount может дублировать supply-order. Поэтому используем его только
            # как резерв на НЕПОКРЫТУЮ часть, если API заявок вернул меньше товара.
            promised_fallback = max(0, promised - supply_api_qty)
            future = current + ret + internal + supply_api_qty + promised_fallback
            sold7 = int(sc["Заказано за 7 дней, шт."].sum()) if not sc.empty else 0
            avg = round(sold7 / 7, 1)
            days_now = int(round(current / avg)) if avg > 0 else None
            days_future = int(round(future / avg)) if avg > 0 else None
            name = ""
            if not s.empty and "Название" in s.columns:
                name = next((str(x) for x in s["Название"] if str(x).strip()), "")
            if not name:
                skus = set(s["SKU Ozon"].astype(str)) if not s.empty else set()
                for sku in skus:
                    if self.name_by_sku.get(sku):
                        name = self.name_by_sku[sku]
                        break
            rows.append({
                "Артикул": article,
                "Выкуп, % (последние завершённые)": buyout_pct,
                "База выкупа, шт.": buyout_base,
                "Среднесуточные продажи за 7 дней, шт.": avg,
                "Запас, дней: сейчас (с учётом пути и заявок)": f"{days_now if days_now is not None else '—'} ({days_future if days_future is not None else '—'})",
                "Остаток, шт.: сейчас (с учётом пути и заявок)": f"{current} ({future})",
                "Название товара": name,
                "Заказано за 7 дней, шт.": sold7,
                "Доступно к продаже сейчас, шт.": current,
                "Возвраты покупателей в пути на склад Ozon, шт.": ret,
                "Перемещение между складами Ozon, шт.": internal,
                "Поставки уже переданы Ozon, шт.": physical,
                "Активные заявки, ещё не переданы Ozon, шт.": requests,
                "Подтверждено Ozon, но не найдено в API заявок, шт.": promised_fallback,
                "Итого с учётом пути и заявок, шт.": future,
                "Запас сейчас, дней": days_now,
                "Запас с учётом пути и заявок, дней": days_future,
                "Дата отчёта": self.target_date,
            })
        df = pd.DataFrame(rows)
        if not df.empty:
            df = df.sort_values(["Среднесуточные продажи за 7 дней, шт.", "Запас сейчас, дней"], ascending=[False, True], na_position="last").reset_index(drop=True)
        return df

    def _build_by_cluster(self, stocks: pd.DataFrame, transit: pd.DataFrame, sales: pd.DataFrame,
                          supplies: pd.DataFrame) -> pd.DataFrame:
        pairs = set()
        for df, ccol in ((stocks, "Кластер"), (sales, "Кластер"), (supplies, "Кластер назначения")):
            if df is not None and not df.empty and "Артикул" in df.columns and ccol in df.columns:
                pairs |= set((str(a), str(c)) for a, c in zip(df["Артикул"], df[ccol]) if str(a).strip() and str(c).strip())
        rows: List[Dict[str, Any]] = []
        for article, cluster_raw in sorted(pairs):
            cluster = self.cluster_maps.canonical(cluster_raw)
            s = stocks[(stocks["Артикул"] == article) & (stocks["Кластер"].map(self.cluster_maps.canonical) == cluster)] if not stocks.empty else pd.DataFrame()
            sale = sales[(sales["Артикул"] == article) & (sales["Кластер"].map(self.cluster_maps.canonical) == cluster)] if not sales.empty else pd.DataFrame()
            sp = supplies[(supplies["Артикул"] == article) & (supplies["Кластер назначения"].map(self.cluster_maps.canonical) == cluster)] if not supplies.empty else pd.DataFrame()
            # Внутренний транзит/возврат относим к кластеру только когда сам API
            # вернул склад/кластер. Пустой кластер не размазываем пропорционально.
            t = pd.DataFrame()
            if transit is not None and not transit.empty and "Кластер" in transit.columns:
                t = transit[(transit["Артикул"] == article) & (transit["Кластер"].astype(str).str.strip() != "") & (transit["Кластер"].map(self.cluster_maps.canonical) == cluster)]
            current = int(s["Доступно сейчас, шт."].sum()) if not s.empty else 0
            cluster_returns = int(t["Возврат покупателя в пути, шт."].sum()) if not t.empty else 0
            cluster_internal = int(t["Перемещение внутри Ozon, шт."].sum()) if not t.empty else 0
            sold7 = int(sale["Заказано за 7 дней, шт."].sum()) if not sale.empty else 0
            avg = round(sold7 / 7, 1)
            future_supply = int(sp["Количество, шт."].sum()) if not sp.empty else 0
            promised = int(s["Подтверждено Ozon в поставках, шт."].sum()) if (not s.empty and "Подтверждено Ozon в поставках, шт." in s.columns) else 0
            promised_fallback = max(0, promised - future_supply)
            future_total = current + cluster_returns + cluster_internal + future_supply + promised_fallback
            days_now = int(round(current / avg)) if avg > 0 else None
            days_future = int(round(future_total / avg)) if avg > 0 else None
            depletion = self.target_date + timedelta(days=math.ceil(current / avg)) if avg > 0 else None

            # Существующие поставки событиями по ETA.
            events: List[Tuple[date, int, str]] = []
            if not sp.empty:
                for _, r in sp.iterrows():
                    eta = r.get("Плановая дата прибытия")
                    if isinstance(eta, pd.Timestamp):
                        eta = eta.date()
                    elif not isinstance(eta, date):
                        eta = parse_any_date(eta)
                    if eta:
                        events.append((eta, int_qty(r.get("Количество, шт.")), str(r.get("Статус") or "")))
            events.sort(key=lambda x: x[0])

            sla_days, sla_source = self.sla.lookup(cluster, "")
            normal_lead = PREPARATION_DAYS + sla_days if sla_days is not None else None
            hypothetical_eta = self.target_date + timedelta(days=normal_lead) if normal_lead is not None else None

            # Если supply-order не раскрыл всю поставку, но stock API показывает promised_amount,
            # не теряем этот товар. Для неизвестного статуса используем консервативный ETA
            # новой поставки: +4 дня на сборку + срок кросс-дока.
            if promised_fallback > 0 and hypothetical_eta is not None:
                events.append((hypothetical_eta, promised_fallback, "Подтверждено Ozon, детали заявки не получены"))
                events.sort(key=lambda x: x[0])

            oos_days_existing, nearest_eta, nearest_qty, nearest_status = self._simulate_oos(current, avg, events)
            if events:
                forecast_eta = nearest_eta
                forecast_qty = nearest_qty
                forecast_status = nearest_status
                oos_days = oos_days_existing
                forecast_basis = "Учитываем существующую заявку/поставку"
            else:
                forecast_eta = hypothetical_eta
                forecast_qty = 0
                forecast_status = "Новой заявки ещё нет"
                if avg > 0 and hypothetical_eta and depletion:
                    oos_days = max(0, int(math.ceil((hypothetical_eta - depletion).total_seconds() / 86400)))
                else:
                    oos_days = None
                forecast_basis = f"Если начать сборку сегодня: {PREPARATION_DAYS} дн. сборка + кросс-док" if hypothetical_eta else "Нет срока маршрута"

            if avg <= 0:
                will_make = "Продаж за 7 дней нет"
            elif forecast_eta is None or depletion is None:
                will_make = "Нет данных"
            else:
                will_make = "Да" if forecast_eta <= depletion else "Нет"

            rows.append({
                "Артикул": article,
                "Кластер": cluster,
                "Среднесуточные продажи за 7 дней, шт.": avg,
                "Запас, дней: сейчас (с учётом заявок)": f"{days_now if days_now is not None else '—'} ({days_future if days_future is not None else '—'})",
                "Остаток, шт.: сейчас (с учётом заявок)": f"{current} ({future_total})",
                "Заказано за 7 дней, шт.": sold7,
                "Остаток сейчас в кластере, шт.": current,
                "Возвраты покупателей в пути в кластер, шт.": cluster_returns,
                "Перемещение внутри Ozon в кластер, шт.": cluster_internal,
                "В активных поставках в кластер, шт.": future_supply,
                "Подтверждено Ozon, но не найдено в API заявок, шт.": promised_fallback,
                "Остаток с учётом пути и поставок, шт.": future_total,
                "Запас сейчас, дней": days_now,
                "Запас с учётом заявок, дней": days_future,
                "Дата ожидаемого окончания остатка": depletion if depletion else None,
                "Ближайшее пополнение, шт.": forecast_qty,
                "Статус ближайшего пополнения": forecast_status,
                "Плановая дата прибытия ближайшего пополнения": forecast_eta,
                "Срок кросс-дока из Москвы — Кавказский, дней": sla_days,
                "Срок до прибытия новой поставки: 4 дня сборки + кросс-док, дней": normal_lead,
                "Источник срока кросс-дока": sla_source,
                "Успеет пополнение до дефицита": will_make,
                "Прогноз дней без остатка": oos_days,
                "Основание прогноза": forecast_basis,
                "Дата отчёта": self.target_date,
            })
        df = pd.DataFrame(rows)
        if not df.empty:
            # Критичные кластеры сверху: сначала OOS, затем короткий запас.
            rank = df["Успеет пополнение до дефицита"].map({"Нет": 0, "Нет данных": 1, "Да": 2, "Продаж за 7 дней нет": 3}).fillna(4)
            df = df.assign(_rank=rank).sort_values(["_rank", "Запас сейчас, дней", "Среднесуточные продажи за 7 дней, шт."], ascending=[True, True, False], na_position="last").drop(columns="_rank").reset_index(drop=True)
        return df

    def _simulate_oos(self, current: int, avg: float, events: List[Tuple[date, int, str]]) -> Tuple[Optional[int], Optional[date], int, str]:
        if not events:
            return None, None, 0, ""
        nearest_eta, nearest_qty, nearest_status = events[0]
        if avg <= 0:
            return 0, nearest_eta, nearest_qty, nearest_status
        inv = float(current)
        cursor = self.target_date
        oos_days = 0.0
        for eta, qty, _status in events:
            if eta < cursor:
                eta = cursor
            delta = (eta - cursor).days
            need = avg * delta
            if inv >= need:
                inv -= need
            else:
                covered = inv / avg if avg > 0 else delta
                oos_days += max(0.0, delta - covered)
                inv = 0.0
            inv += qty
            cursor = eta
        return int(math.ceil(oos_days)), nearest_eta, nearest_qty, nearest_status

    def _cluster_reference(self) -> pd.DataFrame:
        rows = []
        for wh, cluster in sorted(self.cluster_maps.warehouse_name_to_cluster.items()):
            rows.append({"Источник": "API", "Склад / правило": wh, "ID макролокального кластера": "", "Кластер": cluster})
        for mid, cluster in sorted(self.cluster_maps.macro_to_cluster.items()):
            rows.append({"Источник": "API", "Склад / правило": "", "ID макролокального кластера": mid, "Кластер": cluster})
        for tokens, cluster in STATIC_WAREHOUSE_CLUSTER_PATTERNS:
            rows.append({"Источник": "Резерв по названию склада", "Склад / правило": " + ".join(tokens), "ID макролокального кластера": "", "Кластер": cluster})
        return pd.DataFrame(rows)

    def _methodology(self) -> pd.DataFrame:
        return pd.DataFrame([
            {"Показатель": "Среднесуточные продажи за 7 дней", "Как считаем": "Заказанные FBO-штуки за последние 7 календарных дней / 7. quantity FBO автоматически проверяется на масштаб ×1000 и нормализуется только если структура данных это подтверждает."},
            {"Показатель": "Остаток сейчас", "Как считаем": "free_to_sell_amount — реально доступно к продаже на FBO-складах Ozon."},
            {"Показатель": "Остаток в скобках по артикулу", "Как считаем": "Остаток сейчас + возвраты покупателей в пути + внутренние перемещения Ozon + все активные заявки продавца, включая ещё не переданные Ozon."},
            {"Показатель": "Кластер", "Как считаем": "Сначала кластер из API, затем справочник склад→кластер, затем резерв по географическому названию склада. Физическое имя РФЦ не становится отдельным кластером."},
            {"Показатель": "Выкуп, % (последние завершённые)", "Как считаем": f"Для каждого артикула берём до {BUYOUT_TARGET_UNITS} самых свежих завершённых единиц за доступное окно {BUYOUT_LOOKBACK_DAYS} дней. status=cancelled — невыкуп; status=delivered — выкуп, кроме количества, сопоставленного с type=ClientReturn по posting_number + SKU. Cancellation из returns/list повторно не вычитаем. Незавершённые отправления не участвуют. Если завершённых единиц меньше {BUYOUT_TARGET_UNITS}, процент считается по фактической базе, её размер показан рядом."},
            {"Показатель": "Активные поставки", "Как считаем": "Список заявок: /v3/supply-order/list с фильтром по активным статусам; метод возвращает order_ids. Детали: /v3/supply-order/get по order_ids. Состав: /v1/supply-order/bundle."},
            {"Показатель": "promised_amount", "Как считаем": "Используем только как резерв: прибавляем часть promised_amount, которая не покрыта найденными supply-order, чтобы не считать поставку дважды."},
            {"Показатель": "Дата отгрузки", "Как считаем": "Берём начало timeslot поставки из /v3/supply-order/get. Если timeslot отсутствует, дату не выдумываем."},
            {"Показатель": "Плановая дата прибытия", "Как считаем": "Если API Ozon вернул arrival_date/плановую дату — используем её. Иначе: дата отгрузки + верхняя граница срока кросс-дока из переданного тарифа Ozon от 03.08.2026."},
            {"Показатель": "Срок кросс-дока", "Как считаем": f"Для FINICK источник — {CROSSDOCK_ORIGIN}. Сначала берём норматив конкретного конечного склада, если он известен; иначе консервативный максимум по кластеру."},
            {"Показатель": "Новая поставка без заявки", "Как считаем": f"Для оценки, успеет ли новая поставка: {PREPARATION_DAYS} календарных дня на сборку + срок кросс-дока. Дополнительные 5 дней приёмки больше не добавляются."},
            {"Показатель": "Прогноз дней без остатка", "Как считаем": "Сравниваем текущий запас в днях с плановой датой прибытия ближайшего известного пополнения; если заявки нет — с расчётным прибытием новой поставки."},
        ])

    # ---------- Excel / S3 ----------
    def save(self, sheets: Dict[str, pd.DataFrame]) -> Tuple[str, str]:
        data = self._xlsx(sheets)
        prefix = f"Сводные отчёты/{self.store}"
        dated = f"{prefix}/Остатки_и_оборачиваемость_{self.target_date.isoformat()}.xlsx"
        latest = f"{prefix}/Последний.xlsx"
        content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        self.storage.upload(dated, data, content_type)
        self.storage.upload(latest, data, content_type)
        status_key = f"{prefix}/Служебные/Последний_запуск.json"
        self.storage.upload_json(status_key, {
            "status": "OK",
            "script_version": SCRIPT_VERSION,
            "store": self.store,
            "date": self.target_date.isoformat(),
            "created_at": datetime.now(MOSCOW_TZ).isoformat(),
            "file": dated,
            "rows_articles": len(sheets.get("По артикулам", [])),
            "rows_clusters": len(sheets.get("По кластерам", [])),
            "rows_supplies": len(sheets.get("Поставки и прогноз", [])),
            "crossdock_tariff_source": self.sla.source,
            "crossdock_tariff_as_of": self.sla.updated_at,
            "posting_quantity_scale": self._posting_qty_scale,
            "unknown_warehouses": sorted(self._unknown_warehouses),
            "supply_diagnostics": self._supply_diagnostics,
        })
        return dated, latest

    def _xlsx(self, sheets: Dict[str, pd.DataFrame]) -> bytes:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            for sheet_name, source in sheets.items():
                df = source.copy() if source is not None else pd.DataFrame()
                if df.empty:
                    df = pd.DataFrame([{"Статус": "Нет данных"}])
                # Даты оставляем датами — Excel их нормально форматирует.
                df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
                ws = writer.book[sheet_name[:31]]
                ws.auto_filter.ref = ws.dimensions
                ws.freeze_panes = "A2"
                fill = PatternFill("solid", fgColor="1F4E78")
                font = Font(color="FFFFFF", bold=True)
                for cell in ws[1]:
                    cell.fill = fill
                    cell.font = font
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                for col_idx, col in enumerate(df.columns, 1):
                    header = str(col)
                    vals = [len(header)] + [len(str(x)) for x in df[col].dropna().astype(str).head(100)]
                    ws.column_dimensions[get_column_letter(col_idx)].width = min(52, max(10, max(vals, default=10) + 2))
                    if "Среднесуточные" in header or "%" in header:
                        fmt = "0.0"
                    elif any(x in header for x in ("шт.", "дней")) and "Запас, дней:" not in header:
                        fmt = "#,##0"
                    elif "Дата" in header or "Прогноз" in header or "Слот" in header:
                        fmt = "DD.MM.YYYY"
                    else:
                        fmt = None
                    if fmt:
                        for cell in ws[get_column_letter(col_idx)][1:]:
                            cell.number_format = fmt
                if sheet_name in {"По артикулам", "По кластерам"}:
                    ws.freeze_panes = "E2"
                    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=min(5, ws.max_column)):
                        for cell in row:
                            if cell.column in {1, 3, 4}:
                                cell.font = Font(bold=True)
        return buf.getvalue()


def env_first(*names: str, default: str = "") -> str:
    """Возвращает первое непустое значение переменной окружения из списка."""
    for name in names:
        value = (os.getenv(name, "") or "").strip()
        if value:
            return value
    return default


def required_env(store: str, suffix: str) -> str:
    options = [f"OZON_{suffix}_{store}", f"OZON_{store}_{suffix}"]
    value = env_first(*options)
    if value:
        return value
    raise RuntimeError(f"Не найден секрет: {' / '.join(options)}")


def main() -> int:
    load_report_env()
    parser = argparse.ArgumentParser(description="Ozon — отдельный сводный отчёт остатков и оборачиваемости")
    parser.add_argument("--store", default=os.getenv("OZON_STORE", "FINICK"))
    parser.add_argument("--target-date", default=os.getenv("OZON_TARGET_DATE", ""))
    args = parser.parse_args()
    store = args.store.upper().strip()
    target = resolve_target_date(args.target_date)

    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S")
    logging.info("=== Остатки и оборачиваемость | %s | %s ===", store, target)

    client_id = required_env(store, "CLIENT_ID")
    api_key = required_env(store, "API_KEY")
    # Поддерживаем все варианты имён, которые использовались в предыдущих
    # Ozon-сборщиках и в GitHub Secrets. Это позволяет запускать отдельный
    # отчёт без создания новых секретов.
    access = env_first("OZON_YC_ACCESS_KEY_ID", "YC_ACCESS_KEY_ID", "AWS_ACCESS_KEY_ID")
    secret = env_first("OZON_YC_SECRET_ACCESS_KEY", "YC_SECRET_ACCESS_KEY", "AWS_SECRET_ACCESS_KEY")
    if not access or not secret:
        raise RuntimeError(
            "Не найдены ключи Object Storage. Поддерживаются: "
            "OZON_YC_ACCESS_KEY_ID/SECRET_ACCESS_KEY, "
            "YC_ACCESS_KEY_ID/YC_SECRET_ACCESS_KEY или AWS_ACCESS_KEY_ID/AWS_SECRET_ACCESS_KEY"
        )
    bucket = env_first("OZON_YC_BUCKET", "YC_BUCKET_NAME", default=DEFAULT_BUCKET) or DEFAULT_BUCKET
    endpoint = env_first(
        "OZON_YC_ENDPOINT_URL", "YC_ENDPOINT_URL",
        default="https://storage.yandexcloud.net"
    )

    storage = Storage(access, secret, bucket, endpoint)
    client = OzonClient(client_id, api_key)
    builder = ReportBuilder(store, target, client, storage)
    sheets = builder.build()
    dated, latest = builder.save(sheets)
    logging.info("ГОТОВО: s3://%s/%s", bucket, dated)
    logging.info("Последняя версия: s3://%s/%s", bucket, latest)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception:
        logging.exception("Ошибка формирования сводного отчёта")
        raise
