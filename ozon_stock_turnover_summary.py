#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Отдельный сводный отчёт Ozon FBO: остатки, оборачиваемость, кластеры и прогноз OOS.

Это САМОСТОЯТЕЛЬНЫЙ скрипт. Он не запускает рекламу, финансы, отзывы и прочие
модули большого сборщика.

Главный результат в Object Storage:
  Сводные отчёты/<STORE>/Остатки_и_оборачиваемость_YYYY-MM-DD.xlsx
  Сводные отчёты/<STORE>/Последний.xlsx

Что считает:
- среднесуточные продажи за 7 дней по артикулу и по кластеру;
- доступный остаток FBO сейчас;
- товары в пути внутри Ozon и возвраты покупателей на склад;
- все активные FBO-заявки продавца (и уже переданные Ozon, и ещё ожидающие отгрузки);
- запас в днях сейчас и с учётом пути/активных заявок;
- процент выкупа за 30 дней без ранних отмен и без обычных ClientReturn;
- прогноз дефицита по кластерам;
- норматив кросс-дока из «Москва — Кавказский» из Базы знаний Ozon + 4 дня на сборку.

Важно по кросс-доку:
с 16.02.2026 Ozon для кросс-докинговых поставок не возвращает
storage_warehouse.arrival_date. Поэтому прогноз строится по сроку маршрута из
официальной Базы знаний Ozon и фактическому статусу/слоту существующей заявки.
Если срок маршрута не удалось прочитать ни из БЗ, ни из кэша/override, скрипт
НЕ выдумывает срок и помечает прогноз как «Нет срока маршрута».
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
from html import unescape
from typing import Any, Dict, Iterable, List, Mapping, Optional, Sequence, Tuple
from urllib.parse import urljoin
from zoneinfo import ZoneInfo

import boto3
import pandas as pd
import requests
from botocore.client import Config
from botocore.exceptions import ClientError
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

SCRIPT_VERSION = "OZON_STOCK_TURNOVER_SUMMARY_STANDALONE_V1_4_20260807"
MOSCOW_TZ = ZoneInfo("Europe/Moscow")
OZON_API_BASE = "https://api-seller.ozon.ru"
DEFAULT_BUCKET = "ozon-assist"
SELLER_MIN_INTERVAL_SECONDS = 0.55
CROSSDOCK_ORIGIN = "Москва — Кавказский"
CROSSDOCK_KB_URL = "https://seller-edu.ozon.ru/fbo/crossdoking/kross-doking"
PREPARATION_DAYS = 4

ACTIVE_STATES = {
    "DATA_FILLING",
    "READY_TO_SUPPLY",
    "ACCEPTED_AT_SUPPLY_WAREHOUSE",
    "IN_TRANSIT",
    "ACCEPTANCE_AT_STORAGE_WAREHOUSE",
    "REPORTS_CONFIRMATION_AWAITING",
    "REPORT_REJECTED",
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
REFUSAL_MARKERS = (
    "отказался при вручении", "отказалась при вручении", "отказ при вручении",
    "не забрал", "не забрала", "не забрали", "истек срок хранения",
    "истёк срок хранения", "срок хранения", "отказ от получения",
    "отказался от получения", "отказалась от получения", "не пришел",
    "не пришёл", "не пришла", "не востребован", "невыкуп",
)


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


class CrossdockSLAProvider:
    def __init__(self, storage: Storage, store: str):
        self.storage = storage
        self.store = store
        self.cache_key = f"Сводные отчёты/{store}/Справочники/Сроки_кроссдока_Кавказский.json"
        self.rows: Dict[str, int] = {}
        self.source = ""
        self.updated_at = ""

    def load(self) -> Dict[str, int]:
        # 1) Пытаемся каждый запуск прочитать актуальную официальную БЗ Ozon.
        try:
            mapping = self._load_from_ozon_kb()
            if mapping:
                self.rows = mapping
                self.source = "База знаний Ozon — Москва — Кавказский"
                self.updated_at = datetime.now(MOSCOW_TZ).isoformat()
                self.storage.upload_json(self.cache_key, {
                    "source": self.source,
                    "url": CROSSDOCK_KB_URL,
                    "updated_at": self.updated_at,
                    "origin": CROSSDOCK_ORIGIN,
                    "sla_days": mapping,
                })
                return mapping
        except Exception as exc:
            logging.warning("Не удалось обновить сроки кросс-дока из БЗ Ozon: %s", exc)

        # 2) Используем последний успешно сохранённый кэш.
        cached = self.storage.read_json(self.cache_key)
        cached_map = cached.get("sla_days") if isinstance(cached, Mapping) else None
        if isinstance(cached_map, Mapping) and cached_map:
            self.rows = {str(k): int(v) for k, v in cached_map.items() if conservative_days(v) is not None}
            self.source = str(cached.get("source") or "Кэш Базы знаний Ozon")
            self.updated_at = str(cached.get("updated_at") or "")
            if self.rows:
                return self.rows

        # 3) Явный резервный override. Не скрываем его происхождение.
        raw = os.getenv("OZON_CROSSDOCK_SLA_JSON", "").strip()
        if raw:
            try:
                obj = json.loads(raw)
                if isinstance(obj, Mapping):
                    self.rows = {str(k): int(conservative_days(v)) for k, v in obj.items() if conservative_days(v) is not None}
                    self.source = "Ручной резерв OZON_CROSSDOCK_SLA_JSON"
                    self.updated_at = datetime.now(MOSCOW_TZ).isoformat()
            except Exception as exc:
                logging.warning("Некорректный OZON_CROSSDOCK_SLA_JSON: %s", exc)
        return self.rows

    def lookup(self, cluster: str) -> Tuple[Optional[int], str]:
        if not cluster or not self.rows:
            return None, self.source or "Срок не найден"
        best_key, best_score = "", 0.0
        for key in self.rows:
            score = similarity(cluster, key)
            if score > best_score:
                best_key, best_score = key, score
        if best_key and best_score >= 0.55:
            return int(self.rows[best_key]), self.source
        return None, self.source or "Срок не найден"

    def _load_from_ozon_kb(self) -> Dict[str, int]:
        headers = {"User-Agent": "Mozilla/5.0 (compatible; OzonAssist/1.0)"}
        r = requests.get(CROSSDOCK_KB_URL, headers=headers, timeout=45, allow_redirects=True)
        r.raise_for_status()
        html = r.text
        tables: List[pd.DataFrame] = []
        try:
            tables.extend(pd.read_html(io.StringIO(html)))
        except Exception:
            pass

        # Ищем прикреплённые XLSX/CSV — в БЗ таблица сроков иногда вынесена в файл.
        links = set(re.findall(r'''(?:href|src)=["']([^"']+\.(?:xlsx?|csv)(?:\?[^"']*)?)["']''', html, flags=re.I))
        for href in links:
            try:
                url = urljoin(r.url, unescape(href))
                fr = requests.get(url, headers=headers, timeout=45)
                fr.raise_for_status()
                if re.search(r"\.csv(?:\?|$)", url, re.I):
                    tables.append(pd.read_csv(io.BytesIO(fr.content)))
                else:
                    book = pd.ExcelFile(io.BytesIO(fr.content))
                    for sheet in book.sheet_names:
                        tables.append(pd.read_excel(book, sheet_name=sheet, header=None))
            except Exception:
                continue

        # Некоторые страницы хранят таблицу в JSON/HTML без <table>. Превращаем
        # фрагменты вокруг «Кавказский» в псевдотаблицу, если удаётся найти пары.
        mapping: Dict[str, int] = {}
        for table in tables:
            mapping.update(self._extract_from_table(table))
        if mapping:
            return mapping

        text = unescape(re.sub(r"<[^>]+>", " ", html))
        text = re.sub(r"\s+", " ", text)
        # Резервный парсер только для явных конструкций «назначение ... N дней»
        # рядом с Кавказским. Не пытается угадывать числа вне контекста дней.
        pos = norm_name(text).find(norm_name(CROSSDOCK_ORIGIN))
        if pos >= 0:
            fragment = text[max(0, pos - 3000): pos + 30000]
            for m in re.finditer(r"([А-ЯЁA-Z][А-Яа-яЁёA-Za-z0-9 ._\-/]{2,80})\s*[-–—:]?\s*(\d{1,2}(?:\s*[-–—]\s*\d{1,2})?)\s*(?:календарн\w*\s*)?дн", fragment, re.I):
                d = conservative_days(m.group(2))
                name = m.group(1).strip(" -–—:;,")
                if d is not None and name and "кавказ" not in norm_name(name):
                    mapping[name] = d
        return mapping

    def _extract_from_table(self, table: pd.DataFrame) -> Dict[str, int]:
        if table is None or table.empty:
            return {}
        df = table.copy().fillna("")
        df.columns = [str(c).strip() for c in df.columns]
        result: Dict[str, int] = {}

        # Layout 1: одна строка = маршрут (откуда, куда, срок).
        cols = list(df.columns)
        origin_cols = [c for c in cols if any(x in norm_name(c) for x in ("откуда", "пункт отправ", "точка отправ", "склад отправ"))]
        dest_cols = [c for c in cols if any(x in norm_name(c) for x in ("куда", "назначен", "кластер", "склад получ"))]
        day_cols = [c for c in cols if any(x in norm_name(c) for x in ("срок", "дней", "дни", "доставка"))]
        if origin_cols and dest_cols and day_cols:
            for _, row in df.iterrows():
                origin = str(row.get(origin_cols[0], ""))
                if "кавказ" not in norm_name(origin):
                    continue
                dest = str(row.get(dest_cols[0], "")).strip()
                days = conservative_days(row.get(day_cols[0]))
                if dest and days is not None:
                    result[dest] = days
            if result:
                return result

        # Layout 2: матрица — строка «Москва — Кавказский», колонки = назначения.
        for idx in range(len(df)):
            values = [str(x) for x in df.iloc[idx].tolist()]
            origin_index = next((j for j, x in enumerate(values) if "кавказ" in norm_name(x) and "моск" in norm_name(x)), None)
            if origin_index is None:
                continue
            # В качестве заголовков проверяем несколько строк выше и сами columns.
            header_candidates: List[List[str]] = [[str(c) for c in df.columns]]
            for h in range(max(0, idx - 3), idx):
                header_candidates.append([str(x) for x in df.iloc[h].tolist()])
            headers = max(header_candidates, key=lambda hs: sum(bool(norm_name(x)) for x in hs))
            for j, value in enumerate(values):
                if j == origin_index:
                    continue
                days = conservative_days(value)
                dest = headers[j].strip() if j < len(headers) else ""
                if days is not None and dest and not dest.isdigit() and "unnamed" not in dest.lower():
                    result[dest] = days
        return result


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
        self.sla = CrossdockSLAProvider(storage, self.store)

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
        # v1.4: используем актуальный /v2/cluster/list. Его структура:
        # clusters[] -> logistic_clusters[] -> warehouses[].
        # В v1.3 этот ответ парсился как старая v1-структура, поэтому при успешном
        # ответе v2 мы всё равно получали «Кластеров: 0».
        try:
            response = self.client.post("/v2/cluster/list", {})
        except OzonApiError as exc:
            logging.warning("/v2/cluster/list недоступен: %s", exc)
            return

        clusters = extract_items(response, [("clusters",), ("result", "clusters"), ("result", "items"), ("items",)])
        # На случай изменения обёртки ищем словари, похожие именно на кластер v2.
        if not clusters:
            clusters = [dict(x) for x in walk_dicts(response)
                        if isinstance(x, Mapping) and x.get("name") and isinstance(x.get("logistic_clusters"), list)]

        names: List[str] = []
        wh_id: Dict[str, str] = {}
        wh_name: Dict[str, str] = {}
        macro: Dict[str, str] = {}

        for c in clusters:
            cname = str(c.get("name") or c.get("cluster_name") or c.get("macrolocal_cluster_name") or "").strip()
            if not cname:
                continue
            if cname not in names:
                names.append(cname)
            cid = norm_id(c.get("id") or c.get("cluster_id") or c.get("macrolocal_cluster_id") or c.get("macro_cluster_id"))
            if cid:
                macro[cid] = cname

            logistic_clusters = c.get("logistic_clusters")
            if not isinstance(logistic_clusters, list):
                logistic_clusters = []
            # Совместимость, если Ozon когда-нибудь снова вернёт warehouses прямо в кластере.
            direct_warehouses = c.get("warehouses")
            if isinstance(direct_warehouses, list):
                logistic_clusters = list(logistic_clusters) + [{"warehouses": direct_warehouses}]

            for lc in logistic_clusters:
                if not isinstance(lc, Mapping):
                    continue
                if bool(lc.get("is_archived")):
                    continue
                warehouses = lc.get("warehouses")
                if not isinstance(warehouses, list):
                    continue
                for w in warehouses:
                    if not isinstance(w, Mapping):
                        continue
                    wid = norm_id(w.get("warehouse_id") or w.get("id"))
                    wname = str(w.get("name") or w.get("warehouse_name") or "").strip()
                    if wid:
                        wh_id[wid] = cname
                    if wname:
                        wh_name[norm_name(wname)] = cname

        self.cluster_maps = ClusterMaps(names, wh_id, wh_name, macro, "/v2/cluster/list")
        logging.info("Кластеров: %s, складов с привязкой: %s", len(names), len(wh_id) + len(wh_name))
        if not names:
            top_keys = list(response.keys()) if isinstance(response, Mapping) else []
            logging.warning("/v2/cluster/list ответил, но кластеры не распознаны. Верхние поля ответа: %s", top_keys)

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
            cluster = self.cluster_maps.warehouse_to_cluster.get(wh_id, "")
            if not cluster and wh_name:
                cluster = self.cluster_maps.warehouse_name_to_cluster.get(norm_name(wh_name), "")
            if not cluster:
                # Последняя попытка — название склада иногда уже содержит название кластера.
                cluster = self.cluster_maps.canonical(wh_name) if wh_name else "Кластер не определён"
            out.append({
                "SKU Ozon": sku,
                "Артикул": offer or f"SKU {sku}",
                "Название": name,
                "Кластер": cluster or "Кластер не определён",
                "Склад Ozon": wh_name,
                "Доступно сейчас, шт.": int_qty(r.get("free_to_sell_amount")),
                "Подтверждено Ozon в поставках, шт.": int_qty(r.get("promised_amount")),
                "Зарезервировано Ozon, шт.": int_qty(r.get("reserved_amount")),
            })
            if offer:
                self.offer_by_sku[sku] = offer
            if name:
                self.name_by_sku[sku] = name
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
    def fetch_postings(self, start: date, end: date) -> List[Dict[str, Any]]:
        # v3 основной. Старый v2 не используем: с августа 2026 он отключается.
        variants = [
            {
                "dir": "ASC", "filter": {"since": to_ozon_datetime(start), "to": to_ozon_datetime(end, True), "status": ""},
                "limit": 100, "offset": 0,
                "with": {"analytics_data": True, "financial_data": True, "translit": True},
            },
            {
                "dir": "ASC", "filter": {"since": to_ozon_datetime(start), "to": to_ozon_datetime(end, True), "status": ""},
                "limit": 100, "offset": 0,
            },
        ]
        last: Optional[Exception] = None
        for payload in variants:
            try:
                return self.client.offset_pages(
                    "/v3/posting/fbo/list", payload,
                    [("result", "postings"), ("postings",), ("result",)], limit=100, max_pages=1000)
            except OzonApiError as exc:
                last = exc
                if exc.status not in {400, 404, 405, 422}:
                    raise
        if last:
            raise last
        return []

    def sales_7d_by_cluster(self) -> Tuple[pd.DataFrame, List[Dict[str, Any]]]:
        logging.info("5/8 Продажи за 7 дней по кластерам")
        start = self.target_date - timedelta(days=6)
        postings = self.fetch_postings(start, self.target_date)
        rows: List[Dict[str, Any]] = []
        for p in postings:
            status = str(p.get("status") or "").lower()
            if status == "cancelled":
                continue
            financial = p.get("financial_data") if isinstance(p.get("financial_data"), Mapping) else {}
            analytics = p.get("analytics_data") if isinstance(p.get("analytics_data"), Mapping) else {}
            cluster = str(
                financial.get("cluster_to") or analytics.get("cluster_to") or
                p.get("cluster_to") or deep_first(p, ["cluster_to"] ) or "Кластер не определён"
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
                qty = max(1, int_qty(product.get("quantity") or 1))
                rows.append({"Артикул": offer, "SKU Ozon": sku, "Название": name, "Кластер": cluster, "Заказано, шт.": qty})
        df = pd.DataFrame(rows)
        if df.empty:
            return pd.DataFrame(columns=["Артикул", "Кластер", "Заказано за 7 дней, шт.", "Среднесуточные продажи за 7 дней, шт."]), postings
        agg = df.groupby(["Артикул", "Кластер"], as_index=False).agg(
            **{"Заказано за 7 дней, шт.": ("Заказано, шт.", "sum")})
        agg["Среднесуточные продажи за 7 дней, шт."] = (agg["Заказано за 7 дней, шт."] / 7).round(1)
        return agg, postings

    # ---------- выкуп ----------
    def fetch_returns(self, start: date, end: date) -> List[Dict[str, Any]]:
        rows: List[Dict[str, Any]] = []
        last_id: Any = 0
        for _ in range(300):
            payload = {
                "filter": {"logistic_return_date": {"time_from": to_ozon_datetime(start), "time_to": to_ozon_datetime(end, True)}},
                "limit": 500, "last_id": last_id,
            }
            data = self.client.post("/v1/returns/list", payload)
            got = extract_items(data, [("returns",), ("result", "returns"), ("result", "items"), ("items",)])
            rows.extend(got)
            nxt = deep_first(data, ["last_id"])
            if not got or nxt in (None, "", 0, last_id):
                break
            last_id = nxt
        return rows

    def buyout_30d(self) -> Tuple[pd.DataFrame, pd.DataFrame]:
        logging.info("6/8 Процент выкупа за 30 дней")
        start = self.target_date - timedelta(days=29)
        postings = self.fetch_postings(start, self.target_date)
        returns = self.fetch_returns(start, self.target_date)
        bought: Dict[str, int] = {}
        refused: Dict[str, int] = {}
        audit_rows: List[Dict[str, Any]] = []
        for p in postings:
            if str(p.get("status") or "").lower() != "delivered":
                continue
            for product in p.get("products") or []:
                if not isinstance(product, Mapping):
                    continue
                sku = norm_id(product.get("sku") or product.get("product_id"))
                if not sku:
                    continue
                offer = str(product.get("offer_id") or self.offer_by_sku.get(sku, f"SKU {sku}")).strip()
                qty = max(1, int_qty(product.get("quantity") or 1))
                bought[offer] = bought.get(offer, 0) + qty
        for r in returns:
            rtype = str(r.get("type") or r.get("return_type") or "").lower()
            reason = str(r.get("return_reason_name") or r.get("reason_name") or r.get("reason") or "").lower()
            include = rtype == "cancellation" and any(m in reason for m in REFUSAL_MARKERS)
            product = r.get("product") if isinstance(r.get("product"), Mapping) else {}
            sku = norm_id(product.get("sku") or r.get("sku"))
            offer = str(product.get("offer_id") or self.offer_by_sku.get(sku, f"SKU {sku}" if sku else "")).strip()
            qty = max(1, int_qty(product.get("quantity") or r.get("quantity") or 1))
            audit_rows.append({
                "Дата возврата": str(deep_first(r, ["return_date", "logistic_return_date", "final_moment"]) or "")[:10],
                "Артикул": offer,
                "SKU Ozon": sku,
                "Тип возврата": r.get("type") or r.get("return_type"),
                "Причина": r.get("return_reason_name") or r.get("reason_name") or r.get("reason"),
                "Количество, шт.": qty,
                "Считаем невыкупом": "Да" if include else "Нет",
            })
            if include and offer:
                refused[offer] = refused.get(offer, 0) + qty
        result: List[Dict[str, Any]] = []
        for offer in sorted(set(bought) | set(refused)):
            b, f = bought.get(offer, 0), refused.get(offer, 0)
            denom = b + f
            result.append({
                "Артикул": offer,
                "Выкуплено за 30 дней, шт.": b,
                "Невыкуплено после доставки покупателю, шт.": f,
                "Процент выкупа за 30 дней, %": round(b / denom * 100, 1) if denom else None,
            })
        return pd.DataFrame(result), pd.DataFrame(audit_rows)

    # ---------- поставки ----------
    @staticmethod
    def _extract_timeslot(obj: Any) -> Tuple[Optional[date], Optional[date]]:
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

    def supply_orders(self) -> pd.DataFrame:
        logging.info("7/8 Активные FBO-заявки и поставки")
        created_from = (self.target_date - timedelta(days=180)).isoformat()
        created_to = (self.target_date + timedelta(days=30)).isoformat()
        variants = [
            {"filter": {"states": sorted(ACTIVE_STATES), "created_date_from": created_from, "created_date_to": created_to},
             "limit": 100, "last_id": "", "sort_by": 1, "sort_direction": 2},
            {"filter": {"states": sorted(ACTIVE_STATES)}, "limit": 100, "last_id": "", "sort_by": 1, "sort_direction": 2},
        ]
        list_items: List[Dict[str, Any]] = []
        last_exc: Optional[Exception] = None
        for payload in variants:
            try:
                list_items = self.client.cursor_pages(
                    "/v3/supply-order/list", payload,
                    [("result", "items"), ("items",), ("result", "orders"), ("orders",)],
                    cursor_field="last_id", limit=100)
                break
            except Exception as exc:
                last_exc = exc
        if not list_items and last_exc:
            logging.warning("Список поставок не получен: %s", last_exc)
            return pd.DataFrame()

        seen: set[str] = set()
        rows: List[Dict[str, Any]] = []
        for idx, short in enumerate(list_items, 1):
            oid = norm_id(short.get("supply_order_id") or short.get("order_id") or short.get("id"))
            if not oid or oid in seen:
                continue
            seen.add(oid)
            try:
                try:
                    detail = self.client.post("/v3/supply-order/get", {"order_id": oid})
                except OzonApiError as exc:
                    if exc.status in {400, 422}:
                        detail = self.client.post("/v3/supply-order/get", {"supply_order_id": oid})
                    else:
                        raise
            except Exception as exc:
                logging.warning("Заявка %s: детали недоступны: %s", oid, exc)
                continue

            roots = []
            root = detail.get("result", detail) if isinstance(detail, Mapping) else detail
            if isinstance(root, Mapping):
                if isinstance(root.get("orders"), list):
                    roots = [x for x in root["orders"] if isinstance(x, Mapping)]
                else:
                    roots = [root]
            elif isinstance(root, list):
                roots = [x for x in root if isinstance(x, Mapping)]
            if not roots:
                roots = [short]

            for order in roots:
                order_state = str(order.get("state") or short.get("state") or "").upper()
                if order_state not in ACTIVE_STATES:
                    continue
                order_slot_from, order_slot_to = self._extract_timeslot(order.get("timeslot") or order)
                supplies = order.get("supplies") if isinstance(order.get("supplies"), list) else []
                if not supplies and isinstance(short.get("supplies"), list):
                    supplies = short.get("supplies")
                for supply in supplies:
                    if not isinstance(supply, Mapping):
                        continue
                    state = str(supply.get("supply_state") or supply.get("state") or order_state).upper()
                    if state not in ACTIVE_STATES:
                        continue
                    supply_slot_from, supply_slot_to = self._extract_timeslot(supply.get("timeslot") or supply)
                    slot_from = order_slot_from or supply_slot_from
                    slot_to = order_slot_to or supply_slot_to
                    bundle_id = norm_id(supply.get("bundle_id"))
                    storage_wh = supply.get("storage_warehouse") if isinstance(supply.get("storage_warehouse"), Mapping) else {}
                    dropoff_wh = supply.get("drop_off_warehouse") if isinstance(supply.get("drop_off_warehouse"), Mapping) else {}
                    storage_id = norm_id(supply.get("storage_warehouse_id") or storage_wh.get("warehouse_id") or storage_wh.get("id"))
                    dropoff_id = norm_id(order.get("dropoff_warehouse_id") or supply.get("dropoff_warehouse_id") or dropoff_wh.get("warehouse_id") or dropoff_wh.get("id"))
                    storage_name = str(storage_wh.get("name") or supply.get("storage_warehouse_name") or "").strip()
                    dropoff_name = str(dropoff_wh.get("name") or order.get("dropoff_warehouse_name") or "").strip()
                    macro_id = norm_id(supply.get("macrolocal_cluster_id") or order.get("macrolocal_cluster_id") or short.get("macrolocal_cluster_id"))
                    cluster = self.cluster_maps.macro_to_cluster.get(macro_id, "")
                    if not cluster and storage_id:
                        cluster = self.cluster_maps.warehouse_to_cluster.get(storage_id, "")
                    if not cluster and storage_name:
                        cluster = self.cluster_maps.warehouse_name_to_cluster.get(norm_name(storage_name), "")
                    cluster = self.cluster_maps.canonical(cluster) if cluster else "Кластер не определён"
                    route = "Кросс-док"
                    if dropoff_id and storage_id and dropoff_id == storage_id:
                        route = "Прямая"
                    elif storage_name and dropoff_name and similarity(storage_name, dropoff_name) >= 0.9:
                        route = "Прямая"
                    arrival_api = parse_any_date(storage_wh.get("arrival_date") or supply.get("arrival_date"))

                    products: List[Dict[str, Any]] = []
                    if bundle_id:
                        try:
                            products = self._bundle_items(bundle_id)
                        except Exception as exc:
                            logging.warning("Состав поставки bundle=%s: %s", bundle_id, exc)
                    if not products:
                        continue
                    for product in products:
                        sku = norm_id(product.get("sku") or product.get("product_id"))
                        if not sku:
                            continue
                        offer = str(product.get("contractor_item_code") or product.get("offer_id") or self.offer_by_sku.get(sku, f"SKU {sku}")).strip()
                        qty = int_qty(product.get("quantity"))
                        eta, eta_basis, sla_days = self._estimate_supply_eta(
                            state=state, route=route, cluster=cluster, slot_from=slot_from,
                            arrival_api=arrival_api, detail=order, supply=supply)
                        rows.append({
                            "Артикул": offer,
                            "SKU Ozon": sku,
                            "Название": str(product.get("name") or self.name_by_sku.get(sku, "")),
                            "Кластер назначения": cluster,
                            "Количество, шт.": qty,
                            "Статус": STATUS_RU.get(state, state),
                            "Статус API": state,
                            "Уже передано Ozon": "Да" if state in PHYSICAL_STATES else "Нет",
                            "Тип поставки": route,
                            "Точка отгрузки": dropoff_name or (CROSSDOCK_ORIGIN if route == "Кросс-док" else ""),
                            "Склад назначения": storage_name,
                            "Слот отгрузки / приёмки": slot_from,
                            "Дата прибытия от API": arrival_api,
                            "Прогноз прибытия": eta,
                            "Основание прогноза": eta_basis,
                            "Срок кросс-дока из Кавказского, дней": sla_days,
                            "ID заявки": oid,
                            "ID поставки": norm_id(supply.get("supply_id") or supply.get("id")),
                        })
            if idx % 20 == 0:
                logging.info("Обработано заявок: %s/%s", idx, len(list_items))
        return pd.DataFrame(rows)

    def _estimate_supply_eta(self, state: str, route: str, cluster: str, slot_from: Optional[date],
                             arrival_api: Optional[date], detail: Mapping[str, Any], supply: Mapping[str, Any]) -> Tuple[Optional[date], str, Optional[int]]:
        if arrival_api:
            return arrival_api, "Дата прибытия из API Ozon", None
        if state == "ACCEPTANCE_AT_STORAGE_WAREHOUSE":
            return self.target_date, "Уже идёт приёмка на складе назначения", 0
        if route == "Прямая":
            if slot_from:
                return max(self.target_date, slot_from), "Слот прямой поставки", 0
            return None, "Нет даты прямой поставки в API", None

        sla_days, source = self.sla.lookup(cluster)
        if sla_days is None:
            return None, "Нет срока маршрута в Базе знаний Ozon/кэше", None

        # Для уже переданной Ozon поставки не добавляем повторно 4 дня сборки.
        # Началом маршрута считаем фактический/плановый слот отгрузки, если он известен.
        if state in PHYSICAL_STATES:
            if slot_from:
                eta = slot_from + timedelta(days=sla_days)
                return max(self.target_date, eta), f"Слот передачи Ozon + {sla_days} дн. кросс-дока ({source})", sla_days
            return self.target_date + timedelta(days=sla_days), f"Поставка уже у Ozon; оставшийся срок оценён консервативно как до {sla_days} дн. ({source})", sla_days

        # Заявка есть, но Ozon ещё не принял товар. Если есть слот — это наша реальная
        # дата передачи. Если слота нет, считаем, что на сборку нужно 4 дня.
        if slot_from:
            handoff = max(slot_from, self.target_date)
            return handoff + timedelta(days=sla_days), f"Существующий слот + {sla_days} дн. кросс-дока ({source})", sla_days
        handoff = self.target_date + timedelta(days=PREPARATION_DAYS)
        return handoff + timedelta(days=sla_days), f"{PREPARATION_DAYS} дн. на сборку + {sla_days} дн. кросс-дока ({source})", sla_days

    # ---------- сборка ----------
    def build(self) -> Dict[str, pd.DataFrame]:
        self.fetch_products()
        self.fetch_clusters()
        sla_map = self.sla.load()
        logging.info("Сроков кросс-дока из Кавказского загружено: %s", len(sla_map))
        stocks = self.fetch_stock_warehouses()
        transit = self.fetch_transit_stocks()
        sales_cluster, _ = self.sales_7d_by_cluster()
        buyout, buyout_audit = self.buyout_30d()
        supplies = self.supply_orders()
        logging.info("8/8 Формирование сводки")

        by_article = self._build_by_article(stocks, transit, sales_cluster, buyout, supplies)
        by_cluster = self._build_by_cluster(stocks, transit, sales_cluster, supplies)
        method = self._methodology()
        cluster_ref = self._cluster_reference()
        return {
            "По артикулам": by_article,
            "По кластерам": by_cluster,
            "Поставки и прогноз": supplies,
            "Расчёт выкупа": buyout_audit,
            "Справочник кластеров": cluster_ref,
            "Методика": method,
        }

    def _build_by_article(self, stocks: pd.DataFrame, transit: pd.DataFrame, sales_cluster: pd.DataFrame,
                          buyout: pd.DataFrame, supplies: pd.DataFrame) -> pd.DataFrame:
        articles = set()
        for df in (stocks, transit, sales_cluster, buyout, supplies):
            if df is not None and not df.empty and "Артикул" in df.columns:
                articles |= set(str(x) for x in df["Артикул"].dropna() if str(x).strip())
        rows: List[Dict[str, Any]] = []
        for article in sorted(articles):
            s = stocks[stocks["Артикул"] == article] if not stocks.empty else pd.DataFrame()
            t = transit[transit["Артикул"] == article] if not transit.empty else pd.DataFrame()
            sc = sales_cluster[sales_cluster["Артикул"] == article] if not sales_cluster.empty else pd.DataFrame()
            sp = supplies[supplies["Артикул"] == article] if not supplies.empty else pd.DataFrame()
            b = buyout[buyout["Артикул"] == article] if not buyout.empty else pd.DataFrame()
            current = int(s["Доступно сейчас, шт."].sum()) if not s.empty else 0
            ret = int(t["Возврат покупателя в пути, шт."].sum()) if not t.empty else 0
            internal = int(t["Перемещение внутри Ozon, шт."].sum()) if not t.empty else 0
            physical = int(sp.loc[sp["Уже передано Ozon"] == "Да", "Количество, шт."].sum()) if not sp.empty else 0
            requests = int(sp.loc[sp["Уже передано Ozon"] == "Нет", "Количество, шт."].sum()) if not sp.empty else 0
            # promised_amount не прибавляем отдельно: одна и та же активная поставка может
            # уже быть отражена в promised_amount и в supply-order; иначе был бы двойной счёт.
            future = current + ret + internal + physical + requests
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
            pct = None
            bought = refused = 0
            if not b.empty:
                bought = int(b["Выкуплено за 30 дней, шт."].sum())
                refused = int(b["Невыкуплено после доставки покупателю, шт."].sum())
                pct = round(bought / (bought + refused) * 100, 1) if bought + refused else None
            rows.append({
                "Артикул": article,
                "Среднесуточные продажи за 7 дней, шт.": avg,
                "Запас, дней: сейчас (с учётом пути и заявок)": f"{days_now if days_now is not None else '—'} ({days_future if days_future is not None else '—'})",
                "Остаток, шт.: сейчас (с учётом пути и заявок)": f"{current} ({future})",
                "Процент выкупа за 30 дней, %": pct,
                "Название товара": name,
                "Заказано за 7 дней, шт.": sold7,
                "Доступно к продаже сейчас, шт.": current,
                "Возвраты покупателей в пути на склад Ozon, шт.": ret,
                "Перемещение между складами Ozon, шт.": internal,
                "Поставки уже переданы Ozon, шт.": physical,
                "Активные заявки, ещё не переданы Ozon, шт.": requests,
                "Итого с учётом пути и заявок, шт.": future,
                "Запас сейчас, дней": days_now,
                "Запас с учётом пути и заявок, дней": days_future,
                "Выкуплено за 30 дней, шт.": bought,
                "Невыкуплено после доставки покупателю, шт.": refused,
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
            future_total = current + cluster_returns + cluster_internal + future_supply
            days_now = int(round(current / avg)) if avg > 0 else None
            days_future = int(round(future_total / avg)) if avg > 0 else None
            depletion = self.target_date + timedelta(days=math.ceil(current / avg)) if avg > 0 else None

            # Существующие поставки событиями по ETA.
            events: List[Tuple[date, int, str]] = []
            if not sp.empty:
                for _, r in sp.iterrows():
                    eta = r.get("Прогноз прибытия")
                    if isinstance(eta, pd.Timestamp):
                        eta = eta.date()
                    elif not isinstance(eta, date):
                        eta = parse_any_date(eta)
                    if eta:
                        events.append((eta, int_qty(r.get("Количество, шт.")), str(r.get("Статус") or "")))
            events.sort(key=lambda x: x[0])

            sla_days, sla_source = self.sla.lookup(cluster)
            normal_lead = PREPARATION_DAYS + sla_days if sla_days is not None else None
            hypothetical_eta = self.target_date + timedelta(days=normal_lead) if normal_lead is not None else None

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
                "Остаток с учётом пути и поставок, шт.": future_total,
                "Запас сейчас, дней": days_now,
                "Запас с учётом заявок, дней": days_future,
                "Дата ожидаемого окончания остатка": depletion if depletion else None,
                "Ближайшее пополнение, шт.": forecast_qty,
                "Статус ближайшего пополнения": forecast_status,
                "Прогноз прибытия пополнения": forecast_eta,
                "Срок кросс-дока из Москвы — Кавказский, дней": sla_days,
                "Срок пополнения новой поставкой с учётом 4 дней сборки, дней": normal_lead,
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
            rows.append({"Склад Ozon (нормализовано)": wh, "Кластер": cluster})
        for mid, cluster in sorted(self.cluster_maps.macro_to_cluster.items()):
            rows.append({"ID макролокального кластера": mid, "Кластер": cluster})
        return pd.DataFrame(rows)

    def _methodology(self) -> pd.DataFrame:
        return pd.DataFrame([
            {"Показатель": "Среднесуточные продажи за 7 дней", "Как считаем": "Заказы FBO за последние 7 календарных дней / 7. По кластеру используем cluster_to из отправления."},
            {"Показатель": "Остаток сейчас", "Как считаем": "free_to_sell_amount — реально доступно к продаже на FBO-складах Ozon."},
            {"Показатель": "Остаток в скобках по артикулу", "Как считаем": "Остаток сейчас + возвраты покупателей в пути + внутренние перемещения Ozon + ВСЕ активные заявки продавца (уже переданные и ещё не переданные Ozon)."},
            {"Показатель": "Остаток в скобках по кластеру", "Как считаем": "Остаток кластера + активные заявки/поставки, у которых удалось определить этот кластер назначения. Общесетевой транзит без кластера искусственно не распределяем."},
            {"Показатель": "promised_amount", "Как считаем": "Показывается источником Ozon, но отдельно к активным supply-order не прибавляется, чтобы одну поставку не посчитать дважды."},
            {"Показатель": "Процент выкупа", "Как считаем": "Delivered / (Delivered + явный отказ/неполучение после прибытия покупателю). Ранние отмены и ClientReturn после покупки исключены."},
            {"Показатель": "Срок новой поставки", "Как считаем": f"{PREPARATION_DAYS} дня на сборку + срок кросс-дока из официальной Базы знаний Ozon для точки «{CROSSDOCK_ORIGIN}». Берём верхнюю границу, если указан диапазон."},
            {"Показатель": "Существующая заявка", "Как считаем": "Если товар уже передан Ozon — 4 дня сборки повторно не добавляем. Если есть будущий слот — считаем от слота. Если заявка есть без слота — от даты отчёта + 4 дня."},
            {"Показатель": "Дата прибытия кросс-дока", "Как считаем": "С 16.02.2026 Ozon возвращает storage_warehouse.arrival_date=null для кросс-дока, поэтому это прогноз, а не обещанная дата API."},
            {"Показатель": "Прогноз дней без остатка", "Как считаем": "Моделируем расход по среднесуточным продажам и даты/количество известных пополнений. Если заявки нет — сравниваем запас с нормативом новой поставки."},
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
            "crossdock_sla_source": self.sla.source,
            "crossdock_sla_updated_at": self.sla.updated_at,
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
                    elif "Дата" in header or "Прогноз прибытия" in header or "Слот" in header:
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
