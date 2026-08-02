#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Ozon FBO Daily Data Collector v5

Назначение:
- ежедневный read-only сбор максимально доступных нефинансовых данных Ozon Seller API;
- отдельное хранение каждого отчёта в Yandex Object Storage (бакет ozon-assist);
- недельные Excel-файлы с обновлением/дедупликацией;
- мастер-справочники как актуальные снимки;
- test: тестовая выгрузка только целевой даты;
- daily: ежедневное обновление последних 90 календарных дней;
- archive: ZIP строго за выбранную дату;
- магазин TOPFACE, архитектура готова к добавлению других магазинов.

Правило целевой даты (как в действующем WB-сборщике):
- OZON_TARGET_DATE / --target-date задана явно -> используем её;
- запуск по Москве с 15:00 -> вчера;
- запуск до 15:00 -> позавчера.

Важно:
- финансовые методы намеренно не вызываются;
- Performance API намеренно не вызывается в v2;
- API Ozon меняется. Необязательные методы помечены optional: отсутствие права/метода
  фиксируется в диагностике и не ломает остальные отчёты.
"""
from __future__ import annotations

import argparse
import io
import json
import logging
import os
import re
import sys
import tempfile
import time
import traceback
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, Iterable, List, Mapping, Optional, Sequence, Tuple
from zoneinfo import ZoneInfo

import boto3
import pandas as pd
import requests
from botocore.client import Config
from botocore.exceptions import ClientError
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.utils import get_column_letter

SCRIPT_VERSION = "OZON_FBO_ALL_DATA_V8_NO_PRODUCT_ATTRIBUTES_20260802"
MOSCOW_TZ = ZoneInfo("Europe/Moscow")
OZON_API_BASE = "https://api-seller.ozon.ru"
DEFAULT_BUCKET = "ozon-assist"
DEFAULT_STORE = "TOPFACE"
DEFAULT_DAILY_LOOKBACK_DAYS = 90
DEFAULT_RETRY_DAYS = 7
MAX_EXCEL_CELL = 32000


# ---------------------------------------------------------------------------
# Environment / dates
# ---------------------------------------------------------------------------

def load_report_env() -> Dict[str, str]:
    """Загружает многострочный REPORT_ENV без перезаписи уже заданных env."""
    raw = os.getenv("REPORT_ENV", "") or ""
    loaded: Dict[str, str] = {}
    for line in raw.splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("export "):
            line = line[len("export "):].strip()
        if "=" not in line:
            continue
        key, value = line.split("=", 1)
        key, value = key.strip(), value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value
            loaded[key] = value
    return loaded


def parse_iso_date(value: str) -> date:
    return datetime.strptime(value.strip(), "%Y-%m-%d").date()


def resolve_target_date(forced: str = "", now_msk: Optional[datetime] = None) -> date:
    forced = (forced or os.getenv("OZON_TARGET_DATE", "") or "").strip()
    if forced:
        return parse_iso_date(forced)
    now_msk = now_msk or datetime.now(MOSCOW_TZ)
    return now_msk.date() - timedelta(days=1 if now_msk.hour >= 15 else 2)


def iso_week(d: date) -> Tuple[int, int]:
    y, w, _ = d.isocalendar()
    return int(y), int(w)


def week_filename(prefix: str, d: date) -> str:
    year, week = iso_week(d)
    return f"{prefix}_{year}-W{week:02d}.xlsx"


def to_ozon_datetime(d: date, end: bool = False) -> str:
    suffix = "23:59:59.999Z" if end else "00:00:00.000Z"
    return f"{d.isoformat()}T{suffix}"


def mode_period(mode: str, target_date: date) -> Tuple[date, date]:
    """Возвращает включительный период получения событий."""
    normalized = str(mode or "").strip().lower()
    if normalized == "daily":
        return target_date - timedelta(days=DEFAULT_DAILY_LOOKBACK_DAYS - 1), target_date
    # test и archive строго за один календарный день; daily — 90 дней включительно.
    return target_date, target_date


def iter_days(start: date, end: date) -> Iterable[date]:
    current = start
    while current <= end:
        yield current
        current += timedelta(days=1)


def iter_date_chunks(start: date, end: date, chunk_days: int = 1) -> Iterable[Tuple[date, date]]:
    """Разбивает диапазон, чтобы не упираться в MAX_OFFSET_EXCEEDED."""
    current = start
    while current <= end:
        chunk_end = min(end, current + timedelta(days=max(1, chunk_days) - 1))
        yield current, chunk_end
        current = chunk_end + timedelta(days=1)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def safe_json(value: Any) -> str:
    try:
        text = json.dumps(value, ensure_ascii=False, default=str, separators=(",", ":"))
    except Exception:
        text = str(value)
    return text[:MAX_EXCEL_CELL]


def scalarize(value: Any) -> Any:
    if isinstance(value, (dict, list, tuple)):
        return safe_json(value)
    return value


def flatten_record(record: Mapping[str, Any], prefix: str = "") -> Dict[str, Any]:
    """Плоские скаляры; вложенные списки/словари сохраняются JSON для сохранения всех данных."""
    out: Dict[str, Any] = {}
    for key, value in record.items():
        name = f"{prefix}{key}" if prefix else str(key)
        if isinstance(value, dict):
            # Делаем и плоские дочерние поля, и исходный JSON для полной сохранности.
            out[name] = safe_json(value)
            for child_key, child_val in value.items():
                child_name = f"{name}.{child_key}"
                out[child_name] = scalarize(child_val)
        else:
            out[name] = scalarize(value)
    return out


def records_to_df(records: Iterable[Mapping[str, Any]], extra: Optional[Dict[str, Any]] = None) -> pd.DataFrame:
    rows: List[Dict[str, Any]] = []
    for rec in records:
        if not isinstance(rec, Mapping):
            rows.append({"value": scalarize(rec)})
            continue
        row = flatten_record(rec)
        if extra:
            row.update(extra)
        rows.append(row)
    return pd.DataFrame(rows)


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out.columns = [str(c).strip() for c in out.columns]
    return out


def first_existing(df: pd.DataFrame, candidates: Sequence[str]) -> Optional[str]:
    lookup = {str(c).strip().lower(): c for c in df.columns}
    for candidate in candidates:
        actual = lookup.get(candidate.strip().lower())
        if actual is not None:
            return actual
    return None


def ensure_date_column(df: pd.DataFrame, column: str, value: date) -> pd.DataFrame:
    out = df.copy()
    out[column] = value.isoformat()
    return out


def dedupe_merge(old: pd.DataFrame, new: pd.DataFrame, keys: Sequence[str]) -> pd.DataFrame:
    if old is None or old.empty:
        combined = new.copy()
    elif new is None or new.empty:
        combined = old.copy()
    else:
        combined = pd.concat([old, new], ignore_index=True, sort=False)
    if combined.empty:
        return combined
    valid = [k for k in keys if k in combined.columns]
    if valid:
        combined = combined.drop_duplicates(subset=valid, keep="last")
    else:
        combined = combined.drop_duplicates(keep="last")
    return combined.reset_index(drop=True)


def extract_items(data: Any, paths: Sequence[Sequence[str]]) -> List[Dict[str, Any]]:
    """Извлекает список из нескольких известных форматов ответа Ozon."""
    for path in paths:
        cur = data
        ok = True
        for key in path:
            if not isinstance(cur, Mapping) or key not in cur:
                ok = False
                break
            cur = cur[key]
        if ok and isinstance(cur, list):
            return [x for x in cur if isinstance(x, Mapping)]
    if isinstance(data, list):
        return [x for x in data if isinstance(x, Mapping)]
    return []


def extract_cursor(data: Any) -> str:
    candidates = [
        ("result", "cursor"), ("cursor",), ("result", "last_id"), ("last_id",),
        ("result", "next_cursor"), ("next_cursor",),
    ]
    for path in candidates:
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
    for path in [("result", "total"), ("total",), ("result", "total_count"), ("total_count",)]:
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


# ---------------------------------------------------------------------------
# S3
# ---------------------------------------------------------------------------

class S3Storage:
    def __init__(self, access_key: str, secret_key: str, bucket: str, endpoint: str):
        self.bucket = bucket
        self.client = boto3.client(
            "s3",
            endpoint_url=endpoint,
            aws_access_key_id=access_key,
            aws_secret_access_key=secret_key,
            region_name="ru-central1",
            config=Config(signature_version="s3v4", read_timeout=300, connect_timeout=60,
                          retries={"max_attempts": 7, "mode": "standard"}),
        )

    def ensure_bucket(self) -> None:
        self.client.head_bucket(Bucket=self.bucket)

    def exists(self, key: str) -> bool:
        try:
            self.client.head_object(Bucket=self.bucket, Key=key)
            return True
        except ClientError as exc:
            code = str(exc.response.get("Error", {}).get("Code", ""))
            if code in {"404", "NoSuchKey", "NotFound"}:
                return False
            raise

    def read_bytes(self, key: str) -> bytes:
        return self.client.get_object(Bucket=self.bucket, Key=key)["Body"].read()

    def read_excel(self, key: str) -> pd.DataFrame:
        if not self.exists(key):
            return pd.DataFrame()
        return pd.read_excel(io.BytesIO(self.read_bytes(key)))

    def upload_bytes(self, key: str, data: bytes, content_type: Optional[str] = None) -> None:
        kwargs: Dict[str, Any] = {"Bucket": self.bucket, "Key": key, "Body": data}
        if content_type:
            kwargs["ContentType"] = content_type
        self.client.put_object(**kwargs)

    def upload_file(self, local_path: str, key: str) -> None:
        self.client.upload_file(local_path, self.bucket, key)

    def upload_json(self, key: str, value: Any) -> None:
        self.upload_bytes(key, json.dumps(value, ensure_ascii=False, indent=2, default=str).encode("utf-8"),
                          "application/json")

    def list_keys(self, prefix: str) -> List[str]:
        out: List[str] = []
        token: Optional[str] = None
        while True:
            kwargs: Dict[str, Any] = {"Bucket": self.bucket, "Prefix": prefix}
            if token:
                kwargs["ContinuationToken"] = token
            resp = self.client.list_objects_v2(**kwargs)
            out.extend([x["Key"] for x in resp.get("Contents", [])])
            if not resp.get("IsTruncated"):
                break
            token = resp.get("NextContinuationToken")
        return out


# ---------------------------------------------------------------------------
# API
# ---------------------------------------------------------------------------

class OzonApiError(RuntimeError):
    def __init__(self, method: str, path: str, status: int, message: str, response_text: str = ""):
        super().__init__(f"{method} {path}: HTTP {status}: {message}")
        self.method = method
        self.path = path
        self.status = status
        self.response_text = response_text


class OzonSellerClient:
    def __init__(self, client_id: str, api_key: str, timeout: int = 180):
        self.base = OZON_API_BASE.rstrip("/")
        self.timeout = timeout
        self.session = requests.Session()
        self.session.headers.update({
            "Client-Id": str(client_id),
            "Api-Key": str(api_key),
            "Content-Type": "application/json",
            "Accept": "application/json",
            "User-Agent": f"ozon-assist/{SCRIPT_VERSION}",
        })

    def post(self, path: str, payload: Mapping[str, Any], retries: int = 6) -> Dict[str, Any]:
        url = self.base + path
        last_error = ""
        for attempt in range(1, retries + 1):
            try:
                resp = self.session.post(url, json=payload, timeout=self.timeout)
            except requests.RequestException as exc:
                last_error = str(exc)
                if attempt == retries:
                    raise OzonApiError("POST", path, 0, last_error) from exc
                time.sleep(min(60, 2 ** attempt))
                continue

            if resp.status_code == 200:
                try:
                    return resp.json()
                except Exception as exc:
                    raise OzonApiError("POST", path, 200, "Ответ не является JSON", resp.text[:2000]) from exc

            text = resp.text[:6000]
            if resp.status_code in {429, 500, 502, 503, 504} and attempt < retries:
                wait = resp.headers.get("Retry-After")
                try:
                    seconds = max(1, int(float(wait))) if wait else min(90, 2 ** attempt)
                except Exception:
                    seconds = min(90, 2 ** attempt)
                logging.warning("Повтор %s/%s %s после HTTP %s через %s сек", attempt, retries, path,
                                resp.status_code, seconds)
                time.sleep(seconds)
                continue

            try:
                body = resp.json()
                msg = body.get("message") or body.get("error") or body.get("code") or text
            except Exception:
                msg = text
            raise OzonApiError("POST", path, resp.status_code, str(msg), text)

        raise OzonApiError("POST", path, 0, last_error or "Неизвестная ошибка")

    def cursor_pages(
        self,
        path: str,
        payload: Dict[str, Any],
        item_paths: Sequence[Sequence[str]],
        limit: int = 1000,
        max_pages: int = 500,
        cursor_field: str = "cursor",
    ) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        all_items: List[Dict[str, Any]] = []
        raw_pages: List[Dict[str, Any]] = []
        cursor = str(payload.get(cursor_field, "") or "")
        for page in range(1, max_pages + 1):
            body = dict(payload)
            body["limit"] = min(limit, int(body.get("limit", limit)))
            body[cursor_field] = cursor
            data = self.post(path, body)
            raw_pages.append(data)
            items = extract_items(data, item_paths)
            all_items.extend(items)
            next_cursor = extract_cursor(data)
            if not items or not next_cursor or next_cursor == cursor:
                break
            cursor = next_cursor
            time.sleep(0.12)
        return all_items, raw_pages

    def offset_pages(
        self,
        path: str,
        payload: Dict[str, Any],
        item_paths: Sequence[Sequence[str]],
        limit: int = 1000,
        max_pages: int = 500,
    ) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        all_items: List[Dict[str, Any]] = []
        raw_pages: List[Dict[str, Any]] = []
        offset = int(payload.get("offset", 0) or 0)
        for page in range(1, max_pages + 1):
            body = dict(payload)
            body["limit"] = min(limit, int(body.get("limit", limit)))
            body["offset"] = offset
            data = self.post(path, body)
            raw_pages.append(data)
            items = extract_items(data, item_paths)
            all_items.extend(items)
            total = extract_total(data)
            if not items or len(items) < body["limit"] or (total is not None and len(all_items) >= total):
                break
            offset += len(items)
            time.sleep(0.12)
        return all_items, raw_pages


# ---------------------------------------------------------------------------
# Report definitions
# ---------------------------------------------------------------------------

@dataclass
class ReportResult:
    code: str
    title: str
    folder: str
    filename_prefix: str
    df: pd.DataFrame = field(default_factory=pd.DataFrame)
    raw: Any = None
    date_column: str = "Дата"
    keys: Sequence[str] = field(default_factory=list)
    snapshot: bool = False
    optional: bool = False
    method_paths: List[str] = field(default_factory=list)
    status: str = "OK"
    message: str = ""
    s3_key: str = ""
    local_path: str = ""


# ---------------------------------------------------------------------------
# Collector
# ---------------------------------------------------------------------------

class OzonFboCollector:
    def __init__(self, client: OzonSellerClient, storage: S3Storage, store: str, target_date: date,
                 mode: str, workdir: Path):
        self.client = client
        self.storage = storage
        self.store = store.upper().strip()
        self.target_date = target_date
        self.mode = str(mode or "test").strip().lower()
        self.period_from, self.period_to = mode_period(self.mode, target_date)
        self.workdir = workdir
        self.workdir.mkdir(parents=True, exist_ok=True)
        self.run_id = datetime.now(MOSCOW_TZ).strftime("%Y%m%d_%H%M%S")
        self.results: List[ReportResult] = []
        self.errors: List[Dict[str, Any]] = []
        self.catalog_items: List[Dict[str, Any]] = []
        self.catalog_ids: List[int] = []
        self.offer_ids: List[str] = []
        self.sku_ids: List[int] = []

    @property
    def base_prefix(self) -> str:
        return "Отчёты"

    def log_error(self, code: str, path: str, exc: Exception, optional: bool) -> None:
        row = {
            "run_id": self.run_id,
            "report": code,
            "path": path,
            "optional": optional,
            "error_type": type(exc).__name__,
            "error": str(exc),
            "traceback": traceback.format_exc(limit=8),
        }
        if isinstance(exc, OzonApiError):
            row["http_status"] = exc.status
            row["response_text"] = exc.response_text
        self.errors.append(row)
        level = logging.WARNING if optional else logging.ERROR
        logging.log(level, "%s: %s", code, exc)

    def _run(self, code: str, title: str, folder: str, filename_prefix: str, fn,
             date_column: str, keys: Sequence[str], snapshot: bool = False,
             optional: bool = False) -> ReportResult:
        logging.info("=== %s ===", title)
        result = ReportResult(code=code, title=title, folder=folder,
                              filename_prefix=filename_prefix, date_column=date_column,
                              keys=list(keys), snapshot=snapshot, optional=optional)
        try:
            df, raw, methods = fn()
            result.df = normalize_columns(df if df is not None else pd.DataFrame())
            result.raw = raw
            result.method_paths = list(methods)
            if result.df.empty:
                result.status = "EMPTY"
                result.message = "Метод доступен, но строки отсутствуют"
            logging.info("%s: %s строк", title, len(result.df))
        except Exception as exc:
            self.log_error(code, title, exc, optional)
            result.status = "SKIPPED_OPTIONAL" if optional else "ERROR"
            result.message = str(exc)
        self.results.append(result)
        return result

    # ------------------------- products -------------------------
    def fetch_product_catalog(self) -> Tuple[pd.DataFrame, Any, List[str]]:
        items, raw = self.client.cursor_pages(
            "/v3/product/list",
            {"filter": {"visibility": "ALL"}, "limit": 1000, "last_id": ""},
            [("result", "items"), ("items",)],
            cursor_field="last_id",
        )
        # cursor_pages expects generic cursor extraction and body field last_id works here.
        self.catalog_items = items
        ids: List[int] = []
        offers: List[str] = []
        for item in items:
            try:
                if item.get("product_id") is not None:
                    ids.append(int(item["product_id"]))
            except Exception:
                pass
            if item.get("offer_id") not in (None, ""):
                offers.append(str(item["offer_id"]))
        self.catalog_ids = sorted(set(ids))
        self.offer_ids = sorted(set(offers))
        df = records_to_df(items, {"Дата снимка": self.target_date.isoformat()})
        return df, raw, ["/v3/product/list"]

    def _batch(self, values: Sequence[Any], size: int) -> Iterable[List[Any]]:
        for i in range(0, len(values), size):
            yield list(values[i:i + size])

    def fetch_product_info(self) -> Tuple[pd.DataFrame, Any, List[str]]:
        if not self.offer_ids and not self.catalog_ids:
            return pd.DataFrame(), [], ["/v3/product/info/list"]
        rows: List[Dict[str, Any]] = []
        raw: List[Any] = []
        # v3 accepts offer_id/product_id lists. Use offer_id to keep seller identity stable.
        source = self.offer_ids or self.catalog_ids
        for batch in self._batch(source, 100):
            payload = {"offer_id": batch, "product_id": [], "sku": []} if self.offer_ids else {
                "offer_id": [], "product_id": batch, "sku": []}
            data = self.client.post("/v3/product/info/list", payload)
            raw.append(data)
            batch_items = extract_items(data, [("items",), ("result", "items")])
            rows.extend(batch_items)
            for item in batch_items:
                for candidate in ("sku", "fbo_sku", "fbs_sku"):
                    try:
                        value = item.get(candidate)
                        if value not in (None, "", 0):
                            self.sku_ids.append(int(value))
                    except Exception:
                        pass
            time.sleep(0.12)
        self.sku_ids = sorted(set(self.sku_ids))
        return records_to_df(rows, {"Дата снимка": self.target_date.isoformat()}), raw, ["/v3/product/info/list"]

    def fetch_prices(self) -> Tuple[pd.DataFrame, Any, List[str]]:
        items, raw = self.client.cursor_pages(
            "/v5/product/info/prices",
            {"filter": {"visibility": "ALL"}, "limit": 1000, "cursor": ""},
            [("items",), ("result", "items")],
        )
        df = records_to_df(items, {"Дата снимка": self.target_date.isoformat()})
        # Расчётный аналог поддержки скидки, только когда поля доступны.
        seller_col = first_existing(df, ["price.marketing_seller_price", "marketing_seller_price", "price"])
        buyer_col = first_existing(df, ["price.marketing_price", "marketing_price", "customer_price"])
        if seller_col and buyer_col:
            seller = pd.to_numeric(df[seller_col], errors="coerce")
            buyer = pd.to_numeric(df[buyer_col], errors="coerce")
            df["Расчётная поддержка скидки Ozon, руб"] = (seller - buyer).clip(lower=0)
            df["Расчётная поддержка скидки Ozon, %"] = ((seller - buyer) / seller.replace(0, pd.NA) * 100).clip(lower=0)
        return df, raw, ["/v5/product/info/prices"]

    # ------------------------- stocks -------------------------
    def fetch_stocks(self) -> Tuple[pd.DataFrame, Any, List[str]]:
        items, raw = self.client.cursor_pages(
            "/v4/product/info/stocks",
            {"filter": {"visibility": "ALL"}, "limit": 1000, "cursor": ""},
            [("items",), ("result", "items")],
        )
        rows: List[Dict[str, Any]] = []
        for item in items:
            base = {k: v for k, v in item.items() if k != "stocks"}
            stocks = item.get("stocks")
            if isinstance(stocks, list) and stocks:
                for stock in stocks:
                    row = dict(base)
                    if isinstance(stock, Mapping):
                        row.update({f"stock.{k}": v for k, v in stock.items()})
                    else:
                        row["stock"] = stock
                    rows.append(row)
            else:
                rows.append(item)
        return records_to_df(rows, {"Дата снимка": self.target_date.isoformat()}), raw, ["/v4/product/info/stocks"]

    def fetch_stock_on_warehouses(self) -> Tuple[pd.DataFrame, Any, List[str]]:
        # Расширенный складской разрез. Не во всех кабинетах/версиях доступен.
        items, raw = self.client.offset_pages(
            "/v2/analytics/stock_on_warehouses",
            {"limit": 1000, "offset": 0, "warehouse_type": "ALL"},
            [("result", "rows"), ("rows",), ("result", "items")],
        )
        return records_to_df(items, {"Дата снимка": self.target_date.isoformat()}), raw, ["/v2/analytics/stock_on_warehouses"]

    def fetch_turnover(self) -> Tuple[pd.DataFrame, Any, List[str]]:
        data = self.client.post("/v1/analytics/turnover/stocks", {"limit": 1000, "offset": 0})
        items = extract_items(data, [("items",), ("result", "items"), ("result", "rows"), ("rows",)])
        # Если метод offset-пагинируемый, добираем страницы.
        if len(items) >= 1000:
            items, raw = self.client.offset_pages(
                "/v1/analytics/turnover/stocks", {"limit": 1000, "offset": 0},
                [("items",), ("result", "items"), ("result", "rows"), ("rows",)]
            )
        else:
            raw = [data]
        return records_to_df(items, {"Дата снимка": self.target_date.isoformat()}), raw, ["/v1/analytics/turnover/stocks"]

    # ------------------------- orders/returns -------------------------
    def fetch_fbo_postings(self) -> Tuple[pd.DataFrame, Any, List[str]]:
        """FBO-заказы. Диапазон режется по дням, чтобы не получить MAX_OFFSET_EXCEEDED."""
        all_postings: List[Dict[str, Any]] = []
        raw_all: List[Any] = []
        for day_from, day_to in iter_date_chunks(self.period_from, self.period_to, chunk_days=1):
            items, raw = self.client.offset_pages(
                "/v2/posting/fbo/list",
                {
                    "dir": "ASC",
                    "filter": {
                        "since": to_ozon_datetime(day_from),
                        "to": to_ozon_datetime(day_to, end=True),
                        "status": "",
                    },
                    "limit": 1000,
                    "offset": 0,
                    "translit": True,
                    "with": {"analytics_data": True, "financial_data": False},
                },
                [("result",), ("result", "postings"), ("postings",)],
                max_pages=100,
            )
            raw_all.extend(raw)
            all_postings.extend(items)
            logging.info("Заказы FBO: получен день %s, отправлений %s", day_from, len(items))

        rows: List[Dict[str, Any]] = []
        for posting in all_postings:
            base = {k: v for k, v in posting.items() if k != "products"}
            products = posting.get("products")
            if isinstance(products, list) and products:
                for product in products:
                    row = dict(base)
                    if isinstance(product, Mapping):
                        row.update({f"product.{k}": v for k, v in product.items()})
                    else:
                        row["product"] = product
                    rows.append(row)
            else:
                rows.append(posting)
        df = records_to_df(rows, {
            "Дата обновления снимка": self.target_date.isoformat(),
            "Период с": self.period_from.isoformat(),
            "Период по": self.period_to.isoformat(),
        })
        return df, raw_all, ["/v2/posting/fbo/list"]

    def fetch_returns(self) -> Tuple[pd.DataFrame, Any, List[str]]:
        since = self.period_from
        payload = {
            "filter": {
                "logistic_return_date": {"time_from": to_ozon_datetime(since),
                                         "time_to": to_ozon_datetime(self.period_to, end=True)},
            },
            "limit": 500,
            "last_id": 0,
        }
        # Метод имеет разные форматы пагинации между версиями. Сначала один вызов,
        # затем при наличии last_id продолжаем вручную.
        rows: List[Dict[str, Any]] = []
        raw: List[Any] = []
        last_id: Any = 0
        for _ in range(500):
            body = dict(payload)
            body["last_id"] = last_id
            data = self.client.post("/v1/returns/list", body)
            raw.append(data)
            items = extract_items(data, [("returns",), ("result", "returns"), ("result", "items"), ("items",)])
            rows.extend(items)
            new_last = None
            for path in [("last_id",), ("result", "last_id")]:
                cur = data
                ok = True
                for key in path:
                    if not isinstance(cur, Mapping) or key not in cur:
                        ok = False
                        break
                    cur = cur[key]
                if ok:
                    new_last = cur
                    break
            if not items or new_last in (None, "", 0, last_id):
                break
            last_id = new_last
            time.sleep(0.12)
        return records_to_df(rows, {"Дата обновления снимка": self.target_date.isoformat()}), raw, ["/v1/returns/list"]

    # ------------------------- analytics funnel -------------------------
    def fetch_funnel(self) -> Tuple[pd.DataFrame, Any, List[str]]:
        dimensions = ["day", "sku"]
        metrics = [
            "hits_view_search", "hits_view_pdp", "hits_tocart", "ordered_units", "revenue",
            "cancellations", "delivered_units", "returns", "session_view", "conv_tocart",
            "conv_tocart_from_search", "conv_order", "position_category",
        ]
        payload = {
            "date_from": self.period_from.isoformat(),
            "date_to": self.period_to.isoformat(),
            "dimension": dimensions,
            "metrics": metrics,
            "filters": [],
            "sort": [{"key": "ordered_units", "order": "DESC"}],
            "limit": 1000,
            "offset": 0,
        }
        try:
            items, raw = self.client.offset_pages(
                "/v1/analytics/data", payload,
                [("result", "data"), ("data",), ("result", "rows"), ("rows",)],
            )
        except OzonApiError as exc:
            # Метрики могут меняться. Повтор с консервативным ядром.
            if exc.status not in {400, 422}:
                raise
            logging.warning("Расширенная воронка отклонена, повторяем с базовыми метриками")
            payload["metrics"] = ["hits_view_search", "hits_view_pdp", "hits_tocart", "ordered_units", "revenue"]
            items, raw = self.client.offset_pages(
                "/v1/analytics/data", payload,
                [("result", "data"), ("data",), ("result", "rows"), ("rows",)],
            )
        rows: List[Dict[str, Any]] = []
        for item in items:
            row: Dict[str, Any] = {}
            dims = item.get("dimensions") or item.get("dimension")
            mets = item.get("metrics")
            if isinstance(dims, list):
                for i, val in enumerate(dims):
                    key = dimensions[i] if i < len(dimensions) else f"dimension_{i}"
                    if isinstance(val, Mapping):
                        row[key] = val.get("name") or val.get("id") or safe_json(val)
                        row[f"{key}_raw"] = safe_json(val)
                    else:
                        row[key] = val
            if isinstance(mets, list):
                for i, val in enumerate(mets):
                    key = payload["metrics"][i] if i < len(payload["metrics"]) else f"metric_{i}"
                    row[key] = scalarize(val)
            for k, v in item.items():
                if k not in {"dimensions", "dimension", "metrics"}:
                    row[k] = scalarize(v)
            rows.append(row)
        df = records_to_df(rows, {
            "Период с": self.period_from.isoformat(),
            "Период по": self.period_to.isoformat(),
        })
        if "day" in df.columns and "Дата" not in df.columns:
            df["Дата"] = df["day"].astype(str).str[:10]
        elif "Дата" not in df.columns:
            df["Дата"] = self.target_date.isoformat()
        return df, raw, ["/v1/analytics/data"]

    # ------------------------- search -------------------------
    def fetch_product_queries(self) -> Tuple[pd.DataFrame, Any, List[str]]:
        """Сводная поисковая аналитика с page_size до 1000."""
        endpoint = "/v1/analytics/product-queries"
        all_items: List[Dict[str, Any]] = []
        raw: List[Any] = []
        page = 1
        page_size = 1000
        for _ in range(500):
            payload = {
                "date_from": to_ozon_datetime(self.period_from),
                "date_to": to_ozon_datetime(self.period_to, end=True),
                "page": page,
                "page_size": page_size,
                "sort": {"key": "orders", "order": "DESC"},
            }
            data = self.client.post(endpoint, payload)
            raw.append(data)
            items = extract_items(
                data,
                [("items",), ("result", "items"), ("result", "rows"), ("rows",)],
            )
            all_items.extend(items)
            if not items or len(items) < page_size:
                break
            page += 1
            time.sleep(0.12)

        return records_to_df(all_items, {
            "Дата": self.target_date.isoformat(),
            "Период с": self.period_from.isoformat(),
            "Период по": self.period_to.isoformat(),
        }), raw, [endpoint]

    def fetch_product_query_details(self) -> Tuple[pd.DataFrame, Any, List[str]]:
        """Товар × запрос. Максимальный page_size метода — 100."""
        rows: List[Dict[str, Any]] = []
        raw: List[Any] = []

        source: List[Tuple[str, Any]] = [("sku", x) for x in self.sku_ids]
        if not source:
            source = [("offer_id", x) for x in self.offer_ids]
        if not source:
            return pd.DataFrame(), raw, ["/v1/analytics/product-queries/details"]

        endpoint = "/v1/analytics/product-queries/details"
        for idx, (id_field, id_value) in enumerate(source, start=1):
            page = 1
            try:
                while page <= 100:
                    payload = {
                        "date_from": to_ozon_datetime(self.period_from),
                        "date_to": to_ozon_datetime(self.period_to, end=True),
                        id_field: id_value,
                        "page": page,
                        "page_size": 100,
                    }
                    data = self.client.post(endpoint, payload)
                    raw.append(data)
                    items = extract_items(
                        data,
                        [("items",), ("result", "items"), ("result", "rows"), ("rows",)],
                    )
                    for item in items:
                        row = dict(item)
                        row.setdefault(id_field, id_value)
                        rows.append(row)
                    if not items or len(items) < 100:
                        break
                    page += 1
                    time.sleep(0.10)
            except OzonApiError as exc:
                if exc.status in {400, 404, 422}:
                    self.log_error(
                        "search_query_details_item",
                        f"{id_field}={id_value}",
                        exc,
                        True,
                    )
                    if idx == 1:
                        logging.warning(
                            "Детализация поисковых запросов недоступна с текущей схемой; цикл остановлен"
                        )
                        break
                    continue
                raise

            if idx % 50 == 0:
                logging.info(
                    "Поисковые запросы: обработано товаров %s/%s",
                    idx,
                    len(source),
                )
            time.sleep(0.15)

        return records_to_df(rows, {
            "Дата": self.target_date.isoformat(),
            "Период с": self.period_from.isoformat(),
            "Период по": self.period_to.isoformat(),
        }), raw, [endpoint]

    def fetch_search_queries_top(self) -> Tuple[pd.DataFrame, Any, List[str]]:
        endpoint = "/v1/search-queries/top"
        all_items: List[Dict[str, Any]] = []
        raw: List[Any] = []
        offset = 0
        limit = 50
        for _ in range(500):
            payload = {
                "date_from": to_ozon_datetime(self.period_from),
                "date_to": to_ozon_datetime(self.period_to, end=True),
                "limit": limit,
                "offset": offset,
            }
            data = self.client.post(endpoint, payload)
            raw.append(data)
            items = extract_items(
                data,
                [("items",), ("result", "items"), ("result", "rows"), ("rows",)],
            )
            all_items.extend(items)
            if not items or len(items) < limit:
                break
            offset += len(items)
            time.sleep(0.12)

        return records_to_df(all_items, {
            "Дата": self.target_date.isoformat(),
            "Период с": self.period_from.isoformat(),
            "Период по": self.period_to.isoformat(),
        }), raw, [endpoint]

    # ------------------------- supplies/warehouses -------------------------
    def fetch_supply_orders(self) -> Tuple[pd.DataFrame, Any, List[str]]:
        payload = {
            "filter": {
                "states": [],
                "created_date_from": self.period_from.isoformat(),
                "created_date_to": (self.period_to + timedelta(days=1)).isoformat(),
            },
            "limit": 100,
            "last_id": "",
            "sort_by": "CREATED_AT",
            "sort_direction": "DESC",
        }
        # list
        items, raw_list = self.client.cursor_pages(
            "/v3/supply-order/list", payload,
            [("result", "items"), ("items",)], limit=100, cursor_field="last_id",
        )
        # get details, сохраняем по одной строке на товар поставки.
        rows: List[Dict[str, Any]] = []
        raw_details: List[Any] = []
        for i, item in enumerate(items, start=1):
            order_id = item.get("order_id") or item.get("supply_order_id") or item.get("id")
            base = dict(item)
            if not order_id:
                rows.append(base)
                continue
            try:
                detail = self.client.post("/v3/supply-order/get", {"order_id": order_id})
                raw_details.append(detail)
                root = detail.get("result", detail) if isinstance(detail, Mapping) else {}
                products = []
                if isinstance(root, Mapping):
                    products = root.get("items") or root.get("products") or root.get("supply_order_items") or []
                if isinstance(products, list) and products:
                    for product in products:
                        row = dict(base)
                        row["detail"] = safe_json(root)
                        if isinstance(product, Mapping):
                            row.update({f"product.{k}": v for k, v in product.items()})
                        rows.append(row)
                else:
                    row = dict(base)
                    row["detail"] = safe_json(root)
                    rows.append(row)
            except OzonApiError as exc:
                self.log_error("supply_get", str(order_id), exc, True)
                rows.append(base)
            if i % 50 == 0:
                logging.info("Поставки: детализация %s/%s", i, len(items))
            time.sleep(0.15)
        return records_to_df(rows, {"Дата обновления снимка": self.target_date.isoformat()}), {
            "list": raw_list, "details": raw_details
        }, ["/v3/supply-order/list", "/v3/supply-order/get"]

    def fetch_warehouses(self) -> Tuple[pd.DataFrame, Any, List[str]]:
        """Старый /v1/warehouse/list отключён. Справочник строим из отчёта остатков по складам."""
        source_result = next(
            (r for r in self.results if r.code == "stocks_warehouses" and not r.df.empty),
            None,
        )
        if source_result is None:
            return pd.DataFrame(), {"source": "stocks_warehouses", "rows": 0}, [
                "/v2/analytics/stock_on_warehouses"
            ]

        df = source_result.df.copy()
        candidates = [
            "warehouse_id", "warehouse_name", "cluster_name",
            "warehouse", "cluster", "region",
        ]
        cols = [c for c in candidates if c in df.columns]
        if not cols:
            return pd.DataFrame(), {"source": "stocks_warehouses", "rows": len(df)}, [
                "/v2/analytics/stock_on_warehouses"
            ]

        result = df[cols].drop_duplicates().reset_index(drop=True)
        result["Дата снимка"] = self.target_date.isoformat()
        return result, {
            "source": "stocks_warehouses",
            "rows": len(df),
        }, ["/v2/analytics/stock_on_warehouses"]

    # ------------------------- collect/save -------------------------
    def collect_all(self) -> None:
        # Порядок важен: товарный справочник нужен для детализации поисковых запросов.
        self._run("products", "Справочник товаров", "Товары", "Товары",
                  self.fetch_product_catalog, "Дата снимка",
                  ["Дата снимка", "product_id", "offer_id"], snapshot=True)
        self._run("product_info", "Подробная информация о товарах", "Информация о товарах", "Информация_о_товарах",
                  self.fetch_product_info, "Дата снимка",
                  ["Дата снимка", "id", "product_id", "offer_id"], snapshot=True, optional=True)
        self._run("prices", "Цены", "Цены", "Цены", self.fetch_prices, "Дата снимка",
                  ["Дата снимка", "product_id", "offer_id"])
        self._run("stocks", "Остатки FBO", "Остатки", "Остатки",
                  self.fetch_stocks, "Дата снимка",
                  ["Дата снимка", "product_id", "offer_id", "stock.warehouse_id", "stock.type"])
        self._run("stocks_warehouses", "Остатки по складам", "Остатки по складам", "Остатки_по_складам",
                  self.fetch_stock_on_warehouses, "Дата снимка",
                  ["Дата снимка", "sku", "warehouse_name", "warehouse_id"], optional=True)
        self._run("turnover", "Оборачиваемость Ozon", "Оборачиваемость Ozon", "Оборачиваемость_Ozon",
                  self.fetch_turnover, "Дата снимка",
                  ["Дата снимка", "sku", "offer_id", "product_id"], optional=True)
        self._run("orders", "Заказы FBO", "Заказы", "Заказы",
                  self.fetch_fbo_postings, "Дата обновления снимка",
                  ["posting_number", "product.sku", "product.offer_id", "product.name"])
        self._run("returns", "Возвраты", "Возвраты", "Возвраты",
                  self.fetch_returns, "Дата обновления снимка",
                  ["id", "return_id", "posting_number", "sku"], optional=True)
        self._run("funnel", "Воронка продаж", "Воронка продаж", "Воронка_продаж",
                  self.fetch_funnel, "Дата", ["Дата", "sku"])
        self._run("product_queries", "Поисковые запросы — сводная", "Поисковые запросы", "Поисковые_запросы_сводная",
                  self.fetch_product_queries, "Дата", ["Дата", "offer_id", "sku", "query"], optional=True)
        self._run("product_query_details", "Поисковые запросы — товар × запрос", "Поисковые запросы по товарам",
                  "Поисковые_запросы_по_товарам", self.fetch_product_query_details, "Дата",
                  ["Дата", "offer_id", "sku", "query", "search_query"], optional=True)
        self._run("search_top", "Популярные поисковые запросы", "Поисковая аналитика", "Популярные_запросы",
                  self.fetch_search_queries_top, "Дата", ["Дата", "query", "search_query"], optional=True)
        self._run("supplies", "Поставки FBO", "Поставки", "Поставки",
                  self.fetch_supply_orders, "Дата обновления снимка",
                  ["order_id", "supply_order_id", "product.sku", "product.offer_id"], optional=True)
        self._run("warehouses", "Справочник складов", "Склады", "Склады",
                  self.fetch_warehouses, "Дата снимка", ["warehouse_id", "name"], snapshot=True, optional=True)

    @staticmethod
    def _clean_excel_value(value: Any) -> Any:
        """Удаляет управляющие символы XML, запрещённые внутри XLSX."""
        if value is None:
            return value
        if isinstance(value, str):
            return ILLEGAL_CHARACTERS_RE.sub("", value)
        return value

    def _sanitize_dataframe_for_excel(self, df: pd.DataFrame) -> pd.DataFrame:
        clean = df.copy()
        clean.columns = [
            ILLEGAL_CHARACTERS_RE.sub("", str(col)) for col in clean.columns
        ]
        object_cols = clean.select_dtypes(include=["object", "string"]).columns
        for col in object_cols:
            clean[col] = clean[col].map(self._clean_excel_value)
        return clean

    def _excel_bytes(self, df: pd.DataFrame, sheet_name: str = "Данные") -> bytes:
        buffer = io.BytesIO()
        df = self._sanitize_dataframe_for_excel(df)
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            safe_sheet = re.sub(r"[\\/*?:\[\]]", "_", sheet_name)[:31] or "Данные"
            safe_sheet = ILLEGAL_CHARACTERS_RE.sub("", safe_sheet)
            df.to_excel(writer, sheet_name=safe_sheet, index=False)
            ws = writer.book[safe_sheet]
            ws.freeze_panes = "A2"
            header_fill = PatternFill("solid", fgColor="1F4E78")
            header_font = Font(color="FFFFFF", bold=True)
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            for col_idx, col in enumerate(df.columns, start=1):
                sample = [len(str(col))] + [len(str(x)) for x in df[col].dropna().astype(str).head(150)]
                ws.column_dimensions[get_column_letter(col_idx)].width = min(60, max(10, max(sample, default=10) + 2))
        return buffer.getvalue()

    def _detect_row_date_column(self, df: pd.DataFrame, preferred: str) -> Optional[str]:
        candidates = [
            preferred, "Дата", "day", "created_at", "in_process_at", "shipment_date",
            "logistic_return_date", "return_date", "posting_date", "Дата заказа",
            "Дата создания", "Дата события",
        ]
        for col in candidates:
            if col and col in df.columns:
                parsed = pd.to_datetime(df[col], errors="coerce", utc=True)
                if parsed.notna().any():
                    return col
        return None

    def _save_weekly_partitioned(self, result: ReportResult, df: pd.DataFrame) -> None:
        """Разносит 90-дневные события по ISO-неделям; агрегаты без даты идут в неделю target."""
        date_col = self._detect_row_date_column(df, result.date_column)
        if not date_col:
            partitions = [(self.target_date, df)]
        else:
            work = df.copy()
            parsed = pd.to_datetime(work[date_col], errors="coerce", utc=True)
            work["_partition_date"] = parsed.dt.date
            valid = work[work["_partition_date"].notna()].copy()
            invalid = work[work["_partition_date"].isna()].drop(columns=["_partition_date"], errors="ignore")
            partitions = []
            if not valid.empty:
                valid["_week_key"] = valid["_partition_date"].map(lambda d: iso_week(d))
                for _, part in valid.groupby("_week_key", dropna=False):
                    part_date = part["_partition_date"].iloc[0]
                    partitions.append((part_date, part.drop(columns=["_partition_date", "_week_key"], errors="ignore")))
            if not invalid.empty:
                partitions.append((self.target_date, invalid))

        written_keys = []
        for partition_date, part_df in partitions:
            filename = week_filename(result.filename_prefix, partition_date)
            key = f"{self.base_prefix}/{result.folder}/{self.store}/Недельные/{filename}"
            old = self.storage.read_excel(key)
            merged = dedupe_merge(old, part_df, result.keys)
            self.storage.upload_bytes(
                key, self._excel_bytes(merged, result.title),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            written_keys.append(key)
            logging.info("Сохранено: s3://%s/%s (%s строк)", self.storage.bucket, key, len(merged))
        result.s3_key = "; ".join(written_keys)

    def save_results(self) -> None:
        archive_local_files: List[Tuple[str, str]] = []
        for result in self.results:
            if result.status in {"ERROR", "SKIPPED_OPTIONAL"}:
                if self.mode != "archive":
                    continue
                df = pd.DataFrame([{
                    "Статус": result.status,
                    "Отчёт": result.title,
                    "Дата": self.target_date.isoformat(),
                    "Период с": self.period_from.isoformat(),
                    "Период по": self.period_to.isoformat(),
                    "Ошибка": result.message,
                    "Методы": ", ".join(result.method_paths),
                }])
            else:
                df = result.df.copy()
            if df.empty:
                # В archive режиме сохраняем и пустой отчёт, чтобы было видно, что метод отработал.
                if self.mode != "archive":
                    continue
                df = pd.DataFrame([{
                    "Статус": "Нет строк",
                    "Дата": self.target_date.isoformat(),
                    "Методы": ", ".join(result.method_paths),
                }])

            if self.mode == "archive":
                local_name = f"{result.filename_prefix}_{self.target_date.isoformat()}.xlsx"
                local_path = self.workdir / local_name
                local_path.write_bytes(self._excel_bytes(df, result.title))
                result.local_path = str(local_path)
                archive_local_files.append((str(local_path), f"{result.folder}/{local_name}"))
                continue

            if result.snapshot:
                key = f"{self.base_prefix}/{result.folder}/{self.store}/{result.filename_prefix}.xlsx"
                merged = df
            else:
                self._save_weekly_partitioned(result, df)
                # Сырые ответы сохраняются ниже.
                merged = None
                key = result.s3_key

            if result.snapshot:
                self.storage.upload_bytes(key, self._excel_bytes(merged, result.title),
                                          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                result.s3_key = key
                logging.info("Сохранено: s3://%s/%s (%s строк)", self.storage.bucket, key, len(merged))

            # Сырые ответы — отдельная служебная папка, не смешиваем с отчётами.
            if result.raw is not None:
                raw_key = (f"{self.base_prefix}/Служебные/{self.store}/Сырые ответы/{self.target_date.isoformat()}/"
                           f"{result.code}_{self.run_id}.json")
                self.storage.upload_json(raw_key, result.raw)

        if self.mode == "archive":
            archive_dir = self.workdir / f"Архив_{self.store}_{self.target_date.isoformat()}"
            archive_dir.mkdir(exist_ok=True)
            zip_name = f"Архив_всех_отчётов_{self.store}_{self.target_date.isoformat()}_{self.run_id}.zip"
            zip_path = self.workdir / zip_name
            manifest = self.build_manifest()
            manifest_path = self.workdir / "manifest.json"
            manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
            with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as zf:
                for local_path, archive_name in archive_local_files:
                    zf.write(local_path, archive_name)
                zf.write(manifest_path, "Служебные/manifest.json")
                if self.errors:
                    err_path = self.workdir / "errors.json"
                    err_path.write_text(json.dumps(self.errors, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
                    zf.write(err_path, "Служебные/errors.json")
            archive_key = (f"{self.base_prefix}/Архив/{self.store}/{self.target_date:%Y/%m}/"
                           f"{zip_name}")
            self.storage.upload_file(str(zip_path), archive_key)
            logging.info("Архив сохранён: s3://%s/%s", self.storage.bucket, archive_key)
            # Удобный указатель на последний архив выбранного дня.
            self.storage.upload_json(
                f"{self.base_prefix}/Архив/{self.store}/Последний_архив.json",
                {"date": self.target_date.isoformat(), "key": archive_key, "run_id": self.run_id,
                 "created_at": datetime.now(MOSCOW_TZ).isoformat()},
            )

    def build_manifest(self) -> Dict[str, Any]:
        return {
            "script_version": SCRIPT_VERSION,
            "run_id": self.run_id,
            "store": self.store,
            "mode": self.mode,
            "target_date": self.target_date.isoformat(),
            "period_from": self.period_from.isoformat(),
            "period_to": self.period_to.isoformat(),
            "started_at": datetime.now(MOSCOW_TZ).isoformat(),
            "reports": [
                {
                    "code": r.code,
                    "title": r.title,
                    "status": r.status,
                    "rows": len(r.df),
                    "methods": r.method_paths,
                    "s3_key": r.s3_key,
                    "message": r.message,
                } for r in self.results
            ],
            "errors_count": len(self.errors),
        }

    def save_diagnostics(self) -> None:
        manifest = self.build_manifest()
        base = f"{self.base_prefix}/Служебные/{self.store}"
        self.storage.upload_json(f"{base}/Запуски/{self.target_date.isoformat()}_{self.run_id}.json", manifest)
        self.storage.upload_json(f"{base}/Последний_запуск.json", manifest)
        if self.errors:
            self.storage.upload_json(f"{base}/Ошибки_API/{self.target_date.isoformat()}_{self.run_id}.json", self.errors)
        required_errors = [r for r in self.results if r.status == "ERROR" and not r.optional]
        if required_errors:
            names = ", ".join(r.title for r in required_errors)
            if self.mode == "daily":
                raise RuntimeError(
                    "Есть ошибки обязательных отчётов в daily: "
                    + names
                    + ". См. Служебные/Ошибки_API"
                )
            logging.error(
                "В режиме %s есть ошибки обязательных отчётов: %s. "
                "Архив/тест сохранён с diagnostics и workflow не прерывается.",
                self.mode,
                names,
            )


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def env_first(*names: str, default: str = "") -> str:
    for name in names:
        val = os.getenv(name)
        if val not in (None, ""):
            return str(val)
    return default


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Ozon FBO: ежедневный сбор всех нефинансовых данных")
    p.add_argument("--mode", choices=["test", "daily", "archive"], default=env_first("OZON_MODE", default="test"))
    p.add_argument("--target-date", default=env_first("OZON_TARGET_DATE"))
    p.add_argument("--store", default=env_first("OZON_STORE", "STORE", default=DEFAULT_STORE))
    p.add_argument("--bucket", default=env_first("OZON_YC_BUCKET", "YC_BUCKET_NAME", default=DEFAULT_BUCKET))
    p.add_argument("--workdir", default="output_ozon_v1")
    return p


def main() -> int:
    load_report_env()
    args = build_parser().parse_args()
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    logging.info("VERSION: %s", SCRIPT_VERSION)

    store = str(args.store).upper().strip() or DEFAULT_STORE
    client_id = env_first(f"OZON_CLIENT_ID_{store}", "OZON_CLIENT_ID")
    api_key = env_first(f"OZON_API_KEY_{store}", "OZON_API_KEY")
    if not client_id or not api_key:
        raise RuntimeError(f"Не заданы OZON_CLIENT_ID_{store} и/или OZON_API_KEY_{store}")

    access_key = env_first("OZON_YC_ACCESS_KEY_ID", "YC_ACCESS_KEY_ID", "AWS_ACCESS_KEY_ID")
    secret_key = env_first("OZON_YC_SECRET_ACCESS_KEY", "YC_SECRET_ACCESS_KEY", "AWS_SECRET_ACCESS_KEY")
    endpoint = env_first("OZON_YC_ENDPOINT_URL", "YC_ENDPOINT_URL", default="https://storage.yandexcloud.net")
    if not access_key or not secret_key:
        raise RuntimeError("Не заданы ключи Yandex Object Storage")

    target = resolve_target_date(args.target_date)
    logging.info("Магазин: %s", store)
    logging.info("Режим: %s", args.mode)
    logging.info("Целевая дата: %s", target)
    p_from, p_to = mode_period(args.mode, target)
    logging.info("Период событий: %s — %s", p_from, p_to)
    if args.mode in {"test", "archive"} and p_from != p_to:
        raise RuntimeError(f"Режим {args.mode} обязан работать строго за 1 день")
    logging.info("Бакет: %s", args.bucket)

    storage = S3Storage(access_key, secret_key, args.bucket, endpoint)
    storage.ensure_bucket()
    client = OzonSellerClient(client_id, api_key)
    collector = OzonFboCollector(client, storage, store, target, args.mode, Path(args.workdir))
    collector.collect_all()
    collector.save_results()
    collector.save_diagnostics()

    ok = sum(r.status == "OK" for r in collector.results)
    empty = sum(r.status == "EMPTY" for r in collector.results)
    skipped = sum(r.status == "SKIPPED_OPTIONAL" for r in collector.results)
    failed = sum(r.status == "ERROR" for r in collector.results)
    logging.info("Готово. OK=%s, EMPTY=%s, optional skipped=%s, errors=%s", ok, empty, skipped, failed)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        logging.exception("Критическая ошибка: %s", exc)
        raise
