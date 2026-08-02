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

SCRIPT_VERSION = "OZON_FBO_ALL_DATA_V13_PERFORMANCE_BATCHES_OBJECTS_20260802"
MOSCOW_TZ = ZoneInfo("Europe/Moscow")
OZON_API_BASE = "https://api-seller.ozon.ru"
OZON_PERFORMANCE_API_BASE = "https://api-performance.ozon.ru"
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


def batched(values: Sequence[Any], size: int) -> Iterable[List[Any]]:
    """Разбивает список на пакеты допустимого размера API."""
    size = max(1, int(size))
    for idx in range(0, len(values), size):
        yield list(values[idx:idx + size])

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



class OzonPerformanceClient:
    """Клиент Performance API.

    Авторизация:
    POST /api/client/token
    client_id + client_secret + grant_type=client_credentials.

    Методы чтения сделаны с fallback-вариантами, потому что Ozon постепенно
    переводит рекламные инструменты между версиями API. Все успешные и
    неуспешные ответы сохраняются сборщиком в raw JSON.
    """

    def __init__(
        self,
        client_id: str,
        client_secret: str,
        timeout: int = 180,
    ):
        self.base = OZON_PERFORMANCE_API_BASE.rstrip("/")
        self.client_id = str(client_id)
        self.client_secret = str(client_secret)
        self.timeout = timeout
        self.session = requests.Session()
        self.session.headers.update({
            "Accept": "application/json",
            "User-Agent": f"ozon-assist/{SCRIPT_VERSION}",
        })
        self.access_token = ""
        self.token_expires_at = 0.0

    def _ensure_token(self) -> None:
        if self.access_token and time.time() < self.token_expires_at - 60:
            return

        url = self.base + "/api/client/token"
        payload = {
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "grant_type": "client_credentials",
        }
        response = self.session.post(
            url,
            json=payload,
            headers={"Content-Type": "application/json"},
            timeout=self.timeout,
        )
        if response.status_code != 200:
            raise OzonApiError(
                "POST",
                "/api/client/token",
                response.status_code,
                response.text[:2000],
                response.text[:6000],
            )

        try:
            body = response.json()
        except Exception as exc:
            raise OzonApiError(
                "POST",
                "/api/client/token",
                200,
                "Ответ token не является JSON",
                response.text[:2000],
            ) from exc

        token = (
            body.get("access_token")
            or body.get("accessToken")
            or body.get("token")
        )
        if not token:
            raise OzonApiError(
                "POST",
                "/api/client/token",
                200,
                "В ответе отсутствует access_token",
                safe_json(body),
            )

        self.access_token = str(token)
        expires_in = body.get("expires_in") or body.get("expiresIn") or 1800
        try:
            expires_seconds = int(expires_in)
        except Exception:
            expires_seconds = 1800
        self.token_expires_at = time.time() + max(300, expires_seconds)
        self.session.headers.update({
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json",
        })

    def request(
        self,
        method: str,
        path: str,
        *,
        payload: Optional[Mapping[str, Any]] = None,
        params: Optional[Mapping[str, Any]] = None,
        retries: int = 6,
    ) -> Any:
        self._ensure_token()
        url = self.base + path
        last_error = ""

        for attempt in range(1, retries + 1):
            try:
                response = self.session.request(
                    method.upper(),
                    url,
                    json=dict(payload) if payload is not None else None,
                    params=dict(params) if params is not None else None,
                    timeout=self.timeout,
                )
            except requests.RequestException as exc:
                last_error = str(exc)
                if attempt == retries:
                    raise OzonApiError(method.upper(), path, 0, last_error) from exc
                time.sleep(min(60, 2 ** attempt))
                continue

            if response.status_code in {200, 201, 202}:
                if not response.text.strip():
                    return {}
                try:
                    return response.json()
                except Exception:
                    return {
                        "_content_type": response.headers.get("Content-Type", ""),
                        "_text": response.text,
                    }

            if response.status_code == 401 and attempt < retries:
                self.access_token = ""
                self.token_expires_at = 0.0
                self._ensure_token()
                continue

            if response.status_code in {429, 500, 502, 503, 504} and attempt < retries:
                retry_after = response.headers.get("Retry-After")
                try:
                    seconds = int(float(retry_after)) if retry_after else min(90, 2 ** attempt)
                except Exception:
                    seconds = min(90, 2 ** attempt)
                logging.warning(
                    "Performance API повтор %s/%s %s после HTTP %s через %s сек",
                    attempt,
                    retries,
                    path,
                    response.status_code,
                    seconds,
                )
                time.sleep(max(1, seconds))
                continue

            text_body = response.text[:6000]
            try:
                body = response.json()
                message = (
                    body.get("message")
                    or body.get("error")
                    or body.get("code")
                    or text_body
                )
            except Exception:
                message = text_body
            raise OzonApiError(
                method.upper(),
                path,
                response.status_code,
                str(message),
                text_body,
            )

        raise OzonApiError(method.upper(), path, 0, last_error or "Неизвестная ошибка")

    def get(self, path: str, params: Optional[Mapping[str, Any]] = None) -> Any:
        return self.request("GET", path, params=params)

    def post(self, path: str, payload: Mapping[str, Any]) -> Any:
        return self.request("POST", path, payload=payload)

    def try_variants(
        self,
        variants: Sequence[Tuple[str, str, Optional[Mapping[str, Any]], Optional[Mapping[str, Any]]]],
    ) -> Tuple[Any, Dict[str, Any]]:
        """Пробует варианты method/path/payload/params до первого успеха."""
        attempts: List[Dict[str, Any]] = []
        last_exc: Optional[Exception] = None

        for method, path, payload, params in variants:
            try:
                data = self.request(
                    method,
                    path,
                    payload=payload,
                    params=params,
                )
                meta = {
                    "chosen_method": method,
                    "chosen_path": path,
                    "chosen_payload": payload,
                    "chosen_params": params,
                    "attempts": attempts,
                }
                return data, meta
            except OzonApiError as exc:
                last_exc = exc
                attempts.append({
                    "method": method,
                    "path": path,
                    "payload": payload,
                    "params": params,
                    "status": exc.status,
                    "error": str(exc),
                    "response_text": exc.response_text,
                })
                if exc.status in {400, 404, 405, 422}:
                    continue
                raise

        if last_exc:
            raise last_exc
        raise RuntimeError("Не задано ни одного варианта Performance API")

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
    def __init__(
        self,
        client: OzonSellerClient,
        storage: S3Storage,
        store: str,
        target_date: date,
        mode: str,
        workdir: Path,
        performance_client: Optional[OzonPerformanceClient] = None,
    ):
        self.client = client
        self.performance_client = performance_client
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
    def _search_period_candidates(self) -> List[Tuple[date, date]]:
        """Периоды для Premium-поиска.

        Остальные отчёты сохраняют режим test/archive=1 день и daily=90 дней.
        Только поисковая аналитика при отсутствии данных последовательно
        проверяет 1, 7, 14 и 30 календарных дней, заканчивающихся target_date.
        """
        candidates: List[Tuple[date, date]] = []
        for days in (1, 7, 14, 30):
            period_to = self.target_date
            period_from = period_to - timedelta(days=days - 1)
            pair = (period_from, period_to)
            if pair not in candidates:
                candidates.append(pair)
        return candidates

    @staticmethod
    def _is_no_search_data_error(exc: Exception) -> bool:
        message = str(exc).lower()
        return (
            "there is no data for the specified period" in message
            or "getpremiumanalyticsperiod" in message
        )

    def fetch_product_queries(self) -> Tuple[pd.DataFrame, Any, List[str]]:
        """Premium: сводная поисковая аналитика с подбором допустимого периода."""
        endpoint = "/v1/analytics/product-queries"
        sku_values = sorted({int(x) for x in self.sku_ids if x not in (None, "", 0)})
        if not sku_values:
            return pd.DataFrame(), {"warning": "sku_ids is empty"}, [endpoint]

        attempts_raw: List[Any] = []
        last_no_data_error: Optional[Exception] = None

        for search_from, search_to in self._search_period_candidates():
            logging.info(
                "Поисковые запросы — сводная: пробуем период %s — %s",
                search_from,
                search_to,
            )
            all_items: List[Dict[str, Any]] = []
            period_raw: List[Any] = []
            page_size = 1000

            try:
                for batch_no, sku_batch in enumerate(batched(sku_values, 1000), start=1):
                    page = 1
                    while page <= 500:
                        payload = {
                            "date_from": to_ozon_datetime(search_from),
                            "date_to": to_ozon_datetime(search_to, end=True),
                            "skus": sku_batch,
                            "page": page,
                            "page_size": page_size,
                        }
                        data = self.client.post(endpoint, payload)
                        period_raw.append({
                            "request": payload,
                            "response": data,
                        })
                        items = extract_items(
                            data,
                            [
                                ("items",),
                                ("result", "items"),
                                ("result", "rows"),
                                ("rows",),
                            ],
                        )
                        for item in items:
                            row = dict(item) if isinstance(item, Mapping) else {"value": item}
                            row.setdefault("sku_batch_no", batch_no)
                            row.setdefault("api_page", page)
                            all_items.append(row)

                        if not items or len(items) < page_size:
                            break
                        page += 1
                        time.sleep(0.12)

                attempts_raw.append({
                    "period_from": search_from.isoformat(),
                    "period_to": search_to.isoformat(),
                    "status": "OK",
                    "rows": len(all_items),
                    "responses": period_raw,
                })

                # Даже пустой успешный ответ считаем валидным периодом.
                df = records_to_df(all_items, {
                    "Дата снимка": self.target_date.isoformat(),
                    "Период с": search_from.isoformat(),
                    "Период по": search_to.isoformat(),
                    "Длина периода, дней": (search_to - search_from).days + 1,
                    "Тип подписки": "Premium",
                })
                logging.info(
                    "Поисковые запросы — сводная: выбран период %s — %s, строк %s",
                    search_from,
                    search_to,
                    len(df),
                )
                return df, attempts_raw, [endpoint]

            except OzonApiError as exc:
                attempts_raw.append({
                    "period_from": search_from.isoformat(),
                    "period_to": search_to.isoformat(),
                    "status": "ERROR",
                    "error": str(exc),
                    "responses": period_raw,
                })
                if self._is_no_search_data_error(exc):
                    last_no_data_error = exc
                    logging.warning(
                        "Поисковые запросы — сводная: данных за %s дней нет, пробуем больший период",
                        (search_to - search_from).days + 1,
                    )
                    continue
                raise

        if last_no_data_error:
            raise last_no_data_error
        return pd.DataFrame(), attempts_raw, [endpoint]

    def fetch_product_query_details(self) -> Tuple[pd.DataFrame, Any, List[str]]:
        """Premium: SKU × запрос с limit_by_sku=15 и подбором периода."""
        endpoint = "/v1/analytics/product-queries/details"
        sku_values = sorted({int(x) for x in self.sku_ids if x not in (None, "", 0)})
        if not sku_values:
            return pd.DataFrame(), {"warning": "sku_ids is empty"}, [endpoint]

        page_size = 100
        limit_by_sku = 15
        attempts_raw: List[Any] = []
        last_no_data_error: Optional[Exception] = None

        for search_from, search_to in self._search_period_candidates():
            logging.info(
                "Поисковые запросы — детализация: пробуем период %s — %s",
                search_from,
                search_to,
            )
            all_items: List[Dict[str, Any]] = []
            period_raw: List[Any] = []

            try:
                for batch_no, sku_batch in enumerate(batched(sku_values, 100), start=1):
                    page = 1
                    previous_signature: Optional[str] = None
                    batch_rows_before = len(all_items)

                    while page <= 1000:
                        payload = {
                            "date_from": to_ozon_datetime(search_from),
                            "date_to": to_ozon_datetime(search_to, end=True),
                            "skus": sku_batch,
                            "page": page,
                            "page_size": page_size,
                            "limit_by_sku": limit_by_sku,
                        }
                        data = self.client.post(endpoint, payload)
                        items = extract_items(
                            data,
                            [
                                ("items",),
                                ("result", "items"),
                                ("result", "rows"),
                                ("rows",),
                            ],
                        )

                        signature_parts: List[str] = []
                        for item in items[:20]:
                            if isinstance(item, Mapping):
                                signature_parts.append(
                                    json.dumps(item, ensure_ascii=False, sort_keys=True, default=str)
                                )
                            else:
                                signature_parts.append(str(item))
                        page_signature = "|".join(signature_parts)

                        pagination_meta: Dict[str, Any] = {
                            "batch_no": batch_no,
                            "page": page,
                            "page_size": page_size,
                            "limit_by_sku": limit_by_sku,
                            "skus_count": len(sku_batch),
                            "items_count": len(items),
                            "same_as_previous_page": bool(
                                previous_signature is not None
                                and page_signature == previous_signature
                            ),
                        }

                        if isinstance(data, Mapping):
                            for key in (
                                "total", "total_count", "has_next", "next_page",
                                "next_cursor", "cursor"
                            ):
                                if key in data:
                                    pagination_meta[key] = data.get(key)
                            result_obj = data.get("result")
                            if isinstance(result_obj, Mapping):
                                for key in (
                                    "total", "total_count", "has_next", "next_page",
                                    "next_cursor", "cursor"
                                ):
                                    if key in result_obj:
                                        pagination_meta[f"result.{key}"] = result_obj.get(key)

                        period_raw.append({
                            "request": payload,
                            "pagination": pagination_meta,
                            "response": data,
                        })

                        for item in items:
                            row = dict(item) if isinstance(item, Mapping) else {"value": item}
                            row.setdefault("sku_batch_no", batch_no)
                            row.setdefault("api_page", page)
                            row.setdefault("limit_by_sku", limit_by_sku)
                            all_items.append(row)

                        if (
                            previous_signature is not None
                            and page_signature == previous_signature
                        ):
                            logging.warning(
                                "Поисковые запросы: пакет %s, страница %s повторяет предыдущую; "
                                "пагинация остановлена",
                                batch_no,
                                page,
                            )
                            break
                        previous_signature = page_signature

                        has_next = None
                        next_page = None
                        total = None
                        if isinstance(data, Mapping):
                            has_next = data.get("has_next")
                            next_page = data.get("next_page")
                            total = data.get("total", data.get("total_count"))
                            result_obj = data.get("result")
                            if isinstance(result_obj, Mapping):
                                if has_next is None:
                                    has_next = result_obj.get("has_next")
                                if next_page is None:
                                    next_page = result_obj.get("next_page")
                                if total is None:
                                    total = result_obj.get(
                                        "total", result_obj.get("total_count")
                                    )

                        if not items:
                            break
                        if has_next is False:
                            break
                        if (
                            isinstance(total, (int, float))
                            and len(all_items) - batch_rows_before >= int(total)
                        ):
                            break
                        if (
                            len(items) < page_size
                            and has_next not in (True, 1, "true", "True")
                        ):
                            break

                        page = (
                            int(next_page)
                            if isinstance(next_page, int) and next_page > page
                            else page + 1
                        )
                        time.sleep(0.12)

                    logging.info(
                        "Поисковые запросы — детализация: пакет %s, SKU %s, добавлено %s",
                        batch_no,
                        len(sku_batch),
                        len(all_items) - batch_rows_before,
                    )
                    time.sleep(0.15)

                attempts_raw.append({
                    "period_from": search_from.isoformat(),
                    "period_to": search_to.isoformat(),
                    "status": "OK",
                    "rows": len(all_items),
                    "responses": period_raw,
                })

                df = records_to_df(all_items, {
                    "Дата снимка": self.target_date.isoformat(),
                    "Период с": search_from.isoformat(),
                    "Период по": search_to.isoformat(),
                    "Длина периода, дней": (search_to - search_from).days + 1,
                    "Тип подписки": "Premium",
                    "Лимит запросов на SKU": limit_by_sku,
                })

                aliases = {
                    "Поисковый запрос": ["query", "search_query", "query_text", "text"],
                    "Позиция": [
                        "position", "avg_position", "average_position", "median_position"
                    ],
                    "Пользователи поиска": [
                        "unique_search_users", "search_users"
                    ],
                    "Просмотры карточки": [
                        "unique_view_users", "view_users"
                    ],
                    "Конверсия в просмотр": [
                        "view_conversion", "conversion"
                    ],
                    "Продажи по запросу, руб": [
                        "gmv", "revenue", "order_sum"
                    ],
                    "Заказы по запросу, шт": [
                        "orders", "orders_count", "ordered_units"
                    ],
                }
                for target, candidates in aliases.items():
                    if target in df.columns:
                        continue
                    for candidate in candidates:
                        if candidate in df.columns:
                            df[target] = df[candidate]
                            break

                if (
                    "Пользователи поиска" in df.columns
                    and "Просмотры карточки" in df.columns
                    and "CTR-аналог, %" not in df.columns
                ):
                    search_users = pd.to_numeric(
                        df["Пользователи поиска"], errors="coerce"
                    )
                    view_users = pd.to_numeric(
                        df["Просмотры карточки"], errors="coerce"
                    )
                    df["CTR-аналог, %"] = view_users.div(
                        search_users.where(search_users.ne(0))
                    ).mul(100)

                logging.info(
                    "Поисковые запросы — детализация: выбран период %s — %s, строк %s",
                    search_from,
                    search_to,
                    len(df),
                )
                return df, attempts_raw, [endpoint]

            except OzonApiError as exc:
                attempts_raw.append({
                    "period_from": search_from.isoformat(),
                    "period_to": search_to.isoformat(),
                    "status": "ERROR",
                    "error": str(exc),
                    "responses": period_raw,
                })
                if self._is_no_search_data_error(exc):
                    last_no_data_error = exc
                    logging.warning(
                        "Поисковые запросы — детализация: данных за %s дней нет, пробуем больший период",
                        (search_to - search_from).days + 1,
                    )
                    continue
                raise

        if last_no_data_error:
            raise last_no_data_error
        return pd.DataFrame(), attempts_raw, [endpoint]

    # ------------------------- supplies/warehouses -------------------------
    def fetch_supply_orders(self) -> Tuple[pd.DataFrame, Any, List[str]]:
        """Поставки FBO с подбором совместимого варианта сортировки v3."""
        endpoint_list = "/v3/supply-order/list"
        base_filter = {
            "states": [],
            "created_date_from": self.period_from.isoformat(),
            "created_date_to": (self.period_to + timedelta(days=1)).isoformat(),
        }

        payload_variants: List[Dict[str, Any]] = [
            {
                "filter": base_filter,
                "limit": 100,
                "last_id": "",
                "sort_by": "CREATION_DATE",
                "sort_direction": "DESC",
            },
            {
                "filter": base_filter,
                "limit": 100,
                "last_id": "",
                "sort_by": "CREATED_AT",
                "sort_direction": "DESC",
            },
            {
                "filter": base_filter,
                "limit": 100,
                "last_id": "",
                "sort_by": 1,
                "sort_direction": 2,
            },
            {
                "filter": base_filter,
                "limit": 100,
                "last_id": "",
            },
        ]

        items: List[Dict[str, Any]] = []
        raw_list: List[Any] = []
        last_exc: Optional[Exception] = None
        chosen_payload: Optional[Dict[str, Any]] = None

        for variant_no, payload in enumerate(payload_variants, start=1):
            try:
                logging.info(
                    "Поставки FBO: пробуем вариант payload %s: sort_by=%r, sort_direction=%r",
                    variant_no,
                    payload.get("sort_by"),
                    payload.get("sort_direction"),
                )
                items, pages = self.client.cursor_pages(
                    endpoint_list,
                    payload,
                    [("result", "items"), ("items",)],
                    limit=100,
                    cursor_field="last_id",
                )
                raw_list.append({
                    "variant_no": variant_no,
                    "request": payload,
                    "status": "OK",
                    "responses": pages,
                })
                chosen_payload = payload
                break
            except OzonApiError as exc:
                last_exc = exc
                raw_list.append({
                    "variant_no": variant_no,
                    "request": payload,
                    "status": "ERROR",
                    "error": str(exc),
                })
                message = str(exc).lower()
                if exc.status == 400 and (
                    "sortby" in message
                    or "sort_by" in message
                    or "sortdirection" in message
                    or "sort_direction" in message
                ):
                    continue
                raise

        if chosen_payload is None:
            if last_exc:
                raise last_exc
            return pd.DataFrame(), raw_list, [endpoint_list]

        rows: List[Dict[str, Any]] = []
        raw_details: List[Any] = []

        for i, item in enumerate(items, start=1):
            order_id = (
                item.get("order_id")
                or item.get("supply_order_id")
                or item.get("id")
            )
            base = dict(item)
            if not order_id:
                rows.append(base)
                continue

            try:
                detail = self.client.post(
                    "/v3/supply-order/get",
                    {"order_id": order_id},
                )
                raw_details.append({
                    "order_id": order_id,
                    "response": detail,
                })
                root = (
                    detail.get("result", detail)
                    if isinstance(detail, Mapping)
                    else {}
                )
                products: Any = []
                if isinstance(root, Mapping):
                    products = (
                        root.get("items")
                        or root.get("products")
                        or root.get("supply_order_items")
                        or []
                    )

                if isinstance(products, list) and products:
                    for product in products:
                        row = dict(base)
                        if isinstance(product, Mapping):
                            row.update({
                                f"product.{k}": v
                                for k, v in product.items()
                            })
                        row["detail"] = safe_json(root)
                        rows.append(row)
                else:
                    row = dict(base)
                    row["detail"] = safe_json(root)
                    rows.append(row)

            except Exception as exc:
                self.log_error(
                    "supply_detail",
                    f"order_id={order_id}",
                    exc,
                    optional=True,
                )
                base["detail_error"] = str(exc)
                rows.append(base)

            if i % 20 == 0:
                logging.info(
                    "Поставки FBO: обработано деталей %s/%s",
                    i,
                    len(items),
                )
            time.sleep(0.12)

        raw = {
            "chosen_list_payload": chosen_payload,
            "list_attempts": raw_list,
            "details": raw_details,
        }
        df = records_to_df(rows, {
            "Дата снимка": self.target_date.isoformat(),
            "Период с": self.period_from.isoformat(),
            "Период по": self.period_to.isoformat(),
        })
        return df, raw, [endpoint_list, "/v3/supply-order/get"]

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

    # ------------------------- Performance API -------------------------

    def _require_performance(self) -> OzonPerformanceClient:
        if self.performance_client is None:
            raise RuntimeError(
                "Не заданы OZON_PERFORMANCE_CLIENT_ID и OZON_PERFORMANCE_CLIENT_SECRET"
            )
        return self.performance_client

    @staticmethod
    def _extract_performance_items(data: Any) -> List[Dict[str, Any]]:
        if isinstance(data, list):
            return [
                dict(item) if isinstance(item, Mapping) else {"value": item}
                for item in data
            ]
        if not isinstance(data, Mapping):
            return []

        paths = [
            ("list",),
            ("items",),
            ("campaigns",),
            ("products",),
            ("rows",),
            ("result",),
            ("result", "items"),
            ("result", "campaigns"),
            ("result", "products"),
            ("data",),
            ("data", "items"),
            ("data", "campaigns"),
            ("data", "products"),
        ]
        items = extract_items(data, paths)
        if items:
            return items

        # Иногда один объект возвращается без массива.
        if any(
            key in data
            for key in ("id", "campaignId", "campaign_id", "sku", "productId")
        ):
            return [dict(data)]
        return []

    @staticmethod
    def _find_object_lists(value: Any) -> List[Dict[str, Any]]:
        """Ищет списки рекламных объектов в произвольной структуре ответа."""
        found: List[Dict[str, Any]] = []

        def walk(node: Any, path: str = "") -> None:
            if isinstance(node, list):
                for item in node:
                    if isinstance(item, Mapping):
                        keys = {str(k).lower() for k in item.keys()}
                        if keys.intersection({
                            "sku", "productid", "product_id", "offerid", "offer_id",
                            "objectid", "object_id", "bid", "price", "status"
                        }):
                            row = dict(item)
                            row.setdefault("_source_path", path)
                            found.append(row)
                        walk(item, path)
            elif isinstance(node, Mapping):
                for key, child in node.items():
                    child_path = f"{path}.{key}" if path else str(key)
                    walk(child, child_path)

        walk(value)
        return found

    def fetch_ad_campaigns(self) -> Tuple[pd.DataFrame, Any, List[str]]:
        client = self._require_performance()
        data, meta = client.try_variants([
            ("GET", "/api/client/campaign", None, None),
            ("GET", "/api/client/campaigns", None, None),
            ("POST", "/api/client/campaign", {}, None),
        ])
        items = self._extract_performance_items(data)
        df = records_to_df(items, {
            "Дата снимка": self.target_date.isoformat(),
        })
        raw = {"meta": meta, "response": data}
        return df, raw, [meta["chosen_path"]]

    def _campaign_ids_from_results(self) -> List[str]:
        result = next(
            (r for r in self.results if r.code == "ad_campaigns" and not r.df.empty),
            None,
        )
        if result is None:
            return []

        candidates = [
            "id", "campaignId", "campaign_id", "campaign.id", "advCampaignId"
        ]
        ids: List[str] = []
        for column in candidates:
            if column not in result.df.columns:
                continue
            for value in result.df[column].dropna().tolist():
                text_value = str(value).strip()
                if text_value and text_value.lower() != "nan":
                    ids.append(text_value)
        return sorted(set(ids))

    def fetch_ad_products_bids(self) -> Tuple[pd.DataFrame, Any, List[str]]:
        """Получает товары и ставки по кампаниям.

        В v13 ошибка одной кампании больше не прерывает весь отчёт. Сначала
        читаем карточку кампании, затем пробуем актуальные object-endpoint'ы.
        """
        client = self._require_performance()
        campaign_ids = self._campaign_ids_from_results()
        if not campaign_ids:
            return pd.DataFrame(), {"warning": "campaign_ids is empty"}, []

        rows: List[Dict[str, Any]] = []
        diagnostics: List[Dict[str, Any]] = []
        raw: List[Any] = []
        used_paths: List[str] = []

        for index, campaign_id in enumerate(campaign_ids, start=1):
            campaign_rows: List[Dict[str, Any]] = []
            campaign_raw: Dict[str, Any] = {"campaign_id": campaign_id}

            # 1. Карточка кампании часто уже содержит objects/products.
            try:
                detail, detail_meta = client.try_variants([
                    ("GET", f"/api/client/campaign/{campaign_id}", None, None),
                ])
                used_paths.append(detail_meta["chosen_path"])
                campaign_raw["campaign_detail"] = detail
                campaign_raw["campaign_detail_meta"] = detail_meta
                campaign_rows.extend(self._find_object_lists(detail))
            except Exception as exc:
                campaign_raw["campaign_detail_error"] = str(exc)

            # 2. Отдельные методы объектов.
            object_data = None
            object_meta = None
            try:
                object_data, object_meta = client.try_variants([
                    ("GET", f"/api/client/campaign/{campaign_id}/objects", None, None),
                    ("GET", f"/api/client/campaign/{campaign_id}/object", None, None),
                    ("GET", f"/api/client/campaign/{campaign_id}/objects/info", None, None),
                    ("GET", f"/api/client/campaign/{campaign_id}/products", None, None),
                ])
                used_paths.append(object_meta["chosen_path"])
                campaign_raw["objects_response"] = object_data
                campaign_raw["objects_meta"] = object_meta
                extracted = self._extract_performance_items(object_data)
                if not extracted:
                    extracted = self._find_object_lists(object_data)
                campaign_rows.extend(extracted)
            except Exception as exc:
                campaign_raw["objects_error"] = str(exc)

            # 3. Конкурентные ставки — необязательны.
            competitive_data: Any = {}
            competitive_items: List[Dict[str, Any]] = []
            try:
                competitive_data, competitive_meta = client.try_variants([
                    ("GET", f"/api/client/campaign/{campaign_id}/objects/bids/competitive", None, None),
                    ("GET", f"/api/client/campaign/{campaign_id}/bids/competitive", None, None),
                    ("GET", f"/api/client/campaign/{campaign_id}/competitive-bids", None, None),
                ])
                used_paths.append(competitive_meta["chosen_path"])
                campaign_raw["competitive_response"] = competitive_data
                campaign_raw["competitive_meta"] = competitive_meta
                competitive_items = self._extract_performance_items(competitive_data)
                if not competitive_items:
                    competitive_items = self._find_object_lists(competitive_data)
            except Exception as exc:
                campaign_raw["competitive_error"] = str(exc)

            # Дедупликация объектов внутри кампании.
            seen: set[str] = set()
            unique_rows: List[Dict[str, Any]] = []
            for item in campaign_rows:
                key_value = (
                    item.get("sku")
                    or item.get("productId")
                    or item.get("product_id")
                    or item.get("offerId")
                    or item.get("offer_id")
                    or item.get("objectId")
                    or item.get("object_id")
                )
                signature = f"{key_value}|{safe_json(item)}"
                if signature in seen:
                    continue
                seen.add(signature)
                unique_rows.append(item)

            competitive_by_key: Dict[str, Dict[str, Any]] = {}
            for item in competitive_items:
                key_value = (
                    item.get("sku")
                    or item.get("productId")
                    or item.get("product_id")
                    or item.get("offerId")
                    or item.get("offer_id")
                    or item.get("objectId")
                    or item.get("object_id")
                )
                if key_value is not None:
                    competitive_by_key[str(key_value)] = item

            for item in unique_rows:
                row: Dict[str, Any] = {
                    "Дата снимка": self.target_date.isoformat(),
                    "campaign_id": campaign_id,
                }
                row.update(item)
                key_value = (
                    item.get("sku")
                    or item.get("productId")
                    or item.get("product_id")
                    or item.get("offerId")
                    or item.get("offer_id")
                    or item.get("objectId")
                    or item.get("object_id")
                )
                competitive = competitive_by_key.get(str(key_value), {})
                for field, value in competitive.items():
                    if field not in row:
                        row[f"competitive.{field}"] = value
                rows.append(row)

            if not unique_rows:
                diagnostics.append({
                    "Дата снимка": self.target_date.isoformat(),
                    "campaign_id": campaign_id,
                    "status": "NO_OBJECTS",
                    "campaign_detail_error": campaign_raw.get("campaign_detail_error", ""),
                    "objects_error": campaign_raw.get("objects_error", ""),
                    "competitive_error": campaign_raw.get("competitive_error", ""),
                })

            raw.append(campaign_raw)

            if index % 20 == 0:
                logging.info(
                    "Performance API: товары и ставки %s/%s кампаний, строк %s",
                    index,
                    len(campaign_ids),
                    len(rows),
                )
            time.sleep(0.10)

        if rows:
            df = records_to_df(rows, {"Дата снимка": self.target_date.isoformat()})
        else:
            # Не скрываем результат: сохраняем диагностический Excel.
            df = records_to_df(diagnostics, {
                "Дата снимка": self.target_date.isoformat(),
            })

        return df, {"campaigns": raw, "diagnostics": diagnostics}, sorted(set(used_paths))

    def _poll_statistics_report(
        self,
        client: OzonPerformanceClient,
        report_id: str,
        max_attempts: int = 60,
    ) -> Tuple[Any, List[Any], str]:
        raw: List[Any] = []
        variants = [
            f"/api/client/statistics/{report_id}",
            f"/api/client/statistics/report/{report_id}",
            f"/api/client/statistics/{report_id}/data",
        ]
        last_path = variants[0]

        for attempt in range(1, max_attempts + 1):
            for path in variants:
                try:
                    data = client.get(path)
                    last_path = path
                    raw.append({
                        "attempt": attempt,
                        "path": path,
                        "response": data,
                    })

                    status = ""
                    if isinstance(data, Mapping):
                        status = str(
                            data.get("status")
                            or data.get("state")
                            or data.get("reportStatus")
                            or ""
                        ).upper()

                    if status in {"ERROR", "FAILED", "CANCELLED"}:
                        raise RuntimeError(
                            f"Performance report {report_id}: status={status}"
                        )

                    items = self._extract_performance_items(data)
                    if items:
                        return data, raw, path

                    if status in {
                        "OK", "DONE", "READY", "COMPLETED", "SUCCESS"
                    }:
                        return data, raw, path

                except OzonApiError as exc:
                    if exc.status in {404, 405}:
                        continue
                    raise

            time.sleep(min(15, 2 + attempt // 5))

        raise RuntimeError(
            f"Performance report {report_id} не готов после {max_attempts} проверок"
        )

    def fetch_ad_statistics(self) -> Tuple[pd.DataFrame, Any, List[str]]:
        """Дневная статистика Performance API пакетами кампаний.

        Основная причина ошибки v12 — отправка всех 189 кампаний одним запросом.
        В v13 кампании обрабатываются пакетами по 10, а ошибка одного пакета
        не прерывает остальные.
        """
        client = self._require_performance()
        campaign_ids = self._campaign_ids_from_results()
        if not campaign_ids:
            return pd.DataFrame(), {"warning": "campaign_ids is empty"}, []

        date_from = self.period_from.isoformat()
        date_to = self.period_to.isoformat()

        all_rows: List[Dict[str, Any]] = []
        raw_batches: List[Any] = []
        diagnostics: List[Dict[str, Any]] = []
        used_paths: List[str] = []

        for batch_no, campaign_batch in enumerate(batched(campaign_ids, 10), start=1):
            # ID обычно числовые; готовим обе формы.
            int_ids: List[Any] = []
            for value in campaign_batch:
                try:
                    int_ids.append(int(value))
                except Exception:
                    int_ids.append(value)

            payload_variants = [
                {
                    "campaigns": int_ids,
                    "dateFrom": date_from,
                    "dateTo": date_to,
                    "groupBy": "DATE",
                },
                {
                    "campaigns": [str(x) for x in campaign_batch],
                    "dateFrom": date_from,
                    "dateTo": date_to,
                    "groupBy": "DATE",
                },
                {
                    "campaignIds": int_ids,
                    "dateFrom": date_from,
                    "dateTo": date_to,
                    "groupBy": "DATE",
                },
            ]

            batch_success = False
            batch_errors: List[Any] = []

            for payload in payload_variants:
                try:
                    data, meta = client.try_variants([
                        ("POST", "/api/client/statistics", payload, None),
                        ("POST", "/api/client/statistics/json", payload, None),
                    ])
                    used_paths.append(meta["chosen_path"])

                    direct_items = self._extract_performance_items(data)
                    if direct_items:
                        for item in direct_items:
                            row = dict(item)
                            row.setdefault("statistics_batch_no", batch_no)
                            all_rows.append(row)
                        raw_batches.append({
                            "batch_no": batch_no,
                            "campaigns": campaign_batch,
                            "request": payload,
                            "meta": meta,
                            "response": data,
                            "mode": "direct",
                        })
                        batch_success = True
                        break

                    report_id = None
                    if isinstance(data, Mapping):
                        report_id = (
                            data.get("UUID")
                            or data.get("uuid")
                            or data.get("reportId")
                            or data.get("report_id")
                            or data.get("id")
                        )

                    if not report_id:
                        batch_errors.append({
                            "request": payload,
                            "error": "Ответ без строк и без report_id",
                            "response": data,
                        })
                        continue

                    final_data, poll_raw, poll_path = self._poll_statistics_report(
                        client,
                        str(report_id),
                    )
                    used_paths.append(poll_path)
                    items = self._extract_performance_items(final_data)
                    for item in items:
                        row = dict(item)
                        row.setdefault("statistics_batch_no", batch_no)
                        row.setdefault("report_id", str(report_id))
                        all_rows.append(row)

                    raw_batches.append({
                        "batch_no": batch_no,
                        "campaigns": campaign_batch,
                        "request": payload,
                        "meta": meta,
                        "create_response": data,
                        "report_id": report_id,
                        "poll": poll_raw,
                        "final_response": final_data,
                        "mode": "async",
                    })
                    batch_success = True
                    break

                except Exception as exc:
                    batch_errors.append({
                        "request": payload,
                        "error": str(exc),
                    })

            if not batch_success:
                diagnostics.append({
                    "Дата выгрузки": self.target_date.isoformat(),
                    "batch_no": batch_no,
                    "campaigns": ",".join(map(str, campaign_batch)),
                    "status": "ERROR",
                    "errors": safe_json(batch_errors),
                })
                raw_batches.append({
                    "batch_no": batch_no,
                    "campaigns": campaign_batch,
                    "status": "ERROR",
                    "errors": batch_errors,
                })

            logging.info(
                "Performance API: статистика пакет %s/%s, успех=%s, строк всего=%s",
                batch_no,
                (len(campaign_ids) + 9) // 10,
                batch_success,
                len(all_rows),
            )
            time.sleep(0.20)

        if all_rows:
            df = records_to_df(all_rows, {
                "Дата выгрузки": self.target_date.isoformat(),
                "Период с": date_from,
                "Период по": date_to,
            })
        else:
            # Сохраняем диагностический Excel вместо полного отсутствия папки.
            df = records_to_df(diagnostics, {
                "Дата выгрузки": self.target_date.isoformat(),
                "Период с": date_from,
                "Период по": date_to,
            })

        return df, {
            "batches": raw_batches,
            "diagnostics": diagnostics,
        }, sorted(set(used_paths))

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
        if self.performance_client is not None:
            self._run(
                "ad_campaigns",
                "Реклама — кампании",
                "Реклама/Кампании",
                "Рекламные_кампании",
                self.fetch_ad_campaigns,
                "Дата снимка",
                ["Дата снимка", "id", "campaignId", "campaign_id"],
                snapshot=True,
                optional=True,
            )
            self._run(
                "ad_products_bids",
                "Реклама — товары и ставки",
                "Реклама/Товары и ставки",
                "Реклама_товары_и_ставки",
                self.fetch_ad_products_bids,
                "Дата снимка",
                ["Дата снимка", "campaign_id", "sku", "productId", "offerId"],
                optional=False,
            )
            self._run(
                "ad_statistics",
                "Реклама — дневная статистика",
                "Реклама/Статистика",
                "Реклама_статистика",
                self.fetch_ad_statistics,
                "Дата",
                ["Дата", "date", "campaignId", "campaign_id", "sku", "batch_no"],
                optional=False,
            )
        else:
            logging.warning(
                "Performance API отключён: ключи рекламного аккаунта не заданы"
            )
        self._run("product_queries", "Поисковые запросы — сводная", "Поисковые запросы", "Поисковые_запросы_сводная",
                  self.fetch_product_queries, "Дата", ["Дата", "offer_id", "sku", "query", "search_query", "query_text"], optional=True)
        self._run("product_query_details", "Поисковые запросы — товар × запрос", "Поисковые запросы по товарам",
                  "Поисковые_запросы_по_товарам", self.fetch_product_query_details, "Дата",
                  ["Дата", "offer_id", "sku", "query", "search_query", "query_text"], optional=True)
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
    performance_client_id = env_first(
        f"OZON_PERFORMANCE_CLIENT_ID_{store}",
        "OZON_PERFORMANCE_CLIENT_ID",
    )
    performance_client_secret = env_first(
        f"OZON_PERFORMANCE_CLIENT_SECRET_{store}",
        "OZON_PERFORMANCE_CLIENT_SECRET",
    )
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
    performance_client: Optional[OzonPerformanceClient] = None
    if performance_client_id and performance_client_secret:
        performance_client = OzonPerformanceClient(
            performance_client_id,
            performance_client_secret,
        )
        logging.info("Performance API: ключи найдены, рекламные отчёты включены")
    else:
        logging.warning(
            "Performance API: ключи не найдены, рекламные отчёты будут пропущены"
        )

    collector = OzonFboCollector(
        client,
        storage,
        store,
        target,
        args.mode,
        Path(args.workdir),
        performance_client=performance_client,
    )
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
