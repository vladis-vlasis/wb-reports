#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Ассистент WB — FINICK / Реклама, единый аналитический контур V4.

Что делает эта версия
---------------------
1. Один код вместо отдельных assistant_wb.py и upravlenie_reklamoy.py для FINICK.
2. Аналитика остаётся READ-ONLY: реальные ставки, бюджеты и статусы РК НЕ меняются.
3. Каждый день строится:
   - Предмет -> Артикул -> РК -> CPM-поисковые кластеры;
   - 1Д: вчера против среднего дня предыдущих 7 дней;
   - LIVE 7Д: последние 7 завершённых дней против предыдущих 7;
   - MATURE 7Д: исключаем сегодня и последние 3 дня, затем 7Д vs предыдущие 7Д;
   - экономический предел ставки CPC/CPM;
   - post-check повышений/снижений ставок;
   - сравнение собственного трафика с динамикой спроса WB;
   - переоценка изменения ставки через 14 дней.
4. Сохраняются только 3 постоянных файла:
   - state.sqlite3                — состояние/история;
   - ui_payload.json             — данные для веб-приложения;
   - Техническая_аналитика.xlsx  — проверка расчётов и работоспособности.

Источники FINICK в Yandex Object Storage
----------------------------------------
- Отчёты/Заказы/FINICK/Недельные/
- Отчёты/Реклама/FINICK/Недельные/
- Отчёты/Остатки/FINICK/Недельные/
- Отчёты/Поисковые запросы/FINICK/Запросы с WB/
  Здесь имя файла — дата выгрузки, а фактическая дата данных читается ИЗ ФАЙЛА.
  Например файл от 07.08 содержит выбранный период 06.08.

WB Promotion API, только чтение в текущем режиме
------------------------------------------------
- POST /api/advert/v1/bids/min          минимальные ставки
- GET  /adv/v1/budget                   бюджет РК
- GET  /adv/v1/balance                  баланс кабинета продвижения
- POST /adv/v0/normquery/get-bids       ставки CPM-поисковых кластеров
- POST /adv/v1/normquery/stats          дневная статистика поисковых кластеров CPC/CPM

Методы записи уже заложены, но заблокированы ALLOW_AD_WRITES=False:
- PATCH /api/advert/v1/bids             ставка РК
- POST  /adv/v0/normquery/bids          ставки поисковых кластеров
- POST  /adv/v1/budget/deposit          пополнение бюджета
- GET   /adv/v0/pause                   пауза РК
- GET   /adv/v0/start                   запуск РК
"""
from __future__ import annotations

import argparse
import io
import json
import math
import os
import re
import sqlite3
import tempfile
import time
import uuid
import zipfile
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple
from zoneinfo import ZoneInfo

import boto3
import numpy as np
import pandas as pd
import requests


# -----------------------------------------------------------------------------
# Конфигурация
# -----------------------------------------------------------------------------
VERSION = "assistant-wb-finick-unified-v4-analytics-readonly-2026-08-07"
STORE = "FINICK"
MOSCOW = ZoneInfo("Europe/Moscow")
S3_ENDPOINT_DEFAULT = "https://storage.yandexcloud.net"
WB_BASE_URL = "https://advert-api.wildberries.ru"

# Железный предохранитель текущей версии.
ALLOW_AD_WRITES = False

TARGET_DRR_PCT = float(os.getenv("FINICK_TARGET_DRR_PCT", "15"))
MATURE_LAG_DAYS = 3
LIVE_WINDOW_DAYS = 7
MATURE_WINDOW_DAYS = 7
FINAL_RECHECK_AFTER_DAYS = 14
POSTCHECK_DAYS = 3
TRAFFIC_GOOD_PCT = 10.0
TRAFFIC_WEAK_PCT = 5.0
REVENUE_DROP_ALERT_PCT = 15.0

# CPM-эксперимент: максимум 3 повышения. Первые два шага +40, третья попытка +100.
CPM_RAISE_STEPS = (40, 40, 100)
CPC_RAISE_STEP = 1
MAX_RAISE_ATTEMPTS = 3

ORDERS_PREFIX = "Отчёты/Заказы/FINICK/Недельные/"
ADS_PREFIX = "Отчёты/Реклама/FINICK/Недельные/"
STOCKS_PREFIX = "Отчёты/Остатки/FINICK/Недельные/"
MARKET_PREFIX = "Отчёты/Поисковые запросы/FINICK/Запросы с WB/"

OUTPUT_PREFIX = "Служебные файлы/Ассистент WB/FINICK/Реклама/"
STATE_KEY = OUTPUT_PREFIX + "state.sqlite3"
UI_KEY = OUTPUT_PREFIX + "ui_payload.json"
TECH_XLSX_KEY = OUTPUT_PREFIX + "Техническая_аналитика.xlsx"

# WB endpoints.
EP_BIDS = "/api/advert/v1/bids"
EP_BIDS_MIN = "/api/advert/v1/bids/min"
EP_BUDGET = "/adv/v1/budget"
EP_BUDGET_DEPOSIT = "/adv/v1/budget/deposit"
EP_BALANCE = "/adv/v1/balance"
EP_PAUSE = "/adv/v0/pause"
EP_START = "/adv/v0/start"
EP_QUERY_BIDS = "/adv/v0/normquery/get-bids"
EP_QUERY_SET_BIDS = "/adv/v0/normquery/bids"
EP_QUERY_STATS = "/adv/v1/normquery/stats"

TREND_FLAT_EPS_PCT = 0.5


@dataclass(frozen=True)
class PeriodWindow:
    code: str
    label: str
    current_from: date
    current_to: date
    base_from: date
    base_to: date
    base_daily_average: bool = False


@dataclass
class RunContext:
    as_of: date
    now_msk: datetime
    run_id: str


# -----------------------------------------------------------------------------
# Env / S3
# -----------------------------------------------------------------------------
def load_report_env() -> List[str]:
    raw = os.getenv("REPORT_ENV", "") or ""
    allowed = {
        "YC_ACCESS_KEY_ID",
        "YC_SECRET_ACCESS_KEY",
        "YC_BUCKET_NAME",
        "YC_S3_ENDPOINT",
        "FINICK_API_WB",
    }
    loaded: List[str] = []
    for raw_line in raw.splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("export "):
            line = line[7:].strip()
        if "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key in allowed and value and not os.getenv(key):
            os.environ[key] = value
            loaded.append(key)
    return loaded


def env_required(name: str) -> str:
    value = (os.getenv(name) or "").strip()
    if not value:
        raise RuntimeError(f"Не задан обязательный env/secret: {name}")
    return value


def make_s3():
    endpoint = (os.getenv("YC_S3_ENDPOINT") or "").strip() or S3_ENDPOINT_DEFAULT
    if not endpoint.startswith(("http://", "https://")):
        raise RuntimeError(f"Некорректный YC_S3_ENDPOINT: {endpoint!r}")
    return boto3.client(
        "s3",
        endpoint_url=endpoint,
        aws_access_key_id=env_required("YC_ACCESS_KEY_ID"),
        aws_secret_access_key=env_required("YC_SECRET_ACCESS_KEY"),
        region_name="ru-central1",
    )


def s3_exists(s3, bucket: str, key: str) -> bool:
    try:
        s3.head_object(Bucket=bucket, Key=key)
        return True
    except Exception:
        return False


def s3_download(s3, bucket: str, key: str) -> bytes:
    return s3.get_object(Bucket=bucket, Key=key)["Body"].read()


def s3_upload_bytes(s3, bucket: str, key: str, payload: bytes, content_type: str) -> None:
    s3.put_object(Bucket=bucket, Key=key, Body=payload, ContentType=content_type)


def list_s3_keys(s3, bucket: str, prefix: str, extensions: Sequence[str], limit: int) -> List[str]:
    exts = tuple(x.lower() for x in extensions)
    paginator = s3.get_paginator("list_objects_v2")
    rows: List[Tuple[datetime, str]] = []
    for page in paginator.paginate(Bucket=bucket, Prefix=prefix):
        for obj in page.get("Contents", []):
            key = str(obj.get("Key") or "")
            if key.lower().endswith(exts) and "~$" not in key:
                rows.append((obj.get("LastModified") or datetime.min.replace(tzinfo=ZoneInfo("UTC")), key))
    rows.sort(key=lambda x: x[0], reverse=True)
    return [k for _, k in rows[:limit]]


# -----------------------------------------------------------------------------
# Общие преобразования
# -----------------------------------------------------------------------------
def norm_text(v: Any) -> str:
    text = str(v or "").strip().lower().replace("ё", "е")
    return re.sub(r"\s+", " ", text)


def norm_col(v: Any) -> str:
    return re.sub(r"[^a-zа-я0-9]+", "", norm_text(v))


def clean_article(v: Any) -> str:
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return ""
    s = str(v).strip().replace("\\", "/")
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return s


def clean_int(v: Any) -> Optional[int]:
    try:
        if v is None or (isinstance(v, float) and math.isnan(v)):
            return None
        return int(float(v))
    except Exception:
        return None


def as_float(v: Any, default: float = np.nan) -> float:
    try:
        x = float(v)
        return x if math.isfinite(x) else default
    except Exception:
        return default


def to_num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(
        s.astype(str)
        .str.replace("\u00a0", " ", regex=False)
        .str.replace("₽", "", regex=False)
        .str.replace("%", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.replace(",", ".", regex=False),
        errors="coerce",
    )


def find_col(df: pd.DataFrame, aliases: Sequence[str]) -> Optional[str]:
    if df is None or df.empty:
        return None
    mapping = {norm_col(c): c for c in df.columns}
    for alias in aliases:
        key = norm_col(alias)
        if key in mapping:
            return mapping[key]
    return None


def col(df: pd.DataFrame, aliases: Sequence[str], default: Any = "") -> pd.Series:
    c = find_col(df, aliases)
    if c is None:
        return pd.Series([default] * len(df), index=df.index)
    return df[c]


def parse_bool(v: Any) -> bool:
    s = norm_text(v)
    if s in {"true", "1", "да", "yes", "истина"}:
        return True
    if s in {"false", "0", "нет", "no", "ложь", ""}:
        return False
    return bool(v)


def safe_div(a: Any, b: Any, default: float = np.nan) -> float:
    aa, bb = as_float(a), as_float(b)
    if pd.isna(aa) or pd.isna(bb) or bb == 0:
        return default
    return aa / bb


def pct_delta(cur: Any, base: Any) -> float:
    r = safe_div(cur, base)
    return (r - 1.0) * 100.0 if pd.notna(r) else np.nan


def canon_subject(v: Any) -> str:
    # Не пытаемся переименовывать товарные категории семантически.
    # Нормализация нужна прежде всего для регистра/пробелов.
    s = re.sub(r"\s+", " ", str(v or "").strip())
    if not s:
        return ""
    return s[:1].upper() + s[1:]


def query_norm(v: Any) -> str:
    return norm_text(v)


def jsonable(v: Any) -> Any:
    if v is None:
        return None
    if isinstance(v, (np.integer,)):
        return int(v)
    if isinstance(v, (np.floating,)):
        if np.isnan(v) or np.isinf(v):
            return None
        return float(v)
    if isinstance(v, (pd.Timestamp, datetime, date)):
        return v.isoformat()
    try:
        if pd.isna(v):
            return None
    except Exception:
        pass
    return v


def records(df: pd.DataFrame) -> List[Dict[str, Any]]:
    if df is None or df.empty:
        return []
    return [{k: jsonable(v) for k, v in row.items()} for row in df.to_dict(orient="records")]


# -----------------------------------------------------------------------------
# Чтение Excel / ZIP из S3
# -----------------------------------------------------------------------------
def read_book_sheets(raw: bytes) -> Dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(io.BytesIO(raw))
    out: Dict[str, pd.DataFrame] = {}
    for sh in xls.sheet_names:
        try:
            out[sh] = pd.read_excel(xls, sheet_name=sh)
        except Exception:
            pass
    return out


def choose_sheet(sheets: Dict[str, pd.DataFrame], preferred: Sequence[str]) -> pd.DataFrame:
    by_norm = {norm_col(k): v for k, v in sheets.items()}
    for p in preferred:
        if norm_col(p) in by_norm:
            return by_norm[norm_col(p)].copy()
    for df in sheets.values():
        if df is not None and not df.empty:
            return df.copy()
    return pd.DataFrame()


def normalize_orders(sheets: Dict[str, pd.DataFrame], key: str) -> pd.DataFrame:
    df = choose_sheet(sheets, ["Заказы", "Заказы FBS", "orders", "Лист1", "Sheet1"])
    if df.empty:
        return pd.DataFrame()
    # Основной Orders-источник даёт finishedPrice после СПП. Специализированная FBS-выгрузка
    # может вместо него содержать цены API в копейках — используем их только как fallback.
    finished = pd.Series(np.nan, index=df.index, dtype=float)
    price_source = pd.Series([""] * len(df), index=df.index, dtype=object)
    price_candidates = [
        ("finishedPrice", "finishedPrice"), ("finished_price", "finishedPrice"),
        ("Цена продажи", "Цена продажи"), ("Цена покупателя", "Цена покупателя"),
        ("Конвертированная финальная цена API", "FBS converted final API"),
        ("Конвертированная цена API", "FBS converted API"),
        ("Финальная цена API", "FBS final API"), ("Цена API", "FBS API"),
    ]
    seen_cols=set()
    for alias,label in price_candidates:
        c=find_col(df,[alias])
        if not c or c in seen_cols:
            continue
        seen_cols.add(c)
        vals=to_num(df[c])
        if vals.notna().any() and float(vals.dropna().median()) > 10000:
            vals=vals/100.0
        mask=finished.isna() & vals.notna()
        finished.loc[mask]=vals.loc[mask]
        price_source.loc[mask]=label
    price_source = price_source.replace("", "нет цены")
    cancel_col = find_col(df, ["isCancel", "is_cancel", "Отменено", "Отказ/отмена FBS"])
    srid_col = find_col(df, ["srid", "Srid", "SRID", "orderUid", "order_uid", "ID заказа FBS"])
    out = pd.DataFrame({
        "day": pd.to_datetime(col(df, ["date", "Дата", "Дата заказа", "Дата заказа покупателем"]), errors="coerce").dt.date,
        "supplier_article": col(df, ["supplierArticle", "Артикул продавца", "Артикул поставщика", "Артикул"]).map(clean_article),
        "nm_id": to_num(col(df, ["nmId", "nmID", "Артикул WB", "Код номенклатуры"])).astype("Int64"),
        "subject_raw": col(df, ["subject", "Предмет", "Название предмета"]).astype(str).str.strip(),
        "category": col(df, ["category", "Категория"]).astype(str).str.strip(),
        "finished_price": finished,
        "price_source": price_source,
        "is_cancel": df[cancel_col].map(parse_bool) if cancel_col else False,
        "srid": df[srid_col].astype(str) if srid_col else "",
        "g_number": col(df, ["gNumber", "g_number", "supplyId"]).astype(str),
        "order_channel": "FBS" if find_col(df, ["ID заказа FBS", "orderUid"]) else "ORDERS",
        "source_key": key,
    })
    out = out[out["day"].notna() & out["nm_id"].notna()].copy()
    out["nm_id"] = out["nm_id"].astype(int)
    out = out[~out["is_cancel"]].copy()
    return out


def normalize_ads(sheets: Dict[str, pd.DataFrame], key: str) -> pd.DataFrame:
    df = choose_sheet(sheets, ["Статистика_Ежедневно", "Статистика ежедневно", "daily", "Реклама", "Лист1"])
    if df.empty:
        return pd.DataFrame()
    out = pd.DataFrame({
        "day": pd.to_datetime(col(df, ["Дата", "date", "day", "dt"]), errors="coerce").dt.date,
        "campaign_id": to_num(col(df, ["ID кампании", "campaign_id", "advertId", "advert_id"])).astype("Int64"),
        "nm_id": to_num(col(df, ["Артикул WB", "nmID", "nmId", "nm_id"])).astype("Int64"),
        "campaign_name": col(df, ["Название", "campaign_name", "Кампания"]).astype(str),
        "subject_raw": col(df, ["Название предмета", "Предмет", "subject", "subject_norm"]).astype(str).str.strip(),
        "impressions": to_num(col(df, ["Показы", "impressions", "shows"])).fillna(0.0),
        "clicks": to_num(col(df, ["Клики", "clicks", "Переходы"])).fillna(0.0),
        "orders": to_num(col(df, ["Заказы", "orders"])).fillna(0.0),
        "spend": to_num(col(df, ["Расход", "Затраты", "spend"])).fillna(0.0),
        "order_sum": to_num(col(df, ["Сумма заказов", "order_sum", "revenue"])).fillna(0.0),
        "source_key": key,
    })
    out = out[out["day"].notna() & out["campaign_id"].notna()].copy()
    out["campaign_id"] = out["campaign_id"].astype(int)
    out["nm_id"] = out["nm_id"].astype("Int64")
    return out


def normalize_campaign_list(sheets: Dict[str, pd.DataFrame], key: str) -> pd.DataFrame:
    df = choose_sheet(sheets, ["Список_кампаний", "Список кампаний", "campaigns", "Кампании"])
    if df.empty:
        return pd.DataFrame()
    out = pd.DataFrame({
        "campaign_id": to_num(col(df, ["ID кампании", "campaign_id", "advertId", "advert_id"])).astype("Int64"),
        "nm_id": to_num(col(df, ["Артикул WB", "nm_id", "nmId", "nmID"])).astype("Int64"),
        "campaign_name": col(df, ["Название", "campaign_name", "Кампания"]).astype(str),
        "campaign_status": col(df, ["Статус", "campaign_status", "status"]).astype(str),
        "payment_type": col(df, ["Тип оплаты", "payment_type"]).astype(str),
        "bid_type": col(df, ["Тип ставки", "bid_type"]).astype(str),
        "search_bid": to_num(col(df, ["Ставка в поиске (руб)", "Ставка в поиске", "search_bid"])),
        "reco_bid": to_num(col(df, ["Ставка в рекомендациях (руб)", "Ставка в рекомендациях", "reco_bid"])),
        "subject_raw": col(df, ["Название предмета", "Предмет", "subject", "subject_norm"]).astype(str).str.strip(),
        "supplier_article_raw": col(df, ["Артикул продавца", "supplier_article", "supplierArticle"]).map(clean_article),
        "source_key": key,
    })
    out = out[out["campaign_id"].notna()].copy()
    out["campaign_id"] = out["campaign_id"].astype(int)
    out["nm_id"] = out["nm_id"].astype("Int64")
    return out


def normalize_stocks(sheets: Dict[str, pd.DataFrame], key: str) -> pd.DataFrame:
    df = choose_sheet(sheets, ["Остатки", "stocks", "Лист1", "Sheet1"])
    if df.empty:
        return pd.DataFrame()
    out = pd.DataFrame({
        "supplier_article": col(df, ["Артикул продавца", "supplierArticle", "Артикул поставщика"]).map(clean_article),
        "nm_id": to_num(col(df, ["Артикул WB", "nmId", "nmID"])).astype("Int64"),
        "subject": col(df, ["Предмет", "subject", "Название предмета"]).astype(str).str.strip(),
        "category": col(df, ["Категория", "category"]).astype(str).str.strip(),
        "brand": col(df, ["Бренд", "brand"]).astype(str).str.strip(),
        "source_key": key,
    })
    out = out[out["nm_id"].notna()].copy()
    out["nm_id"] = out["nm_id"].astype(int)
    return out.drop_duplicates("nm_id", keep="last")


def load_weekly_source(s3, bucket: str, prefix: str, limit: int, normalizer) -> Tuple[pd.DataFrame, List[str]]:
    frames: List[pd.DataFrame] = []
    used: List[str] = []
    keys = list_s3_keys(s3, bucket, prefix, [".xlsx", ".xlsm"], limit)
    for key in reversed(keys):
        try:
            part = normalizer(read_book_sheets(s3_download(s3, bucket, key)), key)
            if part is not None and not part.empty:
                frames.append(part)
                used.append(key)
        except Exception as exc:
            print(f"WARN source {key}: {exc}", flush=True)
    return (pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()), used


def load_ads_and_campaigns(s3, bucket: str, limit: int = 8) -> Tuple[pd.DataFrame, pd.DataFrame, List[str]]:
    ads_parts: List[pd.DataFrame] = []
    campaign_parts: List[pd.DataFrame] = []
    used: List[str] = []
    keys = list_s3_keys(s3, bucket, ADS_PREFIX, [".xlsx", ".xlsm"], limit)
    for key in reversed(keys):
        try:
            sheets = read_book_sheets(s3_download(s3, bucket, key))
            a = normalize_ads(sheets, key)
            c = normalize_campaign_list(sheets, key)
            if not a.empty:
                ads_parts.append(a)
                used.append(key)
            if not c.empty:
                campaign_parts.append(c)
        except Exception as exc:
            print(f"WARN ads {key}: {exc}", flush=True)
    ads = pd.concat(ads_parts, ignore_index=True) if ads_parts else pd.DataFrame()
    campaigns = pd.concat(campaign_parts, ignore_index=True) if campaign_parts else pd.DataFrame()
    if not campaigns.empty:
        campaigns["_seq"] = np.arange(len(campaigns))
        campaigns = campaigns.sort_values("_seq").drop_duplicates(["campaign_id", "nm_id"], keep="last").drop(columns="_seq")
    return ads, campaigns, used


def _extract_period_date(value: Any) -> Optional[date]:
    text = str(value or "")
    # WB экспорт использует DD-MM-YYYY; допускаем точки и слеши.
    m = re.search(r"(\d{2})[-./](\d{2})[-./](\d{4})", text)
    if not m:
        return None
    try:
        return date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    except Exception:
        return None


def _market_xlsx_to_rows(raw: bytes, source_key: str) -> pd.DataFrame:
    """Парсит выгрузку «Аналитика поиска» WB.

    В БД кладём и выбранный, и предыдущий день. Выбранный период имеет priority=2,
    предыдущий — priority=1. Когда на следующий день появится собственный файл для
    предыдущей даты, он перезапишет fallback.
    """
    bio = io.BytesIO(raw)
    xls = pd.ExcelFile(bio)
    general_name = next((s for s in xls.sheet_names if norm_text(s) == "общая информация"), xls.sheet_names[0])
    detail_name = next((s for s in xls.sheet_names if "деталь" in norm_text(s)), None)
    if not detail_name:
        return pd.DataFrame()

    general = pd.read_excel(xls, sheet_name=general_name, header=None)
    selected_day: Optional[date] = None
    previous_day: Optional[date] = None
    subject = ""
    for _, row in general.iterrows():
        label = norm_text(row.iloc[0] if len(row) else "")
        value = row.iloc[1] if len(row) > 1 else ""
        if "выбранный период" in label:
            selected_day = _extract_period_date(value)
        elif "предыдущий период" in label:
            previous_day = _extract_period_date(value)
        elif label == "предмет":
            subject = canon_subject(value)

    detail = pd.read_excel(xls, sheet_name=detail_name, header=1)
    if detail.empty:
        return pd.DataFrame()

    qcol = find_col(detail, ["Поисковый запрос"])
    if qcol is None:
        return pd.DataFrame()
    rows: List[Dict[str, Any]] = []

    def add_role(day_value: Optional[date], suffix: str, priority: int) -> None:
        if day_value is None:
            return
        prev = " (предыдущий период)" if suffix == "previous" else ""
        aliases = {
            "query_count": ["Количество запросов" + prev],
            "avg_daily_queries": ["Запросов в среднем за день" + prev],
            "card_clicks": ["Перешли в карточку товара" + prev],
            "baskets": ["Добавили в корзину" + prev],
            "basket_cr_pct": ["Конверсия в корзину" + prev],
            "orders": ["Заказали товаров" + prev],
            "order_cr_pct": ["Конверсия в заказ" + prev],
            "ordered_products": ["Предметов с заказами по запросу" + prev],
            "products_found": ["Количество товаров" + prev],
        }
        cols = {k: find_col(detail, a) for k, a in aliases.items()}
        for _, r in detail.iterrows():
            q = query_norm(r.get(qcol, ""))
            if not q:
                continue
            rec: Dict[str, Any] = {
                "day": day_value,
                "subject": subject,
                "query": q,
                "source_key": source_key,
                "source_role": suffix,
                "source_priority": priority,
            }
            for name, c in cols.items():
                rec[name] = as_float(r.get(c), np.nan) if c else np.nan
            # Fallback subject from "Больше всего заказов в предмете" only for diagnostics.
            top_col = find_col(detail, ["Больше всего заказов в предмете"])
            rec["top_subject"] = canon_subject(r.get(top_col, "")) if top_col else subject
            rows.append(rec)

    add_role(selected_day, "selected", 2)
    add_role(previous_day, "previous", 1)
    return pd.DataFrame(rows)


def load_market_files(s3, bucket: str, limit: int = 35) -> Tuple[pd.DataFrame, List[str]]:
    frames: List[pd.DataFrame] = []
    used: List[str] = []
    keys = list_s3_keys(s3, bucket, MARKET_PREFIX, [".xlsx", ".xlsm", ".zip"], limit)
    for key in reversed(keys):
        try:
            raw = s3_download(s3, bucket, key)
            if key.lower().endswith(".zip"):
                with zipfile.ZipFile(io.BytesIO(raw)) as z:
                    found = False
                    for info in z.infolist():
                        if info.filename.lower().endswith((".xlsx", ".xlsm")):
                            part = _market_xlsx_to_rows(z.read(info), key + "::" + info.filename)
                            if not part.empty:
                                frames.append(part)
                                found = True
                    if found:
                        used.append(key)
            else:
                part = _market_xlsx_to_rows(raw, key)
                if not part.empty:
                    frames.append(part)
                    used.append(key)
        except Exception as exc:
            print(f"WARN market {key}: {exc}", flush=True)
    return (pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()), used


# -----------------------------------------------------------------------------
# Каноническая карта товара и дедупликация
# -----------------------------------------------------------------------------
def build_product_map(stocks: pd.DataFrame, orders: pd.DataFrame, ads: Optional[pd.DataFrame] = None, campaigns: Optional[pd.DataFrame] = None) -> pd.DataFrame:
    """Каноническая карта nmID -> артикул -> предмет.

    Приоритет источников: остатки > заказы > список РК > ежедневная реклама.
    Каждое поле выбирается отдельно: пустой предмет в заказах не должен
    перекрывать непустой предмет из списка РК/рекламы.
    """
    parts: List[pd.DataFrame] = []
    if stocks is not None and not stocks.empty:
        x = stocks[["nm_id", "supplier_article", "subject", "category"]].copy(); x["source_priority"] = 40; parts.append(x)
    if orders is not None and not orders.empty:
        x = orders[["nm_id", "supplier_article", "subject_raw", "category"]].rename(columns={"subject_raw": "subject"}); x["source_priority"] = 30; parts.append(x)
    if campaigns is not None and not campaigns.empty:
        x = campaigns[["nm_id", "supplier_article_raw", "subject_raw"]].rename(columns={"supplier_article_raw":"supplier_article", "subject_raw":"subject"}).copy()
        x["category"] = ""; x["source_priority"] = 20; parts.append(x)
    if ads is not None and not ads.empty:
        x = ads[["nm_id", "subject_raw"]].rename(columns={"subject_raw":"subject"}).copy()
        x["supplier_article"] = ""; x["category"] = ""; x["source_priority"] = 10; parts.append(x)
    if not parts:
        return pd.DataFrame(columns=["nm_id", "supplier_article", "subject", "category"])
    df = pd.concat(parts, ignore_index=True, sort=False)
    df = df[df["nm_id"].notna()].copy()
    for c in ["supplier_article", "subject", "category"]:
        if c not in df.columns: df[c] = ""
        df[c] = df[c].fillna("").astype(str).str.strip()
    df["subject"] = df["subject"].map(canon_subject)
    rows=[]
    for nm,g in df.groupby("nm_id", dropna=False):
        g=g.sort_values("source_priority", ascending=False)
        rec={"nm_id":int(nm)}
        for field in ["supplier_article","subject","category"]:
            vals=g.loc[g[field].astype(str).str.strip().ne(""), field]
            rec[field]=vals.iloc[0] if len(vals) else ""
        rows.append(rec)
    return pd.DataFrame(rows)


def canonicalize_products(df: pd.DataFrame, pmap: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df.copy() if df is not None else pd.DataFrame()
    out = df.copy()
    if "nm_id" not in out.columns or pmap is None or pmap.empty:
        if "subject_raw" in out.columns:
            out["subject"] = out["subject_raw"].map(canon_subject)
        return out
    p = pmap.rename(columns={"supplier_article": "map_article", "subject": "map_subject", "category": "map_category"})
    out = out.merge(p, on="nm_id", how="left")
    # Ключевой фикс: product map по nmID — источник истины и ПЕРЕОПРЕДЕЛЯЕТ текст из рекламы.
    out["supplier_article"] = out["map_article"].fillna(out.get("supplier_article_raw", out.get("supplier_article", ""))).map(clean_article)
    raw_subj = out.get("subject_raw", out.get("subject", pd.Series([""] * len(out), index=out.index)))
    out["subject"] = out["map_subject"].fillna(raw_subj).map(canon_subject)
    # Не оставляем пустую карточку предмета в UI. При этом pmap сохраняет исходный
    # пустой subject, поэтому Data Quality всё равно покажет, какие nmID не сопоставлены.
    out["subject"] = out["subject"].fillna("").astype(str).str.strip().replace("", "Предмет не определён")
    out["category"] = out["map_category"].fillna(out.get("category", pd.Series([""] * len(out), index=out.index)))
    return out.drop(columns=[c for c in ["map_article", "map_subject", "map_category"] if c in out.columns])


def dedupe_orders_global(orders: pd.DataFrame) -> pd.DataFrame:
    if orders is None or orders.empty:
        return pd.DataFrame()
    out = orders.copy()
    srid_text = out["srid"].fillna("").astype(str).str.strip()
    srid_ok = srid_text.str.len().gt(5) & ~srid_text.str.lower().isin({"nan", "none", "nat"})
    # srid — единственный надёжный глобальный ID заказа. Если его нет (часто FBS),
    # одинаковые SKU/цена/дата НЕ являются дублем: это могут быть разные реальные заказы.
    a = out[srid_ok].drop_duplicates("srid", keep="last")
    b = out[~srid_ok].copy()
    return pd.concat([a, b], ignore_index=True)


def dedupe_ads_global(ads: pd.DataFrame) -> pd.DataFrame:
    if ads is None or ads.empty:
        return pd.DataFrame()
    out = ads.copy()
    out["_seq"] = np.arange(len(out))
    return out.sort_values("_seq").drop_duplicates(["day", "campaign_id", "nm_id"], keep="last").drop(columns="_seq")


# -----------------------------------------------------------------------------
# WB Promotion API — read now, write methods parked behind guard
# -----------------------------------------------------------------------------
class WBPromotionClient:
    def __init__(self, token: str, allow_writes: bool = False):
        self.token = token.strip()
        self.allow_writes = bool(allow_writes and ALLOW_AD_WRITES)
        self.log: List[Dict[str, Any]] = []
        self._last_query_stats_call = 0.0
        self._last_endpoint_call: Dict[str, float] = {}

    def _rate_wait(self, endpoint: str, interval_sec: float) -> None:
        last = self._last_endpoint_call.get(endpoint, 0.0)
        wait = float(interval_sec) - (time.monotonic() - last)
        if last > 0 and wait > 0:
            time.sleep(wait)
        self._last_endpoint_call[endpoint] = time.monotonic()

    @property
    def headers(self) -> Dict[str, str]:
        return {"Authorization": self.token, "Content-Type": "application/json"}

    def _call(self, method: str, endpoint: str, *, params: Optional[dict] = None, payload: Optional[dict] = None, timeout: int = 60, write: bool = False) -> Tuple[Optional[Any], int, str]:
        if write and not self.allow_writes:
            preview = {"method": method, "endpoint": endpoint, "params": params, "payload": payload}
            self.log.append({"at": datetime.now(MOSCOW).isoformat(), "method": method, "endpoint": endpoint, "status": "blocked_read_only", "request": json.dumps(preview, ensure_ascii=False), "response": "ALLOW_AD_WRITES=False"})
            return None, 0, "blocked_read_only"
        try:
            resp = requests.request(method, WB_BASE_URL + endpoint, headers=self.headers, params=params, json=payload, timeout=timeout)
            text = resp.text[:4000]
            self.log.append({"at": datetime.now(MOSCOW).isoformat(), "method": method, "endpoint": endpoint, "status": str(resp.status_code), "request": json.dumps({"params": params, "payload": payload}, ensure_ascii=False), "response": text})
            data = None
            if text:
                try:
                    data = resp.json()
                except Exception:
                    data = text
            return data, resp.status_code, text
        except Exception as exc:
            self.log.append({"at": datetime.now(MOSCOW).isoformat(), "method": method, "endpoint": endpoint, "status": "exception", "request": json.dumps({"params": params, "payload": payload}, ensure_ascii=False), "response": repr(exc)[:4000]})
            return None, -1, repr(exc)

    def get_min_bids(self, advert_id: int, nm_ids: Sequence[int], payment_type: str, placement_types: Sequence[str]) -> Dict[int, Dict[str, float]]:
        self._rate_wait(EP_BIDS_MIN, 3.05)
        payload = {"advert_id": int(advert_id), "nm_ids": [int(x) for x in nm_ids][:100], "payment_type": payment_type, "placement_types": list(placement_types)}
        data, status, _ = self._call("POST", EP_BIDS_MIN, payload=payload)
        out: Dict[int, Dict[str, float]] = {}
        if status < 200 or status >= 300 or not isinstance(data, dict):
            return out
        for item in data.get("bids", []) or []:
            nm = clean_int(item.get("nm_id"))
            if not nm:
                continue
            vals: Dict[str, float] = {}
            for b in item.get("bids", []) or []:
                typ = str(b.get("type") or "")
                kopecks = as_float(b.get("value"))
                if typ and pd.notna(kopecks):
                    vals[typ] = kopecks / 100.0
            out[nm] = vals
        return out

    def get_budget(self, advert_id: int) -> Dict[str, Any]:
        self._rate_wait(EP_BUDGET, 0.27)
        data, status, _ = self._call("GET", EP_BUDGET, params={"id": int(advert_id)})
        return data if 200 <= status < 300 and isinstance(data, dict) else {}

    def get_balance(self) -> Dict[str, Any]:
        self._rate_wait(EP_BALANCE, 1.05)
        data, status, _ = self._call("GET", EP_BALANCE)
        return data if 200 <= status < 300 and isinstance(data, dict) else {}

    def get_query_bids(self, items: Sequence[Tuple[int, int]]) -> pd.DataFrame:
        rows: List[Dict[str, Any]] = []
        uniq = sorted(set((int(a), int(n)) for a, n in items if a and n))
        for i in range(0, len(uniq), 100):
            chunk = uniq[i:i+100]
            payload = {"items": [{"advert_id": a, "nm_id": n} for a, n in chunk]}
            self._rate_wait(EP_QUERY_BIDS, 0.22)
            data, status, _ = self._call("POST", EP_QUERY_BIDS, payload=payload)
            if not (200 <= status < 300) or data is None:
                continue
            candidates = data.get("bids", data.get("items", [])) if isinstance(data, dict) else data
            if not isinstance(candidates, list):
                continue
            for item in candidates:
                if not isinstance(item, dict):
                    continue
                q = item.get("norm_query") or item.get("normQuery") or item.get("query")
                cid = clean_int(item.get("advert_id") or item.get("advertId"))
                nm = clean_int(item.get("nm_id") or item.get("nmId"))
                bid = as_float(item.get("bid"))
                if q and cid and nm:
                    rows.append({"campaign_id": cid, "nm_id": nm, "query": query_norm(q), "query_bid": bid, "bid_source": "wb_get_bids"})
            time.sleep(0.25)
        return pd.DataFrame(rows).drop_duplicates(["campaign_id", "nm_id", "query"], keep="last") if rows else pd.DataFrame(columns=["campaign_id", "nm_id", "query", "query_bid", "bid_source"])

    @staticmethod
    def _flatten_query_stats(obj: Any, default_cid: Optional[int] = None, default_nm: Optional[int] = None, parent_query: str = "") -> List[Dict[str, Any]]:
        rows: List[Dict[str, Any]] = []
        if isinstance(obj, list):
            for x in obj:
                rows.extend(WBPromotionClient._flatten_query_stats(x, default_cid, default_nm, parent_query))
            return rows
        if not isinstance(obj, dict):
            return rows
        cid = clean_int(obj.get("advert_id") or obj.get("advertId")) or default_cid
        nm = clean_int(obj.get("nm_id") or obj.get("nmId")) or default_nm
        q = obj.get("norm_query") or obj.get("normQuery") or obj.get("query") or obj.get("name") or parent_query
        aliases = {
            "impressions": ["views", "impressions", "shows", "show", "shw"],
            "clicks": ["clicks", "click", "clk"],
            "orders": ["orders", "shks", "ordered", "order", "orders_count", "orderCount"],
            "spend": ["spend", "expense", "cost", "costs", "ad_spend", "sum"],
            "ctr": ["ctr"],
            "cpc": ["cpc"],
            "cpm": ["cpm"],
            "position": ["position", "avg_position", "avgPosition", "median_position", "medianPosition", "pos"],
            "visibility": ["visibility", "visible", "visibility_pct", "visibilityPercent"],
        }
        lower = {str(k).lower(): k for k in obj.keys()}
        metrics: Dict[str, Any] = {}
        for out_name, names in aliases.items():
            for name in names:
                if name.lower() in lower:
                    metrics[out_name] = obj.get(lower[name.lower()])
                    break
        if q and cid and nm and metrics:
            row = {"campaign_id": cid, "nm_id": nm, "query": query_norm(q)}
            row.update(metrics)
            rows.append(row)
        for v in obj.values():
            if isinstance(v, (dict, list)):
                rows.extend(WBPromotionClient._flatten_query_stats(v, cid, nm, str(q or parent_query)))
        return rows

    def get_query_stats(self, items: Sequence[Tuple[int, int]], date_from: date, date_to: date) -> pd.DataFrame:
        """Дневная статистика поисковых кластеров WB /adv/v1/normquery/stats.

        Метод используется и для CPM, и для CPC. Для CPC WB не возвращает views/CTR/CPM,
        поэтому эти поля остаются NaN — нули не подставляем.
        """
        uniq = sorted(set((int(a), int(n)) for a, n in items if a and n))
        rows: List[Dict[str, Any]] = []
        for i in range(0, len(uniq), 100):
            chunk = uniq[i:i+100]
            # В V1 по официальной схеме items используют camelCase advertId/nmId.
            payload = {
                "from": date_from.isoformat(),
                "to": date_to.isoformat(),
                "items": [{"advertId": a, "nmId": n} for a, n in chunk],
            }
            # Лимит normquery/stats: интервал около 6 секунд.
            wait = 6.1 - (time.monotonic() - self._last_query_stats_call)
            if self._last_query_stats_call > 0 and wait > 0:
                time.sleep(wait)
            data, status, _ = self._call("POST", EP_QUERY_STATS, payload=payload)
            self._last_query_stats_call = time.monotonic()
            if not (200 <= status < 300) or not isinstance(data, dict):
                continue
            for item in data.get("items", []) or []:
                if not isinstance(item, dict):
                    continue
                cid = clean_int(item.get("advertId") or item.get("advert_id"))
                nm = clean_int(item.get("nmId") or item.get("nm_id"))
                if not cid or not nm:
                    continue
                for daily in item.get("dailyStats", item.get("daily_stats", [])) or []:
                    if not isinstance(daily, dict):
                        continue
                    stat = daily.get("stat") or daily.get("stats") or {}
                    stats_list = stat if isinstance(stat, list) else [stat]
                    for st in stats_list:
                        if not isinstance(st, dict):
                            continue
                        q = st.get("normQuery") or st.get("norm_query") or st.get("query")
                        if not q:
                            continue
                        rows.append({
                            "day": pd.to_datetime(daily.get("date"), errors="coerce").date() if pd.notna(pd.to_datetime(daily.get("date"), errors="coerce")) else None,
                            "campaign_id": cid,
                            "nm_id": nm,
                            "query": query_norm(q),
                            "impressions": as_float(st.get("views"), np.nan),
                            "clicks": as_float(st.get("clicks"), np.nan),
                            "orders": as_float(st.get("orders"), np.nan),
                            "ordered_units": as_float(st.get("shks"), np.nan),
                            "spend": as_float(st.get("spend"), np.nan),
                            "ctr_api": as_float(st.get("ctr"), np.nan),
                            "cpc_api": as_float(st.get("cpc"), np.nan),
                            "cpm_api": as_float(st.get("cpm"), np.nan),
                            "position": as_float(st.get("avgPos", st.get("avg_pos")), np.nan),
                            "visibility": as_float(st.get("visibility", st.get("visibilityPct")), np.nan),
                        })
        if not rows:
            return pd.DataFrame(columns=["campaign_id", "nm_id", "query", "impressions", "clicks", "orders", "ordered_units", "spend", "ctr", "cpc", "cpm", "position", "visibility", "days_seen"])
        raw = pd.DataFrame(rows)

        def sum_or_nan(series: pd.Series) -> float:
            x = pd.to_numeric(series, errors="coerce")
            return float(x.sum()) if x.notna().any() else np.nan

        def weighted(g: pd.DataFrame, name: str) -> float:
            vals = pd.to_numeric(g[name], errors="coerce")
            mask = vals.notna()
            if not mask.any():
                return np.nan
            w_impr = pd.to_numeric(g["impressions"], errors="coerce").fillna(0.0)
            w_click = pd.to_numeric(g["clicks"], errors="coerce").fillna(0.0)
            weights = w_impr if w_impr[mask].sum() > 0 else w_click
            if weights[mask].sum() > 0:
                return float((vals[mask] * weights[mask]).sum() / weights[mask].sum())
            return float(vals[mask].mean())

        agg_rows: List[Dict[str, Any]] = []
        for keys, g in raw.groupby(["campaign_id", "nm_id", "query"], dropna=False):
            impr = sum_or_nan(g["impressions"]); clicks = sum_or_nan(g["clicks"]); orders = sum_or_nan(g["orders"]); units = sum_or_nan(g["ordered_units"]); spend = sum_or_nan(g["spend"])
            ctr = safe_div(clicks * 100.0, impr) if pd.notna(impr) else weighted(g, "ctr_api")
            cpc = safe_div(spend, clicks) if pd.notna(spend) and pd.notna(clicks) else weighted(g, "cpc_api")
            cpm = safe_div(spend * 1000.0, impr) if pd.notna(spend) and pd.notna(impr) else weighted(g, "cpm_api")
            agg_rows.append({
                "campaign_id": int(keys[0]), "nm_id": int(keys[1]), "query": keys[2],
                "impressions": impr, "clicks": clicks, "orders": orders, "ordered_units": units, "spend": spend,
                "ctr": ctr, "cpc": cpc, "cpm": cpm, "position": weighted(g, "position"), "visibility": weighted(g, "visibility"),
                "days_seen": int(pd.Series(g["day"]).dropna().nunique()),
            })
        return pd.DataFrame(agg_rows)

    # --- Будущие write методы. Сейчас всегда preview/blocked. ---
    def set_campaign_bid(self, advert_id: int, nm_id: int, bid_rub: float, placement: str) -> Tuple[Optional[Any], int, str]:
        payload = {"bids": [{"advert_id": int(advert_id), "nm_bids": [{"nm_id": int(nm_id), "bid_kopecks": int(round(float(bid_rub) * 100)), "placement": placement}]}]}
        return self._call("PATCH", EP_BIDS, payload=payload, write=True)

    def set_query_bids(self, rows: Sequence[Dict[str, Any]]) -> Tuple[Optional[Any], int, str]:
        payload = {"bids": [{"advert_id": int(r["campaign_id"]), "nm_id": int(r["nm_id"]), "norm_query": str(r["query"]), "bid": int(round(float(r["bid_rub"])))} for r in rows][:100]}
        return self._call("POST", EP_QUERY_SET_BIDS, payload=payload, write=True)

    def deposit_budget(self, advert_id: int, amount_rub: int, cashback_sum: Optional[int] = None, cashback_percent: Optional[int] = None) -> Tuple[Optional[Any], int, str]:
        payload: Dict[str, Any] = {"sum": int(amount_rub)}
        if cashback_sum is not None:
            payload["cashback_sum"] = int(cashback_sum)
        if cashback_percent is not None:
            payload["cashback_percent"] = int(cashback_percent)
        return self._call("POST", EP_BUDGET_DEPOSIT, params={"id": int(advert_id)}, payload=payload, write=True)

    def pause_campaign(self, advert_id: int) -> Tuple[Optional[Any], int, str]:
        return self._call("GET", EP_PAUSE, params={"id": int(advert_id)}, write=True)

    def start_campaign(self, advert_id: int) -> Tuple[Optional[Any], int, str]:
        return self._call("GET", EP_START, params={"id": int(advert_id)}, write=True)


# -----------------------------------------------------------------------------
# SQLite state
# -----------------------------------------------------------------------------
class StateDB:
    def __init__(self, path: Path):
        self.path = path
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.conn = sqlite3.connect(str(path))
        self.conn.row_factory = sqlite3.Row
        self._init_schema()

    def close(self) -> None:
        self.conn.commit()
        self.conn.close()

    def _init_schema(self) -> None:
        self.conn.executescript("""
        CREATE TABLE IF NOT EXISTS runs (
            run_id TEXT PRIMARY KEY,
            as_of TEXT NOT NULL,
            generated_at TEXT NOT NULL,
            status TEXT NOT NULL,
            note TEXT
        );
        CREATE TABLE IF NOT EXISTS market_daily (
            day TEXT NOT NULL,
            subject TEXT NOT NULL,
            query TEXT NOT NULL,
            query_count REAL,
            avg_daily_queries REAL,
            card_clicks REAL,
            baskets REAL,
            basket_cr_pct REAL,
            orders REAL,
            order_cr_pct REAL,
            ordered_products REAL,
            products_found REAL,
            source_key TEXT,
            source_priority INTEGER NOT NULL DEFAULT 1,
            PRIMARY KEY(day, subject, query)
        );
        CREATE TABLE IF NOT EXISTS campaign_snapshots (
            as_of TEXT NOT NULL,
            campaign_id INTEGER NOT NULL,
            nm_id INTEGER,
            supplier_article TEXT,
            subject TEXT,
            campaign_type TEXT,
            status TEXT,
            bid_current REAL,
            min_bid REAL,
            budget_total REAL,
            live7_impressions REAL,
            live7_clicks REAL,
            live7_spend REAL,
            live7_orders REAL,
            live7_order_sum REAL,
            live7_ctr REAL,
            live7_cr REAL,
            live7_cpo REAL,
            live7_drr REAL,
            max_bid_live REAL,
            PRIMARY KEY(as_of, campaign_id, nm_id)
        );
        CREATE TABLE IF NOT EXISTS query_snapshots (
            as_of TEXT NOT NULL,
            stats_day TEXT NOT NULL,
            campaign_id INTEGER NOT NULL,
            nm_id INTEGER NOT NULL,
            query TEXT NOT NULL,
            query_bid REAL,
            impressions REAL,
            clicks REAL,
            spend REAL,
            orders REAL,
            position REAL,
            visibility REAL,
            impression_share_pct REAL,
            click_share_pct REAL,
            market_queries REAL,
            market_growth_pct REAL,
            PRIMARY KEY(as_of, stats_day, campaign_id, nm_id, query)
        );
        CREATE TABLE IF NOT EXISTS bid_changes (
            event_id TEXT PRIMARY KEY,
            detected_on TEXT NOT NULL,
            effective_day TEXT NOT NULL,
            campaign_id INTEGER NOT NULL,
            nm_id INTEGER,
            supplier_article TEXT,
            subject TEXT,
            campaign_type TEXT,
            old_bid REAL NOT NULL,
            new_bid REAL NOT NULL,
            direction TEXT NOT NULL,
            attempt_no INTEGER NOT NULL DEFAULT 1,
            traffic_verdict TEXT,
            traffic_reason TEXT,
            mature_verdict TEXT,
            mature_reason TEXT,
            final14_verdict TEXT,
            final14_reason TEXT,
            last_evaluated_on TEXT
        );
        CREATE TABLE IF NOT EXISTS event_query_baselines (
            event_id TEXT NOT NULL,
            campaign_id INTEGER NOT NULL,
            nm_id INTEGER NOT NULL,
            query TEXT NOT NULL,
            impressions_avg REAL,
            clicks_avg REAL,
            position REAL,
            visibility REAL,
            market_queries_avg REAL,
            PRIMARY KEY(event_id, query)
        );
        CREATE TABLE IF NOT EXISTS api_read_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            run_id TEXT,
            at TEXT,
            method TEXT,
            endpoint TEXT,
            status TEXT,
            request TEXT,
            response TEXT
        );
        CREATE INDEX IF NOT EXISTS idx_bid_changes_campaign ON bid_changes(campaign_id, effective_day);
        CREATE INDEX IF NOT EXISTS idx_market_daily_day ON market_daily(day);
        """)
        self.conn.commit()

    def start_run(self, ctx: RunContext) -> None:
        self.conn.execute("INSERT OR REPLACE INTO runs(run_id,as_of,generated_at,status,note) VALUES(?,?,?,?,?)", (ctx.run_id, ctx.as_of.isoformat(), ctx.now_msk.isoformat(), "running", ""))
        self.conn.commit()

    def finish_run(self, ctx: RunContext, status: str, note: str = "") -> None:
        self.conn.execute("UPDATE runs SET status=?, note=? WHERE run_id=?", (status, note[:2000], ctx.run_id))
        self.conn.commit()

    def upsert_market(self, df: pd.DataFrame) -> None:
        if df is None or df.empty:
            return
        sql = """
        INSERT INTO market_daily(day,subject,query,query_count,avg_daily_queries,card_clicks,baskets,basket_cr_pct,orders,order_cr_pct,ordered_products,products_found,source_key,source_priority)
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        ON CONFLICT(day,subject,query) DO UPDATE SET
          query_count=excluded.query_count, avg_daily_queries=excluded.avg_daily_queries,
          card_clicks=excluded.card_clicks, baskets=excluded.baskets,
          basket_cr_pct=excluded.basket_cr_pct, orders=excluded.orders,
          order_cr_pct=excluded.order_cr_pct, ordered_products=excluded.ordered_products,
          products_found=excluded.products_found, source_key=excluded.source_key,
          source_priority=excluded.source_priority
        WHERE excluded.source_priority >= market_daily.source_priority
        """
        rows = []
        for _, r in df.iterrows():
            rows.append((
                r["day"].isoformat() if isinstance(r["day"], date) else str(r["day"]), canon_subject(r.get("subject")), query_norm(r.get("query")),
                jsonable(r.get("query_count")), jsonable(r.get("avg_daily_queries")), jsonable(r.get("card_clicks")), jsonable(r.get("baskets")),
                jsonable(r.get("basket_cr_pct")), jsonable(r.get("orders")), jsonable(r.get("order_cr_pct")), jsonable(r.get("ordered_products")),
                jsonable(r.get("products_found")), str(r.get("source_key") or ""), int(r.get("source_priority") or 1),
            ))
        self.conn.executemany(sql, rows)
        self.conn.commit()

    def market_df(self, start: date, end: date) -> pd.DataFrame:
        df = pd.read_sql_query("SELECT * FROM market_daily WHERE day BETWEEN ? AND ?", self.conn, params=[start.isoformat(), end.isoformat()])
        if not df.empty:
            df["day"] = pd.to_datetime(df["day"]).dt.date
        return df

    def last_campaign_snapshot(self, campaign_id: int, nm_id: Optional[int]) -> Optional[Dict[str, Any]]:
        row = self.conn.execute(
            "SELECT * FROM campaign_snapshots WHERE campaign_id=? AND COALESCE(nm_id,-1)=COALESCE(?,-1) ORDER BY as_of DESC LIMIT 1",
            (int(campaign_id), nm_id),
        ).fetchone()
        return dict(row) if row else None

    def insert_campaign_snapshot(self, r: Dict[str, Any], as_of: date) -> None:
        cols = ["campaign_id","nm_id","supplier_article","subject","campaign_type","campaign_status","bid_current","min_bid","budget_total","live7_impressions","live7_clicks","live7_spend","live7_orders","live7_order_sum","live7_ctr","live7_cr","live7_cpo","live7_drr","max_bid_live"]
        vals = [r.get(c) for c in cols]
        self.conn.execute(
            "INSERT OR REPLACE INTO campaign_snapshots(as_of,campaign_id,nm_id,supplier_article,subject,campaign_type,status,bid_current,min_bid,budget_total,live7_impressions,live7_clicks,live7_spend,live7_orders,live7_order_sum,live7_ctr,live7_cr,live7_cpo,live7_drr,max_bid_live) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            [as_of.isoformat(), vals[0], vals[1], vals[2], vals[3], vals[4], vals[5], *vals[6:]],
        )

    def last_change(self, campaign_id: int) -> Optional[Dict[str, Any]]:
        row = self.conn.execute("SELECT * FROM bid_changes WHERE campaign_id=? ORDER BY effective_day DESC, detected_on DESC LIMIT 1", (int(campaign_id),)).fetchone()
        return dict(row) if row else None

    def changes_df(self, start: Optional[date] = None, end: Optional[date] = None) -> pd.DataFrame:
        sql = "SELECT * FROM bid_changes"
        params: List[Any] = []
        where: List[str] = []
        if start:
            where.append("effective_day>=?")
            params.append(start.isoformat())
        if end:
            where.append("effective_day<=?")
            params.append(end.isoformat())
        if where:
            sql += " WHERE " + " AND ".join(where)
        sql += " ORDER BY effective_day DESC, detected_on DESC"
        return pd.read_sql_query(sql, self.conn, params=params)

    def detect_change(self, current: Dict[str, Any], as_of: date) -> Optional[str]:
        cid = int(current["campaign_id"])
        nm = clean_int(current.get("nm_id"))
        bid = as_float(current.get("bid_current"))
        if pd.isna(bid) or bid <= 0:
            return None
        prev = self.last_campaign_snapshot(cid, nm)
        if not prev:
            return None
        old = as_float(prev.get("bid_current"))
        if pd.isna(old) or old <= 0 or abs(old - bid) < 1e-9:
            return None
        direction = "raise" if bid > old else "lower"
        # Изменение обнаружено текущим запуском. Не приписываем его вчерашнему трафику:
        # первый post-check начнётся только когда появятся данные после даты обнаружения.
        effective_day = as_of
        # Не создаём второй event, если этот переход уже был зафиксирован.
        existing = self.conn.execute(
            "SELECT event_id FROM bid_changes WHERE campaign_id=? AND old_bid=? AND new_bid=? AND effective_day=? LIMIT 1",
            (cid, old, bid, effective_day.isoformat()),
        ).fetchone()
        if existing:
            return str(existing[0])
        attempt_no = 1
        last = self.last_change(cid)
        if direction == "raise" and last and last.get("direction") == "raise":
            last_day = pd.to_datetime(last.get("effective_day"), errors="coerce")
            if pd.notna(last_day) and (pd.Timestamp(effective_day) - last_day).days <= 14:
                attempt_no = min(MAX_RAISE_ATTEMPTS, int(last.get("attempt_no") or 1) + 1)
        event_id = str(uuid.uuid4())
        self.conn.execute(
            "INSERT INTO bid_changes(event_id,detected_on,effective_day,campaign_id,nm_id,supplier_article,subject,campaign_type,old_bid,new_bid,direction,attempt_no,last_evaluated_on) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (event_id, as_of.isoformat(), effective_day.isoformat(), cid, nm, current.get("supplier_article"), current.get("subject"), current.get("campaign_type"), old, bid, direction, attempt_no, as_of.isoformat()),
        )
        self.conn.commit()
        return event_id

    def save_event_query_baseline(self, event_id: str, base_df: pd.DataFrame, market: pd.DataFrame) -> None:
        if base_df is None or base_df.empty:
            return
        market_query = market.groupby("query", as_index=True)["query_count"].mean().to_dict() if market is not None and not market.empty else {}
        rows = []
        for _, r in base_df.iterrows():
            rows.append((event_id, int(r["campaign_id"]), int(r["nm_id"]), query_norm(r["query"]), safe_div(r.get("impressions"), 7.0), safe_div(r.get("clicks"), 7.0), jsonable(r.get("position")), jsonable(r.get("visibility")), jsonable(market_query.get(query_norm(r["query"])))) )
        self.conn.executemany("INSERT OR REPLACE INTO event_query_baselines(event_id,campaign_id,nm_id,query,impressions_avg,clicks_avg,position,visibility,market_queries_avg) VALUES(?,?,?,?,?,?,?,?,?)", rows)
        self.conn.commit()

    def event_query_baselines(self, event_id: str) -> pd.DataFrame:
        return pd.read_sql_query("SELECT * FROM event_query_baselines WHERE event_id=?", self.conn, params=[event_id])

    def update_event(self, event_id: str, **fields: Any) -> None:
        allowed = {"traffic_verdict","traffic_reason","mature_verdict","mature_reason","final14_verdict","final14_reason","last_evaluated_on"}
        use = {k: v for k, v in fields.items() if k in allowed}
        if not use:
            return
        sql = "UPDATE bid_changes SET " + ",".join(f"{k}=?" for k in use) + " WHERE event_id=?"
        self.conn.execute(sql, list(use.values()) + [event_id])
        self.conn.commit()

    def insert_query_snapshot(self, as_of: date, stats_day: date, q: pd.DataFrame) -> None:
        if q is None or q.empty:
            return
        rows = []
        for _, r in q.iterrows():
            rows.append((as_of.isoformat(), stats_day.isoformat(), int(r["campaign_id"]), int(r["nm_id"]), query_norm(r["query"]), jsonable(r.get("query_bid")), jsonable(r.get("impressions")), jsonable(r.get("clicks")), jsonable(r.get("spend")), jsonable(r.get("orders")), jsonable(r.get("position")), jsonable(r.get("visibility")), jsonable(r.get("impression_share_pct")), jsonable(r.get("click_share_pct")), jsonable(r.get("market_queries")), jsonable(r.get("market_growth_pct"))))
        self.conn.executemany("INSERT OR REPLACE INTO query_snapshots(as_of,stats_day,campaign_id,nm_id,query,query_bid,impressions,clicks,spend,orders,position,visibility,impression_share_pct,click_share_pct,market_queries,market_growth_pct) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", rows)
        self.conn.commit()

    def write_api_log(self, run_id: str, log_rows: Sequence[Dict[str, Any]]) -> None:
        rows = [(run_id, str(r.get("at") or ""), str(r.get("method") or ""), str(r.get("endpoint") or ""), str(r.get("status") or ""), str(r.get("request") or "")[:5000], str(r.get("response") or "")[:5000]) for r in log_rows]
        if rows:
            self.conn.executemany("INSERT INTO api_read_log(run_id,at,method,endpoint,status,request,response) VALUES(?,?,?,?,?,?,?)", rows)
            # Держим лог компактным: только последние 2000 строк.
            self.conn.execute("DELETE FROM api_read_log WHERE id NOT IN (SELECT id FROM api_read_log ORDER BY id DESC LIMIT 2000)")
            self.conn.commit()


# -----------------------------------------------------------------------------
# Периоды и базовые агрегаты
# -----------------------------------------------------------------------------
def build_windows(as_of: date) -> Dict[str, PeriodWindow]:
    yday = as_of - timedelta(days=1)
    mature_end = as_of - timedelta(days=MATURE_LAG_DAYS + 1)
    mature_start = mature_end - timedelta(days=MATURE_WINDOW_DAYS - 1)
    mature_base_end = mature_start - timedelta(days=1)
    mature_base_start = mature_base_end - timedelta(days=MATURE_WINDOW_DAYS - 1)
    return {
        "1d": PeriodWindow("1d", "Вчера vs средний день предыдущих 7 дней", yday, yday, as_of - timedelta(days=8), as_of - timedelta(days=2), True),
        "live7": PeriodWindow("live7", "Последние 7 завершённых дней vs предыдущие 7 дней", as_of - timedelta(days=7), yday, as_of - timedelta(days=14), as_of - timedelta(days=8), False),
        "mature7": PeriodWindow("mature7", "Зрелые 7 дней: последние 3 дня исключены", mature_start, mature_end, mature_base_start, mature_base_end, False),
    }


def date_filter(df: pd.DataFrame, start: date, end: date) -> pd.DataFrame:
    if df is None or df.empty or "day" not in df.columns:
        return pd.DataFrame(columns=df.columns if df is not None else [])
    return df[(df["day"] >= start) & (df["day"] <= end)].copy()


def aggregate_metrics(df: pd.DataFrame, keys: List[str], value_cols: List[str], w: PeriodWindow) -> pd.DataFrame:
    cur = date_filter(df, w.current_from, w.current_to)
    base = date_filter(df, w.base_from, w.base_to)
    c = cur.groupby(keys, dropna=False, as_index=False)[value_cols].sum() if not cur.empty else pd.DataFrame(columns=keys + value_cols)
    b = base.groupby(keys, dropna=False, as_index=False)[value_cols].sum() if not base.empty else pd.DataFrame(columns=keys + value_cols)
    if w.base_daily_average and not b.empty:
        b[value_cols] = b[value_cols] / 7.0
    c = c.rename(columns={x: x + "_current" for x in value_cols})
    b = b.rename(columns={x: x + "_base" for x in value_cols})
    out = c.merge(b, on=keys, how="outer")
    for x in value_cols:
        for suf in ["current", "base"]:
            cc = x + "_" + suf
            if cc not in out.columns:
                out[cc] = 0.0
            out[cc] = pd.to_numeric(out[cc], errors="coerce").fillna(0.0)
    return out


def add_ad_derived(out: pd.DataFrame, prefix: str = "") -> pd.DataFrame:
    if out.empty:
        return out
    p = prefix
    for suf in ["current", "base"]:
        impr = out[f"{p}impressions_{suf}"]
        clicks = out[f"{p}clicks_{suf}"]
        orders = out[f"{p}orders_{suf}"]
        spend = out[f"{p}spend_{suf}"]
        revenue = out[f"{p}order_sum_{suf}"]
        out[f"{p}ctr_{suf}"] = np.where(impr > 0, clicks / impr * 100.0, np.nan)
        out[f"{p}cr_{suf}"] = np.where(clicks > 0, orders / clicks * 100.0, np.nan)
        out[f"{p}cpc_{suf}"] = np.where(clicks > 0, spend / clicks, np.nan)
        out[f"{p}cpm_{suf}"] = np.where(impr > 0, spend / impr * 1000.0, np.nan)
        out[f"{p}cpo_{suf}"] = np.where(orders > 0, spend / orders, np.nan)
        out[f"{p}drr_{suf}"] = np.where(revenue > 0, spend / revenue * 100.0, np.nan)
    return out


def add_deltas(df: pd.DataFrame, metrics: Sequence[str]) -> pd.DataFrame:
    out = df.copy()
    for m in metrics:
        c, b = m + "_current", m + "_base"
        if c in out.columns and b in out.columns:
            out[m + "_delta_pct"] = [pct_delta(x, y) for x, y in zip(out[c], out[b])]
    return out


# -----------------------------------------------------------------------------
# Типы кампаний, ставки и экономический предел
# -----------------------------------------------------------------------------
def campaign_status_code(value: Any) -> Optional[int]:
    text = str(value or "").strip().lower()
    try:
        num = int(float(text))
        if num in {-1, 4, 7, 8, 9, 11}:
            return num
    except Exception:
        pass
    if text in {"активна", "активен", "active"}:
        return 9
    if "пауз" in text or text == "paused":
        return 11
    if "готов" in text or text == "ready":
        return 4
    if "заверш" in text or text == "finished":
        return 7
    if "отмен" in text or text == "cancelled":
        return 8
    return None


def is_active_campaign_status(value: Any) -> bool:
    return campaign_status_code(value) == 9


def is_api_relevant_campaign_status(value: Any) -> bool:
    # Для чтения текущих ставок/бюджета достаточно активных, готовых и паузных РК.
    return campaign_status_code(value) in {4, 9, 11}


def campaign_type(row: pd.Series) -> str:
    payment = norm_text(row.get("payment_type"))
    search_bid = as_float(row.get("search_bid"), 0.0)
    reco_bid = as_float(row.get("reco_bid"), 0.0)
    if "cpc" in payment:
        return "CPC"
    if search_bid > 0 and reco_bid > 0:
        return "CPM ПОИСК + ПОЛКИ"
    if search_bid > 0:
        return "CPM ПОИСК"
    if reco_bid > 0:
        return "CPM ПОЛКИ"
    return "CPM"


def primary_campaign_bid(row: pd.Series) -> float:
    typ = str(row.get("campaign_type") or "")
    search_bid = as_float(row.get("search_bid"))
    reco_bid = as_float(row.get("reco_bid"))
    if typ == "CPM ПОЛКИ" and pd.notna(reco_bid) and reco_bid > 0:
        return reco_bid
    if pd.notna(search_bid) and search_bid > 0:
        return search_bid
    return reco_bid


def economic_bid_cap(campaign_type_value: str, impressions: float, clicks: float, orders: float, order_sum: float, target_drr_pct: float = TARGET_DRR_PCT) -> Dict[str, Any]:
    target = target_drr_pct / 100.0
    ctr = safe_div(clicks, impressions)
    cr = safe_div(orders, clicks)
    avg_order = safe_div(order_sum, orders)
    if campaign_type_value.startswith("CPC"):
        direct = safe_div(order_sum * target, clicks)
        funnel = avg_order * cr * target if pd.notna(avg_order) and pd.notna(cr) else np.nan
        return {"max_bid": direct, "max_bid_funnel": funnel, "unit": "₽/клик", "ctr": ctr * 100 if pd.notna(ctr) else np.nan, "cr": cr * 100 if pd.notna(cr) else np.nan, "avg_ad_order_value": avg_order}
    direct = safe_div(order_sum * target * 1000.0, impressions)
    funnel = avg_order * ctr * cr * target * 1000.0 if pd.notna(avg_order) and pd.notna(ctr) and pd.notna(cr) else np.nan
    return {"max_bid": direct, "max_bid_funnel": funnel, "unit": "₽/1000 показов", "ctr": ctr * 100 if pd.notna(ctr) else np.nan, "cr": cr * 100 if pd.notna(cr) else np.nan, "avg_ad_order_value": avg_order}


def recommendation_next_raise(row: pd.Series) -> Tuple[Optional[float], str]:
    current = as_float(row.get("bid_current"))
    max_bid = as_float(row.get("max_bid_live"))
    if pd.isna(current) or current <= 0:
        return None, "Нет текущей ставки"
    typ = str(row.get("campaign_type") or "")
    last_attempt = int(row.get("last_raise_attempt_no") or 0)
    if last_attempt >= MAX_RAISE_ATTEMPTS:
        return None, f"Лимит {MAX_RAISE_ATTEMPTS} попытки повышения уже достигнут"
    next_attempt = last_attempt + 1
    if typ.startswith("CPM"):
        step = CPM_RAISE_STEPS[next_attempt - 1]
    else:
        step = CPC_RAISE_STEP
    target = current + step
    if pd.notna(max_bid) and max_bid > 0:
        target = min(target, max_bid)
    if target <= current + 1e-9:
        return None, "Экономический предел не позволяет следующий шаг"
    return round(float(target), 2), f"попытка {next_attempt}/{MAX_RAISE_ATTEMPTS}, шаг +{step:g} ₽"


# -----------------------------------------------------------------------------
# Market WB
# -----------------------------------------------------------------------------
def market_subject_daily(market: pd.DataFrame) -> pd.DataFrame:
    if market is None or market.empty:
        return pd.DataFrame(columns=["day", "subject", "market_queries", "market_card_clicks", "market_orders"])
    return market.groupby(["day", "subject"], as_index=False).agg(
        market_queries=("query_count", "sum"),
        market_card_clicks=("card_clicks", "sum"),
        market_orders=("orders", "sum"),
    )


def market_query_period(market: pd.DataFrame, start: date, end: date) -> pd.DataFrame:
    m = market[(market["day"] >= start) & (market["day"] <= end)].copy() if market is not None and not market.empty else pd.DataFrame()
    if m.empty:
        return pd.DataFrame(columns=["query", "market_queries", "market_card_clicks", "market_orders"])
    return m.groupby("query", as_index=False).agg(market_queries=("query_count", "mean"), market_card_clicks=("card_clicks", "mean"), market_orders=("orders", "mean"))


def attach_market_to_query_stats(q: pd.DataFrame, market: pd.DataFrame, cur_day: date, base_start: date, base_end: date) -> pd.DataFrame:
    if q is None or q.empty:
        return q
    out = q.copy()
    cur = market_query_period(market, cur_day, cur_day).rename(columns={"market_queries": "market_queries_current", "market_card_clicks": "market_card_clicks_current", "market_orders": "market_orders_current"})
    base = market_query_period(market, base_start, base_end).rename(columns={"market_queries": "market_queries_base", "market_card_clicks": "market_card_clicks_base", "market_orders": "market_orders_base"})
    out = out.merge(cur, on="query", how="left").merge(base, on="query", how="left")
    out["market_growth_pct"] = [pct_delta(a, b) for a, b in zip(out["market_queries_current"], out["market_queries_base"])]
    return out


# -----------------------------------------------------------------------------
# Таблицы: предметы, артикулы, кампании
# -----------------------------------------------------------------------------
def build_subject_article_tables(orders: pd.DataFrame, ads: pd.DataFrame, w: PeriodWindow, as_of: date) -> Tuple[pd.DataFrame, pd.DataFrame]:
    sales_sub = aggregate_metrics(orders.rename(columns={"finished_price": "sales_sum"}), ["subject"], ["sales_sum"], w)
    ads_sub = aggregate_metrics(ads, ["subject"], ["impressions", "clicks", "orders", "spend", "order_sum"], w)
    subjects = sales_sub.merge(ads_sub, on="subject", how="outer")
    subjects = add_ad_derived(subjects)
    subjects["overall_drr_current"] = np.where(subjects["sales_sum_current"] > 0, subjects["spend_current"] / subjects["sales_sum_current"] * 100.0, np.nan)
    subjects["overall_drr_base"] = np.where(subjects["sales_sum_base"] > 0, subjects["spend_base"] / subjects["sales_sum_base"] * 100.0, np.nan)
    subjects = add_deltas(subjects, ["sales_sum","impressions","clicks","order_sum","spend","ctr","cr","cpc","cpm","cpo","drr","overall_drr"])

    keys = ["subject", "supplier_article", "nm_id"]
    sales_art = aggregate_metrics(orders.rename(columns={"finished_price": "sales_sum"}), keys, ["sales_sum"], w)
    ads_art = aggregate_metrics(ads, keys, ["impressions", "clicks", "orders", "spend", "order_sum"], w)
    articles = sales_art.merge(ads_art, on=keys, how="outer")
    articles = add_ad_derived(articles)
    articles["overall_drr_current"] = np.where(articles["sales_sum_current"] > 0, articles["spend_current"] / articles["sales_sum_current"] * 100.0, np.nan)
    articles["overall_drr_base"] = np.where(articles["sales_sum_base"] > 0, articles["spend_base"] / articles["sales_sum_base"] * 100.0, np.nan)
    articles = add_deltas(articles, ["sales_sum","impressions","clicks","order_sum","spend","ctr","cr","cpc","cpm","cpo","drr","overall_drr"])
    last30 = date_filter(orders, as_of - timedelta(days=30), as_of - timedelta(days=1))
    if not last30.empty:
        rank = last30.groupby(["supplier_article", "nm_id"], as_index=False).agg(orders_count_30d=("finished_price", "size"), order_sum_30d=("finished_price", "sum"))
        articles = articles.merge(rank, on=["supplier_article", "nm_id"], how="left")
    else:
        articles["orders_count_30d"] = 0
        articles["order_sum_30d"] = 0
    return subjects.sort_values("sales_sum_current", ascending=False, na_position="last"), articles.sort_values(["subject","orders_count_30d","order_sum_30d"], ascending=[True,False,False], na_position="last")


def _campaign_period_metrics(ads: pd.DataFrame, w: PeriodWindow, suffix: str) -> pd.DataFrame:
    keys = ["campaign_id", "nm_id", "supplier_article", "subject"]
    out = aggregate_metrics(ads, keys, ["impressions","clicks","orders","spend","order_sum"], w)
    out = add_ad_derived(out)
    # Переименовываем current/base поля в явные window-имена.
    ren = {}
    for name in ["impressions","clicks","orders","spend","order_sum","ctr","cr","cpc","cpm","cpo","drr"]:
        ren[name+"_current"] = f"{suffix}_{name}"
        ren[name+"_base"] = f"{suffix}_base_{name}"
    return out.rename(columns=ren)


def build_campaign_table(ads: pd.DataFrame, campaigns: pd.DataFrame, windows: Dict[str, PeriodWindow], wb_min: Dict[Tuple[int,int], float], budgets: Dict[int, Dict[str,Any]], market: pd.DataFrame, state: StateDB) -> pd.DataFrame:
    live = _campaign_period_metrics(ads, windows["live7"], "live7")
    one = _campaign_period_metrics(ads, windows["1d"], "yday")
    mature = _campaign_period_metrics(ads, windows["mature7"], "mature7")
    out = live.merge(one, on=["campaign_id","nm_id","supplier_article","subject"], how="outer").merge(mature, on=["campaign_id","nm_id","supplier_article","subject"], how="outer")
    if campaigns is not None and not campaigns.empty:
        cm = campaigns.copy()
        cm["campaign_type"] = cm.apply(campaign_type, axis=1)
        cm["bid_current"] = cm.apply(primary_campaign_bid, axis=1)
        keep = ["campaign_id","nm_id","campaign_name","campaign_status","payment_type","bid_type","search_bid","reco_bid","campaign_type","bid_current"]
        out = out.merge(cm[keep].drop_duplicates(["campaign_id","nm_id"]), on=["campaign_id","nm_id"], how="left")
    else:
        out["campaign_type"] = ""
        out["bid_current"] = np.nan
        out["campaign_status"] = ""
        out["campaign_name"] = ""

    out["min_bid"] = [wb_min.get((int(cid), int(nm))) if pd.notna(nm) else np.nan for cid, nm in zip(out["campaign_id"], out["nm_id"])]
    out["budget_total"] = [as_float((budgets.get(int(cid)) or {}).get("total")) for cid in out["campaign_id"]]

    # Экономический предел: primary = LIVE7, mature = контроль из зрелого окна.
    caps_live, caps_mature = [], []
    for _, r in out.iterrows():
        caps_live.append(economic_bid_cap(str(r.get("campaign_type") or ""), r.get("live7_impressions"), r.get("live7_clicks"), r.get("live7_orders"), r.get("live7_order_sum")))
        caps_mature.append(economic_bid_cap(str(r.get("campaign_type") or ""), r.get("mature7_impressions"), r.get("mature7_clicks"), r.get("mature7_orders"), r.get("mature7_order_sum")))
    out["max_bid_live"] = [x["max_bid"] for x in caps_live]
    out["max_bid_funnel_live"] = [x["max_bid_funnel"] for x in caps_live]
    out["max_bid_mature"] = [x["max_bid"] for x in caps_mature]
    out["max_bid_funnel_mature"] = [x["max_bid_funnel"] for x in caps_mature]
    out["max_bid_unit"] = [x["unit"] for x in caps_live]
    out["avg_ad_order_value_live"] = [x["avg_ad_order_value"] for x in caps_live]
    out["avg_ad_order_value_mature"] = [x["avg_ad_order_value"] for x in caps_mature]
    out["bid_headroom_pct"] = [pct_delta(m, b) if pd.notna(m) and pd.notna(b) and b > 0 else np.nan for m, b in zip(out["max_bid_live"], out["bid_current"])]

    # Рынок: для каждого предмета вчера vs предыдущий день/7д average.
    subject_daily = market_subject_daily(market)
    yday = windows["1d"].current_to
    base_start, base_end = windows["1d"].base_from, windows["1d"].base_to
    if not subject_daily.empty:
        curm = subject_daily[subject_daily["day"].eq(yday)].groupby("subject", as_index=False).agg(market_queries_current=("market_queries","sum"), market_card_clicks_current=("market_card_clicks","sum"))
        basem = subject_daily[(subject_daily["day"]>=base_start)&(subject_daily["day"]<=base_end)].groupby("subject", as_index=False).agg(market_queries_base=("market_queries","mean"), market_card_clicks_base=("market_card_clicks","mean"))
        out = out.merge(curm, on="subject", how="left").merge(basem, on="subject", how="left")
    else:
        for c in ["market_queries_current","market_card_clicks_current","market_queries_base","market_card_clicks_base"]:
            out[c] = np.nan
    out["market_queries_delta_pct"] = [pct_delta(a,b) for a,b in zip(out["market_queries_current"],out["market_queries_base"])]
    out["traffic_clicks_delta_pct"] = [pct_delta(a,b) for a,b in zip(out["yday_clicks"],out["yday_base_clicks"])]
    out["traffic_impressions_delta_pct"] = [pct_delta(a,b) for a,b in zip(out["yday_impressions"],out["yday_base_impressions"])]
    out["traffic_clicks_vs_market_pct"] = [((1+a/100)/(1+m/100)-1)*100 if pd.notna(a) and pd.notna(m) and m>-100 else np.nan for a,m in zip(out["traffic_clicks_delta_pct"],out["market_queries_delta_pct"])]
    out["traffic_impressions_vs_market_pct"] = [((1+a/100)/(1+m/100)-1)*100 if pd.notna(a) and pd.notna(m) and m>-100 else np.nan for a,m in zip(out["traffic_impressions_delta_pct"],out["market_queries_delta_pct"])]

    # Полные динамики LIVE7/MATURE7 для объяснимого решения.
    for pref in ["live7", "mature7"]:
        for metric in ["impressions","clicks","orders","spend","order_sum","ctr","cr","cpc","cpm","cpo","drr"]:
            cur_col, base_col = f"{pref}_{metric}", f"{pref}_base_{metric}"
            if cur_col in out.columns and base_col in out.columns:
                out[f"{pref}_{metric}_delta_pct"] = [pct_delta(a,b) for a,b in zip(out[cur_col], out[base_col])]
        if f"{pref}_drr" in out.columns and f"{pref}_base_drr" in out.columns:
            out[f"{pref}_drr_delta_pp"] = pd.to_numeric(out[f"{pref}_drr"], errors="coerce") - pd.to_numeric(out[f"{pref}_base_drr"], errors="coerce")

    # Рынок на зрелом финансовом окне: средний дневной спрос текущих 7Д vs предыдущих 7Д.
    if not subject_daily.empty:
        mw = windows["mature7"]
        mcur = subject_daily[(subject_daily["day"]>=mw.current_from)&(subject_daily["day"]<=mw.current_to)].groupby("subject", as_index=False).agg(mature_market_queries_current=("market_queries","mean"))
        mbase = subject_daily[(subject_daily["day"]>=mw.base_from)&(subject_daily["day"]<=mw.base_to)].groupby("subject", as_index=False).agg(mature_market_queries_base=("market_queries","mean"))
        out = out.merge(mcur, on="subject", how="left").merge(mbase, on="subject", how="left")
    else:
        out["mature_market_queries_current"] = np.nan; out["mature_market_queries_base"] = np.nan
    out["mature_market_queries_delta_pct"] = [pct_delta(a,b) for a,b in zip(out["mature_market_queries_current"], out["mature_market_queries_base"])]

    # Компактный контроль бюджета: сколько дней хватит текущего бюджета при расходе последних 7Д.
    out["avg_daily_spend_7d"] = pd.to_numeric(out["live7_spend"], errors="coerce") / 7.0
    out["budget_runway_days"] = [safe_div(b,a) for b,a in zip(out["budget_total"], out["avg_daily_spend_7d"])]
    out["suggested_topup_to_3d"] = [max(0.0, 3.0*a-b) if pd.notna(a) and pd.notna(b) else np.nan for a,b in zip(out["avg_daily_spend_7d"], out["budget_total"])]
    out["budget_status"] = ["LOW" if pd.notna(x) and x < 2 else "OK" if pd.notna(x) else "UNKNOWN" for x in out["budget_runway_days"]]

    # История последних попыток.
    last_attempts, last_change_days, last_change_dirs = [], [], []
    for cid in out["campaign_id"]:
        last = state.last_change(int(cid))
        last_attempts.append(int(last.get("attempt_no") or 0) if last and last.get("direction") == "raise" else 0)
        last_change_days.append(last.get("effective_day") if last else None)
        last_change_dirs.append(last.get("direction") if last else None)
    out["last_raise_attempt_no"] = last_attempts
    out["last_bid_change_day"] = last_change_days
    out["last_bid_change_direction"] = last_change_dirs
    return out


# -----------------------------------------------------------------------------
# CPM query analytics
# -----------------------------------------------------------------------------
def build_search_query_table(campaigns: pd.DataFrame, query_bids: pd.DataFrame, q_yday: pd.DataFrame, q_base7: pd.DataFrame, q_live7: pd.DataFrame, market: pd.DataFrame, windows: Dict[str, PeriodWindow]) -> pd.DataFrame:
    # Статистика кластеров доступна для CPC и поисковой части CPM. Для CPC WB может не отдавать показы/CTR/CPM.
    ctype = campaigns["campaign_type"].astype(str)
    search = campaigns[ctype.eq("CPC") | (ctype.str.startswith("CPM") & ctype.str.contains("ПОИСК", na=False))].copy()
    if search.empty:
        return pd.DataFrame()
    pairs = search[["campaign_id","nm_id","supplier_article","subject","campaign_type","bid_current"]].drop_duplicates(["campaign_id","nm_id"])
    pairs["payment_model"] = np.where(pairs["campaign_type"].astype(str).eq("CPC"), "cpc", "cpm")

    # Union queries from stats/bids.
    parts = []
    for df in [query_bids, q_yday, q_base7, q_live7]:
        if df is not None and not df.empty:
            parts.append(df[["campaign_id","nm_id","query"]].copy())
    if not parts:
        return pd.DataFrame()
    q = pd.concat(parts, ignore_index=True).drop_duplicates(["campaign_id","nm_id","query"])
    q = q.merge(pairs, on=["campaign_id","nm_id"], how="inner")
    if not query_bids.empty:
        q = q.merge(query_bids[["campaign_id","nm_id","query","query_bid","bid_source"]], on=["campaign_id","nm_id","query"], how="left")
    else:
        q["query_bid"] = np.nan
        q["bid_source"] = ""

    def merge_stats(base: pd.DataFrame, df: pd.DataFrame, suffix: str) -> pd.DataFrame:
        if df is None or df.empty:
            for c in ["impressions","clicks","orders","spend","ctr","cpc","cpm","position","visibility"]:
                base[f"{c}_{suffix}"] = np.nan
            return base
        d = df.copy().rename(columns={c: f"{c}_{suffix}" for c in ["impressions","clicks","orders","spend","ctr","cpc","cpm","position","visibility"] if c in df.columns})
        cols = ["campaign_id","nm_id","query"] + [c for c in d.columns if c.endswith("_"+suffix)]
        return base.merge(d[cols], on=["campaign_id","nm_id","query"], how="left")

    q = merge_stats(q, q_yday, "yday")
    q = merge_stats(q, q_base7, "base7")
    q = merge_stats(q, q_live7, "live7")
    # baseline per day from 7-day aggregate.
    for c in ["impressions","clicks","orders","spend"]:
        if f"{c}_base7" in q.columns:
            q[f"{c}_base_daily"] = q[f"{c}_base7"] / 7.0
    q["impression_share_pct"] = np.where(q.groupby("campaign_id")["impressions_yday"].transform("sum") > 0, q["impressions_yday"] / q.groupby("campaign_id")["impressions_yday"].transform("sum") * 100.0, np.nan)
    q["click_share_pct"] = np.where(q.groupby("campaign_id")["clicks_yday"].transform("sum") > 0, q["clicks_yday"] / q.groupby("campaign_id")["clicks_yday"].transform("sum") * 100.0, np.nan)
    q["impressions_delta_pct"] = [pct_delta(a,b) for a,b in zip(q["impressions_yday"],q["impressions_base_daily"])]
    q["clicks_delta_pct"] = [pct_delta(a,b) for a,b in zip(q["clicks_yday"],q["clicks_base_daily"])]
    q = attach_market_to_query_stats(q, market, windows["1d"].current_to, windows["1d"].base_from, windows["1d"].base_to)
    q["clicks_vs_market_pct"] = [((1+a/100)/(1+m/100)-1)*100 if pd.notna(a) and pd.notna(m) and m>-100 else np.nan for a,m in zip(q["clicks_delta_pct"],q["market_growth_pct"])]
    q["impressions_vs_market_pct"] = [((1+a/100)/(1+m/100)-1)*100 if pd.notna(a) and pd.notna(m) and m>-100 else np.nan for a,m in zip(q["impressions_delta_pct"],q["market_growth_pct"])]
    q["position_delta"] = q.get("position_yday", pd.Series(np.nan, index=q.index)) - q.get("position_base7", pd.Series(np.nan, index=q.index))
    stat_cols = [c for c in ["impressions_yday","clicks_yday","position_yday","orders_yday","spend_yday"] if c in q.columns]
    q["query_stats_available"] = q[stat_cols].notna().any(axis=1) if stat_cols else False
    # У CPC отсутствующие impressions — это «нет поля API», а не ноль.
    return q.sort_values(["campaign_id","clicks_yday","impressions_yday","query"], ascending=[True,False,False,True], na_position="last")


def campaign_position_summary(campaign_table: pd.DataFrame, queries: pd.DataFrame) -> pd.DataFrame:
    out = campaign_table.copy()
    if queries is None or queries.empty:
        out["avg_position_7d"] = np.nan
        out["avg_visibility_7d"] = np.nan
        return out
    rows = []
    for cid, g in queries.groupby("campaign_id"):
        pos = pd.to_numeric(g.get("position_live7"), errors="coerce")
        vis = pd.to_numeric(g.get("visibility_live7"), errors="coerce")
        w = pd.to_numeric(g.get("impressions_live7"), errors="coerce").fillna(0.0)
        def wavg(vals: pd.Series) -> float:
            mask = vals.notna()
            if not mask.any():
                return np.nan
            if w[mask].sum() > 0:
                return float((vals[mask] * w[mask]).sum() / w[mask].sum())
            return float(vals[mask].mean())
        rows.append({"campaign_id": int(cid), "avg_position_7d": wavg(pos), "avg_visibility_7d": wavg(vis)})
    return out.merge(pd.DataFrame(rows), on="campaign_id", how="left")


def attach_query_event_postchecks(queries: pd.DataFrame, events: pd.DataFrame, state: StateDB, as_of: date) -> pd.DataFrame:
    """Оценивает каждый поисковый кластер после последнего изменения ставки РК.

    Сравнение идёт с замороженным 7-дневным baseline события, а не с плавающим окном.
    Поэтому D+1/D+2/D+3 можно корректно сопоставлять с состоянием ДО изменения ставки.
    """
    if queries is None or queries.empty:
        return queries
    q = queries.copy()
    default_cols = [
        "event_id","event_day","event_direction","postcheck_day","event_baseline_impressions","event_baseline_clicks",
        "event_baseline_position","event_baseline_market_queries","event_impressions_delta_pct","event_clicks_delta_pct",
        "event_market_delta_pct","event_impressions_vs_market_pct","event_clicks_vs_market_pct","event_position_improvement",
        "query_postcheck_verdict","query_postcheck_reason",
    ]
    for c in default_cols:
        q[c] = None if c in {"event_id","event_day","event_direction","query_postcheck_verdict","query_postcheck_reason"} else np.nan
    if events is None or events.empty:
        return q
    last = events.sort_values(["effective_day","detected_on"]).drop_duplicates("campaign_id", keep="last")
    event_map = {int(r.campaign_id): r._asdict() for r in last.itertuples(index=False)}
    baseline_cache: Dict[str, pd.DataFrame] = {}
    stats_day = as_of - timedelta(days=1)
    for idx, r in q.iterrows():
        cid = int(r["campaign_id"]); ev = event_map.get(cid)
        if not ev:
            continue
        event_day = pd.to_datetime(ev.get("effective_day"), errors="coerce")
        if pd.isna(event_day):
            continue
        event_date = event_day.date()
        days_after = (stats_day - event_date).days
        if days_after < 0 or days_after >= POSTCHECK_DAYS:
            continue
        eid = str(ev.get("event_id") or "")
        if not eid:
            continue
        if eid not in baseline_cache:
            b = state.event_query_baselines(eid)
            baseline_cache[eid] = b.set_index("query") if not b.empty else b
        bdf = baseline_cache[eid]
        qn = query_norm(r.get("query"))
        if bdf is None or bdf.empty or qn not in bdf.index:
            q.at[idx,"event_id"] = eid; q.at[idx,"event_day"] = event_date.isoformat(); q.at[idx,"event_direction"] = ev.get("direction"); q.at[idx,"postcheck_day"] = days_after + 1
            q.at[idx,"query_postcheck_verdict"] = "NEW_QUERY_AFTER_CHANGE"
            q.at[idx,"query_postcheck_reason"] = "Кластер появился после изменения ставки или отсутствовал в baseline."
            continue
        br = bdf.loc[qn]
        if isinstance(br, pd.DataFrame): br = br.iloc[-1]
        base_imp = as_float(br.get("impressions_avg")); base_clk = as_float(br.get("clicks_avg")); base_pos = as_float(br.get("position")); base_market = as_float(br.get("market_queries_avg"))
        cur_imp = as_float(r.get("impressions_yday")); cur_clk = as_float(r.get("clicks_yday")); cur_pos = as_float(r.get("position_yday")); cur_market = as_float(r.get("market_queries_current"))
        imp_delta = pct_delta(cur_imp, base_imp); clk_delta = pct_delta(cur_clk, base_clk); market_delta = pct_delta(cur_market, base_market)
        imp_adj = ((1+imp_delta/100)/(1+market_delta/100)-1)*100 if pd.notna(imp_delta) and pd.notna(market_delta) and market_delta>-100 else imp_delta
        clk_adj = ((1+clk_delta/100)/(1+market_delta/100)-1)*100 if pd.notna(clk_delta) and pd.notna(market_delta) and market_delta>-100 else clk_delta
        pos_improve = base_pos-cur_pos if pd.notna(base_pos) and pd.notna(cur_pos) else np.nan
        direction = str(ev.get("direction") or "")
        if direction == "raise":
            if (pd.notna(clk_adj) and clk_adj >= TRAFFIC_GOOD_PCT) or (pd.notna(imp_adj) and imp_adj >= TRAFFIC_GOOD_PCT) or (pd.notna(pos_improve) and pos_improve >= 1):
                verdict = "QUERY_RAISE_GOOD"
            elif (pd.isna(clk_adj) or clk_adj < TRAFFIC_WEAK_PCT) and (pd.isna(imp_adj) or imp_adj < TRAFFIC_WEAK_PCT) and (pd.isna(pos_improve) or pos_improve < 1):
                verdict = "QUERY_RAISE_NO_EFFECT"
            else:
                verdict = "QUERY_RAISE_WEAK"
        else:
            if (pd.isna(clk_adj) or clk_adj >= -TRAFFIC_WEAK_PCT) and (pd.isna(imp_adj) or imp_adj >= -TRAFFIC_WEAK_PCT): verdict = "QUERY_LOWER_GOOD"
            elif (pd.notna(clk_adj) and clk_adj <= -15) or (pd.notna(imp_adj) and imp_adj <= -20): verdict = "QUERY_LOWER_HARMFUL"
            else: verdict = "QUERY_LOWER_WEAK"
        q.at[idx,"event_id"] = eid; q.at[idx,"event_day"] = event_date.isoformat(); q.at[idx,"event_direction"] = direction; q.at[idx,"postcheck_day"] = days_after+1
        q.at[idx,"event_baseline_impressions"] = base_imp; q.at[idx,"event_baseline_clicks"] = base_clk; q.at[idx,"event_baseline_position"] = base_pos; q.at[idx,"event_baseline_market_queries"] = base_market
        q.at[idx,"event_impressions_delta_pct"] = imp_delta; q.at[idx,"event_clicks_delta_pct"] = clk_delta; q.at[idx,"event_market_delta_pct"] = market_delta
        q.at[idx,"event_impressions_vs_market_pct"] = imp_adj; q.at[idx,"event_clicks_vs_market_pct"] = clk_adj; q.at[idx,"event_position_improvement"] = pos_improve
        q.at[idx,"query_postcheck_verdict"] = verdict
        q.at[idx,"query_postcheck_reason"] = f"D+{days_after+1}: показы vs рынок {imp_adj:.1f}%" if pd.notna(imp_adj) else f"D+{days_after+1}: клики vs рынок {clk_adj:.1f}%" if pd.notna(clk_adj) else f"D+{days_after+1}: данных трафика недостаточно"
    return q


def attach_query_postcheck_summary(campaigns: pd.DataFrame, queries: pd.DataFrame) -> pd.DataFrame:
    out = campaigns.copy()
    if queries is None or queries.empty or "query_postcheck_verdict" not in queries.columns:
        out["query_postcheck_good"] = 0; out["query_postcheck_no_effect"] = 0; out["query_postcheck_harmful"] = 0; out["query_postcheck_total"] = 0
        return out
    rows=[]
    for cid,g in queries.groupby("campaign_id"):
        v=g["query_postcheck_verdict"].fillna("").astype(str)
        rows.append({"campaign_id":int(cid),"query_postcheck_good":int(v.str.contains("GOOD").sum()),"query_postcheck_no_effect":int(v.str.contains("NO_EFFECT|WEAK").sum()),"query_postcheck_harmful":int(v.str.contains("HARMFUL").sum()),"query_postcheck_total":int(v.ne("").sum())})
    return out.merge(pd.DataFrame(rows), on="campaign_id", how="left").fillna({"query_postcheck_good":0,"query_postcheck_no_effect":0,"query_postcheck_harmful":0,"query_postcheck_total":0})


# -----------------------------------------------------------------------------
# Вердикты post-check и рекомендации
# -----------------------------------------------------------------------------
def _traffic_verdict(direction: str, impressions_delta: float, clicks_delta: float, market_delta: float, avg_position: float) -> Tuple[str, str]:
    imp_adj = ((1+impressions_delta/100)/(1+market_delta/100)-1)*100 if pd.notna(impressions_delta) and pd.notna(market_delta) and market_delta>-100 else impressions_delta
    click_adj = ((1+clicks_delta/100)/(1+market_delta/100)-1)*100 if pd.notna(clicks_delta) and pd.notna(market_delta) and market_delta>-100 else clicks_delta
    pos_note = "позиция неизвестна"
    if pd.notna(avg_position):
        pos_note = "позиция TOP-5" if avg_position <= 5 else "позиция TOP-10" if avg_position <= 10 else f"средняя позиция {avg_position:.1f}"
    if direction == "raise":
        if (pd.notna(click_adj) and click_adj >= TRAFFIC_GOOD_PCT) or (pd.notna(imp_adj) and imp_adj >= TRAFFIC_GOOD_PCT):
            return "TRAFFIC_RAISE_GOOD", f"Повышение дало прирост трафика выше рынка: показы {imp_adj:.1f}% / клики {click_adj:.1f}%; {pos_note}. Ставку держать, экономику оценить после дозревания заказов."
        if (pd.isna(click_adj) or click_adj < TRAFFIC_WEAK_PCT) and (pd.isna(imp_adj) or imp_adj < TRAFFIC_WEAK_PCT):
            return "TRAFFIC_RAISE_NO_EFFECT", f"Заметного эффекта повышения нет: показы {imp_adj if pd.notna(imp_adj) else float('nan'):.1f}% / клики {click_adj if pd.notna(click_adj) else float('nan'):.1f}% относительно рынка; {pos_note}."
        return "TRAFFIC_RAISE_WEAK", f"Эффект повышения слабый/неоднозначный: показы {imp_adj:.1f}% / клики {click_adj:.1f}%; {pos_note}."
    # lower
    if (pd.notna(click_adj) and click_adj >= -TRAFFIC_WEAK_PCT) and (pd.notna(imp_adj) and imp_adj >= -TRAFFIC_WEAK_PCT):
        return "TRAFFIC_LOWER_GOOD", f"После снижения трафик почти сохранён относительно рынка: показы {imp_adj:.1f}% / клики {click_adj:.1f}%. Снижение выглядит удачным."
    if (pd.notna(click_adj) and click_adj <= -15) or (pd.notna(imp_adj) and imp_adj <= -20):
        return "TRAFFIC_LOWER_HARMFUL", f"После снижения потеря трафика заметно выше рынка: показы {imp_adj:.1f}% / клики {click_adj:.1f}%. Возможно, ставку снизили слишком сильно."
    return "TRAFFIC_LOWER_WEAK", f"После снижения есть умеренная потеря трафика: показы {imp_adj:.1f}% / клики {click_adj:.1f}%. Нужна зрелая экономика."


def attach_period_bid_activity(campaigns: pd.DataFrame, state: StateDB, windows: Dict[str, PeriodWindow]) -> pd.DataFrame:
    out = campaigns.copy()
    events = state.changes_df()
    for c in ["mature_bid_changes_current","mature_bid_changes_base"]:
        out[c] = 0
    out["mature_bid_changes_summary_current"] = ""
    out["mature_bid_changes_summary_base"] = ""
    if events.empty or out.empty:
        return out
    events["event_day"] = pd.to_datetime(events["effective_day"], errors="coerce").dt.date
    mw = windows["mature7"]
    for idx,r in out.iterrows():
        cid=int(r["campaign_id"]); e=events[events["campaign_id"].eq(cid)].copy()
        cur=e[(e["event_day"]>=mw.current_from)&(e["event_day"]<=mw.current_to)]
        base=e[(e["event_day"]>=mw.base_from)&(e["event_day"]<=mw.base_to)]
        def summary(part: pd.DataFrame) -> str:
            if part.empty: return ""
            p=part.sort_values("event_day")
            return "; ".join(f"{float(x.old_bid):g}→{float(x.new_bid):g} {'↑' if str(x.direction)=='raise' else '↓'} {x.event_day}" for x in p.itertuples())
        out.at[idx,"mature_bid_changes_current"] = len(cur); out.at[idx,"mature_bid_changes_base"] = len(base)
        out.at[idx,"mature_bid_changes_summary_current"] = summary(cur); out.at[idx,"mature_bid_changes_summary_base"] = summary(base)
    return out


def _finance_verdict(direction: str, row: pd.Series) -> Tuple[str, str]:
    """Зрелый финансовый вывод: 7Д без последних 3 дней vs предыдущие 7Д.

    Здесь уже можно учитывать заказы/ДРР, но дополнительно нормализуем продажи на
    изменение спроса WB и показываем CTR/CR + фактическую стоимость трафика.
    """
    spend = pct_delta(row.get("mature7_spend"), row.get("mature7_base_spend"))
    clicks = pct_delta(row.get("mature7_clicks"), row.get("mature7_base_clicks"))
    impressions = pct_delta(row.get("mature7_impressions"), row.get("mature7_base_impressions"))
    orders = pct_delta(row.get("mature7_orders"), row.get("mature7_base_orders"))
    revenue = pct_delta(row.get("mature7_order_sum"), row.get("mature7_base_order_sum"))
    ctr = pct_delta(row.get("mature7_ctr"), row.get("mature7_base_ctr"))
    cr = pct_delta(row.get("mature7_cr"), row.get("mature7_base_cr"))
    cpo = pct_delta(row.get("mature7_cpo"), row.get("mature7_base_cpo"))
    market = as_float(row.get("mature_market_queries_delta_pct"))
    drr_cur, drr_base = as_float(row.get("mature7_drr")), as_float(row.get("mature7_base_drr"))
    drr_pp = drr_cur - drr_base if pd.notna(drr_cur) and pd.notna(drr_base) else np.nan
    typ = str(row.get("campaign_type") or "")
    traffic_cost_name = "CPC" if typ.startswith("CPC") else "CPM"
    traffic_cost = pct_delta(row.get("mature7_cpc" if typ.startswith("CPC") else "mature7_cpm"), row.get("mature7_base_cpc" if typ.startswith("CPC") else "mature7_base_cpm"))

    def vs_market(delta: float) -> float:
        return ((1+delta/100)/(1+market/100)-1)*100 if pd.notna(delta) and pd.notna(market) and market>-100 else delta
    orders_adj, revenue_adj = vs_market(orders), vs_market(revenue)
    def f(v: Any) -> str:
        x=as_float(v); return "н/д" if pd.isna(x) else f"{x:+.1f}%"
    def drr_text() -> str:
        return "н/д" if pd.isna(drr_base) or pd.isna(drr_cur) else f"{drr_base:.1f}%→{drr_cur:.1f}%"
    changes = str(row.get("mature_bid_changes_summary_current") or "").strip() or "ставка в зрелом окне не менялась"
    context = (f"показы {f(impressions)}, клики {f(clicks)}, CTR {f(ctr)}, CR {f(cr)}, {traffic_cost_name} {f(traffic_cost)}, "
               f"расход {f(spend)}, заказы {f(orders)} (к рынку {f(orders_adj)}), выручка {f(revenue)} (к рынку {f(revenue_adj)}), "
               f"CPO {f(cpo)}, ДРР {drr_text()}, рынок WB {f(market)}. Изменения ставки: {changes}.")
    if direction == "raise":
        doubtful = pd.notna(spend) and spend > 10 and (pd.isna(revenue_adj) or revenue_adj < 5) and (pd.isna(orders_adj) or orders_adj < 5) and ((pd.notna(cpo) and cpo > 10) or (pd.notna(drr_pp) and drr_pp > 2))
        good = ((pd.notna(revenue_adj) and revenue_adj >= 10) or (pd.notna(orders_adj) and orders_adj >= 10)) and (pd.isna(drr_pp) or drr_pp <= 2) and (pd.isna(cpo) or cpo <= 10)
        if doubtful:
            return "MATURE_RAISE_DOUBTFUL", "Повышение сомнительное: " + context + " Рост расходов не дал сопоставимого роста зрелых заказов относительно рынка; пробуем снижаться."
        if good:
            return "MATURE_RAISE_GOOD", "Повышение подтверждено зрелой экономикой: " + context
        return "MATURE_RAISE_WAIT", "Зрелая экономика неоднозначна: " + context
    good_lower = pd.notna(spend) and spend < -8 and (pd.isna(orders_adj) or orders_adj >= -5) and (pd.isna(revenue_adj) or revenue_adj >= -5) and ((pd.notna(cpo) and cpo < -5) or (pd.notna(drr_pp) and drr_pp < 0))
    harmful = (pd.notna(orders_adj) and orders_adj < -20) or (pd.notna(revenue_adj) and revenue_adj < -20)
    if good_lower:
        return "MATURE_LOWER_GOOD", "Снижение подтверждено: " + context + " Экономия получена без существенной потери зрелых продаж относительно рынка."
    if harmful:
        return "MATURE_LOWER_HARMFUL", "Снижение оказалось слишком сильным: " + context + " Рассмотреть возврат предыдущей ставки."
    return "MATURE_LOWER_WAIT", "Эффект снижения пока неоднозначен: " + context


def build_recommendations(campaigns: pd.DataFrame) -> pd.DataFrame:
    if campaigns.empty:
        return campaigns
    out = campaigns.copy()
    actions, targets, reasons = [], [], []
    budget_actions=[]
    for _, r in out.iterrows():
        last_verdict = str(r.get("last_traffic_verdict") or "")
        mature = str(r.get("last_mature_verdict") or "")
        current = as_float(r.get("bid_current"))
        max_bid = as_float(r.get("max_bid_live"))
        min_bid = as_float(r.get("min_bid"), 0)
        attempt = int(as_float(r.get("last_raise_attempt_no"), 0) or 0)
        typ = str(r.get("campaign_type") or "")
        avg_pos = as_float(r.get("avg_position_7d"))
        mature_impr = as_float(r.get("mature7_impressions"), 0)
        mature_orders = as_float(r.get("mature7_orders"), 0)
        mature_drr = as_float(r.get("mature7_drr"))
        status_value = r.get("campaign_status")

        if not is_active_campaign_status(status_value):
            budget_actions.append("HOLD_BUDGET")
            actions.append("HOLD_INACTIVE")
            targets.append(np.nan)
            reasons.append(f"РК не активна (статус {status_value}); ставки и паузы не рекомендуем")
            continue

        if str(r.get("budget_status")) == "LOW":
            amount = as_float(r.get("suggested_topup_to_3d"))
            budget_actions.append(f"TOPUP_CANDIDATE {math.ceil(amount):d} ₽" if pd.notna(amount) and amount>0 else "TOPUP_CANDIDATE")
        else:
            budget_actions.append("HOLD_BUDGET")

        # Зрелый 5000+ и 0 заказов — кандидат на паузу, но только рекомендация.
        if mature_impr >= 5000 and mature_orders <= 0:
            actions.append("PAUSE_CANDIDATE"); targets.append(np.nan); reasons.append("Зрелое окно: >=5000 показов и 0 заказов. Пауза только как рекомендация; запись в WB отключена.")
            continue
        if mature == "MATURE_RAISE_DOUBTFUL":
            step = 40 if typ.startswith("CPM") else 1
            target = max(min_bid, current - step) if pd.notna(current) else np.nan
            actions.append("LOWER"); targets.append(target); reasons.append("Зрелая экономика показала, что предыдущее повышение не окупилось")
            continue
        if mature == "MATURE_LOWER_HARMFUL":
            actions.append("RAISE_RESTORE"); targets.append(as_float(r.get("last_event_old_bid"))); reasons.append("Зрелая экономика показала чрезмерную потерю продаж после снижения; вернуть предыдущую ставку")
            continue
        if mature == "MATURE_LOWER_GOOD" or last_verdict == "TRAFFIC_LOWER_GOOD":
            actions.append("HOLD"); targets.append(np.nan); reasons.append("Снижение выглядит удачным; новую ставку пока держим")
            continue
        if last_verdict in {"TRAFFIC_RAISE_NO_EFFECT", "TRAFFIC_RAISE_WEAK"}:
            if attempt >= MAX_RAISE_ATTEMPTS:
                actions.append("STOP_RAISE_TEST"); targets.append(np.nan); reasons.append("Три попытки повышения не дали достаточного прироста трафика; выше не идём")
            else:
                target, reason = recommendation_next_raise(r)
                if target is not None:
                    pos_note = ""
                    if pd.notna(avg_pos) and avg_pos <= 5: pos_note = " Позиция уже TOP-5: ожидаемый эффект ограничен, но тест допустим."
                    elif pd.notna(avg_pos) and avg_pos <= 10: pos_note = " Позиция TOP-10: тестируем умеренно."
                    actions.append("RAISE_TEST"); targets.append(target); reasons.append("Нет достаточного прироста трафика; " + reason + pos_note)
                else:
                    actions.append("HOLD"); targets.append(np.nan); reasons.append(reason)
            continue
        if last_verdict == "TRAFFIC_RAISE_GOOD":
            actions.append("HOLD_WAIT_MATURE"); targets.append(np.nan); reasons.append("Рост трафика подтверждён; ждём зрелые заказы/ДРР")
            continue
        if last_verdict == "TRAFFIC_LOWER_HARMFUL":
            actions.append("RAISE_RESTORE"); targets.append(as_float(r.get("last_event_old_bid"))); reasons.append("Трафик после снижения просел сильнее рынка; рекомендован возврат предыдущей ставки")
            continue
        # Без недавнего эксперимента: тест только при заметном экономическом запасе.
        if attempt >= MAX_RAISE_ATTEMPTS:
            actions.append("STOP_RAISE_TEST"); targets.append(np.nan); reasons.append("Лимит 3 попытки повышения уже достигнут")
        elif pd.notna(current) and pd.notna(max_bid) and max_bid > current * 1.1 and as_float(r.get("live7_drr"), 999) < TARGET_DRR_PCT:
            target, step_reason = recommendation_next_raise(r)
            pos_note = ""
            if pd.notna(avg_pos) and avg_pos <= 5: pos_note = " Позиция уже TOP-5, поэтому сильный прирост маловероятен; тест всё равно разрешён."
            elif pd.notna(avg_pos) and avg_pos <= 10: pos_note = " Средняя позиция TOP-10; повышение тестируем осторожно."
            actions.append("CAN_TEST_RAISE" if target is not None else "HOLD"); targets.append(target if target is not None else np.nan); reasons.append(("Есть экономический запас ставки; " + step_reason + pos_note) if target is not None else step_reason)
        elif pd.notna(mature_drr) and mature_drr >= 20 and pd.notna(current) and current > min_bid:
            step = 40 if typ.startswith("CPM") else 1
            actions.append("LOWER"); targets.append(max(min_bid,current-step)); reasons.append("Зрелый ДРР >=20%; сначала снижаем ставку, а не наращиваем трафик")
        else:
            actions.append("HOLD"); targets.append(np.nan); reasons.append("Нет сильного сигнала для изменения ставки")
    out["recommendation"] = actions
    out["recommended_bid"] = targets
    out["recommendation_reason"] = reasons
    out["budget_recommendation"] = budget_actions
    out["write_enabled"] = False
    return out


def _event_finance_comparison_row(
    before: pd.DataFrame,
    after: pd.DataFrame,
    campaign_type_value: str,
    market: pd.DataFrame,
    subject: str,
    before_from: date,
    before_to: date,
    after_from: date,
    after_to: date,
) -> pd.Series:
    """Готовит event-aligned финансовый срез 7Д после изменения vs 7Д до него."""
    def agg(g: pd.DataFrame) -> Dict[str, float]:
        vals = {x: float(pd.to_numeric(g[x], errors="coerce").fillna(0).sum()) for x in ["impressions","clicks","orders","spend","order_sum"]}
        vals["ctr"] = safe_div(vals["clicks"] * 100.0, vals["impressions"])
        vals["cr"] = safe_div(vals["orders"] * 100.0, vals["clicks"])
        vals["cpc"] = safe_div(vals["spend"], vals["clicks"])
        vals["cpm"] = safe_div(vals["spend"] * 1000.0, vals["impressions"])
        vals["cpo"] = safe_div(vals["spend"], vals["orders"])
        vals["drr"] = safe_div(vals["spend"] * 100.0, vals["order_sum"])
        return vals
    b, a = agg(before), agg(after)
    ms = market_subject_daily(market)
    market_cur = market_base = np.nan
    if not ms.empty:
        sm = ms[ms["subject"].eq(subject)]
        cur_vals = sm[(sm["day"]>=after_from)&(sm["day"]<=after_to)]["market_queries"]
        base_vals = sm[(sm["day"]>=before_from)&(sm["day"]<=before_to)]["market_queries"]
        if not cur_vals.empty: market_cur = float(cur_vals.mean())
        if not base_vals.empty: market_base = float(base_vals.mean())
    return pd.Series({
        "campaign_type": campaign_type_value,
        "mature7_impressions": a["impressions"], "mature7_base_impressions": b["impressions"],
        "mature7_clicks": a["clicks"], "mature7_base_clicks": b["clicks"],
        "mature7_orders": a["orders"], "mature7_base_orders": b["orders"],
        "mature7_spend": a["spend"], "mature7_base_spend": b["spend"],
        "mature7_order_sum": a["order_sum"], "mature7_base_order_sum": b["order_sum"],
        "mature7_ctr": a["ctr"], "mature7_base_ctr": b["ctr"],
        "mature7_cr": a["cr"], "mature7_base_cr": b["cr"],
        "mature7_cpc": a["cpc"], "mature7_base_cpc": b["cpc"],
        "mature7_cpm": a["cpm"], "mature7_base_cpm": b["cpm"],
        "mature7_cpo": a["cpo"], "mature7_base_cpo": b["cpo"],
        "mature7_drr": a["drr"], "mature7_base_drr": b["drr"],
        "mature_market_queries_current": market_cur,
        "mature_market_queries_base": market_base,
        "mature_market_queries_delta_pct": pct_delta(market_cur, market_base),
        "mature_bid_changes_summary_current": "event-aligned post-window",
    })


def evaluate_events(state: StateDB, campaigns: pd.DataFrame, ads: pd.DataFrame, market: pd.DataFrame, as_of: date) -> pd.DataFrame:
    events = state.changes_df()
    if events.empty:
        return events
    cmap = campaigns.set_index("campaign_id").to_dict("index") if not campaigns.empty else {}
    for _, ev in events.iterrows():
        eid = str(ev["event_id"])
        cid = int(ev["campaign_id"])
        effective = pd.to_datetime(ev["effective_day"]).date()
        direction = str(ev["direction"])
        row = cmap.get(cid)
        if not row:
            continue
        # Traffic post-check: среднее фактических дней после изменения vs средний день 7д до него.
        after_end = min(as_of - timedelta(days=1), effective + timedelta(days=POSTCHECK_DAYS-1))
        if after_end >= effective:
            after = date_filter(ads[ads["campaign_id"].eq(cid)], effective, after_end)
            before = date_filter(ads[ads["campaign_id"].eq(cid)], effective - timedelta(days=7), effective - timedelta(days=1))
            if not after.empty and not before.empty:
                after_days = max(1, (after_end - effective).days + 1)
                imp_cur = after["impressions"].sum() / after_days
                clk_cur = after["clicks"].sum() / after_days
                imp_base = before["impressions"].sum() / 7.0
                clk_base = before["clicks"].sum() / 7.0
                imp_delta = pct_delta(imp_cur, imp_base)
                clk_delta = pct_delta(clk_cur, clk_base)
                # Market subject index over same windows.
                subj = str(ev.get("subject") or row.get("subject") or "")
                ms = market_subject_daily(market)
                ma = ms[(ms["subject"].eq(subj)) & (ms["day"]>=effective) & (ms["day"]<=after_end)]["market_queries"].mean() if not ms.empty else np.nan
                mb = ms[(ms["subject"].eq(subj)) & (ms["day"]>=effective-timedelta(days=7)) & (ms["day"]<=effective-timedelta(days=1))]["market_queries"].mean() if not ms.empty else np.nan
                market_delta = pct_delta(ma, mb)
                avg_pos = as_float(row.get("avg_position_7d"))
                verdict, reason = _traffic_verdict(direction, imp_delta, clk_delta, market_delta, avg_pos)
                state.update_event(eid, traffic_verdict=verdict, traffic_reason=reason, last_evaluated_on=as_of.isoformat())
        # Mature verdict конкретного изменения: ждём, пока ВСЕ 7 дней после изменения
        # выйдут из 3-дневного лага. Иначе финансовый вывод относится к периоду ДО ставки.
        days_since = (as_of - effective).days
        post_from, post_to = effective, effective + timedelta(days=MATURE_WINDOW_DAYS-1)
        pre_from, pre_to = effective - timedelta(days=MATURE_WINDOW_DAYS), effective - timedelta(days=1)
        event_ads = ads[ads["campaign_id"].eq(cid)]
        if days_since >= MATURE_WINDOW_DAYS + MATURE_LAG_DAYS:
            before = date_filter(event_ads, pre_from, pre_to)
            after = date_filter(event_ads, post_from, post_to)
            if not before.empty and not after.empty:
                tmp = _event_finance_comparison_row(before, after, str(row.get("campaign_type") or ""), market, str(ev.get("subject") or row.get("subject") or ""), pre_from, pre_to, post_from, post_to)
                verdict, reason = _finance_verdict(direction, tmp)
                state.update_event(eid, mature_verdict=verdict, mature_reason="Зрелый event-window: " + reason, last_evaluated_on=as_of.isoformat())
        # Финальная переоценка D+14: повторно пересчитываем те же 7 дней после изменения.
        # К этому моменту отложенная атрибуция заказов WB могла дополнить исторические строки.
        if days_since >= FINAL_RECHECK_AFTER_DAYS:
            before = date_filter(event_ads, pre_from, pre_to)
            after = date_filter(event_ads, post_from, post_to)
            if not before.empty and not after.empty:
                tmp = _event_finance_comparison_row(before, after, str(row.get("campaign_type") or ""), market, str(ev.get("subject") or row.get("subject") or ""), pre_from, pre_to, post_from, post_to)
                verdict, reason = _finance_verdict(direction, tmp)
                state.update_event(eid, final14_verdict="FINAL14_" + verdict, final14_reason="Переоценка через 14 дней: " + reason, last_evaluated_on=as_of.isoformat())
    return state.changes_df()


def attach_latest_event_to_campaigns(campaigns: pd.DataFrame, events: pd.DataFrame) -> pd.DataFrame:
    if campaigns.empty:
        return campaigns
    out = campaigns.copy()
    if events is None or events.empty:
        for c in ["last_event_id","last_event_day","last_event_old_bid","last_event_new_bid","last_event_direction","last_traffic_verdict","last_traffic_reason","last_mature_verdict","last_mature_reason","last_final14_verdict","last_final14_reason"]:
            out[c] = None
        return out
    e = events.sort_values(["effective_day","detected_on"]).drop_duplicates("campaign_id", keep="last").rename(columns={
        "event_id":"last_event_id", "effective_day":"last_event_day", "old_bid":"last_event_old_bid", "new_bid":"last_event_new_bid", "direction":"last_event_direction",
        "traffic_verdict":"last_traffic_verdict", "traffic_reason":"last_traffic_reason", "mature_verdict":"last_mature_verdict", "mature_reason":"last_mature_reason",
        "final14_verdict":"last_final14_verdict", "final14_reason":"last_final14_reason", "attempt_no":"last_event_attempt_no",
    })
    keep = ["campaign_id"] + [c for c in e.columns if c.startswith("last_")]
    return out.merge(e[keep], on="campaign_id", how="left")


# -----------------------------------------------------------------------------
# Data quality
# -----------------------------------------------------------------------------
def missing_days(df: pd.DataFrame, start: date, end: date) -> List[str]:
    if df is None or df.empty or "day" not in df.columns:
        return [(start + timedelta(days=i)).isoformat() for i in range((end-start).days+1)]
    have = set(df["day"].dropna())
    return [d.isoformat() for d in (start + timedelta(days=i) for i in range((end-start).days+1)) if d not in have]


def source_quality(
    name: str,
    df: pd.DataFrame,
    keys: Sequence[str],
    required: Optional[PeriodWindow] = None,
    note: str = "",
    require_every_day: bool = True,
) -> Dict[str, Any]:
    min_day = min(df["day"]) if df is not None and not df.empty and "day" in df.columns else None
    max_day = max(df["day"]) if df is not None and not df.empty and "day" in df.columns else None
    miss: List[str] = []
    status = "OK" if df is not None and not df.empty else "MISSING"
    if required and df is not None and not df.empty and "day" in df.columns:
        # Для заказов отсутствие строк в конкретный день может означать реальные 0 заказов,
        # а не дырку в выгрузке. Там проверяем только границы покрытия. Для рекламы/рынка
        # по-прежнему требуем ежедневные строки, чтобы не получить ложные -100%.
        if require_every_day:
            miss = missing_days(df, required.current_from, required.current_to)
            if miss:
                status = "INCOMPLETE"
        elif min_day is None or max_day is None or min_day > required.current_from or max_day < required.current_to:
            status = "INCOMPLETE"
    return {"source": name, "status": status, "rows": 0 if df is None else len(df), "min_date": min_day, "max_date": max_day, "missing_required_dates": ", ".join(miss), "note": note, "keys": "\n".join(keys)}


# -----------------------------------------------------------------------------
# Output
# -----------------------------------------------------------------------------
def build_ui_payload(ctx: RunContext, windows: Dict[str, PeriodWindow], subjects: Dict[str,pd.DataFrame], articles: Dict[str,pd.DataFrame], campaigns: pd.DataFrame, queries: pd.DataFrame, events: pd.DataFrame, quality: pd.DataFrame, market: pd.DataFrame, balance: Dict[str, Any]) -> Dict[str, Any]:
    market_last = market_subject_daily(market)
    latest_market_day = max(market_last["day"]) if not market_last.empty else None
    market_latest = market_last[market_last["day"].eq(latest_market_day)] if latest_market_day else pd.DataFrame()
    return {
        "meta": {
            "version": VERSION,
            "store": STORE,
            "generated_at": ctx.now_msk.isoformat(),
            "as_of": ctx.as_of.isoformat(),
            "write_enabled": False,
            "target_drr_pct": TARGET_DRR_PCT,
            "mature_lag_days": MATURE_LAG_DAYS,
            "final_recheck_days": FINAL_RECHECK_AFTER_DAYS,
            "cpm_raise_steps": list(CPM_RAISE_STEPS),
        },
        "windows": {k: {"label":v.label,"current_from":v.current_from.isoformat(),"current_to":v.current_to.isoformat(),"base_from":v.base_from.isoformat(),"base_to":v.base_to.isoformat(),"base_daily_average":v.base_daily_average} for k,v in windows.items()},
        "data_quality": records(quality),
        "promotion_balance": {k: jsonable(v) for k,v in (balance or {}).items()},
        "market_latest": records(market_latest),
        "periods": {
            code: {"subjects":records(subjects[code]), "articles":records(articles[code])}
            for code in ["1d","live7","mature7"]
        },
        "campaigns": records(campaigns),
        "search_queries": records(queries),
        "cpm_queries": records(queries[queries.get("payment_model", pd.Series(index=queries.index, dtype=str)).eq("cpm")]) if not queries.empty else [],
        "postchecks": records(events),
    }


def write_technical_xlsx(path: Path, ctx: RunContext, windows: Dict[str, PeriodWindow], subjects: Dict[str,pd.DataFrame], articles: Dict[str,pd.DataFrame], campaigns: pd.DataFrame, queries: pd.DataFrame, events: pd.DataFrame, quality: pd.DataFrame, api_log: pd.DataFrame, market: pd.DataFrame) -> None:
    # Один диагностический workbook вместо множества служебных Excel.
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        summary = pd.DataFrame([
            ["Версия", VERSION], ["Магазин", STORE], ["Дата расчёта", ctx.as_of.isoformat()], ["Запись в WB", "ОТКЛЮЧЕНА"],
            ["Целевой ДРР", TARGET_DRR_PCT], ["Шаги CPM-теста", " / ".join(map(str, CPM_RAISE_STEPS))],
            ["Зрелое окно", f"{windows['mature7'].current_from}..{windows['mature7'].current_to}"],
            ["Зрелая база", f"{windows['mature7'].base_from}..{windows['mature7'].base_to}"],
        ], columns=["Параметр","Значение"])
        summary.to_excel(writer, sheet_name="00_Сводка", index=False)
        # Сводим предметы и артикулы в один лист с явным уровнем.
        pa = []
        for code in ["1d","live7","mature7"]:
            s = subjects[code].copy(); s.insert(0,"level","subject"); s.insert(1,"period",code); pa.append(s)
            a = articles[code].copy(); a.insert(0,"level","article"); a.insert(1,"period",code); pa.append(a)
        pd.concat(pa, ignore_index=True, sort=False).to_excel(writer, sheet_name="10_Предметы_Артикулы", index=False)
        campaigns.to_excel(writer, sheet_name="20_РК", index=False)
        (queries if not queries.empty else pd.DataFrame({"note":["Статистика поисковых кластеров недоступна/нет поисковых РК"]})).to_excel(writer, sheet_name="30_Поисковые_Запросы", index=False)
        (events if not events.empty else pd.DataFrame({"note":["Изменений ставок пока не зафиксировано"]})).to_excel(writer, sheet_name="40_PostCheck", index=False)
        market_tail = market.sort_values(["day","query"], ascending=[False,True]).head(5000) if market is not None and not market.empty else pd.DataFrame({"note":["Нет данных рынка WB"]})
        market_tail.to_excel(writer, sheet_name="50_Рынок_WB", index=False)
        diag = quality.copy()
        diag.to_excel(writer, sheet_name="90_Диагностика", index=False, startrow=0)
        start = len(diag) + 3
        (api_log if not api_log.empty else pd.DataFrame({"note":["WB API не вызывался/нет лога"]})).to_excel(writer, sheet_name="90_Диагностика", index=False, startrow=start)
        wb = writer.book
        header = wb.add_format({"bold":True,"bg_color":"#0F274F","font_color":"#FFFFFF","border":1})
        for name, ws in writer.sheets.items():
            ws.freeze_panes(1,0)
            ws.set_row(0,24,header)
            ws.set_column(0, min(getattr(ws,"dim_colmax",25),60), 15)
            ws.set_column(0,1,24)
            if name in {"40_PostCheck","90_Диагностика"}:
                ws.set_column(0, min(getattr(ws,"dim_colmax",25),60), 22)


# -----------------------------------------------------------------------------
# Основной run
# -----------------------------------------------------------------------------
def run(as_of: date, out_dir: Path, upload: bool = True, api_reads: bool = True) -> Tuple[Path, Path, Path]:
    now = datetime.now(MOSCOW)
    ctx = RunContext(as_of=as_of, now_msk=now, run_id=str(uuid.uuid4()))
    windows = build_windows(as_of)
    s3 = make_s3()
    bucket = env_required("YC_BUCKET_NAME")

    out_dir.mkdir(parents=True, exist_ok=True)
    state_path = out_dir / "state.sqlite3"
    if upload and s3_exists(s3, bucket, STATE_KEY):
        state_path.write_bytes(s3_download(s3, bucket, STATE_KEY))
    state = StateDB(state_path)
    state.start_run(ctx)

    try:
        # 1. Источники.
        orders, order_keys = load_weekly_source(s3, bucket, ORDERS_PREFIX, 10, normalize_orders)
        ads, campaign_meta, ad_keys = load_ads_and_campaigns(s3, bucket, 10)
        stocks, stock_keys = load_weekly_source(s3, bucket, STOCKS_PREFIX, 3, normalize_stocks)
        market_new, market_keys = load_market_files(s3, bucket, 40)

        orders = dedupe_orders_global(orders)
        pmap = build_product_map(stocks, orders, ads, campaign_meta)
        orders = canonicalize_products(orders, pmap)
        ads = canonicalize_products(ads, pmap)
        campaign_meta = canonicalize_products(campaign_meta, pmap)
        ads = dedupe_ads_global(ads)
        if not campaign_meta.empty:
            campaign_meta["campaign_type"] = campaign_meta.apply(campaign_type, axis=1)
            campaign_meta["bid_current"] = campaign_meta.apply(primary_campaign_bid, axis=1)

        # 2. Рынок WB: фактические даты внутри файла, не дата имени файла.
        state.upsert_market(market_new)
        market = state.market_df(as_of - timedelta(days=60), as_of - timedelta(days=1))

        # 3. Read-only WB API.
        client: Optional[WBPromotionClient] = None
        query_bids = pd.DataFrame(columns=["campaign_id","nm_id","query","query_bid","bid_source"])
        q_yday = pd.DataFrame(); q_base7 = pd.DataFrame(); q_live7 = pd.DataFrame()
        min_bid_map: Dict[Tuple[int,int], float] = {}
        budgets: Dict[int, Dict[str,Any]] = {}
        balance: Dict[str,Any] = {}
        if api_reads and (os.getenv("FINICK_API_WB") or "").strip() and not campaign_meta.empty:
            client = WBPromotionClient(env_required("FINICK_API_WB"), allow_writes=False)
            balance = client.get_balance()
            # Бюджет/min-bid читаем только для РК, которые ещё могут работать: ready/active/paused.
            api_meta = campaign_meta[campaign_meta["campaign_status"].map(is_api_relevant_campaign_status)].drop_duplicates(["campaign_id","nm_id"]).copy()
            for _, c in api_meta.iterrows():
                cid, nm = int(c["campaign_id"]), clean_int(c.get("nm_id"))
                if not nm:
                    continue
                budgets[cid] = client.get_budget(cid)
                typ = str(c.get("campaign_type") or "")
                payment = "cpc" if typ.startswith("CPC") else "cpm"
                placements = ["search"] if typ in {"CPC","CPM ПОИСК"} else ["recommendation"] if typ == "CPM ПОЛКИ" else ["search","combined","recommendation"]
                mb = client.get_min_bids(cid, [nm], payment, placements)
                vals = list((mb.get(nm) or {}).values())
                if vals:
                    min_bid_map[(cid,nm)] = max(vals) if typ == "CPM ПОИСК + ПОЛКИ" else min(vals)
                time.sleep(0.15)
            # Search-cluster stats: CPC + поисковая часть CPM.
            # Индивидуальные ставки кластеров читаем только для CPM: CPC не поддерживает cluster bids.
            # Query stats нужны для моментальной оценки только по активным РК.
            active_meta = campaign_meta[campaign_meta["campaign_status"].map(is_active_campaign_status)].copy()
            ctype = active_meta["campaign_type"].astype(str)
            cpm_search = active_meta[ctype.str.startswith("CPM") & ctype.str.contains("ПОИСК", na=False)]
            cpm_pairs = [(int(r.campaign_id), int(r.nm_id)) for r in cpm_search.itertuples() if not pd.isna(r.nm_id)]
            if cpm_pairs:
                query_bids = client.get_query_bids(cpm_pairs)
            search_stats = active_meta[ctype.eq("CPC") | (ctype.str.startswith("CPM") & ctype.str.contains("ПОИСК", na=False))]
            stats_pairs = [(int(r.campaign_id), int(r.nm_id)) for r in search_stats.itertuples() if not pd.isna(r.nm_id)]
            if stats_pairs:
                q_yday = client.get_query_stats(stats_pairs, windows["1d"].current_from, windows["1d"].current_to)
                q_base7 = client.get_query_stats(stats_pairs, windows["1d"].base_from, windows["1d"].base_to)
                q_live7 = client.get_query_stats(stats_pairs, windows["live7"].current_from, windows["live7"].current_to)

        # 4. UI hierarchy subjects/articles.
        subjects: Dict[str,pd.DataFrame] = {}
        articles: Dict[str,pd.DataFrame] = {}
        for code in ["1d","live7","mature7"]:
            subjects[code], articles[code] = build_subject_article_tables(orders, ads, windows[code], as_of)

        # 5. Campaign analytics.
        campaigns = build_campaign_table(ads, campaign_meta, windows, min_bid_map, budgets, market, state)
        queries = build_search_query_table(campaigns, query_bids, q_yday, q_base7, q_live7, market, windows)
        campaigns = campaign_position_summary(campaigns, queries)

        # 6. Detect bid changes BEFORE inserting today's snapshot.
        new_events: List[str] = []
        for _, r in campaigns.iterrows():
            eid = state.detect_change(r.to_dict(), as_of)
            if eid:
                new_events.append(eid)
                # Baseline изменения = последние 7 завершённых дней ДО обнаружения ставки.
                if not q_live7.empty:
                    base = q_live7[(q_live7["campaign_id"].eq(int(r["campaign_id"]))) & (q_live7["nm_id"].eq(int(r["nm_id"])))]
                    market_base = market[(market["day"]>=windows["live7"].current_from)&(market["day"]<=windows["live7"].current_to)] if not market.empty else pd.DataFrame()
                    state.save_event_query_baseline(eid, base, market_base)

        # 7. Query snapshot for yesterday.
        if not queries.empty:
            state.insert_query_snapshot(as_of, windows["1d"].current_to, queries.rename(columns={
                "impressions_yday":"impressions","clicks_yday":"clicks","spend_yday":"spend","orders_yday":"orders",
                "position_yday":"position","visibility_yday":"visibility","market_queries_current":"market_queries",
            }))

        # 8. Campaign snapshot; после этого состояние готово к следующему дню.
        for _, r in campaigns.iterrows():
            state.insert_campaign_snapshot(r.to_dict(), as_of)
        state.conn.commit()

        # 9. Post-check events and recommendations.
        campaigns = attach_period_bid_activity(campaigns, state, windows)
        events = evaluate_events(state, campaigns, ads, market, as_of)
        campaigns = attach_latest_event_to_campaigns(campaigns, events)
        queries = attach_query_event_postchecks(queries, events, state, as_of)
        campaigns = attach_query_postcheck_summary(campaigns, queries)
        # last_raise_attempt_no должен отражать event, который только что обнаружили.
        if "last_event_attempt_no" in campaigns.columns:
            campaigns["last_raise_attempt_no"] = np.where(campaigns["last_event_direction"].eq("raise"), campaigns["last_event_attempt_no"].fillna(campaigns["last_raise_attempt_no"]), campaigns["last_raise_attempt_no"])
        campaigns = build_recommendations(campaigns)

        # 10. Data Quality.
        quality_rows = [
            source_quality("WB orders", orders, order_keys, windows["mature7"], "finishedPrice; отмены исключены; FBO/FBS дедуп по srid/orderUid. Отсутствие заказов в отдельный день не считается дыркой.", require_every_day=False),
            source_quality("WB ads daily", ads, ad_keys, windows["1d"], "РК: показы/клики/заказы/расход/сумма заказов"),
            source_quality("WB stock product map", stocks, stock_keys, None, "Основной справочник nmID -> артикул -> предмет"),
            source_quality("WB market search demand", market, market_keys, windows["1d"], "Дата берётся из выбранного периода внутри файла; имя файла = дата выгрузки"),
        ]
        unmapped = pmap[(pmap["subject"].fillna("").astype(str).str.strip().eq("")) | (pmap["supplier_article"].fillna("").astype(str).str.strip().eq(""))].copy() if not pmap.empty else pmap.copy()
        unmapped_ids = ", ".join(map(str, unmapped["nm_id"].head(30).tolist())) if not unmapped.empty and "nm_id" in unmapped.columns else ""
        quality_rows.append({
            "source": "Canonical product mapping",
            "status": "OK" if unmapped.empty else "FALLBACK",
            "rows": len(pmap),
            "min_date": None,
            "max_date": None,
            "missing_required_dates": "",
            "note": f"nmID сопоставляется с приоритетом остатки > заказы > список РК > реклама. Неполных nmID: {len(unmapped)}" + (f"; первые: {unmapped_ids}" if unmapped_ids else ""),
            "keys": "runtime canonical map",
        })
        q_status = "OK" if not queries.empty and queries.get("query_stats_available", pd.Series(dtype=bool)).any() else ("MISSING" if queries.empty else "UNAVAILABLE")
        quality_rows.append({"source":"WB search query API","status":q_status,"rows":len(queries),"min_date":None,"max_date":None,"missing_required_dates":"","note":"/adv/v1/normquery/stats: CPC + CPM, с разбивкой по дням. Для CPC views/CTR/CPM могут отсутствовать — это не ноль.","keys":"WB API"})
        quality = pd.DataFrame(quality_rows)

        # API log in one technical file + SQLite, no separate Лог_API.xlsx.
        api_log = pd.DataFrame(client.log) if client else pd.DataFrame()
        if client:
            state.write_api_log(ctx.run_id, client.log)

        # 11. Outputs.
        ui_payload = build_ui_payload(ctx, windows, subjects, articles, campaigns, queries, events, quality, market, balance)
        ui_path = out_dir / "ui_payload.json"
        tech_path = out_dir / "Техническая_аналитика.xlsx"
        ui_path.write_text(json.dumps(ui_payload, ensure_ascii=False, indent=2, default=jsonable), encoding="utf-8")
        write_technical_xlsx(tech_path, ctx, windows, subjects, articles, campaigns, queries, events, quality, api_log, market)

        state.finish_run(ctx, "ok", f"campaigns={len(campaigns)}; queries={len(queries)}; events={len(events)}")
        state.close()

        if upload:
            s3_upload_bytes(s3, bucket, STATE_KEY, state_path.read_bytes(), "application/x-sqlite3")
            s3_upload_bytes(s3, bucket, UI_KEY, ui_path.read_bytes(), "application/json; charset=utf-8")
            s3_upload_bytes(s3, bucket, TECH_XLSX_KEY, tech_path.read_bytes(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        print(f"Готово: {tech_path}")
        print(f"Готово: {ui_path}")
        print(f"Готово: {state_path}")
        print("Запись ставок/пауз/бюджета: ОТКЛЮЧЕНА")
        return tech_path, ui_path, state_path
    except Exception as exc:
        try:
            state.finish_run(ctx, "error", repr(exc))
            state.close()
            if upload and state_path.exists():
                s3_upload_bytes(s3, bucket, STATE_KEY, state_path.read_bytes(), "application/x-sqlite3")
        except Exception:
            pass
        raise


def main(argv: Optional[Sequence[str]] = None) -> int:
    loaded = load_report_env()
    if loaded:
        print("REPORT_ENV: подхвачены секреты: " + ", ".join(sorted(loaded)), flush=True)
    p = argparse.ArgumentParser(description="Ассистент WB FINICK — единая аналитика рекламы, без записи в WB")
    p.add_argument("--as-of", default="", help="Дата расчёта YYYY-MM-DD; по умолчанию сегодня по Москве")
    p.add_argument("--out-dir", default="assistant_output")
    p.add_argument("--no-upload", action="store_true")
    p.add_argument("--no-api", action="store_true", help="Не читать Promotion API; только S3 отчёты")
    args = p.parse_args(list(argv) if argv is not None else None)
    as_of = datetime.strptime(args.as_of, "%Y-%m-%d").date() if args.as_of else datetime.now(MOSCOW).date()
    run(as_of, Path(args.out_dir), upload=not args.no_upload, api_reads=not args.no_api)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
