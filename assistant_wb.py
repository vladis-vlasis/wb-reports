#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Ассистент WB — аналитика рекламы FINICK, расчётный слой V1.

Задача модуля:
- ничего не меняет в рекламе;
- читает уже собранные WB-данные FINICK из Yandex Object Storage;
- строит расчёты для будущего веб-интерфейса:
  Предмет -> Артикул -> Рекламная кампания -> Поисковый запрос;
- поддерживает два режима периода: 1Д и 7Д;
- сегодня никогда не входит в расчёт;
- 1Д = вчера против среднего дня предыдущих 7 дней;
- 7Д = последние 7 завершённых дней против предыдущих 7 завершённых дней;
- один и тот же итоговый файл перезаписывается при каждом запуске (09:00/12:00/16:00).

Источники WB FINICK:
1) Заказы:
   Отчёты/Заказы/FINICK/Недельные/
   finishedPrice = фактическая цена покупателя после СПП.
2) Реклама по РК:
   Отчёты/Реклама/FINICK/Недельные/
3) Карта товаров / предметов:
   Отчёты/Остатки/FINICK/Недельные/
4) Позиции и видимость поисковых запросов:
   Отчёты/Поисковые запросы/FINICK/Недельные/
5) Актуальный список CPM-запросов/ставок (если есть):
   Служебные файлы/Ассистент WB/FINICK/CPM_адаптивное_управление.xlsx

Выход:
- Ассистент_WB_FINICK_Реклама_технический.xlsx
- Ассистент_WB_FINICK_Реклама_ui.json
- загрузка этих же файлов в S3:
  Служебные файлы/Ассистент WB/FINICK/Реклама/
"""
from __future__ import annotations

import argparse
import io
import json
import math
import os
import re
import sys
import tempfile
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple
from zoneinfo import ZoneInfo

import boto3
import numpy as np
import pandas as pd

VERSION = "assistant-wb-finick-ads-analytics-v2-github-2026-08-07"
STORE = "FINICK"
MOSCOW = ZoneInfo("Europe/Moscow")
S3_ENDPOINT = os.getenv("YC_S3_ENDPOINT", "https://storage.yandexcloud.net")

ORDERS_PREFIX = "Отчёты/Заказы/FINICK/Недельные/"
ADS_PREFIX = "Отчёты/Реклама/FINICK/Недельные/"
STOCKS_PREFIX = "Отчёты/Остатки/FINICK/Недельные/"
KEYWORDS_PREFIXES = [
    "Отчёты/Поисковые запросы/FINICK/Недельные/",
    "Отчёты/Позиции по Ключам/FINICK/Недельные/",
]
MANAGER_KEY = "Служебные файлы/Ассистент WB/FINICK/CPM_адаптивное_управление.xlsx"
OUTPUT_PREFIX = "Служебные файлы/Ассистент WB/FINICK/Реклама/"
OUTPUT_CURRENT_PREFIX = OUTPUT_PREFIX + "current/"
OUTPUT_HISTORY_PREFIX = OUTPUT_PREFIX + "history/"
OUTPUT_XLSX_KEY = OUTPUT_CURRENT_PREFIX + "Техническая_аналитика.xlsx"
OUTPUT_JSON_KEY = OUTPUT_CURRENT_PREFIX + "ui_payload.json"
OUTPUT_META_KEY = OUTPUT_CURRENT_PREFIX + "meta.json"

# Динамика: менее 0.5% считаем практически без изменения.
TREND_FLAT_EPS_PCT = 0.5


@dataclass(frozen=True)
class PeriodWindow:
    code: str
    label: str
    current_from: date
    current_to: date
    base_from: date
    base_to: date
    base_is_daily_average: bool


def load_report_env() -> List[str]:
    """Подхватывает недостающие секреты из REPORT_ENV.

    Это повторяет старую схему GitHub Actions, но находится внутри Python,
    поэтому тот же файл позже можно перенести на сервер без переписывания
    бизнес-логики. Значения секретов никогда не печатаются.
    """
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
    value = os.getenv(name, "").strip()
    if not value:
        raise RuntimeError(f"Не задан обязательный env/secret: {name}")
    return value


def make_s3():
    return boto3.client(
        "s3",
        endpoint_url=S3_ENDPOINT,
        aws_access_key_id=env_required("YC_ACCESS_KEY_ID"),
        aws_secret_access_key=env_required("YC_SECRET_ACCESS_KEY"),
        region_name="ru-central1",
    )


def norm_text(v: Any) -> str:
    s = str(v or "").strip().replace("ё", "е").lower()
    return re.sub(r"\s+", " ", s)


def norm_col(v: Any) -> str:
    return re.sub(r"[^a-zа-я0-9]+", "", norm_text(v))


def clean_article(v: Any) -> str:
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return ""
    text = str(v).strip()
    if text.endswith(".0") and re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return text


def clean_int(v: Any) -> Optional[int]:
    try:
        if v is None or (isinstance(v, float) and math.isnan(v)):
            return None
        return int(float(v))
    except Exception:
        return None


def to_num(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.astype(str)
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


def list_keys(s3, bucket: str, prefix: str, limit: int = 20) -> List[str]:
    paginator = s3.get_paginator("list_objects_v2")
    objects: List[Tuple[datetime, str]] = []
    for page in paginator.paginate(Bucket=bucket, Prefix=prefix):
        for obj in page.get("Contents", []):
            key = str(obj.get("Key", ""))
            if key.lower().endswith(".xlsx"):
                objects.append((obj.get("LastModified") or datetime.min.replace(tzinfo=ZoneInfo("UTC")), key))
    objects.sort(key=lambda x: x[0], reverse=True)
    return [k for _, k in objects[:limit]]


def key_exists(s3, bucket: str, key: str) -> bool:
    try:
        s3.head_object(Bucket=bucket, Key=key)
        return True
    except Exception:
        return False


def download_bytes(s3, bucket: str, key: str) -> bytes:
    return s3.get_object(Bucket=bucket, Key=key)["Body"].read()


def read_xlsx_bytes(raw: bytes, sheet: Any = 0) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(raw), sheet_name=sheet)


def read_book_sheets(raw: bytes) -> Dict[str, pd.DataFrame]:
    bio = io.BytesIO(raw)
    xls = pd.ExcelFile(bio)
    out: Dict[str, pd.DataFrame] = {}
    for sh in xls.sheet_names:
        try:
            out[sh] = pd.read_excel(xls, sheet_name=sh)
        except Exception:
            continue
    return out


def concat_weekly(s3, bucket: str, prefix: str, limit: int, normalizer) -> Tuple[pd.DataFrame, List[str]]:
    frames: List[pd.DataFrame] = []
    used: List[str] = []
    for key in reversed(list_keys(s3, bucket, prefix, limit=limit)):
        try:
            raw = download_bytes(s3, bucket, key)
            sheets = read_book_sheets(raw)
            part = normalizer(sheets, key)
            if part is not None and not part.empty:
                frames.append(part)
                used.append(key)
        except Exception as exc:
            print(f"WARN: {key}: {exc}", flush=True)
    if not frames:
        return pd.DataFrame(), used
    return pd.concat(frames, ignore_index=True), used


def choose_sheet(sheets: Dict[str, pd.DataFrame], preferred: Sequence[str]) -> pd.DataFrame:
    if not sheets:
        return pd.DataFrame()
    by_norm = {norm_col(k): v for k, v in sheets.items()}
    for name in preferred:
        if norm_col(name) in by_norm:
            return by_norm[norm_col(name)].copy()
    # первый непустой лист
    for df in sheets.values():
        if df is not None and not df.empty:
            return df.copy()
    return pd.DataFrame()


def normalize_orders(sheets: Dict[str, pd.DataFrame], key: str) -> pd.DataFrame:
    df = choose_sheet(sheets, ["Заказы", "orders", "Лист1", "Sheet1"])
    if df.empty:
        return pd.DataFrame()
    out = pd.DataFrame({
        "day": pd.to_datetime(col(df, ["date", "Дата", "Дата заказа", "Дата заказа покупателем"]), errors="coerce").dt.date,
        "supplier_article": col(df, ["supplierArticle", "Артикул продавца", "Артикул поставщика", "Артикул"]).map(clean_article),
        "nm_id": to_num(col(df, ["nmId", "nmID", "Артикул WB", "Код номенклатуры"])).astype("Int64"),
        "subject": col(df, ["subject", "Предмет", "Название предмета"]).astype(str).str.strip(),
        "category": col(df, ["category", "Категория"]).astype(str).str.strip(),
        "finished_price": to_num(col(df, ["finishedPrice", "finished_price", "Цена продажи", "Цена покупателя"])).fillna(0.0),
        "is_cancel": col(df, ["isCancel", "is_cancel", "Отменено"]).map(parse_bool),
        "srid": col(df, ["srid", "Srid", "SRID"]).astype(str),
        "g_number": col(df, ["gNumber", "g_number"]).astype(str),
        "source_key": key,
    })
    out = out[out["day"].notna() & out["nm_id"].notna()].copy()
    out["nm_id"] = out["nm_id"].astype(int)
    # Для "заказано" исключаем отменённые строки.
    out = out[~out["is_cancel"]].copy()
    # Dedupe: srid — лучший ключ; если пуст, fallback по набору полей.
    has_srid = out["srid"].astype(str).str.len().gt(5)
    a = out[has_srid].drop_duplicates(["srid"], keep="last")
    b = out[~has_srid].drop_duplicates(["day", "nm_id", "supplier_article", "finished_price", "g_number"], keep="last")
    return pd.concat([a, b], ignore_index=True)


def normalize_ads(sheets: Dict[str, pd.DataFrame], key: str) -> pd.DataFrame:
    df = choose_sheet(sheets, ["Статистика_Ежедневно", "Статистика ежедневно", "daily", "Реклама", "Лист1"])
    if df.empty:
        return pd.DataFrame()
    out = pd.DataFrame({
        "day": pd.to_datetime(col(df, ["Дата", "date", "day", "dt"]), errors="coerce").dt.date,
        "campaign_id": to_num(col(df, ["ID кампании", "campaign_id", "advertId", "advert_id"])).astype("Int64"),
        "nm_id": to_num(col(df, ["Артикул WB", "nmID", "nmId", "nm_id"])).astype("Int64"),
        "campaign_name": col(df, ["Название", "campaign_name", "Кампания"]).astype(str),
        "subject": col(df, ["Название предмета", "Предмет", "subject", "subject_norm"]).astype(str).str.strip(),
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
    # В перекрывающихся выгрузках одна дата/РК/nm может повториться: последняя версия должна победить.
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


def normalize_keywords(sheets: Dict[str, pd.DataFrame], key: str) -> pd.DataFrame:
    # Файл может иметь разные названия листов, ищем лист по наличию поискового запроса.
    candidates: List[pd.DataFrame] = []
    for df in sheets.values():
        if df is None or df.empty:
            continue
        if find_col(df, ["Поисковый запрос", "query", "search_query", "query_text"]):
            candidates.append(df)
    if not candidates:
        return pd.DataFrame()
    df = max(candidates, key=len).copy()
    out = pd.DataFrame({
        "day": pd.to_datetime(col(df, ["Дата", "date", "day", "dt"]), errors="coerce").dt.date,
        "supplier_article": col(df, ["Артикул продавца", "Артикул", "supplierArticle"]).map(clean_article),
        "nm_id": to_num(col(df, ["Артикул WB", "nmId", "nmID"])).astype("Int64"),
        "subject": col(df, ["Предмет", "Название предмета", "subject"]).astype(str).str.strip(),
        "query": col(df, ["Поисковый запрос", "query", "query_text", "search_query"]).astype(str).map(norm_text),
        "position": to_num(col(df, ["Медианная позиция", "Позиция", "median_position"])),
        "visibility": to_num(col(df, ["Видимость %", "Видимость", "visibility_pct"])),
        "frequency": to_num(col(df, ["Частота запросов", "Частотность", "frequency", "Частота за неделю"])),
        "search_impressions": to_num(col(df, ["Показы", "impressions"])),
        "search_clicks": to_num(col(df, ["Переходы в карточку", "Клики", "clicks"])),
        "search_orders": to_num(col(df, ["Заказы", "orders"])),
        "source_key": key,
    })
    out = out[out["day"].notna() & out["query"].ne("")].copy()
    out["nm_id"] = out["nm_id"].astype("Int64")
    return out


def normalize_manager_campaigns(sheets: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    if "РК_2ч_управление" not in sheets:
        return pd.DataFrame()
    df = sheets["РК_2ч_управление"].copy()
    if df.empty:
        return df
    out = pd.DataFrame({
        "campaign_id": to_num(col(df, ["campaign_id", "ID кампании"])).astype("Int64"),
        "nm_id": to_num(col(df, ["nm_id", "nmId", "Артикул WB"])).astype("Int64"),
        "campaign_name": col(df, ["campaign_name", "Название"]).astype(str),
        "campaign_status": col(df, ["campaign_status", "Статус"]).astype(str),
        "payment_type": col(df, ["payment_type", "Тип оплаты"]).astype(str),
        "bid_type": col(df, ["bid_type", "Тип ставки"]).astype(str),
        "search_bid": to_num(col(df, ["search_bid", "Ставка в поиске"])),
        "reco_bid": to_num(col(df, ["reco_bid", "Ставка в рекомендациях"])),
        "subject": col(df, ["subject_norm", "Предмет"]).astype(str),
        "supplier_article": col(df, ["supplier_article", "Артикул продавца"]).map(clean_article),
    })
    out = out[out["campaign_id"].notna()].copy()
    out["campaign_id"] = out["campaign_id"].astype(int)
    out["nm_id"] = out["nm_id"].astype("Int64")
    return out.drop_duplicates(["campaign_id", "nm_id"], keep="last")


def normalize_manager_queries(sheets: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    if "Запросы_2ч_управление" not in sheets:
        return pd.DataFrame()
    df = sheets["Запросы_2ч_управление"].copy()
    if df.empty:
        return df
    out = pd.DataFrame({
        "campaign_id": to_num(col(df, ["campaign_id", "advertId", "ID кампании"])).astype("Int64"),
        "nm_id": to_num(col(df, ["nm_id", "nmId", "Артикул WB"])).astype("Int64"),
        "query": col(df, ["norm_query_clean", "norm_query", "Поисковый запрос"]).astype(str).map(norm_text),
        "query_bid": to_num(col(df, ["current_bid_rub", "query_bid", "Ставка"])),
        "ad_impressions": to_num(col(df, ["impressions", "Показы"])),
        "ad_clicks": to_num(col(df, ["clicks", "Клики"])),
        "ad_spend": to_num(col(df, ["spend", "Расход"])),
        "ad_orders": to_num(col(df, ["orders", "Заказы"])),
        "position_manager": to_num(col(df, ["position", "Позиция"])),
        "visibility_manager": to_num(col(df, ["visibility", "Видимость"])),
        "efficiency_source": col(df, ["efficiency_source"]).astype(str),
        "keep_query": col(df, ["keep_query"]).map(parse_bool),
        "supplier_article": col(df, ["supplier_article", "Артикул продавца"]).map(clean_article),
    })
    out = out[out["campaign_id"].notna() & out["query"].ne("")].copy()
    out["campaign_id"] = out["campaign_id"].astype(int)
    out["nm_id"] = out["nm_id"].astype("Int64")
    return out.drop_duplicates(["campaign_id", "nm_id", "query"], keep="last")


def build_windows(as_of: date) -> Dict[str, PeriodWindow]:
    yesterday = as_of - timedelta(days=1)
    return {
        "1d": PeriodWindow(
            code="1d",
            label="Вчера vs средний день предыдущих 7 дней",
            current_from=yesterday,
            current_to=yesterday,
            base_from=as_of - timedelta(days=8),
            base_to=as_of - timedelta(days=2),
            base_is_daily_average=True,
        ),
        "7d": PeriodWindow(
            code="7d",
            label="Последние 7 завершённых дней vs предыдущие 7 дней",
            current_from=as_of - timedelta(days=7),
            current_to=yesterday,
            base_from=as_of - timedelta(days=14),
            base_to=as_of - timedelta(days=8),
            base_is_daily_average=False,
        ),
    }


def safe_div(a: Any, b: Any) -> Optional[float]:
    try:
        aa, bb = float(a), float(b)
        if not math.isfinite(aa) or not math.isfinite(bb) or bb == 0:
            return None
        return aa / bb
    except Exception:
        return None


def delta_pct(current: Any, base: Any) -> Optional[float]:
    ratio = safe_div(current, base)
    return None if ratio is None else (ratio - 1.0) * 100.0


def trend_payload(current: Any, base: Any, lower_is_better: bool = False) -> Dict[str, Any]:
    try:
        cur = float(current) if current is not None and not pd.isna(current) else None
        bas = float(base) if base is not None and not pd.isna(base) else None
    except Exception:
        cur, bas = None, None
    if cur is None or bas is None:
        return {"base": bas, "delta_abs": None, "delta_pct": None, "direction": "na", "tone": "neutral"}
    d_abs = cur - bas
    d_pct = delta_pct(cur, bas)
    if d_pct is None:
        direction = "up" if d_abs > 0 else "down" if d_abs < 0 else "flat"
    elif abs(d_pct) < TREND_FLAT_EPS_PCT:
        direction = "flat"
    else:
        direction = "up" if d_pct > 0 else "down"
    if direction == "flat":
        tone = "neutral"
    else:
        improvement = direction == "down" if lower_is_better else direction == "up"
        tone = "good" if improvement else "bad"
    return {"base": bas, "delta_abs": d_abs, "delta_pct": d_pct, "direction": direction, "tone": tone}


def drr_band(drr: Optional[float], spend: float = 0.0, orders: Optional[float] = None) -> str:
    if drr is None:
        if spend > 0 and orders is not None and orders <= 0:
            return "bad"
        return "neutral"
    if drr < 15.0:
        return "good"
    if drr < 20.0:
        return "warning"
    return "bad"


def campaign_type(row: pd.Series) -> str:
    payment = norm_text(row.get("payment_type", ""))
    bid_type = norm_text(row.get("bid_type", ""))
    search_bid = float(row.get("search_bid", 0) or 0)
    reco_bid = float(row.get("reco_bid", 0) or 0)
    name = norm_text(row.get("campaign_name", ""))
    if "cpc" in payment or "cpc" in bid_type:
        return "CPC"
    if search_bid > 0 and reco_bid > 0:
        return "CPM ПОИСК + ПОЛКИ"
    if search_bid > 0:
        return "CPM ПОИСК"
    if reco_bid > 0:
        return "CPM ПОЛКИ"
    if "полк" in name or "рек" in name or "рекомен" in name:
        return "CPM ПОЛКИ"
    return "CPM ПОИСК"


def product_map(stocks: pd.DataFrame, orders: pd.DataFrame) -> pd.DataFrame:
    parts: List[pd.DataFrame] = []
    if stocks is not None and not stocks.empty:
        parts.append(stocks[["nm_id", "supplier_article", "subject", "category"]].copy())
    if orders is not None and not orders.empty:
        p = orders[["nm_id", "supplier_article", "subject", "category"]].copy()
        parts.append(p)
    if not parts:
        return pd.DataFrame(columns=["nm_id", "supplier_article", "subject", "category"])
    df = pd.concat(parts, ignore_index=True)
    for c in ["supplier_article", "subject", "category"]:
        df[c] = df[c].fillna("").astype(str).str.strip()
    # Предпочитаем непустые последние значения.
    df["quality"] = (
        df["supplier_article"].ne("").astype(int) * 4
        + df["subject"].ne("").astype(int) * 2
        + df["category"].ne("").astype(int)
    )
    df = df.sort_values(["nm_id", "quality"]).drop_duplicates("nm_id", keep="last")
    return df.drop(columns=["quality"])


def add_product_fields(df: pd.DataFrame, pmap: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df.copy() if df is not None else pd.DataFrame()
    out = df.copy()
    if pmap is not None and not pmap.empty and "nm_id" in out.columns:
        p = pmap.rename(columns={"supplier_article": "map_article", "subject": "map_subject", "category": "map_category"})
        out = out.merge(p, on="nm_id", how="left")
        if "supplier_article" not in out.columns:
            out["supplier_article"] = ""
        if "subject" not in out.columns:
            out["subject"] = ""
        out["supplier_article"] = out["supplier_article"].fillna("").astype(str)
        out["subject"] = out["subject"].fillna("").astype(str)
        out["supplier_article"] = np.where(out["supplier_article"].str.strip().eq(""), out["map_article"].fillna(""), out["supplier_article"])
        out["subject"] = np.where(out["subject"].str.strip().eq(""), out["map_subject"].fillna(""), out["subject"])
        out["category"] = out.get("category", pd.Series([""] * len(out))).fillna("").astype(str)
        out["category"] = np.where(out["category"].str.strip().eq(""), out["map_category"].fillna(""), out["category"])
        out = out.drop(columns=[c for c in ["map_article", "map_subject", "map_category"] if c in out.columns])
    return out


def dedupe_ads(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    out = df.copy()
    out["_order"] = np.arange(len(out))
    return out.sort_values("_order").drop_duplicates(["day", "campaign_id", "nm_id"], keep="last").drop(columns=["_order"])


def date_filter(df: pd.DataFrame, start: date, end: date) -> pd.DataFrame:
    if df is None or df.empty or "day" not in df.columns:
        return pd.DataFrame(columns=df.columns if df is not None else [])
    return df[(df["day"] >= start) & (df["day"] <= end)].copy()


def aggregate_base_days(df: pd.DataFrame, keys: List[str], value_cols: List[str], w: PeriodWindow) -> pd.DataFrame:
    base = date_filter(df, w.base_from, w.base_to)
    if base.empty:
        return pd.DataFrame(columns=keys + value_cols)
    g = base.groupby(keys, dropna=False, as_index=False)[value_cols].sum()
    if w.base_is_daily_average:
        # База 1Д — средний календарный день ровно за 7 дней, включая нулевые дни.
        for c in value_cols:
            g[c] = g[c] / 7.0
    return g


def aggregate_current(df: pd.DataFrame, keys: List[str], value_cols: List[str], w: PeriodWindow) -> pd.DataFrame:
    cur = date_filter(df, w.current_from, w.current_to)
    if cur.empty:
        return pd.DataFrame(columns=keys + value_cols)
    return cur.groupby(keys, dropna=False, as_index=False)[value_cols].sum()


def merge_metrics(cur: pd.DataFrame, base: pd.DataFrame, keys: List[str], metrics: List[str], prefix: str = "") -> pd.DataFrame:
    if cur.empty and base.empty:
        return pd.DataFrame(columns=keys)
    c = cur.rename(columns={m: f"{prefix}{m}_current" for m in metrics})
    b = base.rename(columns={m: f"{prefix}{m}_base" for m in metrics})
    out = c.merge(b, on=keys, how="outer")
    for m in metrics:
        out[f"{prefix}{m}_current"] = pd.to_numeric(out.get(f"{prefix}{m}_current"), errors="coerce").fillna(0.0)
        out[f"{prefix}{m}_base"] = pd.to_numeric(out.get(f"{prefix}{m}_base"), errors="coerce").fillna(0.0)
    return out


def build_sales_metrics(orders: pd.DataFrame, keys: List[str], w: PeriodWindow) -> pd.DataFrame:
    vals = ["finished_price"]
    cur = aggregate_current(orders, keys, vals, w).rename(columns={"finished_price": "sales_sum"})
    base = aggregate_base_days(orders, keys, vals, w).rename(columns={"finished_price": "sales_sum"})
    return merge_metrics(cur, base, keys, ["sales_sum"])


def build_ads_metrics(ads: pd.DataFrame, keys: List[str], w: PeriodWindow) -> pd.DataFrame:
    vals = ["impressions", "clicks", "orders", "spend", "order_sum"]
    cur = aggregate_current(ads, keys, vals, w)
    base = aggregate_base_days(ads, keys, vals, w)
    return merge_metrics(cur, base, keys, vals, prefix="ad_")


def enrich_common(out: pd.DataFrame) -> pd.DataFrame:
    if out.empty:
        return out
    # Текущие агрегаты.
    out["ad_ctr_current"] = np.where(out["ad_impressions_current"] > 0, out["ad_clicks_current"] / out["ad_impressions_current"] * 100.0, np.nan)
    out["ad_ctr_base"] = np.where(out["ad_impressions_base"] > 0, out["ad_clicks_base"] / out["ad_impressions_base"] * 100.0, np.nan)
    out["ad_drr_current"] = np.where(out["ad_order_sum_current"] > 0, out["ad_spend_current"] / out["ad_order_sum_current"] * 100.0, np.nan)
    out["ad_drr_base"] = np.where(out["ad_order_sum_base"] > 0, out["ad_spend_base"] / out["ad_order_sum_base"] * 100.0, np.nan)
    if "sales_sum_current" in out.columns:
        out["overall_drr_current"] = np.where(out["sales_sum_current"] > 0, out["ad_spend_current"] / out["sales_sum_current"] * 100.0, np.nan)
        out["overall_drr_base"] = np.where(out["sales_sum_base"] > 0, out["ad_spend_base"] / out["sales_sum_base"] * 100.0, np.nan)
    return out


def add_dynamics(out: pd.DataFrame, metric_specs: Dict[str, bool]) -> pd.DataFrame:
    if out.empty:
        return out
    for metric, lower_is_better in metric_specs.items():
        ccol = metric + "_current"
        bcol = metric + "_base"
        if ccol not in out.columns or bcol not in out.columns:
            continue
        trends = [trend_payload(c, b, lower_is_better) for c, b in zip(out[ccol], out[bcol])]
        out[metric + "_delta_abs"] = [x["delta_abs"] for x in trends]
        out[metric + "_delta_pct"] = [x["delta_pct"] for x in trends]
        out[metric + "_direction"] = [x["direction"] for x in trends]
        out[metric + "_tone"] = [x["tone"] for x in trends]
    return out


def build_subject_table(orders: pd.DataFrame, ads: pd.DataFrame, w: PeriodWindow) -> pd.DataFrame:
    keys = ["subject"]
    sales = build_sales_metrics(orders, keys, w)
    ad = build_ads_metrics(ads, keys, w)
    out = sales.merge(ad, on=keys, how="outer")
    out = enrich_common(out)
    out = add_dynamics(out, {
        "sales_sum": False,
        "ad_impressions": False,
        "ad_clicks": False,
        "ad_order_sum": False,
        "ad_spend": True,
        "ad_ctr": False,
        "ad_drr": True,
        "overall_drr": True,
    })
    # 30 дней для сортировки предметов.
    return out.sort_values("sales_sum_current", ascending=False, na_position="last")


def build_article_table(orders: pd.DataFrame, ads: pd.DataFrame, w: PeriodWindow, as_of: date) -> pd.DataFrame:
    keys = ["subject", "supplier_article", "nm_id"]
    sales = build_sales_metrics(orders, keys, w)
    ad = build_ads_metrics(ads, keys, w)
    out = sales.merge(ad, on=keys, how="outer")
    out = enrich_common(out)
    out = add_dynamics(out, {
        "sales_sum": False,
        "ad_impressions": False,
        "ad_clicks": False,
        "ad_order_sum": False,
        "ad_spend": True,
        "ad_ctr": False,
        "ad_drr": True,
        "overall_drr": True,
    })
    # Цена покупателя после СПП: текущий период, fallback на последние 7 завершённых дней.
    cur_orders = date_filter(orders, w.current_from, w.current_to)
    price_cur = cur_orders.groupby(["supplier_article", "nm_id"], as_index=False).agg(avg_finished_price=("finished_price", "mean")) if not cur_orders.empty else pd.DataFrame(columns=["supplier_article", "nm_id", "avg_finished_price"])
    last7 = date_filter(orders, as_of - timedelta(days=7), as_of - timedelta(days=1))
    price_7 = last7.groupby(["supplier_article", "nm_id"], as_index=False).agg(avg_finished_price_7d=("finished_price", "mean")) if not last7.empty else pd.DataFrame(columns=["supplier_article", "nm_id", "avg_finished_price_7d"])
    out = out.merge(price_cur, on=["supplier_article", "nm_id"], how="left").merge(price_7, on=["supplier_article", "nm_id"], how="left")
    has_selected_price = out["avg_finished_price"].notna()
    out["avg_finished_price"] = out["avg_finished_price"].fillna(out["avg_finished_price_7d"])
    out["price_source"] = np.where(
        has_selected_price,
        "selected_period",
        np.where(out["avg_finished_price_7d"].notna(), "last7", "missing"),
    )
    # Сортировка карточек по количеству заказов за последние 30 завершённых дней; сумма — tie-breaker.
    last30 = date_filter(orders, as_of - timedelta(days=30), as_of - timedelta(days=1))
    if not last30.empty:
        rank = last30.groupby(["supplier_article", "nm_id"], as_index=False).agg(
            orders_count_30d=("finished_price", "size"),
            order_sum_30d=("finished_price", "sum"),
        )
        out = out.merge(rank, on=["supplier_article", "nm_id"], how="left")
    else:
        out["orders_count_30d"] = 0.0
        out["order_sum_30d"] = 0.0
    out["orders_count_30d"] = pd.to_numeric(out["orders_count_30d"], errors="coerce").fillna(0.0)
    out["order_sum_30d"] = pd.to_numeric(out["order_sum_30d"], errors="coerce").fillna(0.0)
    return out.sort_values(["subject", "orders_count_30d", "order_sum_30d"], ascending=[True, False, False])


def build_campaign_table(ads: pd.DataFrame, campaign_meta: pd.DataFrame, w: PeriodWindow) -> pd.DataFrame:
    keys = ["campaign_id", "nm_id", "supplier_article", "subject"]
    out = build_ads_metrics(ads, keys, w)
    out = enrich_common(out)
    if campaign_meta is not None and not campaign_meta.empty:
        meta = campaign_meta.copy()
        meta["campaign_type"] = meta.apply(campaign_type, axis=1)
        keep = ["campaign_id", "nm_id", "campaign_name", "campaign_status", "payment_type", "bid_type", "search_bid", "reco_bid", "campaign_type"]
        out = out.merge(meta[keep].drop_duplicates(["campaign_id", "nm_id"]), on=["campaign_id", "nm_id"], how="left")
    else:
        out["campaign_name"] = ""
        out["campaign_status"] = ""
        out["campaign_type"] = ""
    out = add_dynamics(out, {
        "ad_impressions": False,
        "ad_clicks": False,
        "ad_orders": False,
        "ad_order_sum": False,
        "ad_spend": True,
        "ad_ctr": False,
        "ad_drr": True,
    })
    out["has_search_queries"] = out["campaign_type"].astype(str).str.contains("ПОИСК", na=False)
    return out.sort_values(["subject", "supplier_article", "ad_order_sum_current", "ad_spend_current"], ascending=[True, True, False, False])


def weighted_keyword_stats(keywords: pd.DataFrame, w: PeriodWindow) -> Tuple[pd.DataFrame, pd.DataFrame]:
    if keywords is None or keywords.empty:
        cols = ["supplier_article", "nm_id", "query", "position", "visibility"]
        return pd.DataFrame(columns=cols), pd.DataFrame(columns=cols)

    def agg(part: pd.DataFrame, average_daily: bool) -> pd.DataFrame:
        if part.empty:
            return pd.DataFrame(columns=["supplier_article", "nm_id", "query", "position", "visibility", "frequency"])
        # Позиция/видимость: среднее по дневным наблюдениям; если есть частота — используем её как вес.
        rows = []
        for keys, g in part.groupby(["supplier_article", "nm_id", "query"], dropna=False):
            weight = pd.to_numeric(g["frequency"], errors="coerce").fillna(0.0)
            def wavg(name: str):
                vals = pd.to_numeric(g[name], errors="coerce")
                mask = vals.notna()
                if not mask.any():
                    return np.nan
                ww = weight.where(mask, 0.0)
                if ww.sum() > 0:
                    return float((vals.fillna(0.0) * ww).sum() / ww.sum())
                return float(vals[mask].mean())
            rows.append({
                "supplier_article": keys[0],
                "nm_id": keys[1],
                "query": keys[2],
                "position": wavg("position"),
                "visibility": wavg("visibility"),
                "frequency": float(pd.to_numeric(g["frequency"], errors="coerce").fillna(0.0).mean()),
            })
        return pd.DataFrame(rows)

    cur = agg(date_filter(keywords, w.current_from, w.current_to), False)
    base = agg(date_filter(keywords, w.base_from, w.base_to), w.base_is_daily_average)
    return cur, base


def build_query_table(
    manager_queries: pd.DataFrame,
    keywords: pd.DataFrame,
    article_table: pd.DataFrame,
    campaign_table: pd.DataFrame,
    w: PeriodWindow,
) -> pd.DataFrame:
    if manager_queries is None or manager_queries.empty:
        return pd.DataFrame()
    q = manager_queries.copy()
    # Карта артикулов/предметов из кампаний.
    camp_map_cols = ["campaign_id", "nm_id", "supplier_article", "subject", "campaign_type", "campaign_name"]
    camp_map = campaign_table[camp_map_cols].drop_duplicates(["campaign_id", "nm_id"]) if not campaign_table.empty else pd.DataFrame(columns=camp_map_cols)
    q = q.merge(camp_map, on=["campaign_id", "nm_id"], how="left", suffixes=("", "_campaign"))
    q["supplier_article"] = np.where(q["supplier_article"].astype(str).str.strip().eq(""), q["supplier_article_campaign"].fillna(""), q["supplier_article"])
    q = q.drop(columns=[c for c in ["supplier_article_campaign"] if c in q.columns])

    # Позиции/видимость из независимого WB search-report.
    pos_cur, pos_base = weighted_keyword_stats(keywords, w)
    pos_cur = pos_cur.rename(columns={"position": "position_current", "visibility": "visibility_current", "frequency": "frequency_current"})
    pos_base = pos_base.rename(columns={"position": "position_base", "visibility": "visibility_base", "frequency": "frequency_base"})
    q = q.merge(pos_cur, on=["supplier_article", "nm_id", "query"], how="left")
    q = q.merge(pos_base, on=["supplier_article", "nm_id", "query"], how="left")
    q["position_current"] = q["position_current"].fillna(q["position_manager"])
    q["visibility_current"] = q["visibility_current"].fillna(q["visibility_manager"])

    # В текущем менеджере query-level stats могут быть недоступны (WB /normquery/stats => items:null).
    # Не выдаём нули за реальные метрики: если источник no_efficiency_data, делаем NA.
    no_stats = q["efficiency_source"].astype(str).str.contains("no_efficiency_data", case=False, na=True)
    for c in ["ad_impressions", "ad_clicks", "ad_spend", "ad_orders"]:
        q.loc[no_stats, c] = np.nan
    q["query_stats_available"] = ~no_stats & q[["ad_impressions", "ad_clicks", "ad_spend", "ad_orders"]].notna().any(axis=1)

    # Цена товара для CPO/ДРР запроса.
    prices = article_table[["supplier_article", "nm_id", "avg_finished_price"]].drop_duplicates(["supplier_article", "nm_id"]) if not article_table.empty else pd.DataFrame(columns=["supplier_article", "nm_id", "avg_finished_price"])
    q = q.merge(prices, on=["supplier_article", "nm_id"], how="left")
    q["ctr_current"] = np.where(q["ad_impressions"] > 0, q["ad_clicks"] / q["ad_impressions"] * 100.0, np.nan)
    q["cpo_current"] = np.where(q["ad_orders"] > 0, q["ad_spend"] / q["ad_orders"], np.nan)
    q["query_drr_current"] = np.where((q["cpo_current"].notna()) & (q["avg_finished_price"] > 0), q["cpo_current"] / q["avg_finished_price"] * 100.0, np.nan)
    q["query_drr_band"] = [drr_band(d, float(s or 0), float(o) if not pd.isna(o) else None) for d, s, o in zip(q["query_drr_current"], q["ad_spend"], q["ad_orders"])]
    q["position_delta"] = q["position_current"] - q["position_base"]
    q["visibility_delta"] = q["visibility_current"] - q["visibility_base"]
    # Для позиции уменьшение — улучшение.
    q["position_tone"] = np.where(q["position_delta"].isna(), "neutral", np.where(q["position_delta"] < 0, "good", np.where(q["position_delta"] > 0, "bad", "neutral")))
    q["visibility_tone"] = np.where(q["visibility_delta"].isna(), "neutral", np.where(q["visibility_delta"] > 0, "good", np.where(q["visibility_delta"] < 0, "bad", "neutral")))

    # Сортировка по заказам от большего к меньшему, как запрошено.
    q["sort_orders"] = pd.to_numeric(q["ad_orders"], errors="coerce").fillna(-1)
    q["sort_clicks"] = pd.to_numeric(q["ad_clicks"], errors="coerce").fillna(-1)
    return q.sort_values(["campaign_id", "sort_orders", "sort_clicks", "query"], ascending=[True, False, False, True])


def data_quality_table(
    source_keys: Dict[str, List[str]],
    orders: pd.DataFrame,
    ads: pd.DataFrame,
    stocks: pd.DataFrame,
    keywords: pd.DataFrame,
    manager_queries: pd.DataFrame,
) -> pd.DataFrame:
    return pd.DataFrame([
        {"source": "WB orders", "status": "OK" if not orders.empty else "MISSING", "rows": len(orders), "note": "finishedPrice после СПП; отмены исключены", "keys": "\n".join(source_keys.get("orders", []))},
        {"source": "WB ads daily", "status": "OK" if not ads.empty else "MISSING", "rows": len(ads), "note": "РК: показы/клики/заказы/расход/сумма заказов", "keys": "\n".join(source_keys.get("ads", []))},
        {"source": "WB product map", "status": "OK" if not stocks.empty else "FALLBACK", "rows": len(stocks), "note": "nmID -> артикул продавца -> предмет", "keys": "\n".join(source_keys.get("stocks", []))},
        {"source": "WB search positions", "status": "OK" if not keywords.empty else "MISSING", "rows": len(keywords), "note": "позиция/видимость; НЕ рекламный расход", "keys": "\n".join(source_keys.get("keywords", []))},
        {"source": "WB CPM current queries", "status": "OK" if not manager_queries.empty else "MISSING", "rows": len(manager_queries), "note": "актуальный список запросов/ставок; query-level spend может быть недоступен", "keys": MANAGER_KEY},
    ])


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
    if pd.isna(v):
        return None
    return v


def records(df: pd.DataFrame) -> List[Dict[str, Any]]:
    if df is None or df.empty:
        return []
    return [{k: jsonable(v) for k, v in row.items()} for row in df.to_dict(orient="records")]


def build_payload(as_of: date, windows: Dict[str, PeriodWindow], tables: Dict[str, Dict[str, pd.DataFrame]], quality: pd.DataFrame) -> Dict[str, Any]:
    payload: Dict[str, Any] = {
        "meta": {
            "version": VERSION,
            "store": STORE,
            "generated_at": datetime.now(MOSCOW).isoformat(),
            "as_of": as_of.isoformat(),
            "today_excluded": True,
            "query_drr_thresholds": {"good_lt": 15.0, "warning_from": 15.0, "warning_to_lt": 20.0, "bad_from": 20.0},
        },
        "data_quality": records(quality),
        "periods": {},
    }
    for code, w in windows.items():
        payload["periods"][code] = {
            "window": {
                "label": w.label,
                "current_from": w.current_from.isoformat(),
                "current_to": w.current_to.isoformat(),
                "base_from": w.base_from.isoformat(),
                "base_to": w.base_to.isoformat(),
                "base_is_daily_average": w.base_is_daily_average,
            },
            "subjects": records(tables[code]["subjects"]),
            "articles": records(tables[code]["articles"]),
            "campaigns": records(tables[code]["campaigns"]),
            "queries": records(tables[code]["queries"]),
        }
    return payload


def write_excel(path: Path, as_of: date, windows: Dict[str, PeriodWindow], tables: Dict[str, Dict[str, pd.DataFrame]], quality: pd.DataFrame, pmap: pd.DataFrame) -> None:
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        params = []
        for code, w in windows.items():
            params.append({
                "period": code,
                "label": w.label,
                "current_from": w.current_from,
                "current_to": w.current_to,
                "base_from": w.base_from,
                "base_to": w.base_to,
                "base_mode": "avg/day" if w.base_is_daily_average else "sum",
            })
        pd.DataFrame(params).to_excel(writer, sheet_name="00_Окна", index=False)
        quality.to_excel(writer, sheet_name="01_DataQuality", index=False)
        pmap.to_excel(writer, sheet_name="02_ProductMap", index=False)
        for code in ["1d", "7d"]:
            tables[code]["subjects"].to_excel(writer, sheet_name=f"10_Предмет_{code.upper()}", index=False)
            tables[code]["articles"].to_excel(writer, sheet_name=f"20_Артикул_{code.upper()}", index=False)
            tables[code]["campaigns"].to_excel(writer, sheet_name=f"30_РК_{code.upper()}", index=False)
            tables[code]["queries"].to_excel(writer, sheet_name=f"40_Запросы_{code.upper()}", index=False)
        # Служебный лист формул.
        formulas = pd.DataFrame([
            ["Общий ДРР предмета/артикула", "Расход всех РК / сумма ВСЕХ заказов finishedPrice * 100", "верхняя карточка"],
            ["Рекламный ДРР", "Расход РК / сумма заказов, атрибутированных рекламе * 100", "нижний рекламный блок"],
            ["CTR", "Клики / Показы * 100", "РК и запрос"],
            ["CPO запроса", "Расход запроса / Заказы запроса", "только при наличии query-level ad stats"],
            ["ДРР запроса", "CPO / средний finishedPrice артикула * 100", "<15 зелёный; 15-20 оранжевый; >=20 красный"],
            ["1Д динамика", "Вчера / средний день предыдущих 7 дней - 1", "сегодня исключён"],
            ["7Д динамика", "Последние 7 завершённых дней / предыдущие 7 дней - 1", "сегодня исключён"],
        ], columns=["metric", "formula", "note"])
        formulas.to_excel(writer, sheet_name="99_Формулы", index=False)
        # Базовое оформление.
        wb = writer.book
        header = wb.add_format({"bold": True, "bg_color": "#0F274F", "font_color": "#FFFFFF", "border": 1})
        money = wb.add_format({"num_format": '#,##0 "₽"'})
        pct = wb.add_format({"num_format": "0.0%"})
        num1 = wb.add_format({"num_format": "0.0"})
        wrap = wb.add_format({"text_wrap": True, "valign": "top"})
        for sheet_name, ws in writer.sheets.items():
            # header row
            try:
                max_col = writer.sheets[sheet_name].dim_colmax
            except Exception:
                max_col = 30
            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, max(1, ws.dim_rowmax), max(0, ws.dim_colmax))
            ws.set_row(0, 24, header)
            ws.set_column(0, min(ws.dim_colmax, 60), 14)
            ws.set_column(0, 1, 22)
            if sheet_name in {"01_DataQuality", "99_Формулы"}:
                ws.set_column(0, ws.dim_colmax, 28, wrap)


def upload_file(s3, bucket: str, local_path: Path, key: str, content_type: str) -> None:
    s3.upload_file(str(local_path), bucket, key, ExtraArgs={"ContentType": content_type})


def run(as_of: date, out_dir: Path, upload: bool = True) -> Tuple[Path, Path]:
    s3 = make_s3()
    bucket = env_required("YC_BUCKET_NAME")
    source_keys: Dict[str, List[str]] = {"orders": [], "ads": [], "stocks": [], "keywords": []}

    orders, source_keys["orders"] = concat_weekly(s3, bucket, ORDERS_PREFIX, 5, normalize_orders)
    ads, source_keys["ads"] = concat_weekly(s3, bucket, ADS_PREFIX, 5, normalize_ads)
    stocks, source_keys["stocks"] = concat_weekly(s3, bucket, STOCKS_PREFIX, 2, normalize_stocks)
    keywords = pd.DataFrame()
    for prefix in KEYWORDS_PREFIXES:
        part, keys = concat_weekly(s3, bucket, prefix, 5, normalize_keywords)
        if not part.empty:
            keywords = pd.concat([keywords, part], ignore_index=True) if not keywords.empty else part
            source_keys["keywords"].extend(keys)

    manager_campaigns = pd.DataFrame()
    manager_queries = pd.DataFrame()
    if key_exists(s3, bucket, MANAGER_KEY):
        try:
            msheets = read_book_sheets(download_bytes(s3, bucket, MANAGER_KEY))
            manager_campaigns = normalize_manager_campaigns(msheets)
            manager_queries = normalize_manager_queries(msheets)
        except Exception as exc:
            print(f"WARN manager workbook: {exc}", flush=True)

    # Product map is canonical link nmID -> article -> subject.
    pmap = product_map(stocks, orders)
    orders = add_product_fields(orders, pmap)
    ads = add_product_fields(ads, pmap)
    manager_campaigns = add_product_fields(manager_campaigns, pmap)
    manager_queries = add_product_fields(manager_queries, pmap)
    keywords = add_product_fields(keywords, pmap)

    ads = dedupe_ads(ads)
    # Метаданные кампаний: если manager ещё не видел часть РК, создаём минимальные строки из ads.
    if ads is not None and not ads.empty:
        fallback = ads[["campaign_id", "nm_id", "campaign_name", "subject", "supplier_article"]].drop_duplicates(["campaign_id", "nm_id"], keep="last")
        for c in ["campaign_status", "payment_type", "bid_type"]:
            fallback[c] = ""
        for c in ["search_bid", "reco_bid"]:
            fallback[c] = 0.0
        if manager_campaigns.empty:
            manager_campaigns = fallback
        else:
            existing = set(zip(manager_campaigns["campaign_id"].astype(int), manager_campaigns["nm_id"].fillna(-1).astype(int)))
            add = fallback[[ (int(r.campaign_id), int(r.nm_id) if not pd.isna(r.nm_id) else -1) not in existing for r in fallback.itertuples() ]]
            if not add.empty:
                manager_campaigns = pd.concat([manager_campaigns, add], ignore_index=True)

    # В ads campaign_name/subject/article должны быть заполнены.
    if not ads.empty:
        cmap = manager_campaigns[["campaign_id", "nm_id", "campaign_name", "subject", "supplier_article"]].drop_duplicates(["campaign_id", "nm_id"])
        ads = ads.drop(columns=[c for c in ["campaign_name", "subject", "supplier_article"] if c in ads.columns]).merge(cmap, on=["campaign_id", "nm_id"], how="left")
        ads = add_product_fields(ads, pmap)

    windows = build_windows(as_of)
    tables: Dict[str, Dict[str, pd.DataFrame]] = {}
    for code, w in windows.items():
        articles = build_article_table(orders, ads, w, as_of)
        campaigns = build_campaign_table(ads, manager_campaigns, w)
        subjects = build_subject_table(orders, ads, w)
        queries = build_query_table(manager_queries, keywords, articles, campaigns, w)
        tables[code] = {"subjects": subjects, "articles": articles, "campaigns": campaigns, "queries": queries}

    quality = data_quality_table(source_keys, orders, ads, stocks, keywords, manager_queries)
    payload = build_payload(as_of, windows, tables, quality)

    out_dir.mkdir(parents=True, exist_ok=True)
    xlsx_path = out_dir / "Ассистент_WB_FINICK_Реклама_технический.xlsx"
    json_path = out_dir / "Ассистент_WB_FINICK_Реклама_ui.json"
    write_excel(xlsx_path, as_of, windows, tables, quality, pmap)
    json_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    if upload:
        # current/ всегда содержит последнюю актуальную версию для веб-приложения.
        upload_file(s3, bucket, xlsx_path, OUTPUT_XLSX_KEY, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        upload_file(s3, bucket, json_path, OUTPUT_JSON_KEY, "application/json; charset=utf-8")

        # history/YYYY-MM-DD/ в течение одного дня перезаписывается в 09/12/16 МСК.
        # На следующий день создаётся новый дневной срез.
        hist_prefix = OUTPUT_HISTORY_PREFIX + f"{as_of:%Y-%m-%d}/"
        hist_xlsx_key = hist_prefix + "Техническая_аналитика.xlsx"
        hist_json_key = hist_prefix + "ui_payload.json"
        upload_file(s3, bucket, xlsx_path, hist_xlsx_key, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        upload_file(s3, bucket, json_path, hist_json_key, "application/json; charset=utf-8")

        meta = {
            "version": VERSION,
            "store": STORE,
            "as_of": as_of.isoformat(),
            "generated_at_msk": datetime.now(MOSCOW).isoformat(),
            "today_excluded": True,
            "periods": {
                code: {
                    "current_from": w.current_from.isoformat(),
                    "current_to": w.current_to.isoformat(),
                    "base_from": w.base_from.isoformat(),
                    "base_to": w.base_to.isoformat(),
                    "base_is_daily_average": w.base_is_daily_average,
                }
                for code, w in windows.items()
            },
            "outputs": {
                "current_xlsx": OUTPUT_XLSX_KEY,
                "current_json": OUTPUT_JSON_KEY,
                "history_xlsx": hist_xlsx_key,
                "history_json": hist_json_key,
            },
        }
        s3.put_object(
            Bucket=bucket,
            Key=OUTPUT_META_KEY,
            Body=json.dumps(meta, ensure_ascii=False, indent=2).encode("utf-8"),
            ContentType="application/json; charset=utf-8",
        )
        print(f"S3 current: s3://{bucket}/{OUTPUT_XLSX_KEY}")
        print(f"S3 current: s3://{bucket}/{OUTPUT_JSON_KEY}")
        print(f"S3 history: s3://{bucket}/{hist_json_key}")

    print(f"Готово: {xlsx_path}")
    print(f"Готово: {json_path}")
    return xlsx_path, json_path


def main(argv: Optional[Sequence[str]] = None) -> int:
    loaded = load_report_env()
    if loaded:
        print("REPORT_ENV: подхвачены секреты: " + ", ".join(sorted(loaded)), flush=True)
    p = argparse.ArgumentParser(description="Ассистент WB — расчёт аналитики рекламы FINICK")
    p.add_argument("--as-of", default="", help="Дата расчёта YYYY-MM-DD; по умолчанию сегодня по Москве")
    p.add_argument("--out-dir", default=".")
    p.add_argument("--no-upload", action="store_true")
    args = p.parse_args(list(argv) if argv is not None else None)
    as_of = datetime.strptime(args.as_of, "%Y-%m-%d").date() if args.as_of else datetime.now(MOSCOW).date()
    run(as_of, Path(args.out_dir), upload=not args.no_upload)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
