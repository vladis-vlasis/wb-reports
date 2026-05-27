#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
assistant_wb_ads_manager.py

Новый скрипт управления рекламными ставками WB для магазина TOPFACE.

Логика принятия решений строго бинарная:
- ДРР кампании >= 10.0% -> снизить ставку;
- ДРР кампании < 10.0% -> повысить ставку.

Активная управляемая строка ставки не удерживается без действия, кроме технических
исключений с фиксированным reason_code.
"""

from __future__ import annotations

import argparse
import io
import json
import math
import os
import re
import sys
import time
import uuid
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from typing import Any, Dict, Iterable, List, Optional, Tuple

import boto3
import pandas as pd
import requests
from botocore.exceptions import ClientError


# =============================
# Константы проекта
# =============================

SCRIPT_NAME = "upravlenie_reklamoy.py"
SCRIPT_VERSION = "strict-drr-v19-economics-zero-base-hotfix-dtype-fix-2026-05-27"
STORE_NAME = "TOPFACE"
DRR_LIMIT_PCT = 10.0
TECHNICAL_BID_FLOOR_RUB = 1.0
ANALYSIS_WINDOW_DAYS = 5

# ДРР теперь не общий для всех: для помад, блесков и косметических карандашей допустимый порог 15%.
SUBJECT_DRR_LIMITS = {
    "кисти косметические": 10.0,
    "помады": 15.0,
    "блески": 15.0,
    "косметические карандаши": 15.0,
}

# Пауза запрещена для кистей. Пауза разрешена только для этих предметов и только при достаточной статистике.
PAUSE_ALLOWED_SUBJECTS = {"помады", "блески", "косметические карандаши"}
PAUSE_MIN_IMPRESSIONS = int(os.environ.get("WB_PAUSE_MIN_IMPRESSIONS", "10000") or 10000)
PAUSE_ANALYSIS_DAYS = int(os.environ.get("WB_PAUSE_ANALYSIS_DAYS", "21") or 21)
AUTO_APPLY_PAUSE_REASON_CODES = {"PAUSE_MIN_BID_HIGH_DRR_21D_10000"}

# Разгон слабых кампаний: цель — выйти на >=1000 показов/день при расходе <=500 ₽/день.
RAMP_TARGET_IMPRESSIONS_PER_DAY = float(os.environ.get("WB_RAMP_TARGET_IMPRESSIONS_PER_DAY", "1000") or 1000)
RAMP_MAX_SPEND_PER_DAY = float(os.environ.get("WB_RAMP_MAX_SPEND_PER_DAY", "500") or 500)
RAMP_CHECK_DAYS = int(os.environ.get("WB_RAMP_CHECK_DAYS", "7") or 7)

# Эксперимент: одна лучшая РК на товарную группу для помад/блесков/карандашей.
ONE_CAMPAIGN_EXPERIMENT_SUBJECTS = {"помады", "блески", "косметические карандаши"}
ONE_CAMPAIGN_TARGET_POSITION = int(os.environ.get("WB_ONE_CAMPAIGN_TARGET_POSITION", "10") or 10)
ONE_CAMPAIGN_CHECK_DAYS = (7, 10, 14)

MANAGED_SUBJECTS = {
    "кисти косметические",
    "помады",
    "блески",
    "косметические карандаши",
}

SERVICE_PREFIX = "Служебные файлы/Ассистент WB/TOPFACE/"
ADS_MAIN_KEY = "Отчёты/Реклама/TOPFACE/Анализ рекламы.xlsx"
ADS_WEEKLY_PREFIX = "Отчёты/Реклама/TOPFACE/Недельные/"

BID_HISTORY_KEY = SERVICE_PREFIX + "История_ставок.xlsx"
PAUSE_HISTORY_KEY = SERVICE_PREFIX + "История_пауз.xlsx"
RUN_OUTPUT_KEY = SERVICE_PREFIX + "Итог_последнего_запуска.xlsx"
PREVIEW_OUTPUT_KEY = SERVICE_PREFIX + "Предпросмотр_последнего_запуска.xlsx"
SUMMARY_JSON_KEY = SERVICE_PREFIX + "Сводка_последнего_запуска.json"
API_LOG_KEY = SERVICE_PREFIX + "Лог_API.xlsx"

KEYWORDS_WEEKLY_PREFIX = "Отчёты/Поисковые запросы/TOPFACE/Недельные/"
FUNNEL_KEY = "Отчёты/Воронка продаж/TOPFACE/Воронка продаж.xlsx"
ECONOMICS_KEY = "Отчёты/Финансовые показатели/TOPFACE/Экономика.xlsx"
PRICE_HISTORY_KEY = SERVICE_PREFIX + "История_изменений_цен.xlsx"
ONE_CAMPAIGN_EXPERIMENT_HISTORY_KEY = SERVICE_PREFIX + "История_эксперимента_1РК.xlsx"

WB_PRICES_BASE_URL = "https://discounts-prices-api.wildberries.ru"
WB_PRICES_LIST_ENDPOINT = "/api/v2/list/goods/filter"
WB_PRICE_UPLOAD_ENDPOINT = "/api/v2/upload/task"

DEFAULT_SELLER_DISCOUNT_PCT = 26
DEFAULT_PRICE_RAISE_STEP_PP = 1
DEFAULT_MIN_SELLER_DISCOUNT_PCT = int(os.environ.get("WB_PRICE_MIN_SELLER_DISCOUNT_PCT", "25") or 25)
PRICE_TEST_SUBJECTS = {"помады", "блески", "косметические карандаши"}
MAX_PRICE_TEST_ITEMS_PER_RUN = int(os.environ.get("WB_MAX_PRICE_TEST_ITEMS_PER_RUN", "30") or 30)
# Условная ВП после рекламы: вычитаем себестоимость, если она есть в Экономике.
ECONOMICS_SUBTRACT_COGS = str(os.environ.get("WB_ECONOMICS_SUBTRACT_COGS", "1")).strip().lower() not in {"0", "false", "no", "нет"}


WB_ADVERT_BASE_URL = "https://advert-api.wildberries.ru"
WB_BIDS_ENDPOINT = "/api/advert/v1/bids"
WB_BIDS_MIN_ENDPOINT = "/api/advert/v1/bids/min"

MIN_BID_COLUMNS = [
    "run_datetime",
    "campaign_id",
    "nm_id",
    "placement",
    "payment_type",
    "min_bid_rub",
    "api_status",
    "response_text",
]

RENAME_CAMPAIGN_COLUMNS = [
    "run_datetime",
    "campaign_id",
    "current_name",
    "target_name",
    "supplier_article",
    "nm_ids",
    "subjects",
    "rename_action",
    "reason_code",
    "api_status",
    "response_text",
]
WB_PAUSE_ENDPOINT = "/adv/v0/pause"
WB_START_ENDPOINT = "/adv/v0/start"
WB_RENAME_ENDPOINT = "/adv/v0/rename"

BID_HISTORY_COLUMNS = [
    "event_id",
    "run_datetime",
    "event_date",
    "campaign_id",
    "nm_id",
    "supplier_article",
    "subject_norm",
    "placement",
    "old_bid_rub",
    "new_bid_rub",
    "direction",
    "reason_code",
    "spend_before",
    "revenue_before",
    "orders_before",
    "impressions_before",
    "clicks_before",
    "drr_before",
    "gp_before",
    "postcheck_status",
    "final_verdict",
    "d1_verdict",
    "d3_verdict",
    "d1_check_date",
    "d3_check_date",
    "d7_verdict",
    "d7_check_date",
]

KEYWORD_POSITION_COLUMNS = [
    "date", "nm_id", "supplier_article", "subject_norm", "query_text", "filter_type",
    "query_freq", "median_position", "visibility_pct", "clicks_to_card", "keyword_orders",
    "keyword_group", "orders_share", "orders_cum_share",
]

KEYWORD_EFFECT_COLUMNS = [
    "event_id", "campaign_id", "nm_id", "supplier_article", "placement", "direction", "event_date",
    "keyword_group", "queries_count",
    "position_before", "position_d1", "position_d3", "position_delta_d1", "position_delta_d3",
    "visibility_before", "visibility_d1", "visibility_d3", "visibility_delta_d1", "visibility_delta_d3",
    "keyword_orders_before", "keyword_orders_d1", "keyword_orders_d3",
    "query_freq_before", "query_freq_d1", "query_freq_d3",
    "keyword_verdict_d1", "keyword_verdict_d3", "keyword_comment",
]

PRICE_HISTORY_COLUMNS = [
    "price_event_id", "run_datetime", "event_date", "nm_id", "supplier_article", "subject_norm",
    "old_discount", "new_discount", "direction", "reason_code",
    "orders_before", "impressions_before", "clicks_before", "ctr_before",
    "card_views_before", "add_to_cart_before", "funnel_orders_before",
    "add_to_cart_conv_before", "cart_to_order_conv_before", "funnel_missing",
    "postcheck_status", "final_verdict", "d2_verdict", "d2_check_date",
    "api_status", "api_response",
]

PRICE_DECISION_COLUMNS = [
    "nm_id", "supplier_article", "subject_norm", "current_discount", "new_discount", "price_action",
    "reason_code", "reason_text", "orders", "impressions", "clicks", "ctr_pct",
    "card_views", "add_to_cart", "funnel_orders", "add_to_cart_conv", "cart_to_order_conv", "funnel_missing",
    "previous_price_event_id", "price_postcheck_status",
]

BID_RAMP_MONITOR_COLUMNS = [
    "campaign_id", "nm_id", "supplier_article", "subject_norm", "placement", "campaign_status",
    "ramp_candidate", "ramp_mode_status", "ramp_applied_in_current_run", "ramp_reason_group",
    "current_bid_rub", "min_bid_rub", "new_bid_rub", "api_status",
    "reason_code", "reason_text",
    "wait_status", "wait_rule", "wait_until_date", "wait_days_left",
    "last_bid_change_date", "last_bid_change_old_bid", "last_bid_change_new_bid", "last_bid_change_reason_code",
    "impressions", "avg_impressions_per_day", "spend", "avg_spend_per_day", "orders", "revenue",
    "campaign_drr_pct", "drr_limit_pct", "last21_impressions", "last21_drr_pct",
    "target_impressions_per_day", "max_spend_per_day", "check_days", "monitor_status",
]

BID_CAMPAIGN_COMPARE_COLUMNS = [
    "campaign_id", "nm_id", "supplier_article", "subject_norm", "placement", "campaign_status",
    "economics_match_method", "economics_product_group", "economics_avg_price", "economics_commission_pct",
    "economics_acquiring_pct", "economics_vat_per_unit", "economics_logistics_per_unit", "economics_cogs_per_unit",
    "comparison_status", "last_bid_change_date", "old_bid_rub", "new_bid_rub", "bid_change_reason_code",
    "before_period", "after_period", "before_days", "after_days",
    "current_action", "current_reason_code", "current_reason_text",
    "postcheck_status", "final_verdict", "target_bid_action", "target_bid_action_text", "recommended_next_bid_rub",
    "before_impressions", "after_impressions", "before_impressions_per_day", "after_impressions_per_day", "impressions_delta_pct",
    "before_clicks", "after_clicks", "before_clicks_per_day", "after_clicks_per_day", "clicks_delta_pct",
    "before_ctr_pct", "after_ctr_pct", "ctr_delta_pp",
    "before_orders_qty", "after_orders_qty", "before_orders_qty_per_day", "after_orders_qty_per_day", "orders_qty_delta_pct",
    "before_orders_sum_rub", "after_orders_sum_rub", "before_orders_sum_rub_per_day", "after_orders_sum_rub_per_day", "orders_sum_delta_pct",
    "before_gp_after_ads_rub", "after_gp_after_ads_rub", "before_gp_after_ads_rub_per_day", "after_gp_after_ads_rub_per_day", "gp_after_ads_delta_pct",
    "before_ad_spend_rub", "after_ad_spend_rub", "before_ad_spend_rub_per_day", "after_ad_spend_rub_per_day", "ad_spend_delta_pct",
    "before_drr_pct", "after_drr_pct", "drr_delta_pp",
    "core80_queries_count",
    "before_core80_avg_position", "after_core80_avg_position", "core80_position_delta",
    "before_core80_visibility_pct", "after_core80_visibility_pct", "core80_visibility_delta_pp",
    "before_core80_query_freq", "after_core80_query_freq",
    "before_card_views", "after_card_views", "before_card_views_per_day", "after_card_views_per_day",
    "before_traffic_share_pct", "after_traffic_share_pct", "traffic_share_delta_pp",
    "diagnostic_conclusion",
]

ONE_CAMPAIGN_EXPERIMENT_COLUMNS = [
    "experiment_id", "run_datetime", "subject_norm", "product_group",
    "selected_campaign_id", "selected_nm_id", "selected_supplier_article", "selected_placement",
    "selected_bid_rub", "recommended_new_bid_rub", "selection_basis",
    "group_campaigns_count", "campaigns_to_pause", "pause_campaign_ids",
    "group_gp_before", "group_orders_before", "group_revenue_before", "group_spend_before",
    "group_impressions_before", "group_clicks_before", "group_ctr_before", "group_conversion_before",
    "selected_gp_before", "selected_orders_before", "selected_revenue_before", "selected_spend_before",
    "selected_impressions_before", "selected_clicks_before", "selected_ctr_before", "selected_conversion_before",
    "core_queries_count", "core_median_position", "core_target_position",
    "recommended_action", "reason_code", "reason_text", "check_days",
]

PAUSE_HISTORY_COLUMNS = [
    "pause_event_id",
    "pause_date",
    "campaign_id",
    "nm_id",
    "placement",
    "supplier_article",
    "subject_norm",
    "reason_code",
    "impressions_before_pause",
    "clicks_before_pause",
    "spend_before_pause",
    "revenue_before_pause",
    "orders_before_pause",
    "drr_before_pause",
    "gp_before_pause",
    "status",
    "next_check_date",
    "api_status",
]

DECISION_COLUMNS = [
    "campaign_id",
    "nm_id",
    "supplier_article",
    "subject_norm",
    "placement",
    "campaign_status",
    "current_bid_rub",
    "min_bid_rub",
    "drr_limit_pct",
    "avg_impressions_per_day",
    "avg_spend_per_day",
    "last21_impressions",
    "last21_spend",
    "last21_revenue",
    "last21_orders",
    "last21_drr_pct",
    "last21_avg_impressions_per_day",
    "last21_avg_spend_per_day",
    "new_bid_rub",
    "action",
    "reason_code",
    "reason_text",
    "spend",
    "revenue",
    "orders",
    "impressions",
    "clicks",
    "campaign_drr_pct",
    "cpo",
    "ctr_pct",
    "gp_after_ads",
    "previous_event_id",
    "postcheck_status",
    "last_bid_change_event_id",
    "last_bid_change_date",
    "last_bid_change_old_bid",
    "last_bid_change_new_bid",
    "last_bid_change_direction",
    "last_bid_change_reason_code",
    "wait_rule",
    "wait_until_date",
    "wait_days_left",
    "wait_status",
    "pause_decision",
    "ramp_candidate",
    "ramp_status",
    "ramp_applied_in_current_run",
    "ramp_api_status",
]

COLUMN_ALIASES: Dict[str, List[str]] = {
    "date": ["Дата", "date", "day", "День"],
    "campaign_id": ["ID кампании", "advertId", "advert_id", "campaign_id", "Кампания ID"],
    "campaign_name": ["Название", "Название кампании", "name", "campaign_name"],
    "campaign_status": ["Статус", "status", "Статус кампании"],
    "nm_id": ["nmId", "nm_id", "Номенклатура WB", "Артикул WB", "Товар"],
    "supplier_article": ["Артикул продавца", "supplier_article", "supplierArticle", "Артикул"],
    "subject_norm": ["Предмет", "subject", "subject_norm", "Название предмета"],
    "placement": ["Плейсмент", "placement", "Тип кампании", "Место размещения", "placement_norm"],
    "current_bid_rub": ["Текущая ставка, ₽", "Текущая ставка", "Ставка", "Ставка в поиске (руб)", "Ставка в рекомендациях (руб)", "bid", "cpc", "cpm"],
    "impressions": ["Показы", "views", "impressions"],
    "clicks": ["Клики", "clicks"],
    "orders": ["Заказы РК", "Заказы", "orders"],
    "spend": ["Расход", "Расходы", "Затраты", "Расход РК", "ad_spend"],
    "revenue": ["Выручка РК", "Продажи РК", "Сумма заказов", "Сумма заказов, ₽", "Заказано на сумму", "Заказано на сумму, ₽", "ordersSumRub", "sum_price", "sumPrice", "sales", "revenue", "GMV"],
    "gp_after_ads": ["ВП кампании", "Валовая прибыль после рекламы", "ВП после рекламы", "gross_profit"],
}


# =============================
# Конфигурация и S3
# =============================

@dataclass
class Config:
    yc_access_key_id: str
    yc_secret_access_key: str
    yc_bucket_name: str
    wb_promo_key: str
    s3_endpoint_url: str = "https://storage.yandexcloud.net"
    wb_base_url: str = WB_ADVERT_BASE_URL


@dataclass
class RunContext:
    mode: str
    dry_run: bool
    apply_pause: bool
    apply_start: bool
    apply_price: bool
    apply_experiment: bool
    run_datetime: datetime
    mature_end: date
    current_start: date
    current_end: date
    base_start: date
    base_end: date


def require_env(name: str) -> str:
    value = os.environ.get(name, "").strip()
    if not value:
        raise RuntimeError(f"Не задан обязательный secret/env: {name}")
    return value


def load_config() -> Config:
    return Config(
        yc_access_key_id=require_env("YC_ACCESS_KEY_ID"),
        yc_secret_access_key=require_env("YC_SECRET_ACCESS_KEY"),
        yc_bucket_name=require_env("YC_BUCKET_NAME"),
        wb_promo_key=require_env("WB_PROMO_KEY_TOPFACE"),
    )


def make_s3_client(config: Config):
    return boto3.client(
        "s3",
        endpoint_url=config.s3_endpoint_url,
        aws_access_key_id=config.yc_access_key_id,
        aws_secret_access_key=config.yc_secret_access_key,
    )


def s3_key_exists(s3_client, bucket: str, key: str) -> bool:
    try:
        s3_client.head_object(Bucket=bucket, Key=key)
        return True
    except ClientError as exc:
        code = str(exc.response.get("Error", {}).get("Code", ""))
        if code in {"404", "NoSuchKey", "NotFound"}:
            return False
        raise


def read_s3_bytes(s3_client, bucket: str, key: str) -> bytes:
    obj = s3_client.get_object(Bucket=bucket, Key=key)
    return obj["Body"].read()


def upload_s3_bytes(s3_client, bucket: str, key: str, payload: bytes, content_type: Optional[str] = None) -> None:
    extra: Dict[str, Any] = {}
    if content_type:
        extra["ContentType"] = content_type
    s3_client.put_object(Bucket=bucket, Key=key, Body=payload, **extra)


def list_s3_keys(s3_client, bucket: str, prefix: str) -> List[str]:
    keys: List[str] = []
    continuation_token: Optional[str] = None
    while True:
        kwargs: Dict[str, Any] = {"Bucket": bucket, "Prefix": prefix}
        if continuation_token:
            kwargs["ContinuationToken"] = continuation_token
        resp = s3_client.list_objects_v2(**kwargs)
        for item in resp.get("Contents", []):
            key = item.get("Key", "")
            if key:
                keys.append(key)
        if not resp.get("IsTruncated"):
            break
        continuation_token = resp.get("NextContinuationToken")
    return keys


# =============================
# Helper-функции колонок
# =============================

def _norm_col_name(value: Any) -> str:
    text = str(value).strip().lower()
    text = re.sub(r"\s+", " ", text)
    text = text.replace("ё", "е")
    return text


def find_col(df: pd.DataFrame, aliases: Iterable[str]) -> Optional[str]:
    """Возвращает имя первой найденной колонки из списка aliases или None."""
    if df is None or df.empty and len(df.columns) == 0:
        return None
    by_norm = {_norm_col_name(col): col for col in df.columns}
    for alias in aliases:
        found = by_norm.get(_norm_col_name(alias))
        if found is not None:
            return found
    return None


def series_or_default(df: pd.DataFrame, aliases: Iterable[str], default: Any = "") -> pd.Series:
    """Всегда возвращает pandas.Series длины len(df), даже если колонка отсутствует."""
    col = find_col(df, aliases)
    if col is None:
        return pd.Series([default] * len(df), index=df.index)
    return df[col]


def numeric_series(df: pd.DataFrame, aliases: Iterable[str], default: float = 0.0) -> pd.Series:
    """Возвращает числовой Series; все нечисловые значения -> default."""
    src = series_or_default(df, aliases, default=default)
    text = src.astype(str).str.replace("\u00a0", "", regex=False).str.replace(" ", "", regex=False)
    text = text.str.replace(",", ".", regex=False)
    text = text.str.replace(r"[^0-9.\-]", "", regex=True)
    num = pd.to_numeric(text, errors="coerce")
    return num.fillna(default).astype(float)


def parse_date_series(values: pd.Series) -> pd.Series:
    """Без warning разбирает даты из WB-отчётов: ISO, dd.mm.yyyy и Excel datetime."""
    if not isinstance(values, pd.Series):
        values = pd.Series(values)
    raw = values.copy()
    result = pd.Series(pd.NaT, index=raw.index, dtype="datetime64[ns]")
    text = raw.astype(str).str.strip()

    iso_mask = text.str.fullmatch(r"\d{4}-\d{2}-\d{2}", na=False)
    if iso_mask.any():
        result.loc[iso_mask] = pd.to_datetime(text.loc[iso_mask], format="%Y-%m-%d", errors="coerce")

    dot_mask = result.isna() & text.str.fullmatch(r"\d{2}\.\d{2}\.\d{4}", na=False)
    if dot_mask.any():
        result.loc[dot_mask] = pd.to_datetime(text.loc[dot_mask], format="%d.%m.%Y", errors="coerce")

    slash_mask = result.isna() & text.str.fullmatch(r"\d{2}/\d{2}/\d{4}", na=False)
    if slash_mask.any():
        result.loc[slash_mask] = pd.to_datetime(text.loc[slash_mask], format="%d/%m/%Y", errors="coerce")

    remaining = result.isna() & raw.notna() & text.ne("") & text.ne("NaT") & text.ne("nan")
    if remaining.any():
        result.loc[remaining] = pd.to_datetime(raw.loc[remaining], errors="coerce")
    return result.dt.date


def _clean_id_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "nat"}:
        return ""
    text = text.replace("\u00a0", " ").strip()
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return text


def _text_series(df: pd.DataFrame, aliases: Iterable[str], default: str = "") -> pd.Series:
    src = series_or_default(df, aliases, default=default)
    return src.map(_clean_text_value)


def _clean_text_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    text = str(value).strip()
    if text.lower() in {"nan", "none", "nat"}:
        return ""
    return re.sub(r"\s+", " ", text)


def normalize_subject_value(value: Any) -> str:
    text = _clean_text_value(value).replace("ё", "е").strip()
    return re.sub(r"\s+", " ", text)


def normalize_placement_value(value: Any) -> str:
    text = _clean_text_value(value).replace("ё", "е").lower()
    if not text:
        return ""
    if "search" in text or "поиск" in text:
        return "search"
    if "recommend" in text or "рекоменд" in text:
        return "recommendations"
    if "combined" in text or "комбин" in text or "cpm" in text:
        return "combined"
    return text


def normalize_columns(df: pd.DataFrame, source_type: str) -> pd.DataFrame:
    """Создаёт канонические поля и не удаляет исходные колонки."""
    result = df.copy()
    result["source_type"] = source_type
    result["_row_id"] = range(len(result))

    date_src = series_or_default(result, COLUMN_ALIASES["date"], default=pd.NaT)
    result["date"] = parse_date_series(date_src)

    result["campaign_id"] = series_or_default(result, COLUMN_ALIASES["campaign_id"], default="").map(_clean_id_value)
    result["campaign_name"] = _text_series(result, COLUMN_ALIASES["campaign_name"], default="")
    result["campaign_status"] = _text_series(result, COLUMN_ALIASES["campaign_status"], default="")
    result["nm_id"] = series_or_default(result, COLUMN_ALIASES["nm_id"], default="").map(_clean_id_value)
    result["supplier_article"] = _text_series(result, COLUMN_ALIASES["supplier_article"], default="")
    result["subject_norm"] = series_or_default(result, COLUMN_ALIASES["subject_norm"], default="").map(normalize_subject_value)
    result["placement"] = series_or_default(result, COLUMN_ALIASES["placement"], default="").map(normalize_placement_value)

    for metric in ["current_bid_rub", "impressions", "clicks", "orders", "spend", "revenue"]:
        result[metric] = numeric_series(result, COLUMN_ALIASES[metric], default=0.0)

    gp_col = find_col(result, COLUMN_ALIASES["gp_after_ads"])
    if gp_col is None:
        result["gp_after_ads"] = pd.Series([float("nan")] * len(result), index=result.index)
    else:
        gp_text = result[gp_col].astype(str).str.replace("\u00a0", "", regex=False).str.replace(" ", "", regex=False)
        gp_text = gp_text.str.replace(",", ".", regex=False)
        gp_text = gp_text.str.replace(r"[^0-9.\-]", "", regex=True)
        result["gp_after_ads"] = pd.to_numeric(gp_text, errors="coerce")

    return result


# =============================
# Загрузка Excel-данных
# =============================

def read_excel_bytes_as_sheets(payload: bytes) -> Dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(io.BytesIO(payload))
    return {sheet_name: pd.read_excel(io.BytesIO(payload), sheet_name=sheet_name) for sheet_name in xls.sheet_names}


def first_sheet_by_name(sheets: Dict[str, pd.DataFrame], wanted_name: str) -> pd.DataFrame:
    wanted_norm = _norm_col_name(wanted_name)
    for name, df in sheets.items():
        if _norm_col_name(name) == wanted_norm:
            return df.copy()
    return pd.DataFrame()


def read_excel_sheets_as_frame(sheets: Dict[str, pd.DataFrame], source_name: str) -> pd.DataFrame:
    frames: List[pd.DataFrame] = []
    for sheet_name, sheet_df in sheets.items():
        if sheet_df is None or sheet_df.empty:
            continue
        local_df = sheet_df.copy()
        local_df["source_file"] = source_name
        local_df["source_sheet"] = str(sheet_name)
        frames.append(local_df)
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True, sort=False)


def read_excel_bytes_as_frame(payload: bytes, source_name: str) -> pd.DataFrame:
    sheets = read_excel_bytes_as_sheets(payload)
    return read_excel_sheets_as_frame(sheets, source_name)


def derive_campaign_placement_and_bid(campaigns_norm: pd.DataFrame, campaigns_raw: pd.DataFrame) -> pd.DataFrame:
    result = campaigns_norm.copy()
    search_bid = numeric_series(campaigns_raw, [
        "Ставка в поиске (руб)", "Ставка поиск, руб", "Ставка поиск", "bid_search_rub", "search_bid",
        "Ставка в поиске", "Ставка в поиске ₽",
    ], default=0.0)
    reco_bid = numeric_series(campaigns_raw, [
        "Ставка в рекомендациях (руб)", "Ставка рекомендации, руб", "Ставка рекомендации", "bid_reco_rub",
        "reco_bid", "recommendation_bid", "Ставка в рекомендациях", "Ставка в рекомендациях ₽",
    ], default=0.0)
    direct_bid = numeric_series(campaigns_raw, COLUMN_ALIASES["current_bid_rub"], default=0.0)

    placements: List[str] = []
    bids: List[float] = []
    for idx in result.index:
        placement_raw = result.at[idx, "placement"] if "placement" in result.columns else ""
        placement = normalize_placement_value(placement_raw)
        s_bid = float(search_bid.loc[idx] if idx in search_bid.index else 0.0)
        r_bid = float(reco_bid.loc[idx] if idx in reco_bid.index else 0.0)
        d_bid = float(direct_bid.loc[idx] if idx in direct_bid.index else 0.0)

        if not placement:
            if s_bid > 0 and r_bid > 0:
                placement = "combined"
            elif s_bid > 0:
                placement = "search"
            elif r_bid > 0:
                placement = "recommendations"
            else:
                placement = "search"
        if placement == "recommendation":
            placement = "recommendations"

        if placement == "recommendations" and r_bid > 0:
            bid = r_bid
        elif placement in {"search", "combined"} and s_bid > 0:
            bid = s_bid
        elif d_bid > 0:
            bid = d_bid
        elif s_bid > 0:
            bid = s_bid
        else:
            bid = r_bid

        placements.append(placement)
        bids.append(float(bid or 0.0))
    result["placement"] = placements
    result["current_bid_rub"] = bids
    return result


def normalize_ads_analysis_sheets(sheets: Dict[str, pd.DataFrame], source_name: str) -> pd.DataFrame:
    """Специально читает Анализ рекламы.xlsx: метрики из Статистика_Ежедневно, ставки/status из Список_кампаний."""
    daily_raw = first_sheet_by_name(sheets, "Статистика_Ежедневно")
    campaigns_raw = first_sheet_by_name(sheets, "Список_кампаний")

    if daily_raw.empty:
        combined = read_excel_sheets_as_frame(sheets, source_name)
        return normalize_columns(combined, source_type="ads_generic")

    daily = normalize_columns(daily_raw, source_type="ads_daily")
    daily["source_file"] = source_name
    daily["source_sheet"] = "Статистика_Ежедневно"

    if "revenue" in daily.columns and float(pd.to_numeric(daily["revenue"], errors="coerce").fillna(0).sum()) == 0:
        sum_orders_col = find_col(daily_raw, ["Сумма заказов", "Сумма заказов, ₽", "Заказано на сумму", "Заказано на сумму, ₽"])
        if sum_orders_col:
            daily["revenue"] = numeric_series(daily_raw, [sum_orders_col], default=0.0)

    if campaigns_raw.empty:
        if daily["campaign_status"].astype(str).str.strip().eq("").all():
            daily["campaign_status"] = "Активна"
        if daily["placement"].astype(str).str.strip().eq("").all():
            daily["placement"] = "search"
        return daily

    campaigns = normalize_columns(campaigns_raw, source_type="ads_campaigns")
    campaigns = derive_campaign_placement_and_bid(campaigns, campaigns_raw)
    campaigns["source_file"] = source_name
    campaigns["source_sheet"] = "Список_кампаний"

    keep_cols = [
        "campaign_id", "nm_id", "placement", "campaign_status", "campaign_name",
        "current_bid_rub", "supplier_article", "subject_norm",
    ]
    for col in keep_cols:
        if col not in campaigns.columns:
            campaigns[col] = ""

    campaigns_dim = campaigns[keep_cols].copy()
    campaigns_dim = campaigns_dim[
        campaigns_dim["campaign_id"].map(_clean_id_value).ne("")
        & campaigns_dim["nm_id"].map(_clean_id_value).ne("")
    ].copy()
    if not campaigns_dim.empty:
        campaigns_dim["campaign_id"] = campaigns_dim["campaign_id"].map(_clean_id_value)
        campaigns_dim["nm_id"] = campaigns_dim["nm_id"].map(_clean_id_value)
        campaigns_dim["placement"] = campaigns_dim["placement"].map(normalize_placement_value).replace({"recommendation": "recommendations"})
        campaigns_dim["_status_rank"] = campaigns_dim["campaign_status"].map(lambda x: 1 if is_active_campaign(x) else 0)
        campaigns_dim = campaigns_dim.sort_values(["campaign_id", "nm_id", "_status_rank"], ascending=[True, True, False])
        campaigns_dim = campaigns_dim.drop_duplicates(["campaign_id", "nm_id", "placement"], keep="first")
        campaigns_dim = campaigns_dim.drop(columns=["_status_rank"], errors="ignore")

    daily["campaign_id"] = daily["campaign_id"].map(_clean_id_value)
    daily["nm_id"] = daily["nm_id"].map(_clean_id_value)

    if campaigns_dim.empty:
        if daily["campaign_status"].astype(str).str.strip().eq("").all():
            daily["campaign_status"] = "Активна"
        if daily["placement"].astype(str).str.strip().eq("").all():
            daily["placement"] = "search"
        return daily

    metric_cols = ["date", "campaign_id", "nm_id", "impressions", "clicks", "orders", "spend", "revenue", "gp_after_ads", "_row_id", "source_file", "source_sheet"]
    for col in ["supplier_article", "subject_norm", "campaign_name", "campaign_status", "placement", "current_bid_rub"]:
        if col not in daily.columns:
            daily[col] = "" if col != "current_bid_rub" else 0.0
    metric_cols.extend(["supplier_article", "subject_norm", "campaign_name", "campaign_status", "placement", "current_bid_rub"])
    metric_cols = [c for c in metric_cols if c in daily.columns]

    merged = daily[metric_cols].merge(
        campaigns_dim,
        on=["campaign_id", "nm_id"],
        how="left",
        suffixes=("", "_campaign"),
    )

    for col in ["placement", "campaign_status", "campaign_name", "supplier_article", "subject_norm"]:
        camp_col = f"{col}_campaign"
        if camp_col in merged.columns:
            base = merged[col].fillna("").astype(str) if col in merged.columns else pd.Series([""] * len(merged), index=merged.index)
            camp = merged[camp_col].fillna("").astype(str)
            merged[col] = base.where(base.str.strip().ne(""), camp)
    if "current_bid_rub_campaign" in merged.columns:
        base_bid = pd.to_numeric(merged.get("current_bid_rub", 0), errors="coerce").fillna(0.0)
        camp_bid = pd.to_numeric(merged["current_bid_rub_campaign"], errors="coerce").fillna(0.0)
        merged["current_bid_rub"] = base_bid.where(base_bid > 0, camp_bid)

    drop_cols = [c for c in merged.columns if c.endswith("_campaign")]
    merged = merged.drop(columns=drop_cols, errors="ignore")
    merged["placement"] = merged["placement"].map(normalize_placement_value).replace({"recommendation": "recommendations"})
    return merged


def load_ads_report(s3_client, config: Config) -> pd.DataFrame:
    if s3_key_exists(s3_client, config.yc_bucket_name, ADS_MAIN_KEY):
        payload = read_s3_bytes(s3_client, config.yc_bucket_name, ADS_MAIN_KEY)
        sheets = read_excel_bytes_as_sheets(payload)
        raw_all = normalize_ads_analysis_sheets(sheets, ADS_MAIN_KEY)
        if raw_all.empty:
            raise RuntimeError(f"Основной рекламный отчёт пустой: {ADS_MAIN_KEY}")
        before_subjects = raw_all.get("subject_norm", pd.Series(dtype=str)).map(str).value_counts().head(20).to_dict()
        raw = filter_managed_subject_rows(raw_all)
        if raw.empty:
            raise RuntimeError(
                "После фильтра 4 управляемых предметов не осталось строк. "
                f"Проверь поле subject_norm/Предмет в {ADS_MAIN_KEY}. Предметы в файле: {before_subjects}"
            )
        after_subjects = raw.get("subject_norm", pd.Series(dtype=str)).map(str).value_counts().to_dict()
        print(
            "Диагностика загрузки рекламы: "
            f"листы={list(sheets.keys())}; "
            f"строк после нормализации={len(raw_all)}; "
            f"строк после фильтра 4 предметов={len(raw)}; "
            f"предметы после фильтра={json.dumps(after_subjects, ensure_ascii=False)}; "
            f"валидных campaign_id={raw['campaign_id'].map(_clean_id_value).ne('').sum() if 'campaign_id' in raw.columns else 0}; "
            f"валидных nm_id={raw['nm_id'].map(_clean_id_value).ne('').sum() if 'nm_id' in raw.columns else 0}; "
            f"валидных placement={raw['placement'].astype(str).str.strip().ne('').sum() if 'placement' in raw.columns else 0}; "
            f"валидных ставок={(pd.to_numeric(raw['current_bid_rub'], errors='coerce').fillna(0) > 0).sum() if 'current_bid_rub' in raw.columns else 0}; "
            f"активных={raw['campaign_status'].map(is_active_campaign).sum() if 'campaign_status' in raw.columns else 0}",
            flush=True,
        )
        return raw

    weekly_keys = [
        key for key in list_s3_keys(s3_client, config.yc_bucket_name, ADS_WEEKLY_PREFIX)
        if key.lower().endswith(".xlsx") and not key.endswith("/~$")
    ]
    weekly_keys = sorted(weekly_keys, reverse=True)[:8]
    frames: List[pd.DataFrame] = []
    for key in weekly_keys:
        payload = read_s3_bytes(s3_client, config.yc_bucket_name, key)
        sheets = read_excel_bytes_as_sheets(payload)
        raw = normalize_ads_analysis_sheets(sheets, key)
        if not raw.empty:
            frames.append(raw)
    if not frames:
        raise RuntimeError(
            "Не найден основной рекламный отчёт и нет непустых fallback-файлов: "
            f"{ADS_MAIN_KEY}; {ADS_WEEKLY_PREFIX}"
        )
    result_all = pd.concat(frames, ignore_index=True, sort=False)
    result = filter_managed_subject_rows(result_all)
    print(
        f"Диагностика fallback-рекламы: файлов={len(frames)}, "
        f"строк до фильтра={len(result_all)}, строк после фильтра 4 предметов={len(result)}",
        flush=True,
    )
    if result.empty:
        raise RuntimeError("Fallback-реклама найдена, но после фильтра 4 управляемых предметов строк не осталось")
    return result


def coerce_history_columns_object(df: pd.DataFrame, columns: List[str]) -> pd.DataFrame:
    """
    Pandas 3.x не разрешает записывать текстовый verdict/status в колонку,
    которую Excel прочитал как float64 из-за пустых ячеек.
    Исторические журналы содержат смешанные поля: числа, даты, статусы, ответы API.
    Для безопасного дозаполнения приводим колонки журнала к object.
    Числовые расчёты ниже всё равно выполняются через pd.to_numeric(...).
    """
    if df is None or not isinstance(df, pd.DataFrame):
        return pd.DataFrame(columns=columns).astype({c: "object" for c in columns})
    out = df.copy()
    for col in columns:
        if col not in out.columns:
            out[col] = pd.Series([""] * len(out), index=out.index, dtype="object")
        else:
            out[col] = out[col].astype("object")
    return out


def load_excel_table_from_s3(s3_client, config: Config, key: str, columns: List[str]) -> pd.DataFrame:
    if not s3_key_exists(s3_client, config.yc_bucket_name, key):
        return coerce_history_columns_object(pd.DataFrame(columns=columns), columns)
    payload = read_s3_bytes(s3_client, config.yc_bucket_name, key)
    try:
        df = pd.read_excel(io.BytesIO(payload))
    except Exception:
        return coerce_history_columns_object(pd.DataFrame(columns=columns), columns)
    return coerce_history_columns_object(df, columns)


def load_bid_history(s3_client, config: Config) -> pd.DataFrame:
    return load_excel_table_from_s3(s3_client, config, BID_HISTORY_KEY, BID_HISTORY_COLUMNS)


def load_pause_history(s3_client, config: Config) -> pd.DataFrame:
    return load_excel_table_from_s3(s3_client, config, PAUSE_HISTORY_KEY, PAUSE_HISTORY_COLUMNS)



# =============================
# Мониторинг ключевых фраз для оценки эффекта ставок
# =============================

def _keyword_numeric(df: pd.DataFrame, aliases: Iterable[str], default: float = 0.0) -> pd.Series:
    return numeric_series(df, aliases, default=default)


def _keyword_text(df: pd.DataFrame, aliases: Iterable[str], default: str = "") -> pd.Series:
    return _text_series(df, aliases, default=default)


def normalize_keyword_positions(df: pd.DataFrame, source_name: str, sheet_name: str) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=KEYWORD_POSITION_COLUMNS)
    result = pd.DataFrame(index=df.index)
    result["source_file"] = source_name
    result["source_sheet"] = sheet_name
    result["date"] = parse_date_series(series_or_default(df, ["Дата", "date", "День", "day"], default=pd.NaT))
    result["nm_id"] = series_or_default(df, ["Артикул WB", "nmID", "nmId", "nm_id", "Номенклатура WB"], default="").map(_clean_id_value)
    result["supplier_article"] = _keyword_text(df, ["Артикул продавца", "supplierArticle", "Артикул", "vendorCode"], default="")
    result["subject_norm"] = series_or_default(df, ["Предмет", "Название предмета", "subject", "subject_norm"], default="").map(normalize_subject_value)
    result["query_text"] = _keyword_text(df, ["Поисковый запрос", "Запрос", "Ключевая фраза", "Ключевая фраза/слово", "keyword", "query", "search_query"], default="")
    result["filter_type"] = _keyword_text(df, ["Фильтр", "filter", "Тип", "type"], default="")
    result["query_freq"] = _keyword_numeric(df, ["Частота запросов", "Частота за неделю", "Частотность", "query_freq", "demand_week"], default=0.0)
    result["median_position"] = _keyword_numeric(df, ["Медианная позиция", "Позиция", "median_position", "position"], default=0.0)
    result["visibility_pct"] = _keyword_numeric(df, ["Видимость %", "Видимость, %", "visibility_pct", "visibility"], default=0.0)
    result["clicks_to_card"] = _keyword_numeric(df, ["Переходы в карточку", "Переходы", "clicks_to_card", "clicks"], default=0.0)
    result["keyword_orders"] = _keyword_numeric(df, ["Заказы", "Заказы по запросу", "keyword_orders", "orders"], default=0.0)
    result["query_text_norm"] = result["query_text"].astype(str).str.strip().str.lower().str.replace("ё", "е", regex=False)
    result = result[result["nm_id"].map(_clean_id_value).ne("") & result["query_text_norm"].ne("")].copy()
    result = filter_managed_subject_rows(result)
    return result


def load_keyword_positions(s3_client, config: Config) -> pd.DataFrame:
    keys = [k for k in list_s3_keys(s3_client, config.yc_bucket_name, KEYWORDS_WEEKLY_PREFIX) if k.lower().endswith(".xlsx")]
    keys = sorted(keys, reverse=True)[:8]
    frames: List[pd.DataFrame] = []
    for key in keys:
        try:
            payload = read_s3_bytes(s3_client, config.yc_bucket_name, key)
            sheets = read_excel_bytes_as_sheets(payload)
            target_sheets = []
            for sh_name, sh_df in sheets.items():
                sh_norm = _norm_col_name(sh_name)
                if "позиции" in sh_norm or "ключ" in sh_norm or "запрос" in sh_norm:
                    target_sheets.append((sh_name, sh_df))
            if not target_sheets and sheets:
                first_name = next(iter(sheets.keys()))
                target_sheets = [(first_name, sheets[first_name])]
            for sh_name, sh_df in target_sheets:
                norm = normalize_keyword_positions(sh_df, key, str(sh_name))
                if not norm.empty:
                    frames.append(norm)
        except Exception as exc:
            print(f"Предупреждение: не удалось прочитать поисковые запросы {key}: {exc}", flush=True)
            continue
    if not frames:
        return pd.DataFrame(columns=KEYWORD_POSITION_COLUMNS)
    result = pd.concat(frames, ignore_index=True, sort=False)
    result = result.drop_duplicates(subset=["date", "nm_id", "supplier_article", "query_text_norm", "filter_type"], keep="last")
    return result


def classify_core_keywords(keyword_df: pd.DataFrame, ctx: RunContext) -> pd.DataFrame:
    if keyword_df is None or keyword_df.empty:
        return pd.DataFrame(columns=KEYWORD_POSITION_COLUMNS)
    local = keyword_df.copy()
    if has_valid_dates(local):
        # Берём базу классификации: последние зрелые данные до текущего mature_end.
        local = local[local["date"].notna() & (local["date"] <= ctx.mature_end)].copy()
    if local.empty:
        return pd.DataFrame(columns=KEYWORD_POSITION_COLUMNS)
    grp = local.groupby(["nm_id", "supplier_article", "subject_norm", "query_text_norm"], dropna=False).agg(
        query_text=("query_text", "last"),
        keyword_orders=("keyword_orders", "sum"),
        query_freq=("query_freq", "sum"),
        median_position=("median_position", "median"),
        visibility_pct=("visibility_pct", "mean"),
        clicks_to_card=("clicks_to_card", "sum"),
    ).reset_index()
    grp["keyword_group"] = "TAIL_20"
    grp["orders_share"] = 0.0
    grp["orders_cum_share"] = 0.0
    out_frames: List[pd.DataFrame] = []
    for _, g in grp.groupby(["nm_id", "supplier_article"], dropna=False):
        g = g.sort_values(["keyword_orders", "query_freq", "clicks_to_card"], ascending=False).copy()
        total_orders = float(g["keyword_orders"].sum())
        if total_orders > 0:
            g["orders_share"] = g["keyword_orders"] / total_orders
            g["orders_cum_share"] = g["orders_share"].cumsum()
            # CORE_80: фразы до достижения 80% заказов + первая фраза, которая пересекла порог.
            g["keyword_group"] = "TAIL_20"
            core_mask = g["orders_cum_share"] <= 0.80
            if not core_mask.any() and len(g) > 0:
                core_mask.iloc[0] = True
            elif core_mask.any():
                first_tail_idx = g.index[~core_mask]
                if len(first_tail_idx) > 0:
                    core_mask.loc[first_tail_idx[0]] = True
            g.loc[core_mask, "keyword_group"] = "CORE_80"
        else:
            # Если заказов нет, CORE_80 не назначаем, чтобы не имитировать продающие запросы.
            g["keyword_group"] = "TAIL_20"
        out_frames.append(g)
    result = pd.concat(out_frames, ignore_index=True, sort=False)
    return result


def _keyword_window_agg(keyword_df: pd.DataFrame, core_map: pd.DataFrame, nm_id: str, supplier_article: str, start_date: date, end_date: date, group_name: str) -> Dict[str, float]:
    empty = {"queries_count": 0.0, "position": 0.0, "visibility": 0.0, "keyword_orders": 0.0, "query_freq": 0.0}
    if keyword_df is None or keyword_df.empty or core_map is None or core_map.empty:
        return empty
    core = core_map[(core_map["nm_id"].astype(str) == str(nm_id)) & (core_map["keyword_group"] == group_name)].copy()
    if supplier_article:
        core_art = core[core["supplier_article"].astype(str) == str(supplier_article)]
        if not core_art.empty:
            core = core_art
    if core.empty:
        return empty
    queries = set(core["query_text_norm"].astype(str).tolist())
    part = keyword_df[(keyword_df["nm_id"].astype(str) == str(nm_id)) & (keyword_df["query_text_norm"].astype(str).isin(queries))].copy()
    if supplier_article:
        part_art = part[part["supplier_article"].astype(str) == str(supplier_article)]
        if not part_art.empty:
            part = part_art
    if has_valid_dates(part):
        part = part[(part["date"] >= start_date) & (part["date"] <= end_date)].copy()
    if part.empty:
        return {"queries_count": float(len(queries)), "position": 0.0, "visibility": 0.0, "keyword_orders": 0.0, "query_freq": 0.0}
    # Для позиции меньше = лучше. Берём медиану по фразам/дням.
    pos = pd.to_numeric(part["median_position"], errors="coerce")
    pos = pos[pos > 0]
    return {
        "queries_count": float(len(queries)),
        "position": float(pos.median()) if not pos.empty else 0.0,
        "visibility": float(pd.to_numeric(part["visibility_pct"], errors="coerce").fillna(0).mean()),
        "keyword_orders": float(pd.to_numeric(part["keyword_orders"], errors="coerce").fillna(0).sum()),
        "query_freq": float(pd.to_numeric(part["query_freq"], errors="coerce").fillna(0).sum()),
    }


def _position_delta(after: float, before: float) -> float:
    before = float(before or 0)
    after = float(after or 0)
    if before == 0 or after == 0:
        return 0.0
    # Положительное значение = позиция улучшилась, потому что число позиции уменьшилось.
    return before - after


def _keyword_verdict(direction: str, before: Dict[str, float], after: Dict[str, float], horizon: str) -> Tuple[str, str]:
    if before.get("queries_count", 0) <= 0:
        return "NO_CORE_KEYWORDS", "нет классифицированных продающих фраз"
    pos_delta = _position_delta(after.get("position", 0), before.get("position", 0))
    vis_delta = float(after.get("visibility", 0) or 0) - float(before.get("visibility", 0) or 0)
    orders_before = float(before.get("keyword_orders", 0) or 0)
    orders_after = float(after.get("keyword_orders", 0) or 0)
    direction = str(direction or "").lower()
    if direction == "raise":
        if pos_delta > 0 or vis_delta >= 2 or (orders_before == 0 and orders_after > 0) or orders_after >= orders_before:
            return f"RAISE_KEYWORDS_OK_{horizon}", f"позиция Δ={pos_delta:.2f}, видимость Δ={vis_delta:.2f} п.п., заказы {orders_before:.0f}->{orders_after:.0f}"
        return f"RAISE_KEYWORDS_WEAK_{horizon}", f"нет улучшения CORE_80: позиция Δ={pos_delta:.2f}, видимость Δ={vis_delta:.2f} п.п."
    if direction == "lower":
        if pos_delta < -3 or vis_delta <= -5 or (orders_before > 0 and orders_after < orders_before * 0.80):
            return f"LOWER_KEYWORDS_RISK_{horizon}", f"просадка CORE_80: позиция Δ={pos_delta:.2f}, видимость Δ={vis_delta:.2f} п.п., заказы {orders_before:.0f}->{orders_after:.0f}"
        return f"LOWER_KEYWORDS_OK_{horizon}", f"CORE_80 без критичной просадки: позиция Δ={pos_delta:.2f}, видимость Δ={vis_delta:.2f} п.п."
    return f"KEYWORDS_MONITOR_{horizon}", f"позиция Δ={pos_delta:.2f}, видимость Δ={vis_delta:.2f} п.п."


def build_keyword_effects(bid_history: pd.DataFrame, keyword_df: pd.DataFrame, core_map: pd.DataFrame, ctx: RunContext) -> pd.DataFrame:
    if bid_history is None or bid_history.empty:
        return pd.DataFrame(columns=KEYWORD_EFFECT_COLUMNS)
    rows: List[Dict[str, Any]] = []
    for _, event in bid_history.iterrows():
        event_date = pd.to_datetime(event.get("event_date"), errors="coerce")
        if pd.isna(event_date):
            continue
        event_day = event_date.date()
        nm_id = _clean_id_value(event.get("nm_id", ""))
        supplier_article = _clean_text_value(event.get("supplier_article", ""))
        direction = _clean_text_value(event.get("direction", "")).lower()
        before_start = event_day - timedelta(days=5)
        before_end = event_day - timedelta(days=1)
        d1_day = event_day + timedelta(days=1)
        d3_start = event_day + timedelta(days=1)
        d3_end = event_day + timedelta(days=3)
        for group_name in ["CORE_80", "TAIL_20"]:
            before = _keyword_window_agg(keyword_df, core_map, nm_id, supplier_article, before_start, before_end, group_name)
            d1 = _keyword_window_agg(keyword_df, core_map, nm_id, supplier_article, d1_day, d1_day, group_name) if ctx.mature_end >= d1_day else {"queries_count": before.get("queries_count", 0), "position": 0, "visibility": 0, "keyword_orders": 0, "query_freq": 0}
            d3 = _keyword_window_agg(keyword_df, core_map, nm_id, supplier_article, d3_start, d3_end, group_name) if ctx.mature_end >= d3_end else {"queries_count": before.get("queries_count", 0), "position": 0, "visibility": 0, "keyword_orders": 0, "query_freq": 0}
            v1, c1 = _keyword_verdict(direction, before, d1, "D1") if ctx.mature_end >= d1_day else ("WAIT_D1", "зрелый D+1 ещё не доступен")
            v3, c3 = _keyword_verdict(direction, before, d3, "D3") if ctx.mature_end >= d3_end else ("WAIT_D3", "зрелый D+3 ещё не доступен")
            rows.append({
                "event_id": event.get("event_id", ""),
                "campaign_id": event.get("campaign_id", ""),
                "nm_id": nm_id,
                "supplier_article": supplier_article,
                "placement": event.get("placement", ""),
                "direction": direction,
                "event_date": event.get("event_date", ""),
                "keyword_group": group_name,
                "queries_count": int(before.get("queries_count", 0) or 0),
                "position_before": before.get("position", 0),
                "position_d1": d1.get("position", 0),
                "position_d3": d3.get("position", 0),
                "position_delta_d1": _position_delta(d1.get("position", 0), before.get("position", 0)),
                "position_delta_d3": _position_delta(d3.get("position", 0), before.get("position", 0)),
                "visibility_before": before.get("visibility", 0),
                "visibility_d1": d1.get("visibility", 0),
                "visibility_d3": d3.get("visibility", 0),
                "visibility_delta_d1": float(d1.get("visibility", 0) or 0) - float(before.get("visibility", 0) or 0),
                "visibility_delta_d3": float(d3.get("visibility", 0) or 0) - float(before.get("visibility", 0) or 0),
                "keyword_orders_before": before.get("keyword_orders", 0),
                "keyword_orders_d1": d1.get("keyword_orders", 0),
                "keyword_orders_d3": d3.get("keyword_orders", 0),
                "query_freq_before": before.get("query_freq", 0),
                "query_freq_d1": d1.get("query_freq", 0),
                "query_freq_d3": d3.get("query_freq", 0),
                "keyword_verdict_d1": v1,
                "keyword_verdict_d3": v3,
                "keyword_comment": c3 if not str(v3).startswith("WAIT") else c1,
            })
    return pd.DataFrame(rows, columns=KEYWORD_EFFECT_COLUMNS)


def enrich_effects_with_keyword_monitoring(effect_df: pd.DataFrame, keyword_effects: pd.DataFrame) -> pd.DataFrame:
    if effect_df is None or effect_df.empty or keyword_effects is None or keyword_effects.empty:
        return effect_df if effect_df is not None else pd.DataFrame()
    core = keyword_effects[keyword_effects["keyword_group"] == "CORE_80"].copy()
    if core.empty:
        return effect_df
    cols = [
        "event_id", "queries_count", "position_before", "position_d1", "position_d3", "position_delta_d1", "position_delta_d3",
        "visibility_before", "visibility_d1", "visibility_d3", "visibility_delta_d1", "visibility_delta_d3",
        "keyword_orders_before", "keyword_orders_d1", "keyword_orders_d3", "keyword_verdict_d1", "keyword_verdict_d3", "keyword_comment",
    ]
    core = core[cols].drop_duplicates(subset=["event_id"], keep="last")
    renamed = {c: f"core80_{c}" for c in cols if c != "event_id"}
    core = core.rename(columns=renamed)
    return effect_df.merge(core, on="event_id", how="left")


# =============================
# Контур изменения цены через скидку продавца
# =============================

def load_price_history(s3_client, config: Config) -> pd.DataFrame:
    return load_excel_table_from_s3(s3_client, config, PRICE_HISTORY_KEY, PRICE_HISTORY_COLUMNS)


def normalize_funnel_report(df: pd.DataFrame) -> pd.DataFrame:
    """
    Нормализует файл воронки.

    Важно: воронку НЕ фильтруем по subject_norm внутри этой функции.
    В реальном файле воронки предмета может не быть или он может называться иначе.
    Если здесь применить filter_managed_subject_rows(), строки с пустым subject_norm
    полностью исчезают, и ценовой блок ошибочно уходит в режим funnel_missing=True.
    Предмет подтягивается позже через nm_id из рекламных метрик / price API.
    """
    if df is None or df.empty:
        return pd.DataFrame()
    result = pd.DataFrame(index=df.index)
    result["date"] = parse_date_series(series_or_default(df, ["Дата", "date", "dt", "day", "День"], default=pd.NaT))
    result["nm_id"] = series_or_default(df, ["Артикул WB", "Номенклатура WB", "nmID", "nmId", "nm_id", "nm", "НМ"], default="").map(_clean_id_value)
    result["supplier_article"] = _text_series(df, ["Артикул продавца", "supplierArticle", "supplier_article", "Артикул", "vendorCode", "vendor_code"], default="")
    result["subject_norm"] = series_or_default(df, ["Предмет", "Название предмета", "subject", "subject_norm"], default="").map(normalize_subject_value)
    result["card_views"] = numeric_series(df, [
        "Переходы в карточку", "Переходы", "Просмотры карточки", "Просмотры карточек",
        "Карточку посмотрели", "openCardCount", "open_card_count", "openCard", "open_card"
    ], default=0.0)
    result["add_to_cart"] = numeric_series(df, [
        "Добавления в корзину", "Добавили в корзину", "Корзины", "В корзину",
        "addToCartCount", "add_to_cart_count", "addToCart", "add_to_cart"
    ], default=0.0)
    result["funnel_orders"] = numeric_series(df, [
        "Заказы", "Заказали", "Заказано", "ordersCount", "orders_count", "orders"
    ], default=0.0)
    # Если конверсии есть в отчёте, используем их; иначе считаем ниже.
    result["add_to_cart_conv"] = numeric_series(df, [
        "Конверсия в корзину", "Конверсия в корзину %", "Конверсия корзины",
        "addToCartConversion", "add_to_cart_conversion"
    ], default=0.0)
    result["cart_to_order_conv"] = numeric_series(df, [
        "Конверсия в заказ", "Конверсия корзина-заказ", "Конверсия корзины в заказ",
        "cartToOrderConversion", "cart_to_order_conversion"
    ], default=0.0)
    result = result[result["nm_id"].map(_clean_id_value).ne("")].copy()
    result["add_to_cart_conv"] = result.apply(lambda r: safe_ctr_pct(r.get("add_to_cart", 0), r.get("card_views", 0)) if float(r.get("add_to_cart_conv", 0) or 0) == 0 else float(r.get("add_to_cart_conv", 0) or 0), axis=1)
    result["cart_to_order_conv"] = result.apply(lambda r: safe_ctr_pct(r.get("funnel_orders", 0), r.get("add_to_cart", 0)) if float(r.get("cart_to_order_conv", 0) or 0) == 0 else float(r.get("cart_to_order_conv", 0) or 0), axis=1)
    return result


def load_funnel_report(s3_client, config: Config) -> pd.DataFrame:
    if not s3_key_exists(s3_client, config.yc_bucket_name, FUNNEL_KEY):
        print(f"Диагностика воронки: файл не найден: {FUNNEL_KEY}", flush=True)
        return pd.DataFrame()
    try:
        payload = read_s3_bytes(s3_client, config.yc_bucket_name, FUNNEL_KEY)
        sheets = read_excel_bytes_as_sheets(payload)
        frames: List[pd.DataFrame] = []
        diag_rows: List[Dict[str, Any]] = []
        for sheet_name, df in sheets.items():
            raw_rows = 0 if df is None else len(df)
            norm = normalize_funnel_report(df)
            diag_rows.append({
                "sheet": sheet_name,
                "raw_rows": raw_rows,
                "normalized_rows": len(norm),
                "rows_with_nm_id": int(norm["nm_id"].map(_clean_id_value).ne("").sum()) if not norm.empty and "nm_id" in norm.columns else 0,
                "rows_with_date": int(norm["date"].notna().sum()) if not norm.empty and "date" in norm.columns else 0,
                "card_views_sum": float(pd.to_numeric(norm.get("card_views", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()) if not norm.empty else 0.0,
                "add_to_cart_sum": float(pd.to_numeric(norm.get("add_to_cart", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()) if not norm.empty else 0.0,
                "orders_sum": float(pd.to_numeric(norm.get("funnel_orders", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()) if not norm.empty else 0.0,
            })
            if not norm.empty:
                norm["source_sheet"] = sheet_name
                frames.append(norm)
        if diag_rows:
            total_norm = sum(int(r["normalized_rows"]) for r in diag_rows)
            total_views = sum(float(r["card_views_sum"]) for r in diag_rows)
            total_cart = sum(float(r["add_to_cart_sum"]) for r in diag_rows)
            total_orders = sum(float(r["orders_sum"]) for r in diag_rows)
            print(
                f"Диагностика воронки: файл найден; листов={len(diag_rows)}; строк после нормализации={total_norm}; "
                f"переходы={total_views:.0f}; корзины={total_cart:.0f}; заказы={total_orders:.0f}",
                flush=True,
            )
        return pd.concat(frames, ignore_index=True, sort=False) if frames else pd.DataFrame()
    except Exception as exc:
        print(f"Предупреждение: не удалось прочитать воронку {FUNNEL_KEY}: {exc}", flush=True)
        return pd.DataFrame()




# =============================
# Экономика: условная ВП после рекламы для диагностики ставок
# =============================

def normalize_economics_report(df: pd.DataFrame, source_sheet: str = "") -> pd.DataFrame:
    """Нормализует файл Экономика.xlsx.

    Для управления ставками нужна не точная бухгалтерская прибыль, а устойчивые параметры юнит-экономики:
    комиссия WB %, эквайринг %, НДС/ед, средняя логистика/ед, себестоимость/ед и средняя цена.
    Если точного артикула нет, ниже используется fallback по товарной группе: 901/5 -> среднее по 901.
    """
    if df is None or df.empty:
        return pd.DataFrame()
    result = pd.DataFrame(index=df.index)
    result["source_sheet"] = source_sheet
    result["week"] = _text_series(df, ["Неделя", "week"], default="")
    result["nm_id"] = series_or_default(df, ["Артикул WB", "nmID", "nmId", "nm_id", "Номенклатура WB"], default="").map(_clean_id_value)
    result["supplier_article"] = _text_series(df, ["Артикул продавца", "supplierArticle", "supplier_article", "Артикул", "vendorCode"], default="")
    result["subject_norm"] = series_or_default(df, ["Предмет", "Название предмета", "subject", "subject_norm"], default="").map(normalize_subject_value)
    result["sales_qty"] = numeric_series(df, ["Чистые продажи, шт", "Продажи, шт", "sales_qty", "quantity"], default=0.0)

    gross_revenue = numeric_series(df, ["Валовая выручка", "Выручка", "revenue"], default=0.0)
    avg_price = numeric_series(df, ["Средняя цена продажи", "avg_price", "Средняя цена"], default=0.0)
    result["avg_price"] = avg_price
    mask_price = (result["avg_price"] <= 0) & (result["sales_qty"] > 0) & (gross_revenue > 0)
    result.loc[mask_price, "avg_price"] = gross_revenue.loc[mask_price] / result.loc[mask_price, "sales_qty"]

    commission_pct = numeric_series(df, ["Комиссия WB, %", "Комиссия WB %", "commission_pct"], default=0.0)
    commission_total = numeric_series(df, ["Комиссия WB", "commission"], default=0.0)
    result["commission_pct"] = commission_pct
    mask_comm = (result["commission_pct"] <= 0) & (gross_revenue > 0) & (commission_total > 0)
    result.loc[mask_comm, "commission_pct"] = commission_total.loc[mask_comm] / gross_revenue.loc[mask_comm] * 100.0

    acquiring_pct = numeric_series(df, ["Эквайринг, %", "Эквайринг %", "acquiring_pct"], default=0.0)
    acquiring_total = numeric_series(df, ["Эквайринг", "acquiring"], default=0.0)
    result["acquiring_pct"] = acquiring_pct
    mask_acq = (result["acquiring_pct"] <= 0) & (gross_revenue > 0) & (acquiring_total > 0)
    result.loc[mask_acq, "acquiring_pct"] = acquiring_total.loc[mask_acq] / gross_revenue.loc[mask_acq] * 100.0

    vat_per_unit = numeric_series(df, ["НДС, руб/ед", "НДС руб/ед", "vat_per_unit"], default=0.0)
    vat_total = numeric_series(df, ["НДС", "vat"], default=0.0)
    result["vat_per_unit"] = vat_per_unit
    mask_vat = (result["vat_per_unit"] <= 0) & (result["sales_qty"] > 0) & (vat_total > 0)
    result.loc[mask_vat, "vat_per_unit"] = vat_total.loc[mask_vat] / result.loc[mask_vat, "sales_qty"]

    logistics_direct_unit = numeric_series(df, ["Логистика прямая, руб/ед", "Логистика прямая руб/ед"], default=0.0)
    logistics_return_unit = numeric_series(df, ["Логистика обратная, руб/ед", "Логистика обратная руб/ед"], default=0.0)
    logistics_direct_total = numeric_series(df, ["Логистика прямая"], default=0.0)
    logistics_return_total = numeric_series(df, ["Логистика обратная"], default=0.0)
    result["logistics_per_unit"] = logistics_direct_unit + logistics_return_unit
    mask_log = (result["logistics_per_unit"] <= 0) & (result["sales_qty"] > 0) & ((logistics_direct_total + logistics_return_total) > 0)
    result.loc[mask_log, "logistics_per_unit"] = (logistics_direct_total.loc[mask_log] + logistics_return_total.loc[mask_log]) / result.loc[mask_log, "sales_qty"]

    cogs_per_unit = numeric_series(df, ["Себестоимость, руб", "Себестоимость", "cogs_per_unit"], default=0.0)
    cogs_total = numeric_series(df, ["Себестоимость всего", "cogs_total"], default=0.0)
    result["cogs_per_unit"] = cogs_per_unit
    mask_cogs = (result["cogs_per_unit"] <= 0) & (result["sales_qty"] > 0) & (cogs_total > 0)
    result.loc[mask_cogs, "cogs_per_unit"] = cogs_total.loc[mask_cogs] / result.loc[mask_cogs, "sales_qty"]

    result["product_group"] = result["supplier_article"].map(product_group_from_article)
    result = result[(result["nm_id"].ne("") | result["supplier_article"].map(_clean_text_value).ne(""))].copy()
    for col in ["avg_price", "commission_pct", "acquiring_pct", "vat_per_unit", "logistics_per_unit", "cogs_per_unit", "sales_qty"]:
        result[col] = pd.to_numeric(result[col], errors="coerce").fillna(0.0)
    return result


def load_economics_report(s3_client, config: Config) -> pd.DataFrame:
    if not s3_key_exists(s3_client, config.yc_bucket_name, ECONOMICS_KEY):
        print(f"Диагностика экономики: файл не найден: {ECONOMICS_KEY}", flush=True)
        return pd.DataFrame()
    try:
        payload = read_s3_bytes(s3_client, config.yc_bucket_name, ECONOMICS_KEY)
        sheets = read_excel_bytes_as_sheets(payload)
        frames: List[pd.DataFrame] = []
        for sheet_name, df in sheets.items():
            sh_norm = _norm_col_name(sheet_name)
            if not ("юнит" in sh_norm or "общий факт" in sh_norm):
                continue
            norm = normalize_economics_report(df, source_sheet=str(sheet_name))
            if not norm.empty:
                frames.append(norm)
        result = pd.concat(frames, ignore_index=True, sort=False) if frames else pd.DataFrame()
        if not result.empty:
            print(
                "Диагностика экономики: "
                f"файл найден; листов использовано={len(frames)}; строк={len(result)}; "
                f"nm_id={result['nm_id'].map(_clean_id_value).ne('').sum()}; "
                f"групп={result['product_group'].map(_clean_text_value).ne('').sum()}",
                flush=True,
            )
        else:
            print("Диагностика экономики: файл найден, но нужные листы/строки не распознаны", flush=True)
        return result
    except Exception as exc:
        print(f"Предупреждение: не удалось прочитать экономику {ECONOMICS_KEY}: {exc}", flush=True)
        return pd.DataFrame()


def _weighted_mean_numeric(df: pd.DataFrame, value_col: str, weight_col: str = "sales_qty") -> float:
    if df is None or df.empty or value_col not in df.columns:
        return 0.0
    vals = pd.to_numeric(df[value_col], errors="coerce")
    weights = pd.to_numeric(df.get(weight_col, pd.Series([1.0] * len(df), index=df.index)), errors="coerce").fillna(0.0)
    mask = vals.notna() & (weights > 0)
    if mask.any() and float(weights.loc[mask].sum()) > 0:
        return float((vals.loc[mask] * weights.loc[mask]).sum() / weights.loc[mask].sum())
    vals = vals.dropna()
    return float(vals.mean()) if not vals.empty else 0.0


def _economics_metric_from_group(g: pd.DataFrame, method: str) -> Dict[str, Any]:
    if g is None or g.empty:
        return {}
    supplier_article = _clean_text_value(g["supplier_article"].dropna().astype(str).iloc[0]) if "supplier_article" in g.columns and not g["supplier_article"].dropna().empty else ""
    product_group = _clean_text_value(g["product_group"].dropna().astype(str).iloc[0]) if "product_group" in g.columns and not g["product_group"].dropna().empty else product_group_from_article(supplier_article)
    return {
        "economics_match_method": method,
        "supplier_article_from_economics": normalize_article_for_campaign_name(supplier_article) or supplier_article,
        "economics_product_group": product_group,
        "economics_avg_price": _weighted_mean_numeric(g, "avg_price"),
        "economics_commission_pct": _weighted_mean_numeric(g, "commission_pct"),
        "economics_acquiring_pct": _weighted_mean_numeric(g, "acquiring_pct"),
        "economics_vat_per_unit": _weighted_mean_numeric(g, "vat_per_unit"),
        "economics_logistics_per_unit": _weighted_mean_numeric(g, "logistics_per_unit"),
        "economics_cogs_per_unit": _weighted_mean_numeric(g, "cogs_per_unit") if ECONOMICS_SUBTRACT_COGS else 0.0,
    }


def build_economics_lookup(economics_df: pd.DataFrame) -> Dict[str, Dict[str, Dict[str, Any]]]:
    lookup: Dict[str, Dict[str, Dict[str, Any]]] = {"nm": {}, "article": {}, "group": {}}
    if economics_df is None or economics_df.empty:
        return lookup
    local = economics_df.copy()
    # Обычно файл недельный. Берём последнюю неделю, чтобы не усреднять старые условия комиссии/логистики.
    if "week" in local.columns and local["week"].map(_clean_text_value).ne("").any():
        latest_week = sorted(local["week"].map(_clean_text_value).dropna().unique())[-1]
        local = local[local["week"].map(_clean_text_value).eq(latest_week)].copy()
    for nm_id, g in local[local["nm_id"].map(_clean_id_value).ne("")].groupby("nm_id", dropna=False):
        lookup["nm"][_clean_id_value(nm_id)] = _economics_metric_from_group(g, "exact_nm_id")
    tmp = local.copy()
    tmp["article_norm"] = tmp["supplier_article"].map(normalize_article_for_campaign_name)
    for art, g in tmp[tmp["article_norm"].map(_clean_text_value).ne("")].groupby("article_norm", dropna=False):
        lookup["article"][_clean_text_value(art)] = _economics_metric_from_group(g, "exact_supplier_article")
    for grp, g in local[local["product_group"].map(_clean_text_value).ne("")].groupby("product_group", dropna=False):
        lookup["group"][_clean_text_value(grp)] = _economics_metric_from_group(g, "avg_product_group")
    return lookup


def lookup_economics_metrics(economics_lookup: Dict[str, Dict[str, Dict[str, Any]]], nm_id: Any, supplier_article: Any) -> Dict[str, Any]:
    nm = _clean_id_value(nm_id)
    art = normalize_article_for_campaign_name(supplier_article) or _clean_text_value(supplier_article)
    grp = product_group_from_article(art or supplier_article)
    if nm and nm in economics_lookup.get("nm", {}):
        return economics_lookup["nm"][nm]
    if art and art in economics_lookup.get("article", {}):
        return economics_lookup["article"][art]
    if grp and grp in economics_lookup.get("group", {}):
        return economics_lookup["group"][grp]
    return {}


def estimate_gp_after_ads_from_economics(revenue: Any, orders: Any, ad_spend: Any, econ: Dict[str, Any]) -> float:
    revenue_f = money_or_zero(revenue)
    orders_f = money_or_zero(orders)
    spend_f = money_or_zero(ad_spend)
    if not econ or revenue_f <= 0:
        return float("nan")
    avg_price = float(econ.get("economics_avg_price", 0) or 0)
    units = orders_f
    if units <= 0 and avg_price > 0:
        units = revenue_f / avg_price
    commission = revenue_f * float(econ.get("economics_commission_pct", 0) or 0) / 100.0
    acquiring = revenue_f * float(econ.get("economics_acquiring_pct", 0) or 0) / 100.0
    vat = units * float(econ.get("economics_vat_per_unit", 0) or 0)
    logistics = units * float(econ.get("economics_logistics_per_unit", 0) or 0)
    cogs = units * float(econ.get("economics_cogs_per_unit", 0) or 0)
    return float(revenue_f - commission - acquiring - vat - logistics - cogs - spend_f)


def enrich_ads_with_estimated_gp(ads_df: pd.DataFrame, economics_df: pd.DataFrame) -> pd.DataFrame:
    """Добавляет условную ВП после рекламы в рекламные дневные строки.

    Формула: сумма заказов - комиссия WB% - эквайринг% - НДС/ед*заказы - логистика/ед*заказы
    - себестоимость/ед*заказы - расход рекламы. Если точного SKU нет, берём среднее по группе артикула
    (например, 901/5 -> среднее по всем 901).
    """
    if ads_df is None or ads_df.empty:
        return ads_df if ads_df is not None else pd.DataFrame()
    result = ads_df.copy()
    for col in ["economics_match_method", "economics_product_group", "economics_avg_price", "economics_commission_pct", "economics_acquiring_pct", "economics_vat_per_unit", "economics_logistics_per_unit", "economics_cogs_per_unit"]:
        if col not in result.columns:
            result[col] = "" if col in {"economics_match_method", "economics_product_group"} else float("nan")
    lookup = build_economics_lookup(economics_df)
    if not any(lookup.values()):
        print("Диагностика экономики: lookup пустой, ВП после рекламы не рассчитана", flush=True)
        return result
    matched = 0
    gp_values: List[float] = []
    for idx, row in result.iterrows():
        econ = lookup_economics_metrics(lookup, row.get("nm_id", ""), row.get("supplier_article", ""))
        gp = estimate_gp_after_ads_from_economics(row.get("revenue", 0), row.get("orders", 0), row.get("spend", 0), econ)
        gp_values.append(gp)
        if econ:
            matched += 1
            if not _clean_text_value(row.get("supplier_article", "")) and _clean_text_value(econ.get("supplier_article_from_economics", "")):
                result.at[idx, "supplier_article"] = econ.get("supplier_article_from_economics", "")
            for col in ["economics_match_method", "economics_product_group", "economics_avg_price", "economics_commission_pct", "economics_acquiring_pct", "economics_vat_per_unit", "economics_logistics_per_unit", "economics_cogs_per_unit"]:
                result.at[idx, col] = econ.get(col, "")
    result["estimated_gp_after_ads"] = gp_values
    # Заполняем gp_after_ads условной экономикой, если в рекламном отчёте ВП нет или она пустая.
    if "gp_after_ads" not in result.columns:
        result["gp_after_ads"] = result["estimated_gp_after_ads"]
    else:
        existing = pd.to_numeric(result["gp_after_ads"], errors="coerce")
        estimated = pd.to_numeric(result["estimated_gp_after_ads"], errors="coerce")
        result["gp_after_ads"] = existing.where(existing.notna(), estimated)
    print(
        f"Диагностика экономики: ВП после рекламы рассчитана для {matched} из {len(result)} строк рекламы; "
        f"себестоимость вычитаем={'да' if ECONOMICS_SUBTRACT_COGS else 'нет'}",
        flush=True,
    )
    return result


def enrich_supplier_articles_from_economics(df: pd.DataFrame, economics_df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty or economics_df is None or economics_df.empty or "nm_id" not in df.columns:
        return df if df is not None else pd.DataFrame()
    lookup = build_economics_lookup(economics_df)
    result = df.copy()
    if "supplier_article" not in result.columns:
        result["supplier_article"] = ""
    for idx, row in result.iterrows():
        if _clean_text_value(row.get("supplier_article", "")):
            continue
        econ = lookup_economics_metrics(lookup, row.get("nm_id", ""), "")
        art = _clean_text_value(econ.get("supplier_article_from_economics", "")) if econ else ""
        if art:
            result.at[idx, "supplier_article"] = art
    return result


def _economics_info_for_key(ads_df: pd.DataFrame, key: Tuple[str, str, str]) -> Dict[str, Any]:
    empty = {
        "economics_match_method": "", "economics_product_group": "", "economics_avg_price": "",
        "economics_commission_pct": "", "economics_acquiring_pct": "", "economics_vat_per_unit": "",
        "economics_logistics_per_unit": "", "economics_cogs_per_unit": "",
    }
    if ads_df is None or ads_df.empty:
        return empty
    campaign_id, nm_id, placement = key
    part = ads_df[
        (ads_df["campaign_id"].astype(str).map(_clean_id_value).eq(campaign_id))
        & (ads_df["nm_id"].astype(str).map(_clean_id_value).eq(nm_id))
        & (ads_df["placement"].astype(str).map(normalize_placement_value).eq(placement))
    ].copy()
    if part.empty:
        return empty
    for col in empty:
        if col in part.columns:
            vals = [_clean_text_value(x) for x in part[col].tolist() if _clean_text_value(x)]
            if vals:
                empty[col] = vals[-1]
    return empty

def aggregate_funnel_metrics(funnel_df: pd.DataFrame, start_date: date, end_date: date) -> pd.DataFrame:
    cols = ["nm_id", "card_views", "add_to_cart", "funnel_orders", "add_to_cart_conv", "cart_to_order_conv"]
    if funnel_df is None or funnel_df.empty:
        return pd.DataFrame(columns=cols)
    part = funnel_df.copy()
    if has_valid_dates(part):
        part = part[(part["date"] >= start_date) & (part["date"] <= end_date)].copy()
    if part.empty:
        return pd.DataFrame(columns=cols)
    agg = part.groupby("nm_id", dropna=False).agg(
        card_views=("card_views", "sum"),
        add_to_cart=("add_to_cart", "sum"),
        funnel_orders=("funnel_orders", "sum"),
    ).reset_index()
    agg["add_to_cart_conv"] = [safe_ctr_pct(a, v) for a, v in zip(agg["add_to_cart"], agg["card_views"])]
    agg["cart_to_order_conv"] = [safe_ctr_pct(o, a) for o, a in zip(agg["funnel_orders"], agg["add_to_cart"])]
    return agg


def fetch_current_goods_prices(config: Config, ctx: RunContext) -> Tuple[pd.DataFrame, pd.DataFrame]:
    url = WB_PRICES_BASE_URL + WB_PRICES_LIST_ENDPOINT
    rows: List[Dict[str, Any]] = []
    logs: List[Dict[str, Any]] = []
    limit = 1000
    offset = 0
    if ctx.mode == "preview" or ctx.dry_run:
        # Список текущих цен нужен даже в dry-run, это чтение без изменения данных.
        pass
    for page in range(1, 50):
        params = {"limit": limit, "offset": offset}
        endpoint = f"{WB_PRICES_LIST_ENDPOINT}?limit={limit}&offset={offset}"
        try:
            resp = requests.get(url, params=params, headers=wb_headers(config), timeout=90)
            logs.append(api_log_row(ctx.run_datetime, "GET", endpoint, "", str(resp.status_code), resp.text[:4000]))
            if resp.status_code != 200:
                break
            payload = resp.json()
            data = payload.get("data", payload) if isinstance(payload, dict) else payload
            if isinstance(data, dict):
                items = data.get("listGoods") or data.get("goods") or data.get("items") or data.get("data") or []
            elif isinstance(data, list):
                items = data
            else:
                items = []
            if not items:
                break
            for item in items:
                sizes = item.get("sizes") if isinstance(item.get("sizes"), list) else []
                first_size = sizes[0] if sizes else {}
                rows.append({
                    "nm_id": _clean_id_value(item.get("nmID") or item.get("nmId") or item.get("nm_id")),
                    "supplier_article_api": _clean_text_value(item.get("vendorCode") or item.get("supplierArticle") or item.get("article") or ""),
                    "current_wb_price": float(pd.to_numeric(pd.Series([item.get("price", first_size.get("price", 0))]), errors="coerce").fillna(0).iloc[0]),
                    "current_discount": float(pd.to_numeric(pd.Series([item.get("discount", first_size.get("discount", 0))]), errors="coerce").fillna(0).iloc[0]),
                })
            if len(items) < limit:
                break
            offset += limit
            time.sleep(0.4)
        except Exception as exc:
            logs.append(api_log_row(ctx.run_datetime, "GET", endpoint, "", "exception", repr(exc)))
            break
    df = pd.DataFrame(rows)
    if not df.empty:
        df = df[df["nm_id"].map(_clean_id_value).ne("")].drop_duplicates(subset=["nm_id"], keep="last")
    return df, pd.DataFrame(logs)


def _price_pending_events(price_history: pd.DataFrame) -> Dict[str, Dict[str, Any]]:
    if price_history is None or price_history.empty:
        return {}
    local = price_history.copy()
    for col in PRICE_HISTORY_COLUMNS:
        if col not in local.columns:
            local[col] = ""
    local["event_date_parsed"] = pd.to_datetime(local["event_date"], errors="coerce")
    local["run_dt_parsed"] = pd.to_datetime(local["run_datetime"], errors="coerce")
    local = local.sort_values(["event_date_parsed", "run_dt_parsed"], na_position="first")
    out: Dict[str, Dict[str, Any]] = {}
    for _, row in local.iterrows():
        status = _clean_text_value(row.get("postcheck_status", "")).lower()
        if status != "resolved":
            out[_clean_id_value(row.get("nm_id", ""))] = row.to_dict()
    return out


def evaluate_price_postchecks(price_history: pd.DataFrame, ads_df: pd.DataFrame, funnel_df: pd.DataFrame, ctx: RunContext) -> Tuple[pd.DataFrame, pd.DataFrame]:
    if price_history is None or price_history.empty:
        return pd.DataFrame(columns=PRICE_HISTORY_COLUMNS), pd.DataFrame()
    updated = coerce_history_columns_object(price_history.copy(), PRICE_HISTORY_COLUMNS)
    for col in PRICE_HISTORY_COLUMNS:
        if col not in updated.columns:
            updated[col] = pd.Series([""] * len(updated), index=updated.index, dtype="object")
        else:
            updated[col] = updated[col].astype("object")
    rows: List[Dict[str, Any]] = []
    funnel_by_day = funnel_df if funnel_df is not None else pd.DataFrame()
    for idx, row in updated.iterrows():
        event_dt = pd.to_datetime(row.get("event_date"), errors="coerce")
        if pd.isna(event_dt):
            continue
        event_day = event_dt.date()
        d2_end = event_day + timedelta(days=2)
        nm_id = _clean_id_value(row.get("nm_id", ""))
        if ctx.mature_end < d2_end:
            continue
        before_days = 5.0
        after_days = 2.0
        before = {
            "orders": float(pd.to_numeric(pd.Series([row.get("orders_before", 0)]), errors="coerce").fillna(0).iloc[0]) / before_days,
            "impressions": float(pd.to_numeric(pd.Series([row.get("impressions_before", 0)]), errors="coerce").fillna(0).iloc[0]) / before_days,
            "clicks": float(pd.to_numeric(pd.Series([row.get("clicks_before", 0)]), errors="coerce").fillna(0).iloc[0]) / before_days,
            "ctr": float(pd.to_numeric(pd.Series([row.get("ctr_before", 0)]), errors="coerce").fillna(0).iloc[0]),
            "card_views": float(pd.to_numeric(pd.Series([row.get("card_views_before", 0)]), errors="coerce").fillna(0).iloc[0]) / before_days,
            "add_to_cart_conv": float(pd.to_numeric(pd.Series([row.get("add_to_cart_conv_before", 0)]), errors="coerce").fillna(0).iloc[0]),
            "cart_to_order_conv": float(pd.to_numeric(pd.Series([row.get("cart_to_order_conv_before", 0)]), errors="coerce").fillna(0).iloc[0]),
        }
        ad_after = ads_df[(ads_df["nm_id"].astype(str) == str(nm_id))].copy() if ads_df is not None and not ads_df.empty else pd.DataFrame()
        if not ad_after.empty and has_valid_dates(ad_after):
            ad_after = ad_after[(ad_after["date"] >= event_day + timedelta(days=1)) & (ad_after["date"] <= d2_end)].copy()
        after_orders = float(pd.to_numeric(ad_after.get("orders", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()) / after_days if not ad_after.empty else 0.0
        after_impressions = float(pd.to_numeric(ad_after.get("impressions", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()) / after_days if not ad_after.empty else 0.0
        after_clicks = float(pd.to_numeric(ad_after.get("clicks", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()) / after_days if not ad_after.empty else 0.0
        after_ctr = safe_ctr_pct(after_clicks, after_impressions)
        f_after = funnel_by_day[funnel_by_day["nm_id"].astype(str) == str(nm_id)].copy() if not funnel_by_day.empty else pd.DataFrame()
        if not f_after.empty and has_valid_dates(f_after):
            f_after = f_after[(f_after["date"] >= event_day + timedelta(days=1)) & (f_after["date"] <= d2_end)].copy()
        card_views_after = float(pd.to_numeric(f_after.get("card_views", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()) / after_days if not f_after.empty else 0.0
        add_to_cart_after = float(pd.to_numeric(f_after.get("add_to_cart", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()) / after_days if not f_after.empty else 0.0
        funnel_orders_after = float(pd.to_numeric(f_after.get("funnel_orders", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()) / after_days if not f_after.empty else 0.0
        add_to_cart_conv_after = safe_ctr_pct(add_to_cart_after, card_views_after)
        cart_to_order_conv_after = safe_ctr_pct(funnel_orders_after, add_to_cart_after)
        funnel_missing_event = str(row.get("funnel_missing", "")).strip().lower() in {"true", "1", "yes", "да"}
        has_funnel_after = not f_after.empty and (card_views_after > 0 or add_to_cart_after > 0 or funnel_orders_after > 0)
        use_funnel_verdict = (not funnel_missing_event) and has_funnel_after and before["card_views"] > 0
        traffic_ratio = (card_views_after / before["card_views"]) if before["card_views"] > 0 else ((after_clicks / before["clicks"]) if before["clicks"] > 0 else 1.0)
        orders_ratio = (after_orders / before["orders"]) if before["orders"] > 0 else (1.0 if after_orders > 0 else 0.0)
        clicks_ratio = (after_clicks / before["clicks"]) if before["clicks"] > 0 else (1.0 if after_clicks > 0 else 0.0)
        ctr_ratio = (after_ctr / before["ctr"]) if before["ctr"] > 0 else 1.0
        atc_conv_ratio = (add_to_cart_conv_after / before["add_to_cart_conv"]) if before["add_to_cart_conv"] > 0 else 1.0
        cto_conv_ratio = (cart_to_order_conv_after / before["cart_to_order_conv"]) if before["cart_to_order_conv"] > 0 else 1.0
        if use_funnel_verdict:
            if orders_ratio >= 0.95 and atc_conv_ratio >= 0.90 and cto_conv_ratio >= 0.90:
                verdict = "PRICE_RAISE_GOOD"
                comment = "цена повышена, заказы и конверсии удержались"
            elif orders_ratio < 0.90 and traffic_ratio < 0.85 and atc_conv_ratio >= 0.90 and cto_conv_ratio >= 0.90:
                verdict = "PRICE_RAISE_TRAFFIC_DROP"
                comment = "заказы просели на фоне падения трафика; конверсии не доказывают вред цены"
            elif orders_ratio < 0.90 and (atc_conv_ratio < 0.90 or cto_conv_ratio < 0.90) and traffic_ratio >= 0.85:
                verdict = "PRICE_RAISE_BAD_CONVERSION_DROP"
                comment = "трафик удержался, но конверсия упала; нужен откат скидки"
            else:
                verdict = "PRICE_RAISE_MIXED"
                comment = "смешанный эффект; без нового повышения до следующей проверки"
        else:
            if before["orders"] <= 0 and before["clicks"] < 20:
                verdict = "PRICE_CHECK_NOT_ENOUGH_DATA"
                comment = "воронка отсутствует и мало рекламных данных; цену не откатываем автоматически"
            elif orders_ratio >= 0.95 and ctr_ratio >= 0.90:
                verdict = "PRICE_RAISE_GOOD_ADS_ONLY"
                comment = "воронка отсутствует; по рекламе заказы/CTR удержались"
            elif orders_ratio < 0.90 and clicks_ratio < 0.85 and ctr_ratio >= 0.90:
                verdict = "PRICE_EFFECT_UNCLEAR_TRAFFIC_DROP"
                comment = "воронка отсутствует; падение заказов совпало с падением кликов/трафика, откат не автоматический"
            elif orders_ratio < 0.80 and clicks_ratio >= 0.85:
                verdict = "PRICE_RAISE_BAD_ADS_ORDERS_DROP"
                comment = "воронка отсутствует; трафик удержался, но заказы просели, нужен откат скидки"
            else:
                verdict = "PRICE_RAISE_MIXED_ADS_ONLY"
                comment = "воронка отсутствует; смешанный эффект по рекламе, без нового повышения"
        updated.at[idx, "postcheck_status"] = "resolved"
        updated.at[idx, "final_verdict"] = verdict
        updated.at[idx, "d2_verdict"] = verdict
        updated.at[idx, "d2_check_date"] = str(ctx.mature_end)
        rows.append({
            "price_event_id": row.get("price_event_id", ""), "nm_id": nm_id, "supplier_article": row.get("supplier_article", ""),
            "subject_norm": row.get("subject_norm", ""), "event_date": row.get("event_date", ""),
            "old_discount": row.get("old_discount", ""), "new_discount": row.get("new_discount", ""),
            "orders_before_daily": before["orders"], "orders_after_daily": after_orders, "orders_ratio": orders_ratio,
            "traffic_ratio": traffic_ratio, "ctr_before": before["ctr"], "ctr_after": after_ctr,
            "card_views_before_daily": before["card_views"], "card_views_after_daily": card_views_after,
            "add_to_cart_conv_before": before["add_to_cart_conv"], "add_to_cart_conv_after": add_to_cart_conv_after, "add_to_cart_conv_ratio": atc_conv_ratio,
            "cart_to_order_conv_before": before["cart_to_order_conv"], "cart_to_order_conv_after": cart_to_order_conv_after, "cart_to_order_conv_ratio": cto_conv_ratio,
            "funnel_missing": (not use_funnel_verdict),
            "verdict": verdict, "comment": comment,
        })
    return updated[PRICE_HISTORY_COLUMNS], pd.DataFrame(rows)


def build_price_decisions(metrics_df: pd.DataFrame, funnel_current: pd.DataFrame, goods_prices: pd.DataFrame, price_history: pd.DataFrame, config: Config, ctx: RunContext) -> pd.DataFrame:
    """Формирует решения по тестовому повышению цены через скидку продавца.

    Правила:
    - работаем только с Помадами, Блесками и Косметическими карандашами;
    - текущую скидку продавца берём только из WB Discounts & Prices API;
    - если скидка из API не получена, цену не меняем;
    - повышение цены = снижение фактической скидки продавца на 1 п.п.;
    - ниже DEFAULT_MIN_SELLER_DISCOUNT_PCT не опускаемся;
    - если воронки нет, ценовой тест разрешён ограниченно: оцениваем по рекламе/заказам, а в отчётах ставим funnel_missing=True;
    - не больше MAX_PRICE_TEST_ITEMS_PER_RUN новых price-test за один запуск; по умолчанию 30, можно переопределить env WB_MAX_PRICE_TEST_ITEMS_PER_RUN;
    - если есть незавершённый price post-check, товар не трогаем.
    """
    if metrics_df is None or metrics_df.empty:
        return pd.DataFrame(columns=PRICE_DECISION_COLUMNS)

    base = metrics_df[["nm_id", "supplier_article", "subject_norm", "orders", "impressions", "clicks", "ctr_pct"]].copy()
    base["subject_norm"] = base["subject_norm"].map(lambda x: normalize_subject_value(x).lower())
    base = base[base["subject_norm"].isin(PRICE_TEST_SUBJECTS)].copy()
    if base.empty:
        return pd.DataFrame(columns=PRICE_DECISION_COLUMNS)

    base = base.groupby(["nm_id", "supplier_article", "subject_norm"], dropna=False).agg(
        orders=("orders", "sum"),
        impressions=("impressions", "sum"),
        clicks=("clicks", "sum"),
        ctr_pct=("ctr_pct", "mean"),
    ).reset_index()

    has_funnel = funnel_current is not None and not funnel_current.empty
    if has_funnel:
        base = base.merge(funnel_current, on="nm_id", how="left")

    for col in ["card_views", "add_to_cart", "funnel_orders", "add_to_cart_conv", "cart_to_order_conv"]:
        if col not in base.columns:
            base[col] = 0.0
        base[col] = pd.to_numeric(base[col], errors="coerce").fillna(0.0)

    if goods_prices is not None and not goods_prices.empty:
        gp = goods_prices[["nm_id", "current_discount", "current_wb_price"]].copy()
        gp["current_discount"] = pd.to_numeric(gp["current_discount"], errors="coerce")
        gp["current_wb_price"] = pd.to_numeric(gp["current_wb_price"], errors="coerce")
        base = base.merge(gp, on="nm_id", how="left")
    else:
        base["current_discount"] = float("nan")
        base["current_wb_price"] = float("nan")

    pending = _price_pending_events(price_history)
    rows: List[Dict[str, Any]] = []
    candidate_indices: List[int] = []

    # Сначала строим все строки и помечаем потенциальные новые тесты.
    for idx, row in base.iterrows():
        nm_id = _clean_id_value(row.get("nm_id", ""))
        subject_norm = normalize_subject_value(row.get("subject_norm", "")).lower()
        current_discount_raw = pd.to_numeric(pd.Series([row.get("current_discount")]), errors="coerce").iloc[0]
        current_discount = float(current_discount_raw) if not pd.isna(current_discount_raw) else float("nan")
        action = "Без изменений"
        reason_code = "PRICE_MONITOR_ONLY"
        new_discount: Optional[float] = None
        prev_event_id = ""
        post_status = ""
        pending_event = pending.get(nm_id)

        if subject_norm not in PRICE_TEST_SUBJECTS:
            reason_code = "PRICE_NOT_TARGET_SUBJECT"
        elif pd.isna(current_discount) or current_discount <= 0:
            reason_code = "NO_CURRENT_DISCOUNT_FROM_WB_API"
        elif pending_event:
            prev_event_id = _clean_text_value(pending_event.get("price_event_id", ""))
            post_status = _clean_text_value(pending_event.get("postcheck_status", "pending")) or "pending"
            reason_code = "PRICE_WAIT_D2_POSTCHECK"
        else:
            latest_bad = None
            if price_history is not None and not price_history.empty:
                ph = price_history[price_history["nm_id"].astype(str) == str(nm_id)].copy()
                if not ph.empty:
                    ph["run_dt"] = pd.to_datetime(ph["run_datetime"], errors="coerce")
                    ph = ph.sort_values("run_dt")
                    last = ph.tail(1).iloc[0]
                    if _clean_text_value(last.get("final_verdict", "")) == "PRICE_RAISE_BAD_CONVERSION_DROP":
                        latest_bad = last
            if latest_bad is not None:
                old_discount = float(pd.to_numeric(pd.Series([latest_bad.get("old_discount", current_discount)]), errors="coerce").fillna(current_discount).iloc[0])
                if current_discount < old_discount:
                    action = "Вернуть скидку"
                    new_discount = old_discount
                    reason_code = "PRICE_RAISE_BAD_REVERT"
                else:
                    reason_code = "PRICE_ALREADY_REVERTED_AFTER_BAD"
            elif current_discount <= DEFAULT_MIN_SELLER_DISCOUNT_PCT:
                reason_code = "PRICE_MIN_DISCOUNT_REACHED"
            elif float(row.get("orders", 0) or 0) <= 0:
                reason_code = "PRICE_NO_ORDERS_FOR_TEST"
            else:
                action = "Повысить цену"
                new_discount = max(DEFAULT_MIN_SELLER_DISCOUNT_PCT, current_discount - DEFAULT_PRICE_RAISE_STEP_PP)
                reason_code = "PRICE_RAISE_1PP_TEST" if has_funnel else "PRICE_RAISE_1PP_TEST_NO_FUNNEL"

        reason_text = (
            f"скидка_WB_API={current_discount if not pd.isna(current_discount) else 'н/д'}%; "
            f"скидка по умолчанию={DEFAULT_SELLER_DISCOUNT_PCT}%; новая скидка={new_discount if new_discount is not None else 'н/д'}; "
            f"предмет={subject_norm}; заказы={float(row.get('orders',0) or 0):.0f}; показы={float(row.get('impressions',0) or 0):.0f}; "
            f"клики={float(row.get('clicks',0) or 0):.0f}; CTR={float(row.get('ctr_pct',0) or 0):.2f}%; "
            f"просмотры карточки={float(row.get('card_views',0) or 0):.0f}; add_to_cart_conv={float(row.get('add_to_cart_conv',0) or 0):.2f}%; "
            f"cart_to_order_conv={float(row.get('cart_to_order_conv',0) or 0):.2f}%; "
            f"funnel_missing={str(not has_funnel).lower()}"
        )
        rows.append({
            "nm_id": nm_id,
            "supplier_article": row.get("supplier_article", ""),
            "subject_norm": subject_norm,
            "current_discount": current_discount,
            "new_discount": new_discount,
            "price_action": action,
            "reason_code": reason_code,
            "reason_text": reason_text,
            "orders": row.get("orders", 0),
            "impressions": row.get("impressions", 0),
            "clicks": row.get("clicks", 0),
            "ctr_pct": row.get("ctr_pct", 0),
            "card_views": row.get("card_views", 0),
            "add_to_cart": row.get("add_to_cart", 0),
            "funnel_orders": row.get("funnel_orders", 0),
            "add_to_cart_conv": row.get("add_to_cart_conv", 0),
            "cart_to_order_conv": row.get("cart_to_order_conv", 0),
            "funnel_missing": bool(not has_funnel),
            "previous_price_event_id": prev_event_id,
            "price_postcheck_status": post_status,
        })
        if action == "Повысить цену" and reason_code in {"PRICE_RAISE_1PP_TEST", "PRICE_RAISE_1PP_TEST_NO_FUNNEL"}:
            candidate_indices.append(len(rows) - 1)

    # Ограничиваем новые ценовые тесты за запуск, чтобы не менять 85 товаров одним пакетом.
    if len(candidate_indices) > MAX_PRICE_TEST_ITEMS_PER_RUN:
        # Оставляем товары с наибольшим числом заказов, остальным ставим ожидание лимита.
        ranked = sorted(candidate_indices, key=lambda i: float(rows[i].get("orders", 0) or 0), reverse=True)
        allowed = set(ranked[:MAX_PRICE_TEST_ITEMS_PER_RUN])
        for i in candidate_indices:
            if i not in allowed:
                rows[i]["price_action"] = "Без изменений"
                rows[i]["new_discount"] = None
                rows[i]["reason_code"] = "PRICE_TEST_LIMIT_PER_RUN"
                rows[i]["reason_text"] += f"; лимит новых price-test за запуск={MAX_PRICE_TEST_ITEMS_PER_RUN}"

    return pd.DataFrame(rows, columns=PRICE_DECISION_COLUMNS)


def apply_price_changes(price_decisions: pd.DataFrame, goods_prices: pd.DataFrame, config: Config, ctx: RunContext, apply_price: bool) -> Tuple[pd.DataFrame, pd.DataFrame]:
    if price_decisions is None or price_decisions.empty:
        return pd.DataFrame(), pd.DataFrame()
    to_send = price_decisions[price_decisions["price_action"].isin(["Повысить цену", "Вернуть скидку"])].copy()
    if to_send.empty:
        return pd.DataFrame(), pd.DataFrame()
    payload_rows: List[Dict[str, Any]] = []
    prices_map = {}
    if goods_prices is not None and not goods_prices.empty:
        prices_map = dict(zip(goods_prices["nm_id"].astype(str), goods_prices.get("current_wb_price", pd.Series(dtype=float))))
    sent_rows: List[Dict[str, Any]] = []
    for _, row in to_send.iterrows():
        nm_id_int = to_int_id(row.get("nm_id"))
        new_discount = int(round(float(row.get("new_discount")))) if pd.notna(row.get("new_discount")) else None
        price = prices_map.get(str(row.get("nm_id")), 0)
        price_int = int(round(float(price or 0)))
        if nm_id_int is None or new_discount is None or price_int <= 0:
            continue
        payload_rows.append({"nmID": nm_id_int, "price": price_int, "discount": new_discount})
        sent_rows.append(row.to_dict())
    if not payload_rows:
        return pd.DataFrame(), pd.DataFrame()
    api_logs: List[Dict[str, Any]] = []
    if ctx.mode == "preview" or ctx.dry_run or not apply_price:
        api_logs.append(api_log_row(ctx.run_datetime, "POST", WB_PRICE_UPLOAD_ENDPOINT, {"rows": len(payload_rows)}, "not_sent", "Цены не отправлялись: preview/dry-run или apply_price=False"))
        return pd.DataFrame(), pd.DataFrame(api_logs)
    url = WB_PRICES_BASE_URL + WB_PRICE_UPLOAD_ENDPOINT
    for start in range(0, len(payload_rows), 1000):
        batch = payload_rows[start:start+1000]
        payload = {"data": batch}
        try:
            resp = requests.post(url, headers=wb_headers(config), json=payload, timeout=120)
            api_logs.append(api_log_row(ctx.run_datetime, "POST", WB_PRICE_UPLOAD_ENDPOINT, json.dumps(payload, ensure_ascii=False), str(resp.status_code), resp.text[:4000]))
            status = str(resp.status_code)
            response = resp.text[:1000]
        except Exception as exc:
            api_logs.append(api_log_row(ctx.run_datetime, "POST", WB_PRICE_UPLOAD_ENDPOINT, json.dumps(payload, ensure_ascii=False), "exception", repr(exc)))
            status = "exception"
            response = repr(exc)
        for r in sent_rows[start:start+len(batch)]:
            r["api_status"] = status
            r["api_response"] = response
    return pd.DataFrame(sent_rows), pd.DataFrame(api_logs)


def record_price_events(applied_price_changes: pd.DataFrame, price_history: pd.DataFrame, ctx: RunContext) -> pd.DataFrame:
    if price_history is None or price_history.empty:
        history = pd.DataFrame(columns=PRICE_HISTORY_COLUMNS)
    else:
        history = price_history.copy()
    if applied_price_changes is None or applied_price_changes.empty:
        return history[PRICE_HISTORY_COLUMNS]
    rows: List[Dict[str, Any]] = []
    for _, row in applied_price_changes.iterrows():
        status = _clean_text_value(row.get("api_status", ""))
        is_success = status.isdigit() and 200 <= int(status) < 300
        if status == "dry_run_or_not_applied":
            continue
        if not is_success:
            continue
        direction = "price_raise" if row.get("price_action") == "Повысить цену" else "price_rollback"
        rows.append({
            "price_event_id": str(uuid.uuid4()), "run_datetime": ctx.run_datetime.strftime("%Y-%m-%d %H:%M:%S"), "event_date": ctx.run_datetime.date().isoformat(),
            "nm_id": row.get("nm_id", ""), "supplier_article": row.get("supplier_article", ""), "subject_norm": row.get("subject_norm", ""),
            "old_discount": row.get("current_discount", ""), "new_discount": row.get("new_discount", ""), "direction": direction, "reason_code": row.get("reason_code", ""),
            "orders_before": row.get("orders", 0), "impressions_before": row.get("impressions", 0), "clicks_before": row.get("clicks", 0), "ctr_before": row.get("ctr_pct", 0),
            "card_views_before": row.get("card_views", 0), "add_to_cart_before": row.get("add_to_cart", 0), "funnel_orders_before": row.get("funnel_orders", 0),
            "add_to_cart_conv_before": row.get("add_to_cart_conv", 0), "cart_to_order_conv_before": row.get("cart_to_order_conv", 0),
            "funnel_missing": row.get("funnel_missing", False),
            "postcheck_status": "pending", "final_verdict": "", "d2_verdict": "", "d2_check_date": "", "api_status": row.get("api_status", ""), "api_response": row.get("api_response", ""),
        })
    if rows:
        history = pd.concat([history, pd.DataFrame(rows)], ignore_index=True, sort=False)
    for col in PRICE_HISTORY_COLUMNS:
        if col not in history.columns:
            history[col] = ""
    return history[PRICE_HISTORY_COLUMNS]


# =============================
# Окна, агрегация и метрики
# =============================

def build_windows(run_date: Optional[date] = None) -> Tuple[date, date, date, date, date]:
    today = run_date or date.today()
    mature_end = today - timedelta(days=3)
    current_start = mature_end - timedelta(days=4)
    current_end = mature_end
    base_end = current_start - timedelta(days=1)
    base_start = base_end - timedelta(days=4)
    return mature_end, current_start, current_end, base_start, base_end


def build_run_context(args: argparse.Namespace) -> RunContext:
    mature_end, current_start, current_end, base_start, base_end = build_windows()
    return RunContext(
        mode=args.command,
        dry_run=bool(args.dry_run),
        apply_pause=bool(args.apply_pause),
        apply_start=bool(args.apply_start),
        apply_price=bool(getattr(args, "apply_price", False)),
        apply_experiment=bool(getattr(args, "apply_experiment", False)),
        run_datetime=datetime.now(),
        mature_end=mature_end,
        current_start=current_start,
        current_end=current_end,
        base_start=base_start,
        base_end=base_end,
    )


def has_valid_dates(df: pd.DataFrame) -> bool:
    return "date" in df.columns and df["date"].notna().any()


def filter_by_date_window(df: pd.DataFrame, start_date: date, end_date: date) -> pd.DataFrame:
    if not has_valid_dates(df):
        return df.copy()
    mask = (df["date"] >= start_date) & (df["date"] <= end_date)
    return df.loc[mask].copy()


def latest_nonempty_value(df: pd.DataFrame, group_keys: List[str], col: str) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=group_keys + [col])
    local = df[group_keys + ["date", "_row_id", col]].copy()
    local["_has_value"] = local[col].map(lambda x: _clean_text_value(x) != "" if not isinstance(x, (int, float)) else not pd.isna(x))
    local = local.sort_values(group_keys + ["_has_value", "date", "_row_id"])
    latest = local.groupby(group_keys, dropna=False).tail(1)
    return latest[group_keys + [col]]


def latest_numeric_value(df: pd.DataFrame, group_keys: List[str], col: str) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=group_keys + [col])
    local = df[group_keys + ["date", "_row_id", col]].copy()
    local = local.sort_values(group_keys + ["date", "_row_id"])
    latest = local.groupby(group_keys, dropna=False).tail(1)
    return latest[group_keys + [col]]


def aggregate_window_metrics(df: pd.DataFrame, group_keys: List[str], prefix: str = "") -> pd.DataFrame:
    if df.empty:
        cols = group_keys + [
            f"{prefix}spend",
            f"{prefix}revenue",
            f"{prefix}orders",
            f"{prefix}impressions",
            f"{prefix}clicks",
            f"{prefix}gp_after_ads",
        ]
        return pd.DataFrame(columns=cols)

    agg = df.groupby(group_keys, dropna=False).agg(
        spend=("spend", "sum"),
        revenue=("revenue", "sum"),
        orders=("orders", "sum"),
        impressions=("impressions", "sum"),
        clicks=("clicks", "sum"),
        gp_after_ads=("gp_after_ads", "sum"),
    ).reset_index()

    # Если ВП отсутствует во всех строках группы, sum даёт 0. Возвращаем NaN для таких групп.
    gp_present = df.groupby(group_keys, dropna=False)["gp_after_ads"].apply(lambda s: s.notna().any()).reset_index(name="_gp_present")
    agg = agg.merge(gp_present, on=group_keys, how="left")
    agg.loc[~agg["_gp_present"].fillna(False), "gp_after_ads"] = float("nan")
    agg = agg.drop(columns=["_gp_present"])

    if prefix:
        rename_map = {col: f"{prefix}{col}" for col in ["spend", "revenue", "orders", "impressions", "clicks", "gp_after_ads"]}
        agg = agg.rename(columns=rename_map)
    return agg


def safe_drr_pct(spend: float, revenue: float) -> float:
    spend = float(spend or 0)
    revenue = float(revenue or 0)
    if revenue == 0 and spend > 0:
        return 999.0
    if revenue == 0 and spend == 0:
        return 0.0
    return spend / revenue * 100.0


def safe_cpo(spend: float, orders: float) -> float:
    spend = float(spend or 0)
    orders = float(orders or 0)
    if orders == 0 and spend > 0:
        return 999999.0
    if orders == 0:
        return 0.0
    return spend / orders


def safe_ctr_pct(clicks: float, impressions: float) -> float:
    clicks = float(clicks or 0)
    impressions = float(impressions or 0)
    if impressions == 0:
        return 0.0
    return clicks / impressions * 100.0


def growth_pct_or_status(current: float, base: float) -> Tuple[Optional[float], str]:
    """Безопасный расчёт изменения к базе.

    Для обычных неотрицательных метрик возвращает % роста.
    Для ВП/GP база может быть 0 или отрицательной: в этих случаях процент
    классической формулой current / base считать нельзя — иначе получаем
    division by zero или вводящий в заблуждение знак. Поэтому возвращаем
    None и статус, а в отчёте показываем причину.
    """
    current_num = pd.to_numeric(pd.Series([current]), errors="coerce").iloc[0]
    base_num = pd.to_numeric(pd.Series([base]), errors="coerce").iloc[0]
    current = 0.0 if pd.isna(current_num) else float(current_num)
    base = 0.0 if pd.isna(base_num) else float(base_num)

    if abs(base) < 1e-9:
        if abs(current) < 1e-9:
            return None, "ZERO_BASE"
        if current > 0:
            return None, "NEW_ACTIVITY"
        return None, "NEGATIVE_CURRENT_ZERO_BASE"

    # Для отрицательной базы процент роста ВП не интерпретируем как обычный growth %.
    if base < 0:
        if current >= 0:
            return None, "FROM_NEGATIVE_TO_NONNEGATIVE"
        return ((current - base) / abs(base)) * 100.0, "NEGATIVE_BASE"

    return (current / base - 1.0) * 100.0, "OK"


def compute_metrics(df: pd.DataFrame) -> pd.DataFrame:
    result = df.copy()
    result["campaign_drr_pct"] = [safe_drr_pct(s, r) for s, r in zip(result["spend"], result["revenue"])]
    result["cpo"] = [safe_cpo(s, o) for s, o in zip(result["spend"], result["orders"])]
    result["ctr_pct"] = [safe_ctr_pct(c, i) for c, i in zip(result["clicks"], result["impressions"])]

    for metric in ["spend", "revenue", "orders", "impressions", "clicks", "gp_after_ads"]:
        base_col = f"base_{metric}"
        if base_col in result.columns:
            growth_values: List[Optional[float]] = []
            growth_statuses: List[str] = []
            for current_value, base_value in zip(result[metric], result[base_col]):
                if pd.isna(current_value):
                    current_value = 0.0
                if pd.isna(base_value):
                    base_value = 0.0
                growth, status = growth_pct_or_status(float(current_value), float(base_value))
                growth_values.append(growth)
                growth_statuses.append(status)
            result[f"{metric}_growth_pct"] = growth_values
            result[f"{metric}_growth_status"] = growth_statuses
    return result


def aggregate_campaign_metrics(ads_df: pd.DataFrame, ctx: RunContext) -> pd.DataFrame:
    ads_df = filter_managed_subject_rows(ads_df)
    group_keys = ["campaign_id", "nm_id", "placement"]
    current_df = filter_by_date_window(ads_df, ctx.current_start, ctx.current_end)
    base_df = filter_by_date_window(ads_df, ctx.base_start, ctx.base_end)

    if current_df.empty and not ads_df.empty and not has_valid_dates(ads_df):
        current_df = ads_df.copy()

    current_metrics = aggregate_window_metrics(current_df, group_keys, prefix="")
    base_metrics = aggregate_window_metrics(base_df, group_keys, prefix="base_")

    pause21_start = ctx.mature_end - timedelta(days=PAUSE_ANALYSIS_DAYS - 1)
    pause21_df = filter_by_date_window(ads_df, pause21_start, ctx.mature_end) if has_valid_dates(ads_df) else ads_df.copy()
    pause21_metrics = aggregate_window_metrics(pause21_df, group_keys, prefix="last21_")

    if current_metrics.empty:
        return pd.DataFrame(columns=DECISION_COLUMNS)

    result = current_metrics.merge(base_metrics, on=group_keys, how="left")
    result = result.merge(pause21_metrics, on=group_keys, how="left")

    source_for_dims = current_df if not current_df.empty else ads_df
    for col in ["campaign_name", "campaign_status", "supplier_article", "subject_norm"]:
        result = result.merge(latest_nonempty_value(source_for_dims, group_keys, col), on=group_keys, how="left")
    result = result.merge(latest_numeric_value(source_for_dims, group_keys, "current_bid_rub"), on=group_keys, how="left")

    for metric in ["base_spend", "base_revenue", "base_orders", "base_impressions", "base_clicks",
                   "last21_spend", "last21_revenue", "last21_orders", "last21_impressions", "last21_clicks"]:
        if metric not in result.columns:
            result[metric] = 0.0
        result[metric] = pd.to_numeric(result[metric], errors="coerce").fillna(0.0)

    if "base_gp_after_ads" not in result.columns:
        result["base_gp_after_ads"] = float("nan")

    result = compute_metrics(result)
    result["drr_limit_pct"] = result["subject_norm"].map(drr_limit_for_subject)
    days = safe_window_days(ctx)
    result["avg_impressions_per_day"] = pd.to_numeric(result.get("impressions", 0), errors="coerce").fillna(0.0) / days
    result["avg_spend_per_day"] = pd.to_numeric(result.get("spend", 0), errors="coerce").fillna(0.0) / days
    result["last21_drr_pct"] = [safe_drr_pct(s, r) for s, r in zip(result.get("last21_spend", pd.Series([0]*len(result))), result.get("last21_revenue", pd.Series([0]*len(result))))]
    result["last21_avg_impressions_per_day"] = pd.to_numeric(result.get("last21_impressions", 0), errors="coerce").fillna(0.0) / float(PAUSE_ANALYSIS_DAYS)
    result["last21_avg_spend_per_day"] = pd.to_numeric(result.get("last21_spend", 0), errors="coerce").fillna(0.0) / float(PAUSE_ANALYSIS_DAYS)
    return result


# =============================
# Post-check изменений ставки
# =============================

def make_key(row: pd.Series | Dict[str, Any]) -> Tuple[str, str, str]:
    return (
        _clean_id_value(row.get("campaign_id", "")),
        _clean_id_value(row.get("nm_id", "")),
        normalize_placement_value(row.get("placement", "")),
    )


def is_ramp_event_reason(reason_code: Any) -> bool:
    text = _clean_text_value(reason_code).upper()
    return text.startswith("RAMP_") or text.startswith("LOW_BID_NO_SPEND_NO_ORDERS_RAMP")


def wait_days_for_event(row: pd.Series | Dict[str, Any]) -> int:
    reason_code = _clean_text_value(row.get("reason_code", ""))
    return int(RAMP_CHECK_DAYS) if is_ramp_event_reason(reason_code) else 3


def wait_rule_for_event(row: pd.Series | Dict[str, Any]) -> str:
    reason_code = _clean_text_value(row.get("reason_code", ""))
    direction = _clean_text_value(row.get("direction", "")).lower()
    if is_ramp_event_reason(reason_code):
        return f"WAIT_D{int(RAMP_CHECK_DAYS)}_RAMP_CHECK"
    if direction == "raise":
        return "WAIT_D3_RAISE_CHECK"
    if direction == "lower":
        return "WAIT_D3_LOWER_CHECK"
    return "WAIT_D3_BID_CHECK"


def pending_wait_info(row: pd.Series | Dict[str, Any], ctx: Optional[RunContext] = None) -> Dict[str, Any]:
    event_dt = pd.to_datetime(row.get("event_date", ""), errors="coerce")
    wait_days = wait_days_for_event(row)
    wait_rule = wait_rule_for_event(row)
    if pd.isna(event_dt):
        return {
            "active_wait": False,
            "wait_rule": wait_rule,
            "wait_until_date": "",
            "wait_days_left": 0,
            "wait_status": "NO_EVENT_DATE",
        }
    event_day = event_dt.date()
    wait_until = event_day + timedelta(days=wait_days)
    mature_end = ctx.mature_end if ctx is not None else date.today()
    active_wait = mature_end < wait_until
    days_left = max((wait_until - mature_end).days, 0)
    return {
        "active_wait": bool(active_wait),
        "wait_rule": wait_rule,
        "wait_until_date": wait_until.isoformat(),
        "wait_days_left": int(days_left),
        "wait_status": "WAIT_ACTIVE" if active_wait else "WAIT_EXPIRED",
    }


def load_pending_events(bid_history: pd.DataFrame, ctx: Optional[RunContext] = None) -> Dict[Tuple[str, str, str], Dict[str, Any]]:
    """Возвращает только реально активные ожидания.

    Старые записи с postcheck_status != resolved больше не должны бесконечно блокировать ставку.
    Обычное изменение ставки ждём D+3, разгон показов ждём D+7.
    Если срок уже наступил, строка не попадает в pending и снова идёт в обычную логику ДРР/разгона.
    """
    if bid_history.empty:
        return {}
    local = bid_history.copy()
    for col in BID_HISTORY_COLUMNS:
        if col not in local.columns:
            local[col] = ""
    local["event_date_parsed"] = pd.to_datetime(local["event_date"], errors="coerce").dt.date
    local["run_dt_parsed"] = pd.to_datetime(local["run_datetime"], errors="coerce")
    local = local.sort_values(["event_date_parsed", "run_dt_parsed"], na_position="first")
    pending: Dict[Tuple[str, str, str], Dict[str, Any]] = {}
    for _, row in local.iterrows():
        status = _clean_text_value(row.get("postcheck_status", "")).lower()
        verdict = _clean_text_value(row.get("final_verdict", ""))
        if status == "resolved" or verdict in {"RAISE_NO_TRAFFIC_GROWTH"}:
            continue
        old_bid = pd.to_numeric(pd.Series([row.get("old_bid_rub", None)]), errors="coerce").iloc[0]
        new_bid = pd.to_numeric(pd.Series([row.get("new_bid_rub", None)]), errors="coerce").iloc[0]
        if pd.isna(old_bid) or pd.isna(new_bid) or float(old_bid) == float(new_bid):
            continue
        expected_step, _ = bid_step_rub(row.get("placement", ""))
        if abs(abs(float(new_bid) - float(old_bid)) - float(expected_step)) > 0.05:
            continue
        info = pending_wait_info(row, ctx)
        if not info.get("active_wait", False):
            continue
        event = row.to_dict()
        event.update(info)
        pending[make_key(row)] = event
    return pending

def latest_postcheck_results(bid_history: pd.DataFrame) -> Dict[Tuple[str, str, str], Dict[str, Any]]:
    if bid_history.empty:
        return {}
    local = bid_history.copy()
    local["event_date_parsed"] = pd.to_datetime(local["event_date"], errors="coerce").dt.date
    local["run_dt_parsed"] = pd.to_datetime(local["run_datetime"], errors="coerce")
    local = local.sort_values(["event_date_parsed", "run_dt_parsed"], na_position="first")
    latest: Dict[Tuple[str, str, str], Dict[str, Any]] = {}
    for _, row in local.iterrows():
        latest[make_key(row)] = row.to_dict()
    return latest


def aggregate_after_event(
    ads_df: pd.DataFrame,
    key: Tuple[str, str, str],
    start_date: date,
    end_date: date,
) -> Dict[str, float]:
    if ads_df.empty:
        return {"spend": 0.0, "revenue": 0.0, "orders": 0.0, "impressions": 0.0, "clicks": 0.0, "gp_after_ads": float("nan")}
    if not has_valid_dates(ads_df):
        return {"spend": 0.0, "revenue": 0.0, "orders": 0.0, "impressions": 0.0, "clicks": 0.0, "gp_after_ads": float("nan")}
    campaign_id, nm_id, placement = key
    mask = (
        (ads_df["campaign_id"] == campaign_id)
        & (ads_df["nm_id"] == nm_id)
        & (ads_df["placement"] == placement)
        & (ads_df["date"] >= start_date)
        & (ads_df["date"] <= end_date)
    )
    part = ads_df.loc[mask]
    if part.empty:
        return {"spend": 0.0, "revenue": 0.0, "orders": 0.0, "impressions": 0.0, "clicks": 0.0, "gp_after_ads": float("nan")}
    gp = part["gp_after_ads"].sum() if part["gp_after_ads"].notna().any() else float("nan")
    return {
        "spend": float(part["spend"].sum()),
        "revenue": float(part["revenue"].sum()),
        "orders": float(part["orders"].sum()),
        "impressions": float(part["impressions"].sum()),
        "clicks": float(part["clicks"].sum()),
        "gp_after_ads": gp,
    }


def grew_enough(after: float, before: float, factor: float) -> bool:
    before = float(before or 0)
    after = float(after or 0)
    if before == 0:
        return after > 0
    return after >= before * factor


def retained_enough(after: float, before: float, factor: float) -> bool:
    before = float(before or 0)
    after = float(after or 0)
    if before == 0:
        return after >= 0
    return after >= before * factor


def ge_metric(after: float, before: float) -> bool:
    if pd.isna(after) and pd.isna(before):
        return True
    if pd.isna(after) or pd.isna(before):
        return False
    return float(after) >= float(before)


def ge_metric_factor(after: float, before: float, factor: float) -> bool:
    if pd.isna(after) and pd.isna(before):
        return True
    if pd.isna(after) or pd.isna(before):
        return False
    return float(after) >= float(before) * factor


def lt_metric_factor(after: float, before: float, factor: float) -> bool:
    if pd.isna(after) or pd.isna(before):
        return False
    return float(after) < float(before) * factor



def bid_effect_empty_frame() -> pd.DataFrame:
    return pd.DataFrame(columns=[
        "event_id", "campaign_id", "nm_id", "supplier_article", "subject_norm", "placement", "direction",
        "event_date", "old_bid_rub", "new_bid_rub", "reason_code",
        "before_window_days", "after_d1_days", "after_d3_days", "after_d7_days",
        "before_spend", "after_d1_spend", "after_d3_spend", "after_d7_spend",
        "before_spend_per_day", "after_d1_spend_per_day", "after_d3_spend_per_day", "after_d7_spend_per_day",
        "before_revenue", "after_d1_revenue", "after_d3_revenue", "after_d7_revenue",
        "before_revenue_per_day", "after_d1_revenue_per_day", "after_d3_revenue_per_day", "after_d7_revenue_per_day",
        "before_orders", "after_d1_orders", "after_d3_orders", "after_d7_orders",
        "before_orders_per_day", "after_d1_orders_per_day", "after_d3_orders_per_day", "after_d7_orders_per_day",
        "before_impressions", "after_d1_impressions", "after_d3_impressions", "after_d7_impressions",
        "before_impressions_per_day", "after_d1_impressions_per_day", "after_d3_impressions_per_day", "after_d7_impressions_per_day",
        "before_clicks", "after_d1_clicks", "after_d3_clicks", "after_d7_clicks",
        "before_clicks_per_day", "after_d1_clicks_per_day", "after_d3_clicks_per_day", "after_d7_clicks_per_day",
        "before_ctr_pct", "after_d1_ctr_pct", "after_d3_ctr_pct", "after_d7_ctr_pct",
        "before_drr_pct", "after_d1_drr_pct", "after_d3_drr_pct", "after_d7_drr_pct",
        "before_gp", "after_d1_gp", "after_d3_gp", "after_d7_gp",
        "before_gp_per_day", "after_d1_gp_per_day", "after_d3_gp_per_day", "after_d7_gp_per_day",
        "d1_verdict", "d3_verdict", "d7_verdict", "final_verdict", "postcheck_status",
        "target_bid_action", "target_bid_action_text", "recommended_next_bid_rub",
        "traffic_delta_d1_pct", "traffic_delta_d3_pct", "orders_delta_d3_pct", "revenue_delta_d3_pct", "gp_delta_d3_pct",
        # обратная совместимость со старым листом
        "drr_before", "drr_after_d3", "impressions_after_d1", "clicks_after_d1", "orders_after_d3",
        "revenue_after_d3", "spend_after_d3", "gp_after_d3",
    ])


def _event_num(row: pd.Series | Dict[str, Any], col: str, default: float = 0.0) -> float:
    value = pd.to_numeric(pd.Series([row.get(col, default)]), errors="coerce").iloc[0]
    if pd.isna(value):
        return default
    return float(value)


def _per_day(metrics: Dict[str, float], days: float) -> Dict[str, float]:
    d = max(float(days or 1.0), 1.0)
    return {
        "spend": float(metrics.get("spend", 0.0) or 0.0) / d,
        "revenue": float(metrics.get("revenue", 0.0) or 0.0) / d,
        "orders": float(metrics.get("orders", 0.0) or 0.0) / d,
        "impressions": float(metrics.get("impressions", 0.0) or 0.0) / d,
        "clicks": float(metrics.get("clicks", 0.0) or 0.0) / d,
        "gp_after_ads": (float(metrics.get("gp_after_ads", 0.0)) / d) if not pd.isna(metrics.get("gp_after_ads", float("nan"))) else float("nan"),
    }


def _pct_delta(after: float, before: float) -> Optional[float]:
    before = float(before or 0.0)
    after = float(after or 0.0)
    if before == 0:
        return None
    return (after / before - 1.0) * 100.0


def _metric_ratio(after: float, before: float) -> float:
    before = float(before or 0.0)
    after = float(after or 0.0)
    if before == 0:
        return 1.0 if after >= 0 else 0.0
    return after / before


def _postcheck_target_action(
    verdict: str,
    direction: str,
    is_ramp_event: bool,
    current_new_bid: float,
    step: float,
) -> Tuple[str, str, Any]:
    verdict = _clean_text_value(verdict)
    direction = _clean_text_value(direction).lower()
    if not verdict:
        return "WAIT_POSTCHECK", "ждём созревания данных для оценки изменения ставки", ""
    if verdict.startswith("WAIT") or verdict in {"RAMP_D3_MONITOR_WEEK"}:
        return "WAIT_POSTCHECK", "ждём финальную оценку изменения ставки", ""
    if verdict in {"RAISE_BAD", "RAMP_NEGATIVE_GP_D7", "RAMP_SPEND_OVER_LIMIT_D7"}:
        next_bid = round(max(float(current_new_bid or 0.0) - float(step or 1.0), TECHNICAL_BID_FLOOR_RUB), 2)
        return "REVERT_TO_PREVIOUS_BID", "эффект плохой: откатить ставку на предыдущий уровень", next_bid
    if verdict == "LOWER_BAD":
        next_bid = round(float(current_new_bid or 0.0) + float(step or 1.0), 2)
        return "REVERT_TO_PREVIOUS_BID", "снижение ухудшило трафик/экономику: откатить ставку вверх", next_bid
    if verdict in {"RAISE_GOOD"}:
        return "HOLD_BID_LEVEL", "повышение сработало: оставить новый уровень, следующий шаг только по обычной логике ДРР/разгона", ""
    if verdict in {"RAISE_D3_MIXED", "RAISE_NO_TRAFFIC_GROWTH"}:
        return "HOLD_BID_LEVEL", "эффект неоднозначный: оставить ставку, не делать автоматический откат без ухудшения экономики", ""
    if verdict in {"LOWER_GOOD"}:
        return "HOLD_BID_LEVEL", "снижение сработало: оставить новый уровень; дальше снижать только если ДРР всё ещё выше лимита", ""
    if verdict in {"LOWER_D3_MIXED"}:
        return "HOLD_BID_LEVEL", "эффект снижения смешанный: оставить ставку до следующего окна", ""
    if verdict == "RAMP_GOOD_D7":
        return "STOP_RAMP_HOLD", "цель разгона достигнута: остановить разгон и оставить ставку", ""
    if verdict == "RAMP_NEEDS_MORE_BID_D7":
        next_bid = round(float(current_new_bid or 0.0) + float(step or 1.0), 2)
        return "CONTINUE_RAMP", "цель 1000 показов/день не достигнута, расход в лимите: продолжить разгон следующим шагом", next_bid
    return "HOLD_BID_LEVEL", f"вердикт {verdict}: оставить ставку до следующего окна", ""


def evaluate_postchecks(ads_df: pd.DataFrame, bid_history: pd.DataFrame, ctx: RunContext) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Оценивает эффект изменения ставки нормальным сравнением ДО/ПОСЛЕ.

    Важно: значения ДО в истории — это 5-дневная база на момент изменения ставки.
    Старый код сравнивал D+1 за один день с этой 5-дневной суммой и мог ошибочно ставить
    RAISE_NO_TRAFFIC_GROWTH. Здесь все сравнения делаются в формате среднее/день:
    - ДО: сумма за 5 дней / 5;
    - D+1: один зрелый день;
    - D+3: сумма за 3 дня / 3;
    - D+7 для разгона: сумма за 7 дней / 7.
    """
    if bid_history.empty:
        return bid_history.copy(), bid_effect_empty_frame()

    updated = coerce_history_columns_object(bid_history.copy(), BID_HISTORY_COLUMNS)
    for col in BID_HISTORY_COLUMNS:
        if col not in updated.columns:
            updated[col] = pd.Series([""] * len(updated), index=updated.index, dtype="object")
        else:
            updated[col] = updated[col].astype("object")

    effects: List[Dict[str, Any]] = []
    before_days = float(ANALYSIS_WINDOW_DAYS)

    for idx, row in updated.iterrows():
        event_date = pd.to_datetime(row.get("event_date"), errors="coerce")
        if pd.isna(event_date):
            continue
        event_day = event_date.date()
        key = make_key(row)
        direction = _clean_text_value(row.get("direction", "")).lower()
        event_reason_code = _clean_text_value(row.get("reason_code", ""))
        is_ramp_event = is_ramp_event_reason(event_reason_code)
        step, _ = bid_step_rub(row.get("placement", ""))
        drr_limit = drr_limit_for_subject(row.get("subject_norm", ""))

        status = _clean_text_value(row.get("postcheck_status", "")) or "pending"
        final_verdict = _clean_text_value(row.get("final_verdict", ""))
        d1_verdict = _clean_text_value(row.get("d1_verdict", ""))
        d3_verdict = _clean_text_value(row.get("d3_verdict", ""))
        d7_verdict = _clean_text_value(row.get("d7_verdict", ""))

        before = {
            "spend": _event_num(row, "spend_before", 0.0),
            "revenue": _event_num(row, "revenue_before", 0.0),
            "orders": _event_num(row, "orders_before", 0.0),
            "impressions": _event_num(row, "impressions_before", 0.0),
            "clicks": _event_num(row, "clicks_before", 0.0),
            "gp_after_ads": _event_num(row, "gp_before", float("nan")),
        }
        before_drr = _event_num(row, "drr_before", safe_drr_pct(before["spend"], before["revenue"]))
        before_daily = _per_day(before, before_days)
        before_ctr = safe_ctr_pct(before["clicks"], before["impressions"])

        d1_after = {"spend": 0.0, "revenue": 0.0, "orders": 0.0, "impressions": 0.0, "clicks": 0.0, "gp_after_ads": float("nan")}
        d3_after = {"spend": 0.0, "revenue": 0.0, "orders": 0.0, "impressions": 0.0, "clicks": 0.0, "gp_after_ads": float("nan")}
        d7_after = {"spend": 0.0, "revenue": 0.0, "orders": 0.0, "impressions": 0.0, "clicks": 0.0, "gp_after_ads": float("nan")}
        d1_daily = _per_day(d1_after, 1.0)
        d3_daily = _per_day(d3_after, 3.0)
        d7_daily = _per_day(d7_after, float(RAMP_CHECK_DAYS))
        drr_after_d1 = 0.0
        drr_after_d3 = 0.0
        drr_after_d7 = 0.0

        # D+1: только ранний сигнал по трафику, без финального отката обычного повышения.
        d1_day = event_day + timedelta(days=1)
        if ctx.mature_end >= d1_day and status not in {"resolved"}:
            d1_after = aggregate_after_event(ads_df, key, d1_day, d1_day)
            d1_daily = _per_day(d1_after, 1.0)
            drr_after_d1 = safe_drr_pct(d1_after["spend"], d1_after["revenue"])
            d1_traffic_grew = (
                grew_enough(d1_daily["impressions"], before_daily["impressions"], 1.05)
                or grew_enough(d1_daily["clicks"], before_daily["clicks"], 1.05)
            )
            if direction == "raise":
                if is_ramp_event:
                    d1_verdict = "RAMP_D1_TRAFFIC_GROWTH" if d1_traffic_grew else "RAMP_D1_WAIT_WEEK"
                    status = "d1_done"
                else:
                    d1_verdict = "RAISE_D1_OK" if d1_traffic_grew else "RAISE_D1_NO_TRAFFIC_YET"
                    status = "d1_done"
            elif direction == "lower":
                traffic_retained = (
                    retained_enough(d1_daily["impressions"], before_daily["impressions"], 0.80)
                    and retained_enough(d1_daily["clicks"], before_daily["clicks"], 0.80)
                )
                d1_verdict = "LOWER_D1_OK" if traffic_retained else "LOWER_TRAFFIC_DROP_RISK"
                status = "d1_done"
            updated.at[idx, "d1_verdict"] = d1_verdict
            updated.at[idx, "d1_check_date"] = str(ctx.mature_end)

        # D+3: финальная оценка обычного повышения/снижения. Сравнение только среднее/день.
        d3_end = event_day + timedelta(days=3)
        if ctx.mature_end >= d3_end and status not in {"resolved"}:
            d3_after = aggregate_after_event(ads_df, key, event_day + timedelta(days=1), d3_end)
            d3_daily = _per_day(d3_after, 3.0)
            drr_after_d3 = safe_drr_pct(d3_after["spend"], d3_after["revenue"])
            d3_traffic_grew = (
                grew_enough(d3_daily["impressions"], before_daily["impressions"], 1.05)
                or grew_enough(d3_daily["clicks"], before_daily["clicks"], 1.05)
            )
            traffic_retained = (
                retained_enough(d3_daily["impressions"], before_daily["impressions"], 0.80)
                and retained_enough(d3_daily["clicks"], before_daily["clicks"], 0.80)
            )
            orders_retained = retained_enough(d3_daily["orders"], before_daily["orders"], 0.90)
            revenue_retained = retained_enough(d3_daily["revenue"], before_daily["revenue"], 0.90)
            gp_retained = ge_metric_factor(d3_daily["gp_after_ads"], before_daily["gp_after_ads"], 0.90)
            gp_bad = lt_metric_factor(d3_daily["gp_after_ads"], before_daily["gp_after_ads"], 0.80)
            drr_ok = drr_after_d3 <= drr_limit

            if direction == "raise":
                if is_ramp_event:
                    d3_verdict = "RAMP_D3_MONITOR_WEEK"
                    final_verdict = ""
                    status = "d3_done"
                elif (not drr_ok) and (not orders_retained or not revenue_retained or gp_bad):
                    d3_verdict = "RAISE_BAD"
                    final_verdict = "RAISE_BAD"
                    status = "resolved"
                elif d3_traffic_grew and drr_ok and orders_retained and revenue_retained and gp_retained:
                    d3_verdict = "RAISE_GOOD"
                    final_verdict = "RAISE_GOOD"
                    status = "resolved"
                elif (not d3_traffic_grew) and drr_ok and orders_retained and revenue_retained:
                    d3_verdict = "RAISE_NO_TRAFFIC_GROWTH"
                    final_verdict = "RAISE_NO_TRAFFIC_GROWTH"
                    status = "resolved"
                else:
                    d3_verdict = "RAISE_D3_MIXED"
                    final_verdict = "RAISE_D3_MIXED"
                    status = "resolved"
            elif direction == "lower":
                if drr_after_d3 < before_drr and orders_retained and gp_retained and traffic_retained:
                    d3_verdict = "LOWER_GOOD"
                    final_verdict = "LOWER_GOOD"
                elif (not traffic_retained or not orders_retained or gp_bad) and drr_after_d3 >= before_drr:
                    d3_verdict = "LOWER_BAD"
                    final_verdict = "LOWER_BAD"
                else:
                    d3_verdict = "LOWER_D3_MIXED"
                    final_verdict = "LOWER_D3_MIXED"
                status = "resolved"
            updated.at[idx, "d3_verdict"] = d3_verdict
            updated.at[idx, "d3_check_date"] = str(ctx.mature_end)

        # D+7: финальная оценка разгона.
        if is_ramp_event:
            d7_end = event_day + timedelta(days=RAMP_CHECK_DAYS)
            if ctx.mature_end >= d7_end and status not in {"resolved"}:
                d7_after = aggregate_after_event(ads_df, key, event_day + timedelta(days=1), d7_end)
                d7_daily = _per_day(d7_after, float(RAMP_CHECK_DAYS))
                drr_after_d7 = safe_drr_pct(d7_after["spend"], d7_after["revenue"])
                avg_imp_d7 = d7_daily["impressions"]
                avg_spend_d7 = d7_daily["spend"]
                gp_d7_daily = d7_daily["gp_after_ads"]
                if not pd.isna(gp_d7_daily) and float(gp_d7_daily) < 0:
                    d7_verdict = "RAMP_NEGATIVE_GP_D7"
                    final_verdict = "RAMP_NEGATIVE_GP_D7"
                elif avg_spend_d7 > RAMP_MAX_SPEND_PER_DAY:
                    d7_verdict = "RAMP_SPEND_OVER_LIMIT_D7"
                    final_verdict = "RAMP_SPEND_OVER_LIMIT_D7"
                elif avg_imp_d7 >= RAMP_TARGET_IMPRESSIONS_PER_DAY and (pd.isna(gp_d7_daily) or float(gp_d7_daily) >= 0):
                    d7_verdict = "RAMP_GOOD_D7"
                    final_verdict = "RAMP_GOOD_D7"
                else:
                    d7_verdict = "RAMP_NEEDS_MORE_BID_D7"
                    final_verdict = "RAMP_NEEDS_MORE_BID_D7"
                updated.at[idx, "d7_verdict"] = d7_verdict
                updated.at[idx, "d7_check_date"] = str(ctx.mature_end)
                status = "resolved"

        # Для отчёта считаем доступные after-окна независимо от текущего status, если данные уже зрелые.
        if ctx.mature_end >= d1_day:
            d1_after = aggregate_after_event(ads_df, key, d1_day, d1_day)
            d1_daily = _per_day(d1_after, 1.0)
            drr_after_d1 = safe_drr_pct(d1_after["spend"], d1_after["revenue"])
        if ctx.mature_end >= d3_end:
            d3_after = aggregate_after_event(ads_df, key, event_day + timedelta(days=1), d3_end)
            d3_daily = _per_day(d3_after, 3.0)
            drr_after_d3 = safe_drr_pct(d3_after["spend"], d3_after["revenue"])
        if is_ramp_event:
            d7_end = event_day + timedelta(days=RAMP_CHECK_DAYS)
            if ctx.mature_end >= d7_end:
                d7_after = aggregate_after_event(ads_df, key, event_day + timedelta(days=1), d7_end)
                d7_daily = _per_day(d7_after, float(RAMP_CHECK_DAYS))
                drr_after_d7 = safe_drr_pct(d7_after["spend"], d7_after["revenue"])

        updated.at[idx, "postcheck_status"] = status
        updated.at[idx, "final_verdict"] = final_verdict

        target_action, target_text, recommended_next_bid = _postcheck_target_action(
            final_verdict, direction, is_ramp_event, _event_num(row, "new_bid_rub", 0.0), step
        )

        after_d1_ctr = safe_ctr_pct(d1_after["clicks"], d1_after["impressions"])
        after_d3_ctr = safe_ctr_pct(d3_after["clicks"], d3_after["impressions"])
        after_d7_ctr = safe_ctr_pct(d7_after["clicks"], d7_after["impressions"])

        effects.append({
            "event_id": row.get("event_id", ""),
            "campaign_id": row.get("campaign_id", ""),
            "nm_id": row.get("nm_id", ""),
            "supplier_article": row.get("supplier_article", ""),
            "subject_norm": row.get("subject_norm", ""),
            "placement": row.get("placement", ""),
            "direction": direction,
            "event_date": row.get("event_date", ""),
            "old_bid_rub": row.get("old_bid_rub", ""),
            "new_bid_rub": row.get("new_bid_rub", ""),
            "reason_code": event_reason_code,
            "before_window_days": before_days,
            "after_d1_days": 1 if ctx.mature_end >= d1_day else 0,
            "after_d3_days": 3 if ctx.mature_end >= d3_end else 0,
            "after_d7_days": RAMP_CHECK_DAYS if is_ramp_event and ctx.mature_end >= event_day + timedelta(days=RAMP_CHECK_DAYS) else 0,
            "before_spend": before["spend"], "after_d1_spend": d1_after["spend"], "after_d3_spend": d3_after["spend"], "after_d7_spend": d7_after["spend"],
            "before_spend_per_day": before_daily["spend"], "after_d1_spend_per_day": d1_daily["spend"], "after_d3_spend_per_day": d3_daily["spend"], "after_d7_spend_per_day": d7_daily["spend"],
            "before_revenue": before["revenue"], "after_d1_revenue": d1_after["revenue"], "after_d3_revenue": d3_after["revenue"], "after_d7_revenue": d7_after["revenue"],
            "before_revenue_per_day": before_daily["revenue"], "after_d1_revenue_per_day": d1_daily["revenue"], "after_d3_revenue_per_day": d3_daily["revenue"], "after_d7_revenue_per_day": d7_daily["revenue"],
            "before_orders": before["orders"], "after_d1_orders": d1_after["orders"], "after_d3_orders": d3_after["orders"], "after_d7_orders": d7_after["orders"],
            "before_orders_per_day": before_daily["orders"], "after_d1_orders_per_day": d1_daily["orders"], "after_d3_orders_per_day": d3_daily["orders"], "after_d7_orders_per_day": d7_daily["orders"],
            "before_impressions": before["impressions"], "after_d1_impressions": d1_after["impressions"], "after_d3_impressions": d3_after["impressions"], "after_d7_impressions": d7_after["impressions"],
            "before_impressions_per_day": before_daily["impressions"], "after_d1_impressions_per_day": d1_daily["impressions"], "after_d3_impressions_per_day": d3_daily["impressions"], "after_d7_impressions_per_day": d7_daily["impressions"],
            "before_clicks": before["clicks"], "after_d1_clicks": d1_after["clicks"], "after_d3_clicks": d3_after["clicks"], "after_d7_clicks": d7_after["clicks"],
            "before_clicks_per_day": before_daily["clicks"], "after_d1_clicks_per_day": d1_daily["clicks"], "after_d3_clicks_per_day": d3_daily["clicks"], "after_d7_clicks_per_day": d7_daily["clicks"],
            "before_ctr_pct": before_ctr, "after_d1_ctr_pct": after_d1_ctr, "after_d3_ctr_pct": after_d3_ctr, "after_d7_ctr_pct": after_d7_ctr,
            "before_drr_pct": before_drr, "after_d1_drr_pct": drr_after_d1, "after_d3_drr_pct": drr_after_d3, "after_d7_drr_pct": drr_after_d7,
            "before_gp": before["gp_after_ads"], "after_d1_gp": d1_after["gp_after_ads"], "after_d3_gp": d3_after["gp_after_ads"], "after_d7_gp": d7_after["gp_after_ads"],
            "before_gp_per_day": before_daily["gp_after_ads"], "after_d1_gp_per_day": d1_daily["gp_after_ads"], "after_d3_gp_per_day": d3_daily["gp_after_ads"], "after_d7_gp_per_day": d7_daily["gp_after_ads"],
            "d1_verdict": d1_verdict,
            "d3_verdict": d3_verdict,
            "d7_verdict": _clean_text_value(updated.at[idx, "d7_verdict"]) if "d7_verdict" in updated.columns else d7_verdict,
            "final_verdict": final_verdict,
            "postcheck_status": status,
            "target_bid_action": target_action,
            "target_bid_action_text": target_text,
            "recommended_next_bid_rub": recommended_next_bid,
            "traffic_delta_d1_pct": _pct_delta(d1_daily["impressions"], before_daily["impressions"]),
            "traffic_delta_d3_pct": _pct_delta(d3_daily["impressions"], before_daily["impressions"]),
            "orders_delta_d3_pct": _pct_delta(d3_daily["orders"], before_daily["orders"]),
            "revenue_delta_d3_pct": _pct_delta(d3_daily["revenue"], before_daily["revenue"]),
            "gp_delta_d3_pct": _pct_delta(d3_daily["gp_after_ads"], before_daily["gp_after_ads"]),
            # обратная совместимость
            "drr_before": before_drr,
            "drr_after_d3": drr_after_d3,
            "impressions_after_d1": d1_after.get("impressions", 0.0),
            "clicks_after_d1": d1_after.get("clicks", 0.0),
            "orders_after_d3": d3_after.get("orders", 0.0),
            "revenue_after_d3": d3_after.get("revenue", 0.0),
            "spend_after_d3": d3_after.get("spend", 0.0),
            "gp_after_d3": d3_after.get("gp_after_ads", float("nan")),
        })

    effect_df = pd.DataFrame(effects)
    if effect_df.empty:
        effect_df = bid_effect_empty_frame()
    return updated[BID_HISTORY_COLUMNS], effect_df


# =============================
# Решения по ставкам
# =============================

def is_active_campaign(status_value: Any) -> bool:
    text = _clean_text_value(status_value).replace("ё", "е").lower()
    if not text:
        return False
    if text in {"9", "9.0"}:
        return True
    if "active" in text:
        return True
    if "актив" in text and "неактив" not in text and "не актив" not in text:
        return True
    return False


def is_managed_subject(subject_value: Any) -> bool:
    text = normalize_subject_value(subject_value).lower()
    return text in MANAGED_SUBJECTS


def drr_limit_for_subject(subject_value: Any) -> float:
    """Возвращает допустимый ДРР для предмета в процентах."""
    subject = normalize_subject_value(subject_value).lower()
    return float(SUBJECT_DRR_LIMITS.get(subject, DRR_LIMIT_PCT))


def is_pause_allowed_subject(subject_value: Any) -> bool:
    """Кисти не паузим никогда; пауза только для помад/блесков/карандашей."""
    subject = normalize_subject_value(subject_value).lower()
    return subject in PAUSE_ALLOWED_SUBJECTS


def is_one_campaign_experiment_subject(subject_value: Any) -> bool:
    subject = normalize_subject_value(subject_value).lower()
    return subject in ONE_CAMPAIGN_EXPERIMENT_SUBJECTS


def product_group_from_article(value: Any) -> str:
    """617/1, PT617.001, 617_1 -> 617. Если артикул пустой — пустая группа."""
    text = _clean_text_value(value).upper().replace(" ", "")
    if not text:
        return ""
    m = re.search(r"(\d{3,4})", text)
    if m:
        return m.group(1).lstrip("0") or m.group(1)
    return text.split("/")[0].split(".")[0].split("_")[0][:32]


def safe_window_days(ctx: Optional[RunContext] = None) -> float:
    if ctx is None:
        return float(ANALYSIS_WINDOW_DAYS)
    days = (ctx.current_end - ctx.current_start).days + 1
    return float(days if days > 0 else ANALYSIS_WINDOW_DAYS)


def money_or_zero(value: Any) -> float:
    return float(pd.to_numeric(pd.Series([value]), errors="coerce").fillna(0).iloc[0])


def filter_managed_subject_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Оставляет только 4 управляемых предмета.

    Жёсткое правило проекта: скрипт не должен принимать решения, строить паузы,
    post-check или API-вызовы по любым другим предметам.
    """
    if df is None or df.empty:
        return pd.DataFrame() if df is None else df.copy()
    if "subject_norm" not in df.columns:
        return df.iloc[0:0].copy()
    result = df.copy()
    result["subject_norm"] = result["subject_norm"].map(normalize_subject_value)
    mask = result["subject_norm"].map(is_managed_subject)
    return result.loc[mask].copy()


def make_key_set_from_ads(ads_df: pd.DataFrame) -> set[Tuple[str, str, str]]:
    if ads_df is None or ads_df.empty:
        return set()
    required = {"campaign_id", "nm_id", "placement"}
    if not required.issubset(set(ads_df.columns)):
        return set()
    keys: set[Tuple[str, str, str]] = set()
    for _, row in ads_df.iterrows():
        key = make_key(row)
        if all(key):
            keys.add(key)
    return keys


def filter_bid_history_managed_only(bid_history: pd.DataFrame, ads_df: pd.DataFrame) -> pd.DataFrame:
    if bid_history is None or bid_history.empty:
        return pd.DataFrame(columns=BID_HISTORY_COLUMNS)
    result = bid_history.copy()
    for col in BID_HISTORY_COLUMNS:
        if col not in result.columns:
            result[col] = ""
    if "subject_norm" in result.columns and result["subject_norm"].map(_clean_text_value).ne("").any():
        result = result.loc[result["subject_norm"].map(is_managed_subject)].copy()
    else:
        managed_keys = make_key_set_from_ads(ads_df)
        result = result.loc[result.apply(lambda r: make_key(r) in managed_keys, axis=1)].copy() if managed_keys else result.iloc[0:0].copy()
    return result[BID_HISTORY_COLUMNS]


def filter_pause_history_managed_only(pause_history: pd.DataFrame, ads_df: pd.DataFrame) -> pd.DataFrame:
    if pause_history is None or pause_history.empty:
        return pd.DataFrame(columns=PAUSE_HISTORY_COLUMNS)
    result = pause_history.copy()
    for col in PAUSE_HISTORY_COLUMNS:
        if col not in result.columns:
            result[col] = ""
    managed_keys = make_key_set_from_ads(ads_df)
    if not managed_keys:
        return result.iloc[0:0][PAUSE_HISTORY_COLUMNS].copy()
    result = result.loc[result.apply(lambda r: make_key(r) in managed_keys, axis=1)].copy()
    return result[PAUSE_HISTORY_COLUMNS]


def bid_step_rub(placement: Any) -> Tuple[float, str]:
    placement_norm = normalize_placement_value(placement)
    if placement_norm in {"search", "recommendations"}:
        return 1.0, ""
    if placement_norm == "combined":
        return 5.0, ""
    return 1.0, "UNKNOWN_PLACEMENT_DEFAULT_STEP"


def format_float(value: Any, digits: int = 2) -> str:
    if value is None or pd.isna(value):
        return "н/д"
    return f"{float(value):.{digits}f}"


def build_reason_text(row: pd.Series, action: str, new_bid: Optional[float], extra: str = "") -> str:
    old_bid = row.get("current_bid_rub", float("nan"))
    parts = [
        f"ДРР={format_float(row.get('campaign_drr_pct'), 2)}%",
        f"ставка={format_float(old_bid, 2)} ₽",
        f"мин. WB={format_float(row.get('min_bid_rub'), 2)} ₽" if 'min_bid_rub' in row.index and not pd.isna(row.get('min_bid_rub')) else "мин. WB=н/д",
        f"новая ставка={format_float(new_bid, 2)} ₽" if new_bid is not None else "новая ставка=н/д",
        f"расход={format_float(row.get('spend'), 2)} ₽",
        f"выручка={format_float(row.get('revenue'), 2)} ₽",
        f"заказы={format_float(row.get('orders'), 0)}",
    ]
    if extra:
        parts.append(extra)
    return f"{action}: " + "; ".join(parts)


def technical_hold(reason_code: str, row: pd.Series, reason: str) -> Dict[str, Any]:
    return {
        "action": "Без изменений",
        "new_bid_rub": None,
        "reason_code": reason_code,
        "reason_text": build_reason_text(row, "Без изменений", None, reason),
        "pause_decision": "",
    }


def decide_action(row: pd.Series, pending_event: Optional[Dict[str, Any]] = None, postcheck_result: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """
    Возвращает dict:
    {
        "action": "Повысить" | "Снизить" | "Без изменений",
        "new_bid_rub": float | None,
        "reason_code": str,
        "reason_text": str,
        "pause_decision": str | ""
    }
    """
    campaign_id = _clean_id_value(row.get("campaign_id", ""))
    nm_id = _clean_id_value(row.get("nm_id", ""))
    placement = normalize_placement_value(row.get("placement", ""))
    current_bid = row.get("current_bid_rub", float("nan"))

    if not campaign_id or not nm_id or not placement or pd.isna(current_bid) or float(current_bid) <= 0:
        return technical_hold("MISSING_KEY", row, "нет campaign_id / nm_id / placement или текущей ставки")

    if not is_active_campaign(row.get("campaign_status", "")):
        return technical_hold("NOT_ACTIVE", row, "кампания не активна")

    if not is_managed_subject(row.get("subject_norm", "")):
        return technical_hold("NOT_MANAGED_SUBJECT", row, "предмет не входит в управляемые")

    final_verdict = _clean_text_value((postcheck_result or {}).get("final_verdict", ""))
    previous_event_id = _clean_text_value((postcheck_result or {}).get("event_id", ""))

    step, step_reason = bid_step_rub(placement)
    current_bid_float = float(current_bid)
    drr_limit = drr_limit_for_subject(row.get("subject_norm", ""))
    impressions = money_or_zero(row.get("impressions", 0))
    spend = money_or_zero(row.get("spend", 0))
    revenue = money_or_zero(row.get("revenue", 0))
    orders = money_or_zero(row.get("orders", 0))
    avg_impressions_per_day = money_or_zero(row.get("avg_impressions_per_day", impressions / ANALYSIS_WINDOW_DAYS))
    avg_spend_per_day = money_or_zero(row.get("avg_spend_per_day", spend / ANALYSIS_WINDOW_DAYS))

    if final_verdict in {"RAISE_BAD", "RAISE_NO_TRAFFIC_GROWTH", "RAMP_NEGATIVE_GP_D7", "RAMP_SPEND_OVER_LIMIT_D7"}:
        new_bid = round(current_bid_float - step, 2)
        reason_code = "RAISE_FAILED_REVERT"
        if step_reason:
            reason_code += f"__{step_reason}"
        if new_bid < TECHNICAL_BID_FLOOR_RUB:
            return {
                "action": "Без изменений",
                "new_bid_rub": None,
                "reason_code": "TECHNICAL_FLOOR_REACHED",
                "reason_text": build_reason_text(row, "Без изменений", None, f"откат после {final_verdict}; новая ставка ниже 1 ₽; previous_event_id={previous_event_id}"),
                "pause_decision": "PAUSE_CANDIDATE",
            }
        return {
            "action": "Снизить",
            "new_bid_rub": new_bid,
            "reason_code": reason_code,
            "reason_text": build_reason_text(row, "Снизить", new_bid, f"откат после {final_verdict}; previous_event_id={previous_event_id}"),
            "pause_decision": "",
        }

    if final_verdict == "LOWER_BAD":
        new_bid = round(current_bid_float + step, 2)
        reason_code = "LOWER_FAILED_REVERT"
        if step_reason:
            reason_code += f"__{step_reason}"
        return {
            "action": "Повысить",
            "new_bid_rub": new_bid,
            "reason_code": reason_code,
            "reason_text": build_reason_text(row, "Повысить", new_bid, f"откат после {final_verdict}; previous_event_id={previous_event_id}"),
            "pause_decision": "",
        }

    if pending_event is not None:
        wait_rule = _clean_text_value(pending_event.get("wait_rule", "WAIT_POSTCHECK"))
        wait_until = _clean_text_value(pending_event.get("wait_until_date", ""))
        wait_days_left = _clean_text_value(pending_event.get("wait_days_left", ""))
        event_date = _clean_text_value(pending_event.get("event_date", ""))
        old_bid = _clean_text_value(pending_event.get("old_bid_rub", ""))
        new_bid_prev = _clean_text_value(pending_event.get("new_bid_rub", ""))
        prev_reason = _clean_text_value(pending_event.get("reason_code", ""))
        return technical_hold(
            wait_rule,
            row,
            f"ждём post-check: event_id={pending_event.get('event_id', '')}; "
            f"последняя правка={event_date}; ставка {old_bid}→{new_bid_prev}; "
            f"правило={prev_reason}; ждём до {wait_until}; осталось дней={wait_days_left}"
        )

    # Отдельная механика разгона слабых кампаний: если ставка настолько низкая,
    # что нет расхода/заказов или нет 1000 показов в день при расходе до 500 ₽/день, повышаем ставку и смотрим неделю.
    if avg_impressions_per_day < RAMP_TARGET_IMPRESSIONS_PER_DAY and avg_spend_per_day <= RAMP_MAX_SPEND_PER_DAY:
        new_bid = round(current_bid_float + step, 2)
        reason_code = "RAMP_TO_1000_IMPRESSIONS_PER_DAY"
        if spend == 0 and orders == 0:
            reason_code = "LOW_BID_NO_SPEND_NO_ORDERS_RAMP"
        if step_reason:
            reason_code += f"__{step_reason}"
        return {
            "action": "Повысить",
            "new_bid_rub": new_bid,
            "reason_code": reason_code,
            "reason_text": build_reason_text(row, "Повысить", new_bid, f"разгон показов: {avg_impressions_per_day:.0f}/день < {RAMP_TARGET_IMPRESSIONS_PER_DAY:.0f}/день; расход {avg_spend_per_day:.0f} ₽/день <= {RAMP_MAX_SPEND_PER_DAY:.0f} ₽/день; проверка D+{RAMP_CHECK_DAYS}"),
            "pause_decision": "",
        }

    drr = float(row.get("campaign_drr_pct", 0) or 0)
    if drr >= drr_limit:
        new_bid = round(current_bid_float - step, 2)
        reason_code = "DRR_GE_LIMIT_REDUCE"
        if step_reason:
            reason_code += f"__{step_reason}"
        if new_bid < TECHNICAL_BID_FLOOR_RUB:
            return {
                "action": "Без изменений",
                "new_bid_rub": None,
                "reason_code": "TECHNICAL_FLOOR_REACHED",
                "reason_text": build_reason_text(row, "Без изменений", None, "требуется снижение, но новая ставка ниже 1 ₽"),
                "pause_decision": "",
            }
        return {
            "action": "Снизить",
            "new_bid_rub": new_bid,
            "reason_code": reason_code,
            "reason_text": build_reason_text(row, "Снизить", new_bid, f"ДРР >= допустимого порога {drr_limit:.1f}%"),
            "pause_decision": "",
        }

    new_bid = round(current_bid_float + step, 2)
    reason_code = "DRR_LT_LIMIT_GROW"
    if step_reason:
        reason_code += f"__{step_reason}"
    return {
        "action": "Повысить",
        "new_bid_rub": new_bid,
        "reason_code": reason_code,
        "reason_text": build_reason_text(row, "Повысить", new_bid, f"ДРР < допустимого порога {drr_limit:.1f}%"),
        "pause_decision": "",
    }



def business_min_bid_rub(placement: Any) -> float:
    """Бизнес-минимум для отчёта и решений: поиск/CPC не ниже 4 ₽, полки/combined не ниже 80 ₽."""
    placement_norm = normalize_placement_value(placement)
    if placement_norm == "combined":
        return 80.0
    if placement_norm in {"search", "recommendations"}:
        return 4.0
    return 1.0


def is_ramp_related_reason(reason_code: Any, wait_rule: Any = "", last_reason_code: Any = "") -> bool:
    text = " ".join([
        _clean_text_value(reason_code).upper(),
        _clean_text_value(wait_rule).upper(),
        _clean_text_value(last_reason_code).upper(),
    ])
    return "RAMP" in text or "РАЗГОН" in text or "LOW_BID_NO_SPEND_NO_ORDERS" in text or "WAIT_D7" in text


def is_ramp_candidate_by_metrics(row: pd.Series | Dict[str, Any]) -> bool:
    """Кампания подходит под режим разгона по текущим метрикам: мало показов и расход в лимите."""
    if not is_active_campaign(row.get("campaign_status", "")):
        return False
    if not is_managed_subject(row.get("subject_norm", "")):
        return False
    avg_imp = money_or_zero(row.get("avg_impressions_per_day", 0))
    avg_spend = money_or_zero(row.get("avg_spend_per_day", 0))
    return avg_imp < RAMP_TARGET_IMPRESSIONS_PER_DAY and avg_spend <= RAMP_MAX_SPEND_PER_DAY


def classify_ramp_status(row: pd.Series | Dict[str, Any]) -> Tuple[str, str, bool]:
    """Возвращает статус режима Разгон для отчёта: статус, группа причины, применён ли режим."""
    reason_code = _clean_text_value(row.get("reason_code", ""))
    wait_rule = _clean_text_value(row.get("wait_rule", ""))
    wait_status = _clean_text_value(row.get("wait_status", ""))
    last_reason = _clean_text_value(row.get("last_bid_change_reason_code", ""))
    action = _clean_text_value(row.get("action", ""))
    api_status = _clean_text_value(row.get("api_status", row.get("ramp_api_status", "")))
    candidate = is_ramp_candidate_by_metrics(row)
    related = is_ramp_related_reason(reason_code, wait_rule, last_reason)
    reason_upper = reason_code.upper()
    wait_upper = wait_rule.upper()
    api_ok = api_status.isdigit() and 200 <= int(api_status) < 300

    if api_ok and action == "Повысить" and related:
        return "РАЗГОН_ПРИМЕНЕН_СЕЙЧАС_API_200", "APPLIED_NOW", True
    if ("WAIT_D7_RAMP" in reason_upper or "WAIT_D7_RAMP" in wait_upper or "RAMP" in last_reason.upper()) and wait_status == "WAIT_ACTIVE":
        return "РАЗГОН_АКТИВЕН_ЖДЕМ_D7", "ACTIVE_WAIT_D7", True
    if related and "WB_MIN_BID_NOT_ALLOWED" in reason_upper:
        return "РАЗГОН_ПОДХОДИТ_НО_БЛОК_MIN_WB", "BLOCKED_BY_MIN_BID", False
    if action == "Повысить" and related:
        return "РАЗГОН_К_ОТПРАВКЕ", "TO_SEND", False
    if candidate and wait_status == "WAIT_ACTIVE" and not related:
        return "РАЗГОН_ПОДХОДИТ_НО_ЖДЕМ_ДРУГОЙ_POSTCHECK", "WAIT_OTHER_POSTCHECK", False
    if candidate and reason_code in {"NOT_ACTIVE", "MISSING_KEY", "NOT_MANAGED_SUBJECT"}:
        return "РАЗГОН_НЕ_МОЖЕТ_БЫТЬ_ПРИМЕНЕН", "TECHNICAL_BLOCK", False
    if candidate:
        return "РАЗГОН_ПОДХОДИТ_ПО_МЕТРИКАМ_НО_НЕ_ВКЛЮЧЕН", "CANDIDATE_BY_METRICS", False
    if related:
        return "РАЗГОН_КОНТРОЛЬ", "RELATED_CONTROL", False
    return "", "", False

def build_decisions(metrics_df: pd.DataFrame, pending_events: Dict[Tuple[str, str, str], Dict[str, Any]], postcheck_results: Dict[Tuple[str, str, str], Dict[str, Any]], ctx: Optional[RunContext] = None) -> pd.DataFrame:
    if metrics_df.empty:
        return pd.DataFrame(columns=DECISION_COLUMNS)
    rows: List[Dict[str, Any]] = []
    for _, row in metrics_df.iterrows():
        key = make_key(row)
        pending = pending_events.get(key)
        latest_result = postcheck_results.get(key)
        reference_event = pending or latest_result or {}
        decision = decide_action(row, pending_event=pending, postcheck_result=latest_result)
        previous_event_id = _clean_text_value((latest_result or {}).get("event_id", ""))
        postcheck_status = _clean_text_value((latest_result or {}).get("postcheck_status", ""))
        wait_info = pending_wait_info(reference_event, ctx) if reference_event else {
            "wait_rule": "", "wait_until_date": "", "wait_days_left": "", "wait_status": "NO_PREVIOUS_EVENT"
        }
        if pending is None and reference_event:
            wait_info["wait_status"] = "WAIT_EXPIRED_OR_RESOLVED"
        out = {
            "campaign_id": row.get("campaign_id", ""),
            "nm_id": row.get("nm_id", ""),
            "supplier_article": row.get("supplier_article", ""),
            "subject_norm": row.get("subject_norm", ""),
            "placement": row.get("placement", ""),
            "campaign_status": row.get("campaign_status", ""),
            "current_bid_rub": row.get("current_bid_rub", 0),
            "min_bid_rub": row.get("min_bid_rub", float("nan")),
            "drr_limit_pct": row.get("drr_limit_pct", drr_limit_for_subject(row.get("subject_norm", ""))),
            "avg_impressions_per_day": row.get("avg_impressions_per_day", 0),
            "avg_spend_per_day": row.get("avg_spend_per_day", 0),
            "last21_impressions": row.get("last21_impressions", 0),
            "last21_spend": row.get("last21_spend", 0),
            "last21_revenue": row.get("last21_revenue", 0),
            "last21_orders": row.get("last21_orders", 0),
            "last21_drr_pct": row.get("last21_drr_pct", 0),
            "last21_avg_impressions_per_day": row.get("last21_avg_impressions_per_day", 0),
            "last21_avg_spend_per_day": row.get("last21_avg_spend_per_day", 0),
            "new_bid_rub": decision.get("new_bid_rub"),
            "action": decision.get("action", "Без изменений"),
            "reason_code": decision.get("reason_code", ""),
            "reason_text": decision.get("reason_text", ""),
            "spend": row.get("spend", 0),
            "revenue": row.get("revenue", 0),
            "orders": row.get("orders", 0),
            "impressions": row.get("impressions", 0),
            "clicks": row.get("clicks", 0),
            "campaign_drr_pct": row.get("campaign_drr_pct", 0),
            "cpo": row.get("cpo", 0),
            "ctr_pct": row.get("ctr_pct", 0),
            "gp_after_ads": row.get("gp_after_ads", float("nan")),
            "previous_event_id": previous_event_id,
            "postcheck_status": postcheck_status,
            "last_bid_change_event_id": _clean_text_value(reference_event.get("event_id", "")) if reference_event else "",
            "last_bid_change_date": _clean_text_value(reference_event.get("event_date", "")) if reference_event else "",
            "last_bid_change_old_bid": reference_event.get("old_bid_rub", "") if reference_event else "",
            "last_bid_change_new_bid": reference_event.get("new_bid_rub", "") if reference_event else "",
            "last_bid_change_direction": _clean_text_value(reference_event.get("direction", "")) if reference_event else "",
            "last_bid_change_reason_code": _clean_text_value(reference_event.get("reason_code", "")) if reference_event else "",
            "wait_rule": wait_info.get("wait_rule", ""),
            "wait_until_date": wait_info.get("wait_until_date", ""),
            "wait_days_left": wait_info.get("wait_days_left", ""),
            "wait_status": wait_info.get("wait_status", ""),
            "pause_decision": decision.get("pause_decision", ""),
        }
        ramp_status, ramp_reason_group, ramp_applied = classify_ramp_status(out)
        out["ramp_candidate"] = bool(is_ramp_candidate_by_metrics(out))
        out["ramp_status"] = ramp_status
        out["ramp_applied_in_current_run"] = bool(ramp_applied and _clean_text_value(out.get("ramp_api_status", "")))
        out["ramp_api_status"] = ""
        rows.append(out)
    result = pd.DataFrame(rows)
    for col in DECISION_COLUMNS:
        if col not in result.columns:
            result[col] = ""
    return result[DECISION_COLUMNS]


# =============================
# API ставок и запись истории
# =============================

def wb_headers(config: Config) -> Dict[str, str]:
    return {
        "Authorization": config.wb_promo_key,
        "Content-Type": "application/json",
    }


def to_int_id(value: Any) -> Optional[int]:
    text = _clean_id_value(value)
    if re.fullmatch(r"\d+", text):
        return int(text)
    return None


def build_bid_payload(row: pd.Series) -> Optional[Dict[str, Any]]:
    advert_id = to_int_id(row.get("campaign_id", ""))
    nm_id = to_int_id(row.get("nm_id", ""))
    placement = normalize_placement_value(row.get("placement", ""))
    new_bid = row.get("new_bid_rub")
    if advert_id is None or nm_id is None or not placement or pd.isna(new_bid):
        return None
    bid_kopecks = int(round(float(new_bid) * 100))
    return {
        "bids": [
            {
                "advert_id": advert_id,
                "nm_bids": [
                    {
                        "nm_id": nm_id,
                        "bid_kopecks": bid_kopecks,
                        "placement": placement,
                    }
                ],
            }
        ]
    }


def api_log_row(run_datetime: datetime, method: str, endpoint: str, payload: Any, status: str, response_text: str, campaign_id: Any = "", nm_id: Any = "", placement: Any = "") -> Dict[str, Any]:
    return {
        "run_datetime": run_datetime.strftime("%Y-%m-%d %H:%M:%S"),
        "method": method,
        "endpoint": endpoint,
        "campaign_id": campaign_id,
        "nm_id": nm_id,
        "placement": placement,
        "payload": json.dumps(payload, ensure_ascii=False) if payload not in (None, "") else "",
        "api_status": status,
        "response_text": str(response_text)[:1000],
    }



def placement_for_min_endpoint(value: Any) -> str:
    placement = normalize_placement_value(value)
    if placement == "recommendations":
        return "recommendation"
    if placement in {"search", "combined", "recommendation"}:
        return placement
    return "search"


def infer_payment_type_for_min(row: pd.Series) -> str:
    placement = normalize_placement_value(row.get("placement", ""))
    # В текущем отчёте нет надёжной отдельной колонки payment_type, поэтому для WB min endpoint:
    # combined считаем CPM, search/recommendations считаем CPC.
    return "cpm" if placement == "combined" else "cpc"


def fetch_wb_min_bids_for_decisions(decisions: pd.DataFrame, config: Config, ctx: RunContext) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Получает минимальные ставки WB для активных управляемых строк перед отправкой изменений."""
    empty_min = pd.DataFrame(columns=MIN_BID_COLUMNS)
    empty_log = pd.DataFrame(columns=["run_datetime", "method", "endpoint", "campaign_id", "nm_id", "placement", "payload", "api_status", "response_text"])
    if decisions is None or decisions.empty:
        return empty_min, empty_log

    work = decisions.copy()
    # Минимальные ставки нужны не только перед отправкой, но и для объяснения WAIT/TECHNICAL_FLOOR:
    # если текущая ставка уже равна минимальной WB, код не должен писать "ждём" вместо "снижать нельзя".
    work = work[
        work["campaign_id"].map(_clean_id_value).ne("")
        & work["nm_id"].map(_clean_id_value).ne("")
        & work["placement"].map(normalize_placement_value).ne("")
        & work["subject_norm"].map(is_managed_subject)
        & work["campaign_status"].map(is_active_campaign)
    ].copy()
    if work.empty:
        return empty_min, empty_log

    work["payment_type_for_min"] = work.apply(infer_payment_type_for_min, axis=1)
    work["placement_for_min"] = work["placement"].map(placement_for_min_endpoint)
    work["campaign_id_int"] = work["campaign_id"].map(to_int_id)
    work["nm_id_int"] = work["nm_id"].map(to_int_id)
    work = work.dropna(subset=["campaign_id_int", "nm_id_int"]).copy()
    if work.empty:
        return empty_min, empty_log

    url = config.wb_base_url.rstrip("/") + WB_BIDS_MIN_ENDPOINT
    min_rows: List[Dict[str, Any]] = []
    api_logs: List[Dict[str, Any]] = []

    if not config.wb_promo_key:
        api_logs.append(api_log_row(ctx.run_datetime, "POST", WB_BIDS_MIN_ENDPOINT, {}, "skipped", "Нет WB_PROMO_KEY_TOPFACE"))
        return pd.DataFrame(min_rows, columns=MIN_BID_COLUMNS), pd.DataFrame(api_logs)

    for (campaign_id, payment_type), grp in work.groupby(["campaign_id_int", "payment_type_for_min"], dropna=True):
        nm_ids_all = sorted({int(x) for x in grp["nm_id_int"].tolist() if pd.notna(x) and int(x) > 0})
        placement_types = sorted({placement_for_min_endpoint(x) for x in grp["placement_for_min"].tolist() if _clean_text_value(x)})
        if not nm_ids_all:
            continue
        for offset in range(0, len(nm_ids_all), 100):
            nm_chunk = nm_ids_all[offset:offset + 100]
            payload = {
                "advert_id": int(campaign_id),
                "nm_ids": nm_chunk,
                "payment_type": str(payment_type),
                "placement_types": placement_types or ["search"],
            }
            try:
                resp = requests.post(url, headers=wb_headers(config), json=payload, timeout=60)
                api_logs.append(api_log_row(ctx.run_datetime, "POST", WB_BIDS_MIN_ENDPOINT, payload, str(resp.status_code), resp.text, campaign_id=campaign_id))
                if 200 <= resp.status_code < 300:
                    try:
                        data = resp.json()
                    except Exception:
                        data = {}
                    for item in data.get("bids", []) or []:
                        nm_id = to_int_id(item.get("nm_id", ""))
                        if nm_id is None:
                            continue
                        for bid_item in item.get("bids", []) or []:
                            placement_type = placement_for_min_endpoint(bid_item.get("type", ""))
                            value_kopecks = pd.to_numeric(bid_item.get("value", None), errors="coerce")
                            if pd.isna(value_kopecks) or float(value_kopecks) <= 0:
                                continue
                            min_rows.append({
                                "run_datetime": ctx.run_datetime.strftime("%Y-%m-%d %H:%M:%S"),
                                "campaign_id": str(int(campaign_id)),
                                "nm_id": str(int(nm_id)),
                                "placement": "recommendations" if placement_type == "recommendation" else placement_type,
                                "payment_type": str(payment_type),
                                "min_bid_rub": round(float(value_kopecks) / 100.0, 2),
                                "api_status": str(resp.status_code),
                                "response_text": "",
                            })
                # endpoint имеет ограничение по частоте; выдерживаем паузу как в старом рабочем коде.
                time.sleep(3.1)
            except Exception as exc:
                api_logs.append(api_log_row(ctx.run_datetime, "POST", WB_BIDS_MIN_ENDPOINT, payload, "exception", repr(exc), campaign_id=campaign_id))

    min_df = pd.DataFrame(min_rows, columns=MIN_BID_COLUMNS).drop_duplicates() if min_rows else empty_min
    return min_df, pd.DataFrame(api_logs)


def enrich_decisions_with_min_bids(decisions: pd.DataFrame, min_bids_df: pd.DataFrame) -> pd.DataFrame:
    if decisions is None or decisions.empty:
        return pd.DataFrame(columns=DECISION_COLUMNS)
    result = decisions.copy()
    if "min_bid_rub" not in result.columns:
        result["min_bid_rub"] = float("nan")
    if min_bids_df is None or min_bids_df.empty:
        return result[DECISION_COLUMNS]

    lookup: Dict[Tuple[str, str, str], float] = {}
    for _, r in min_bids_df.iterrows():
        key = (
            _clean_id_value(r.get("campaign_id", "")),
            _clean_id_value(r.get("nm_id", "")),
            normalize_placement_value(r.get("placement", "")),
        )
        val = pd.to_numeric(r.get("min_bid_rub", None), errors="coerce")
        if all(key) and not pd.isna(val) and float(val) > 0:
            lookup[key] = float(val)

    for idx, row in result.iterrows():
        key = make_key(row)
        min_bid = lookup.get(key)
        if min_bid is None:
            continue
        # Используем эффективный минимум: максимум из WB API min и нашего бизнес-минимума.
        # Для поиска/рекомендаций бизнес-минимум 4 ₽, для combined/полок 80 ₽.
        min_bid = max(float(min_bid), business_min_bid_rub(row.get("placement", "")))
        result.at[idx, "min_bid_rub"] = round(min_bid, 2)

        current_bid = pd.to_numeric(row.get("current_bid_rub", None), errors="coerce")
        new_bid = pd.to_numeric(row.get("new_bid_rub", None), errors="coerce")
        if pd.isna(current_bid):
            continue
        current_bid_f = float(current_bid)

        drr_current = pd.to_numeric(row.get("campaign_drr_pct", 0), errors="coerce")
        if pd.isna(drr_current):
            drr_current = 0.0
        drr_limit = pd.to_numeric(row.get("drr_limit_pct", drr_limit_for_subject(row.get("subject_norm", ""))), errors="coerce")
        if pd.isna(drr_limit) or float(drr_limit) <= 0:
            drr_limit = drr_limit_for_subject(row.get("subject_norm", ""))
        drr_limit_f = float(drr_limit)

        last21_impressions = money_or_zero(row.get("last21_impressions", row.get("impressions", 0)))
        last21_drr = pd.to_numeric(row.get("last21_drr_pct", drr_current), errors="coerce")
        if pd.isna(last21_drr):
            last21_drr = float(drr_current)
        last21_drr_f = float(last21_drr)
        avg_impressions_per_day = money_or_zero(row.get("avg_impressions_per_day", 0))
        avg_spend_per_day = money_or_zero(row.get("avg_spend_per_day", 0))
        at_wb_min_bid = current_bid_f <= float(min_bid) + 0.001
        pause_subject = is_pause_allowed_subject(row.get("subject_norm", ""))

        # Главное правило: если ставка уже минимальная WB, а экономика стабильно плохая
        # по 21 дню и статистики достаточно — не ждём, а ставим РК на паузу.
        if (
            pause_subject
            and at_wb_min_bid
            and last21_impressions >= PAUSE_MIN_IMPRESSIONS
            and last21_drr_f > drr_limit_f
        ):
            result.at[idx, "action"] = "Без изменений"
            result.at[idx, "new_bid_rub"] = None
            result.at[idx, "reason_code"] = "PAUSE_MIN_BID_HIGH_DRR_21D_10000"
            result.at[idx, "reason_text"] = build_reason_text(
                result.loc[idx],
                "Без изменений",
                None,
                (
                    f"пауза: ставка {current_bid_f:.2f} ₽ уже на минимуме WB {float(min_bid):.2f} ₽; "
                    f"ДРР за {PAUSE_ANALYSIS_DAYS} дней {last21_drr_f:.2f}% > лимита {drr_limit_f:.1f}%; "
                    f"показов за {PAUSE_ANALYSIS_DAYS} дней {last21_impressions:.0f} >= {PAUSE_MIN_IMPRESSIONS}; ждать нечего"
                ),
            )
            result.at[idx, "wait_status"] = "NO_WAIT_PAUSE_MIN_BID_21D"
            result.at[idx, "wait_rule"] = "PAUSE_MIN_BID_HIGH_DRR_21D_10000"
            result.at[idx, "wait_until_date"] = ""
            result.at[idx, "wait_days_left"] = 0
            result.at[idx, "pause_decision"] = "PAUSE_CANDIDATE"
            continue

        # Если ставка минимальная, ДРР высокий, но показов за 21 день меньше 10 000 —
        # это не пауза. Включаем общий разгон, если расход в лимите 500 ₽/день.
        if (
            pause_subject
            and at_wb_min_bid
            and last21_impressions < PAUSE_MIN_IMPRESSIONS
            and last21_drr_f > drr_limit_f
            and avg_spend_per_day <= RAMP_MAX_SPEND_PER_DAY
            and avg_impressions_per_day < RAMP_TARGET_IMPRESSIONS_PER_DAY
        ):
            step, step_reason = bid_step_rub(row.get("placement", ""))
            ramp_bid = round(current_bid_f + step, 2)
            rc = "RAMP_MIN_BID_HIGH_DRR_UNDER_10000_IMPRESSIONS_21D"
            if step_reason:
                rc += f"__{step_reason}"
            result.at[idx, "action"] = "Повысить"
            result.at[idx, "new_bid_rub"] = ramp_bid
            result.at[idx, "reason_code"] = rc
            result.at[idx, "reason_text"] = build_reason_text(
                result.loc[idx],
                "Повысить",
                ramp_bid,
                (
                    f"разгон вместо паузы: ставка на минимуме WB {float(min_bid):.2f} ₽, "
                    f"ДРР за {PAUSE_ANALYSIS_DAYS} дней {last21_drr_f:.2f}% > лимита {drr_limit_f:.1f}%, "
                    f"но показов {last21_impressions:.0f} < {PAUSE_MIN_IMPRESSIONS}; "
                    f"расход {avg_spend_per_day:.0f} ₽/день <= {RAMP_MAX_SPEND_PER_DAY:.0f} ₽/день; проверка D+{RAMP_CHECK_DAYS}"
                ),
            )
            result.at[idx, "wait_status"] = "NO_WAIT_RAMP_UNDER_10000_21D"
            result.at[idx, "pause_decision"] = ""
            continue

        # Если ДРР плохой, но ставка уже на минималке WB, это не ожидание post-check.
        # Если статистики для паузы не хватает и разгон не разрешён лимитом, показываем честную причину.
        if _clean_text_value(row.get("reason_code", "")).startswith("WAIT_") and float(drr_current) >= drr_limit_f and at_wb_min_bid:
            result.at[idx, "action"] = "Без изменений"
            result.at[idx, "new_bid_rub"] = None
            result.at[idx, "reason_code"] = "WB_MIN_BID_REACHED"
            result.at[idx, "reason_text"] = build_reason_text(result.loc[idx], "Без изменений", None, f"ДРР {float(drr_current):.2f}% >= лимита {drr_limit_f:.1f}%, но текущая ставка {current_bid_f:.2f} ₽ не выше минимальной WB {float(min_bid):.2f} ₽; показов за {PAUSE_ANALYSIS_DAYS} дней {last21_impressions:.0f}; правило паузы/разгона не выполнено")
            result.at[idx, "wait_status"] = "NO_WAIT_MIN_BID_REACHED"
            result.at[idx, "pause_decision"] = ""
            continue

        if row.get("action") != "Снизить":
            # Для повышения/разгона нельзя отправлять ставку ниже минимума WB: WB отклонит запрос.
            # Но если это именно Разгон, корректно поднимаем ставку сразу до effective min_bid,
            # иначе combined/полки с текущей ставкой 3-5 ₽ никогда не выйдут в рабочую минимальную ставку 80 ₽.
            if not pd.isna(new_bid) and float(new_bid) < float(min_bid) and row.get("action") == "Повысить":
                rc = _clean_text_value(result.at[idx, "reason_code"])
                if is_ramp_related_reason(rc) or is_ramp_candidate_by_metrics(row):
                    adjusted_bid = round(float(min_bid), 2)
                    result.at[idx, "new_bid_rub"] = adjusted_bid
                    result.at[idx, "reason_code"] = (rc + "__TO_EFFECTIVE_MIN_BID").strip("_")
                    result.at[idx, "reason_text"] = build_reason_text(
                        result.loc[idx],
                        "Повысить",
                        adjusted_bid,
                        f"разгон: расчётная ставка {float(new_bid):.2f} ₽ ниже effective min_bid {float(min_bid):.2f} ₽; ставим сразу минимально допустимую ставку"
                    )
                    result.at[idx, "pause_decision"] = ""
                else:
                    result.at[idx, "action"] = "Без изменений"
                    result.at[idx, "new_bid_rub"] = None
                    result.at[idx, "reason_code"] = (rc + "__WB_MIN_BID_NOT_ALLOWED").strip("_")
                    result.at[idx, "reason_text"] = build_reason_text(result.loc[idx], "Без изменений", None, f"расчётная ставка {float(new_bid):.2f} ₽ ниже минимально допустимой WB/effective min {float(min_bid):.2f} ₽; не отправляем заведомо невалидную ставку")
                    result.at[idx, "pause_decision"] = ""
            continue
        if pd.isna(new_bid):
            continue
        new_bid_f = float(new_bid)

        if at_wb_min_bid:
            result.at[idx, "action"] = "Без изменений"
            result.at[idx, "new_bid_rub"] = None
            result.at[idx, "reason_code"] = "WB_MIN_BID_REACHED"
            result.at[idx, "reason_text"] = build_reason_text(result.loc[idx], "Без изменений", None, f"текущая ставка не выше минимально допустимой WB {float(min_bid):.2f} ₽; снижение не отправляем; показов за {PAUSE_ANALYSIS_DAYS} дней {last21_impressions:.0f}")
            result.at[idx, "pause_decision"] = ""
            continue

        if new_bid_f < float(min_bid):
            result.at[idx, "action"] = "Без изменений"
            result.at[idx, "new_bid_rub"] = None
            rc = _clean_text_value(result.at[idx, "reason_code"])
            result.at[idx, "reason_code"] = (rc + "__WB_MIN_BID_NOT_ALLOWED").strip("_")
            result.at[idx, "reason_text"] = build_reason_text(result.loc[idx], "Без изменений", None, f"расчётная ставка {new_bid_f:.2f} ₽ ниже минимально допустимой WB {float(min_bid):.2f} ₽; не отправляем заведомо невалидную ставку")
            result.at[idx, "pause_decision"] = ""

    for col in DECISION_COLUMNS:
        if col not in result.columns:
            result[col] = ""
    return result[DECISION_COLUMNS]

def apply_bid_changes(decisions: pd.DataFrame, config: Config, ctx: RunContext) -> Tuple[pd.DataFrame, pd.DataFrame]:
    candidates = decisions[decisions["action"].isin(["Повысить", "Снизить"])].copy() if not decisions.empty else pd.DataFrame(columns=decisions.columns)
    api_logs: List[Dict[str, Any]] = []
    changed_rows: List[Dict[str, Any]] = []

    if candidates.empty:
        return pd.DataFrame(columns=decisions.columns.tolist() + ["api_status"]), pd.DataFrame(api_logs)

    url = config.wb_base_url.rstrip("/") + WB_BIDS_ENDPOINT
    for _, row in candidates.iterrows():
        payload = build_bid_payload(row)
        if payload is None:
            api_logs.append(api_log_row(ctx.run_datetime, "PATCH", WB_BIDS_ENDPOINT, {}, "payload_error", "Не удалось собрать payload", row.get("campaign_id"), row.get("nm_id"), row.get("placement")))
            continue

        if ctx.mode == "preview":
            api_logs.append(api_log_row(ctx.run_datetime, "PATCH", WB_BIDS_ENDPOINT, payload, "preview_no_call", "Предпросмотр без API-вызова", row.get("campaign_id"), row.get("nm_id"), row.get("placement")))
            continue

        if ctx.dry_run:
            api_logs.append(api_log_row(ctx.run_datetime, "PATCH", WB_BIDS_ENDPOINT, payload, "dry_run_no_call", "run --dry-run без API-вызова", row.get("campaign_id"), row.get("nm_id"), row.get("placement")))
            continue

        try:
            resp = requests.patch(url, headers=wb_headers(config), json=payload, timeout=60)
            status = str(resp.status_code)
            api_logs.append(api_log_row(ctx.run_datetime, "PATCH", WB_BIDS_ENDPOINT, payload, status, resp.text, row.get("campaign_id"), row.get("nm_id"), row.get("placement")))
            if 200 <= resp.status_code < 300:
                changed = row.to_dict()
                changed["api_status"] = status
                changed_rows.append(changed)
        except Exception as exc:
            api_logs.append(api_log_row(ctx.run_datetime, "PATCH", WB_BIDS_ENDPOINT, payload, "exception", repr(exc), row.get("campaign_id"), row.get("nm_id"), row.get("placement")))

    changed_df = pd.DataFrame(changed_rows)
    api_log_df = pd.DataFrame(api_logs)
    return changed_df, api_log_df


def record_bid_events(successful_changes: pd.DataFrame, bid_history: pd.DataFrame, ctx: RunContext) -> pd.DataFrame:
    if successful_changes.empty:
        return bid_history[BID_HISTORY_COLUMNS].copy() if not bid_history.empty else pd.DataFrame(columns=BID_HISTORY_COLUMNS)

    rows: List[Dict[str, Any]] = []
    for _, row in successful_changes.iterrows():
        action = _clean_text_value(row.get("action", ""))
        direction = "raise" if action == "Повысить" else "lower"
        rows.append({
            "event_id": str(uuid.uuid4()),
            "run_datetime": ctx.run_datetime.strftime("%Y-%m-%d %H:%M:%S"),
            "event_date": ctx.run_datetime.date().isoformat(),
            "campaign_id": row.get("campaign_id", ""),
            "nm_id": row.get("nm_id", ""),
            "supplier_article": row.get("supplier_article", ""),
            "subject_norm": row.get("subject_norm", ""),
            "placement": row.get("placement", ""),
            "old_bid_rub": row.get("current_bid_rub", 0),
            "new_bid_rub": row.get("new_bid_rub", 0),
            "direction": direction,
            "reason_code": row.get("reason_code", ""),
            "spend_before": row.get("spend", 0),
            "revenue_before": row.get("revenue", 0),
            "orders_before": row.get("orders", 0),
            "impressions_before": row.get("impressions", 0),
            "clicks_before": row.get("clicks", 0),
            "drr_before": row.get("campaign_drr_pct", 0),
            "gp_before": row.get("gp_after_ads", float("nan")),
            "postcheck_status": "pending",
            "final_verdict": "",
            "d1_verdict": "",
            "d3_verdict": "",
            "d1_check_date": "",
            "d3_check_date": "",
        })
    additions = pd.DataFrame(rows)
    base = bid_history.copy()
    for col in BID_HISTORY_COLUMNS:
        if col not in base.columns:
            base[col] = ""
        if col not in additions.columns:
            additions[col] = ""
    return pd.concat([base[BID_HISTORY_COLUMNS], additions[BID_HISTORY_COLUMNS]], ignore_index=True)



# =============================
# Переименование рекламных кампаний
# =============================

def normalize_article_for_campaign_name(value: Any) -> str:
    """
    Приводит артикул продавца к короткому имени кампании.
    Примеры: PT155.009K -> 155/9; PT156.001 -> 156/1; 155/001 -> 155/1.
    Если артикул уже в нормальном формате или содержит буквенную часть, возвращаем аккуратно очищенный текст.
    """
    text = _clean_text_value(value).replace(" ", "").strip()
    if not text:
        return ""
    upper = text.upper()

    # PT155.009K / PT155.009 -> 155/9
    m = re.fullmatch(r"PT(\d{2,5})[\._\-/](\d{1,4})([A-ZА-Я]*)", upper)
    if m:
        code = str(int(m.group(1))) if m.group(1).isdigit() else m.group(1)
        shade_raw = m.group(2)
        shade = str(int(shade_raw)) if shade_raw.isdigit() else shade_raw
        suffix = m.group(3) or ""
        return f"{code}/{shade}{suffix}"

    # 155.009K / 155_009 / 155-009 -> 155/9K
    m = re.fullmatch(r"(\d{2,5})[\._\-/](\d{1,4})([A-ZА-Я]*)", upper)
    if m:
        code = str(int(m.group(1))) if m.group(1).isdigit() else m.group(1)
        shade_raw = m.group(2)
        shade = str(int(shade_raw)) if shade_raw.isdigit() else shade_raw
        suffix = m.group(3) or ""
        return f"{code}/{shade}{suffix}"

    # PT901.F26 -> 901/F26; если уже F26 — оставляем F26.
    m = re.fullmatch(r"PT(\d{2,5})[\._\-/]([A-ZА-Я]+\d+)", upper)
    if m:
        code = str(int(m.group(1))) if m.group(1).isdigit() else m.group(1)
        return f"{code}/{m.group(2)}"

    # Если уже есть слэш, убираем лидирующие нули у числового оттенка.
    m = re.fullmatch(r"(\d{2,5})/(\d{1,4})([A-ZА-Я]*)", upper)
    if m:
        code = str(int(m.group(1))) if m.group(1).isdigit() else m.group(1)
        shade = str(int(m.group(2))) if m.group(2).isdigit() else m.group(2)
        return f"{code}/{shade}{m.group(3) or ''}"

    return text[:100]


def _best_article_by_nm(metrics_df: pd.DataFrame, keyword_core_df: Optional[pd.DataFrame], goods_prices: Optional[pd.DataFrame]) -> Dict[str, str]:
    """Собирает nm_id -> артикул из отчёта рекламы, поисковых запросов и WB price API."""
    article_by_nm: Dict[str, str] = {}

    def add_mapping(df: Optional[pd.DataFrame], nm_col: str, art_col: str) -> None:
        if df is None or df.empty or nm_col not in df.columns or art_col not in df.columns:
            return
        local = df[[nm_col, art_col]].copy()
        local[nm_col] = local[nm_col].map(_clean_id_value)
        local[art_col] = local[art_col].map(_clean_text_value)
        local = local[(local[nm_col] != "") & (local[art_col] != "")]
        # Берём самое частое непустое значение по nm_id.
        for nm_id, grp in local.groupby(nm_col, dropna=False):
            if nm_id in article_by_nm and article_by_nm[nm_id]:
                continue
            vals = grp[art_col].astype(str).str.strip()
            if vals.empty:
                continue
            article_by_nm[nm_id] = normalize_article_for_campaign_name(vals.value_counts().index[0])

    add_mapping(metrics_df, "nm_id", "supplier_article")
    add_mapping(keyword_core_df, "nm_id", "supplier_article")
    add_mapping(goods_prices, "nm_id", "supplier_article_api")
    return article_by_nm


def build_campaign_rename_plan(
    metrics_df: pd.DataFrame,
    keyword_core_df: Optional[pd.DataFrame],
    goods_prices: Optional[pd.DataFrame],
    ctx: RunContext,
) -> pd.DataFrame:
    """
    Строит план переименования РК обратно в короткий артикул продавца.
    Безопасное правило: переименовываем только кампании, где однозначно найден один артикул.
    Если в campaign_id несколько разных артикулов — не трогаем, чтобы не назвать сборную кампанию неверно.
    """
    if metrics_df is None or metrics_df.empty:
        return pd.DataFrame(columns=RENAME_CAMPAIGN_COLUMNS)
    article_by_nm = _best_article_by_nm(metrics_df, keyword_core_df, goods_prices)
    work = metrics_df.copy()
    for col in ["campaign_id", "nm_id", "campaign_name", "supplier_article", "subject_norm"]:
        if col not in work.columns:
            work[col] = ""
    work["campaign_id"] = work["campaign_id"].map(_clean_id_value)
    work["nm_id"] = work["nm_id"].map(_clean_id_value)
    work["current_name_clean"] = work["campaign_name"].map(_clean_text_value)
    work["article_for_name"] = work.apply(
        lambda r: normalize_article_for_campaign_name(r.get("supplier_article", "")) or article_by_nm.get(_clean_id_value(r.get("nm_id", "")), ""),
        axis=1,
    )
    work = work[(work["campaign_id"] != "") & (work["article_for_name"] != "")].copy()
    if work.empty:
        return pd.DataFrame(columns=RENAME_CAMPAIGN_COLUMNS)

    rows: List[Dict[str, Any]] = []
    for campaign_id, grp in work.groupby("campaign_id", dropna=False):
        current_names = [_clean_text_value(x) for x in grp.get("current_name_clean", pd.Series(dtype=str)).tolist() if _clean_text_value(x)]
        current_name = current_names[0] if current_names else ""
        articles = sorted({normalize_article_for_campaign_name(x) for x in grp.get("article_for_name", pd.Series(dtype=str)).tolist() if normalize_article_for_campaign_name(x)})
        nm_ids = sorted({_clean_id_value(x) for x in grp.get("nm_id", pd.Series(dtype=str)).tolist() if _clean_id_value(x)})
        subjects = sorted({normalize_subject_value(x) for x in grp.get("subject_norm", pd.Series(dtype=str)).tolist() if normalize_subject_value(x)})

        if len(articles) != 1:
            rows.append({
                "run_datetime": ctx.run_datetime.strftime("%Y-%m-%d %H:%M:%S"),
                "campaign_id": campaign_id,
                "current_name": current_name,
                "target_name": "",
                "supplier_article": "; ".join(articles[:10]),
                "nm_ids": "; ".join(nm_ids[:20]),
                "subjects": "; ".join(subjects[:10]),
                "rename_action": "Без изменений",
                "reason_code": "MULTIPLE_ARTICLES_IN_CAMPAIGN" if articles else "NO_SUPPLIER_ARTICLE",
                "api_status": "not_sent",
                "response_text": "Кампания содержит несколько артикулов или артикул не найден; автоматическое переименование небезопасно",
            })
            continue

        target_name = articles[0]
        if current_name == target_name:
            action = "Без изменений"
            reason_code = "ALREADY_NAMED_BY_ARTICLE"
            api_status = "not_sent"
            response_text = "Название уже равно артикулу продавца"
        else:
            action = "Переименовать"
            reason_code = "RENAME_TO_SUPPLIER_ARTICLE"
            api_status = ""
            response_text = ""
        rows.append({
            "run_datetime": ctx.run_datetime.strftime("%Y-%m-%d %H:%M:%S"),
            "campaign_id": campaign_id,
            "current_name": current_name,
            "target_name": target_name,
            "supplier_article": target_name,
            "nm_ids": "; ".join(nm_ids[:20]),
            "subjects": "; ".join(subjects[:10]),
            "rename_action": action,
            "reason_code": reason_code,
            "api_status": api_status,
            "response_text": response_text,
        })
    out = pd.DataFrame(rows)
    for col in RENAME_CAMPAIGN_COLUMNS:
        if col not in out.columns:
            out[col] = ""
    return out[RENAME_CAMPAIGN_COLUMNS]


def apply_campaign_renames(rename_plan: pd.DataFrame, config: Config, ctx: RunContext) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """POST /adv/v0/rename. В run переименовывает РК в артикул продавца; в preview/dry-run только показывает план."""
    if rename_plan is None or rename_plan.empty:
        return pd.DataFrame(columns=RENAME_CAMPAIGN_COLUMNS), pd.DataFrame()
    result = rename_plan.copy()
    api_logs: List[Dict[str, Any]] = []
    candidates = result[result["rename_action"].astype(str) == "Переименовать"].copy()
    if candidates.empty:
        return result, pd.DataFrame(api_logs)

    url = config.wb_base_url.rstrip("/") + WB_RENAME_ENDPOINT
    for idx, row in candidates.iterrows():
        advert_id = to_int_id(row.get("campaign_id", ""))
        target_name = _clean_text_value(row.get("target_name", ""))
        if advert_id is None or not target_name:
            result.loc[idx, "api_status"] = "payload_error"
            result.loc[idx, "response_text"] = "Не удалось собрать payload для rename"
            api_logs.append(api_log_row(ctx.run_datetime, "POST", WB_RENAME_ENDPOINT, {}, "payload_error", "Не удалось собрать payload для rename", campaign_id=row.get("campaign_id", "")))
            continue
        payload = {"advertId": advert_id, "name": target_name}

        if ctx.mode == "preview":
            result.loc[idx, "api_status"] = "preview_no_call"
            result.loc[idx, "response_text"] = "Предпросмотр без API-вызова"
            api_logs.append(api_log_row(ctx.run_datetime, "POST", WB_RENAME_ENDPOINT, payload, "preview_no_call", "Предпросмотр без API-вызова", campaign_id=row.get("campaign_id", "")))
            continue
        if ctx.dry_run:
            result.loc[idx, "api_status"] = "dry_run_no_call"
            result.loc[idx, "response_text"] = "run --dry-run без API-вызова"
            api_logs.append(api_log_row(ctx.run_datetime, "POST", WB_RENAME_ENDPOINT, payload, "dry_run_no_call", "run --dry-run без API-вызова", campaign_id=row.get("campaign_id", "")))
            continue

        try:
            resp = requests.post(url, headers=wb_headers(config), json=payload, timeout=60)
            status = str(resp.status_code)
            result.loc[idx, "api_status"] = status
            result.loc[idx, "response_text"] = str(resp.text)[:1000]
            api_logs.append(api_log_row(ctx.run_datetime, "POST", WB_RENAME_ENDPOINT, payload, status, resp.text, campaign_id=row.get("campaign_id", "")))
            if 200 <= resp.status_code < 300:
                result.loc[idx, "reason_code"] = "RENAMED_TO_SUPPLIER_ARTICLE"
            else:
                result.loc[idx, "reason_code"] = "RENAME_API_ERROR"
        except Exception as exc:
            result.loc[idx, "api_status"] = "exception"
            result.loc[idx, "response_text"] = repr(exc)[:1000]
            result.loc[idx, "reason_code"] = "RENAME_API_EXCEPTION"
            api_logs.append(api_log_row(ctx.run_datetime, "POST", WB_RENAME_ENDPOINT, payload, "exception", repr(exc), campaign_id=row.get("campaign_id", "")))
        time.sleep(0.2)

    for col in RENAME_CAMPAIGN_COLUMNS:
        if col not in result.columns:
            result[col] = ""
    return result[RENAME_CAMPAIGN_COLUMNS], pd.DataFrame(api_logs)

# =============================
# Паузы и запуск обратно
# =============================

def consecutive_lowers_for_key(bid_history: pd.DataFrame, key: Tuple[str, str, str]) -> int:
    if bid_history.empty:
        return 0
    local = bid_history.copy()
    local["event_date_parsed"] = pd.to_datetime(local["event_date"], errors="coerce").dt.date
    local["run_dt_parsed"] = pd.to_datetime(local["run_datetime"], errors="coerce")
    local = local[local.apply(lambda r: make_key(r) == key, axis=1)].sort_values(["event_date_parsed", "run_dt_parsed"], ascending=False)
    count = 0
    for _, row in local.iterrows():
        if _clean_text_value(row.get("direction", "")).lower() == "lower":
            count += 1
        else:
            break
    return count


def latest_lower_event_for_key(bid_history: pd.DataFrame, key: Tuple[str, str, str]) -> Optional[Dict[str, Any]]:
    if bid_history.empty:
        return None
    local = bid_history.copy()
    local = local[local.apply(lambda r: make_key(r) == key and _clean_text_value(r.get("direction", "")).lower() == "lower", axis=1)]
    if local.empty:
        return None
    local["event_date_parsed"] = pd.to_datetime(local["event_date"], errors="coerce").dt.date
    local["run_dt_parsed"] = pd.to_datetime(local["run_datetime"], errors="coerce")
    local = local.sort_values(["event_date_parsed", "run_dt_parsed"], ascending=False)
    return local.iloc[0].to_dict()


def build_pause_candidates(decisions: pd.DataFrame, bid_history: pd.DataFrame) -> pd.DataFrame:
    """Строгая постановка на паузу по согласованному правилу v10.

    Новая пауза разрешена только если одновременно выполнено:
    - предмет: Помады / Блески / Косметические карандаши;
    - кампания активна;
    - текущая ставка уже на минимальной WB;
    - ДРР за последние 21 зрелых дня выше лимита категории (15%);
    - показов за последние 21 зрелых дня >= 10 000.

    Если показов < 10 000, пауза не формируется: такая строка должна идти в разгон показов.
    Кисти не паузятся никогда.
    """
    if decisions is None or decisions.empty:
        return pd.DataFrame(columns=PAUSE_HISTORY_COLUMNS)

    work = decisions.copy()
    for col in DECISION_COLUMNS:
        if col not in work.columns:
            work[col] = ""
    work["subject_norm"] = work["subject_norm"].map(normalize_subject_value)
    work = work[work["subject_norm"].map(is_pause_allowed_subject)].copy()
    work = work[work["campaign_status"].map(is_active_campaign)].copy()
    if work.empty:
        return pd.DataFrame(columns=PAUSE_HISTORY_COLUMNS)

    candidates: List[Dict[str, Any]] = []
    for campaign_id, g in work.groupby("campaign_id", dropna=False):
        campaign_id_clean = _clean_id_value(campaign_id)
        if not campaign_id_clean:
            continue

        direct = g[
            (g["pause_decision"].map(_clean_text_value) == "PAUSE_CANDIDATE")
            & (g["reason_code"].map(_clean_text_value) == "PAUSE_MIN_BID_HIGH_DRR_21D_10000")
        ].copy()
        if direct.empty:
            continue

        last21_impressions = money_or_zero(g.get("last21_impressions", pd.Series(dtype=float)).sum())
        last21_clicks = money_or_zero(g.get("last21_clicks", pd.Series(dtype=float)).sum()) if "last21_clicks" in g.columns else money_or_zero(g.get("clicks", pd.Series(dtype=float)).sum())
        last21_spend = money_or_zero(g.get("last21_spend", pd.Series(dtype=float)).sum())
        last21_revenue = money_or_zero(g.get("last21_revenue", pd.Series(dtype=float)).sum())
        last21_orders = money_or_zero(g.get("last21_orders", pd.Series(dtype=float)).sum())
        campaign_drr_21d = safe_drr_pct(last21_spend, last21_revenue)
        max_limit = max(drr_limit_for_subject(x) for x in g["subject_norm"].dropna().unique()) if not g.empty else 15.0

        # Повторная защита на уровне campaign_id: без 10 000 показов паузы нет.
        if last21_impressions < PAUSE_MIN_IMPRESSIONS:
            continue
        if campaign_drr_21d <= max_limit:
            continue

        # Основная строка для отображения — та, где сработало правило, с максимальными показами/расходом.
        sort_cols = [c for c in ["last21_impressions", "last21_spend", "impressions", "spend"] if c in direct.columns]
        main = direct.sort_values(sort_cols, ascending=False).iloc[0] if sort_cols else direct.iloc[0]
        gp_series = pd.to_numeric(g.get("last21_gp_after_ads", g.get("gp_after_ads", pd.Series(dtype=float))), errors="coerce")
        gp = float(gp_series.sum()) if gp_series.notna().any() else float("nan")

        candidates.append({
            "pause_event_id": str(uuid.uuid4()),
            "pause_date": date.today().isoformat(),
            "campaign_id": campaign_id_clean,
            "nm_id": main.get("nm_id", ""),
            "placement": main.get("placement", ""),
            "supplier_article": main.get("supplier_article", ""),
            "subject_norm": main.get("subject_norm", ""),
            "reason_code": "PAUSE_MIN_BID_HIGH_DRR_21D_10000",
            "impressions_before_pause": last21_impressions,
            "clicks_before_pause": last21_clicks,
            "spend_before_pause": last21_spend,
            "revenue_before_pause": last21_revenue,
            "orders_before_pause": last21_orders,
            "drr_before_pause": campaign_drr_21d,
            "gp_before_pause": gp,
            "status": "candidate",
            "next_check_date": (date.today() + timedelta(days=1)).isoformat(),
            "api_status": "",
        })
    return pd.DataFrame(candidates, columns=PAUSE_HISTORY_COLUMNS)

def apply_pause_actions(pause_candidates: pd.DataFrame, config: Config, ctx: RunContext) -> Tuple[pd.DataFrame, pd.DataFrame]:
    if pause_candidates.empty:
        return pause_candidates.copy(), pd.DataFrame()

    result = pause_candidates.copy()
    api_logs: List[Dict[str, Any]] = []
    if ctx.mode == "preview":
        result["api_status"] = "preview_no_call"
        return result, pd.DataFrame(api_logs)
    if ctx.dry_run:
        result["api_status"] = "dry_run_no_call"
        return result, pd.DataFrame(api_logs)

    # В обычном run без флага автоматически применяем только строгое согласованное правило v10.
    # Остальные возможные кандидаты остаются кандидатами и не уходят в API без явного --apply-pause.
    if ctx.apply_pause:
        to_send = result.copy()
        result["api_status"] = "not_sent"
    else:
        auto_mask = result["reason_code"].map(_clean_text_value).isin(AUTO_APPLY_PAUSE_REASON_CODES)
        if ctx.mode == "run" and auto_mask.any():
            result["api_status"] = "not_applied_without_flag"
            result.loc[auto_mask, "api_status"] = "auto_apply_v10_pending"
            to_send = result.loc[auto_mask].copy()
        else:
            result["api_status"] = "not_applied_without_flag"
            return result, pd.DataFrame(api_logs)

    url_base = config.wb_base_url.rstrip("/") + WB_PAUSE_ENDPOINT
    status_by_campaign: Dict[str, Tuple[str, str]] = {}
    for campaign_id in sorted(to_send["campaign_id"].map(_clean_id_value).unique()):
        advert_id = to_int_id(campaign_id)
        if advert_id is None:
            status_by_campaign[campaign_id] = ("payload_error", "campaign_id не является числом")
            continue
        endpoint = f"{WB_PAUSE_ENDPOINT}?id={advert_id}"
        try:
            resp = requests.get(url_base, params={"id": advert_id}, headers=wb_headers(config), timeout=60)
            status_by_campaign[campaign_id] = (str(resp.status_code), resp.text)
            api_logs.append(api_log_row(ctx.run_datetime, "GET", endpoint, "", str(resp.status_code), resp.text, campaign_id=campaign_id))
        except Exception as exc:
            status_by_campaign[campaign_id] = ("exception", repr(exc))
            api_logs.append(api_log_row(ctx.run_datetime, "GET", endpoint, "", "exception", repr(exc), campaign_id=campaign_id))

    statuses: List[str] = []
    final_statuses: List[str] = []
    send_campaigns = set(to_send["campaign_id"].map(_clean_id_value).unique())
    for _, row in result.iterrows():
        campaign_id_clean = _clean_id_value(row.get("campaign_id", ""))
        if campaign_id_clean not in send_campaigns:
            statuses.append(_clean_text_value(row.get("api_status", "not_applied_without_flag")) or "not_applied_without_flag")
            final_statuses.append(_clean_text_value(row.get("status", "candidate")) or "candidate")
            continue
        api_status, _ = status_by_campaign.get(campaign_id_clean, ("not_sent", ""))
        statuses.append(api_status)
        if api_status.isdigit() and 200 <= int(api_status) < 300:
            final_statuses.append("paused")
        else:
            final_statuses.append("candidate")
    result["api_status"] = statuses
    result["status"] = final_statuses
    return result, pd.DataFrame(api_logs)

def latest_pause_records(pause_history: pd.DataFrame) -> pd.DataFrame:
    if pause_history.empty:
        return pd.DataFrame(columns=PAUSE_HISTORY_COLUMNS)
    local = pause_history.copy()
    local["pause_date_parsed"] = pd.to_datetime(local["pause_date"], errors="coerce")
    local = local.sort_values(["campaign_id", "nm_id", "placement", "pause_date_parsed"])
    return local.groupby(["campaign_id", "nm_id", "placement"], dropna=False).tail(1)


def build_start_candidates(pause_history: pd.DataFrame, ads_df: pd.DataFrame, ctx: RunContext) -> pd.DataFrame:
    latest = latest_pause_records(pause_history)
    if latest.empty:
        return pd.DataFrame(columns=PAUSE_HISTORY_COLUMNS)
    rows: List[Dict[str, Any]] = []
    for _, row in latest.iterrows():
        status = _clean_text_value(row.get("status", "")).lower()
        if status not in {"paused", "keep_paused"}:
            continue
        pause_date = pd.to_datetime(row.get("pause_date", ""), errors="coerce")
        if pd.isna(pause_date):
            continue
        key = make_key(row)
        after = aggregate_after_event(ads_df, key, pause_date.date() + timedelta(days=1), ctx.mature_end)
        drr_after = safe_drr_pct(after["spend"], after["revenue"])
        gp_after = after["gp_after_ads"]
        if (after["revenue"] > 0 and after["orders"] > 0 and drr_after < DRR_LIMIT_PCT) or (not pd.isna(gp_after) and gp_after > 0):
            rows.append({
                "pause_event_id": str(uuid.uuid4()),
                "pause_date": date.today().isoformat(),
                "campaign_id": row.get("campaign_id", ""),
                "nm_id": row.get("nm_id", ""),
                "placement": row.get("placement", ""),
                "supplier_article": row.get("supplier_article", ""),
                "reason_code": "START_AFTER_ECONOMY_RECOVERY",
                "spend_before_pause": after["spend"],
                "revenue_before_pause": after["revenue"],
                "orders_before_pause": after["orders"],
                "drr_before_pause": drr_after,
                "gp_before_pause": gp_after,
                "status": "restart_candidate",
                "next_check_date": (date.today() + timedelta(days=1)).isoformat(),
                "api_status": "",
            })
    return pd.DataFrame(rows, columns=PAUSE_HISTORY_COLUMNS)



def build_wrong_subject_pause_rollback_candidates(pause_history_raw: pd.DataFrame, managed_ads_df: pd.DataFrame, ctx: RunContext) -> pd.DataFrame:
    """Возвращает кандидатов на запуск кампаний, ошибочно поставленных на паузу вне 4 управляемых предметов."""
    if pause_history_raw is None or pause_history_raw.empty:
        return pd.DataFrame(columns=PAUSE_HISTORY_COLUMNS)
    latest = latest_pause_records(pause_history_raw)
    if latest.empty:
        return pd.DataFrame(columns=PAUSE_HISTORY_COLUMNS)

    managed_keys: set[Tuple[str, str, str]] = set()
    if managed_ads_df is not None and not managed_ads_df.empty:
        for _, r in managed_ads_df[["campaign_id", "nm_id", "placement"]].drop_duplicates().iterrows():
            key = make_key(r)
            if all(key):
                managed_keys.add(key)

    rows: List[Dict[str, Any]] = []
    for _, row in latest.iterrows():
        status = _clean_text_value(row.get("status", "")).lower()
        if status not in {"paused", "keep_paused"}:
            continue
        key = make_key(row)
        if not all(key):
            continue
        if key in managed_keys:
            continue
        rows.append({
            "pause_event_id": str(uuid.uuid4()),
            "pause_date": date.today().isoformat(),
            "campaign_id": row.get("campaign_id", ""),
            "nm_id": row.get("nm_id", ""),
            "placement": row.get("placement", ""),
            "supplier_article": row.get("supplier_article", ""),
            "reason_code": "ROLLBACK_WRONG_SUBJECT_PAUSE",
            "spend_before_pause": row.get("spend_before_pause", 0),
            "revenue_before_pause": row.get("revenue_before_pause", 0),
            "orders_before_pause": row.get("orders_before_pause", 0),
            "drr_before_pause": row.get("drr_before_pause", 0),
            "gp_before_pause": row.get("gp_before_pause", float("nan")),
            "status": "restart_candidate",
            "next_check_date": date.today().isoformat(),
            "api_status": "",
        })
    return pd.DataFrame(rows, columns=PAUSE_HISTORY_COLUMNS)


def apply_start_actions(start_candidates: pd.DataFrame, config: Config, ctx: RunContext) -> Tuple[pd.DataFrame, pd.DataFrame]:
    if start_candidates.empty:
        return start_candidates.copy(), pd.DataFrame()
    result = start_candidates.copy()
    api_logs: List[Dict[str, Any]] = []
    if ctx.mode == "preview":
        result["api_status"] = "preview_no_call"
        return result, pd.DataFrame(api_logs)
    if ctx.dry_run:
        result["api_status"] = "dry_run_no_call"
        return result, pd.DataFrame(api_logs)
    if not ctx.apply_start:
        rollback_mask = result.get("reason_code", pd.Series(dtype=str)).astype(str).eq("ROLLBACK_WRONG_SUBJECT_PAUSE")
        if not rollback_mask.any():
            result["api_status"] = "not_applied_without_flag"
            return result, pd.DataFrame(api_logs)
        # Технический rollback ошибочных пауз прошлой версии запускаем автоматически в обычном run.
        result.loc[~rollback_mask, "api_status"] = "not_applied_without_flag"
        result = result.loc[rollback_mask].copy()

    url_base = config.wb_base_url.rstrip("/") + WB_START_ENDPOINT
    status_by_campaign: Dict[str, Tuple[str, str]] = {}
    for campaign_id in sorted(result["campaign_id"].map(_clean_id_value).unique()):
        advert_id = to_int_id(campaign_id)
        if advert_id is None:
            status_by_campaign[campaign_id] = ("payload_error", "campaign_id не является числом")
            continue
        endpoint = f"{WB_START_ENDPOINT}?id={advert_id}"
        try:
            resp = requests.get(url_base, params={"id": advert_id}, headers=wb_headers(config), timeout=60)
            status_by_campaign[campaign_id] = (str(resp.status_code), resp.text)
            api_logs.append(api_log_row(ctx.run_datetime, "GET", endpoint, "", str(resp.status_code), resp.text, campaign_id=campaign_id))
        except Exception as exc:
            status_by_campaign[campaign_id] = ("exception", repr(exc))
            api_logs.append(api_log_row(ctx.run_datetime, "GET", endpoint, "", "exception", repr(exc), campaign_id=campaign_id))

    statuses: List[str] = []
    final_statuses: List[str] = []
    for _, row in result.iterrows():
        api_status, _ = status_by_campaign.get(_clean_id_value(row.get("campaign_id", "")), ("not_sent", ""))
        statuses.append(api_status)
        if api_status.isdigit() and 200 <= int(api_status) < 300:
            final_statuses.append("started")
        else:
            final_statuses.append("restart_candidate")
    result["api_status"] = statuses
    result["status"] = final_statuses
    return result, pd.DataFrame(api_logs)


# =============================
# Запись Excel / JSON
# =============================

def dataframe_to_excel_bytes(sheets: Dict[str, pd.DataFrame]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            safe_name = sheet_name[:31]
            local_df = df.copy() if df is not None else pd.DataFrame()
            local_df.to_excel(writer, index=False, sheet_name=safe_name)
            ws = writer.book[safe_name]
            ws.freeze_panes = "A2"
            for col_cells in ws.columns:
                max_len = 0
                col_letter = col_cells[0].column_letter
                for cell in col_cells:
                    value = cell.value
                    if value is not None:
                        max_len = max(max_len, len(str(value)))
                ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 60)
    return output.getvalue()


def save_table_to_s3_excel(s3_client, config: Config, key: str, df: pd.DataFrame) -> None:
    payload = dataframe_to_excel_bytes({"Лист1": df})
    upload_s3_bytes(s3_client, config.yc_bucket_name, key, payload, content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def append_api_log_to_s3(s3_client, config: Config, new_log: pd.DataFrame) -> pd.DataFrame:
    existing = load_excel_table_from_s3(s3_client, config, API_LOG_KEY, [
        "run_datetime", "method", "endpoint", "campaign_id", "nm_id", "placement", "payload", "api_status", "response_text"
    ])
    if new_log is None or new_log.empty:
        combined = existing
    else:
        combined = pd.concat([existing, new_log], ignore_index=True, sort=False)
    save_table_to_s3_excel(s3_client, config, API_LOG_KEY, combined)
    return combined



# =============================
# Разгон показов и эксперимент 1 РК на товарную группу
# =============================


def enrich_supplier_articles_from_rename_plan(df: pd.DataFrame, rename_plan: Optional[pd.DataFrame]) -> pd.DataFrame:
    """Подтягивает короткий артикул продавца из плана переименования в основные листы отчёта."""
    if df is None or df.empty or rename_plan is None or rename_plan.empty:
        return df if df is not None else pd.DataFrame()
    if "campaign_id" not in df.columns:
        return df
    result = df.copy()
    if "supplier_article" not in result.columns:
        result["supplier_article"] = ""
    rp = rename_plan.copy()
    if "campaign_id" not in rp.columns or "supplier_article" not in rp.columns:
        return result
    rp["campaign_id_clean"] = rp["campaign_id"].map(_clean_id_value)
    article_map = {
        _clean_id_value(r.get("campaign_id", "")): _clean_text_value(r.get("supplier_article", ""))
        for _, r in rp.iterrows()
        if _clean_id_value(r.get("campaign_id", "")) and _clean_text_value(r.get("supplier_article", ""))
    }
    for idx, row in result.iterrows():
        cur = _clean_text_value(row.get("supplier_article", ""))
        if cur:
            continue
        art = article_map.get(_clean_id_value(row.get("campaign_id", "")), "")
        if art:
            result.at[idx, "supplier_article"] = art
    return result


def enrich_decisions_with_bid_api_status(decisions: pd.DataFrame, successful_changes: Optional[pd.DataFrame]) -> pd.DataFrame:
    """Фиксирует в Решениях, какие ставки реально ушли в WB, и обновляет ramp_status."""
    if decisions is None or decisions.empty:
        return decisions if decisions is not None else pd.DataFrame(columns=DECISION_COLUMNS)
    result = decisions.copy()
    if "ramp_api_status" not in result.columns:
        result["ramp_api_status"] = ""
    if "ramp_applied_in_current_run" not in result.columns:
        result["ramp_applied_in_current_run"] = False
    success_map: Dict[Tuple[str, str, str], str] = {}
    if successful_changes is not None and not successful_changes.empty:
        for _, r in successful_changes.iterrows():
            success_map[make_key(r)] = _clean_text_value(r.get("api_status", ""))
    for idx, row in result.iterrows():
        key = make_key(row)
        if key in success_map:
            result.at[idx, "ramp_api_status"] = success_map[key]
        tmp = result.loc[idx].to_dict()
        if key in success_map:
            tmp["api_status"] = success_map[key]
        status, reason_group, applied = classify_ramp_status(tmp)
        result.at[idx, "ramp_candidate"] = bool(is_ramp_candidate_by_metrics(tmp))
        result.at[idx, "ramp_status"] = status
        result.at[idx, "ramp_applied_in_current_run"] = bool(key in success_map and is_ramp_related_reason(tmp.get("reason_code", ""), tmp.get("wait_rule", ""), tmp.get("last_bid_change_reason_code", "")))
    for col in DECISION_COLUMNS:
        if col not in result.columns:
            result[col] = ""
    return result[DECISION_COLUMNS]


def build_bid_ramp_monitor(decisions: pd.DataFrame) -> pd.DataFrame:
    """Лист Разгон_показов: показывает не только отправленные изменения, но и весь статус режима Разгон.

    В лист попадают:
    - кампании, которые подходят под разгон по метрикам: <1000 показов/день и расход <=500 ₽/день;
    - кампании, где режим Разгон уже активен и ждёт D+7;
    - кампании, где Разгон был заблокирован/скорректирован минимумом ставки WB.
    """
    if decisions is None or decisions.empty:
        return pd.DataFrame(columns=BID_RAMP_MONITOR_COLUMNS)
    rows: List[Dict[str, Any]] = []
    for _, row in decisions.iterrows():
        rc = _clean_text_value(row.get("reason_code", ""))
        related = is_ramp_related_reason(rc, row.get("wait_rule", ""), row.get("last_bid_change_reason_code", ""))
        candidate = is_ramp_candidate_by_metrics(row)
        if not related and not candidate:
            continue
        status, reason_group, applied = classify_ramp_status(row)
        rows.append({
            "campaign_id": row.get("campaign_id", ""),
            "nm_id": row.get("nm_id", ""),
            "supplier_article": row.get("supplier_article", ""),
            "subject_norm": row.get("subject_norm", ""),
            "placement": row.get("placement", ""),
            "campaign_status": row.get("campaign_status", ""),
            "ramp_candidate": bool(candidate),
            "ramp_mode_status": status,
            "ramp_applied_in_current_run": bool(row.get("ramp_applied_in_current_run", False)),
            "ramp_reason_group": reason_group,
            "current_bid_rub": row.get("current_bid_rub", 0),
            "min_bid_rub": row.get("min_bid_rub", ""),
            "new_bid_rub": row.get("new_bid_rub", ""),
            "api_status": row.get("ramp_api_status", ""),
            "reason_code": rc,
            "reason_text": row.get("reason_text", ""),
            "wait_status": row.get("wait_status", ""),
            "wait_rule": row.get("wait_rule", ""),
            "wait_until_date": row.get("wait_until_date", ""),
            "wait_days_left": row.get("wait_days_left", ""),
            "last_bid_change_date": row.get("last_bid_change_date", ""),
            "last_bid_change_old_bid": row.get("last_bid_change_old_bid", ""),
            "last_bid_change_new_bid": row.get("last_bid_change_new_bid", ""),
            "last_bid_change_reason_code": row.get("last_bid_change_reason_code", ""),
            "impressions": row.get("impressions", 0),
            "avg_impressions_per_day": row.get("avg_impressions_per_day", 0),
            "spend": row.get("spend", 0),
            "avg_spend_per_day": row.get("avg_spend_per_day", 0),
            "orders": row.get("orders", 0),
            "revenue": row.get("revenue", 0),
            "campaign_drr_pct": row.get("campaign_drr_pct", 0),
            "drr_limit_pct": row.get("drr_limit_pct", 0),
            "last21_impressions": row.get("last21_impressions", 0),
            "last21_drr_pct": row.get("last21_drr_pct", 0),
            "target_impressions_per_day": RAMP_TARGET_IMPRESSIONS_PER_DAY,
            "max_spend_per_day": RAMP_MAX_SPEND_PER_DAY,
            "check_days": RAMP_CHECK_DAYS,
            "monitor_status": status or "not_ramp",
        })
    out = pd.DataFrame(rows)
    for col in BID_RAMP_MONITOR_COLUMNS:
        if col not in out.columns:
            out[col] = ""
    if not out.empty:
        out = out.sort_values(["ramp_reason_group", "subject_norm", "avg_impressions_per_day"], ascending=[True, True, True])
    return out[BID_RAMP_MONITOR_COLUMNS]

def keyword_core_stats_for_group(keyword_core_df: pd.DataFrame, subject_norm: str, product_group: str, supplier_articles: Iterable[Any]) -> Dict[str, Any]:
    if keyword_core_df is None or keyword_core_df.empty:
        return {"core_queries_count": 0, "core_median_position": float("nan")}
    work = keyword_core_df.copy()
    if "keyword_group" in work.columns:
        work = work[work["keyword_group"] == "CORE_80"].copy()
    if work.empty:
        return {"core_queries_count": 0, "core_median_position": float("nan")}
    work["subject_norm"] = work.get("subject_norm", "").map(normalize_subject_value)
    articles = set(_clean_text_value(x) for x in supplier_articles if _clean_text_value(x))
    if "supplier_article" in work.columns and articles:
        work = work[work["supplier_article"].astype(str).isin(articles)].copy()
    else:
        work["product_group"] = work.get("supplier_article", "").map(product_group_from_article)
        work = work[(work["subject_norm"] == subject_norm) & (work["product_group"] == product_group)].copy()
    if work.empty:
        return {"core_queries_count": 0, "core_median_position": float("nan")}
    return {
        "core_queries_count": int(work.get("query_text", pd.Series(dtype=str)).nunique()),
        "core_median_position": float(pd.to_numeric(work.get("median_position", pd.Series(dtype=float)), errors="coerce").median()),
    }


def build_one_campaign_experiment(metrics_df: pd.DataFrame, keyword_core_df: pd.DataFrame) -> pd.DataFrame:
    """Формирует отдельный экспериментальный план: одна лучшая РК на товарную группу.

    Для помад/блесков/косметических карандашей выбираем campaign_id с максимальными кликами,
    затем CTR, затем заказами. Остальные campaign_id в этой товарной группе рекомендуются к паузе
    только в рамках экспериментального блока. Главная цель — вывести выбранную РК в топ-10 по CORE_80.
    """
    if metrics_df is None or metrics_df.empty:
        return pd.DataFrame(columns=ONE_CAMPAIGN_EXPERIMENT_COLUMNS)
    work = metrics_df.copy()
    work["subject_norm"] = work["subject_norm"].map(normalize_subject_value)
    work = work[work["subject_norm"].map(is_one_campaign_experiment_subject)].copy()
    work = work[work["campaign_status"].map(is_active_campaign)].copy()
    if work.empty:
        return pd.DataFrame(columns=ONE_CAMPAIGN_EXPERIMENT_COLUMNS)
    work["product_group"] = work["supplier_article"].map(product_group_from_article)
    work = work[work["product_group"].astype(str).str.strip() != ""].copy()
    if work.empty:
        return pd.DataFrame(columns=ONE_CAMPAIGN_EXPERIMENT_COLUMNS)

    rows: List[Dict[str, Any]] = []
    for (subject, product_group), g in work.groupby(["subject_norm", "product_group"], dropna=False):
        campaign_count = int(g["campaign_id"].astype(str).nunique())
        if campaign_count < 2:
            continue
        g = g.copy()
        g["ctr_for_sort"] = pd.to_numeric(g.get("ctr_pct", 0), errors="coerce").fillna(0.0)
        g["clicks_for_sort"] = pd.to_numeric(g.get("clicks", 0), errors="coerce").fillna(0.0)
        g["orders_for_sort"] = pd.to_numeric(g.get("orders", 0), errors="coerce").fillna(0.0)
        g["spend_for_sort"] = pd.to_numeric(g.get("spend", 0), errors="coerce").fillna(0.0)
        selected = g.sort_values(["clicks_for_sort", "ctr_for_sort", "orders_for_sort", "spend_for_sort"], ascending=False).iloc[0]
        selected_campaign_id = _clean_id_value(selected.get("campaign_id", ""))
        to_pause = sorted(set(_clean_id_value(x) for x in g["campaign_id"].tolist()) - {selected_campaign_id})
        if not to_pause:
            continue

        group_spend = money_or_zero(g["spend"].sum())
        group_revenue = money_or_zero(g["revenue"].sum())
        group_orders = money_or_zero(g["orders"].sum())
        group_impressions = money_or_zero(g["impressions"].sum())
        group_clicks = money_or_zero(g["clicks"].sum())
        group_gp_series = pd.to_numeric(g.get("gp_after_ads", pd.Series(dtype=float)), errors="coerce")
        group_gp = float(group_gp_series.sum()) if group_gp_series.notna().any() else float("nan")
        group_ctr = safe_ctr_pct(group_clicks, group_impressions)
        group_conversion = (group_orders / group_clicks * 100.0) if group_clicks > 0 else 0.0

        selected_gp = money_or_zero(selected.get("gp_after_ads", 0))
        selected_clicks = money_or_zero(selected.get("clicks", 0))
        selected_orders = money_or_zero(selected.get("orders", 0))
        selected_impressions = money_or_zero(selected.get("impressions", 0))
        selected_conversion = (selected_orders / selected_clicks * 100.0) if selected_clicks > 0 else 0.0
        step, _ = bid_step_rub(selected.get("placement", ""))
        selected_bid = money_or_zero(selected.get("current_bid_rub", 0))
        recommended_new_bid = round(selected_bid + step, 2) if selected_bid > 0 else ""

        kw_stats = keyword_core_stats_for_group(keyword_core_df, subject, product_group, g["supplier_article"].unique())
        rows.append({
            "experiment_id": str(uuid.uuid4()),
            "run_datetime": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "subject_norm": subject,
            "product_group": product_group,
            "selected_campaign_id": selected_campaign_id,
            "selected_nm_id": selected.get("nm_id", ""),
            "selected_supplier_article": selected.get("supplier_article", ""),
            "selected_placement": selected.get("placement", ""),
            "selected_bid_rub": selected_bid,
            "recommended_new_bid_rub": recommended_new_bid,
            "selection_basis": "max_clicks_then_ctr_then_orders",
            "group_campaigns_count": campaign_count,
            "campaigns_to_pause": len(to_pause),
            "pause_campaign_ids": ",".join(to_pause),
            "group_gp_before": group_gp,
            "group_orders_before": group_orders,
            "group_revenue_before": group_revenue,
            "group_spend_before": group_spend,
            "group_impressions_before": group_impressions,
            "group_clicks_before": group_clicks,
            "group_ctr_before": group_ctr,
            "group_conversion_before": group_conversion,
            "selected_gp_before": selected_gp,
            "selected_orders_before": selected_orders,
            "selected_revenue_before": selected.get("revenue", 0),
            "selected_spend_before": selected.get("spend", 0),
            "selected_impressions_before": selected_impressions,
            "selected_clicks_before": selected_clicks,
            "selected_ctr_before": selected.get("ctr_pct", 0),
            "selected_conversion_before": selected_conversion,
            "core_queries_count": kw_stats["core_queries_count"],
            "core_median_position": kw_stats["core_median_position"],
            "core_target_position": ONE_CAMPAIGN_TARGET_POSITION,
            "recommended_action": "leave_one_campaign_raise_selected_pause_others",
            "reason_code": "ONE_CAMPAIGN_PER_PRODUCT_GROUP_EXPERIMENT",
            "reason_text": (
                f"выбрана РК {selected_campaign_id}: клики={selected_clicks:.0f}, CTR={money_or_zero(selected.get('ctr_pct',0)):.2f}%; "
                f"остальные РК группы к паузе={len(to_pause)}; цель — топ-{ONE_CAMPAIGN_TARGET_POSITION} по CORE_80; "
                f"сравнение через 7/10/14 дней: ВП группы, показы, клики, CTR, конверсия"
            ),
            "check_days": "/".join(str(x) for x in ONE_CAMPAIGN_CHECK_DAYS),
        })
    return pd.DataFrame(rows, columns=ONE_CAMPAIGN_EXPERIMENT_COLUMNS)


def build_experiment_bid_decisions(one_campaign_experiment: pd.DataFrame, decisions: pd.DataFrame) -> pd.DataFrame:
    if one_campaign_experiment is None or one_campaign_experiment.empty:
        return pd.DataFrame(columns=DECISION_COLUMNS)
    decision_lookup = decisions.copy() if decisions is not None else pd.DataFrame(columns=DECISION_COLUMNS)
    rows: List[Dict[str, Any]] = []
    for _, exp in one_campaign_experiment.iterrows():
        new_bid = money_or_zero(exp.get("recommended_new_bid_rub", 0))
        if new_bid <= 0:
            continue
        selected_campaign_id = _clean_id_value(exp.get("selected_campaign_id", ""))
        selected_nm_id = _clean_id_value(exp.get("selected_nm_id", ""))
        selected_placement = normalize_placement_value(exp.get("selected_placement", ""))
        base = pd.Series(dtype=object)
        if not decision_lookup.empty:
            m = decision_lookup[
                (decision_lookup["campaign_id"].astype(str).map(_clean_id_value) == selected_campaign_id)
                & (decision_lookup["nm_id"].astype(str).map(_clean_id_value) == selected_nm_id)
                & (decision_lookup["placement"].astype(str).map(normalize_placement_value) == selected_placement)
            ]
            if not m.empty:
                base = m.iloc[0]
        row = {col: base.get(col, "") if isinstance(base, pd.Series) else "" for col in DECISION_COLUMNS}
        row.update({
            "campaign_id": selected_campaign_id,
            "nm_id": selected_nm_id,
            "supplier_article": exp.get("selected_supplier_article", row.get("supplier_article", "")),
            "subject_norm": exp.get("subject_norm", row.get("subject_norm", "")),
            "placement": selected_placement,
            "current_bid_rub": exp.get("selected_bid_rub", row.get("current_bid_rub", 0)),
            "new_bid_rub": new_bid,
            "action": "Повысить",
            "reason_code": "EXPERIMENT_ONE_CAMPAIGN_RAISE_TO_TOP10_CORE80",
            "reason_text": exp.get("reason_text", "эксперимент 1 РК на товарную группу: повышаем выбранную РК"),
            "pause_decision": "",
        })
        rows.append(row)
    return pd.DataFrame(rows, columns=DECISION_COLUMNS)


def build_experiment_pause_candidates(one_campaign_experiment: pd.DataFrame, metrics_df: pd.DataFrame) -> pd.DataFrame:
    if one_campaign_experiment is None or one_campaign_experiment.empty or metrics_df is None or metrics_df.empty:
        return pd.DataFrame(columns=PAUSE_HISTORY_COLUMNS)
    metrics = metrics_df.copy()
    metrics["campaign_id_clean"] = metrics["campaign_id"].astype(str).map(_clean_id_value)
    rows: List[Dict[str, Any]] = []
    for _, exp in one_campaign_experiment.iterrows():
        pause_ids = [x.strip() for x in str(exp.get("pause_campaign_ids", "")).split(",") if x.strip()]
        for campaign_id in pause_ids:
            g = metrics[metrics["campaign_id_clean"] == _clean_id_value(campaign_id)].copy()
            if g.empty:
                continue
            main = g.sort_values(["spend", "clicks", "impressions"], ascending=False).iloc[0]
            rows.append({
                "pause_event_id": str(uuid.uuid4()),
                "pause_date": date.today().isoformat(),
                "campaign_id": _clean_id_value(campaign_id),
                "nm_id": main.get("nm_id", ""),
                "placement": main.get("placement", ""),
                "supplier_article": main.get("supplier_article", ""),
                "subject_norm": main.get("subject_norm", exp.get("subject_norm", "")),
                "reason_code": "EXPERIMENT_ONE_CAMPAIGN_PAUSE_DUPLICATE",
                "impressions_before_pause": money_or_zero(g["impressions"].sum()),
                "clicks_before_pause": money_or_zero(g["clicks"].sum()),
                "spend_before_pause": money_or_zero(g["spend"].sum()),
                "revenue_before_pause": money_or_zero(g["revenue"].sum()),
                "orders_before_pause": money_or_zero(g["orders"].sum()),
                "drr_before_pause": safe_drr_pct(money_or_zero(g["spend"].sum()), money_or_zero(g["revenue"].sum())),
                "gp_before_pause": money_or_zero(pd.to_numeric(g.get("gp_after_ads", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()),
                "status": "candidate",
                "next_check_date": (date.today() + timedelta(days=7)).isoformat(),
                "api_status": "",
            })
    return pd.DataFrame(rows, columns=PAUSE_HISTORY_COLUMNS)

def build_summary(ctx: RunContext, decisions: pd.DataFrame, successful_changes: pd.DataFrame, pause_candidates: pd.DataFrame, applied_pauses: pd.DataFrame, start_candidates: pd.DataFrame, applied_starts: pd.DataFrame) -> Dict[str, Any]:
    changed_count = int(len(successful_changes)) if successful_changes is not None else 0
    pause_applied_count = 0
    if applied_pauses is not None and not applied_pauses.empty:
        pause_applied_count = int((applied_pauses["status"] == "paused").sum())
    start_applied_count = 0
    if applied_starts is not None and not applied_starts.empty:
        start_applied_count = int((applied_starts["status"] == "started").sum())
    recommendation_count = 0
    if decisions is not None and not decisions.empty:
        recommendation_count = int(decisions["action"].isin(["Повысить", "Снизить"]).sum())

    return {
        "Режим": ctx.mode if not ctx.dry_run else f"{ctx.mode} --dry-run",
        "Дата формирования": ctx.run_datetime.strftime("%Y-%m-%d %H:%M:%S"),
        "Всего рекомендаций": recommendation_count,
        "Изменённых ставок": changed_count,
        "Блоков отправки ставок": recommendation_count if ctx.mode == "run" and not ctx.dry_run else 0,
        "Кандидатов на паузу": int(len(pause_candidates)) if pause_candidates is not None else 0,
        "Поставлено на паузу": pause_applied_count,
        "Кандидатов на запуск": int(len(start_candidates)) if start_candidates is not None else 0,
        "Запущено обратно": start_applied_count,
        "Текущее окно с": ctx.current_start.isoformat(),
        "Текущее окно по": ctx.current_end.isoformat(),
        "База с": ctx.base_start.isoformat(),
        "База по": ctx.base_end.isoformat(),
    }



# =============================
# Диагностика РК: нормальное сравнение 7 дней до/после изменения ставки
# =============================

def _date_window_label(start_date: Optional[date], end_date: Optional[date]) -> str:
    if not start_date or not end_date or end_date < start_date:
        return ""
    return f"{start_date.isoformat()}..{end_date.isoformat()}"


def _window_days_count(start_date: Optional[date], end_date: Optional[date]) -> int:
    if not start_date or not end_date or end_date < start_date:
        return 0
    return int((end_date - start_date).days + 1)


def _sum_window_campaign_metrics(ads_df: pd.DataFrame, key: Tuple[str, str, str], start_date: Optional[date], end_date: Optional[date]) -> Dict[str, float]:
    if not start_date or not end_date or end_date < start_date:
        return {"spend": 0.0, "revenue": 0.0, "orders": 0.0, "impressions": 0.0, "clicks": 0.0, "gp_after_ads": float("nan")}
    return aggregate_after_event(ads_df, key, start_date, end_date)


def _safe_per_day_value(value: Any, days: int) -> float:
    if days <= 0:
        return 0.0
    val = pd.to_numeric(pd.Series([value]), errors="coerce").iloc[0]
    if pd.isna(val):
        return float("nan")
    return float(val) / float(days)


def _safe_delta_pct(after: Any, before: Any) -> Any:
    before_num = pd.to_numeric(pd.Series([before]), errors="coerce").iloc[0]
    after_num = pd.to_numeric(pd.Series([after]), errors="coerce").iloc[0]
    if pd.isna(before_num) or pd.isna(after_num) or float(before_num) == 0:
        return ""
    return (float(after_num) / float(before_num) - 1.0) * 100.0


def _safe_delta_pp(after: Any, before: Any) -> Any:
    before_num = pd.to_numeric(pd.Series([before]), errors="coerce").iloc[0]
    after_num = pd.to_numeric(pd.Series([after]), errors="coerce").iloc[0]
    if pd.isna(before_num) or pd.isna(after_num):
        return ""
    return float(after_num) - float(before_num)


def _weighted_average(values: pd.Series, weights: pd.Series) -> float:
    vals = pd.to_numeric(values, errors="coerce")
    w = pd.to_numeric(weights, errors="coerce").fillna(0.0)
    mask = vals.notna() & (w > 0)
    if mask.any() and float(w.loc[mask].sum()) > 0:
        return float((vals.loc[mask] * w.loc[mask]).sum() / w.loc[mask].sum())
    vals = vals.dropna()
    return float(vals.mean()) if not vals.empty else 0.0


def _core80_queries_for_nm(keyword_core_df: pd.DataFrame, nm_id: str, supplier_article: str = "") -> set[str]:
    if keyword_core_df is None or keyword_core_df.empty or "query_text_norm" not in keyword_core_df.columns:
        return set()
    core = keyword_core_df.copy()
    if "keyword_group" in core.columns:
        core = core[core["keyword_group"].astype(str).eq("CORE_80")].copy()
    core = core[core["nm_id"].astype(str).map(_clean_id_value).eq(_clean_id_value(nm_id))].copy() if "nm_id" in core.columns else core.iloc[0:0].copy()
    if supplier_article and "supplier_article" in core.columns:
        art_core = core[core["supplier_article"].astype(str).map(_clean_text_value).eq(_clean_text_value(supplier_article))].copy()
        if not art_core.empty:
            core = art_core
    return set(core["query_text_norm"].astype(str).str.strip().str.lower().tolist())


def _aggregate_core80_keyword_window(keyword_df: pd.DataFrame, keyword_core_df: pd.DataFrame, nm_id: str, supplier_article: str, start_date: Optional[date], end_date: Optional[date]) -> Dict[str, float]:
    empty = {"queries_count": 0.0, "avg_position": 0.0, "visibility_pct": 0.0, "query_freq": 0.0, "clicks_to_card": 0.0, "keyword_orders": 0.0}
    if keyword_df is None or keyword_df.empty or not start_date or not end_date or end_date < start_date:
        return empty
    queries = _core80_queries_for_nm(keyword_core_df, nm_id, supplier_article)
    if not queries:
        return empty
    part = keyword_df.copy()
    if "nm_id" in part.columns:
        part = part[part["nm_id"].astype(str).map(_clean_id_value).eq(_clean_id_value(nm_id))].copy()
    if supplier_article and "supplier_article" in part.columns:
        art_part = part[part["supplier_article"].astype(str).map(_clean_text_value).eq(_clean_text_value(supplier_article))].copy()
        if not art_part.empty:
            part = art_part
    if "query_text_norm" in part.columns:
        part = part[part["query_text_norm"].astype(str).str.strip().str.lower().isin(queries)].copy()
    if has_valid_dates(part):
        part = part[(part["date"] >= start_date) & (part["date"] <= end_date)].copy()
    if part.empty:
        return {**empty, "queries_count": float(len(queries))}
    query_freq = pd.to_numeric(part.get("query_freq", pd.Series(dtype=float)), errors="coerce").fillna(0.0)
    return {
        "queries_count": float(len(queries)),
        "avg_position": _weighted_average(part.get("median_position", pd.Series(dtype=float)), query_freq),
        "visibility_pct": _weighted_average(part.get("visibility_pct", pd.Series(dtype=float)), query_freq),
        "query_freq": float(query_freq.sum()),
        "clicks_to_card": float(pd.to_numeric(part.get("clicks_to_card", pd.Series(dtype=float)), errors="coerce").fillna(0.0).sum()),
        "keyword_orders": float(pd.to_numeric(part.get("keyword_orders", pd.Series(dtype=float)), errors="coerce").fillna(0.0).sum()),
    }


def _aggregate_funnel_card_views_window(funnel_df: pd.DataFrame, nm_id: str, start_date: Optional[date], end_date: Optional[date]) -> float:
    if funnel_df is None or funnel_df.empty or not start_date or not end_date or end_date < start_date:
        return 0.0
    part = funnel_df.copy()
    if "nm_id" not in part.columns:
        return 0.0
    part = part[part["nm_id"].astype(str).map(_clean_id_value).eq(_clean_id_value(nm_id))].copy()
    if has_valid_dates(part):
        part = part[(part["date"] >= start_date) & (part["date"] <= end_date)].copy()
    if part.empty:
        return 0.0
    return float(pd.to_numeric(part.get("card_views", pd.Series(dtype=float)), errors="coerce").fillna(0.0).sum())


def _latest_bid_event_by_key(bid_history: pd.DataFrame) -> Dict[Tuple[str, str, str], Dict[str, Any]]:
    if bid_history is None or bid_history.empty:
        return {}
    local = bid_history.copy()
    for col in BID_HISTORY_COLUMNS:
        if col not in local.columns:
            local[col] = ""
    local["event_date_parsed"] = pd.to_datetime(local["event_date"], errors="coerce")
    local["run_dt_parsed"] = pd.to_datetime(local["run_datetime"], errors="coerce")
    local = local.sort_values(["event_date_parsed", "run_dt_parsed"], na_position="first")
    out: Dict[Tuple[str, str, str], Dict[str, Any]] = {}
    for _, row in local.iterrows():
        key = make_key(row)
        if all(key):
            out[key] = row.to_dict()
    return out


def _effect_by_event_id(effect_df: Optional[pd.DataFrame]) -> Dict[str, Dict[str, Any]]:
    if effect_df is None or effect_df.empty or "event_id" not in effect_df.columns:
        return {}
    out: Dict[str, Dict[str, Any]] = {}
    for _, row in effect_df.iterrows():
        eid = _clean_text_value(row.get("event_id", ""))
        if eid:
            out[eid] = row.to_dict()
    return out


def _bid_compare_conclusion(
    action: str,
    reason_code: str,
    final_verdict: str,
    target_action: str,
    before_gp_day: Any,
    after_gp_day: Any,
    before_drr: float,
    after_drr: float,
    before_pos: float,
    after_pos: float,
    before_visibility: float,
    after_visibility: float,
) -> str:
    gp_delta = _safe_delta_pct(after_gp_day, before_gp_day)
    drr_delta = _safe_delta_pp(after_drr, before_drr)
    pos_delta = _safe_delta_pp(after_pos, before_pos)
    vis_delta = _safe_delta_pp(after_visibility, before_visibility)
    gp_text = "условная ВП после рекламы/день н/д"
    if gp_delta != "":
        gp_text = f"условная ВП после рекламы/день {'выросла' if float(gp_delta) >= 0 else 'упала'} на {float(gp_delta):.1f}%"
    drr_text = f"ДРР {before_drr:.2f}%→{after_drr:.2f}%" if before_drr or after_drr else "ДРР н/д"
    pos_text = "позиция CORE80 н/д"
    if pos_delta != "":
        pos_text = f"позиция CORE80 {before_pos:.1f}→{after_pos:.1f} ({float(pos_delta):+.1f}; меньше лучше)"
    vis_text = "видимость CORE80 н/д"
    if vis_delta != "":
        vis_text = f"видимость {before_visibility:.1f}%→{after_visibility:.1f}% ({float(vis_delta):+.1f} п.п.)"
    decision_text = f"Решение кода: {action or 'Без изменений'} / {reason_code or 'без reason_code'}"
    if final_verdict:
        decision_text += f"; post-check={final_verdict}"
    if target_action:
        decision_text += f"; целевое действие={target_action}"
    return "; ".join([gp_text, drr_text, pos_text, vis_text, decision_text])


def build_campaign_7d_comparison(
    decisions: pd.DataFrame,
    bid_history: pd.DataFrame,
    ads_df: pd.DataFrame,
    keyword_df: pd.DataFrame,
    keyword_core_df: pd.DataFrame,
    funnel_df: pd.DataFrame,
    effect_df: Optional[pd.DataFrame],
    ctx: RunContext,
) -> pd.DataFrame:
    """Строит диагностический лист: РК × nm × placement, сравнение 7 дней до/после последней ставки.

    Если по строке есть последняя успешная правка ставки, окно ДО = 7 дней до event_date,
    окно ПОСЛЕ = до 7 зрелых дней после event_date. Если истории правки нет, сравниваем
    последние 7 зрелых дней с предыдущими 7 днями. Метрики CORE_80 берутся по продающим
    запросам, которые формируют 80% заказов по SKU; доля трафика = переходы в карточку
    из воронки / частотность CORE_80.
    """
    if decisions is None or decisions.empty:
        return pd.DataFrame(columns=BID_CAMPAIGN_COMPARE_COLUMNS)
    latest_events = _latest_bid_event_by_key(bid_history)
    effects = _effect_by_event_id(effect_df)
    rows: List[Dict[str, Any]] = []
    current_after_end = ctx.mature_end
    current_after_start = current_after_end - timedelta(days=6)
    current_before_end = current_after_start - timedelta(days=1)
    current_before_start = current_before_end - timedelta(days=6)

    for _, decision in decisions.iterrows():
        key = make_key(decision)
        if not all(key):
            continue
        event = latest_events.get(key)
        comparison_status = "NO_BID_CHANGE_HISTORY_COMPARE_LAST_7D"
        old_bid = ""
        new_bid = ""
        change_reason = ""
        event_date_text = ""
        if event:
            event_dt = pd.to_datetime(event.get("event_date", ""), errors="coerce")
            if not pd.isna(event_dt):
                event_day = event_dt.date()
                before_start = event_day - timedelta(days=7)
                before_end = event_day - timedelta(days=1)
                after_start = event_day + timedelta(days=1)
                after_end = min(ctx.mature_end, event_day + timedelta(days=7))
                comparison_status = "BID_CHANGE_7D_BEFORE_AFTER"
                event_date_text = event_day.isoformat()
            else:
                before_start, before_end = current_before_start, current_before_end
                after_start, after_end = current_after_start, current_after_end
            old_bid = event.get("old_bid_rub", "")
            new_bid = event.get("new_bid_rub", "")
            change_reason = _clean_text_value(event.get("reason_code", ""))
        else:
            before_start, before_end = current_before_start, current_before_end
            after_start, after_end = current_after_start, current_after_end

        before_days = _window_days_count(before_start, before_end)
        after_days = _window_days_count(after_start, after_end)
        before_metrics = _sum_window_campaign_metrics(ads_df, key, before_start, before_end)
        after_metrics = _sum_window_campaign_metrics(ads_df, key, after_start, after_end)
        before_daily = _per_day(before_metrics, before_days or 1)
        after_daily = _per_day(after_metrics, after_days or 1)
        before_ctr = safe_ctr_pct(before_metrics.get("clicks", 0), before_metrics.get("impressions", 0))
        after_ctr = safe_ctr_pct(after_metrics.get("clicks", 0), after_metrics.get("impressions", 0))
        before_drr = safe_drr_pct(before_metrics.get("spend", 0), before_metrics.get("revenue", 0))
        after_drr = safe_drr_pct(after_metrics.get("spend", 0), after_metrics.get("revenue", 0))

        nm_id = _clean_id_value(decision.get("nm_id", ""))
        supplier_article = _clean_text_value(decision.get("supplier_article", ""))
        before_kw = _aggregate_core80_keyword_window(keyword_df, keyword_core_df, nm_id, supplier_article, before_start, before_end)
        after_kw = _aggregate_core80_keyword_window(keyword_df, keyword_core_df, nm_id, supplier_article, after_start, after_end)
        before_card_views = _aggregate_funnel_card_views_window(funnel_df, nm_id, before_start, before_end)
        after_card_views = _aggregate_funnel_card_views_window(funnel_df, nm_id, after_start, after_end)
        before_share = safe_ctr_pct(before_card_views, before_kw.get("query_freq", 0))
        after_share = safe_ctr_pct(after_card_views, after_kw.get("query_freq", 0))
        econ_info = _economics_info_for_key(ads_df, key)

        if after_days <= 0:
            comparison_status = "WAIT_AFTER_DATA" if event else comparison_status

        event_id = _clean_text_value(event.get("event_id", "")) if event else ""
        eff = effects.get(event_id, {}) if event_id else {}
        final_verdict = _clean_text_value(eff.get("final_verdict", event.get("final_verdict", "") if event else ""))
        target_action = _clean_text_value(eff.get("target_bid_action", ""))
        target_text = _clean_text_value(eff.get("target_bid_action_text", ""))
        recommended_next = eff.get("recommended_next_bid_rub", "") if eff else ""
        action = _clean_text_value(decision.get("action", ""))
        reason_code = _clean_text_value(decision.get("reason_code", ""))
        reason_text = _clean_text_value(decision.get("reason_text", ""))
        conclusion = _bid_compare_conclusion(
            action, reason_code, final_verdict, target_action,
            before_daily.get("gp_after_ads", float("nan")), after_daily.get("gp_after_ads", float("nan")),
            before_drr, after_drr,
            before_kw.get("avg_position", 0.0), after_kw.get("avg_position", 0.0),
            before_kw.get("visibility_pct", 0.0), after_kw.get("visibility_pct", 0.0),
        )

        rows.append({
            "campaign_id": decision.get("campaign_id", ""),
            "nm_id": nm_id,
            "supplier_article": supplier_article,
            "subject_norm": decision.get("subject_norm", ""),
            "placement": decision.get("placement", ""),
            "campaign_status": decision.get("campaign_status", ""),
            "economics_match_method": econ_info.get("economics_match_method", ""),
            "economics_product_group": econ_info.get("economics_product_group", ""),
            "economics_avg_price": econ_info.get("economics_avg_price", ""),
            "economics_commission_pct": econ_info.get("economics_commission_pct", ""),
            "economics_acquiring_pct": econ_info.get("economics_acquiring_pct", ""),
            "economics_vat_per_unit": econ_info.get("economics_vat_per_unit", ""),
            "economics_logistics_per_unit": econ_info.get("economics_logistics_per_unit", ""),
            "economics_cogs_per_unit": econ_info.get("economics_cogs_per_unit", ""),
            "comparison_status": comparison_status,
            "last_bid_change_date": event_date_text,
            "old_bid_rub": old_bid,
            "new_bid_rub": new_bid,
            "bid_change_reason_code": change_reason,
            "before_period": _date_window_label(before_start, before_end),
            "after_period": _date_window_label(after_start, after_end),
            "before_days": before_days,
            "after_days": after_days,
            "current_action": action,
            "current_reason_code": reason_code,
            "current_reason_text": reason_text,
            "postcheck_status": decision.get("postcheck_status", ""),
            "final_verdict": final_verdict,
            "target_bid_action": target_action,
            "target_bid_action_text": target_text,
            "recommended_next_bid_rub": recommended_next,
            "before_impressions": before_metrics.get("impressions", 0),
            "after_impressions": after_metrics.get("impressions", 0),
            "before_impressions_per_day": before_daily.get("impressions", 0),
            "after_impressions_per_day": after_daily.get("impressions", 0),
            "impressions_delta_pct": _safe_delta_pct(after_daily.get("impressions", 0), before_daily.get("impressions", 0)),
            "before_clicks": before_metrics.get("clicks", 0),
            "after_clicks": after_metrics.get("clicks", 0),
            "before_clicks_per_day": before_daily.get("clicks", 0),
            "after_clicks_per_day": after_daily.get("clicks", 0),
            "clicks_delta_pct": _safe_delta_pct(after_daily.get("clicks", 0), before_daily.get("clicks", 0)),
            "before_ctr_pct": before_ctr,
            "after_ctr_pct": after_ctr,
            "ctr_delta_pp": _safe_delta_pp(after_ctr, before_ctr),
            "before_orders_qty": before_metrics.get("orders", 0),
            "after_orders_qty": after_metrics.get("orders", 0),
            "before_orders_qty_per_day": before_daily.get("orders", 0),
            "after_orders_qty_per_day": after_daily.get("orders", 0),
            "orders_qty_delta_pct": _safe_delta_pct(after_daily.get("orders", 0), before_daily.get("orders", 0)),
            "before_orders_sum_rub": before_metrics.get("revenue", 0),
            "after_orders_sum_rub": after_metrics.get("revenue", 0),
            "before_orders_sum_rub_per_day": before_daily.get("revenue", 0),
            "after_orders_sum_rub_per_day": after_daily.get("revenue", 0),
            "orders_sum_delta_pct": _safe_delta_pct(after_daily.get("revenue", 0), before_daily.get("revenue", 0)),
            "before_gp_after_ads_rub": before_metrics.get("gp_after_ads", float("nan")),
            "after_gp_after_ads_rub": after_metrics.get("gp_after_ads", float("nan")),
            "before_gp_after_ads_rub_per_day": before_daily.get("gp_after_ads", float("nan")),
            "after_gp_after_ads_rub_per_day": after_daily.get("gp_after_ads", float("nan")),
            "gp_after_ads_delta_pct": _safe_delta_pct(after_daily.get("gp_after_ads", float("nan")), before_daily.get("gp_after_ads", float("nan"))),
            "before_ad_spend_rub": before_metrics.get("spend", 0),
            "after_ad_spend_rub": after_metrics.get("spend", 0),
            "before_ad_spend_rub_per_day": before_daily.get("spend", 0),
            "after_ad_spend_rub_per_day": after_daily.get("spend", 0),
            "ad_spend_delta_pct": _safe_delta_pct(after_daily.get("spend", 0), before_daily.get("spend", 0)),
            "before_drr_pct": before_drr,
            "after_drr_pct": after_drr,
            "drr_delta_pp": _safe_delta_pp(after_drr, before_drr),
            "core80_queries_count": max(before_kw.get("queries_count", 0), after_kw.get("queries_count", 0)),
            "before_core80_avg_position": before_kw.get("avg_position", 0),
            "after_core80_avg_position": after_kw.get("avg_position", 0),
            "core80_position_delta": _safe_delta_pp(after_kw.get("avg_position", 0), before_kw.get("avg_position", 0)),
            "before_core80_visibility_pct": before_kw.get("visibility_pct", 0),
            "after_core80_visibility_pct": after_kw.get("visibility_pct", 0),
            "core80_visibility_delta_pp": _safe_delta_pp(after_kw.get("visibility_pct", 0), before_kw.get("visibility_pct", 0)),
            "before_core80_query_freq": before_kw.get("query_freq", 0),
            "after_core80_query_freq": after_kw.get("query_freq", 0),
            "before_card_views": before_card_views,
            "after_card_views": after_card_views,
            "before_card_views_per_day": _safe_per_day_value(before_card_views, before_days),
            "after_card_views_per_day": _safe_per_day_value(after_card_views, after_days),
            "before_traffic_share_pct": before_share,
            "after_traffic_share_pct": after_share,
            "traffic_share_delta_pp": _safe_delta_pp(after_share, before_share),
            "diagnostic_conclusion": conclusion,
        })
    out = pd.DataFrame(rows)
    for col in BID_CAMPAIGN_COMPARE_COLUMNS:
        if col not in out.columns:
            out[col] = ""
    if not out.empty:
        after_cols = [c for c in out.columns if c.startswith("after_")] + [
            "impressions_delta_pct", "clicks_delta_pct", "ctr_delta_pp", "orders_qty_delta_pct",
            "orders_sum_delta_pct", "gp_after_ads_delta_pct", "ad_spend_delta_pct", "drr_delta_pp",
            "core80_position_delta", "core80_visibility_delta_pp", "traffic_share_delta_pp",
        ]
        no_after = pd.to_numeric(out.get("after_days", pd.Series(dtype=float)), errors="coerce").fillna(0).le(0)
        if no_after.any():
            # pandas в GitHub Actions стал строже: нельзя записывать строку "" в int/float-колонки.
            # Для строк без зрелого after-периода в Excel должны быть пустые ячейки, поэтому
            # переводим только эти output-колонки в object и затем проставляем пустые значения.
            after_cols = [c for c in after_cols if c in out.columns]
            for col in after_cols:
                out[col] = out[col].astype("object")
                out.loc[no_after, col] = ""
            out.loc[no_after, "after_period"] = ""
            out.loc[no_after, "diagnostic_conclusion"] = (
                "WAIT_AFTER_DATA: после изменения ставки ещё нет зрелых данных для сравнения; "
                "нули не используются как эффект изменения"
            )
        out = out.sort_values(["subject_norm", "supplier_article", "campaign_id", "placement"], ascending=[True, True, True, True])
    return out[BID_CAMPAIGN_COMPARE_COLUMNS]


def write_outputs(
    s3_client,
    config: Config,
    ctx: RunContext,
    decisions: pd.DataFrame,
    bid_history: pd.DataFrame,
    effect_df: pd.DataFrame,
    pause_candidates: pd.DataFrame,
    pause_history: pd.DataFrame,
    successful_changes: pd.DataFrame,
    api_log: pd.DataFrame,
    start_candidates: pd.DataFrame,
    applied_pauses: pd.DataFrame,
    applied_starts: pd.DataFrame,
    min_bids_df: pd.DataFrame,
    keyword_core_df: Optional[pd.DataFrame] = None,
    keyword_effects_df: Optional[pd.DataFrame] = None,
    price_decisions: Optional[pd.DataFrame] = None,
    price_history: Optional[pd.DataFrame] = None,
    price_effects_df: Optional[pd.DataFrame] = None,
    applied_price_changes: Optional[pd.DataFrame] = None,
    bid_ramp_monitor: Optional[pd.DataFrame] = None,
    one_campaign_experiment: Optional[pd.DataFrame] = None,
    rename_plan: Optional[pd.DataFrame] = None,
    bid_campaign_compare: Optional[pd.DataFrame] = None,
) -> Dict[str, Any]:
    summary = build_summary(ctx, decisions, successful_changes, pause_candidates, applied_pauses, start_candidates, applied_starts)
    summary["Ключевых фраз CORE_80"] = int(len(keyword_core_df[keyword_core_df["keyword_group"] == "CORE_80"])) if keyword_core_df is not None and not keyword_core_df.empty and "keyword_group" in keyword_core_df.columns else 0
    summary["Рекомендаций по цене"] = int(price_decisions["price_action"].isin(["Повысить цену", "Вернуть скидку"]).sum()) if price_decisions is not None and not price_decisions.empty and "price_action" in price_decisions.columns else 0
    summary["Ценовых тестов без воронки"] = int(price_decisions.get("funnel_missing", pd.Series(dtype=bool)).astype(bool).sum()) if price_decisions is not None and not price_decisions.empty and "funnel_missing" in price_decisions.columns else 0
    summary["Изменений цены отправлено"] = int(applied_price_changes["api_status"].astype(str).str.fullmatch(r"2\d\d", na=False).sum()) if applied_price_changes is not None and not applied_price_changes.empty and "api_status" in applied_price_changes.columns else 0
    summary["Скидка продавца по умолчанию"] = DEFAULT_SELLER_DISCOUNT_PCT
    summary["Минимальная скидка продавца"] = DEFAULT_MIN_SELLER_DISCOUNT_PCT
    summary["Кандидатов на разгон показов"] = int(len(bid_ramp_monitor)) if bid_ramp_monitor is not None else 0
    if bid_ramp_monitor is not None and not bid_ramp_monitor.empty and "ramp_mode_status" in bid_ramp_monitor.columns:
        ramp_counts = bid_ramp_monitor["ramp_mode_status"].astype(str).value_counts().to_dict()
        summary["Разгон: активен ждём D+7"] = int(ramp_counts.get("РАЗГОН_АКТИВЕН_ЖДЕМ_D7", 0))
        summary["Разгон: применён сейчас"] = int(ramp_counts.get("РАЗГОН_ПРИМЕНЕН_СЕЙЧАС_API_200", 0))
        summary["Разгон: заблокирован min WB"] = int(ramp_counts.get("РАЗГОН_ПОДХОДИТ_НО_БЛОК_MIN_WB", 0))
        summary["Разгон: подходит по метрикам, но ждёт другой post-check"] = int(ramp_counts.get("РАЗГОН_ПОДХОДИТ_НО_ЖДЕМ_ДРУГОЙ_POSTCHECK", 0))
    summary["Эксперимент 1РК групп"] = int(len(one_campaign_experiment)) if one_campaign_experiment is not None else 0
    summary["Порог паузы по показам"] = PAUSE_MIN_IMPRESSIONS
    summary["Окно проверки паузы, дней"] = PAUSE_ANALYSIS_DAYS
    summary["Правило автопаузы"] = "минимальная ставка WB + ДРР > лимита за 21 день + показы >= 10000"
    summary["Кисти паузим"] = "нет"
    summary["Кандидатов на переименование РК"] = 0
    summary["Переименовано РК"] = 0
    summary["Переименование РК"] = "отключено"
    summary["Строк сравнения РК 7д"] = int(len(bid_campaign_compare)) if bid_campaign_compare is not None else 0
    summary_df = pd.DataFrame([{"Показатель": k, "Значение": v} for k, v in summary.items()])

    sheets = {
        "Решения": decisions if decisions is not None else pd.DataFrame(columns=DECISION_COLUMNS),
        "История_изменений_ставок": bid_history if bid_history is not None else pd.DataFrame(columns=BID_HISTORY_COLUMNS),
        "Эффект_изменения_ставки": effect_df if effect_df is not None else pd.DataFrame(),
        "Оценка_изменения_ставок": effect_df if effect_df is not None else pd.DataFrame(),
        "Сравнение_РК_7дней": bid_campaign_compare if bid_campaign_compare is not None else pd.DataFrame(columns=BID_CAMPAIGN_COMPARE_COLUMNS),
        "Ключевые_фразы_80": keyword_core_df if keyword_core_df is not None else pd.DataFrame(columns=KEYWORD_POSITION_COLUMNS),
        "Эффект_по_ключевым_фразам": keyword_effects_df if keyword_effects_df is not None else pd.DataFrame(columns=KEYWORD_EFFECT_COLUMNS),
        "Разгон_показов": bid_ramp_monitor if bid_ramp_monitor is not None else pd.DataFrame(columns=BID_RAMP_MONITOR_COLUMNS),
        "Эксперимент_1РК_на_товар": one_campaign_experiment if one_campaign_experiment is not None else pd.DataFrame(columns=ONE_CAMPAIGN_EXPERIMENT_COLUMNS),
        "Решения_по_цене": price_decisions if price_decisions is not None else pd.DataFrame(columns=PRICE_DECISION_COLUMNS),
        "История_изменений_цен": price_history if price_history is not None else pd.DataFrame(columns=PRICE_HISTORY_COLUMNS),
        "Эффект_изменения_цен": price_effects_df if price_effects_df is not None else pd.DataFrame(),
        "Оценка_изменения_цен": price_effects_df if price_effects_df is not None else pd.DataFrame(),
        "Фактически_изменённые_цены": applied_price_changes if applied_price_changes is not None else pd.DataFrame(),
        "Кандидаты_на_паузу": pause_candidates if pause_candidates is not None else pd.DataFrame(columns=PAUSE_HISTORY_COLUMNS),
        "История_пауз": pause_history if pause_history is not None else pd.DataFrame(columns=PAUSE_HISTORY_COLUMNS),
        "Фактически_изменённые_ставки": successful_changes if successful_changes is not None else pd.DataFrame(),
        "Кандидаты_на_запуск": start_candidates if start_candidates is not None else pd.DataFrame(columns=PAUSE_HISTORY_COLUMNS),
        "Минимальные_ставки_WB": min_bids_df if min_bids_df is not None else pd.DataFrame(columns=MIN_BID_COLUMNS),
        "Переименование_РК": rename_plan if rename_plan is not None else pd.DataFrame(columns=RENAME_CAMPAIGN_COLUMNS),
        "Лог_API": api_log if api_log is not None else pd.DataFrame(),
        "Сводка": summary_df,
    }
    payload = dataframe_to_excel_bytes(sheets)
    output_key = PREVIEW_OUTPUT_KEY if ctx.mode == "preview" else RUN_OUTPUT_KEY
    upload_s3_bytes(s3_client, config.yc_bucket_name, output_key, payload, content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    upload_s3_bytes(
        s3_client,
        config.yc_bucket_name,
        SUMMARY_JSON_KEY,
        json.dumps(summary, ensure_ascii=False, indent=2).encode("utf-8"),
        content_type="application/json; charset=utf-8",
    )
    return summary


# =============================
# CLI и main
# =============================

def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="WB TOPFACE strict ad bid manager")
    subparsers = parser.add_subparsers(dest="command", required=True)

    preview = subparsers.add_parser("preview", help="Предпросмотр без отправки изменений ставок")
    preview.set_defaults(dry_run=False, apply_pause=False, apply_start=False, apply_price=False, apply_experiment=False, skip_price=False)

    run = subparsers.add_parser("run", help="Боевой расчёт и отправка изменений ставок")
    run.add_argument("--dry-run", action="store_true", help="Расчёт run-режима без реальных API-вызовов")
    run.add_argument("--apply-pause", action="store_true", help="Разрешить отправку pause для кандидатов")
    run.add_argument("--apply-start", action="store_true", help="Разрешить отправку start для кандидатов")
    run.add_argument("--apply-price", action="store_true", help="Разрешить отправку изменений скидки продавца через Discounts & Prices API")
    run.add_argument("--apply-experiment", action="store_true", help="Разрешить применение экспериментального блока 1 РК на товарную группу")
    run.add_argument("--skip-price", action="store_true", help="Не строить контур цен в этом запуске")
    return parser.parse_args(argv)


def print_summary(summary: Dict[str, Any]) -> None:
    print("=== Сводка запуска ===")
    for key, value in summary.items():
        print(f"{key}: {value}")


def main(argv: Optional[List[str]] = None) -> int:
    args = parse_args(argv)
    config = load_config()
    ctx = build_run_context(args)
    s3_client = make_s3_client(config)

    print(f"[{ctx.run_datetime:%Y-%m-%d %H:%M:%S}] Старт {SCRIPT_NAME}: версия={SCRIPT_VERSION}, режим={ctx.mode}, dry_run={ctx.dry_run}")
    print(f"Окна: база {ctx.base_start}..{ctx.base_end}; текущее {ctx.current_start}..{ctx.current_end}; mature_end={ctx.mature_end}")

    ads_df = load_ads_report(s3_client, config)
    print(f"Рекламный отчёт загружен: {len(ads_df):,} строк".replace(",", " "))

    keyword_df = load_keyword_positions(s3_client, config)
    keyword_core_df = classify_core_keywords(keyword_df, ctx)
    print(
        f"Мониторинг ключевых фраз: строк={len(keyword_df):,}; CORE_80={int((keyword_core_df.get('keyword_group', pd.Series(dtype=str)) == 'CORE_80').sum()) if not keyword_core_df.empty else 0}".replace(",", " "),
        flush=True,
    )

    funnel_df = load_funnel_report(s3_client, config)
    print(f"Воронка продаж загружена для price-check: {len(funnel_df):,} строк".replace(",", " "), flush=True)

    economics_df = load_economics_report(s3_client, config)
    ads_df = enrich_ads_with_estimated_gp(ads_df, economics_df)

    bid_history_raw = load_bid_history(s3_client, config)
    pause_history_raw = load_pause_history(s3_client, config)
    price_history_raw = load_price_history(s3_client, config)
    bid_history = filter_bid_history_managed_only(bid_history_raw, ads_df)
    pause_history = filter_pause_history_managed_only(pause_history_raw, ads_df)
    print(
        (
            f"История ставок: {len(bid_history):,} строк из {len(bid_history_raw):,} после фильтра 4 предметов; "
            f"история пауз: {len(pause_history):,} строк из {len(pause_history_raw):,} после фильтра 4 предметов; "
            f"история цен: {len(price_history_raw):,} строк"
        ).replace(",", " ")
    )

    bid_history, effect_df = evaluate_postchecks(ads_df, bid_history, ctx)
    keyword_effects_df = build_keyword_effects(bid_history, keyword_df, keyword_core_df, ctx)
    effect_df = enrich_effects_with_keyword_monitoring(effect_df, keyword_effects_df)

    # Price post-check D+2 по истории изменения скидки продавца.
    managed_nmids = set(ads_df["nm_id"].astype(str).unique()) if ads_df is not None and not ads_df.empty and "nm_id" in ads_df.columns else set()
    price_history_for_check = price_history_raw.copy() if price_history_raw is not None else pd.DataFrame(columns=PRICE_HISTORY_COLUMNS)
    for col in PRICE_HISTORY_COLUMNS:
        if col not in price_history_for_check.columns:
            price_history_for_check[col] = ""
    if managed_nmids and not price_history_for_check.empty:
        price_history_for_check = price_history_for_check[price_history_for_check["nm_id"].astype(str).isin(managed_nmids)].copy()
    price_history, price_effects_df = evaluate_price_postchecks(price_history_for_check, ads_df, funnel_df, ctx)

    pending_events = load_pending_events(bid_history, ctx)
    postcheck_results = latest_postcheck_results(bid_history)

    metrics_df = aggregate_campaign_metrics(ads_df, ctx)
    metrics_df = enrich_supplier_articles_from_economics(metrics_df, economics_df)
    print(f"Диагностика агрегации: строк метрик={len(metrics_df)}", flush=True)
    if not metrics_df.empty:
        print("Диагностика метрик по статусам: " + json.dumps(metrics_df.get("campaign_status", pd.Series(dtype=str)).map(str).value_counts().head(10).to_dict(), ensure_ascii=False), flush=True)
        print("Диагностика метрик по предметам: " + json.dumps(metrics_df.get("subject_norm", pd.Series(dtype=str)).map(str).value_counts().head(10).to_dict(), ensure_ascii=False), flush=True)
    decisions = build_decisions(metrics_df, pending_events, postcheck_results, ctx)
    decisions = enrich_supplier_articles_from_economics(decisions, economics_df)
    min_bids_df, min_bid_api_log = fetch_wb_min_bids_for_decisions(decisions, config, ctx)
    decisions = enrich_decisions_with_min_bids(decisions, min_bids_df)
    if not min_bids_df.empty:
        print(f"Диагностика WB min bids: получено строк={len(min_bids_df)}", flush=True)
    if not decisions.empty:
        print("Диагностика решений action: " + json.dumps(decisions["action"].value_counts().to_dict(), ensure_ascii=False), flush=True)
        print("Диагностика решений reason_code: " + json.dumps(decisions["reason_code"].value_counts().head(10).to_dict(), ensure_ascii=False), flush=True)

    bid_ramp_monitor = build_bid_ramp_monitor(decisions)
    if not bid_ramp_monitor.empty:
        print(f"Диагностика разгона показов до API: строк статуса={len(bid_ramp_monitor)}", flush=True)

    one_campaign_experiment = build_one_campaign_experiment(metrics_df, keyword_core_df)
    if not one_campaign_experiment.empty:
        print(f"Диагностика эксперимента 1РК: товарных групп={len(one_campaign_experiment)}", flush=True)

    price_list_api_log = pd.DataFrame()
    price_decisions = pd.DataFrame(columns=PRICE_DECISION_COLUMNS)
    applied_price_changes = pd.DataFrame()
    price_api_log = pd.DataFrame()
    rename_plan = pd.DataFrame(columns=RENAME_CAMPAIGN_COLUMNS)
    rename_api_log = pd.DataFrame()
    goods_prices = pd.DataFrame()
    if not getattr(args, "skip_price", False):
        goods_prices, price_list_api_log = fetch_current_goods_prices(config, ctx)
        funnel_current = aggregate_funnel_metrics(funnel_df, ctx.current_start, ctx.current_end)
        price_decisions = build_price_decisions(metrics_df, funnel_current, goods_prices, price_history, config, ctx)
        if not price_decisions.empty:
            print("Диагностика цен action: " + json.dumps(price_decisions["price_action"].value_counts().to_dict(), ensure_ascii=False), flush=True)
            print("Диагностика цен reason_code: " + json.dumps(price_decisions["reason_code"].value_counts().head(10).to_dict(), ensure_ascii=False), flush=True)
        apply_price_now = (ctx.mode == "run" and not ctx.dry_run) or bool(getattr(args, "apply_price", False))
        applied_price_changes, price_api_log = apply_price_changes(price_decisions, goods_prices, config, ctx, apply_price=apply_price_now)
        price_history = record_price_events(applied_price_changes, price_history, ctx)

    # Переименование РК отключено: кампании уже вернули к артикулам, больше не отправляем /adv/v0/rename.
    rename_plan = pd.DataFrame(columns=RENAME_CAMPAIGN_COLUMNS)
    rename_api_log = pd.DataFrame()
    decisions = enrich_supplier_articles_from_economics(decisions, economics_df)

    pause_candidates = build_pause_candidates(decisions, bid_history)

    if ctx.apply_experiment and one_campaign_experiment is not None and not one_campaign_experiment.empty:
        experiment_bid_decisions = build_experiment_bid_decisions(one_campaign_experiment, decisions)
        if not experiment_bid_decisions.empty:
            exp_keys = set(experiment_bid_decisions.apply(make_key, axis=1).tolist())
            decisions = decisions.loc[~decisions.apply(lambda r: make_key(r) in exp_keys, axis=1)].copy() if not decisions.empty else decisions
            decisions = pd.concat([decisions, experiment_bid_decisions], ignore_index=True, sort=False)
            print(f"Диагностика эксперимента 1РК: ставок к применению={len(experiment_bid_decisions)}", flush=True)
        experiment_pause_candidates = build_experiment_pause_candidates(one_campaign_experiment, metrics_df)
        if not experiment_pause_candidates.empty:
            pause_candidates = pd.concat([pause_candidates, experiment_pause_candidates], ignore_index=True, sort=False)
            print(f"Диагностика эксперимента 1РК: пауз кандидатов={len(experiment_pause_candidates)}", flush=True)

    successful_changes, bid_api_log = apply_bid_changes(decisions, config, ctx)
    bid_history = record_bid_events(successful_changes, bid_history, ctx)
    decisions = enrich_decisions_with_bid_api_status(decisions, successful_changes)
    bid_campaign_compare = build_campaign_7d_comparison(
        decisions=decisions,
        bid_history=bid_history,
        ads_df=ads_df,
        keyword_df=keyword_df,
        keyword_core_df=keyword_core_df,
        funnel_df=funnel_df,
        effect_df=effect_df,
        ctx=ctx,
    )
    if not bid_campaign_compare.empty:
        print(f"Диагностика сравнения РК 7д: строк={len(bid_campaign_compare)}", flush=True)
    bid_ramp_monitor = build_bid_ramp_monitor(decisions)
    if not bid_ramp_monitor.empty:
        print("Диагностика разгона статус: " + json.dumps(bid_ramp_monitor["ramp_mode_status"].value_counts().to_dict(), ensure_ascii=False), flush=True)

    applied_pauses, pause_api_log = apply_pause_actions(pause_candidates, config, ctx)
    if not applied_pauses.empty:
        pause_history = pd.concat([pause_history, applied_pauses[PAUSE_HISTORY_COLUMNS]], ignore_index=True, sort=False)

    normal_start_candidates = build_start_candidates(pause_history, ads_df, ctx)
    rollback_start_candidates = build_wrong_subject_pause_rollback_candidates(pause_history_raw, ads_df, ctx)
    if not rollback_start_candidates.empty:
        print(f"Диагностика rollback ошибочных пауз вне 4 предметов: кандидатов на запуск={len(rollback_start_candidates)}", flush=True)
    start_candidates = pd.concat(
        [df for df in [normal_start_candidates, rollback_start_candidates] if df is not None and not df.empty],
        ignore_index=True,
        sort=False,
    ) if any(df is not None and not df.empty for df in [normal_start_candidates, rollback_start_candidates]) else pd.DataFrame(columns=PAUSE_HISTORY_COLUMNS)
    applied_starts, start_api_log = apply_start_actions(start_candidates, config, ctx)
    if not applied_starts.empty:
        pause_history = pd.concat([pause_history, applied_starts[PAUSE_HISTORY_COLUMNS]], ignore_index=True, sort=False)

    all_api_log = pd.concat(
        [df for df in [min_bid_api_log, bid_api_log, pause_api_log, start_api_log, price_list_api_log, price_api_log, rename_api_log] if df is not None and not df.empty],
        ignore_index=True,
        sort=False,
    ) if any(df is not None and not df.empty for df in [min_bid_api_log, bid_api_log, pause_api_log, start_api_log, price_list_api_log, price_api_log, rename_api_log]) else pd.DataFrame()
    full_api_log = append_api_log_to_s3(s3_client, config, all_api_log)

    save_table_to_s3_excel(s3_client, config, BID_HISTORY_KEY, bid_history[BID_HISTORY_COLUMNS])
    save_table_to_s3_excel(s3_client, config, PAUSE_HISTORY_KEY, pause_history[PAUSE_HISTORY_COLUMNS] if not pause_history.empty else pd.DataFrame(columns=PAUSE_HISTORY_COLUMNS))
    save_table_to_s3_excel(s3_client, config, PRICE_HISTORY_KEY, price_history[PRICE_HISTORY_COLUMNS] if price_history is not None and not price_history.empty else pd.DataFrame(columns=PRICE_HISTORY_COLUMNS))

    summary = write_outputs(
        s3_client=s3_client,
        config=config,
        ctx=ctx,
        decisions=decisions,
        bid_history=bid_history[BID_HISTORY_COLUMNS],
        effect_df=effect_df,
        pause_candidates=pause_candidates,
        pause_history=pause_history[PAUSE_HISTORY_COLUMNS] if not pause_history.empty else pd.DataFrame(columns=PAUSE_HISTORY_COLUMNS),
        successful_changes=successful_changes,
        api_log=full_api_log,
        start_candidates=start_candidates,
        applied_pauses=applied_pauses,
        applied_starts=applied_starts,
        min_bids_df=min_bids_df,
        keyword_core_df=keyword_core_df,
        keyword_effects_df=keyword_effects_df,
        price_decisions=price_decisions,
        price_history=price_history[PRICE_HISTORY_COLUMNS] if price_history is not None and not price_history.empty else pd.DataFrame(columns=PRICE_HISTORY_COLUMNS),
        price_effects_df=price_effects_df,
        applied_price_changes=applied_price_changes,
        bid_ramp_monitor=bid_ramp_monitor,
        one_campaign_experiment=one_campaign_experiment,
        rename_plan=rename_plan,
        bid_campaign_compare=bid_campaign_compare,
    )
    print_summary(summary)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ОШИБКА: {exc}", file=sys.stderr)
        raise
