#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FIX46_CORE_RAMP_PAUSE_20260611
WB Ads Manager — decision engine for ставки, CORE-трафик, разгон, паузы и возвраты.

Что изменено по новой логике:
- ABC-рентабельность полностью исключена из принятия решений.
- Финансовая оценка использует зрелое окно: последние 7 дней, исключая последние 3 дня лага.
  Например as_of=2026-06-11 => current=2026-06-01..2026-06-07, base=2026-05-25..2026-05-31.
- Search/CPC: ставка целая, шаг 1 рубль.
- Combined/CPM/полки: минимум 80 рублей, шаг 6 рублей. Значения CPC из статистики не считаются ставкой.
- Предельная ставка считается через прогнозный ДРР 16% и среднюю цену продажи finishedPrice из Orders.
- CORE оценивается не только кликами/позициями, но и кликами на 1 заказ и CPO запроса.
- Разгон: не больше 1 CPC и 1 CPM на товарный блок одновременно, до 5000 показов, затем пауза на дозревание.
- Пауза: не просто срез расхода, а pause_and_reallocate внутри product_root+placement.

Скрипт не отправляет API-вызовы сам. Он формирует решения и payload-preview.
Дальше текущий боевой runner должен отправлять только строки action in {raise, lower, pause, start}.

Пример:
python assistant_wb_ads_manager__FIX46_CORE_RAMP_PAUSE_20260611.py \
  --ads "Реклама_2026-W24.xlsx" \
  --orders "Заказы_2026-W24.xlsx" \
  --previous-output "Итог_последнего_запуска.xlsx" \
  --bid-history "История_ставок.xlsx" \
  --pause-history "История_пауз.xlsx" \
  --as-of 2026-06-11 \
  --out "wb_ads_decisions_FIX46.xlsx"
"""

from __future__ import annotations

import argparse
import math
import os
import re
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import numpy as np
import pandas as pd

VERSION = "FIX46_CORE_RAMP_PAUSE_20260611"

# -------------------------
# Business constants
# -------------------------
DRR_RAISE_GATE_PCT = 10.0          # ДРР < 10 => повышение можно рассматривать, но не автоматически
DRR_PAUSE_LIMIT_PCT = 15.0         # 14д + 10000 показов + ДРР > 15 => pause candidate
DRR_FORECAST_CAP_PCT = 16.0        # потолок прогнозного ДРР для предельной ставки
SEARCH_MIN_BID_RUB = 4             # WB min can be fetched externally; safe default for CPC
SEARCH_STEP_RUB = 1
COMBINED_MIN_BID_RUB = 80
COMBINED_STEP_RUB = 6
RAMP_TARGET_IMPRESSIONS = 5000
RAMP_MAX_ACTIVE_CPC_PER_BLOCK = 1
RAMP_MAX_ACTIVE_CPM_PER_BLOCK = 1
NEW_NO_PAUSE_DAYS = 14
MATURE_WINDOW_DAYS = 7
DATA_LAG_DAYS = 3
PAUSE_WINDOW_DAYS = 14
POST_PAUSE_CHECK_DAYS = 7
FLAGSHIP_MAX_PER_PRODUCT = 3
FLAGSHIP_TARGET_POSITION = 10
FLAGSHIP_BEST_POSITION = 3

TARGET_SUBJECTS = {
    "кисти косметические": "Кисти косметические",
    "помады": "Помады",
    "блески": "Блески",
    "карандаши": "Косметические карандаши",
    "косметические карандаши": "Косметические карандаши",
}

# -------------------------
# Column aliases
# -------------------------
ALIASES = {
    "campaign_id": ["campaign_id", "ID кампании", "id кампании", "advertId", "advert_id"],
    "nm_id": ["nm_id", "nmId", "Артикул WB", "Номенклатура WB", "nmID"],
    "supplier_article": ["supplier_article", "supplierArticle", "Артикул продавца", "Артикул"],
    "subject": ["subject", "subject_norm", "Название предмета", "Предмет", "category", "Категория"],
    "day": ["day", "Дата", "date", "dt"],
    "campaign_name": ["Название", "name", "campaign_name", "Кампания"],
    "campaign_status": ["Статус", "campaign_status", "status"],
    "payment_type": ["Тип оплаты", "payment_type", "placement", "Тип кампании"],
    "bid_type": ["Тип ставки", "bid_type"],
    "search_bid": ["Ставка в поиске (руб)", "search_bid", "bid_search", "Ставка в поиске"],
    "reco_bid": ["Ставка в рекомендациях (руб)", "reco_bid", "bid_recommendations", "Ставка в рекомендациях"],
    "impressions": ["Показы", "impressions", "shows"],
    "clicks": ["Клики", "clicks", "Переходы", "transitions", "clicks_to_card"],
    "ctr_pct": ["CTR", "ctr", "CTR, %", "ctr_pct"],
    "cpc": ["CPC", "cpc"],
    "orders": ["Заказы", "orders", "ad_orders", "keyword_orders"],
    "spend": ["Расход", "Затраты", "spend", "ad_spend"],
    "order_sum": ["Сумма заказов", "revenue", "order_sum", "ad_order_sum"],
    "drr_pct": ["ДРР", "Доля затрат", "drr", "campaign_drr_pct", "Доля затрат, %"],
    "frequency": ["query_freq", "Частотность", "frequency", "freq"],
    "query": ["query_text", "query_text_norm", "Поисковый запрос", "search_query", "Запрос"],
    "median_position": ["median_position", "Медианная позиция", "Позиция", "position"],
    "visibility_pct": ["visibility_pct", "Видимость", "visibility"],
    "finished_price": ["finishedPrice", "finished_price", "Цена продажи", "finished price"],
    "price_with_disc": ["priceWithDisc", "price_with_disc"],
    "is_cancel": ["isCancel", "is_cancel", "Отменено"],
}


def norm_col(x: Any) -> str:
    return re.sub(r"\s+", " ", str(x).strip())


def find_col(df: pd.DataFrame, logical: str) -> Optional[str]:
    aliases = ALIASES.get(logical, [logical])
    lower_map = {norm_col(c).lower(): c for c in df.columns}
    for a in aliases:
        key = norm_col(a).lower()
        if key in lower_map:
            return lower_map[key]
    # fuzzy normalized
    normalized = {re.sub(r"[^a-zа-я0-9]+", "", norm_col(c).lower()): c for c in df.columns}
    for a in aliases:
        key = re.sub(r"[^a-zа-я0-9]+", "", norm_col(a).lower())
        if key in normalized:
            return normalized[key]
    return None


def s(df: pd.DataFrame, logical: str, default: Any = np.nan) -> pd.Series:
    col = find_col(df, logical)
    if col is None:
        return pd.Series([default] * len(df), index=df.index)
    return df[col]


def to_num(x: pd.Series) -> pd.Series:
    if x is None:
        return pd.Series(dtype=float)
    return pd.to_numeric(
        x.astype(str)
        .str.replace("\u00a0", " ", regex=False)
        .str.replace("₽", "", regex=False)
        .str.replace("%", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.replace(",", ".", regex=False),
        errors="coerce",
    )


def to_date(x: pd.Series) -> pd.Series:
    return pd.to_datetime(x, errors="coerce", dayfirst=True).dt.normalize()


def clean_article(v: Any) -> str:
    if pd.isna(v):
        return ""
    return str(v).strip().replace("\\", "/")


def product_root(article: Any) -> str:
    a = clean_article(article)
    if not a:
        return ""
    # PT901.F26 => 901, 155/12 => 155, 617/1 => 617
    m = re.search(r"(\d{3})", a)
    return m.group(1) if m else a.split("/")[0].split(".")[0]


def canon_subject(v: Any) -> str:
    if pd.isna(v):
        return ""
    x = str(v).strip().lower()
    x = x.replace("ё", "е")
    if "кист" in x:
        return "Кисти косметические"
    if "помад" in x:
        return "Помады"
    if "блес" in x:
        return "Блески"
    if "карандаш" in x:
        return "Косметические карандаши"
    return str(v).strip()


def placement_from_payment(v: Any, bid_type: Any = "") -> str:
    x = f"{v} {bid_type}".lower()
    if "cpc" in x or "поиск" in x:
        return "search"
    if "cpm" in x or "единая" in x or "пол" in x or "combined" in x:
        return "combined"
    # default from old code terms
    if "search" in x:
        return "search"
    return "combined"


def safe_div(a: float, b: float, default: float = np.nan) -> float:
    try:
        if b is None or pd.isna(b) or float(b) == 0:
            return default
        return float(a) / float(b)
    except Exception:
        return default


def pct(a: float, b: float) -> float:
    return safe_div(a, b, np.nan) * 100.0


def floor_to_combined_grid(value: float) -> Optional[int]:
    if pd.isna(value):
        return None
    if value < COMBINED_MIN_BID_RUB:
        return COMBINED_MIN_BID_RUB
    steps = math.floor((value - COMBINED_MIN_BID_RUB) / COMBINED_STEP_RUB)
    return int(COMBINED_MIN_BID_RUB + steps * COMBINED_STEP_RUB)


def ceil_to_combined_grid(value: float) -> Optional[int]:
    if pd.isna(value):
        return None
    if value <= COMBINED_MIN_BID_RUB:
        return COMBINED_MIN_BID_RUB
    steps = math.ceil((value - COMBINED_MIN_BID_RUB) / COMBINED_STEP_RUB)
    return int(COMBINED_MIN_BID_RUB + steps * COMBINED_STEP_RUB)


def date_windows(as_of: pd.Timestamp) -> Dict[str, pd.Timestamp]:
    as_of = pd.Timestamp(as_of).normalize()
    current_end = as_of - pd.Timedelta(days=DATA_LAG_DAYS + 1)  # 11 => 7
    current_start = current_end - pd.Timedelta(days=MATURE_WINDOW_DAYS - 1)
    base_end = current_start - pd.Timedelta(days=1)
    base_start = base_end - pd.Timedelta(days=MATURE_WINDOW_DAYS - 1)
    pause_end = as_of - pd.Timedelta(days=1)
    pause_start = pause_end - pd.Timedelta(days=PAUSE_WINDOW_DAYS - 1)
    return {
        "as_of": as_of,
        "current_start": current_start,
        "current_end": current_end,
        "base_start": base_start,
        "base_end": base_end,
        "pause_start": pause_start,
        "pause_end": pause_end,
    }


def in_window(df: pd.DataFrame, start: pd.Timestamp, end: pd.Timestamp, day_col: str = "day") -> pd.DataFrame:
    if df.empty or day_col not in df.columns:
        return df.iloc[0:0].copy()
    day = pd.to_datetime(df[day_col], errors="coerce").dt.normalize()
    return df[(day >= start) & (day <= end)].copy()


def read_sheet(path: Optional[str], preferred: Sequence[str] = ()) -> pd.DataFrame:
    if not path or not os.path.exists(path):
        return pd.DataFrame()
    xl = pd.ExcelFile(path)
    for p in preferred:
        for sh in xl.sheet_names:
            if sh.strip().lower() == p.strip().lower():
                return pd.read_excel(path, sheet_name=sh)
    # if preferred not found, first sheet
    return pd.read_excel(path, sheet_name=xl.sheet_names[0])


def expand_input_paths(spec: Optional[str]) -> List[str]:
    """Accepts one path, comma/semicolon separated paths, or glob masks."""
    if not spec:
        return []
    parts = re.split(r"[;,]", str(spec))
    paths: List[str] = []
    for part in parts:
        p = part.strip().strip('"').strip("'")
        if not p:
            continue
        if any(ch in p for ch in ["*", "?", "["]):
            paths.extend([str(x) for x in sorted(Path().glob(p))])
        else:
            paths.append(p)
    # keep existing only, preserve order and dedupe
    seen = set()
    out = []
    for p in paths:
        if p not in seen and os.path.exists(p):
            out.append(p)
            seen.add(p)
    return out


def read_previous_sheet(path: Optional[str], sheet_name: str) -> pd.DataFrame:
    if not path or not os.path.exists(path):
        return pd.DataFrame()
    try:
        xl = pd.ExcelFile(path)
        for sh in xl.sheet_names:
            if sh.strip().lower() == sheet_name.strip().lower():
                return pd.read_excel(path, sheet_name=sh)
    except Exception:
        pass
    return pd.DataFrame()


# -------------------------
# Normalizers
# -------------------------

def load_ads_daily(path: str) -> Tuple[pd.DataFrame, pd.DataFrame]:
    paths = expand_input_paths(path)
    if not paths:
        raise RuntimeError(f"Не найден ни один файл рекламы по --ads={path}")
    daily_frames: List[pd.DataFrame] = []
    campaign_frames: List[pd.DataFrame] = []
    for pth in paths:
        daily_raw = read_sheet(pth, ["Статистика_Ежедневно", "daily", "Статистика ежедневно"])
        campaigns_raw = read_sheet(pth, ["Список_кампаний", "campaigns", "Кампании"])
        if daily_raw.empty:
            continue
        daily = pd.DataFrame({
            "campaign_id": to_num(s(daily_raw, "campaign_id")).astype("Int64"),
            "nm_id": to_num(s(daily_raw, "nm_id")).astype("Int64"),
            "campaign_name": s(daily_raw, "campaign_name", "").astype(str),
            "subject_norm": s(daily_raw, "subject", "").map(canon_subject),
            "day": to_date(s(daily_raw, "day")),
            "impressions": to_num(s(daily_raw, "impressions")).fillna(0.0),
            "clicks": to_num(s(daily_raw, "clicks")).fillna(0.0),
            "ctr_pct_raw": to_num(s(daily_raw, "ctr_pct")),
            "cpc_stat": to_num(s(daily_raw, "cpc")),
            "orders": to_num(s(daily_raw, "orders")).fillna(0.0),
            "spend": to_num(s(daily_raw, "spend")).fillna(0.0),
            "order_sum": to_num(s(daily_raw, "order_sum")).fillna(0.0),
            "drr_pct_raw": to_num(s(daily_raw, "drr_pct")),
            "source_file": Path(pth).name,
        })
        daily = daily[daily["campaign_id"].notna()].copy()
        if not daily.empty:
            daily["campaign_id"] = daily["campaign_id"].astype(int)
            daily["nm_id"] = daily["nm_id"].astype("Int64")
            daily_frames.append(daily)
        if not campaigns_raw.empty:
            c = normalize_campaigns(campaigns_raw)
            c["source_file"] = Path(pth).name
            campaign_frames.append(c)
    if not daily_frames:
        raise RuntimeError(f"Не найден лист статистики рекламы в файлах: {paths}")
    out = pd.concat(daily_frames, ignore_index=True)
    # Dedupe exact day+campaign rows across overlapping exports; keep latest file in input order.
    out["_source_order"] = out["source_file"].map({Path(p).name: i for i, p in enumerate(paths)})
    out = out.sort_values("_source_order").drop_duplicates(["campaign_id", "day"], keep="last").drop(columns=["_source_order"])

    camp = pd.concat(campaign_frames, ignore_index=True) if campaign_frames else pd.DataFrame()
    if not camp.empty:
        camp["_source_order"] = camp["source_file"].map({Path(p).name: i for i, p in enumerate(paths)})
        camp = camp.sort_values("_source_order").drop_duplicates("campaign_id", keep="last").drop(columns=["_source_order"])
        out = out.merge(camp[["campaign_id", "supplier_article", "product_root", "placement", "campaign_status", "real_bid_rub", "search_bid", "reco_bid"]], on="campaign_id", how="left")
    else:
        out["supplier_article"] = ""
        out["product_root"] = ""
        out["placement"] = ""
        out["campaign_status"] = ""
        out["real_bid_rub"] = np.nan
        out["search_bid"] = np.nan
        out["reco_bid"] = np.nan
    out["placement"] = out["placement"].fillna("").replace("", "search")
    return out, camp


def normalize_campaigns(df: pd.DataFrame) -> pd.DataFrame:
    out = pd.DataFrame({
        "campaign_id": to_num(s(df, "campaign_id")).astype("Int64"),
        "campaign_name": s(df, "campaign_name", "").astype(str),
        "campaign_status": s(df, "campaign_status", "").astype(str),
        "payment_type": s(df, "payment_type", "").astype(str),
        "bid_type": s(df, "bid_type", "").astype(str),
        "search_bid": to_num(s(df, "search_bid")),
        "reco_bid": to_num(s(df, "reco_bid")),
        "subject_norm": s(df, "subject", "").map(canon_subject),
        "nm_id": to_num(s(df, "nm_id")).astype("Int64"),
    })
    out = out[out["campaign_id"].notna()].copy()
    out["campaign_id"] = out["campaign_id"].astype(int)
    out["placement"] = [placement_from_payment(a, b) for a, b in zip(out["payment_type"], out["bid_type"])]
    out["supplier_article"] = out["campaign_name"].map(extract_article_from_campaign_name)
    out["product_root"] = out["supplier_article"].map(product_root)

    # Real bid: search uses search_bid, combined/CPM uses max of available bids and never below 80.
    real = []
    for _, r in out.iterrows():
        if r["placement"] == "search":
            val = r["search_bid"]
            if pd.isna(val):
                val = r["reco_bid"]
            real.append(float(round(val)) if pd.notna(val) else np.nan)
        else:
            vals = [v for v in [r["search_bid"], r["reco_bid"]] if pd.notna(v)]
            val = max(vals) if vals else np.nan
            if pd.notna(val):
                val = max(float(val), COMBINED_MIN_BID_RUB)
            real.append(val)
    out["real_bid_rub"] = real
    out["is_active"] = out["campaign_status"].str.lower().str.contains("актив|active", na=False)
    return out


def extract_article_from_campaign_name(name: Any) -> str:
    if pd.isna(name):
        return ""
    text = str(name).strip()
    # campaign names like 155/15, PT901.F26, 156/16
    m = re.search(r"(PT\s*\d{3}[\.\-/]?[A-Z]?\d*|\d{3}\s*/\s*\d+|\d{3}[\./][A-Za-z0-9]+)", text, flags=re.I)
    if m:
        return m.group(1).replace(" ", "")
    # fallback root only
    m = re.search(r"\b(\d{3})\b", text)
    return m.group(1) if m else ""


def load_orders(path: Optional[str]) -> pd.DataFrame:
    paths = expand_input_paths(path)
    frames: List[pd.DataFrame] = []
    for pth in paths:
        df = read_sheet(pth, ["Заказы", "orders"])
        if df.empty:
            continue
        out = pd.DataFrame({
            "day": to_date(s(df, "day")),
            "supplier_article": s(df, "supplier_article", "").map(clean_article),
            "nm_id": to_num(s(df, "nm_id")).astype("Int64"),
            "subject_norm": s(df, "subject", "").map(canon_subject),
            "finished_price": to_num(s(df, "finished_price")),
            "price_with_disc": to_num(s(df, "price_with_disc")),
            "is_cancel": s(df, "is_cancel", False),
            "source_file": Path(pth).name,
        })
        out["product_root"] = out["supplier_article"].map(product_root)
        cancel = out["is_cancel"].astype(str).str.lower().isin(["true", "1", "да", "yes"])
        out = out[~cancel].copy()
        frames.append(out)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def load_bid_history(path: Optional[str]) -> pd.DataFrame:
    df = read_sheet(path, ["Лист1", "История_изменений_ставок", "История ставок"])
    if df.empty:
        return pd.DataFrame(columns=["campaign_id", "event_date", "old_bid_rub", "new_bid_rub", "direction", "reason_code"])
    out = pd.DataFrame({
        "campaign_id": to_num(s(df, "campaign_id")).astype("Int64"),
        "event_date": to_date(s(df, "day", np.nan) if find_col(df, "day") else df.get("event_date", pd.Series([np.nan]*len(df)))),
        "old_bid_rub": to_num(df.get("old_bid_rub", pd.Series([np.nan]*len(df)))),
        "new_bid_rub": to_num(df.get("new_bid_rub", pd.Series([np.nan]*len(df)))),
        "direction": df.get("direction", pd.Series([""]*len(df))).astype(str),
        "reason_code": df.get("reason_code", pd.Series([""]*len(df))).astype(str),
    })
    out = out[out["campaign_id"].notna()].copy()
    out["campaign_id"] = out["campaign_id"].astype(int)
    return out


def load_pause_history(path: Optional[str]) -> pd.DataFrame:
    df = read_sheet(path, ["Лист1", "История_пауз", "История пауз"])
    if df.empty:
        return pd.DataFrame(columns=["campaign_id", "pause_date", "status"])
    out = pd.DataFrame({
        "campaign_id": to_num(s(df, "campaign_id")).astype("Int64"),
        "pause_date": to_date(df.get("pause_date", pd.Series([np.nan]*len(df)))),
        "status": df.get("status", pd.Series([""]*len(df))).astype(str),
        "reason_code": df.get("reason_code", pd.Series([""]*len(df))).astype(str),
        "api_status": df.get("api_status", pd.Series([""]*len(df))).astype(str),
    })
    out = out[out["campaign_id"].notna()].copy()
    out["campaign_id"] = out["campaign_id"].astype(int)
    return out


def load_keywords_from_previous(path: Optional[str]) -> pd.DataFrame:
    """Loads CORE key phrases from previous output if available."""
    df = read_previous_sheet(path, "Ключевые_фразы_80")
    if df.empty:
        return pd.DataFrame()
    out = pd.DataFrame({
        "nm_id": to_num(s(df, "nm_id")).astype("Int64"),
        "supplier_article": s(df, "supplier_article", "").map(clean_article),
        "subject_norm": s(df, "subject", "").map(canon_subject),
        "query_text": s(df, "query", "").astype(str),
        "orders": to_num(s(df, "orders")).fillna(0.0),
        "frequency": to_num(s(df, "frequency")).fillna(0.0),
        "median_position": to_num(s(df, "median_position")),
        "visibility_pct": to_num(s(df, "visibility_pct")),
        "clicks": to_num(s(df, "clicks")).fillna(0.0),
        "keyword_group": df.get("keyword_group", pd.Series([""]*len(df))).astype(str),
        "orders_share": to_num(df.get("orders_share", pd.Series([np.nan]*len(df)))),
        "orders_cum_share": to_num(df.get("orders_cum_share", pd.Series([np.nan]*len(df)))),
    })
    out["product_root"] = out["supplier_article"].map(product_root)
    out = out[out["query_text"].ne("")].copy()
    # The previous output may contain tens of thousands of tail queries. For decision logic we need
    # CORE only; tail queries are used later only as diagnostics. This keeps the run fast/stable.
    if "keyword_group" in out.columns and out["keyword_group"].astype(str).str.len().gt(0).any():
        core_mask = out["keyword_group"].astype(str).str.upper().str.contains("CORE")
        if core_mask.any():
            out = out[core_mask].copy()
    elif "orders_cum_share" in out.columns:
        out = out[out["orders_cum_share"].fillna(999) <= 0.85].copy()
    # Safety cap: no more than top 30 queries per product by orders/frequency.
    if not out.empty:
        out = (out.sort_values(["product_root", "orders", "frequency"], ascending=[True, False, False])
                 .groupby("product_root", as_index=False)
                 .head(30)
                 .reset_index(drop=True))
    return out


# -------------------------
# Metrics
# -------------------------

def campaign_window_metrics(ads: pd.DataFrame, start: pd.Timestamp, end: pd.Timestamp, suffix: str) -> pd.DataFrame:
    w = in_window(ads, start, end)
    if w.empty:
        return pd.DataFrame(columns=["campaign_id"])
    grp = w.groupby("campaign_id", as_index=False).agg(
        impressions=("impressions", "sum"),
        clicks=("clicks", "sum"),
        orders=("orders", "sum"),
        spend=("spend", "sum"),
        order_sum=("order_sum", "sum"),
    )
    grp[f"ctr_pct_{suffix}"] = np.where(grp["impressions"] > 0, grp["clicks"] / grp["impressions"] * 100, np.nan)
    grp[f"cpc_{suffix}"] = np.where(grp["clicks"] > 0, grp["spend"] / grp["clicks"], np.nan)
    grp[f"cpo_{suffix}"] = np.where(grp["orders"] > 0, grp["spend"] / grp["orders"], np.nan)
    grp[f"clicks_per_order_{suffix}"] = np.where(grp["orders"] > 0, grp["clicks"] / grp["orders"], np.nan)
    grp[f"impressions_per_order_{suffix}"] = np.where(grp["orders"] > 0, grp["impressions"] / grp["orders"], np.nan)
    grp[f"drr_pct_{suffix}"] = np.where(grp["order_sum"] > 0, grp["spend"] / grp["order_sum"] * 100, np.nan)
    return grp.rename(columns={
        "impressions": f"impressions_{suffix}",
        "clicks": f"clicks_{suffix}",
        "orders": f"orders_{suffix}",
        "spend": f"spend_{suffix}",
        "order_sum": f"order_sum_{suffix}",
    })


def avg_price_by_product(orders: pd.DataFrame, start: pd.Timestamp, end: pd.Timestamp) -> pd.DataFrame:
    if orders.empty:
        return pd.DataFrame(columns=["product_root", "nm_id", "avg_finished_price"])
    w = in_window(orders, start, end)
    if w.empty:
        return pd.DataFrame(columns=["product_root", "nm_id", "avg_finished_price"])
    by_nm = w.groupby(["product_root", "nm_id"], dropna=False, as_index=False).agg(
        avg_finished_price=("finished_price", "mean"),
        orders_rows=("finished_price", "size"),
    )
    by_root = w.groupby("product_root", as_index=False).agg(
        avg_finished_price_root=("finished_price", "mean"),
        orders_rows_root=("finished_price", "size"),
    )
    return by_nm.merge(by_root, on="product_root", how="left")


def build_campaign_base(ads: pd.DataFrame, campaigns: pd.DataFrame, orders: pd.DataFrame, bid_history: pd.DataFrame, windows: Dict[str, pd.Timestamp]) -> pd.DataFrame:
    current = campaign_window_metrics(ads, windows["current_start"], windows["current_end"], "cur")
    base = campaign_window_metrics(ads, windows["base_start"], windows["base_end"], "base")
    p14 = campaign_window_metrics(ads, windows["pause_start"], windows["pause_end"], "14d")

    if campaigns.empty:
        ids = pd.concat([current[["campaign_id"]], base[["campaign_id"]], p14[["campaign_id"]]], ignore_index=True).drop_duplicates()
        df = ids
    else:
        df = campaigns.copy()

    for part in [current, base, p14]:
        df = df.merge(part, on="campaign_id", how="left")

    # First seen by ads stats/history.
    first_seen_ads = ads.groupby("campaign_id", as_index=False).agg(first_seen=("day", "min")) if not ads.empty else pd.DataFrame(columns=["campaign_id", "first_seen"])
    df = df.merge(first_seen_ads, on="campaign_id", how="left")

    if not bid_history.empty:
        last = bid_history.dropna(subset=["event_date"]).sort_values("event_date").groupby("campaign_id", as_index=False).tail(1)
        last = last[["campaign_id", "event_date", "old_bid_rub", "new_bid_rub", "direction", "reason_code"]].rename(columns={
            "event_date": "last_bid_change_date",
            "old_bid_rub": "last_old_bid_rub",
            "new_bid_rub": "last_new_bid_rub",
            "direction": "last_bid_direction",
            "reason_code": "last_bid_reason_code",
        })
        df = df.merge(last, on="campaign_id", how="left")
    else:
        df["last_bid_change_date"] = pd.NaT
        df["last_old_bid_rub"] = np.nan
        df["last_new_bid_rub"] = np.nan
        df["last_bid_direction"] = ""
        df["last_bid_reason_code"] = ""

    # Avg prices.
    prices = avg_price_by_product(orders, windows["current_start"], windows["current_end"])
    if not prices.empty and "nm_id" in df.columns:
        df = df.merge(prices[["product_root", "nm_id", "avg_finished_price", "avg_finished_price_root"]], on=["product_root", "nm_id"], how="left")
        df["avg_finished_price"] = df["avg_finished_price"].fillna(df["avg_finished_price_root"])
    else:
        df["avg_finished_price"] = np.nan

    # Basic defaults and expected metric columns even when a window has no rows.
    metric_cols = []
    for suf in ["cur", "base", "14d"]:
        metric_cols += [
            f"impressions_{suf}", f"clicks_{suf}", f"orders_{suf}", f"spend_{suf}", f"order_sum_{suf}",
            f"ctr_pct_{suf}", f"cpc_{suf}", f"cpo_{suf}", f"clicks_per_order_{suf}",
            f"impressions_per_order_{suf}", f"drr_pct_{suf}",
        ]
    for col in metric_cols:
        if col not in df.columns:
            df[col] = np.nan
    for col in [c for c in df.columns if any(x in c for x in ["impressions_", "clicks_", "orders_", "spend_", "order_sum_"])]:
        df[col] = to_num(df[col]).fillna(0.0)

    # Derived statuses.
    df["days_since_first_seen"] = (windows["as_of"] - pd.to_datetime(df["first_seen"], errors="coerce")).dt.days
    df["is_new"] = df["days_since_first_seen"].fillna(9999) < NEW_NO_PAUSE_DAYS
    df["days_since_last_bid_change"] = (windows["as_of"] - pd.to_datetime(df["last_bid_change_date"], errors="coerce")).dt.days
    df["recent_bid_change"] = df["days_since_last_bid_change"].between(0, 1, inclusive="both")
    return df


def compute_bid_caps(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    caps = []
    reasons = []
    forecast_cpo_next = []
    forecast_drr_next = []

    # category competition cap by subject+placement.
    comp = out.groupby(["subject_norm", "placement"], dropna=False).agg(category_max_bid=("real_bid_rub", "max"), category_median_bid=("real_bid_rub", "median")).reset_index()
    out = out.merge(comp, on=["subject_norm", "placement"], how="left")

    for _, r in out.iterrows():
        avg_price = r.get("avg_finished_price", np.nan)
        placement = r.get("placement", "")
        current_bid = r.get("real_bid_rub", np.nan)
        cat_cap = r.get("category_max_bid", np.nan)
        if placement == "search":
            clicks_per_order = r.get("clicks_per_order_cur", np.nan)
            if pd.notna(avg_price) and avg_price > 0 and pd.notna(clicks_per_order) and clicks_per_order > 0:
                target_cpo = avg_price * DRR_FORECAST_CAP_PCT / 100.0
                max_bid = math.floor(target_cpo / clicks_per_order)
                max_bid = max(max_bid, SEARCH_MIN_BID_RUB)
                reason = "CAP_BY_16_DRR_CLICKS_PER_ORDER"
            else:
                # Not enough orders: use category competition +2, but mark weak cap.
                max_bid = math.floor((cat_cap if pd.notna(cat_cap) else current_bid if pd.notna(current_bid) else SEARCH_MIN_BID_RUB) + 2)
                reason = "CAP_BY_CATEGORY_MAX_PLUS_2_NO_STABLE_CPO"
            next_bid = (round(current_bid) + SEARCH_STEP_RUB) if pd.notna(current_bid) else np.nan
            cpo_next = (r.get("clicks_per_order_cur", np.nan) * next_bid) if pd.notna(next_bid) and pd.notna(r.get("clicks_per_order_cur", np.nan)) else np.nan
            drr_next = pct(cpo_next, avg_price) if pd.notna(cpo_next) and pd.notna(avg_price) and avg_price > 0 else np.nan
        else:
            impressions_per_order = r.get("impressions_per_order_cur", np.nan)
            if pd.notna(avg_price) and avg_price > 0 and pd.notna(impressions_per_order) and impressions_per_order > 0:
                target_cpo = avg_price * DRR_FORECAST_CAP_PCT / 100.0
                raw_cpm = target_cpo * 1000.0 / impressions_per_order
                max_bid = floor_to_combined_grid(raw_cpm)
                reason = "CAP_BY_16_DRR_IMPRESSIONS_PER_ORDER"
            else:
                cat = cat_cap if pd.notna(cat_cap) else current_bid if pd.notna(current_bid) else COMBINED_MIN_BID_RUB
                max_bid = floor_to_combined_grid(cat + 2 * COMBINED_STEP_RUB)
                reason = "CAP_BY_CATEGORY_MAX_PLUS_2_STEPS_NO_STABLE_CPO"
            current_effective = max(current_bid if pd.notna(current_bid) else COMBINED_MIN_BID_RUB, COMBINED_MIN_BID_RUB)
            next_bid = ceil_to_combined_grid(current_effective + COMBINED_STEP_RUB)
            cpo_next = (r.get("impressions_per_order_cur", np.nan) * next_bid / 1000.0) if pd.notna(next_bid) and pd.notna(r.get("impressions_per_order_cur", np.nan)) else np.nan
            drr_next = pct(cpo_next, avg_price) if pd.notna(cpo_next) and pd.notna(avg_price) and avg_price > 0 else np.nan
        caps.append(max_bid)
        reasons.append(reason)
        forecast_cpo_next.append(cpo_next)
        forecast_drr_next.append(drr_next)
    out["max_allowed_bid_rub"] = caps
    out["max_bid_reason"] = reasons
    out["forecast_cpo_next_step"] = forecast_cpo_next
    out["forecast_drr_next_step_pct"] = forecast_drr_next
    return out


def build_core_efficiency(keywords: pd.DataFrame, campaigns: pd.DataFrame, campaign_metrics: pd.DataFrame) -> pd.DataFrame:
    if keywords.empty:
        return pd.DataFrame()
    rows = []
    # Use product-level CPC proxy: median current CPC from search campaigns of same product.
    cpc_by_product = campaign_metrics[campaign_metrics["placement"].eq("search")].groupby("product_root", as_index=True)["cpc_cur"].median().to_dict()
    for _, r in keywords.iterrows():
        root = r.get("product_root", "")
        cpc = cpc_by_product.get(root, np.nan)
        clicks = float(r.get("clicks", 0) or 0)
        orders = float(r.get("orders", 0) or 0)
        clicks_per_order = clicks / orders if orders > 0 else np.inf if clicks > 0 else np.nan
        query_cpo = clicks_per_order * cpc if pd.notna(cpc) and np.isfinite(clicks_per_order) else np.inf if clicks > 0 and orders <= 0 else np.nan
        frequency = float(r.get("frequency", 0) or 0)
        click_share = clicks / frequency * 100 if frequency > 0 else np.nan
        # classification done after ranking
        rows.append({
            "product_root": root,
            "nm_id": r.get("nm_id", pd.NA),
            "supplier_article": r.get("supplier_article", ""),
            "subject_norm": r.get("subject_norm", ""),
            "query_text": r.get("query_text", ""),
            "keyword_group": r.get("keyword_group", ""),
            "frequency": frequency,
            "clicks": clicks,
            "orders": orders,
            "clicks_per_order": clicks_per_order,
            "proxy_cpc": cpc,
            "query_cpo": query_cpo,
            "median_position": r.get("median_position", np.nan),
            "visibility_pct": r.get("visibility_pct", np.nan),
            "our_click_share_pct": click_share,
            "orders_share": r.get("orders_share", np.nan),
            "orders_cum_share": r.get("orders_cum_share", np.nan),
        })
    out = pd.DataFrame(rows)
    if out.empty:
        return out
    # Pick 1-3 flagships by orders, frequency and acceptable query CPO.
    out["flagship_score"] = (
        out["orders"].fillna(0) * 1000
        + np.log1p(out["frequency"].fillna(0)) * 50
        - out["clicks_per_order"].replace(np.inf, 999).fillna(999) * 5
        - out["median_position"].fillna(99)
    )
    out["query_role"] = "core_secondary"
    for root, idxs in out.groupby("product_root").groups.items():
        part = out.loc[list(idxs)].copy()
        part = part[part["orders"].fillna(0) > 0]
        top = part.sort_values("flagship_score", ascending=False).head(FLAGSHIP_MAX_PER_PRODUCT).index
        out.loc[top, "query_role"] = "flagship"
    out.loc[(out["orders"].fillna(0) <= 0) & (out["clicks"].fillna(0) > 0), "query_role"] = "bad_no_orders"
    out.loc[(out["clicks_per_order"].replace(np.inf, 999) >= 20) & out["orders"].fillna(0).gt(0), "query_role"] = "broad_expensive"
    return out


def summarize_core_by_product(core: pd.DataFrame) -> pd.DataFrame:
    if core.empty:
        return pd.DataFrame(columns=["product_root"])
    g = core.groupby("product_root", as_index=False).agg(
        flagship_queries=("query_text", lambda x: "; ".join(list(x.head(3)))),
        flagship_orders=("orders", "sum"),
        flagship_clicks=("clicks", "sum"),
        flagship_frequency=("frequency", "sum"),
        flagship_position=("median_position", "median"),
        flagship_visibility_pct=("visibility_pct", "mean"),
        flagship_cpo=("query_cpo", "median"),
        flagship_clicks_per_order=("clicks_per_order", "median"),
    )
    return g


def rank_blocks(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    # A block is product_root + placement. For brushes product_root normally 901 but individual article can still be handled by supplier_article if needed.
    out["block_key"] = out["product_root"].fillna("") + "|" + out["placement"].fillna("")
    out["cpo_cur_safe"] = out["cpo_cur"].replace([np.inf, -np.inf], np.nan)
    out["block_score"] = (
        out["orders_cur"].fillna(0) * 1000
        + out["clicks_cur"].fillna(0) * 5
        + out["ctr_pct_cur"].fillna(0) * 20
        - out["cpo_cur_safe"].fillna(999) * 2
        - out["drr_pct_cur"].fillna(999) * 5
    )
    out["block_rank"] = out.groupby("block_key")["block_score"].rank(method="first", ascending=False)
    out["active_in_block"] = out.groupby("block_key")["is_active"].transform("sum") if "is_active" in out.columns else 1
    out["is_block_leader"] = out["block_rank"].eq(1)
    out["is_block_top3"] = out["block_rank"].le(3)
    return out


# -------------------------
# Decisions
# -------------------------

def decide_all(campaigns: pd.DataFrame, core: pd.DataFrame, pause_history: pd.DataFrame, windows: Dict[str, pd.Timestamp]) -> pd.DataFrame:
    df = campaigns.copy()
    core_summary = summarize_core_by_product(core[core["query_role"].eq("flagship")]) if not core.empty else pd.DataFrame()
    if not core_summary.empty:
        df = df.merge(core_summary, on="product_root", how="left")
    else:
        for c in ["flagship_queries", "flagship_orders", "flagship_clicks", "flagship_frequency", "flagship_position", "flagship_visibility_pct", "flagship_cpo", "flagship_clicks_per_order"]:
            df[c] = np.nan

    df = rank_blocks(df)
    df = select_ramp_slots(df)

    decisions = []
    for _, r in df.iterrows():
        decisions.append(decide_campaign(r, pause_history, windows))
    res = pd.DataFrame(decisions)
    return df.merge(res, on="campaign_id", how="left")


def select_ramp_slots(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["ramp_slot_selected"] = False
    out["ramp_queue_rank"] = np.nan
    # Candidate: active, not leader or leader with low volume, not too high DRR, below 5000 impressions in mature window.
    for (root, placement), part in out.groupby(["product_root", "placement"], dropna=False):
        if str(root) == "":
            continue
        max_slots = RAMP_MAX_ACTIVE_CPC_PER_BLOCK if placement == "search" else RAMP_MAX_ACTIVE_CPM_PER_BLOCK
        cand = part.copy()
        cand = cand[cand["is_active"].fillna(True)]
        cand = cand[cand["impressions_cur"].fillna(0) < RAMP_TARGET_IMPRESSIONS]
        cand = cand[~cand["is_new"].fillna(False) | cand["is_new"].fillna(False)]
        if cand.empty:
            continue
        cand["ramp_score"] = (
            cand["ctr_pct_cur"].fillna(0) * 100
            + cand["clicks_cur"].fillna(0)
            + cand["orders_cur"].fillna(0) * 50
            - cand["drr_pct_cur"].fillna(0) * 2
        )
        cand = cand.sort_values("ramp_score", ascending=False)
        ranks = range(1, len(cand) + 1)
        out.loc[cand.index, "ramp_queue_rank"] = list(ranks)
        selected = cand.head(max_slots).index
        out.loc[selected, "ramp_slot_selected"] = True
    return out


def decide_campaign(r: pd.Series, pause_history: pd.DataFrame, windows: Dict[str, pd.Timestamp]) -> Dict[str, Any]:
    cid = int(r["campaign_id"])
    placement = r.get("placement", "")
    active = bool(r.get("is_active", True))
    current_bid = r.get("real_bid_rub", np.nan)
    max_bid = r.get("max_allowed_bid_rub", np.nan)
    drr_cur = r.get("drr_pct_cur", np.nan)
    drr_14 = r.get("drr_pct_14d", np.nan)
    impressions_14 = r.get("impressions_14d", 0.0)
    impressions_cur = r.get("impressions_cur", 0.0)
    orders_cur = r.get("orders_cur", 0.0)
    cpo_cur = r.get("cpo_cur", np.nan)
    ctr_cur = r.get("ctr_pct_cur", np.nan)
    new_status = bool(r.get("is_new", False))
    is_leader = bool(r.get("is_block_leader", False))
    is_top3 = bool(r.get("is_block_top3", False))
    active_in_block = int(r.get("active_in_block", 1) or 1)
    ramp_slot = bool(r.get("ramp_slot_selected", False))
    recent_bid = bool(r.get("recent_bid_change", False))
    subject = str(r.get("subject_norm", ""))
    root = str(r.get("product_root", ""))

    # Normalize bid and step.
    if placement == "search":
        step = SEARCH_STEP_RUB
        min_bid = SEARCH_MIN_BID_RUB
        bid_effective = int(round(current_bid)) if pd.notna(current_bid) else min_bid
        next_up = bid_effective + step
        next_down = max(min_bid, bid_effective - step)
    else:
        step = COMBINED_STEP_RUB
        min_bid = COMBINED_MIN_BID_RUB
        bid_effective = max(float(current_bid) if pd.notna(current_bid) else min_bid, min_bid)
        bid_effective = ceil_to_combined_grid(bid_effective)
        next_up = bid_effective + step
        next_down = max(min_bid, bid_effective - step)

    # paused campaigns: start candidates.
    last_pause_status = get_last_pause_status(pause_history, cid)
    if not active:
        start_decision = decide_start_candidate(r, last_pause_status, windows)
        return {"campaign_id": cid, **start_decision}

    # Never pause NEW for 14 days.
    if new_status:
        if ramp_slot and can_raise(next_up, max_bid):
            return decision(cid, "raise", next_up, "NEW_RAMP_RAISE_TO_TRAFFIC", f"NEW<{NEW_NO_PAUSE_DAYS}д: пауза запрещена; разгон по одному шагу до 5000 показов; ставка {bid_effective}->{next_up}; cap={max_bid}")
        return decision(cid, "hold", bid_effective, "NEW_UNDER_14D_NO_PAUSE", f"NEW<{NEW_NO_PAUSE_DAYS}д: пауза запрещена; ждём трафик/позиции/клики")

    # Hard pause: high mature traffic and high DRR, but keep leader/last active.
    if (impressions_14 >= 10000 and pd.notna(drr_14) and drr_14 > DRR_PAUSE_LIMIT_PCT and active_in_block > 1 and not is_leader):
        return decision(cid, "pause", bid_effective, "PAUSE_HIGH_DRR_14D_10000_REALLOCATE", f"14д: показы={impressions_14:.0f} >=10000, ДРР={drr_14:.1f}% >15%; не лидер блока; пауза для перелива бюджета в лидеров")

    # Low-volume ramp flow: if not enough impressions, do not judge by orders yet. Ramp only selected slots, pause/queue others.
    if impressions_cur < RAMP_TARGET_IMPRESSIONS:
        if ramp_slot and can_raise(next_up, max_bid):
            return decision(cid, "raise", next_up, "RAMP_ACTIVE_TO_5000_IMPRESSIONS", f"Разгон: показов в зрелом окне {impressions_cur:.0f}<5000; слот выбран по CTR/кликам; ставка {bid_effective}->{next_up}; cap={max_bid}")
        if not is_leader and active_in_block > 1:
            return decision(cid, "pause", bid_effective, "PAUSE_TO_RAMP_QUEUE_LOW_VOLUME", f"Не лидер и не выбран в текущий слот разгона; ставим в очередь/паузу, чтобы не разгонять много РК одновременно")
        return decision(cid, "hold", bid_effective, "LOW_VOLUME_LEADER_HOLD", f"Лидер/единственная РК блока, но показов {impressions_cur:.0f}<5000; держим без резких действий")

    # If enough impressions and high DRR: reduce first unless min bid; not pause leader.
    if pd.notna(drr_cur) and drr_cur > DRR_RAISE_GATE_PCT:
        # Special case: flagship position is bad and query economics are acceptable, allow a controlled flagship test instead of lowering.
        flagship_position = r.get("flagship_position", np.nan)
        flagship_cpo = r.get("flagship_cpo", np.nan)
        avg_price = r.get("avg_finished_price", np.nan)
        flagship_drr = pct(flagship_cpo, avg_price) if pd.notna(flagship_cpo) and pd.notna(avg_price) and avg_price > 0 else np.nan
        if (pd.notna(flagship_position) and flagship_position > FLAGSHIP_TARGET_POSITION and pd.notna(flagship_drr) and flagship_drr <= DRR_FORECAST_CAP_PCT and can_raise(next_up, max_bid) and not recent_bid):
            return decision(cid, "raise", next_up, "TEST_RAISE_DRR_HIGH_FLAGSHIP_BAD_POSITION", f"ДРР кампании {drr_cur:.1f}% >10%, но флагманские запросы ниже топ-10 и прогнозный ДРР флагмана {flagship_drr:.1f}%<=16%; тестируем рост ставки {bid_effective}->{next_up}")
        if bid_effective > min_bid and not recent_bid:
            return decision(cid, "lower", next_down, "DRR_GT_10_LOWER_ONE_STEP", f"ДРР зрелого окна {drr_cur:.1f}% >10%; сначала сушим на 1 шаг: {bid_effective}->{next_down}; следим за CORE-кликами")
        return decision(cid, "hold", bid_effective, "DRR_GT_10_MIN_OR_RECENT_CHANGE_HOLD", f"ДРР {drr_cur:.1f}% >10%, но ставка на минимуме или недавно менялась; hold")

    # DRR <= 10: raise only if cap allows and not too recent. This is a test, not automatic scaling.
    if pd.notna(drr_cur) and drr_cur <= DRR_RAISE_GATE_PCT:
        if recent_bid:
            return decision(cid, "hold", bid_effective, "WAIT_AFTER_RECENT_BID_CHANGE", f"ДРР {drr_cur:.1f}% <=10%, но ставка менялась недавно; ждём быстрые метрики CORE")
        if can_raise(next_up, max_bid):
            if placement == "search":
                return decision(cid, "raise", next_up, "TEST_RAISE_CORE_TRAFFIC_ONE_STEP", f"ДРР {drr_cur:.1f}%<=10%; тест +1р: {bid_effective}->{next_up}; cap={max_bid}; успех = рост доли флагманских CORE-кликов/показов без ухода в широкий мусор")
            return decision(cid, "raise", next_up, "TEST_RAISE_SHELF_TRAFFIC_ONE_STEP", f"ДРР {drr_cur:.1f}%<=10%; тест CPM +6р: {bid_effective}->{next_up}; cap={max_bid}; успех = рост показов/кликов и CTR не падает")
        return decision(cid, "hold", bid_effective, "RAISE_BLOCKED_BY_FORECAST_DRR_CAP", f"ДРР {drr_cur:.1f}%<=10%, но следующий шаг {next_up} выше предельной ставки {max_bid}; hold")

    return decision(cid, "hold", bid_effective, "NO_STABLE_DRR_HOLD", "Нет стабильного ДРР в зрелом окне; без изменений")


def can_raise(next_bid: float, max_bid: Any) -> bool:
    if pd.isna(next_bid):
        return False
    if pd.isna(max_bid):
        return False
    return float(next_bid) <= float(max_bid)


def get_last_pause_status(pause_history: pd.DataFrame, campaign_id: int) -> Dict[str, Any]:
    if pause_history.empty:
        return {}
    p = pause_history[pause_history["campaign_id"].eq(campaign_id)].copy()
    if p.empty:
        return {}
    p = p.sort_values("pause_date")
    last = p.iloc[-1]
    return last.to_dict()


def decide_start_candidate(r: pd.Series, last_pause: Dict[str, Any], windows: Dict[str, pd.Timestamp]) -> Dict[str, Any]:
    cid = int(r["campaign_id"])
    placement = r.get("placement", "")
    min_bid = SEARCH_MIN_BID_RUB if placement == "search" else COMBINED_MIN_BID_RUB
    drr_14 = r.get("drr_pct_14d", np.nan)
    orders_14 = r.get("orders_14d", 0.0)
    pause_date = pd.to_datetime(last_pause.get("pause_date", pd.NaT), errors="coerce")
    days_paused = (windows["as_of"] - pause_date).days if pd.notna(pause_date) else np.nan
    if pd.notna(drr_14) and drr_14 <= DRR_PAUSE_LIMIT_PCT and orders_14 > 0 and (pd.isna(days_paused) or days_paused >= POST_PAUSE_CHECK_DAYS):
        return decision(cid, "start", min_bid, "START_AFTER_DRR_RECOVERY", f"Paused РК: заказы дозрели={orders_14:.0f}, ДРР14={drr_14:.1f}%<=15%; запуск по минимальной ставке {min_bid}")
    return decision(cid, "hold_paused", min_bid, "KEEP_PAUSED_WAIT_RECOVERY", f"Paused РК: возврат не подтверждён; orders14={orders_14:.0f}, drr14={drr_14 if pd.notna(drr_14) else 'н/д'}")


def decision(campaign_id: int, action: str, new_bid: Any, reason_code: str, reason_text: str) -> Dict[str, Any]:
    return {
        "campaign_id": campaign_id,
        "action": action,
        "new_bid_rub": new_bid,
        "reason_code": reason_code,
        "reason_text": reason_text,
    }


def build_payload_preview(decisions: pd.DataFrame) -> pd.DataFrame:
    if decisions.empty:
        return pd.DataFrame()
    rows = []
    for _, r in decisions.iterrows():
        action = r.get("action", "hold")
        if action not in {"raise", "lower", "pause", "start"}:
            continue
        cid = r.get("campaign_id")
        if action in {"raise", "lower"}:
            rows.append({
                "api_action": "PATCH_BID",
                "campaign_id": cid,
                "new_bid_rub": r.get("new_bid_rub"),
                "placement": r.get("placement"),
                "reason_code": r.get("reason_code"),
            })
        elif action == "pause":
            rows.append({
                "api_action": "PAUSE",
                "campaign_id": cid,
                "new_bid_rub": np.nan,
                "placement": r.get("placement"),
                "reason_code": r.get("reason_code"),
            })
        elif action == "start":
            rows.append({
                "api_action": "START",
                "campaign_id": cid,
                "new_bid_rub": r.get("new_bid_rub"),
                "placement": r.get("placement"),
                "reason_code": r.get("reason_code"),
            })
    return pd.DataFrame(rows)


def block_reallocation_summary(decisions: pd.DataFrame) -> pd.DataFrame:
    if decisions.empty:
        return pd.DataFrame()
    g = decisions.groupby(["product_root", "placement"], dropna=False).agg(
        active_campaigns=("campaign_id", "count"),
        spend_cur=("spend_cur", "sum"),
        orders_cur=("orders_cur", "sum"),
        clicks_cur=("clicks_cur", "sum"),
        pause_count=("action", lambda x: (x == "pause").sum()),
        raise_count=("action", lambda x: (x == "raise").sum()),
        lower_count=("action", lambda x: (x == "lower").sum()),
        leaders=("campaign_id", lambda x: ",".join(map(str, list(x.head(3))))),
    ).reset_index()
    g["cpo_cur"] = np.where(g["orders_cur"] > 0, g["spend_cur"] / g["orders_cur"], np.nan)
    g["method_comment"] = "Цель: pause_and_reallocate, расход не просто срезать, а перелить в лидеров; успех = orders_index/spend_index > 1"
    return g


def summary_table(decisions: pd.DataFrame, windows: Dict[str, pd.Timestamp]) -> pd.DataFrame:
    rows = [
        ["version", VERSION],
        ["as_of", windows["as_of"].date().isoformat()],
        ["current_window", f"{windows['current_start'].date()}..{windows['current_end'].date()}"],
        ["base_window", f"{windows['base_start'].date()}..{windows['base_end'].date()}"],
        ["pause_window", f"{windows['pause_start'].date()}..{windows['pause_end'].date()}"],
        ["drr_raise_gate_pct", DRR_RAISE_GATE_PCT],
        ["drr_pause_limit_pct", DRR_PAUSE_LIMIT_PCT],
        ["forecast_drr_cap_pct", DRR_FORECAST_CAP_PCT],
        ["ramp_target_impressions", RAMP_TARGET_IMPRESSIONS],
        ["abc_profitability_used", "NO"],
    ]
    if not decisions.empty and "action" in decisions.columns:
        for action, cnt in decisions["action"].value_counts(dropna=False).items():
            rows.append([f"action_{action}", int(cnt)])
    return pd.DataFrame(rows, columns=["metric", "value"])


def write_outputs(path: str, decisions: pd.DataFrame, core: pd.DataFrame, payload: pd.DataFrame, windows: Dict[str, pd.Timestamp]) -> None:
    out_path = Path(path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        summary_table(decisions, windows).to_excel(writer, sheet_name="Сводка", index=False)
        cols_dec = [
            "campaign_id", "supplier_article", "product_root", "subject_norm", "placement", "campaign_status", "is_new",
            "real_bid_rub", "max_allowed_bid_rub", "new_bid_rub", "action", "reason_code", "reason_text",
            "impressions_cur", "clicks_cur", "orders_cur", "spend_cur", "order_sum_cur", "drr_pct_cur", "cpo_cur", "ctr_pct_cur",
            "impressions_14d", "orders_14d", "spend_14d", "order_sum_14d", "drr_pct_14d",
            "avg_finished_price", "forecast_cpo_next_step", "forecast_drr_next_step_pct", "max_bid_reason",
            "flagship_queries", "flagship_position", "flagship_cpo", "flagship_clicks_per_order", "block_rank", "ramp_queue_rank", "ramp_slot_selected",
            "last_bid_change_date", "days_since_last_bid_change", "last_bid_reason_code",
        ]
        existing = [c for c in cols_dec if c in decisions.columns]
        decisions[existing].to_excel(writer, sheet_name="Решения", index=False)
        if not core.empty:
            core.to_excel(writer, sheet_name="CORE_эффективность_по_окнам", index=False)
        else:
            pd.DataFrame({"note": ["CORE source not provided. Pass --previous-output with Ключевые_фразы_80 or add search query daily source."]}).to_excel(writer, sheet_name="CORE_эффективность_по_окнам", index=False)
        cap_cols = [c for c in ["campaign_id", "supplier_article", "placement", "real_bid_rub", "max_allowed_bid_rub", "forecast_cpo_next_step", "forecast_drr_next_step_pct", "avg_finished_price", "clicks_per_order_cur", "impressions_per_order_cur", "max_bid_reason"] if c in decisions.columns]
        decisions[cap_cols].to_excel(writer, sheet_name="Предельные_ставки", index=False)
        pause_cols = [c for c in ["campaign_id", "supplier_article", "product_root", "placement", "action", "reason_code", "reason_text", "impressions_14d", "drr_pct_14d", "active_in_block", "is_block_leader", "is_new"] if c in decisions.columns]
        decisions[decisions["action"].isin(["pause", "start", "hold_paused"])][pause_cols].to_excel(writer, sheet_name="Паузы_и_возвраты", index=False)
        ramp_cols = [c for c in ["campaign_id", "supplier_article", "product_root", "placement", "ramp_queue_rank", "ramp_slot_selected", "action", "reason_code", "impressions_cur", "ctr_pct_cur", "orders_cur", "drr_pct_cur", "real_bid_rub", "new_bid_rub", "max_allowed_bid_rub"] if c in decisions.columns]
        decisions.sort_values(["product_root", "placement", "ramp_queue_rank"], na_position="last")[ramp_cols].to_excel(writer, sheet_name="Разгон_очередь", index=False)
        block_reallocation_summary(decisions).to_excel(writer, sheet_name="Блоки_перелива", index=False)
        payload.to_excel(writer, sheet_name="API_payload_preview", index=False)


# -------------------------
# CLI
# -------------------------

def run(args: argparse.Namespace) -> int:
    as_of = pd.Timestamp(args.as_of).normalize() if args.as_of else pd.Timestamp(datetime.utcnow().date())
    windows = date_windows(as_of)

    ads, campaigns = load_ads_daily(args.ads)
    orders = load_orders(args.orders)
    bid_history = load_bid_history(args.bid_history)
    pause_history = load_pause_history(args.pause_history)
    keywords = load_keywords_from_previous(args.previous_output)

    campaign_base = build_campaign_base(ads, campaigns, orders, bid_history, windows)
    campaign_base = compute_bid_caps(campaign_base)
    core = build_core_efficiency(keywords, campaigns, campaign_base)
    decisions = decide_all(campaign_base, core, pause_history, windows)
    payload = build_payload_preview(decisions)

    write_outputs(args.out, decisions, core, payload, windows)

    if args.print_summary:
        print(summary_table(decisions, windows).to_string(index=False))
        print(f"Output: {args.out}")
        print(f"API actions preview: {len(payload)}")
    return 0


def build_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="WB Ads Manager FIX46 decision engine")
    p.add_argument("--ads", required=True, help="WB advertising Excel: Реклама_YYYY-WNN.xlsx")
    p.add_argument("--orders", required=False, default=None, help="WB orders Excel for avg finishedPrice")
    p.add_argument("--previous-output", required=False, default=None, help="Previous output workbook with Ключевые_фразы_80")
    p.add_argument("--bid-history", required=False, default=None, help="История_ставок.xlsx")
    p.add_argument("--pause-history", required=False, default=None, help="История_пауз.xlsx")
    p.add_argument("--as-of", required=False, default=None, help="Decision date YYYY-MM-DD. Example: 2026-06-11")
    p.add_argument("--out", required=False, default=f"wb_ads_decisions_{VERSION}.xlsx")
    p.add_argument("--print-summary", action="store_true")
    return p


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = build_arg_parser()
    args = parser.parse_args(argv)
    return run(args)


if __name__ == "__main__":
    raise SystemExit(main())
