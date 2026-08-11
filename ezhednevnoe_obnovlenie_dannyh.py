#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Ежедневный сбор данных Wildberries с сохранением в Yandex Cloud Object Storage.
Данные хранятся только в недельных файлах (кроме воронки продаж и 1С).
Автоматическое получение всех артикулов из заказов для отчёта по ключам.
Формат для keywords: Неделя ГГГГ-WНН.xlsx
Финансовые показатели: новый метод POST /api/finance/v1/sales-reports/detailed, в ежедневном режиме только целевая дата.
Всегда читается первый лист в файле.
Поисковые запросы: загружается целевая дата по правилу времени запуска.
Реклама: каждый день заново получает последние 14 дат статистики и хранит ежедневные снимки для анализа лага.
Отчёт 1c_stocks временно исключён из списка (можно вернуть позже).
Добавлен agent_catalog: единый ZIP для ИИ-агента с карточками WB, характеристиками, размерами, фото и AI-friendly паспортом по каждому артикулу продавца.
Для TOPFACE/MISSTAIS используются основные WB-токены; FBS по этим магазинам собирается только по подтверждённым складам в Липецке (ID 1728667/1935990). Для Finance можно задать отдельные WB_FINANCE_KEY_TOPFACE/WB_FINANCE_KEY_MISSTAIS. Для FINICK используется FINICK_API_WB; finance и keywords для FINICK отключены. FINICK при первом запуске догружает 7 полностью завершённых недель истории (без текущей недели) для доступных исторических отчётов.
"""

import os
import io
import json
import time
import uuid
import zipfile
import tempfile
import traceback
import re
import sys
import argparse
import hashlib
import mimetypes
from urllib.parse import urlparse
from datetime import datetime, timedelta
from typing import List, Dict, Optional, Tuple, Any, Set
import warnings
from collections import defaultdict

import pandas as pd
import requests
import boto3
from botocore.client import Config
from botocore.exceptions import ClientError
import pytz

warnings.simplefilter(action='ignore', category=FutureWarning)

SCRIPT_VERSION = "2026-08-11_v42_AGENT_CATALOG_ARCHIVE"

# Для TOPFACE и MISSTAIS FBS собираем только с двух подтверждённых липецких
# складов продавца. Набор общий намеренно: токен каждого магазина видит только
# свои склады, поэтому это не требует жёстко привязывать скрин к конкретному магазину.
# Все остальные FBS-склады этих магазинов игнорируются. FINICK не ограничиваем.
FBS_LIPETSK_WAREHOUSE_IDS = {1728667, 1935990}
FBS_LIPETSK_ONLY_STORES = {'TOPFACE', 'MISSTAIS'}


def parse_date_yyyy_mm_dd(value: str) -> datetime.date:
    return datetime.strptime(value, "%Y-%m-%d").date()


def resolve_target_data_date(now_msk: Optional[datetime] = None) -> datetime.date:
    """Целевая дата выгрузки.

    Правило:
    - если явно задан WB_TARGET_DATE=YYYY-MM-DD, используем его;
    - если запуск по МСК после 15:00 включительно, выгружаем вчера;
    - если запуск по МСК до 15:00, выгружаем предыдущий полный день, то есть позавчера.

    Пример: запуск 26 числа в 02:00 МСК -> данные за 24 число.
    Пример: запуск 26 числа в 15:00 МСК -> данные за 25 число.
    """
    forced = (os.environ.get("WB_TARGET_DATE", "") or "").strip()
    if forced:
        return parse_date_yyyy_mm_dd(forced)

    if now_msk is None:
        now_msk = datetime.now(pytz.timezone("Europe/Moscow"))

    base_shift_days = 1 if now_msk.hour >= 15 else 2
    return now_msk.date() - timedelta(days=base_shift_days)



# ========================== КЛАСС ДЛЯ РАБОТЫ С YANDEX CLOUD ==========================

class S3Storage:
    """Клиент для работы с S3-совместимым хранилищем Yandex Cloud."""

    def __init__(self, access_key: str, secret_key: str, bucket_name: str):
        self.bucket = bucket_name
        self.s3 = boto3.client(
            's3',
            endpoint_url='https://storage.yandexcloud.net',
            aws_access_key_id=access_key,
            aws_secret_access_key=secret_key,
            region_name='ru-central1',
            config=Config(
                signature_version='s3v4',
                read_timeout=300,
                connect_timeout=60,
                retries={'max_attempts': 5}
            )
        )
        print(f"🔑 DEBUG: подключение к Yandex Cloud, Access Key (первые 5 символов): {access_key[:5]}...")

    def read_excel(self, key: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
        try:
            obj = self.s3.get_object(Bucket=self.bucket, Key=key)
            data = obj['Body'].read()
            df = pd.read_excel(io.BytesIO(data), sheet_name=sheet_name)
            return df
        except ClientError as e:
            if e.response['Error']['Code'] == 'NoSuchKey':
                return pd.DataFrame()
            else:
                raise e
        except Exception as e:
            print(f"Ошибка чтения {key}: {e}")
            return pd.DataFrame()

    def write_excel(self, key: str, df: pd.DataFrame, sheet_name: str = 'Data'):
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            tmp_path = tmp.name
        try:
            df.to_excel(tmp_path, index=False, sheet_name=sheet_name)
            self.upload_file(tmp_path, key)
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

    def write_excel_multi(self, key: str, sheets: Dict[str, pd.DataFrame]):
        """
        Сохраняет несколько листов в один Excel-файл.
        sheets: словарь {имя_листа: DataFrame}
        """
        if not sheets:
            return
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            tmp_path = tmp.name
        try:
            with pd.ExcelWriter(tmp_path, engine='openpyxl') as writer:
                for sheet_name, df in sheets.items():
                    if not df.empty:
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
            self.upload_file(tmp_path, key)
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

    def upload_file(self, local_path: str, key: str):
        self.s3.upload_file(local_path, self.bucket, key)

    def read_json(self, key: str) -> Optional[dict]:
        try:
            obj = self.s3.get_object(Bucket=self.bucket, Key=key)
            raw = obj['Body'].read()
            return json.loads(raw.decode('utf-8'))
        except ClientError as e:
            if e.response['Error']['Code'] in {'NoSuchKey', '404'}:
                return None
            raise
        except Exception as e:
            print(f"Ошибка чтения JSON {key}: {e}")
            return None

    def write_json(self, key: str, data: dict):
        payload = json.dumps(data, ensure_ascii=False, indent=2, default=str).encode('utf-8')
        self.s3.put_object(
            Bucket=self.bucket,
            Key=key,
            Body=payload,
            ContentType='application/json; charset=utf-8',
        )

    def file_exists(self, key: str) -> bool:
        try:
            self.s3.head_object(Bucket=self.bucket, Key=key)
            return True
        except ClientError:
            return False

    def list_files(self, prefix: str) -> List[str]:
        try:
            response = self.s3.list_objects_v2(Bucket=self.bucket, Prefix=prefix)
            if 'Contents' in response:
                return [obj['Key'] for obj in response['Contents']]
            else:
                return []
        except Exception as e:
            print(f"Ошибка при list_files: {e}")
            return []


# ====================== ОСНОВНОЙ КЛАСС СБОРЩИКА ДАННЫХ ======================

class WildberriesDailyUpdater:
    def __init__(self, api_keys: Dict[str, Dict[str, str]], s3: S3Storage):
        self.api_keys = api_keys
        self.s3 = s3
        self.start_time = datetime.now(pytz.timezone('Europe/Moscow'))
        self.target_date = resolve_target_data_date(self.start_time)
        self.data_period_days = 90
        self.keyword_errors = []  # для сбора ошибок поисковых запросов

        self.reports_config = {
            'orders': {
                'name': 'Заказы',
                'folder': 'Заказы',
                'date_column': 'date',
                'id_columns': ['date', 'gNumber', 'srid'],
                'api_url': 'https://statistics-api.wildberries.ru/api/v1/supplier/orders',
                'api_method': 'GET',
                'key_type': 'promo',
            },
            'stocks': {
                'name': 'Остатки',
                'folder': 'Остатки',
                'date_column': 'Дата запроса',
                'id_columns': ['Дата запроса', 'Артикул WB', 'Склад'],
                'api_url': 'https://seller-analytics-api.wildberries.ru/api/v1/warehouse_remains',
                'api_method': 'GET',
                'key_type': 'promo',
            },
            'fbs_orders': {
                'name': 'Заказы FBS',
                'folder': 'Заказы',
                'date_column': 'Дата заказа',
                'id_columns': ['ID заказа FBS'],
                'api_url': 'https://marketplace-api.wildberries.ru/api/v3/orders',
                'api_method': 'GET',
                'key_type': 'marketplace',
            },
            'fbs_stocks': {
                'name': 'Остатки FBS',
                'folder': 'Остатки',
                'date_column': 'Дата запроса',
                'id_columns': ['Дата запроса', 'ID склада продавца', 'chrtId'],
                'api_url': 'https://marketplace-api.wildberries.ru/api/v3/stocks',
                'api_method': 'POST',
                'key_type': 'marketplace',
            },
            'finance': {
                'name': 'Финансовые показатели',
                'folder': 'Финансовые показатели',
                'date_column': 'rr_dt',
                'id_columns': ['rr_dt', 'rrd_id', 'nm_id'],
                'api_url': 'https://finance-api.wildberries.ru/api/finance/v1/sales-reports/detailed',
                'api_method': 'POST',
                'key_type': 'finance',
            },
            'keywords': {
                'name': 'Позиции по Ключам',
                'folder': 'Поисковые запросы',
                'date_column': 'Дата',
                'id_columns': ['Дата', 'Поисковый запрос', 'Артикул WB', 'Фильтр'],
                'api_url': 'https://seller-analytics-api.wildberries.ru/api/v2/search-report/product/search-texts',
                'api_method': 'POST',
                'key_type': 'promo',
            },
            'funnel': {
                'name': 'Воронка продаж',
                'folder': 'Воронка продаж',
                'filename': 'Воронка продаж.xlsx',
                'date_column': 'dt',
                'id_columns': ['dt', 'nmID'],
                'api_url': 'https://seller-analytics-api.wildberries.ru/api/v2/nm-report/downloads',
                'api_method': 'POST',
                'key_type': 'promo',
                'retention_days': 90,
            },
            'adverts': {
                'name': 'Реклама',
                'folder': 'Реклама',
                'date_column': 'Дата',
                'id_columns': ['ID кампании', 'Дата'],
                'api_url': 'https://advert-api.wildberries.ru/api/advert/v2/adverts',
                'api_method': 'GET',
                'key_type': 'promo',
                'retention_days': 30,
            },
            '1c_stocks': {
                'name': 'Остатки 1С',
                'folder': 'Остатки',
                'filename': 'Остатки_1С.xlsx',
                'date_column': None,
                'id_columns': [],
                'api_url': None,
                'key_type': None,
            }
        }

        self.delays = {
            'orders': 65,
            'stocks': 65,
            'fbs_orders': 1,
            'fbs_stocks': 1,
            'finance': 65,
            'keywords': 90,
            'funnel': 30,
            'adverts': 30,
            '1c_stocks': 0,
            'agent_catalog': 0,
        }

        # v22: поисковые запросы выгружаем по всем артикулам из заказов, без ограничения 4 категориями.
        self.target_subjects = []
        self.log(f"VERSION: {SCRIPT_VERSION}")
        self.log(f"🚀 Запуск обновления данных. Время: {self.start_time}")
        self.log(f"📅 Целевая дата выгрузки: {self.target_date:%Y-%m-%d}")

    # ====================== ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ ======================
    def log(self, message: str, level: str = "INFO", end: str = "\n"):
        print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] [{level}] {message}", end=end, flush=True)

    def _get_week_start(self, date: datetime) -> datetime:
        return date - timedelta(days=date.weekday())

    def _get_weekly_key(self, store_name: str, report_type: str, date: datetime) -> str:
        year, week, _ = date.isocalendar()
        config = self.reports_config[report_type]
        if report_type == 'keywords':
            filename = f"Неделя {year}-W{week:02d}.xlsx"
        else:
            filename = f"{config['name']}_{year}-W{week:02d}.xlsx"
        return f"Отчёты/{config['folder']}/{store_name}/Недельные/{filename}"

    def _load_weekly_data(self, store_name: str, report_type: str, week_date: datetime) -> pd.DataFrame:
        key = self._get_weekly_key(store_name, report_type, week_date)
        self.log(f"📥 Загрузка недельного файла: {key}")
        try:
            df = self.s3.read_excel(key, sheet_name=0)
            if df.empty:
                self.log(f"ℹ️ Файл пуст")
                return df
            self.log(f"📋 Колонки в файле: {list(df.columns)}")
            date_col = self.reports_config[report_type]['date_column']
            if date_col in df.columns:
                df[date_col] = pd.to_datetime(df[date_col]).dt.strftime('%Y-%m-%d')
                unique_dates = sorted(df[date_col].unique())
                self.log(f"📊 В файле {len(df)} записей, даты: {unique_dates}")
                if report_type == 'keywords' and 'Фильтр' in df.columns and 'Артикул WB' in df.columns:
                    filters_present = df['Фильтр'].unique()
                    articles_count = df['Артикул WB'].nunique()
                    self.log(f"🔍 Фильтры в файле: {list(filters_present)}, уникальных артикулов: {articles_count}")
            else:
                self.log(f"⚠️ Колонка даты '{date_col}' не найдена")
            return df
        except Exception as e:
            self.log(f"⚠️ Ошибка загрузки {key}: {e}")
            return pd.DataFrame()

    def _save_weekly_data(self, df: pd.DataFrame, store_name: str, report_type: str, week_date: datetime) -> bool:
        if df.empty:
            return True
        key = self._get_weekly_key(store_name, report_type, week_date)
        config = self.reports_config[report_type]

        before = len(df)
        if config['id_columns']:
            existing_cols = [c for c in config['id_columns'] if c in df.columns]
            if existing_cols:
                df = df.drop_duplicates(subset=existing_cols, keep='last')
                after = len(df)
                if before > after:
                    self.log(f"🔍 Удалено дубликатов в недельном файле: {before - after}")

        try:
            self.s3.write_excel(key, df, sheet_name=config['name'])
            self.log(f"✅ Недельный файл сохранён: {key}, записей: {len(df)}")
            return True
        except Exception as e:
            self.log(f"❌ Ошибка сохранения {key}: {e}")
            traceback.print_exc()
            return False

    def _get_date_range_90_days(self) -> Tuple[datetime.date, datetime.date]:
        today = datetime.now(pytz.timezone('Europe/Moscow')).date()
        end_date = today - timedelta(days=1)
        start_date = end_date - timedelta(days=self.data_period_days - 1)
        return start_date, end_date

    def _get_yesterday_range(self) -> Tuple[datetime.date, datetime.date]:
        return self.target_date, self.target_date

    def _get_daily_or_backfill_range(self, env_name: str) -> Tuple[datetime.date, datetime.date]:
        """По умолчанию берём только целевую дату. Историческую догрузку включаем только явно через env YYYY-MM-DD."""
        end_date = self.target_date
        start_date = _parse_optional_date_env(env_name) or end_date
        if start_date > end_date:
            start_date = end_date
        return start_date, end_date

    def _get_last_completed_weeks_range(self, weeks: int = 7) -> Tuple[datetime.date, datetime.date]:
        """Диапазон последних полностью завершённых недель, не включая текущую неделю.

        Для запуска в пятницу 2026-08-07:
        - текущая неделя начинается 2026-08-03 и НЕ входит в историю;
        - 7 завершённых недель = 2026-06-15 .. 2026-08-02.
        """
        current_date = self.start_time.date()
        current_week_start = current_date - timedelta(days=current_date.weekday())
        end_date = current_week_start - timedelta(days=1)
        start_date = end_date - timedelta(days=weeks * 7 - 1)
        return start_date, end_date

    @staticmethod
    def _split_date_range(start_date: datetime.date, end_date: datetime.date, max_days: int) -> List[Tuple[datetime.date, datetime.date]]:
        """Разбить диапазон на куски не длиннее max_days включительно."""
        if start_date > end_date:
            return []
        result = []
        cur = start_date
        while cur <= end_date:
            chunk_end = min(end_date, cur + timedelta(days=max_days - 1))
            result.append((cur, chunk_end))
            cur = chunk_end + timedelta(days=1)
        return result

    def _get_date_range_last_n_days(self, n: int) -> Tuple[datetime.date, datetime.date]:
        today = datetime.now(pytz.timezone('Europe/Moscow')).date()
        end_date = today - timedelta(days=1)
        start_date = end_date - timedelta(days=n - 1)
        return start_date, end_date

    def _get_articles_by_subjects(
        self,
        store_name: str,
        subjects: List[str],
        min_order_date: Optional[datetime.date] = None,
    ) -> List[int]:
        """Собрать nmId из недельных файлов заказов.

        v22:
        - если subjects пустой список, берём ВСЕ артикулы из заказов, без фильтра по категориям;
        - если subjects передан, оставляем старую логику фильтрации по предметам;
        - для MISSTAIS можно передать min_order_date=2026-06-08, чтобы не тащить старые SKU.
        """
        use_subject_filter = bool(subjects)

        if use_subject_filter:
            if min_order_date:
                self.log(
                    f"🔍 Сбор артикулов из заказов по категориям: {subjects}; "
                    f"дата заказа >= {min_order_date:%Y-%m-%d}"
                )
            else:
                self.log(f"🔍 Сбор артикулов из заказов по категориям: {subjects}")
        else:
            if min_order_date:
                self.log(
                    f"🔍 Сбор ВСЕХ артикулов из заказов без фильтра по категориям; "
                    f"дата заказа >= {min_order_date:%Y-%m-%d}"
                )
            else:
                self.log("🔍 Сбор ВСЕХ артикулов из заказов без фильтра по категориям")

        prefix = f"Отчёты/Заказы/{store_name}/Недельные/"
        all_files = self.s3.list_files(prefix)
        if not all_files:
            self.log("⚠️ Не найдено недельных файлов заказов")
            return []

        articles_set = set()
        possible_nm_cols = ['nmId', 'nmID', 'Артикул WB', 'Артикул']
        possible_subj_cols = ['subject', 'Предмет', 'subjectName', 'Название предмета']
        possible_date_cols = ['date', 'Дата', 'Дата заказа', 'createdAt', 'Дата создания']

        for file_key in all_files:
            self.log(f"📄 Обработка файла: {file_key}")
            try:
                df = self.s3.read_excel(file_key, sheet_name=0)
                if df.empty:
                    continue

                nm_col = next((col for col in possible_nm_cols if col in df.columns), None)
                if nm_col is None:
                    self.log(f"⚠️ В файле {file_key} не найдена колонка с артикулом")
                    continue

                if min_order_date:
                    date_col = next((col for col in possible_date_cols if col in df.columns), None)
                    if date_col is None:
                        self.log(
                            f"⚠️ В файле {file_key} не найдена колонка даты заказа; "
                            f"файл пропущен для фильтрации с {min_order_date:%Y-%m-%d}"
                        )
                        continue
                    dates = pd.to_datetime(df[date_col], errors='coerce').dt.date
                    before = len(df)
                    df = df.loc[dates >= min_order_date].copy()
                    if df.empty:
                        self.log(f"ℹ️ В файле {file_key} нет заказов после {min_order_date:%Y-%m-%d}")
                        continue
                    self.log(f"   ↳ после фильтра по дате осталось строк: {len(df)} из {before}")

                if use_subject_filter:
                    subj_col = next((col for col in possible_subj_cols if col in df.columns), None)
                    if subj_col is None:
                        self.log(f"⚠️ В файле {file_key} не найдена колонка предмета")
                        continue

                    df[subj_col] = df[subj_col].astype(str).str.lower().str.strip()
                    target_lower = [s.lower() for s in subjects]
                    mask = df[subj_col].isin(target_lower)
                    filtered = df.loc[mask, nm_col].dropna().unique()
                else:
                    filtered = df[nm_col].dropna().unique()

                for val in filtered:
                    try:
                        articles_set.add(int(val))
                    except (ValueError, TypeError):
                        continue

            except Exception as e:
                self.log(f"❌ Ошибка при обработке файла {file_key}: {e}")
                continue

        articles = sorted(articles_set)
        if use_subject_filter:
            self.log(f"✅ Собрано {len(articles)} уникальных артикулов из заказов по выбранным категориям")
        else:
            self.log(f"✅ Собрано {len(articles)} уникальных артикулов из заказов по ВСЕМ категориям")
        return articles

    # ====================== МЕТОДЫ ДЛЯ КАЖДОГО ОТЧЁТА ======================
    def _rate_limit_wait_seconds(self, resp, default_seconds: int = 65, max_seconds: int = 900) -> int:
        """Время ожидания при 429 по заголовкам WB."""
        candidates = []
        for header in ("X-Ratelimit-Retry", "X-RateLimit-Retry", "X-Ratelimit-Reset", "X-RateLimit-Reset", "Retry-After"):
            raw = resp.headers.get(header)
            if raw is None:
                continue
            try:
                value = int(float(str(raw).strip()))
                if value > 0:
                    candidates.append(value)
            except Exception:
                continue
        wait = max(candidates) if candidates else default_seconds
        return max(5, min(wait, max_seconds))

    def _make_request(self, config: dict, headers: dict, date_str: str, **kwargs) -> Optional[Any]:
        url = config['api_url']
        method = config['api_method']
        params = {}
        payload = None

        if config['name'] == 'Заказы':
            params = {"dateFrom": date_str, "flag": 1}
        elif config['name'] == 'Остатки':
            params = {"dateFrom": date_str}
        elif config['name'] == 'Финансовые показатели':
            return self._fetch_finance_day(config, headers, date_str)

        max_attempts = 5
        for attempt in range(max_attempts):
            try:
                if method == 'GET':
                    resp = requests.get(url, headers=headers, params=params, timeout=120)
                else:
                    resp = requests.post(url, headers=headers, json=payload, timeout=120)

                if resp.status_code == 200:
                    return resp.json()
                elif resp.status_code == 429:
                    wait = 60 * (attempt + 1)
                    self.log(f"    ⚠ Лимит запросов (429), попытка {attempt+1}/{max_attempts}, ждём {wait} сек...")
                    time.sleep(wait)
                elif resp.status_code == 204:
                    return []
                elif resp.status_code in (502, 503, 504):
                    wait = 30 * (attempt + 1)
                    self.log(f"    ⚠ Ошибка шлюза {resp.status_code}, попытка {attempt+1}/{max_attempts}, ждём {wait} сек...")
                    time.sleep(wait)
                else:
                    self.log(f"    ❌ Ошибка {resp.status_code}: {resp.text[:200]}")
                    if attempt < max_attempts - 1:
                        time.sleep(10)
                    else:
                        return None
            except Exception as e:
                self.log(f"    ❌ Исключение при запросе: {e}")
                if attempt < max_attempts - 1:
                    time.sleep(10)
                else:
                    return None
        return None

    def _map_finance_v1_row_to_old_format(self, row: dict) -> dict:
        """Привести camelCase из нового finance-api к старым snake_case колонкам.

        Это нужно, чтобы недельные файлы оставались совместимыми со старой структурой:
        rr_dt, rrd_id, nm_id, supplier_oper_name, ppvz_for_pay и т.д.
        """
        mapping = {
            "reportId": "realizationreport_id",
            "dateFrom": "date_from",
            "dateTo": "date_to",
            "createDate": "create_dt",
            "currency": "currency_name",
            "rrdId": "rrd_id",
            "giId": "gi_id",
            "dlvPrc": "dlv_prc",
            "fixTariffDateFrom": "fix_tariff_date_from",
            "fixTariffDateTo": "fix_tariff_date_to",
            "subjectName": "subject_name",
            "nmId": "nm_id",
            "brandName": "brand_name",
            "vendorCode": "sa_name",
            "techSize": "ts_name",
            "sku": "barcode",
            "docTypeName": "doc_type_name",
            "retailPrice": "retail_price",
            "retailAmount": "retail_amount",
            "salePercent": "sale_percent",
            "commissionPercent": "commission_percent",
            "officeName": "office_name",
            "sellerOperName": "supplier_oper_name",
            "orderDt": "order_dt",
            "saleDt": "sale_dt",
            "rrDate": "rr_dt",
            "shkId": "shk_id",
            "retailPriceWithDisc": "retail_price_withdisc_rub",
            "deliveryAmount": "delivery_amount",
            "returnAmount": "return_amount",
            "deliveryService": "delivery_rub",
            "giBoxTypeName": "gi_box_type_name",
            "productDiscountForReport": "product_discount_for_report",
            "sellerPromo": "supplier_promo",
            "spp": "ppvz_spp_prc",
            "kvwBase": "ppvz_kvw_prc_base",
            "kvw": "ppvz_kvw_prc",
            "supRatingUp": "sup_rating_prc_up",
            "isKgvpV2": "is_kgvp_v2",
            "ppvzSalesCommission": "ppvz_sales_commission",
            "forPay": "ppvz_for_pay",
            "ppvzReward": "ppvz_reward",
            "acquiringFee": "acquiring_fee",
            "acquiringPercent": "acquiring_percent",
            "paymentProcessing": "payment_processing",
            "acquiringBank": "acquiring_bank",
            "vw": "ppvz_vw",
            "vwNds": "ppvz_vw_nds",
            "ppvzOfficeName": "ppvz_office_name",
            "ppvzOfficeId": "ppvz_office_id",
            "ppvzSupplierName": "ppvz_supplier_name",
            "ppvzSupplierInn": "ppvz_inn",
            "declarationNumber": "declaration_number",
            "bonusTypeName": "bonus_type_name",
            "stickerId": "sticker_id",
            "penalty": "penalty",
            "additionalPayment": "additional_payment",
            "rebillLogisticCost": "rebill_logistic_cost",
            "rebillLogisticOrg": "rebill_logistic_org",
            "paidStorage": "storage_fee",
            "deduction": "deduction",
            "paidAcceptance": "acceptance",
            "orderId": "order_id",
            "srid": "srid",
            "articleSubstitution": "article_substitution",
            "salePriceAffiliatedDiscountPrc": "sale_price_affiliated_discount_prc",
            "salePriceWholesaleDiscountPrc": "sale_price_wholesale_discount_prc",
            "cashbackAmount": "cashback_amount",
            "cashbackDiscount": "cashback_discount",
            "cashbackCommissionChange": "cashback_commission_change",
            "paymentSchedule": "payment_schedule",
            "deliveryMethod": "delivery_method",
            "sellerPromoId": "seller_promo_id",
            "sellerPromoDiscount": "seller_promo_discount",
            "loyaltyId": "loyalty_id",
            "loyaltyDiscount": "loyalty_discount",
            "uuidPromocode": "uuid_promocode",
            "salePricePromocodeDiscountPrc": "sale_price_promocode_discount_prc",
            "agencyVat": "agency_vat",
            "orderUid": "order_uid",
        }

        out = {}
        for key, value in row.items():
            out[mapping.get(key, key)] = value

        # В новом методе title есть, в старом его не было. Оставляем как полезное новое поле.
        if "title" in row:
            out["title"] = row.get("title")

        return out

    def _finance_wait_429(self, resp, attempt: int, default_seconds: int = 65) -> None:
        wait = default_seconds * attempt
        for header in ("X-Ratelimit-Retry", "X-RateLimit-Retry", "X-Ratelimit-Reset", "X-RateLimit-Reset", "Retry-After"):
            raw = resp.headers.get(header)
            if not raw:
                continue
            try:
                wait = max(wait, int(float(str(raw).strip())))
            except Exception:
                pass
        wait = min(max(wait, 30), 900)
        self.log(f"    ⚠ Finance API 429, попытка {attempt}, ждём {wait} сек. headers={dict(resp.headers)}")
        time.sleep(wait)

    def _fetch_finance_day(self, config: dict, headers: dict, date_str: str) -> List[dict]:
        """Новый финансовый метод WB с 15.07: POST /api/finance/v1/sales-reports/detailed."""
        url = config['api_url']
        all_items = []
        rrd_id = 0
        limit = 100000

        while True:
            payload = {
                "dateFrom": date_str,
                "dateTo": date_str,
                "limit": limit,
                "rrdId": rrd_id,
                "period": "daily"
            }

            max_attempts = 5
            page_loaded = False

            for attempt in range(1, max_attempts + 1):
                try:
                    resp = requests.post(url, headers=headers, json=payload, timeout=180)

                    if resp.status_code == 200:
                        data = resp.json()
                        if not data:
                            return all_items

                        if not isinstance(data, list):
                            self.log(f"    ❌ Finance API вернул не список: {str(data)[:1000]}")
                            return all_items

                        mapped = [self._map_finance_v1_row_to_old_format(item) for item in data]
                        all_items.extend(mapped)

                        last_rrd_id = 0
                        for item in data[::-1]:
                            try:
                                last_rrd_id = int(item.get("rrdId") or item.get("rrd_id") or 0)
                            except Exception:
                                last_rrd_id = 0
                            if last_rrd_id:
                                break

                        self.log(f"    ✅ Finance API: получено {len(data)} строк, rrdId={last_rrd_id}")

                        if len(data) < limit or last_rrd_id <= rrd_id:
                            return all_items

                        rrd_id = last_rrd_id
                        page_loaded = True
                        break

                    if resp.status_code == 204:
                        return all_items

                    if resp.status_code == 429:
                        self._finance_wait_429(resp, attempt, default_seconds=65)
                        continue

                    if resp.status_code in (401, 403):
                        self.log(
                            f"    ❌ Finance API {resp.status_code}: проверь токен. "
                            f"Новый метод требует токен категории Finance. Ответ: {resp.text[:1000]}"
                        )
                        return all_items

                    self.log(f"    ❌ Finance API HTTP {resp.status_code}: {resp.text[:1000]}")
                    if attempt < max_attempts:
                        time.sleep(15 * attempt)
                    else:
                        return all_items

                except Exception as e:
                    self.log(f"    ❌ Исключение Finance API: {e}")
                    if attempt < max_attempts:
                        time.sleep(15 * attempt)
                    else:
                        return all_items

            if not page_loaded:
                return all_items

    # ---------- Заказы ----------
    def update_orders(self, store_name: str) -> bool:
        self.log(f"\n📌 ОБНОВЛЕНИЕ: Заказы для магазина {store_name}")

        explicit_backfill = _parse_optional_date_env("WB_ORDERS_BACKFILL_FROM")
        if store_name == "FINICK" and explicit_backfill is None:
            history_start, history_end = self._get_last_completed_weeks_range(7)
            history_dates = [
                history_start + timedelta(days=i)
                for i in range((history_end - history_start).days + 1)
            ]
            # Текущую неделю не догружаем как историю. Добавляем только целевую дату ежедневного запуска.
            all_dates = history_dates + ([self.target_date] if self.target_date not in history_dates else [])
            self.log(
                f"📚 FINICK: проверяем/догружаем 7 завершённых недель истории заказов "
                f"{history_start:%Y-%m-%d} — {history_end:%Y-%m-%d}; "
                f"текущая неделя исключена. Отдельно целевая дата: {self.target_date:%Y-%m-%d}"
            )
        else:
            start_date, end_date = self._get_daily_or_backfill_range("WB_ORDERS_BACKFILL_FROM")
            self.log(f"📅 Диапазон заказов: {start_date:%Y-%m-%d} — {end_date:%Y-%m-%d}")
            all_dates = [start_date + timedelta(days=i) for i in range((end_date - start_date).days + 1)]

        weeks = defaultdict(list)
        for d in all_dates:
            week_start = self._get_week_start(datetime.combine(d, datetime.min.time()))
            weeks[week_start].append(d)

        config = self.reports_config['orders']
        api_key = self.api_keys[store_name][config['key_type']]
        headers = {"Authorization": api_key.strip()}

        total_loaded = 0
        for week_start, dates in sorted(weeks.items()):
            self.log(f"📅 Обработка недели, начинающейся {week_start.strftime('%Y-%m-%d')}")
            weekly_df = self._load_weekly_data(store_name, 'orders', week_start)
            if not weekly_df.empty:
                existing_dates = set(
                    pd.to_datetime(weekly_df['date'], errors='coerce').dt.date.dropna().unique()
                ) if 'date' in weekly_df.columns else set()
            else:
                existing_dates = set()

            dates_to_load = [d for d in sorted(set(dates)) if d not in existing_dates]
            if not dates_to_load:
                self.log("✅ Все нужные дни недели уже загружены")
                continue

            self.log(f"📅 Недостающие дни: {[d.strftime('%Y-%m-%d') for d in dates_to_load]}")
            new_data = []

            for idx, date in enumerate(dates_to_load):
                date_str = date.strftime('%Y-%m-%d')
                self.log(f"📅 Загрузка дня: {date_str}")
                data = self._make_request(config, headers, date_str)
                if data and isinstance(data, list):
                    day_df = pd.DataFrame(data)
                    if not day_df.empty:
                        # supplier/orders с flag=1 обычно возвращает день dateFrom, но дополнительно
                        # страхуемся и оставляем только строки нужной даты заказа.
                        if 'date' in day_df.columns:
                            parsed_date = pd.to_datetime(day_df['date'], errors='coerce').dt.date
                            exact_df = day_df.loc[parsed_date == date].copy()
                            if not exact_df.empty:
                                day_df = exact_df
                            day_df['date'] = pd.to_datetime(day_df['date'], errors='coerce').dt.strftime('%Y-%m-%d')

                        day_df['store'] = store_name
                        new_data.append(day_df)
                        total_loaded += len(day_df)
                        self.log(f"✅ Получено {len(day_df)} записей")
                    else:
                        self.log(f"ℹ️ Нет данных за {date_str}")
                else:
                    self.log(f"⚠️ Не удалось получить данные за {date_str}")

                if idx < len(dates_to_load) - 1:
                    time.sleep(self.delays['orders'])

            if new_data:
                new_df = pd.concat(new_data, ignore_index=True)
                if weekly_df.empty:
                    weekly_df = new_df
                else:
                    weekly_df = pd.concat([weekly_df, new_df], ignore_index=True)
                self._save_weekly_data(weekly_df, store_name, 'orders', week_start)
            else:
                self.log("ℹ️ Нет новых данных за неделю")

        self.log(f"✅ Заказы обновлены. Новых строк в этом запуске: {total_loaded}")
        return True

    # ---------- Остатки ----------
    def _create_warehouse_remains_task(self, headers: dict) -> Optional[str]:
        """Создать задачу нового отчёта warehouse_remains вместо deprecated /supplier/stocks."""
        url = "https://seller-analytics-api.wildberries.ru/api/v1/warehouse_remains"
        params = {
            "locale": "ru",
            "groupByBrand": "true",
            "groupBySubject": "true",
            "groupBySa": "true",
            "groupByNm": "true",
            "groupByBarcode": "true",
            "groupBySize": "true",
        }

        for attempt in range(1, 4):
            try:
                resp = requests.get(url, headers=headers, params=params, timeout=120)
                if resp.status_code == 200:
                    data = resp.json()
                    task_id = (data.get("data") or {}).get("taskId")
                    if not task_id:
                        self.log(f"❌ warehouse_remains не вернул taskId: {str(data)[:1000]}")
                        return None
                    self.log(f"✅ Задача warehouse_remains создана: {task_id}")
                    return task_id

                if resp.status_code == 429:
                    wait = self._rate_limit_wait_seconds(resp, default_seconds=90, max_seconds=900)
                    self.log(
                        f"⚠️ warehouse_remains create: 429, попытка {attempt}/3, "
                        f"ждём {wait} сек. headers={dict(resp.headers)}"
                    )
                    time.sleep(wait)
                    continue

                self.log(f"❌ warehouse_remains create HTTP {resp.status_code}: {resp.text[:1000]}")
                return None
            except Exception as e:
                self.log(f"❌ Ошибка warehouse_remains create: {e}")
                if attempt < 3:
                    time.sleep(30)
                else:
                    return None

        self.log("❌ warehouse_remains create: лимит 429 не снялся после 3 попыток, остатки пропущены")
        return None

    def _wait_warehouse_remains_task(self, task_id: str, headers: dict) -> bool:
        """Дождаться готовности отчёта warehouse_remains."""
        url = f"https://seller-analytics-api.wildberries.ru/api/v1/warehouse_remains/tasks/{task_id}/status"

        for attempt in range(1, 31):
            try:
                resp = requests.get(url, headers=headers, timeout=60)
                if resp.status_code == 200:
                    data = resp.json()
                    status = ((data.get("data") or {}).get("status") or "").lower()
                    self.log(f"⏳ warehouse_remains status: {status or 'unknown'} ({attempt}/30)")
                    if status == "done":
                        return True
                    if status in {"canceled", "cancelled", "failed", "error"}:
                        self.log(f"❌ warehouse_remains завершился статусом {status}: {str(data)[:1000]}")
                        return False
                    time.sleep(10)
                    continue

                if resp.status_code == 429:
                    wait = self._rate_limit_wait_seconds(resp, default_seconds=15, max_seconds=300)
                    self.log(
                        f"⚠️ warehouse_remains status: 429, попытка {attempt}/30, "
                        f"ждём {wait} сек. headers={dict(resp.headers)}"
                    )
                    time.sleep(wait)
                    continue

                self.log(f"❌ warehouse_remains status HTTP {resp.status_code}: {resp.text[:1000]}")
                return False
            except Exception as e:
                self.log(f"❌ Ошибка warehouse_remains status: {e}")
                time.sleep(10)

        self.log("❌ warehouse_remains не подготовился за 30 попыток")
        return False

    def _download_warehouse_remains_task(self, task_id: str, headers: dict) -> Optional[List[dict]]:
        """Скачать готовый отчёт warehouse_remains."""
        url = f"https://seller-analytics-api.wildberries.ru/api/v1/warehouse_remains/tasks/{task_id}/download"

        for attempt in range(1, 4):
            try:
                resp = requests.get(url, headers=headers, timeout=180)
                if resp.status_code == 200:
                    data = resp.json()
                    if isinstance(data, list):
                        self.log(f"✅ warehouse_remains скачан, строк верхнего уровня: {len(data)}")
                        return data
                    self.log(f"❌ warehouse_remains download вернул не список: {str(data)[:1000]}")
                    return None

                if resp.status_code == 204:
                    self.log("ℹ️ warehouse_remains download: нет данных")
                    return []

                if resp.status_code == 429:
                    wait = self._rate_limit_wait_seconds(resp, default_seconds=90, max_seconds=900)
                    self.log(
                        f"⚠️ warehouse_remains download: 429, попытка {attempt}/3, "
                        f"ждём {wait} сек. headers={dict(resp.headers)}"
                    )
                    time.sleep(wait)
                    continue

                self.log(f"❌ warehouse_remains download HTTP {resp.status_code}: {resp.text[:1000]}")
                return None
            except Exception as e:
                self.log(f"❌ Ошибка warehouse_remains download: {e}")
                if attempt < 3:
                    time.sleep(30)
                else:
                    return None

        self.log("❌ warehouse_remains download: лимит 429 не снялся после 3 попыток, остатки пропущены")
        return None

    def _warehouse_remains_to_stocks_df(self, data: List[dict], store_name: str, target_date_str: str) -> pd.DataFrame:
        """Преобразовать новый warehouse_remains в старую структуру листа Остатки."""
        rows = []
        now_str = datetime.now(pytz.timezone('Europe/Moscow')).strftime('%Y-%m-%d %H:%M:%S')

        inway_to_names = {"в пути до получателей", "в пути к клиенту", "в пути до клиента"}
        inway_from_names = {"в пути возвраты на склад wb", "в пути от клиента", "в пути возвраты"}
        total_names = {"всего находится на складах", "итого", "всего"}

        for item in data or []:
            warehouses = item.get("warehouses") or []

            for wh in warehouses:
                wh_name = str(wh.get("warehouseName", "") or "").strip()
                q = int(wh.get("quantity") or 0)
                wh_name_l = wh_name.lower()

                if wh_name_l in total_names:
                    continue

                if wh_name_l in inway_to_names:
                    row_qty, row_in_to, row_in_from = 0, q, 0
                elif wh_name_l in inway_from_names:
                    row_qty, row_in_to, row_in_from = 0, 0, q
                else:
                    row_qty, row_in_to, row_in_from = q, 0, 0

                rows.append({
                    'Дата последнего изменения': now_str,
                    'Склад': wh_name,
                    'Артикул продавца': item.get('vendorCode', ''),
                    'Артикул WB': item.get('nmId', ''),
                    'Баркод': item.get('barcode', ''),
                    'Доступно для продажи': row_qty,
                    'В пути к клиенту': row_in_to,
                    'В пути от клиента': row_in_from,
                    'Полное количество': row_qty + row_in_to + row_in_from,
                    'Категория': '',
                    'Предмет': item.get('subjectName', ''),
                    'Бренд': item.get('brand', ''),
                    'Размер': item.get('techSize', ''),
                    'Цена': 0,
                    'Скидка': 0,
                    'Договор поставки': '',
                    'Договор реализации': '',
                    'Код контракта': '',
                    'Дата запроса': target_date_str,
                    'Магазин': store_name,
                    'Дата сбора': now_str,
                })

            if not warehouses:
                rows.append({
                    'Дата последнего изменения': now_str,
                    'Склад': '',
                    'Артикул продавца': item.get('vendorCode', ''),
                    'Артикул WB': item.get('nmId', ''),
                    'Баркод': item.get('barcode', ''),
                    'Доступно для продажи': 0,
                    'В пути к клиенту': 0,
                    'В пути от клиента': 0,
                    'Полное количество': 0,
                    'Категория': '',
                    'Предмет': item.get('subjectName', ''),
                    'Бренд': item.get('brand', ''),
                    'Размер': item.get('techSize', ''),
                    'Цена': 0,
                    'Скидка': 0,
                    'Договор поставки': '',
                    'Договор реализации': '',
                    'Код контракта': '',
                    'Дата запроса': target_date_str,
                    'Магазин': store_name,
                    'Дата сбора': now_str,
                })

        return pd.DataFrame(rows)

    def update_stocks(self, store_name: str) -> bool:
        """Обновление остатков через новый warehouse_remains.

        Старый endpoint /api/v1/supplier/stocks deprecated и даёт постоянные 429/недоступность.
        Новый метод асинхронный: create task -> wait status -> download.
        """
        self.log(f"\\n📌 ОБНОВЛЕНИЕ: Остатки для магазина {store_name}")
        target_date = self.target_date
        target_date_str = target_date.strftime('%Y-%m-%d')
        week_start = self._get_week_start(datetime.combine(target_date, datetime.min.time()))

        weekly_df = self._load_weekly_data(store_name, 'stocks', week_start)
        if not weekly_df.empty and 'Дата запроса' in weekly_df.columns:
            existing_dates = set(pd.to_datetime(weekly_df['Дата запроса'], errors='coerce').dt.date.dropna().unique())
        else:
            existing_dates = set()

        if target_date in existing_dates:
            self.log(f"✅ Данные за {target_date_str} уже есть в недельном файле, пропускаем")
            return True

        self.log(f"📅 Загрузка остатков за {target_date_str} через warehouse_remains...")

        api_key = self.api_keys[store_name][self.reports_config['stocks']['key_type']]
        headers = {"Authorization": api_key.strip()}

        task_id = self._create_warehouse_remains_task(headers)
        if not task_id:
            self.log("⚠️ Остатки пропущены: не удалось создать задачу warehouse_remains")
            return False

        if not self._wait_warehouse_remains_task(task_id, headers):
            self.log("⚠️ Остатки пропущены: задача warehouse_remains не готова")
            return False

        raw_data = self._download_warehouse_remains_task(task_id, headers)
        if raw_data is None:
            self.log("⚠️ Остатки пропущены: не удалось скачать warehouse_remains")
            return False

        df_day = self._warehouse_remains_to_stocks_df(raw_data, store_name, target_date_str)
        if df_day.empty:
            self.log(f"ℹ️ Нет данных остатков за {target_date_str}")
            return True

        dedup_cols = ['Дата запроса', 'Артикул WB', 'Баркод', 'Склад']
        existing_cols = [c for c in dedup_cols if c in df_day.columns]
        if existing_cols:
            before = len(df_day)
            df_day = df_day.drop_duplicates(subset=existing_cols, keep='last')
            removed = before - len(df_day)
            if removed:
                self.log(f"🔍 Удалено дубликатов в дневных остатках: {removed}")

        if weekly_df.empty:
            weekly_df = df_day
        else:
            weekly_df = pd.concat([weekly_df, df_day], ignore_index=True)
            if existing_cols:
                before = len(weekly_df)
                weekly_df = weekly_df.drop_duplicates(subset=existing_cols, keep='last')
                removed = before - len(weekly_df)
                if removed:
                    self.log(f"🔍 Удалено дубликатов в недельном файле остатков: {removed}")

        self._save_weekly_data(weekly_df, store_name, 'stocks', week_start)
        self.log(f"✅ Остатки за {target_date_str} добавлены в недельный файл через warehouse_remains")
        return True

    # ---------- Финансовые показатели ----------
    def update_finance(self, store_name: str) -> bool:
        self.log(f"\n📌 ОБНОВЛЕНИЕ: Финансовые показатели для магазина {store_name} (ежедневный режим: только целевую дату)")
        config = self.reports_config['finance']
        start_date, end_date = self._get_daily_or_backfill_range("WB_FINANCE_BACKFILL_FROM")
        self.log(f"📅 Диапазон финансов: {start_date:%Y-%m-%d} — {end_date:%Y-%m-%d}")

        all_dates = [start_date + timedelta(days=i) for i in range((end_date - start_date).days + 1)]
        weeks = defaultdict(list)
        for d in all_dates:
            week_start = self._get_week_start(datetime.combine(d, datetime.min.time()))
            weeks[week_start].append(d)

        api_key = self.api_keys[store_name][config['key_type']]
        headers = {"Authorization": api_key.strip(), "Content-Type": "application/json"}
        total_loaded_rows = 0
        total_loaded_days = 0

        for week_start in sorted(weeks.keys()):
            dates = weeks[week_start]
            self.log(f"📅 Обработка финансовой недели, начинающейся {week_start.strftime('%Y-%m-%d')}")
            weekly_df = self._load_weekly_data(store_name, 'finance', week_start)

            if not weekly_df.empty and 'rr_dt' in weekly_df.columns:
                existing_dates = set(pd.to_datetime(weekly_df['rr_dt'], errors='coerce').dt.date.dropna().unique())
            else:
                existing_dates = set()

            dates_to_load = [d for d in dates if d not in existing_dates]
            if not dates_to_load:
                self.log("✅ Все дни финансовой недели уже загружены")
                continue

            self.log(f"📅 Недостающие дни финансовой недели: {[d.strftime('%Y-%m-%d') for d in dates_to_load]}")
            new_data = []

            for date in dates_to_load:
                date_str = date.strftime('%Y-%m-%d')
                self.log(f"📅 Загрузка финансового дня: {date_str}")
                day_data = self._fetch_finance_day(config, headers, date_str)
                if day_data:
                    day_df = pd.DataFrame(day_data)
                    day_df['store'] = store_name
                    if 'rr_dt' in day_df.columns:
                        day_df['rr_dt'] = pd.to_datetime(day_df['rr_dt'], errors='coerce').dt.strftime('%Y-%m-%d')
                    new_data.append(day_df)
                    total_loaded_rows += len(day_df)
                    total_loaded_days += 1
                    self.log(f"✅ Получено {len(day_df)} записей")
                else:
                    self.log(f"ℹ️ Нет финансовых данных за {date_str}")

                if date != dates_to_load[-1]:
                    time.sleep(self.delays['finance'])

            if new_data:
                new_df = pd.concat(new_data, ignore_index=True)
                if weekly_df.empty:
                    weekly_df = new_df
                else:
                    weekly_df = pd.concat([weekly_df, new_df], ignore_index=True)

                id_cols = [col for col in config.get('id_columns', []) if col in weekly_df.columns]
                if id_cols:
                    before = len(weekly_df)
                    weekly_df = weekly_df.drop_duplicates(subset=id_cols, keep='last')
                    removed = before - len(weekly_df)
                    if removed:
                        self.log(f"🔍 Удалено дубликатов в финансовой неделе: {removed}")

                self._save_weekly_data(weekly_df, store_name, 'finance', week_start)
            else:
                self.log("ℹ️ Нет новых финансовых данных за неделю")

            # Пауза между неделями только если впереди есть ещё недели с недостающими днями.
            # Основная защита лимита уже стоит между дневными запросами.

        self.log(f"✅ Финансовые показатели обновлены. Загружено дней: {total_loaded_days}, строк: {total_loaded_rows}")
        return True

    # ---------- Повторные попытки для поисковых запросов ----------
    def _retry_keyword_errors(self, store_name: str):
        if not self.keyword_errors:
            return

        self.log(f"\n🔄 Повторная загрузка для {len(self.keyword_errors)} ошибочных комбинаций...")
        api_key = self.api_keys[store_name]['promo']
        headers = {"Authorization": f"Bearer {api_key.strip()}", "Content-Type": "application/json"}
        url = self.reports_config['keywords']['api_url']
        filters = ["orders", "openCard", "addToCart"]

        # Группируем по дате и фильтру
        by_date_filter = defaultdict(list)
        for date_str, nm_id, filter_field in self.keyword_errors:
            by_date_filter[(date_str, filter_field)].append(nm_id)

        new_errors = []
        for (date_str, filter_field), nm_ids in by_date_filter.items():
            nm_ids = list(set(nm_ids))
            self.log(f"📅 {date_str} | Фильтр {filter_field} | артикулов: {len(nm_ids)}")

            batches = [nm_ids[i:i+50] for i in range(0, len(nm_ids), 50)]
            for batch in batches:
                date_obj = datetime.strptime(date_str, '%Y-%m-%d').date()
                past_date_str = (date_obj - timedelta(days=1)).strftime('%Y-%m-%d')
                payload = {
                    "currentPeriod": {"start": date_str, "end": date_str},
                    "pastPeriod": {"start": past_date_str, "end": past_date_str},
                    "nmIds": batch,
                    "topOrderBy": filter_field,
                    "includeSubstitutedSKUs": False,
                    "includeSearchTexts": True,
                    "orderBy": {"field": "avgPosition", "mode": "asc"},
                    "limit": 30
                }
                max_retries = 5
                for attempt in range(max_retries):
                    try:
                        resp = requests.post(url, headers=headers, json=payload, timeout=120)
                        if resp.status_code == 200:
                            data = resp.json()
                            items = data.get('data', {}).get('items', [])
                            if items:
                                batch_data = []
                                for item in items:
                                    text = item.get('text', '').strip()
                                    if not text:
                                        continue
                                    row = {
                                        "Дата": date_str,
                                        "Магазин": store_name,
                                        "Поисковый запрос": text,
                                        "Фильтр": filter_field,
                                        "Артикул WB": item.get("nmId", ""),
                                        "Предмет": item.get("subjectName", ""),
                                        "Бренд": item.get("brandName", ""),
                                        "Артикул продавца": item.get("vendorCode", ""),
                                        "Название товара": item.get("name", ""),
                                        "Рейтинг карточки": item.get("rating", 0),
                                        "Рейтинг отзывов": item.get("feedbackRating", 0),
                                        "Частота запросов": item.get("frequency", {}).get("current", 0),
                                        "Частота динамика %": item.get("frequency", {}).get("dynamics", 0),
                                        "Частота за неделю": item.get("weekFrequency", 0),
                                        "Медианная позиция": item.get("medianPosition", {}).get("current", 0),
                                        "Медианная позиция динамика %": item.get("medianPosition", {}).get("dynamics", 0),
                                        "Средняя позиция": item.get("avgPosition", {}).get("current", 0),
                                        "Средняя позиция динамика %": item.get("avgPosition", {}).get("dynamics", 0),
                                        "Переходы в карточку": item.get("openCard", {}).get("current", 0),
                                        "Переходы динамика %": item.get("openCard", {}).get("dynamics", 0),
                                        "% выше конкурентов (переходы)": item.get("openCard", {}).get("percentile", 0),
                                        "Добавления в корзину": item.get("addToCart", {}).get("current", 0),
                                        "Добавления динамика %": item.get("addToCart", {}).get("dynamics", 0),
                                        "% выше конкурентов (добавления)": item.get("addToCart", {}).get("percentile", 0),
                                        "Заказы": item.get("orders", {}).get("current", 0),
                                        "Заказы динамика %": item.get("orders", {}).get("dynamics", 0),
                                        "% выше конкурентов (заказы)": item.get("orders", {}).get("percentile", 0),
                                        "Конверсия в заказ %": item.get("cartToOrder", {}).get("current", 0),
                                        "Конверсия в заказ динамика %": item.get("cartToOrder", {}).get("dynamics", 0),
                                        "% выше конкурентов (конв. в заказ)": item.get("cartToOrder", {}).get("percentile", 0),
                                        "Конверсия в корзину %": item.get("openToCart", {}).get("current", 0),
                                        "Конверсия в корзину динамика %": item.get("openToCart", {}).get("dynamics", 0),
                                        "% выше конкурентов (конв. в корзину)": item.get("openToCart", {}).get("percentile", 0),
                                        "Видимость %": item.get("visibility", {}).get("current", 0),
                                        "Видимость динамика %": item.get("visibility", {}).get("dynamics", 0),
                                        "Есть рейтинг карточки": item.get("isCardRated", False),
                                        "Минимальная цена": item.get("price", {}).get("minPrice", 0),
                                        "Максимальная цена": item.get("price", {}).get("maxPrice", 0),
                                    }
                                    batch_data.append(row)
                                if batch_data:
                                    # Сохраняем в соответствующий недельный файл
                                    date_obj = datetime.strptime(date_str, '%Y-%m-%d')
                                    week_start = self._get_week_start(date_obj)
                                    weekly_df = self._load_weekly_data(store_name, 'keywords', week_start)
                                    new_df = pd.DataFrame(batch_data)
                                    if weekly_df.empty:
                                        weekly_df = new_df
                                    else:
                                        weekly_df = pd.concat([weekly_df, new_df], ignore_index=True)
                                    self._save_weekly_data(weekly_df, store_name, 'keywords', week_start)
                            break
                        elif resp.status_code in (429, 502, 503, 504):
                            wait = 60 * (attempt + 1)
                            self.log(f"    ⚠ Ошибка {resp.status_code}, повтор через {wait} сек...")
                            time.sleep(wait)
                        else:
                            self.log(f"    ❌ Ошибка {resp.status_code}: {resp.text[:1000]}, пропускаем")
                            for nm_id in batch:
                                new_errors.append((date_str, nm_id, filter_field))
                            break
                    except Exception as e:
                        self.log(f"    ❌ Исключение: {e}")
                        if attempt < max_retries - 1:
                            time.sleep(10)
                        else:
                            for nm_id in batch:
                                new_errors.append((date_str, nm_id, filter_field))
                        break
                time.sleep(30)

        self.keyword_errors = new_errors
        if self.keyword_errors:
            self.log(f"⚠️ После повторов осталось {len(self.keyword_errors)} ошибок")
        else:
            self.log("✅ Все ошибки устранены")

    # ---------- Позиции по ключам ----------
    def _update_keywords_for_date(self, store_name: str, target_date: datetime.date, articles: List[int]) -> bool:
        """Загрузить поисковые запросы за одну дату по переданному списку nmId."""
        target_date_str = target_date.strftime('%Y-%m-%d')
        self.log("")
        self.log(f"📌 ОБНОВЛЕНИЕ: Позиции по ключам для магазина {store_name} за {target_date_str}")

        week_start = self._get_week_start(datetime.combine(target_date, datetime.min.time()))
        self.log(f"📅 Неделя начинается: {week_start.strftime('%Y-%m-%d')}")

        weekly_df = self._load_weekly_data(store_name, 'keywords', week_start)

        existing_keys = set()
        if not weekly_df.empty and 'Дата' in weekly_df.columns:
            day_df = weekly_df[weekly_df['Дата'].astype(str) == target_date_str].copy()
            if not day_df.empty and 'Артикул WB' in day_df.columns and 'Фильтр' in day_df.columns:
                day_df['Артикул WB'] = pd.to_numeric(day_df['Артикул WB'], errors='coerce')
                day_df = day_df.dropna(subset=['Артикул WB'])
                day_df['Артикул WB'] = day_df['Артикул WB'].astype(int)
                for _, row in day_df.iterrows():
                    existing_keys.add((target_date_str, int(row['Артикул WB']), str(row['Фильтр'])))
                self.log(f"🔍 В недельном файле найдено {len(existing_keys)} комбинаций за {target_date_str}")
            else:
                self.log(f"ℹ️ За {target_date_str} в недельном файле записей нет")

        filters = ["orders", "openCard", "addToCart"]

        missing_by_filter = {f: [] for f in filters}
        for nm_id in articles:
            for f in filters:
                if (target_date_str, nm_id, f) not in existing_keys:
                    missing_by_filter[f].append(nm_id)

        total_missing = sum(len(v) for v in missing_by_filter.values())
        if total_missing == 0:
            self.log(f"✅ Все данные за {target_date_str} уже загружены полностью.")
            return True

        self.log(
            f"📅 Необходимо загрузить комбинаций: {total_missing}; "
            f"артикулов всего: {len(articles)}"
        )

        api_key = self.api_keys[store_name]['promo']
        headers = {"Authorization": f"Bearer {api_key.strip()}", "Content-Type": "application/json"}
        url = self.reports_config['keywords']['api_url']

        self.keyword_errors = []
        new_data = []

        for filter_idx, filter_field in enumerate(filters, 1):
            nm_ids_for_filter = sorted(set(missing_by_filter.get(filter_field, [])))
            if not nm_ids_for_filter:
                self.log(f"✅ Фильтр {filter_field}: всё уже загружено")
                continue

            batches = [nm_ids_for_filter[i:i+50] for i in range(0, len(nm_ids_for_filter), 50)]
            self.log(f"🔍 Фильтр {filter_field}: {len(nm_ids_for_filter)} артикулов, батчей: {len(batches)}")

            for batch_idx, batch in enumerate(batches, 1):
                self.log(f"  📦 Батч {batch_idx}/{len(batches)}: {len(batch)} артикулов")
                past_date_str = (target_date - timedelta(days=1)).strftime('%Y-%m-%d')
                payload = {
                    "currentPeriod": {"start": target_date_str, "end": target_date_str},
                    "pastPeriod": {"start": past_date_str, "end": past_date_str},
                    "nmIds": batch,
                    "topOrderBy": filter_field,
                    "includeSubstitutedSKUs": False,
                    "includeSearchTexts": True,
                    "orderBy": {"field": "avgPosition", "mode": "asc"},
                    "limit": 30
                }

                max_retries = 5
                success = False
                for attempt in range(max_retries):
                    try:
                        resp = requests.post(url, headers=headers, json=payload, timeout=120)
                        if resp.status_code == 200:
                            data = resp.json()
                            items = data.get('data', {}).get('items', [])
                            batch_data = []
                            for item in items:
                                text_value = item.get('text', '').strip()
                                if not text_value:
                                    continue
                                row = {
                                    "Дата": target_date_str,
                                    "Магазин": store_name,
                                    "Поисковый запрос": text_value,
                                    "Фильтр": filter_field,
                                    "Артикул WB": item.get("nmId", ""),
                                    "Предмет": item.get("subjectName", ""),
                                    "Бренд": item.get("brandName", ""),
                                    "Артикул продавца": item.get("vendorCode", ""),
                                    "Название товара": item.get("name", ""),
                                    "Рейтинг карточки": item.get("rating", 0),
                                    "Рейтинг отзывов": item.get("feedbackRating", 0),
                                    "Частота запросов": item.get("frequency", {}).get("current", 0),
                                    "Частота динамика %": item.get("frequency", {}).get("dynamics", 0),
                                    "Частота за неделю": item.get("weekFrequency", 0),
                                    "Медианная позиция": item.get("medianPosition", {}).get("current", 0),
                                    "Медианная позиция динамика %": item.get("medianPosition", {}).get("dynamics", 0),
                                    "Средняя позиция": item.get("avgPosition", {}).get("current", 0),
                                    "Средняя позиция динамика %": item.get("avgPosition", {}).get("dynamics", 0),
                                    "Переходы в карточку": item.get("openCard", {}).get("current", 0),
                                    "Переходы динамика %": item.get("openCard", {}).get("dynamics", 0),
                                    "% выше конкурентов (переходы)": item.get("openCard", {}).get("percentile", 0),
                                    "Добавления в корзину": item.get("addToCart", {}).get("current", 0),
                                    "Добавления динамика %": item.get("addToCart", {}).get("dynamics", 0),
                                    "% выше конкурентов (добавления)": item.get("addToCart", {}).get("percentile", 0),
                                    "Заказы": item.get("orders", {}).get("current", 0),
                                    "Заказы динамика %": item.get("orders", {}).get("dynamics", 0),
                                    "% выше конкурентов (заказы)": item.get("orders", {}).get("percentile", 0),
                                    "Конверсия в заказ %": item.get("cartToOrder", {}).get("current", 0),
                                    "Конверсия в заказ динамика %": item.get("cartToOrder", {}).get("dynamics", 0),
                                    "% выше конкурентов (конв. в заказ)": item.get("cartToOrder", {}).get("percentile", 0),
                                    "Конверсия в корзину %": item.get("openToCart", {}).get("current", 0),
                                    "Конверсия в корзину динамика %": item.get("openToCart", {}).get("dynamics", 0),
                                    "% выше конкурентов (конв. в корзину)": item.get("openToCart", {}).get("percentile", 0),
                                    "Видимость %": item.get("visibility", {}).get("current", 0),
                                    "Видимость динамика %": item.get("visibility", {}).get("dynamics", 0),
                                    "Есть рейтинг карточки": item.get("isCardRated", False),
                                    "Минимальная цена": item.get("price", {}).get("minPrice", 0),
                                    "Максимальная цена": item.get("price", {}).get("maxPrice", 0),
                                }
                                batch_data.append(row)

                            if batch_data:
                                new_data.append(pd.DataFrame(batch_data))
                            self.log(f"    ✓ {len(items)} записей")
                            success = True
                            break

                        elif resp.status_code == 429:
                            wait = 60 * (attempt + 1)
                            self.log(f"    ⚠ Лимит 429, попытка {attempt+1}/{max_retries}, ждём {wait} сек...")
                            time.sleep(wait)
                        elif resp.status_code in (502, 503, 504):
                            wait = 30 * (attempt + 1)
                            self.log(f"    ⚠ Ошибка шлюза {resp.status_code}, попытка {attempt+1}/{max_retries}, ждём {wait} сек...")
                            time.sleep(wait)
                        else:
                            self.log(f"    ❌ Ошибка {resp.status_code}: {resp.text[:1000]}")
                            break

                    except Exception as e:
                        self.log(f"    ❌ Исключение: {e}")
                        if attempt < max_retries - 1:
                            time.sleep(10)
                        else:
                            break

                if not success:
                    for nm_id in batch:
                        self.keyword_errors.append((target_date_str, nm_id, filter_field))

                # Пауза между запросами к search-texts API.
                time.sleep(30)

            # Дополнительная пауза между фильтрами.
            if filter_idx < len(filters):
                self.log("    ⏳ Пауза 30 сек между фильтрами...")
                time.sleep(30)

        if new_data:
            new_df = pd.concat(new_data, ignore_index=True)
            if weekly_df.empty:
                weekly_df = new_df
            else:
                weekly_df = pd.concat([weekly_df, new_df], ignore_index=True)

            self._save_weekly_data(weekly_df, store_name, 'keywords', week_start)
            self.log(f"✅ Данные за {target_date_str} успешно добавлены в недельный файл")
        else:
            self.log(f"ℹ️ Нет новых данных для {target_date_str}")

        if self.keyword_errors:
            self._retry_keyword_errors(store_name)

        if self.keyword_errors:
            self.log(f"❌ Поисковые запросы не собраны полностью за {target_date_str}: осталось ошибок {len(self.keyword_errors)}")
            return False

        return True

    def update_keywords(self, store_name: str) -> bool:
        self.log("")
        self.log(f"📌 ОБНОВЛЕНИЕ: Позиции по ключам для магазина {store_name}")

        end_date = self.target_date

        # v24:
        # Ежедневный запуск должен собирать только целевую дату.
        # Историческую догрузку с 2026-06-01 включаем только явно через env:
        # WB_KEYWORDS_BACKFILL_FROM=YYYY-MM-DD
        backfill_from = _parse_optional_date_env("WB_KEYWORDS_BACKFILL_FROM")
        if backfill_from:
            start_date = backfill_from
            self.log(
                f"📅 Ручная догрузка keywords включена: {start_date:%Y-%m-%d} — {end_date:%Y-%m-%d}. "
                f"Если данные за день уже есть, день будет пропущен."
            )
        else:
            start_date = end_date
            self.log(
                f"📅 Ежедневный режим keywords: загружаем только целевую дату {end_date:%Y-%m-%d}. "
                f"Историческая догрузка отключена."
            )

        if end_date < start_date:
            self.log(f"⏭️ Конечная дата {end_date:%Y-%m-%d} раньше старта {start_date:%Y-%m-%d}; пропускаем")
            return True

        # Все категории/предметы. Артикулы берём из заказов с 1 июня,
        # чтобы в keywords попадали новые категории, но сам daily-цикл не гонял историю каждый день.
        articles = self._get_articles_by_subjects(
            store_name,
            self.target_subjects,
            min_order_date=KEYWORDS_DEFAULT_START_DATE,
        )
        if not articles:
            self.log("⚠️ Не найдено артикулов из заказов. Отчёт будет пропущен.")
            return False

        self.log(f"📦 Актуальных артикулов для keywords: {len(articles)}")

        overall_success = True
        total_days = (end_date - start_date).days + 1

        for idx in range(total_days):
            target_date = start_date + timedelta(days=idx)
            ok = self._update_keywords_for_date(store_name, target_date, articles)
            if not ok:
                overall_success = False

            if idx < total_days - 1:
                self.log("⏳ Пауза 30 секунд перед следующей датой keywords...")
                time.sleep(30)

        return overall_success

    # ---------- Воронка продаж ----------
    def update_funnel(self, store_name: str) -> bool:
        """Воронка: для FINICK — бесплатный v3 API; для остальных магазинов сохраняем прежний CSV/Jam-метод."""
        if store_name == "FINICK":
            return self._update_funnel_finick_free(store_name)
        return self._update_funnel_jam(store_name)

    def _update_funnel_finick_free(self, store_name: str) -> bool:
        """FINICK: бесплатная воронка через /api/analytics/v3/sales-funnel/products.

        v38:
        - при первом запуске догружаем 7 полностью завершённых недель + target_date;
        - сохраняем прогресс ПОСЛЕ КАЖДОГО успешно загруженного дня;
        - 502/503/504 считаем временными ошибками и повторяем запрос;
        - если отдельная дата после повторов не загрузилась, продолжаем остальные даты;
        - следующий запуск увидит сохранённые даты и повторит только пропущенные.
        """
        self.log(f"\n📌 ОБНОВЛЕНИЕ: Воронка продаж для магазина {store_name} (FREE v3, resume-safe)")
        config = self.reports_config['funnel']
        key = f"Отчёты/{config['folder']}/{store_name}/{config['filename']}"

        if self.s3.file_exists(key):
            df_existing = self.s3.read_excel(key, sheet_name=0)
        else:
            df_existing = pd.DataFrame()
            self.log("⚠️ Файл воронки FINICK не найден, будет создан")

        existing_dates: Set[datetime.date] = set()
        if not df_existing.empty and 'dt' in df_existing.columns:
            existing_dates = set(
                pd.to_datetime(df_existing['dt'], errors='coerce').dt.date.dropna().unique()
            )

        explicit_backfill = _parse_optional_date_env("WB_FUNNEL_BACKFILL_FROM")
        if explicit_backfill:
            requested_dates = [
                explicit_backfill + timedelta(days=i)
                for i in range((self.target_date - explicit_backfill).days + 1)
            ] if explicit_backfill <= self.target_date else [self.target_date]
            self.log(
                f"📚 FINICK funnel: ручная догрузка {requested_dates[0]:%Y-%m-%d} — "
                f"{requested_dates[-1]:%Y-%m-%d}"
            )
        else:
            history_start, history_end = self._get_last_completed_weeks_range(7)
            history_dates = [
                history_start + timedelta(days=i)
                for i in range((history_end - history_start).days + 1)
            ]
            requested_dates = history_dates + ([self.target_date] if self.target_date not in history_dates else [])
            self.log(
                f"📚 FINICK funnel: 7 завершённых недель {history_start:%Y-%m-%d} — "
                f"{history_end:%Y-%m-%d}; текущая неделя исключена. "
                f"Отдельно target_date={self.target_date:%Y-%m-%d}"
            )

        dates_to_load = [d for d in requested_dates if d not in existing_dates]
        if not dates_to_load:
            self.log("✅ Воронка FINICK уже содержит все требуемые исторические дни и целевую дату")
            return True

        self.log(
            f"📅 Воронка FINICK: отсутствует {len(dates_to_load)} дней. "
            f"Будут загружены только недостающие даты."
        )

        api_key = self.api_keys[store_name][config['key_type']]
        headers = {
            "Authorization": f"Bearer {api_key.strip()}",
            "Content-Type": "application/json",
        }
        url = "https://seller-analytics-api.wildberries.ru/api/analytics/v3/sales-funnel/products"

        total_added_rows = 0
        failed_dates: List[str] = []

        for date_idx, target_date in enumerate(sorted(dates_to_load), start=1):
            date_str = target_date.strftime("%Y-%m-%d")
            self.log(f"📅 Funnel FREE {date_idx}/{len(dates_to_load)}: {date_str}")
            offset = 0
            page_size = 1000
            day_data_rows: List[dict] = []
            day_failed = False

            while True:
                payload = {
                    "selectedPeriod": {"start": date_str, "end": date_str},
                    "nmIds": [],
                    "brandNames": [],
                    "subjectIds": [],
                    "tagIds": [],
                    "skipDeletedNm": False,
                    "orderBy": {"field": "openCard", "mode": "desc"},
                    "limit": page_size,
                    "offset": offset,
                }

                response_data = None
                max_attempts = 5
                for attempt in range(1, max_attempts + 1):
                    try:
                        resp = requests.post(url, headers=headers, json=payload, timeout=150)
                        if resp.status_code == 200:
                            response_data = resp.json()
                            break

                        if resp.status_code == 429:
                            wait = min(20 * attempt, 120)
                            self.log(
                                f"    ⚠️ Funnel FREE 429 за {date_str}, "
                                f"попытка {attempt}/{max_attempts}, ждём {wait} сек"
                            )
                            time.sleep(wait)
                            continue

                        if resp.status_code in (502, 503, 504):
                            wait = min(30 * attempt, 150)
                            self.log(
                                f"    ⚠️ Funnel FREE HTTP {resp.status_code} за {date_str}, "
                                f"попытка {attempt}/{max_attempts}, ждём {wait} сек. "
                                f"Ответ: {resp.text[:500]}"
                            )
                            time.sleep(wait)
                            continue

                        self.log(
                            f"    ❌ Funnel FREE HTTP {resp.status_code} за {date_str}: "
                            f"{resp.text[:1500]}"
                        )
                        day_failed = True
                        break

                    except Exception as e:
                        if attempt == max_attempts:
                            self.log(
                                f"    ❌ Funnel FREE: ошибка запроса за {date_str} "
                                f"после {max_attempts} попыток: {e}"
                            )
                            day_failed = True
                            break
                        wait = min(15 * attempt, 90)
                        self.log(
                            f"    ⚠️ Funnel FREE: исключение за {date_str}, "
                            f"попытка {attempt}/{max_attempts}: {e}; ждём {wait} сек"
                        )
                        time.sleep(wait)

                if day_failed:
                    break
                if response_data is None:
                    day_failed = True
                    break

                data_obj = response_data.get("data", response_data) if isinstance(response_data, dict) else {}
                products = data_obj.get("products", []) if isinstance(data_obj, dict) else []
                currency = data_obj.get("currency", "RUB") if isinstance(data_obj, dict) else "RUB"

                if not products:
                    break

                for item in products:
                    product = item.get("product") or {}
                    statistic = item.get("statistic") or {}
                    selected = statistic.get("selected") or {}
                    conversions = selected.get("conversions") or {}

                    day_data_rows.append({
                        "nmID": product.get("nmId", ""),
                        "dt": date_str,
                        "openCardCount": selected.get("openCount", 0),
                        "addToCartCount": selected.get("cartCount", 0),
                        "ordersCount": selected.get("orderCount", 0),
                        "ordersSumRub": selected.get("orderSum", 0),
                        "buyoutsCount": selected.get("buyoutCount", 0),
                        "buyoutsSumRub": selected.get("buyoutSum", 0),
                        "cancelCount": selected.get("cancelCount", 0),
                        "cancelSumRub": selected.get("cancelSum", 0),
                        "addToCartConversion": selected.get(
                            "addToCartConversion",
                            selected.get("openToCartPercent", 0),
                        ),
                        "cartToOrderConversion": selected.get(
                            "cartToOrderConversion",
                            conversions.get("cartToOrderPercent", 0),
                        ),
                        "buyoutPercent": selected.get(
                            "buyoutPercent",
                            conversions.get("buyoutPercent", 0),
                        ),
                        "addToWishlist": selected.get(
                            "addToWishlist",
                            selected.get("addToWishlistCount", 0),
                        ),
                        "currency": currency,
                        "title": product.get("title", ""),
                        "vendorCode": product.get("vendorCode", ""),
                        "brandName": product.get("brandName", ""),
                        "subjectId": product.get("subjectId", ""),
                        "subjectName": product.get("subjectName", ""),
                        "productRating": product.get("productRating", 0),
                        "feedbackRating": product.get("feedbackRating", 0),
                        "store": store_name,
                    })

                if len(products) < page_size:
                    break
                offset += page_size
                time.sleep(20)

            if day_failed:
                failed_dates.append(date_str)
                self.log(
                    f"    ❌ Funnel FREE за {date_str} не загружен. "
                    f"Переходим к следующей дате; следующий запуск повторит этот день."
                )
                if date_idx < len(dates_to_load):
                    time.sleep(30)
                continue

            # Критично для resume: сохраняем каждый успешно завершённый день сразу.
            day_rows_count = len(day_data_rows)
            if day_data_rows:
                day_df = pd.DataFrame(day_data_rows)
                if df_existing.empty:
                    combined = day_df
                else:
                    combined = pd.concat([df_existing, day_df], ignore_index=True)

                if 'dt' in combined.columns:
                    combined['dt'] = pd.to_datetime(combined['dt'], errors='coerce').dt.strftime('%Y-%m-%d')
                dedup_cols = [c for c in ['dt', 'nmID'] if c in combined.columns]
                if dedup_cols:
                    combined = combined.drop_duplicates(subset=dedup_cols, keep='last')
                if 'dt' in combined.columns:
                    combined = combined.sort_values(['dt', 'nmID'] if 'nmID' in combined.columns else ['dt'])

                self.s3.write_excel(key, combined, sheet_name=config['name'])
                df_existing = combined
                existing_dates.add(target_date)
                total_added_rows += day_rows_count
                self.log(
                    f"    ✅ Funnel FREE за {date_str}: {day_rows_count} товарных строк; "
                    f"день сразу сохранён в {key}"
                )
            else:
                self.log(f"    ℹ️ Funnel FREE за {date_str}: API вернул 0 товарных строк")

            if date_idx < len(dates_to_load):
                time.sleep(20)

        self.log(
            f"📊 Funnel FINICK: за запуск добавлено {total_added_rows} строк; "
            f"успешно сохранённых дат в файле: {len(existing_dates)}"
        )

        if failed_dates:
            self.log(
                f"❌ Funnel FINICK завершён с пропусками. Не загружены даты: "
                f"{', '.join(failed_dates)}. Следующий запуск повторит только их."
            )
            return False

        self.log("✅ Funnel FINICK: все требуемые даты загружены")
        return True

    # ---------- Воронка продаж ----------
    def _update_funnel_jam(self, store_name: str) -> bool:
        self.log(f"\n📌 ОБНОВЛЕНИЕ: Воронка продаж для магазина {store_name}")
        config = self.reports_config['funnel']
        key = f"Отчёты/{config['folder']}/{store_name}/{config['filename']}"
        if self.s3.file_exists(key):
            df_existing = self.s3.read_excel(key, sheet_name=0)
            if not df_existing.empty:
                date_col = config['date_column']
                if date_col in df_existing.columns:
                    df_existing[date_col] = pd.to_datetime(df_existing[date_col])
                    max_date = df_existing[date_col].max()
                    # Проверяем, есть ли данные за целевую дату
                    target_date = self.target_date
                    if max_date and max_date.date() >= target_date:
                        self.log("✅ Данные воронки уже актуальны")
                        return True
                    else:
                        self.log(f"⚠️ Данные воронки устарели: последняя дата {max_date.date()}, требуется обновление до {target_date}")
                else:
                    self.log("⚠️ В файле воронки нет колонки с датой, требуется обновление")
            else:
                self.log("⚠️ Файл воронки пуст, требуется обновление")
        else:
            self.log("⚠️ Файл воронки не найден, будет создан")

        self.log("🔄 Запуск формирования отчёта воронки...")
        api_key = self.api_keys[store_name][config['key_type']]
        headers = {"Authorization": f"Bearer {api_key.strip()}", "Content-Type": "application/json"}

        start_date, end_date = self._get_daily_or_backfill_range("WB_FUNNEL_BACKFILL_FROM")
        self.log(f"📅 Диапазон воронки: {start_date:%Y-%m-%d} — {end_date:%Y-%m-%d}")
        start_str = start_date.strftime("%Y-%m-%d")
        end_str = end_date.strftime("%Y-%m-%d")
        report_id = str(uuid.uuid4())

        create_payload = {
            "id": report_id,
            "reportType": "DETAIL_HISTORY_REPORT",
            "userReportName": "Воронка продаж",
            "params": {
                "nmIDs": [],
                "subjectIds": [],
                "brandNames": [],
                "tagIds": [],
                "startDate": start_str,
                "endDate": end_str,
                "timezone": "Europe/Moscow",
                "aggregationLevel": "day",
                "skipDeletedNm": False
            }
        }

        try:
            resp = requests.post(config['api_url'], headers=headers, json=create_payload, timeout=60)
            if resp.status_code != 200:
                self.log(f"❌ Ошибка создания отчёта: HTTP {resp.status_code}: {resp.text[:1500]}")
                return False
        except Exception as e:
            self.log(f"❌ Ошибка соединения: {e}")
            return False

        self.log("⏳ Ожидание готовности отчёта (до 30 попыток)...")
        download_url = f"https://seller-analytics-api.wildberries.ru/api/v2/nm-report/downloads/file/{report_id}"
        for attempt in range(1, 31):
            time.sleep(30)
            try:
                resp = requests.get(download_url, headers=headers, stream=True, timeout=120)
                if resp.status_code == 200:
                    self.log("✅ Отчёт готов, скачиваю...")
                    zip_data = io.BytesIO(resp.content)
                    with zipfile.ZipFile(zip_data, 'r') as zf:
                        for name in zf.namelist():
                            with zf.open(name) as f:
                                content = f.read()
                                for enc in ['utf-8', 'utf-8-sig', 'cp1251', 'windows-1251']:
                                    try:
                                        text = content.decode(enc)
                                        break
                                    except:
                                        continue
                                else:
                                    self.log("⚠️ Не удалось декодировать файл")
                                    continue
                                for sep in [',', ';', '\t']:
                                    try:
                                        df = pd.read_csv(io.StringIO(text), delimiter=sep)
                                        if len(df.columns) > 1:
                                            break
                                    except:
                                        continue
                                else:
                                    self.log("⚠️ Не удалось прочитать CSV")
                                    continue
                                df['store'] = store_name
                                if 'dt' in df.columns:
                                    df['dt'] = pd.to_datetime(df['dt']).dt.strftime('%Y-%m-%d')
                                self.s3.write_excel(key, df, sheet_name=config['name'])
                                self.log(f"✅ Воронка продаж сохранена: {key}")
                                return True
                elif resp.status_code == 202:
                    self.log(f"⏳ Отчёт ещё не готов, попытка {attempt}/30")
                else:
                    self.log(f"⚠️ Статус {resp.status_code}")
            except Exception as e:
                self.log(f"⚠️ Ошибка при скачивании: {e}")

        self.log("❌ Не удалось получить отчёт воронки")
        return False

    # ---------- Реклама (получение данных напрямую из API) ----------
    def update_adverts(self, store_name: str) -> bool:
        """
        Обновление рекламы с rolling-refresh последних 14 дат статистики.

        Правило v40:
        - каждый ежедневный запуск заново запрашивает target_date и 13 предыдущих дат;
        - поэтому одна дата статистики получает до 14 последовательных снимков
          (снимок №1, №2, ... №14) в разные дни запуска;
        - недельные файлы и Анализ рекламы.xlsx хранят ПОСЛЕДНЕЕ актуальное значение WB;
        - История_рекламы_14дней.xlsx хранит ВСЕ 14 снимков и позволяет измерять лаг/пересчёты WB.

        Для FINICK первоначальная догрузка 7 завершённых недель сохраняется,
        но лаг-история записывает только снимки №1..14, чтобы backfill не выдавать
        за реальные ежедневные наблюдения.
        """
        self.log(f"\n📌 ОБНОВЛЕНИЕ: Реклама для магазина {store_name}")
        config = self.reports_config['adverts']

        # Для FINICK при первом запуске догружаем 7 полностью завершённых недель,
        # не включая текущую. Отмечаем недели, для которых недельного файла ещё нет.
        finick_missing_week_ranges: List[Tuple[datetime.date, datetime.date]] = []
        if store_name == "FINICK" and not _parse_optional_date_env("WB_ADVERTS_BACKFILL_FROM"):
            history_start, history_end = self._get_last_completed_weeks_range(7)
            for week_idx in range(7):
                ws = history_start + timedelta(days=week_idx * 7)
                we = ws + timedelta(days=6)
                week_key = self._get_weekly_key(
                    store_name,
                    'adverts',
                    datetime.combine(ws, datetime.min.time()),
                )
                if not self.s3.file_exists(week_key):
                    finick_missing_week_ranges.append((ws, we))

            if finick_missing_week_ranges:
                self.log(
                    f"📚 FINICK: требуется догрузить рекламную историю за "
                    f"{len(finick_missing_week_ranges)} из 7 завершённых недель "
                    f"({history_start:%Y-%m-%d} — {history_end:%Y-%m-%d})."
                )
            else:
                self.log("✅ FINICK: недельные файлы рекламы за 7 завершённых недель уже есть")

        # v40: рекламу НИКОГДА не пропускаем только потому, что файл уже актуален.
        # WB пересчитывает исторические показатели, поэтому один и тот же день
        # нужно повторно получать 14 ежедневных запусков подряд.
        analytics_key = f"Отчёты/{config['folder']}/{store_name}/Анализ рекламы.xlsx"
        history_key = f"Отчёты/{config['folder']}/{store_name}/История_рекламы_14дней.xlsx"
        if self.s3.file_exists(analytics_key):
            self.log(
                "🔄 Анализ рекламы уже существует, но rolling-refresh обязателен: "
                "повторно запрашиваем последние 14 дат, чтобы поймать пересчёты WB"
            )
        else:
            self.log("⚠️ Аналитический файл не найден, будет создан")
        if self.s3.file_exists(history_key):
            self.log("📚 История_рекламы_14дней.xlsx найдена — добавим очередной ежедневный снимок")
        else:
            self.log("📚 История_рекламы_14дней.xlsx не найдена — создадим и начнём накопление 14 снимков")

        api_key = self.api_keys[store_name][config['key_type']]
        headers = {"Authorization": f"Bearer {api_key.strip()}"}

        # 1. Получаем список кампаний.
        # v40: статус 7 (завершённые) нужен ВСЕМ магазинам: кампания могла завершиться
        # вчера, но её показатели за последние 14 дней WB ещё может пересчитывать.
        self.log("📋 Запрос списка рекламных кампаний...")
        all_adverts = []
        statuses = "7,9,11"
        for payment_type in ['cpm', 'cpc']:
            url = f"{config['api_url']}?statuses={statuses}&payment_type={payment_type}"
            try:
                resp = requests.get(url, headers=headers, timeout=30)
                if resp.status_code == 200:
                    data = resp.json()
                    adverts = data.get('adverts', [])
                    all_adverts.extend(adverts)
                    self.log(f"✅ Получено кампаний для {payment_type}: {len(adverts)}")
                else:
                    self.log(f"⚠️ Не удалось получить список кампаний для {payment_type}: {resp.status_code}")
                time.sleep(0.5)
            except Exception as e:
                self.log(f"❌ Ошибка при запросе кампаний: {e}")
                return False

        if not all_adverts:
            self.log("❌ Не получено ни одной кампании. Отчёт пропущен.")
            return False

        self.log(f"✅ Всего получено кампаний: {len(all_adverts)}")

        # 2. Извлекаем информацию о кампаниях (ID, название, предмет, артикул, тип оплаты, статус, ставки и т.д.)
        campaign_ids = []
        campaign_info = {}  # id -> {'name': ..., 'subject': ..., 'article': ..., 'payment_type': ..., 'bid_type': ..., 'status': ..., 'search_bid': ..., 'recommendations_bid': ...}
        campaigns_list_rows = []  # для сохранения в лист Список_кампаний

        for adv in all_adverts:
            adv_id = adv.get('id')
            if not adv_id:
                continue
            settings = adv.get('settings', {})
            name = settings.get('name', '')
            payment_type = settings.get('payment_type', '')
            bid_type = adv.get('bid_type', '')
            status = 'Активна' if adv.get('status') == 9 else 'На паузе' if adv.get('status') == 11 else str(adv.get('status'))

            # Пытаемся получить предмет и артикул из nm_settings
            subject = ''
            article = ''
            search_bid = 0
            recommendations_bid = 0
            nm_settings = adv.get('nm_settings', [])
            if nm_settings:
                first_nm = nm_settings[0]
                subject_obj = first_nm.get('subject', {})
                if subject_obj:
                    subject = subject_obj.get('name', '')
                article = first_nm.get('nm_id', '')
                bids_kopecks = first_nm.get('bids_kopecks', {})
                if bids_kopecks:
                    search_bid = bids_kopecks.get('search', 0) / 100
                    recommendations_bid = bids_kopecks.get('recommendations', 0) / 100

            campaign_info[adv_id] = {
                'name': name,
                'subject': subject,
                'article': article,
                'payment_type': payment_type,
                'bid_type': bid_type,
                'status': status,
                'search_bid': search_bid,
                'recommendations_bid': recommendations_bid
            }
            campaign_ids.append(adv_id)

            # Добавляем строку для листа Список_кампаний
            campaigns_list_rows.append({
                'ID кампании': adv_id,
                'Название': name,
                'Статус': status,
                'Тип оплаты': payment_type,
                'Тип ставки': bid_type,
                'Ставка в поиске (руб)': search_bid,
                'Ставка в рекомендациях (руб)': recommendations_bid,
                'Название предмета': subject,
                'Артикул WB': article
            })

        self.log(f"📊 Получено {len(campaign_ids)} кампаний с информацией")

        # 3. Определяем периоды статистики.
        explicit_adverts_backfill = _parse_optional_date_env("WB_ADVERTS_BACKFILL_FROM")
        if explicit_adverts_backfill is not None:
            # Ручной backfill оставляем отдельным режимом. Он обновляет актуальные значения,
            # но старые даты с номером снимка >14 не попадут в лаг-историю.
            backfill_start = min(explicit_adverts_backfill, self.target_date)
            requested_periods = self._split_date_range(backfill_start, self.target_date, 31)
            self.log(
                f"📚 Реклама: ручной backfill {backfill_start:%Y-%m-%d} — "
                f"{self.target_date:%Y-%m-%d}"
            )
        else:
            # Главная логика v40: на каждом ежедневном запуске обновляем ровно 14 дат
            # статистики: target_date и 13 предыдущих.
            rolling_end = self.target_date
            rolling_start = rolling_end - timedelta(days=13)
            requested_periods = []

            # FINICK по-прежнему может в первый раз восстановить 7 завершённых недель.
            # Эти периоды нужны для основного Анализа рекламы, но не создают фальшивую
            # историческую "лестницу" — _update_adverts_history оставит только снимки 1..14.
            if store_name == "FINICK" and finick_missing_week_ranges:
                requested_periods.extend(finick_missing_week_ranges)

            requested_periods.append((rolling_start, rolling_end))
            self.log(
                f"🔁 Rolling 14 дней рекламы: {rolling_start:%Y-%m-%d} — "
                f"{rolling_end:%Y-%m-%d}. Каждая дата будет перечитываться "
                f"один раз в день до снимка №14."
            )

        # adv/v3/fullstats принимает максимум 31 день за запрос.
        safe_periods: List[Tuple[datetime.date, datetime.date]] = []
        for p_start, p_end in requested_periods:
            safe_periods.extend(self._split_date_range(p_start, p_end, 31))

        # Убираем дубли периодов, сохраняя порядок.
        dedup_periods = []
        seen_periods = set()
        for p_start, p_end in safe_periods:
            key_period = (p_start, p_end)
            if key_period not in seen_periods:
                seen_periods.add(key_period)
                dedup_periods.append(key_period)
        requested_periods = dedup_periods

        self.log(
            "📅 Периоды рекламы: " +
            ", ".join(f"{a:%Y-%m-%d}..{b:%Y-%m-%d}" for a, b in requested_periods)
        )

        # Набор разрешённых дат нужен ниже, чтобы не захватывать текущую неделю между
        # последней завершённой неделей и target_date.
        requested_dates: Set[str] = set()
        for p_start, p_end in requested_periods:
            for i in range((p_end - p_start).days + 1):
                requested_dates.add((p_start + timedelta(days=i)).strftime('%Y-%m-%d'))

        # 4. Загружаем статистику для всех кампаний по каждому периоду.
        all_stats = []
        stats_url = "https://advert-api.wildberries.ru/adv/v3/fullstats"

        for period_idx, (p_start, p_end) in enumerate(requested_periods, start=1):
            start_str = p_start.strftime('%Y-%m-%d')
            end_str = p_end.strftime('%Y-%m-%d')
            self.log(
                f"📅 Реклама период {period_idx}/{len(requested_periods)}: "
                f"{start_str} — {end_str}"
            )

            # API допускает до 50 ID, оставляем консервативные чанки по 30.
            for i in range(0, len(campaign_ids), 30):
                chunk = campaign_ids[i:i+30]
                ids_param = ','.join(map(str, chunk))
                params = {
                    'ids': ids_param,
                    'beginDate': start_str,
                    'endDate': end_str,
                }

                retries = 0
                success = False
                while retries < 3 and not success:
                    try:
                        self.log(
                            f"⏳ Запрос статистики кампаний "
                            f"{i+1}-{min(i+30, len(campaign_ids))}..."
                        )
                        resp = requests.get(stats_url, headers=headers, params=params, timeout=60)
                        if resp.status_code == 200:
                            data = resp.json()
                            if data:
                                all_stats.extend(data)
                                self.log(f"✅ Получены данные для {len(data)} кампаний")
                            else:
                                self.log("ℹ️ Нет данных для этой группы")
                            success = True
                        elif resp.status_code == 429:
                            retries += 1
                            wait = 60 * retries
                            self.log(f"    ⚠️ Лимит, ждём {wait} сек...")
                            time.sleep(wait)
                        else:
                            self.log(
                                f"❌ Ошибка рекламы HTTP {resp.status_code} "
                                f"за {start_str}..{end_str}: {resp.text[:1000]}"
                            )
                            break
                    except Exception as e:
                        self.log(f"❌ Исключение: {e}")
                        break

                # Лимит fullstats — 3 запроса в минуту.
                time.sleep(20)

        if not all_stats:
            self.log("⚠️ Не получено статистических данных.")
            return False

        self.log(f"📊 Получена статистика для {len(all_stats)} записей кампаний/периодов")

        # 5. Преобразуем полученную статистику в DataFrame (ежедневная) и итоговую
        daily_rows = []
        summary_rows = []  # для итогового листа по кампаниям за период
        for camp in all_stats:
            camp_id = camp.get('advertId')
            if not camp_id:
                continue
            info = campaign_info.get(camp_id, {})
            subject = info.get('subject', '')
            name = info.get('name', '')
            article = info.get('article', '')
            days = camp.get('days', [])
            # Для итоговой статистики суммируем показатели по кампании за период
            camp_summary = {
                'ID кампании': camp_id,
                'Артикул WB': article,
                'Название': name,
                'Название предмета': subject,
                'Показы': 0,
                'Клики': 0,
                'CTR': 0,
                'CPC': 0,
                'Заказы': 0,
                'CR': 0,
                'Расход': 0,
                'ATBS': 0,
                'SHKS': 0,
                'Сумма заказов': 0,
                'Отменено': 0,
                'ДРР': 0
            }
            for day in days:
                day_date = day.get('date', '').split('T')[0]
                if not day_date or day_date not in requested_dates:
                    continue
                row = {
                    'ID кампании': camp_id,
                    'Артикул WB': article,
                    'Название': name,
                    'Название предмета': subject,
                    'Дата': day_date,
                    'Показы': day.get('views', 0),
                    'Клики': day.get('clicks', 0),
                    'CTR': day.get('ctr', 0),
                    'CPC': day.get('cpc', 0),
                    'Заказы': day.get('orders', 0),
                    'CR': day.get('cr', 0),
                    'Расход': day.get('sum', 0),
                    'ATBS': day.get('atbs', 0),
                    'SHKS': day.get('shks', 0),
                    'Сумма заказов': day.get('sum_price', 0),
                    'Отменено': day.get('canceled', 0),
                }
                if row['Сумма заказов'] > 0:
                    row['ДРР'] = round(row['Расход'] / (row['Сумма заказов'] * 0.88) * 100, 2)
                else:
                    row['ДРР'] = 0
                daily_rows.append(row)

                # Добавляем к итоговым суммам
                camp_summary['Показы'] += row['Показы']
                camp_summary['Клики'] += row['Клики']
                camp_summary['Заказы'] += row['Заказы']
                camp_summary['Расход'] += row['Расход']
                camp_summary['Сумма заказов'] += row['Сумма заказов']
                camp_summary['ATBS'] += row['ATBS']
                camp_summary['SHKS'] += row['SHKS']
                camp_summary['Отменено'] += row['Отменено']

            # Если были дни, считаем средние/итоговые метрики для кампании
            if days:
                camp_summary['CTR'] = round((camp_summary['Клики'] / camp_summary['Показы'] * 100) if camp_summary['Показы'] > 0 else 0, 2)
                camp_summary['CPC'] = round((camp_summary['Расход'] / camp_summary['Клики']) if camp_summary['Клики'] > 0 else 0, 2)
                camp_summary['CR'] = round((camp_summary['Заказы'] / camp_summary['Клики'] * 100) if camp_summary['Клики'] > 0 else 0, 2)
                if camp_summary['Сумма заказов'] > 0:
                    camp_summary['ДРР'] = round(camp_summary['Расход'] / (camp_summary['Сумма заказов'] * 0.88) * 100, 2)
                summary_rows.append(camp_summary)

        if not daily_rows:
            self.log("⚠️ Нет ежедневных данных для сохранения.")
            return False

        daily_df = pd.DataFrame(daily_rows)
        # FINICK backfill и rolling-окно могут пересекаться по датам. Для актуального слоя
        # одна кампания + одна дата должны существовать только один раз.
        if not daily_df.empty:
            daily_df['Дата'] = pd.to_datetime(daily_df['Дата'], errors='coerce').dt.strftime('%Y-%m-%d')
            daily_df = daily_df.drop_duplicates(subset=['ID кампании', 'Дата'], keep='last')
            daily_df = daily_df.sort_values(['Дата', 'ID кампании']).reset_index(drop=True)

        summary_df = pd.DataFrame(summary_rows)
        campaigns_df = pd.DataFrame(campaigns_list_rows)

        self.log(
            f"📊 Сформировано {len(daily_df)} уникальных ежедневных записей; "
            f"сырых итоговых записей по кампаниям/периодам: {len(summary_df)}"
        )

        # 6. Группируем только реально запрошенные даты по неделям.
        weeks = defaultdict(list)
        for date_str in sorted(requested_dates):
            d = datetime.strptime(date_str, '%Y-%m-%d').date()
            week_start = self._get_week_start(datetime.combine(d, datetime.min.time()))
            weeks[week_start].append(d)

        for week_start, dates in weeks.items():
            week_dates = [d.strftime('%Y-%m-%d') for d in dates]
            week_daily_df = daily_df[daily_df['Дата'].isin(week_dates)].copy()
            if week_daily_df.empty:
                continue

            # Загружаем существующий недельный файл (если есть) и объединяем по дням
            existing_week_df = self._load_weekly_data(store_name, 'adverts', week_start)
            if not existing_week_df.empty:
                # Объединяем, удаляем дубликаты по ID кампании и дате
                combined_daily = pd.concat([existing_week_df, week_daily_df], ignore_index=True)
                combined_daily = combined_daily.drop_duplicates(subset=['ID кампании', 'Дата'], keep='last')
            else:
                combined_daily = week_daily_df

            # Для итогового листа за неделю: агрегируем по кампаниям
            week_summary = combined_daily.groupby('ID кампании').agg({
                'Артикул WB': 'first',
                'Название': 'first',
                'Название предмета': 'first',
                'Показы': 'sum',
                'Клики': 'sum',
                'Заказы': 'sum',
                'Расход': 'sum',
                'Сумма заказов': 'sum',
                'ATBS': 'sum',
                'SHKS': 'sum',
                'Отменено': 'sum'
            }).reset_index()
            week_summary['CTR'] = (week_summary['Клики'] / week_summary['Показы'] * 100).round(2)
            week_summary['CPC'] = (week_summary['Расход'] / week_summary['Клики']).round(2)
            week_summary['CR'] = (week_summary['Заказы'] / week_summary['Клики'] * 100).round(2)
            week_summary['ДРР'] = (week_summary['Расход'] / (week_summary['Сумма заказов'] * 0.88) * 100).round(2)
            week_summary.fillna(0, inplace=True)

            # Для списка кампаний за неделю – берём актуальные данные из campaigns_df
            week_campaigns = campaigns_df[campaigns_df['ID кампании'].isin(combined_daily['ID кампании'].unique())].copy()

            # Сохраняем недельный файл с несколькими листами
            weekly_key = self._get_weekly_key(store_name, 'adverts', week_start)
            sheets = {
                'Статистика_Ежедневно': combined_daily,
                'Статистика_Итого': week_summary,
                'Список_кампаний': week_campaigns
            }
            self.s3.write_excel_multi(weekly_key, sheets)
            self.log(f"✅ Недельный файл с несколькими листами сохранён: {weekly_key}")

        # 7. Дополнительно формируем отчёты по категориям и единый аналитический файл
        if not daily_df.empty:
            # Отчёт по категориям за каждый день
            daily_cat = daily_df.groupby(['Дата', 'Название предмета']).agg({
                'Показы': 'sum',
                'Клики': 'sum',
                'Заказы': 'sum',
                'Расход': 'sum',
                'Сумма заказов': 'sum'
            }).reset_index()
            daily_cat['CTR'] = (daily_cat['Клики'] / daily_cat['Показы'] * 100).round(2)
            daily_cat['CPC'] = (daily_cat['Расход'] / daily_cat['Клики']).round(2)
            daily_cat['CR'] = (daily_cat['Заказы'] / daily_cat['Клики'] * 100).round(2)
            daily_cat['ROI'] = ((daily_cat['Сумма заказов'] - daily_cat['Расход']) / daily_cat['Расход'] * 100).round(2)
            daily_cat['ДРР'] = (daily_cat['Расход'] / (daily_cat['Сумма заказов'] * 0.88) * 100).round(2)
            daily_cat = daily_cat.sort_values(['Дата', 'Расход'], ascending=[True, False])

            # Итоговый отчёт по категориям
            summary_cat = daily_df.groupby('Название предмета').agg({
                'Показы': 'sum',
                'Клики': 'sum',
                'Заказы': 'sum',
                'Расход': 'sum',
                'Сумма заказов': 'sum'
            }).reset_index()
            summary_cat['CTR'] = (summary_cat['Клики'] / summary_cat['Показы'] * 100).round(2)
            summary_cat['CPC'] = (summary_cat['Расход'] / summary_cat['Клики']).round(2)
            summary_cat['CR'] = (summary_cat['Заказы'] / summary_cat['Клики'] * 100).round(2)
            summary_cat['ROI'] = ((summary_cat['Сумма заказов'] - summary_cat['Расход']) / summary_cat['Расход'] * 100).round(2)
            summary_cat['ДРР'] = (summary_cat['Расход'] / (summary_cat['Сумма заказов'] * 0.88) * 100).round(2)
            summary_cat = summary_cat.sort_values('Расход', ascending=False)

            # Сохраняем единый аналитический файл со всеми листами.
            # v28: daily_df за вчера объединяем с существующей историей, чтобы не затирать файл одним днём.
            analytics_key = f"Отчёты/{config['folder']}/{store_name}/Анализ рекламы.xlsx"
            analytics_daily = daily_df.copy()
            if self.s3.file_exists(analytics_key):
                try:
                    old_daily = self.s3.read_excel(analytics_key, sheet_name='Статистика_Ежедневно')
                    if not old_daily.empty:
                        analytics_daily = pd.concat([old_daily, analytics_daily], ignore_index=True)
                        if 'Дата' in analytics_daily.columns:
                            analytics_daily['Дата'] = pd.to_datetime(analytics_daily['Дата'], errors='coerce').dt.strftime('%Y-%m-%d')
                            retention_days = 56 if store_name == "FINICK" else 30
                            cutoff = (
                                datetime.now(pytz.timezone('Europe/Moscow')).date()
                                - timedelta(days=retention_days)
                            ).strftime('%Y-%m-%d')
                            analytics_daily = analytics_daily[analytics_daily['Дата'] >= cutoff]
                        analytics_daily = analytics_daily.drop_duplicates(subset=['ID кампании', 'Дата'], keep='last')
                except Exception as e:
                    self.log(f"⚠️ Не удалось объединить старую ежедневную рекламу: {e}")

            analytics_summary = analytics_daily.groupby('ID кампании').agg({
                'Артикул WB': 'first',
                'Название': 'first',
                'Название предмета': 'first',
                'Показы': 'sum',
                'Клики': 'sum',
                'Заказы': 'sum',
                'Расход': 'sum',
                'Сумма заказов': 'sum',
                'ATBS': 'sum',
                'SHKS': 'sum',
                'Отменено': 'sum'
            }).reset_index()
            analytics_summary['CTR'] = (analytics_summary['Клики'] / analytics_summary['Показы'] * 100).round(2)
            analytics_summary['CPC'] = (analytics_summary['Расход'] / analytics_summary['Клики']).round(2)
            analytics_summary['CR'] = (analytics_summary['Заказы'] / analytics_summary['Клики'] * 100).round(2)
            analytics_summary['ДРР'] = (analytics_summary['Расход'] / (analytics_summary['Сумма заказов'] * 0.88) * 100).round(2)
            analytics_summary = analytics_summary.replace([float('inf'), -float('inf')], 0).fillna(0)

            daily_cat = analytics_daily.groupby(['Дата', 'Название предмета']).agg({
                'Показы': 'sum',
                'Клики': 'sum',
                'Заказы': 'sum',
                'Расход': 'sum',
                'Сумма заказов': 'sum'
            }).reset_index()
            daily_cat['CTR'] = (daily_cat['Клики'] / daily_cat['Показы'] * 100).round(2)
            daily_cat['CPC'] = (daily_cat['Расход'] / daily_cat['Клики']).round(2)
            daily_cat['CR'] = (daily_cat['Заказы'] / daily_cat['Клики'] * 100).round(2)
            daily_cat['ROI'] = ((daily_cat['Сумма заказов'] - daily_cat['Расход']) / daily_cat['Расход'] * 100).round(2)
            daily_cat['ДРР'] = (daily_cat['Расход'] / (daily_cat['Сумма заказов'] * 0.88) * 100).round(2)
            daily_cat = daily_cat.replace([float('inf'), -float('inf')], 0).fillna(0).sort_values(['Дата', 'Расход'], ascending=[True, False])

            summary_cat = analytics_daily.groupby('Название предмета').agg({
                'Показы': 'sum',
                'Клики': 'sum',
                'Заказы': 'sum',
                'Расход': 'sum',
                'Сумма заказов': 'sum'
            }).reset_index()
            summary_cat['CTR'] = (summary_cat['Клики'] / summary_cat['Показы'] * 100).round(2)
            summary_cat['CPC'] = (summary_cat['Расход'] / summary_cat['Клики']).round(2)
            summary_cat['CR'] = (summary_cat['Заказы'] / summary_cat['Клики'] * 100).round(2)
            summary_cat['ROI'] = ((summary_cat['Сумма заказов'] - summary_cat['Расход']) / summary_cat['Расход'] * 100).round(2)
            summary_cat['ДРР'] = (summary_cat['Расход'] / (summary_cat['Сумма заказов'] * 0.88) * 100).round(2)
            summary_cat = summary_cat.replace([float('inf'), -float('inf')], 0).fillna(0).sort_values('Расход', ascending=False)

            sheets_analytics = {
                'Статистика_Ежедневно': analytics_daily,
                'Статистика_Итого': analytics_summary,
                'Список_кампаний': campaigns_df,
                'Отчет_по_Категории': daily_cat,
                'Отчет_по_Категории_Итог': summary_cat
            }
            self.s3.write_excel_multi(analytics_key, sheets_analytics)
            self.log(f"📊 Аналитический отчёт сохранён: {analytics_key} (листы: {', '.join(sheets_analytics.keys())})")

        # 8. Сохраняем историческую таблицу за последние 14 дней (накопление)
        self._update_adverts_history(store_name, daily_df)

        self.log("✅ Реклама успешно обновлена")
        return True

    def _update_adverts_history(self, store_name: str, new_daily_df: pd.DataFrame):
        """История лага рекламной статистики WB.

        Для каждой даты статистики сохраняем до 14 последовательных ежедневных снимков.
        Пример при target_date=2026-08-06:
        - 07.08 получаем снимок №1 для статистики 06.08;
        - 08.08 — снимок №2 той же статистики;
        - ...;
        - 20.08 — снимок №14.

        История не является "последними 14 календарными днями хранения".
        14 — это число наблюдений за КАЖДОЙ датой статистики. Для анализа во времени
        оставляем до 90 дней дат статистики, каждая из которых может иметь 14 снимков.

        Файл:
        Отчёты/Реклама/{store_name}/История_рекламы_14дней.xlsx

        Листы:
        - История: сырые снимки кампаний;
        - Сводка_лага: суммы показателей по дате статистики и номеру снимка;
        - Контроль_14_снимков: сколько снимков уже накоплено для каждой даты.
        """
        config = self.reports_config['adverts']
        history_key = f"Отчёты/{config['folder']}/{store_name}/История_рекламы_14дней.xlsx"

        if new_daily_df is None or new_daily_df.empty:
            self.log("ℹ️ История рекламы: новых строк для снимка нет")
            return

        now_msk = datetime.now(pytz.timezone('Europe/Moscow'))
        request_ts = now_msk.strftime('%Y-%m-%d %H:%M:%S')
        snapshot_date = now_msk.date()

        # Новый снимок.
        new_rows = new_daily_df.copy()
        new_rows['Дата'] = pd.to_datetime(new_rows['Дата'], errors='coerce').dt.date
        new_rows = new_rows[new_rows['Дата'].notna()].copy()

        new_rows['Дата снимка'] = snapshot_date.strftime('%Y-%m-%d')
        new_rows['Дата целевой выгрузки'] = self.target_date.strftime('%Y-%m-%d')
        new_rows['Дата запроса'] = request_ts
        new_rows['Лаг, дней'] = new_rows['Дата'].apply(
            lambda d: (snapshot_date - d).days if d else None
        )
        # Снимок № — это именно фактический лаг в календарных днях.
        # 06.08, запрошенный 07.08 = снимок №1; 08.08 = №2; ...; 20.08 = №14.
        new_rows['Снимок №'] = new_rows['Лаг, дней']

        # В лаг-историю берём только реальные наблюдения на 1..14 день после даты статистики.
        # Исторический backfill нужен для Анализа рекламы, но не должен притворяться
        # своевременным ежедневным наблюдением.
        before_filter = len(new_rows)
        new_rows = new_rows[
            new_rows['Лаг, дней'].apply(lambda x: pd.notna(x) and 1 <= int(x) <= 14)
        ].copy()
        skipped_backfill = before_filter - len(new_rows)
        if skipped_backfill:
            self.log(
                f"ℹ️ История рекламы: {skipped_backfill} строк backfill не записаны в лаг-историю "
                f"(снимок вне диапазона 1..14)"
            )

        if new_rows.empty:
            self.log("ℹ️ История рекламы: в этом запуске нет строк для снимков №1..14")
            return

        # Возвращаем дату статистики к строковому виду перед объединением/Excel.
        new_rows['Дата'] = new_rows['Дата'].astype(str)

        # Загружаем существующую историю. Поддерживаем старый одно-листовый файл.
        if self.s3.file_exists(history_key):
            try:
                df_history = self.s3.read_excel(history_key, sheet_name='История')
                if df_history.empty:
                    df_history = self.s3.read_excel(history_key, sheet_name=0)
            except Exception:
                try:
                    df_history = self.s3.read_excel(history_key, sheet_name=0)
                except Exception:
                    df_history = pd.DataFrame()
        else:
            df_history = pd.DataFrame()

        # Миграция существующего файла v39: добавляем лаговые поля из Даты запроса.
        if not df_history.empty:
            if 'Дата' in df_history.columns:
                old_stat_dates = pd.to_datetime(df_history['Дата'], errors='coerce')
            else:
                old_stat_dates = pd.Series(pd.NaT, index=df_history.index)

            if 'Дата снимка' not in df_history.columns:
                if 'Дата запроса' in df_history.columns:
                    old_request_dates = pd.to_datetime(df_history['Дата запроса'], errors='coerce')
                    df_history['Дата снимка'] = old_request_dates.dt.strftime('%Y-%m-%d')
                else:
                    df_history['Дата снимка'] = ''

            old_snapshot_dates = pd.to_datetime(df_history['Дата снимка'], errors='coerce')

            if 'Лаг, дней' not in df_history.columns:
                df_history['Лаг, дней'] = (
                    old_snapshot_dates.dt.normalize() - old_stat_dates.dt.normalize()
                ).dt.days

            if 'Снимок №' not in df_history.columns:
                # Для старой истории ближайшее корректное приближение — фактический лаг.
                df_history['Снимок №'] = pd.to_numeric(df_history['Лаг, дней'], errors='coerce')

            if 'Дата целевой выгрузки' not in df_history.columns:
                # Для старых строк точного target_date не было в файле.
                # Восстанавливаем его из даты статистики + номер снимка - 1.
                snap_num = pd.to_numeric(df_history['Снимок №'], errors='coerce')
                restored_target = old_stat_dates + pd.to_timedelta(snap_num.fillna(1) - 1, unit='D')
                df_history['Дата целевой выгрузки'] = restored_target.dt.strftime('%Y-%m-%d')

        if not df_history.empty:
            combined = pd.concat([df_history, new_rows], ignore_index=True, sort=False)
        else:
            combined = new_rows.copy()

        # Нормализация типов.
        combined['Дата'] = pd.to_datetime(combined['Дата'], errors='coerce').dt.strftime('%Y-%m-%d')
        combined['Дата снимка'] = pd.to_datetime(combined['Дата снимка'], errors='coerce').dt.strftime('%Y-%m-%d')
        combined['Снимок №'] = pd.to_numeric(combined['Снимок №'], errors='coerce')
        combined['Лаг, дней'] = pd.to_numeric(combined['Лаг, дней'], errors='coerce')

        # В истории держим только реальные лаги/снимки 1..14.
        combined = combined[
            combined['Лаг, дней'].apply(lambda x: pd.notna(x) and 1 <= int(x) <= 14)
        ].copy()
        combined['Снимок №'] = combined['Лаг, дней']

        # Повторный запуск с тем же target_date обновляет тот же номер снимка, а не создаёт дубль.
        history_key_cols = [c for c in ['ID кампании', 'Дата', 'Снимок №'] if c in combined.columns]
        if len(history_key_cols) == 3:
            combined = combined.drop_duplicates(subset=history_key_cols, keep='last')

        # Не даём Excel-файлу расти бесконечно: храним 90 последних дат статистики,
        # но для каждой такой даты сохраняем все её снимки №1..14.
        cutoff_stat_date = self.target_date - timedelta(days=89)
        stat_dates = pd.to_datetime(combined['Дата'], errors='coerce').dt.date
        combined = combined.loc[stat_dates >= cutoff_stat_date].copy()

        sort_cols = [c for c in ['Дата', 'Снимок №', 'ID кампании'] if c in combined.columns]
        if sort_cols:
            combined = combined.sort_values(sort_cols).reset_index(drop=True)

        # Сводка лага — удобно сравнивать, как одна и та же дата менялась от снимка 1 к 14.
        metric_cols = [
            c for c in [
                'Показы', 'Клики', 'Заказы', 'Расход', 'ATBS', 'SHKS',
                'Сумма заказов', 'Отменено'
            ] if c in combined.columns
        ]
        agg_map = {c: 'sum' for c in metric_cols}
        if 'ID кампании' in combined.columns:
            agg_map['ID кампании'] = 'nunique'

        summary = combined.groupby(
            ['Дата', 'Снимок №'], dropna=False
        ).agg(agg_map).reset_index() if agg_map else pd.DataFrame()

        if not summary.empty:
            if 'ID кампании' in summary.columns:
                summary = summary.rename(columns={'ID кампании': 'Кампаний'})

            # Дата фактического снимка / запроса — берём последнюю для этого шага.
            snapshot_meta = combined.groupby(['Дата', 'Снимок №'], dropna=False).agg({
                'Дата снимка': 'max',
                'Дата запроса': 'max',
                'Лаг, дней': 'max',
            }).reset_index()
            summary = summary.merge(snapshot_meta, on=['Дата', 'Снимок №'], how='left')

            # Изменение к предыдущему снимку той же даты.
            summary = summary.sort_values(['Дата', 'Снимок №']).reset_index(drop=True)
            for metric in metric_cols:
                summary[f'Δ {metric} к пред. снимку'] = (
                    summary.groupby('Дата')[metric].diff().fillna(0)
                )

        # Контроль полноты: должны накопить 14 разных номеров снимка на каждую дату.
        control = combined.groupby('Дата').agg(
            **{
                'Снимков накоплено': ('Снимок №', 'nunique'),
                'Первый снимок': ('Снимок №', 'min'),
                'Последний снимок': ('Снимок №', 'max'),
                'Последняя дата снимка': ('Дата снимка', 'max'),
            }
        ).reset_index()
        control['Статус 14 снимков'] = control['Снимков накоплено'].apply(
            lambda n: 'ГОТОВО 14/14' if int(n) >= 14 else f'{int(n)}/14'
        )
        control = control.sort_values('Дата', ascending=False).reset_index(drop=True)

        sheets = {
            'История': combined,
            'Сводка_лага': summary,
            'Контроль_14_снимков': control,
        }
        self.s3.write_excel_multi(history_key, sheets)

        # Лог по target_date: сколько снимков уже накоплено для самой свежей даты.
        target_str = self.target_date.strftime('%Y-%m-%d')
        target_control = control[control['Дата'] == target_str]
        if not target_control.empty:
            count = int(target_control.iloc[0]['Снимков накоплено'])
            status = target_control.iloc[0]['Статус 14 снимков']
            self.log(
                f"📚 История рекламы: дата {target_str} получила снимок №1; "
                f"сейчас накоплено {count}/14 ({status})"
            )

        self.log(
            f"📊 История_рекламы_14дней.xlsx обновлена: {len(combined)} строк снимков; "
            f"дат статистики в 90-дневном окне: {control['Дата'].nunique() if not control.empty else 0}"
        )

    # ---------- Остатки из 1С (отключено, метод оставлен для возможности возврата) ----------
    def update_1c_stocks(self, store_name: str = '1С') -> bool:
        self.log(f"\n📌 ОБНОВЛЕНИЕ: Остатки из 1С для магазина {store_name}")
        config = self.reports_config['1c_stocks']

        url_1c = os.environ.get('URL_1C_STOCKS')
        username = os.environ.get('_1C_USER')
        password = os.environ.get('_1C_PASSWORD')

        if not url_1c:
            self.log("❌ Переменная окружения URL_1C_STOCKS не задана. Пропускаем.")
            return False

        auth = None
        if username and password:
            auth = (username, password)
            self.log(f"🔐 Используется базовая аутентификация для пользователя {username}")

        google_match = re.search(r'docs\.google\.com/spreadsheets/d/([a-zA-Z0-9_-]+)(?:/.*?gid=(\d+))?', url_1c)
        if google_match:
            spreadsheet_id = google_match.group(1)
            gid = google_match.group(2)
            if not gid:
                self.log("❌ В ссылке на Google Sheets не найден параметр gid. Укажите ссылку на конкретный лист.")
                return False
            download_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=xlsx&gid={gid}"
            self.log(f"📄 Обнаружена Google Sheets, gid={gid}. Будет скачан лист с этим gid.")
        else:
            download_url = url_1c
            self.log("📄 Используется прямая ссылка на файл.")

        tmp_path = None
        try:
            self.log(f"📥 Скачивание файла из: {download_url}")
            resp = requests.get(download_url, auth=auth, timeout=120, stream=True, allow_redirects=True)
            if resp.status_code != 200:
                self.log(f"❌ Ошибка при скачивании: HTTP {resp.status_code}")
                return False

            with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
                tmp_path = tmp.name
                for chunk in resp.iter_content(chunk_size=8192):
                    tmp.write(chunk)
            self.log(f"📦 Файл временно сохранён: {tmp_path}")

            key = f"Отчёты/{config['folder']}/{store_name}/{config['filename']}"
            self.log(f"☁️ Загрузка в бакет: {key}")
            self.s3.upload_file(tmp_path, key)
            self.log(f"✅ Файл успешно сохранён в бакет: {key}")

            return True

        except Exception as e:
            self.log(f"❌ Исключение при обработке: {e}")
            traceback.print_exc()
            return False
        finally:
            if tmp_path and os.path.exists(tmp_path):
                os.remove(tmp_path)
                self.log("🧹 Временный файл удалён")


    # ====================== ИИ-АГЕНТ: КАРТОЧКИ / ХАРАКТЕРИСТИКИ / ФОТО ======================

    @staticmethod
    def _agent_env_set(name: str) -> Set[str]:
        """Прочитать список значений из env: запятая/точка с запятой/перенос строки."""
        raw = (os.environ.get(name, "") or "").strip()
        if not raw:
            return set()
        return {x.strip() for x in re.split(r"[,;\n\r]+", raw) if x.strip()}

    @staticmethod
    def _agent_safe_name(value: Any, fallback: str = "UNKNOWN") -> str:
        text = str(value or "").strip() or fallback
        # Windows/S3/ZIP-safe имя папки, но сохраняем кириллицу и пробелы.
        text = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', text)
        text = re.sub(r'\s+', ' ', text).strip(' .')
        return (text[:140] or fallback)

    @staticmethod
    def _agent_json_text(value: Any) -> str:
        if value is None:
            return ""
        if isinstance(value, str):
            return value.strip()
        if isinstance(value, (int, float, bool)):
            return str(value)
        return json.dumps(value, ensure_ascii=False, default=str)

    def _agent_archive_keys(self, store_name: str) -> Tuple[str, str]:
        base = f"Отзывы/Обучение/{store_name}"
        return (
            f"{base}/WB_Агент_Ассортимент_{store_name}.zip",
            f"{base}/WB_Агент_Ассортимент_{store_name}.manifest.json",
        )

    def _fetch_all_content_cards(self, store_name: str) -> Optional[List[dict]]:
        """Получить все карточки продавца из Content API.

        Используется POST /content/v2/get/cards/list. В ответе WB уже отдаёт
        описание, characteristics, sizes, photos, video, tags и прочие поля карточки.
        Сохраняем ответ карточки целиком, чтобы не потерять новые поля API.
        """
        token = (self.api_keys[store_name].get('content') or self.api_keys[store_name].get('promo') or '').strip()
        if not token:
            self.log(f"❌ agent_catalog {store_name}: не найден Content-токен")
            return None

        url = "https://content-api.wildberries.ru/content/v2/get/cards/list"
        headers = {"Authorization": token, "Content-Type": "application/json"}
        cursor: Dict[str, Any] = {'limit': 100}
        last_cursor: Optional[Tuple[Any, Any]] = None
        cards_all: List[dict] = []

        for page in range(1, 1000):
            payload = {
                'settings': {
                    'sort': {'ascending': True},
                    'filter': {'withPhoto': -1},
                    'cursor': cursor,
                }
            }

            data = None
            for attempt in range(1, 6):
                try:
                    resp = requests.post(url, headers=headers, json=payload, timeout=120)
                    if resp.status_code == 200:
                        data = resp.json()
                        break
                    if resp.status_code == 429:
                        wait = self._rate_limit_wait_seconds(resp, default_seconds=max(5, attempt * 5), max_seconds=120)
                        self.log(f"⚠️ agent_catalog cards/list: 429, попытка {attempt}/5, ждём {wait} сек")
                        time.sleep(wait)
                        continue
                    if resp.status_code in (500, 502, 503, 504):
                        wait = 10 * attempt
                        self.log(f"⚠️ agent_catalog cards/list: HTTP {resp.status_code}, попытка {attempt}/5, ждём {wait} сек")
                        time.sleep(wait)
                        continue
                    if resp.status_code in (401, 403):
                        self.log(
                            f"❌ agent_catalog {store_name}: Content API вернул {resp.status_code}. "
                            f"Нужен токен с категорией Content. Ответ: {resp.text[:700]}"
                        )
                        return None
                    self.log(f"❌ agent_catalog cards/list HTTP {resp.status_code}: {resp.text[:1000]}")
                    return None
                except Exception as e:
                    if attempt >= 5:
                        self.log(f"❌ agent_catalog cards/list: {e}")
                        return None
                    time.sleep(5 * attempt)

            if data is None:
                return None

            cards = data.get('cards') or []
            cards_all.extend(cards)
            self.log(f"📦 agent_catalog: страница {page}, карточек {len(cards)}, всего {len(cards_all)}")

            cur = data.get('cursor') or {}
            if not cards or len(cards) < 100:
                break
            next_cursor = (cur.get('updatedAt'), cur.get('nmID', cur.get('nmId')))
            if not next_cursor[0] or not next_cursor[1] or next_cursor == last_cursor:
                break
            last_cursor = next_cursor
            cursor = {'limit': 100, 'updatedAt': next_cursor[0], 'nmID': next_cursor[1]}
            time.sleep(0.65)

        # На случай повторов курсора оставляем одну последнюю карточку на nmID.
        dedup: Dict[str, dict] = {}
        for card in cards_all:
            nm_id = str(card.get('nmID', card.get('nmId', '')) or '').strip()
            key = nm_id or f"vendor:{card.get('vendorCode', '')}:{len(dedup)}"
            dedup[key] = card
        return list(dedup.values())

    def _agent_filter_cards(self, cards: List[dict]) -> List[dict]:
        """Опциональные фильтры для выведенных из оборота/служебных карточек.

        WB_AGENT_INCLUDE_VENDOR_CODES — если задан, экспортируются только эти артикулы продавца.
        WB_AGENT_EXCLUDE_VENDOR_CODES — исключить артикулы продавца.
        WB_AGENT_EXCLUDE_NMIDS — исключить nmID.
        """
        include_vendor = self._agent_env_set('WB_AGENT_INCLUDE_VENDOR_CODES')
        exclude_vendor = self._agent_env_set('WB_AGENT_EXCLUDE_VENDOR_CODES')
        exclude_nmids = self._agent_env_set('WB_AGENT_EXCLUDE_NMIDS')

        result = []
        skipped = 0
        for card in cards or []:
            vendor = str(card.get('vendorCode', '') or '').strip()
            nm_id = str(card.get('nmID', card.get('nmId', '')) or '').strip()
            if include_vendor and vendor not in include_vendor:
                skipped += 1
                continue
            if vendor in exclude_vendor or nm_id in exclude_nmids:
                skipped += 1
                continue
            result.append(card)

        if skipped:
            self.log(f"🧹 agent_catalog: исключено карточек фильтрами: {skipped}")
        return result

    @staticmethod
    def _agent_photo_urls(card: dict) -> List[dict]:
        photos = card.get('photos') or []
        result: List[dict] = []
        seen: Set[str] = set()
        preferred = ['big', 'c516x688', 'c246x328', 'square', 'tm']

        for idx, photo in enumerate(photos, start=1):
            url = ''
            variants = {}
            if isinstance(photo, str):
                url = photo.strip()
                variants = {'source': url}
            elif isinstance(photo, dict):
                variants = photo
                for key in preferred:
                    candidate = photo.get(key)
                    if isinstance(candidate, str) and candidate.startswith(('http://', 'https://')):
                        url = candidate
                        break
                if not url:
                    for key, candidate in photo.items():
                        if isinstance(candidate, str) and candidate.startswith(('http://', 'https://')):
                            url = candidate
                            break
            if url and url not in seen:
                seen.add(url)
                result.append({'index': idx, 'url': url, 'variants': variants})
        return result

    @staticmethod
    def _agent_ext_from_response(url: str, content_type: str) -> str:
        ct = (content_type or '').split(';')[0].strip().lower()
        ext = mimetypes.guess_extension(ct) if ct else None
        if ext == '.jpe':
            ext = '.jpg'
        if ext in {'.jpg', '.jpeg', '.png', '.webp', '.gif'}:
            return ext
        path_ext = os.path.splitext(urlparse(url).path)[1].lower()
        if path_ext in {'.jpg', '.jpeg', '.png', '.webp', '.gif'}:
            return path_ext
        return '.jpg'

    def _agent_download_media(self, url: str, dst_without_ext: str) -> Tuple[Optional[str], Optional[str]]:
        """Скачать изображение с retry. Возвращает (путь, ошибка)."""
        for attempt in range(1, 4):
            try:
                resp = requests.get(
                    url,
                    timeout=90,
                    stream=True,
                    headers={'User-Agent': 'Mozilla/5.0 WB-Agent-Catalog/1.0'},
                )
                if resp.status_code == 200:
                    ext = self._agent_ext_from_response(url, resp.headers.get('Content-Type', ''))
                    dst = dst_without_ext + ext
                    with open(dst, 'wb') as f:
                        for chunk in resp.iter_content(chunk_size=1024 * 256):
                            if chunk:
                                f.write(chunk)
                    if os.path.getsize(dst) <= 0:
                        return None, 'скачан пустой файл'
                    return dst, None
                if resp.status_code in (429, 500, 502, 503, 504) and attempt < 3:
                    time.sleep(3 * attempt)
                    continue
                return None, f"HTTP {resp.status_code}"
            except Exception as e:
                if attempt >= 3:
                    return None, str(e)
                time.sleep(3 * attempt)
        return None, 'неизвестная ошибка'

    def _agent_product_markdown(self, card: dict) -> str:
        nm_id = card.get('nmID', card.get('nmId', ''))
        vendor = card.get('vendorCode', '')
        title = card.get('title', '')
        brand = card.get('brand', card.get('brandName', ''))
        subject = card.get('subjectName', '')
        description = card.get('description', '')
        characteristics = card.get('characteristics') or []
        sizes = card.get('sizes') or []
        dimensions = card.get('dimensions') or {}
        tags = card.get('tags') or []
        need_kiz = card.get('needKiz', '')

        lines = [
            f"# {vendor or nm_id} — {title}",
            "",
            "## Идентификация",
            f"- Артикул продавца: {vendor}",
            f"- Артикул WB (nmID): {nm_id}",
            f"- Бренд: {brand}",
            f"- Предмет: {subject}",
            f"- Требуется маркировка КИЗ: {need_kiz}",
            f"- Создано: {card.get('createdAt', '')}",
            f"- Обновлено: {card.get('updatedAt', '')}",
            "",
            "## Название",
            str(title or ''),
            "",
            "## Описание",
            str(description or ''),
            "",
            "## Габариты",
        ]
        if isinstance(dimensions, dict) and dimensions:
            for key, value in dimensions.items():
                lines.append(f"- {key}: {self._agent_json_text(value)}")
        else:
            lines.append("- Нет данных")

        lines += ["", "## Характеристики"]
        if characteristics:
            for item in characteristics:
                if isinstance(item, dict):
                    name = item.get('name') or item.get('id') or 'Характеристика'
                    value = item.get('value', item.get('values', ''))
                    lines.append(f"- {name}: {self._agent_json_text(value)}")
                else:
                    lines.append(f"- {self._agent_json_text(item)}")
        else:
            lines.append("- Нет данных")

        lines += ["", "## Размеры и баркоды"]
        if sizes:
            for size in sizes:
                if isinstance(size, dict):
                    lines.append(
                        "- " + "; ".join([
                            f"chrtID={size.get('chrtID', size.get('chrtId', ''))}",
                            f"techSize={size.get('techSize', '')}",
                            f"wbSize={size.get('wbSize', '')}",
                            f"skus={self._agent_json_text(size.get('skus') or [])}",
                        ])
                    )
                else:
                    lines.append(f"- {self._agent_json_text(size)}")
        else:
            lines.append("- Нет данных")

        lines += [
            "",
            "## Медиа",
            f"- Фотографий в карточке: {len(card.get('photos') or [])}",
            f"- Видео: {self._agent_json_text(card.get('video', ''))}",
            "",
            "## Теги",
            self._agent_json_text(tags) or "Нет данных",
            "",
            "## Примечание для ИИ-агента",
            "Факты в этом файле получены из текущей карточки продавца через WB Content API. "
            "Для спорных характеристик приоритет имеет карточка_WB.json и актуальные данные API. "
            "Не делай вывод о материале, размере или совместимости только по внешнему виду фотографии.",
            "",
        ]
        return "\n".join(lines)

    def _agent_catalog_signature(self, cards: List[dict]) -> str:
        canonical = json.dumps(cards, ensure_ascii=False, sort_keys=True, separators=(',', ':'), default=str)
        return hashlib.sha256(canonical.encode('utf-8')).hexdigest()

    def update_agent_catalog(self, store_name: str) -> bool:
        """Создать один ZIP для базы знаний ИИ-агента, разбитый по артикулам продавца.

        Структура архива:
          Каталог.xlsx
          manifest.json
          README.txt
          <Артикул продавца>/
              паспорт_товара.md
              карточка_WB.json
              характеристики.json
              размеры_и_баркоды.json
              медиа.json
              фото/01.jpg ...

        Архив перезаписывается по одному стабильному ключу в Object Storage.
        Если карточки не изменились, фотографии повторно не скачиваются.
        """
        self.log("")
        self.log(f"📌 ОБНОВЛЕНИЕ: Архив знаний ИИ-агента для магазина {store_name}")

        cards = self._fetch_all_content_cards(store_name)
        if cards is None:
            return False
        cards = self._agent_filter_cards(cards)
        if not cards:
            self.log("⚠️ agent_catalog: после фильтров не осталось карточек")
            return False

        cards = sorted(
            cards,
            key=lambda c: (str(c.get('vendorCode', '') or '').lower(), int(c.get('nmID', c.get('nmId', 0)) or 0))
        )
        signature = self._agent_catalog_signature(cards)
        archive_key, manifest_key = self._agent_archive_keys(store_name)
        force = (os.environ.get('WB_AGENT_CATALOG_FORCE', '') or '').strip().lower() in {'1', 'true', 'yes', 'y', 'да'}
        old_manifest = self.s3.read_json(manifest_key) if self.s3.file_exists(manifest_key) else None

        if (
            not force
            and old_manifest
            and old_manifest.get('signature') == signature
            and bool(old_manifest.get('complete'))
            and self.s3.file_exists(archive_key)
        ):
            self.log(
                f"✅ agent_catalog: карточки не изменились ({len(cards)} шт.), "
                f"готовый архив уже есть: {archive_key}"
            )
            return True

        max_photos_raw = (os.environ.get('WB_AGENT_MAX_PHOTOS', '') or '').strip()
        try:
            max_photos = int(max_photos_raw) if max_photos_raw else 0
        except ValueError:
            max_photos = 0
        # 0 = все фотографии.

        now_msk = datetime.now(pytz.timezone('Europe/Moscow'))
        catalog_rows: List[dict] = []
        photo_errors: List[dict] = []
        total_photos = 0
        downloaded_photos = 0
        used_folder_names: Set[str] = set()

        with tempfile.TemporaryDirectory(prefix='wb_agent_catalog_') as tmp_dir:
            root_dir = os.path.join(tmp_dir, f"WB_Агент_Ассортимент_{store_name}")
            os.makedirs(root_dir, exist_ok=True)

            for idx, card in enumerate(cards, start=1):
                nm_id = card.get('nmID', card.get('nmId', ''))
                vendor = str(card.get('vendorCode', '') or '').strip()
                base_folder = self._agent_safe_name(vendor, fallback=f"nm_{nm_id}")
                folder_name = base_folder
                if folder_name in used_folder_names:
                    folder_name = self._agent_safe_name(f"{base_folder}__nm{nm_id}")
                used_folder_names.add(folder_name)

                product_dir = os.path.join(root_dir, folder_name)
                photos_dir = os.path.join(product_dir, 'фото')
                os.makedirs(photos_dir, exist_ok=True)

                # Источник истины — карточка целиком.
                with open(os.path.join(product_dir, 'карточка_WB.json'), 'w', encoding='utf-8') as f:
                    json.dump(card, f, ensure_ascii=False, indent=2, default=str)

                with open(os.path.join(product_dir, 'характеристики.json'), 'w', encoding='utf-8') as f:
                    json.dump(card.get('characteristics') or [], f, ensure_ascii=False, indent=2, default=str)

                with open(os.path.join(product_dir, 'размеры_и_баркоды.json'), 'w', encoding='utf-8') as f:
                    json.dump(card.get('sizes') or [], f, ensure_ascii=False, indent=2, default=str)

                with open(os.path.join(product_dir, 'паспорт_товара.md'), 'w', encoding='utf-8') as f:
                    f.write(self._agent_product_markdown(card))

                photo_items = self._agent_photo_urls(card)
                if max_photos > 0:
                    photo_items = photo_items[:max_photos]
                media_manifest = {
                    'nmID': nm_id,
                    'vendorCode': vendor,
                    'photos': photo_items,
                    'video': card.get('video'),
                }

                for photo_pos, media in enumerate(photo_items, start=1):
                    total_photos += 1
                    dst_no_ext = os.path.join(photos_dir, f"{photo_pos:02d}")
                    local_path, error = self._agent_download_media(media['url'], dst_no_ext)
                    if local_path:
                        downloaded_photos += 1
                        media['downloaded_file'] = f"фото/{os.path.basename(local_path)}"
                        media['download_error'] = None
                    else:
                        media['downloaded_file'] = None
                        media['download_error'] = error
                        photo_errors.append({
                            'vendorCode': vendor,
                            'nmID': nm_id,
                            'photo_index': photo_pos,
                            'url': media['url'],
                            'error': error,
                        })

                with open(os.path.join(product_dir, 'медиа.json'), 'w', encoding='utf-8') as f:
                    json.dump(media_manifest, f, ensure_ascii=False, indent=2, default=str)

                catalog_rows.append({
                    'Магазин': store_name,
                    'Артикул продавца': vendor,
                    'Артикул WB': nm_id,
                    'Название': card.get('title', ''),
                    'Бренд': card.get('brand', card.get('brandName', '')),
                    'Предмет': card.get('subjectName', ''),
                    'Описание': card.get('description', ''),
                    'Фото в карточке': len(card.get('photos') or []),
                    'Фото в архиве': len([x for x in photo_items if x.get('downloaded_file')]),
                    'Характеристик': len(card.get('characteristics') or []),
                    'Размеров': len(card.get('sizes') or []),
                    'Требуется КИЗ': card.get('needKiz', ''),
                    'Создано': card.get('createdAt', ''),
                    'Обновлено': card.get('updatedAt', ''),
                    'Папка в архиве': folder_name,
                })
                self.log(
                    f"   [{idx}/{len(cards)}] {vendor or nm_id}: "
                    f"характеристик={len(card.get('characteristics') or [])}, "
                    f"фото={len(photo_items)}"
                )

            catalog_df = pd.DataFrame(catalog_rows)
            catalog_path = os.path.join(root_dir, 'Каталог.xlsx')
            with pd.ExcelWriter(catalog_path, engine='openpyxl') as writer:
                catalog_df.to_excel(writer, sheet_name='Товары', index=False)
                if photo_errors:
                    pd.DataFrame(photo_errors).to_excel(writer, sheet_name='Ошибки фото', index=False)

            manifest = {
                'version': SCRIPT_VERSION,
                'store': store_name,
                'generated_at_msk': now_msk.strftime('%Y-%m-%d %H:%M:%S'),
                'signature': signature,
                'cards': len(cards),
                'photos_total': total_photos,
                'photos_downloaded': downloaded_photos,
                'photo_errors': len(photo_errors),
                'complete': len(photo_errors) == 0,
                'archive_key': archive_key,
                'source': 'WB Content API POST /content/v2/get/cards/list',
                'filters': {
                    'include_vendor_codes': sorted(self._agent_env_set('WB_AGENT_INCLUDE_VENDOR_CODES')),
                    'exclude_vendor_codes': sorted(self._agent_env_set('WB_AGENT_EXCLUDE_VENDOR_CODES')),
                    'exclude_nmids': sorted(self._agent_env_set('WB_AGENT_EXCLUDE_NMIDS')),
                    'max_photos_per_product': max_photos,
                },
            }
            with open(os.path.join(root_dir, 'manifest.json'), 'w', encoding='utf-8') as f:
                json.dump(manifest, f, ensure_ascii=False, indent=2, default=str)

            readme = (
                "АРХИВ БАЗЫ ЗНАНИЙ ИИ-АГЕНТА WB\n\n"
                "Каждая папка соответствует артикулу продавца.\n"
                "паспорт_товара.md — удобный текст для ИИ/человека.\n"
                "карточка_WB.json — полная актуальная карточка из Content API.\n"
                "характеристики.json — характеристики товара.\n"
                "размеры_и_баркоды.json — chrtID, размеры и SKU/баркоды.\n"
                "медиа.json — URL всех медиа и имена скачанных фотографий.\n"
                "фото/ — изображения карточки в максимальном доступном качестве из ответа WB.\n\n"
                "Архив перезаписывается при изменении карточек и не создаёт ежедневных дублей.\n"
            )
            with open(os.path.join(root_dir, 'README.txt'), 'w', encoding='utf-8') as f:
                f.write(readme)

            zip_path = os.path.join(tmp_dir, f"WB_Агент_Ассортимент_{store_name}.zip")
            with zipfile.ZipFile(zip_path, 'w', compression=zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
                for dirpath, _, filenames in os.walk(root_dir):
                    for filename in filenames:
                        full_path = os.path.join(dirpath, filename)
                        arcname = os.path.relpath(full_path, root_dir)
                        zf.write(full_path, arcname=arcname)

            archive_size_mb = os.path.getsize(zip_path) / 1024 / 1024
            self.s3.upload_file(zip_path, archive_key)
            self.s3.write_json(manifest_key, manifest)

        self.log(
            f"✅ agent_catalog сохранён: {archive_key}; товаров={len(cards)}, "
            f"фото={downloaded_photos}/{total_photos}, размер={archive_size_mb:.1f} МБ"
        )
        if photo_errors:
            self.log(
                f"⚠️ agent_catalog: не скачано фотографий {len(photo_errors)}. "
                f"Список ошибок есть в Каталог.xlsx; следующий ежедневный запуск повторит сбор, "
                f"пока manifest.complete=false."
            )
        return True


    # ====================== FBS: ЗАКАЗЫ / ОСТАТКИ / ПОСТАВКИ ======================

    def _fbs_headers(self, store_name: str) -> dict:
        token = (self.api_keys[store_name].get('marketplace') or '').strip()
        return {"Authorization": token, "Content-Type": "application/json"}

    def _fbs_request_json(
        self,
        method: str,
        url: str,
        headers: dict,
        *,
        params: Optional[dict] = None,
        payload: Optional[dict] = None,
        timeout: int = 120,
        max_attempts: int = 5,
        context: str = "FBS",
    ) -> Tuple[Optional[Any], Optional[int]]:
        """Унифицированный запрос к Marketplace API с retry на 429/5xx.

        Возвращает (json, status_code). При 204 возвращает ({}, 204).
        При 401/403 не ретраим: чаще всего у токена нет категории Marketplace.
        """
        for attempt in range(1, max_attempts + 1):
            try:
                if method.upper() == 'GET':
                    resp = requests.get(url, headers=headers, params=params, timeout=timeout)
                else:
                    resp = requests.post(url, headers=headers, params=params, json=payload, timeout=timeout)

                if resp.status_code == 200:
                    try:
                        return resp.json(), 200
                    except Exception:
                        self.log(f"❌ {context}: ответ 200, но JSON не разобран: {resp.text[:1000]}")
                        return None, 200
                if resp.status_code == 204:
                    return {}, 204
                if resp.status_code in (401, 403):
                    self.log(
                        f"⏭️ {context}: API недоступен для текущего токена ({resp.status_code}). "
                        f"Проверь категорию/права токена. Ответ: {resp.text[:700]}"
                    )
                    return None, resp.status_code
                if resp.status_code == 429:
                    wait = self._rate_limit_wait_seconds(resp, default_seconds=max(5, 5 * attempt), max_seconds=120)
                    self.log(f"⚠️ {context}: 429, попытка {attempt}/{max_attempts}, ждём {wait} сек")
                    time.sleep(wait)
                    continue
                if resp.status_code in (500, 502, 503, 504):
                    wait = 10 * attempt
                    self.log(f"⚠️ {context}: HTTP {resp.status_code}, попытка {attempt}/{max_attempts}, ждём {wait} сек")
                    time.sleep(wait)
                    continue

                self.log(f"❌ {context}: HTTP {resp.status_code}: {resp.text[:1000]}")
                return None, resp.status_code
            except Exception as e:
                if attempt >= max_attempts:
                    self.log(f"❌ {context}: исключение после {attempt} попыток: {e}")
                    return None, None
                wait = 5 * attempt
                self.log(f"⚠️ {context}: исключение {e}; повтор через {wait} сек")
                time.sleep(wait)
        return None, None

    def _fbs_completed_7_weeks_range(self) -> Tuple[datetime.date, datetime.date]:
        """7 полностью завершённых недель, не считая текущую неделю по МСК."""
        today_msk = self.start_time.date()
        current_monday = today_msk - timedelta(days=today_msk.weekday())
        end_date = current_monday - timedelta(days=1)
        start_date = current_monday - timedelta(weeks=7)
        return start_date, end_date

    def _fbs_orders_weekly_key(self, store_name: str, week_date: datetime.date) -> str:
        year, week, _ = week_date.isocalendar()
        return f"Отчёты/Заказы/{store_name}/Недельные/Заказы_FBS_{year}-W{week:02d}.xlsx"

    def _fbs_stocks_weekly_key(self, store_name: str, week_date: datetime.date) -> str:
        year, week, _ = week_date.isocalendar()
        return f"Отчёты/Остатки/{store_name}/Недельные/Остатки_FBS_{year}-W{week:02d}.xlsx"

    def _fbs_registry_key(self, store_name: str) -> str:
        return f"Отчёты/FBS/{store_name}/Реестр_FBS.xlsx"

    def _fbs_supplies_key(self, store_name: str) -> str:
        return f"Отчёты/FBS/{store_name}/Поставки_FBS.xlsx"

    def _fbs_to_msk_str(self, value: Any) -> str:
        if value is None or value == "":
            return ""
        try:
            ts = pd.to_datetime(value, utc=True, errors='coerce')
            if pd.isna(ts):
                return ""
            return ts.tz_convert('Europe/Moscow').strftime('%Y-%m-%d %H:%M:%S')
        except Exception:
            return ""

    def _fbs_allowed_warehouse_ids(self, store_name: str) -> Optional[Set[int]]:
        """Разрешённые FBS-склады для магазина.

        TOPFACE и MISSTAIS: только подтверждённые липецкие склады со скриншотов
        (1728667 и 1935990). У каждого токена WB видны только собственные склады.
        Для остальных магазинов фильтр не применяется.
        """
        if str(store_name or '').upper() in FBS_LIPETSK_ONLY_STORES:
            return set(FBS_LIPETSK_WAREHOUSE_IDS)
        return None

    def _filter_fbs_orders_by_warehouse(self, store_name: str, orders: List[dict]) -> List[dict]:
        allowed = self._fbs_allowed_warehouse_ids(store_name)
        if not allowed:
            return list(orders or [])
        before = len(orders or [])
        filtered = []
        skipped_ids = set()
        for order in orders or []:
            try:
                wid = int(order.get('warehouseId') or 0)
            except Exception:
                wid = 0
            if wid in allowed:
                filtered.append(order)
            else:
                skipped_ids.add(wid)
        self.log(
            f"🏬 FBS {store_name}: фильтр Липецк — оставлено заказов {len(filtered)} из {before}; "
            f"прочие склады игнорируются"
        )
        if skipped_ids:
            self.log(f"⏭️ Игнорируем ID FBS-складов: {sorted(x for x in skipped_ids if x)}")
        return filtered

    def _filter_fbs_df_by_warehouse(self, store_name: str, df: pd.DataFrame) -> pd.DataFrame:
        allowed = self._fbs_allowed_warehouse_ids(store_name)
        if not allowed or df is None or df.empty or 'ID склада продавца' not in df.columns:
            return df.copy() if isinstance(df, pd.DataFrame) else pd.DataFrame()
        ids = pd.to_numeric(df['ID склада продавца'], errors='coerce')
        return df[ids.isin(sorted(allowed))].copy()

    def _prune_fbs_saved_files_to_allowed_warehouses(self, store_name: str) -> None:
        """Удалить из уже накопленных FBS-файлов TOPFACE/MISSTAIS строки других складов.

        Это нужно один раз после перехода на правило «только Липецк», но безопасно
        выполнять и далее: файлы переписываются только если реально найдены лишние строки.
        """
        allowed = self._fbs_allowed_warehouse_ids(store_name)
        if not allowed:
            return

        candidates = [self._fbs_registry_key(store_name)]
        candidates += [
            k for k in self.s3.list_files(f"Отчёты/Заказы/{store_name}/Недельные/")
            if '/Заказы_FBS_' in k or k.split('/')[-1].startswith('Заказы_FBS_')
        ]
        candidates += [
            k for k in self.s3.list_files(f"Отчёты/Остатки/{store_name}/Недельные/")
            if '/Остатки_FBS_' in k or k.split('/')[-1].startswith('Остатки_FBS_')
        ]

        seen = set()
        for key in candidates:
            if key in seen or not self.s3.file_exists(key):
                continue
            seen.add(key)
            df = self.s3.read_excel(key, sheet_name=0)
            if df.empty or 'ID склада продавца' not in df.columns:
                continue
            filtered = self._filter_fbs_df_by_warehouse(store_name, df)
            if len(filtered) == len(df):
                continue
            sheet = 'Реестр FBS' if key.endswith('/Реестр_FBS.xlsx') else ('Остатки FBS' if 'Остатки_FBS_' in key else 'Заказы FBS')
            self.s3.write_excel(key, filtered, sheet_name=sheet)
            self.log(
                f"🧹 {store_name}: удалены строки других FBS-складов из {key}: "
                f"{len(df)} → {len(filtered)}"
            )

    def _fetch_fbs_orders_range(
        self,
        store_name: str,
        start_date: datetime.date,
        end_date: datetime.date,
    ) -> Tuple[Optional[List[dict]], Optional[int]]:
        """Получить FBS-заказы. WB разрешает максимум 30 календарных дней за запрос."""
        headers = self._fbs_headers(store_name)
        url = "https://marketplace-api.wildberries.ru/api/v3/orders"
        tz_msk = pytz.timezone('Europe/Moscow')
        all_orders: List[dict] = []
        chunk_start = start_date

        while chunk_start <= end_date:
            chunk_end = min(chunk_start + timedelta(days=29), end_date)
            from_dt = tz_msk.localize(datetime.combine(chunk_start, datetime.min.time())).astimezone(pytz.UTC)
            to_dt = tz_msk.localize(datetime.combine(chunk_end, datetime.max.time().replace(microsecond=0))).astimezone(pytz.UTC)
            next_value = 0
            seen_next = set()
            self.log(f"📦 FBS заказы: запрос периода {chunk_start:%Y-%m-%d} — {chunk_end:%Y-%m-%d}")

            while True:
                params = {
                    'limit': 1000,
                    'next': next_value,
                    'dateFrom': int(from_dt.timestamp()),
                    'dateTo': int(to_dt.timestamp()),
                }
                data, status = self._fbs_request_json(
                    'GET', url, headers, params=params, timeout=120,
                    context=f"FBS orders {chunk_start:%Y-%m-%d}..{chunk_end:%Y-%m-%d}"
                )
                if data is None:
                    return None, status
                batch = data.get('orders') or []
                if batch:
                    all_orders.extend(batch)
                new_next = int(data.get('next') or 0)
                if not batch or new_next == 0 or new_next == next_value or new_next in seen_next:
                    break
                seen_next.add(new_next)
                next_value = new_next
                time.sleep(0.25)

            chunk_start = chunk_end + timedelta(days=1)
            if chunk_start <= end_date:
                time.sleep(0.5)

        # На практике API иногда может повторить граничную строку при пагинации.
        dedup = {}
        for order in all_orders:
            oid = order.get('id')
            if oid is not None:
                dedup[int(oid)] = order
        return list(dedup.values()), 200

    def _fetch_fbs_statuses(self, store_name: str, order_ids: List[int]) -> Tuple[Optional[Dict[int, dict]], Optional[int]]:
        if not order_ids:
            return {}, 200
        headers = self._fbs_headers(store_name)
        url = "https://marketplace-api.wildberries.ru/api/v3/orders/status"
        result: Dict[int, dict] = {}
        unique_ids = sorted(set(int(x) for x in order_ids if x))
        for i in range(0, len(unique_ids), 1000):
            batch = unique_ids[i:i+1000]
            data, status = self._fbs_request_json(
                'POST', url, headers, payload={'orders': batch}, timeout=120,
                context=f"FBS statuses {i+1}-{i+len(batch)}"
            )
            if data is None:
                return None, status
            for item in data.get('orders') or []:
                oid = item.get('id')
                if oid is not None:
                    result[int(oid)] = item
            time.sleep(0.25)
        return result, 200

    def _fetch_fbs_supplies(self, store_name: str) -> Tuple[Optional[List[dict]], Optional[int]]:
        headers = self._fbs_headers(store_name)
        url = "https://marketplace-api.wildberries.ru/api/v3/supplies"
        next_value = 0
        supplies: List[dict] = []
        seen_next = set()
        for _ in range(200):
            data, status = self._fbs_request_json(
                'GET', url, headers, params={'limit': 1000, 'next': next_value}, timeout=120,
                context="FBS supplies"
            )
            if data is None:
                return None, status
            batch = data.get('supplies') or []
            supplies.extend(batch)
            new_next = int(data.get('next') or 0)
            if not batch or new_next == 0 or new_next == next_value or new_next in seen_next:
                break
            seen_next.add(new_next)
            next_value = new_next
            time.sleep(0.25)
        dedup = {}
        for sp in supplies:
            sid = str(sp.get('id') or '').strip()
            if sid:
                dedup[sid] = sp
        return list(dedup.values()), 200

    def _save_fbs_supplies(self, store_name: str, supplies: List[dict]) -> Dict[str, dict]:
        now_str = datetime.now(pytz.timezone('Europe/Moscow')).strftime('%Y-%m-%d %H:%M:%S')
        rows = []
        supply_map: Dict[str, dict] = {}
        for sp in supplies or []:
            sid = str(sp.get('id') or '').strip()
            if not sid:
                continue
            row = {
                'ID поставки': sid,
                'Название': sp.get('name', ''),
                'Создана': self._fbs_to_msk_str(sp.get('createdAt')),
                'Закрыта / передана в доставку': self._fbs_to_msk_str(sp.get('closedAt')),
                'Сканирование WB': self._fbs_to_msk_str(sp.get('scanDt')),
                'done': bool(sp.get('done', False)),
                'isB2b': bool(sp.get('isB2b', sp.get('isB2B', False))),
                'cargoType': sp.get('cargoType', ''),
                'crossBorderType': sp.get('crossBorderType', ''),
                'destinationOfficeId': sp.get('destinationOfficeId', ''),
                'Дата обновления': now_str,
                'Магазин': store_name,
            }
            rows.append(row)
            supply_map[sid] = row

        if rows:
            df = pd.DataFrame(rows).drop_duplicates(subset=['ID поставки'], keep='last')
            df = df.sort_values(['Создана', 'ID поставки'], ascending=[False, True])
            self.s3.write_excel(self._fbs_supplies_key(store_name), df, sheet_name='Поставки FBS')
            self.log(f"✅ FBS поставки сохранены: {self._fbs_supplies_key(store_name)}, строк: {len(df)}")
        else:
            self.log("ℹ️ FBS поставок пока нет")
        return supply_map

    def _fbs_result_label(self, wb_status: str) -> str:
        mapping = {
            'sold': 'Выкуплен',
            'canceled_by_client': 'Отказ покупателя при получении',
            'declined_by_client': 'Отмена покупателем в первый час',
            'canceled': 'Отменён',
            'defect': 'Отмена из-за брака',
            'ready_for_pickup': 'Ожидает покупателя в ПВЗ',
            'sorted': 'Отсортирован WB',
            'waiting': 'В работе',
            'postponed_delivery': 'Доставка перенесена',
            'accepted_by_carrier': 'Передан перевозчику',
            'sent_to_carrier': 'Направлен перевозчику',
        }
        return mapping.get(str(wb_status or '').strip(), str(wb_status or '').strip() or 'Неизвестно')

    def _build_fbs_orders_df(
        self,
        store_name: str,
        orders: List[dict],
        statuses: Dict[int, dict],
        supply_map: Dict[str, dict],
        existing_registry: Optional[pd.DataFrame] = None,
    ) -> pd.DataFrame:
        now_str = datetime.now(pytz.timezone('Europe/Moscow')).strftime('%Y-%m-%d %H:%M:%S')
        old_sold_seen: Dict[int, str] = {}
        old_cancel_seen: Dict[int, str] = {}
        if existing_registry is not None and not existing_registry.empty and 'ID заказа FBS' in existing_registry.columns:
            for _, r in existing_registry.iterrows():
                try:
                    oid = int(r.get('ID заказа FBS'))
                except Exception:
                    continue
                old_sold_seen[oid] = str(r.get('Первое обнаружение продажи', '') or '')
                old_cancel_seen[oid] = str(r.get('Первое обнаружение отказа/отмены', '') or '')

        rows = []
        for order in orders or []:
            try:
                oid = int(order.get('id'))
            except Exception:
                continue
            status = statuses.get(oid, {})
            supplier_status = str(status.get('supplierStatus') or '')
            wb_status = str(status.get('wbStatus') or '')
            supply_id = str(order.get('supplyId') or '').strip()
            sp = supply_map.get(supply_id, {}) if supply_id else {}

            order_dt = self._fbs_to_msk_str(order.get('createdAt'))
            closed_dt = str(sp.get('Закрыта / передана в доставку', '') or '')
            scan_dt = str(sp.get('Сканирование WB', '') or '')

            hours_to_ship = None
            hours_to_scan = None
            try:
                if order_dt and closed_dt:
                    hours_to_ship = round((pd.to_datetime(closed_dt) - pd.to_datetime(order_dt)).total_seconds() / 3600, 2)
                if order_dt and scan_dt:
                    hours_to_scan = round((pd.to_datetime(scan_dt) - pd.to_datetime(order_dt)).total_seconds() / 3600, 2)
            except Exception:
                pass

            first_sold = old_sold_seen.get(oid, '')
            if wb_status == 'sold' and not first_sold:
                first_sold = now_str
            first_cancel = old_cancel_seen.get(oid, '')
            if wb_status in {'canceled', 'canceled_by_client', 'declined_by_client', 'defect'} and not first_cancel:
                first_cancel = now_str

            skus = order.get('skus') or []
            offices = order.get('offices') or []
            addr = order.get('address') or {}
            row = {
                'ID заказа FBS': oid,
                'orderUid': order.get('orderUid', ''),
                'Дата заказа': order_dt[:10] if order_dt else '',
                'Время заказа': order_dt,
                'Артикул продавца': order.get('article', ''),
                'Артикул WB': order.get('nmId', ''),
                'chrtId': order.get('chrtId', ''),
                'Баркод': skus[0] if skus else '',
                'Баркоды': ', '.join(str(x) for x in skus),
                'ID склада продавца': order.get('warehouseId', ''),
                'officeId': order.get('officeId', ''),
                'Офисы': ', '.join(str(x) for x in offices),
                'deliveryType': order.get('deliveryType', ''),
                'supplyId': supply_id,
                'Поставка создана': sp.get('Создана', ''),
                'Время отгрузки (closedAt)': closed_dt,
                'Время сканирования WB (scanDt)': scan_dt,
                'Часов от заказа до отгрузки': hours_to_ship if hours_to_ship is not None else '',
                'Часов от заказа до сканирования WB': hours_to_scan if hours_to_scan is not None else '',
                'Отгружен': 'Да' if (closed_dt or supplier_status == 'complete') else 'Нет',
                'supplierStatus': supplier_status,
                'wbStatus': wb_status,
                'Результат': self._fbs_result_label(wb_status),
                'Продажа FBS': 1 if wb_status == 'sold' else 0,
                'Отказ/отмена FBS': 1 if wb_status in {'canceled', 'canceled_by_client', 'declined_by_client', 'defect'} else 0,
                'Первое обнаружение продажи': first_sold,
                'Первое обнаружение отказа/отмены': first_cancel,
                'Цена API': order.get('price', ''),
                'Финальная цена API': order.get('finalPrice', ''),
                'Конвертированная цена API': order.get('convertedPrice', ''),
                'Конвертированная финальная цена API': order.get('convertedFinalPrice', ''),
                'currencyCode': order.get('currencyCode', ''),
                'convertedCurrencyCode': order.get('convertedCurrencyCode', ''),
                'scanPrice': order.get('scanPrice', ''),
                'cargoType': order.get('cargoType', ''),
                'crossBorderType': order.get('crossBorderType', ''),
                'isZeroOrder': order.get('isZeroOrder', False),
                'isB2B': (order.get('options') or {}).get('isB2B', (order.get('options') or {}).get('isB2b', False)),
                'Комментарий': order.get('comment', ''),
                'Адрес': addr.get('fullAddress', '') if isinstance(addr, dict) else '',
                'Дата обновления статуса': now_str,
                'Магазин': store_name,
            }
            rows.append(row)
        return pd.DataFrame(rows)

    def _upsert_fbs_registry(self, store_name: str, new_df: pd.DataFrame) -> pd.DataFrame:
        key = self._fbs_registry_key(store_name)
        old_df = self.s3.read_excel(key, sheet_name=0) if self.s3.file_exists(key) else pd.DataFrame()
        if new_df.empty:
            return old_df
        if old_df.empty:
            combined = new_df.copy()
        else:
            # Не теряем исторические строки вне текущего окна обновления.
            new_ids = set(pd.to_numeric(new_df['ID заказа FBS'], errors='coerce').dropna().astype('int64').tolist())
            old_ids = pd.to_numeric(old_df.get('ID заказа FBS'), errors='coerce') if 'ID заказа FBS' in old_df.columns else pd.Series(dtype='float64')
            keep_mask = ~old_ids.fillna(-1).astype('int64').isin(new_ids) if len(old_ids) else pd.Series([True] * len(old_df), index=old_df.index)
            combined = pd.concat([old_df.loc[keep_mask].copy(), new_df], ignore_index=True, sort=False)
        combined = combined.drop_duplicates(subset=['ID заказа FBS'], keep='last')
        if 'Время заказа' in combined.columns:
            combined['_sort_dt'] = pd.to_datetime(combined['Время заказа'], errors='coerce')
            combined = combined.sort_values('_sort_dt', ascending=False).drop(columns=['_sort_dt'])
        self.s3.write_excel(key, combined, sheet_name='Реестр FBS')
        self.log(f"✅ FBS реестр сохранён: {key}, заказов: {len(combined)}")
        return combined

    def _save_fbs_orders_weekly(self, store_name: str, refreshed_df: pd.DataFrame) -> None:
        if refreshed_df.empty or 'Дата заказа' not in refreshed_df.columns:
            return
        tmp = refreshed_df.copy()
        tmp['_date'] = pd.to_datetime(tmp['Дата заказа'], errors='coerce').dt.date
        tmp = tmp[tmp['_date'].notna()].copy()
        tmp['_week'] = tmp['_date'].apply(lambda d: d - timedelta(days=d.weekday()))

        for week_start, week_df in tmp.groupby('_week'):
            key = self._fbs_orders_weekly_key(store_name, week_start)
            existing = self.s3.read_excel(key, sheet_name=0) if self.s3.file_exists(key) else pd.DataFrame()
            existing = self._filter_fbs_df_by_warehouse(store_name, existing)
            week_df = week_df.drop(columns=['_date', '_week'])
            if existing.empty:
                combined = week_df.copy()
            else:
                new_ids = set(pd.to_numeric(week_df['ID заказа FBS'], errors='coerce').dropna().astype('int64').tolist())
                old_ids = pd.to_numeric(existing.get('ID заказа FBS'), errors='coerce') if 'ID заказа FBS' in existing.columns else pd.Series(dtype='float64')
                keep_mask = ~old_ids.fillna(-1).astype('int64').isin(new_ids) if len(old_ids) else pd.Series([True] * len(existing), index=existing.index)
                combined = pd.concat([existing.loc[keep_mask].copy(), week_df], ignore_index=True, sort=False)
            combined = combined.drop_duplicates(subset=['ID заказа FBS'], keep='last')
            if 'Время заказа' in combined.columns:
                combined['_sort_dt'] = pd.to_datetime(combined['Время заказа'], errors='coerce')
                combined = combined.sort_values('_sort_dt').drop(columns=['_sort_dt'])
            self.s3.write_excel(key, combined, sheet_name='Заказы FBS')
            self.log(f"✅ FBS заказы сохранены: {key}, строк: {len(combined)}")

    def update_fbs_orders(self, store_name: str) -> bool:
        """FBS-заказы + статусы + поставки + единый накопительный реестр.

        Первый запуск: 7 полностью завершённых недель + целевая дата/текущая неделя.
        Последующие: обновляем последние 30 дней, чтобы статусы sold/отказов успевали измениться.
        """
        self.log("")
        self.log(f"📌 ОБНОВЛЕНИЕ: Заказы FBS / поставки / реестр для магазина {store_name}")
        self._prune_fbs_saved_files_to_allowed_warehouses(store_name)
        registry_key = self._fbs_registry_key(store_name)
        existing_registry = self.s3.read_excel(registry_key, sheet_name=0) if self.s3.file_exists(registry_key) else pd.DataFrame()
        existing_registry = self._filter_fbs_df_by_warehouse(store_name, existing_registry)

        if existing_registry.empty:
            hist_start, hist_end = self._fbs_completed_7_weeks_range()
            start_date = hist_start
            end_date = max(hist_end, self.target_date)
            self.log(
                f"📚 Первый FBS запуск: собираем 7 завершённых недель {hist_start:%Y-%m-%d} — {hist_end:%Y-%m-%d} "
                f"и данные до target_date={self.target_date:%Y-%m-%d}"
            )
        else:
            end_date = self.target_date
            start_date = end_date - timedelta(days=29)
            self.log(f"🔄 FBS статусы: обновляем скользящее окно 30 дней {start_date:%Y-%m-%d} — {end_date:%Y-%m-%d}")

        orders, status_code = self._fetch_fbs_orders_range(store_name, start_date, end_date)
        if orders is None:
            if status_code in (401, 403):
                self.log(f"⏭️ FBS заказы {store_name} пропущены: текущий токен не имеет категории Marketplace")
                return True
            return False

        orders = self._filter_fbs_orders_by_warehouse(store_name, orders)
        self.log(f"✅ FBS заказов после фильтра склада: {len(orders)}")
        order_ids = [int(o['id']) for o in orders if o.get('id') is not None]
        statuses, status_code = self._fetch_fbs_statuses(store_name, order_ids)
        if statuses is None:
            if status_code in (401, 403):
                self.log(f"⏭️ FBS статусы {store_name} пропущены: нет Marketplace-доступа")
                return True
            return False
        self.log(f"✅ FBS статусов получено: {len(statuses)}")

        supplies, status_code = self._fetch_fbs_supplies(store_name)
        if supplies is None:
            if status_code in (401, 403):
                self.log(f"⏭️ FBS поставки {store_name} пропущены: нет Marketplace-доступа")
                return True
            return False
        allowed_supply_ids = {str(o.get('supplyId') or '').strip() for o in orders if str(o.get('supplyId') or '').strip()}
        if self._fbs_allowed_warehouse_ids(store_name):
            supplies = [sp for sp in supplies if str(sp.get('id') or '').strip() in allowed_supply_ids]
            self.log(f"🏬 FBS {store_name}: в файл поставок попадут только поставки липецкого склада: {len(supplies)}")
        supply_map = self._save_fbs_supplies(store_name, supplies)

        refreshed_df = self._build_fbs_orders_df(store_name, orders, statuses, supply_map, existing_registry)
        self._save_fbs_orders_weekly(store_name, refreshed_df)
        registry_df = self._upsert_fbs_registry(store_name, refreshed_df)

        sold = int(pd.to_numeric(registry_df.get('Продажа FBS', pd.Series(dtype='float64')), errors='coerce').fillna(0).sum()) if not registry_df.empty else 0
        canceled = int(pd.to_numeric(registry_df.get('Отказ/отмена FBS', pd.Series(dtype='float64')), errors='coerce').fillna(0).sum()) if not registry_df.empty else 0
        shipped = int((registry_df.get('Отгружен', pd.Series(dtype='object')).astype(str) == 'Да').sum()) if not registry_df.empty else 0
        self.log(f"📊 FBS реестр: всего {len(registry_df)}, отгружено {shipped}, выкуплено {sold}, отказ/отмена {canceled}")
        return True

    def _fetch_all_product_cards_for_fbs(self, store_name: str) -> Tuple[List[dict], bool]:
        """Получить chrtId всех карточек для запроса FBS-остатков.

        cards/list доступен по Content или Promotion токену. Возвращает (rows, full_access).
        """
        token = (self.api_keys[store_name].get('promo') or '').strip()
        headers = {"Authorization": token, "Content-Type": "application/json"}
        url = "https://content-api.wildberries.ru/content/v2/get/cards/list"
        cursor = {'limit': 100}
        all_rows: List[dict] = []
        last_cursor = None

        for page in range(1, 1000):
            payload = {
                'settings': {
                    'sort': {'ascending': True},
                    'filter': {'withPhoto': -1},
                    'cursor': cursor,
                }
            }
            data, status = self._fbs_request_json(
                'POST', url, headers, payload=payload, timeout=120,
                context=f"FBS cards chrtId page {page}"
            )
            if data is None:
                return all_rows, False
            cards = data.get('cards') or []
            for card in cards:
                nm_id = card.get('nmID', card.get('nmId', ''))
                vendor = card.get('vendorCode', '')
                title = card.get('title', '')
                brand = card.get('brand', '')
                subject = card.get('subjectName', '')
                for size in card.get('sizes') or []:
                    chrt_id = size.get('chrtID', size.get('chrtId'))
                    if chrt_id is None:
                        continue
                    skus = size.get('skus') or []
                    all_rows.append({
                        'chrtId': int(chrt_id),
                        'Артикул WB': nm_id,
                        'Артикул продавца': vendor,
                        'Название': title,
                        'Бренд': brand,
                        'Предмет': subject,
                        'Размер': size.get('techSize', size.get('wbSize', '')),
                        'Баркод': skus[0] if skus else '',
                        'Баркоды': ', '.join(str(x) for x in skus),
                    })

            cur = data.get('cursor') or {}
            if not cards or len(cards) < 100:
                break
            next_cursor = (cur.get('updatedAt'), cur.get('nmID', cur.get('nmId')))
            if not next_cursor[0] or not next_cursor[1] or next_cursor == last_cursor:
                break
            last_cursor = next_cursor
            cursor = {'limit': 100, 'updatedAt': next_cursor[0], 'nmID': next_cursor[1]}
            time.sleep(0.25)

        dedup = {int(r['chrtId']): r for r in all_rows if r.get('chrtId') is not None}
        return list(dedup.values()), True

    def _known_fbs_products_from_registry(self, store_name: str) -> List[dict]:
        key = self._fbs_registry_key(store_name)
        if not self.s3.file_exists(key):
            return []
        df = self.s3.read_excel(key, sheet_name=0)
        if df.empty or 'chrtId' not in df.columns:
            return []
        rows = []
        for _, r in df.iterrows():
            try:
                chrt_id = int(r.get('chrtId'))
            except Exception:
                continue
            rows.append({
                'chrtId': chrt_id,
                'Артикул WB': r.get('Артикул WB', ''),
                'Артикул продавца': r.get('Артикул продавца', ''),
                'Название': '',
                'Бренд': '',
                'Предмет': '',
                'Размер': '',
                'Баркод': r.get('Баркод', ''),
                'Баркоды': r.get('Баркоды', ''),
            })
        dedup = {int(r['chrtId']): r for r in rows}
        return list(dedup.values())

    def _fetch_fbs_warehouses(self, store_name: str) -> Tuple[Optional[List[dict]], Optional[int]]:
        data, status = self._fbs_request_json(
            'GET', 'https://marketplace-api.wildberries.ru/api/v3/warehouses',
            self._fbs_headers(store_name), timeout=120, context='FBS warehouses'
        )
        if data is None:
            return None, status
        if not isinstance(data, list):
            self.log(f"❌ FBS warehouses: ожидался список, получено {type(data).__name__}")
            return None, status
        # deliveryType=1 — склады FBS; deliveryType=3 — DBW.
        warehouses = [w for w in data if int(w.get('deliveryType') or 0) == 1 and not bool(w.get('isDeleting', False))]

        allowed = self._fbs_allowed_warehouse_ids(store_name)
        if allowed:
            before = len(warehouses)
            warehouses = [w for w in warehouses if int(w.get('id') or 0) in allowed]
            matched = [int(w.get('id') or 0) for w in warehouses]
            self.log(
                f"🏬 FBS {store_name}: используем только липецкий склад. "
                f"Найдено разрешённых складов: {len(warehouses)} из {before}; ID={matched}"
            )
        return warehouses, 200

    def update_fbs_stocks(self, store_name: str) -> bool:
        """Текущий snapshot остатков на складах продавца FBS.

        Важно: API /api/v3/stocks/{warehouseId} возвращает текущий остаток, а не остаток задним числом.
        Поэтому `Дата запроса` — фактическая дата сбора по МСК, а target_date сохраняем отдельно.
        """
        self.log("")
        self.log(f"📌 ОБНОВЛЕНИЕ: Остатки FBS для магазина {store_name}")
        warehouses, status_code = self._fetch_fbs_warehouses(store_name)
        if warehouses is None:
            if status_code in (401, 403):
                self.log(f"⏭️ Остатки FBS {store_name} пропущены: текущий токен не имеет категории Marketplace")
                return True
            return False
        if not warehouses:
            self.log(f"ℹ️ У {store_name} не найдено активных складов FBS (deliveryType=1)")
            return True
        self.log(f"✅ Найдено FBS-складов продавца: {len(warehouses)}")

        product_rows, full_cards = self._fetch_all_product_cards_for_fbs(store_name)
        if not product_rows:
            product_rows = self._known_fbs_products_from_registry(store_name)
            full_cards = False
        if not product_rows:
            self.log("❌ Не удалось получить chrtId товаров ни из cards/list, ни из FBS-реестра")
            return False
        if full_cards:
            self.log(f"✅ Для FBS-остатков получено chrtId всех карточек: {len(product_rows)}")
        else:
            self.log(
                f"⚠️ cards/list недоступен: остатки будут собраны только по {len(product_rows)} chrtId, "
                f"которые уже встречались в FBS-заказах"
            )

        product_map = {int(r['chrtId']): r for r in product_rows}
        chrt_ids = sorted(product_map.keys())
        now_msk = datetime.now(pytz.timezone('Europe/Moscow'))
        request_date = now_msk.date()
        request_ts = now_msk.strftime('%Y-%m-%d %H:%M:%S')
        rows = []
        headers = self._fbs_headers(store_name)

        for wh_idx, wh in enumerate(warehouses, start=1):
            warehouse_id = int(wh.get('id'))
            wh_name = str(wh.get('name') or '')
            self.log(f"🏬 FBS склад {wh_idx}/{len(warehouses)}: {wh_name} (id={warehouse_id}), chrtId={len(chrt_ids)}")
            amounts: Dict[int, int] = {}
            for i in range(0, len(chrt_ids), 1000):
                batch = chrt_ids[i:i+1000]
                data, status = self._fbs_request_json(
                    'POST', f'https://marketplace-api.wildberries.ru/api/v3/stocks/{warehouse_id}',
                    headers, payload={'chrtIds': batch}, timeout=120,
                    context=f"FBS stocks warehouse={warehouse_id} batch={i//1000+1}"
                )
                if data is None:
                    if status in (401, 403):
                        self.log(f"⏭️ Остатки FBS {store_name}: нет Marketplace-доступа")
                        return True
                    return False
                for item in data.get('stocks') or []:
                    try:
                        amounts[int(item.get('chrtId'))] = int(item.get('amount') or 0)
                    except Exception:
                        continue
                time.sleep(0.25)

            # Пишем и нулевые остатки, чтобы snapshot был полным.
            for chrt_id in chrt_ids:
                prod = product_map[chrt_id]
                rows.append({
                    'Дата запроса': request_date.strftime('%Y-%m-%d'),
                    'Дата/время сбора': request_ts,
                    'Целевая дата запуска': self.target_date.strftime('%Y-%m-%d'),
                    'Магазин': store_name,
                    'Тип': 'FBS',
                    'ID склада продавца': warehouse_id,
                    'Склад продавца': wh_name,
                    'officeId': wh.get('officeId', ''),
                    'deliveryType': wh.get('deliveryType', ''),
                    'cargoType склада': wh.get('cargoType', ''),
                    'isProcessing': wh.get('isProcessing', ''),
                    'Артикул WB': prod.get('Артикул WB', ''),
                    'Артикул продавца': prod.get('Артикул продавца', ''),
                    'Название': prod.get('Название', ''),
                    'Бренд': prod.get('Бренд', ''),
                    'Предмет': prod.get('Предмет', ''),
                    'chrtId': chrt_id,
                    'Размер': prod.get('Размер', ''),
                    'Баркод': prod.get('Баркод', ''),
                    'Баркоды': prod.get('Баркоды', ''),
                    'Остаток FBS': int(amounts.get(chrt_id, 0)),
                    'Полнота справочника': 'полный cards/list' if full_cards else 'только товары из FBS-заказов',
                })

        df_day = pd.DataFrame(rows)
        if df_day.empty:
            self.log("ℹ️ FBS остатки: данных нет")
            return True

        key = self._fbs_stocks_weekly_key(store_name, request_date)
        existing = self.s3.read_excel(key, sheet_name=0) if self.s3.file_exists(key) else pd.DataFrame()
        existing = self._filter_fbs_df_by_warehouse(store_name, existing)
        if existing.empty:
            combined = df_day
        else:
            dedup_cols = ['Дата запроса', 'ID склада продавца', 'chrtId']
            combined = pd.concat([existing, df_day], ignore_index=True, sort=False)
            combined = combined.drop_duplicates(subset=dedup_cols, keep='last')
        combined = combined.sort_values(['Дата запроса', 'Склад продавца', 'Артикул продавца', 'chrtId'])
        self.s3.write_excel(key, combined, sheet_name='Остатки FBS')
        total_qty = int(pd.to_numeric(df_day['Остаток FBS'], errors='coerce').fillna(0).sum())
        self.log(f"✅ FBS остатки сохранены: {key}, строк snapshot: {len(df_day)}, всего единиц: {total_qty}")
        return True


    # ====================== ОСНОВНОЙ ЗАПУСК ======================
    def run_daily_update(self, store_name: str, reports: List[str] = None):
        # Исключаем 1c_stocks из списка по умолчанию.
        all_reports = ['orders', 'stocks', 'fbs_orders', 'fbs_stocks', 'finance', 'funnel', 'adverts', 'keywords', 'agent_catalog']
        disabled_for_finick = {'finance', 'keywords'}

        if reports is None:
            reports = list(all_reports)

        if store_name == "FINICK":
            before = list(reports)
            reports = [r for r in reports if r not in disabled_for_finick]
            skipped = [r for r in before if r in disabled_for_finick]
            if skipped:
                self.log(
                    "⏭️ FINICK: временно отключены недоступные методы: "
                    + ", ".join(skipped)
                    + ". Остальные магазины работают без изменений."
                )

        self.log(f"🚀 Начало обновления для магазина {store_name}. Запрошенные отчёты: {reports}")
        overall_success = True
        for report in reports:
            self.log(f"➡️ Переход к отчёту: {report}")
            method_name = f"update_{report}"
            if hasattr(self, method_name):
                method = getattr(self, method_name)
                try:
                    success = method(store_name)
                    if not success:
                        overall_success = False
                    self.log(f"📊 Отчёт {report}: {'✅' if success else '❌'}")
                except Exception as e:
                    overall_success = False
                    self.log(f"❌ Критическая ошибка в {report}: {e}")
                    traceback.print_exc()
                    self.log(f"📊 Отчёт {report}: ❌ (исключение)")
            else:
                overall_success = False
                self.log(f"⚠️ Неизвестный тип отчёта: {report}")
            if report != reports[-1]:
                self.log(f"⏳ Пауза 30 секунд перед следующим отчётом...")
                time.sleep(30)

        if overall_success:
            self.log("✅ Обновление завершено успешно")
        else:
            self.log("❌ Обновление завершено с ошибками в доступных отчётах")
        return overall_success

    def log_section(self, title: str):
        self.log("")
        self.log("=" * 80)
        self.log(f"📌 {title}")
        self.log("=" * 80)



# ========================== НАСТРОЙКИ МАГАЗИНОВ ==========================

STORE_SECRET_ENV = {
    'TOPFACE': 'WB_PROMO_KEY_TOPFACE',
    'MISSTAIS': 'WB_KEY_MISSTAIS',
    'FINICK': 'FINICK_API_WB',
}

# Для FBS нужен токен категории Marketplace. Если отдельный токен не задан,
# пробуем основной токен магазина, чтобы не ломать текущую конфигурацию.
STORE_MARKETPLACE_SECRET_ENV = {
    'TOPFACE': 'WB_MARKETPLACE_KEY_TOPFACE',
    'MISSTAIS': 'WB_MARKETPLACE_KEY_MISSTAIS',
    'FINICK': 'FINICK_API_WB',
}

# Для карточек/характеристик/медиа ИИ-агента нужен токен категории Content.
# Если отдельный токен не задан, пробуем основной токен магазина.
STORE_CONTENT_SECRET_ENV = {
    'TOPFACE': 'WB_CONTENT_KEY_TOPFACE',
    'MISSTAIS': 'WB_CONTENT_KEY_MISSTAIS',
    'FINICK': 'WB_CONTENT_KEY_FINICK',
}

# Новый финансовый API с 15 июля требует токен категории Finance.
# Если отдельного finance-токена нет, берём прежний токен магазина как fallback.
STORE_FINANCE_SECRET_ENV = {
    'TOPFACE': 'WB_FINANCE_KEY_TOPFACE',
    'MISSTAIS': 'WB_FINANCE_KEY_MISSTAIS',
    'FINICK': 'FINICK_API_WB',
}

STORE_ALIASES = {
    'TOPFACE': 'TOPFACE',
    'TF': 'TOPFACE',
    'MISSTAIS': 'MISSTAIS',
    'MISS TAIS': 'MISSTAIS',
    'MISS_TAIS': 'MISSTAIS',
    'MISS-TAIS': 'MISSTAIS',
    'MT': 'MISSTAIS',
    'FINICK': 'FINICK',
    'ALL': 'ALL',
}

KEYWORDS_DEFAULT_START_DATE = datetime(2026, 6, 1).date()
MISSTAIS_KEYWORDS_START_DATE = KEYWORDS_DEFAULT_START_DATE


def _parse_optional_date_env(env_name: str) -> Optional[datetime.date]:
    raw = (os.environ.get(env_name, "") or "").strip()
    if not raw:
        return None
    try:
        return datetime.strptime(raw, "%Y-%m-%d").date()
    except ValueError:
        raise ValueError(f"Некорректная дата в {env_name}: {raw}. Нужен формат YYYY-MM-DD")


def normalize_store_name(value: str) -> str:
    raw = (value or 'TOPFACE').strip().upper().replace('-', '_')
    raw = re.sub(r"\s+", " ", raw)
    return STORE_ALIASES.get(raw, raw)


def get_store_secret_env(store_name: str) -> str:
    store = normalize_store_name(store_name)
    if store not in STORE_SECRET_ENV:
        raise ValueError(f"Неизвестный магазин: {store_name}. Поддерживаются: {', '.join(STORE_SECRET_ENV)}")
    return STORE_SECRET_ENV[store]


def build_api_keys_for_stores(stores: List[str]) -> Dict[str, Dict[str, str]]:
    api_keys: Dict[str, Dict[str, str]] = {}
    for store in stores:
        secret_env = get_store_secret_env(store)
        key_value = os.environ.get(secret_env, '').strip()
        if not key_value:
            raise ValueError(f"Для магазина {store} не задан secret/env {secret_env}")

        finance_secret_env = STORE_FINANCE_SECRET_ENV.get(store)
        finance_key_value = os.environ.get(finance_secret_env or '', '').strip() if finance_secret_env else ''
        if finance_key_value:
            if finance_secret_env == secret_env:
                print(f"✅ Для {store} основной токен {secret_env} используется для всех отчётов, включая Finance")
            else:
                print(f"✅ Для {store} используется отдельный finance-token: {finance_secret_env}")
        else:
            finance_key_value = key_value
            print(f"⚠️ Для {store} отдельный finance-token не задан, используем основной токен {secret_env}. "
                  f"Если новый финансовый метод вернёт 401/403, нужен токен категории Finance.")

        marketplace_secret_env = STORE_MARKETPLACE_SECRET_ENV.get(store)
        marketplace_key_value = os.environ.get(marketplace_secret_env or '', '').strip() if marketplace_secret_env else ''
        if marketplace_key_value:
            if marketplace_secret_env == secret_env:
                print(f"✅ Для {store} основной токен {secret_env} используется для Marketplace/FBS")
            else:
                print(f"✅ Для {store} используется отдельный Marketplace/FBS token: {marketplace_secret_env}")
        else:
            marketplace_key_value = key_value
            print(f"ℹ️ Для {store} отдельный Marketplace/FBS token не задан, пробуем основной токен {secret_env}")

        content_secret_env = STORE_CONTENT_SECRET_ENV.get(store)
        content_key_value = os.environ.get(content_secret_env or '', '').strip() if content_secret_env else ''
        if content_key_value:
            if content_secret_env == secret_env:
                print(f"✅ Для {store} основной токен {secret_env} используется для Content API")
            else:
                print(f"✅ Для {store} используется отдельный Content token: {content_secret_env}")
        else:
            content_key_value = key_value
            print(
                f"ℹ️ Для {store} отдельный Content token не задан, пробуем основной токен {secret_env}. "
                f"Для agent_catalog у него должна быть категория Content."
            )

        api_keys[store] = {
            'promo': key_value,
            'stats': key_value,
            'finance': finance_key_value,
            'marketplace': marketplace_key_value,
            'content': content_key_value,
        }
    return api_keys

# ========================== МЕНЮ ДЛЯ РУЧНОГО ЗАПУСКА ==========================

def show_menu() -> int:
    """Отображает меню и возвращает выбор пользователя."""
    print("\n" + "="*60)
    print("ВЫБЕРИТЕ ДЕЙСТВИЕ:")
    print("="*60)
    print("1. Полное ежедневное обновление (все отчёты)")
    print("2. Обновить конкретный отчёт")
    print("3. Выход")
    print("="*60)
    while True:
        try:
            choice = int(input("Введите номер действия (1-3): "))
            if 1 <= choice <= 3:
                return choice
            else:
                print("Ошибка: введите число от 1 до 3.")
        except (EOFError, KeyboardInterrupt):
            # В неинтерактивном режиме или при прерывании возвращаем 1 (полное обновление)
            return 1
        except ValueError:
            print("Ошибка: введите число.")

def run_specific_report(updater: WildberriesDailyUpdater, store: str):
    """Подменю для выбора конкретного отчёта."""
    reports = ['orders', 'stocks', 'fbs_orders', 'fbs_stocks', 'finance', 'funnel', 'adverts', 'keywords', 'agent_catalog']
    print("\n" + "="*60)
    print("ДОСТУПНЫЕ ОТЧЁТЫ:")
    for i, report in enumerate(reports, 1):
        print(f"{i}. {report}")
    print("0. Назад")
    print("="*60)
    while True:
        try:
            choice = int(input("Выберите номер отчёта: "))
            if choice == 0:
                return
            if 1 <= choice <= len(reports):
                selected = reports[choice-1]
                if store == "FINICK" and selected in {'finance', 'keywords'}:
                    updater.log(
                        f"⏭️ Отчёт {selected} для FINICK временно отключён: "
                        f"нет необходимого доступа/подписки."
                    )
                    return
                updater.log(f"➡️ Запуск обновления отчёта: {selected}")
                method = getattr(updater, f"update_{selected}")
                success = method(store)
                updater.log(f"📊 Отчёт {selected}: {'✅' if success else '❌'}")
                return
            else:
                print(f"Ошибка: введите число от 0 до {len(reports)}.")
        except (EOFError, KeyboardInterrupt):
            return
        except ValueError:
            print("Ошибка: введите число.")

def main():
    """Основная функция запуска с поддержкой меню и магазинов TOPFACE/MISSTAIS/FINICK/ALL."""
    parser = argparse.ArgumentParser(description='Wildberries Daily Updater')
    parser.add_argument('--full', action='store_true', help='Полное ежедневное обновление (все отчёты)')
    parser.add_argument('--report', type=str, choices=['orders', 'stocks', 'fbs_orders', 'fbs_stocks', 'finance', 'funnel', 'adverts', 'keywords', 'agent_catalog'],
                        help='Обновить конкретный отчёт')
    parser.add_argument('--store', type=str, default='TOPFACE', help='Магазин: TOPFACE, MISSTAIS, FINICK или ALL')

    args = parser.parse_args()

    store_arg = normalize_store_name(args.store)
    stores = list(STORE_SECRET_ENV.keys()) if store_arg == 'ALL' else [store_arg]

    required_env = [
        'YC_ACCESS_KEY_ID',
        'YC_SECRET_ACCESS_KEY',
        'YC_BUCKET_NAME',
    ]
    for store in stores:
        required_env.append(get_store_secret_env(store))

    missing = [var for var in required_env if not os.environ.get(var)]
    if missing:
        print(f"❌ Отсутствуют переменные окружения: {missing}")
        exit(1)

    s3 = S3Storage(
        access_key=os.environ['YC_ACCESS_KEY_ID'],
        secret_key=os.environ['YC_SECRET_ACCESS_KEY'],
        bucket_name=os.environ['YC_BUCKET_NAME']
    )

    api_keys = build_api_keys_for_stores(stores)
    updater = WildberriesDailyUpdater(api_keys, s3)

    def run_for_store(store: str) -> bool:
        if args.full:
            return updater.run_daily_update(store)
        if args.report:
            if store == "FINICK" and args.report in {'finance', 'keywords'}:
                updater.log(
                    f"⏭️ Отчёт {args.report} для FINICK временно отключён: "
                    f"нет необходимого доступа/подписки. Другие магазины не затронуты."
                )
                return True
            updater.log(f"➡️ Запуск обновления отчёта: {args.report} | магазин {store}")
            method = getattr(updater, f"update_{args.report}")
            success = method(store)
            updater.log(f"📊 Отчёт {args.report} | {store}: {'✅' if success else '❌'}")
            return success

        if sys.stdin.isatty() and len(stores) == 1:
            while True:
                choice = show_menu()
                if choice == 1:
                    updater.run_daily_update(store)
                elif choice == 2:
                    run_specific_report(updater, store)
                elif choice == 3:
                    print("Выход из программы.")
                    break
                print("\n" + "="*60)
                input("Нажмите Enter, чтобы вернуться в меню...")
        else:
            updater.log("🚀 Запуск в неинтерактивном режиме: выполняем полное ежедневное обновление")
            return updater.run_daily_update(store)

        return True

    overall_run_success = True
    for idx, store in enumerate(stores, start=1):
        if len(stores) > 1:
            updater.log(f"===== Магазин {idx}/{len(stores)}: {store} =====")
        store_success = run_for_store(store)
        if store_success is False:
            overall_run_success = False

    if not overall_run_success:
        sys.exit(1)

if __name__ == "__main__":
    main()
