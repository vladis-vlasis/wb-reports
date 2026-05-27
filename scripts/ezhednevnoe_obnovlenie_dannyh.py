#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Ежедневный сбор данных Wildberries с сохранением в Yandex Cloud Object Storage.
Данные хранятся только в недельных файлах (кроме воронки продаж и 1С).
Автоматическое получение артикулов из заказов для отчёта по ключам.
Формат для keywords: Неделя ГГГГ-WНН.xlsx
Финансовые показатели: проверяется только последняя неделя.
Всегда читается первый лист в файле.
Поисковые запросы: загружается ТОЛЬКО предыдущий день (вчера).
Реклама: получает кампании из API, статистика за последние 30 дней, формирует отчёты по категориям.
Отчёт 1c_stocks временно исключён из списка (можно вернуть позже).
Для всех методов Wildberries используется единый ключ из переменной WB_PROMO_KEY_TOPFACE.
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
                'api_url': 'https://statistics-api.wildberries.ru/api/v1/supplier/stocks',
                'api_method': 'GET',
                'key_type': 'promo',
            },
            'finance': {
                'name': 'Финансовые показатели',
                'folder': 'Финансовые показатели',
                'date_column': 'rr_dt',
                'id_columns': ['rr_dt', 'rrd_id', 'nm_id'],
                'api_url': 'https://statistics-api.wildberries.ru/api/v5/supplier/reportDetailByPeriod',
                'api_method': 'GET',
                'key_type': 'promo',
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
            'finance': 65,
            'keywords': 90,
            'funnel': 30,
            'adverts': 30,
            '1c_stocks': 0,
        }

        self.target_subjects = ['Помады', 'Косметические карандаши', 'Кисти косметические', 'Блески']
        self.log(f"🚀 Запуск обновления данных. Время: {self.start_time}")

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

    def _get_date_range_last_n_days(self, n: int) -> Tuple[datetime.date, datetime.date]:
        today = datetime.now(pytz.timezone('Europe/Moscow')).date()
        end_date = today - timedelta(days=1)
        start_date = end_date - timedelta(days=n - 1)
        return start_date, end_date

    def _get_articles_by_subjects(self, store_name: str, subjects: List[str]) -> List[int]:
        self.log(f"🔍 Сбор артикулов из заказов по категориям: {subjects}")
        prefix = f"Отчёты/Заказы/{store_name}/Недельные/"
        all_files = self.s3.list_files(prefix)
        if not all_files:
            self.log("⚠️ Не найдено недельных файлов заказов")
            return []

        articles_set = set()
        possible_nm_cols = ['nmId', 'nmID', 'Артикул WB', 'Артикул']
        possible_subj_cols = ['subject', 'Предмет', 'subjectName', 'Название предмета']

        for file_key in all_files:
            self.log(f"📄 Обработка файла: {file_key}")
            try:
                df = self.s3.read_excel(file_key, sheet_name=0)
                if df.empty:
                    continue

                nm_col = None
                for col in possible_nm_cols:
                    if col in df.columns:
                        nm_col = col
                        break
                subj_col = None
                for col in possible_subj_cols:
                    if col in df.columns:
                        subj_col = col
                        break

                if nm_col is None or subj_col is None:
                    self.log(f"⚠️ В файле {file_key} не найдены колонки с артикулом или предметом")
                    continue

                df[subj_col] = df[subj_col].astype(str).str.lower().str.strip()
                target_lower = [s.lower() for s in subjects]

                mask = df[subj_col].isin(target_lower)
                filtered = df.loc[mask, nm_col].dropna().unique()
                for val in filtered:
                    try:
                        articles_set.add(int(val))
                    except (ValueError, TypeError):
                        continue

            except Exception as e:
                self.log(f"❌ Ошибка при обработке файла {file_key}: {e}")
                continue

        articles = list(articles_set)
        self.log(f"✅ Собрано {len(articles)} уникальных артикулов из заказов")
        return articles

    # ====================== МЕТОДЫ ДЛЯ КАЖДОГО ОТЧЁТА ======================
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

    def _fetch_finance_day(self, config: dict, headers: dict, date_str: str) -> List[dict]:
        url = config['api_url']
        all_items = []
        rrdid = 0
        limit = 100000
        max_attempts = 3
        while True:
            params = {
                "dateFrom": date_str,
                "dateTo": date_str,
                "limit": limit,
                "rrdid": rrdid,
                "period": "daily"
            }
            for attempt in range(max_attempts):
                try:
                    resp = requests.get(url, headers=headers, params=params, timeout=120)
                    if resp.status_code == 200:
                        data = resp.json()
                        if not data:
                            return all_items
                        all_items.extend(data)
                        last_rrdid = data[-1].get("rrd_id", 0)
                        if len(data) < limit or last_rrdid <= rrdid:
                            return all_items
                        rrdid = last_rrdid
                        break
                    elif resp.status_code == 204:
                        return all_items
                    elif resp.status_code == 429:
                        wait = 60 * (attempt + 1)
                        self.log(f"    ⚠ Лимит, попытка {attempt+1}/{max_attempts}, ждём {wait} сек...")
                        time.sleep(wait)
                    else:
                        self.log(f"    ❌ Ошибка {resp.status_code}: {resp.text[:200]}")
                        if attempt == max_attempts - 1:
                            return all_items
                        time.sleep(10)
                except Exception as e:
                    self.log(f"    ❌ Исключение: {e}")
                    if attempt == max_attempts - 1:
                        return all_items
                    time.sleep(10)
        return all_items

    # ---------- Заказы ----------
    def update_orders(self, store_name: str) -> bool:
        self.log(f"\n📌 ОБНОВЛЕНИЕ: Заказы для магазина {store_name}")
        config = self.reports_config['orders']
        start_date, end_date = self._get_date_range_90_days()
        all_dates = [start_date + timedelta(days=i) for i in range((end_date - start_date).days + 1)]

        weeks = defaultdict(list)
        for d in all_dates:
            week_start = self._get_week_start(datetime.combine(d, datetime.min.time()))
            weeks[week_start].append(d)

        api_key = self.api_keys[store_name][config['key_type']]
        headers = {"Authorization": api_key.strip()}

        for week_start, dates in weeks.items():
            self.log(f"📅 Обработка недели, начинающейся {week_start.strftime('%Y-%m-%d')}")
            weekly_df = self._load_weekly_data(store_name, 'orders', week_start)
            if not weekly_df.empty:
                existing_dates = set(pd.to_datetime(weekly_df['date']).dt.date.unique()) if 'date' in weekly_df.columns else set()
            else:
                existing_dates = set()

            dates_to_load = [d for d in dates if d not in existing_dates]
            if not dates_to_load:
                self.log(f"✅ Все дни недели уже загружены")
                continue

            self.log(f"📅 Недостающие дни: {[d.strftime('%Y-%m-%d') for d in dates_to_load]}")
            new_data = []
            for date in dates_to_load:
                date_str = date.strftime('%Y-%m-%d')
                self.log(f"📅 Загрузка дня: {date_str}")
                data = self._make_request(config, headers, date_str)
                if data and isinstance(data, list):
                    day_df = pd.DataFrame(data)
                    if not day_df.empty:
                        day_df['store'] = store_name
                        if 'date' in day_df.columns:
                            day_df['date'] = pd.to_datetime(day_df['date']).dt.strftime('%Y-%m-%d')
                        new_data.append(day_df)
                        self.log(f"✅ Получено {len(day_df)} записей")
                    else:
                        self.log(f"ℹ️ Нет данных за {date_str}")
                else:
                    self.log(f"⚠️ Не удалось получить данные за {date_str}")

                if date != dates_to_load[-1]:
                    time.sleep(self.delays['orders'])

            if new_data:
                new_df = pd.concat(new_data, ignore_index=True)
                if weekly_df.empty:
                    weekly_df = new_df
                else:
                    weekly_df = pd.concat([weekly_df, new_df], ignore_index=True)
                self._save_weekly_data(weekly_df, store_name, 'orders', week_start)
            else:
                self.log(f"ℹ️ Нет новых данных за неделю")
        return True

    # ---------- Остатки (исправленная версия с пагинацией) ----------
    def update_stocks(self, store_name: str) -> bool:
        """Обновление остатков (ежедневный срез) с поддержкой пагинации."""
        self.log(f"\n📌 ОБНОВЛЕНИЕ: Остатки для магазина {store_name}")
        config = self.reports_config['stocks']

        # Целевая дата — вчера
        target_date = (datetime.now() - timedelta(days=1)).date()
        week_start = self._get_week_start(datetime.combine(target_date, datetime.min.time()))
        target_date_str = target_date.strftime('%Y-%m-%d')

        # Загружаем существующий недельный файл
        weekly_df = self._load_weekly_data(store_name, 'stocks', week_start)
        if not weekly_df.empty and 'Дата запроса' in weekly_df.columns:
            existing_dates = set(pd.to_datetime(weekly_df['Дата запроса']).dt.date.unique())
        else:
            existing_dates = set()

        if target_date in existing_dates:
            self.log(f"✅ Данные за {target_date_str} уже есть в недельном файле, пропускаем")
            return True

        self.log(f"📅 Загрузка остатков за {target_date_str}...")

        api_key = self.api_keys[store_name][config['key_type']]
        headers = {"Authorization": api_key.strip()}

        # Параметры для пагинации: начинаем с очень ранней даты
        date_from = "2000-01-01T00:00:00"
        all_data = []
        page = 1
        max_pages = 50  # предохранитель

        while page <= max_pages:
            params = {"dateFrom": date_from}
            self.log(f"  Страница {page}, dateFrom={date_from}")

            try:
                resp = requests.get(config['api_url'], headers=headers, params=params, timeout=60)
                if resp.status_code == 200:
                    data = resp.json()
                    if not data:  # пустой массив – все данные получены
                        self.log(f"  ✅ Завершено: получено пустых данных")
                        break

                    all_data.extend(data)
                    self.log(f"  ➕ Добавлено {len(data)} записей, всего {len(all_data)}")

                    # Берём lastChangeDate последней записи для следующего запроса
                    last_item = data[-1]
                    date_from = last_item.get("lastChangeDate")
                    if not date_from:
                        self.log("  ⚠️ В последней записи нет lastChangeDate, прерываем пагинацию")
                        break

                    # Лимит 1 запрос в минуту, ждём 65 сек перед следующим
                    time.sleep(self.delays['stocks'])
                    page += 1

                elif resp.status_code == 429:
                    self.log(f"  ⚠️ Лимит запросов (429), ждём 65 сек...")
                    time.sleep(65)
                    # повторяем запрос с теми же параметрами
                    continue
                else:
                    self.log(f"  ❌ Ошибка {resp.status_code}: {resp.text[:200]}")
                    return False
            except Exception as e:
                self.log(f"  ❌ Исключение при запросе: {e}")
                return False

        if not all_data:
            self.log(f"ℹ️ Нет данных за {target_date_str}")
            return True

        # Создаём DataFrame и добавляем служебные колонки
        df_day = pd.DataFrame(all_data)
        df_day['Дата запроса'] = target_date_str
        df_day['Магазин'] = store_name
        df_day['Дата сбора'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        # Переименование колонок
        rename_map = {
            'lastChangeDate': 'Дата последнего изменения',
            'warehouseName': 'Склад',
            'supplierArticle': 'Артикул продавца',
            'nmId': 'Артикул WB',
            'barcode': 'Баркод',
            'quantity': 'Доступно для продажи',
            'inWayToClient': 'В пути к клиенту',
            'inWayFromClient': 'В пути от клиента',
            'quantityFull': 'Полное количество',
            'category': 'Категория',
            'subject': 'Предмет',
            'brand': 'Бренд',
            'techSize': 'Размер',
            'Price': 'Цена',
            'Discount': 'Скидка',
            'isSupply': 'Договор поставки',
            'isRealization': 'Договор реализации',
            'SCCode': 'Код контракта'
        }
        df_day.rename(columns={k: v for k, v in rename_map.items() if k in df_day.columns}, inplace=True)

        # Дедупликация по уникальному набору полей (дата + артикул WB + склад)
        dedup_cols = ['Дата запроса', 'Артикул WB', 'Склад']
        existing_cols = [c for c in dedup_cols if c in df_day.columns]
        if existing_cols:
            before = len(df_day)
            df_day = df_day.drop_duplicates(subset=existing_cols, keep='last')
            after = len(df_day)
            if before > after:
                self.log(f"🔍 Удалено дубликатов в дневных данных: {before - after}")

        # Объединяем с недельным файлом
        if weekly_df.empty:
            weekly_df = df_day
        else:
            weekly_df = pd.concat([weekly_df, df_day], ignore_index=True)
            # Дедупликация во всём недельном файле
            if existing_cols:
                before_week = len(weekly_df)
                weekly_df = weekly_df.drop_duplicates(subset=existing_cols, keep='last')
                after_week = len(weekly_df)
                if before_week > after_week:
                    self.log(f"🔍 Удалено дубликатов в недельном файле: {before_week - after_week}")

        # Сохраняем
        self._save_weekly_data(weekly_df, store_name, 'stocks', week_start)
        self.log(f"✅ Данные за {target_date_str} добавлены в недельный файл")
        return True

    # ---------- Финансовые показатели ----------
    def update_finance(self, store_name: str) -> bool:
        self.log(f"\n📌 ОБНОВЛЕНИЕ: Финансовые показатели для магазина {store_name} (оптимизировано)")
        config = self.reports_config['finance']
        today = datetime.now(pytz.timezone('Europe/Moscow')).date()
        last_date = today - timedelta(days=1)
        last_week_start = self._get_week_start(datetime.combine(last_date, datetime.min.time()))

        self.log(f"📅 Обработка последней недели, начинающейся {last_week_start.strftime('%Y-%m-%d')}")
        weekly_df = self._load_weekly_data(store_name, 'finance', last_week_start)

        if not weekly_df.empty:
            existing_dates = set(pd.to_datetime(weekly_df['rr_dt']).dt.date.unique()) if 'rr_dt' in weekly_df.columns else set()
        else:
            existing_dates = set()

        required_dates = []
        current = last_week_start.date()
        while current <= last_date:
            required_dates.append(current)
            current += timedelta(days=1)

        dates_to_load = [d for d in required_dates if d not in existing_dates]
        if not dates_to_load:
            self.log(f"✅ Все дни последней недели уже загружены")
        else:
            self.log(f"📅 Недостающие дни последней недели: {[d.strftime('%Y-%m-%d') for d in dates_to_load]}")
            api_key = self.api_keys[store_name][config['key_type']]
            headers = {"Authorization": f"Bearer {api_key.strip()}"}
            new_data = []

            for date in dates_to_load:
                date_str = date.strftime('%Y-%m-%d')
                self.log(f"📅 Загрузка дня: {date_str}")
                day_data = self._fetch_finance_day(config, headers, date_str)
                if day_data:
                    day_df = pd.DataFrame(day_data)
                    day_df['store'] = store_name
                    if 'rr_dt' in day_df.columns:
                        day_df['rr_dt'] = pd.to_datetime(day_df['rr_dt']).dt.strftime('%Y-%m-%d')
                    new_data.append(day_df)
                    self.log(f"✅ Получено {len(day_df)} записей")
                else:
                    self.log(f"ℹ️ Нет данных за {date_str}")

                if date != dates_to_load[-1]:
                    time.sleep(self.delays['finance'])

            if new_data:
                new_df = pd.concat(new_data, ignore_index=True)
                if weekly_df.empty:
                    weekly_df = new_df
                else:
                    weekly_df = pd.concat([weekly_df, new_df], ignore_index=True)
                self._save_weekly_data(weekly_df, store_name, 'finance', last_week_start)
            else:
                self.log(f"ℹ️ Нет новых данных за последнюю неделю")

        # Проверка наличия файлов для остальных недель
        start_date, end_date = self._get_date_range_90_days()
        all_dates = [start_date + timedelta(days=i) for i in range((end_date - start_date).days + 1)]
        weeks = set()
        for d in all_dates:
            week_start = self._get_week_start(datetime.combine(d, datetime.min.time()))
            weeks.add(week_start)
        weeks.discard(last_week_start)

        for week_start in weeks:
            key = self._get_weekly_key(store_name, 'finance', week_start)
            if not self.s3.file_exists(key):
                self.log(f"⚠️ Отсутствует файл за неделю {week_start.strftime('%Y-%m-%d')}. Возможно, потребуется историческая загрузка.")

        self.log("✅ Финансовые показатели успешно обновлены")
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
                payload = {
                    "currentPeriod": {"start": date_str, "end": date_str},
                    "nmIds": batch,
                    "topOrderBy": filter_field,
                    "includeSubstitutedSKUs": False,
                    "includeSearchTexts": True,
                    "orderBy": {"field": filter_field, "mode": "desc"},
                    "limit": 100
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
                            self.log(f"    ❌ Ошибка {resp.status_code}, пропускаем")
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

    # ---------- Позиции по ключам (загрузка только за предыдущий день) ----------
    def update_keywords(self, store_name: str) -> bool:
        self.log(f"\n📌 ОБНОВЛЕНИЕ: Позиции по ключам для магазина {store_name} (только за вчера)")

        # 1. Получаем актуальные артикулы из заказов
        articles = self._get_articles_by_subjects(store_name, self.target_subjects)
        if not articles:
            self.log("⚠️ Не найдено артикулов из заказов. Отчёт будет пропущен.")
            return False

        self.log(f"📦 Актуальных артикулов: {len(articles)}")

        # 2. Определяем целевую дату – вчера
        target_date = (datetime.now(pytz.timezone('Europe/Moscow')) - timedelta(days=1)).date()
        target_date_str = target_date.strftime('%Y-%m-%d')
        self.log(f"📅 Целевая дата: {target_date_str}")

        # 3. Определяем неделю, к которой относится целевая дата
        week_start = self._get_week_start(datetime.combine(target_date, datetime.min.time()))
        self.log(f"📅 Неделя начинается: {week_start.strftime('%Y-%m-%d')}")

        # 4. Загружаем существующий недельный файл (если есть)
        weekly_df = self._load_weekly_data(store_name, 'keywords', week_start)

        # 5. Формируем множество существующих комбинаций (дата, артикул, фильтр) для целевой даты
        existing_keys = set()
        if not weekly_df.empty:
            # Фильтруем только строки за целевую дату
            day_df = weekly_df[weekly_df['Дата'] == target_date_str].copy()
            if not day_df.empty:
                day_df['Артикул WB'] = day_df['Артикул WB'].astype(int)
                for _, row in day_df.iterrows():
                    nm = row['Артикул WB']
                    f = row['Фильтр']
                    existing_keys.add((target_date_str, nm, f))
                self.log(f"🔍 В недельном файле найдено {len(existing_keys)} записей за {target_date_str}")
            else:
                self.log(f"ℹ️ За {target_date_str} в недельном файле записей нет")

        filters = ["orders", "openCard", "addToCart"]

        # 6. Определяем, каких фильтров не хватает для каждого артикула
        missing_articles = []
        for nm_id in articles:
            missing_filters = []
            for f in filters:
                if (target_date_str, nm_id, f) not in existing_keys:
                    missing_filters.append(f)
            if missing_filters:
                missing_articles.append(nm_id)
                if len(missing_articles) <= 3:
                    self.log(f"❌ Для артикула {nm_id} пропущены фильтры: {missing_filters}")

        if not missing_articles:
            self.log(f"✅ Все данные за {target_date_str} уже загружены полностью.")
            return True

        self.log(f"📅 Необходимо загрузить данные для {len(missing_articles)} артикулов")

        # 7. Загружаем недостающие данные
        api_key = self.api_keys[store_name]['promo']
        headers = {"Authorization": f"Bearer {api_key.strip()}", "Content-Type": "application/json"}
        url = self.reports_config['keywords']['api_url']

        # Сброс списка ошибок перед началом
        self.keyword_errors = []

        new_data = []
        batches = [missing_articles[i:i+50] for i in range(0, len(missing_articles), 50)]
        for batch_idx, batch in enumerate(batches, 1):
            self.log(f"  📦 Батч {batch_idx}/{len(batches)}: {len(batch)} артикулов")
            batch_data = []
            for filter_field in filters:
                self.log(f"    🔍 Фильтр {filter_field}", end="")
                payload = {
                    "currentPeriod": {"start": target_date_str, "end": target_date_str},
                    "nmIds": batch,
                    "topOrderBy": filter_field,
                    "includeSubstitutedSKUs": False,
                    "includeSearchTexts": True,
                    "orderBy": {"field": filter_field, "mode": "desc"},
                    "limit": 100
                }
                max_retries = 5
                success = False
                for attempt in range(max_retries):
                    try:
                        resp = requests.post(url, headers=headers, json=payload, timeout=120)
                        if resp.status_code == 200:
                            data = resp.json()
                            items = data.get('data', {}).get('items', [])
                            for item in items:
                                text = item.get('text', '').strip()
                                if not text:
                                    continue
                                row = {
                                    "Дата": target_date_str,
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
                            self.log(f" -> ✓ {len(items)} записей")
                            success = True
                            break
                        elif resp.status_code == 429:
                            wait = 60 * (attempt + 1)
                            self.log(f" -> ⚠ Лимит, попытка {attempt+1}/{max_retries}, ждём {wait} сек...")
                            time.sleep(wait)
                        elif resp.status_code in (502, 503, 504):
                            wait = 30 * (attempt + 1)
                            self.log(f" -> ⚠ Ошибка шлюза {resp.status_code}, попытка {attempt+1}/{max_retries}, ждём {wait} сек...")
                            time.sleep(wait)
                        else:
                            self.log(f" -> ❌ Ошибка {resp.status_code}")
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
                if filter_field != filters[-1]:
                    time.sleep(30)

            if batch_data:
                batch_df = pd.DataFrame(batch_data)
                new_data.append(batch_df)

            if batch_idx < len(batches):
                self.log("    ⏳ Пауза 30 сек между батчами...")
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

        return True

    # ---------- Воронка продаж ----------
    def update_funnel(self, store_name: str) -> bool:
        self.log(f"\n📌 ОБНОВЛЕНИЕ: Воронка продаж для магазина {store_name}")
        config = self.reports_config['funnel']
        key = f"Отчёты/{config['folder']}/{store_name}/{config['filename']}"
        if self.s3.file_exists(key):
            df_existing = self.s3.read_excel(key, sheet_name=0)
            if not df_existing.empty:
                start_date, _ = self._get_date_range_90_days()
                date_col = config['date_column']
                if date_col in df_existing.columns:
                    df_existing[date_col] = pd.to_datetime(df_existing[date_col])
                    max_date = df_existing[date_col].max()
                    # Проверяем, есть ли данные за вчерашний день
                    yesterday = (datetime.now(pytz.timezone('Europe/Moscow')) - timedelta(days=1)).date()
                    if max_date and max_date.date() >= yesterday:
                        self.log("✅ Данные воронки уже актуальны")
                        return True
                    else:
                        self.log(f"⚠️ Данные воронки устарели: последняя дата {max_date.date()}, требуется обновление до {yesterday}")
                else:
                    self.log("⚠️ В файле воронки нет колонки с датой, требуется обновление")
            else:
                self.log("⚠️ Файл воронки пуст, требуется обновление")
        else:
            self.log("⚠️ Файл воронки не найден, будет создан")

        self.log("🔄 Запуск формирования отчёта воронки...")
        api_key = self.api_keys[store_name][config['key_type']]
        headers = {"Authorization": f"Bearer {api_key.strip()}", "Content-Type": "application/json"}

        start_date, end_date = self._get_date_range_90_days()
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
                self.log(f"❌ Ошибка создания отчёта: {resp.status_code}")
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
        Обновление данных по рекламным кампаниям понедельно.
        Получает список кампаний напрямую из API Wildberries.
        Сохраняет несколько листов в недельный файл:
        - Статистика_Ежедневно
        - Статистика_Итого
        - Список_кампаний
        Также ведёт историческую таблицу за последние 14 дней с датой запроса.
        """
        self.log(f"\n📌 ОБНОВЛЕНИЕ: Реклама для магазина {store_name}")
        config = self.reports_config['adverts']

        # Проверяем актуальность и полноту аналитического файла
        analytics_key = f"Отчёты/{config['folder']}/{store_name}/Анализ рекламы.xlsx"
        required_sheets = ['Статистика_Ежедневно', 'Статистика_Итого', 'Список_кампаний', 'Отчет_по_Категории', 'Отчет_по_Категории_Итог']

        if self.s3.file_exists(analytics_key):
            try:
                sheets_present = True
                all_sheets_non_empty = True
                latest_date_ok = False

                for sheet in required_sheets:
                    df = self.s3.read_excel(analytics_key, sheet_name=sheet)
                    if df.empty:
                        all_sheets_non_empty = False
                        self.log(f"⚠️ Лист '{sheet}' в аналитическом файле пуст, требуется обновление")
                        sheets_present = False
                        break

                if sheets_present and all_sheets_non_empty:
                    daily_df = self.s3.read_excel(analytics_key, sheet_name='Статистика_Ежедневно')
                    if 'Дата' in daily_df.columns:
                        max_date = pd.to_datetime(daily_df['Дата']).max().date()
                        yesterday = (datetime.now() - timedelta(days=1)).date()
                        if max_date >= yesterday:
                            latest_date_ok = True
                        else:
                            self.log(f"⚠️ Данные в аналитическом файле устарели: последняя дата {max_date}, требуется обновление до {yesterday}")

                if sheets_present and all_sheets_non_empty and latest_date_ok:
                    self.log("✅ Данные рекламы актуальны и полны. Пропускаем обновление.")
                    return True
            except Exception as e:
                self.log(f"⚠️ Ошибка при проверке аналитического файла: {e}, продолжаем обновление")
        else:
            self.log("⚠️ Аналитический файл не найден, будет создан")

        api_key = self.api_keys[store_name][config['key_type']]
        headers = {"Authorization": f"Bearer {api_key.strip()}"}

        # 1. Получаем список всех кампаний (статусы 9 - активные, 11 - на паузе)
        self.log("📋 Запрос списка рекламных кампаний...")
        all_adverts = []
        for payment_type in ['cpm', 'cpc']:
            url = f"{config['api_url']}?statuses=9,11&payment_type={payment_type}"
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

        # 3. Определяем диапазон дат для статистики (последние 30 дней)
        end_date = (datetime.now() - timedelta(days=1)).date()
        start_date = end_date - timedelta(days=29)  # 30 дней включая end_date
        start_str = start_date.strftime('%Y-%m-%d')
        end_str = end_date.strftime('%Y-%m-%d')
        self.log(f"📅 Запрашиваем статистику за период: {start_str} - {end_str}")

        # 4. Загружаем статистику для всех кампаний
        all_stats = []  # список ответов от API (каждый ответ содержит данные по группе кампаний)
        stats_url = "https://advert-api.wildberries.ru/adv/v3/fullstats"
        # Разбиваем ID на группы по 30
        for i in range(0, len(campaign_ids), 30):
            chunk = campaign_ids[i:i+30]
            ids_param = ','.join(map(str, chunk))
            params = {
                'ids': ids_param,
                'beginDate': start_str,
                'endDate': end_str
            }
            retries = 0
            success = False
            while retries < 3 and not success:
                try:
                    self.log(f"⏳ Запрос статистики для кампаний {i+1}-{min(i+30, len(campaign_ids))}...")
                    resp = requests.get(stats_url, headers=headers, params=params, timeout=60)
                    if resp.status_code == 200:
                        data = resp.json()
                        if data:
                            all_stats.extend(data)
                            self.log(f"✅ Получены данные для {len(data)} кампаний")
                        else:
                            self.log(f"ℹ️ Нет данных для этой группы")
                        success = True
                    elif resp.status_code == 429:
                        retries += 1
                        wait = 60 * retries
                        self.log(f"    ⚠️ Лимит, ждём {wait} сек...")
                        time.sleep(wait)
                    else:
                        self.log(f"❌ Ошибка {resp.status_code}: {resp.text[:200]}")
                        break
                except Exception as e:
                    self.log(f"❌ Исключение: {e}")
                    break
            time.sleep(30)  # пауза между группами

        if not all_stats:
            self.log("⚠️ Не получено статистических данных.")
            return False

        self.log(f"📊 Получена статистика для {len(all_stats)} записей кампаний")

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
                if not day_date or day_date < start_str or day_date > end_str:
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
        summary_df = pd.DataFrame(summary_rows)
        campaigns_df = pd.DataFrame(campaigns_list_rows)

        self.log(f"📊 Сформировано {len(daily_df)} ежедневных записей, {len(summary_df)} итоговых записей по кампаниям")

        # 6. Группируем по неделям и сохраняем недельные файлы с несколькими листами
        weeks = defaultdict(list)
        for d in [start_date + timedelta(days=i) for i in range((end_date - start_date).days + 1)]:
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

            # Сохраняем единый аналитический файл со всеми листами
            analytics_key = f"Отчёты/{config['folder']}/{store_name}/Анализ рекламы.xlsx"
            sheets_analytics = {
                'Статистика_Ежедневно': daily_df,
                'Статистика_Итого': summary_df,
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
        """
        Ведёт историческую таблицу рекламных данных за последние 14 дней с датой запроса.
        Файл: Отчёты/Реклама/{store_name}/История_рекламы_14дней.xlsx
        При каждом запуске добавляет новые строки с текущей датой запроса.
        """
        config = self.reports_config['adverts']
        history_key = f"Отчёты/{config['folder']}/{store_name}/История_рекламы_14дней.xlsx"
        request_date = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        # Загружаем существующую историю
        if self.s3.file_exists(history_key):
            try:
                df_history = self.s3.read_excel(history_key, sheet_name=0)
            except:
                df_history = pd.DataFrame()
        else:
            df_history = pd.DataFrame()

        # Добавляем столбец с датой запроса к новым данным
        new_rows = new_daily_df.copy()
        new_rows['Дата запроса'] = request_date

        # Объединяем
        if not df_history.empty:
            combined = pd.concat([df_history, new_rows], ignore_index=True)
        else:
            combined = new_rows

        # Оставляем только последние 14 дней по дате статистики (поле 'Дата')
        cutoff_date = (datetime.now() - timedelta(days=14)).date()
        combined['Дата'] = pd.to_datetime(combined['Дата']).dt.date
        combined = combined[combined['Дата'] >= cutoff_date]
        # Возвращаем строковый формат даты
        combined['Дата'] = combined['Дата'].astype(str)

        # Сохраняем
        self.s3.write_excel(history_key, combined, sheet_name='История')
        self.log(f"📊 История рекламы за 14 дней обновлена, всего записей: {len(combined)}")

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

    # ====================== ОСНОВНОЙ ЗАПУСК ======================
    def run_daily_update(self, store_name: str, reports: List[str] = None):
        # Исключаем 1c_stocks из списка по умолчанию (можно вернуть позже, добавив в список)
        all_reports = ['orders', 'stocks', 'finance', 'funnel', 'adverts', 'keywords']
        if reports is None:
            reports = all_reports

        self.log(f"🚀 Начало обновления для магазина {store_name}. Запрошенные отчёты: {reports}")
        for report in reports:
            self.log(f"➡️ Переход к отчёту: {report}")
            method_name = f"update_{report}"
            if hasattr(self, method_name):
                method = getattr(self, method_name)
                try:
                    success = method(store_name)
                    self.log(f"📊 Отчёт {report}: {'✅' if success else '❌'}")
                except Exception as e:
                    self.log(f"❌ Критическая ошибка в {report}: {e}")
                    traceback.print_exc()
                    self.log(f"📊 Отчёт {report}: ❌ (исключение)")
            else:
                self.log(f"⚠️ Неизвестный тип отчёта: {report}")
            if report != reports[-1]:
                self.log(f"⏳ Пауза 30 секунд перед следующим отчётом...")
                time.sleep(30)

        self.log("✅ Обновление завершено")

    def log_section(self, title: str):
        self.log("")
        self.log("=" * 80)
        self.log(f"📌 {title}")
        self.log("=" * 80)


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
    reports = ['orders', 'stocks', 'finance', 'funnel', 'adverts', 'keywords']
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
    """Основная функция запуска с поддержкой меню."""
    parser = argparse.ArgumentParser(description='Wildberries Daily Updater')
    parser.add_argument('--full', action='store_true', help='Полное ежедневное обновление (все отчёты)')
    parser.add_argument('--report', type=str, choices=['orders', 'stocks', 'finance', 'funnel', 'adverts', 'keywords'], 
                        help='Обновить конкретный отчёт')
    parser.add_argument('--store', type=str, default='TOPFACE', help='Название магазина (по умолчанию TOPFACE)')
    
    args = parser.parse_args()

    required_env = [
        'YC_ACCESS_KEY_ID',
        'YC_SECRET_ACCESS_KEY',
        'YC_BUCKET_NAME',
        'WB_PROMO_KEY_TOPFACE'
    ]
    missing = [var for var in required_env if not os.environ.get(var)]
    if missing:
        print(f"❌ Отсутствуют переменные окружения: {missing}")
        exit(1)

    s3 = S3Storage(
        access_key=os.environ['YC_ACCESS_KEY_ID'],
        secret_key=os.environ['YC_SECRET_ACCESS_KEY'],
        bucket_name=os.environ['YC_BUCKET_NAME']
    )

    api_keys = {
        args.store: {
            # Один универсальный ключ для всех методов API
            'promo': os.environ['WB_PROMO_KEY_TOPFACE'],
            # Оставляем alias 'stats' для обратной совместимости внутренних обращений,
            # но он указывает на тот же универсальный ключ.
            'stats': os.environ['WB_PROMO_KEY_TOPFACE'],
        }
    }

    updater = WildberriesDailyUpdater(api_keys, s3)
    store = args.store

    # Если переданы аргументы, выполняем соответствующее действие
    if args.full:
        updater.run_daily_update(store)
        return
    if args.report:
        updater.log(f"➡️ Запуск обновления отчёта: {args.report}")
        method = getattr(updater, f"update_{args.report}")
        success = method(store)
        updater.log(f"📊 Отчёт {args.report}: {'✅' if success else '❌'}")
        return

    # Если аргументы не переданы, проверяем интерактивность
    if sys.stdin.isatty():
        # Интерактивный режим: показываем меню
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
        # Неинтерактивный режим без аргументов: по умолчанию полное обновление
        updater.log("🚀 Запуск в неинтерактивном режиме: выполняем полное ежедневное обновление")
        updater.run_daily_update(store)

if __name__ == "__main__":
    main()
