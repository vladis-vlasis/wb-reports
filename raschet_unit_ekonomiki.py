import os
import io
import tempfile
import traceback
from datetime import datetime, timedelta
from typing import Dict, List, Tuple

import boto3
import pandas as pd
import pytz
from botocore.client import Config
from botocore.exceptions import ClientError


# =========================================================
# НАСТРОЙКИ
# =========================================================

МАГАЗИН = "TOPFACE"
ЧАСОВОЙ_ПОЯС = "Europe/Moscow"

СТАВКА_НДС = 7.0
СТАВКА_НАЛОГА_НА_ПРИБЫЛЬ = 15.0

МИН_КФ_ЛОГИСТИКИ = 0.8
ПОРОГ_ДОРОГОГО_СКЛАДА = 1.6
ЦЕЛЕВОЙ_КФ_ЛОГИСТИКИ = 1.4

ГЛУБИНА_ПРИЕМКИ_НЕДЕЛЬ = 9
ГЛУБИНА_ХРАНЕНИЯ_НЕДЕЛЬ = 13

ПУТЬ_ЭКОНОМИКА = f"Отчёты/Финансовые показатели/{МАГАЗИН}/Экономика.xlsx"
ПРЕФИКС_ФИНАНСЫ = f"Отчёты/Финансовые показатели/{МАГАЗИН}/Недельные/"
ПРЕФИКС_ОСТАТКИ = f"Отчёты/Остатки/{МАГАЗИН}/Недельные/"
ПУТЬ_РЕКЛАМА_АНАЛИЗ = f"Отчёты/Реклама/{МАГАЗИН}/Анализ рекламы.xlsx"
ПРЕФИКС_РЕКЛАМА_НЕДЕЛЬНЫЕ = f"Отчёты/Реклама/{МАГАЗИН}/Недельные/"
ПУТЬ_СЕБЕСТОИМОСТЬ = "Отчёты/Себестоимость/Себестоимость.xlsx"

# Флаг для разового пересчёта всех последних недель (установить True при первом запуске после изменения логики)
ПЕРЕСЧИТАТЬ_ИСТОРИЮ = True  # Измените на True для однократного пересчёта

ЛИСТ_ЮНИТ = "Юнит экономика"
ЛИСТ_ФАКТ = "Общий факт за неделю"
ЛИСТ_АНАЛИЗ = "Анализ неделя к неделе"
ЛИСТ_СКЛАДЫ = "Склады и коэффициенты"

ОПЕРАЦИИ_ПРОДАЖА = {
    "Продажа",
    "Компенсация ущерба",
    "Добровольная компенсация при возврате",
}

ОПЕРАЦИИ_ВОЗВРАТ = {
    "Возврат",
}

ПОДСКАЗКИ_ПРЯМАЯ_ЛОГИСТИКА = [
    "к клиенту при продаже",
    "к клиенту",
]

ПОДСКАЗКИ_ОБРАТНАЯ_ЛОГИСТИКА = [
    "от клиента при отмене",
    "от клиента при возврате",
    "к клиенту при отмене",
    "возврат товара",
    "возврат",
    "от клиента",
]

КОЛОНКИ_ЮНИТ = [
    "Неделя",
    "Артикул WB",
    "Артикул продавца",
    "Предмет",
    "Бренд",
    "Продажи, шт",
    "Возвраты, шт",
    "Чистые продажи, шт",
    "Процент выкупа",
    "Средняя цена продажи",
    "Средняя цена покупателя",
    "СПП, %",
    "Комиссия WB, %",
    "Эквайринг, %",
    "Комиссия WB, руб/ед",
    "Эквайринг, руб/ед",
    "Логистика прямая, руб/ед",
    "Логистика обратная, руб/ед",
    "Хранение, руб/ед",
    "Приёмка, руб/ед",
    "Штрафы и удержания, руб/ед",
    "Реклама, руб/ед",
    "Прочие расходы, руб/ед",
    "Себестоимость, руб",
    "НДС, руб/ед",
    "Валовая прибыль, руб/ед",
    "Чистая прибыль, руб/ед",
    "Валовая рентабельность, %",
    "Чистая рентабельность, %",
]

КОЛОНКИ_ФАКТ = [
    "Неделя",
    "Артикул WB",
    "Артикул продавца",
    "Предмет",
    "Бренд",
    "Продажи, шт",
    "Возвраты, шт",
    "Чистые продажи, шт",
    "Процент выкупа",
    "Валовая выручка",
    "Средняя цена продажи",
    "Средняя цена покупателя",
    "СПП, %",
    "Комиссия WB",
    "Комиссия WB, %",
    "Эквайринг",
    "Эквайринг, %",
    "Логистика прямая",
    "Логистика обратная",
    "Хранение",
    "Приёмка",
    "Штрафы",
    "Удержания",
    "Реклама",
    "Рекламные заказы, шт",
    "Рекламная выручка",
    "Рекламные показы",
    "Рекламные клики",
    "Прочие расходы",
    "Себестоимость, руб",
    "Себестоимость всего",
    "Валовая прибыль",
    "НДС",
    "Прибыль до налога",
    "Налог на прибыль",
    "Чистая прибыль",
    "Валовая рентабельность, %",
    "Чистая рентабельность, %",
    "Ставка НДС, %",
    "Ставка налога на прибыль, %",
]

КОЛОНКИ_АНАЛИЗ = [
    "Раздел",
    "Неделя",
    "Артикул WB",
    "Артикул продавца",
    "Предмет",
    "Показатель",
    "Чистая прибыль текущая",
    "Чистая прибыль предыдущая",
    "Изменение чистой прибыли",
    "Изменение выручки",
    "Изменение рекламы",
    "Изменение комиссии",
    "Изменение логистики",
    "Изменение СПП",
    "Изменение вашей цены",
    "Комментарий",
]

КОЛОНКИ_СКЛАДЫ = [
    "Неделя",
    "Склад",
    "Средний коэффициент",
    "Количество продаж, шт",
    "Логистика факт",
    "Логистика при коэффициенте 1,4",
    "Переплата",
    "Переплата на единицу",
]


# =========================================================
# БАЗОВЫЕ ФУНКЦИИ
# =========================================================

def лог(msg: str):
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}", flush=True)


def сейчас_мск():
    return datetime.now(pytz.timezone(ЧАСОВОЙ_ПОЯС))


def округлить(значение, знаков=2):
    try:
        if pd.isna(значение):
            return 0.0
        return round(float(значение), знаков)
    except Exception:
        return 0.0


def безопасное_деление(a, b, знаков=6):
    try:
        a = float(a)
        b = float(b)
        if b == 0:
            return 0.0
        return round(a / b, знаков)
    except Exception:
        return 0.0


def текст(x) -> str:
    if pd.isna(x):
        return ""
    return str(x).strip().lower()


def неделя_код(дата: datetime.date) -> str:
    год, неделя, _ = дата.isocalendar()
    return f"{год}-W{неделя:02d}"


def последняя_полная_неделя() -> Tuple[datetime.date, datetime.date]:
    сегодня = сейчас_мск().date()
    понедельник_текущей = сегодня - timedelta(days=сегодня.weekday())
    конец_прошлой = понедельник_текущей - timedelta(days=1)
    начало_прошлой = конец_прошлой - timedelta(days=6)
    return начало_прошлой, конец_прошлой


def путь_финансы_неделя(начало_недели: datetime.date) -> str:
    return f"{ПРЕФИКС_ФИНАНСЫ}Финансовые показатели_{неделя_код(начало_недели)}.xlsx"


def путь_остатки_неделя(начало_недели: datetime.date) -> str:
    return f"{ПРЕФИКС_ОСТАТКИ}Остатки_{неделя_код(начало_недели)}.xlsx"


def путь_реклама_неделя(начало_недели: datetime.date) -> str:
    return f"{ПРЕФИКС_РЕКЛАМА_НЕДЕЛЬНЫЕ}Реклама_{неделя_код(начало_недели)}.xlsx"


def в_дату(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce").dt.date


def в_датавремя(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce")


def привести_к_числам(df: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
    for col in cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    return df


def мода_или_последнее(series: pd.Series):
    s = series.dropna().astype(str).str.strip()
    if s.empty:
        return ""
    moda = s.mode()
    if not moda.empty:
        return moda.iloc[0]
    return s.iloc[-1]


def ндс_из_цены_с_ндс(выручка_с_ндс: float, ставка_ндс: float) -> float:
    return округлить(выручка_с_ндс * ставка_ндс / (100.0 + ставка_ндс), 6)


def знак_строки(doc_type_name: str, supplier_oper_name: str) -> int:
    doc = текст(doc_type_name)
    oper = текст(supplier_oper_name)

    if supplier_oper_name in ОПЕРАЦИИ_ПРОДАЖА or doc == "продажа":
        return 1
    if supplier_oper_name in ОПЕРАЦИИ_ВОЗВРАТ or doc == "возврат":
        return -1
    if oper in {x.lower() for x in ОПЕРАЦИИ_ПРОДАЖА}:
        return 1
    if oper in {x.lower() for x in ОПЕРАЦИИ_ВОЗВРАТ}:
        return -1
    return 0


def тип_логистики(row) -> str:
    supplier_oper = текст(row.get("supplier_oper_name", ""))
    bonus_type = текст(row.get("bonus_type_name", ""))
    delivery_amount = float(row.get("delivery_amount", 0) or 0)
    return_amount = float(row.get("return_amount", 0) or 0)

    if supplier_oper != "логистика":
        return "нет"

    for hint in ПОДСКАЗКИ_ОБРАТНАЯ_ЛОГИСТИКА:
        if hint in bonus_type:
            return "обратная"

    for hint in ПОДСКАЗКИ_ПРЯМАЯ_ЛОГИСТИКА:
        if hint in bonus_type:
            return "прямая"

    if return_amount > 0:
        return "обратная"
    if delivery_amount > 0:
        return "прямая"

    return "прямая"


def список_недель(последняя_неделя: datetime.date, сколько: int) -> List[datetime.date]:
    weeks = []
    current = последняя_неделя
    for _ in range(сколько):
        weeks.append(current)
        current = current - timedelta(days=7)
    return sorted(weeks)


# =========================================================
# S3 / OBJECT STORAGE
# =========================================================

class S3Storage:
    def __init__(self, access_key: str, secret_key: str, bucket_name: str):
        self.bucket = bucket_name
        self.s3 = boto3.client(
            "s3",
            endpoint_url="https://storage.yandexcloud.net",
            aws_access_key_id=access_key,
            aws_secret_access_key=secret_key,
            region_name="ru-central1",
            config=Config(
                signature_version="s3v4",
                read_timeout=300,
                connect_timeout=60,
                retries={"max_attempts": 5},
            ),
        )

    def file_exists(self, key: str) -> bool:
        try:
            self.s3.head_object(Bucket=self.bucket, Key=key)
            return True
        except ClientError:
            return False

    def read_excel(self, key: str, sheet_name=0):
        obj = self.s3.get_object(Bucket=self.bucket, Key=key)
        data = obj["Body"].read()
        return pd.read_excel(io.BytesIO(data), sheet_name=sheet_name)

    def read_excel_all_sheets(self, key: str) -> Dict[str, pd.DataFrame]:
        obj = self.s3.get_object(Bucket=self.bucket, Key=key)
        data = obj["Body"].read()
        return pd.read_excel(io.BytesIO(data), sheet_name=None)

    def write_excel_sheets(self, key: str, sheets: Dict[str, pd.DataFrame]):
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp_path = tmp.name
        try:
            with pd.ExcelWriter(tmp_path, engine="openpyxl") as writer:
                for sheet_name, df in sheets.items():
                    safe_sheet = str(sheet_name)[:31]
                    if df is None:
                        df = pd.DataFrame()
                    df.to_excel(writer, index=False, sheet_name=safe_sheet)
            self.s3.upload_file(tmp_path, self.bucket, key)
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)


# =========================================================
# ОЧИСТКА СХЕМЫ ЛИСТОВ
# =========================================================

def подготовить_пустую_схему(columns: List[str]) -> pd.DataFrame:
    return pd.DataFrame(columns=columns)


def очистить_схему(existing_df: pd.DataFrame, allowed_columns: List[str]) -> pd.DataFrame:
    if existing_df is None or existing_df.empty:
        return подготовить_пустую_схему(allowed_columns)

    out = existing_df.copy()

    # Оставляем только нужные русские столбцы
    keep_cols = [c for c in allowed_columns if c in out.columns]
    out = out[keep_cols].copy()

    # Добавляем недостающие колонки
    for col in allowed_columns:
        if col not in out.columns:
            out[col] = pd.NA

    out = out[allowed_columns].copy()
    return out


# =========================================================
# СЕБЕСТОИМОСТЬ
# =========================================================

def нормализовать_себестоимость(cost_df: pd.DataFrame) -> pd.DataFrame:
    if cost_df.empty:
        return pd.DataFrame(columns=["Артикул WB", "Себестоимость, руб"])

    df = cost_df.copy()
    original_columns = list(df.columns)

    normalized = {}
    for col in df.columns:
        norm = str(col).strip().lower().replace("ё", "е")
        normalized[col] = norm

    nm_col = None
    cost_col = None

    for col, norm in normalized.items():
        if norm in ["nm_id", "nmid", "артикул wb", "артикул вб", "код wb"]:
            nm_col = col
            break

    if nm_col is None:
        for col, norm in normalized.items():
            if "артикул wb" in norm or "артикул вб" in norm:
                nm_col = col
                break

    for col, norm in normalized.items():
        if norm in ["cost_price", "себестоимость", "стоимость", "cost", "закупочная цена"]:
            cost_col = col
            break

    if cost_col is None:
        for col, norm in normalized.items():
            if "себестоим" in norm or "стоимость" in norm or norm.startswith("cost"):
                cost_col = col
                break

    if nm_col is None:
        for col in original_columns:
            if "вб" in str(col).lower():
                nm_col = col
                break

    if cost_col is None and len(original_columns) >= 4:
        cost_col = original_columns[-1]

    if nm_col is None or cost_col is None:
        лог(f"⚠️ Не удалось определить колонки в Себестоимости. Колонки: {original_columns}")
        return pd.DataFrame(columns=["Артикул WB", "Себестоимость, руб"])

    df = df.rename(columns={nm_col: "Артикул WB", cost_col: "Себестоимость, руб"})
    df["Артикул WB"] = pd.to_numeric(df["Артикул WB"], errors="coerce")
    df["Себестоимость, руб"] = pd.to_numeric(df["Себестоимость, руб"], errors="coerce")

    df = df[["Артикул WB", "Себестоимость, руб"]].dropna(subset=["Артикул WB"])
    df["Артикул WB"] = df["Артикул WB"].astype("int64")
    df["Себестоимость, руб"] = df["Себестоимость, руб"].fillna(0.0)

    return df.drop_duplicates(subset=["Артикул WB"], keep="last")


def прочитать_себестоимость(s3: S3Storage) -> pd.DataFrame:
    лог("📥 Шаг: чтение себестоимости")
    if not s3.file_exists(ПУТЬ_СЕБЕСТОИМОСТЬ):
        лог(f"⚠️ Не найден файл себестоимости: {ПУТЬ_СЕБЕСТОИМОСТЬ}")
        return pd.DataFrame(columns=["Артикул WB", "Себестоимость, руб"])

    raw = s3.read_excel(ПУТЬ_СЕБЕСТОИМОСТЬ, sheet_name=0)
    if raw.empty:
        return pd.DataFrame(columns=["Артикул WB", "Себестоимость, руб"])

    out = нормализовать_себестоимость(raw)
    лог(f"✅ Себестоимость загружена, SKU: {len(out)}")
    return out


# =========================================================
# ЧТЕНИЕ ИСТОЧНИКОВ
# =========================================================

def прочитать_финансы_недели(s3: S3Storage, начало_недели: datetime.date) -> pd.DataFrame:
    key = путь_финансы_неделя(начало_недели)
    if not s3.file_exists(key):
        лог(f"⚠️ Не найден фин. отчёт: {key}")
        return pd.DataFrame()

    df = s3.read_excel(key, sheet_name=0)
    if df.empty:
        return df

    numeric_cols = [
        "nm_id", "quantity", "retail_price", "retail_amount", "retail_price_withdisc_rub",
        "commission_percent", "ppvz_for_pay", "acquiring_fee", "acquiring_percent",
        "delivery_rub", "delivery_amount", "return_amount", "penalty", "additional_payment",
        "rebill_logistic_cost", "storage_fee", "deduction", "acceptance", "ppvz_spp_prc", "dlv_prc"
    ]
    df = привести_к_числам(df, numeric_cols)

    if "rr_dt" in df.columns:
        df["rr_dt"] = в_дату(df["rr_dt"])
    if "sale_dt" in df.columns:
        df["sale_dt"] = в_датавремя(df["sale_dt"])
    if "order_dt" in df.columns:
        df["order_dt"] = в_датавремя(df["order_dt"])
    if "nm_id" in df.columns:
        df["nm_id"] = pd.to_numeric(df["nm_id"], errors="coerce")

    return df


def прочитать_остатки_недели(s3: S3Storage, начало_недели: datetime.date) -> pd.DataFrame:
    key = путь_остатки_неделя(начало_недели)
    if not s3.file_exists(key):
        лог(f"⚠️ Не найден отчёт остатков: {key}")
        return pd.DataFrame()

    df = s3.read_excel(key, sheet_name=0)
    if df.empty:
        return df

    if "Дата сбора" in df.columns:
        df["Дата сбора"] = в_дату(df["Дата сбора"])
    elif "Дата запроса" in df.columns:
        df["Дата сбора"] = в_дату(df["Дата запроса"])

    df = привести_к_числам(df, ["Артикул WB", "Доступно для продажи", "Полное количество"])
    if "Артикул WB" in df.columns:
        df["Артикул WB"] = pd.to_numeric(df["Артикул WB"], errors="coerce")

    return df


def подготовить_рекламный_лист(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()

    aliases = {
        "Артикул WB": ["Артикул WB", "Артикул", "nm_id"],
        "Дата": ["Дата", "День"],
        "Расход": ["Расход", "Затраты"],
        "Сумма заказов": ["Сумма заказов", "Выручка"],
        "Показы": ["Показы", "Просмотры"],
        "Клики": ["Клики"],
        "Заказы": ["Заказы"],
    }

    rename_map = {}
    lower_cols = {str(c).strip().lower(): c for c in df.columns}

    for target, variants in aliases.items():
        for v in variants:
            key = v.strip().lower()
            if key in lower_cols:
                rename_map[lower_cols[key]] = target
                break

    df = df.rename(columns=rename_map)

    required = ["Артикул WB", "Дата", "Расход"]
    if not all(col in df.columns for col in required):
        return pd.DataFrame()

    for col in ["Артикул WB", "Расход", "Сумма заказов", "Показы", "Клики", "Заказы"]:
        if col not in df.columns:
            df[col] = 0

    df["Дата"] = в_дату(df["Дата"])
    df = привести_к_числам(df, ["Артикул WB", "Расход", "Сумма заказов", "Показы", "Клики", "Заказы"])
    df["Артикул WB"] = pd.to_numeric(df["Артикул WB"], errors="coerce")
    return df


def найти_рекламный_лист_с_данными(sheets: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    preferred = ["Статистика_Ежедневно", "Data", "Sheet1"]
    for name in preferred:
        if name in sheets:
            prepared = подготовить_рекламный_лист(sheets[name].copy())
            if not prepared.empty:
                return prepared

    for _, df in sheets.items():
        prepared = подготовить_рекламный_лист(df.copy())
        if not prepared.empty:
            return prepared

    return pd.DataFrame()


def прочитать_рекламу_недели(s3: S3Storage, начало_недели: datetime.date, конец_недели: datetime.date) -> pd.DataFrame:
    лог("📣 Шаг: чтение рекламы")
    try:
        sheets = s3.read_excel_all_sheets(ПУТЬ_РЕКЛАМА_АНАЛИЗ)
        df = найти_рекламный_лист_с_данными(sheets)
        if not df.empty:
            out = df[(df["Дата"] >= начало_недели) & (df["Дата"] <= конец_недели)].copy()
            лог(f"✅ Реклама прочитана из Анализ рекламы.xlsx, строк: {len(out)}")
            return out
        else:
            лог("⚠️ В Анализ рекламы.xlsx не найден лист с данными по SKU. Пытаюсь weekly-файл.")
    except Exception as e:
        лог(f"⚠️ Не удалось прочитать Анализ рекламы.xlsx: {e}")

    weekly_key = путь_реклама_неделя(начало_недели)
    if not s3.file_exists(weekly_key):
        лог(f"⚠️ Не найден weekly-рекламный файл: {weekly_key}. Продолжаю без рекламы по SKU.")
        return pd.DataFrame()

    try:
        sheets = s3.read_excel_all_sheets(weekly_key)
        df = найти_рекламный_лист_с_данными(sheets)
        if not df.empty:
            out = df[(df["Дата"] >= начало_недели) & (df["Дата"] <= конец_недели)].copy()
            лог(f"✅ Реклама прочитана из weekly-файла, строк: {len(out)}")
            return out
        else:
            лог("⚠️ В weekly-рекламном файле тоже нет данных по SKU. Продолжаю без рекламы по SKU.")
    except Exception as e:
        лог(f"⚠️ Не удалось прочитать weekly рекламу: {e}")

    return pd.DataFrame()


# =========================================================
# ПОДГОТОВКА ФИНАНСОВЫХ СТРОК
# =========================================================

def подготовить_финансовые_строки(fin_df: pd.DataFrame) -> pd.DataFrame:
    if fin_df.empty:
        return pd.DataFrame()

    df = fin_df.copy()

    df["supplier_oper_name_norm"] = df.get("supplier_oper_name", "").astype(str).str.strip()
    df["doc_type_name_norm"] = df.get("doc_type_name", "").astype(str).str.strip()
    df["subject_name"] = df.get("subject_name", "").astype(str)
    df["brand_name"] = df.get("brand_name", "").astype(str)
    df["sa_name"] = df.get("sa_name", "").astype(str)
    df["office_name"] = df.get("office_name", "").astype(str)
    df["bonus_type_name"] = df.get("bonus_type_name", "").astype(str)

    df["sign"] = df.apply(
        lambda r: знак_строки(r.get("doc_type_name_norm", ""), r.get("supplier_oper_name_norm", "")),
        axis=1
    )

    df["signed_revenue"] = df["retail_price_withdisc_rub"] * df["sign"]
    df["signed_buyer_price"] = df["retail_amount"] * df["sign"]
    df["signed_quantity"] = df["quantity"] * df["sign"]

    commission_raw = (df["retail_price_withdisc_rub"] - df["ppvz_for_pay"] - df["acquiring_fee"]).fillna(0)
    df["signed_commission"] = commission_raw.abs() * df["sign"]
    df["signed_acquiring"] = df["acquiring_fee"].abs() * df["sign"]

    return df


# =========================================================
# ХРАНЕНИЕ И СКЛАДЫ
# =========================================================

def распределить_хранение(
    stocks_df: pd.DataFrame,
    total_storage_week: float,
    начало_недели: datetime.date,
    конец_недели: datetime.date
) -> pd.DataFrame:
    if stocks_df.empty or "Дата сбора" not in stocks_df.columns or "Артикул WB" not in stocks_df.columns:
        return pd.DataFrame(columns=["Артикул WB", "Хранение"])

    df = stocks_df.copy()
    df = df[df["Дата сбора"].notna()].copy()

    if df.empty:
        return pd.DataFrame(columns=["Артикул WB", "Хранение"])

    in_week = df[(df["Дата сбора"] >= начало_недели) & (df["Дата сбора"] <= конец_недели)].copy()

    if in_week.empty:
        target_date = df["Дата сбора"].min()
        in_week = df[df["Дата сбора"] == target_date].copy()
    else:
        target_date = in_week["Дата сбора"].min()
        in_week = in_week[in_week["Дата сбора"] == target_date].copy()

    qty_col = "Доступно для продажи" if "Доступно для продажи" in in_week.columns else "Полное количество"
    in_week[qty_col] = pd.to_numeric(in_week[qty_col], errors="coerce").fillna(0)
    in_week["Артикул WB"] = pd.to_numeric(in_week["Артикул WB"], errors="coerce")
    in_week = in_week.dropna(subset=["Артикул WB"]).copy()
    in_week["Артикул WB"] = in_week["Артикул WB"].astype("int64")

    agg = (
        in_week.groupby("Артикул WB", as_index=False)[qty_col]
        .sum()
        .rename(columns={qty_col: "Остаток, шт"})
    )

    total_stock_units = agg["Остаток, шт"].sum()
    if total_stock_units <= 0:
        agg["Хранение"] = 0.0
        return agg[["Артикул WB", "Хранение"]]

    agg["Хранение"] = agg["Остаток, шт"] * total_storage_week / total_stock_units
    return agg[["Артикул WB", "Хранение"]]


def анализ_складов(logistic_rows: pd.DataFrame, начало_недели: datetime.date) -> pd.DataFrame:
    if logistic_rows.empty:
        return подготовить_пустую_схему(КОЛОНКИ_СКЛАДЫ)

    df = logistic_rows.copy()
    df["dlv_prc"] = pd.to_numeric(df["dlv_prc"], errors="coerce")
    df["delivery_rub"] = pd.to_numeric(df["delivery_rub"], errors="coerce").fillna(0)
    df["quantity"] = pd.to_numeric(df["quantity"], errors="coerce").fillna(0)

    df["тип_логистики"] = df.apply(тип_логистики, axis=1)
    df = df[df["тип_логистики"] == "прямая"].copy()
    df = df[df["dlv_prc"].notna()].copy()
    df = df[df["dlv_prc"] >= МИН_КФ_ЛОГИСТИКИ].copy()

    if df.empty:
        return подготовить_пустую_схему(КОЛОНКИ_СКЛАДЫ)

    def пересчёт_строки(row):
        actual = abs(float(row["delivery_rub"]))
        coeff = float(row["dlv_prc"])
        if coeff <= 0:
            return actual, actual, 0.0
        if coeff > ПОРОГ_ДОРОГОГО_СКЛАДА:
            base = actual / coeff
            recalc = base * ЦЕЛЕВОЙ_КФ_ЛОГИСТИКИ
            overpay = max(0.0, actual - recalc)
            return actual, recalc, overpay
        return actual, actual, 0.0

    tmp = df.apply(
        lambda r: pd.Series(пересчёт_строки(r), index=["Логистика факт", "Логистика при коэффициенте 1,4", "Переплата"]),
        axis=1
    )
    df = pd.concat([df, tmp], axis=1)

    out = (
        df.groupby("office_name", as_index=False)
        .agg(
            **{
                "Средний коэффициент": ("dlv_prc", "mean"),
                "Количество продаж, шт": ("quantity", "sum"),
                "Логистика факт": ("Логистика факт", "sum"),
                "Логистика при коэффициенте 1,4": ("Логистика при коэффициенте 1,4", "sum"),
                "Переплата": ("Переплата", "sum"),
            }
        )
        .rename(columns={"office_name": "Склад"})
    )

    out["Переплата на единицу"] = out.apply(
        lambda r: безопасное_деление(r["Переплата"], r["Количество продаж, шт"], 6),
        axis=1
    )
    out["Неделя"] = неделя_код(начало_недели)

    out = out[КОЛОНКИ_СКЛАДЫ].sort_values(["Неделя", "Переплата"], ascending=[False, False]).reset_index(drop=True)
    return out


# =========================================================
# FACT / UNIT (ОСНОВНОЙ РАСЧЁТ)
# =========================================================

def посчитать_факт_и_юнит(
    fin_df: pd.DataFrame,
    advert_df: pd.DataFrame,
    cost_df: pd.DataFrame,
    stocks_df: pd.DataFrame,
    начало_недели: datetime.date,
    конец_недели: datetime.date,
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    if fin_df.empty:
        return подготовить_пустую_схему(КОЛОНКИ_ФАКТ), подготовить_пустую_схему(КОЛОНКИ_ЮНИТ), подготовить_пустую_схему(КОЛОНКИ_СКЛАДЫ)

    лог("📦 Шаг: подготовка финансовых строк")
    df = подготовить_финансовые_строки(fin_df)

    product_rows = df[df["nm_id"].notna()].copy()
    product_rows["nm_id"] = product_rows["nm_id"].astype("int64")

    meta_agg = (
        product_rows.groupby("nm_id", as_index=False)
        .agg(
            **{
                "Артикул продавца": ("sa_name", lambda x: мода_или_последнее(x)),
                "Предмет": ("subject_name", lambda x: мода_или_последнее(x)),
                "Бренд": ("brand_name", lambda x: мода_или_последнее(x)),
            }
        )
        .rename(columns={"nm_id": "Артикул WB"})
    )

    signed_rows = product_rows[product_rows["sign"] != 0].copy()

    продажи = (
        signed_rows[signed_rows["sign"] == 1]
        .groupby("nm_id")["quantity"]
        .sum()
        .rename("Продажи, шт")
    )
    возвраты = (
        signed_rows[signed_rows["sign"] == -1]
        .groupby("nm_id")["quantity"]
        .sum()
        .rename("Возвраты, шт")
    )
    чистые_продажи = (
        signed_rows.groupby("nm_id")["signed_quantity"]
        .sum()
        .rename("Чистые продажи, шт")
    )

    валовая_выручка = (
        signed_rows.groupby("nm_id")["signed_revenue"]
        .sum()
        .rename("Валовая выручка")
    )
    цена_покупателя_сумма = (
        signed_rows.groupby("nm_id")["signed_buyer_price"]
        .sum()
        .rename("Сумма цены покупателя")
    )

    комиссия = (
        signed_rows.groupby("nm_id")["signed_commission"]
        .sum()
        .rename("Комиссия WB")
    )
    эквайринг = (
        signed_rows.groupby("nm_id")["signed_acquiring"]
        .sum()
        .rename("Эквайринг")
    )

    средняя_цена_продажи = (валовая_выручка / чистые_продажи.replace(0, pd.NA)).rename("Средняя цена продажи")
    средняя_цена_покупателя = (цена_покупателя_сумма / чистые_продажи.replace(0, pd.NA)).rename("Средняя цена покупателя")

    sale_like = product_rows[
        product_rows.apply(
            lambda r: знак_строки(r.get("doc_type_name_norm", ""), r.get("supplier_oper_name_norm", "")) == 1,
            axis=1
        )
    ].copy()

    sale_like = sale_like.sort_values(["nm_id", "sale_dt", "rr_dt"], ascending=[True, False, False])

    последняя_комиссия_процент = (
        sale_like.groupby("nm_id")["commission_percent"]
        .first()
        .rename("Комиссия WB, %")
    )

    средний_спп = (
        sale_like.groupby("nm_id")["ppvz_spp_prc"]
        .mean()
        .rename("СПП, %")
    )

    средний_эквайринг_процент = (
        sale_like.groupby("nm_id")["acquiring_percent"]
        .mean()
        .rename("Эквайринг, %")
    )

    logistic_rows = product_rows[product_rows["supplier_oper_name_norm"].str.lower() == "логистика"].copy()
    if not logistic_rows.empty:
        logistic_rows["тип_логистики"] = logistic_rows.apply(тип_логистики, axis=1)
    else:
        logistic_rows["тип_логистики"] = []

    логистика_прямая = (
        logistic_rows[logistic_rows["тип_логистики"] == "прямая"]
        .groupby("nm_id")["delivery_rub"]
        .sum()
        .abs()
        .rename("Логистика прямая")
    )

    логистика_обратная = (
        logistic_rows[logistic_rows["тип_логистики"] == "обратная"]
        .groupby("nm_id")["delivery_rub"]
        .sum()
        .abs()
        .rename("Логистика обратная")
    )

    обратные_события = (
        logistic_rows[logistic_rows["тип_логистики"] == "обратная"]
        .groupby("nm_id")["return_amount"]
        .sum()
        .rename("_обратные_события")
    )

    # Нам нужна только acceptance для приёмки
    total_acceptance_store = abs(pd.to_numeric(df.get("acceptance", 0), errors="coerce").fillna(0).sum())

    if not advert_df.empty and "Артикул WB" in advert_df.columns:
        advert_agg = (
            advert_df.groupby("Артикул WB", as_index=False)
            .agg(
                **{
                    "Реклама": ("Расход", "sum"),
                    "Рекламные заказы, шт": ("Заказы", "sum"),
                    "Рекламная выручка": ("Сумма заказов", "sum"),
                    "Рекламные показы": ("Показы", "sum"),
                    "Рекламные клики": ("Клики", "sum"),
                }
            )
        )
        advert_agg["Артикул WB"] = pd.to_numeric(advert_agg["Артикул WB"], errors="coerce").fillna(0).astype("int64")
    else:
        advert_agg = pd.DataFrame(columns=[
            "Артикул WB", "Реклама", "Рекламные заказы, шт",
            "Рекламная выручка", "Рекламные показы", "Рекламные клики"
        ])

    if cost_df.empty:
        cost_df = pd.DataFrame(columns=["Артикул WB", "Себестоимость, руб"])

    base = meta_agg.copy()

    series_list = [
        продажи, возвраты, чистые_продажи,
        валовая_выручка, цена_покупателя_сумма,
        комиссия, эквайринг,
        средняя_цена_продажи, средняя_цена_покупателя,
        последняя_комиссия_процент, средний_спп, средний_эквайринг_процент,
        логистика_прямая, логистика_обратная, обратные_события,
    ]

    for s in series_list:
        base = base.merge(s.reset_index().rename(columns={"nm_id": "Артикул WB"}), on="Артикул WB", how="left")

    base = base.merge(cost_df, on="Артикул WB", how="left")
    base = base.merge(advert_agg, on="Артикул WB", how="left")

    if "Себестоимость, руб" not in base.columns:
        base["Себестоимость, руб"] = 0.0

    num_fill_cols = [
        "Продажи, шт", "Возвраты, шт", "Чистые продажи, шт",
        "Валовая выручка", "Сумма цены покупателя",
        "Комиссия WB", "Эквайринг",
        "Средняя цена продажи", "Средняя цена покупателя",
        "Комиссия WB, %", "СПП, %", "Эквайринг, %",
        "Логистика прямая", "Логистика обратная", "_обратные_события",
        "Себестоимость, руб", "Реклама", "Рекламные заказы, шт",
        "Рекламная выручка", "Рекламные показы", "Рекламные клики",
    ]
    for c in num_fill_cols:
        if c in base.columns:
            base[c] = pd.to_numeric(base[c], errors="coerce").fillna(0)

    base["Оценка заказанных, шт"] = base["Продажи, шт"] + base["_обратные_события"]
    base["Процент выкупа"] = base.apply(
        lambda r: округлить(безопасное_деление(r["Продажи, шт"] * 100.0, r["Оценка заказанных, шт"], 6), 2)
        if r["Оценка заказанных, шт"] > 0 else 0.0,
        axis=1
    )

    base["Себестоимость всего"] = base["Чистые продажи, шт"] * base["Себестоимость, руб"]

    # --- НОВЫЙ РАСЧЁТ ХРАНЕНИЯ И ПРОЧИХ РАСХОДОВ ---
    # Убедимся, что средняя цена продажи есть и заполнена
    if "Средняя цена продажи" not in base.columns:
        base["Средняя цена продажи"] = 0.0
    base["Средняя цена продажи"] = pd.to_numeric(base["Средняя цена продажи"], errors="coerce").fillna(0)

    # Хранение = 0.5% от цены за единицу * количество чистых продаж
    base["Хранение"] = base["Средняя цена продажи"] * 0.005 * base["Чистые продажи, шт"]
    # Прочие расходы = 1% от цены за единицу * количество чистых продаж
    base["Прочие расходы"] = base["Средняя цена продажи"] * 0.01 * base["Чистые продажи, шт"]

    # Штрафы и удержания игнорируем
    base["Штрафы"] = 0.0
    base["Удержания"] = 0.0

    # Расчёт приёмки (по acceptance)
    total_units_sold_store = base["Продажи, шт"].sum()
    приемка_на_ед = безопасное_деление(total_acceptance_store, total_units_sold_store, 6)
    base["Приёмка"] = base["Продажи, шт"] * приемка_на_ед

    base["НДС"] = base["Валовая выручка"].apply(lambda x: ндс_из_цены_с_ндс(x, СТАВКА_НДС))

    base["Валовая прибыль"] = (
        base["Валовая выручка"]
        - base["Себестоимость всего"]
        - base["Комиссия WB"]
        - base["Эквайринг"]
        - base["Логистика прямая"]
        - base["Логистика обратная"]
        - base["Хранение"]                # новое значение
        - base["Приёмка"]
        - base["Штрафы"]                   # 0
        - base["Удержания"]                 # 0
        - base["Реклама"]
        - base["Прочие расходы"]            # новая статья
    )

    base["Прибыль до налога"] = base["Валовая прибыль"] - base["НДС"]
    base["Налог на прибыль"] = base["Прибыль до налога"].apply(
        lambda x: max(0.0, x) * СТАВКА_НАЛОГА_НА_ПРИБЫЛЬ / 100.0
    )
    base["Чистая прибыль"] = base["Прибыль до налога"] - base["Налог на прибыль"]

    base["Валовая рентабельность, %"] = base.apply(
        lambda r: округлить(безопасное_деление(r["Валовая прибыль"] * 100.0, r["Валовая выручка"], 6), 2)
        if r["Валовая выручка"] else 0.0,
        axis=1
    )
    base["Чистая рентабельность, %"] = base.apply(
        lambda r: округлить(безопасное_деление(r["Чистая прибыль"] * 100.0, r["Валовая выручка"], 6), 2)
        if r["Валовая выручка"] else 0.0,
        axis=1
    )

    unit = base.copy()
    unit["Неделя"] = неделя_код(начало_недели)
    unit["Комиссия WB, руб/ед"] = unit.apply(lambda r: безопасное_деление(r["Комиссия WB"], r["Чистые продажи, шт"], 6), axis=1)
    unit["Эквайринг, руб/ед"] = unit.apply(lambda r: безопасное_деление(r["Эквайринг"], r["Чистые продажи, шт"], 6), axis=1)
    unit["Логистика прямая, руб/ед"] = unit.apply(lambda r: безопасное_деление(r["Логистика прямая"], r["Чистые продажи, шт"], 6), axis=1)
    unit["Логистика обратная, руб/ед"] = unit.apply(lambda r: безопасное_деление(r["Логистика обратная"], r["Чистые продажи, шт"], 6), axis=1)
    unit["Хранение, руб/ед"] = unit.apply(lambda r: безопасное_деление(r["Хранение"], r["Чистые продажи, шт"], 6), axis=1)
    unit["Приёмка, руб/ед"] = unit.apply(lambda r: безопасное_деление(r["Приёмка"], r["Чистые продажи, шт"], 6), axis=1)
    unit["Штрафы и удержания, руб/ед"] = 0.0   # обнуляем
    unit["Реклама, руб/ед"] = unit.apply(lambda r: безопасное_деление(r["Реклама"], r["Чистые продажи, шт"], 6), axis=1)
    unit["Прочие расходы, руб/ед"] = unit.apply(lambda r: безопасное_деление(r["Прочие расходы"], r["Чистые продажи, шт"], 6), axis=1)  # новая
    unit["НДС, руб/ед"] = unit.apply(lambda r: безопасное_деление(r["НДС"], r["Чистые продажи, шт"], 6), axis=1)
    unit["Валовая прибыль, руб/ед"] = unit.apply(lambda r: безопасное_деление(r["Валовая прибыль"], r["Чистые продажи, шт"], 6), axis=1)
    unit["Чистая прибыль, руб/ед"] = unit.apply(lambda r: безопасное_деление(r["Чистая прибыль"], r["Чистые продажи, шт"], 6), axis=1)
    unit = unit[КОЛОНКИ_ЮНИТ].copy()

    факт = base.copy()
    факт["Неделя"] = неделя_код(начало_недели)
    факт["Ставка НДС, %"] = СТАВКА_НДС
    факт["Ставка налога на прибыль, %"] = СТАВКА_НАЛОГА_НА_ПРИБЫЛЬ
    факт = факт[КОЛОНКИ_ФАКТ].copy()

    склады = анализ_складов(logistic_rows, начало_недели)

    return факт, unit, склады


# =========================================================
# ПРИЕМКА ЗА 9 НЕДЕЛЬ
# =========================================================

def норматив_приемки_за_9_недель(s3: S3Storage, последняя_неделя: datetime.date) -> float:
    лог("📦 Шаг: расчёт норматива приёмки за 9 недель")
    weeks = список_недель(последняя_неделя, ГЛУБИНА_ПРИЕМКИ_НЕДЕЛЬ)
    total_acceptance = 0.0
    total_units_sold = 0.0

    for i, ws in enumerate(weeks, start=1):
        лог(f"   → приёмка: неделя {i}/{len(weeks)} ({неделя_код(ws)})")
        df = прочитать_финансы_недели(s3, ws)
        if df.empty:
            continue

        df = подготовить_финансовые_строки(df)
        signed_rows = df[df["sign"] != 0].copy()
        sales = signed_rows[signed_rows["sign"] == 1]
        units = pd.to_numeric(sales["quantity"], errors="coerce").fillna(0).sum()

        total_units_sold += units
        total_acceptance += abs(pd.to_numeric(df.get("acceptance", 0), errors="coerce").fillna(0).sum())

    return безопасное_деление(total_acceptance, total_units_sold, 6)


def применить_норматив_приемки(unit_df: pd.DataFrame, приемка_на_ед_9н: float) -> pd.DataFrame:
    if unit_df.empty:
        return unit_df

    df = unit_df.copy()
    old_acceptance = df["Приёмка, руб/ед"].copy()
    df["Приёмка, руб/ед"] = приемка_на_ед_9н

    delta = df["Приёмка, руб/ед"] - old_acceptance
    df["Валовая прибыль, руб/ед"] = df["Валовая прибыль, руб/ед"] - delta
    df["Чистая прибыль, руб/ед"] = df["Чистая прибыль, руб/ед"] - delta

    return df


# =========================================================
# АНАЛИЗ НЕДЕЛЯ К НЕДЕЛЕ
# =========================================================

def объяснение_магазин(cur_store: Dict[str, float], prev_store: Dict[str, float]) -> str:
    deltas = {
        "выручка": cur_store["Валовая выручка"] - prev_store["Валовая выручка"],
        "реклама": cur_store["Реклама"] - prev_store["Реклама"],
        "комиссия": cur_store["Комиссия WB"] - prev_store["Комиссия WB"],
        "логистика": (cur_store["Логистика прямая"] + cur_store["Логистика обратная"]) -
                     (prev_store["Логистика прямая"] + prev_store["Логистика обратная"]),
        "себестоимость": cur_store["Себестоимость всего"] - prev_store["Себестоимость всего"],
        "хранение": cur_store["Хранение"] - prev_store["Хранение"],
        "приёмка": cur_store["Приёмка"] - prev_store["Приёмка"],
    }

    top_change = max(deltas, key=lambda k: abs(deltas[k]))

    if deltas["выручка"] > 0 and cur_store["Чистая прибыль"] > prev_store["Чистая прибыль"]:
        return f"Прибыль выросла. Основной драйвер — рост выручки. Сильнее всего изменилась статья: {top_change}."
    if deltas["выручка"] < 0 and cur_store["Чистая прибыль"] < prev_store["Чистая прибыль"]:
        return f"Прибыль снизилась вместе с выручкой. Сильнее всего изменилась статья: {top_change}."
    if cur_store["Реклама"] > prev_store["Реклама"] and cur_store["Валовая выручка"] <= prev_store["Валовая выручка"]:
        return "Прибыль снизилась: рекламные расходы выросли быстрее выручки."
    if (cur_store["Логистика прямая"] + cur_store["Логистика обратная"]) > (prev_store["Логистика прямая"] + prev_store["Логистика обратная"]):
        return "Прибыль снизилась: выросли логистические расходы."
    return f"Динамика смешанная. Наиболее сильное влияние оказала статья: {top_change}."


def объяснение_sku(row) -> str:
    delta_profit = row["Изменение чистой прибыли"]
    delta_revenue = row["Изменение выручки"]
    delta_ads = row["Изменение рекламы"]
    delta_comm = row["Изменение комиссии"]
    delta_log = row["Изменение логистики"]
    delta_spp = row["Изменение СПП"]
    delta_price = row["Изменение вашей цены"]

    reasons = []

    if delta_revenue > 0 and delta_profit > 0:
        reasons.append("рост выручки")
    if delta_revenue < 0 and delta_profit < 0:
        reasons.append("снижение выручки")
    if delta_ads > 0 and delta_profit < 0:
        reasons.append("рост рекламных расходов")
    if delta_comm > 0 and delta_profit < 0:
        reasons.append("рост комиссии WB")
    if delta_log > 0 and delta_profit < 0:
        reasons.append("рост логистики")
    if delta_price > 0:
        reasons.append("рост вашей цены")
    if delta_price < 0:
        reasons.append("снижение вашей цены")
    if delta_spp > 0:
        reasons.append("рост СПП WB")
    if delta_spp < 0:
        reasons.append("снижение СПП WB")

    if not reasons:
        return "Динамика без одного явного драйвера."

    return "; ".join(reasons[:3])


def построить_анализ_неделя_к_неделе(current_fact: pd.DataFrame, prev_fact: pd.DataFrame) -> pd.DataFrame:
    rows = []

    if current_fact.empty:
        return подготовить_пустую_схему(КОЛОНКИ_АНАЛИЗ)

    cur = current_fact.copy()
    prev = prev_fact.copy() if prev_fact is not None else pd.DataFrame()
    current_week = cur["Неделя"].iloc[0]

    summary_cols = [
        "Валовая выручка", "Комиссия WB", "Эквайринг", "Логистика прямая",
        "Логистика обратная", "Хранение", "Приёмка", "Штрафы", "Удержания",
        "Реклама", "Себестоимость всего", "Валовая прибыль", "НДС", "Налог на прибыль", "Чистая прибыль"
    ]

    cur_store = {c: cur[c].sum() for c in summary_cols}
    if not prev.empty:
        prev_store = {c: prev[c].sum() for c in summary_cols}
    else:
        prev_store = {c: 0.0 for c in summary_cols}

    rows.append({
        "Раздел": "Итог магазина",
        "Неделя": current_week,
        "Артикул WB": "",
        "Артикул продавца": "",
        "Предмет": "",
        "Показатель": "Итог по магазину",
        "Чистая прибыль текущая": округлить(cur_store["Чистая прибыль"], 2),
        "Чистая прибыль предыдущая": округлить(prev_store["Чистая прибыль"], 2),
        "Изменение чистой прибыли": округлить(cur_store["Чистая прибыль"] - prev_store["Чистая прибыль"], 2),
        "Изменение выручки": округлить(cur_store["Валовая выручка"] - prev_store["Валовая выручка"], 2),
        "Изменение рекламы": округлить(cur_store["Реклама"] - prev_store["Реклама"], 2),
        "Изменение комиссии": округлить(cur_store["Комиссия WB"] - prev_store["Комиссия WB"], 2),
        "Изменение логистики": округлить(
            (cur_store["Логистика прямая"] + cur_store["Логистика обратная"]) -
            (prev_store["Логистика прямая"] + prev_store["Логистика обратная"]),
            2
        ),
        "Изменение СПП": "",
        "Изменение вашей цены": "",
        "Комментарий": объяснение_магазин(cur_store, prev_store),
    })

    if not prev.empty:
        merge_cols = [
            "Артикул WB", "Артикул продавца", "Предмет", "Чистая прибыль", "Валовая выручка",
            "Реклама", "Комиссия WB", "Логистика прямая", "Логистика обратная",
            "СПП, %", "Средняя цена продажи"
        ]

        left = cur[merge_cols].copy()
        right = prev[merge_cols].copy()
        merged = left.merge(right, on="Артикул WB", how="outer", suffixes=("_тек", "_пред")).fillna(0)

        merged["Изменение чистой прибыли"] = merged["Чистая прибыль_тек"] - merged["Чистая прибыль_пред"]
        merged["Изменение выручки"] = merged["Валовая выручка_тек"] - merged["Валовая выручка_пред"]
        merged["Изменение рекламы"] = merged["Реклама_тек"] - merged["Реклама_пред"]
        merged["Изменение комиссии"] = merged["Комиссия WB_тек"] - merged["Комиссия WB_пред"]
        merged["Изменение логистики"] = (
            (merged["Логистика прямая_тек"] + merged["Логистика обратная_тек"]) -
            (merged["Логистика прямая_пред"] + merged["Логистика обратная_пред"])
        )
        merged["Изменение СПП"] = merged["СПП, %_тек"] - merged["СПП, %_пред"]
        merged["Изменение вашей цены"] = merged["Средняя цена продажи_тек"] - merged["Средняя цена продажи_пред"]

        merged["Артикул продавца"] = merged.get("Артикул продавца_тек", "").replace(0, "")
        merged["Предмет"] = merged.get("Предмет_тек", "").replace(0, "")
        merged["Комментарий"] = merged.apply(объяснение_sku, axis=1)

        лидеры_роста = merged.sort_values("Изменение чистой прибыли", ascending=False).head(20)
        лидеры_падения = merged.sort_values("Изменение чистой прибыли", ascending=True).head(20)

        for _, r in лидеры_роста.iterrows():
            rows.append({
                "Раздел": "Лидеры роста",
                "Неделя": current_week,
                "Артикул WB": int(r["Артикул WB"]) if pd.notna(r["Артикул WB"]) and r["Артикул WB"] != "" else "",
                "Артикул продавца": r.get("Артикул продавца", ""),
                "Предмет": r.get("Предмет", ""),
                "Показатель": "Рост прибыли",
                "Чистая прибыль текущая": округлить(r["Чистая прибыль_тек"], 2),
                "Чистая прибыль предыдущая": округлить(r["Чистая прибыль_пред"], 2),
                "Изменение чистой прибыли": округлить(r["Изменение чистой прибыли"], 2),
                "Изменение выручки": округлить(r["Изменение выручки"], 2),
                "Изменение рекламы": округлить(r["Изменение рекламы"], 2),
                "Изменение комиссии": округлить(r["Изменение комиссии"], 2),
                "Изменение логистики": округлить(r["Изменение логистики"], 2),
                "Изменение СПП": округлить(r["Изменение СПП"], 2),
                "Изменение вашей цены": округлить(r["Изменение вашей цены"], 2),
                "Комментарий": r["Комментарий"],
            })

        for _, r in лидеры_падения.iterrows():
            rows.append({
                "Раздел": "Лидеры падения",
                "Неделя": current_week,
                "Артикул WB": int(r["Артикул WB"]) if pd.notna(r["Артикул WB"]) and r["Артикул WB"] != "" else "",
                "Артикул продавца": r.get("Артикул продавца", ""),
                "Предмет": r.get("Предмет", ""),
                "Показатель": "Падение прибыли",
                "Чистая прибыль текущая": округлить(r["Чистая прибыль_тек"], 2),
                "Чистая прибыль предыдущая": округлить(r["Чистая прибыль_пред"], 2),
                "Изменение чистой прибыли": округлить(r["Изменение чистой прибыли"], 2),
                "Изменение выручки": округлить(r["Изменение выручки"], 2),
                "Изменение рекламы": округлить(r["Изменение рекламы"], 2),
                "Изменение комиссии": округлить(r["Изменение комиссии"], 2),
                "Изменение логистики": округлить(r["Изменение логистики"], 2),
                "Изменение СПП": округлить(r["Изменение СПП"], 2),
                "Изменение вашей цены": округлить(r["Изменение вашей цены"], 2),
                "Комментарий": r["Комментарий"],
            })

    out = pd.DataFrame(rows)
    for col in КОЛОНКИ_АНАЛИЗ:
        if col not in out.columns:
            out[col] = pd.NA
    out = out[КОЛОНКИ_АНАЛИЗ].copy()
    return out


# =========================================================
# ИСТОРИЯ / ХРАНЕНИЕ 13 НЕДЕЛЬ
# =========================================================

def объединить_с_удержанием(
    existing_df: pd.DataFrame,
    new_df: pd.DataFrame,
    key_cols: List[str],
    retention_weeks: int = ГЛУБИНА_ХРАНЕНИЯ_НЕДЕЛЬ,
) -> pd.DataFrame:
    if new_df is None or new_df.empty:
        return existing_df if existing_df is not None else pd.DataFrame()

    if existing_df is None or existing_df.empty:
        combined = new_df.copy()
    else:
        combined = pd.concat([existing_df, new_df], ignore_index=True)

    if "Неделя" in combined.columns:
        combined = combined.drop_duplicates(subset=key_cols, keep="last")
        weeks_sorted = sorted([w for w in combined["Неделя"].dropna().astype(str).unique()])
        if len(weeks_sorted) > retention_weeks:
            keep_weeks = set(weeks_sorted[-retention_weeks:])
            combined = combined[combined["Неделя"].astype(str).isin(keep_weeks)].copy()

    return combined.reset_index(drop=True)


def загрузить_существующую_экономику(s3: S3Storage) -> Dict[str, pd.DataFrame]:
    if not s3.file_exists(ПУТЬ_ЭКОНОМИКА):
        return {
            ЛИСТ_ЮНИТ: подготовить_пустую_схему(КОЛОНКИ_ЮНИТ),
            ЛИСТ_ФАКТ: подготовить_пустую_схему(КОЛОНКИ_ФАКТ),
            ЛИСТ_АНАЛИЗ: подготовить_пустую_схему(КОЛОНКИ_АНАЛИЗ),
            ЛИСТ_СКЛАДЫ: подготовить_пустую_схему(КОЛОНКИ_СКЛАДЫ),
        }

    try:
        sheets = s3.read_excel_all_sheets(ПУТЬ_ЭКОНОМИКА)
        return {
            ЛИСТ_ЮНИТ: очистить_схему(sheets.get(ЛИСТ_ЮНИТ, pd.DataFrame()), КОЛОНКИ_ЮНИТ),
            ЛИСТ_ФАКТ: очистить_схему(sheets.get(ЛИСТ_ФАКТ, pd.DataFrame()), КОЛОНКИ_ФАКТ),
            ЛИСТ_АНАЛИЗ: очистить_схему(sheets.get(ЛИСТ_АНАЛИЗ, pd.DataFrame()), КОЛОНКИ_АНАЛИЗ),
            ЛИСТ_СКЛАДЫ: очистить_схему(sheets.get(ЛИСТ_СКЛАДЫ, pd.DataFrame()), КОЛОНКИ_СКЛАДЫ),
        }
    except Exception as e:
        лог(f"⚠️ Не удалось прочитать существующий Экономика.xlsx: {e}")
        return {
            ЛИСТ_ЮНИТ: подготовить_пустую_схему(КОЛОНКИ_ЮНИТ),
            ЛИСТ_ФАКТ: подготовить_пустую_схему(КОЛОНКИ_ФАКТ),
            ЛИСТ_АНАЛИЗ: подготовить_пустую_схему(КОЛОНКИ_АНАЛИЗ),
            ЛИСТ_СКЛАДЫ: подготовить_пустую_схему(КОЛОНКИ_СКЛАДЫ),
        }


# =========================================================
# ОКРУГЛЕНИЕ
# =========================================================

def округлить_таблицу(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    out = df.copy()
    for col in out.columns:
        if pd.api.types.is_numeric_dtype(out[col]):
            out[col] = out[col].apply(lambda x: округлить(x, 2))
    return out


# =========================================================
# ОСНОВНОЙ КЛАСС
# =========================================================

class КалькуляторЭкономики:
    def __init__(self, s3: S3Storage):
        self.s3 = s3

    def run(self):
        начало_недели, конец_недели = последняя_полная_неделя()
        начало_пред = начало_недели - timedelta(days=7)
        конец_пред = конец_недели - timedelta(days=7)

        лог("=" * 80)
        лог(f"📌 Запуск weekly-экономики для магазина {МАГАЗИН}")
        лог(f"📅 Целевая неделя: {неделя_код(начало_недели)}")
        лог("=" * 80)

        # Разовый пересчёт всех последних недель
        if ПЕРЕСЧИТАТЬ_ИСТОРИЮ:
            лог("🔄 Режим пересчёта истории: будут пересчитаны все недели за последние 13 недель")
            недели_для_пересчёта = список_недель(начало_недели, ГЛУБИНА_ХРАНЕНИЯ_НЕДЕЛЬ)
            все_факт = []
            все_юнит = []
            все_склады = []
            for nd in недели_для_пересчёта:
                лог(f"⏳ Пересчёт недели {неделя_код(nd)}")
                фин = прочитать_финансы_недели(self.s3, nd)
                if фин.empty:
                    лог(f"⚠️ Нет финансов за неделю {неделя_код(nd)}, пропускаем")
                    continue
                ост = прочитать_остатки_недели(self.s3, nd)
                рек = прочитать_рекламу_недели(self.s3, nd, nd + timedelta(days=6))
                себест = прочитать_себестоимость(self.s3)
                ф, у, ск = посчитать_факт_и_юнит(
                    fin_df=фин,
                    advert_df=рек,
                    cost_df=себест,
                    stocks_df=ост,
                    начало_недели=nd,
                    конец_недели=nd + timedelta(days=6),
                )
                if not ф.empty:
                    все_факт.append(ф)
                    все_юнит.append(у)
                    все_склады.append(ск)

            if все_факт:
                факт = pd.concat(все_факт, ignore_index=True)
                юнит = pd.concat(все_юнит, ignore_index=True)
                склады = pd.concat(все_склады, ignore_index=True)
                # Применим норматив приёмки ко всем юнитам
                приемка_норм_9н = норматив_приемки_за_9_недель(self.s3, начало_недели)
                лог(f"📦 Норматив приёмки за {ГЛУБИНА_ПРИЕМКИ_НЕДЕЛЬ} недель: {приемка_норм_9н:.6f} руб/ед")
                юнит = применить_норматив_приемки(юнит, приемка_норм_9н)
                # Анализ строить не будем, так как для истории он не нужен (или можно построить для последней недели)
                анализ = подготовить_пустую_схему(КОЛОНКИ_АНАЛИЗ)
                # Сохраняем, игнорируя существующий файл
                sheets_to_write = {
                    ЛИСТ_ЮНИТ: округлить_таблицу(юнит.sort_values(["Неделя", "Чистая прибыль, руб/ед"], ascending=[False, False]).reset_index(drop=True)),
                    ЛИСТ_ФАКТ: округлить_таблицу(факт.sort_values(["Неделя", "Чистая прибыль"], ascending=[False, False]).reset_index(drop=True)),
                    ЛИСТ_АНАЛИЗ: округлить_таблицу(анализ),
                    ЛИСТ_СКЛАДЫ: округлить_таблицу(склады.sort_values(["Неделя", "Переплата"], ascending=[False, False]).reset_index(drop=True)),
                }
                лог("💾 Шаг: запись пересчитанной истории в Экономика.xlsx")
                self.s3.write_excel_sheets(ПУТЬ_ЭКОНОМИКА, sheets_to_write)
                лог("✅ Пересчёт истории завершён.")
                return
            else:
                лог("❌ Не удалось пересчитать ни одной недели, завершаю работу.")
                return

        # Обычный режим (только текущая неделя)
        лог("📥 Шаг: чтение финансов")
        финансы = прочитать_финансы_недели(self.s3, начало_недели)
        if финансы.empty:
            raise RuntimeError("Нет финансовых данных за целевую неделю")
        лог(f"✅ Финансы загружены, строк: {len(финансы)}")

        лог("📥 Шаг: чтение остатков")
        остатки = прочитать_остатки_недели(self.s3, начало_недели)
        лог(f"ℹ️ Остатки: строк {len(остатки)}")

        реклама = прочитать_рекламу_недели(self.s3, начало_недели, конец_недели)
        себестоимость = прочитать_себестоимость(self.s3)

        факт, юнит, склады = посчитать_факт_и_юнит(
            fin_df=финансы,
            advert_df=реклама,
            cost_df=себестоимость,
            stocks_df=остатки,
            начало_недели=начало_недели,
            конец_недели=конец_недели,
        )

        if факт.empty:
            raise RuntimeError("Не удалось сформировать weekly fact")

        приемка_норм_9н = норматив_приемки_за_9_недель(self.s3, начало_недели)
        лог(f"📦 Норматив приёмки за {ГЛУБИНА_ПРИЕМКИ_НЕДЕЛЬ} недель: {приемка_норм_9н:.6f} руб/ед")

        юнит = применить_норматив_приемки(юнит, приемка_норм_9н)

        лог("📥 Шаг: чтение предыдущей недели")
        финансы_пред = прочитать_финансы_недели(self.s3, начало_пред)
        if not финансы_пред.empty:
            остатки_пред = прочитать_остатки_недели(self.s3, начало_пред)
            реклама_пред = прочитать_рекламу_недели(self.s3, начало_пред, конец_пред)
            факт_пред, _, _ = посчитать_факт_и_юнит(
                fin_df=финансы_пред,
                advert_df=реклама_пред,
                cost_df=себестоимость,
                stocks_df=остатки_пред,
                начало_недели=начало_пред,
                конец_недели=конец_пред,
            )
        else:
            факт_пред = подготовить_пустую_схему(КОЛОНКИ_ФАКТ)

        лог("📊 Шаг: построение анализа")
        анализ = построить_анализ_неделя_к_неделе(факт, факт_пред)

        лог("📂 Шаг: загрузка существующего файла Экономика.xlsx")
        existing = загрузить_существующую_экономику(self.s3)

        лог("🧩 Шаг: объединение истории")
        юнит_all = объединить_с_удержанием(
            existing[ЛИСТ_ЮНИТ], юнит,
            key_cols=["Неделя", "Артикул WB"],
            retention_weeks=ГЛУБИНА_ХРАНЕНИЯ_НЕДЕЛЬ,
        )

        факт_all = объединить_с_удержанием(
            existing[ЛИСТ_ФАКТ], факт,
            key_cols=["Неделя", "Артикул WB"],
            retention_weeks=ГЛУБИНА_ХРАНЕНИЯ_НЕДЕЛЬ,
        )

        анализ_all = объединить_с_удержанием(
            existing[ЛИСТ_АНАЛИЗ], анализ,
            key_cols=["Неделя", "Раздел", "Артикул WB", "Показатель"],
            retention_weeks=ГЛУБИНА_ХРАНЕНИЯ_НЕДЕЛЬ,
        )

        склады_all = объединить_с_удержанием(
            existing[ЛИСТ_СКЛАДЫ], склады,
            key_cols=["Неделя", "Склад"],
            retention_weeks=ГЛУБИНА_ХРАНЕНИЯ_НЕДЕЛЬ,
        )

        юнит_all = очистить_схему(юнит_all, КОЛОНКИ_ЮНИТ)
        факт_all = очистить_схему(факт_all, КОЛОНКИ_ФАКТ)
        анализ_all = очистить_схему(анализ_all, КОЛОНКИ_АНАЛИЗ)
        склады_all = очистить_схему(склады_all, КОЛОНКИ_СКЛАДЫ)

        sheets_to_write = {
            ЛИСТ_ЮНИТ: округлить_таблицу(
                юнит_all.sort_values(["Неделя", "Чистая прибыль, руб/ед"], ascending=[False, False]).reset_index(drop=True)
            ),
            ЛИСТ_ФАКТ: округлить_таблицу(
                факт_all.sort_values(["Неделя", "Чистая прибыль"], ascending=[False, False]).reset_index(drop=True)
            ),
            ЛИСТ_АНАЛИЗ: округлить_таблицу(
                анализ_all.sort_values(["Неделя", "Раздел"], ascending=[False, True]).reset_index(drop=True)
            ),
            ЛИСТ_СКЛАДЫ: округлить_таблицу(
                склады_all.sort_values(["Неделя", "Переплата"], ascending=[False, False]).reset_index(drop=True)
            ),
        }

        лог("💾 Шаг: запись Экономика.xlsx")
        self.s3.write_excel_sheets(ПУТЬ_ЭКОНОМИКА, sheets_to_write)

        total_revenue = факт["Валовая выручка"].sum()
        total_gp = факт["Валовая прибыль"].sum()
        total_net = факт["Чистая прибыль"].sum()

        лог(f"✅ Экономика сохранена: {ПУТЬ_ЭКОНОМИКА}")
        лог(f"📊 Выручка недели: {total_revenue:,.2f}")
        лог(f"📊 Валовая прибыль недели: {total_gp:,.2f}")
        лог(f"📊 Чистая прибыль недели: {total_net:,.2f}")


# =========================================================
# ENTRYPOINT
# =========================================================

def main():
    required_env = ["YC_ACCESS_KEY_ID", "YC_SECRET_ACCESS_KEY", "YC_BUCKET_NAME"]
    missing = [var for var in required_env if not os.environ.get(var)]
    if missing:
        raise RuntimeError(f"Отсутствуют переменные окружения: {missing}")

    s3 = S3Storage(
        access_key=os.environ["YC_ACCESS_KEY_ID"],
        secret_key=os.environ["YC_SECRET_ACCESS_KEY"],
        bucket_name=os.environ["YC_BUCKET_NAME"],
    )

    calc = КалькуляторЭкономики(s3)
    calc.run()


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        лог(f"❌ Критическая ошибка: {e}")
        traceback.print_exc()
        raise
