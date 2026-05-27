# WB TOPFACE — автоматизация отчетов и рекламы

Готовая структура для нового репозитория GitHub.

## Что куда положено

| Назначение | Новый файл | Исходный файл |
|---|---|---|
| Ежедневное обновление данных | `scripts/ezhednevnoe_obnovlenie_dannyh.py` | `wb_updater (5).py` |
| Управление рекламой | `scripts/upravlenie_reklamoy.py` | `assistant_wb_ads_manager (38).py` |
| Отчет по оборачиваемости / остаткам | `scripts/otchet_po_oborachivaemosti.py` | `wb_stock_days_report (21).py` |
| Расчет юнит-экономики | `scripts/raschet_unit_ekonomiki.py` | `economics_weekly (1).py` |

## GitHub Actions

В папке `.github/workflows/` лежат 4 workflow:

1. `01_ezhednevnoe_obnovlenie_dannyh.yml` — ежедневное обновление данных.
2. `02_raschet_unit_ekonomiki.yml` — расчет `Экономика.xlsx`.
3. `03_otchet_po_oborachivaemosti.yml` — отчет по оборачиваемости.
4. `04_upravlenie_reklamoy.yml` — управление рекламой.

В интерфейсе GitHub Actions названия будут на русском:
- `Ежедневное обновление данных`
- `Расчет юнит-экономики`
- `Отчет по оборачиваемости`
- `Управление рекламой`

## Обязательные Secrets

Добавить в GitHub: `Settings` → `Secrets and variables` → `Actions` → `New repository secret`.

| Secret | Для чего нужен |
|---|---|
| `YC_ACCESS_KEY_ID` | Access Key Yandex Object Storage |
| `YC_SECRET_ACCESS_KEY` | Secret Key Yandex Object Storage |
| `YC_BUCKET_NAME` | Имя bucket в Object Storage |
| `WB_PROMO_KEY_TOPFACE` | Единый API-ключ Wildberries для TOPFACE |

## Дополнительные Secrets

Эти secrets нужны только если используешь соответствующие функции.

| Secret | Когда нужен |
|---|---|
| `URL_1C_STOCKS` | Если ежедневный updater должен забирать остатки 1С по URL |
| `_1C_USER` | Логин для 1С |
| `_1C_PASSWORD` | Пароль для 1С |
| `TELEGRAM_BOT_TOKEN` | Отправка отчета по оборачиваемости в Telegram |
| `TELEGRAM_CHAT_ID` | Чат Telegram для отправки отчета |
| `YC_ENDPOINT_URL` | Можно не задавать, если используется стандартный `https://storage.yandexcloud.net` |

## Дополнительные Variables

Добавлять в GitHub: `Settings` → `Secrets and variables` → `Actions` → вкладка `Variables`.

| Variable | По умолчанию в коде |
|---|---|
| `WB_FORCE_SEND` | `false` |
| `WB_SEND_REDISTRIBUTION_ALWAYS` | `false` |
| `WB_REDISTRIBUTION_TEMPLATE_KEY` | `Отчёты/Остатки/Перераспределение/Перераспределения.xlsx` |
| `WB_REDISTRIBUTION_LOOKBACK_DAYS` | `14` |
| `WB_REDISTRIBUTION_TARGET_DAYS` | `21` |
| `WB_STOP_LIST_KEY` | пусто |
| `WB_PAUSE_MIN_IMPRESSIONS` | `10000` |
| `WB_PAUSE_ANALYSIS_DAYS` | `21` |
| `WB_RAMP_TARGET_IMPRESSIONS_PER_DAY` | `1000` |
| `WB_RAMP_MAX_SPEND_PER_DAY` | `500` |
| `WB_RAMP_CHECK_DAYS` | `7` |
| `WB_ONE_CAMPAIGN_TARGET_POSITION` | `10` |
| `WB_PRICE_MIN_SELLER_DISCOUNT_PCT` | `25` |
| `WB_MAX_PRICE_TEST_ITEMS_PER_RUN` | `30` |
| `WB_ECONOMICS_SUBTRACT_COGS` | `1` |

## Как быстро загрузить в новый репозиторий

1. Создай пустой репозиторий на GitHub.
2. Распакуй архив в любую папку.
3. Выполни команды:

```bash
git init
git add .
git commit -m "Добавить WB автоматизацию"
git branch -M main
git remote add origin https://github.com/USER/REPO.git
git push -u origin main
```

4. Добавь secrets.
5. Открой вкладку `Actions` и запусти нужный workflow вручную через `Run workflow`.

## Важное по рекламе

Workflow `Управление рекламой` по расписанию запускает:

```bash
python scripts/upravlenie_reklamoy.py run
```

Ручной запуск по умолчанию стоит в режиме `preview`, чтобы сначала проверить предпросмотр без изменений.
