from __future__ import annotations

import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from playwright.sync_api import BrowserContext, sync_playwright

ANALYTICS_URL = "https://seller.ozon.ru/app/analytics"
CATEGORY_NAME = "Миски для животных"
CATEGORY_ID = "95189"
PERIOD = "period_7_days"

ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
SCREENSHOT_DIR = ROOT / "screenshots"

API_URLS = {
    "category_tree": (
        "https://seller.ozon.ru/api/site/"
        "seller-analytics/sales-funnel/v3/get_category_tree"
    ),
    "compare_dynamics": (
        "https://seller.ozon.ru/api/site/"
        "seller-analytics/sales-funnel/v2/sales_funnel/compare_dynamics"
    ),
    "chart_revenue": (
        "https://seller.ozon.ru/api/site/"
        "seller-analytics/sales-funnel/v2/get_chart_revenue"
    ),
}


def load_auth() -> dict[str, Any]:
    raw = os.getenv("OZON_AUTH_JSON", "").strip()
    if not raw:
        raise RuntimeError("В GitHub не задан Secret SECRET_OZON_AUTH_JSON")

    try:
        auth = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise RuntimeError("Секрет содержит некорректный JSON") from exc

    if not auth.get("company_id"):
        raise RuntimeError("В секрете отсутствует company_id")

    if not isinstance(auth.get("cookies"), dict) or not auth["cookies"]:
        raise RuntimeError("В секрете отсутствует объект cookies")

    return auth


def add_cookies(context: BrowserContext, cookies_map: dict[str, str]) -> None:
    cookies = []

    for name, value in cookies_map.items():
        if not value:
            continue

        cookies.append(
            {
                "name": str(name),
                "value": str(value),
                "domain": "seller.ozon.ru",
                "path": "/",
                "secure": True,
                "sameSite": "Lax",
            }
        )

    context.add_cookies(cookies)


def post_json(
    context: BrowserContext,
    url: str,
    payload: dict[str, Any],
    headers: dict[str, str],
) -> Any:
    response = context.request.post(
        url,
        data=payload,
        headers=headers,
        timeout=60_000,
    )

    if not response.ok:
        body = response.text()
        raise RuntimeError(
            f"Ozon API {url} вернул {response.status}: {body[:500]}"
        )

    return response.json()


def main() -> int:
    auth = load_auth()
    company_id = str(auth["company_id"])

    DATA_DIR.mkdir(exist_ok=True)
    SCREENSHOT_DIR.mkdir(exist_ok=True)

    now = datetime.now(timezone.utc)
    stamp = now.strftime("%Y-%m-%d_%H-%M-%S")

    headers = {
        "accept": "application/json, text/plain, */*",
        "accept-language": "ru",
        "content-type": "application/json",
        "origin": "https://seller.ozon.ru",
        "referer": ANALYTICS_URL,
        "x-o3-app-name": "seller-ui",
        "x-o3-company-id": company_id,
        "x-o3-language": "ru",
        "x-o3-page-type": "analytics",
    }

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=True)
        context = browser.new_context(
            viewport={"width": 1600, "height": 1200},
            locale="ru-RU",
        )
        add_cookies(context, auth["cookies"])

        try:
            # Категорию через интерфейс больше не выбираем.
            # Сразу передаём её ID в запросы Ozon.
            result = {
                "collected_at_utc": now.isoformat(),
                "category": CATEGORY_NAME,
                "category_id": CATEGORY_ID,
                "period": PERIOD,
                "category_tree": post_json(
                    context,
                    API_URLS["category_tree"],
                    {"period": PERIOD},
                    headers,
                ),
                "compare_dynamics": post_json(
                    context,
                    API_URLS["compare_dynamics"],
                    {
                        "description_type_id": CATEGORY_ID,
                        "period": PERIOD,
                        "price_range": 0,
                    },
                    headers,
                ),
                "chart_revenue": post_json(
                    context,
                    API_URLS["chart_revenue"],
                    {
                        "compare_type": "PREV_PERIOD",
                        "description_type_id": CATEGORY_ID,
                        "period": PERIOD,
                        "price_range": 0,
                    },
                    headers,
                ),
            }

            daily_path = DATA_DIR / f"ozon_sales_funnel_{stamp}.json"
            daily_path.write_text(
                json.dumps(result, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )

            history_path = DATA_DIR / "ozon_sales_funnel_history.jsonl"
            with history_path.open("a", encoding="utf-8") as file:
                file.write(json.dumps(result, ensure_ascii=False) + "\n")

            print(f"Готово: {daily_path}")

        except Exception as exc:
            print(f"Ошибка: {exc}", file=sys.stderr)
            return 1

        finally:
            browser.close()

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
