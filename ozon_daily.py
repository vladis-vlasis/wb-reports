from __future__ import annotations

import json
import os
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from playwright.sync_api import BrowserContext, Page, sync_playwright

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
        raise RuntimeError("В GitHub не задан Secret OZON_AUTH_JSON")

    try:
        auth = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise RuntimeError("OZON_AUTH_JSON содержит некорректный JSON") from exc

    if not auth.get("company_id"):
        raise RuntimeError("В OZON_AUTH_JSON отсутствует company_id")

    if not isinstance(auth.get("cookies"), dict) or not auth["cookies"]:
        raise RuntimeError("В OZON_AUTH_JSON отсутствует объект cookies")

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


def choose_category(page: Page) -> None:
    category_control = page.get_by_text(re.compile(r"^Категория"))
    if category_control.count() == 0:
        category_control = page.get_by_role(
            "button",
            name=re.compile("Категория", re.IGNORECASE),
        )

    category_control.first.click(timeout=30_000)
    page.wait_for_timeout(700)

    option = page.get_by_text(CATEGORY_NAME, exact=True)
    if option.count() == 0:
        raise RuntimeError(f"Категория не найдена: {CATEGORY_NAME}")

    option.last.click(timeout=30_000)
    page.wait_for_timeout(4_000)


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
            f"Ozon API вернул {response.status}: {body[:500]}"
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
        page = context.new_page()

        try:
            page.goto(
                ANALYTICS_URL,
                wait_until="domcontentloaded",
                timeout=120_000,
            )
            page.wait_for_timeout(6_000)

            if "login" in page.url.lower():
                raise RuntimeError(
                    "Сессия Ozon истекла. Обновите OZON_AUTH_JSON."
                )

            choose_category(page)

            screenshot_path = (
                SCREENSHOT_DIR
                / f"ozon_sales_funnel_{stamp}.png"
            )
            page.screenshot(
                path=str(screenshot_path),
                full_page=True,
            )

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
                file.write(
                    json.dumps(result, ensure_ascii=False) + "\n"
                )

            print(f"Готово: {daily_path}")
            print(f"Скриншот: {screenshot_path}")

        except Exception as exc:
            error_path = SCREENSHOT_DIR / f"error_{stamp}.png"
            try:
                page.screenshot(path=str(error_path), full_page=True)
            except Exception:
                pass

            print(f"Ошибка: {exc}", file=sys.stderr)
            return 1

        finally:
            browser.close()

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
