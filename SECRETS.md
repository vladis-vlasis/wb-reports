# Секреты для GitHub Actions

## Минимальный набор

```text
YC_ACCESS_KEY_ID
YC_SECRET_ACCESS_KEY
YC_BUCKET_NAME
WB_PROMO_KEY_TOPFACE
```

## Если нужна 1С

```text
URL_1C_STOCKS
_1C_USER
_1C_PASSWORD
```

## Если нужна отправка отчета по оборачиваемости в Telegram

```text
TELEGRAM_BOT_TOKEN
TELEGRAM_CHAT_ID
```

## Что можно оставить без заполнения

`YC_ENDPOINT_URL` можно не задавать, если Object Storage — стандартный Yandex Cloud.

Все `WB_*` настройки из Variables можно не добавлять: в коде уже есть значения по умолчанию.
