# VERSION: v7_20260610_UNIQUE_FILE_NO_REPORT_ENV_STOP
# Скачиваемый файл намеренно имеет уникальное имя.
# В GitHub нужно полностью заменить содержимое:
# .github/workflows/01_ezhednevnoe_obnovlenie_dannyh.yml

name: Ежедневное обновление данных MISSTAIS v7

on:
  workflow_dispatch:
    inputs:
      store:
        description: "Магазин"
        required: true
        default: "ALL"
        type: choice
        options:
          - ALL
          - TOPFACE
          - MISSTAIS
  schedule:
    # Каждый день в 10:00 МСК = 07:00 UTC
    - cron: "0 7 * * *"

jobs:
  run:
    name: Запустить ежедневное обновление данных MISSTAIS v7
    runs-on: ubuntu-latest

    steps:
      - name: Скачать репозиторий
        uses: actions/checkout@v4

      - name: Установить Python
        uses: actions/setup-python@v5
        with:
          python-version: "3.11"

      - name: Установить зависимости
        run: |
          python -m pip install --upgrade pip
          pip install -r requirements.txt

      - name: Загрузить secrets v7
        env:
          REPORT_ENV: ${{ secrets.REPORT_ENV }}
          YC_ACCESS_KEY_ID: ${{ secrets.YC_ACCESS_KEY_ID }}
          YC_SECRET_ACCESS_KEY: ${{ secrets.YC_SECRET_ACCESS_KEY }}
          YC_BUCKET_NAME: ${{ secrets.YC_BUCKET_NAME }}
          WB_PROMO_KEY_TOPFACE: ${{ secrets.WB_PROMO_KEY_TOPFACE }}
          WB_KEY_MISSTAIS: ${{ secrets.WB_KEY_MISSTAIS }}
          TELEGRAM_BOT_TOKEN: ${{ secrets.TELEGRAM_BOT_TOKEN }}
          TELEGRAM_CHAT_ID: ${{ secrets.TELEGRAM_CHAT_ID }}
          TORGSTAT_ABC_URL: ${{ secrets.TORGSTAT_ABC_URL }}
        run: |
          echo "__RUNNING_YML_VERSION_v7_20260610_UNIQUE_FILE_NO_REPORT_ENV_STOP__"

          if [ -n "$REPORT_ENV" ]; then
            printf '%s\n' "$REPORT_ENV" | sed 's/\r$//' >> "$GITHUB_ENV"
            echo "REPORT_ENV загружен"
          else
            echo "::warning::REPORT_ENV пустой или недоступен. Это НЕ останавливает workflow."
          fi

          if [ -n "$YC_ACCESS_KEY_ID" ]; then echo "YC_ACCESS_KEY_ID=$YC_ACCESS_KEY_ID" >> "$GITHUB_ENV"; fi
          if [ -n "$YC_SECRET_ACCESS_KEY" ]; then echo "YC_SECRET_ACCESS_KEY=$YC_SECRET_ACCESS_KEY" >> "$GITHUB_ENV"; fi
          if [ -n "$YC_BUCKET_NAME" ]; then echo "YC_BUCKET_NAME=$YC_BUCKET_NAME" >> "$GITHUB_ENV"; fi
          if [ -n "$WB_PROMO_KEY_TOPFACE" ]; then echo "WB_PROMO_KEY_TOPFACE=$WB_PROMO_KEY_TOPFACE" >> "$GITHUB_ENV"; fi
          if [ -n "$WB_KEY_MISSTAIS" ]; then echo "WB_KEY_MISSTAIS=$WB_KEY_MISSTAIS" >> "$GITHUB_ENV"; fi
          if [ -n "$TELEGRAM_BOT_TOKEN" ]; then echo "TELEGRAM_BOT_TOKEN=$TELEGRAM_BOT_TOKEN" >> "$GITHUB_ENV"; fi
          if [ -n "$TELEGRAM_CHAT_ID" ]; then echo "TELEGRAM_CHAT_ID=$TELEGRAM_CHAT_ID" >> "$GITHUB_ENV"; fi
          if [ -n "$TORGSTAT_ABC_URL" ]; then echo "TORGSTAT_ABC_URL=$TORGSTAT_ABC_URL" >> "$GITHUB_ENV"; fi

      - name: Проверить обязательные переменные v7
        env:
          STORE_INPUT: ${{ github.event.inputs.store || 'ALL' }}
        run: |
          echo "__CHECKING_YML_VERSION_v7_20260610_UNIQUE_FILE_NO_REPORT_ENV_STOP__"
          missing=0

          for var in YC_ACCESS_KEY_ID YC_SECRET_ACCESS_KEY YC_BUCKET_NAME; do
            if [ -z "${!var}" ]; then
              echo "::error::Не найдена переменная $var. Добавь её в REPORT_ENV или отдельным GitHub Secret."
              missing=1
            else
              echo "$var загружена"
            fi
          done

          if [ "$STORE_INPUT" = "ALL" ] || [ "$STORE_INPUT" = "TOPFACE" ]; then
            if [ -z "$WB_PROMO_KEY_TOPFACE" ]; then
              echo "::error::Не найден WB_PROMO_KEY_TOPFACE для TOPFACE"
              missing=1
            else
              echo "WB_PROMO_KEY_TOPFACE загружен"
            fi
          fi

          if [ "$STORE_INPUT" = "ALL" ] || [ "$STORE_INPUT" = "MISSTAIS" ]; then
            if [ -z "$WB_KEY_MISSTAIS" ]; then
              echo "::error::Не найден WB_KEY_MISSTAIS для MISSTAIS"
              missing=1
            else
              echo "WB_KEY_MISSTAIS загружен"
            fi
          fi

          if [ "$missing" -ne 0 ]; then
            exit 1
          fi

      - name: Запустить ежедневное обновление данных v7
        env:
          STORE_INPUT: ${{ github.event.inputs.store || 'ALL' }}
        run: |
          echo "__START_PYTHON_FROM_YML_VERSION_v7_20260610_UNIQUE_FILE_NO_REPORT_ENV_STOP__"
          python ezhednevnoe_obnovlenie_dannyh.py --full --store "$STORE_INPUT"
