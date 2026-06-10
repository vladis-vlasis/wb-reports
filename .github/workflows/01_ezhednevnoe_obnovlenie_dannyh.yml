# VERSION: v11_restore_working_20260610
# Основа: последняя рабочая схема запуска через REPORT_ENV.
# Изменения минимальные: добавлен WB_KEY_MISSTAIS, запуск --store ALL, cron 10:00 МСК.

name: Ежедневное обновление данных

on:
  workflow_dispatch:
  schedule:
    # 10:00 МСК = 07:00 UTC
    - cron: "0 7 * * *"

jobs:
  run:
    name: Запустить ежедневное обновление данных
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

      - name: Загрузить переменные из REPORT_ENV и отдельные secrets
        env:
          REPORT_ENV: ${{ secrets.REPORT_ENV }}
          WB_PROMO_KEY_TOPFACE: ${{ secrets.WB_PROMO_KEY_TOPFACE }}
          WB_KEY_MISSTAIS: ${{ secrets.WB_KEY_MISSTAIS }}
          TELEGRAM_BOT_TOKEN: ${{ secrets.TELEGRAM_BOT_TOKEN }}
          TELEGRAM_CHAT_ID: ${{ secrets.TELEGRAM_CHAT_ID }}
          TORGSTAT_ABC_URL: ${{ secrets.TORGSTAT_ABC_URL }}
        run: |
          echo "__RUNNING_YML_VERSION_v11_restore_working_01__"

          if [ -z "$REPORT_ENV" ]; then
            echo "::error::Не задан secret REPORT_ENV"
            exit 1
          fi

          printf '%s
' "$REPORT_ENV" | sed 's/
$//' >> "$GITHUB_ENV"

          if [ -n "$WB_PROMO_KEY_TOPFACE" ]; then
            echo "WB_PROMO_KEY_TOPFACE=$WB_PROMO_KEY_TOPFACE" >> "$GITHUB_ENV"
          fi
          if [ -n "$WB_KEY_MISSTAIS" ]; then
            echo "WB_KEY_MISSTAIS=$WB_KEY_MISSTAIS" >> "$GITHUB_ENV"
          fi
          if [ -n "$TELEGRAM_BOT_TOKEN" ]; then
            echo "TELEGRAM_BOT_TOKEN=$TELEGRAM_BOT_TOKEN" >> "$GITHUB_ENV"
          fi
          if [ -n "$TELEGRAM_CHAT_ID" ]; then
            echo "TELEGRAM_CHAT_ID=$TELEGRAM_CHAT_ID" >> "$GITHUB_ENV"
          fi
          if [ -n "$TORGSTAT_ABC_URL" ]; then
            echo "TORGSTAT_ABC_URL=$TORGSTAT_ABC_URL" >> "$GITHUB_ENV"
          fi

      - name: Запустить ежедневное обновление данных
        run: |
          python ezhednevnoe_obnovlenie_dannyh.py --full --store ALL
