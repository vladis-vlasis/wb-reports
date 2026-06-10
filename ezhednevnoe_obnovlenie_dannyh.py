name: Ежедневное обновление данных

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

      - name: Загрузить переменные из REPORT_ENV и отдельных secrets
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
          if [ -n "$REPORT_ENV" ]; then
            printf '%s\n' "$REPORT_ENV" | sed 's/\r$//' >> "$GITHUB_ENV"
            echo "REPORT_ENV загружен"
          else
            echo "::warning::REPORT_ENV пустой или недоступен для этого workflow. Пробую отдельные secrets."
          fi

          add_env_if_set() {
            name="$1"
            value="$2"
            if [ -n "$value" ]; then
              echo "$name=$value" >> "$GITHUB_ENV"
              echo "$name загружен"
            fi
          }

          add_env_if_set "YC_ACCESS_KEY_ID" "$YC_ACCESS_KEY_ID"
          add_env_if_set "YC_SECRET_ACCESS_KEY" "$YC_SECRET_ACCESS_KEY"
          add_env_if_set "YC_BUCKET_NAME" "$YC_BUCKET_NAME"
          add_env_if_set "WB_PROMO_KEY_TOPFACE" "$WB_PROMO_KEY_TOPFACE"
          add_env_if_set "WB_KEY_MISSTAIS" "$WB_KEY_MISSTAIS"
          add_env_if_set "TELEGRAM_BOT_TOKEN" "$TELEGRAM_BOT_TOKEN"
          add_env_if_set "TELEGRAM_CHAT_ID" "$TELEGRAM_CHAT_ID"
          add_env_if_set "TORGSTAT_ABC_URL" "$TORGSTAT_ABC_URL"

      - name: Проверить доступ к Object Storage
        run: |
          missing=0
          for var in YC_ACCESS_KEY_ID YC_SECRET_ACCESS_KEY YC_BUCKET_NAME; do
            if [ -z "${!var}" ]; then
              echo "::error::Не найдена переменная $var. Она должна быть либо внутри REPORT_ENV, либо отдельным GitHub Secret."
              missing=1
            fi
          done

          if [ "$missing" -ne 0 ]; then
            exit 1
          fi

      - name: Проверить API-ключи магазинов
        env:
          STORE_INPUT: ${{ github.event.inputs.store || 'ALL' }}
        run: |
          missing=0

          if [ "$STORE_INPUT" = "ALL" ] || [ "$STORE_INPUT" = "TOPFACE" ]; then
            if [ -z "$WB_PROMO_KEY_TOPFACE" ]; then
              echo "::error::Не найден WB_PROMO_KEY_TOPFACE для TOPFACE"
              missing=1
            fi
          fi

          if [ "$STORE_INPUT" = "ALL" ] || [ "$STORE_INPUT" = "MISSTAIS" ]; then
            if [ -z "$WB_KEY_MISSTAIS" ]; then
              echo "::error::Не найден WB_KEY_MISSTAIS для MISSTAIS"
              missing=1
            fi
          fi

          if [ "$missing" -ne 0 ]; then
            exit 1
          fi

      - name: Запустить ежедневное обновление данных
        env:
          STORE_INPUT: ${{ github.event.inputs.store || 'ALL' }}
        run: |
          python ezhednevnoe_obnovlenie_dannyh.py --full --store "$STORE_INPUT"
