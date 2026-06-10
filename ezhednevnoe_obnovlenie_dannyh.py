# VERSION: v10_20260610_KEEP_REPORT_ENV
# ВАЖНО: REPORT_ENV сохранён. Он нужен для Yandex Cloud Object Storage.
# В GitHub нужно полностью заменить содержимое:
# .github/workflows/01_ezhednevnoe_obnovlenie_dannyh.yml

name: Ежедневное обновление данных MISSTAIS v10

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
    name: Запустить ежедневное обновление данных MISSTAIS v10
    runs-on: ubuntu-latest

    steps:
      - name: Маркер версии v10
        run: |
          echo "__RUNNING_YML_VERSION_v10_20260610_KEEP_REPORT_ENV__"

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

      - name: Загрузить REPORT_ENV и ключи магазинов v10
        env:
          REPORT_ENV: ${{ secrets.REPORT_ENV }}
          WB_PROMO_KEY_TOPFACE: ${{ secrets.WB_PROMO_KEY_TOPFACE }}
          WB_KEY_MISSTAIS: ${{ secrets.WB_KEY_MISSTAIS }}
          TELEGRAM_BOT_TOKEN: ${{ secrets.TELEGRAM_BOT_TOKEN }}
          TELEGRAM_CHAT_ID: ${{ secrets.TELEGRAM_CHAT_ID }}
          TORGSTAT_ABC_URL: ${{ secrets.TORGSTAT_ABC_URL }}
        run: |
          echo "__LOAD_REPORT_ENV_STEP_v10__"

          if test -z "$REPORT_ENV"; then
            echo "::error::GitHub не передал secret REPORT_ENV в этот запуск. REPORT_ENV нужен: внутри него лежат YC_ACCESS_KEY_ID, YC_SECRET_ACCESS_KEY, YC_BUCKET_NAME и другие переменные для Yandex Cloud."
            echo "::error::Проверь, что REPORT_ENV создан именно в Settings -> Secrets and variables -> Actions для этого репозитория/ветки. Если это Environment secret, workflow должен быть привязан к нужному environment."
            exit 1
          fi

          printf '%s\n' "$REPORT_ENV" | sed 's/\r$//' >> "$GITHUB_ENV"
          echo "REPORT_ENV загружен в GITHUB_ENV"

          if test -n "$WB_PROMO_KEY_TOPFACE"; then echo "WB_PROMO_KEY_TOPFACE=$WB_PROMO_KEY_TOPFACE" >> "$GITHUB_ENV"; fi
          if test -n "$WB_KEY_MISSTAIS"; then echo "WB_KEY_MISSTAIS=$WB_KEY_MISSTAIS" >> "$GITHUB_ENV"; fi
          if test -n "$TELEGRAM_BOT_TOKEN"; then echo "TELEGRAM_BOT_TOKEN=$TELEGRAM_BOT_TOKEN" >> "$GITHUB_ENV"; fi
          if test -n "$TELEGRAM_CHAT_ID"; then echo "TELEGRAM_CHAT_ID=$TELEGRAM_CHAT_ID" >> "$GITHUB_ENV"; fi
          if test -n "$TORGSTAT_ABC_URL"; then echo "TORGSTAT_ABC_URL=$TORGSTAT_ABC_URL" >> "$GITHUB_ENV"; fi

      - name: Проверить переменные после REPORT_ENV v10
        env:
          STORE_INPUT_RAW: ${{ github.event.inputs.store }}
        run: |
          echo "__CHECK_ENV_AFTER_REPORT_ENV_v10__"

          STORE_INPUT="${STORE_INPUT_RAW:-ALL}"
          echo "STORE_INPUT=$STORE_INPUT"

          missing=0

          for var in YC_ACCESS_KEY_ID YC_SECRET_ACCESS_KEY YC_BUCKET_NAME; do
            if test -z "${!var}"; then
              echo "::error::После загрузки REPORT_ENV не найдена переменная $var. Проверь содержимое REPORT_ENV."
              missing=1
            else
              echo "$var найден"
            fi
          done

          if [ "$STORE_INPUT" = "ALL" ] || [ "$STORE_INPUT" = "TOPFACE" ]; then
            if test -z "$WB_PROMO_KEY_TOPFACE"; then
              echo "::error::Не найден WB_PROMO_KEY_TOPFACE для TOPFACE"
              missing=1
            else
              echo "WB_PROMO_KEY_TOPFACE найден"
            fi
          fi

          if [ "$STORE_INPUT" = "ALL" ] || [ "$STORE_INPUT" = "MISSTAIS" ]; then
            if test -z "$WB_KEY_MISSTAIS"; then
              echo "::error::Не найден WB_KEY_MISSTAIS для MISSTAIS"
              missing=1
            else
              echo "WB_KEY_MISSTAIS найден"
            fi
          fi

          if [ "$missing" -ne 0 ]; then
            exit 1
          fi

      - name: Запустить ежедневное обновление данных v10
        env:
          STORE_INPUT_RAW: ${{ github.event.inputs.store }}
        run: |
          echo "__START_PYTHON_FROM_YML_VERSION_v10_20260610_KEEP_REPORT_ENV__"

          STORE_INPUT="${STORE_INPUT_RAW:-ALL}"
          python ezhednevnoe_obnovlenie_dannyh.py --full --store "$STORE_INPUT"
