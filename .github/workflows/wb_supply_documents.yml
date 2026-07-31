name: Документы по поставкам TOPFACE V5

run-name: Поставки и УПД TOPFACE V5 — ${{ github.event.inputs.mode || 'auto' }}

on:
  schedule:
    # Каждый день в 19:30 МСК / 16:30 UTC.
    - cron: "30 16 * * *"
  workflow_dispatch:
    inputs:
      mode:
        description: "Режим запуска"
        required: true
        default: "auto"
        type: choice
        options:
          - auto
          - bootstrap
      lookback_days:
        description: "Сколько дней поставок проверять"
        required: true
        default: "45"
      document_lookback_days:
        description: "Сколько дней документов WB проверять"
        required: true
        default: "21"
      force_notify:
        description: "Повторно прислать актуальные принятые поставки"
        required: false
        default: false
        type: boolean
      dry_run:
        description: "Проверить настройки без запросов к WB"
        required: false
        default: false
        type: boolean

concurrency:
  group: wb-supply-documents-topface
  cancel-in-progress: false

jobs:
  wb-supply-documents:
    name: Контроль поставок, УПД и актов TOPFACE
    runs-on: ubuntu-latest
    timeout-minutes: 90

    # Используются только уже существующие секреты репозитория.
    env:
      REPORT_ENV: ${{ secrets.REPORT_ENV }}
      WB_PROMO_KEY_TOPFACE: ${{ secrets.WB_PROMO_KEY_TOPFACE }}
      TELEGRAM_BOT_TOKEN: ${{ secrets.TELEGRAM_BOT_TOKEN }}
      TELEGRAM_CHAT_ID: ${{ secrets.TELEGRAM_CHAT_ID }}

      AWS_DEFAULT_REGION: ru-central1
      TZ: Europe/Moscow
      PYTHONUNBUFFERED: "1"

      WB_SUPPLY_DOCS_ROOT: "Документы по поставкам/TOPFACE"
      WB_NOTIFY_STATUS_IDS: "4,5,6"
      WB_DEEP_LOOKBACK_DAYS: "21"
      WB_MAX_DEEP_SUPPLIES: "180"
      WB_MAX_DOCUMENT_DOWNLOADS: "20"
      WB_FIRST_RUN_NOTIFY_DAYS: "3"
      TELEGRAM_SEND_UPD_FILES: "1"

    steps:
      - name: Checkout repository
        uses: actions/checkout@v4

      - name: Setup Python
        uses: actions/setup-python@v5
        with:
          python-version: "3.13"

      - name: Install dependencies
        shell: bash
        run: |
          set -euo pipefail
          python -m pip install --upgrade pip
          python -m pip install requests boto3 openpyxl

      # Повторяет уже работающую схему других workflow:
      # REPORT_ENV разбирается и переносится в GITHUB_ENV, а секрет TOPFACE
      # одновременно публикуется под именем, которое ожидает старая версия скрипта.
      - name: Load existing secrets and WB aliases
        shell: bash
        run: |
          set -euo pipefail
          python - <<'PY'
          import os

          allowed_from_report_env = {
              "YC_ACCESS_KEY_ID",
              "YC_SECRET_ACCESS_KEY",
              "YC_BUCKET_NAME",
              "YC_ENDPOINT_URL",
              "TELEGRAM_MESSAGE_THREAD_ID",
              "TELEGRAM_THREAD_ID",
          }

          report_env = os.getenv("REPORT_ENV", "") or ""
          parsed = {}
          normalized = report_env.replace("\r\n", "\n").replace("\r", "\n")
          if "\n" not in normalized and "\\n" in normalized:
              normalized = normalized.replace("\\n", "\n")

          for raw_line in normalized.splitlines():
              line = raw_line.strip()
              if not line or line.startswith("#"):
                  continue
              if line.startswith("export "):
                  line = line[len("export "):].strip()
              if "=" not in line:
                  continue
              key, value = line.split("=", 1)
              key = key.strip().lstrip("\ufeff")
              value = value.strip().strip('"').strip("'")
              if key in allowed_from_report_env and value:
                  parsed[key] = value

          wb_token = (os.getenv("WB_PROMO_KEY_TOPFACE", "") or "").strip()
          if not wb_token:
              # На случай, если токен когда-нибудь перенесут внутрь REPORT_ENV.
              for raw_line in normalized.splitlines():
                  line = raw_line.strip()
                  if line.startswith("export "):
                      line = line[len("export "):].strip()
                  if "=" not in line:
                      continue
                  key, value = line.split("=", 1)
                  if key.strip() == "WB_PROMO_KEY_TOPFACE":
                      wb_token = value.strip().strip('"').strip("'")
                      break

          github_env = os.environ.get("GITHUB_ENV")
          if not github_env:
              raise SystemExit("GITHUB_ENV is not available")

          loaded_names = []
          with open(github_env, "a", encoding="utf-8") as f:
              for key in sorted(allowed_from_report_env):
                  value = (os.getenv(key, "") or parsed.get(key, "")).strip()
                  if value:
                      f.write(f"{key}={value}\n")
                      loaded_names.append(key)

              if wb_token:
                  # Совместимость сразу со всеми версиями wb_supply_documents.py.
                  for key in (
                      "WB_PROMO_KEY_TOPFACE",
                      "WB_API_TOKEN",
                      "TOPFACE_WB_API_TOKEN",
                      "WB_TOKEN",
                      "WILDBERRIES_API_TOKEN",
                  ):
                      f.write(f"{key}={wb_token}\n")
                  loaded_names.extend([
                      "WB_PROMO_KEY_TOPFACE",
                      "WB_API_TOKEN",
                      "TOPFACE_WB_API_TOKEN",
                      "WB_TOKEN",
                      "WILDBERRIES_API_TOKEN",
                  ])

          print("Loaded environment names: " + ", ".join(sorted(set(loaded_names))))
          PY

      - name: Verify existing secrets and script
        shell: bash
        run: |
          set -euo pipefail
          python - <<'PY'
          import os
          from pathlib import Path

          required = [
              "YC_ACCESS_KEY_ID",
              "YC_SECRET_ACCESS_KEY",
              "YC_BUCKET_NAME",
              "YC_ENDPOINT_URL",
              "WB_API_TOKEN",
              "TELEGRAM_BOT_TOKEN",
              "TELEGRAM_CHAT_ID",
          ]
          missing = [name for name in required if not (os.getenv(name, "") or "").strip()]
          if missing:
              raise SystemExit("Не найдены существующие настройки: " + ", ".join(missing))

          script = Path("wb_supply_documents_TOPFACE_v5_20260731.py")
          if not script.is_file():
              raise SystemExit("Не найден ./wb_supply_documents_TOPFACE_v5_20260731.py")

          text = script.read_text(encoding="utf-8", errors="replace")
          if "WB_SUPPLY_DOCUMENTS_TOPFACE" not in text:
              raise SystemExit("В ./wb_supply_documents_TOPFACE_v5_20260731.py нет маркера проекта WB_SUPPLY_DOCUMENTS_TOPFACE")

          print("Все существующие секреты доступны.")
          print("WB token alias: WB_PROMO_KEY_TOPFACE -> WB_API_TOKEN; значение не выводится.")
          print("Script: ./wb_supply_documents_TOPFACE_v5_20260731.py")
          PY

          python -m py_compile ./wb_supply_documents_TOPFACE_v5_20260731.py
          python ./wb_supply_documents_TOPFACE_v5_20260731.py --self-test
          echo "WB_SUPPLY_DOCUMENTS_SCRIPT=./wb_supply_documents_TOPFACE_v5_20260731.py" >> "$GITHUB_ENV"

      - name: Run supply and document monitor
        shell: bash
        run: |
          set -euo pipefail

          MODE="${{ github.event.inputs.mode || 'auto' }}"
          LOOKBACK="${{ github.event.inputs.lookback_days || '45' }}"
          DOC_LOOKBACK="${{ github.event.inputs.document_lookback_days || '21' }}"
          FORCE_NOTIFY="${{ github.event.inputs.force_notify || 'false' }}"
          DRY_RUN="${{ github.event.inputs.dry_run || 'false' }}"

          ARGS=(
            --lookback-days "$LOOKBACK"
            --document-lookback-days "$DOC_LOOKBACK"
            --root-prefix "Документы по поставкам/TOPFACE"
          )

          if [ "$MODE" = "bootstrap" ]; then
            ARGS+=(--first-run-notify-days "$LOOKBACK")
          fi
          if [ "$FORCE_NOTIFY" = "true" ]; then
            ARGS+=(--force-notify)
          fi
          if [ "$DRY_RUN" = "true" ]; then
            ARGS+=(--dry-run)
          fi

          echo "script=$WB_SUPPLY_DOCUMENTS_SCRIPT"
          echo "mode=$MODE lookback=$LOOKBACK document_lookback=$DOC_LOOKBACK force_notify=$FORCE_NOTIFY dry_run=$DRY_RUN"
          python "$WB_SUPPLY_DOCUMENTS_SCRIPT" "${ARGS[@]}"
