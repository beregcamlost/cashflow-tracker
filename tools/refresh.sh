#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"

[ -f .env ] && . ./.env
: "${CASHFLOW_SHEET_ID:?falta CASHFLOW_SHEET_ID (define tools/.env, ver .env.example)}"
: "${CASA_SHEET_ID:?falta CASA_SHEET_ID (define tools/.env, ver .env.example)}"

OUT_DIR="${1:-out}"
mkdir -p "$OUT_DIR"

rclone backend copyid gdrive-personal: "$CASHFLOW_SHEET_ID" "$OUT_DIR/Cashflow.xlsx" \
  --drive-export-formats xlsx 2>/dev/null
rclone backend copyid gdrive-personal: "$CASA_SHEET_ID" "$OUT_DIR/Gordo.xlsx" \
  --drive-export-formats xlsx 2>/dev/null

if [ ! -d .venv ]; then
  python3 -m venv .venv
  .venv/bin/pip install -q pandas openpyxl
fi

.venv/bin/python analisis.py "$OUT_DIR/Cashflow.xlsx" "$OUT_DIR/Gordo.xlsx"
