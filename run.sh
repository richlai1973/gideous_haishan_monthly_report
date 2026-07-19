#!/usr/bin/env bash
# 啟動月例會報告產出系統。用法：bash run.sh
set -euo pipefail
cd "$(dirname "$0")"

if [ ! -d .venv ]; then
  echo "▸ 首次執行，建立虛擬環境…"
  python3 -m venv .venv
fi
source .venv/bin/activate

if ! python -c "import fastapi, docx, openpyxl" 2>/dev/null; then
  echo "▸ 安裝套件…"
  pip install -q -r requirements.txt
fi

if python -c "import pytest" 2>/dev/null; then
  echo "▸ 執行測試…"
  python -m pytest tests/ -q
fi

echo
echo "▸ 啟動：http://127.0.0.1:8848"
python app.py
