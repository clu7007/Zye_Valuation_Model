#!/bin/bash
# macOS 雙擊入口（對應 Windows 的 run_income_model.bat）
cd "$(dirname "$0")" || exit 1

echo "=========================================="
echo "Taiwan Stock Income Model Generator (macOS)"
echo "=========================================="
echo
echo "Enter stock codes separated by spaces."
echo "Example: 2330"
echo "Example: 2330 3532 2454"
echo

read -r -p "Stock codes: " STOCKS

if [ -z "$STOCKS" ]; then
    echo
    echo "No stock code entered."
    read -r -p "Press Enter to close..."
    exit 1
fi

# 找 Python 3（優先 3.12，其次 python3）
PY=""
for cand in python3.12 python3; do
    if command -v "$cand" >/dev/null 2>&1; then PY="$cand"; break; fi
done
if [ -z "$PY" ]; then
    echo "找不到 Python 3。請先安裝：brew install python@3.12 （或到 python.org 下載）"
    read -r -p "Press Enter to close..."
    exit 1
fi

# Windows 的 .venv 在 Mac 上不能用，所以 Mac 用獨立的 .venv-mac
if [ ! -x ".venv-mac/bin/python" ]; then
    echo "First run: creating .venv-mac and installing requirements..."
    "$PY" -m venv .venv-mac || { read -r -p "venv 建立失敗，按 Enter 關閉..."; exit 1; }
    .venv-mac/bin/python -m pip install --quiet --upgrade pip
    .venv-mac/bin/python -m pip install -r requirements.txt || { read -r -p "套件安裝失敗，按 Enter 關閉..."; exit 1; }
fi

echo
echo "Generating Excel model for: $STOCKS"
echo

# shellcheck disable=SC2086  # 刻意讓 $STOCKS 依空白拆成多個參數
.venv-mac/bin/python scripts/build_income_statement_model.py --stocks $STOCKS --start_year 2024
STATUS=$?

if [ $STATUS -ne 0 ]; then
    echo
    echo "Failed to generate model (exit code $STATUS). 請看上方錯誤訊息。"
    read -r -p "Press Enter to close..."
    exit $STATUS
fi

echo
echo "Done. Opening output folder..."
open output/models

echo
read -r -p "Press Enter to close..."
