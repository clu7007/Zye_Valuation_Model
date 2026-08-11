"""Download Taiwan stock financial statements via FinMind and save as CSV."""

from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

import pandas as pd
import requests


PROJECT_ROOT = Path(__file__).resolve().parents[1]
OUTPUT_DIR = PROJECT_ROOT / "output" / "finmind"
FINMIND_API_URL = "https://api.finmindtrade.com/api/v4/data"

# FinMind dataset names -> (output filename suffix, 中文名稱)
STATEMENTS: dict[str, tuple[str, str]] = {
    "income_statement": ("TaiwanStockFinancialStatements", "綜合損益表"),
    "balance_sheet": ("TaiwanStockBalanceSheet", "資產負債表"),
    "cash_flow": ("TaiwanStockCashFlowsStatement", "現金流量表"),
}

# Default history window
DEFAULT_START_DATE = "2019-01-01"

# Seconds to wait between API calls to avoid hitting rate limits
REQUEST_DELAY = 1.5


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="使用 FinMind 下載台股三大財報並存成 CSV。"
    )
    parser.add_argument(
        "--stocks",
        nargs="+",
        required=True,
        metavar="STOCK_ID",
        help="一或多個台股代號，例如 --stocks 2330 7856",
    )
    parser.add_argument(
        "--start_date",
        default=DEFAULT_START_DATE,
        help=f"資料起始日期，格式 YYYY-MM-DD（預設 {DEFAULT_START_DATE}）",
    )
    return parser.parse_args()


def fetch_statement(
    stock_id: str,
    dataset: str,
    start_date: str,
) -> pd.DataFrame:
    params: dict[str, str] = {
        "dataset": dataset,
        "data_id": stock_id,
        "start_date": start_date,
    }
    response = requests.get(FINMIND_API_URL, params=params, timeout=30)
    response.raise_for_status()

    payload = response.json()
    status = payload.get("status")
    if status != 200:
        msg = payload.get("msg", "unknown error")
        raise RuntimeError(f"FinMind API 錯誤（status={status}）：{msg}")

    records = payload.get("data", [])
    if not records:
        raise RuntimeError(
            f"FinMind 回傳空資料：stock_id={stock_id}, dataset={dataset}"
        )

    return pd.DataFrame(records)


def save_csv(df: pd.DataFrame, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(path, index=False, encoding="utf-8-sig")


def download_stock(stock_id: str, start_date: str) -> None:
    print(f"\n[{stock_id}] 開始下載...")
    for suffix, (dataset, label) in STATEMENTS.items():
        output_path = OUTPUT_DIR / f"{stock_id}_{suffix}.csv"
        try:
            df = fetch_statement(stock_id, dataset, start_date)
            save_csv(df, output_path)
            print(f"  ✓ {label} → {output_path.relative_to(PROJECT_ROOT)}")
        except RuntimeError as err:
            print(f"  ✗ {label} 失敗：{err}", file=sys.stderr)
        except requests.RequestException as err:
            print(f"  ✗ {label} 網路錯誤：{err}", file=sys.stderr)
        time.sleep(REQUEST_DELAY)


def main() -> None:
    args = parse_args()
    stocks: list[str] = [s.strip() for s in args.stocks if s.strip()]
    if not stocks:
        print("Error: 請至少提供一個股票代號。", file=sys.stderr)
        sys.exit(1)

    print(f"資料起始日期：{args.start_date}")
    print(f"輸出資料夾：{OUTPUT_DIR.relative_to(PROJECT_ROOT)}")
    print(f"待下載股票：{', '.join(stocks)}")

    for stock_id in stocks:
        download_stock(stock_id, args.start_date)

    print("\n完成。")


if __name__ == "__main__":
    main()
