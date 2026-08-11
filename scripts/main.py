"""Command-line entry point for Taiwan stock quarterly financial exports."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd

from data_sources.csv_loader import load_csv_financial_data
from data_sources.excel_loader import load_excel_financial_data
from data_sources.sample_data import generate_sample_financial_data
from excel_export.exporter import export_quarterly_financials
from financial_analysis.metrics import calculate_growth, calculate_margins


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "output" / "excel"
DEFAULT_DEBUG_DIR = PROJECT_ROOT / "output" / "debug"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate a quarterly financial Excel workbook for any Taiwan stock ID."
    )
    parser.add_argument(
        "--stock_id",
        required=True,
        help="Taiwan stock ID to export, for example 2330, 2317, 2454, or 7722.",
    )
    parser.add_argument(
        "--source",
        default="csv",
        choices=["csv", "excel", "sample"],
        help="Data source to use. Use csv or excel for real local files.",
    )
    parser.add_argument(
        "--input_path",
        help="Local CSV or Excel input path, required when --source is csv or excel.",
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        help="Print raw and calculated DataFrames to the terminal for audit.",
    )
    parser.add_argument(
        "--save_raw",
        action="store_true",
        help="Save raw and calculated audit CSV files under output/debug.",
    )
    return parser.parse_args()


def validate_stock_id(stock_id: str) -> str:
    cleaned = stock_id.strip()
    if not cleaned:
        raise ValueError("--stock_id is required and cannot be blank.")
    if not cleaned.isalnum():
        raise ValueError("--stock_id should contain only letters and numbers.")
    return cleaned


def build_output_path(stock_id: str) -> Path:
    return DEFAULT_OUTPUT_DIR / f"{stock_id}_quarterly_financials.xlsx"


def resolve_input_path(input_path: str | None) -> Path | None:
    if input_path is None:
        return None
    path = Path(input_path)
    if path.is_absolute():
        return path
    return PROJECT_ROOT / path


def load_income_statement(
    stock_id: str, source: str, input_path: Path | None
) -> pd.DataFrame:
    if source in {"csv", "excel"} and input_path is None:
        raise ValueError(f"--input_path is required when --source is {source}.")

    if source == "csv":
        return load_csv_financial_data(input_path)
    if source == "excel":
        return load_excel_financial_data(input_path)
    if source == "sample":
        return generate_sample_financial_data(stock_id)
    raise ValueError(f"Unsupported source: {source}")


def print_debug_report(
    stock_id: str,
    source: str,
    input_path: Path | None,
    income_statement: pd.DataFrame,
    qoq_growth: pd.DataFrame,
    yoy_growth: pd.DataFrame,
    margins: pd.DataFrame,
    output_path: Path,
) -> None:
    print("\n=== Debug / Audit Report ===")
    print(f"stock_id: {stock_id}")
    print(f"source: {source}")
    print(f"input_path: {input_path if input_path else '(not used)'}")
    print(f"quarters: {', '.join(income_statement.columns)}")
    print("\n--- Raw DataFrame ---")
    print(income_statement.to_string())
    print("\n--- QoQ DataFrame ---")
    print(qoq_growth.to_string())
    print("\n--- YoY DataFrame ---")
    print(yoy_growth.to_string())
    print("\n--- Margin DataFrame ---")
    print(margins.to_string())
    print(f"\noutput_path: {output_path}")
    print("=== End Debug / Audit Report ===\n")


def save_audit_files(
    stock_id: str,
    income_statement: pd.DataFrame,
    qoq_growth: pd.DataFrame,
    yoy_growth: pd.DataFrame,
    margins: pd.DataFrame,
) -> dict[str, Path]:
    DEFAULT_DEBUG_DIR.mkdir(parents=True, exist_ok=True)

    audit_files = {
        "raw_data": DEFAULT_DEBUG_DIR / f"{stock_id}_raw_data.csv",
        "qoq": DEFAULT_DEBUG_DIR / f"{stock_id}_qoq.csv",
        "yoy": DEFAULT_DEBUG_DIR / f"{stock_id}_yoy.csv",
        "margins": DEFAULT_DEBUG_DIR / f"{stock_id}_margins.csv",
    }

    income_statement.to_csv(audit_files["raw_data"], index_label="Financial item")
    qoq_growth.to_csv(audit_files["qoq"], index_label="Financial item")
    yoy_growth.to_csv(audit_files["yoy"], index_label="Financial item")
    margins.to_csv(audit_files["margins"], index_label="Financial item")
    return audit_files


def run() -> None:
    args = parse_args()
    stock_id = validate_stock_id(args.stock_id)
    source = args.source
    input_path = resolve_input_path(args.input_path)

    income_statement = load_income_statement(stock_id, source, input_path)
    qoq_growth = calculate_growth(income_statement, periods=1)
    yoy_growth = calculate_growth(income_statement, periods=4)
    margins = calculate_margins(income_statement)

    output_path = build_output_path(stock_id)
    export_quarterly_financials(
        stock_id=stock_id,
        income_statement=income_statement,
        qoq_growth=qoq_growth,
        yoy_growth=yoy_growth,
        margins=margins,
        output_path=output_path,
    )

    if args.save_raw:
        audit_files = save_audit_files(
            stock_id=stock_id,
            income_statement=income_statement,
            qoq_growth=qoq_growth,
            yoy_growth=yoy_growth,
            margins=margins,
        )
        print("Saved audit files:")
        for label, path in audit_files.items():
            print(f"- {label}: {path}")

    if args.debug:
        print_debug_report(
            stock_id=stock_id,
            source=source,
            input_path=input_path,
            income_statement=income_statement,
            qoq_growth=qoq_growth,
            yoy_growth=yoy_growth,
            margins=margins,
            output_path=output_path,
        )

    print(f"Created Excel file: {output_path}")


def main() -> None:
    try:
        run()
    except (FileNotFoundError, ValueError, pd.errors.ParserError) as error:
        print(f"Error: {error}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
