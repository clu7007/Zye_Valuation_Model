"""Shared schema validation for local financial statement inputs."""

from __future__ import annotations

from pathlib import Path

import pandas as pd


REQUIRED_QUARTERS = [
    "2024Q1",
    "2024Q2",
    "2024Q3",
    "2024Q4",
    "2025Q1",
    "2025Q2",
    "2025Q3",
    "2025Q4",
    "2026Q1",
]

REQUIRED_ITEMS = [
    "Revenue",
    "Cost of revenue",
    "Gross profit",
    "Operating expenses",
    "Operating income",
    "Non-operating income and expenses",
    "Profit before tax",
    "Tax expense",
    "Net income",
    "Net income attributable to parent company",
    "EPS",
]


def validate_financial_input(raw_data: pd.DataFrame, input_path: str | Path) -> pd.DataFrame:
    """Validate and normalize local CSV/Excel financial statement input."""

    source_name = str(input_path)
    data = raw_data.copy()
    data.columns = [str(column).strip() for column in data.columns]

    missing_columns = ["Item"] if "Item" not in data.columns else []
    missing_columns.extend(
        quarter for quarter in REQUIRED_QUARTERS if quarter not in data.columns
    )
    if missing_columns:
        raise ValueError(
            f"{source_name} is missing required column(s): "
            f"{', '.join(missing_columns)}. Expected first column 'Item' plus "
            f"quarters: {', '.join(REQUIRED_QUARTERS)}."
        )

    data["Item"] = data["Item"].astype(str).str.strip()
    duplicated_items = data.loc[data["Item"].duplicated(), "Item"].tolist()
    if duplicated_items:
        raise ValueError(
            f"{source_name} has duplicate Item row(s): {', '.join(duplicated_items)}."
        )

    available_items = set(data["Item"])
    missing_items = [item for item in REQUIRED_ITEMS if item not in available_items]
    if missing_items:
        raise ValueError(
            f"{source_name} is missing required Item row(s): "
            f"{', '.join(missing_items)}."
        )

    financials = data.set_index("Item").loc[REQUIRED_ITEMS, REQUIRED_QUARTERS]
    for quarter in REQUIRED_QUARTERS:
        converted = pd.to_numeric(financials[quarter], errors="coerce")
        bad_items = financials.index[converted.isna() & financials[quarter].notna()].tolist()
        if bad_items:
            raise ValueError(
                f"{source_name} has non-numeric value(s) in {quarter} for Item row(s): "
                f"{', '.join(bad_items)}."
            )
        financials[quarter] = converted

    return financials
