"""Financial metric calculations for quarterly Taiwan stock analysis."""

from __future__ import annotations

import pandas as pd


def calculate_growth(financials: pd.DataFrame, periods: int) -> pd.DataFrame:
    """Calculate quarter-over-quarter or year-over-year growth by row."""

    if periods <= 0:
        raise ValueError("periods must be greater than zero.")
    growth = financials.astype(float).pct_change(periods=periods, axis=1)
    return growth.replace([float("inf"), float("-inf")], pd.NA)


def _safe_divide(numerator: pd.Series, denominator: pd.Series) -> pd.Series:
    result = numerator.astype(float) / denominator.astype(float)
    return result.replace([float("inf"), float("-inf")], pd.NA)


def calculate_margins(financials: pd.DataFrame) -> pd.DataFrame:
    """Calculate profitability margins as a percentage of revenue."""

    revenue = financials.loc["Revenue"]
    margins = pd.DataFrame(index=[], columns=financials.columns)
    margins.loc["Gross margin"] = _safe_divide(financials.loc["Gross profit"], revenue)
    margins.loc["Operating margin"] = _safe_divide(
        financials.loc["Operating income"], revenue
    )
    margins.loc["Pre-tax margin"] = _safe_divide(
        financials.loc["Profit before tax"], revenue
    )
    margins.loc["Net margin"] = _safe_divide(financials.loc["Net income"], revenue)
    return margins
