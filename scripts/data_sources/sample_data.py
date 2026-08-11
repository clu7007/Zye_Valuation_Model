"""Synthetic financial data source for the first MVP.

Future data sources should keep the same stock_id-driven interface so Goodinfo,
MOPS, CSV, and Excel loaders can be swapped in without changing the analysis or
export layers.
"""

from __future__ import annotations

import hashlib
import random

import pandas as pd

from data_sources.schema import REQUIRED_ITEMS, REQUIRED_QUARTERS


def _seed_from_stock_id(stock_id: str) -> int:
    digest = hashlib.sha256(stock_id.encode("utf-8")).hexdigest()
    return int(digest[:12], 16)


def generate_sample_financial_data(stock_id: str) -> pd.DataFrame:
    """Generate deterministic sample quarterly financial data for any stock_id.

    The values are synthetic and should not be used as real company financials.
    They are intentionally stock_id-dependent so different tickers produce
    different Excel outputs while sharing the same pipeline.
    """

    if not stock_id or not stock_id.strip():
        raise ValueError("stock_id is required to generate sample financial data.")

    rng = random.Random(_seed_from_stock_id(stock_id.strip()))
    base_revenue = rng.randint(3_000, 30_000)
    quarterly_growth = rng.uniform(-0.015, 0.045)
    gross_margin_base = rng.uniform(0.28, 0.55)
    opex_ratio_base = rng.uniform(0.08, 0.22)
    tax_rate = rng.uniform(0.16, 0.22)
    shares_outstanding_m = rng.randint(500, 8_000)

    rows: dict[str, list[float]] = {item: [] for item in REQUIRED_ITEMS}

    for index, _quarter in enumerate(REQUIRED_QUARTERS):
        seasonality = [0.97, 1.02, 1.04, 1.08][index % 4]
        noise = rng.uniform(0.965, 1.04)
        revenue = base_revenue * ((1 + quarterly_growth) ** index) * seasonality * noise

        gross_margin = gross_margin_base + rng.uniform(-0.025, 0.025)
        opex_ratio = opex_ratio_base + rng.uniform(-0.015, 0.015)
        non_operating_ratio = rng.uniform(-0.018, 0.024)

        cost_of_revenue = revenue * (1 - gross_margin)
        gross_profit = revenue - cost_of_revenue
        operating_expenses = revenue * opex_ratio
        operating_income = gross_profit - operating_expenses
        non_operating_income = revenue * non_operating_ratio
        profit_before_tax = operating_income + non_operating_income
        tax_expense = max(profit_before_tax * tax_rate, 0)
        net_income = profit_before_tax - tax_expense
        parent_net_income = net_income * rng.uniform(0.965, 1.0)
        eps = parent_net_income / shares_outstanding_m

        rows["Revenue"].append(round(revenue, 0))
        rows["Cost of revenue"].append(round(cost_of_revenue, 0))
        rows["Gross profit"].append(round(gross_profit, 0))
        rows["Operating expenses"].append(round(operating_expenses, 0))
        rows["Operating income"].append(round(operating_income, 0))
        rows["Non-operating income and expenses"].append(round(non_operating_income, 0))
        rows["Profit before tax"].append(round(profit_before_tax, 0))
        rows["Tax expense"].append(round(tax_expense, 0))
        rows["Net income"].append(round(net_income, 0))
        rows["Net income attributable to parent company"].append(round(parent_net_income, 0))
        rows["EPS"].append(round(eps, 2))

    return pd.DataFrame(rows, index=REQUIRED_QUARTERS).T
