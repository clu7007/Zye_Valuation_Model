"""Build quarterly income statement Excel model from FinMind data.
Sheets: Income Model | Assumptions | Dashboard | Raw Data

Usage:
    python scripts/build_income_statement_model.py --stocks 2330
    python scripts/build_income_statement_model.py --stocks 2330 7856 --start_year 2024
"""
from __future__ import annotations

import argparse
import re
import sys
import time
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
import requests
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.data_source import AxDataSource, StrRef
from openpyxl.chart.series import SeriesLabel
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText, Text
from openpyxl.chart.title import Title
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.legend import LegendEntry
from openpyxl.drawing.fill import GradientFillProperties, GradientStop, LinearShadeProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.formatting.rule import FormulaRule
from openpyxl.drawing.text import (
    CharacterProperties,
    Font as DrawFont,
    Paragraph,
    ParagraphProperties,
    RegularTextRun,
)
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

PROJECT_ROOT = Path(__file__).resolve().parents[1]
OUTPUT_DIR   = PROJECT_ROOT / "output" / "models"
FINMIND_URL  = "https://api.finmindtrade.com/api/v4/data"
REQUEST_DELAY = 1.5
UNIT = 1e8  # NT$ -> 億元

BAND_LOOKBACK_YEARS     = 5    # PE / PB band statistics window
DIVIDEND_LOOKBACK_YEARS = 10   # Dividend history table window

INCOME_MODEL_SHEET  = "Income Model"
ASSUMPTIONS_SHEET   = "Assumptions"
REVENUE_BUILD_SHEET = "Revenue Build"
DASHBOARD_SHEET    = "Dashboard"
PE_BAND_SHEET       = "PE Band"
PB_BAND_SHEET       = "PB Band"
DIVIDEND_SHEET      = "Dividend History"
TURNOVER_SHEET      = "Turnover Days"
RAW_DATA_SHEET      = "Raw Data"

# ── Income items ──────────────────────────────────────────────────────────────
# (FinMind type, display label, is_balance_sheet)
INCOME_ITEMS: list[tuple[str, str, bool]] = [
    ("Revenue",                             "營業收入淨額",          False),
    ("CostOfGoodsSold",                     "營業成本",              False),
    ("GrossProfit",                         "營業毛利淨額",          False),
    ("OperatingExpenses",                   "營業費用",              False),
    ("OperatingIncome",                     "營業淨利(損)",          False),
    ("TotalNonoperatingIncomeAndExpense",   "營業外收入及支出合計",  False),
    ("PreTaxIncome",                        "稅前淨利",              False),
    ("TAX",                                 "所得稅費用(利益)",      False),
    ("IncomeFromContinuingOperations",      "繼續營業單位稅後淨利",  False),
    ("EquityAttributableToOwnersOfParent",  "稅後淨利-歸屬母公司",  False),
    ("OrdinaryShare",                       "普通股本",              True),
    ("EPS",                                 "每股盈餘(元)",          False),
]
N_ITEMS  = len(INCOME_ITEMS)
EPS_TYPE = "EPS"
BS_TYPES = {"OrdinaryShare"}

# Bold + thick-top-border on these rows
HIGHLIGHT_TYPES = {
    "Revenue", "GrossProfit", "OperatingIncome",
    "PreTaxIncome", "EquityAttributableToOwnersOfParent", "EPS",
}

# ── Column definitions ────────────────────────────────────────────────────────
# col_type: annual_hist | qtr_helper | qtr_hist | qtr_fcast | annual_fcast
COLUMNS: list[tuple[str, str, str]] = [
    ("2024",      "2024",    "annual_hist"),
    ("24Q1",      "24Q1",    "qtr_helper"),
    ("24Q2",      "24Q2",    "qtr_helper"),
    ("24Q3",      "24Q3",    "qtr_helper"),
    ("24Q4",      "24Q4",    "qtr_helper"),
    ("25Q1",      "25Q1",    "qtr_hist"),
    ("25Q2",      "25Q2",    "qtr_hist"),
    ("25Q3",      "25Q3",    "qtr_hist"),
    ("25Q4",      "25Q4",    "qtr_hist"),
    ("2025",      "2025",    "annual_hist"),
    ("26Q1",      "26Q1",    "qtr_hist"),
    ("26Q2(F)",   "26Q2F",   "qtr_fcast"),
    ("26Q3(F)",   "26Q3F",   "qtr_fcast"),
    ("26Q4(F)",   "26Q4F",   "qtr_fcast"),
    ("2026(F)",   "2026F",   "annual_fcast"),
    ("2027(F)",   "2027F",   "annual_fcast"),
    ("2028(F)",   "2028F",   "annual_fcast"),
    ("2029(F)",   "2029F",   "annual_fcast"),
]

LABEL_COL      = 1
FIRST_DATA_COL = 2
COL_IDX: dict[str, int] = {c[1]: FIRST_DATA_COL + i for i, c in enumerate(COLUMNS)}
COL_LTR: dict[str, str] = {k: get_column_letter(v) for k, v in COL_IDX.items()}
COL_HDR: dict[str, str] = {c[1]: c[0] for c in COLUMNS}
LAST_COL = FIRST_DATA_COL + len(COLUMNS) - 1   # 19 = S

# Quarterly-only column keys, in chronological order (excludes annual columns).
QTR_KEYS: list[str] = [key for _, key, ctype in COLUMNS
                        if ctype in ("qtr_helper", "qtr_hist", "qtr_fcast")]
# First quarter with a full trailing 4-quarter window inside QTR_KEYS (24Q4) through
# the last modeled quarter (26Q4F) — the window used for PE/PB Band charts.
BAND_QTR_KEYS: list[str] = QTR_KEYS[3:]
BAND_HIST_QTR_KEYS: list[str] = [k for k in BAND_QTR_KEYS if not k.endswith("F")]
# Turnover-days window: needs only 1 trailing quarter (not 4), so it can start one
# quarter earlier than the PE/PB band window — meaningfully more history given the
# model's default 2024 start_year.
TURN_QTR_KEYS: list[str] = QTR_KEYS[1:]
TURN_HIST_QTR_KEYS: list[str] = [k for k in TURN_QTR_KEYS if not k.endswith("F")]
# Annual forecast column immediately after the last explicit quarterly forecast column
# (26Q4F) — used to proxy quarters beyond the model's explicit quarterly horizon.
PROXY_ANNUAL_KEY = "2027F"

ANNUAL_KEYS = {"2024", "2025", "2026F", "2027F", "2028F", "2029F"}

# ── Revenue Build column layout ────────────────────────────────────────────────
# Same quarters/years as Income Model minus the 4 hidden 24Q helper columns
# (Revenue Build doesn't need 2024's quarterly split — only the annual total).
RB_COLUMNS: list[tuple[str, str, str]] = [(hdr, key, ctype) for hdr, key, ctype in COLUMNS
                                           if ctype != "qtr_helper"]
RB_COL_IDX: dict[str, int] = {key: 2 + i for i, (_, key, _) in enumerate(RB_COLUMNS)}
RB_COL_LTR: dict[str, str] = {key: get_column_letter(v) for key, v in RB_COL_IDX.items()}
RB_CTYPE:   dict[str, str] = {key: ctype for _, key, ctype in RB_COLUMNS}
RB_LAST_COL = 1 + len(RB_COLUMNS)

RB_HIST_KEYS     = ["2024", "25Q1", "25Q2", "25Q3", "25Q4", "2025", "26Q1"]
RB_DRIVER_Q_KEYS = ["26Q2F", "26Q3F", "26Q4F"]
RB_DRIVER_A_KEYS = ["2027F", "2028F", "2029F"]
RB_Q_PRIOR_KEY   = {"26Q2F": "25Q2", "26Q3F": "25Q3", "26Q4F": "25Q4"}    # same-quarter-last-year base
RB_ANN_PRIOR_COL = {"2027F": "2026F", "2028F": "2027F", "2029F": "2028F"}  # prior column within Revenue Build's own forecast row

# Revenue Build row numbers (fixed layout — see build_revenue_build_sheet).
RB_HIST_REVENUE_ROW     = 5    # historical Revenue (references Income Model)
RB_VOLUME_ROW           = 10
RB_ASP_ROW              = 11
RB_MIX_ROW              = 12
RB_FX_ROW               = 13
RB_IMPLIED_GROWTH_ROW   = 14
RB_OVERRIDE_ROW         = 15
RB_APPLIED_GROWTH_ROW   = 16
RB_FORECAST_REVENUE_ROW = 18

# PE Band cell addresses (fixed layout — see build_pe_band_sheet) reused by
# Scenario Analysis and Backtest so both stay live if these are hand-edited.
PE_BAND_MEAN_CELL  = "$B$4"
PE_BAND_SD_CELL    = "$B$5"
PE_BAND_LATEST_PRICE_ROW = 26   # = READ_HDR+1 in build_pe_band_sheet, given the current 6-row BAND_ROWS layout

SCENARIO_SHEET = "Scenario Analysis"
BACKTEST_SHEET = "Backtest"
SCENARIO_YEARS = ["2026F", "2027F", "2028F", "2029F"]

# ── Margin analysis items ─────────────────────────────────────────────────────
MARGIN_ITEMS: list[tuple[str, str, str]] = [
    ("毛利率",     "GrossProfit",                    "Revenue"),
    ("營業費用率", "OperatingExpenses",              "Revenue"),
    ("營業利益率", "OperatingIncome",                "Revenue"),
    ("稅前淨利率", "PreTaxIncome",                   "Revenue"),
    ("稅後淨利率", "EquityAttributableToOwnersOfParent", "Revenue"),
]
HIGHLIGHT_MARGINS = {"毛利率", "營業利益率", "稅後淨利率"}

# ── Row layout (Income Model) ─────────────────────────────────────────────────
TITLE_ROW    = 1
HDR_ROW      = 2
DATA_START   = 3
DATA_END     = DATA_START + N_ITEMS - 1        # 14

QOQ_HDR      = DATA_END + 2                    # 16
QOQ_START    = QOQ_HDR + 1                     # 17
QOQ_END      = QOQ_START + N_ITEMS - 1         # 28

YOY_HDR      = QOQ_END + 2                     # 30
YOY_START    = YOY_HDR + 1                     # 31
YOY_END      = YOY_START + N_ITEMS - 1         # 42

MARGIN_HDR   = YOY_END + 2                     # 44
MARGIN_START = MARGIN_HDR + 1                  # 45
MARGIN_END   = MARGIN_START + len(MARGIN_ITEMS) - 1  # 49

# ── Assumptions cell addresses ────────────────────────────────────────────────
# Quarterly section rows  (cols B=26Q2F, C=26Q3F, D=26Q4F)
A_Q_REV_R    = 5
A_Q_GM_R     = 6    # gross margin rate
A_Q_OPEX_R   = 7    # opex ratio
A_Q_NONOP_R  = 8    # non-op income (億元)
A_Q_TAX_R    = 9    # tax rate
A_Q_CAP_R    = 10   # ordinary share (億元)
A_Q_COL      = {"26Q2F": "B", "26Q3F": "C", "26Q4F": "D"}

# Annual section rows  (cols B=2027F, C=2028F, D=2029F)
A_ANN_GROWTH_R = 17
A_ANN_GM_R     = 18
A_ANN_OPEX_R   = 19
A_ANN_NONOP_R  = 20
A_ANN_TAX_R    = 21
A_ANN_CAP_R    = 22
A_ANN_COL      = {"2027F": "B", "2028F": "C", "2029F": "D"}

# Section C: Valuation assumptions (single-value cells, col B)
A_BAND_LOOKBACK_R = 28   # PE/PB band lookback window, years (bake-time only)
A_PAYOUT_R        = 29   # dividend payout ratio assumption (drives PB Band forward BVPS, live)

# ── Styles ────────────────────────────────────────────────────────────────────
_FN = "Arial Narrow"

def _f(bold: bool = False, size: int = 9, color: str = "000000") -> Font:
    return Font(name=_FN, bold=bold, size=size, color=color)

FILL_WINE      = PatternFill("solid", fgColor="7A0000")
FILL_CREAM     = PatternFill("solid", fgColor="FAF9F4")
FILL_HIST_ANN  = PatternFill("solid", fgColor="D4CCBA")
FILL_FCAST_ANN = PatternFill("solid", fgColor="E8DFC8")
FILL_FCAST_QTR = PatternFill("solid", fgColor="F5EAD0")
FILL_INPUT     = PatternFill("solid", fgColor="FFF5E0")
FILL_SUB_HDR   = PatternFill("solid", fgColor="B0998F")

ALIGN_C = Alignment(horizontal="center", vertical="center")
ALIGN_L = Alignment(horizontal="left",   vertical="center")
ALIGN_R = Alignment(horizontal="right",  vertical="center")
ALIGN_L_WRAP = Alignment(horizontal="left", vertical="center", wrap_text=True)

_th  = Side(style="thin",   color="C9C0B0")
_tk  = Side(style="medium", color="7A0000")
_hr  = Side(style="hair",   color="C9C0B0")

BORDER_D  = Border(bottom=_th, top=_hr,  left=_hr,  right=_hr)
BORDER_HL = Border(bottom=_th, top=_tk,  left=_hr,  right=_hr)

FMT_100M = "#,##0.0"
FMT_EPS  = "0.00"
FMT_PCT  = "0.0%"
FMT_PRICE = "0.0"
FMT_MULT  = "0.00\"x\""
FMT_DAYS  = "0.0"
FMT_YEAR  = "0"

# Wine-family palette shared by every chart in the workbook.
PALETTE       = ["7A0000", "A03333", "C06666", "D49999"]
ACCENT_PRICE  = "1A1A1A"   # near-black — reserved for "actual price" lines only

# ── Col fill by type ──────────────────────────────────────────────────────────
def _cfill(ctype: str) -> PatternFill:
    if ctype == "annual_hist":  return FILL_HIST_ANN
    if ctype == "annual_fcast": return FILL_FCAST_ANN
    if ctype == "qtr_fcast":    return FILL_FCAST_QTR
    return FILL_CREAM

# ── Data fetching ─────────────────────────────────────────────────────────────
def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser()
    p.add_argument("--stocks", nargs="+", required=True)
    p.add_argument("--start_year", type=int, default=2022)
    return p.parse_args()


def fetch_finmind(stock_id: str, dataset: str, start_year: int) -> pd.DataFrame:
    r = requests.get(FINMIND_URL,
                     params={"dataset": dataset, "data_id": stock_id,
                             "start_date": f"{start_year}-01-01"},
                     timeout=30)
    r.raise_for_status()
    payload = r.json()
    if payload.get("status") != 200:
        raise RuntimeError(f"FinMind error: {payload.get('msg')}")
    rows = payload.get("data", [])
    if not rows:
        raise RuntimeError(f"No data: {stock_id}/{dataset}")
    return pd.DataFrame(rows)


def fetch_finmind_optional(stock_id: str, dataset: str, start_year: int) -> pd.DataFrame | None:
    """Like fetch_finmind, but degrades gracefully: prints a warning and returns
    None instead of raising, so one missing valuation dataset never aborts the
    whole stock (Income Model / Dashboard still build fully)."""
    try:
        return fetch_finmind(stock_id, dataset, start_year)
    except (RuntimeError, requests.RequestException) as e:
        print(f"  WARNING [{stock_id}]: {dataset} unavailable ({e}) — related sheet(s) will be left blank.")
        return None


_MQ = {"03": "Q1", "06": "Q2", "09": "Q3", "12": "Q4"}


def _qkey(date_str: str) -> str | None:
    parts = str(date_str).split("-")
    if len(parts) < 2:
        return None
    y, m = parts[0], parts[1]
    q = _MQ.get(m)
    return f"{y[2:4]}{q}" if q else None


def build_data(income_df: pd.DataFrame, bs_df: pd.DataFrame) -> dict[str, dict[str, float]]:
    result: dict[str, dict[str, float]] = {}
    income_types = {t for t, _, bs in INCOME_ITEMS if not bs}
    for df, wanted in [(income_df, income_types), (bs_df, BS_TYPES)]:
        if df is None or df.empty:
            continue
        for _, row in df[df["type"].isin(wanted)].iterrows():
            qk = _qkey(str(row["date"]))
            if qk:
                result.setdefault(str(row["type"]), {})[qk] = float(row["value"])
    return result


# Balance-sheet items needed only for valuation sheets (working capital + book value).
# NOTE: TaiwanStockBalanceSheet reuses the type name "EquityAttributableToOwnersOfParent"
# for total equity attributable to parent (a balance figure) — the *same* type name is
# also used by TaiwanStockFinancialStatements for quarterly net income attributable to
# parent (a flow figure). build_data() already consumes the income-statement version
# under that key, so this dict renames the balance-sheet version to "BS_Equity" to avoid
# silently overwriting it.
BS_EXTRA_TYPES = {
    "Inventories":                        "Inventories",
    "AccountsReceivableNet":              "AccountsReceivable",
    "AccountsPayable":                    "AccountsPayable",
    "EquityAttributableToOwnersOfParent": "BS_Equity",
}


def build_bs_extra_data(bs_df: pd.DataFrame | None) -> dict[str, dict[str, float]]:
    """Quarterly working-capital + book-value balances, in 億元, keyed by our own
    internal names (see BS_EXTRA_TYPES) to avoid the FinMind name collision above."""
    result: dict[str, dict[str, float]] = {}
    if bs_df is None or bs_df.empty:
        return result
    wanted = set(BS_EXTRA_TYPES)
    for _, row in bs_df[bs_df["type"].isin(wanted)].iterrows():
        qk = _qkey(str(row["date"]))
        if qk:
            key = BS_EXTRA_TYPES[str(row["type"])]
            result.setdefault(key, {})[qk] = float(row["value"]) / UNIT
    return result


# ── Market data helpers (price / PER / PBR / dividends) ──────────────────────
_QEND_MD = {1: (3, 31), 2: (6, 30), 3: (9, 30), 4: (12, 31)}


def _quarter_end_date(key: str) -> date:
    """'24Q4' / '26Q2F' -> the calendar last day of that quarter."""
    m = re.match(r"(\d{2})Q([1-4])F?$", key)
    if not m:
        raise ValueError(f"Not a quarter key: {key}")
    year = 2000 + int(m.group(1))
    month, day = _QEND_MD[int(m.group(2))]
    return date(year, month, day)


def _price_on_or_before(price_df: pd.DataFrame | None, target: date, lookback_days: int = 10) -> float | None:
    """Most recent trading-day close on or before `target` (handles holidays)."""
    if price_df is None or price_df.empty:
        return None
    d = price_df[["date", "close"]].copy()
    d["_d"] = pd.to_datetime(d["date"]).dt.date
    window = d[(d["_d"] <= target) & (d["_d"] >= target - timedelta(days=lookback_days))]
    if window.empty:
        return None
    return float(window.sort_values("_d").iloc[-1]["close"])


def _latest_price(price_df: pd.DataFrame | None) -> tuple[float, date] | None:
    if price_df is None or price_df.empty:
        return None
    d = price_df.sort_values("date").iloc[-1]
    return float(d["close"]), pd.to_datetime(d["date"]).date()


def _band_stats(ratio_series: "pd.Series", lookback_years: int) -> dict | None:
    """Mean/SD of a daily PER or PBR series over the trailing `lookback_years`,
    with a light outlier trim (drop values outside 3x of the median) so a single
    earnings-collapse quarter doesn't blow the band out — standard practice for
    forward-multiple band charts."""
    s = pd.to_numeric(ratio_series, errors="coerce").dropna()
    s = s[s > 0]
    if s.empty:
        return None
    med = float(s.median())
    s = s[(s >= med / 3) & (s <= med * 3)]
    if s.empty:
        return None
    return {
        "mean": float(s.mean()),
        "sd":   float(s.std(ddof=1)) if len(s) > 1 else 0.0,
        "n":    int(len(s)),
    }


def _unit(tc: str, v: float) -> float:
    return v if tc == EPS_TYPE else v / UNIT


def _annual(tc: str, data: dict, yr: str) -> float | None:
    td = data.get(tc, {})
    vals = [td.get(f"{yr}Q{q}") for q in range(1, 5)]
    if all(v is None for v in vals):
        return None
    if tc in BS_TYPES:
        v = td.get(f"{yr}Q4")
        return None if v is None else _unit(tc, v)
    return _unit(tc, sum(v for v in vals if v is not None))


def cell_val(tc: str, key: str, data: dict) -> float | None:
    if key == "2024":
        return _annual(tc, data, "24")
    if key == "2025":
        return _annual(tc, data, "25")
    raw = data.get(tc, {}).get(key)
    return None if raw is None else _unit(tc, raw)

# ── Formula helpers ───────────────────────────────────────────────────────────
def _ie(f: str) -> str:
    return f'=IFERROR({f},"")'


def _a_q(col: str, row: int) -> str:        # Assumptions quarterly cell
    return f"Assumptions!${col}${row}"


def _a_ann(col: str, row: int) -> str:      # Assumptions annual cell
    return f"Assumptions!${col}${row}"


def _qtr_fcast_formula(tc: str, cl: str, ir: dict, ac: str, key: str) -> str:
    """Formula for 26Q2F / 26Q3F / 26Q4F column cells."""
    rev = ir["Revenue"]; gp = ir["GrossProfit"]; opex = ir["OperatingExpenses"]
    oi  = ir["OperatingIncome"]; nonop = ir["TotalNonoperatingIncomeAndExpense"]
    pt  = ir["PreTaxIncome"]; tax = ir["TAX"]; ni = ir["IncomeFromContinuingOperations"]
    nip = ir["EquityAttributableToOwnersOfParent"]; cap = ir["OrdinaryShare"]
    a = lambda r: _a_q(ac, r)
    c = lambda r: f"{cl}{r}"
    if tc == "Revenue":
        return f"=IFERROR('{REVENUE_BUILD_SHEET}'!{RB_COL_LTR[key]}{RB_FORECAST_REVENUE_ROW},\"\")"
    elif tc == "GrossProfit":
        return _ie(f"{c(rev)}*{a(A_Q_GM_R)}")
    elif tc == "CostOfGoodsSold":
        return _ie(f"{c(rev)}-{c(gp)}")
    elif tc == "OperatingExpenses":
        return _ie(f"{c(rev)}*{a(A_Q_OPEX_R)}")
    elif tc == "OperatingIncome":
        return _ie(f"{c(gp)}-{c(opex)}")
    elif tc == "TotalNonoperatingIncomeAndExpense":
        return f"={a(A_Q_NONOP_R)}"
    elif tc == "PreTaxIncome":
        return _ie(f"{c(oi)}+{c(nonop)}")
    elif tc == "TAX":
        return _ie(f"{c(pt)}*{a(A_Q_TAX_R)}")
    elif tc == "IncomeFromContinuingOperations":
        return _ie(f"{c(pt)}-{c(tax)}")
    elif tc == "EquityAttributableToOwnersOfParent":
        return _ie(f"{c(ni)}")
    elif tc == "OrdinaryShare":
        return f"={a(A_Q_CAP_R)}"
    elif tc == "EPS":
        return _ie(f"{c(nip)}/{c(cap)}*10")
    return ""


def _ann_fcast_formula(tc: str, cl: str, pcl: str, ir: dict, ac: str, key: str) -> str:
    """Formula for 2027F / 2028F / 2029F column cells."""
    rev = ir["Revenue"]; gp = ir["GrossProfit"]; opex = ir["OperatingExpenses"]
    oi  = ir["OperatingIncome"]; nonop = ir["TotalNonoperatingIncomeAndExpense"]
    pt  = ir["PreTaxIncome"]; tax = ir["TAX"]; ni = ir["IncomeFromContinuingOperations"]
    nip = ir["EquityAttributableToOwnersOfParent"]; cap = ir["OrdinaryShare"]
    a = lambda r: _a_ann(ac, r)
    c = lambda r: f"{cl}{r}"
    p = lambda r: f"{pcl}{r}"
    if tc == "Revenue":
        return f"=IFERROR('{REVENUE_BUILD_SHEET}'!{RB_COL_LTR[key]}{RB_FORECAST_REVENUE_ROW},\"\")"
    elif tc == "GrossProfit":
        return _ie(f"{c(rev)}*{a(A_ANN_GM_R)}")
    elif tc == "CostOfGoodsSold":
        return _ie(f"{c(rev)}-{c(gp)}")
    elif tc == "OperatingExpenses":
        return _ie(f"{c(rev)}*{a(A_ANN_OPEX_R)}")
    elif tc == "OperatingIncome":
        return _ie(f"{c(gp)}-{c(opex)}")
    elif tc == "TotalNonoperatingIncomeAndExpense":
        return f"={a(A_ANN_NONOP_R)}"
    elif tc == "PreTaxIncome":
        return _ie(f"{c(oi)}+{c(nonop)}")
    elif tc == "TAX":
        return _ie(f"{c(pt)}*{a(A_ANN_TAX_R)}")
    elif tc == "IncomeFromContinuingOperations":
        return _ie(f"{c(pt)}-{c(tax)}")
    elif tc == "EquityAttributableToOwnersOfParent":
        return _ie(f"{c(ni)}")
    elif tc == "OrdinaryShare":
        return f"={a(A_ANN_CAP_R)}"
    elif tc == "EPS":
        return _ie(f"{c(nip)}/{c(cap)}*10")
    return ""


def _2026f_formula(tc: str, row: int, ir: dict) -> str:
    """Sum formula for 2026(F) annual column."""
    L, M, N, O = COL_LTR["26Q1"], COL_LTR["26Q2F"], COL_LTR["26Q3F"], COL_LTR["26Q4F"]
    P = COL_LTR["2026F"]
    if tc in BS_TYPES:
        return f"={O}{row}"
    if tc == EPS_TYPE:
        nip = ir["EquityAttributableToOwnersOfParent"]
        cap = ir["OrdinaryShare"]
        return _ie(f"{P}{nip}/{P}{cap}*10")
    return _ie(f"{L}{row}+{M}{row}+{N}{row}+{O}{row}")


# ── QoQ / YoY ────────────────────────────────────────────────────────────────
_QOQ_PREV = {
    "25Q1": "24Q4", "25Q2": "25Q1", "25Q3": "25Q2", "25Q4": "25Q3",
    "26Q1": "25Q4", "26Q2F": "26Q1", "26Q3F": "26Q2F", "26Q4F": "26Q3F",
}

_YOY_PREV = {
    "25Q1": "24Q1", "25Q2": "24Q2", "25Q3": "24Q3", "25Q4": "24Q4",
    "2025": "2024",
    "26Q1": "25Q1", "26Q2F": "25Q2", "26Q3F": "25Q3", "26Q4F": "25Q4",
    "2026F": "2025", "2027F": "2026F", "2028F": "2027F", "2029F": "2028F",
}


def _pct(cl: str, cr: int, pl: str, pr: int) -> str:
    return f"=IFERROR(({cl}{cr}-{pl}{pr})/ABS({pl}{pr}),\"\")"


def f_qoq(key: str, src_row: int) -> str | None:
    prev = _QOQ_PREV.get(key)
    return None if prev is None else _pct(COL_LTR[key], src_row, COL_LTR[prev], src_row)


def f_yoy(key: str, src_row: int) -> str | None:
    prev = _YOY_PREV.get(key)
    return None if prev is None else _pct(COL_LTR[key], src_row, COL_LTR[prev], src_row)


def _ntm_formula(key: str, eps_row: int, data: dict | None = None) -> str:
    """12-month-forward EPS as of `key`'s quarter-end: sum of the next 4 quarterly
    EPS cells in Income Model. Where fewer than 4 explicit quarterly columns remain
    (only true for 26Q1 and the 3 forecast quarters), the missing quarters are
    proxied at PROXY_ANNUAL_KEY/4 each — the standard way to extend a quarterly
    model past its explicit forecast horizon.

    Income Model's 26Q2F/26Q3F/26Q4F columns are a fixed template, not date-aware —
    they keep running the forecast formula even after FinMind has real reported
    results for that quarter (this happens routinely, since a quarter is usually
    reported within ~1.5 quarters of its end). When `data` is supplied and FinMind
    already has an actual EPS for one of the `needed` quarters, that real value is
    used directly instead of referencing the (by then stale) forecast cell — this
    is what keeps the PE Band's historical window from drifting away from reality."""
    i = QTR_KEYS.index(key)
    needed = QTR_KEYS[i + 1: i + 5]
    n_missing = 4 - len(needed)
    terms = []
    for k in needed:
        actual = data.get("EPS", {}).get(k[:-1]) if (data and k.endswith("F")) else None
        terms.append(f"({actual!r})" if actual is not None else f"'{INCOME_MODEL_SHEET}'!{COL_LTR[k]}{eps_row}")
    if n_missing:
        proxy_col = COL_LTR[PROXY_ANNUAL_KEY]
        terms.append(f"({n_missing}*'{INCOME_MODEL_SHEET}'!{proxy_col}{eps_row}/4)")
    return "=IFERROR(" + "+".join(terms) + ',"")'

# ── Section header helper ─────────────────────────────────────────────────────
def _sec(ws, row: int, text: str, fill: PatternFill = None, end_col: int = None) -> None:
    ec = end_col or LAST_COL
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ec)
    c = ws.cell(row, 1, text)
    c.font   = _f(bold=True, size=9, color="FFFFFF")
    c.fill   = fill or FILL_WINE
    c.alignment = ALIGN_L
    ws.row_dimensions[row].height = 16

# ── Shared chart styling ──────────────────────────────────────────────────────
# Column-width -> cm and row-height -> cm conversions (Calibri 11 default digit
# width), used to size/space charts so they provably don't overlap without a
# visual renderer available in this environment:
#   px  = width * 7 + 5                (Excel default-font column width formula)
#   cm  = px * 2.54 / 96
#   row_cm = row_height_pt * 0.03528
CHART_H = 9.0     # cm
CHART_W = 14.5    # cm
CHART_ROW_GAP = 20   # rows between chart-block anchors; 20 * 15pt row = 10.58cm > 9.0cm chart height

def _chart_title(text: str, size: int = 1600, bold: bool = True, color: str = "7A0000") -> Title:
    cp  = CharacterProperties(latin=DrawFont(typeface=_FN), sz=size, b=bold, solidFill=color)
    run = RegularTextRun(rPr=cp, t=text)
    para = Paragraph(pPr=ParagraphProperties(defRPr=cp), r=[run])
    t = Title(tx=Text(rich=RichText(p=[para])))
    t.overlay = True   # float over the plot area instead of reserving its own band —
    return t            # paired with the pinned plot-area rect below, this is what stops
                         # the title from ever colliding with the plot as text length varies.


# Manual layout rectangles (fractions of the chart canvas). The title uses overlay
# (floats over the plot instead of reserving its own band) paired with a pinned plot
# rect — that combination is what stops the title from ever colliding with the plot
# as text length varies, since Excel no longer has to *guess* how much room to leave.
# The legend, by contrast, gets its own RESERVED column on the right (not overlaid) —
# an earlier attempt at also floating the legend in a fixed corner looked fine on the
# one reference chart it was copied from, but overlapped data on charts with a
# different shape (a fixed corner isn't safe for a legend across 16 different charts
# with different data shapes; a dedicated side column is). Plot width narrows when a
# legend is present to make room for that column.
_PLOT_Y, _PLOT_H = 0.08, 0.84           # ~8% top margin for the floating title, ~8% bottom for x-axis labels
_PLOT_X_NARROW, _PLOT_W_NARROW = 0.07, 0.66   # when a legend column is present
_PLOT_X_WIDE, _PLOT_W_WIDE     = 0.07, 0.90   # single-series charts, no legend
_LEGEND_X, _LEGEND_W = 0.76, 0.22
_LEGEND_Y, _LEGEND_H = 0.10, 0.80
CHART_BG = "FAF9F4"   # same cream as the sheet's own cell background — the chart
                       # canvas should blend into the page, not sit on it as a white box.


def _plot_border() -> LineProperties:
    """Plot area border: a soft tan-to-wine gradient instead of a flat solid line —
    a flat border was the harsh edge the title's overlay was visually 'colliding'
    with; a gradient reads as a frame, not a hard stop."""
    grad = GradientFillProperties(
        gsLst=[GradientStop(pos=0, srgbClr="D9D2C2"), GradientStop(pos=100000, srgbClr="7A0000")],
        lin=LinearShadeProperties(ang=0, scaled=True),
    )
    return LineProperties(gradFill=grad, w=9525)


def _manual_plot_layout(legend: bool) -> Layout:
    x, w = (_PLOT_X_NARROW, _PLOT_W_NARROW) if legend else (_PLOT_X_WIDE, _PLOT_W_WIDE)
    return Layout(manualLayout=ManualLayout(layoutTarget="inner", xMode="edge", yMode="edge",
                                             wMode="factor", hMode="factor",
                                             x=x, y=_PLOT_Y, w=w, h=_PLOT_H))


def _manual_legend_layout() -> Layout:
    return Layout(manualLayout=ManualLayout(xMode="edge", yMode="edge", wMode="factor", hMode="factor",
                                             x=_LEGEND_X, y=_LEGEND_Y, w=_LEGEND_W, h=_LEGEND_H))


def _axis_text_props(size: int = 800, color: str = "555046") -> RichText:
    cp = CharacterProperties(latin=DrawFont(typeface=_FN), sz=size, solidFill=color)
    para = Paragraph(pPr=ParagraphProperties(defRPr=cp), endParaRPr=cp)
    return RichText(p=[para])


def _style_chart(ch, title: str, y_fmt: str, legend: bool = False) -> None:
    """Apply one consistent, non-overlapping look to every chart in the workbook —
    pinned plot-area rect + an overlaying title (never collides regardless of title
    length), and a legend in its own reserved side column (never overlaps data
    regardless of data shape). Gridlines are left as a bare <majorGridlines/> element
    so Excel renders them using the chart style's own default treatment (style=10's
    built-in soft gradient) instead of a hand-picked flat color — that built-in look
    is what "有質感" was asking for; a previous flat solid-color override was masking it."""
    ch.title  = _chart_title(title)
    ch.style  = 10
    ch.height = CHART_H
    ch.width  = CHART_W
    ch.y_axis.numFmt = y_fmt
    ch.y_axis.delete = False
    ch.x_axis.delete = False
    try:
        ch.graphical_properties = GraphicalProperties(solidFill=CHART_BG)
        ch.plot_area.spPr = GraphicalProperties(solidFill=CHART_BG, ln=_plot_border())
    except Exception:
        pass
    try:
        ch.y_axis.majorGridlines = ChartLines()   # bare -> inherits the style's own gridline look
        ch.x_axis.majorGridlines = None
        ch.y_axis.txPr = _axis_text_props()
        ch.x_axis.txPr = _axis_text_props()
    except Exception:
        pass
    try:
        # NOTE: must be set on ch.layout, not ch.plot_area.layout — openpyxl's
        # ChartBase._write() does `self.plot_area.layout = self.layout` at save
        # time, silently discarding anything assigned directly to plot_area.layout.
        ch.layout = _manual_plot_layout(legend)
    except Exception:
        pass
    if legend:
        ch.legend.position = "r"
        ch.legend.overlay  = False
        ch.legend.layout   = _manual_legend_layout()
        try:
            ch.legend.txPr = _axis_text_props(size=800, color="333333")
        except Exception:
            pass
    else:
        ch.legend = None


def _style_series_line(s, color: str, width_emu: int = 22000, dash: str | None = None) -> None:
    try:
        s.smooth = False
        s.marker.symbol = "none"
        s.graphicalProperties.line.solidFill = color
        s.graphicalProperties.line.width = width_emu
        if dash:
            s.graphicalProperties.line.dashStyle = dash
    except Exception:
        pass


def _style_series_bar(s, color: str) -> None:
    try:
        s.graphicalProperties.solidFill = color
        s.graphicalProperties.line.solidFill = color
    except Exception:
        pass

# ── Build Income Model ────────────────────────────────────────────────────────
def build_model_sheet(wb: Workbook, stock_id: str, data: dict) -> dict:
    ws = wb.active
    ws.title = INCOME_MODEL_SHEET

    ir  = {tc: DATA_START + i for i, (tc, _, _) in enumerate(INCOME_ITEMS)}
    mr  = {ml: MARGIN_START + i for i, (ml, _, _) in enumerate(MARGIN_ITEMS)}
    yir = {tc: YOY_START + i for i, (tc, _, _) in enumerate(INCOME_ITEMS)}

    ann_prev_key = {"2027F": "2026F", "2028F": "2027F", "2029F": "2028F"}

    # Cream background sweep
    for r in range(1, MARGIN_END + 5):
        for ci in range(1, LAST_COL + 1):
            ws.cell(r, ci).fill = FILL_CREAM

    # Title
    ws.merge_cells(start_row=TITLE_ROW, start_column=1, end_row=TITLE_ROW, end_column=LAST_COL)
    t = ws.cell(TITLE_ROW, 1, f"{stock_id}  季度損益表   單位：新台幣億元")
    t.font = _f(bold=True, size=12, color="7A0000"); t.alignment = ALIGN_C
    ws.row_dimensions[TITLE_ROW].height = 22

    # Header row
    h = ws.cell(HDR_ROW, LABEL_COL, "損益項目")
    h.font = _f(bold=True, size=9, color="FFFFFF"); h.fill = FILL_WINE; h.alignment = ALIGN_C
    for hdr_txt, key, ctype in COLUMNS:
        ci = COL_IDX[key]
        c  = ws.cell(HDR_ROW, ci, hdr_txt)
        c.alignment = ALIGN_C
        if ctype in ("annual_hist", "annual_fcast"):
            c.font = _f(bold=True, size=9, color="FFFFFF"); c.fill = FILL_WINE
        elif ctype == "qtr_fcast":
            c.font = _f(bold=True, size=9); c.fill = FILL_FCAST_QTR
        else:
            c.font = _f(bold=False, size=9); c.fill = _cfill(ctype)
    ws.row_dimensions[HDR_ROW].height = 17

    # ── Data rows ─────────────────────────────────────────────────────────────
    for i, (tc, label, _) in enumerate(INCOME_ITEMS):
        row   = DATA_START + i
        is_hl = tc in HIGHLIGHT_TYPES
        fmt   = FMT_EPS if tc == EPS_TYPE else FMT_100M
        border = BORDER_HL if is_hl else BORDER_D

        lc = ws.cell(row, LABEL_COL, label)
        lc.font = _f(bold=is_hl, size=9); lc.fill = FILL_CREAM
        lc.alignment = ALIGN_L; lc.border = border
        ws.row_dimensions[row].height = 15

        for _, key, ctype in COLUMNS:
            ci   = COL_IDX[key]
            cell = ws.cell(row, ci)
            cell.number_format = fmt
            cell.alignment     = ALIGN_R
            cell.fill          = _cfill(ctype)
            is_ann = key in ANNUAL_KEYS
            cell.font   = _f(bold=(is_hl or is_ann), size=9)
            cell.border = border

            if ctype == "qtr_fcast":
                cell.value = _qtr_fcast_formula(tc, COL_LTR[key], ir, A_Q_COL[key], key)
            elif ctype == "annual_fcast":
                if key == "2026F":
                    cell.value = _2026f_formula(tc, row, ir)
                else:
                    pk  = ann_prev_key[key]
                    cell.value = _ann_fcast_formula(
                        tc, COL_LTR[key], COL_LTR[pk], ir, A_ANN_COL[key], key
                    )
            else:
                cell.value = cell_val(tc, key, data)

    # ── QoQ ───────────────────────────────────────────────────────────────────
    _sec(ws, QOQ_HDR, "季度成長分析 QoQ %", FILL_SUB_HDR)
    for i, (tc, label, _) in enumerate(INCOME_ITEMS):
        row    = QOQ_START + i
        src    = ir[tc]
        is_hl  = tc in HIGHLIGHT_TYPES
        border = BORDER_HL if is_hl else BORDER_D
        lc = ws.cell(row, LABEL_COL, label)
        lc.font = _f(bold=is_hl, size=9); lc.fill = FILL_CREAM
        lc.alignment = ALIGN_L; lc.border = border
        ws.row_dimensions[row].height = 15
        for _, key, ctype in COLUMNS:
            ci = COL_IDX[key]
            c  = ws.cell(row, ci)
            c.number_format = FMT_PCT; c.alignment = ALIGN_R
            c.fill = _cfill(ctype); c.font = _f(size=9); c.border = border
            f = f_qoq(key, src)
            if f: c.value = f

    # ── YoY ───────────────────────────────────────────────────────────────────
    _sec(ws, YOY_HDR, "年成長分析 YoY %", FILL_SUB_HDR)
    for i, (tc, label, _) in enumerate(INCOME_ITEMS):
        row    = YOY_START + i
        src    = ir[tc]
        is_hl  = tc in HIGHLIGHT_TYPES
        border = BORDER_HL if is_hl else BORDER_D
        lc = ws.cell(row, LABEL_COL, label)
        lc.font = _f(bold=is_hl, size=9); lc.fill = FILL_CREAM
        lc.alignment = ALIGN_L; lc.border = border
        ws.row_dimensions[row].height = 15
        for _, key, ctype in COLUMNS:
            ci = COL_IDX[key]
            c  = ws.cell(row, ci)
            c.number_format = FMT_PCT; c.alignment = ALIGN_R
            c.fill = _cfill(ctype); c.font = _f(size=9); c.border = border
            f = f_yoy(key, src)
            if f: c.value = f

    # ── Margin ────────────────────────────────────────────────────────────────
    _sec(ws, MARGIN_HDR, "獲利分析 %", FILL_SUB_HDR)
    for i, (ml, num_tc, den_tc) in enumerate(MARGIN_ITEMS):
        row    = MARGIN_START + i
        is_hl  = ml in HIGHLIGHT_MARGINS
        border = BORDER_HL if is_hl else BORDER_D
        lc = ws.cell(row, LABEL_COL, ml)
        lc.font = _f(bold=is_hl, size=9); lc.fill = FILL_CREAM
        lc.alignment = ALIGN_L; lc.border = border
        ws.row_dimensions[row].height = 15
        num_r = ir[num_tc]; den_r = ir[den_tc]
        for _, key, ctype in COLUMNS:
            ci = COL_IDX[key]
            c  = ws.cell(row, ci)
            c.number_format = FMT_PCT; c.alignment = ALIGN_R
            c.fill = _cfill(ctype)
            is_ann = key in ANNUAL_KEYS
            c.font = _f(bold=(is_hl or is_ann), size=9); c.border = border
            ltr = COL_LTR[key]
            c.value = f'=IFERROR({ltr}{num_r}/{ltr}{den_r},"")'

    # ── Column widths / freeze ────────────────────────────────────────────────
    ws.column_dimensions[get_column_letter(LABEL_COL)].width = 26
    for _, key, ctype in COLUMNS:
        cd = ws.column_dimensions[get_column_letter(COL_IDX[key])]
        if ctype == "qtr_helper":
            cd.hidden = True
        else:
            cd.width = 10
    ws.freeze_panes = ws.cell(HDR_ROW + 1, LABEL_COL + 1)

    return {"item_row": ir, "margin_row": mr, "yoy_item_row": yir}

# ── Build Assumptions ─────────────────────────────────────────────────────────
def build_assumptions_sheet(wb: Workbook, stock_id: str, data: dict,
                             payout_def: float | None, band_lookback_years: int) -> None:
    ws = wb.create_sheet(ASSUMPTIONS_SHEET)
    ws.sheet_properties.tabColor = "7A0000"

    for r in range(1, 32):
        for c in range(1, 8):
            ws.cell(r, c).fill = FILL_CREAM

    def _lbl(row: int, text: str) -> None:
        c = ws.cell(row, 1, text)
        c.font = _f(size=9); c.fill = FILL_CREAM
        c.alignment = ALIGN_L; c.border = BORDER_D
        ws.row_dimensions[row].height = 15

    def _hdr(row: int, labels: list, start_col: int = 1) -> None:
        for j, lbl in enumerate(labels, start=start_col):
            c = ws.cell(row, j, lbl)
            c.font = _f(bold=True, size=9, color="FFFFFF")
            c.fill = FILL_WINE; c.alignment = ALIGN_C; c.border = BORDER_D
        ws.row_dimensions[row].height = 16

    def _inp(row: int, col: int, value=None, fmt: str = FMT_100M) -> None:
        c = ws.cell(row, col, value)
        c.font = _f(size=9); c.fill = FILL_INPUT
        c.alignment = ALIGN_R; c.border = BORDER_D
        c.number_format = fmt

    # ── Compute smart defaults from historical data ──────────────────────────
    def _ratio(num_tc: str, den_tc: str) -> float | None:
        nd, dd = data.get(num_tc, {}), data.get(den_tc, {})
        for qk in ["26Q1", "25Q4", "25Q3", "25Q2"]:
            n, d = nd.get(qk), dd.get(qk)
            if n and d and d != 0:
                return round(n / d, 4)
        return None

    def _latest(tc: str) -> float | None:
        td = data.get(tc, {})
        for qk in ["26Q1", "25Q4", "25Q3", "25Q2"]:
            v = td.get(qk)
            if v is not None:
                return round(v / UNIT, 1) if tc != EPS_TYPE else v
        return None

    gm_def   = _ratio("GrossProfit", "Revenue")
    opex_def = _ratio("OperatingExpenses", "Revenue")
    tax_def  = _ratio("TAX", "PreTaxIncome")
    cap_def  = _latest("OrdinaryShare")

    # Title
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=7)
    t = ws.cell(1, 1, f"{stock_id}  預測假設   請編輯黃色格子")
    t.font = _f(bold=True, size=12, color="7A0000"); t.alignment = ALIGN_C
    ws.row_dimensions[1].height = 22

    ws.row_dimensions[2].height = 6   # spacer

    # ── Section A: Quarterly 26Q2F–Q4F (rows 3–11) ──────────────────────────
    _sec(ws, 3, "短期季度假設 (2026 Q2 ~ Q4)  — 公式將從此處引用", end_col=7)
    _hdr(4, ["假設項目", "26Q2(F)", "26Q3(F)", "26Q4(F)"])

    # 營業收入淨額 (Revenue) is intentionally NOT in this list — Revenue Build is now
    # the single source of truth for revenue, so its Assumptions row is built as a
    # read-only formula link below instead of a free-input cell (see A_Q_REV_R block).
    q_params = [
        (A_Q_GM_R,    "毛利率",                       gm_def,   FMT_PCT),
        (A_Q_OPEX_R,  "營業費用率",                   opex_def, FMT_PCT),
        (A_Q_NONOP_R, "營業外收入及支出合計 (億元)",  0,        FMT_100M),
        (A_Q_TAX_R,   "稅率",                         tax_def,  FMT_PCT),
        (A_Q_CAP_R,   "普通股本 (億元)",              cap_def,  FMT_100M),
    ]
    for row, label, default, fmt in q_params:
        _lbl(row, label)
        for col_offset in range(2, 5):   # B, C, D
            _inp(row, col_offset, default, fmt)

    # 營業收入淨額 — read-only, linked to Revenue Build (see build_revenue_build_sheet)
    rev_lbl = ws.cell(A_Q_REV_R, 1, "營業收入淨額 (億元) — 由 Revenue Build 計算")
    rev_lbl.font = _f(size=9); rev_lbl.fill = FILL_CREAM
    rev_lbl.alignment = ALIGN_L; rev_lbl.border = BORDER_D
    ws.row_dimensions[A_Q_REV_R].height = 15
    for key, col in A_Q_COL.items():
        c = ws[f"{col}{A_Q_REV_R}"]
        c.value = f"=IFERROR('{REVENUE_BUILD_SHEET}'!{RB_COL_LTR[key]}{RB_FORECAST_REVENUE_ROW},\"\")"
        c.font = _f(size=9); c.fill = FILL_CREAM
        c.alignment = ALIGN_R; c.border = BORDER_D; c.number_format = FMT_100M

    ws.row_dimensions[11].height = 6    # spacer

    # note
    ws.merge_cells(start_row=12, start_column=1, end_row=12, end_column=7)
    n = ws.cell(12, 1, "Income Model 26Q2~Q4 各損益項目均引用上表假設推算。請直接在黃色格子輸入數值。")
    n.font = _f(size=8, color="666666"); n.fill = FILL_CREAM
    ws.row_dimensions[12].height = 13

    ws.merge_cells(start_row=13, start_column=1, end_row=13, end_column=7)
    n_rev = ws.cell(13, 1,
        "營收成長率由 Revenue Build 中出貨量、平均售價、產品組合、匯率等營運驅動因子推導；"
        "如需人工覆蓋，請於 Revenue Build 的「人工覆蓋」欄位輸入。")
    n_rev.font = _f(size=8, color="7A0000"); n_rev.fill = FILL_CREAM
    ws.row_dimensions[13].height = 26

    # ── Section B: Annual 2027F–2029F (rows 14–22) ──────────────────────────
    _sec(ws, 14, "年度長期假設 (2027F ~ 2029F)  — 公式將從此處引用", end_col=7)
    _hdr(15, ["假設項目", "2027(F)", "2028(F)", "2029(F)"])

    # 營收成長率 YoY is intentionally NOT in this list — same reason as above.
    ann_params = [
        (A_ANN_GM_R,     "毛利率",                          gm_def,   FMT_PCT),
        (A_ANN_OPEX_R,   "營業費用率",                      opex_def, FMT_PCT),
        (A_ANN_NONOP_R,  "營業外收入及支出合計 (億元)",     0,        FMT_100M),
        (A_ANN_TAX_R,    "稅率",                            tax_def,  FMT_PCT),
        (A_ANN_CAP_R,    "普通股本 (億元)",                 cap_def,  FMT_100M),
    ]
    for row, label, default, fmt in ann_params:
        _lbl(row, label)
        for col_offset in range(2, 5):
            _inp(row, col_offset, default, fmt)

    # 營收成長率 YoY — read-only, linked to Revenue Build
    growth_lbl = ws.cell(A_ANN_GROWTH_R, 1, "營收成長率 YoY（由 Revenue Build 計算）")
    growth_lbl.font = _f(size=9); growth_lbl.fill = FILL_CREAM
    growth_lbl.alignment = ALIGN_L; growth_lbl.border = BORDER_D
    ws.row_dimensions[A_ANN_GROWTH_R].height = 15
    for key, col in A_ANN_COL.items():
        c = ws[f"{col}{A_ANN_GROWTH_R}"]
        c.value = f"=IFERROR('{REVENUE_BUILD_SHEET}'!{RB_COL_LTR[key]}{RB_APPLIED_GROWTH_ROW},\"\")"
        c.font = _f(size=9); c.fill = FILL_CREAM
        c.alignment = ALIGN_R; c.border = BORDER_D; c.number_format = FMT_PCT

    ws.row_dimensions[23].height = 6

    # note
    ws.merge_cells(start_row=24, start_column=1, end_row=24, end_column=7)
    n2 = ws.cell(24, 1,
        "說明：修改後請儲存，Income Model 各損益科目與 Dashboard 圖表將自動更新。"
        "  預設值已參照最近一季歷史數據填入。")
    n2.font = _f(size=8, color="7A0000"); n2.fill = FILL_CREAM
    ws.row_dimensions[24].height = 13

    ws.row_dimensions[25].height = 6

    # ── Section C: Valuation assumptions (rows 26-30) ────────────────────────
    _sec(ws, 26, "估值假設 Valuation Assumptions", end_col=7)
    _lbl(A_BAND_LOOKBACK_R, "PE/PB 河流圖回溯年數 (年)")
    _inp(A_BAND_LOOKBACK_R, 2, band_lookback_years, FMT_YEAR)
    _lbl(A_PAYOUT_R, "股利發放率假設 (預測 BVPS 用)")
    _inp(A_PAYOUT_R, 2, payout_def, FMT_PCT)

    ws.merge_cells(start_row=30, start_column=1, end_row=30, end_column=7)
    n3 = ws.cell(30, 1,
        "說明：股利發放率會即時連動 PB Band 的預測 BVPS 公式，修改後儲存即可看到 PB Band 更新。"
        "回溯年數僅影響「產生報表當下」抓取的 PER/PBR 歷史資料範圍，修改後需重新執行程式才會生效（非即時連動）。"
        "發放率預設值 = 歷史各年度 (合計股利 / EPS) 之平均。")
    n3.font = _f(size=8, color="7A0000"); n3.fill = FILL_CREAM
    ws.row_dimensions[30].height = 26

    ws.column_dimensions["A"].width = 28
    for col in ["B", "C", "D", "E", "F", "G"]:
        ws.column_dimensions[col].width = 14

    ws.freeze_panes = "B5"

# ── Build Revenue Build ────────────────────────────────────────────────────────
def _rb_row_label(ws, row: int, text: str, bold: bool = False, fill: PatternFill = FILL_CREAM) -> None:
    # A leading '=', '+', '-', or '@' makes Excel/openpyxl treat a plain-text label as
    # a FORMULA instead of a string — Excel then fails to parse the text as a formula
    # and the cell shows broken/garbled content. Guard against it here.
    if text and text[0] in "=+-@":
        raise ValueError(f"Row {row} label starts with '{text[0]}' — Excel would read this "
                          f"as a formula, not text: {text!r}")
    c = ws.cell(row, 1, text)
    c.font = _f(bold=bold, size=9); c.fill = fill
    # Wrapped + a row tall enough for 2 lines: these bilingual labels are too long for
    # column A at some row widths, and since column B always has real content next to
    # them, Excel clips (rather than overflows) any label wider than the column — wrap
    # is the only way to guarantee the text is never cut off without hand-tuning column
    # widths per label, which is exactly what's gone wrong here more than once already.
    c.alignment = ALIGN_L_WRAP; c.border = BORDER_HL if bold else BORDER_D
    ws.row_dimensions[row].height = 28


def _rb_write_header(ws, row: int) -> None:
    h = ws.cell(row, 1, "項目 Item")
    h.font = _f(bold=True, size=9, color="FFFFFF"); h.fill = FILL_WINE
    h.alignment = ALIGN_C; h.border = BORDER_D
    for hdr_txt, key, ctype in RB_COLUMNS:
        ci = RB_COL_IDX[key]
        c = ws.cell(row, ci, hdr_txt)
        c.alignment = ALIGN_C; c.border = BORDER_D
        if ctype in ("annual_hist", "annual_fcast"):
            c.font = _f(bold=True, size=9, color="FFFFFF"); c.fill = FILL_WINE
        elif ctype == "qtr_fcast":
            c.font = _f(bold=True, size=9); c.fill = FILL_FCAST_QTR
        else:
            c.font = _f(size=9); c.fill = _cfill(ctype)
    ws.row_dimensions[row].height = 16


def _rb_data_cell(ws, row: int, key: str, value=None, fmt: str = FMT_100M,
                   fill: PatternFill = None, bold: bool = False, editable: bool = False) -> None:
    ci = RB_COL_IDX[key]
    c = ws.cell(row, ci, value)
    c.number_format = fmt; c.alignment = ALIGN_R
    c.font = _f(bold=bold, size=9)
    c.fill = fill if fill is not None else (FILL_INPUT if editable else FILL_CREAM)
    c.border = BORDER_HL if bold else BORDER_D


def _rb_dash_cell(ws, row: int, key: str) -> None:
    ci = RB_COL_IDX[key]
    c = ws.cell(row, ci, "-")
    c.font = _f(size=9, color="9A9186"); c.fill = FILL_CREAM
    c.alignment = ALIGN_C; c.border = BORDER_D


def _rb_na_cell(ws, row: int, key: str) -> None:
    ci = RB_COL_IDX[key]
    c = ws.cell(row, ci, "N/A")
    c.font = _f(size=9, color="9A9186"); c.fill = FILL_CREAM
    c.alignment = ALIGN_C; c.border = BORDER_D


def build_revenue_build_sheet(wb: Workbook, stock_id: str, ir: dict, data: dict) -> None:
    """The single source of truth for future revenue. Income Model's 26Q2F-26Q4F
    and 2027F-2029F Revenue cells reference this sheet's row RB_FORECAST_REVENUE_ROW
    (see _qtr_fcast_formula / _ann_fcast_formula), and Assumptions' revenue rows
    become read-only links into it (see build_assumptions_sheet)."""
    ws = wb.create_sheet(REVENUE_BUILD_SHEET, index=wb.sheetnames.index(ASSUMPTIONS_SHEET) + 1)
    ws.sheet_properties.tabColor = "7A0000"

    rev_row = ir["Revenue"]

    # Quarterly driver placeholder: most recent actual YoY (26Q1 vs 25Q1) — a
    # brand-new model has no prior revenue assumption to preserve, so the most
    # defensible neutral starting point is the latest real trend, not 0%.
    q1_26 = data.get("Revenue", {}).get("26Q1")
    q1_25 = data.get("Revenue", {}).get("25Q1")
    qtr_growth_default = (q1_26 / q1_25 - 1) if (q1_26 and q1_25) else 0.0
    # Annual driver placeholder: a flat, conservative long-term growth rate —
    # matches this model's long-standing default for out-year growth.
    ann_growth_default = 0.05

    for r in range(1, 62):
        for ci in range(1, RB_LAST_COL + 2):
            ws.cell(r, ci).fill = FILL_CREAM

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=RB_LAST_COL)
    t = ws.cell(1, 1, f"{stock_id}  營收驅動因子預測模型 Revenue Build")
    t.font = _f(bold=True, size=12, color="7A0000"); t.alignment = ALIGN_C
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 6

    # ── A. Historical Revenue ─────────────────────────────────────────────
    _sec(ws, 3, "A. 歷史營收 Historical Revenue (引用 Income Model，非重複輸入)", end_col=RB_LAST_COL)
    _rb_write_header(ws, 4)
    _rb_row_label(ws, RB_HIST_REVENUE_ROW, "營業收入淨額 Revenue (億元)", bold=True)
    _rb_row_label(ws, 6, "YoY 成長率 (實際)")
    for key in RB_HIST_KEYS:
        im_col = COL_LTR[key]
        _rb_data_cell(ws, RB_HIST_REVENUE_ROW, key, f"=IFERROR('{INCOME_MODEL_SHEET}'!{im_col}{rev_row},\"\")",
                      fmt=FMT_100M, fill=_cfill(RB_CTYPE[key]), bold=True)
    for key in RB_DRIVER_Q_KEYS + RB_DRIVER_A_KEYS:
        _rb_dash_cell(ws, RB_HIST_REVENUE_ROW, key)
    yoy_pairs = [
        ("25Q1", f"'{INCOME_MODEL_SHEET}'!{COL_LTR['24Q1']}{rev_row}"),
        ("25Q2", f"'{INCOME_MODEL_SHEET}'!{COL_LTR['24Q2']}{rev_row}"),
        ("25Q3", f"'{INCOME_MODEL_SHEET}'!{COL_LTR['24Q3']}{rev_row}"),
        ("25Q4", f"'{INCOME_MODEL_SHEET}'!{COL_LTR['24Q4']}{rev_row}"),
        ("2025", f"{RB_COL_LTR['2024']}{RB_HIST_REVENUE_ROW}"),
        ("26Q1", f"{RB_COL_LTR['25Q1']}{RB_HIST_REVENUE_ROW}"),
    ]
    _rb_dash_cell(ws, 6, "2024")
    for key, prior_ref in yoy_pairs:
        col = RB_COL_LTR[key]
        _rb_data_cell(ws, 6, key, f'=IFERROR({col}{RB_HIST_REVENUE_ROW}/{prior_ref}-1,"")', fmt=FMT_PCT)
    for key in RB_DRIVER_Q_KEYS + RB_DRIVER_A_KEYS:
        _rb_dash_cell(ws, 6, key)

    # ── B. Revenue Driver Build ───────────────────────────────────────────
    ws.row_dimensions[7].height = 6
    _sec(ws, 8, "B. 營收驅動因子預測 Revenue Driver Build — 黃色為可編輯假設", end_col=RB_LAST_COL)
    _rb_write_header(ws, 9)

    for key in RB_HIST_KEYS + ["2026F"]:
        for row in (RB_VOLUME_ROW, RB_ASP_ROW, RB_MIX_ROW, RB_FX_ROW, RB_OVERRIDE_ROW):
            _rb_dash_cell(ws, row, key)

    _rb_row_label(ws, RB_VOLUME_ROW, "出貨量／銷量成長率 Volume Growth YoY (Volume 貢獻度)")
    _rb_row_label(ws, RB_ASP_ROW, "平均售價成長率 ASP Growth YoY (ASP 貢獻度)")
    _rb_row_label(ws, RB_MIX_ROW, "產品組合影響 Product Mix Impact (Mix 貢獻度)")
    _rb_row_label(ws, RB_FX_ROW, "匯率／其他影響 FX/Other Impact (匯率/其他貢獻度)")
    _rb_row_label(ws, RB_IMPLIED_GROWTH_ROW, "隱含營收成長率 Implied Revenue Growth YoY (Driver 加總試算)", bold=True)
    _rb_row_label(ws, RB_OVERRIDE_ROW, "人工覆蓋營收成長率 Manual Override YoY (留空 = 採用 Driver 推導)")
    _rb_row_label(ws, RB_APPLIED_GROWTH_ROW, "實際採用營收成長率 Applied Growth YoY (實際套用公式)", bold=True)
    _rb_row_label(ws, RB_FORECAST_REVENUE_ROW, "營收預測 Forecast Revenue (億元)", bold=True)

    for key in RB_DRIVER_Q_KEYS + RB_DRIVER_A_KEYS:
        vol_default = qtr_growth_default if key in RB_DRIVER_Q_KEYS else ann_growth_default
        _rb_data_cell(ws, RB_VOLUME_ROW, key, round(vol_default, 4), fmt=FMT_PCT, editable=True)
        _rb_data_cell(ws, RB_ASP_ROW, key, 0.0, fmt=FMT_PCT, editable=True)
        _rb_data_cell(ws, RB_MIX_ROW, key, 0.0, fmt=FMT_PCT, editable=True)
        _rb_data_cell(ws, RB_FX_ROW, key, 0.0, fmt=FMT_PCT, editable=True)
        col = RB_COL_LTR[key]
        _rb_data_cell(ws, RB_IMPLIED_GROWTH_ROW, key,
                      f"=IFERROR((1+{col}{RB_VOLUME_ROW})*(1+{col}{RB_ASP_ROW})*(1+{col}{RB_MIX_ROW})*(1+{col}{RB_FX_ROW})-1,\"\")",
                      fmt=FMT_PCT, fill=FILL_FCAST_QTR, bold=True)
        _rb_data_cell(ws, RB_OVERRIDE_ROW, key, None, fmt=FMT_PCT, editable=True)
        _rb_data_cell(ws, RB_APPLIED_GROWTH_ROW, key,
                      f'=IFERROR(IF({col}{RB_OVERRIDE_ROW}<>"",{col}{RB_OVERRIDE_ROW},{col}{RB_IMPLIED_GROWTH_ROW}),"")',
                      fmt=FMT_PCT, fill=FILL_FCAST_QTR, bold=True)

    for key in RB_HIST_KEYS:
        _rb_dash_cell(ws, RB_FORECAST_REVENUE_ROW, key)
    for key in RB_DRIVER_Q_KEYS:
        col, prior_col = RB_COL_LTR[key], RB_COL_LTR[RB_Q_PRIOR_KEY[key]]
        _rb_data_cell(ws, RB_FORECAST_REVENUE_ROW, key,
                      f'=IFERROR({prior_col}{RB_HIST_REVENUE_ROW}*(1+{col}{RB_APPLIED_GROWTH_ROW}),"")',
                      fmt=FMT_100M, fill=FILL_FCAST_QTR, bold=True)
    q26_1, q26_2, q26_3, q26_4 = RB_COL_LTR["26Q1"], RB_COL_LTR["26Q2F"], RB_COL_LTR["26Q3F"], RB_COL_LTR["26Q4F"]
    _rb_data_cell(ws, RB_FORECAST_REVENUE_ROW, "2026F",
                  f'=IFERROR({q26_1}{RB_HIST_REVENUE_ROW}+{q26_2}{RB_FORECAST_REVENUE_ROW}+{q26_3}{RB_FORECAST_REVENUE_ROW}+{q26_4}{RB_FORECAST_REVENUE_ROW},"")',
                  fmt=FMT_100M, fill=_cfill("annual_fcast"), bold=True)
    for key in RB_DRIVER_A_KEYS:
        col, prior_col = RB_COL_LTR[key], RB_COL_LTR[RB_ANN_PRIOR_COL[key]]
        _rb_data_cell(ws, RB_FORECAST_REVENUE_ROW, key,
                      f'=IFERROR({prior_col}{RB_FORECAST_REVENUE_ROW}*(1+{col}{RB_APPLIED_GROWTH_ROW}),"")',
                      fmt=FMT_100M, fill=_cfill("annual_fcast"), bold=True)

    ws.row_dimensions[17].height = 6
    ws.row_dimensions[19].height = 6

    # ── C. Segment / Product Mix Framework ────────────────────────────────
    # Every company has a different product mix / weighting / gross margin by
    # segment, so this section is a free-form "what-if" calculator: fill in your
    # own view of Mix %, Volume, ASP and Gross Margin per segment, and the blended
    # margin row shows the resulting weighted margin for reference. It does not
    # automatically overwrite Assumptions' Gross Margin cells (those stay a direct,
    # freely-editable input) — copy the number across yourself if you want to use
    # it, so nothing here can silently change existing output.
    NUM_SEGMENTS = 5
    SEG_SUBROWS = ["營收佔比 %", "出貨量成長率", "平均售價 (ASP) 成長率", "毛利率 %"]
    SEG_ROWS_PER_BLOCK = 1 + len(SEG_SUBROWS)   # 1 name-input row + 4 driver subrows
    SEG_BASE_ROW = 23
    SEG_NAME_ROWS   = [SEG_BASE_ROW + SEG_ROWS_PER_BLOCK * i for i in range(NUM_SEGMENTS)]  # name row of each segment
    SEG_MIX_ROWS    = [r + 1 for r in SEG_NAME_ROWS]                            # Mix % row of each segment
    SEG_MARGIN_ROWS = [r + 3 for r in SEG_MIX_ROWS]                             # Gross Margin % row of each segment
    SEG_LAST_ROW = SEG_BASE_ROW + SEG_ROWS_PER_BLOCK * NUM_SEGMENTS - 1
    MIX_TOTAL_ROW, BLEND_MARGIN_ROW, SEG_NOTE_ROW = SEG_LAST_ROW + 1, SEG_LAST_ROW + 2, SEG_LAST_ROW + 3

    _sec(ws, 20, "C. 產品結構試算 Segment / Product Mix Framework — 依公司自訂，僅供參考不自動覆蓋 Assumptions", end_col=RB_LAST_COL)
    ws.merge_cells(start_row=21, start_column=1, end_row=21, end_column=RB_LAST_COL)
    note1 = ws.cell(21, 1,
        "目前無可靠歷史 Segment Revenue 拆分數據，故採用「簡化 Driver Mode」：以下 Segment 假設僅供情境分析與框架延伸，"
        "預設不影響第 B 段 Total Revenue Forecast（Total Revenue 由 Volume/ASP/Mix/FX 直接推導）。"
        "每家公司的產品線名稱、組合、占比、毛利率皆不同，下方黃色格（含產品線名稱本身）請依自己的預估自由填寫，未填寫前不代表任何特定產品。")
    note1.font = _f(size=8, color="666666"); note1.fill = FILL_CREAM
    ws.row_dimensions[21].height = 26

    _rb_write_header(ws, 22)
    for si, name_row in enumerate(SEG_NAME_ROWS):
        _rb_row_label(ws, name_row, f"產品線 {si + 1} 名稱 Segment {si + 1} Name（請自行輸入）", bold=True)
        ws.merge_cells(start_row=name_row, start_column=2, end_row=name_row, end_column=RB_LAST_COL)
        name_cell = ws.cell(name_row, 2)
        name_cell.font = _f(size=9, bold=True)
        name_cell.fill = FILL_INPUT
        name_cell.alignment = ALIGN_L
        name_cell.border = BORDER_D
        for j, sub in enumerate(SEG_SUBROWS):
            row = name_row + 1 + j
            _rb_row_label(ws, row, sub)
            for key in RB_HIST_KEYS + ["2026F"]:
                _rb_na_cell(ws, row, key)
            for key in RB_DRIVER_Q_KEYS + RB_DRIVER_A_KEYS:
                _rb_data_cell(ws, row, key, None, fmt=FMT_PCT, editable=True)

    _rb_row_label(ws, MIX_TOTAL_ROW, "產品結構佔比合計 Segment Mix Total (應 = 100% 或 N/A)", bold=True)
    for key in RB_HIST_KEYS + ["2026F"]:
        _rb_na_cell(ws, MIX_TOTAL_ROW, key)
    for key in RB_DRIVER_Q_KEYS + RB_DRIVER_A_KEYS:
        col = RB_COL_LTR[key]
        mix_refs = [f"{col}{r}" for r in SEG_MIX_ROWS]
        f = f'=IF(COUNT({",".join(mix_refs)})=0,"N/A",{"+".join(mix_refs)})'
        _rb_data_cell(ws, MIX_TOTAL_ROW, key, f, fmt=FMT_PCT, fill=FILL_FCAST_QTR, bold=True)

    _rb_row_label(ws, BLEND_MARGIN_ROW, "加權毛利率試算 Blended Gross Margin (Mix % x Segment 毛利率)", bold=True)
    for key in RB_HIST_KEYS + ["2026F"]:
        _rb_na_cell(ws, BLEND_MARGIN_ROW, key)
    for key in RB_DRIVER_Q_KEYS + RB_DRIVER_A_KEYS:
        col = RB_COL_LTR[key]
        mix_refs = [f"{col}{r}" for r in SEG_MIX_ROWS]
        weighted_terms = [f"{col}{mr}*{col}{gr}" for mr, gr in zip(SEG_MIX_ROWS, SEG_MARGIN_ROWS)]
        denom = f'SUM({",".join(mix_refs)})'
        f = (f'=IFERROR(IF({denom}=0,"N/A",'
             f'({"+".join(weighted_terms)})/{denom}),"N/A")')
        _rb_data_cell(ws, BLEND_MARGIN_ROW, key, f, fmt=FMT_PCT, fill=FILL_FCAST_QTR, bold=True)

    ws.merge_cells(start_row=SEG_NOTE_ROW, start_column=1, end_row=SEG_NOTE_ROW, end_column=RB_LAST_COL)
    seg_note = ws.cell(SEG_NOTE_ROW, 1,
        "說明：Blended Gross Margin 為試算參考值，不會自動套用到 Assumptions。"
        "如果你認為這個依產品結構算出來的毛利率更貼近實際，請自行複製到 Assumptions 對應欄位（第 6 列／第 18 列）的黃色格取代原本數值；"
        "如未填寫任何 Segment 假設，此列顯示 N/A，Assumptions 毛利率維持原本直接輸入的方式不受影響。")
    seg_note.font = _f(size=8, color="7A0000"); seg_note.fill = FILL_CREAM
    ws.row_dimensions[SEG_NOTE_ROW].height = 26

    # ── D. Model Checks ────────────────────────────────────────────────────
    # Label merged across A:D (result moved to column E) — a single narrow column-A
    # cell isn't wide enough for a bilingual label once Chinese text is involved,
    # and column B can't be used for overflow because it already holds the result.
    CHECK_HDR_ROW = SEG_NOTE_ROW + 2
    ws.row_dimensions[SEG_NOTE_ROW + 1].height = 6
    _sec(ws, CHECK_HDR_ROW, "D. 模型檢查 Model Checks", end_col=RB_LAST_COL)
    CHECK_ROW0 = CHECK_HDR_ROW + 1
    CHECK_RESULT_COL = 5   # column E
    q2, q3, q4, a26 = RB_COL_LTR["26Q2F"], RB_COL_LTR["26Q3F"], RB_COL_LTR["26Q4F"], RB_COL_LTR["2026F"]
    m1, n1, o1 = RB_COL_LTR["2027F"], RB_COL_LTR["2028F"], RB_COL_LTR["2029F"]
    im_m, im_o = COL_LTR["26Q2F"], COL_LTR["26Q4F"]
    im_q, im_s = COL_LTR["2027F"], COL_LTR["2029F"]

    checks = [
        (CHECK_ROW0 + 0, "1. 季度加總檢查 Quarterly Sum Check：2026F = 26Q1 + Q2F + Q3F + Q4F",
             f'=IF(ABS({a26}{RB_FORECAST_REVENUE_ROW}-({q26_1}{RB_HIST_REVENUE_ROW}+{q2}{RB_FORECAST_REVENUE_ROW}+{q3}{RB_FORECAST_REVENUE_ROW}+{q4}{RB_FORECAST_REVENUE_ROW}))<0.01,"OK","ERROR")'),
        (CHECK_ROW0 + 1, "2. 產品結構佔比檢查 Segment Mix Check：合計 = 100% 或 N/A",
             f'=IF(COUNTIF({q2}{MIX_TOTAL_ROW}:{q4}{MIX_TOTAL_ROW},"N/A")+COUNTIF({m1}{MIX_TOTAL_ROW}:{o1}{MIX_TOTAL_ROW},"N/A")=6,"N/A",'
             f'IF(SUMPRODUCT(ISNUMBER({q2}{MIX_TOTAL_ROW}:{q4}{MIX_TOTAL_ROW})*(ABS(N({q2}{MIX_TOTAL_ROW}:{q4}{MIX_TOTAL_ROW})-1)>0.01))'
             f'+SUMPRODUCT(ISNUMBER({m1}{MIX_TOTAL_ROW}:{o1}{MIX_TOTAL_ROW})*(ABS(N({m1}{MIX_TOTAL_ROW}:{o1}{MIX_TOTAL_ROW})-1)>0.01))>0,"ERROR","OK"))'),
        (CHECK_ROW0 + 2, "3. 營收連結檢查 Revenue Link Check：Revenue Build 與 Income Model 一致",
             f'=IF(AND(SUMPRODUCT(ABS(({q2}{RB_FORECAST_REVENUE_ROW}:{q4}{RB_FORECAST_REVENUE_ROW})-(\'{INCOME_MODEL_SHEET}\'!{im_m}{rev_row}:{im_o}{rev_row})))<0.01,'
             f'SUMPRODUCT(ABS(({m1}{RB_FORECAST_REVENUE_ROW}:{o1}{RB_FORECAST_REVENUE_ROW})-(\'{INCOME_MODEL_SHEET}\'!{im_q}{rev_row}:{im_s}{rev_row})))<0.01),"OK","ERROR")'),
        (CHECK_ROW0 + 3, "4. 成長率計算檢查 Growth Calculation Check：Driver 公式無錯誤值",
             f'=IF(SUMPRODUCT(--ISERROR({q2}{RB_IMPLIED_GROWTH_ROW}:{q4}{RB_IMPLIED_GROWTH_ROW}))+SUMPRODUCT(--ISERROR({m1}{RB_IMPLIED_GROWTH_ROW}:{o1}{RB_IMPLIED_GROWTH_ROW}))=0,"OK","ERROR")'),
    ]
    for row, label, formula in checks:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=CHECK_RESULT_COL - 1)
        _rb_row_label(ws, row, label, bold=True)
        c = ws.cell(row, CHECK_RESULT_COL, formula)
        c.font = _f(bold=True, size=9); c.fill = FILL_CREAM
        c.alignment = ALIGN_C; c.border = BORDER_HL

    for row, *_rest in checks:
        rng = f"{get_column_letter(CHECK_RESULT_COL)}{row}"
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{rng}="OK"'], font=Font(color="1E7B34", bold=True)))
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{rng}="ERROR"'], font=Font(color="B00020", bold=True)))

    # CHECK_ROW0+0..+3 are the 4 check rows themselves — the spacer must be the row
    # AFTER that block (CHECK_ROW0+4), not "NOTE2_ROW-1" computed from a NOTE2_ROW
    # that was only 1 row past the last check: that collided with (and silently
    # re-shrank) the last check row's own height, which is exactly what "row 43's
    # text gets eaten" turned out to be.
    NOTE2_ROW = CHECK_ROW0 + 5
    ws.row_dimensions[CHECK_ROW0 + 4].height = 6
    ws.merge_cells(start_row=NOTE2_ROW, start_column=1, end_row=NOTE2_ROW, end_column=RB_LAST_COL)
    note2 = ws.cell(NOTE2_ROW, 1,
        "說明：Revenue Build 為 Income Model 未來營收的唯一來源 (Single Source of Truth)。"
        "季度預測 = 去年同期營收 x (1+Applied Growth)；年度預測 = 前一年度營收 x (1+Applied Growth)。"
        "季度出貨量成長率初始值 = 最近一期實際 YoY（近期實際趨勢作為中性起點）；"
        "年度出貨量成長率初始值 = 5%（保守長期成長假設，可自行調整）。"
        "平均售價／產品組合／匯率初始值均為 0%。如需調整，請直接修改本表黃色格子，或於「人工覆蓋」列強制指定成長率。")
    note2.font = _f(size=8, color="7A0000"); note2.fill = FILL_CREAM
    ws.row_dimensions[NOTE2_ROW].height = 26

    ws.column_dimensions["A"].width = 42
    for _, key, _c in RB_COLUMNS:
        ws.column_dimensions[RB_COL_LTR[key]].width = 11
    ws.freeze_panes = "B5"


# ── Build Dashboard ───────────────────────────────────────────────────────────
def build_dashboard_sheet(wb: Workbook, stock_id: str, layout: dict) -> None:
    ws = wb.create_sheet(DASHBOARD_SHEET)
    ws.sheet_properties.tabColor = "7A0000"

    ir  = layout["item_row"]
    mr  = layout["margin_row"]
    yir = layout["yoy_item_row"]

    ANN_YEARS  = ["2024", "2025", "2026(F)", "2027(F)", "2028(F)", "2029(F)"]
    ANN_IM_COL = {
        "2024":    COL_LTR["2024"],
        "2025":    COL_LTR["2025"],
        "2026(F)": COL_LTR["2026F"],
        "2027(F)": COL_LTR["2027F"],
        "2028(F)": COL_LTR["2028F"],
        "2029(F)": COL_LTR["2029F"],
    }

    # ── Cream background ─────────────────────────────────────────────────────
    for r in range(1, 80):
        for ci in range(1, 22):
            ws.cell(r, ci).fill = FILL_CREAM

    # ── Title ────────────────────────────────────────────────────────────────
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=18)
    t = ws.cell(1, 1, f"{stock_id}  Financial Dashboard   單位：億元")
    t.font = _f(bold=True, size=13, color="7A0000"); t.alignment = ALIGN_C
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 5

    # ── Source Table (rows 3-13, cols A-G) ───────────────────────────────────
    # Col A holds English metric names — used directly as chart series labels.
    # Cols B-G reference Income Model annual columns.
    ST_HDR   = 3   # header / year-label row
    ST_FIRST = 4   # first data row

    # (legend_name, income_model_section, source_key, number_format)
    SRC_METRICS = [
        ("Revenue",          "data",   "Revenue",                            FMT_100M),
        ("Gross Profit",     "data",   "GrossProfit",                        FMT_100M),
        ("Operating Income", "data",   "OperatingIncome",                    FMT_100M),
        ("Net Income",       "data",   "EquityAttributableToOwnersOfParent", FMT_100M),
        ("EPS",              "data",   "EPS",                                FMT_EPS),
        ("Revenue YoY",      "yoy",    "Revenue",                            FMT_PCT),
        ("Gross Margin",     "margin", "毛利率",                             FMT_PCT),
        ("Op. Margin",       "margin", "營業利益率",                         FMT_PCT),
        ("Net Margin",       "margin", "稅後淨利率",                         FMT_PCT),
        ("EPS YoY",          "yoy",    "EPS",                                FMT_PCT),
    ]

    # Header row
    h = ws.cell(ST_HDR, 1, "Dashboard Source Table")
    h.font = _f(bold=True, size=9, color="FFFFFF"); h.fill = FILL_WINE
    h.alignment = ALIGN_L; h.border = BORDER_D
    ws.row_dimensions[ST_HDR].height = 16
    for j, yr in enumerate(ANN_YEARS, start=2):
        c = ws.cell(ST_HDR, j, yr)
        c.font = _f(bold=True, size=9, color="FFFFFF"); c.fill = FILL_WINE
        c.alignment = ALIGN_C; c.border = BORDER_D

    # Data rows
    src_row: dict[str, int] = {}
    for idx, (en_name, section, src_key, fmt) in enumerate(SRC_METRICS):
        row = ST_FIRST + idx
        src_row[en_name] = row
        is_hl  = en_name in ("Revenue", "EPS", "Gross Margin", "Op. Margin")
        border = BORDER_HL if is_hl else BORDER_D

        lc = ws.cell(row, 1, en_name)
        lc.font = _f(bold=is_hl, size=9); lc.fill = FILL_CREAM
        lc.alignment = ALIGN_L; lc.border = border
        ws.row_dimensions[row].height = 15

        for j, yr in enumerate(ANN_YEARS, start=2):
            im_col = ANN_IM_COL[yr]
            if section == "data":
                im_row = ir[src_key]
            elif section == "yoy":
                im_row = yir[src_key]
            else:
                im_row = mr[src_key]
            cell = ws.cell(row, j)
            cell.value         = f"='Income Model'!{im_col}{im_row}"
            cell.number_format = fmt
            cell.alignment     = ALIGN_R
            cell.font          = _f(bold=is_hl, size=9)
            cell.fill          = FILL_FCAST_ANN if "(F)" in yr else FILL_HIST_ANN
            cell.border        = border

    # ── Forecast Summary (rows 3-10, cols I-M) ───────────────────────────────
    # Cols: I=9 (label), J=10 (2026F), K=11 (2027F), L=12 (2028F), M=13 (2029F)
    # Source table col indices: D=4 (2026F), E=5 (2027F), F=6 (2028F), G=7 (2029F)
    FCST_YEARS    = ["2026(F)", "2027(F)", "2028(F)", "2029(F)"]
    FCST_LBL_COL  = 9    # col I
    FCST_YEAR_COL = {yr: FCST_LBL_COL + 1 + i for i, yr in enumerate(FCST_YEARS)}
    FCST_SRC_COL  = {yr: ANN_YEARS.index(yr) + 2 for yr in FCST_YEARS}  # col index in source table

    ws.merge_cells(start_row=ST_HDR, start_column=FCST_LBL_COL,
                   end_row=ST_HDR, end_column=FCST_LBL_COL + 4)
    fsh = ws.cell(ST_HDR, FCST_LBL_COL, "Forecast Summary (2026F - 2029F)")
    fsh.font = _f(bold=True, size=9, color="FFFFFF"); fsh.fill = FILL_WINE
    fsh.alignment = ALIGN_L; fsh.border = BORDER_D

    # Year sub-headers (row 4, cols J-M) — col I left as spacer/blank
    ws.cell(ST_FIRST, FCST_LBL_COL).fill = FILL_WINE
    ws.cell(ST_FIRST, FCST_LBL_COL).border = BORDER_D
    for yr in FCST_YEARS:
        c = ws.cell(ST_FIRST, FCST_YEAR_COL[yr], yr)
        c.font = _f(bold=True, size=8, color="FFFFFF"); c.fill = FILL_WINE
        c.alignment = ALIGN_C; c.border = BORDER_D

    # 5 key summary rows (rows 5-9)
    FCST_ROWS = [
        ("Revenue (億元)", "Revenue",     FMT_100M),
        ("Revenue YoY",    "Revenue YoY", FMT_PCT),
        ("Gross Margin",   "Gross Margin",FMT_PCT),
        ("Op. Margin",     "Op. Margin",  FMT_PCT),
        ("EPS (元)",       "EPS",         FMT_EPS),
    ]
    for mi, (lbl, src_name, fmt) in enumerate(FCST_ROWS):
        row    = ST_FIRST + 1 + mi
        is_hl  = mi in (0, 4)
        lc = ws.cell(row, FCST_LBL_COL, lbl)
        lc.font = _f(bold=is_hl, size=8); lc.fill = FILL_CREAM
        lc.alignment = ALIGN_L; lc.border = BORDER_D
        ws.row_dimensions[row].height = 15
        st_row = src_row[src_name]
        for yr in FCST_YEARS:
            src_col = get_column_letter(FCST_SRC_COL[yr])
            cell = ws.cell(row, FCST_YEAR_COL[yr])
            cell.value         = f"={src_col}{st_row}"
            cell.number_format = fmt
            cell.alignment     = ALIGN_R
            cell.font          = _f(bold=is_hl, size=8)
            cell.fill          = FILL_FCAST_ANN
            cell.border        = BORDER_D

    # ── Charts (2-column × 3-row layout, starting row 16) ────────────────────
    # Categories: year-label row (row 3, cols B-G) — use StrRef so "(F)" text parses correctly
    cats_f = f"'{ws.title}'!$B${ST_HDR}:$G${ST_HDR}"

    R0 = 16
    ANCHORS = [
        (f"B{R0}",                    f"J{R0}"),
        (f"B{R0 + CHART_ROW_GAP}",     f"J{R0 + CHART_ROW_GAP}"),
        (f"B{R0 + 2 * CHART_ROW_GAP}", f"J{R0 + 2 * CHART_ROW_GAP}"),
    ]

    def _apply_cats(ch) -> None:
        """Attach StrRef categories to every series in the chart."""
        for s in ch.series:
            try:
                s.cat = AxDataSource(strRef=StrRef(f=cats_f))
            except Exception:
                pass

    def _lc(title: str, y_fmt: str, metrics: list) -> LineChart:
        ch = LineChart()
        for idx, (name, color) in enumerate(metrics):
            ref = Reference(ws, min_col=2, max_col=7,
                            min_row=src_row[name], max_row=src_row[name])
            ch.add_data(ref, from_rows=True)   # 1 row → exactly 1 series
            s = ch.series[idx]
            s.tx = SeriesLabel(v=name)
            _style_series_line(s, color)
        _style_chart(ch, title, y_fmt, legend=len(metrics) > 1)
        _apply_cats(ch)
        return ch

    def _bc(title: str, metrics: list) -> BarChart:
        ch = BarChart()
        ch.type = "col"
        ch.grouping = "clustered"
        for idx, (name, color) in enumerate(metrics):
            ref = Reference(ws, min_col=2, max_col=7,
                            min_row=src_row[name], max_row=src_row[name])
            ch.add_data(ref, from_rows=True)
            s = ch.series[idx]
            s.tx = SeriesLabel(v=name)
            _style_series_bar(s, color)
        _style_chart(ch, title, "#,##0", legend=len(metrics) > 1)
        _apply_cats(ch)
        return ch

    # Row 1 — Revenue Trend (left) | Gross Profit Trend (right)
    ws.add_chart(_lc("Revenue Trend (億元)", "#,##0",
                     [("Revenue", PALETTE[0])]), ANCHORS[0][0])
    ws.add_chart(_lc("Gross Profit Trend (億元)", "#,##0",
                     [("Gross Profit", PALETTE[0])]), ANCHORS[0][1])

    # Row 2 — Operating Income Trend (left) | EPS Trend (right)
    ws.add_chart(_lc("Operating Income Trend (億元)", "#,##0",
                     [("Operating Income", PALETTE[0])]), ANCHORS[1][0])
    ws.add_chart(_lc("EPS Trend (元)", "0.00",
                     [("EPS", PALETTE[0])]), ANCHORS[1][1])

    # Row 3 — Margin Trend (left) | P&L Comparison (right)
    ws.add_chart(_lc("Margin Trend", "0%", [
        ("Gross Margin", PALETTE[0]),
        ("Op. Margin",   PALETTE[1]),
        ("Net Margin",   PALETTE[2]),
    ]), ANCHORS[2][0])
    ws.add_chart(_bc("P&L Comparison (億元)", [
        ("Revenue",          PALETTE[0]),
        ("Gross Profit",     PALETTE[1]),
        ("Operating Income", PALETTE[2]),
        ("Net Income",       PALETTE[3]),
    ]), ANCHORS[2][1])

    # ── Column widths ─────────────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 20
    for col in ["B", "C", "D", "E", "F", "G", "H"]:
        ws.column_dimensions[col].width = 10
    for col in ["I", "J", "K", "L", "M"]:
        ws.column_dimensions[col].width = 13

# ── Build PE Band ────────────────────────────────────────────────────────────
def _qfill(key: str) -> PatternFill:
    return FILL_FCAST_QTR if key.endswith("F") else FILL_CREAM


def _band_row_label(ws, row: int, text: str, bold: bool = False) -> None:
    # A leading '=', '+', '-', or '@' makes Excel/openpyxl treat a plain-text label as
    # a FORMULA instead of a string — guard against it (see _rb_row_label for the
    # incident this is copied from).
    if text and text[0] in "=+-@":
        raise ValueError(f"Row {row} label starts with '{text[0]}' — Excel would read this "
                          f"as a formula, not text: {text!r}")
    c = ws.cell(row, 1, text)
    c.font = _f(bold=bold, size=9); c.fill = FILL_CREAM
    c.alignment = ALIGN_L; c.border = BORDER_HL if bold else BORDER_D
    ws.row_dimensions[row].height = 15


def _band_quarter_header(ws, row: int, keys: list[str]) -> None:
    lc = ws.cell(row, 1, "季度")
    lc.font = _f(bold=True, size=9, color="FFFFFF"); lc.fill = FILL_WINE
    lc.alignment = ALIGN_C; lc.border = BORDER_D
    for i, key in enumerate(keys):
        c = ws.cell(row, 2 + i, COL_HDR[key])
        c.font = _f(bold=True, size=9, color="FFFFFF" if not key.endswith("F") else "3A2A1A")
        c.fill = FILL_WINE if not key.endswith("F") else FILL_FCAST_QTR
        c.alignment = ALIGN_C; c.border = BORDER_D
    ws.row_dimensions[row].height = 16


def build_pe_band_sheet(wb: Workbook, stock_id: str, ir: dict, data: dict, price_df: pd.DataFrame | None,
                         per_df: pd.DataFrame | None, band_lookback_years: int) -> None:
    ws = wb.create_sheet(PE_BAND_SHEET)
    ws.sheet_properties.tabColor = "7A0000"
    keys = BAND_QTR_KEYS
    n = len(keys)
    last_col = 1 + n

    for r in range(1, 40):
        for ci in range(1, last_col + 2):
            ws.cell(r, ci).fill = FILL_CREAM

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    t = ws.cell(1, 1, f"{stock_id}  12個月前瞻本益比河流圖 (12M Forward P/E Band)   單位：新台幣元")
    t.font = _f(bold=True, size=12, color="7A0000"); t.alignment = ALIGN_C
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 6

    # ── PER statistics (bake-time, from FinMind daily data) ──────────────────
    stats = _band_stats(per_df["PER"], band_lookback_years) if per_df is not None and not per_df.empty else None
    _sec(ws, 3, f"PER 統計 (Trailing {band_lookback_years} 年日資料，來源：FinMind)", end_col=last_col)
    STAT_MEAN_R, STAT_SD_R, STAT_N_R, STAT_RANGE_R = 4, 5, 6, 7
    for row, label in [(STAT_MEAN_R, "平均 PER (x)"), (STAT_SD_R, "標準差 (x)"),
                        (STAT_N_R, "樣本數 (交易日)"), (STAT_RANGE_R, "資料期間")]:
        _band_row_label(ws, row, label)
    if stats:
        for row, val, fmt, editable in [(STAT_MEAN_R, stats["mean"], FMT_MULT, True),
                                         (STAT_SD_R, stats["sd"], FMT_MULT, True),
                                         (STAT_N_R, stats["n"], FMT_YEAR, False)]:
            c = ws.cell(row, 2, round(val, 4) if isinstance(val, float) else val)
            c.number_format = fmt; c.alignment = ALIGN_R; c.border = BORDER_D
            c.font = _f(size=9); c.fill = FILL_INPUT if editable else FILL_CREAM
        d0, d1 = pd.to_datetime(per_df["date"]).min().date(), pd.to_datetime(per_df["date"]).max().date()
        rc = ws.cell(STAT_RANGE_R, 2, f"{d0} ~ {d1}")
        rc.font = _f(size=9); rc.alignment = ALIGN_L; rc.border = BORDER_D; rc.fill = FILL_CREAM
    else:
        nc = ws.cell(STAT_MEAN_R, 2, "無法取得 FinMind PER 資料")
        nc.font = _f(size=9, color="7A0000"); nc.alignment = ALIGN_L; nc.border = BORDER_D
    note = ws.cell(8, 1, "平均值/標準差為黃色可編輯格 — 可手動覆寫；重新執行程式會以最新資料回填預設值。"
                         "已排除偏離中位數 3 倍以上的極端值（例如單季獲利驟降造成的異常本益比）。"
                         "NTM EPS 於超出季度預測範圍時（如 27Q1F），以次一年度 EPS 預測 ÷4 作為 proxy，待建立完整季度預測後可替換。")
    note.font = _f(size=8, color="666666"); note.fill = FILL_CREAM
    ws.row_dimensions[8].height = 26

    # ── NTM EPS & implied forward PE ──────────────────────────────────────────
    ws.row_dimensions[9].height = 6
    _sec(ws, 10, "12個月前瞻 EPS (NTM EPS) 與隱含 Forward P/E", end_col=last_col)
    _band_quarter_header(ws, 11, keys)
    NTM_R, PRICE_R, IMPPE_R = 12, 13, 14
    _band_row_label(ws, NTM_R, "NTM EPS (元)", bold=True)
    _band_row_label(ws, PRICE_R, "季底收盤價 (元)")
    _band_row_label(ws, IMPPE_R, "隱含 Forward P/E (x)", bold=True)
    # A quarter counts as "reported" for price/multiple purposes based on whether we
    # can actually find a trading-day close for it — NOT on the model's static
    # 26Q2F/26Q3F/26Q4F template labels, which stay "forecast" even once FinMind has
    # real data for that quarter (see _ntm_formula for the same issue on the EPS side).
    price_found: set[str] = set()
    for i, key in enumerate(keys):
        col = get_column_letter(2 + i)
        fill = _qfill(key)
        ntm_c = ws.cell(NTM_R, 2 + i, _ntm_formula(key, ir["EPS"], data))
        ntm_c.number_format = FMT_EPS; ntm_c.alignment = ALIGN_R; ntm_c.font = _f(bold=True, size=9)
        ntm_c.fill = fill; ntm_c.border = BORDER_HL

        price_cell = ws.cell(PRICE_R, 2 + i)
        price_cell.number_format = FMT_PRICE; price_cell.alignment = ALIGN_R
        price_cell.font = _f(size=9); price_cell.fill = fill; price_cell.border = BORDER_D
        px = _price_on_or_before(price_df, _quarter_end_date(key))
        if px is not None:
            price_cell.value = round(px, 2)
            price_found.add(key)

        pe_cell = ws.cell(IMPPE_R, 2 + i)
        pe_cell.number_format = FMT_MULT; pe_cell.alignment = ALIGN_R
        pe_cell.font = _f(bold=True, size=9); pe_cell.fill = fill; pe_cell.border = BORDER_HL
        if key in price_found:
            pe_cell.value = f'=IFERROR({col}{PRICE_R}/{col}{NTM_R},"")'

    # ── Fair value band ────────────────────────────────────────────────────────
    ws.row_dimensions[15].height = 6
    _sec(ws, 16, "估值區間 Fair Value Band (股價，元)", end_col=last_col)
    _band_quarter_header(ws, 17, keys)
    BAND_ROWS = [
        ("股價 (實際)",      None),
        ("Mean − 2SD",       -2),
        ("Mean − 1SD",       -1),
        ("Mean (平均本益比)", 0),
        ("Mean + 1SD",        1),
        ("Mean + 2SD",        2),
    ]
    ACTUAL_R = 18
    for bi, (label, k) in enumerate(BAND_ROWS):
        row = ACTUAL_R + bi
        _band_row_label(ws, row, label, bold=(k == 0))
        for i, key in enumerate(keys):
            col = get_column_letter(2 + i)
            cell = ws.cell(row, 2 + i)
            cell.number_format = FMT_PRICE; cell.alignment = ALIGN_R
            cell.font = _f(bold=(k == 0), size=9)
            cell.fill = _qfill(key); cell.border = BORDER_HL if k == 0 else BORDER_D
            if k is None:
                if key in price_found:
                    cell.value = f"={col}{PRICE_R}"
            elif stats:
                cell.value = f'=IFERROR({col}{NTM_R}*($B${STAT_MEAN_R}+{k}*$B${STAT_SD_R}),"")'
    BAND_LAST_R = ACTUAL_R + len(BAND_ROWS) - 1

    # ── Current valuation read ────────────────────────────────────────────────
    ws.row_dimensions[BAND_LAST_R + 1].height = 6
    READ_HDR = BAND_LAST_R + 2
    _sec(ws, READ_HDR, "目前估值位置 Current Valuation", end_col=last_col)
    latest = _latest_price(price_df)
    rows_lbl = ["最新收盤價 (元)", "最新收盤日期", "目前 Forward P/E (x)", "目前價格 vs. 平均 PER 偏離", "估值判讀"]
    for j, lbl in enumerate(rows_lbl):
        _band_row_label(ws, READ_HDR + 1 + j, lbl)
    R_PX, R_DT, R_PE, R_DEV, R_READ = (READ_HDR + 1, READ_HDR + 2, READ_HDR + 3, READ_HDR + 4, READ_HDR + 5)
    cur_ntm_col = get_column_letter(2 + keys.index("26Q1"))
    if latest:
        px, d = latest
        c = ws.cell(R_PX, 2, round(px, 2)); c.number_format = FMT_PRICE
        c.fill = FILL_INPUT; c.font = _f(size=9); c.alignment = ALIGN_R; c.border = BORDER_D
        c2 = ws.cell(R_DT, 2, str(d)); c2.fill = FILL_CREAM; c2.font = _f(size=9)
        c2.alignment = ALIGN_L; c2.border = BORDER_D
        pe_f = f"=IFERROR($B${R_PX}/{cur_ntm_col}{NTM_R},\"\")"
        c3 = ws.cell(R_PE, 2, pe_f); c3.number_format = FMT_MULT
        c3.fill = FILL_CREAM; c3.font = _f(bold=True, size=9); c3.alignment = ALIGN_R; c3.border = BORDER_HL
        if stats:
            dev_f = f'=IFERROR(($B${R_PE}-$B${STAT_MEAN_R})/$B${STAT_MEAN_R},"")'
            c4 = ws.cell(R_DEV, 2, dev_f); c4.number_format = FMT_PCT
            c4.fill = FILL_CREAM; c4.font = _f(size=9); c4.alignment = ALIGN_R; c4.border = BORDER_D
            read_f = (f'=IFERROR(IF($B${R_PE}<$B${STAT_MEAN_R}-$B${STAT_SD_R},"偏低 / 具吸引力",'
                      f'IF($B${R_PE}>$B${STAT_MEAN_R}+$B${STAT_SD_R},"偏高 / 需留意","合理區間")),"")')
            c5 = ws.cell(R_READ, 2, read_f)
            c5.fill = FILL_CREAM; c5.font = _f(bold=True, size=9, color="7A0000")
            c5.alignment = ALIGN_L; c5.border = BORDER_D
    else:
        ws.cell(R_PX, 2, "無法取得 FinMind 股價資料").font = _f(size=9, color="7A0000")

    # ── Chart ──────────────────────────────────────────────────────────────────
    CHART_ROW = R_READ + 3
    cats_f = f"'{ws.title}'!$B$17:${get_column_letter(1 + n)}$17"
    ch = LineChart()
    band_colors = {
        "股價 (實際)": ACCENT_PRICE, "Mean − 2SD": PALETTE[2], "Mean − 1SD": PALETTE[1],
        "Mean (平均本益比)": PALETTE[0], "Mean + 1SD": PALETTE[1], "Mean + 2SD": PALETTE[2],
    }
    band_dash = {"股價 (實際)": None, "Mean − 2SD": "sysDot", "Mean − 1SD": "dash",
                 "Mean (平均本益比)": None, "Mean + 1SD": "dash", "Mean + 2SD": "sysDot"}
    for idx, (label, _k) in enumerate(BAND_ROWS):
        row = ACTUAL_R + idx
        ref = Reference(ws, min_col=2, max_col=1 + n, min_row=row, max_row=row)
        ch.add_data(ref, from_rows=True)
        s = ch.series[idx]
        s.tx = SeriesLabel(v=label)
        width = 26000 if label == "股價 (實際)" else 16000
        _style_series_line(s, band_colors[label], width_emu=width, dash=band_dash[label])
    for s in ch.series:
        try:
            s.cat = AxDataSource(strRef=StrRef(f=cats_f))
        except Exception:
            pass
    _style_chart(ch, "12M Forward P/E Band", FMT_PRICE, legend=True)
    ch.height, ch.width = 11.5, 23.0
    ws.add_chart(ch, f"A{CHART_ROW}")

    ws.column_dimensions["A"].width = 22
    for i in range(n):
        ws.column_dimensions[get_column_letter(2 + i)].width = 11

# ── Build PB Band ────────────────────────────────────────────────────────────
def build_pb_band_sheet(wb: Workbook, stock_id: str, ir: dict, data: dict, bs_extra: dict,
                         price_df: pd.DataFrame | None, per_df: pd.DataFrame | None,
                         band_lookback_years: int) -> None:
    ws = wb.create_sheet(PB_BAND_SHEET)
    ws.sheet_properties.tabColor = "7A0000"
    keys = BAND_QTR_KEYS
    n = len(keys)
    last_col = 1 + n

    for r in range(1, 40):
        for ci in range(1, last_col + 2):
            ws.cell(r, ci).fill = FILL_CREAM

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    t = ws.cell(1, 1, f"{stock_id}  12個月前瞻股價淨值比河流圖 (12M Forward P/B Band)   單位：新台幣元")
    t.font = _f(bold=True, size=12, color="7A0000"); t.alignment = ALIGN_C
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 6

    # ── PBR statistics ─────────────────────────────────────────────────────────
    stats = _band_stats(per_df["PBR"], band_lookback_years) if per_df is not None and not per_df.empty else None
    _sec(ws, 3, f"PBR 統計 (Trailing {band_lookback_years} 年日資料，來源：FinMind)", end_col=last_col)
    STAT_MEAN_R, STAT_SD_R, STAT_N_R, STAT_RANGE_R = 4, 5, 6, 7
    for row, label in [(STAT_MEAN_R, "平均 PBR (x)"), (STAT_SD_R, "標準差 (x)"),
                        (STAT_N_R, "樣本數 (交易日)"), (STAT_RANGE_R, "資料期間")]:
        _band_row_label(ws, row, label)
    if stats:
        for row, val, fmt, editable in [(STAT_MEAN_R, stats["mean"], FMT_MULT, True),
                                         (STAT_SD_R, stats["sd"], FMT_MULT, True),
                                         (STAT_N_R, stats["n"], FMT_YEAR, False)]:
            c = ws.cell(row, 2, round(val, 4) if isinstance(val, float) else val)
            c.number_format = fmt; c.alignment = ALIGN_R; c.border = BORDER_D
            c.font = _f(size=9); c.fill = FILL_INPUT if editable else FILL_CREAM
        d0, d1 = pd.to_datetime(per_df["date"]).min().date(), pd.to_datetime(per_df["date"]).max().date()
        rc = ws.cell(STAT_RANGE_R, 2, f"{d0} ~ {d1}")
        rc.font = _f(size=9); rc.alignment = ALIGN_L; rc.border = BORDER_D; rc.fill = FILL_CREAM
    else:
        nc = ws.cell(STAT_MEAN_R, 2, "無法取得 FinMind PBR 資料")
        nc.font = _f(size=9, color="7A0000"); nc.alignment = ALIGN_L; nc.border = BORDER_D
    note = ws.cell(8, 1, "平均值/標準差為黃色可編輯格 — 可手動覆寫；重新執行程式會以最新資料回填預設值。"
                         "已排除偏離中位數 3 倍以上的極端值。")
    note.font = _f(size=8, color="666666"); note.fill = FILL_CREAM
    ws.row_dimensions[8].height = 13

    # ── Book value per share (BVPS) roll-forward ──────────────────────────────
    ws.row_dimensions[9].height = 6
    _sec(ws, 10, "每股淨值 (BVPS) — 歷史實際 + 預測滾動推算", end_col=last_col)
    _band_quarter_header(ws, 11, keys)
    BVPS_R, PRICE_R, IMPPB_R = 12, 13, 14
    _band_row_label(ws, BVPS_R, "BVPS (元)", bold=True)
    _band_row_label(ws, PRICE_R, "季底收盤價 (元)")
    _band_row_label(ws, IMPPB_R, "隱含 Forward P/B (x)", bold=True)
    payout_ref = f"'{ASSUMPTIONS_SHEET}'!$B${A_PAYOUT_R}"
    # Same staleness issue as PE Band's NTM EPS: Income Model's 26Q2F/26Q3F/26Q4F
    # stay labelled "forecast" even after FinMind has real reported balance-sheet
    # data for that quarter. Check real data availability directly rather than the
    # model's static historical/forecast key split.
    price_found: set[str] = set()
    for i, key in enumerate(keys):
        col = get_column_letter(2 + i)
        fill = _qfill(key)
        plain_key = key[:-1] if key.endswith("F") else key

        bvps_c = ws.cell(BVPS_R, 2 + i)
        bvps_c.number_format = FMT_PRICE; bvps_c.alignment = ALIGN_R
        bvps_c.font = _f(bold=True, size=9); bvps_c.fill = fill; bvps_c.border = BORDER_HL
        eq = bs_extra.get("BS_Equity", {}).get(plain_key)
        cap_raw = data.get("OrdinaryShare", {}).get(plain_key)   # raw NT$ — bs_extra values are already 億元
        if eq is not None and cap_raw:
            bvps_c.value = round(eq / (cap_raw / UNIT) * 10, 2)
        elif i > 0:
            # Roll forward from the previous column's own BVPS — only valid once
            # there IS a previous column in this row to roll forward from.
            prev_col = get_column_letter(2 + i - 1)
            eps_cell = f"'{INCOME_MODEL_SHEET}'!{COL_LTR[key]}{ir['EPS']}"
            bvps_c.value = f'=IFERROR({prev_col}{BVPS_R}+{eps_cell}*(1-{payout_ref}),"")'

        price_cell = ws.cell(PRICE_R, 2 + i)
        price_cell.number_format = FMT_PRICE; price_cell.alignment = ALIGN_R
        price_cell.font = _f(size=9); price_cell.fill = fill; price_cell.border = BORDER_D
        px = _price_on_or_before(price_df, _quarter_end_date(key))
        if px is not None:
            price_cell.value = round(px, 2)
            price_found.add(key)

        pb_cell = ws.cell(IMPPB_R, 2 + i)
        pb_cell.number_format = FMT_MULT; pb_cell.alignment = ALIGN_R
        pb_cell.font = _f(bold=True, size=9); pb_cell.fill = fill; pb_cell.border = BORDER_HL
        if key in price_found:
            pb_cell.value = f'=IFERROR({col}{PRICE_R}/{col}{BVPS_R},"")'

    note2 = ws.cell(15, 1, "說明：預測季度 BVPS = 前一季 BVPS + 該季 EPS 預測 × (1 − 股利發放率假設，見 Assumptions)。"
                           "此為簡化的保留盈餘滾動法，非完整資產負債表預測（本模型未預測其他權益項目變動）。")
    note2.font = _f(size=8, color="666666"); note2.fill = FILL_CREAM
    ws.row_dimensions[15].height = 13

    # ── Fair value band ────────────────────────────────────────────────────────
    ws.row_dimensions[16].height = 6
    _sec(ws, 17, "估值區間 Fair Value Band (股價，元)", end_col=last_col)
    _band_quarter_header(ws, 18, keys)
    BAND_ROWS = [
        ("股價 (實際)",       None),
        ("Mean − 2SD",        -2),
        ("Mean − 1SD",        -1),
        ("Mean (平均淨值比)",  0),
        ("Mean + 1SD",         1),
        ("Mean + 2SD",         2),
    ]
    ACTUAL_R = 19
    for bi, (label, k) in enumerate(BAND_ROWS):
        row = ACTUAL_R + bi
        _band_row_label(ws, row, label, bold=(k == 0))
        for i, key in enumerate(keys):
            col = get_column_letter(2 + i)
            cell = ws.cell(row, 2 + i)
            cell.number_format = FMT_PRICE; cell.alignment = ALIGN_R
            cell.font = _f(bold=(k == 0), size=9)
            cell.fill = _qfill(key); cell.border = BORDER_HL if k == 0 else BORDER_D
            if k is None:
                if key in price_found:
                    cell.value = f"={col}{PRICE_R}"
            elif stats:
                cell.value = f'=IFERROR({col}{BVPS_R}*($B${STAT_MEAN_R}+{k}*$B${STAT_SD_R}),"")'
    BAND_LAST_R = ACTUAL_R + len(BAND_ROWS) - 1

    # ── Current valuation read ────────────────────────────────────────────────
    ws.row_dimensions[BAND_LAST_R + 1].height = 6
    READ_HDR = BAND_LAST_R + 2
    _sec(ws, READ_HDR, "目前估值位置 Current Valuation", end_col=last_col)
    latest = _latest_price(price_df)
    rows_lbl = ["最新收盤價 (元)", "最新收盤日期", "目前 Forward P/B (x)", "目前價格 vs. 平均 PBR 偏離", "估值判讀"]
    for j, lbl in enumerate(rows_lbl):
        _band_row_label(ws, READ_HDR + 1 + j, lbl)
    R_PX, R_DT, R_PB, R_DEV, R_READ = (READ_HDR + 1, READ_HDR + 2, READ_HDR + 3, READ_HDR + 4, READ_HDR + 5)
    cur_bvps_col = get_column_letter(2 + keys.index("26Q1"))
    if latest:
        px, d = latest
        c = ws.cell(R_PX, 2, round(px, 2)); c.number_format = FMT_PRICE
        c.fill = FILL_INPUT; c.font = _f(size=9); c.alignment = ALIGN_R; c.border = BORDER_D
        c2 = ws.cell(R_DT, 2, str(d)); c2.fill = FILL_CREAM; c2.font = _f(size=9)
        c2.alignment = ALIGN_L; c2.border = BORDER_D
        pb_f = f"=IFERROR($B${R_PX}/{cur_bvps_col}{BVPS_R},\"\")"
        c3 = ws.cell(R_PB, 2, pb_f); c3.number_format = FMT_MULT
        c3.fill = FILL_CREAM; c3.font = _f(bold=True, size=9); c3.alignment = ALIGN_R; c3.border = BORDER_HL
        if stats:
            dev_f = f'=IFERROR(($B${R_PB}-$B${STAT_MEAN_R})/$B${STAT_MEAN_R},"")'
            c4 = ws.cell(R_DEV, 2, dev_f); c4.number_format = FMT_PCT
            c4.fill = FILL_CREAM; c4.font = _f(size=9); c4.alignment = ALIGN_R; c4.border = BORDER_D
            read_f = (f'=IFERROR(IF($B${R_PB}<$B${STAT_MEAN_R}-$B${STAT_SD_R},"偏低 / 具吸引力",'
                      f'IF($B${R_PB}>$B${STAT_MEAN_R}+$B${STAT_SD_R},"偏高 / 需留意","合理區間")),"")')
            c5 = ws.cell(R_READ, 2, read_f)
            c5.fill = FILL_CREAM; c5.font = _f(bold=True, size=9, color="7A0000")
            c5.alignment = ALIGN_L; c5.border = BORDER_D
    else:
        ws.cell(R_PX, 2, "無法取得 FinMind 股價資料").font = _f(size=9, color="7A0000")

    # ── Chart ──────────────────────────────────────────────────────────────────
    CHART_ROW = R_READ + 3
    cats_f = f"'{ws.title}'!$B$18:${get_column_letter(1 + n)}$18"
    ch = LineChart()
    band_colors = {
        "股價 (實際)": ACCENT_PRICE, "Mean − 2SD": PALETTE[2], "Mean − 1SD": PALETTE[1],
        "Mean (平均淨值比)": PALETTE[0], "Mean + 1SD": PALETTE[1], "Mean + 2SD": PALETTE[2],
    }
    band_dash = {"股價 (實際)": None, "Mean − 2SD": "sysDot", "Mean − 1SD": "dash",
                 "Mean (平均淨值比)": None, "Mean + 1SD": "dash", "Mean + 2SD": "sysDot"}
    for idx, (label, _k) in enumerate(BAND_ROWS):
        row = ACTUAL_R + idx
        ref = Reference(ws, min_col=2, max_col=1 + n, min_row=row, max_row=row)
        ch.add_data(ref, from_rows=True)
        s = ch.series[idx]
        s.tx = SeriesLabel(v=label)
        width = 26000 if label == "股價 (實際)" else 16000
        _style_series_line(s, band_colors[label], width_emu=width, dash=band_dash[label])
    for s in ch.series:
        try:
            s.cat = AxDataSource(strRef=StrRef(f=cats_f))
        except Exception:
            pass
    _style_chart(ch, "12M Forward P/B Band", FMT_PRICE, legend=True)
    ch.height, ch.width = 11.5, 23.0
    ws.add_chart(ch, f"A{CHART_ROW}")

    ws.column_dimensions["A"].width = 22
    for i in range(n):
        ws.column_dimensions[get_column_letter(2 + i)].width = 11

# ── Build Scenario Analysis ─────────────────────────────────────────────────
def build_scenario_analysis_sheet(wb: Workbook, stock_id: str, ir: dict) -> None:
    """Bull / Base / Bear scenarios. Base is a live read-only link to the single
    source of truth (Revenue Build for 2027F-2029F growth, Income Model's own
    2026F column, Assumptions for margin/opex/tax/capital) — it is never a second
    copy of an assumption. Bull/Bear are expressed as an editable DELTA vs Base
    (not an absolute override), so the actual Bull/Bear rate is itself a formula
    (Base + delta) and needs no Python-side recomputation of Base's value.
    Operating expense ratio, non-operating income, tax rate and capital are held
    the SAME across all three scenarios (only revenue growth and gross margin are
    varied) — this is the standard simplification for a quick scenario table."""
    ws = wb.create_sheet(SCENARIO_SHEET)
    ws.sheet_properties.tabColor = "7A0000"

    SC_COL = {"2026F": "B", "2027F": "C", "2028F": "D", "2029F": "E"}
    LAST_COL = 5

    for r in range(1, 50):
        for ci in range(1, LAST_COL + 2):
            ws.cell(r, ci).fill = FILL_CREAM

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=LAST_COL)
    t = ws.cell(1, 1, f"{stock_id}  情境分析 Scenario Analysis (Bull / Base / Bear)")
    t.font = _f(bold=True, size=12, color="7A0000"); t.alignment = ALIGN_C
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 6

    def _hdr(row: int) -> None:
        h = ws.cell(row, 1, "項目 Item")
        h.font = _f(bold=True, size=9, color="FFFFFF"); h.fill = FILL_WINE
        h.alignment = ALIGN_C; h.border = BORDER_D
        for yr, col in SC_COL.items():
            c = ws.cell(row, {"B": 2, "C": 3, "D": 4, "E": 5}[col], yr.replace("F", "(F)"))
            c.font = _f(bold=True, size=9, color="FFFFFF"); c.fill = FILL_WINE
            c.alignment = ALIGN_C; c.border = BORDER_D
        ws.row_dimensions[row].height = 16

    def _cell(row: int, yr: str, value=None, fmt: str = FMT_PCT, fill=None, bold: bool = False, editable: bool = False):
        col = {"B": 2, "C": 3, "D": 4, "E": 5}[SC_COL[yr]]
        c = ws.cell(row, col, value)
        c.number_format = fmt; c.alignment = ALIGN_R; c.font = _f(bold=bold, size=9)
        c.fill = fill if fill is not None else (FILL_INPUT if editable else FILL_CREAM)
        c.border = BORDER_HL if bold else BORDER_D
        return c

    rev_row, gp_row = ir["Revenue"], ir["GrossProfit"]
    opex_row, oi_row = ir["OperatingExpenses"], ir["OperatingIncome"]
    nonop_row, pretax_row = ir["TotalNonoperatingIncomeAndExpense"], ir["PreTaxIncome"]
    tax_row, ni_row, cap_row = ir["TAX"], ir["EquityAttributableToOwnersOfParent"], ir["OrdinaryShare"]
    IM = INCOME_MODEL_SHEET
    c2026 = COL_LTR["2026F"]

    BASE_GROWTH = {
        "2026F": f"'{IM}'!{c2026}{rev_row}/'{IM}'!{COL_LTR['2025']}{rev_row}-1",
        "2027F": f"'{REVENUE_BUILD_SHEET}'!{RB_COL_LTR['2027F']}{RB_APPLIED_GROWTH_ROW}",
        "2028F": f"'{REVENUE_BUILD_SHEET}'!{RB_COL_LTR['2028F']}{RB_APPLIED_GROWTH_ROW}",
        "2029F": f"'{REVENUE_BUILD_SHEET}'!{RB_COL_LTR['2029F']}{RB_APPLIED_GROWTH_ROW}",
    }
    BASE_MARGIN = {
        "2026F": f"'{IM}'!{c2026}{gp_row}/'{IM}'!{c2026}{rev_row}",
        "2027F": f"'{ASSUMPTIONS_SHEET}'!$B${A_ANN_GM_R}",
        "2028F": f"'{ASSUMPTIONS_SHEET}'!$C${A_ANN_GM_R}",
        "2029F": f"'{ASSUMPTIONS_SHEET}'!$D${A_ANN_GM_R}",
    }
    OPEX_REF = {
        "2026F": f"'{IM}'!{c2026}{opex_row}/'{IM}'!{c2026}{rev_row}",
        "2027F": f"'{ASSUMPTIONS_SHEET}'!$B${A_ANN_OPEX_R}",
        "2028F": f"'{ASSUMPTIONS_SHEET}'!$C${A_ANN_OPEX_R}",
        "2029F": f"'{ASSUMPTIONS_SHEET}'!$D${A_ANN_OPEX_R}",
    }
    NONOP_REF = {
        "2026F": f"'{IM}'!{c2026}{nonop_row}",
        "2027F": f"'{ASSUMPTIONS_SHEET}'!$B${A_ANN_NONOP_R}",
        "2028F": f"'{ASSUMPTIONS_SHEET}'!$C${A_ANN_NONOP_R}",
        "2029F": f"'{ASSUMPTIONS_SHEET}'!$D${A_ANN_NONOP_R}",
    }
    TAX_REF = {
        "2026F": f"'{IM}'!{c2026}{tax_row}/'{IM}'!{c2026}{pretax_row}",
        "2027F": f"'{ASSUMPTIONS_SHEET}'!$B${A_ANN_TAX_R}",
        "2028F": f"'{ASSUMPTIONS_SHEET}'!$C${A_ANN_TAX_R}",
        "2029F": f"'{ASSUMPTIONS_SHEET}'!$D${A_ANN_TAX_R}",
    }
    CAP_REF = {
        "2026F": f"'{IM}'!{c2026}{cap_row}",
        "2027F": f"'{ASSUMPTIONS_SHEET}'!$B${A_ANN_CAP_R}",
        "2028F": f"'{ASSUMPTIONS_SHEET}'!$C${A_ANN_CAP_R}",
        "2029F": f"'{ASSUMPTIONS_SHEET}'!$D${A_ANN_CAP_R}",
    }
    PRIOR_ACTUAL_REV = f"'{IM}'!{COL_LTR['2025']}{rev_row}"

    # ── A. Scenario assumptions ────────────────────────────────────────────
    _sec(ws, 3, "A. 情境假設 Scenario Assumptions — 黃色為可編輯假設 (Bull/Bear 為相對 Base 的調整幅度)", end_col=LAST_COL)
    _hdr(4)
    BASE_GROWTH_R, BULL_DELTA_R, BULL_GROWTH_R, BEAR_DELTA_R, BEAR_GROWTH_R = 5, 6, 7, 8, 9
    _band_row_label(ws, BASE_GROWTH_R, "Base 營收成長率 YoY（連動主模型）", bold=True)
    _band_row_label(ws, BULL_DELTA_R, "Bull 營收成長率調整 Δ（相對 Base）")
    _band_row_label(ws, BULL_GROWTH_R, "Bull 營收成長率 YoY（= Base + Δ）", bold=True)
    _band_row_label(ws, BEAR_DELTA_R, "Bear 營收成長率調整 Δ（相對 Base）")
    _band_row_label(ws, BEAR_GROWTH_R, "Bear 營收成長率 YoY（= Base + Δ）", bold=True)
    ws.row_dimensions[10].height = 6
    BASE_MARGIN_R, BULL_MDELTA_R, BULL_MARGIN_R, BEAR_MDELTA_R, BEAR_MARGIN_R = 11, 12, 13, 14, 15
    _band_row_label(ws, BASE_MARGIN_R, "Base 毛利率（連動主模型）", bold=True)
    _band_row_label(ws, BULL_MDELTA_R, "Bull 毛利率調整 Δ（相對 Base）")
    _band_row_label(ws, BULL_MARGIN_R, "Bull 毛利率（= Base + Δ）", bold=True)
    _band_row_label(ws, BEAR_MDELTA_R, "Bear 毛利率調整 Δ（相對 Base）")
    _band_row_label(ws, BEAR_MARGIN_R, "Bear 毛利率（= Base + Δ）", bold=True)
    ws.row_dimensions[16].height = 6

    for yr in SCENARIO_YEARS:
        col = SC_COL[yr]
        _cell(BASE_GROWTH_R, yr, f"=IFERROR({BASE_GROWTH[yr]},\"\")", bold=True)
        _cell(BULL_DELTA_R, yr, 0.05, editable=True)
        _cell(BULL_GROWTH_R, yr, f"=IFERROR({col}{BASE_GROWTH_R}+{col}{BULL_DELTA_R},\"\")", bold=True)
        _cell(BEAR_DELTA_R, yr, -0.05, editable=True)
        _cell(BEAR_GROWTH_R, yr, f"=IFERROR({col}{BASE_GROWTH_R}+{col}{BEAR_DELTA_R},\"\")", bold=True)
        _cell(BASE_MARGIN_R, yr, f"=IFERROR({BASE_MARGIN[yr]},\"\")", bold=True)
        _cell(BULL_MDELTA_R, yr, 0.02, editable=True)
        _cell(BULL_MARGIN_R, yr, f"=IFERROR({col}{BASE_MARGIN_R}+{col}{BULL_MDELTA_R},\"\")", bold=True)
        _cell(BEAR_MDELTA_R, yr, -0.02, editable=True)
        _cell(BEAR_MARGIN_R, yr, f"=IFERROR({col}{BASE_MARGIN_R}+{col}{BEAR_MDELTA_R},\"\")", bold=True)

    # ── B. Scenario P&L ────────────────────────────────────────────────────
    _sec(ws, 17, "B. 情境損益推算 Scenario P&L（營業費用率／業外收支／稅率／股本三情境共用主模型數值）", end_col=LAST_COL)
    _hdr(18)

    def _pnl_block(base_row: int, growth_row: int, margin_row: int, label_prefix: str) -> dict:
        rev_r, gp_r, oi_r, ni_r, eps_r = base_row, base_row + 1, base_row + 2, base_row + 3, base_row + 4
        _band_row_label(ws, rev_r, f"{label_prefix}－營業收入淨額 Revenue (億元)", bold=True)
        _band_row_label(ws, gp_r, f"{label_prefix}－營業毛利 Gross Profit (億元)")
        _band_row_label(ws, oi_r, f"{label_prefix}－營業淨利 Operating Income (億元)")
        _band_row_label(ws, ni_r, f"{label_prefix}－稅後淨利 Net Income (億元)")
        _band_row_label(ws, eps_r, f"{label_prefix}－每股盈餘 EPS (元)", bold=True)
        prior_ref = PRIOR_ACTUAL_REV
        for yr in SCENARIO_YEARS:
            col = SC_COL[yr]
            _cell(rev_r, yr, f'=IFERROR({prior_ref}*(1+{col}{growth_row}),"")', fmt=FMT_100M, bold=True)
            _cell(gp_r, yr, f'=IFERROR({col}{rev_r}*{col}{margin_row},"")', fmt=FMT_100M)
            _cell(oi_r, yr, f'=IFERROR({col}{gp_r}-{col}{rev_r}*({OPEX_REF[yr]}),"")', fmt=FMT_100M)
            _cell(ni_r, yr, f'=IFERROR(({col}{oi_r}+({NONOP_REF[yr]}))*(1-({TAX_REF[yr]})),"")', fmt=FMT_100M)
            _cell(eps_r, yr, f'=IFERROR({col}{ni_r}/({CAP_REF[yr]})*10,"")', fmt=FMT_EPS, bold=True)
            prior_ref = f"{col}{rev_r}"
        return {"rev": rev_r, "gp": gp_r, "oi": oi_r, "ni": ni_r, "eps": eps_r}

    bull_rows = _pnl_block(19, BULL_GROWTH_R, BULL_MARGIN_R, "Bull")
    ws.row_dimensions[24].height = 6
    base_rows = _pnl_block(25, BASE_GROWTH_R, BASE_MARGIN_R, "Base")
    ws.row_dimensions[30].height = 6
    bear_rows = _pnl_block(31, BEAR_GROWTH_R, BEAR_MARGIN_R, "Bear")
    ws.row_dimensions[36].height = 6

    # ── C. Implied valuation ───────────────────────────────────────────────
    _sec(ws, 37, "C. 情境估值區間 Implied Valuation（EPS x PE Band 統計倍數：Bull=Mean+1SD／Base=Mean／Bear=Mean-1SD）", end_col=LAST_COL)
    _hdr(38)
    BULL_TP_R, BASE_TP_R, BEAR_TP_R = 39, 40, 41
    _band_row_label(ws, BULL_TP_R, "Bull 目標價 (元)", bold=True)
    _band_row_label(ws, BASE_TP_R, "Base 目標價 (元)", bold=True)
    _band_row_label(ws, BEAR_TP_R, "Bear 目標價 (元)", bold=True)
    for yr in SCENARIO_YEARS:
        col = SC_COL[yr]
        _cell(BULL_TP_R, yr, f"=IFERROR({col}{bull_rows['eps']}*('PE Band'!{PE_BAND_MEAN_CELL}+'PE Band'!{PE_BAND_SD_CELL}),\"\")",
              fmt=FMT_PRICE, fill=FILL_FCAST_QTR, bold=True)
        _cell(BASE_TP_R, yr, f"=IFERROR({col}{base_rows['eps']}*'PE Band'!{PE_BAND_MEAN_CELL},\"\")",
              fmt=FMT_PRICE, fill=FILL_FCAST_QTR, bold=True)
        _cell(BEAR_TP_R, yr, f"=IFERROR({col}{bear_rows['eps']}*('PE Band'!{PE_BAND_MEAN_CELL}-'PE Band'!{PE_BAND_SD_CELL}),\"\")",
              fmt=FMT_PRICE, fill=FILL_FCAST_QTR, bold=True)

    ws.row_dimensions[42].height = 6
    CUR_PX_R, BULL_UP_R, BEAR_DOWN_R = 43, 44, 45
    _band_row_label(ws, CUR_PX_R, "目前股價 (元，引用 PE Band 最新收盤價)")
    _band_row_label(ws, BULL_UP_R, "Bull 潛在漲幅 vs 目前股價")
    _band_row_label(ws, BEAR_DOWN_R, "Bear 潛在跌幅 vs 目前股價")
    for yr in SCENARIO_YEARS:
        col = SC_COL[yr]
        _cell(CUR_PX_R, yr, f"=IFERROR('PE Band'!$B${PE_BAND_LATEST_PRICE_ROW},\"\")", fmt=FMT_PRICE)
        _cell(BULL_UP_R, yr, f'=IFERROR({col}{BULL_TP_R}/{col}{CUR_PX_R}-1,"")', fmt=FMT_PCT)
        _cell(BEAR_DOWN_R, yr, f'=IFERROR({col}{BEAR_TP_R}/{col}{CUR_PX_R}-1,"")', fmt=FMT_PCT)

    ws.row_dimensions[46].height = 6
    ws.merge_cells(start_row=47, start_column=1, end_row=47, end_column=LAST_COL)
    note = ws.cell(47, 1,
        "說明：Base 情境直接連動主模型（Revenue Build／Income Model／Assumptions），不是另一份假設；"
        "Bull／Bear 只需調整營收成長率與毛利率兩項，其餘假設（費用率／業外／稅率／股本）三情境共用，"
        "為情境分析的標準簡化做法。目標價 = 該情境 EPS x PE Band 統計倍數，Bull/Bear 同時疊加較高/較低的 EPS 與較高/較低的本益比，"
        "屬於較保守（雙重疊加）的區間估法。PE Band 的平均值/標準差若被手動覆寫，本頁會自動連動更新。")
    note.font = _f(size=8, color="7A0000"); note.fill = FILL_CREAM
    ws.row_dimensions[47].height = 26
    ws.row_dimensions[48].height = 6

    # ── Chart helper row (football-field range chart needs a hidden base-offset
    # series + two stacked spans — not meant to be edited, just chart source data) ──
    HIDDEN_BASE_R, SPAN_LO_R, SPAN_HI_R = 49, 50, 51
    _band_row_label(ws, HIDDEN_BASE_R, "Bear 目標價 (圖表輔助列，勿刪除)")
    _band_row_label(ws, SPAN_LO_R, "Bear→Base 區間 (圖表輔助列)")
    _band_row_label(ws, SPAN_HI_R, "Base→Bull 區間 (圖表輔助列)")
    for yr in SCENARIO_YEARS:
        col = SC_COL[yr]
        _cell(HIDDEN_BASE_R, yr, f"=IFERROR({col}{BEAR_TP_R},\"\")", fmt=FMT_PRICE)
        _cell(SPAN_LO_R, yr, f'=IFERROR({col}{BASE_TP_R}-{col}{BEAR_TP_R},"")', fmt=FMT_PRICE)
        _cell(SPAN_HI_R, yr, f'=IFERROR({col}{BULL_TP_R}-{col}{BASE_TP_R},"")', fmt=FMT_PRICE)
    ws.row_dimensions[52].height = 6

    for r in range(1, 92):
        for ci in range(1, 17):
            ws.cell(r, ci).fill = FILL_CREAM

    # ── Charts (2x2 grid) ─────────────────────────────────────────────────────
    cats_f = f"'{ws.title}'!$B$38:$E$38"

    def _apply_cats(ch) -> None:
        for s in ch.series:
            try:
                s.cat = AxDataSource(strRef=StrRef(f=cats_f))
            except Exception:
                pass

    R0 = 53
    ANCHORS = [(f"A{R0}", f"J{R0}"), (f"A{R0 + CHART_ROW_GAP}", f"J{R0 + CHART_ROW_GAP}")]

    # 1. Target price trend (top-left)
    ch1 = LineChart()
    for row, label, color in [(BULL_TP_R, "Bull 目標價", PALETTE[1]), (BASE_TP_R, "Base 目標價", ACCENT_PRICE),
                               (BEAR_TP_R, "Bear 目標價", PALETTE[2]), (CUR_PX_R, "目前股價", "1A1A1A")]:
        ref = Reference(ws, min_col=2, max_col=LAST_COL, min_row=row, max_row=row)
        ch1.add_data(ref, from_rows=True)
        s = ch1.series[len(ch1.series) - 1]
        s.tx = SeriesLabel(v=label)
        _style_series_line(s, color, dash="dash" if label == "目前股價" else None)
    _apply_cats(ch1)
    _style_chart(ch1, "情境目標價趨勢 Scenario Target Price", FMT_PRICE, legend=True)
    ws.add_chart(ch1, ANCHORS[0][0])

    # 2. EPS comparison (top-right)
    ch2 = BarChart()
    ch2.type = "col"; ch2.grouping = "clustered"
    for row, label, color in [(bull_rows["eps"], "Bull EPS", PALETTE[1]), (base_rows["eps"], "Base EPS", PALETTE[0]),
                               (bear_rows["eps"], "Bear EPS", PALETTE[2])]:
        ref = Reference(ws, min_col=2, max_col=LAST_COL, min_row=row, max_row=row)
        ch2.add_data(ref, from_rows=True)
        s = ch2.series[len(ch2.series) - 1]
        s.tx = SeriesLabel(v=label)
        _style_series_bar(s, color)
    _apply_cats(ch2)
    _style_chart(ch2, "情境 EPS 比較 Scenario EPS", FMT_EPS, legend=True)
    ws.add_chart(ch2, ANCHORS[0][1])

    # 3. Upside / downside % (bottom-left)
    ch3 = BarChart()
    ch3.type = "col"; ch3.grouping = "clustered"
    for row, label, color in [(BULL_UP_R, "Bull 潛在漲幅", PALETTE[1]), (BEAR_DOWN_R, "Bear 潛在跌幅", PALETTE[2])]:
        ref = Reference(ws, min_col=2, max_col=LAST_COL, min_row=row, max_row=row)
        ch3.add_data(ref, from_rows=True)
        s = ch3.series[len(ch3.series) - 1]
        s.tx = SeriesLabel(v=label)
        _style_series_bar(s, color)
    _apply_cats(ch3)
    _style_chart(ch3, "潛在漲跌幅 Upside / Downside vs 目前股價", FMT_PCT, legend=True)
    ws.add_chart(ch3, ANCHORS[1][0])

    # 4. Football-field target price range (bottom-right) — stacked bar with a
    # transparent base-offset series so each column floats from Bear to Bull price.
    ch4 = BarChart()
    ch4.type = "col"; ch4.grouping = "stacked"; ch4.overlap = 100
    for row, label, color, hidden in [(HIDDEN_BASE_R, "(隱藏基準)", None, True),
                                       (SPAN_LO_R, "Bear→Base", PALETTE[2], False),
                                       (SPAN_HI_R, "Base→Bull", PALETTE[1], False)]:
        ref = Reference(ws, min_col=2, max_col=LAST_COL, min_row=row, max_row=row)
        ch4.add_data(ref, from_rows=True)
        s = ch4.series[len(ch4.series) - 1]
        s.tx = SeriesLabel(v=label)
        if hidden:
            try:
                s.graphicalProperties.noFill = True
                s.graphicalProperties.line.noFill = True
            except Exception:
                pass
        else:
            _style_series_bar(s, color)
    _apply_cats(ch4)
    _style_chart(ch4, "情境目標價區間 Target Price Range (Bear-Base-Bull)", FMT_PRICE, legend=True)
    ch4.legend.legendEntry = [LegendEntry(idx=0, delete=True)]   # hide the invisible offset series' entry
    ws.add_chart(ch4, ANCHORS[1][1])

    ws.column_dimensions["A"].width = 40
    for col in ["B", "C", "D", "E"]:
        ws.column_dimensions[col].width = 13
    ws.freeze_panes = "B5"


# ── Build Backtest ─────────────────────────────────────────────────────────────
def _weekly_backtest_series(price_df: pd.DataFrame | None, per_df: pd.DataFrame | None,
                             lookback_years: int) -> pd.DataFrame | None:
    """Weekly (last trading day per week) close price + trailing PER, over the same
    lookback window PE Band's Mean/SD are computed from — so the backtest and the
    signal it tests are measuring the same period, and outlier PER values (<=0,
    a loss-making quarter) are dropped since they aren't a meaningful valuation."""
    if price_df is None or price_df.empty or per_df is None or per_df.empty:
        return None
    p = price_df[["date", "close"]].copy()
    r = per_df[["date", "PER"]].copy()
    merged = pd.merge(p, r, on="date", how="inner")
    merged["date"] = pd.to_datetime(merged["date"])
    merged = merged[merged["PER"] > 0]
    if merged.empty:
        return None
    cutoff = merged["date"].max() - pd.DateOffset(years=lookback_years)
    merged = merged[merged["date"] >= cutoff]
    merged = merged.set_index("date").sort_index()
    weekly = merged.resample("W-FRI").last().dropna().reset_index()
    return weekly if len(weekly) >= 10 else None


def build_backtest_sheet(wb: Workbook, stock_id: str, price_df: pd.DataFrame | None,
                          per_df: pd.DataFrame | None, band_lookback_years: int) -> None:
    """Naive backtest of the PE Band valuation signal (buy below Mean-1SD, sell
    above Mean+1SD, hold otherwise) vs. buy-and-hold, on weekly-sampled data. Every
    row is a live Excel formula chained off the row above (position/return/equity/
    drawdown), and the signal formula references PE Band's own Mean/SD cells
    directly — so this recalculates automatically if those are hand-edited.
    This is intentionally simple: no transaction costs, no slippage, no dividend
    reinvestment, single security, and the Mean/SD being tested were themselves
    fitted on this same historical window (not a true walk-forward test) — see the
    in-sheet note for the full caveat list."""
    ws = wb.create_sheet(BACKTEST_SHEET)
    ws.sheet_properties.tabColor = "7A0000"

    weekly = _weekly_backtest_series(price_df, per_df, band_lookback_years)

    for r in range(1, 37):
        for ci in range(1, 14):
            ws.cell(r, ci).fill = FILL_CREAM

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=13)
    t = ws.cell(1, 1, f"{stock_id}  估值訊號回測 Backtest — PE Band 訊號 vs. 買進持有")
    t.font = _f(bold=True, size=12, color="7A0000"); t.alignment = ALIGN_C
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 6

    _sec(ws, 3, "A. 回測摘要 Backtest Summary", end_col=13)
    hdr = ["項目", "策略 (PE Band 訊號)", "買進持有 Buy & Hold"]
    for j, txt in enumerate(hdr, start=1):
        c = ws.cell(4, j, txt)
        c.font = _f(bold=True, size=9, color="FFFFFF"); c.fill = FILL_WINE
        c.alignment = ALIGN_C; c.border = BORDER_D
    ws.row_dimensions[4].height = 16

    if weekly is None:
        nc = ws.cell(5, 1, "無法取得足夠的 FinMind 股價／PER 資料，無法執行回測。")
        nc.font = _f(size=9, color="7A0000"); nc.alignment = ALIGN_L
        ws.column_dimensions["A"].width = 30
        return

    # Chart gets its own dedicated rows (14-33) so it never overlaps the summary
    # table, the note row, or the detail table's own section header below — the
    # exact overlap bug fixed earlier on Dashboard, avoided here by construction
    # (default row height ~14.4pt x 20 rows = 10.16cm > 9.0cm chart height, with
    # a buffer margin instead of an exact fit).
    CHART_ROW = 14
    HDR_ROW = 35
    FIRST_ROW = HDR_ROW + 1
    n = len(weekly)
    LAST_ROW = FIRST_ROW + n - 1

    # ── A. Summary (formulas reference the detail table built below) ─────────
    summary_rows = [
        (5, "總報酬率 Total Return", f"=G{LAST_ROW}-1", f"=I{LAST_ROW}-1", FMT_PCT),
        (6, "年化報酬率 CAGR", f"=(G{LAST_ROW})^(365.25/(A{LAST_ROW}-A{FIRST_ROW}))-1",
             f"=(I{LAST_ROW})^(365.25/(A{LAST_ROW}-A{FIRST_ROW}))-1", FMT_PCT),
        (7, "最大回撤 Max Drawdown", f"=MIN(K{FIRST_ROW}:K{LAST_ROW})", f"=MIN(M{FIRST_ROW}:M{LAST_ROW})", FMT_PCT),
        (8, "交易次數 Number of Trades",
             f"=SUMPRODUCT((E{FIRST_ROW+1}:E{LAST_ROW}=1)*(E{FIRST_ROW}:E{LAST_ROW-1}=0))", "N/A", FMT_YEAR),
    ]
    for row, label, f_strat, f_bh, fmt in summary_rows:
        lc = ws.cell(row, 1, label)
        lc.font = _f(size=9); lc.fill = FILL_CREAM; lc.alignment = ALIGN_L; lc.border = BORDER_D
        for col, formula in [(2, f_strat), (3, f_bh)]:
            c = ws.cell(row, col, f"=IFERROR({formula[1:]},\"\")" if formula != "N/A" else "N/A")
            c.number_format = fmt; c.alignment = ALIGN_R; c.font = _f(bold=True, size=9)
            c.fill = FILL_FCAST_QTR; c.border = BORDER_HL
        ws.row_dimensions[row].height = 15

    period_lbl = ws.cell(9, 1, "回測期間 Backtest Period")
    period_lbl.font = _f(size=9); period_lbl.fill = FILL_CREAM; period_lbl.alignment = ALIGN_L; period_lbl.border = BORDER_D
    period_val = ws.cell(9, 2, f"{weekly['date'].min().date()} ~ {weekly['date'].max().date()}")
    period_val.font = _f(size=9); period_val.fill = FILL_CREAM; period_val.alignment = ALIGN_L; period_val.border = BORDER_D
    ws.merge_cells(start_row=9, start_column=2, end_row=9, end_column=3)
    ws.row_dimensions[9].height = 15

    sample_lbl = ws.cell(10, 1, "取樣頻率 Sampling")
    sample_lbl.font = _f(size=9); sample_lbl.fill = FILL_CREAM; sample_lbl.alignment = ALIGN_L; sample_lbl.border = BORDER_D
    sample_val = ws.cell(10, 2, f"週 (每週最後一個交易日，共 {n} 筆)")
    sample_val.font = _f(size=9); sample_val.fill = FILL_CREAM; sample_val.alignment = ALIGN_L; sample_val.border = BORDER_D
    ws.merge_cells(start_row=10, start_column=2, end_row=10, end_column=3)
    ws.row_dimensions[10].height = 15

    ws.row_dimensions[11].height = 6
    ws.merge_cells(start_row=12, start_column=1, end_row=12, end_column=13)
    note = ws.cell(12, 1,
        "說明：策略 = 股價低於 PE Band 之 Mean-1SD 買進、高於 Mean+1SD 賣出、其餘區間持有原部位（0 或 1，不做空）。"
        "訊號直接引用 PE Band 的平均值/標準差儲存格，兩邊若手動調整會連動更新。"
        "本回測為簡化模擬：不計交易成本與滑價、不含股利再投入、僅單一標的、且統計用的平均值/標準差是用同一段歷史區間算出來的"
        "（非真正樣本外的 walk-forward 測試），結果僅供參考，不是保證未來績效。")
    note.font = _f(size=8, color="7A0000"); note.fill = FILL_CREAM
    ws.row_dimensions[12].height = 26
    ws.row_dimensions[13].height = 6

    # ── B. Weekly detail table ────────────────────────────────────────────────
    _sec(ws, 34, "B. 週資料明細 Weekly Detail", end_col=13)
    cols = ["日期 Date", "收盤價 Close", "PER (trailing)", "訊號 Signal", "部位 Position",
            "策略週報酬 Strategy Ret", "策略淨值 Strategy Equity", "買進持有週報酬 B&H Ret",
            "買進持有淨值 B&H Equity", "策略歷史高點 Strategy Peak", "策略回撤 Strategy DD",
            "B&H 歷史高點 B&H Peak", "B&H 回撤 B&H DD"]
    for j, txt in enumerate(cols, start=1):
        c = ws.cell(HDR_ROW, j, txt)
        c.font = _f(bold=True, size=8, color="FFFFFF"); c.fill = FILL_WINE
        c.alignment = ALIGN_C; c.border = BORDER_D
    ws.row_dimensions[HDR_ROW].height = 26

    for i, row in enumerate(weekly.itertuples(index=False)):
        r = FIRST_ROW + i
        is_first = (i == 0)
        a = ws.cell(r, 1, row.date.date()); a.number_format = "yyyy-mm-dd"
        b = ws.cell(r, 2, round(float(row.close), 2))
        c_per = ws.cell(r, 3, round(float(row.PER), 2))
        for cell in (a, b, c_per):
            cell.font = _f(size=8); cell.fill = FILL_CREAM; cell.alignment = ALIGN_R; cell.border = BORDER_D

        d = ws.cell(r, 4, f'=IF(C{r}<\'PE Band\'!{PE_BAND_MEAN_CELL}-\'PE Band\'!{PE_BAND_SD_CELL},"BUY",'
                          f'IF(C{r}>\'PE Band\'!{PE_BAND_MEAN_CELL}+\'PE Band\'!{PE_BAND_SD_CELL},"SELL","HOLD"))')
        e = ws.cell(r, 5, f'=IF(D{r}="BUY",1,IF(D{r}="SELL",0,0))' if is_first
                    else f'=IF(D{r}="BUY",1,IF(D{r}="SELL",0,E{r-1}))')
        f_ = ws.cell(r, 6, 0 if is_first else f'=E{r-1}*(B{r}/B{r-1}-1)')
        g = ws.cell(r, 7, 1 if is_first else f'=G{r-1}*(1+F{r})')
        h = ws.cell(r, 8, 0 if is_first else f'=B{r}/B{r-1}-1')
        i_ = ws.cell(r, 9, 1 if is_first else f'=I{r-1}*(1+H{r})')
        j_ = ws.cell(r, 10, f'=MAX(G${FIRST_ROW}:G{r})')
        k = ws.cell(r, 11, f'=G{r}/J{r}-1')
        l = ws.cell(r, 12, f'=MAX(I${FIRST_ROW}:I{r})')
        m = ws.cell(r, 13, f'=I{r}/L{r}-1')
        for cell, fmt in [(d, None), (e, "0"), (f_, FMT_PCT), (g, "0.000"), (h, FMT_PCT),
                          (i_, "0.000"), (j_, "0.000"), (k, FMT_PCT), (l, "0.000"), (m, FMT_PCT)]:
            cell.font = _f(size=8); cell.fill = FILL_CREAM; cell.alignment = ALIGN_R; cell.border = BORDER_D
            if fmt:
                cell.number_format = fmt
        d.alignment = ALIGN_C

    # ── Chart: strategy vs buy-and-hold equity curve ───────────────────────────
    ch = LineChart()
    for col, label, color in [(7, "策略淨值 Strategy", PALETTE[0]), (9, "買進持有 Buy & Hold", "1A1A1A")]:
        ref = Reference(ws, min_col=col, max_col=col, min_row=HDR_ROW, max_row=LAST_ROW)
        ch.add_data(ref, titles_from_data=True)
        s = ch.series[len(ch.series) - 1]
        _style_series_line(s, color)
    cats = Reference(ws, min_col=1, max_col=1, min_row=FIRST_ROW, max_row=LAST_ROW)
    ch.set_categories(cats)
    _style_chart(ch, "Strategy vs Buy & Hold Equity Curve", "0.00", legend=True)
    ch.height, ch.width = 9.0, 22.0
    ws.add_chart(ch, f"A{CHART_ROW}")

    ws.column_dimensions["A"].width = 26
    for col in ["B", "C"]:
        ws.column_dimensions[col].width = 14
    for col in ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]:
        ws.column_dimensions[col].width = 12
    ws.freeze_panes = f"B{FIRST_ROW}"


# ── Build Dividend History ────────────────────────────────────────────────────
def _dividend_year(row: pd.Series) -> int | None:
    for key in ("CashExDividendTradingDate", "StockExDividendTradingDate", "date"):
        v = row.get(key)
        if v:
            ts = pd.to_datetime(v, errors="coerce")
            if pd.notna(ts):
                return int(ts.year)
    return None


def _annual_eps_complete(data: dict, yy: str) -> float | None:
    """Full-year EPS only if all 4 quarters are actual data — unlike _annual(),
    this deliberately returns None for an in-progress year (e.g. the current year,
    where only Q1 is reported) instead of silently summing a partial year, which
    would understate EPS and blow up any ratio computed from it (payout ratio,
    dividend-table EPS column)."""
    vals = [data.get(EPS_TYPE, {}).get(f"{yy}Q{q}") for q in range(1, 5)]
    if any(v is None for v in vals):
        return None
    return sum(vals)


def payout_ratio_default(dividend_df: pd.DataFrame | None, data: dict) -> float | None:
    """Smart default for Assumptions!B29 — average historical (total DPS / EPS)
    across whichever years overlap the fetched EPS window."""
    if dividend_df is None or dividend_df.empty:
        return None
    d = dividend_df.copy()
    d["_year"] = d.apply(_dividend_year, axis=1)
    d = d.dropna(subset=["_year"])
    d["_year"] = d["_year"].astype(int)
    ratios = []
    for year, g in d.groupby("_year"):
        cash  = float(pd.to_numeric(g.get("CashEarningsDistribution"), errors="coerce").fillna(0).sum())
        stock = float(pd.to_numeric(g.get("StockEarningsDistribution"), errors="coerce").fillna(0).sum())
        eps = _annual_eps_complete(data, f"{year % 100:02d}")
        if eps:
            ratios.append((cash + stock) / eps)
    if not ratios:
        return None
    return round(sum(ratios) / len(ratios), 4)


def build_dividend_sheet(wb: Workbook, stock_id: str, data: dict,
                          dividend_df: pd.DataFrame | None, price_df: pd.DataFrame | None) -> None:
    ws = wb.create_sheet(DIVIDEND_SHEET)
    ws.sheet_properties.tabColor = "7A0000"

    for r in range(1, 32):
        for ci in range(1, 10):
            ws.cell(r, ci).fill = FILL_CREAM

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
    t = ws.cell(1, 1, f"{stock_id}  股利政策與配息紀錄 Dividend History")
    t.font = _f(bold=True, size=12, color="7A0000"); t.alignment = ALIGN_C
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 6

    records: list[dict] = []
    if dividend_df is not None and not dividend_df.empty:
        d = dividend_df.copy()
        d["_year"] = d.apply(_dividend_year, axis=1)
        d = d.dropna(subset=["_year"])
        d["_year"] = d["_year"].astype(int)
        for year, g in d.groupby("_year"):
            cash_dps  = float(pd.to_numeric(g.get("CashEarningsDistribution"), errors="coerce").fillna(0).sum())
            stock_dps = float(pd.to_numeric(g.get("StockEarningsDistribution"), errors="coerce").fillna(0).sum())
            exdiv_dates = []
            for col in ("CashExDividendTradingDate", "StockExDividendTradingDate"):
                if col in g:
                    exdiv_dates += pd.to_datetime(g[col], errors="coerce").dropna().dt.date.tolist()
            ref_price = _price_on_or_before(price_df, max(exdiv_dates)) if exdiv_dates else None
            eps = _annual_eps_complete(data, f"{year % 100:02d}")
            total = cash_dps + stock_dps
            records.append({
                "year": year, "cash": cash_dps, "stock": stock_dps, "total": total, "eps": eps,
                "payout": (total / eps) if eps else None,
                "ref_price": ref_price,
                "yield": (cash_dps / ref_price) if (ref_price and cash_dps) else None,
            })
        records.sort(key=lambda r: r["year"])

    _sec(ws, 3, "年度股利紀錄 (依除息年度彙總，來源：FinMind)", end_col=8)
    hdrs = ["年度", "現金股利(元)", "股票股利(元)", "合計股利(元)", "EPS(元)", "配息率", "參考股價(元)", "殖利率"]
    for j, h in enumerate(hdrs, start=1):
        c = ws.cell(4, j, h)
        c.font = _f(bold=True, size=9, color="FFFFFF"); c.fill = FILL_WINE
        c.alignment = ALIGN_C; c.border = BORDER_D
    ws.row_dimensions[4].height = 16

    R0 = 5
    for i, rec in enumerate(records):
        row = R0 + i
        vals = [rec["year"], rec["cash"], rec["stock"], rec["total"], rec["eps"],
                rec["payout"], rec["ref_price"], rec["yield"]]
        fmts = [FMT_YEAR, FMT_EPS, FMT_EPS, FMT_EPS, FMT_EPS, FMT_PCT, FMT_PRICE, FMT_PCT]
        for j, (v, fmt) in enumerate(zip(vals, fmts), start=1):
            c = ws.cell(row, j, round(v, 4) if isinstance(v, float) else v)
            c.number_format = fmt; c.alignment = ALIGN_R if j > 1 else ALIGN_C
            c.font = _f(bold=(j in (1, 4)), size=9); c.fill = FILL_CREAM; c.border = BORDER_D
        ws.row_dimensions[row].height = 15
    last_row = R0 + len(records) - 1 if records else R0

    if not records:
        nc = ws.cell(R0, 1, "無法取得 FinMind 股利資料")
        nc.font = _f(size=9, color="7A0000"); nc.alignment = ALIGN_L

    note = ws.cell(last_row + 2, 1,
                    "說明：年度依除息日彙總（同年多次配息會加總，適用季配息個股）。"
                    "配息率 = 合計股利 / 當年度 EPS；殖利率 = 現金股利 / 除息日參考股價。"
                    "EPS 僅涵蓋 Income Model 抓取範圍內、且四季資料齊全的年度（受 --start_year 影響）；"
                    "較早年度與當年度（僅部分季度已公布）皆留白，避免以不完整全年 EPS 誤算配息率。")
    note.font = _f(size=8, color="666666"); note.fill = FILL_CREAM
    ws.row_dimensions[last_row + 2].height = 13

    if records:
        cats = Reference(ws, min_col=1, min_row=R0, max_row=last_row)
        chart_row = last_row + 4

        bar = BarChart(); bar.type = "col"
        bar.add_data(Reference(ws, min_col=2, min_row=4, max_row=last_row), titles_from_data=True)
        bar.set_categories(cats)
        _style_series_bar(bar.series[0], PALETTE[0])
        _style_chart(bar, "Cash Dividend Per Share (元)", FMT_EPS, legend=False)
        ws.add_chart(bar, f"A{chart_row}")

        line = LineChart()
        line.add_data(Reference(ws, min_col=6, min_row=4, max_row=last_row), titles_from_data=True)
        line.add_data(Reference(ws, min_col=8, min_row=4, max_row=last_row), titles_from_data=True)
        line.set_categories(cats)
        _style_series_line(line.series[0], PALETTE[0])
        _style_series_line(line.series[1], ACCENT_PRICE)
        _style_chart(line, "Payout Ratio & Dividend Yield", FMT_PCT, legend=True)
        ws.add_chart(line, f"J{chart_row}")

    ws.column_dimensions["A"].width = 10
    for col in ["B", "C", "D", "E", "F", "G", "H"]:
        ws.column_dimensions[col].width = 13

# ── Build Turnover Days ───────────────────────────────────────────────────────
def build_turnover_sheet(wb: Workbook, stock_id: str, data: dict, bs_extra: dict) -> None:
    ws = wb.create_sheet(TURNOVER_SHEET)
    ws.sheet_properties.tabColor = "7A0000"
    keys = TURN_QTR_KEYS
    n = len(keys)
    last_col = 1 + n

    for r in range(1, 30):
        for ci in range(1, last_col + 2):
            ws.cell(r, ci).fill = FILL_CREAM

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    t = ws.cell(1, 1, f"{stock_id}  存貨與應收帳款週轉天數 Working Capital Turnover Days")
    t.font = _f(bold=True, size=12, color="7A0000"); t.alignment = ALIGN_C
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 6

    _sec(ws, 3, "資產負債表餘額 (億元，來源：FinMind)", end_col=last_col)
    _band_quarter_header(ws, 4, keys)
    INV_R, AR_R, AP_R = 5, 6, 7
    for row, label in [(INV_R, "存貨"), (AR_R, "應收帳款淨額"), (AP_R, "應付帳款")]:
        _band_row_label(ws, row, label)

    ws.row_dimensions[8].height = 6
    _sec(ws, 9, "週轉天數 Turnover Days (單季年化，需前一季餘額)", end_col=last_col)
    _band_quarter_header(ws, 10, keys)
    DIO_R, DSO_R, DPO_R, CCC_R = 11, 12, 13, 14
    for row, label in [(DIO_R, "存貨周轉天數 DIO"), (DSO_R, "應收帳款周轉天數 DSO"),
                        (DPO_R, "應付帳款周轉天數 DPO"), (CCC_R, "現金循環週期 CCC (DIO+DSO-DPO)")]:
        _band_row_label(ws, row, label, bold=(row == CCC_R))

    inv  = bs_extra.get("Inventories", {})
    ar   = bs_extra.get("AccountsReceivable", {})
    ap   = bs_extra.get("AccountsPayable", {})
    cogs = data.get("CostOfGoodsSold", {})
    rev  = data.get("Revenue", {})

    def _days(avg_bal: float | None, flow_q: float | None) -> float | None:
        if avg_bal is None or not flow_q:
            return None
        return avg_bal / (flow_q / UNIT * 4) * 365

    for i, key in enumerate(keys):
        col = 2 + i
        fill = _qfill(key)
        prev_key = TURN_QTR_KEYS[i - 1] if i > 0 else None
        is_hist = key in TURN_HIST_QTR_KEYS and prev_key is not None

        inv_v = inv.get(key); ar_v = ar.get(key); ap_v = ap.get(key)
        for row, v in [(INV_R, inv_v), (AR_R, ar_v), (AP_R, ap_v)]:
            c = ws.cell(row, col, round(v, 2) if v is not None else None)
            c.number_format = FMT_100M; c.alignment = ALIGN_R
            c.font = _f(size=9); c.fill = fill; c.border = BORDER_D

        dio = dso = dpo = None
        if is_hist:
            cogs_q = cogs.get(key)
            rev_q  = rev.get(key)
            if inv_v is not None and inv.get(prev_key) is not None:
                dio = _days((inv_v + inv[prev_key]) / 2, cogs_q)
            if ar_v is not None and ar.get(prev_key) is not None:
                dso = _days((ar_v + ar[prev_key]) / 2, rev_q)
            if ap_v is not None and ap.get(prev_key) is not None:
                dpo = _days((ap_v + ap[prev_key]) / 2, cogs_q)
        for row, v in [(DIO_R, dio), (DSO_R, dso), (DPO_R, dpo)]:
            c = ws.cell(row, col, round(v, 1) if v is not None else None)
            c.number_format = FMT_DAYS; c.alignment = ALIGN_R
            c.font = _f(size=9); c.fill = fill; c.border = BORDER_D
        ccc = (dio + dso - dpo) if (dio is not None and dso is not None and dpo is not None) else None
        cc = ws.cell(CCC_R, col, round(ccc, 1) if ccc is not None else None)
        cc.number_format = FMT_DAYS; cc.alignment = ALIGN_R
        cc.font = _f(bold=True, size=9); cc.fill = fill; cc.border = BORDER_HL

    note = ws.cell(15, 1,
                    "公式：週轉天數 = 平均餘額(本季、上季) / (本季成本或營收 × 4) × 365 — 單季年化法，"
                    "較 TTM 平均法所需歷史資料更短，能涵蓋更多季度。"
                    "預測欄位 (F) 留白：本模型未預測資產負債表科目（存貨/應收/應付），故無法估算未來週轉天數。")
    note.font = _f(size=8, color="666666"); note.fill = FILL_CREAM
    ws.row_dimensions[15].height = 26
    ws.merge_cells(start_row=15, start_column=1, end_row=15, end_column=last_col)

    # ── Chart ──────────────────────────────────────────────────────────────────
    cats_f = f"'{ws.title}'!$B$10:${get_column_letter(1 + n)}$10"
    ch = LineChart()
    for row, label, color in [(DIO_R, "DIO", PALETTE[0]), (DSO_R, "DSO", PALETTE[1]), (DPO_R, "DPO", PALETTE[2])]:
        ref = Reference(ws, min_col=2, max_col=1 + n, min_row=row, max_row=row)
        ch.add_data(ref, from_rows=True)
        s = ch.series[len(ch.series) - 1]
        s.tx = SeriesLabel(v=label)
        _style_series_line(s, color)
    for s in ch.series:
        try:
            s.cat = AxDataSource(strRef=StrRef(f=cats_f))
        except Exception:
            pass
    _style_chart(ch, "Turnover Days Trend", FMT_DAYS, legend=True)
    ws.add_chart(ch, "A17")

    ws.column_dimensions["A"].width = 24
    for i in range(n):
        ws.column_dimensions[get_column_letter(2 + i)].width = 10

# ── Build Raw Data ────────────────────────────────────────────────────────────
def build_raw_data_sheet(wb: Workbook, income_df: pd.DataFrame, bs_df: pd.DataFrame,
                          price_df: pd.DataFrame | None = None, per_df: pd.DataFrame | None = None,
                          dividend_df: pd.DataFrame | None = None) -> None:
    ws = wb.create_sheet(RAW_DATA_SHEET)
    ws.sheet_state = "hidden"
    col_offset = 1
    datasets = [(income_df, "Income Statement"), (bs_df, "Balance Sheet"),
                (price_df, "Stock Price"), (per_df, "PER-PBR"), (dividend_df, "Dividend")]
    for df, label in datasets:
        if df is None or df.empty:
            continue
        ws.cell(1, col_offset, label).font = _f(bold=True, size=9)
        for ci, col in enumerate(df.columns, start=col_offset):
            ws.cell(2, ci, col).font = _f(bold=True, size=9)
        for ri, row in enumerate(df.itertuples(index=False), start=3):
            for ci, val in enumerate(row, start=col_offset):
                ws.cell(ri, ci, val).font = _f(size=9)
        col_offset += len(df.columns) + 2

# ── Process & Main ────────────────────────────────────────────────────────────
def process_stock(stock_id: str, start_year: int) -> None:
    print(f"\n[{stock_id}] Downloading income statement...")
    income_df = fetch_finmind(stock_id, "TaiwanStockFinancialStatements", start_year)
    time.sleep(REQUEST_DELAY)

    print(f"[{stock_id}] Downloading balance sheet...")
    bs_df = fetch_finmind(stock_id, "TaiwanStockBalanceSheet", start_year)
    time.sleep(REQUEST_DELAY)

    data = build_data(income_df, bs_df)
    if not data.get("OrdinaryShare"):
        print(f"  WARNING [{stock_id}]: No OrdinaryShare data — EPS formulas may be empty.")
    bs_extra = build_bs_extra_data(bs_df)

    # Valuation sheets need a longer market-data history than --start_year provides,
    # so they fetch independently and degrade gracefully (warn + blank sheet, not crash).
    this_year = date.today().year
    print(f"[{stock_id}] Downloading price history...")
    price_df = fetch_finmind_optional(stock_id, "TaiwanStockPrice", this_year - BAND_LOOKBACK_YEARS)
    time.sleep(REQUEST_DELAY)

    print(f"[{stock_id}] Downloading PER/PBR history...")
    per_df = fetch_finmind_optional(stock_id, "TaiwanStockPER", this_year - BAND_LOOKBACK_YEARS)
    time.sleep(REQUEST_DELAY)

    print(f"[{stock_id}] Downloading dividend history...")
    dividend_df = fetch_finmind_optional(stock_id, "TaiwanStockDividend", this_year - DIVIDEND_LOOKBACK_YEARS)

    payout_def = payout_ratio_default(dividend_df, data)

    print(f"[{stock_id}] Building Excel...")
    wb     = Workbook()
    layout = build_model_sheet(wb, stock_id, data)
    build_assumptions_sheet(wb, stock_id, data, payout_def, BAND_LOOKBACK_YEARS)
    build_revenue_build_sheet(wb, stock_id, layout["item_row"], data)
    build_dashboard_sheet(wb, stock_id, layout)
    build_pe_band_sheet(wb, stock_id, layout["item_row"], data, price_df, per_df, BAND_LOOKBACK_YEARS)
    build_pb_band_sheet(wb, stock_id, layout["item_row"], data, bs_extra, price_df, per_df, BAND_LOOKBACK_YEARS)
    build_scenario_analysis_sheet(wb, stock_id, layout["item_row"])
    build_backtest_sheet(wb, stock_id, price_df, per_df, BAND_LOOKBACK_YEARS)
    build_dividend_sheet(wb, stock_id, data, dividend_df, price_df)
    build_turnover_sheet(wb, stock_id, data, bs_extra)
    build_raw_data_sheet(wb, income_df, bs_df, price_df, per_df, dividend_df)

    try:
        wb.calculation.calcMode = "auto"
        wb.calculation.fullCalcOnLoad = True
    except Exception:
        pass

    out = OUTPUT_DIR / f"{stock_id}_income_model.xlsx"
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))
    print(f"[{stock_id}] Saved -> {out.relative_to(PROJECT_ROOT)}")


def main() -> None:
    args   = parse_args()
    stocks = [s.strip() for s in args.stocks if s.strip()]
    if not stocks:
        print("Error: provide at least one stock ID", file=sys.stderr)
        sys.exit(1)
    print(f"Stocks: {', '.join(stocks)}  start_year: {args.start_year}")
    for stock_id in stocks:
        try:
            process_stock(stock_id, args.start_year)
        except RuntimeError as e:
            print(f"  ERROR {stock_id}: {e}", file=sys.stderr)
        except requests.RequestException as e:
            print(f"  NETWORK ERROR {stock_id}: {e}", file=sys.stderr)
    print("\nDone.")


if __name__ == "__main__":
    main()
