"""Stage 1 upgrade: Revenue Driver Build.

Adds a "Revenue Build" sheet (Historical Revenue / Revenue Driver Build /
Segment Mix Framework / Model Checks) to an *already-generated*
{stock}_income_model.xlsx, and rewires Income Model's 26Q2F-26Q4F / 2027F-2029F
Revenue cells plus the Assumptions sheet's revenue inputs to source from it.

This does NOT regenerate the workbook — it opens the existing file, backs it
up, patches it in place, and saves. It does not touch Balance Sheet, Cash
Flow, DCF, or Scenario logic (out of scope for Stage 1), and does not modify
Dashboard / PE Band / PB Band / Dividend History / Turnover Days formulas
(only adds one explanatory note to PE Band about the NTM EPS proxy).

Usage:
    python scripts/add_revenue_build.py --stock 2059
"""
from __future__ import annotations

import argparse
import shutil
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from build_income_statement_model import (
    COLUMNS, INCOME_MODEL_SHEET, ASSUMPTIONS_SHEET, PE_BAND_SHEET,
    A_Q_REV_R, A_Q_COL, A_ANN_GROWTH_R, A_ANN_COL,
    FILL_WINE, FILL_CREAM, FILL_INPUT, FILL_FCAST_QTR,
    ALIGN_C, ALIGN_L, ALIGN_L_WRAP, ALIGN_R, BORDER_D, BORDER_HL,
    FMT_100M, FMT_PCT,
    _f, _sec, _cfill,
)

REVENUE_BUILD_SHEET = "Revenue Build"
PROJECT_ROOT = Path(__file__).resolve().parents[1]
OUTPUT_DIR = PROJECT_ROOT / "output" / "models"

# Income Model row numbers (INCOME_ITEMS order in build_income_statement_model.py).
IM_REVENUE_ROW = 3
IM_EPS_ROW = 14

# Revenue Build's own column layout = Income Model's COLUMNS minus the 4 hidden
# 24Q helper columns (Revenue Build doesn't need 2024's quarterly split).
RB_COLUMNS = [(hdr, key, ctype) for hdr, key, ctype in COLUMNS if ctype != "qtr_helper"]
RB_COL_IDX = {key: 2 + i for i, (_, key, _) in enumerate(RB_COLUMNS)}
RB_COL_LTR = {key: get_column_letter(v) for key, v in RB_COL_IDX.items()}
RB_CTYPE   = {key: ctype for _, key, ctype in RB_COLUMNS}
RB_LAST_COL = 1 + len(RB_COLUMNS)

# Income Model's own columns (with the hidden-24Q gap), for cross-sheet refs.
IM_COL_IDX = {key: 2 + i for i, (_, key, _) in enumerate(COLUMNS)}
IM_COL_LTR = {key: get_column_letter(v) for key, v in IM_COL_IDX.items()}

HIST_KEYS      = ["2024", "25Q1", "25Q2", "25Q3", "25Q4", "2025", "26Q1"]
DRIVER_Q_KEYS  = ["26Q2F", "26Q3F", "26Q4F"]
DRIVER_A_KEYS  = ["2027F", "2028F", "2029F"]
DRIVER_KEYS    = DRIVER_Q_KEYS + DRIVER_A_KEYS
Q_PRIOR_KEY    = {"26Q2F": "25Q2", "26Q3F": "25Q3", "26Q4F": "25Q4"}   # same-quarter-last-year base
ANN_PRIOR_COL  = {"2027F": "2026F", "2028F": "2027F", "2029F": "2028F"}  # prior column WITHIN Revenue Build's own row 18


def _backup(path: Path) -> Path:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = path.with_name(f"{path.stem}.backup_{ts}{path.suffix}")
    shutil.copy2(path, backup_path)
    return backup_path


def _row_label(ws, row: int, text: str, bold: bool = False, fill=FILL_CREAM) -> None:
    # A leading '=', '+', '-', or '@' makes Excel/openpyxl treat a plain-text label as
    # a FORMULA instead of a string (this bit us once already — Excel then fails to
    # parse the Chinese text as a formula and the cell shows garbled/broken content).
    # Guard against it here so no future label can reintroduce that bug.
    if text and text[0] in "=+-@":
        raise ValueError(f"Row {row} label starts with '{text[0]}' — Excel would read this "
                          f"as a formula, not text: {text!r}")
    c = ws.cell(row, 1, text)
    c.font = _f(bold=bold, size=9); c.fill = fill
    # Wrapped + taller row: these bilingual labels are too long for column A at some
    # widths, and since column B always has real content, Excel clips (rather than
    # overflows) anything too wide for the column — wrap is the only way to guarantee
    # it's never cut off without hand-tuning column widths per label.
    c.alignment = ALIGN_L_WRAP; c.border = BORDER_HL if bold else BORDER_D
    ws.row_dimensions[row].height = 28


def _write_header(ws, row: int) -> None:
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


def _data_cell(ws, row: int, key: str, value=None, fmt: str = FMT_100M,
               fill=None, bold: bool = False, editable: bool = False) -> None:
    ci = RB_COL_IDX[key]
    c = ws.cell(row, ci, value)
    c.number_format = fmt; c.alignment = ALIGN_R
    c.font = _f(bold=bold, size=9)
    c.fill = fill if fill is not None else (FILL_INPUT if editable else FILL_CREAM)
    c.border = BORDER_HL if bold else BORDER_D


def _dash_cell(ws, row: int, key: str) -> None:
    ci = RB_COL_IDX[key]
    c = ws.cell(row, ci, "-")
    c.font = _f(size=9, color="9A9186"); c.fill = FILL_CREAM
    c.alignment = ALIGN_C; c.border = BORDER_D


def _na_cell(ws, row: int, key: str) -> None:
    ci = RB_COL_IDX[key]
    c = ws.cell(row, ci, "N/A")
    c.font = _f(size=9, color="9A9186"); c.fill = FILL_CREAM
    c.alignment = ALIGN_C; c.border = BORDER_D


def add_revenue_build(xlsx_path: Path, stock_id: str) -> Path:
    backup_path = _backup(xlsx_path)
    wb = load_workbook(xlsx_path)

    if REVENUE_BUILD_SHEET in wb.sheetnames:
        raise RuntimeError(f"'{REVENUE_BUILD_SHEET}' already exists in {xlsx_path.name} — "
                            "aborting so we don't silently double-patch an already-upgraded file.")

    ws_im = wb[INCOME_MODEL_SHEET]
    ws_a  = wb[ASSUMPTIONS_SHEET]

    # ── Smart defaults, read from the workbook itself (not re-fetched) ────────
    # Quarterly driver placeholder: most recent actual YoY (26Q1 vs 25Q1) — there is
    # no pre-existing quarterly revenue assumption to preserve (it was blank), so the
    # most defensible "neutral" starting point is the latest real trend, not 0%.
    q1_26 = ws_im.cell(IM_REVENUE_ROW, IM_COL_IDX["26Q1"]).value
    q1_25 = ws_im.cell(IM_REVENUE_ROW, IM_COL_IDX["25Q1"]).value
    qtr_growth_default = (q1_26 / q1_25 - 1) if (isinstance(q1_26, (int, float)) and
                                                  isinstance(q1_25, (int, float)) and q1_25) else 0.0
    # Annual driver placeholder: reuse the EXISTING annual growth assumption values
    # verbatim, so Stage-1 output reproduces whatever forecast already existed there.
    ann_growth_default = {
        key: ws_a[f"{A_ANN_COL[key]}{A_ANN_GROWTH_R}"].value or 0.0
        for key in DRIVER_A_KEYS
    }

    # ═══════════════════════════════════════════════════════════════════════
    # Revenue Build sheet
    # ═══════════════════════════════════════════════════════════════════════
    ws = wb.create_sheet(REVENUE_BUILD_SHEET, index=wb.sheetnames.index(ASSUMPTIONS_SHEET) + 1)
    ws.sheet_properties.tabColor = "7A0000"

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
    _write_header(ws, 4)
    _row_label(ws, 5, "營業收入淨額 Revenue (億元)", bold=True)
    _row_label(ws, 6, "YoY 成長率 (實際)")
    for key in HIST_KEYS:
        im_col = IM_COL_LTR[key]
        _data_cell(ws, 5, key, f"=IFERROR('{INCOME_MODEL_SHEET}'!{im_col}{IM_REVENUE_ROW},\"\")",
                   fmt=FMT_100M, fill=_cfill(RB_CTYPE[key]), bold=True)
    for key in DRIVER_KEYS:
        _dash_cell(ws, 5, key)
    yoy_pairs = [
        ("25Q1", f"'{INCOME_MODEL_SHEET}'!{IM_COL_LTR['24Q1']}{IM_REVENUE_ROW}"),
        ("25Q2", f"'{INCOME_MODEL_SHEET}'!{IM_COL_LTR['24Q2']}{IM_REVENUE_ROW}"),
        ("25Q3", f"'{INCOME_MODEL_SHEET}'!{IM_COL_LTR['24Q3']}{IM_REVENUE_ROW}"),
        ("25Q4", f"'{INCOME_MODEL_SHEET}'!{IM_COL_LTR['24Q4']}{IM_REVENUE_ROW}"),
        ("2025", f"{RB_COL_LTR['2024']}5"),
        ("26Q1", f"{RB_COL_LTR['25Q1']}5"),
    ]
    _dash_cell(ws, 6, "2024")
    for key, prior_ref in yoy_pairs:
        col = RB_COL_LTR[key]
        _data_cell(ws, 6, key, f'=IFERROR({col}5/{prior_ref}-1,"")', fmt=FMT_PCT)
    for key in DRIVER_KEYS:
        _dash_cell(ws, 6, key)

    # ── B. Revenue Driver Build ───────────────────────────────────────────
    ws.row_dimensions[7].height = 6
    _sec(ws, 8, "B. 營收驅動因子預測 Revenue Driver Build — 黃色為可編輯假設", end_col=RB_LAST_COL)
    _write_header(ws, 9)

    for key in HIST_KEYS + ["2026F"]:
        for row in (10, 11, 12, 13, 15):
            _dash_cell(ws, row, key)

    _row_label(ws, 10, "出貨量／銷量成長率 Volume Growth YoY (Volume 貢獻度)")
    _row_label(ws, 11, "平均售價成長率 ASP Growth YoY (ASP 貢獻度)")
    _row_label(ws, 12, "產品組合影響 Product Mix Impact (Mix 貢獻度)")
    _row_label(ws, 13, "匯率／其他影響 FX/Other Impact (匯率/其他貢獻度)")
    _row_label(ws, 14, "隱含營收成長率 Implied Revenue Growth YoY (Driver 加總試算)", bold=True)
    _row_label(ws, 15, "人工覆蓋營收成長率 Manual Override YoY (留空 = 採用 Driver 推導)")
    _row_label(ws, 16, "實際採用營收成長率 Applied Growth YoY (實際套用公式)", bold=True)
    _row_label(ws, 18, "營收預測 Forecast Revenue (億元)", bold=True)

    for key in DRIVER_Q_KEYS + DRIVER_A_KEYS:
        vol_default = qtr_growth_default if key in DRIVER_Q_KEYS else ann_growth_default[key]
        _data_cell(ws, 10, key, round(vol_default, 4), fmt=FMT_PCT, editable=True)
        _data_cell(ws, 11, key, 0.0, fmt=FMT_PCT, editable=True)
        _data_cell(ws, 12, key, 0.0, fmt=FMT_PCT, editable=True)
        _data_cell(ws, 13, key, 0.0, fmt=FMT_PCT, editable=True)
        col = RB_COL_LTR[key]
        _data_cell(ws, 14, key, f"=IFERROR((1+{col}10)*(1+{col}11)*(1+{col}12)*(1+{col}13)-1,\"\")",
                   fmt=FMT_PCT, fill=FILL_FCAST_QTR, bold=True)
        _data_cell(ws, 15, key, None, fmt=FMT_PCT, editable=True)
        _data_cell(ws, 16, key, f'=IFERROR(IF({col}15<>"",{col}15,{col}14),"")',
                   fmt=FMT_PCT, fill=FILL_FCAST_QTR, bold=True)

    for key in HIST_KEYS:
        _dash_cell(ws, 18, key)
    for key in DRIVER_Q_KEYS:
        col, prior_col = RB_COL_LTR[key], RB_COL_LTR[Q_PRIOR_KEY[key]]
        _data_cell(ws, 18, key, f'=IFERROR({prior_col}5*(1+{col}16),"")',
                   fmt=FMT_100M, fill=FILL_FCAST_QTR, bold=True)
    q26_1, q26_2, q26_3, q26_4 = RB_COL_LTR["26Q1"], RB_COL_LTR["26Q2F"], RB_COL_LTR["26Q3F"], RB_COL_LTR["26Q4F"]
    _data_cell(ws, 18, "2026F", f'=IFERROR({q26_1}5+{q26_2}18+{q26_3}18+{q26_4}18,"")',
               fmt=FMT_100M, fill=_cfill("annual_fcast"), bold=True)
    for key in DRIVER_A_KEYS:
        col, prior_col = RB_COL_LTR[key], RB_COL_LTR[ANN_PRIOR_COL[key]]
        _data_cell(ws, 18, key, f'=IFERROR({prior_col}18*(1+{col}16),"")',
                   fmt=FMT_100M, fill=_cfill("annual_fcast"), bold=True)

    ws.row_dimensions[17].height = 6
    ws.row_dimensions[19].height = 6

    # ── C. Segment / Product Mix Framework ────────────────────────────────
    # Every company has a different product mix / weighting / gross margin by
    # segment, so this section is a free-form "what-if" calculator: fill in your
    # own view of Mix %, Volume, ASP and Gross Margin per segment, and row 36
    # shows the resulting blended gross margin for reference. It does not
    # automatically overwrite Assumptions' Gross Margin cells (those stay a
    # direct, freely-editable input as before) — copy the number across yourself
    # if you want to use it, so nothing here can silently change existing output.
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

    _write_header(ws, 22)
    for si, name_row in enumerate(SEG_NAME_ROWS):
        _row_label(ws, name_row, f"產品線 {si + 1} 名稱 Segment {si + 1} Name（請自行輸入）", bold=True)
        ws.merge_cells(start_row=name_row, start_column=2, end_row=name_row, end_column=RB_LAST_COL)
        name_cell = ws.cell(name_row, 2)
        name_cell.font = _f(size=9, bold=True)
        name_cell.fill = FILL_INPUT
        name_cell.alignment = ALIGN_L
        name_cell.border = BORDER_D
        for j, sub in enumerate(SEG_SUBROWS):
            row = name_row + 1 + j
            _row_label(ws, row, sub)
            for key in HIST_KEYS + ["2026F"]:
                _na_cell(ws, row, key)
            for key in DRIVER_Q_KEYS + DRIVER_A_KEYS:
                _data_cell(ws, row, key, None, fmt=FMT_PCT, editable=True)

    _row_label(ws, MIX_TOTAL_ROW, "產品結構佔比合計 Segment Mix Total (應 = 100% 或 N/A)", bold=True)
    for key in HIST_KEYS + ["2026F"]:
        _na_cell(ws, MIX_TOTAL_ROW, key)
    for key in DRIVER_Q_KEYS + DRIVER_A_KEYS:
        col = RB_COL_LTR[key]
        mix_refs = [f"{col}{r}" for r in SEG_MIX_ROWS]
        f = f'=IF(COUNT({",".join(mix_refs)})=0,"N/A",{"+".join(mix_refs)})'
        _data_cell(ws, MIX_TOTAL_ROW, key, f, fmt=FMT_PCT, fill=FILL_FCAST_QTR, bold=True)

    _row_label(ws, BLEND_MARGIN_ROW, "加權毛利率試算 Blended Gross Margin (Mix % x Segment 毛利率)", bold=True)
    for key in HIST_KEYS + ["2026F"]:
        _na_cell(ws, BLEND_MARGIN_ROW, key)
    for key in DRIVER_Q_KEYS + DRIVER_A_KEYS:
        col = RB_COL_LTR[key]
        mix_refs = [f"{col}{r}" for r in SEG_MIX_ROWS]
        weighted_terms = [f"{col}{mr}*{col}{gr}" for mr, gr in zip(SEG_MIX_ROWS, SEG_MARGIN_ROWS)]
        denom = f'SUM({",".join(mix_refs)})'
        f = (f'=IFERROR(IF({denom}=0,"N/A",'
             f'({"+".join(weighted_terms)})/{denom}),"N/A")')
        _data_cell(ws, BLEND_MARGIN_ROW, key, f, fmt=FMT_PCT, fill=FILL_FCAST_QTR, bold=True)

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
    im_m, im_n, im_o = IM_COL_LTR["26Q2F"], IM_COL_LTR["26Q3F"], IM_COL_LTR["26Q4F"]
    im_q, im_r, im_s = IM_COL_LTR["2027F"], IM_COL_LTR["2028F"], IM_COL_LTR["2029F"]

    checks = [
        (CHECK_ROW0 + 0, "1. 季度加總檢查 Quarterly Sum Check：2026F = 26Q1 + Q2F + Q3F + Q4F",
             f'=IF(ABS({a26}18-({q26_1}5+{q2}18+{q3}18+{q4}18))<0.01,"OK","ERROR")'),
        (CHECK_ROW0 + 1, "2. 產品結構佔比檢查 Segment Mix Check：合計 = 100% 或 N/A",
             # Split into two contiguous sub-ranges (I:K, M:O) — column L (2026F) sits
             # between them and is always "N/A" (it's an aggregate column, not a driver
             # period), so a single I:O span would corrupt the count/check below.
             f'=IF(COUNTIF({q2}{MIX_TOTAL_ROW}:{q4}{MIX_TOTAL_ROW},"N/A")+COUNTIF({m1}{MIX_TOTAL_ROW}:{o1}{MIX_TOTAL_ROW},"N/A")=6,"N/A",'
             f'IF(SUMPRODUCT(ISNUMBER({q2}{MIX_TOTAL_ROW}:{q4}{MIX_TOTAL_ROW})*(ABS(N({q2}{MIX_TOTAL_ROW}:{q4}{MIX_TOTAL_ROW})-1)>0.01))'
             f'+SUMPRODUCT(ISNUMBER({m1}{MIX_TOTAL_ROW}:{o1}{MIX_TOTAL_ROW})*(ABS(N({m1}{MIX_TOTAL_ROW}:{o1}{MIX_TOTAL_ROW})-1)>0.01))>0,"ERROR","OK"))'),
        (CHECK_ROW0 + 2, "3. 營收連結檢查 Revenue Link Check：Revenue Build 與 Income Model 一致",
             f'=IF(AND(SUMPRODUCT(ABS(({q2}18:{q4}18)-(\'{INCOME_MODEL_SHEET}\'!{im_m}{IM_REVENUE_ROW}:{im_o}{IM_REVENUE_ROW})))<0.01,'
             f'SUMPRODUCT(ABS(({m1}18:{o1}18)-(\'{INCOME_MODEL_SHEET}\'!{im_q}{IM_REVENUE_ROW}:{im_s}{IM_REVENUE_ROW})))<0.01),"OK","ERROR")'),
        (CHECK_ROW0 + 3, "4. 成長率計算檢查 Growth Calculation Check：Driver 公式無錯誤值",
             f'=IF(SUMPRODUCT(--ISERROR({q2}14:{q4}14))+SUMPRODUCT(--ISERROR({m1}14:{o1}14))=0,"OK","ERROR")'),
    ]
    for row, label, formula in checks:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=CHECK_RESULT_COL - 1)
        _row_label(ws, row, label, bold=True)
        c = ws.cell(row, CHECK_RESULT_COL, formula)
        c.font = _f(bold=True, size=9); c.fill = FILL_CREAM
        c.alignment = ALIGN_C; c.border = BORDER_HL

    for row, *_ in checks:
        rng = f"{get_column_letter(CHECK_RESULT_COL)}{row}"
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{rng}="OK"'], font=Font(color="1E7B34", bold=True)))
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{rng}="ERROR"'], font=Font(color="B00020", bold=True)))

    # CHECK_ROW0+0..+3 are the 4 check rows themselves — the spacer must be the row
    # AFTER that block (CHECK_ROW0+4), not "NOTE2_ROW-1" computed from a NOTE2_ROW
    # that was only 1 row past the last check: that collided with (and silently
    # re-shrank) the last check row's own height.
    NOTE2_ROW = CHECK_ROW0 + 5
    ws.row_dimensions[CHECK_ROW0 + 4].height = 6
    ws.merge_cells(start_row=NOTE2_ROW, start_column=1, end_row=NOTE2_ROW, end_column=RB_LAST_COL)
    note2 = ws.cell(NOTE2_ROW, 1,
        "說明：Revenue Build 為 Income Model 未來營收的唯一來源 (Single Source of Truth)。"
        "季度預測 = 去年同期營收 x (1+實際採用成長率)；年度預測 = 前一年度營收 x (1+實際採用成長率)。"
        f"季度出貨量成長率初始值 = 最近一期實際 YoY（{HIST_KEYS[-1]} 對比去年同季，因原模型該欄位原為空白，"
        "以近期實際趨勢作為中性起點）；年度出貨量成長率初始值 = 原 Assumptions 營收成長率假設（維持原預測不變）。"
        "平均售價／產品組合／匯率初始值均為 0%。如需調整，請直接修改本表黃色格子，或於「人工覆蓋」列強制指定成長率。")
    note2.font = _f(size=8, color="7A0000"); note2.fill = FILL_CREAM
    ws.row_dimensions[NOTE2_ROW].height = 26

    ws.column_dimensions["A"].width = 42
    for _, key, _c in RB_COLUMNS:
        ws.column_dimensions[RB_COL_LTR[key]].width = 11
    ws.freeze_panes = "B5"

    # ═══════════════════════════════════════════════════════════════════════
    # Rewire Income Model — Revenue row only
    # ═══════════════════════════════════════════════════════════════════════
    for key, im_key in [("26Q2F", "26Q2F"), ("26Q3F", "26Q3F"), ("26Q4F", "26Q4F")]:
        cell = ws_im.cell(IM_REVENUE_ROW, IM_COL_IDX[im_key])
        cell.value = f"=IFERROR('{REVENUE_BUILD_SHEET}'!{RB_COL_LTR[key]}18,\"\")"
    for key in DRIVER_A_KEYS:
        cell = ws_im.cell(IM_REVENUE_ROW, IM_COL_IDX[key])
        cell.value = f"=IFERROR('{REVENUE_BUILD_SHEET}'!{RB_COL_LTR[key]}18,\"\")"
    # 2026F (col P) formula untouched — already =IFERROR(L3+M3+N3+O3,"") and now
    # correctly sums 26Q1 actual + the 3 now-live Revenue Build forecast cells.

    # ═══════════════════════════════════════════════════════════════════════
    # Rewire Assumptions — revenue cells become read-only, Revenue-Build-linked
    # ═══════════════════════════════════════════════════════════════════════
    ws_a.cell(A_Q_REV_R, 1, "營業收入淨額 (億元) — 由 Revenue Build 計算")
    for key, col in A_Q_COL.items():
        c = ws_a[f"{col}{A_Q_REV_R}"]
        c.value = f"=IFERROR('{REVENUE_BUILD_SHEET}'!{RB_COL_LTR[key]}18,\"\")"
        c.fill = FILL_CREAM

    ws_a.cell(A_ANN_GROWTH_R, 1, "營收成長率 YoY（由 Revenue Build 計算）")
    for key, col in A_ANN_COL.items():
        c = ws_a[f"{col}{A_ANN_GROWTH_R}"]
        c.value = f"=IFERROR('{REVENUE_BUILD_SHEET}'!{RB_COL_LTR[key]}16,\"\")"
        c.fill = FILL_CREAM

    ws_a.merge_cells(start_row=13, start_column=1, end_row=13, end_column=7)
    note_cell = ws_a.cell(13, 1,
        "營收成長率由 Revenue Build 中出貨量、平均售價、產品組合、匯率等營運驅動因子推導；"
        "如需人工覆蓋，請於 Revenue Build 的「人工覆蓋」欄位輸入。")
    note_cell.font = _f(size=8, color="7A0000"); note_cell.fill = FILL_CREAM
    ws_a.row_dimensions[13].height = 26

    # ═══════════════════════════════════════════════════════════════════════
    # PE Band — additive note only, no formula changes
    # ═══════════════════════════════════════════════════════════════════════
    if PE_BAND_SHEET in wb.sheetnames:
        ws_pe = wb[PE_BAND_SHEET]
        existing = ws_pe.cell(8, 1).value or ""
        ws_pe.cell(8, 1, existing + " NTM EPS 於超出季度預測範圍時（如 27Q1F），"
                   "以次一年度 EPS 預測 ÷4 作為 proxy，待建立完整季度預測後可替換；"
                   "此次 Revenue Build 上線後，26Q2F-26Q4F EPS 已由正常營收推導，Forward P/E 應已回到合理區間。")

    try:
        wb.calculation.calcMode = "auto"
        wb.calculation.fullCalcOnLoad = True
    except Exception:
        pass

    wb.save(xlsx_path)
    return backup_path


def main() -> None:
    p = argparse.ArgumentParser()
    p.add_argument("--stock", required=True)
    args = p.parse_args()
    xlsx_path = OUTPUT_DIR / f"{args.stock}_income_model.xlsx"
    if not xlsx_path.exists():
        raise SystemExit(f"Not found: {xlsx_path}")
    backup = add_revenue_build(xlsx_path, args.stock)
    print(f"Backup saved -> {backup.relative_to(PROJECT_ROOT)}")
    print(f"Upgraded -> {xlsx_path.relative_to(PROJECT_ROOT)}")


if __name__ == "__main__":
    main()
