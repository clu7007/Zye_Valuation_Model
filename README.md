# Taiwan Stock Financial Model Generator

輸入台股股票代碼，自動從 [FinMind](https://finmindtrade.com/) 抓取財報／股價／股利資料，產出一份投行格式的 Excel 分析模型。不需要申請帳號或 API 金鑰。

## 功能

- 輸入股票代碼即自動下載歷史財報、股價、股利資料並建模，支援一次跑多支股票
- **Income Model**：季度損益表歷史（2024 至今）+ 2026Q2F–2029F 預測
- **Assumptions**：毛利率／費用率／稅率等假設輸入（黃色格＝可編輯）
- **Revenue Build**：Volume／ASP／Mix／FX 營收驅動因子預測，是全模型預測營收的唯一來源；內建 5 條空白產品線讓你自行輸入名稱與佔比做產品結構試算
- **Dashboard**：關鍵指標總覽
- **PE Band / PB Band**：本益比／股價淨值比歷史區間圖
- **Scenario Analysis**：樂觀／中性／悲觀情境分析
- **Backtest**：預測回測
- **Dividend History**、**Turnover Days**：股利發放歷史、應收/存貨/應付週轉天數
- **Raw Data**：FinMind 原始資料存底
- 全表公式自動串接（Revenue Build → Income Model → Assumptions → Dashboard），只有黃色格需要手動輸入，其餘改動上游數字會自動連動

## 安裝

需要 Python 3.12（`py -0p` 可檢查已安裝版本）。

```bash
git clone <repo-url>
cd <repo-folder>

python -m venv .venv
.venv\Scripts\activate        # Windows
# source .venv/bin/activate   # macOS / Linux

pip install -r requirements.txt
```

## 使用方式

**方法一：雙擊批次檔（Windows，最簡單）**

雙擊根目錄的 `run_income_model.bat`，依提示輸入股票代碼：

```
Stock codes: 2330
Stock codes: 2330 3532 2454        # 一次多支股票，空格分隔
```

完成後自動開啟 `output\models\` 資料夾。

**方法二：指令列**

```bash
python scripts/build_income_statement_model.py --stocks 2330
python scripts/build_income_statement_model.py --stocks 2330 3532 2454 --start_year 2024
```

- `--stocks`：一或多個台股代碼（必填）
- `--start_year`：歷史資料起始年份（預設 2022）

輸出固定在 `output/models/{股票代碼}_income_model.xlsx`。

**既有舊版模型檔，只想補上 Revenue Build 分頁**（新產生的模型檔已內建，不需要跑這支）：

```bash
python scripts/add_revenue_build.py --stock 2330
```

會先備份原檔（`{股票代碼}_income_model.backup_{時間戳}.xlsx`）再原地修改。

## 已知限制

- FinMind 免費額度有流量限制，短時間查太多股票可能被限速，稍等再試即可
- Revenue Build 的產品結構試算區塊預設空白，需自行填寫，不填不影響主要營收預測
- 目前只覆蓋台股上市櫃代碼，資料完整度取決於 FinMind 對該檔股票的覆蓋範圍

## 專案架構

```text
.
├─ run_income_model.bat                     # 主工具的雙擊入口
├─ scripts/
│  ├─ build_income_statement_model.py       # 主工具：FinMind → Excel 模型
│  ├─ add_revenue_build.py                  # 舊模型檔的一次性升級腳本
│  ├─ main.py                               # legacy：本地 CSV/Excel 匯入工具（見下）
│  └─ data_sources/ financial_analysis/ excel_export/
├─ data/raw/                                # 放本地 CSV/Excel 財報（legacy 工具用）
├─ output/                                  # 產生的 Excel／CSV，不進版控，執行時自動建立
└─ notes/project_log.md
```

## Legacy：本地 CSV／Excel 匯入工具

串接 FinMind 之前的第一版 MVP，改用本地整理好的財報 CSV／Excel 作資料來源，離線測試流程用：

```bash
python scripts/main.py --stock_id 2330 --source csv --input_path data/raw/2330_quarterly.csv
python scripts/main.py --stock_id 2330 --source excel --input_path data/raw/2330_quarterly.xlsx
python scripts/main.py --stock_id 2330 --source sample   # 假資料，僅供流程測試
```

檔案格式：第一欄 `Item`，其餘為季度欄（`2024Q1` 起），必要列項目為 `Revenue`、`Cost of revenue`、`Gross profit`、`Operating expenses`、`Operating income`、`Non-operating income and expenses`、`Profit before tax`、`Tax expense`、`Net income`、`Net income attributable to parent company`、`EPS`。

加上 `--debug --save_raw` 可印出中間計算結果並輸出稽核用 CSV 到 `output/debug/`。輸出位置：`output/excel/{stock_id}_quarterly_financials.xlsx`
