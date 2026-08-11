# Taiwan Stock Financial Model Generator

輸入台股股票代碼，自動從 [FinMind](https://finmindtrade.com/) 抓取財報 / 股價 / 股利資料，產出一份投行格式的 Excel 分析模型。不需要申請帳號或 API 金鑰。

## 這個工具做什麼

```
python scripts/build_income_statement_model.py --stocks 2330
```

會產出 `output/models/2330_income_model.xlsx`，內含：

| 分頁 | 內容 |
|---|---|
| Income Model | 季度損益表歷史（含 2024 至今）+ 2026Q2F–2029F 預測 |
| Assumptions | 毛利率／費用率／稅率等假設輸入（黃色格可編輯） |
| Revenue Build | 營收驅動因子預測（Volume／ASP／Mix／FX），是預測營收的唯一來源；也內建一個可自訂的產品結構試算框架（5 條產品線空格，自己輸入名稱與佔比） |
| Dashboard | 關鍵指標總覽 |
| PE Band / PB Band | 本益比／股價淨值比歷史區間圖 |
| Scenario Analysis | 情境分析（樂觀／中性／悲觀） |
| Backtest | 預測回測 |
| Dividend History | 股利發放歷史 |
| Turnover Days | 應收/存貨/應付週轉天數 |
| Raw Data | FinMind 原始資料存底 |

多年份的歷史資料 + 公式全部自動串接（Revenue Build → Income Model → Assumptions → Dashboard 等），黃色底色的儲存格才是需要手動輸入的假設，其餘都是公式，改動上游數字會自動連動。

## 快速開始

**需要**：Python 3.12（`py -0p` 可以檢查已安裝的版本）

```bash
git clone <repo-url>
cd <repo-folder>

python -m venv .venv
.venv\Scripts\activate        # Windows
# source .venv/bin/activate   # macOS / Linux

pip install -r requirements.txt
```

### 方法一：雙擊批次檔（Windows，最簡單）

雙擊根目錄的 **`run_income_model.bat`**，依提示輸入股票代碼即可：

```
Stock codes: 2330
Stock codes: 2330 3532 2454        # 支援一次多支股票，空格分隔
```

完成後會自動開啟 `output\models\` 資料夾。

### 方法二：指令列

```bash
python scripts/build_income_statement_model.py --stocks 2330
python scripts/build_income_statement_model.py --stocks 2330 3532 2454 --start_year 2024
```

- `--stocks`：一或多個台股代碼（必填）
- `--start_year`：歷史資料起始年份（預設 2022）

輸出固定在 `output/models/{股票代碼}_income_model.xlsx`。

### 已有舊版模型檔，只想補上 Revenue Build 分頁

`scripts/add_revenue_build.py` 是給「已經用舊版腳本產生、還沒有 Revenue Build 分頁」的既有 `.xlsx` 檔案用的一次性升級腳本（新產生的模型檔已經內建 Revenue Build，不需要跑這支）：

```bash
python scripts/add_revenue_build.py --stock 2330
```

會先備份原檔（`{股票代碼}_income_model.backup_{時間戳}.xlsx`）再原地修改。

## 已知限制

- 資料來源是 FinMind 公開端點，免費額度有流量限制；短時間內查太多股票可能會被限速，稍等再試即可。
- Revenue Build 的「產品結構試算」區塊（5 條產品線）預設是空白，需要你依公司實際情況自己填寫產品線名稱與佔比；不填不影響主要的營收預測公式。
- 目前只覆蓋台股（`.tw` 上市櫃代碼），資料完整度取決於 FinMind 該檔股票的覆蓋範圍。

---

## 舊版工具：本地 CSV / Excel 財報匯入（legacy）

在串接 FinMind 之前的第一版 MVP，改用本地 CSV／Excel 檔案作為資料來源。目前主力工具是上面的 FinMind 版本；這支只在你手上已經有整理好的財報 CSV／Excel、或想離線測試流程時使用。

```bash
python scripts/main.py --stock_id 2330 --source csv --input_path data/raw/2330_quarterly.csv
python scripts/main.py --stock_id 2330 --source excel --input_path data/raw/2330_quarterly.xlsx
```

本地檔案格式（第一欄 `Item`，其餘為季度欄）：

```csv
Item,2024Q1,2024Q2,2024Q3,2024Q4,2025Q1,2025Q2,2025Q3,2025Q4,2026Q1
Revenue,1000,1100,1200,1300,1400,1500,1600,1700,1800
Cost of revenue,600,650,700,760,820,880,940,1000,1060
...
```

必要列項目：`Revenue`、`Cost of revenue`、`Gross profit`、`Operating expenses`、`Operating income`、`Non-operating income and expenses`、`Profit before tax`、`Tax expense`、`Net income`、`Net income attributable to parent company`、`EPS`。

也支援 `--source sample` 產生假資料（`generate_sample_financial_data(stock_id)`），純粹用來測試流程，不可用於實際分析。

加上 `--debug --save_raw` 可以在 terminal 印出中間計算結果，並輸出 `output/debug/{stock_id}_{raw_data,qoq,yoy,margins}.csv` 供檢查。

輸出位置：`output/excel/{stock_id}_quarterly_financials.xlsx`

## 專案架構

```text
.
├─ requirements.txt
├─ run_income_model.bat          # 主工具的雙擊入口
├─ scripts/
│  ├─ build_income_statement_model.py   # 主工具：FinMind → Excel 模型
│  ├─ add_revenue_build.py              # 舊模型檔的一次性升級腳本
│  ├─ download_financials.py
│  ├─ main.py                           # legacy：本地 CSV/Excel 匯入工具
│  ├─ data_sources/
│  ├─ financial_analysis/
│  └─ excel_export/
├─ data/
│  ├─ raw/                       # 放本地 CSV/Excel 財報（legacy 工具用）
│  └─ sample/
├─ output/                       # 產生的 Excel／CSV（不進版控，執行時自動建立）
│  ├─ models/                    # 主工具輸出
│  ├─ excel/                     # legacy 工具輸出
│  ├─ debug/
│  └─ finmind/
└─ notes/
   └─ project_log.md
```

## 安裝套件

```bash
pip install -r requirements.txt
```
