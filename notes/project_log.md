# Project Log

## 2026-06-17

- 已完成任意 `stock_id` 參數化架構。
- `7722` 只是測試範例，不是寫死的股票代號。
- 建立第一版 Python / Excel MVP：
  - `scripts/main.py` 使用 `argparse` 接收 `--stock_id`。
  - `scripts/data_sources/sample_data.py` 提供 `generate_sample_financial_data(stock_id: str)`。
  - `scripts/financial_analysis/metrics.py` 計算 QoQ、YoY 與 profitability margins。
  - `scripts/excel_export/exporter.py` 輸出格式化 Excel。
- Excel 輸出檔名會依照股票代號產生，例如 `output/excel/2330_quarterly_financials.xlsx`。
- 下一步是加入真實資料來源，例如 Goodinfo、MOPS、CSV、Excel input。

## 2026-06-17 Debug / Audit Update

- 新增 debug mode：使用 `--debug` 可在 terminal 印出 stock_id、source、季度欄位、原始資料、QoQ、YoY、margins 與輸出檔案路徑。
- 新增 raw data export：使用 `--save_raw` 可輸出 `output/debug/{stock_id}_raw_data.csv`、`{stock_id}_qoq.csv`、`{stock_id}_yoy.csv`、`{stock_id}_margins.csv`。
- 建立 financial calculations 的 audit trail，方便檢查 Excel 數據是資料來源問題、計算問題或輸出格式問題。
- 明確標註目前 `sample` source 是 synthetic sample data，不是真實財報資料。

## 2026-06-17 Local Data Source Update

- 新增 `data/raw/` 作為本地真實財報資料放置位置。
- 新增 CSV 匯入功能：`scripts/data_sources/csv_loader.py`，支援 `--source csv --input_path data/raw/{stock_id}_quarterly.csv`。
- 新增 Excel 匯入功能：`scripts/data_sources/excel_loader.py`，支援 `--source excel --input_path data/raw/{stock_id}_quarterly.xlsx`。
- 新增共用 schema 驗證：`scripts/data_sources/schema.py`，檢查必要季度欄位與財務項目列。
- `csv` / `excel` 成功讀入後會沿用同一套 QoQ、YoY、margin、Excel export、debug、save_raw 流程。
- `sample` source 保留為測試用途，但不再作為主要資料來源。

## 2026-08-11 估值自動化 + Dashboard 圖表修正

- 修正 `build_income_statement_model.py` Dashboard 分頁圖表重疊問題：根本原因是圖表高度
  (8.5cm) 大於錨點列距 (16列×15pt≈8.46cm)，且單一數列圖表未關閉圖例，擠壓繪圖區。
  改為 20 列列距 + 9.0cm×14.5cm 圖表尺寸（已用欄寬/列高換算成 cm 驗證不重疊），單一數列
  圖表關閉圖例，多數列圖表圖例移至底部，圖表標題/座標軸字型統一為 Arial Narrow，格線改為
  淡灰細線，全部透過共用的 `_style_chart()` helper 套用，新分頁圖表沿用同一套風格。
- 新增 4 個全自動估值分頁（原專案完全沒有估值邏輯，此次為新建）：
  - **PE Band** / **PB Band**：12 個月前瞻本益比／股價淨值比河流圖。NTM EPS 以 Income Model
    既有季度 EPS 公式即時運算（超出季度預測範圍時，以下一個年度預測值 ÷4 補足缺口）；
    估值區間統計（平均值、標準差）取自 FinMind 近 5 年日 PER/PBR 資料，寫成 Assumptions
    分頁可編輯的黃色格（含智慧預設值）。BVPS 以「上一季 BVPS + 預測 EPS ×(1−股利發放率假設)」
    滾動推算，發放率假設同樣即時連動。
  - **Dividend History**：依除息年度彙總 FinMind 股利資料，含配息率、殖利率與趨勢圖。
  - **Turnover Days**：存貨/應收/應付週轉天數（單季年化法，較 TTM 平均法需要更短歷史，
    能涵蓋更多季度），預測欄位因本模型未預測資產負債表科目而刻意留白，不亂猜。
  - 新增 FinMind 資料源：`TaiwanStockPrice`、`TaiwanStockPER`、`TaiwanStockDividend`；
    皆包成 `fetch_finmind_optional()`，單一資料源失敗只會讓對應分頁留白並印出警告，
    不會讓整支股票的產出失敗。
  - 修正過程中抓到並修正 2 個計算錯誤：(1) PB Band 的 BVPS 公式股本單位未轉換
    （bs_extra 已除以 UNIT，OrdinaryShare 原始值未除，導致 BVPS 幾乎為 0）；
    (2) 股利發放率預設值與股利表配息率誤用當年度（資料尚未齊全的進行中年度）EPS，
    改用只在四季資料齊全時才計算年度 EPS 的 `_annual_eps_complete()`。
  - 已用 2330（資料齊全）與 7856（PER 資料缺失）各跑一次驗證：公式手動推導核對正確、
    數值量級合理（TSMC BVPS/DIO/DSO/股利歷史均與已知公開數據吻合），缺資料時能優雅降級
    不中斷整支股票的產出。

## 2026-08-11 Revenue Build Stage 1（2059）

- 新增 `scripts/add_revenue_build.py`：獨立、可重複使用的「升級腳本」，直接修改
  已產生的 `{stock}_income_model.xlsx`（先備份、非重建整份 workbook），插入
  Revenue Build 分頁並改接 Income Model / Assumptions 的營收公式。與主產生腳本
  `build_income_statement_model.py` 共用同一套樣式 helper（`_f`/`_sec`/`_cfill`/
  `FILL_*`/`BORDER_*`），確保視覺風格完全一致。
- 根因確認：2059 原本 26Q2F-26Q4F Revenue 因 `Assumptions!B5:D5` 從未填值而恆為 0，
  下游 Gross Profit/OI/NI/EPS 全部歸零，PE Band NTM EPS 幾乎只靠 2027F/4 proxy 撐著，
  造成異常 Forward P/E（用戶回報的「1545x」）。2027F-2029F 的年度成長假設也是 0%
  （等於沒有真正的長期預測）。
- Revenue Build 分頁（Assumptions 之後、Dashboard 之前）：
  A. Historical Revenue（純引用 Income Model，不重複輸入）
  B. Revenue Driver Build（Volume/ASP/Mix/FX 四因子 → Implied Growth → Manual
     Override（留空預設）→ Applied Growth → Forecast Revenue；季度用去年同季 base，
     年度用前一年度 base，避免季節性失真）
  C. Segment Mix Framework（AI/一般伺服器/其他，簡化 Driver Mode，因無可靠歷史
     segment 資料，預設不影響 Total Revenue，僅供未來擴充）
  D. Model Checks（季度加總、Segment Mix 合計、Revenue Build↔Income Model 連結、
     Driver 公式錯誤值，皆為 OK/ERROR 判斷式，並加了條件式格式化上色）
- 中性預設值：季度 Volume Growth = 26Q1 實際 YoY（37.82%，取代原本的空白/0）；
  年度 Volume Growth = 原 Assumptions 營收成長率假設值原封不動搬過來（2059 為 0%，
  維持原預測不變）；ASP/Mix/FX 皆為 0%。
- Income Model 只動了 Revenue 列：M3/N3/O3 改連 Revenue Build!I18/J18/K18，
  Q3/R3/S3 改連 Revenue Build!M18/N18/O18；P3（2026F）公式完全沒動，因為它本來就是
  加總 L3+M3+N3+O3，換源後自動正確。COGS/GP/Opex/OI/Tax/NI/EPS 全部公式未變動。
  Assumptions!B5:D5、B17:D17 改成唯讀公式（連回 Revenue Build），底色從黃改米白，
  標籤加註「由 Revenue Build 計算」。PE Band 只在既有註解格追加一段文字說明
  27Q1F proxy 邏輯，沒有動任何公式。
- 手動核算驗證（依 2059 真實歷史數字推導）：26Q2F Revenue ≈ 58.3 億、EPS ≈ 32.95
  （原本是 0）；NTM EPS（26Q1 時點）≈ 139.7（原本因分子幾乎全靠 1 個 proxy 季度，
  數值異常小）。Model Check 公式、Revenue Link Check 皆設計為可在 Excel 開啟時
  即時驗證一致性。
- 備份檔：`output/models/2059_income_model.backup_20260811_011956.xlsx`。

### 同日修正：Row14 標籤誤判為公式 + Segment 毛利率擴充

- 用戶回報 Revenue Build 出現「公式亂碼」。查證後根因：`_row_label` 把
  `"= Implied Revenue Growth YoY (Driver 推導)"` 寫入儲存格時，Excel/openpyxl
  規則是「開頭為 `=` 一律視為公式」，導致這段中文標籤被誤存成 `<f>` 公式節點而非
  純文字，Excel 開啟時解析失敗顯示異常內容。已在 `_row_label()` 加上防呆
  （標籤開頭若為 `=+-@` 直接拋錯），並移除所有標籤開頭的 `=`／`→` 等符號，
  筆記文字中的 `×`、`−`、`≈`、`Σ` 等特殊符號也一併換成純 ASCII，降低字型缺字風險。
  已用程式逐格掃描確認 Revenue Build 全表不再有任何「文字被誤存為公式」的儲存格。
- 擴充 Segment / Product Mix Framework：每個 Segment 從 3 列（Mix%/Volume/ASP）
  增加到 4 列，新增 **Gross Margin %**；新增 **Blended Gross Margin 加權毛利率試算列**
  （= Σ(Mix% × Segment Margin%) / ΣMix%，未填寫則顯示 N/A）。此為試算參考工具，
  不會自動覆蓋 Assumptions 的毛利率輸入格（仍維持原本直接輸入、自由調整的方式），
  避免產生「一改 Segment 就默默改變既有預測」的風險；如果使用者認同試算結果，
  需自行複製貼上到 Assumptions。Model Checks 區塊隨版面調整下移（row 40-43）。
- 已重新對 2059 執行升級（先還原備份、再套用修正後腳本），逐格驗證公式與數值正確。

### 同日：Revenue Build 併入主產生腳本（bat 檔同步）

- 先前 Revenue Build 只存在於 `scripts/add_revenue_build.py`（事後補丁工具），透過
  `run_income_model.bat` 產生的新股票不會自動有這個分頁。現已把 Revenue Build 的
  建表邏輯搬進 `scripts/build_income_statement_model.py` 主流程
  （新函式 `build_revenue_build_sheet`），並修改 `_qtr_fcast_formula`／
  `_ann_fcast_formula` 讓 Income Model 的 26Q2F-26Q4F、2027F-2029F Revenue
  直接生成時就連 Revenue Build，`build_assumptions_sheet` 的營收相關格也是
  建立當下就是唯讀公式（不再是先建可編輯格、事後才改）。往後不管透過 bat 檔
  或指令列，任何股票代碼產生的 Excel 都會內建 Revenue Build，不需要再額外執行
  補丁腳本。
- 年度 Volume Growth 的中性預設值，在「全新產生」情境下沒有「原本的假設」可以
  保留，改用原本腳本一直以來的保守預設 5%（與舊版 `營收成長率 YoY` 預設值一致）；
  季度 Volume Growth 預設值一樣是最近一期實際 YoY，直接從抓下來的 `data` dict
  計算，不需要像補丁腳本那樣去讀已存在的 Excel 儲存格。
- 同步把 row14 標籤誤判為公式的防呆（`_rb_row_label` 開頭字元檢查）與 Segment
  毛利率試算擴充，都內建到這個共用路徑，`add_revenue_build.py` 仍保留、專門用來
  補丁「這次升級之前產生」的舊檔案（例如 1595/2474/3016/4919/6786/7722/7856），
  兩邊已用系統化掃描確認都不再有任何「文字被誤存成公式」的儲存格。
- 已用 2059、2330（TSMC）分別重新從零產生驗證：Sheet 順序、Income Model／
  Assumptions／PE Band 的連結、Revenue Build 各區塊公式皆正確；2330 的季度
  Volume Growth 預設值算出 35.13%，符合其近期真實營收成長趨勢。

### 同日：PE Band / PB Band「圖表怪怪的」— 根因是預測季度已經過期

- 用戶回報 PE/PB Band 圖表看起來不對。實際重演整條公式鏈後發現：2059 的
  26Q2 2026 真實財報（EPS 74.38、營收 108.3 億）已經進了 FinMind，但 Income
  Model 的欄位模板是固定寫死的（26Q1 永遠是「最新實際」、26Q2F 永遠是「預測」），
  不會隨時間推進自動改口徑。PE Band 算「回顧」某幾季的 NTM EPS 時，需要往後看
  1-3 季，剛好會取到這個已經過期但還在跑預測公式的 26Q2F/26Q3F/26Q4F 儲存格，
  導致 Band 曲線失真。
- 兩個方案給用戶選：(1) 只修 PE/PB Band 讓它們自己判斷有沒有真實數據可用；
  (2) 整個模型改成自動偵測「目前最新實際季度」、連動改所有分頁欄位定義。
  用戶選了 (1)，範圍限定在 PE Band / PB Band 兩個分頁，不動 Income Model／
  Revenue Build／Assumptions 的欄位架構。
- 實作：`_ntm_formula()` 新增 `data` 參數，組 NTM EPS 公式時，凡遇到
  26Q2F/26Q3F/26Q4F 且 FinMind 已有該季實際 EPS，直接把實際數字寫成公式裡的
  常數項，取代原本連到 Income Model 預測儲存格的參照。PB Band 的 BVPS 比照辦理
  （改用 `bs_extra`／`data` 是否有該季實際資產負債表數據判斷，非再看
  `BAND_HIST_QTR_KEYS` 這個寫死的集合）。價格列與隱含倍數列也從「只有歷史季度才填」
  改成「只要抓得到那一季季底交易日收盤價就填」，不再受限於模板上標的是否為
  「(F)」。修完後補了一個邊界情況防呆（若 BVPS 序列第一欄剛好缺實際資料，
  不會誤把標籤欄 A 當成前一欄拿去做滾動公式）。
- 已用 2059（26Q2 已真實入帳）與 2330（26Q2 財報尚未公布、僅股價已有）分別
  重新產生驗證：2059 的 NTM EPS／BVPS 公式正確改用真實數字，2330 則正確維持
  原本的預測公式（股價因為沒有財報公布時滯，仍可以往後多填一欄），確認沒有
  對「真的還沒公布」的情況造成誤判或回歸。

### 同日：Revenue Build 標題／說明改回中文為主 + 修寬度截字問題

- 用戶回饋 Revenue Build 的區塊標題、驅動因子列、Segment 列、Model Checks
  列標籤都改成英文為主（上一輪為了處理「全部變成中文」的誤解而改的），
  現在要求改回中文在前、英文在後的風格，比照 Income Model／Assumptions 既有慣例。
  已把所有區段標題（A/B/C/D）、驅動因子列（出貨量／平均售價／產品組合／匯率／
  隱含成長率／人工覆蓋／實際採用／營收預測）、Segment 列（AI／高階伺服器、
  一般伺服器、其他產品、營收佔比、毛利率等）、Model Checks 四項檢查標籤全部
  改為中文開頭，並同步修正 Assumptions／說明文字裡提到的對應欄位名稱。
- Model Checks 第 4 項標籤太長導致單一欄寬（A 欄）放不下被截斷：改成標籤跨欄
  合併 A:D（四欄），檢查結果（OK/ERROR）搬到 E 欄，並把 A 欄整體寬度從 34
  加到 42，讓其他也變長的中文標籤有更多空間。`build_income_statement_model.py`
  的內建版本與 `add_revenue_build.py` 補丁腳本兩邊同步套用，保持一致。
- 已重新產生 2059 驗證：逐格掃描確認沒有「文字被誤存成公式」的回歸，四個
  Model Check 標籤都正確跨欄合併、檢查結果正確落在 E 欄。

### 同日：新增 Scenario Analysis（情境分析）與 Backtest（估值訊號回測）

- 兩個都是新分頁，位置在 PB Band 之後、Dividend History 之前。
- **Scenario Analysis**：Bull/Base/Bear 三情境。Base 不是另一份假設，是直接連動
  主模型（2026F 讀 Income Model 自己的欄位反推隱含成長率/毛利率；2027F-2029F
  讀 Revenue Build 的 Applied Growth 與 Assumptions 的毛利率）。Bull/Bear 用
  「相對 Base 的調整幅度 Δ」表示（預設營收±5pp、毛利率±2pp），黃色可編輯，
  實際使用的成長率/毛利率是 Base+Δ 的公式，不需要在 Python 端知道 Base 的
  實際數值。營業費用率／業外收支／稅率／股本三情境共用主模型數值（情境分析
  標準簡化作法，只變動營收與毛利率兩個主要變因）。往下算出三情境的完整迷你
  損益（Revenue→GP→OI→NI→EPS），2026F-2029F。目標價 = 情境 EPS × PE Band
  統計倍數，採 Bull配Mean+1SD、Base配Mean、Bear配Mean-1SD 的雙重疊加保守估法。
- **Backtest**：回測「股價跌破 PE Band 的 Mean-1SD 買進、漲破 Mean+1SD 賣出」
  這個訊號策略，vs. 買進持有。用週資料（每週最後一個交易日），跟 PE Band
  的統計回溯年數用同一段區間。整張表逐列都是 Excel 公式（部位／報酬／淨值／
  歷史高點／回撤都是連鎖參照上一列），訊號公式直接連到 PE Band 的平均值/
  標準差儲存格，那邊手動改了這裡會跟著連動更新，不是 Python 端算好貼死的
  數字。摘要區（總報酬率／年化報酬率/CAGR／最大回撤／交易次數）在表格上方，
  權益曲線圖表獨立佔一塊列區（row 14-33），避免跟摘要表、說明文字、明細表
  標題互相重疊（就是先前 Dashboard 圖表重疊那個問題的預防作法）。
- 兩邊都對「文字開頭不能是 =+-@」的防呆做了套用（沿用 Revenue Build 那次
  抓到的教訓），也把同樣的防呆補進 `_band_row_label`（PE Band/PB Band/
  Turnover Days 共用的 helper），一次性堵住這整類 bug。
- 已用 2059、2330 分別重新產生驗證：兩張新分頁都正確生成、無「文字誤存成
  公式」的儲存格、公式手動核對正確（Base 情境的營收鏈公式與 Income Model
  自己的預測公式使用相同來源，理論上應該互相 tie out）、Backtest 逐列公式
  鏈正確（首列種子值、後續列正確參照上一列），259 週的資料範圍與 PE Band
  的 5 年回溯窗口一致。

### 同日：Scenario Analysis 補 4 張圖表

- 原本只有 1 張「情境目標價趨勢」折線圖，用戶要求更直觀，改成 2x2 圖表矩陣
  （沿用 Dashboard 那套已驗證過不會重疊的版面數學：CHART_ROW_GAP=20 列、
  14.5x9.0cm、B/J 欄錨點）：① 情境目標價趨勢（折線）② 情境 EPS 比較（三情境
  長條圖，直接看出獲利能力差異）③ 潛在漲跌幅（Bull 漲幅／Bear 跌幅長條圖，
  正負值一眼看出風險報酬）④ 目標價區間 Football Field（股票研究常見畫法：
  堆疊長條，底部一段透明只是墊高用，上面兩段分別是 Bear→Base、Base→Bull
  區間，形成一根每年「浮動」在 Bear-Bull 之間的長條）。
- Football Field 圖表用了「透明堆疊底座」技巧（該序列 noFill + 圖例隱藏該項），
  存檔前先用一個獨立小測試檔確認 openpyxl 寫出來的 XML 是正確的
  noFill／stacked／legendEntry delete，才套用到正式程式，避免這裡沒辦法
  肉眼預覽就交出去。新增 3 列圖表輔助資料（Bear 目標價、Bear→Base 區間、
  Base→Bull 區間），標註「勿刪除」，不影響其他計算。
- 已用 2059、2330 重新產生驗證：4 張圖表都存在、位置與尺寸經 XML 檢查確認
  無重疊，Football Field 圖表的透明序列與圖例隱藏都正確寫入。

### 同日：全部圖表改用「固定版面矩形」樣式，修掉標題壓線問題

- 用戶把自己在 Excel 手動調整過、覺得好看的圖表複製貼到 2059 檔案新開的
  Sheet1 分頁當參考，要求所有圖表都改成那個樣子。用 zipfile 直接讀圖表的
  原始 XML（openpyxl 讀回來的物件屬性有些看不出全貌）比對出關鍵差異：
  參考圖表的標題是 `overlay="1"`（不占專屬版面、直接浮在繪圖區上方），
  搭配繪圖區用 `manualLayout` 釘死一個固定矩形（x=7.05%／y=6.68%／
  寬88.6%／高85%，相對圖表整體），圖例也用 manualLayout 釘在左上角一小塊
  （x=13.65%／y=12.99%／寬24.48%／高30.02%），同樣 overlay。這跟我原本
  依賴 Excel 自動版面演算法（只設定圖例位置在下方、不管繪圖區怎麼排）
  不一樣——自動版面在圖表本身不大、標題字級偏大時，就是這次一直反覆出現
  「標題跟線重疊」的根本原因，因為 Excel 每次都要重新「猜」要留多少空間，
  猜不準就會撞在一起；手動釘死矩形之後 Excel 不用再猜，是更穩定的做法。
- 套用時抓到一個 openpyxl 的真實陷阱：直接設定 `ch.plot_area.layout` 存檔後
  會整個消失不見，追進 openpyxl 原始碼才發現 `ChartBase._write()` 存檔當下
  會強制執行 `self.plot_area.layout = self.layout`，把 `plot_area.layout`
  蓋掉——正確做法是設定 `ch.layout`（圖表物件本身的屬性），不是
  `ch.plot_area.layout`。已用獨立小測試檔案先確認這個修正生效、且原本
  以為對的寫法真的會在存檔後消失，才套用到正式程式，避免這種讀不出來的
  openpyxl 內部行為又埋一次坑。
- 這次改的是共用的 `_style_chart()` / `_chart_title()` helper，全部 16 張
  圖表（Dashboard 6 張、PE Band、PB Band、Scenario Analysis 4 張、Backtest、
  Dividend History、Turnover Days）都是透過這兩個 helper 產生，一次修好
  全部套用，不用每張圖表分別改。標題字級從 11pt 調整為 16pt（比照參考圖表），
  字體仍是 Arial Narrow、酒紅色 7A0000（沒有跟著參考圖表變成預設黑色——
  這點是我的判斷，猜測參考圖表掉色是複製貼上的副作用而非用戶刻意要拿掉
  品牌色，有落差的話請再說一聲）。
- 已用 2059、2330 重新產生驗證：plotArea/legend 的 manualLayout 與 title
  overlay 皆正確寫入原始 XML（單一數列、無圖例的圖表也正確套用繪圖區矩形，
  只是沒有圖例區塊），全工作簿 16 張圖表數量不變。

### 同日：修正上一版造成的圖例重疊 + 格線改回原生質感

- 用戶回報上一版改壞了：圖例現在會跟線/長條重疊，而且用戶真正要的重點其實
  是格線要更有質感，不是黑色實心。重新檢討後發現兩個問題：
  1. 圖例沿用參考圖表的「浮動在左上角一小塊」manualLayout，那個位置是
     依參考圖表（P&L Comparison 長條圖）當時的資料形狀調出來的，套用到其他
     15 張資料形狀不同的圖表（趨勢線、Football Field 等）就會蓋到資料。
     改成圖例固定在右側一個獨立欄位（x=0.76~0.98），並把繪圖區寬度收窄到
     x=0.07~0.73，兩者中間留 3% 間距、且圖例 `overlay="0"`（不佔用繪圖區、
     非浮動），這樣不管哪張圖表的資料長怎樣，圖例都不會蓋到。單一數列、
     不需要圖例的圖表則讓繪圖區用滿寬度（90%）。
  2. 格線先前是我自己指定純色純寬度蓋掉的；追查參考圖表的原始 XML 才發現
     它的格線之所以好看，其實是因為完全沒有手動指定顏色——`ch.style=10`
     這個 Excel 內建圖表樣式本身就有一套柔和的預設格線效果，我先前又額外
     疊了一層純色格線把它蓋掉了。改成只留 `<majorGridlines/>` 空元素（等於
     跟參考圖表一樣「不覆蓋，直接吃 style=10 的預設值」），不用自己去猜
     漸層要怎麼調。
- title 的 overlay 效果（上一版真正解決「標題壓線」的部分）維持不變，
  沒有被這次修正影響到。
- 已用 2059、2330 重新產生驗證：原始 XML 確認繪圖區/圖例的 manualLayout
  座標正確（有圖例的圖表繪圖區收窄、圖例落在右側獨立欄位且 overlay=0；
  無圖例的圖表繪圖區維持滿寬)，格線都是乾淨的空 `<majorGridlines/>`，
  title overlay 仍然存在。

### 同日：圖表背景改米白 + 修正被我上一版誤縮的繪圖區比例

- 用戶回饋還是沒有高級感、背景要統一 #FAF9F4（跟儲存格背景同一個米白色）、
  比例不對。查了一下發現繪圖區的垂直高度上一版被我改成 0.72（上下各留
  14% margin），比參考圖表原本的 0.85 小了一截——原意是想空間給下方 x 軸
  標籤，結果留過頭了，圖看起來比例矮胖、留白過多，跟參考圖表的舒展比例
  差很多。已改回 y=0.08／h=0.84（上下各約 8% margin，貼近參考圖表原始
  比例），寬度收窄（有圖例欄位時）的部分維持不變。
- 背景色：`ch.graphical_properties`（整個圖表外框背景）與
  `ch.plot_area.spPr`（繪圖區內部背景）都設成 `FAF9F4`，讓圖表整塊融入
  米白色的工作表背景，不會浮一塊視覺上突兀的白色矩形。過程中一度誤判
  `ch.plot_area.spPr` 沒有生效（跟 layout 一樣的陷阱），後來發現是我自己
  找 XML 字串時只往後看了 120 字元、沒看到 `<spPr>` 其實有正確寫在
  `<plotArea>` 結尾（barChart／catAx／valAx 之後），實際上一開始就是對的、
  是我自己找漏了，不是真的 bug。
- 已用 2059、2330 重新產生驗證：原始 XML 確認 FAF9F4 出現兩次（圖表外框
  背景 + 繪圖區背景），繪圖區與圖例的 y/h 座標已改回較寬鬆的比例。

### 同日：繪圖區邊框改漸層

- 用戶明確指出根本問題是 Format Plot Area 的「Border」本身是純色實線，
  這條線跟浮動（overlay）的標題視覺上「撞」在一起才顯得醜，要求把邊框
  改成漸層。新增 `_plot_border()`，用現有的茶棕色 D9D2C2 到酒紅 7A0000
  水平漸層（0.75pt 線寬），套在 `plot_area.spPr.ln` 上，只改「Format Plot
  Area」的邊框，沒有動整個圖表外框的邊框（用戶只點名 Plot Area）。
- 已用 2059、2330 重新產生驗證：原始 XML 確認繪圖區的 `<a:ln>` 內是
  `<a:gradFill>`（D9D2C2 → 7A0000，水平方向），不是純色 `<a:solidFill>`。

### 同日：Revenue Build 第43列高度衝突 bug + A欄長標籤統一改自動換行

- 用戶回報第43列格式還是有問題、字被吃掉，且 A 欄有些格子文字被邊框線
  裁掉。查出兩個問題：
  1. **真正的 bug**：`build_revenue_build_sheet` 裡 `NOTE2_ROW = CHECK_ROW0 + 4`
     算出來剛好跟第 4 個 Model Check（就是第 43 列）同一列，後面那行
     `ws.row_dimensions[NOTE2_ROW - 1].height = 6` 其實是想留一個空白間隔列，
     結果算錯位置，直接把第 43 列（已經被 `_rb_row_label` 設成正常高度）的
     列高又蓋回 6pt——這才是「字被吃掉」的真正原因，不是版面比例或字型問題。
     已改成 `NOTE2_ROW = CHECK_ROW0 + 5`，讓間隔列落在自己獨立的一列，
     不再跟任何 Check 列衝突。逐列檢查過整個函式裡其他的 row_dimensions
     設定，確認沒有其他同類型的列號衝突。`add_revenue_build.py`
     （補丁腳本）裡同一段複製貼上的邏輯有一樣的 bug，一併修掉。
  2. **A 欄長標籤裁字**：Revenue Build 的雙語標籤（中文+英文）常常比 A 欄
     寬，右邊 B 欄又都有資料擋住，文字溢出時 Excel 會直接裁掉而不是外溢。
     改成 `_rb_row_label()`／`_row_label()` 統一套用自動換行
     （`wrap_text=True`）+ 列高從 15pt 提高到 28pt（可放 2 行 9pt 字），
     不用再一個個手動猜每個標籤實際需要多寬的欄位——這件事我已經猜錯
     兩次了，改成自動換行是比繼續猜版面數字更穩妥的做法。
- 已用 2059、2330 重新產生驗證：第 43 列高度正確維持 28pt（不再被壓回
  6pt），44 列是獨立的間隔列、45 列是說明文字，逐格掃描確認沒有「文字被
  誤存成公式」的回歸。
