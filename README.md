# 基甸會海山支會 · 月例會報告產出系統

把每月 10 份月例會 Word 報告的手工作業，變成瀏覽器上的幾分鐘檢核。
依《基甸會月例會報告_Web設計計畫書_完整版》(v2.4) 方案 A（本機）實作。

```
輸入介面 → 檔案解析層 → 統一資料模型(JSON) → docx 產出引擎 → 下載 / Google Drive
```

## 快速開始

```bash
cd gideons-report-app
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
python app.py            # → http://127.0.0.1:8848
```

預設讀寫 `~/Documents/海山支會/`，可用環境變數覆寫：

| 變數 | 預設 | 說明 |
|---|---|---|
| `GIDEONS_BASE_DIR` | `~/Documents/海山支會` | 每月工作資料夾的根目錄 |
| `GIDEONS_CRED_DIR` | `./credentials` | Google OAuth 憑證與 token |
| `GIDEONS_DRIVE_PARENT` | 專案指定資料夾 ID | Drive 上傳的父資料夾 |
| `GIDEONS_PLAN_DIR` | `../贈經計畫` | 年度贈經計畫 Excel 存放處 |

## 使用流程

1. **選月份** — 預設當月，會議日期自動帶入該月第四個禮拜天（可覆寫）。
2. **① 初始化** — 由上月 10 份 docx 複製並重新命名為本月檔名。
3. **上傳資料** — 五大模組分頁，各自支援 `txt / csv / xlsx / pdf / jpeg`。
   每份上傳後先顯示解析結果，**確認才寫入**資料模型（human-in-the-loop）。
4. **產生 10 份報告** — 逐檔更新日期與數據，畫面列出每份的變更明細。
5. **下載 ZIP** 或 **上傳 Google Drive**（自動建立「2026年7月」資料夾、同名檔覆寫）。

### 五大輸入模組

| 模組 | 影響文件 | 主要輸入 |
|---|---|---|
| ① 產生報告 | 全部 10 份 | 月份、會議日期 |
| ② 行政財務 | -2 -3 -4 -9 | Grafana 事工成果表 Excel、現金帳 |
| ③ LINE 與贈經 | -1 -5 -6 | LINE 匯出文字、贈經名單／照片 |
| ④ 代禱 | -7 | 代禱會影像或文字 |
| ⑤ 其他／增量 | -1 -5 -8 | 贈經排程、臨時動議、輪值調整 |

模組⑤走增量更新：`POST /api/generate` 帶 `only: [7, 8]` 只重繪指定文件。

## Google Drive 設定（首次一次即可）

1. [Google Cloud Console](https://console.cloud.google.com/) 建立專案 → 啟用 **Google Drive API**。
2. 「API 和服務 → 憑證」→ 建立 **OAuth 用戶端 ID**，類型選 **桌面應用程式**。
3. 下載 JSON，存成 `credentials/client_secret.json`。
4. 啟動 App，按「連結 Google Drive」，瀏覽器完成授權；
   token 存到 `credentials/token.json`，之後自動續用。

上傳行為：在目標父資料夾底下找「`{年}年{月}`」子資料夾，沒有就建立；
同名 DOCX 存在則 **update 既有檔案**（保留檔案 ID 與分享連結），否則新建。

> `credentials/` 已列入 `.gitignore` — 憑證與 token 絕不進版控。

## 事工成果表（Grafana）

介面「開啟最新年度 Dashboard」按鈕會依今天日期自動算出財年並組好 URL
（財年 6/1–5/31，6 月起屬下一財年；2026-07 → `var-year=2027`）。
登入後匯出 Excel，於模組②上傳即可；系統會另存為
`事工成果表_{年}_{MM}.xlsx` 供產出引擎引用。

### Excel 欄位對照（已用 2026/06 實際匯出檔驗證）

| Row | 內容 | | Col | 內容 |
|---|---|---|---|---|
| 5 | 目標 | | A | 編號 |
| 6 | 成果 | | D / E | 弟兄 / 姊妹 |
| 7 | 差額 | | F-G | 會費日期 |
| 8 | 達成率（小數，1 = 100%） | | H-I | 聖經奉獻 |
| 9+ | 各會員 | | J-K | 巴拿巴奉獻 |
| | | | L-O | 見證日期／講員／教會／金額 |
| | | | P-Q | 贈經類別 / 本數 |

> ⚠️ P/Q 在**會員列**是贈經明細（類別＋本數），與欄標頭「贈經總數／姊妹贈經」
> 語意不同；總數一律取 Row 5-8 的匯總值，不自行加總。

## 各文件更新規則

| # | 文件 | 自動更新 | 備註 |
|---|---|---|---|
| 1 | 議程 | 日期、年度主題 | 臨時動議等標【TODO】 |
| 2 | 事工成果統計表 | 目標／成果／差額／達成率四列 | 直接對映 Excel Row 5-8 |
| 3 | 收入支用統計表 | 標題日期 | 現金帳需人工或上傳 |
| 4 | 各項奉獻 | 各會員會費／聖經奉獻 + 匯總 | 「安息／退會」註記保留 |
| 5 | 贈經事工（除學校） | 標題；列出本月排程 | 內容人工確認 |
| 6 | 學校贈經統計表 | 標題、年度期間 | |
| 7 | 會員及地界教會代禱 | 標題 | 圖片需多模態辨識 |
| 8 | 早禱會及月例會輪值表 | 標題、日期 | |
| 9 | 年度教會見證統計表 | 日期／講員／金額 + **匯總列** | 教會名模糊比對，未對應會警示 |
| 10 | 會員名冊 | 標題 | |

**日期替換鐵則**：先具體日期（`6月28日`→`7月26日`）、再中文月份、再通用年月、
最後斜線年月與財年期間。順序寫死在 `docx_utils.update_dates()`。

**多段落 cell**：`set_cell_text()` 會清掉 cell 內所有段落的所有 run 再寫入
（聖奉欄位常是「13人 / 36,000 / 元」三個 paragraph）。

## 專案結構

```
gideons-report-app/
├─ app.py                 FastAPI 後端 + REST API
├─ static/index.html      單頁前端（五大模組頁籤）
├─ engine/
│  ├─ dates.py            民國年／財年／會議日期／Dashboard URL
│  ├─ docx_utils.py       set_cell_text、跨 run 替換、日期替換
│  ├─ generate.py         複製範本 + 逐檔更新
│  ├─ parse_excel.py      事工成果表解析與分析
│  ├─ parse_files.py      txt/csv/xlsx/pdf/jpeg 分派、LINE 解析
│  └─ drive.py            Google Drive OAuth／建資料夾／覆寫上傳
└─ models/model.py        統一資料模型
```

## API

| 方法 | 路徑 | 說明 |
|---|---|---|
| GET | `/api/status` | 月份衍生值、檔案狀態、Drive 授權狀態 |
| POST | `/api/init` | 複製前月範本並重新命名 |
| POST | `/api/parse` | 上傳並解析單一檔案（不寫入） |
| POST | `/api/commit` | 確認後寫入資料模型 |
| POST | `/api/generate` | 產出 10 份（`only` 可指定增量） |
| GET | `/api/analysis` | 事工成果表分析 |
| GET | `/api/download` | 下載單檔或整月 ZIP |
| POST | `/api/drive/authorize` `/api/drive/upload` | Drive 授權與上傳 |

## 已知限制

- LINE 原始匯出是跨年度聊天記錄，規則式擷取只能做到「篩出當月訊息供人工挑選」；
  結構化欄位（年度主題／贈經排程）需整理過的摘要文字，或後續接 LLM。
- 代禱會影像（-7）與掃描型 PDF 需多模態辨識，目前只收檔並標示待辨識。
- 無 PDF 預覽（需 LibreOffice，本機可自行補上 `soffice --convert-to pdf`）。
- 6 月為新財年，職員名單切換仍需人工確認。

## 資料安全

工作資料含會員個資與奉獻金額。`.gitignore` 已排除 `*.docx`、`*.xlsx`、
`_inputs/`、`credentials/`。推上遠端前請再確認 `git status`。
