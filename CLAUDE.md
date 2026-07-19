# CLAUDE.md — 基甸會海山支會月例會報告產出系統

給在此 repo 工作的 AI 助手與開發者。內容以「程式碼看不出來、但踩過雷」為主。

## 這是什麼

每月把 10 份月例會 Word 報告（`-1議程` ~ `-10會員名冊`）從上月範本複製、
更新日期與數據的自動化系統。FastAPI 後端 + 單頁前端，本機與 Vercel 共用
同一份 `app.py`。

```
python3 run.py                 # 本機 → http://127.0.0.1:8848
python3 -m pytest tests/ -q    # 測試（目前 27 項，改動後必跑）
```

## 架構速覽

```
app.py               後端 + 密碼中介層（唯一的 FastAPI 實例）
api/index.py         Vercel 進入點，只做 sys.path + import app
engine/
  dates.py           民國年、財年(6/1–5/31)、第四個禮拜天、三種月份的推導
  docx_utils.py      set_cell_text（多段落 cell）、日期替換（順序固定）
  generate.py        複製範本 → 逐檔更新；-2/-9 有 API 與 Excel 兩條更新路徑
  parse_excel.py     事工成果表解析，自動辨識兩種版面
  parse_files.py     txt/csv/xlsx/pdf/jpeg 分派；LINE 聊天記錄按月篩選
  grafana.py         免登入查 dashboard 資料源（/api/ds/query）
  build_excel.py     由 API 資料重建官方版面的成果表 Excel
  drive.py           Google Drive OAuth／環境變數授權、覆寫上傳
  storage.py         LocalStorage / DriveStorage 抽象
  auth.py            單一共用密碼（HMAC cookie）
```

儲存層的關鍵設計：`work_dir()` 一律回傳**本機路徑**（雲端是 /tmp，
內容與 Drive 同步），所以 `generate.py` 完全不知道自己跑在哪。
改 generate.py 時不要引入任何對儲存位置的假設。

## 不能違反的領域規則

1. **日期替換順序**：先具體日期（`6月28日→7月26日`）→ 中文月份 → 民國年月
   → 西元年月 → 斜線年月 → 財年期間。順序寫死在 `docx_utils.update_dates()`，
   打亂會把「6月28日」的月份先換掉。

2. **三種月份不可混淆**：報告月份＝當月；範本月份＝上月；
   財年 `var-year`＝月份≥6 則 +1（2026-07 → FY2027）。

3. **年度回退預設關閉**（`fetch_excel(allow_fallback=False)`，有測試守住）。
   FY2027 報表產不出來時退回抓 FY2026，會把上一財年的 184,772 教會聖奉
   寫進本財年報告（正確值 6,400，差 29 倍）。這是真實發生過的 bug。

4. **教會見證必須 FULL OUTER JOIN 場次與奉獻**。FY2027 實測兩邊完全
   不重疊：有場次的沒奉獻、有奉獻的沒場次。只查場次會漏掉全部金額。

5. **多段落 cell**：docx 表格 cell 可能含多個 paragraph（「13人/36,000/元」
   是三段）。更新一律走 `set_cell_text()`，只改第一個 run 會留殘影。

6. **-9 匯總列**（最後兩行 merged cells）最容易被漏，測試有涵蓋。

## 外部系統的實測真相（文件與直覺都錯過）

### 事工成果表有兩種版面，欄位差兩欄
- Grafana 匯出：17 欄，弟兄在 D
- gideons.tw 總會匯出：15 欄，弟兄在 B
硬編欄號會把「會費」讀成「姓名」且不報錯。`parse_excel` 靠 Row 4 的
「日期／講員／名稱」三連欄自動辨識，改版面邏輯前先看兩個 LAYOUTS。

### 年度目標只在總會系統有
Grafana 讀的 `gideons_goal1` 缺新財年資料（回 null），但總會
`/api/targets` 有完整目標。所以 UI 建議使用者上傳總會匯出檔
（Row 5 就是目標），而非依賴 Grafana。

### Grafana dashboard 免登入
`/api/ds/query` 匿名可查（設計計畫書寫「需登入」是錯的）。
`/report/campstat/{支會}?year=` 也匿名可下載，但年度目標未建檔的
財年會回 500——這不代表資料不存在。

### gideons.tw 的「匯出」是前端組的
按鈕呼叫 3 個 API 後用 SheetJS 在瀏覽器組 xlsx，沒有可直接抓的檔案 URL，
且全部端點需登入（NextAuth）。刻意不做自動登入：需保存會員編號與生日，
且是未公開內部 API。維持「使用者自行匯出 → 上傳」。

### Excel 的達成率是小數
`1` = 100%。直接印會變成「1%」。

### LINE 匯出是跨年聊天記錄
兩萬多行、回溯數年。規則式擷取只做「按月篩訊息供人工挑選」；
別試圖從原文抽結構化欄位（曾產生 184 筆重複假排程）。

## Vercel 部署（三次 404 的教訓）

此專案的 Vercel 設定走**傳統模式：function 只認 `api/` 目錄**。
官方文件寫的「根目錄 app.py 自動偵測」是新版框架預設，在這個專案不適用
（建置錯誤原文：doesn't match any Serverless Functions inside the `api` directory）。

必須同時成立，缺一即 404：
1. `api/index.py` 匯出名為 `app` 的 FastAPI 實例
2. `vercel.json` 的 `functions` 鍵是 `api/index.py`（**不要**指定 runtime）
3. `rewrites` 把 `/(.*)` 導向 `/api/index`（否則只有 /api/* 有路由）

另外：**`.gitignore` 不要用 `*.json`**。曾因此把 `vercel.json` 擋在
repo 外，部署看似成功實則整站 404，查了三輪才發現。要擋憑證就指名擋。

## 安全邊界

- 密碼、token、憑證**只走環境變數或 credentials/**，絕不進程式碼與版控。
  `.env` 本機用；Vercel 用平台環境變數。
- 弱密碼判斷（auth.py）刻意用結構特徵而非清單——清單本身會洩漏密碼。
- 測試裡不要出現實際使用的密碼。
- `GOOGLE_REFRESH_TOKEN` 有整個 Drive 的讀寫權限，對外提及時要附警語。
- 雲端環境未設 `APP_PASSWORD` 時回 503，這是刻意的，不要「修好」它。
- 報告內容含會員個資與**具名健康資訊**（-7 代禱項目）。任何新的對外
  輸出路徑（分享連結、公開端點、log）都要先想到這件事。

## 資料檔案位置

- 本機資料：`~/Documents/海山支會/{西元年}年{M}月月例會/`（月份**不**補零）
- Drive 資料：「月例會」資料夾下的 `{西元年}年{MM}月`（月份**補零**）
  ——兩邊命名規則不同，混用會多開空資料夾
- 範本 = 上月資料夾的 11 份 docx（10 份編號 + `-A.行事曆`）
- 年度贈經計畫：`../贈經計畫/{period}_聖經配送計畫.xlsx`，整個財年固定

## 慣例

- 測試：pytest，改 engine/ 必附測試；防呆類修正（如 allow_fallback 預設）
  用測試把預設值釘住
- 訊息與 UI 全部繁體中文
- commit 訊息：第一行結論，內文寫「為什麼」，特別是推翻先前假設時
- Human-in-the-loop：所有解析結果先顯示供確認，按了才寫入資料模型；
  新功能延續這個模式
