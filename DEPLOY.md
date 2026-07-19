# 部署指南

## 一、本機執行（現況，可直接用）

```bash
cd ~/Claude/Projects/基甸會月例會報告/gideons-report-app
bash run.sh
```

首次會自動建 `.venv`、裝套件、跑測試，然後開 http://127.0.0.1:8848。
之後每次都是同一行。

### 每月操作順序

1. 確認上方年月（預設當月）與會議日期（預設第四個禮拜天）
2. 按 **① 初始化 / 複製範本** → 由上月 11 份 docx 複製並改名
3. 分頁 **② 行政財務** → **擷取本財年資料** → 核對數字 → **✓ 確認並套用**
4. 需要 Excel 存查就按 **產生成果表 Excel**
5. 其他模組（③ LINE、④ 代禱、⑤ 其他）視需要上傳
6. 按 **產生 10 份報告**，看畫面列出的變更明細
7. **下載 ZIP** 或 **上傳 Google Drive（覆寫）**

### 疑難排解

| 狀況 | 處理 |
|---|---|
| `No module named fastapi` | `.venv` 沒啟用，直接跑 `bash run.sh` |
| Drive 顯示「待授權」 | 按「連結 Google Drive」；測試模式的 token 7 天會過期 |
| 擷取資料失敗 | 確認網路可連 `gideons-dashboard.pointing.tw` |
| 找不到範本資料夾 | 確認 `~/Documents/海山支會/{年}年{月}月例會` 存在 |

---

## 二、推上 GitHub

```bash
cd ~/Claude/Projects/基甸會月例會報告/gideons-report-app
bash setup-git.sh          # 建 repo、檢查沒有個資、commit
git push -u origin main
```

`setup-git.sh` 會在 commit 前擋下 `.docx`／`.xlsx`／`credentials/`。
`.gitignore` 已涵蓋，但推之前請再看一眼 `git status`。

> ⚠️ repo 若是 public，任何人都看得到程式碼。程式碼本身沒有機密
> （憑證與資料都被排除），但 `engine/grafana.py` 內含 dashboard URL 與
> `var-ref=5394` 等參數。若介意就設成 private。

---

## 三、Vercel — 先看這段再決定

**現在這個版本不能直接丟上 Vercel。** 不是設定問題，是架構上真的不相容。

### 三個硬性阻礙

| # | 問題 | 為什麼 | 要改什麼 |
|---|---|---|---|
| 1 | **讀不到本機資料夾** | serverless 沒有你的 `~/Documents/海山支會`，而「複製上月 11 份 docx 當範本」整套流程都靠它 | 改用 Google Drive 當儲存層 |
| 2 | **寫不了檔案** | serverless 檔案系統唯讀（只有 `/tmp` 可寫且不保存），`_inputs/`、`model.json`、產出的 docx 全都要寫檔 | 產出改為記憶體 → 直接上傳 Drive |
| 3 | **OAuth 流程不同** | `run_local_server()` 要開瀏覽器回呼 localhost，雲端做不到；`token.json` 也無處可存 | 改網頁版 OAuth，token 存環境變數或 Vercel KV |

### 還有一個更該先想清楚的：隱私

這些報告含**會員個資、奉獻金額，以及 -7 代禱項目裡的健康狀況**
（手術、復健、失智、呼吸中止症等，都指名道姓）。

**Vercel 部署預設是公開網址，任何拿到連結的人都能開。**
把這些內容放上去而沒有存取控制，等於公開會員的醫療隱私。

要上雲端，存取控制不是選配：

- Vercel 的 Deployment Protection（密碼保護）需要 **Pro 方案**
- 或自己實作登入（再加一層要維護的東西）
- 也建議先問過支會，這些資料放雲端是否適當

### 若確定要做，正確的架構

好消息是有現成解法：**你的 Drive 上已經有完整的月份資料夾**
（`2026年06月` 裡就有那 11 份 docx）。所以雲端版可以完全不碰本機：

```
Drive「{上月}」資料夾 ──讀範本──▶ Vercel Function ──寫回──▶ Drive「{本月}」資料夾
                                      │
                             Grafana API（免登入，雲端可直連）
```

需要的改動：

1. **`engine/storage.py`** — 抽象出「取範本 / 存產出」，本機與 Drive 兩種實作
2. **OAuth 改網頁流程** — `Flow` + redirect URI 設成 `https://{你的網域}/api/oauth/callback`，
   refresh token 存 Vercel 環境變數
3. **入口與設定** — `api/index.py` 掛 FastAPI，`vercel.json` 設
   `maxDuration`（產 11 份 docx 加查詢，10 秒不太夠，Pro 可到 60 秒）
4. **`requirements.txt` 瘦身** — `pdfplumber` 相依較重，雲端若不做 PDF 解析可拿掉

工作量大約 1–2 天，主要在 1 和 2。

### 我的建議

先在本機用一到兩個月。核心邏輯（日期規則、Grafana 擷取、docx 更新）
與部署方式完全無關，這段時間的驗證不會白費，而且你會更清楚哪些細節
還要調整——現在就搬雲端，等於在還沒穩定的東西上疊一層複雜度。

真的需要遠端存取時，比 Vercel 更省事的選項：

- **Tailscale**：把你的 Mac 加進私有網路，手機／筆電直接連 `127.0.0.1:8848`，
  零改動、不公開、免費
- **Cloudflare Tunnel**：類似效果，可加 Access 做登入控管
- **Railway / Render**：能跑長時間程序、有持久磁碟，比 serverless 更貼近現在的架構

要走 Vercel 這條路的話再跟我說，我可以把 `storage.py` 和網頁版 OAuth 做出來。
