# 部署指南

本機執行、推上 GitHub、部署到 Vercel。三者共用同一份 `app.py`。

- 系統功能與每月操作流程 → [README.md](README.md)
- 開發規範與領域規則 → [CLAUDE.md](CLAUDE.md)

---

## 一、本機執行

```bash
cd ~/Claude/Projects/基甸會月例會報告/gideons-report-app
python3 run.py              # 或 bash run.sh
```

首次會自動建 `.venv`、安裝套件。啟動後開 http://127.0.0.1:8848。

| 參數 | 說明 |
|---|---|
| `--port 9000` | 換連接埠 |
| `--open` | 啟動後自動開瀏覽器 |
| `--reload` | 改程式碼自動重載（開發用） |

### 本機設定

複製 `.env.example` 成 `.env`（已被 `.gitignore` 排除）：

```bash
cp .env.example .env
```

| 變數 | 預設 | 說明 |
|---|---|---|
| `APP_PASSWORD` | 無 | 未設時本機不驗證密碼 |
| `SESSION_SECRET` | 由密碼衍生 | 設了之後換密碼不會登出所有人 |
| `GIDEONS_BASE_DIR` | `~/Documents/海山支會` | 每月工作資料夾根目錄 |
| `GIDEONS_PLAN_DIR` | `../贈經計畫` | 年度聖經配送計畫存放處 |
| `GIDEONS_TEMPLATE_DIR` | `./templates` | 固定範本（找不到上月資料夾時的回退）|

### 固定範本（雲端唯一的範本來源）

`templates/` 裡放一組完整的 11 份 docx，跟程式一起部署。
本機執行時**優先用 `~/Documents/海山支會/{上月}月例會`**，找不到才回退到它；
雲端沒有那個資料夾，所以永遠用它。

畫面「範本來源／範本月份」會顯示這次實際用了哪一組——**務必確認**，
因為固定範本停在某個月份（例如 115年7月），日期替換是以它為基準。

想更新雲端的起點（建議每隔幾個月做一次）：

```bash
cd ~/Claude/Projects/基甸會月例會報告/gideons-report-app
rm -f templates/*.docx
cp ~/Documents/海山支會/2026年8月月例會/月例會議程*.docx templates/
git add templates && git commit -m "更新固定範本至 115年8月" && git push
```

> ⚠️ 這些 docx 含會員個資與 -7 的具名健康資訊，`.gitignore` 對 `templates/`
> 開了例外才進得了版控。**repo 必須維持 private。**

年度贈經計畫同理：`贈經計畫/{period}_聖經配送計畫.xlsx` 也在 repo 裡，
換財年時 commit 一次即可（雲端上傳只在本次工作階段有效）。

### 疑難排解

| 狀況 | 處理 |
|---|---|
| `No module named fastapi` | `.venv` 沒啟用，直接跑 `python3 run.py` |
| `python3 -m venv` 失敗 | macOS 缺 Command Line Tools：`xcode-select --install` |
| 找不到範本資料夾 | 確認 `~/Documents/海山支會/{年}年{月}月例會` 存在；沒有的話會自動用 `templates/` |
| 範本月份不對 | 看畫面「範本來源」。要換就更新 `templates/` 內容 |
| 擷取資料失敗 | 確認網路可連 `gideons-dashboard.pointing.tw` |

---

## 二、推上 GitHub

```bash
bash setup-git.sh          # 建 repo、檢查沒有個資、commit
git push -u origin main
```

`setup-git.sh` 會在 commit 前列出所有版控檔案，並擋下
`.docx` / `.xlsx` / `credentials/` / `token.json`。

> **repo 建議設 private。** 程式碼本身沒有機密（憑證與資料都排除），
> 但 `engine/grafana.py` 含 dashboard URL 與 `var-ref=5394` 等支會識別參數。

> ⚠️ **`.gitignore` 不要用 `*.json`**。曾因此把 `vercel.json` 擋在 repo 外，
> 部署看似成功實則整站 404。要擋憑證就指名擋。

---

## 三、Vercel 部署

兩個版本靠環境變數自動切換，程式碼完全相同：

| | 本機 | Vercel |
|---|---|---|
| 進入點 | `run.py` → `app.py` | `api/index.py` → `app.py` |
| 儲存 | `~/Documents/海山支會`（永久） | `/tmp`，**單次工作階段** |
| 範本 | 上月本機資料夾 →（沒有就）`templates/` | `templates/` |
| 年度計畫 | 專案旁的 `贈經計畫/` | repo 的 `贈經計畫/` |
| 輸出 | 下載 ZIP | 下載 ZIP（唯一管道） |
| 密碼 | 可不設 | **必填**，未設回 503 |

App 偵測到 `VERCEL` 環境變數就把工作區改到 `/tmp`。沒有任何外部儲存、
沒有任何會過期的憑證——這是刻意的。

### 1. 建立專案

1. [vercel.com](https://vercel.com) → **Add New… → Project** → 匯入 repo
2. Team 選 `richardlai1973-3671s-projects`
3. **Root Directory 留空**（repo 根目錄就是 app 本身）
4. 先設環境變數，再 Deploy

### 2. 環境變數

Project → Settings → Environment Variables，全部套用到 Production：

| 變數 | 值 |
|---|---|
| `APP_PASSWORD` | 你設定的密碼 |
| `SESSION_SECRET` | `python3 -c "import secrets;print(secrets.token_urlsafe(32))"` |

只有這兩個。舊版的 `GOOGLE_CLIENT_ID` / `GOOGLE_CLIENT_SECRET` /
`GOOGLE_REFRESH_TOKEN` / `GIDEONS_DRIVE_PARENT` **可以全部刪掉**。

### 3. 驗證

開啟部署網址 → 應出現**密碼登入頁**（看到登入頁就表示路由通了）。
登入後檢查狀態列：「範本來源」應顯示「內建固定範本（115年X月）」、
年度贈經計畫應顯示已就緒，畫面上方會有一行黃字提醒產出只留在本次工作階段。

流程與本機相同，差別是**做完要立刻按「下載 ZIP」**——雲端不留檔。

### Vercel 疑難排解

#### `404: NOT_FOUND`

這個專案的 Vercel 設定走**傳統模式：function 只認 `api/` 目錄**
（官方文件寫的「根目錄 `app.py` 自動偵測」是新版框架預設，此專案不適用）。

以下三項**必須同時成立**，缺一即 404：

1. `api/index.py` 匯出名為 `app` 的 FastAPI 實例
2. `vercel.json` 的 `functions` 鍵是 `api/index.py`，且**不要**指定 `runtime`
3. `rewrites` 把 `/(.*)` 導向 `/api/index`（否則只有 `/api/*` 有路由）

正確內容就是這樣，沒有其他東西：

```json
{
  "$schema": "https://openapi.vercel.sh/vercel.json",
  "functions": { "api/index.py": { "maxDuration": 60 } },
  "rewrites": [{ "source": "/(.*)", "destination": "/api/index" }]
}
```

另外務必確認 `vercel.json` **真的在 repo 裡**（見上方 `.gitignore` 警告）。

#### 建置失敗

> The pattern "app.py" defined in `functions` doesn't match any
> Serverless Functions inside the `api` directory.

→ `functions` 的鍵指到了 `api/` 以外的檔案，改成 `api/index.py`。

#### 部署成功但 500 或頁面空白

看 Vercel 的 **Logs** 分頁。最常見是 `APP_PASSWORD` 漏設（會回 503）。
其他錯誤現在都會以中文 JSON 回傳（`{"detail": "伺服器錯誤：..."}`），
介面直接顯示，不會再只看到「Internal Server Error」。

### 已知限制

| 項目 | 說明 |
|---|---|
| 執行時間 | `maxDuration: 60`，Hobby 方案上限即 60 秒。產 11 份 docx 綽綽有餘 |
| `/tmp` 不保證保留 | **雲端沒有持久儲存**：產出後要立刻下載 ZIP，否則可能要重跑 |
| 範本會過時 | 固定範本停在 commit 當下的月份；隔幾個月更新一次比較好 |
| 部署尺寸 | 若超限可從 `requirements.txt` 移除 `pdfplumber`（只失去 PDF 解析） |

---

## 四、其他遠端存取方式

不一定要用 Vercel。以下更貼近現有架構：

| 方案 | 說明 |
|---|---|
| **Tailscale** | 把 Mac 加進私有網路，手機／筆電直接連 `127.0.0.1:8848`。程式零改動、不公開、免費 |
| **Cloudflare Tunnel** | 類似效果，可加 Access 做登入控管 |
| **Railway / Render** | 有持久磁碟、可跑長時間程序，比 serverless 更貼近原架構 |

---

## 安全提醒

部署到公開網址後，**密碼是唯一的一道門**。而且 `templates/` 裡那 11 份
docx 就在 repo 裡、也在部署的檔案系統裡——**repo 必須 private**。

報告內容含會員個資、奉獻金額，以及 `-7代禱項目` 裡的**具名健康狀況**
（手術、復健、失智等，涉及會員本人與家屬）。這些人同意在支會內部傳閱，
不等於同意公開上網。

因此：

- 密碼用隨機字串，不要用支會名或常見單字
- 網址不要公開張貼
- 懷疑外流時：換 `APP_PASSWORD` 並重新部署
- 「只有我們知道網址」不是存取控制——Vercel 網址會出現在
  Certificate Transparency 公開紀錄，有人專門在掃新網域
