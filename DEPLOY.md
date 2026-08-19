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
| `GIDEONS_CRED_DIR` | `./credentials` | Google OAuth 憑證與 token |
| `GIDEONS_DRIVE_PARENT` | 專案指定 ID | Drive「月例會」資料夾 |

### Google Drive 授權（首次一次即可）

1. [Google Cloud Console](https://console.cloud.google.com/) 建專案 → 啟用 **Google Drive API**
2. 「API 和服務 → 憑證」→ 建立 **OAuth 用戶端 ID**，類型選**桌面應用程式**
3. 下載 JSON，存成 `credentials/client_secret.json`
4. 啟動 App → 按「連結 Google Drive」→ 瀏覽器完成授權
   token 存到 `credentials/token.json`，之後自動續用

> ⚠️ 若同意畫面還在「測試中」，要把自己加進測試使用者，且 refresh token
> **7 天就過期**。這不是小事：token 一過期，雲端版每個要寫 Drive 的動作
> 都會失敗（2026-08 就是這樣整站看起來壞掉）。
> **長期使用請務必把應用程式狀態發布為正式版**，再重新取一次 token。

### 疑難排解

| 狀況 | 處理 |
|---|---|
| `No module named fastapi` | `.venv` 沒啟用，直接跑 `python3 run.py` |
| `python3 -m venv` 失敗 | macOS 缺 Command Line Tools：`xcode-select --install` |
| Drive 顯示「待授權」 | 按「連結 Google Drive」；測試模式 token 7 天過期 |
| 找不到範本資料夾 | 確認 `~/Documents/海山支會/{年}年{月}月例會` 存在 |
| 擷取資料失敗 | 確認網路可連 `gideons-dashboard.pointing.tw` |
| `invalid_grant` / 「授權已失效」 | refresh token 過期。同意畫面**發布為正式版**後重新授權，再更新 Vercel 的 `GOOGLE_REFRESH_TOKEN` |

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
| 儲存 | `~/Documents/海山支會` | Google Drive（`/tmp` 僅單次請求） |
| 授權 | `credentials/token.json` | 環境變數 refresh token |
| 密碼 | 可不設 | **必填**，未設回 503 |
| 範本 | 上月本機資料夾 | Drive 的「{年}年{MM}月」資料夾 |

App 偵測到 `VERCEL` 環境變數就自動切成 Drive 儲存層並改用 `/tmp`。

### 1. 取出 refresh token

在本機已授權的狀態下執行（**輸出含機密，不要貼到任何聊天或公開場合**）：

```bash
python3 -c "
import json
t = json.load(open('credentials/token.json'))
c = json.load(open('credentials/client_secret.json'))['installed']
print('GOOGLE_CLIENT_ID    =', c['client_id'])
print('GOOGLE_CLIENT_SECRET=', c['client_secret'])
print('GOOGLE_REFRESH_TOKEN=', t['refresh_token'])
"
```

> ⚠️ 這組 token 有你 Drive 的**完整讀寫權限**，不限月例會資料夾。
> 若外流，立刻到 [Google 帳號權限頁](https://myaccount.google.com/permissions)
> 移除此應用程式，再重新授權。

### 2. 建立專案

1. [vercel.com](https://vercel.com) → **Add New… → Project** → 匯入 repo
2. Team 選 `richardlai1973-3671s-projects`
3. **Root Directory 留空**（repo 根目錄就是 app 本身）
4. 先設環境變數，再 Deploy

### 3. 環境變數

Project → Settings → Environment Variables，全部套用到 Production：

| 變數 | 值 |
|---|---|
| `APP_PASSWORD` | 你設定的密碼 |
| `SESSION_SECRET` | `python3 -c "import secrets;print(secrets.token_urlsafe(32))"` |
| `GOOGLE_CLIENT_ID` | 步驟 1 取得 |
| `GOOGLE_CLIENT_SECRET` | 步驟 1 取得 |
| `GOOGLE_REFRESH_TOKEN` | 步驟 1 取得 |
| `GIDEONS_DRIVE_PARENT` | `15HBrIm4TOJrIMHo6ydHR2bWHN0ZTrWJ5` |

`STORAGE` 會自動設為 `drive`，可省略。

### 4. 驗證

開啟部署網址 → 應出現**密碼登入頁**（看到登入頁就表示路由通了）。
登入後檢查狀態列：儲存應顯示 Google Drive、範本應顯示已就緒。

流程與本機相同，差別是產出後**自動上傳回 Drive** 的「{年}年{MM}月」資料夾。

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

#### 「Internal Server Error」／狀態列顯示「⛔ 授權已失效」

Drive 的 refresh token 過期或被撤銷（`invalid_grant`）。現在介面會直接寫出原因
與三步驟，照做即可：

1. Google Cloud Console → OAuth 同意畫面 → **發布為正式版**（測試中只有 7 天）
2. 本機跑 `python3 run.py` → 按「連結 Google Drive」重新授權
3. 用本章「1. 取出 refresh token」的指令取新值 → 更新 Vercel 環境變數
   `GOOGLE_REFRESH_TOKEN` → **Redeploy**

授權壞掉時系統不會整個停擺：資料仍寫入雲端暫存區，可正常產出並「下載 ZIP」，
只是離開網頁後不保留，介面會以黃字提醒。

#### 部署成功但 500 或頁面空白

看 Vercel 的 **Logs** 分頁。最常見是環境變數漏設，此時狀態列的
「Drive」會顯示未授權。

### 已知限制

| 項目 | 說明 |
|---|---|
| 執行時間 | `maxDuration: 60`，Hobby 方案上限即 60 秒。產 11 份 docx 加 Drive 上下傳若逼近上限，需改 Pro 或分批 |
| `/tmp` 不保證保留 | 每次都從 Drive 重新同步，不影響正確性，只是稍慢 |
| Drive API 配額 | 每次產出約 20–30 次呼叫，個人帳號額度充足 |
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

部署到公開網址後，**密碼是唯一的一道門**，而 `GOOGLE_REFRESH_TOKEN`
讓這個網站能操作你**整個 Drive**——不只月例會資料夾。

報告內容含會員個資、奉獻金額，以及 `-7代禱項目` 裡的**具名健康狀況**
（手術、復健、失智等，涉及會員本人與家屬）。這些人同意在支會內部傳閱，
不等於同意公開上網。

因此：

- 密碼用隨機字串，不要用支會名或常見單字
- 網址不要公開張貼
- 懷疑外流時：先到 Google 帳號權限頁撤銷授權，再換 `APP_PASSWORD`
- 「只有我們知道網址」不是存取控制——Vercel 網址會出現在
  Certificate Transparency 公開紀錄，有人專門在掃新網域
