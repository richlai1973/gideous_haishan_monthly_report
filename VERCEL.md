# Vercel 部署步驟

兩個版本共用同一份 `app.py`，靠環境變數自動切換：

| | 本機 `python3 run.py` | Vercel |
|---|---|---|
| 進入點 | `run.py` → `app.py` | `app.py`（Vercel 自動偵測） |
| 儲存 | `~/Documents/海山支會` | Google Drive（`/tmp` 只在單次請求內用） |
| 授權 | `credentials/token.json` | 環境變數 refresh token |
| 密碼 | 可不設 | **必填**，未設會回 503 |
| 範本 | 上月本機資料夾 | Drive 的「{年}年{MM}月」資料夾 |

App 偵測到 `VERCEL` 環境變數就自動切成 Drive 儲存層並改用 `/tmp`，
不需要額外的進入點檔案。

> **不要另外建 `api/index.py`。** Vercel 會在
> `app.py`／`index.py`／`server.py`／`main.py`／`wsgi.py`／`asgi.py`
> （根目錄或 `src/`、`app/`、`api/` 內）尋找名為 `app` 的 FastAPI 實例。
> 同時存在兩個進入點會造成路由解析錯誤，出現 `404: NOT_FOUND`。

---

## 1. 取出 refresh token

在本機已授權的狀態下執行（**不要把輸出貼到任何聊天或公開場合**）：

```bash
cd ~/Claude/Projects/基甸會月例會報告/gideons-report-app
python3 -c "
import json
t = json.load(open('credentials/token.json'))
c = json.load(open('credentials/client_secret.json'))['installed']
print('GOOGLE_CLIENT_ID    =', c['client_id'])
print('GOOGLE_CLIENT_SECRET=', c['client_secret'])
print('GOOGLE_REFRESH_TOKEN=', t['refresh_token'])
"
```

> ⚠️ 這組 token 有你 Drive 的**完整讀寫權限**，不限於月例會資料夾。
> 若外流，請立刻到 [Google 帳號權限頁](https://myaccount.google.com/permissions)
> 移除此應用程式的存取權，再重新授權一次。

---

## 2. 建立 Vercel 專案

1. [vercel.com](https://vercel.com) → **Add New… → Project**
2. 匯入 GitHub repo `gideous_haishan_monthly_report`
3. Team 選 `richardlai1973-3671s-projects`，專案名建議 `gideons-haishan-report`
4. Framework Preset 選 **Other**（`vercel.json` 已設定好）
5. 先**不要**按 Deploy，先設環境變數（下一步）

---

## 3. 環境變數

Project → Settings → Environment Variables，全部套用到 Production：

| 變數 | 值 | 說明 |
|---|---|---|
| `APP_PASSWORD` | 你設定的密碼 | 未設會回 503 停止服務 |
| `SESSION_SECRET` | `python3 -c "import secrets;print(secrets.token_urlsafe(32))"` | 讓換密碼不會登出所有人 |
| `GOOGLE_CLIENT_ID` | 步驟 1 取得 | |
| `GOOGLE_CLIENT_SECRET` | 步驟 1 取得 | |
| `GOOGLE_REFRESH_TOKEN` | 步驟 1 取得 | |
| `GIDEONS_DRIVE_PARENT` | `15HBrIm4TOJrIMHo6ydHR2bWHN0ZTrWJ5` | Drive「月例會」資料夾 |
| `STORAGE` | `drive` | 偵測到 Vercel 會自動設定，可省略 |

設完按 **Deploy**。

---

## 疑難排解

### `404: NOT_FOUND`

依序檢查：

1. **有沒有多個進入點**：根目錄 `app.py` 之外若還有 `api/index.py`，
   Vercel 會解析錯誤。只留一個。
2. **`vercel.json` 不要寫 `rewrites`**：整個 FastAPI 會被打包成**單一
   function**，所有路由由 FastAPI 自己處理，不需要也不應該再加 rewrite。
   多加了反而會把請求導到不存在的路徑。
3. **`functions` 不要指定 `runtime`**：只需要 `maxDuration` 之類的設定，
   鍵名要對應實際的進入點檔案（本專案是 `app.py`）。

正確的 `vercel.json` 就是這樣，沒有其他東西：

```json
{
  "$schema": "https://openapi.vercel.sh/vercel.json",
  "functions": { "app.py": { "maxDuration": 60 } }
}
```

### 部署成功但頁面空白／500

到 Vercel 的 **Logs** 分頁看 function 的錯誤訊息。最常見是環境變數漏設，
此時狀態列的「Drive」會顯示未授權。

---

## 4. 驗證

開啟部署網址，應該看到密碼登入頁。登入後檢查狀態列：

- **儲存**：Google Drive
- **範本**：需顯示已就緒。若顯示不存在，確認 Drive 上有
  「{上月}」資料夾（如 `2026年06月`）且裡面有 11 份 docx

流程與本機相同，差別是產出後會**自動上傳回 Drive** 的「{年}年{MM}月」資料夾。

---

## 已知限制

**執行時間**：`vercel.json` 設 `maxDuration: 60`。Hobby 方案上限即 60 秒。
產 11 份 docx 加上 Drive 上下傳，若逼近上限請改用 Pro 或拆成分批產出。

**`/tmp` 不保證保留**：同一個 lambda 實例會沿用，但可能隨時被回收。
程式已設計成每次都從 Drive 重新同步，所以不影響正確性，只是稍慢。

**Drive API 配額**：每次產出約 20–30 次 API 呼叫，個人帳號的額度綽綽有餘。

**PDF 解析**：`pdfplumber` 相依較重，若部署尺寸超限可從
`requirements.txt` 移除（只會失去 PDF 上傳解析，其他功能不受影響）。

---

## 安全提醒

部署後任何人只要有網址就會看到登入頁。密碼是唯一的一道門，而且
`GOOGLE_REFRESH_TOKEN` 讓這個網站能操作你**整個 Drive**——不只月例會資料夾。

因此：

- 密碼請用隨機字串，不要用支會名或常見單字
- 不要把網址公開貼在 LINE 群組以外的地方
- 若懷疑外流，先到 Google 帳號權限頁撤銷授權，再換 `APP_PASSWORD`
