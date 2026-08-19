"""Google Drive 上傳：OAuth 授權、依月份自動建資料夾、DOCX 覆寫上傳。

行為（依需求）：
  1. 在指定的父資料夾底下尋找「2026年7月」子資料夾；不存在則建立。
  2. 上傳 DOCX：同名檔案存在則「更新既有檔案」（覆寫，保留檔案 ID 與連結），
     不存在則新建。

首次使用：
  1. Google Cloud Console 建立 OAuth 2.0 用戶端（類型：桌面應用程式）
  2. 下載 client_secret JSON，存為 credentials/client_secret.json
  3. 啟動 App，於介面按「連結 Google Drive」完成瀏覽器授權
  4. token 會存到 credentials/token.json，之後自動續用
"""

from __future__ import annotations

import os
from typing import Iterable

SCOPES = ["https://www.googleapis.com/auth/drive"]
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
FOLDER_MIME = "application/vnd.google-apps.folder"

# 使用者提供的目標父資料夾
DEFAULT_PARENT_ID = "15HBrIm4TOJrIMHo6ydHR2bWHN0ZTrWJ5"


class DriveNotConfigured(RuntimeError):
    pass


class DriveAuthExpired(DriveNotConfigured):
    """refresh token 失效（過期／被撤銷／換過 OAuth 用戶端）。

    最常見原因：OAuth 同意畫面還停在「測試中」，Google 只發 7 天有效的
    refresh token。到期後每個要動 Drive 的請求都會 invalid_grant，
    以前會一路冒泡成前端的「Internal Server Error」——所以在這裡就轉成
    看得懂、講得出下一步的訊息。
    """


REAUTH_HINT = (
    "Google Drive 授權已失效（invalid_grant：refresh token 過期或被撤銷）。"
    "重新授權三步驟：① Google Cloud Console → OAuth 同意畫面「發布為正式版」"
    "（停在「測試中」的 token 只有 7 天）；② 本機執行 App，按「連結 Google Drive」"
    "重新取得 token；③ 依 DEPLOY.md 第三章把新的 GOOGLE_REFRESH_TOKEN 更新到 "
    "Vercel 環境變數並重新部署。"
)

CLOUD_AUTH_HINT = (
    "雲端環境開不了瀏覽器，無法在這裡完成 Google 授權。請在本機授權後，把 "
    "GOOGLE_CLIENT_ID / GOOGLE_CLIENT_SECRET / GOOGLE_REFRESH_TOKEN "
    "設到 Vercel 環境變數並重新部署（見 DEPLOY.md 第三章）。"
)


def _is_cloud() -> bool:
    from engine.auth import is_cloud
    return is_cloud()


def _import_libs():
    try:
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaFileUpload
        return Request, Credentials, InstalledAppFlow, build, MediaFileUpload
    except ImportError as exc:
        raise DriveNotConfigured(
            "缺少 Google 套件，請執行：pip install -r requirements.txt"
        ) from exc


class DriveClient:
    def __init__(self, cred_dir: str, parent_id: str = DEFAULT_PARENT_ID):
        self.cred_dir = cred_dir
        self.parent_id = parent_id
        self.client_secret = os.path.join(cred_dir, "client_secret.json")
        self.token_path = os.path.join(cred_dir, "token.json")
        self._service = None
        self._last_error: str | None = None

    # ── 狀態 ─────────────────────────────────────────────
    def env_configured(self) -> bool:
        """雲端授權模式：三個環境變數都在就算已設定。"""
        return all(os.environ.get(k, "").strip() for k in
                   ("GOOGLE_CLIENT_ID", "GOOGLE_CLIENT_SECRET", "GOOGLE_REFRESH_TOKEN"))

    def status(self) -> dict:
        """雲端沒有 client_secret.json 也沒有 token.json，授權走環境變數。

        舊版只看檔案，雲端永遠顯示「缺 client_secret.json」——看起來像沒設定，
        實際上是設定好但 token 過期。兩者要分得出來。
        """
        env = self.env_configured()
        token_file = os.path.exists(self.token_path)
        return {
            "mode": "env" if env else ("file" if token_file else "none"),
            "env_configured": env,
            "client_secret_present": os.path.exists(self.client_secret),
            "token_file_present": token_file,
            "authorized": bool(env or token_file),
            "error": self._last_error,
            "reauth_hint": REAUTH_HINT if self._last_error else None,
            "parent_id": self.parent_id,
            "parent_url": f"https://drive.google.com/drive/folders/{self.parent_id}",
        }

    # ── 換 access token（唯一會碰到 invalid_grant 的地方）──
    def _do_refresh(self, creds, Request) -> None:
        from google.auth.exceptions import RefreshError
        try:
            creds.refresh(Request())
        except RefreshError as exc:
            self._last_error = str(exc)
            raise DriveAuthExpired(REAUTH_HINT) from exc
        except Exception as exc:          # 網路、憑證格式等其他問題
            self._last_error = str(exc)
            raise DriveNotConfigured(f"連線 Google Drive 失敗：{exc}") from exc
        self._last_error = None

    # ── 授權 ─────────────────────────────────────────────
    def _creds_from_env(self):
        """雲端用：純環境變數建立憑證，不碰檔案系統。

        Vercel 沒有持久磁碟也無法開瀏覽器，因此把本機授權好的
        refresh token 放進環境變數，直接換取 access token。
        """
        _, Credentials, _, _, _ = _import_libs()
        cid = os.environ.get("GOOGLE_CLIENT_ID", "").strip()
        secret = os.environ.get("GOOGLE_CLIENT_SECRET", "").strip()
        refresh = os.environ.get("GOOGLE_REFRESH_TOKEN", "").strip()
        if not (cid and secret and refresh):
            return None
        return Credentials(
            token=None, refresh_token=refresh, client_id=cid,
            client_secret=secret, scopes=SCOPES,
            token_uri="https://oauth2.googleapis.com/token")

    def authorize(self, interactive: bool = True):
        Request, Credentials, InstalledAppFlow, build, _ = _import_libs()

        # 優先走環境變數（雲端）
        creds = self._creds_from_env()
        if creds is not None:
            self._do_refresh(creds, Request)
            self._service = build("drive", "v3", credentials=creds,
                                  cache_discovery=False)
            return self._service

        creds = None
        if os.path.exists(self.token_path):
            creds = Credentials.from_authorized_user_file(self.token_path, SCOPES)
        if creds and creds.expired and creds.refresh_token:
            try:
                self._do_refresh(creds, Request)
                self._save(creds)
            except DriveNotConfigured:
                # 本機重新授權的入口：舊 token 失效時直接改走瀏覽器流程，
                # 不要逼使用者先手動刪掉 credentials/token.json。
                if not interactive or _is_cloud():
                    raise
                creds = None
        if not creds or not creds.valid:
            if _is_cloud():
                # serverless 開不了瀏覽器；走到這裡代表環境變數沒設或設錯
                raise DriveNotConfigured(CLOUD_AUTH_HINT)
            if not interactive:
                raise DriveNotConfigured("尚未授權 Google Drive，請先於介面點「連結 Google Drive」")
            if not os.path.exists(self.client_secret):
                raise DriveNotConfigured(
                    f"找不到 {self.client_secret}。請至 Google Cloud Console 建立 "
                    "OAuth 用戶端（桌面應用程式），下載 JSON 後存入此路徑。")
            flow = InstalledAppFlow.from_client_secrets_file(self.client_secret, SCOPES)
            creds = flow.run_local_server(port=0, prompt="consent")
            self._save(creds)
        self._service = build("drive", "v3", credentials=creds, cache_discovery=False)
        return self._service

    def _save(self, creds):
        os.makedirs(self.cred_dir, exist_ok=True)
        with open(self.token_path, "w", encoding="utf-8") as f:
            f.write(creds.to_json())

    @property
    def service(self):
        if self._service is None:
            self.authorize(interactive=False)
        return self._service

    # ── 查詢與下載（儲存層用）─────────────────────────────
    def find_folder(self, name: str, parent_id: str | None = None) -> dict | None:
        parent = parent_id or self.parent_id
        safe = name.replace("'", "\\'")
        q = (f"name = '{safe}' and mimeType = '{FOLDER_MIME}' "
             f"and '{parent}' in parents and trashed = false")
        files = self.service.files().list(
            q=q, fields="files(id,name)", pageSize=1,
            supportsAllDrives=True, includeItemsFromAllDrives=True
        ).execute().get("files", [])
        return files[0] if files else None

    def find_file(self, name: str, folder_id: str) -> str | None:
        safe = name.replace("'", "\\'")
        q = f"name = '{safe}' and '{folder_id}' in parents and trashed = false"
        files = self.service.files().list(
            q=q, fields="files(id)", pageSize=1,
            supportsAllDrives=True, includeItemsFromAllDrives=True
        ).execute().get("files", [])
        return files[0]["id"] if files else None

    def list_files(self, folder_id: str) -> list[dict]:
        out, token = [], None
        while True:
            res = self.service.files().list(
                q=f"'{folder_id}' in parents and trashed = false",
                fields="nextPageToken, files(id,name,mimeType,size)",
                pageSize=200, pageToken=token,
                supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
            out += res.get("files", [])
            token = res.get("nextPageToken")
            if not token:
                return out

    def download(self, file_id: str) -> bytes:
        return self.service.files().get_media(
            fileId=file_id, supportsAllDrives=True).execute()

    # ── 資料夾 ───────────────────────────────────────────
    def ensure_folder(self, name: str, parent_id: str | None = None) -> dict:
        """找不到就建立子資料夾，回傳 {id, name, url, created}。"""
        parent = parent_id or self.parent_id
        safe = name.replace("'", "\\'")
        q = (f"name = '{safe}' and mimeType = '{FOLDER_MIME}' "
             f"and '{parent}' in parents and trashed = false")
        res = self.service.files().list(
            q=q, fields="files(id,name)", pageSize=1,
            supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
        files = res.get("files", [])
        if files:
            fid, created = files[0]["id"], False
        else:
            meta = {"name": name, "mimeType": FOLDER_MIME, "parents": [parent]}
            fid = self.service.files().create(
                body=meta, fields="id", supportsAllDrives=True).execute()["id"]
            created = True
        return {"id": fid, "name": name, "created": created,
                "url": f"https://drive.google.com/drive/folders/{fid}"}

    # ── 上傳（覆寫）───────────────────────────────────────
    def upload_file(self, path: str, folder_id: str, mime: str = DOCX_MIME) -> dict:
        _, _, _, _, MediaFileUpload = _import_libs()
        name = os.path.basename(path)
        safe = name.replace("'", "\\'")
        q = f"name = '{safe}' and '{folder_id}' in parents and trashed = false"
        res = self.service.files().list(
            q=q, fields="files(id,name)", pageSize=1,
            supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
        existing = res.get("files", [])
        media = MediaFileUpload(path, mimetype=mime, resumable=False)

        if existing:
            fid = existing[0]["id"]
            f = self.service.files().update(
                fileId=fid, media_body=media,
                fields="id,name,webViewLink,modifiedTime",
                supportsAllDrives=True).execute()
            action = "overwritten"
        else:
            f = self.service.files().create(
                body={"name": name, "parents": [folder_id]}, media_body=media,
                fields="id,name,webViewLink,modifiedTime",
                supportsAllDrives=True).execute()
            action = "created"
        return {"id": f["id"], "name": f["name"], "action": action,
                "url": f.get("webViewLink"), "modified": f.get("modifiedTime")}

    def upload_month(self, paths: Iterable[str], folder_name: str) -> dict:
        """建立（或取得）月份資料夾並覆寫上傳全部檔案。"""
        folder = self.ensure_folder(folder_name)
        uploaded, errors = [], []
        for p in paths:
            try:
                mime = DOCX_MIME if p.lower().endswith(".docx") else None
                uploaded.append(self.upload_file(p, folder["id"], mime or DOCX_MIME))
            except Exception as exc:
                errors.append({"file": os.path.basename(p), "error": str(exc)})
        return {"folder": folder, "uploaded": uploaded, "errors": errors}
