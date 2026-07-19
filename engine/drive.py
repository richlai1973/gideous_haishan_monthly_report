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

    # ── 狀態 ─────────────────────────────────────────────
    def status(self) -> dict:
        return {
            "client_secret_present": os.path.exists(self.client_secret),
            "authorized": os.path.exists(self.token_path),
            "parent_id": self.parent_id,
            "parent_url": f"https://drive.google.com/drive/folders/{self.parent_id}",
        }

    # ── 授權 ─────────────────────────────────────────────
    def authorize(self, interactive: bool = True):
        Request, Credentials, InstalledAppFlow, build, _ = _import_libs()
        creds = None
        if os.path.exists(self.token_path):
            creds = Credentials.from_authorized_user_file(self.token_path, SCOPES)
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
            self._save(creds)
        if not creds or not creds.valid:
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
