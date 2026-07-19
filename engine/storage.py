"""儲存層：本機檔案系統 ／ Google Drive。

雲端（Vercel）沒有持久磁碟，但 Drive 上本來就有每月資料夾
（如「2026年06月」內含 11 份 docx），因此直接把 Drive 當儲存層：

    Drive「{上月}」──下載範本──▶ /tmp 工作區 ──產出──▶ 上傳回 Drive「{本月}」

`/tmp` 只在單次請求內存在，跨請求需要保存的只有 model.json，
同樣存在 Drive 的月份資料夾裡。

這樣做的好處是 **generate.py 完全不用改** —— 它照樣對著本機路徑工作，
差別只在那個路徑是永久資料夾還是 /tmp。
"""

from __future__ import annotations

import json
import os
import shutil
import tempfile
from pathlib import Path

MODEL_NAME = "model.json"


class Storage:
    """介面。work_dir() 回傳一個可直接讀寫 docx 的本機路徑。"""

    mode = "base"

    def work_dir(self, meta) -> Path:
        raise NotImplementedError

    def template_dir(self, meta) -> Path | None:
        raise NotImplementedError

    def load_model(self, meta) -> dict:
        raise NotImplementedError

    def save_model(self, meta, model: dict) -> None:
        raise NotImplementedError

    def publish(self, meta) -> dict:
        """把工作區產出送到最終位置。本機版不需要動作。"""
        return {"published": False}


# ── 本機 ─────────────────────────────────────────────────
class LocalStorage(Storage):
    mode = "local"

    def __init__(self, base_dir: str | Path):
        self.base = Path(base_dir)

    def work_dir(self, meta) -> Path:
        d = self.base / meta.work_dir_name
        d.mkdir(parents=True, exist_ok=True)
        (d / "_inputs").mkdir(exist_ok=True)
        return d

    def template_dir(self, meta) -> Path | None:
        d = self.base / f"{meta.prev_year}年{meta.prev_month}月月例會"
        return d if d.is_dir() else None

    def _model_path(self, meta) -> Path:
        return self.work_dir(meta) / "_inputs" / MODEL_NAME

    def load_model(self, meta) -> dict:
        p = self._model_path(meta)
        if p.exists():
            return json.loads(p.read_text(encoding="utf-8"))
        return {}

    def save_model(self, meta, model: dict) -> None:
        p = self._model_path(meta)
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(json.dumps(model, ensure_ascii=False, indent=2, default=str),
                     encoding="utf-8")


# ── Google Drive（雲端）──────────────────────────────────
class DriveStorage(Storage):
    """把 Drive 月份資料夾當成儲存層。

    每次請求都在 /tmp 建工作區；需要範本時從 Drive 上月資料夾拉下來，
    產出後再 publish() 上傳回 Drive 本月資料夾（同名覆寫）。
    """

    mode = "drive"

    def __init__(self, drive_client, tmp_root: str | None = None):
        self.drive = drive_client
        self.tmp_root = Path(tmp_root or tempfile.gettempdir()) / "gideons"
        self._synced: set[str] = set()

    @staticmethod
    def _safe(fn, default=None):
        """Drive 未設定或暫時失敗時回傳預設值，讓狀態頁仍能顯示而非 500。"""
        try:
            return fn()
        except Exception:
            return default

    # ── /tmp 工作區 ──────────────────────────────────────
    def work_dir(self, meta) -> Path:
        d = self.tmp_root / meta.work_dir_name
        d.mkdir(parents=True, exist_ok=True)
        (d / "_inputs").mkdir(exist_ok=True)
        # /tmp 在同一個 lambda 實例中可能已有殘留，需先確保內容是最新的
        if meta.work_dir_name not in self._synced:
            self._pull(meta, d)
            self._synced.add(meta.work_dir_name)
        return d

    def ready(self) -> tuple[bool, str]:
        """Drive 是否可用。回傳 (可用, 說明)，供狀態端點顯示。"""
        try:
            self.drive.service.files().list(
                q="trashed = false", pageSize=1, fields="files(id)").execute()
            return True, "已連線"
        except Exception as exc:
            return False, str(exc)

    def _pull(self, meta, dest: Path) -> int:
        """把 Drive 本月資料夾既有的檔案拉到工作區（續作已產出的月份用）。"""
        folder = self._safe(lambda: self.drive.find_folder(meta.drive_folder_name))
        if not folder:
            return 0
        n = 0
        for f in (self._safe(lambda: self.drive.list_files(folder["id"]), []) or []):
            if f["name"].endswith((".docx", ".xlsx", ".json")):
                data = self._safe(lambda fid=f["id"]: self.drive.download(fid))
                if data is None:
                    continue
                target = dest / "_inputs" / f["name"] if f["name"] == MODEL_NAME \
                    else dest / f["name"]
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(data)
                n += 1
        return n

    def template_dir(self, meta) -> Path | None:
        """把 Drive 上月資料夾下載成一個暫時的範本目錄。"""
        name = f"{meta.prev_year}年{meta.prev_month:02d}月"
        folder = self._safe(lambda: self.drive.find_folder(name))
        if not folder:
            return None
        d = self.tmp_root / f"_tpl_{name}"
        if d.exists() and any(d.glob("*.docx")):
            return d
        d.mkdir(parents=True, exist_ok=True)
        got = 0
        for f in (self._safe(lambda: self.drive.list_files(folder["id"]), []) or []):
            if f["name"].endswith(".docx"):
                data = self._safe(lambda fid=f["id"]: self.drive.download(fid))
                if data is None:
                    continue
                (d / f["name"]).write_bytes(data)
                got += 1
        return d if got else None

    # ── model.json 存 Drive ──────────────────────────────
    def load_model(self, meta) -> dict:
        local = self.work_dir(meta) / "_inputs" / MODEL_NAME
        if local.exists():
            return json.loads(local.read_text(encoding="utf-8"))
        folder = self._safe(lambda: self.drive.find_folder(meta.drive_folder_name))
        if folder:
            fid = self._safe(lambda: self.drive.find_file(MODEL_NAME, folder["id"]))
            if fid:
                data = self.drive.download(fid)
                local.parent.mkdir(parents=True, exist_ok=True)
                local.write_bytes(data)
                return json.loads(data.decode("utf-8"))
        return {}

    def save_model(self, meta, model: dict) -> None:
        local = self.work_dir(meta) / "_inputs" / MODEL_NAME
        local.parent.mkdir(parents=True, exist_ok=True)
        payload = json.dumps(model, ensure_ascii=False, indent=2, default=str)
        local.write_text(payload, encoding="utf-8")
        folder = self.drive.ensure_folder(meta.drive_folder_name)
        self.drive.upload_file(str(local), folder["id"], "application/json")

    # ── 上傳產出 ─────────────────────────────────────────
    def publish(self, meta) -> dict:
        wd = self.work_dir(meta)
        paths = sorted(str(p) for p in wd.glob("*.docx"))
        paths += [str(p) for p in wd.glob("*.xlsx")]
        if not paths:
            return {"published": False, "error": "工作區無產出檔"}
        res = self.drive.upload_month(paths, meta.drive_folder_name)
        return {"published": True, **res}


# ── 建立 ─────────────────────────────────────────────────
def make_storage(mode: str, base_dir, drive_client) -> Storage:
    if mode == "drive":
        if drive_client is None:
            raise RuntimeError("STORAGE=drive 但未提供 Drive 用戶端")
        return DriveStorage(drive_client)
    return LocalStorage(base_dir)
