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

    def take_warnings(self) -> list[str]:
        """取出並清空累積的警告（本機版永遠沒有）。"""
        return []

    def find_plan(self, period: str) -> Path | None:
        """年度贈經計畫（整個財年固定的學校配送排程）。找不到回傳 None。"""
        raise NotImplementedError

    def save_plan(self, period: str, src: Path) -> Path:
        """存入新年度的贈經計畫。"""
        raise NotImplementedError


# ── 本機 ─────────────────────────────────────────────────
PLAN_FOLDER = "贈經計畫"


def plan_filename(period: str) -> str:
    return f"{period}_聖經配送計畫.xlsx"


class LocalStorage(Storage):
    mode = "local"

    def __init__(self, base_dir: str | Path, plan_dir: str | Path | None = None):
        self.base = Path(base_dir)
        self.plan_dir = Path(plan_dir) if plan_dir else self.base / PLAN_FOLDER

    def find_plan(self, period: str) -> Path | None:
        p = self.plan_dir / plan_filename(period)
        return p if p.exists() else None

    def save_plan(self, period: str, src: Path) -> Path:
        self.plan_dir.mkdir(parents=True, exist_ok=True)
        dst = self.plan_dir / plan_filename(period)
        shutil.copy2(src, dst)
        return dst

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
        self.warnings: list[str] = []

    # ── 警告蒐集 ─────────────────────────────────────────
    # Drive 掛掉時不再讓請求 500：資料照樣寫進 /tmp，只是這次工作階段結束後
    # 不保留。使用者仍能產出並「下載 ZIP」，但**必須看得到**沒同步這件事，
    # 否則就變成 CLAUDE.md 講的「顯示成功卻什麼都沒發生」。
    def _warn(self, msg: str) -> None:
        if msg not in self.warnings:
            self.warnings.append(msg)

    def take_warnings(self) -> list[str]:
        w, self.warnings = self.warnings, []
        return w

    def _safe(self, fn, default=None):
        """Drive 未設定或暫時失敗時回傳預設值，讓狀態頁仍能顯示而非 500。"""
        try:
            return fn()
        except Exception as exc:
            self._warn(f"Google Drive 讀取失敗：{exc}")
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
        try:
            folder = self.drive.ensure_folder(meta.drive_folder_name)
            self.drive.upload_file(str(local), folder["id"], "application/json")
        except Exception as exc:
            self._warn(f"資料模型只留在本次工作階段，未同步到 Drive：{exc}")

    # ── 年度贈經計畫 ─────────────────────────────────────
    def find_plan(self, period: str) -> Path | None:
        """從 Drive 的「贈經計畫」資料夾取得，快取到 /tmp。

        本機版讀專案旁的資料夾，那條路徑在 serverless 不存在，
        因此雲端一律走 Drive。
        """
        name = plan_filename(period)
        cached = self.tmp_root / PLAN_FOLDER / name
        if cached.exists():
            return cached

        folder = self._safe(lambda: self.drive.find_folder(PLAN_FOLDER))
        if not folder:
            return None
        fid = self._safe(lambda: self.drive.find_file(name, folder["id"]))
        if not fid:
            return None
        data = self._safe(lambda: self.drive.download(fid))
        if data is None:
            return None
        cached.parent.mkdir(parents=True, exist_ok=True)
        cached.write_bytes(data)
        return cached

    def save_plan(self, period: str, src: Path) -> Path:
        staged = self.tmp_root / PLAN_FOLDER / plan_filename(period)
        staged.parent.mkdir(parents=True, exist_ok=True)
        if Path(src) != staged:
            shutil.copy2(src, staged)
        try:
            folder = self.drive.ensure_folder(PLAN_FOLDER)
            self.drive.upload_file(str(staged), folder["id"])
        except Exception as exc:
            self._warn(f"贈經計畫只留在本次工作階段，未同步到 Drive：{exc}")
        return staged

    # ── 上傳產出 ─────────────────────────────────────────
    def publish(self, meta) -> dict:
        wd = self.work_dir(meta)
        paths = sorted(str(p) for p in wd.glob("*.docx"))
        paths += [str(p) for p in wd.glob("*.xlsx")]
        if not paths:
            return {"published": False, "error": "工作區無產出檔"}
        try:
            res = self.drive.upload_month(paths, meta.drive_folder_name)
        except Exception as exc:
            self._warn(f"產出未上傳 Drive：{exc}")
            return {"published": False, "error": str(exc)}
        return {"published": True, **res}


# ── 建立 ─────────────────────────────────────────────────
def make_storage(mode: str, base_dir, drive_client, plan_dir=None) -> Storage:
    if mode == "drive":
        if drive_client is None:
            raise RuntimeError("STORAGE=drive 但未提供 Drive 用戶端")
        return DriveStorage(drive_client)
    return LocalStorage(base_dir, plan_dir)
