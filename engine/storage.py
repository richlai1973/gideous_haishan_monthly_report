"""儲存層：本機檔案系統。

三種來源，語意不同，別混在一起：

  工作區 work_dir   本機＝`~/Documents/海山支會/{年}年{月}月例會`
                    雲端＝`/tmp/gideons/...`（**單次工作階段**，隨時會消失）
  範本 template     優先用上月工作資料夾；沒有就退到 repo 內建的固定範本
  年度贈經計畫      整個財年固定，逐一掃 plan_dirs，找到第一份就用

固定範本讓雲端版不必依賴任何外部儲存（原本靠 Google Drive，token 一過期
整站就跟著壞）。代價是它停在某一個月份——所以**範本月份要從檔名讀出來**，
不能假設「範本一定是上個月」，否則日期替換會全部落空。
"""

from __future__ import annotations

import json
import re
import shutil
from pathlib import Path

MODEL_NAME = "model.json"
PLAN_FOLDER = "贈經計畫"

# 月例會議程115年7月-1議程.docx → (115, 7)
_RE_TPL_MONTH = re.compile(r"月例會議程\s*(\d{2,3})\s*年\s*(\d{1,2})\s*月")


def plan_filename(period: str) -> str:
    return f"{period}_聖經配送計畫.xlsx"


def template_month_of(directory: Path) -> tuple[int, int] | None:
    """從範本檔名讀出它是哪一個月（回傳西元年、月）。

    固定範本不會跟著月份走，`meta.prev_*` 若照「上個月」推導就會對不上，
    docx_utils.update_dates() 要換的字串一個都找不到——文件會看似產出成功
    卻整份停留在舊日期。所以範本月份一律以檔名為準。
    """
    for p in sorted(directory.glob("*.docx")):
        m = _RE_TPL_MONTH.search(p.name)
        if m:
            return int(m.group(1)) + 1911, int(m.group(2))
    return None


class Storage:
    """work_dir() 回傳一個可直接讀寫 docx 的本機路徑。"""

    mode = "local"

    def take_warnings(self) -> list[str]:
        return []


class LocalStorage(Storage):
    def __init__(self, base_dir: str | Path,
                 plan_dirs: list[str | Path] | None = None,
                 fixed_template_dir: str | Path | None = None):
        self.base = Path(base_dir)
        self.plan_dirs = [Path(p) for p in (plan_dirs or [])] or [self.base / PLAN_FOLDER]
        self.fixed = Path(fixed_template_dir) if fixed_template_dir else None

    # ── 工作區 ───────────────────────────────────────────
    def work_dir(self, meta) -> Path:
        d = self.base / meta.work_dir_name
        d.mkdir(parents=True, exist_ok=True)
        (d / "_inputs").mkdir(exist_ok=True)
        return d

    # ── 範本 ─────────────────────────────────────────────
    def month_template_dir(self, meta) -> Path | None:
        """上月工作資料夾。內容是上個月真正送出的版本，優先用。"""
        d = self.base / f"{meta.prev_year}年{meta.prev_month}月月例會"
        return d if d.is_dir() and any(d.glob("*.docx")) else None

    def fixed_template_dir(self) -> Path | None:
        """repo 內建的固定範本（雲端唯一的來源）。"""
        if self.fixed and self.fixed.is_dir() and any(self.fixed.glob("*.docx")):
            return self.fixed
        return None

    def template_dir(self, meta) -> Path | None:
        return self.month_template_dir(meta) or self.fixed_template_dir()

    # ── 資料模型 ─────────────────────────────────────────
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

    # ── 年度贈經計畫 ─────────────────────────────────────
    def find_plan(self, period: str) -> Path | None:
        name = plan_filename(period)
        for d in self.plan_dirs:
            p = d / name
            if p.exists():
                return p
        return None

    def save_plan(self, period: str, src: Path) -> Path:
        dst = self.plan_dirs[0] / plan_filename(period)
        dst.parent.mkdir(parents=True, exist_ok=True)
        if Path(src) != dst:
            shutil.copy2(src, dst)
        return dst


def make_storage(base_dir, plan_dirs=None, fixed_template_dir=None) -> Storage:
    return LocalStorage(base_dir, plan_dirs, fixed_template_dir)
