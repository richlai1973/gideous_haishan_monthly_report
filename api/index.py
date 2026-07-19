"""Vercel serverless 入口。

Vercel 會把 `api/` 底下的 Python 檔當成 function，並匯出名為 `app` 的
ASGI 應用程式。這裡只負責設定雲端專用的預設值，實作與本機版共用 app.py。

雲端與本機的差異：
  儲存層   Drive（無持久磁碟）  vs  ~/Documents/海山支會
  授權     環境變數 refresh token vs credentials/token.json
  密碼     必填（未設會回 503）  vs  可不設
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

# 讓 app.py 與 engine/ 可被匯入
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

# 雲端預設：Drive 儲存層、/tmp 當工作區
os.environ.setdefault("STORAGE", "drive")
os.environ.setdefault("GIDEONS_BASE_DIR", "/tmp/gideons")
os.environ.setdefault("GIDEONS_CRED_DIR", "/tmp/gideons-cred")

from app import app  # noqa: E402

__all__ = ["app"]
