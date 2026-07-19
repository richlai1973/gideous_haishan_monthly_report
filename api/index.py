"""Vercel serverless 進入點。

⚠️ 這個檔案必須放在 `api/` 底下。本專案的 Vercel 設定走傳統模式，
   function 只認 `api/` 目錄；把進入點放根目錄會得到

       The pattern "app.py" defined in `functions`
       doesn't match any Serverless Functions inside the `api` directory.

   而根目錄同時要有 `vercel.json` 的 rewrites 把所有路徑導到這裡，
   否則只有 /api/* 會進到 function，首頁會 404。

實作與本機版共用 ../app.py，雲端專用設定由 app.py 偵測 VERCEL 後自動套用。
"""

from __future__ import annotations

import sys
from pathlib import Path

# 讓 app.py 與 engine/ 可被匯入（function 的工作目錄不保證是 repo 根）
ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from app import app  # noqa: E402

__all__ = ["app"]
