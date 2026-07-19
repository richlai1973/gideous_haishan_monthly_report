#!/usr/bin/env python3
"""本機執行入口。

    python3 run.py                # → http://127.0.0.1:8848
    python3 run.py --port 9000
    python3 run.py --open         # 順便開瀏覽器

與雲端版共用同一份 app.py，差別只在儲存層：
本機直接讀寫 ~/Documents/海山支會，雲端走 Google Drive（見 api/index.py）。
"""

from __future__ import annotations

import argparse
import os
import sys
import webbrowser
from pathlib import Path

APP_DIR = Path(__file__).parent
sys.path.insert(0, str(APP_DIR))


def main() -> int:
    p = argparse.ArgumentParser(description="基甸會海山支會 月例會報告產出系統")
    p.add_argument("--host", default="127.0.0.1")
    p.add_argument("--port", type=int, default=8848)
    p.add_argument("--open", action="store_true", help="啟動後開啟瀏覽器")
    p.add_argument("--reload", action="store_true", help="改程式碼自動重載（開發用）")
    args = p.parse_args()

    # 本機一律用本機檔案系統
    os.environ.setdefault("STORAGE", "local")

    try:
        import uvicorn  # noqa: F401
        from app import BASE_DIR, app  # noqa: F401
    except ImportError as exc:
        print(f"❌ 缺少套件：{exc}")
        print("   請先執行：pip install -r requirements.txt")
        return 1

    from engine import auth

    url = f"http://{args.host}:{args.port}"
    print("─" * 52)
    print("  國際基甸會 海山支會 · 月例會報告產出系統")
    print("─" * 52)
    print(f"  資料夾  ：{BASE_DIR}")
    if not BASE_DIR.exists():
        print("            ⚠️ 找不到此資料夾，可用 GIDEONS_BASE_DIR 指定")
    print(f"  密碼    ：{'已啟用' if auth.configured_password() else '未設定（本機不驗證）'}")
    if w := auth.password_strength_warning():
        print(f"            ⚠️ {w}")
    print(f"  網址    ：{url}")
    print("─" * 52)

    if args.open:
        webbrowser.open(url)

    import uvicorn
    uvicorn.run("app:app" if args.reload else "app:app",
                host=args.host, port=args.port, reload=args.reload,
                log_level="info")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
