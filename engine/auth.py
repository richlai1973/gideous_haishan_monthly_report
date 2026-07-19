"""單一共用密碼保護。

密碼**只從環境變數讀取**，絕不寫進程式碼——這個 repo 可能是 public，
而且雲端部署時密碼要能隨時更換而不需改程式。

    本機：  export APP_PASSWORD='...'   或寫進 .env（已被 .gitignore 排除）
    Vercel：Project → Settings → Environment Variables → APP_PASSWORD

未設定 APP_PASSWORD 時**不啟用驗證**，方便本機開發；
但只要偵測到雲端環境（VERCEL 等）就一律要求密碼，避免誤把無防護的版本推上線。
"""

from __future__ import annotations

import hashlib
import hmac
import os
import secrets
import time

COOKIE = "gideons_session"
MAX_AGE = 12 * 60 * 60          # 12 小時
CLOUD_ENVS = ("VERCEL", "RAILWAY_ENVIRONMENT", "RENDER", "FLY_APP_NAME")


def configured_password() -> str | None:
    pw = os.environ.get("APP_PASSWORD", "").strip()
    return pw or None


def is_cloud() -> bool:
    return any(os.environ.get(k) for k in CLOUD_ENVS)


def auth_required() -> bool:
    """有設密碼就啟用；雲端環境即使沒設也視為需要（會擋下所有請求並提示）。"""
    return configured_password() is not None or is_cloud()


def _secret() -> bytes:
    """簽章金鑰。優先用獨立的 SESSION_SECRET，否則由密碼衍生。"""
    s = os.environ.get("SESSION_SECRET", "").strip()
    if s:
        return s.encode()
    pw = configured_password() or ""
    return hashlib.sha256(("gideons-session::" + pw).encode()).digest()


def issue_token() -> str:
    """簽發 `到期時間.簽章`。"""
    exp = str(int(time.time()) + MAX_AGE)
    sig = hmac.new(_secret(), exp.encode(), hashlib.sha256).hexdigest()
    return f"{exp}.{sig}"


def valid_token(token: str | None) -> bool:
    if not token or "." not in token:
        return False
    exp, _, sig = token.partition(".")
    expect = hmac.new(_secret(), exp.encode(), hashlib.sha256).hexdigest()
    if not hmac.compare_digest(sig, expect):   # 定時比較，避免時序攻擊
        return False
    try:
        return int(exp) > time.time()
    except ValueError:
        return False


def check_password(supplied: str) -> bool:
    pw = configured_password()
    if not pw:
        return False
    # 定時比較：避免以回應時間逐字元猜測
    return hmac.compare_digest(supplied.encode(), pw.encode())


def password_strength_warning() -> str | None:
    """密碼太弱時回報，顯示在狀態列提醒承辦人。"""
    pw = configured_password()
    if not pw:
        return None
    weak = {"haishan", "gideons", "海山", "password", "12345678", "gideon"}
    if pw.lower() in weak:
        return ("目前密碼是可猜到的常見字（支會名／組織名）。"
                "報告含具名健康資訊，建議改用隨機密碼，"
                "例如：" + secrets.token_urlsafe(9))
    if len(pw) < 8:
        return "密碼短於 8 碼，建議加長。"
    return None
