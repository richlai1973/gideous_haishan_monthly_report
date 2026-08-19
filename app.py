"""基甸會海山支會 月例會報告產出系統 — FastAPI 後端 + 單頁前端。

啟動：
    pip install -r requirements.txt
    python app.py            # → http://127.0.0.1:8848
"""

from __future__ import annotations

import io
import os
import shutil
import zipfile
from datetime import date
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, Request, Response, UploadFile
from fastapi.responses import (FileResponse, HTMLResponse, JSONResponse,
                               StreamingResponse)
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from engine import auth
from engine import generate as gen
from engine import grafana
from engine.dates import build_meta, dashboard_url, default_report_month, latest_fiscal_year
from engine.parse_files import SUPPORTED, parse_file
from engine.storage import make_storage, template_month_of
from models.model import AFFECTED_DOCS, merge

APP_DIR = Path(__file__).parent

# 本機 .env（已被 .gitignore 排除）；雲端則直接用平台的環境變數
_env_file = APP_DIR / ".env"
if _env_file.exists():
    for _line in _env_file.read_text(encoding="utf-8").splitlines():
        _line = _line.strip()
        if _line and not _line.startswith("#") and "=" in _line:
            _k, _, _v = _line.partition("=")
            os.environ.setdefault(_k.strip(), _v.strip().strip("'\""))
from engine import auth as _auth_probe  # noqa: E402  （需在設定路徑前判斷環境）

# 雲端只有 /tmp 可寫，而且**單次工作階段**用完就沒了。
# 沒有任何外部儲存（Drive 已移除）：範本改用 repo 內建的固定範本，
# 產出結果一律由使用者按「下載 ZIP」帶走。
CLOUD = _auth_probe.is_cloud()
if CLOUD:
    os.environ.setdefault("GIDEONS_BASE_DIR", "/tmp/gideons")

BASE_DIR = Path(os.environ.get("GIDEONS_BASE_DIR",
                               Path.home() / "Documents" / "海山支會"))
# repo 內建：固定範本（11 份 docx）與年度贈經計畫，跟程式一起部署
FIXED_TEMPLATE_DIR = Path(os.environ.get("GIDEONS_TEMPLATE_DIR", APP_DIR / "templates"))
REPO_PLAN_DIR = APP_DIR / "贈經計畫"
# 本機優先寫專案旁的資料夾；雲端只有 /tmp 可寫，repo 內的那份唯讀備用
PLAN_DIRS = ([BASE_DIR / "贈經計畫", REPO_PLAN_DIR] if CLOUD
             else [Path(os.environ.get("GIDEONS_PLAN_DIR", APP_DIR.parent / "贈經計畫")),
                   REPO_PLAN_DIR])

app = FastAPI(title="基甸會海山支會 月例會報告產出系統", version="2.0")
store = make_storage(BASE_DIR, PLAN_DIRS, FIXED_TEMPLATE_DIR)

# ── 錯誤處理 ─────────────────────────────────────────────


def _warnings() -> list[str]:
    return store.take_warnings()


def _detail(msg: str) -> str:
    w = _warnings()
    return msg + ("　（" + "；".join(w) + "）" if w else "")


@app.exception_handler(Exception)
async def _handle_unexpected(request: Request, exc: Exception):
    """沒預料到的錯誤也要講人話——至少講出例外型別與訊息。

    這個站台有密碼保護，訊息只給承辦人看，可讀性優先於隱藏細節。
    """
    return JSONResponse(
        {"detail": f"伺服器錯誤：{type(exc).__name__}: {exc}"[:600]},
        status_code=500)


# ── 密碼保護 ─────────────────────────────────────────────
OPEN_PATHS = {"/login", "/api/login", "/api/auth-status", "/favicon.ico"}


@app.middleware("http")
async def require_password(request: Request, call_next):
    """所有頁面與 API 都要通過密碼；未設定密碼且在本機則放行。"""
    path = request.url.path
    if path in OPEN_PATHS or path.startswith("/static/"):
        return await call_next(request)

    if not auth.auth_required():
        return await call_next(request)

    if not auth.configured_password():
        # 雲端卻沒設密碼 → 明確擋下，不要無防護上線
        return JSONResponse(
            {"detail": "伺服器尚未設定 APP_PASSWORD 環境變數，為保護資料已停止服務。"},
            status_code=503)

    if auth.valid_token(request.cookies.get(auth.COOKIE)):
        return await call_next(request)

    if path.startswith("/api/"):
        return JSONResponse({"detail": "需要登入"}, status_code=401)
    return HTMLResponse(LOGIN_HTML, status_code=401)


@app.get("/api/auth-status")
def auth_status(request: Request):
    return {
        "required": auth.auth_required(),
        "configured": auth.configured_password() is not None,
        "logged_in": auth.valid_token(request.cookies.get(auth.COOKIE)),
        "cloud": auth.is_cloud(),
        "warning": auth.password_strength_warning(),
    }


class LoginReq(BaseModel):
    password: str


@app.post("/api/login")
def api_login(req: LoginReq, response: Response):
    if not auth.check_password(req.password):
        raise HTTPException(401, "密碼錯誤")
    response.set_cookie(
        auth.COOKIE, auth.issue_token(), max_age=auth.MAX_AGE,
        httponly=True, samesite="lax", secure=auth.is_cloud(), path="/")
    return {"ok": True}


@app.post("/api/logout")
def api_logout(response: Response):
    response.delete_cookie(auth.COOKIE, path="/")
    return {"ok": True}


@app.get("/login", response_class=HTMLResponse)
def login_page():
    return LOGIN_HTML


LOGIN_HTML = """<!DOCTYPE html><html lang="zh-Hant"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>基甸會海山支會 月例會報告</title><style>
body{margin:0;height:100vh;display:flex;align-items:center;justify-content:center;
 background:#f6f7f9;font:15px/1.6 -apple-system,"PingFang TC","Noto Sans TC",sans-serif}
.box{background:#fff;border:1px solid #e3e6ea;border-radius:10px;padding:28px 32px;
 width:330px;text-align:center}
h1{font-size:16px;margin:0 0 4px;color:#1f4e79}
p{font-size:12px;color:#6b7684;margin:0 0 18px}
input{width:100%;padding:9px 11px;border:1px solid #e3e6ea;border-radius:6px;
 font:inherit;box-sizing:border-box}
button{width:100%;margin-top:10px;padding:9px;border:0;border-radius:6px;
 background:#1f4e79;color:#fff;font:inherit;cursor:pointer}
button:hover{background:#2e6ca4}
.err{color:#b3261e;font-size:12px;margin-top:10px;min-height:16px}
</style></head><body><div class="box">
<h1>國際基甸會 海山支會</h1><p>月例會報告產出系統</p>
<input id="pw" type="password" placeholder="請輸入密碼" autofocus>
<button id="go">登入</button><div class="err" id="err"></div>
<script>
const go=async()=>{document.getElementById('err').textContent='';
 const r=await fetch('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},
  body:JSON.stringify({password:document.getElementById('pw').value})});
 if(r.ok){location.href='/';}else{document.getElementById('err').textContent='密碼錯誤';}};
document.getElementById('go').onclick=go;
document.getElementById('pw').addEventListener('keydown',e=>{if(e.key==='Enter')go();});
</script></div></body></html>"""


# ── 共用 ─────────────────────────────────────────────────
def _meta(year: int, month: int, meeting_date: str | None = None,
          template_ym: tuple[int, int] | None = None):
    return build_meta(year, month, meeting_date or None, template_ym)


def resolve_template(year: int, month: int, meeting_date: str | None = None):
    """決定這次要用哪一組範本，並讓 meta 的「範本月份」跟著它。

    ① 上月工作資料夾（本機常態，內容是上個月真正送出的版本）
    ② repo 內建固定範本（雲端唯一來源；停在某一個月，例如 115年7月）

    ②的月份不會等於「上個月」，而 `update_dates()` 全部替換都以 meta.prev_*
    為來源字串——所以這裡要把 meta 重建成以**實際範本月份**為準，否則會產出
    一份日期完全沒換到的報告，而且不會報錯。

    回傳 (meta, 範本路徑或 None, 來源說明)。
    """
    meta = _meta(year, month, meeting_date)
    tpl = store.month_template_dir(meta)
    if tpl is not None:
        return meta, tpl, f"上月資料夾（{meta.prev_year}年{meta.prev_month}月）"

    fixed = store.fixed_template_dir()
    if fixed is None:
        return meta, None, "找不到範本"

    ym = template_month_of(fixed)
    if ym and (ym[0], ym[1]) != (meta.prev_year, meta.prev_month):
        meta = _meta(year, month, meeting_date, ym)
    label = f"內建固定範本（{meta.prev_roc_year}年{meta.prev_month}月）"
    return meta, fixed, label


def _work_dir(year: int, month: int) -> Path:
    """工作區。本機＝永久資料夾；雲端＝/tmp，單次工作階段用完即消失。"""
    return store.work_dir(_meta(year, month))


def _docx_paths(work_dir: Path, meta) -> list[str]:
    prefix = f"月例會議程{meta.roc_year}年{meta.report_month}月"
    return sorted((str(p) for p in work_dir.glob(f"{prefix}-*.docx")),
                  key=lambda p: gen.file_number(os.path.basename(p)))


# ── API：狀態 ────────────────────────────────────────────
@app.get("/api/status")
def status(year: int | None = None, month: int | None = None):
    y, m = (year, month) if year and month else default_report_month()
    meta = _meta(y, m)
    wd = _work_dir(y, m)
    files = []
    if wd.exists():
        for p in _docx_paths(wd, meta):
            files.append({"name": os.path.basename(p),
                          "num": gen.file_number(os.path.basename(p)),
                          "size": os.path.getsize(p),
                          "mtime": os.path.getmtime(p)})
    plan = store.find_plan(meta.period)
    _meta_tpl, tpl_dir, tpl_label = resolve_template(y, m)
    return {
        "meta": meta.to_dict(),
        "cloud": CLOUD,
        "storage": "cloud-temp" if CLOUD else "local",
        "base_dir": "雲端暫存區（單次工作階段）" if CLOUD else str(BASE_DIR),
        "base_dir_exists": True if CLOUD else BASE_DIR.exists(),
        "work_dir": str(wd),
        "work_dir_exists": wd.exists(),
        "template_dir": str(tpl_dir) if tpl_dir else None,
        "template_exists": tpl_dir is not None,
        "template_source": tpl_label,
        "template_month": (f"{_meta_tpl.prev_roc_year}年{_meta_tpl.prev_month}月"
                           if tpl_dir else None),
        "files": files,
        "ready": len(files),
        "latest_fiscal_year": latest_fiscal_year(),
        "dashboard_url": dashboard_url(meta.fiscal_year),
        "distribution_plan": {"path": str(plan) if plan else None,
                              "exists": plan is not None,
                              "in_model": bool(store.load_model(meta)
                                               .get("distribution_plan", {}).get("schools")),
                              "period": meta.period,
                              "source": "、".join(str(d) for d in PLAN_DIRS)},
        "model": store.load_model(meta) or None,
        "supported_ext": sorted(SUPPORTED),
        "warnings": _warnings(),
    }


def _ensure_plan_in_model(meta, model) -> str | None:
    """年度贈經計畫是**整個財年的固定參考**，新月份要自動帶進資料模型。

    先前只有「上傳當下」那個月的 model 有 distribution_plan，隔月產生報告時
    -6 學校贈經統計表就是空的——介面卻寫著「一年上傳一次即可」。
    2026-2027 這份計畫涵蓋 2026/6/1–2027/5/31，中間每個月都該吃到同一份。
    """
    cur = model.get("distribution_plan") or {}
    if cur.get("period") == meta.period and cur.get("schools"):
        return None
    p = store.find_plan(meta.period)
    if not p:
        return None
    from engine.parse_plan import parse as parse_plan
    try:
        plan = parse_plan(str(p), meta.period)
    except Exception as exc:
        return f"年度贈經計畫讀取失敗：{exc}"
    if not plan.get("schools"):
        return "；".join(plan.get("warnings", [])) or \
            f"年度贈經計畫（{meta.period}）沒有可用排程"
    merge(model, {"distribution_plan": plan}, f"贈經計畫:{meta.period}（固定參考自動帶入）")
    store.save_model(meta, model)
    from engine.parse_plan import schools_of_month
    n = len(schools_of_month(plan, meta.report_year, meta.report_month))
    return (f"已自動帶入 {meta.period} 年度贈經計畫："
            f"全年 {len(plan['schools'])} 場，本月 {n} 場")


# ── API：① 初始化當月 ────────────────────────────────────
class InitReq(BaseModel):
    year: int
    month: int
    meeting_date: str | None = None


@app.post("/api/init")
def api_init(req: InitReq):
    meta, tpl, tpl_label = resolve_template(req.year, req.month, req.meeting_date)
    if tpl is None:
        raise HTTPException(400, _detail(
            f"找不到範本：上月資料夾與內建固定範本（{FIXED_TEMPLATE_DIR}）都沒有 docx"))
    res = gen.init_month(str(_work_dir(req.year, req.month).parent), meta,
                         template_dir=str(tpl))
    if not res["ok"]:
        raise HTTPException(400, res["error"])
    model = store.load_model(meta)
    model["meta"] = meta.to_dict()
    store.save_model(meta, model)
    note = _ensure_plan_in_model(meta, model)
    res["meta"] = meta.to_dict()
    res["template_source"] = tpl_label
    res["plan_note"] = note
    res["warnings"] = _warnings()
    return res


# ── API：解析上傳檔（先看結果，確認才寫入）──────────────
@app.post("/api/parse")
async def api_parse(module: str = Form(...), year: int = Form(...),
                    month: int = Form(...), file: UploadFile = File(...)):
    meta = _meta(year, month)
    wd = _work_dir(year, month)
    inputs = wd / "_inputs"
    inputs.mkdir(parents=True, exist_ok=True)

    dest = inputs / file.filename
    with open(dest, "wb") as f:
        shutil.copyfileobj(file.file, f)

    ext = dest.suffix.lower()
    if ext not in SUPPORTED:
        raise HTTPException(400, f"不支援的格式：{ext}")

    result = parse_file(str(dest), module, meta.report_year, meta.report_month)
    result["module"] = module
    result["saved_to"] = str(dest)
    result["affected_docs"] = AFFECTED_DOCS.get(module, [])

    # 事工成果表另存標準檔名，供產出時引用
    if result["kind"] == "ministry_excel":
        std = wd / f"事工成果表_{year}_{month:02d}.xlsx"
        shutil.copy2(dest, std)
        result["standard_excel"] = str(std)
    return result


# ── API：自動下載 Grafana 事工成果表（只下載＋解析，不寫入）──
class FetchReq(BaseModel):
    year: int
    month: int
    team: str = grafana.DEFAULT_TEAM
    fiscal_year: int | None = None      # 可覆寫（補做舊年度）
    # ⚠️ 預設 False：回退年度會取到上一財年的數字，寫進本月報告就是錯的
    allow_fallback: bool = False


@app.post("/api/grafana/fetch")
def api_grafana_fetch(req: FetchReq):
    """從 Grafana 擷取本財年資料（免登入），回傳供審閱。

    預設走 **API 查詢**：任何年度都可用，包含目標尚未設定的新財年。
    此端點**不會**寫入資料模型或改動任何 docx；
    承辦人在介面確認數字無誤後，再按「確認並套用」走 /api/commit。
    """
    meta = _meta(req.year, req.month)
    fy = req.fiscal_year or meta.fiscal_year

    try:
        res = grafana.fetch_all(fy, req.team)
    except grafana.GrafanaError as exc:
        raise HTTPException(502, str(exc))

    stats = res["ministry_stats"]
    checks = [
        {"item": "連線 Grafana API", "ok": True, "detail": "免登入查詢成功"},
        {"item": "財年區間", "ok": True, "detail": res["period"]},
        {"item": "成果項目", "ok": bool(stats),
         "detail": f"{len(stats)} 項（{'、'.join(list(stats)[:5])}）" if stats else "查無資料"},
        {"item": "年度目標已設定",
         "ok": bool(stats) and any(v["goal"] is not None for v in stats.values()),
         "detail": "已設定" if (stats and any(v["goal"] is not None for v in stats.values()))
                   else "尚未設定 → 目標／差額／達成率顯示「-」"},
        {"item": "教會見證", "ok": True, "detail": f"{len(res['church_testimony'])} 場"},
        {"item": "會費繳納", "ok": True, "detail": f"{len(res['membership_fees'])} 筆"},
    ]

    # 已設定目標的年度 → 供介面說明 Excel 報表可用範圍
    try:
        goal_years = grafana.available_goal_years()
    except grafana.GrafanaError:
        goal_years = []

    return {
        "source": "api",
        "fiscal_year": fy,
        "requested_fiscal_year": meta.fiscal_year,
        "period": res["period"],
        "team": res["team"],
        "checks": checks,
        "all_ok": all(c["ok"] for c in checks),
        "warnings": res["warnings"],
        "counts": res["counts"],
        "ministry_stats": stats,
        "church_testimony": res["church_testimony"],
        "membership_fees": res["membership_fees"][:100],
        "bible_giving": res["bible_giving"][:100],
        "goal_years": goal_years,
        "excel_available": fy in goal_years,
        "dashboard_url": grafana.dashboard_url(fy),
        "patch": {                        # 確認後才送去 /api/commit
            "ministry_stats_api": stats,
            "church_testimony_api": res["church_testimony"],
            "membership_fees_api": res["membership_fees"],
            "offerings_api": res["offerings_detail"],
            "bible_giving_api": res["bible_giving"],
        },
    }


@app.post("/api/grafana/excel")
def api_grafana_excel(req: FetchReq):
    """產出事工成果表 Excel。

    先試官方報表端點；該年度目標未建檔而回 500 時，
    改由 API 資料**重建**同版面的 Excel，兩者都存成同一個檔名。
    """
    meta = _meta(req.year, req.month)
    fy = req.fiscal_year or meta.fiscal_year
    wd = _work_dir(req.year, req.month)
    wd.mkdir(parents=True, exist_ok=True)
    dest = wd / f"事工成果表_{req.year}_{req.month:02d}.xlsx"

    try:
        res = grafana.fetch_excel(str(dest), fy, req.team, req.allow_fallback)
        if res["ok"]:
            return {**res, "rebuilt": False}
        official_error = res["error"]

        # 官方報表產不出來 → 用 API 重建
        from engine import build_excel
        data = grafana.fetch_all(fy, req.team)
        built = build_excel.build(data, str(dest), req.team)
        return {"ok": True, "source": "rebuilt", "rebuilt": True,
                "official_error": official_error, "fiscal_year": fy,
                "file": dest.name, "path": str(dest),
                "size": dest.stat().st_size, **built,
                "warnings": data["warnings"]}
    except grafana.GrafanaError as exc:
        raise HTTPException(502, str(exc))


# ── API：確認後寫入資料模型 ──────────────────────────────
class CommitReq(BaseModel):
    year: int
    month: int
    module: str
    patch: dict
    source: str = ""


@app.post("/api/commit")
def api_commit(req: CommitReq):
    wd = _work_dir(req.year, req.month)
    if not wd.exists():
        raise HTTPException(400, "尚未初始化當月工作區，請先按「產生報告」")
    meta = _meta(req.year, req.month)
    model = store.load_model(meta)
    merge(model, req.patch, req.source or req.module)
    store.save_model(meta, model)
    return {"ok": True, "model": model,
            "warnings": _warnings(),
            "affected_docs": AFFECTED_DOCS.get(req.module, [])}


# ── API：產出 10 份 docx ─────────────────────────────────
class GenReq(BaseModel):
    year: int
    month: int
    meeting_date: str | None = None
    only: list[int] | None = None      # 增量更新：只重繪指定文件編號


@app.post("/api/generate")
def api_generate(req: GenReq):
    meta, tpl, tpl_label = resolve_template(req.year, req.month, req.meeting_date)
    wd = _work_dir(req.year, req.month)
    if not any(wd.glob("*.docx")):
        if tpl is None:
            raise HTTPException(400, _detail(
                f"找不到範本：上月資料夾與內建固定範本（{FIXED_TEMPLATE_DIR}）都沒有 docx"))
        init = gen.init_month(str(wd.parent), meta, template_dir=str(tpl))
        if not init["ok"]:
            raise HTTPException(400, init["error"])

    model = store.load_model(meta)
    plan_note = _ensure_plan_in_model(meta, model)
    excel = wd / f"事工成果表_{req.year}_{req.month:02d}.xlsx"
    res = gen.generate_all(str(wd), meta, str(excel) if excel.exists() else None, model)
    if not res["ok"]:
        raise HTTPException(400, res["error"])
    if req.only:
        res["results"] = [r for r in res["results"] if r["num"] in req.only]
    res["meta"] = meta.to_dict()
    res["plan_note"] = plan_note
    res["template_source"] = tpl_label
    res["warnings"] = _warnings()
    if CLOUD:
        # /tmp 不保證留到下一個請求，也沒有任何外部儲存可放——只能提醒帶走
        res["warnings"] = res["warnings"] + ["雲端產出只留在本次工作階段，請立即按「下載 ZIP」帶走"]
    return res


@app.post("/api/plan/upload")
async def api_plan_upload(year: int = Form(...), month: int = Form(...),
                          period: str = Form(""), file: UploadFile = File(...)):
    """上傳年度贈經計畫（整個財年固定的學校配送排程）。

    本機存到專案旁的「贈經計畫」資料夾（永久保留）。
    雲端只存得進 /tmp，**單次工作階段有效**——repo 內建的那份才是雲端常態來源，
    所以換年度時記得把新的計畫檔一起 commit 進 `贈經計畫/`。
    """
    meta = _meta(year, month)
    period = period.strip() or meta.period
    if not file.filename.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(400, "年度贈經計畫需為 Excel 檔")

    tmp = _work_dir(year, month) / "_inputs" / file.filename
    tmp.parent.mkdir(parents=True, exist_ok=True)
    with open(tmp, "wb") as f:
        shutil.copyfileobj(file.file, f)

    dest = store.save_plan(period, tmp)

    # 解析並寫入資料模型——只存檔不解析的話，-6 不會有任何變化
    from engine.parse_plan import parse as parse_plan
    try:
        plan = parse_plan(str(dest), period)
    except Exception as exc:
        raise HTTPException(422, f"檔案已存入但解析失敗：{exc}")
    if not plan["schools"]:
        raise HTTPException(
            422, "；".join(plan["warnings"]) or f"找不到 {period} 年度的排程")

    model = store.load_model(meta)
    merge(model, {"distribution_plan": plan}, f"贈經計畫:{period}")
    store.save_model(meta, model)

    from engine.parse_plan import schools_of_month
    # 預覽要看得到「日期待定」的場次，但產出時不會自動填進 -6
    this_month = schools_of_month(plan, meta.report_year, meta.report_month,
                                  include_undated=True)

    return {"ok": True, "period": period, "file": os.path.basename(str(dest)),
            "path": str(dest), "storage": "cloud-temp" if CLOUD else "local",
            "schools": plan["schools"], "total": len(plan["schools"]),
            "this_month": this_month,
            "warnings": plan["warnings"] + _warnings(),
            "affected_docs": [5, 6]}


# ── API：分析檢視（事工成果表）──────────────────────────
@app.get("/api/analysis")
def api_analysis(year: int, month: int):
    wd = _work_dir(year, month)
    excel = wd / f"事工成果表_{year}_{month:02d}.xlsx"
    if not excel.exists():
        raise HTTPException(404, "尚未上傳事工成果表 Excel")
    from engine.parse_excel import MinistryExcel
    mx = MinistryExcel(str(excel))
    return {"file": excel.name, "analysis": mx.analysis(),
            "offerings": mx.offerings(), "churches": mx.churches()}


# ── API：下載 ────────────────────────────────────────────
@app.get("/api/download")
def api_download(year: int, month: int, name: str | None = None):
    meta = _meta(year, month)
    wd = _work_dir(year, month)
    if name:
        p = wd / name
        if not p.exists():
            raise HTTPException(404, "檔案不存在")
        return FileResponse(str(p), filename=name)

    paths = _docx_paths(wd, meta)
    if not paths:
        raise HTTPException(404, "尚無產出檔案")
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for p in paths:
            z.write(p, os.path.basename(p))
    buf.seek(0)
    return StreamingResponse(
        buf, media_type="application/zip",
        headers={"Content-Disposition":
                 f'attachment; filename="gideons-{year}-{month:02d}.zip"'})


# ── 前端 ─────────────────────────────────────────────────
@app.get("/", response_class=HTMLResponse)
def index():
    return (APP_DIR / "static" / "index.html").read_text(encoding="utf-8")


if (APP_DIR / "static").exists():
    app.mount("/static", StaticFiles(directory=str(APP_DIR / "static")), name="static")


if __name__ == "__main__":
    import uvicorn
    print(f"資料夾：{BASE_DIR}")
    print("開啟 → http://127.0.0.1:8848")
    uvicorn.run(app, host="127.0.0.1", port=8848)
