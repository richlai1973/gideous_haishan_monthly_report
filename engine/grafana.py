"""Grafana 事工成果資料擷取（免登入，已實測驗證）。

## 兩條取得資料的路徑

1. **API 查詢（主要）** — `POST /api/ds/query` 直接查 PostgreSQL 資料來源。
   任何年度都可用，**包含目標尚未設定的新財年**。
2. **Excel 報表（次要）** — `GET /report/campstat/{支會}?year={年度}`，
   即 dashboard 上「下載最新成果表」按鈕。**只有已設定年度目標的年度可用**。

## ⚠️ 為什麼預設走 API

實測（2026-07）：
- `gideons_goal1` 的年度只到 **2026**，2027 財年目標尚未建檔。
- 因此 `/report/campstat/?year=2027` 回 **HTTP 500**，2024–2026 則正常。
- 但 2027 的**成果資料真實存在**（弟兄會費 3、姊妹會費 2、教會聖奉 6,400…），
  且與 115年6月-2事工成果統計表 docx 的「成果」列完全吻合。

所以**絕不可**因為 2027 下載失敗就退回抓 2026 —— 那會把上一個財年的
184,772 教會聖奉寫進本財年報告。年度回退預設關閉，只在明確指定時啟用。
"""

from __future__ import annotations

import json
import os
import urllib.error
import urllib.parse
import urllib.request

BASE = "https://gideons-dashboard.pointing.tw"
DATASOURCE = {"type": "grafana-postgresql-datasource", "uid": "de9c9tp9gzg1sb"}
DEFAULT_TEAM = "海山"
TIMEOUT = 60
XLSX_MAGIC = b"PK"

# -2 事工成果統計表：docx 欄索引 → 資料庫 category
MINISTRY_CATEGORIES = {
    1: "增加弟兄", 2: "增加姊妹",
    3: "弟兄會費", 4: "姊妹會費",
    5: "贈送聖經", 6: "教會見證", 7: "教會聖奉",
    8: "弟兄聖奉", 9: "姊妹聖奉", 10: "姊妹贈經",
}


class GrafanaError(RuntimeError):
    pass


# ── 基礎 HTTP ────────────────────────────────────────────
def _request(url: str, data: bytes | None = None,
             content_type: str = "") -> tuple[int, bytes]:
    headers = {"User-Agent": "gideons-report-app/1.0"}
    if content_type:
        headers["Content-Type"] = content_type
    req = urllib.request.Request(url, data=data, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=TIMEOUT) as resp:
            return resp.status, resp.read()
    except urllib.error.HTTPError as exc:
        return exc.code, exc.read()[:500]
    except urllib.error.URLError as exc:
        raise GrafanaError(f"無法連線 Grafana：{exc.reason}") from exc


def fiscal_range(fiscal_year: int) -> tuple[str, str]:
    """財年 6/1 ~ 5/31。2027 → ('2026-06-01', '2027-05-31')"""
    return f"{fiscal_year - 1}-06-01", f"{fiscal_year}-05-31"


def query(sql: str, fiscal_year: int) -> list[dict]:
    """執行 SQL，回傳 [{欄位: 值}]。"""
    start, end = fiscal_range(fiscal_year)
    payload = {
        "queries": [{
            "refId": "A", "datasource": DATASOURCE, "rawSql": sql,
            "format": "table", "rawQuery": True,
            "intervalMs": 60000, "maxDataPoints": 5000,
        }],
        # $__timeFilter 等巨集用得到；本模組 SQL 一律寫死日期，這裡僅為相容
        "from": str(int(_epoch_ms(start))), "to": str(int(_epoch_ms(end))),
    }
    status, body = _request(f"{BASE}/api/ds/query",
                            json.dumps(payload).encode(), "application/json")
    if status != 200:
        raise GrafanaError(f"查詢失敗 HTTP {status}：{body[:200]!r}")

    frames = (json.loads(body).get("results", {}).get("A", {}).get("frames") or [])
    if not frames:
        return []
    frame = frames[0]
    names = [f["name"] for f in frame["schema"]["fields"]]
    cols = frame["data"]["values"]
    if not cols or not names:
        return []
    return [dict(zip(names, row)) for row in zip(*cols)]


def _epoch_ms(date_str: str) -> float:
    from datetime import datetime, timezone
    y, m, d = (int(x) for x in date_str.split("-"))
    return datetime(y, m, d, tzinfo=timezone.utc).timestamp() * 1000


def _sqlstr(s: str) -> str:
    return s.replace("'", "''")


# ── 各項資料 ─────────────────────────────────────────────
def ministry_stats(fiscal_year: int, team: str = DEFAULT_TEAM) -> dict:
    """-2 事工成果統計表：各 category 的目標與成果。

    新財年目標未設定時 goal 為 None，介面與 docx 皆呈現「-」。
    """
    rows = query(
        "SELECT category, goal, value FROM v_gideons_goal_result_joined "
        f"WHERE team = '{_sqlstr(team)}' AND year = {fiscal_year} "
        "ORDER BY display_order", fiscal_year)
    out = {}
    for r in rows:
        goal, val = r.get("goal"), r.get("value")
        diff = (val - goal) if (goal is not None and val is not None) else None
        rate = (val / goal) if (goal not in (None, 0) and val is not None) else None
        out[r["category"]] = {"goal": goal, "value": val, "diff": diff, "rate": rate}
    return out


def church_testimony(fiscal_year: int, team: str = DEFAULT_TEAM) -> list[dict]:
    """-9 年度教會見證：見證場次 **與** 教會奉獻的聯集。

    ⚠️ 兩者不一定同時存在，必須 FULL OUTER JOIN：
      - 有場次無奉獻：愛加倍浸信會（2026/07/05 黃哲斌）
      - 有奉獻無場次：基督國度溪水旁教會 5,900、樹林長老教會（北中）500
    只查場次會漏掉全部實際奉獻金額（實測 FY2027 即是如此）。
    """
    start, end = fiscal_range(fiscal_year)
    t = _sqlstr(team)
    rows = query(
        "WITH ev AS (SELECT track_id, TO_CHAR(datetime,'YYYY/MM/DD') AS d, "
        "              partner, provider, datetime "
        f"           FROM gideons_witness_event_data WHERE team = '{t}' "
        f"             AND datetime BETWEEN '{start}' AND '{end}'), "
        "don AS (SELECT partner AS church, SUM(amount) AS amt "
        f"        FROM gideons_witness_donation WHERE team = '{t}' "
        f"          AND date BETWEEN '{start}' AND '{end}' GROUP BY partner) "
        "SELECT COALESCE(ev.provider, don.church) AS church, "
        "       ev.d AS date, ev.partner AS speaker, don.amt AS amount, "
        "       ev.datetime AS dt "
        "FROM ev FULL OUTER JOIN don ON ev.provider = don.church "
        "ORDER BY ev.datetime NULLS LAST, church", fiscal_year)
    return [{"date": r.get("date"), "speaker": r.get("speaker"),
             "church": r.get("church"), "amount": r.get("amount") or 0,
             "has_event": bool(r.get("date")),
             "donation_only": not r.get("date") and bool(r.get("amount"))}
            for r in rows]


def member_roster(fiscal_year: int, team: str = DEFAULT_TEAM) -> list[dict]:
    """會員名冊與會費狀態（Excel D–G 欄）。

    「弟兄會費／姊妹會費」欄位語意同 dashboard：一般會員顯示最近繳費日期，
    終身／資深顯示類別，安息／退會顯示 None。
    """
    start, end = fiscal_range(fiscal_year)
    t = _sqlstr(team)
    fee = ("(SELECT to_char(date,'MM/DD') FROM v_membership_record "
           "  WHERE ref IN ({ref}) AND date BETWEEN '{s}' AND '{e}' "
           "  ORDER BY date DESC LIMIT 1)")
    b_fee = fee.format(ref='"會員編號", TRIM(LEADING \'0\' FROM "會員編號")',
                       s=start, e=end)
    s_fee = fee.format(ref='"會員編號" || \'A\', '
                           'TRIM(LEADING \'0\' FROM "會員編號") || \'A\'',
                       s=start, e=end)
    rows = query(
        "WITH LIST AS (SELECT \"會員編號\", "
        f" (CASE WHEN \"弟兄安息\" < '{start}' OR \"弟兄類別\"='退會' "
        "        THEN NULL ELSE \"弟兄\" END) AS b, "
        f" (CASE WHEN \"姊妹安息\" < '{start}' OR \"姊妹類別\"='退會' "
        "        THEN NULL ELSE \"姊妹\" END) AS s, "
        " \"弟兄類別\" AS bt, \"姊妹類別\" AS st, "
        " \"弟兄退會\" AS bq, \"姊妹退會\" AS sq, "
        " \"弟兄支會\" AS bteam, \"姊妹支會\" AS steam, "
        " (CASE WHEN (\"弟兄類別\" NOT IN ('終身','資深') AND \"弟兄安息\" IS NULL "
        f"             AND \"弟兄退會\" IS NULL) THEN {b_fee} "
        "        WHEN \"弟兄類別\" LIKE '%安息%' OR \"弟兄類別\"='退會' THEN NULL "
        "        ELSE \"弟兄類別\" END) AS bfee, "
        " (CASE WHEN (\"姊妹類別\" NOT IN ('終身','資深') AND \"姊妹安息\" IS NULL "
        f"             AND \"姊妹退會\" IS NULL) THEN {s_fee} "
        "        WHEN \"姊妹類別\" LIKE '%安息%' OR \"姊妹類別\"='退會' THEN NULL "
        "        ELSE \"姊妹類別\" END) AS sfee "
        f" FROM v_membership_gender_joined WHERE '{t}' IN (\"弟兄支會\",\"姊妹支會\") "
        f"   AND ((\"弟兄入會\"::date <= '{end}' AND (\"弟兄退會\" IS NULL "
        f"         OR \"弟兄退會\"::date >= '{start}')) "
        f"     OR (\"姊妹入會\"::date <= '{end}' AND (\"姊妹退會\" IS NULL "
        f"         OR \"姊妹退會\"::date >= '{start}')))) "
        "SELECT DISTINCT \"會員編號\" AS id, b, s, bfee, sfee, bt, st FROM LIST "
        f"WHERE ((b IS NOT NULL AND bt <> '退會' AND bteam = '{t}') "
        f"    OR (s IS NOT NULL AND st <> '退會' AND steam = '{t}')) "
        "ORDER BY id", fiscal_year)
    return [{"id": r.get("id"), "brother": r.get("b"), "sister": r.get("s"),
             "fee_brother": r.get("bfee"), "fee_sister": r.get("sfee"),
             "type_brother": r.get("bt"), "type_sister": r.get("st")}
            for r in rows]


def offerings_by_member(fiscal_year: int, team: str = DEFAULT_TEAM) -> dict:
    """依 category 分組的奉獻明細（Excel H–K 欄：聖經奉獻、巴拿巴）。"""
    detail = offerings_detail(fiscal_year, team)
    out: dict[str, dict[str, float]] = {}
    for d in detail:
        cat = d["category"]
        if cat not in ("弟兄聖奉", "姊妹聖奉", "巴拿巴"):
            continue
        out.setdefault(cat, {})
        key = str(d["id"] or d["name"])
        out[cat][key] = out[cat].get(key, 0) + (d["amount"] or 0)
    return out


def membership_fees(fiscal_year: int, team: str = DEFAULT_TEAM) -> list[dict]:
    """-4 各項奉獻：會費繳納記錄。"""
    start, end = fiscal_range(fiscal_year)
    rows = query(
        "SELECT TO_CHAR(date,'YYYY-MM-DD') AS date, member, ref, gender, amount "
        f"FROM v_membership_record WHERE team = '{_sqlstr(team)}' "
        f"  AND date BETWEEN '{start}' AND '{end}' AND amount > 0 ORDER BY date",
        fiscal_year)
    return [{"date": r.get("date"), "name": r.get("member"), "id": r.get("ref"),
             "gender": "M" if r.get("gender") == "male" else "F",
             "amount": r.get("amount") or 0} for r in rows]


def offerings_detail(fiscal_year: int, team: str = DEFAULT_TEAM) -> list[dict]:
    """-4 聖經奉獻與巴拿巴明細。"""
    start, end = fiscal_range(fiscal_year)
    rows = query(
        "SELECT category, member, ref, TO_CHAR(date,'YYYY-MM-DD') AS date, amount "
        "FROM v_gideons_goal_result_detail "
        f"WHERE team = '{_sqlstr(team)}' AND date BETWEEN '{start}' AND '{end}' "
        "ORDER BY date", fiscal_year)
    return [{"category": r.get("category"), "name": r.get("member"),
             "id": r.get("ref"), "date": r.get("date"),
             "amount": r.get("amount") or 0} for r in rows]


def bible_giving(fiscal_year: int, team: str = DEFAULT_TEAM) -> list[dict]:
    """-5/-6 贈經紀錄。"""
    start, end = fiscal_range(fiscal_year)
    rows = query(
        "SELECT TO_CHAR(date,'YYYY/MM/DD') AS date, otype, product_name, total_num "
        f"FROM gideons_odoo_bible_data WHERE team = '{_sqlstr(team)}' "
        f"  AND date BETWEEN '{start}' AND '{end}' ORDER BY date", fiscal_year)
    return [{"date": r.get("date"), "kind": r.get("otype"),
             "product": r.get("product_name"), "count": r.get("total_num") or 0}
            for r in rows]


def available_goal_years() -> list[int]:
    """已設定目標的年度（決定 Excel 報表可否產生）。"""
    rows = query('SELECT DISTINCT "年度" AS y FROM gideons_goal1 ORDER BY 1 DESC', 2026)
    return [r["y"] for r in rows if r.get("y")]


# ── 整合擷取（供審閱步驟）─────────────────────────────────
def fetch_all(fiscal_year: int, team: str = DEFAULT_TEAM) -> dict:
    """一次取回本財年全部資料，附帶警示供人工確認。"""
    stats = ministry_stats(fiscal_year, team)
    churches = church_testimony(fiscal_year, team)
    fees = membership_fees(fiscal_year, team)
    offerings = offerings_detail(fiscal_year, team)
    bibles = bible_giving(fiscal_year, team)
    roster = member_roster(fiscal_year, team)
    by_member = offerings_by_member(fiscal_year, team)

    start, end = fiscal_range(fiscal_year)
    warnings = []
    donation_only = [c["church"] for c in churches if c["donation_only"]]
    if donation_only:
        warnings.append(
            f"有奉獻但尚無見證場次的教會 {len(donation_only)} 間："
            + "、".join(map(str, donation_only)) + "。金額已計入教會聖奉合計。")
    if not bibles:
        warnings.append("本財年尚無贈經紀錄，贈經相關欄位會是 0／空白。")
    if not stats:
        warnings.append(f"{fiscal_year} 財年查無任何成果資料，請確認年度是否正確。")
    elif all(v["goal"] is None for v in stats.values()):
        warnings.append(
            f"{fiscal_year} 財年的**目標尚未設定**，報表「目標／差額／達成率」"
            "皆會顯示「-」。這與新財年初期的實際狀況相符，但請確認是否應先於"
            "系統建立年度目標。")
    return {
        "ok": True, "source": "api", "team": team,
        "fiscal_year": fiscal_year, "period": f"{start} ~ {end}",
        "ministry_stats": stats,
        "church_testimony": churches,
        "membership_fees": fees,
        "offerings_detail": offerings,
        "offerings_by_member": by_member,
        "bible_giving": bibles,
        "member_roster": roster,
        "counts": {"會員": len(roster), "教會見證": len(churches),
                   "會費筆數": len(fees), "奉獻筆數": len(offerings),
                   "贈經筆數": len(bibles)},
        "warnings": warnings,
    }


# ── Excel 報表（次要路徑）─────────────────────────────────
def report_url(team: str = DEFAULT_TEAM, year: int = 0) -> str:
    return f"{BASE}/report/campstat/{urllib.parse.quote(team)}?year={year}"


def dashboard_url(fiscal_year: int) -> str:
    fs = fiscal_year - 1
    return (f"{BASE}/d/ceaj87f99x79cd/"
            "e694af-e69c83-e4ba8b-e5b7a5-e68890-e69e9c-e8a1a8"
            f"?orgId=1&from={fs}-05-31T16:00:00.000Z"
            f"&to={fiscal_year}-05-31T15:59:59.999Z"
            f"&timezone=browser&var-year={fiscal_year}"
            "&var-permission=60&var-ref=5394")


def fetch_excel(dest_path: str, fiscal_year: int, team: str = DEFAULT_TEAM,
                allow_fallback: bool = False) -> dict:
    """下載成果表 Excel。

    ⚠️ `allow_fallback` 預設 **False**：年度回退會取到上一財年的數字，
    寫進本月報告就是錯的（見模組說明）。僅在你確知要補做舊月份時才開啟。
    """
    years = [fiscal_year] + ([fiscal_year - 1] if allow_fallback else [])
    attempts = []
    for y in years:
        status, body = _request(report_url(team, y))
        ok = status == 200 and body[:2] == XLSX_MAGIC
        attempts.append({"year": y, "status": status, "ok": ok})
        if not ok:
            continue
        os.makedirs(os.path.dirname(dest_path) or ".", exist_ok=True)
        with open(dest_path, "wb") as f:
            f.write(body)
        start, end = fiscal_range(y)
        out = {"ok": True, "source": "excel", "path": dest_path,
               "file": os.path.basename(dest_path), "team": team, "year": y,
               "requested_year": fiscal_year, "size": len(body),
               "url": report_url(team, y), "period": f"{start} ~ {end}",
               "fallback": y != fiscal_year, "attempts": attempts}
        if out["fallback"]:
            out["warning"] = (f"⚠️ 已改抓 {y} 財年（{out['period']}），"
                              f"這**不是** {fiscal_year} 財年的數字，切勿直接套用。")
        return out

    return {"ok": False, "source": "excel", "requested_year": fiscal_year,
            "attempts": attempts,
            "error": (f"{fiscal_year} 財年無法產生 Excel 報表"
                      "（通常是該年度目標尚未設定，伺服器回 500）。"
                      "請改用 API 擷取。")}
