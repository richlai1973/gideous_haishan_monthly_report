"""產出引擎：複製前月 10 份範本 → 重新命名 → 逐檔更新日期與數據。"""

from __future__ import annotations

import glob
import os
import re
import shutil
from datetime import date

from docx import Document

from .dates import Meta, build_meta
from .docx_utils import get_cell_text, replace_text_in_doc, set_cell_text, update_dates
from .parse_excel import MinistryExcel

DOC_SUFFIXES = [
    "-1議程", "-2事工成果統計表", "-3收入支用統計表", "-4各項奉獻",
    "-5贈經事工(除學校)", "-6學校贈經統計表", "-7會員及地界教會代禱項目",
    "-8早禱會及月例會輪值表", "-9年度教會見證統計表(橫式)", "-10會員名冊(橫式)",
]


def doc_filename(roc: int, month: int, suffix: str) -> str:
    return f"月例會議程{roc}年{month}月{suffix}.docx"


def file_number(fname: str) -> int:
    m = re.search(r"月\s*-(\d+)", fname) or re.search(r"-(\d+)", fname)
    return int(m.group(1)) if m else 0


# ── Step 1：初始化當月工作區 ─────────────────────────────
def init_month(base_dir: str, meta: Meta, template_dir: str | None = None) -> dict:
    """建立當月資料夾、由前月複製 10 份 docx 並重新命名。"""
    work_dir = os.path.join(base_dir, meta.work_dir_name)
    prev_dir = template_dir or os.path.join(base_dir, f"{meta.prev_year}年{meta.prev_month}月月例會")

    if not os.path.isdir(prev_dir):
        return {"ok": False, "error": f"找不到範本資料夾：{prev_dir}", "work_dir": work_dir}

    os.makedirs(work_dir, exist_ok=True)
    os.makedirs(os.path.join(work_dir, "_inputs"), exist_ok=True)

    files, skipped = [], []
    for suffix in DOC_SUFFIXES:
        src_name = doc_filename(meta.prev_roc_year, meta.prev_month, suffix)
        src = os.path.join(prev_dir, src_name)
        if not os.path.exists(src):
            cands = glob.glob(os.path.join(prev_dir, f"*{glob.escape(suffix)}.docx"))
            if not cands:
                skipped.append(suffix)
                continue
            src = cands[0]
        dst = os.path.join(work_dir, doc_filename(meta.roc_year, meta.report_month, suffix))
        if not os.path.exists(dst):
            shutil.copy2(src, dst)
        files.append({"suffix": suffix, "name": os.path.basename(dst),
                      "path": dst, "status": "copied"})

    return {"ok": True, "work_dir": work_dir, "template_dir": prev_dir,
            "files": files, "missing": skipped}


# ── Step 2：逐檔更新 ─────────────────────────────────────
def generate_all(work_dir: str, meta: Meta, excel_path: str | None = None,
                 model: dict | None = None) -> dict:
    """更新工作區內的 10 份 docx。回傳每份的狀態與變更紀錄。"""
    model = model or {}
    data = None
    if excel_path and os.path.exists(excel_path):
        data = MinistryExcel(excel_path)

    prefix = f"月例會議程{meta.roc_year}年{meta.report_month}月"
    paths = sorted(glob.glob(os.path.join(work_dir, f"{prefix}-*.docx")),
                   key=lambda p: file_number(os.path.basename(p)))
    if not paths:
        return {"ok": False, "error": f"工作區找不到 {prefix}-*.docx", "results": []}

    results = []
    for path in paths:
        fname = os.path.basename(path)
        num = file_number(fname)
        doc = Document(path)
        log = update_dates(doc, meta)
        status, notes = "ok", []

        try:
            if num == 1:
                notes += _update_agenda(doc, meta, model)
            elif num == 2 and data:
                notes += _update_ministry(doc, data)
            elif num == 4 and data:
                notes += _update_offerings(doc, data)
            elif num == 5:
                notes += _update_bible_giving(doc, meta, model)
            elif num == 9 and data:
                notes += _update_church_testimony(doc, data)
        except Exception as exc:  # 不讓單一檔失敗中斷整批
            status, notes = "warn", notes + [f"更新時發生例外：{exc}"]

        doc.save(path)
        results.append({
            "file": fname, "num": num, "path": path, "status": status,
            "date_changes": log, "notes": notes,
            "data_updated": bool(notes),
        })

    return {"ok": True, "work_dir": work_dir, "results": results,
            "excel": os.path.basename(excel_path) if excel_path else None}


# ── 各文件更新 ───────────────────────────────────────────
def _find_row(table, keyword: str, col: int = 0):
    for row in table.rows:
        if keyword in get_cell_text(row.cells[col]):
            return row
    return None


def _update_agenda(doc, meta: Meta, model: dict) -> list[str]:
    """-1 議程：年度主題／目標；未定項目標 TODO。"""
    notes = []
    annual = model.get("annual") or {}
    if theme := annual.get("theme"):
        if replace_text_in_doc(doc, "【年度主題】", theme):
            notes.append(f"年度主題 → {theme}")
    if bt := annual.get("bible_target"):
        notes.append(f"年度贈經目標 {bt:,} 本（供人工核對）")
    todos = model.get("agenda_todo", {})
    for key, label in (("motions", "臨時動議"), ("guest_reports", "來賓報告"),
                       ("next_events", "下月活動預告")):
        if items := todos.get(key):
            notes.append(f"{label}：{len(items)} 筆（【TODO】待人工置入）")
    return notes


# -2 docx 欄索引 → Excel 匯總欄語意名稱（已用 2026/06 匯出檔驗證）
MINISTRY_COL_MAP = {
    1: ("members_brother", "人"),   # 會員數 弟兄
    2: ("members_sister", "人"),    # 會員數 姐妹
    3: ("fee_brother", "人"),       # 會費 弟兄
    4: ("fee_sister", "人"),        # 會費 姐妹
    5: ("scripture_total", ""),     # 贈送聖經(本)
    6: ("church_count", "次"),      # 教會見證 次數
    7: ("church_offering", ""),     # 教會聖奉
    8: ("bible_brother", ""),       # 會員聖奉 弟兄
    9: ("bible_sister", ""),        # 會員聖奉 姊妹
    10: ("scripture_sister", ""),   # 姊妹贈經
}

MINISTRY_ROW_LABELS = {"目標": "target", "成果": "actual",
                       "差額": "diff", "達成率": "rate"}


def _fmt_stat(key: str, val, unit: str) -> str:
    if val in (None, ""):
        return "-"
    if key == "rate":
        # Excel 的達成率以小數儲存（1 = 100%）
        return f"{val * 100:.0f}%" if isinstance(val, (int, float)) else f"{val}"
    if isinstance(val, (int, float)):
        return f"{int(round(val)):,}{unit}"
    return f"{val}{unit}"


def _update_ministry(doc, data: MinistryExcel) -> list[str]:
    """-2 事工成果統計表：目標／成果／差額／達成率四列直接對映 Excel Row 5-8。"""
    table = doc.tables[0]
    notes, filled = [], 0

    for label, key in MINISTRY_ROW_LABELS.items():
        row = _find_row(table, label)
        if row is None:
            notes.append(f"找不到「{label}」列")
            continue
        stats = data.summary[key]
        for col, (field, unit) in MINISTRY_COL_MAP.items():
            if col >= len(row.cells):
                continue
            val = stats.get(field)
            set_cell_text(row.cells[col],
                          _fmt_stat(key, val, "" if key in ("rate",) else unit))
            filled += 1

    act = data.summary["actual"]
    notes.append(f"四列共 {filled} 欄已對映 Excel Row 5-8"
                 f"（成果：會費 {act['fee_brother']}/{act['fee_sister']}、"
                 f"贈經 {act['scripture_total']}、教會聖奉 {act['church_offering']}）")
    return notes


def _update_offerings(doc, data: MinistryExcel) -> list[str]:
    """-4 各項奉獻：依會員姓名比對，更新會費日期與聖經奉獻。"""
    table = doc.tables[0]
    updated = 0
    for row in table.rows[2:]:
        cells = row.cells
        if len(cells) < 7:
            continue
        label = get_cell_text(cells[0])
        if label in ("目標", "成果", "差額", "達成率", ""):
            continue
        for name_col, fee_col, bible_col, field in (
            (1, 3, 5, "brother"), (2, 4, 6, "sister"),
        ):
            name = get_cell_text(cells[name_col])
            if not name:
                continue
            m = data.member_by_name(name)
            if not m:
                continue
            fee = m[f"fee_{field}"]
            bible = m[f"bible_{field}"]
            if get_cell_text(cells[fee_col]) in ("安息", "退會"):  # 保留人工註記
                continue
            if fee:
                set_cell_text(cells[fee_col], str(fee))
                updated += 1
            if bible:
                set_cell_text(cells[bible_col], f"{int(bible):,}")
                updated += 1

    # 四列匯總直接取 Excel Row 5-8（會費人數、聖經奉獻）
    labels = 0
    for label, key in (("目標", "target"), ("成果", "actual"),
                       ("差額", "diff"), ("達成率", "rate")):
        row = _find_row(table, label)
        if row is None or len(row.cells) < 7:
            continue
        s = data.summary[key]
        for col, field in ((3, "fee_brother"), (4, "fee_sister"),
                           (5, "bible_brother"), (6, "bible_sister")):
            set_cell_text(row.cells[col], _fmt_stat(key, s.get(field), ""))
        labels += 1
    return [f"各項奉獻已更新 {updated} 個欄位、{labels} 列匯總已對映 Excel"]


_CHURCH_NOISE = ("基督教", "基督", "台灣", "臺灣", "教會", "禮拜堂", "福音中心",
                 "堂", "會", "（北中）", "(北中)", " ", "　")


def _church_key(name: str) -> str:
    """教會名正規化：去掉宗派／型態等雜訊字，只留核心地名。

    例：「板城靈糧堂教會」與「板城靈糧堂」→ 同一 key「板城靈糧」。
    """
    s = str(name or "")
    for w in _CHURCH_NOISE:
        s = s.replace(w, "")
    return s


def _match_church(cname: str, chs: list[dict]):
    """先精確、再正規化包含比對。回傳 (教會資料, 比對方式) 或 (None, '')。"""
    for c in chs:
        if str(c["church"]) == cname:
            return c, "exact"
    k = _church_key(cname)
    if len(k) < 2:
        return None, ""
    for c in chs:
        ck = _church_key(c["church"])
        if ck and (k == ck or k in ck or ck in k):
            return c, "fuzzy"
    return None, ""


def _update_church_testimony(doc, data: MinistryExcel) -> list[str]:
    """-9 年度教會見證統計表：各教會日期／講員／金額 + 最後匯總列（易遺漏）。"""
    table = doc.tables[0]
    chs = data.churches()
    used, exact, fuzzy, unmatched = set(), 0, 0, []

    for row in table.rows[2:]:
        cells = row.cells
        if len(cells) < 12:
            continue
        cname = get_cell_text(cells[1])
        if not cname or "小計" in cname or "目標" in cname:
            continue
        c, how = _match_church(cname, chs)
        if not c:
            unmatched.append(cname)
            continue
        if c["date"]:
            set_cell_text(cells[5], str(c["date"]).replace("-", "/"))
        if c["speaker"]:
            set_cell_text(cells[9], str(c["speaker"]))
        set_cell_text(cells[11], f"{int(c['amount']):,}" if c["amount"] else "")
        used.add(str(c["church"]))
        exact += how == "exact"
        fuzzy += how == "fuzzy"

    # ── 匯總列：一律採 Excel 官方匯總值，不自行加總 ────────
    act, tgt = data.summary["actual"], data.summary["target"]
    total_amount = int(act.get("church_offering") or 0)
    target_amount = int(tgt.get("church_offering") or 0)
    done_count = int(act.get("church_count") or 0)
    target_count = int(tgt.get("church_count") or 0)
    rate = f"{total_amount / target_amount * 100:.0f}%" if target_amount else "-"

    summary_notes = []
    for row in table.rows[-4:]:
        text = get_cell_text(row.cells[0])
        if "小計" in text and len(row.cells) > 11:
            set_cell_text(row.cells[11], f"{total_amount:,}")
            summary_notes.append(f"小計 {total_amount:,}")
        elif "目標間數" in text:
            set_cell_text(row.cells[0],
                          f"目標間數: {target_count or '-'}    已達成: {done_count} 間")
            summary_notes.append(f"間數 {done_count}/{target_count or '-'}")
        elif "目標金額" in text:
            set_cell_text(row.cells[0],
                          f"目標金額: {target_amount:,}    已達成: {total_amount:,} 元"
                          f"    達成率: {rate}")
            summary_notes.append(f"金額 {total_amount:,}（達成率 {rate}）")

    notes = [f"教會見證已更新 {exact + fuzzy} 間（精確 {exact}、模糊 {fuzzy}）",
             "匯總列：" + ("、".join(summary_notes) if summary_notes else "未找到，請人工檢查")]
    if unmatched:
        notes.append(f"⚠️ 未對應到 Excel 的教會 {len(unmatched)} 間："
                     + "、".join(unmatched[:8]))
    left = [c["church"] for c in chs if str(c["church"]) not in used]
    if left:
        notes.append(f"⚠️ Excel 有但表中未列 {len(left)} 間：" + "、".join(map(str, left[:8])))
    return notes


def _update_bible_giving(doc, meta: Meta, model: dict) -> list[str]:
    """-5 贈經事工：列出當月排程（僅回報，內容由人確認後貼入）。"""
    sched = (model.get("bible_giving") or {}).get("schedule") or []
    this_month = [s for s in sched if s.get("date", "") and
                  s["date"][:7] == f"{meta.report_year}-{meta.report_month:02d}"]
    if not this_month:
        return []
    return [f"本月贈經排程 {len(this_month)} 筆："
            + "、".join(f"{s['date'][5:]} {s['target']}" for s in this_month)]


# ── 便利入口 ─────────────────────────────────────────────
def run(base_dir: str, year: int, month: int, meeting_date: str | None = None,
        excel_path: str | None = None, model: dict | None = None) -> dict:
    meta = build_meta(year, month, meeting_date)
    init = init_month(base_dir, meta)
    if not init["ok"]:
        return init
    gen = generate_all(init["work_dir"], meta, excel_path, model)
    gen["meta"] = meta.to_dict()
    gen["init"] = init
    return gen
