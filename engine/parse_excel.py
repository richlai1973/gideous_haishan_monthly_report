"""Grafana 事工成果表 Excel 解析（Sheet：支會各項事工成果表）。

實際欄位（已用 2026/06 匯出檔驗證）：
  Row 3-4 欄位標頭｜Row 5 目標｜Row 6 成果｜Row 7 差額｜Row 8 達成率｜Row 9+ 會員
  A 編號｜D 弟兄 E 姊妹｜F-G 會費日期｜H-I 聖經奉獻｜J-K 巴拿巴
  L 見證日期 M 講員 N 教會 O 奉獻金額｜P 贈經類別 Q 贈經本數

注意：P/Q 在「會員列」是贈經明細（類別＋本數），與該欄標頭
（贈經總數／姊妹贈經）語意不同；總數一律取 Row 5-8 的匯總值。
"""

from __future__ import annotations

import datetime as _dt
import os

import openpyxl

# 會員列欄位
COLS = {
    "id": 1, "brother": 4, "sister": 5,
    "fee_brother": 6, "fee_sister": 7,
    "bible_brother": 8, "bible_sister": 9,
    "barnabas_brother": 10, "barnabas_sister": 11,
    "church_date": 12, "church_speaker": 13,
    "church_name": 14, "church_offering": 15,
    "scripture_kind": 16, "scripture_count": 17,
}

# 匯總列（Row 5-8）欄位 → 語意名稱
SUMMARY_COLS = {
    "members_brother": 4, "members_sister": 5,
    "fee_brother": 6, "fee_sister": 7,
    "bible_brother": 8, "bible_sister": 9,
    "barnabas_brother": 10, "barnabas_sister": 11,
    "church_count": 12, "church_offering": 15,
    "scripture_total": 16, "scripture_sister": 17,
}

SUMMARY_ROWS = {"target": 5, "actual": 6, "diff": 7, "rate": 8}

# 會費欄位可能出現的人工註記（非日期，不可覆寫）
FEE_NOTES = ("安息", "退會", "入會", "停權", "轉會")


def _norm(v):
    if v is None:
        return None
    if isinstance(v, (_dt.datetime, _dt.date)):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if isinstance(v, str):
        return v.strip() or None
    return v


def _num(v) -> float | int:
    """僅取數值，字串／None 視為 0。"""
    return v if isinstance(v, (int, float)) else 0


class NotMinistryExcel(ValueError):
    """此 Excel 不是事工成果表（例：現金帳、贈經計畫）。"""


class MinistryExcel:
    def __init__(self, path: str, strict: bool = True):
        self.path = path
        wb = openpyxl.load_workbook(path, data_only=True)
        named = [n for n in wb.sheetnames if "事工成果" in n]
        self.ws = wb[named[0]] if named else wb[wb.sheetnames[0]]
        self.sheet_name = self.ws.title
        if strict and not named and not self._looks_like_ministry():
            raise NotMinistryExcel(
                f"「{os.path.basename(path)}」不是事工成果表"
                f"（工作表：{', '.join(wb.sheetnames[:3])}）")
        self._parse()

    def _looks_like_ministry(self) -> bool:
        """以 Row 3-4 標頭與 Row 5-8 標籤判定，避免誤認其他 Excel。"""
        ws = self.ws
        head = " ".join(str(ws.cell(r, c).value or "")
                        for r in (1, 3, 4) for c in range(1, 18))
        labels = [str(ws.cell(r, 1).value or "") for r in range(5, 9)]
        has_head = sum(k in head for k in
                       ("會員會費", "聖經奉獻", "教會見證", "贈經", "巴拿巴")) >= 3
        has_labels = ["目標", "成果", "差額", "達成率"] == labels
        return has_head and has_labels

    # ── 解析 ────────────────────────────────────────────
    def _parse(self):
        ws = self.ws
        self.title = _norm(ws.cell(1, 7).value)
        self.date_range = _norm(ws.cell(1, 12).value)
        self.current_members = {
            "brother": _num(ws.cell(4, 4).value),
            "sister": _num(ws.cell(4, 5).value),
        }

        # 匯總（目標/成果/差額/達成率）
        self.summary = {
            key: {name: _norm(ws.cell(row, col).value)
                  for name, col in SUMMARY_COLS.items()}
            for key, row in SUMMARY_ROWS.items()
        }

        # 會員列
        self.members = []
        for row in range(9, ws.max_row + 1):
            if not ws.cell(row, COLS["id"]).value:
                continue
            rec = {k: _norm(ws.cell(row, c).value) for k, c in COLS.items()}
            rec["row"] = row
            rec["id"] = str(rec["id"])
            self.members.append(rec)

    # ── 查詢 ────────────────────────────────────────────
    def member_by_name(self, name: str):
        name = (name or "").strip()
        if not name:
            return None
        for m in self.members:
            for f in ("brother", "sister"):
                val = m[f]
                if val and (name == str(val) or name in str(val)):
                    return m
        return None

    def churches(self) -> list[dict]:
        return [{
            "date": m["church_date"], "speaker": m["church_speaker"],
            "church": m["church_name"], "amount": _num(m["church_offering"]),
        } for m in self.members if m["church_name"]]

    def offerings(self) -> list[dict]:
        out = []
        for m in self.members:
            for gender, nk, fk, bk, kk in (
                ("M", "brother", "fee_brother", "bible_brother", "barnabas_brother"),
                ("F", "sister", "fee_sister", "bible_sister", "barnabas_sister"),
            ):
                name = m[nk]
                if not name or any(n in str(name) for n in FEE_NOTES):
                    continue
                fee = m[fk]
                fee_note = fee if fee and any(n in str(fee) for n in FEE_NOTES) else None
                out.append({
                    "id": m["id"], "name": str(name), "gender": gender,
                    "fee_date": None if fee_note else fee,
                    "fee_note": fee_note,
                    "bible_offering": _num(m[bk]),
                    "barnabas": _num(m[kk]),
                })
        return out

    def scripture_breakdown(self) -> list[dict]:
        """P/Q 欄的贈經明細（類別 → 本數）。"""
        return [{"kind": str(m["scripture_kind"]), "count": _num(m["scripture_count"])}
                for m in self.members if m["scripture_kind"]]

    # ── 分析（供 Web 檢視）───────────────────────────────
    def analysis(self) -> dict:
        offs = self.offerings()
        chs = self.churches()
        act, tgt = self.summary["actual"], self.summary["target"]
        active = [o for o in offs if not o["fee_note"]]
        paid = [o for o in active if o["fee_date"]]

        def rate(a, t):
            a, t = _num(a), _num(t)
            return round(a / t * 100, 1) if t else None

        return {
            "sheet": self.sheet_name,
            "title": self.title,
            "date_range": self.date_range,
            "current_members": self.current_members,
            "member_rows": len(self.members),
            "people": len(active),
            "fee_paid": len(paid),
            "fee_unpaid": len(active) - len(paid),
            "fee_rate": round(len(paid) / len(active) * 100, 1) if active else 0,
            "bible_offering_total": sum(o["bible_offering"] for o in offs),
            "barnabas_total": sum(o["barnabas"] for o in offs),
            "church_count": len(chs),
            "church_offering_total": sum(c["amount"] for c in chs),
            # 官方匯總值（以 Excel 為準，非自行加總）
            "official": {
                "scripture_total": _num(act["scripture_total"]),
                "scripture_target": _num(tgt["scripture_total"]),
                "scripture_rate": rate(act["scripture_total"], tgt["scripture_total"]),
                "sister_scripture": _num(act["scripture_sister"]),
                "sister_target": _num(tgt["scripture_sister"]),
                "church_offering": _num(act["church_offering"]),
                "church_offering_target": _num(tgt["church_offering"]),
                "church_offering_rate": rate(act["church_offering"], tgt["church_offering"]),
                "bible_offering": _num(act["bible_brother"]) + _num(act["bible_sister"]),
                "bible_offering_target": _num(tgt["bible_brother"]) + _num(tgt["bible_sister"]),
                "recruit_actual": _num(act["members_brother"]) + _num(act["members_sister"]),
                "recruit_target": _num(tgt["members_brother"]) + _num(tgt["members_sister"]),
            },
            "summary": self.summary,
            "scripture_breakdown": self.scripture_breakdown(),
            "top_bible_offering": sorted(
                [o for o in offs if o["bible_offering"]],
                key=lambda x: x["bible_offering"], reverse=True)[:10],
            "top_churches": sorted(chs, key=lambda c: c["amount"], reverse=True)[:10],
            "unpaid_members": [o["name"] for o in active if not o["fee_date"]],
        }

    def to_model(self) -> dict:
        return {
            "ministry_stats": self.summary,
            "offerings": self.offerings(),
            "church_testimony": self.churches(),
            "analysis": self.analysis(),
        }
