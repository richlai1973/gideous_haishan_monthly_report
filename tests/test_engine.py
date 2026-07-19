"""核心邏輯單元測試：python3 -m pytest tests/ -q"""

import os
import sys
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from engine.dates import (build_meta, default_report_month, fiscal_year_of,
                          fourth_sunday, roc_year, template_month)
from engine.generate import _church_key, _fmt_stat, _match_church
from engine.parse_files import filter_month, parse_line_text, split_by_day


# ── 日期／財年 ───────────────────────────────────────────
def test_roc_year():
    assert roc_year(2026) == 115


def test_fourth_sunday():
    assert fourth_sunday(2026, 7) == date(2026, 7, 26)
    assert fourth_sunday(2026, 6) == date(2026, 6, 28)


def test_template_month_wraps_january():
    assert template_month(2026, 7) == (2026, 6)
    assert template_month(2026, 1) == (2025, 12)


def test_fiscal_year_boundary():
    assert fiscal_year_of(2026, 5) == 2026    # 5 月 → 本財年
    assert fiscal_year_of(2026, 6) == 2027    # 6 月 → 下一財年
    assert fiscal_year_of(2026, 12) == 2027


def test_build_meta():
    m = build_meta(2026, 7)
    assert (m.roc_year, m.prev_month, m.fiscal_year) == (115, 6, 2027)
    assert m.meeting_date == "2026-07-26"
    assert m.period == "2026-2027"
    assert m.drive_folder_name == "2026年7月"
    assert "var-year=2027" in m.dashboard_url


def test_build_meta_override_meeting_date():
    assert build_meta(2026, 7, "2026-07-19").meeting_date == "2026-07-19"


# ── 達成率格式（Excel 以小數儲存，1 = 100%）──────────────
def test_fmt_rate():
    assert _fmt_stat("rate", 1, "") == "100%"
    assert _fmt_stat("rate", 0.6812, "") == "68%"
    assert _fmt_stat("actual", 184772, "") == "184,772"
    assert _fmt_stat("target", None, "人") == "-"


# ── 教會名模糊比對 ───────────────────────────────────────
def test_church_key_normalisation():
    assert _church_key("板城靈糧堂教會") == _church_key("板城靈糧堂")


def test_match_church():
    chs = [{"church": "板城靈糧堂", "amount": 13300, "date": None, "speaker": None},
           {"church": "貴格會土城教會", "amount": 5500, "date": None, "speaker": None}]
    c, how = _match_church("貴格會土城教會", chs)
    assert how == "exact" and c["amount"] == 5500
    c, how = _match_church("板城靈糧堂教會", chs)
    assert how == "fuzzy" and c["amount"] == 13300
    assert _match_church("完全不存在的教會", chs)[0] is None


# ── LINE 解析 ────────────────────────────────────────────
CURATED = """📖 年度主題與目標 (2026-2027)
• 年度主題：傳道書 10:10
• 支會目標：招募 3位弟兄、2位姊妹
• 贈經目標：3,500本聖經
• 姐妹贈經行程：
 07/18：土城醫院
 08/12(三)：樹林中山路診所
🙏 靈修（實體禱告會時間表）
• 第一週：鶯歌 (週六 16:00)
• 第二週：三峽 (週六 16:00)
• 第三週：樹林 (週六 16:00)
• 第四週：土城 (週日 15:30)
"""


def test_parse_curated_line_summary():
    r = parse_line_text(CURATED, 2026)
    assert r["annual"]["theme"].startswith("傳道書")
    assert r["annual"]["period"] == "2026-2027"
    assert r["annual"]["recruit_target"] == {"brothers": 3, "sisters": 2}
    assert r["annual"]["bible_target"] == 3500
    sched = r["bible_giving"]["schedule"]
    assert {s["date"] for s in sched} == {"2026-07-18", "2026-08-12"}
    assert len(r["prayer_rota"]["weeks"]) == 4


def test_raw_chat_does_not_produce_junk_schedule():
    """原始聊天記錄不該吐出上百筆假排程。"""
    noise = "\n".join(
        "各位基甸勇士：平安！本學期第5次贈經將於2023年11月1日在中正國中舉行，"
        "預計學生中文288本（6箱）茲公佈如下：" for _ in range(50))
    r = parse_line_text(noise, 2026)
    assert len(r.get("bible_giving", {}).get("schedule", [])) == 0


RAW = """2026.06.01 星期一
08:09\t葉忠濶\t圖片
10:58\tCaleb哲斌\t忠濶兄早上好
2026.07.02 星期四
09:00\t王木林\t本週禱告會
"""


def test_split_and_filter_month():
    assert [d for d, _ in split_by_day(RAW)] == ["2026-06-01", "2026-07-02"]
    r = filter_month(RAW, 2026, 6)
    assert r["matched_days"] == 1 and len(r["messages"]) == 2
    assert r["messages"][0]["who"] == "葉忠濶"
    assert r["messages"][1]["text"] == "忠濶兄早上好"
    assert r["range"] == ("2026-06-01", "2026-07-02")
