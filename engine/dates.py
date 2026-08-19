"""日期／民國年／財年衍生計算。

三個月份勿混淆（見設計計畫書 §15.2）：
  ①報告月份 = 送出當下年月
  ②範本月份 = 報告月份的上一個月
  ③財年 var-year = 最新財年（Grafana 用）
"""

from __future__ import annotations

from dataclasses import dataclass, asdict
from datetime import date, timedelta

CHINESE_MONTHS = ["一", "二", "三", "四", "五", "六",
                  "七", "八", "九", "十", "十一", "十二"]

DASHBOARD_BASE = ("https://gideons-dashboard.pointing.tw/d/ceaj87f99x79cd/"
                  "e694af-e69c83-e4ba8b-e5b7a5-e68890-e69e9c-e8a1a8")


def roc_year(western_year: int) -> int:
    """民國年 = 西元年 - 1911。2026 → 115"""
    return western_year - 1911


def chinese_month_name(month: int) -> str:
    """1 → '一月份'"""
    return f"{CHINESE_MONTHS[month - 1]}月份"


def fourth_sunday(year: int, month: int) -> date:
    """當月第四個禮拜天（預設會議日期）。"""
    d = date(year, month, 1)
    while d.weekday() != 6:  # 6 = Sunday
        d += timedelta(days=1)
    return d + timedelta(weeks=3)


def default_report_month(today: date | None = None) -> tuple[int, int]:
    """預設報告月份 = 送出當下年月。2026-07-19 → (2026, 7)"""
    d = today or date.today()
    return d.year, d.month


def template_month(year: int, month: int) -> tuple[int, int]:
    """範本 = 上一個月。(2026,7)→(2026,6)；(2026,1)→(2025,12)"""
    return (year - 1, 12) if month == 1 else (year, month - 1)


def next_month(year: int, month: int) -> tuple[int, int]:
    return (year + 1, 1) if month == 12 else (year, month + 1)


def fiscal_year_of(year: int, month: int) -> int:
    """財年 6/1~5/31；月份 >= 6 屬下一財年。"""
    return year + 1 if month >= 6 else year


def latest_fiscal_year(today: date | None = None) -> int:
    d = today or date.today()
    return fiscal_year_of(d.year, d.month)


def dashboard_url(fiscal_year: int) -> str:
    """組出該財年的 Grafana 事工成果表 dashboard URL。"""
    fs = fiscal_year - 1
    return (f"{DASHBOARD_BASE}?orgId=1"
            f"&from={fs}-05-31T16:00:00.000Z"
            f"&to={fiscal_year}-05-31T15:59:59.999Z"
            f"&timezone=browser&var-year={fiscal_year}"
            f"&var-permission=60&var-ref=5394")


@dataclass
class Meta:
    report_year: int
    report_month: int
    roc_year: int
    meeting_date: str
    next_meeting_date: str
    fiscal_year: int
    fiscal_start_year: int
    prev_year: int
    prev_month: int
    prev_roc_year: int
    prev_meeting_date: str
    period: str                # 財年期間，如 "2026-2027"
    work_dir_name: str         # "2026年7月月例會"
    dashboard_url: str

    def to_dict(self) -> dict:
        return asdict(self)


def build_meta(year: int, month: int, meeting_date: str | None = None,
               template_ym: tuple[int, int] | None = None) -> Meta:
    """由報告年月推導全部衍生欄位。meeting_date 可覆寫（YYYY-MM-DD）。

    `template_ym` 覆寫「範本月份」。預設是上個月，但改用 repo 內建的固定範本時
    範本停在某一個月（如 115年7月），`prev_*` 必須跟著它走——update_dates() 全部
    的替換都以 prev_* 為來源字串，對不上就一個字都換不到。
    """
    py, pm = template_ym or template_month(year, month)
    ny, nm = next_month(year, month)
    fy = fiscal_year_of(year, month)

    if meeting_date:
        y, m, d = (int(x) for x in meeting_date.split("-"))
        md = date(y, m, d)
    else:
        md = fourth_sunday(year, month)

    return Meta(
        report_year=year,
        report_month=month,
        roc_year=roc_year(year),
        meeting_date=md.isoformat(),
        next_meeting_date=fourth_sunday(ny, nm).isoformat(),
        fiscal_year=fy,
        fiscal_start_year=fy - 1,
        prev_year=py,
        prev_month=pm,
        prev_roc_year=roc_year(py),
        prev_meeting_date=fourth_sunday(py, pm).isoformat(),
        period=f"{fy - 1}-{fy}",
        work_dir_name=f"{year}年{month}月月例會",
        dashboard_url=dashboard_url(fy),
    )
