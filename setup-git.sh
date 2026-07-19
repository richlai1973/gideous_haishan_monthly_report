#!/usr/bin/env bash
# 在本機建立 git repo 並推上 GitHub。
# 用法：cd gideons-report-app && bash setup-git.sh
set -euo pipefail

REMOTE="https://github.com/richlai1973/gideous_haishan_monthly_report.git"
cd "$(dirname "$0")"

# 沙箱環境無法刪檔，留下的殘留物在此清掉
rm -rf .git ztest.tmp

git init -q
git branch -M main

# 若尚未設定 git 身分，就地補上（不動全域設定）
git config user.name  >/dev/null 2>&1 || git config user.name  "Richard Lai"
git config user.email >/dev/null 2>&1 || git config user.email "richlai0614@icloud.com"

git add -A

echo "── 即將提交的檔案 ──────────────────────────"
git status --short
echo "────────────────────────────────────────────"
if git status --short | grep -Eqi "\.docx|\.xlsx|token\.json|client_secret|_inputs/"; then
  echo "⚠️  偵測到可能含個資的檔案，已中止。請檢查 .gitignore 後重試。"
  exit 1
fi
echo "✅ 未含 docx／xlsx／憑證等機敏檔"

git commit -q -F - <<'MSG'
月例會報告產出系統 v1.0：Web 介面、docx 引擎、Excel 分析、Drive 上傳

依《Web 設計計畫書 v2.4》方案 A（本機 FastAPI）實作：

- dates.py：民國年、財年 6/1-5/31、第四個禮拜天、Grafana URL 自動組裝
- docx_utils.py：set_cell_text 處理多段落 cell；日期替換依「具體日期 → 中文月份
  → 通用年月 → 斜線年月 → 財年期間」固定順序
- parse_excel.py：事工成果表解析，已用 2026/06 實際匯出檔校正欄位——P/Q 在會員列
  是贈經明細（類別＋本數），非欄標頭語意，總數改取 Row 5-8 匯總
- generate.py：複製範本並更新 10 份；-2/-4 四列直接對映 Excel Row 5-8，
  -9 教會名正規化模糊比對 + 匯總列採官方值，未對應項目主動警示
- parse_files.py：txt/csv/xlsx/pdf/jpeg 分派。LINE 原始匯出為跨年聊天記錄，
  改為篩出當月訊息供人工挑選，並加去重與長度限制避免產生假排程
- drive.py：OAuth、依「{年}年{月}」自動建資料夾、同名檔 update 覆寫
- app.py + static/index.html：五大模組頁籤、解析先確認再寫入、成果表分析、
  增量產出、ZIP 下載、Drive 上傳
- tests/：12 項單元測試（日期邊界、達成率格式、教會比對、LINE 解析）

.gitignore 排除 credentials/、*.docx、*.xlsx、_inputs/（含個資與奉獻金額）。
MSG

git remote add origin "$REMOTE" 2>/dev/null || git remote set-url origin "$REMOTE"

echo
echo "✅ 已建立本機 commit：$(git log --oneline -1)"
echo "推送請執行： git push -u origin main"
