#!/usr/bin/env bash
# 建立本機 git repo 並準備推上 GitHub。
#
# 沙箱環境的資料夾掛載不允許刪檔，git 無法在裡面完成 commit，
# 因此開發過程的 commit 保存在 gideons-report-app.bundle。
# 本腳本優先從 bundle 還原**完整歷史**；bundle 不在時才重新 init 成單一 commit。
#
# 用法：bash setup-git.sh
set -euo pipefail

REMOTE="https://github.com/richlai1973/gideous_haishan_monthly_report.git"
BUNDLE="gideons-report-app.bundle"
cd "$(dirname "$0")"

# 清掉沙箱留下的殘骸
rm -rf .git .gitclone ztest.tmp

# 註：不用 `git bundle verify`——它需要既有 repo 才能執行，
#     此時 .git 剛被刪掉，必定失敗。直接試 clone 才是可靠的判斷。
if [ -f "$BUNDLE" ] && git clone -q -b main "$BUNDLE" .gitclone 2>/dev/null; then
  echo "▸ 由 $BUNDLE 還原完整開發歷史…"
  mv .gitclone/.git .git
  rm -rf .gitclone
  git checkout -q main
else
  rm -rf .gitclone
  echo "▸ 找不到可用的 bundle，改建立單一 commit…"
  git init -q
  git branch -M main
  git config user.name  >/dev/null 2>&1 || git config user.name  "Richard Lai"
  git config user.email >/dev/null 2>&1 || git config user.email "richlai0614@icloud.com"
  git add -A
  git commit -q -m "基甸會海山支會月例會報告產出系統"
fi

# ── 安全檢查：確認沒有個資或憑證進版控 ──────────────────
echo
echo "▸ 版控中的檔案："
git ls-files | sed 's/^/    /'

if git ls-files | grep -Eqi '\.docx$|\.xlsx$|credentials/|token\.json|client_secret'; then
  echo
  echo "❌ 偵測到 docx／xlsx／憑證進了版控，已中止。請檢查 .gitignore。"
  exit 1
fi
echo "✅ 未含 docx／xlsx／憑證等機敏檔"

# ── 遠端 ────────────────────────────────────────────────
git remote add origin "$REMOTE" 2>/dev/null || git remote set-url origin "$REMOTE"

echo
git --no-pager log --oneline
echo
echo "▸ 遠端：$REMOTE"
echo
echo "接著執行推送："
echo "    git push -u origin main"
echo
echo "若遠端已有內容而被拒絕，先確認要不要覆寫，再用："
echo "    git push -u origin main --force"
