#!/usr/bin/env bash
# リポジトリ内で変更があれば commit & push
# scheduled task から呼ばれる
set -euo pipefail

cd "$(dirname "$0")/.."

MSG="${1:-auto update}"

# 何も変更がなければ何もしない
if git diff --quiet && git diff --cached --quiet; then
  echo "no changes"
  exit 0
fi

git add -A
git -c user.email="auto@game-tracker" -c user.name="Game Tracker Bot" \
    commit -m "$MSG ($(date '+%Y-%m-%d %H:%M'))"
git push origin main
echo "pushed"
