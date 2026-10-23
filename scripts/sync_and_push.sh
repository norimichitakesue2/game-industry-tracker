#!/usr/bin/env bash
# リポジトリ内で変更があれば commit & push（scheduled task から呼ばれる）
# - 未コミット変更があれば commit
# - origin/main より先行しているコミットがあれば push（commit済みで呼ばれても push 漏れしない）
# - push に失敗したら exit 2 と PUSH_FAILED を出力。scripts/PUSH_FALLBACK.md の手順でPC経由pushすること
set -uo pipefail

cd "$(dirname "$0")/.."

MSG="${1:-auto update}"

if ! git diff --quiet || ! git diff --cached --quiet || [ -n "$(git ls-files --others --exclude-standard)" ]; then
  git add -A
  git -c user.email="noreply@anthropic.com" -c user.name="Claude" \
      commit -q -m "$MSG ($(date '+%Y-%m-%d %H:%M'))"
fi

git fetch -q origin main 2>/dev/null || true
AHEAD=$(git rev-list --count origin/main..HEAD 2>/dev/null || echo 1)
if [ "$AHEAD" = "0" ]; then
  echo "no changes (nothing to push)"
  exit 0
fi

git pull -q --rebase origin main 2>/dev/null || git rebase --abort 2>/dev/null || true

if git push origin HEAD:main 2>&1 | tee /tmp/sync_push.log; [ "${PIPESTATUS[0]}" = "0" ]; then
  echo "pushed: $(git rev-parse --short HEAD)"
  exit 0
fi

echo ""
echo "=================================================================="
echo "PUSH_FAILED: クラウドからの push が拒否されました（git proxy 403 等）。"
echo "コミットは手元に残っています: $(git rev-parse --short HEAD)"
echo "作業を捨てずに scripts/PUSH_FALLBACK.md の手順でユーザーPC経由で push してください。"
echo "それも不可なら PushNotification で必ず失敗を通知すること（黙って終了しない）。"
echo "=================================================================="
exit 2
