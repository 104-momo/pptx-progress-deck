#!/bin/bash
# 一键创建 GitHub 仓库并推送
# 用法：GITHUB_TOKEN=xxx bash deploy.sh [仓库名] [public|private]
#       bash deploy.sh                    # 也可从 /tmp/gh_token 读
set -e

REPO="${1:-pptx-progress-deck}"
VIS="${2:-private}"

TOKEN="${GITHUB_TOKEN:-}"
if [ -z "$TOKEN" ]; then
  [ -f /tmp/gh_token ] || { echo "✗ 没有 token。请先 export GITHUB_TOKEN=... 或写入 /tmp/gh_token"; exit 1; }
  TOKEN=$(cat /tmp/gh_token)
fi

api() { curl -s --max-time 30 -H "Authorization: Bearer $TOKEN" \
              -H "Accept: application/vnd.github+json" "$@"; }

LOGIN=$(api https://api.github.com/user | python3 -c "import sys,json;print(json.load(sys.stdin)['login'])")
echo "✓ 已认证为：$LOGIN"

if [ "$VIS" = "public" ]; then PRIVATE=False; else PRIVATE=True; fi

# 仓库已存在则跳过创建
CODE=$(curl -s -o /dev/null -w "%{http_code}" --max-time 20 \
       -H "Authorization: Bearer $TOKEN" "https://api.github.com/repos/$LOGIN/$REPO")
if [ "$CODE" = "200" ]; then
  echo "• 仓库 $LOGIN/$REPO 已存在，直接推送"
else
  R=$(api -X POST https://api.github.com/user/repos \
       -d "$(python3 -c "
import json,sys
print(json.dumps({
 'name':'''$REPO''',
 'description':'把 pptx 转成单文件网页课件，内置多班级教学断点标注系统',
 'private':$PRIVATE,'auto_init':False}))")")
  echo "$R" | python3 -c "
import sys,json
d=json.load(sys.stdin)
if 'full_name' in d: print('✓ 仓库已创建：'+d['full_name']+'  ('+(  'private' if d['private'] else 'public')+')')
else: print('✗ 创建失败：'+str(d.get('message'))); sys.exit(1)"
fi

cd "$(dirname "$0")"
git remote remove origin 2>/dev/null || true
git remote add origin "https://oauth2:${TOKEN}@github.com/${LOGIN}/${REPO}.git"
git branch -M main
GIT_TERMINAL_PROMPT=0 git push -u origin main

# 推送后移除明文 token 的 remote，改用 credential 形式
git remote set-url origin "https://github.com/${LOGIN}/${REPO}.git"
echo ""
echo "✓ 推送完成：https://github.com/${LOGIN}/${REPO}"
