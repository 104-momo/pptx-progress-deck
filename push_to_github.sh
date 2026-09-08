#!/usr/bin/env bash
# 把这个仓库推到你自己的 GitHub 账号。
# 在你的电脑上执行（沙箱里访问不到 GitHub）。
#
# 前置：安装 GitHub CLI 并登录   https://cli.github.com/
#         gh auth login
#
# 用法：  ./push_to_github.sh [仓库名]    默认仓库名 pptx-progress-deck
set -e
cd "$(dirname "$0")"

REPO="${1:-pptx-progress-deck}"

if ! command -v gh >/dev/null 2>&1; then
  echo "没装 GitHub CLI。手动方式："
  echo "  1. 在 https://github.com/new 建一个空仓库（不要勾 README）"
  echo "  2. git remote add origin https://github.com/<你的用户名>/$REPO.git"
  echo "  3. git branch -M main && git push -u origin main"
  exit 1
fi

# 建仓库并推送（已存在则只推送）
gh repo create "$REPO" \
  --public \
  --description "把 PowerPoint 转成带多班断点标注的单文件网页课件" \
  --source=. \
  --push 2>/dev/null || {
    echo "仓库可能已存在，改为直接推送…"
    git remote remove origin 2>/dev/null || true
    gh repo create "$REPO" --public --source=. --push --confirm 2>/dev/null || true
  }

echo
echo "完成。仓库地址："
gh repo view "$REPO" --web 2>/dev/null || echo "https://github.com/$(gh api user --jq .login)/$REPO"
