#!/bin/bash

# ==============================================================================
# BatchMail Release Script
# 用于自动化发布新版本：验证环境、构建项目、创建 Git 标签并推送到远程仓库。
# ==============================================================================

# 发生错误时停止脚本
set -e

# 设置自定义 SSH 密钥以解决 GitHub 权限问题
export GIT_SSH_COMMAND='ssh -i /Users/guojc/Library/Application\ Support/UGit/ssh/ugit-created-ssh-key-donnot-delete-bujiC.local -o IdentitiesOnly=yes'

# --- 颜色定义 ---
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m'

log() {
    echo -e "${BLUE}[RELEASE]${NC} $1"
}

# 提示函数
warn() {
    echo -e "${YELLOW}[WARN]${NC} $1"
}

error() {
    echo -e "${RED}[ERROR]${NC} $1"
    exit 1
}

# 1. 基础检查
log "正在验证发布环境..."

# 检测仓库是否初始化
if [ ! -d ".git" ]; then
    error "当前目录不是一个 Git 仓库。"
fi

# 检测主分支名称 (main 或 master)
if git show-ref --verify --quiet refs/heads/main; then
    MAIN_BRANCH="main"
elif git show-ref --verify --quiet refs/heads/master; then
    MAIN_BRANCH="master"
else
    # 如果本地没有，尝试获取远程主分支名
    MAIN_BRANCH=$(git remote show origin | grep 'HEAD branch' | cut -d' ' -f5)
    [ -z "$MAIN_BRANCH" ] && MAIN_BRANCH="master"
fi

# 检查当前是否在主分支
CURRENT_BRANCH=$(git rev-parse --abbrev-ref HEAD)
if [ "$CURRENT_BRANCH" != "$MAIN_BRANCH" ]; then
    error "当前分支是 $CURRENT_BRANCH。发布操作必须在 $MAIN_BRANCH 分支上执行。"
fi

# 检查工作区是否干净
if [ -n "$(git status --porcelain)" ]; then
    warn "工作区有未提交的更改。"
    read -p "是否继续发布？（可能会包含这些更改）(y/N) " confirm
    if [[ $confirm != [yY] ]]; then
        exit 1
    fi
fi

# 2. 版本获取
# 从 package.json 中提取版本号
VERSION=$(grep '"version":' package.json | head -1 | awk -F: '{ print $2 }' | sed 's/[", ]//g')
log "检测到当前版本: ${GREEN}v$VERSION${NC}"

# 询问是否需要手动指定版本号（默认为当前版本）
read -p "确认发布版本号 [$VERSION]: " INPUT_VERSION
VERSION=${INPUT_VERSION:-$VERSION}

# 3. 执行构建流程
echo "------------------------------------------------"
read -p "是否在发布前执行构建 (yarn build)? (y/N) " RUN_BUILD
if [[ $RUN_BUILD == [yY] ]]; then
    log "正在运行构建脚本..."
    yarn build
else
    warn "跳过构建步骤，请确保代码已就绪。"
fi
echo "------------------------------------------------"

# 4. Git 提交与标签
# 如果 package.json 有变动（例如手动更新了版本号尚未提交），先提交
if git status --porcelain | grep "package.json" > /dev/null; then
    log "提交版本更新..."
    git add package.json
    git commit -m "chore: release v$VERSION"
fi

# 处理已存在的标签
if git rev-parse "v$VERSION" >/dev/null 2>&1; then
    warn "标签 v$VERSION 已存在。正在删除本地旧标签并重新创建..."
    git tag -d "v$VERSION"
fi

log "正在创建 Git 标签: ${GREEN}v$VERSION${NC}"
git tag -a "v$VERSION" -m "Release v$VERSION"

# 5. 推送远程
log "正在推送代码到远程仓库 ($MAIN_BRANCH)..."
git push origin "$MAIN_BRANCH"

log "正在推送标签到远程仓库..."
git push origin --tags

echo "------------------------------------------------"
echo -e "${GREEN}🎉 发布成功: v$VERSION${NC}"
echo "------------------------------------------------"
