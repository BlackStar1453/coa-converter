#!/bin/bash
# =============================================================
# COA Converter Web 一键安装脚本
# 用途：clone 仓库后运行此脚本即可完成全部配置
# 兼容全新 Mac（无 Homebrew / 无 Python 3）
# =============================================================

set -e

GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
BLUE='\033[0;34m'
NC='\033[0m'

# macOS bash 3.2 的 echo 不支持 -e，使用 printf 替代
info()  { printf "${BLUE}%s${NC}\n" "$1"; }
ok()    { printf "  ${GREEN}✓${NC} %s\n" "$1"; }
warn()  { printf "  ${YELLOW}⚠${NC} %s\n" "$1"; }
fail()  { printf "  ${RED}✗${NC} %s\n" "$1"; }
step()  { printf "${YELLOW}%s${NC} %s\n" "$1" "$2"; }

echo ""
printf "${BLUE}╔══════════════════════════════════════════╗${NC}\n"
printf "${BLUE}║   COA Converter Web 安装程序              ║${NC}\n"
printf "${BLUE}╚══════════════════════════════════════════╝${NC}\n"
echo ""

# 获取脚本所在目录（即项目根目录）
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
SKILL_SRC="$SCRIPT_DIR/.claude/commands/coa-to-template"
SKILL_DST="$HOME/.claude/commands"

# ---- Step 1: 检查 Git ----
step "[1/5]" "检查 Git..."

if ! command -v git &> /dev/null; then
    warn "未找到 git（你已经 clone 了仓库，说明 git 可用，跳过）"
else
    ok "$(git --version)"
fi

# ---- Step 2: 检查 / 安装 Python 3 ----
step "[2/5]" "检查 Python 环境..."

install_python_guide() {
    echo ""
    fail "未找到 python3，请先安装 Python 3："
    echo ""
    echo "  方法 1（推荐）: 从官网下载安装包"
    echo "    https://www.python.org/downloads/"
    echo "    下载 macOS 安装包 → 双击安装 → 重新运行此脚本"
    echo ""
    echo "  方法 2: 使用 Homebrew（如果已安装 brew）"
    echo "    brew install python3"
    echo ""
    echo "  方法 3: 使用 Xcode Command Line Tools"
    echo "    xcode-select --install"
    echo "    （注意：此方式安装的 Python 可能版本较旧）"
    echo ""
    exit 1
}

if ! command -v python3 &> /dev/null; then
    install_python_guide
fi

# 验证 python3 是否完整可用（Xcode CLT 的 python3 stub 会弹安装窗口）
PYTHON_VERSION=$(python3 --version 2>&1) || install_python_guide
if echo "$PYTHON_VERSION" | grep -q "No developer tools"; then
    fail "检测到 Xcode Command Line Tools 的 Python stub（未真正安装）"
    install_python_guide
fi

ok "$PYTHON_VERSION"

# 检查 venv 模块是否可用
if ! python3 -c "import venv" 2>/dev/null; then
    fail "Python 3 的 venv 模块不可用"
    echo "  请安装完整版 Python 3: https://www.python.org/downloads/"
    exit 1
fi

# ---- Step 3: 创建虚拟环境并安装依赖 ----
step "[3/5]" "配置 Python 虚拟环境..."

VENV_DIR="$SCRIPT_DIR/.venv"
if [ -d "$VENV_DIR" ]; then
    echo "  虚拟环境已存在，更新依赖..."
else
    echo "  创建虚拟环境..."
    python3 -m venv "$VENV_DIR"
fi

source "$VENV_DIR/bin/activate"
pip install --quiet --upgrade pip
pip install --quiet -r "$SCRIPT_DIR/requirements.txt"
deactivate
ok "Python 依赖安装完成"

# ---- Step 4: 创建工作目录 ----
step "[4/5]" "初始化工作目录..."

mkdir -p "$SCRIPT_DIR/input"
mkdir -p "$SCRIPT_DIR/output"
ok "input/ output/ 目录已就绪"

# ---- Step 5: 安装 Claude Code Skill ----
step "[5/5]" "安装 Claude Code Skill..."

if [ -d "$SKILL_SRC" ] && [ -f "$SKILL_SRC/SKILL.md" ]; then
    mkdir -p "$SKILL_DST"
    # 备份旧版 Skill
    if [ -f "$SKILL_DST/coa-to-template.md" ]; then
        mv "$SKILL_DST/coa-to-template.md" "$SKILL_DST/coa-to-template.md.bak"
        warn "旧版 Skill 文件已备份"
    fi
    rm -rf "$SKILL_DST/coa-to-template"
    cp -R "$SKILL_SRC" "$SKILL_DST/coa-to-template"
    ok "Skill 已安装到 $SKILL_DST/coa-to-template/"
else
    warn "未找到 Skill 文件，跳过（不影响 Web 功能）"
fi

# ---- 完成 ----
echo ""
printf "${GREEN}╔══════════════════════════════════════════╗${NC}\n"
printf "${GREEN}║   安装完成！                              ║${NC}\n"
printf "${GREEN}╚══════════════════════════════════════════╝${NC}\n"
echo ""
echo "接下来："
echo "  1. 将模板文件（.xlsx / .docx）放入 $SCRIPT_DIR/templates/"
echo "  2. 启动 Web 服务：bash run.sh"
echo "  3. 或在 Claude Code 中使用 /coa-to-template"
echo ""
printf "项目位置: ${BLUE}%s${NC}\n" "$SCRIPT_DIR"

# 显示局域网 IP
LOCAL_IP=$(ipconfig getifaddr en0 2>/dev/null || ipconfig getifaddr en1 2>/dev/null || hostname -I 2>/dev/null | awk '{print $1}' || echo "127.0.0.1")
printf "Web 地址: ${BLUE}http://%s:5050${NC}\n" "$LOCAL_IP"
echo ""
