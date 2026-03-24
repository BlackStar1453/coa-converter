#!/bin/bash
# =============================================================
# COA Converter Web 一键安装脚本
# 用途：clone 仓库后运行此脚本即可完成全部配置
# 全自动：自动安装 Homebrew、Python 3（如缺失）
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
SKILL_FIX_SRC="$SCRIPT_DIR/.claude/commands/coa-fix-output.md"
SKILL_DST="$HOME/.claude/commands"

# ---- Step 1: 检查 / 安装 Homebrew（仅 macOS）----
step "[1/5]" "检查 Homebrew..."

if [ "$(uname)" = "Darwin" ]; then
    if command -v brew &> /dev/null; then
        ok "Homebrew 已安装"
    else
        echo "  未找到 Homebrew，正在自动安装..."
        echo "  （安装过程中可能需要输入电脑密码）"
        /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"

        # Homebrew 安装后需要加入 PATH（Apple Silicon vs Intel 路径不同）
        if [ -f "/opt/homebrew/bin/brew" ]; then
            eval "$(/opt/homebrew/bin/brew shellenv)"
        elif [ -f "/usr/local/bin/brew" ]; then
            eval "$(/usr/local/bin/brew shellenv)"
        fi

        if command -v brew &> /dev/null; then
            ok "Homebrew 安装完成"
        else
            fail "Homebrew 安装失败，请手动安装: https://brew.sh"
            exit 1
        fi
    fi
else
    ok "非 macOS 系统，跳过 Homebrew"
fi

# ---- Step 2: 检查 / 安装 Python 3 ----
step "[2/5]" "检查 Python 环境..."

NEED_PYTHON=false

if ! command -v python3 &> /dev/null; then
    NEED_PYTHON=true
else
    # 验证 python3 是否完整可用（Xcode CLT 的 stub 会弹窗而非执行）
    PYTHON_VERSION=$(python3 --version 2>&1) || NEED_PYTHON=true
    if echo "$PYTHON_VERSION" | grep -q "No developer tools"; then
        NEED_PYTHON=true
    fi
    # 检查 venv 模块是否可用
    if [ "$NEED_PYTHON" = false ] && ! python3 -c "import venv" 2>/dev/null; then
        NEED_PYTHON=true
    fi
fi

if [ "$NEED_PYTHON" = true ]; then
    echo "  未找到可用的 Python 3，正在自动安装..."
    if [ "$(uname)" = "Darwin" ]; then
        # macOS：通过 Homebrew 安装
        if command -v brew &> /dev/null; then
            brew install python3
        else
            fail "无法自动安装 Python 3（Homebrew 不可用）"
            echo "  请手动下载安装: https://www.python.org/downloads/"
            exit 1
        fi
    else
        # Linux：尝试 apt 或 yum
        if command -v apt-get &> /dev/null; then
            sudo apt-get update -qq && sudo apt-get install -y -qq python3 python3-venv python3-pip
        elif command -v yum &> /dev/null; then
            sudo yum install -y python3 python3-pip
        else
            fail "无法自动安装 Python 3，请手动安装"
            exit 1
        fi
    fi

    # 验证安装结果
    if ! command -v python3 &> /dev/null; then
        fail "Python 3 安装失败"
        echo "  请手动下载安装: https://www.python.org/downloads/"
        exit 1
    fi
fi

PYTHON_VERSION=$(python3 --version 2>&1)
ok "$PYTHON_VERSION"

# 最终确认 venv 可用
if ! python3 -c "import venv" 2>/dev/null; then
    fail "Python 3 的 venv 模块不可用"
    echo "  macOS: brew reinstall python3"
    echo "  Linux: sudo apt install python3-venv"
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

# ---- Step 5: 安装 Claude Code Skills ----
step "[5/5]" "安装 Claude Code Skills..."

mkdir -p "$SKILL_DST"

# 安装 coa-to-template skill
if [ -d "$SKILL_SRC" ] && [ -f "$SKILL_SRC/SKILL.md" ]; then
    # 备份旧版
    if [ -f "$SKILL_DST/coa-to-template.md" ]; then
        mv "$SKILL_DST/coa-to-template.md" "$SKILL_DST/coa-to-template.md.bak"
    fi
    rm -rf "$SKILL_DST/coa-to-template"
    cp -R "$SKILL_SRC" "$SKILL_DST/coa-to-template"
    ok "Skill coa-to-template 已安装"
else
    warn "未找到 coa-to-template Skill 文件，跳过"
fi

# 安装 coa-fix-output skill
if [ -f "$SKILL_FIX_SRC" ]; then
    cp "$SKILL_FIX_SRC" "$SKILL_DST/coa-fix-output.md"
    ok "Skill coa-fix-output 已安装"
else
    warn "未找到 coa-fix-output Skill 文件，跳过"
fi

# ---- 完成 ----
echo ""
printf "${GREEN}╔══════════════════════════════════════════╗${NC}\n"
printf "${GREEN}║   安装完成！                              ║${NC}\n"
printf "${GREEN}╚══════════════════════════════════════════╝${NC}\n"
echo ""
echo "启动 Web 服务："
echo "  bash run.sh"
echo ""
echo "Claude Code Skills："
echo "  /coa-to-template  — COA PDF 转模板"
echo "  /coa-fix-output   — 修复转换输出错误"
echo ""
printf "项目位置: ${BLUE}%s${NC}\n" "$SCRIPT_DIR"

# 显示局域网 IP
LOCAL_IP=$(ipconfig getifaddr en0 2>/dev/null || ipconfig getifaddr en1 2>/dev/null || hostname -I 2>/dev/null | awk '{print $1}' || echo "127.0.0.1")
printf "Web 地址: ${BLUE}http://%s:5050${NC}\n" "$LOCAL_IP"
echo ""
