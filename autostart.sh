#!/bin/bash
# =============================================================
# COA Converter Web — 开机自启配置
# 用途：配置 macOS 开机自动启动 Web 服务
# =============================================================

GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
BLUE='\033[0;34m'
NC='\033[0m'

ok()   { printf "  ${GREEN}✓${NC} %s\n" "$1"; }
fail() { printf "  ${RED}✗${NC} %s\n" "$1"; }

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PLIST_NAME="com.coa-converter-web.server"
PLIST_PATH="$HOME/Library/LaunchAgents/${PLIST_NAME}.plist"

# 确定 python3 路径（优先使用 venv）
if [ -f "$SCRIPT_DIR/.venv/bin/python3" ]; then
    PYTHON_PATH="$SCRIPT_DIR/.venv/bin/python3"
else
    PYTHON_PATH="$(which python3)"
fi

usage() {
    echo ""
    printf "${BLUE}COA Converter Web 开机自启管理${NC}\n"
    echo ""
    echo "用法："
    echo "  bash autostart.sh enable    开启开机自启"
    echo "  bash autostart.sh disable   关闭开机自启"
    echo "  bash autostart.sh status    查看当前状态"
    echo ""
}

do_enable() {
    echo ""
    printf "${BLUE}配置开机自启...${NC}\n"

    mkdir -p "$HOME/Library/LaunchAgents"

    cat > "$PLIST_PATH" << PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>${PLIST_NAME}</string>
    <key>ProgramArguments</key>
    <array>
        <string>${PYTHON_PATH}</string>
        <string>${SCRIPT_DIR}/app.py</string>
    </array>
    <key>WorkingDirectory</key>
    <string>${SCRIPT_DIR}</string>
    <key>RunAtLoad</key>
    <true/>
    <key>KeepAlive</key>
    <true/>
    <key>StandardOutPath</key>
    <string>${SCRIPT_DIR}/server.log</string>
    <key>StandardErrorPath</key>
    <string>${SCRIPT_DIR}/server.log</string>
</dict>
</plist>
PLIST

    # 加载服务（同时立即启动）
    launchctl unload "$PLIST_PATH" 2>/dev/null
    launchctl load "$PLIST_PATH"

    ok "开机自启已开启"
    ok "服务已启动"
    echo ""
    LOCAL_IP=$(ipconfig getifaddr en0 2>/dev/null || ipconfig getifaddr en1 2>/dev/null || echo "127.0.0.1")
    printf "  访问地址: ${BLUE}http://%s:5050${NC}\n" "$LOCAL_IP"
    echo "  日志文件: $SCRIPT_DIR/server.log"
    echo ""
    echo "  电脑每次开机后服务会自动运行，无需手动操作。"
    echo ""
}

do_disable() {
    echo ""
    if [ -f "$PLIST_PATH" ]; then
        launchctl unload "$PLIST_PATH" 2>/dev/null
        rm -f "$PLIST_PATH"
        ok "开机自启已关闭，服务已停止"
    else
        ok "开机自启未配置，无需操作"
    fi
    echo ""
}

do_status() {
    echo ""
    if [ -f "$PLIST_PATH" ]; then
        if launchctl list | grep -q "$PLIST_NAME"; then
            ok "开机自启: 已开启，服务正在运行"
            LOCAL_IP=$(ipconfig getifaddr en0 2>/dev/null || ipconfig getifaddr en1 2>/dev/null || echo "127.0.0.1")
            printf "  访问地址: ${BLUE}http://%s:5050${NC}\n" "$LOCAL_IP"
        else
            printf "  ${YELLOW}⚠${NC} 开机自启: 已配置，但服务未运行\n"
            echo "  尝试重新启动: launchctl load $PLIST_PATH"
        fi
    else
        echo "  开机自启: 未配置"
        echo "  开启方法: bash autostart.sh enable"
    fi
    echo ""
}

case "${1:-}" in
    enable)  do_enable ;;
    disable) do_disable ;;
    status)  do_status ;;
    *)       usage ;;
esac
