#!/bin/bash
# COA Converter Web — Start Script
cd "$(dirname "$0")"

# Activate venv if exists
if [ -d ".venv" ]; then
    source .venv/bin/activate
fi

echo "Starting COA Converter Web..."

# Check if autostart is configured; if not, offer to enable it
PLIST_PATH="$HOME/Library/LaunchAgents/com.coa-converter-web.server.plist"
if [ ! -f "$PLIST_PATH" ]; then
    echo ""
    echo "提示: 尚未配置开机自启。"
    printf "是否开启开机自动启动? (y/N): "
    read -r REPLY
    if [[ "$REPLY" =~ ^[Yy]$ ]]; then
        bash "$(dirname "$0")/autostart.sh" enable
    fi
    echo ""
fi

# Get local IP for LAN access
LOCAL_IP=$(ipconfig getifaddr en0 2>/dev/null || ipconfig getifaddr en1 2>/dev/null || echo "127.0.0.1")
echo "Local:   http://127.0.0.1:5050"
echo "LAN:     http://${LOCAL_IP}:5050"

# Open browser after a short delay
(sleep 1 && open "http://127.0.0.1:5050") &

python3 app.py
