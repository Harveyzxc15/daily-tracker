#!/bin/bash
# 一鍵安裝 — 在 Finder 中雙擊此檔案即可執行
DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$DIR"

echo "======================================"
echo "  每日追蹤主機項目 — 安裝程式"
echo "======================================"
echo ""

# 1. 安裝 Python 套件
echo "▶ 安裝 Python 套件..."
pip3 install -r requirements.txt -q
if [ $? -ne 0 ]; then
    osascript -e 'display alert "安裝失敗" message "pip 安裝失敗，請確認已安裝 Python 3"'
    exit 1
fi
echo "  ✅ Python 套件安裝完成"
echo ""

# 2. 確認 Java
JAVA="/Library/Java/JavaVirtualMachines/jdk1.8.0_251.jdk/Contents/Home/bin/java"
echo "▶ 確認 Java..."
if [ -f "$JAVA" ]; then
    echo "  ✅ Java 已就緒"
else
    echo "  ⚠️  找不到 Java（$JAVA）"
    echo "     請開啟 config.py 修改 JAVA 路徑"
fi
echo ""

# 3. 確認 EPBrowser（EPB 連線 lib，公司內部程式）
echo "▶ 確認 EPBrowser..."
if [ -f "/Library/EPBrowser/EPB/Shell/shell.jar" ]; then
    echo "  ✅ EPBrowser 已就緒（EPB 橋接由本工具自動編譯，不需另放資料夾）"
else
    echo "  ⚠️  找不到 EPBrowser（/Library/EPBrowser）"
    echo "     EPB 查詢需要公司的 EPBrowser 程式，請先安裝"
fi
echo ""

# 4. 建立輸出目錄
echo "▶ 建立輸出目錄..."
mkdir -p "$HOME/daily-tracker-output/北二區"
mkdir -p "$HOME/daily-tracker-output/北一區"
echo "  ✅ 輸出目錄已建立"
echo ""

# 5. 設定開機自動啟動選單列 App
echo "▶ 設定開機自動啟動..."
PLIST="$HOME/Library/LaunchAgents/com.dailytracker.menubar.plist"
PYTHON=$(which python3)
cat > "$PLIST" << PLIST_EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>com.dailytracker.menubar</string>
    <key>ProgramArguments</key>
    <array>
        <string>$PYTHON</string>
        <string>$DIR/menubar.py</string>
    </array>
    <key>RunAtLoad</key>
    <true/>
    <key>KeepAlive</key>
    <dict>
        <key>SuccessfulExit</key>
        <false/>
    </dict>
    <key>ThrottleInterval</key>
    <integer>10</integer>
    <key>StandardOutPath</key>
    <string>/tmp/dailytracker_menubar.log</string>
    <key>StandardErrorPath</key>
    <string>/tmp/dailytracker_menubar.log</string>
</dict>
</plist>
PLIST_EOF
launchctl unload "$PLIST" 2>/dev/null
launchctl load "$PLIST"
echo "  ✅ 開機自動啟動已設定"
echo ""

# 6. 立刻啟動選單列 App
echo "▶ 啟動選單列..."
pkill -f "menubar.py" 2>/dev/null
sleep 0.5
nohup python3 "$DIR/menubar.py" > /tmp/dailytracker_menubar.log 2>&1 &
sleep 2
if pgrep -f "menubar.py" > /dev/null; then
    echo "  ✅ 選單列已啟動（右上角 📊）"
else
    echo "  ⚠️  選單列啟動失敗，請手動執行：python3 $DIR/menubar.py"
fi
echo ""

echo "======================================"
echo "  安裝完成！"
echo "  右上角 📊 就是你的工具選單"
echo "======================================"
echo ""
read -p "按 Enter 關閉..."
