#!/bin/bash
# 一鍵安裝 — 在 Finder 中雙擊此檔案即可執行
cd "$(dirname "$0")"

echo "======================================"
echo "  每日追蹤主機項目 — 安裝程式"
echo "======================================"
echo ""

# 1. 安裝 Python 套件
echo "▶ 安裝 Python 套件..."
pip3 install -r requirements.txt -q
if [ $? -ne 0 ]; then
    echo "❌ pip 安裝失敗，請確認已安裝 Python 3"
    read -p "按 Enter 關閉..."
    exit 1
fi
echo "  ✅ Python 套件安裝完成"
echo ""

# 2. 確認 Java
echo "▶ 確認 Java..."
JAVA="/Library/Java/JavaVirtualMachines/jdk1.8.0_251.jdk/Contents/Home/bin/java"
if [ -f "$JAVA" ]; then
    echo "  ✅ Java 已就緒"
else
    echo "  ⚠️  找不到 Java，請確認路徑："
    echo "     $JAVA"
    echo "  若路徑不同，請開啟 config.py 修改 JAVA 變數"
fi
echo ""

# 3. 確認 EPB App 資料夾
EPB_DIR="$HOME/Desktop/北一區週報-app"
echo "▶ 確認 EPB App 資料夾..."
if [ -d "$EPB_DIR" ]; then
    echo "  ✅ EPB App 資料夾已就緒"
else
    echo "  ⚠️  找不到 EPB App 資料夾："
    echo "     $EPB_DIR"
    echo "  請將 EPB App 資料夾放到桌面，或開啟 config.py 修改 EPB_APP_DIR"
fi
echo ""

# 4. 建立輸出目錄
OUT="$HOME/daily-tracker-output/北二區"
echo "▶ 建立輸出目錄..."
mkdir -p "$OUT"
echo "  ✅ 輸出目錄：$OUT"
echo ""

echo "======================================"
echo "  安裝完成！"
echo ""
echo "  執行方式："
echo "  python3 warranty_report_n2.py             # 自動最近完整週"
echo "  python3 warranty_report_n2.py 5/31 6/4   # 指定日期區間"
echo "======================================"
echo ""
read -p "按 Enter 關閉..."
