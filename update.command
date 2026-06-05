#!/bin/bash
# 更新工具 — 從 GitHub 拉取最新版本
cd "$(dirname "$0")"

echo "======================================"
echo "  更新每日追蹤主機項目"
echo "======================================"
echo ""

echo "▶ 拉取最新版本..."
git pull
if [ $? -ne 0 ]; then
    echo "❌ 更新失敗，請確認網路連線"
    read -p "按 Enter 關閉..."
    exit 1
fi
echo ""

echo "▶ 更新 Python 套件..."
pip3 install -r requirements.txt -q
echo "  ✅ 套件已更新"
echo ""

echo "======================================"
echo "  更新完成！"
echo "======================================"
echo ""
read -p "按 Enter 關閉..."
