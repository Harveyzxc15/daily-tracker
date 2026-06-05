#!/bin/bash
# 每日追蹤主機項目 — 啟動器
# 雙擊此檔案即可選擇要執行的報表
cd "$(dirname "$0")"

echo "======================================"
echo "  每日追蹤主機項目"
echo "======================================"
echo ""

# 自動偵測可執行的工具（以 # TOOL: 開頭的行作為顯示名稱）
declare -a FILES
declare -a NAMES
while IFS= read -r -d '' f; do
    name=$(grep -m1 '^# TOOL:' "$f" | sed 's/^# TOOL: *//')
    if [ -n "$name" ]; then
        FILES+=("$f")
        NAMES+=("$name")
    fi
done < <(find . -maxdepth 1 -name "*.py" -print0 | sort -z)

if [ ${#FILES[@]} -eq 0 ]; then
    echo "找不到可執行的工具"
    read -p "按 Enter 關閉..."
    exit 1
fi

# 顯示選單
for i in "${!NAMES[@]}"; do
    echo "  $((i+1)). ${NAMES[$i]}"
done
echo ""
read -p "請選擇（輸入數字）: " choice

if ! [[ "$choice" =~ ^[0-9]+$ ]] || [ "$choice" -lt 1 ] || [ "$choice" -gt ${#FILES[@]} ]; then
    echo "無效的選擇"
    read -p "按 Enter 關閉..."
    exit 1
fi

SCRIPT="${FILES[$((choice-1))]}"
TOOL_NAME="${NAMES[$((choice-1))]}"

echo ""
echo "▶ 執行：$TOOL_NAME"
echo ""

# 詢問日期區間
read -p "日期區間（直接 Enter = 自動最近完整週，或輸入如 5/31 6/4）: " daterange

echo ""
if [ -z "$daterange" ]; then
    python3 "$SCRIPT"
else
    python3 "$SCRIPT" $daterange
fi

echo ""
read -p "按 Enter 關閉..."
