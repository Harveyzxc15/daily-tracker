#!/bin/bash
# 每日追蹤主機項目 — 啟動器（雙擊執行）
cd "$(dirname "$0")"

# 收集所有工具（# TOOL: 開頭）
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
    osascript -e 'display alert "找不到可執行的工具" message "請確認資料夾內有 .py 工具檔案"'
    exit 1
fi

# 組成 AppleScript list 字串
LIST=""
for name in "${NAMES[@]}"; do
    LIST="${LIST}\"${name}\", "
done
LIST="${LIST%,*}"  # 移除最後一個逗號

# 跳出點選視窗
CHOICE=$(osascript <<EOF
choose from list {$LIST} with title "每日追蹤主機項目" with prompt "請選擇要執行的報表：" OK button name "執行" cancel button name "取消"
EOF
)

[ "$CHOICE" = "false" ] && exit 0

# 找到對應腳本
SCRIPT=""
for i in "${!NAMES[@]}"; do
    if [ "${NAMES[$i]}" = "$CHOICE" ]; then
        SCRIPT="${FILES[$i]}"
        break
    fi
done

# 詢問日期（可留空）
DATERANGE=$(osascript <<EOF
set d to text returned of (display dialog "輸入日期區間（留空 = 自動最近完整週）" default answer "" with title "$CHOICE" buttons {"取消", "執行"} default button "執行")
return d
EOF
)

[ $? -ne 0 ] && exit 0

echo "▶ 執行：$CHOICE"
echo ""

if [ -z "$DATERANGE" ]; then
    python3 "$SCRIPT"
else
    python3 "$SCRIPT" $DATERANGE
fi

read -p "按 Enter 關閉..."
