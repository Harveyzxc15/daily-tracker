"""
config.py — 請依照自己的環境修改此檔案後再執行報表腳本。
"""
from pathlib import Path

# ── EPB 連線設定 ──────────────────────────────────────────────────────────────
# Java 執行檔路徑（確認版本與路徑是否正確）
JAVA = "/Library/Java/JavaVirtualMachines/jdk1.8.0_251.jdk/Contents/Home/bin/java"

# EPB 應用程式資料夾（放在 ~/Desktop/ 下即可，路徑會自動帶入使用者名稱）
EPB_APP_DIR = Path("~/Desktop/北一區週報-app").expanduser()
EPB_CP  = f"{EPB_APP_DIR}:/Library/EPBrowser/EPB/Shell/shell.jar:/Library/EPBrowser/EPB/Shell/lib/*"
EPB_CWD = str(EPB_APP_DIR)

# ── Apple Mail ────────────────────────────────────────────────────────────────
# Mail 本機儲存路徑（macOS 預設，一般不需更改）
MAIL_BASE = Path("~/Library/Mail/V10").expanduser()

# ── 輸出目錄 ──────────────────────────────────────────────────────────────────
# 報表 Excel 輸出的根目錄
OUTPUT_BASE = Path("~/daily-tracker-output").expanduser()
