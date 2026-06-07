"""
config.py — 請依照自己的環境修改此檔案後再執行報表腳本。
"""
import subprocess
from pathlib import Path

# ── EPB 連線設定 ──────────────────────────────────────────────────────────────
# Java 執行檔路徑（確認版本與路徑是否正確）
JAVA  = "/Library/Java/JavaVirtualMachines/jdk1.8.0_251.jdk/Contents/Home/bin/java"
JAVAC = JAVA.replace("/bin/java", "/bin/javac")

# EPB 橋接：直接用 daily-tracker 資料夾本身（內含 EPBReportQuery.java，自動編譯）
# 只需電腦有安裝 EPBrowser（/Library/EPBrowser）即可，不必另外放「北一區週報-app」
EPB_APP_DIR = Path(__file__).resolve().parent
EPB_CP  = f"{EPB_APP_DIR}:/Library/EPBrowser/EPB/Shell/shell.jar:/Library/EPBrowser/EPB/Shell/lib/*"
EPB_CWD = str(EPB_APP_DIR)


def _ensure_epb_compiled():
    """EPBReportQuery.java → .class（缺檔或原始碼較新時自動編譯）。"""
    src = EPB_APP_DIR / "EPBReportQuery.java"
    cls = EPB_APP_DIR / "EPBReportQuery.class"
    if src.exists() and (not cls.exists() or cls.stat().st_mtime < src.stat().st_mtime):
        subprocess.run([JAVAC, "-cp", EPB_CP, str(src)],
                       cwd=str(EPB_APP_DIR), capture_output=True, text=True)


_ensure_epb_compiled()

# ── Apple Mail ────────────────────────────────────────────────────────────────
# Mail 本機儲存路徑（macOS 預設，一般不需更改）
MAIL_BASE = Path("~/Library/Mail/V10").expanduser()

# ── 輸出目錄 ──────────────────────────────────────────────────────────────────
# 報表 Excel 輸出的根目錄
OUTPUT_BASE = Path("~/daily-tracker-output").expanduser()
