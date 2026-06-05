#!/usr/bin/env python3
"""
每日追蹤主機項目 — macOS 選單列
右上角圖示點一下 → 選擇報表執行
"""
import subprocess, glob, os, sys, rumps
from datetime import date, timedelta
from pathlib import Path
import threading

BASE_DIR = Path(__file__).parent
OUT_DIR  = Path("~/daily-tracker-output").expanduser()

# ── 工具函式 ──────────────────────────────────────────────────────────────────
def last_run(pattern: str, subdir: str) -> str:
    files = sorted(glob.glob(str(OUT_DIR / subdir / pattern)),
                   key=os.path.getmtime, reverse=True)
    if not files:
        return "尚未執行"
    from datetime import datetime
    return datetime.fromtimestamp(os.path.getmtime(files[0])).strftime("%-m/%-d %H:%M")

def ask_date_range(title: str):
    script = f'''
    tell application "System Events"
        set r to display dialog ¬
            "輸入起始日與截止日（格式：5/31 6/4）" ¬
            with title "{title}" ¬
            default answer "" ¬
            buttons {{"取消", "執行"}} ¬
            default button "執行"
        return text returned of r
    end tell
    '''
    r = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
    if r.returncode != 0:
        return None
    parts = r.stdout.strip().split()
    if len(parts) == 2:
        return parts[0], parts[1]
    if len(parts) == 1:
        return parts[0], parts[0]
    return None

def run_report(script: str, date_arg: str = "") -> None:
    cmd = f'cd "{BASE_DIR}" && python3 "{script}"'
    if date_arg:
        cmd += f" {date_arg}"
    subprocess.run(["osascript", "-e", f'''
    tell application "Terminal"
        activate
        do script "{cmd}"
    end tell
    '''])

def discover_tools():
    """自動偵測資料夾內有 # TOOL: 標記的 .py 腳本"""
    tools = []
    for f in sorted(BASE_DIR.glob("*.py")):
        if f.name == "menubar.py": continue
        with open(f, encoding="utf-8", errors="ignore") as fh:
            for line in fh:
                if line.startswith("# TOOL:"):
                    tools.append((line[7:].strip(), str(f)))
                    break
    return tools

# ── 選單列 App ────────────────────────────────────────────────────────────────
class DailyTrackerApp(rumps.App):
    def __init__(self):
        super().__init__("📊", quit_button=None)
        self._build_menu()

    def _build_menu(self):
        self.menu.clear()
        tools = discover_tools()
        items = []

        for tool_name, script_path in tools:
            # 從工具名稱推測輸出子目錄（北一區 or 北二區）
            subdir = "北二區" if "北二區" in tool_name else ("北一區" if "北一區" in tool_name else "")
            pattern = "保固搭售率北二區_*.xlsx" if "北二區" in tool_name else "保固搭售率_*.xlsx"

            item = rumps.MenuItem(f"🛡️  {tool_name}")
            if subdir:
                item.add(rumps.MenuItem(f"上次執行：{last_run(pattern, subdir)}"))
                item.add(None)
            item.add(rumps.MenuItem("自動（最近完整週）",
                                    callback=lambda _, s=script_path: run_report(s)))
            item.add(rumps.MenuItem("指定區間…",
                                    callback=lambda _, s=script_path, n=tool_name: self._custom(s, n)))
            items.append(item)

        self.menu = items + [
            None,
            rumps.MenuItem("🔄  更新工具", callback=self._update),
            None,
            rumps.MenuItem("結束", callback=lambda _: rumps.quit_application()),
        ]

    def _custom(self, script_path, tool_name):
        result = ask_date_range(tool_name)
        if result:
            run_report(script_path, f"{result[0]} {result[1]}")

    def _update(self, _):
        rumps.notification("每日追蹤工具", "", "更新中，請稍候…", sound=False)
        def do_update():
            try:
                subprocess.run(["git", "-C", str(BASE_DIR), "pull", "--rebase"],
                               capture_output=True, check=True)
                subprocess.run([sys.executable, "-m", "pip", "install", "-r",
                                str(BASE_DIR / "requirements.txt"), "-q"],
                               capture_output=True, check=True)
                rumps.notification("每日追蹤工具", "", "更新完成！重新啟動選單中…", sound=False)
                # 重新啟動自己
                subprocess.Popen([sys.executable, str(BASE_DIR / "menubar.py")])
                rumps.quit_application()
            except subprocess.CalledProcessError as e:
                rumps.notification("每日追蹤工具", "更新失敗", str(e), sound=False)
        threading.Thread(target=do_update, daemon=True).start()

    @rumps.clicked("結束")
    def quit(self, _):
        rumps.quit_application()


if __name__ == "__main__":
    DailyTrackerApp().run()
