#!/usr/bin/env python3
# TOOL_APP: 📈 北二區週報產生器
"""
北二區週報產生器啟動器（選單列點一下執行）
- 第一次：自動 git clone 北二區週報 + 人流插件到桌面，並引導裝插件
- 之後：git pull 更新 → 啟動本機 server（port 8783）+ 用 Chrome 開網頁
網頁開啟後，選日期、按「產生週報」即可下載 Excel。
"""
import subprocess
from pathlib import Path

REPORT_URL = "https://github.com/Harveyzxc15/north2-weekly-report.git"
REPORT_DIR = Path("~/Desktop/north2-weekly-report").expanduser()
LAUNCHER   = REPORT_DIR / "啟動北二區週報.command"

PLUGIN_URL = "https://github.com/Harveyzxc15/shoppertrak-traffic-plugin.git"
PLUGIN_DIR = Path("~/Desktop/shoppertrak-traffic-plugin").expanduser()


def dialog(msg, title):
    subprocess.run(["osascript", "-e",
        f'display dialog "{msg}" buttons {{"知道了"}} default button 1 with title "{title}"'])


def ensure_report():
    if not REPORT_DIR.exists():
        print("第一次使用，正在下載北二區週報…")
        r = subprocess.run(["git", "clone", REPORT_URL, str(REPORT_DIR)])
        if r.returncode != 0 or not LAUNCHER.exists():
            print("\n✗ 週報下載失敗，請確認網路 / git 是否可用。")
            input("按 Enter 關閉…")
            raise SystemExit(1)
        print("✓ 週報下載完成\n")


def ensure_plugin():
    """下載人流插件；第一次時引導同事在 Chrome 載入（無法全自動，Chrome 限制）。"""
    first_time = not PLUGIN_DIR.exists()
    if first_time:
        print("下載人流插件…")
        if subprocess.run(["git", "clone", PLUGIN_URL, str(PLUGIN_DIR)]).returncode != 0:
            print("⚠ 插件下載失敗（不影響週報，人流可手填）")
            return
    else:
        subprocess.run(["git", "-C", str(PLUGIN_DIR), "pull", "--quiet"])

    if first_time:
        # 開擴充功能頁 + Finder 顯示資料夾 + 步驟提示
        subprocess.run(["open", "-a", "Google Chrome", "chrome://extensions"])
        subprocess.run(["open", "-R", str(PLUGIN_DIR)])
        dialog(
            "人流插件已下載到桌面的 shoppertrak-traffic-plugin 資料夾。\n\n"
            "請在 Chrome 的「擴充功能」頁面（若沒自動開，請在網址列輸入 chrome://extensions ）：\n"
            "1. 開啟右上角的「開發人員模式」\n"
            "2. 點「載入未封裝項目」\n"
            "3. 選擇 Finder 剛跳出的 shoppertrak-traffic-plugin 資料夾\n\n"
            "裝好後登入 ShopperTrak，週報的人流就會自動帶入（沒裝也能用，人流手填）。",
            "安裝人流插件（只需做一次）")


print("=== 北二區週報產生器 ===\n")
ensure_report()
ensure_plugin()

# 啟動（該腳本會：git pull → 清舊 server → 用 Chrome 開 ShopperTrak+週報 → 啟動 server）
print("啟動中…（瀏覽器會自動開啟；若 EPB 尚未測試過，第一次產報請確認 VPN/EPB 連線）\n")
subprocess.run(["bash", str(LAUNCHER)])
