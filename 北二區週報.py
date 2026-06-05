#!/usr/bin/env python3
# TOOL_APP: 📈 北二區週報產生器
"""
北二區週報產生器啟動器（選單列點一下執行）
- 第一次：自動 git clone 北二區週報到桌面
- 之後：git pull 更新 → 啟動本機 server（port 8783）+ 用 Chrome 開網頁
網頁開啟後，選日期、按「產生週報」即可下載 Excel。
"""
import subprocess
from pathlib import Path

REPO_URL = "https://github.com/Harveyzxc15/north2-weekly-report.git"
REPO_DIR = Path("~/Desktop/north2-weekly-report").expanduser()
LAUNCHER = REPO_DIR / "啟動北二區週報.command"

print("=== 北二區週報產生器 ===\n")

# 第一次：自動下載
if not REPO_DIR.exists():
    print("第一次使用，正在下載北二區週報…")
    r = subprocess.run(["git", "clone", REPO_URL, str(REPO_DIR)])
    if r.returncode != 0 or not LAUNCHER.exists():
        print("\n✗ 下載失敗，請確認網路 / git 是否可用。")
        input("按 Enter 關閉…")
        raise SystemExit(1)
    print("✓ 下載完成\n")

# 啟動（該腳本會：git pull → 清舊 server → 用 Chrome 開 ShopperTrak+週報 → 啟動 server）
print("啟動中…（瀏覽器會自動開啟；若 EPB 尚未測試過，第一次產報請確認 VPN/EPB 連線）\n")
subprocess.run(["bash", str(LAUNCHER)])
