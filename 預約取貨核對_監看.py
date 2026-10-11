#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
預約取貨核對 Google 表監看（筆電 launchd 每 30 分鐘）：
Mac mini 整台沒在跑（停電、更新重開機卡在 FileVault）時，mini 自己發不出通知，由筆電發 Discord。
判斷方式：讀試算表檔名「（MM/DD HH:MM 更新・下次 HH:00）」，過了「下次」30 分鐘還沒更新就通知（同一個時段只發一次），
之後看到有更新再發「已恢復」。
"""
import json
import re
from datetime import datetime, timedelta
from pathlib import Path

import requests
from google.auth.transport.requests import AuthorizedSession
from google.oauth2.service_account import Credentials

SHEET_ID = (Path.home() / ".config/studioa/resv_pickup_sheet_id.txt").read_text().strip()  # 表是知道連結就能編輯，ID 不放公開 repo
KEY = Path.home() / ".config/studioa/forecast-service-account.json"
WEBHOOK_FILE = Path.home() / ".config/studioa/resv_pickup_discord_webhook.txt"
STATE_FILE = Path.home() / ".config/studioa/resv_pickup_watch.json"
GRACE = timedelta(minutes=30)  # 一次約 2 分鐘，失敗重試也夠


def notify(text):
    requests.post(WEBHOOK_FILE.read_text().strip(), json={"content": text}, timeout=30)


def parse(name, now):
    """「下次」三種寫法：15:00（同一天）／明天 11:00／10/12 15:00（mini 失敗時寫明日期）"""
    m = re.search(r"（(\d\d)/(\d\d) (\d\d):(\d\d) 更新・下次 (明天 |(\d\d)/(\d\d) )?(\d+):00）", name)
    if not m:
        return None, None
    mo, d, h, mi, prefix, nmo, nd, nh = m.groups()
    year = lambda mo_, d_: now.year - 1 if datetime(now.year, int(mo_), int(d_)) > now + timedelta(days=1) else now.year  # 跨年
    last = datetime(year(mo, d), int(mo), int(d), int(h), int(mi))
    if nmo:
        nxt = datetime(year(nmo, nd), int(nmo), int(nd), int(nh))
    else:
        nxt = last.replace(hour=int(nh), minute=0) + timedelta(days=1 if prefix else 0)
    return last, nxt


def main():
    now = datetime.now()
    cred = Credentials.from_service_account_file(KEY, scopes=["https://www.googleapis.com/auth/drive"])
    name = AuthorizedSession(cred).get(f"https://www.googleapis.com/drive/v3/files/{SHEET_ID}",
                                       params={"fields": "name"}, timeout=60).json()["name"]
    state = json.loads(STATE_FILE.read_text()) if STATE_FILE.exists() else {}
    last, nxt = parse(name, now)
    if last is None:
        if state.get("alert") != "name":
            notify(f"⚠️ 預約取貨核對：Google 表檔名被改過，讀不到更新時間（現在是「{name}」），監看暫停到下次自動更新")
            STATE_FILE.write_text(json.dumps({"alert": "name"}))
        return
    if now > nxt + GRACE:
        if state.get("alert") != f"{nxt:%m/%d %H:%M}":
            notify(f"⚠️ 預約取貨核對：預定 {nxt:%m/%d %H:%M} 的自動更新沒有跑（Google 表停在 {last:%m/%d %H:%M}）。"
                   f"Mac mini 多半沒在運作（停電／重開機卡在 FileVault 解鎖畫面），請去 mini 輸入密碼。")
            STATE_FILE.write_text(json.dumps({"alert": f"{nxt:%m/%d %H:%M}"}))
    elif state.get("alert") and (state["alert"] == "name" or f"{last:%m/%d %H:%M}" >= state["alert"]):
        notify(f"✅ 預約取貨核對已恢復（{last:%m/%d %H:%M} 更新成功）")
        STATE_FILE.write_text("{}")


if __name__ == "__main__":
    main()
