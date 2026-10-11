#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
預約取貨核對 × Google 試算表（Mac mini launchd 定時跑，需 VPN）

1. 讀線上表「待處理清單」「區經理參考」店長填的 V 或略過原因（只寫原因也算），收進略過清單；
   「已略過」分頁補寫的原因會更新，把 V 清掉＝取消略過
2. 重跑 預約取貨核對.py
3. 把新報表覆蓋回同一份 Google 試算表（檔案不變，連結不變）

用法：python3 預約取貨核對_Google.py            # 收 V → 重跑 → 上傳
      python3 預約取貨核對_Google.py --harvest  # 只收 V，不重跑
"""
import json
import re
import sys
import traceback
from datetime import datetime
from pathlib import Path

import gspread
import requests
from google.auth.transport.requests import AuthorizedSession
from google.oauth2.service_account import Credentials

import 預約取貨核對 as core

SHEET_ID = (Path.home() / ".config/studioa/resv_pickup_sheet_id.txt").read_text().strip()  # 表是知道連結就能編輯，ID 不放公開 repo
KEY = Path.home() / ".config/studioa/forecast-service-account.json"
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
WEBHOOK_FILE = Path.home() / ".config/studioa/resv_pickup_discord_webhook.txt"
STATE_FILE = Path.home() / ".config/studioa/resv_pickup_state.json"  # 記上次是否失敗，恢復時才發「已恢復」


def connect():
    cred = Credentials.from_service_account_file(KEY, scopes=["https://www.googleapis.com/auth/spreadsheets",
                                                             "https://www.googleapis.com/auth/drive"])
    return gspread.authorize(cred).open_by_key(SHEET_ID), AuthorizedSession(cred)


NEED = ("類型", "門市", "存貨代碼", "預約單號", "台數", "略過(填V)", "略過原因")


def _tab(tabs, title):
    """讀分頁；分頁被刪或表頭被改就跳過，不要讓整個排程掛掉。"""
    rows = tabs[title].get_all_values() if title in tabs else []
    col = {n: i for i, n in enumerate(rows[0])} if rows else {}
    if rows and not all(k in col for k in NEED):
        print(f"⚠️「{title}」表頭被改過，這次不收")
        return [], {}
    return rows, col


def harvest_google(sh):
    ignores = core.load_ignores()
    tabs = {ws.title: ws for ws in sh.worksheets()}
    changed = 0
    for title in ("待處理清單", "區經理參考"):
        rows, col = _tab(tabs, title)
        if rows:
            changed += core._harvest_rows(rows, col, ignores)
    # 「已略過」分頁：補寫原因 → 更新原因；把 V 清掉 → 取消略過（下次回到待處理清單）
    rows, col = _tab(tabs, "已略過")
    for r in rows[1:]:
        key = core.ignore_key(r[col["類型"]], r[col["門市"]], r[col["存貨代碼"]], r[col["預約單號"]])
        if key not in ignores:
            continue
        if r[col["略過(填V)"]].strip() not in core.SKIP_MARKS:
            del ignores[key]
            changed += 1
            continue
        reason = r[col["略過原因"]].strip()
        if reason and reason != ignores[key][1]:
            ignores[key] = (ignores[key][0], reason)
            changed += 1
    if changed:
        core.save_ignores(ignores)
    return changed


def upload(session, xlsx):
    """用 xlsx 取代線上試算表內容（同一個檔案 ID，格式／公式跟本機報表一樣）。"""
    r = session.patch(f"https://www.googleapis.com/upload/drive/v3/files/{SHEET_ID}",
                      params={"uploadType": "media"}, data=xlsx.read_bytes(), headers={"Content-Type": XLSX},
                      timeout=300)
    r.raise_for_status()
    # 檔名放更新時間：每個分頁最上方都看得到，店長不用切到摘要
    name = f"iPhone 18 預約取貨核對（{datetime.now():%m/%d %H:%M} 更新・下次 {core.next_update()}）"
    session.patch(f"https://www.googleapis.com/drive/v3/files/{SHEET_ID}", json={"name": name},
                  timeout=60).raise_for_status()


def main():
    sh, session = connect()
    n = harvest_google(sh)
    print(f"{datetime.now():%Y-%m-%d %H:%M} 從 Google 表收到 {n} 筆略過／原因更新")
    if "--harvest" in sys.argv:
        return
    out = core.OUT_DIR / f"iPhone18預約取貨核對_{datetime.now():%Y%m%d}.xlsx"
    sys.argv = [sys.argv[0], str(out)]
    core.main()
    # 跑一次約 1～2 分鐘，這段期間店長新打的 V 要先收進來再覆蓋，否則會被洗掉；重跑期間又有新的就再收（最多 3 輪）
    for _ in range(3):
        n = harvest_google(sh)
        if not n:
            break
        print(f"跑報表期間又收到 {n} 筆，重新產出")
        core.main()
    upload(session, out)
    print(f"已更新 https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit")


def notify(text):
    try:
        requests.post(WEBHOOK_FILE.read_text().strip(), json={"content": text[:1900]}, timeout=30)
    except Exception as e:  # 通知失敗不能蓋掉原本的錯誤
        print(f"Discord 通知失敗：{e}")


def mark_failed():
    """檔名標「更新失敗」並把「下次」往後推：店長看得到，筆電監看也不會再重複通知。"""
    try:
        _, session = connect()
        url = f"https://www.googleapis.com/drive/v3/files/{SHEET_ID}"
        name = session.get(url, params={"fields": "name"}, timeout=60).json()["name"].split("⚠️")[0]
        t = core.next_update_dt()  # 失敗日可能跟上次成功不同天，日期寫明，監看才不會算錯
        name = re.sub(r"下次 [^）]+）", f"下次 {t:%m/%d} {t.hour}:00）", name)
        session.patch(url, json={"name": f"{name}⚠️ {datetime.now():%H:%M} 更新失敗"}, timeout=60)
    except Exception as e:
        print(f"檔名標失敗也失敗：{e}")


def hint(err):
    if "EPB" in err:
        return "EPB 連不到 → 多半是 Mac mini 的 VPN 斷了"
    if "googleapis" in err or "gspread" in err:
        return "Google 試算表讀寫失敗"
    if "Authorization" in err or "401" in err or "reservation" in err:
        return "預約系統 API 失敗（token 過期或系統維護）"
    return "未預期的錯誤"


def run():
    state = json.loads(STATE_FILE.read_text()) if STATE_FILE.exists() else {}
    try:
        main()
    except BaseException as e:
        if isinstance(e, KeyboardInterrupt):
            raise
        err = str(e) if isinstance(e, SystemExit) else traceback.format_exc()
        print(err)
        notify(f"⚠️ 預約取貨核對 {datetime.now():%m/%d %H:%M} 自動更新失敗：{hint(err)}\n"
               f"Google 表維持上次的內容；店長已打的 V 不會遺失，下次成功會一起收。\n```{err[-1200:]}```")
        mark_failed()
        STATE_FILE.write_text(json.dumps({"failed_since": state.get("failed_since") or f"{datetime.now():%m/%d %H:%M}"}))
        sys.exit(1)
    if "--harvest" not in sys.argv:  # 只收 V 不算「更新成功」
        if state.get("failed_since"):
            notify(f"✅ 預約取貨核對已恢復（{state['failed_since']} 起失敗，{datetime.now():%m/%d %H:%M} 更新成功）")
        STATE_FILE.write_text("{}")


if __name__ == "__main__":
    run()
