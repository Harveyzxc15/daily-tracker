#!/usr/bin/env python3
"""EPB 連線診斷 — 確認連得到、看得到資料、ARpedia(brand 496) 有沒有資料。
用法：python3 ~/Desktop/daily-tracker/epb_check.py
"""
import subprocess
from config import JAVA, EPB_CP, EPB_CWD

def q(sql):
    r = subprocess.run([JAVA, "-cp", EPB_CP, "EPBReportQuery", sql, "100"],
                       capture_output=True, text=True, cwd=EPB_CWD)
    return r.stdout.strip(), r.stderr.strip()

BASE = ("FROM poslinev_bi l WHERE l.org_id='01' "
        "AND l.shop_id IN ('009','025','050','055','063','064','068') "
        "AND l.doc_date>=TO_DATE('2026-06-01','yyyy-mm-dd') "
        "AND l.doc_date<=TO_DATE('2026-06-06','yyyy-mm-dd')")

print("=== EPB 連線診斷（6/1~6/6 北二區）===\n")

out, err = q(f"SELECT COUNT(*) AS N {BASE}")
print("總筆數      :", out or "(空)")
if err:
    print("\n⚠ 錯誤訊息（STDERR）:\n" + err[:600])
    print("\n→ 有錯誤通常是：VPN 沒連 / EPBrowser 設定。把上面這段貼給 Bob。")
    raise SystemExit

out, _ = q(f"SELECT COUNT(*) AS N {BASE} AND l.brand_id='496'")
print("ARpedia496  :", out or "(空)")

out, _ = q(f"SELECT l.brand_id, COUNT(*) AS N {BASE} GROUP BY l.brand_id ORDER BY N DESC")
print("\n前幾大 brand:")
print("\n".join(out.split("\n")[:11]) if out else "(空)")

print("\n（對照 Bob 機器：總筆數約 5036、ARpedia496 約 27）")
print("把整段結果貼給 Bob。")
