#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
iPhone 18 Pro 預約 × EPB POS 結帳 — 取貨狀態核對（北區 13 店），規則依 9/17 預約系統 SOP

每筆結帳依序對到預約：① POS F10「7 預約條碼」(POSLINE.REF7) ② 備註預約單號 ③ 備註會員編號 ④ 結帳會員＋規格
  依使用者 7 點檢核：A 結帳未改、B 改狀態無消費、疑似代取、銷退仍已取貨、掃錯規格、重複掃、放棄未遞補、逾期未取、
  清點單（已到貨＋保留）vs EPB 庫存、會員總預約數核對
另做門市×規格庫存對帳（不靠會員）當總量交叉檢查。

需 VPN + 已安裝 EPBrowser。用法：python3 預約取貨核對.py [輸出.xlsx]
"""
import json
import os
import re
import subprocess
import sys
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from config import JAVA, EPB_CP, EPB_CWD

RESV_CONFIG = Path.home() / "studioa_reservation_config.json"
OUT_DIR = Path(os.environ.get("RESV_OUT_DIR") or Path.home() / "Desktop")  # Mac mini launchd 讀寫不到 Desktop，改用環境變數
IGNORE_FILE = Path(__file__).resolve().parent / "預約取貨核對_略過清單.xlsx"
SKIP_MARKS = {"V", "v", "Ｖ", "ｖ", "✓", "✔", "Y", "y"}
ACTIVITIES = {
    "3a239ce0-4214-a56a-3c82-28788bac5906": "限定場",
    "3a239ce3-1831-cbcc-8661-d9b711d1fa7a": "全會員",
}
SALES_FROM = "2026-09-18"
UPDATE_HOURS = (11, 15, 19, 22)  # Mac mini 自動更新 Google 表的時段（要跟 com.studioa.resv-pickup-google.plist 一致）
HOURS_TXT = "／".join(map(str, UPDATE_HOURS))
SETTLE_MIN = 10  # 結帳滿幾分鐘才判定「漏改取貨」（給門市改狀態、系統同步的時間）

# (門市碼, 預約系統店名, 顯示名, 區)
STORES = [
    ("004", "士林", "士林", "北一"),
    ("005", "微風", "微風", "北一"),
    ("024", "美麗華", "美麗華", "北一"),
    ("046", "阿波羅", "阿波羅", "北一"),
    ("054", "大葉高島屋", "高島屋", "北一"),
    ("057", "羅東", "羅東", "北一"),
    ("009", "永和", "永和", "北二"),
    ("025", "板橋誠品", "板橋誠品", "北二"),
    ("050", "新西門", "新西門", "北二"),
    ("055", "花蓮", "花蓮", "北二"),
    ("063", "板橋遠百", "板橋遠百", "北二"),
    ("064", "新莊宏匯", "新莊宏匯", "北二"),
    ("068", "新店裕隆城", "新店裕隆城", "北二"),
]
BY_RESV_NAME = {s[1]: s for s in STORES}
BY_CODE = {s[0]: s for s in STORES}
PICKED = "已取貨"


def epb(sql, max_rows=20000):
    r = subprocess.run([JAVA, "-cp", EPB_CP, "EPBReportQuery", sql, str(max_rows)],
                       capture_output=True, text=True, cwd=EPB_CWD)
    lines = [l for l in r.stdout.strip().split("\n") if l]
    if not lines:
        raise SystemExit("EPB 查詢失敗（VPN 沒連？）：\n" + r.stderr.strip()[:600])
    if "MAX_ROWS_REACHED" in r.stderr:
        raise SystemExit("EPB 資料被截斷，請調高 max_rows")
    hdr = lines[0].split("\t")
    return [dict(zip(hdr, l.split("\t"))) for l in lines[1:]]


def fetch_reservations():
    cfg = json.loads(RESV_CONFIG.read_text())
    headers = {"Authorization": cfg["token"]}
    act = "&".join(f"ReservationActivityIds={a}" for a in ACTIVITIES)
    items, skip = [], 0
    while True:
        url = (f"{cfg['base_url']}/reservation-activity/reservation-user-list"
               f"?SkipCount={skip}&MaxResultCount=1000&{act}")
        d = requests.get(url, headers=headers, timeout=180).json()["data"]["userReservationListOutDtos"]
        items += d["items"]
        skip += 1000
        if skip >= d["totalCount"]:
            # 抓取期間門市在改狀態，分頁可能錯位 → 依預約明細 id 去重
            return list({it["productOrderProductShelfId"]: it for it in items}.values())


_HOT_SKUS = None
HOT_HISTORY_FILE = Path(__file__).resolve().parent / "預約取貨核對_熱門歷史.json"
_HOT_SPANS = {}   # 規格 → [[列為熱門的時間, 移出的時間或 None], …]（字串 YYYY-MM-DD HH:MM:SS）


def guid_time(g):
    """預約後台的 GUID 是 ABP 循序 GUID：前 12 碼＝UTC 毫秒（自西元 1 年起），＋8 小時＝台灣時間建立時間，準到秒。
       用 shop-transfer-list 的 creationTime 校準過（10/3）。"""
    return datetime(1, 1, 1) + timedelta(milliseconds=int(g.replace("-", "")[:12], 16), hours=8)


def update_hot_history(items):
    """熱門清單會變（9/29 加 07313593／614／616、10/1 加 07313585、拿掉 07313593／594），
       判斷「當時能不能賣現貨」要看事情發生那時候是不是熱門，不能拿今天的清單套過去。
       後台只給得到目前清單＋每筆設定的建立時間（GUID），刪掉的看不到 → 每次執行把清單記進歷史檔：
       新出現的規格從 GUID 時間起算；不見的規格以這次執行時間當結束（實際移出時間落在上次執行到這次之間）。"""
    spans = json.loads(HOT_HISTORY_FILE.read_text()) if HOT_HISTORY_FILE.exists() else {}
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    since = defaultdict(list)
    for x in items:
        if x.get("stK_ID"):
            since[x["stK_ID"]].append(guid_time(x["hotDriveId"]).strftime("%Y-%m-%d %H:%M:%S"))
    for sku, ts in since.items():
        ss = spans.setdefault(sku, [])
        if not ss or ss[-1][1]:
            ss.append([min(ts), None])
            print(f"  熱門清單新增 {sku}（後台設定時間 {min(ts)}）")
    for sku, ss in spans.items():
        if sku not in since and ss and not ss[-1][1]:
            ss[-1][1] = now
            print(f"  ⚠️ {sku} 已不在熱門清單（{now} 發現，實際移出時間在上次執行之後）")
    HOT_HISTORY_FILE.write_text(json.dumps(spans, ensure_ascii=False, indent=1))
    _HOT_SPANS.clear()
    _HOT_SPANS.update(spans)


def hot_at(sku, t):
    """規格在時間 t（datetime 或 YYYYMMDDHHMMSS／YYYY-MM-DD HH:MM:SS 字串）是不是熱門。"""
    if isinstance(t, str):
        t = datetime.strptime(re.sub(r"\D", "", t)[:14].ljust(14, "0"), "%Y%m%d%H%M%S")
    ts = t.strftime("%Y-%m-%d %H:%M:%S")
    return any(a <= ts and (not b or ts < b) for a, b in _HOT_SPANS.get(sku, ()))


def hot_since(sku):
    """目前這段熱門從什麼時候開始（不在熱門清單回 None）。"""
    ss = _HOT_SPANS.get(sku) or []
    return datetime.strptime(ss[-1][0], "%Y-%m-%d %H:%M:%S") if ss and not ss[-1][1] else None


def fetch_hot_skus():
    """熱門機型以預約後台的官方設定為準（reservation-activity/hot-drive-list），不要從預約紀錄的 isHotDrive 反推：
       反推只看得到「北區有人預約過」的規格，而且熱門與否會決定能不能轉現貨、要不要跨店遞補，判錯代價很大。
       （9/23 實測兩者剛好一致：官方 18 個規格＝36 筆設定，規格×限定場/全會員兩場。）"""
    global _HOT_SKUS
    if _HOT_SKUS is None:
        cfg = json.loads(RESV_CONFIG.read_text())
        d = requests.get(f"{cfg['base_url']}/reservation-activity/hot-drive-list?SkipCount=0&MaxResultCount=500",
                         headers={"Authorization": cfg["token"]}, timeout=180).json()["data"]["hotDriveDataListOutDtos"]
        _HOT_SKUS = {x["stK_ID"] for x in d["items"] if x.get("stK_ID")}
        print(f"官方熱門機型 {len(_HOT_SKUS)} 個規格（後台設定 {d['totalCount']} 筆＝規格×場次）")
        update_hot_history(d["items"])
    return _HOT_SKUS


def fetch_transfers():
    """跨店調撥建議（系統每天凌晨 3 點計算）：用來判斷放棄的機器到底有沒有被排進遞補。"""
    cfg = json.loads(RESV_CONFIG.read_text())
    headers = {"Authorization": cfg["token"]}
    out, skip = [], 0
    while True:
        url = (f"{cfg['base_url']}/reservation-activity/shop-transfer-list"
               f"?SkipCount={skip}&MaxResultCount=500")
        d = requests.get(url, headers=headers, timeout=180).json()["data"]["shopTransferDataListOutDtos"]
        out += d["items"]
        skip += 500
        if skip >= d["totalCount"]:
            return out


def doc_time(d):
    """EPB 單號第 3 段是交易時間（7508-050e05-20260918123030-0007），取不到就退回日期。"""
    parts = d["DOC_ID"].split("-")
    return parts[2] if len(parts) >= 4 else d["DOC_DATE"]


def mask_name(n):
    n = (n or "").strip()
    return n[:1] + "○○" if n else ""


def sku_aliases(skus):
    """同一支手機在 EPB 有多組代碼（07313…＝預約系統用、073S13…、905029…＝另一進貨管道；886…＝未開保，排除），依 MODEL 對回預約系統代碼。"""
    in_sku = ",".join(f"'{s}'" for s in sorted(skus))
    model = {x["MODEL"]: x["STK_ID"] for x in epb(f"select stk_id, model from stkmas where stk_id in ({in_sku})", 500)
             if x["MODEL"]}
    in_model = ",".join(f"'{m}'" for m in model)
    alias = {s: s for s in skus}
    # 品牌 886＝未開保（類別1 1010），不是新機，不算預約機（使用者 10/3）；905 是另一個進貨管道，要算
    for x in epb(f"select stk_id, model from stkmas where model in ({in_model}) and brand_id <> '886'", 2000):
        alias.setdefault(x["STK_ID"], model[x["MODEL"]])
    return alias


def fetch_sales(alias):
    """POS 明細（POSLINE＋表頭 POSMAS）：含 F10「7 預約條碼」＝REF7、行備註與表頭備註（F8）。"""
    in_loc = ",".join(f"'{s[0]}'" for s in STORES)
    in_sku = ",".join(f"'{s}'" for s in sorted(alias))
    rows = epb(f"""select p.doc_id, p.doc_date, p.shop_id loc_id, p.stk_id, p.stk_qty, p.trans_type,
                   p.vip_id cust_id, p.ref7, p.remark line_remark, h.remark head_remark
            from posline p, posmas h
            where p.mas_rec_key = h.rec_key and p.shop_id in ({in_loc})
              and p.trans_type in ('A','E','H')
              and p.doc_date >= to_date('{SALES_FROM}','YYYY-MM-DD')
              and p.stk_id in ({in_sku})""")
    out = []
    for x in rows:   # 尾端空欄位會被截掉，一律用 get
        g = lambda k: (x.get(k) or "").strip()
        remark = " ".join(dict.fromkeys(t for t in (g("LINE_REMARK"), g("HEAD_REMARK")) if t))
        out.append(dict(DOC_ID=g("DOC_ID"), DOC_DATE=g("DOC_DATE"), LOC_ID=g("LOC_ID"), STK_ID=alias[g("STK_ID")],
                        STK_QTY=g("STK_QTY") or "0", TRANS_TYPE=g("TRANS_TYPE"), CUST_ID=g("CUST_ID"),
                        VIP=g("CUST_ID"), REF7=g("REF7"), REMARK=remark))
    return out


def fetch_store_moves(skus, alias):
    """EPB 庫存異動（STOREDTL），依實體倉 SA+門市碼：進貨淨額（調撥進出）、調撥出台數、目前庫存。
       調撥出要分開算：機器調走了，來源店那張預約的「已到貨」不會跟著收回，
       拿淨額判「到貨沒進EPB」會整批假警報（9/21 實測 7 筆有 5 筆是這樣）。"""
    sa = ",".join(f"'SA{s[0]}'" for s in STORES)
    in_sku = ",".join(f"'{a}'" for a, c in sorted(alias.items()) if c in skus)
    inbound, onhand, moved_out = defaultdict(int), defaultdict(int), defaultdict(int)
    for x in epb(f"""select store_id, stk_id, invmove_id, sum(stk_qty) q from storedtl
            where store_id in ({sa}) and stk_id in ({in_sku}) group by store_id, stk_id, invmove_id""", 5000):
        k, q = (x["STORE_ID"][2:], alias[x["STK_ID"]]), int(float(x["Q"]))
        onhand[k] += q
        if x["INVMOVE_ID"] not in ("POSOUT", "POSIN"):
            inbound[k] += q
            if q < 0:
                moved_out[k] += -q
    return inbound, onhand, moved_out


OPEN_ORDER = ("已到貨", "保留", "已配貨", "待付款", "已預約")
VIP_FROM = "2026-09-10"


def fetch_received(alias, cond, label):
    """依總倉出貨單的「我方參考」分類門市進貨台數，但只算已經進到門市庫存異動的。
       STOREDTL.OUR_REF 就是總倉那張出貨單的單號，拿它接回 INVTRNIMAS 才讀得到「我方參考」。
       不能直接用總倉出貨單算：總倉出了貨、門市還沒收（在途）的也會被算進來，
       拿去扣門市端進貨就會憑空少一台（9/23 士林、永和的 HQ VIP 機就是這樣被判成「到貨沒進EPB」）。"""
    in_sa = ",".join(f"'SA{s[0]}'" for s in STORES)
    in_sku = ",".join(f"'{a}'" for a in sorted(alias))
    n = defaultdict(int)
    for x in epb(f"""select d.store_id, d.stk_id, sum(d.stk_qty) q from storedtl d, invtrnimas h
            where d.our_ref = h.doc_id and d.store_id in ({in_sa}) and d.stk_id in ({in_sku})
              and d.invmove_id not in ('POSOUT', 'POSIN') and d.stk_qty > 0
              and h.doc_date >= to_date('{VIP_FROM}','YYYY-MM-DD') and {cond}
            group by d.store_id, d.stk_id""", 5000):
        n[(x["STORE_ID"][2:], alias[x["STK_ID"]])] += int(float(x["Q"]))
    print(f"{label}：北區 {sum(n.values())} 台")
    return n


def fetch_vip_units(alias):
    """總部 VIP 機（總倉出貨單「我方參考」含 Chester／VIP，總倉 SA099→SA999 再到店），不列入核對。"""
    return fetch_received(alias, "(upper(h.our_ref) like '%CHESTER%' or upper(h.our_ref) like '%VIP%')",
                          "總部 VIP 機（Chester／HQ VIP 收貨單，已進店）")


def fetch_spot_units(alias):
    """現貨進貨：收貨單「我方參考」是空白的（預約配貨會寫梯次「2026 iPhone 18 09xx」、VIP 寫 Chester／HQ VIP、
       店間調撥寫調撥單號），依每店每規格台數不列入預約核對。"""
    return fetch_received(alias, "(h.our_ref is null or trim(h.our_ref) = '')",
                          "現貨進貨（收貨單我方參考空白，已進店）")


def link_sales(sales, resv, by_key):
    """依 SOP 把每行結帳對到一張預約：
       ① F10 預約條碼（REF7）② 備註裡的預約單號 ③ 備註裡的會員編號（代取 SAxxxxxxxx）④ 結帳會員＋規格。
       都對不到、備註寫「現貨」的才算現貨。"""
    orders = {r["orderSNo"]: r for r in resv if r.get("orderSNo")}
    for s in sales:
        s["QTY"] = int(float(s["STK_QTY"] or 0))
        s["LINK"], s["HOW"] = None, ""
        s["BAD_REF7"] = s["REF7"] if s["REF7"] and s["REF7"] not in orders else ""
        s["DOA"] = "DOA" in s["REMARK"].upper()
        if s["REF7"] in orders:
            s["LINK"], s["HOW"] = orders[s["REF7"]], "條碼"
            continue
        found = [orders[o] for o in re.findall(r"\d{21}", s["REMARK"]) if o in orders]
        same = [r for r in found if r["stK_ID"] == s["STK_ID"]]
        if same or len(found) == 1:
            s["LINK"], s["HOW"] = (same or found)[0], "備註單號"
            continue
        for vip in re.findall(r"S[AT]T?\d{8}", s["REMARK"]):
            if by_key.get((vip, s["STK_ID"])):
                s["CUST_ID"], s["HOW"] = vip, "備註會員"
                break
    # ①-b 掃到「同一位會員另一張預約」的條碼（兩台一起取、條碼掃反／掃同一張）→ 改對到賣出規格那張
    valid = lambda r: (r["statusName"].startswith(PICKED) or r["statusName"] in OPEN_ORDER
                       or (r["statusName"] == "放棄" and r.get("arrivalEndTime")))   # 到店後放棄、還沒遞補：機器本來就是他的
    scanned = defaultdict(int)
    for s in sales:
        if s["HOW"] == "條碼" and s["QTY"] > 0:
            scanned[s["LINK"]["orderSNo"]] += 1
    for s in sales:
        r = s["LINK"]
        if s["HOW"] == "條碼" and s["STK_ID"] != r["stK_ID"]:
            own = sorted((x for x in by_key.get((r["vipId"], s["STK_ID"]), []) if valid(x)),
                         key=lambda x: scanned[x["orderSNo"]] > 0)
            if own:
                s["LINK"], s["HOW"], s["MIXUP"] = own[0], "條碼掃到同會員另一張", r["orderSNo"]
                scanned[own[0]["orderSNo"]] += 1
    # ④ 會員＋規格：全部結帳／銷退依時間一起處理（含已用條碼對到的，才能正確處理「結帳→銷退→重結」）
    #    結帳：一張預約只配一台，多的留給 VIP／其他檢核；已取消／放棄的不配；優先同店
    #    銷退：先抵銷同會員同規格先前對不到預約的結帳（兩筆都作廢），沒有才退回有配到的那張預約
    used = defaultdict(int)
    floating = defaultdict(list)                  # (會員, 規格) → 對不到預約、尚未被銷退抵銷的結帳
    rank = lambda r: (0 if r["statusName"].startswith(PICKED) else 1 if r["statusName"] in OPEN_ORDER else 2)
    # 已經有結帳掃到條碼的預約，留給那一筆；同會員同規格有兩張時，沒被掃到的那張優先給會員配對，
    # 否則「兩張預約各買一台、第二台掃錯成第一張的條碼」會同時報成「重複掃同一張」＋「B 改狀態無消費」
    barcoded = {x["LINK"]["orderSNo"] for x in sales
                if x["QTY"] > 0 and x["LINK"] and x["HOW"] in ("條碼", "備註單號", "條碼掃到同會員另一張")}
    for s in sales:
        s["VOID"] = False
    for s in sorted((x for x in sales if x["QTY"] != 0), key=doc_time):
        key = (s["CUST_ID"], s["STK_ID"])
        if s["LINK"]:
            if s["QTY"] < 0 and floating[key] and used[s["LINK"]["orderSNo"]] <= 0:
                x = floating[key].pop()
                s["LINK"], s["HOW"], s["VOID"], x["VOID"] = None, "", True, True
                continue
            used[s["LINK"]["orderSNo"]] += s["QTY"]
            continue
        if s["QTY"] > 0:
            cands = sorted((r for r in by_key.get(key, []) if valid(r) and used[r["orderSNo"]] <= 0),
                           key=lambda r: (BY_RESV_NAME[r["shopName"].strip()][0] != s["LOC_ID"],
                                          r["orderSNo"] in barcoded, rank(r)))
            if cands:
                s["LINK"], s["HOW"] = cands[0], s["HOW"] or "會員"
                used[cands[0]["orderSNo"]] += s["QTY"]
            else:
                floating[key].append(s)
        else:
            if floating[key]:
                x = floating[key].pop()
                s["VOID"] = x["VOID"] = True
                continue
            cands = [r for r in by_key.get(key, []) if used[r["orderSNo"]] > 0]
            if cands:
                s["LINK"], s["HOW"] = cands[0], s["HOW"] or "會員"
                used[cands[0]["orderSNo"]] += s["QTY"]
    for s in sales:
        s["STOCK"] = not s["LINK"] and bool(re.search(r"(?<![非不])現貨", s["REMARK"]))
    n = defaultdict(int)
    for s in sales:
        if s["QTY"] > 0:
            n[s["HOW"] or ("現貨" if s["STOCK"] else "沒對到")] += 1
    print("結帳對預約：" + "、".join(f"{k} {v}" for k, v in n.items()))


def build():
    resv = [r for r in fetch_reservations() if (r.get("shopName") or "").strip() in BY_RESV_NAME]
    print(f"北區預約 {len(resv)} 筆")
    by_key = defaultdict(list)  # (會員, 規格) → 預約
    for r in resv:
        if r.get("vipId") and r.get("stK_ID"):
            by_key[(r["vipId"].strip(), r["stK_ID"].strip())].append(r)
    sku_names = {r["stK_ID"]: r["productName"] for r in resv if r.get("stK_ID")}
    hot = fetch_hot_skus()

    alias = sku_aliases(set(sku_names))
    all_sales = fetch_sales(alias)
    print(f"EPB POS 明細 {len(all_sales)} 行（含同型號其他代碼 {sum(1 for a, c in alias.items() if a != c)} 個）")
    link_sales(all_sales, resv, by_key)
    vip_units = fetch_vip_units(alias)
    vip_left = dict(vip_units)
    # 門市用「現貨販售處理」把放棄的機器轉現貨：梯次尾端會加 -ShopSold（9/29 起狀態改叫「現貨販售」，之前是放棄(已遞補)）
    sold_left = defaultdict(int)
    for r in resv:
        if (r.get("batchNo") or "").endswith("ShopSold"):
            sold_left[(BY_RESV_NAME[r["shopName"].strip()][0], r["stK_ID"])] += 1
    print(f"門市轉現貨（梯次 ShopSold）：北區 {sum(sold_left.values())} 台")
    spot_units = fetch_spot_units(alias)
    spot_left = dict(spot_units)
    off_src = defaultdict(list)   # 預約單號 → 沒梯次的結帳用了哪種機器（銷退時還回去）
    for s in sorted(all_sales, key=doc_time):
        s["VIP_UNIT"] = s["SOLD_FILL"] = s["SPOT"] = False
        s["OFF_SRC"] = ""
        k = (s["LOC_ID"], s["STK_ID"])
        # 預約沒有梯次（系統從來沒配機器給他）卻有結帳：門市是拿店裡的 VIP 機／現貨直接賣給排隊的客人、再手動改已取貨。
        # 那台機器要從 VIP／現貨扣，不然 VIP 機一直算「還在店裡」→ 庫存變負數、多出假的「確定賣了沒改」（10/3 高島屋 606／602）
        if s["QTY"] < 0 and s["LINK"] and off_src.get(s["LINK"]["orderSNo"]):
            # 銷退要跟著原本那筆結帳走：那台機器回到 VIP／現貨，不能留在預約這邊（10/3 高島屋 07313588 SA00******）
            s["OFF_SRC"] = src = off_src[s["LINK"]["orderSNo"]].pop()
            if src == "vip":
                vip_left[k] = vip_left.get(k, 0) + 1
            elif src == "spot":
                spot_left[k] = spot_left.get(k, 0) + 1
        if s["QTY"] > 0 and s["LINK"] and not s["LINK"].get("batchNo") and not s["VOID"]:
            if vip_left.get(k, 0) > 0:
                s["OFF_SRC"] = "vip"
                vip_left[k] -= 1
            elif spot_left.get(k, 0) > 0:
                s["OFF_SRC"] = "spot"
                spot_left[k] -= 1
            else:
                s["OFF_SRC"] = "pool"       # 用的是店裡沒配出去的預約機
            off_src[s["LINK"]["orderSNo"]].append(s["OFF_SRC"])
        if s["QTY"] > 0 and not s["LINK"] and not s["STOCK"] and not s["VOID"]:
            if vip_left.get(k, 0) > 0:
                s["VIP_UNIT"] = True
                vip_left[k] -= 1
            elif sold_left.get(k, 0) > 0:   # 轉現貨賣掉的，跟備註寫「現貨」的一樣不列入核對
                s["SOLD_FILL"] = s["STOCK"] = True
                sold_left[k] -= 1
            elif spot_left.get(k, 0) > 0:   # 本來就是現貨進貨的機器賣掉（機器跟預約機同一堆，只在賣出這側當現貨）
                s["SPOT"] = s["STOCK"] = True
                spot_left[k] -= 1
    cutoff = (datetime.now() - timedelta(minutes=SETTLE_MIN)).strftime("%Y%m%d%H%M%S")
    is_recent = lambda d: doc_time(d) > cutoff  # 結帳未滿 SETTLE_MIN 分鐘

    per = defaultdict(list)                       # 預約單號 → 對到的結帳行
    for s in all_sales:
        if s["LINK"]:
            per[s["LINK"]["orderSNo"]].append(s)
    code_of = lambda r: BY_RESV_NAME[r["shopName"].strip()][0]
    items = []                                    # (類型, 預約, 結帳行, 備註)
    no_sale = []                                  # 已取貨但沒有任何結帳對到（四／五）
    eff_ids = set()                               # 各預約有效（沒被銷退抵掉）的結帳
    for r in resv:
        lines = sorted(per.get(r["orderSNo"], []), key=doc_time)
        net = sum(x["QTY"] for x in lines)
        net_ok = sum(x["QTY"] for x in lines if not is_recent(x))
        pos = [x for x in lines if x["QTY"] > 0]
        last = pos[-1] if pos else None
        picked = r["statusName"].startswith(PICKED)
        if net_ok > 0 and not picked:
            how = last["HOW"]
            why = ("有掃預約條碼，但狀態沒有自動改" if how == "條碼"
                   else f"結帳沒掃預約條碼（以{ {'會員': '會員＋規格', '備註單號': '備註單號', '備註會員': '備註會員'}.get(how, how) }對到）")
            if r["statusName"] == "放棄":
                why += "；⚠️ 預約是「放棄」還沒遞補（系統只會自動改已到貨／保留）→ 立刻改已取貨，否則系統會把這台遞補給下一位"
            items.append(("A 結帳未改", r, last, why + (f"；備註：{last['REMARK']}" if last["REMARK"] else "")))
        if picked and net <= 0:
            if pos and (any(x["DOA"] for x in lines) or r.get("doa")):
                # 預約系統自己有 doa 時戳，比門市有沒有在 POS 備註打「DOA」可靠
                items.append(("DOA 待客服改狀態", r, last, "DOA 銷退，依 SOP 由客服改回「已預約」"
                              + (f"（預約系統 DOA 時間 {r['doa'][:16]}）" if r.get("doa") else "（依 POS 備註判斷）")))
            elif pos:
                items.append(("銷退仍已取貨", r, last, "結帳後已銷退，預約還是已取貨：沒取機就改回原狀態"))
            else:
                no_sale.append(r)
        if net > 1:
            items.append(("重複掃同一張", r, last, f"這張預約被結帳 {net} 台"))
        effective = []                            # 沒被銷退抵掉的結帳（結帳→銷退→重結只看最後那次）
        for x in lines:
            if x["QTY"] > 0:
                effective.append(x)
            elif effective:
                same = [y for y in effective if y["STK_ID"] == x["STK_ID"]] or effective
                effective.remove(same[-1])
        eff_ids.update(id(x) for x in effective)
        for x in effective:
            if x["HOW"] in ("條碼", "備註單號") and x["STK_ID"] != r["stK_ID"]:   # 同會員另一張的已自動改對，不算
                items.append(("掃錯規格", r, x, f"賣出的是 {x['STK_ID']}，掃到的預約是 {r['stK_ID']}（SOP 規格不能改）"))
            if x["VIP"] != r["vipId"]:
                items.append(("非本人取機", r, x, (f"結帳會員 {x['VIP']}≠預約會員" if x["VIP"] else "結帳沒打會員號")
                              + "（SOP 限本人領取、須出示本人 LINE 會員卡）"
                              + (f"；備註：{x['REMARK']}" if x["REMARK"] else "")))
            if x["LOC_ID"] != code_of(r):
                items.append(("跨店取機", r, x, f"在 {BY_CODE.get(x['LOC_ID'], ('', '', x['LOC_ID']))[2]} 結帳（SOP 取貨門市不能改）"))
    # 四／五：已取貨但沒有消費 → 同店同規格有對不到預約的結帳（無會員或別人會員）＝疑似代取，否則異常
    pool, pool2 = defaultdict(list), defaultdict(list)
    for s in sorted(all_sales, key=doc_time):
        if s["QTY"] > 0 and not s["LINK"] and not s["VIP_UNIT"] and not s["VOID"]:
            if not s["STOCK"]:
                pool[(s["LOC_ID"], s["STK_ID"])].append(s)
            elif s["SPOT"] or s["SOLD_FILL"]:
                # 被現貨／轉現貨台數吃掉的結帳：優先當現貨，但沒有別的結帳可配時還是要拿來提醒門市比對
                pool2[(s["LOC_ID"], s["STK_ID"])].append(s)
    paired = set()
    for r in no_sale:
        k = (code_of(r), r["stK_ID"])
        cand, spot = pool.get(k), False
        if not cand:
            cand, spot = pool2.get(k), True
        if cand:
            x = cand.pop(0)
            paired.add(id(x))
            items.append(("疑似代取", r, x, f"同店同規格有一筆結帳對不到預約（結帳會員 {x['VIP'] or '未打'}）"
                          "：非本人取機只能購買、不能累積會員，請確認是不是這張的代取；是的話在 POS 備註補「代取 會員號」"
                          + ("；⚠️ 本店該規格有現貨／已轉現貨的機器，這筆也可能只是現貨賣出，請一併確認" if spot else "")))
        else:
            # 把這位會員的實際結帳攤出來（含銷退、含別店），不要只說「買了其他規格」讓人去猜是不是轉現貨賣給他
            buys = sorted((x for x in all_sales if x["VIP"] == r["vipId"]), key=doc_time)
            detail = "；".join(f"{BY_CODE.get(x['LOC_ID'], ('', '', x['LOC_ID']))[2]} {x['STK_ID']} {x['QTY']:+d} {fmt_time(x)}"
                               for x in buys)
            items.append(("B 改狀態無消費", r, None,
                          f"本規格 {r['stK_ID']} 在 EPB 完全查不到結帳；這位會員的全部結帳："
                          + (detail + "（SOP 規格不能改、取貨門市不能改）" if detail else "一筆都沒有")))
    now_dt = datetime.now()
    now = now_dt.strftime("%Y-%m-%d %H:%M:%S")
    transfers = fetch_transfers()
    transfer_from = {x["fromOrderSNo"] for x in transfers if x.get("fromOrderSNo")}
    print(f"跨店調撥建議 {len(transfers)} 筆（放棄的機器已排進遞補的 {len(transfer_from)} 張）")
    # 來源店已確認出貨、預約系統的調撥單卻還沒處理完：機器多半已經在本店，卻還沒接上要遞補的那位客人
    pending_in = defaultdict(list)
    for t in transfers:
        st = BY_RESV_NAME.get((t.get("toShop") or "").strip())
        if st and t.get("stK_ID") and t.get("fromShopConfirmTime"):
            pending_in[(st[0], t["stK_ID"])].append(t)
    by_order = {r["orderSNo"]: r for r in resv if r.get("orderSNo")}
    for s in all_sales:
        if s["QTY"] <= 0:
            continue
        if s["BAD_REF7"] and not s["VOID"] and (not s["LINK"] or id(s) in eff_ids):   # 被銷退掉的那筆不算
            ok = s["LINK"] and s["LINK"]["statusName"].startswith(PICKED)
            items.append(("條碼欄打錯", s["LINK"], s, f"預約條碼欄填了「{s['BAD_REF7']}」，不是預約單號"
                          + ("；預約狀態已手動改對，只需宣導" if ok else "；確認是哪張預約並手動改狀態")))
        if not s["LINK"] and not s["STOCK"] and not s["VIP_UNIT"] and not s["SPOT"] and not s["VOID"] and hot_at(s["STK_ID"], doc_time(s)) and not is_recent(s) and id(s) not in paired:
            # 本店有一張「來源店已出貨、預約系統還沒處理完」的調撥單，卻把熱門機賣給沒有預約的客人：
            # 那台多半就是調撥進來要遞補的機器，賣掉之後要遞補的人會撲空（9/23 阿波羅 07313595 實例）
            tf = pending_in.get((s["LOC_ID"], s["STK_ID"]))
            if tf:
                t = tf[0]
                dst = by_order.get(t.get("toOrderSNo") or "")
                items.append(("調撥機被當現貨賣", dst, s,
                              f"{t.get('fromShop', '')} 已於 {str(t.get('fromShopConfirmTime'))[:16]} 確認出貨給本店"
                              f"（調撥單 {t.get('transferBatchNumber', '')}），預約系統的調撥進度還停在"
                              f"「{t.get('processProgress', '')}」；要遞補的是 {t.get('toOrderSNo', '')}"
                              + (f"（目前狀態「{dst['statusName']}」）" if dst else "")
                              + f"，機器卻在 {fmt_time(s)} 賣給沒有任何預約的 {s['VIP']}"))
            else:
                pass   # 「熱門無預約結帳(參考)」使用者 10/3 決定不列：現貨或 VIP，不算異常

    # 六／七：門市已按到貨、後來放棄或取消、沒有遞補下一位（沒有「(已遞補)」尾綴）；逾期：過了取貨期限還是已到貨／保留
    sold_orders = {o for o, ls in per.items() if sum(x["QTY"] for x in ls) > 0}   # 有有效結帳的預約（機器已賣出）
    queue = defaultdict(list)                     # (門市, 規格) → 排隊中（已預約）依系統順序
    for r in resv:
        if r["statusName"] == "已預約":
            queue[(code_of(r), r["stK_ID"])].append(r)
    for q in queue.values():                      # 會員等級高→非再購→登記早（預約單號前 14 碼＝登記時間）
        q.sort(key=lambda r: (-(r.get("userClass") or 0), bool(r.get("isRepeatPurchase")), r["orderSNo"][:14]))
    for r in resv:
        if r["statusName"] in ("放棄", "已取消") and r.get("batchNo") and r.get("arrivalEndTime") \
                and r["orderSNo"] not in sold_orders:                                   # 已賣出的列在 A，不叫門市遞補
            q = queue.get((code_of(r), r["stK_ID"]), [])
            just_expired = timedelta(0) <= now_dt - datetime.strptime(r["arrivalEndTime"][:19], "%Y-%m-%d %H:%M:%S") < timedelta(hours=12)
            if just_expired:
                items.append(("放棄待系統遞補", r, None, f"取貨期限 {r['arrivalEndTime'][:16]} 剛到期、系統自動放棄，遞補還沒跑"
                              f"（本店 {len(q)} 位排隊）→ 12 小時後仍未遞補才會列為待處理"))
            elif q:
                nxt = q[0]
                items.append(("放棄未遞補(有人排隊)", r, None, f"本店本規格還有 {len(q)} 位在排隊，下一位：{nxt['orderSNo']}"
                              f"（{nxt['vipId']}，{nxt.get('userClassName', '')}）→ 請立即遞補"))
            elif r["stK_ID"] in hot:
                # 規格剛被列為熱門時，之前就放棄的機器要等列熱門後系統跑過遞補才算數，不能從取貨期限起算
                start = max(datetime.strptime(r["arrivalEndTime"][:19], "%Y-%m-%d %H:%M:%S"), hot_since(r["stK_ID"]) or datetime.min)
                late = now_dt - start
                if late > timedelta(hours=24) and r["orderSNo"] not in transfer_from:
                    items.append(("放棄逾24小時未遞補", r, None,
                                  f"取貨期限 {r['arrivalEndTime'][:16]}{'（' + start.strftime('%m/%d %H:%M') + ' 才列為熱門）' if start > datetime.strptime(r['arrivalEndTime'][:19], '%Y-%m-%d %H:%M:%S') else ''} 已過 {int(late.total_seconds() // 3600)} 小時，"
                                  "系統跨店遞補一直沒有排到這台，目前的調撥建議清單裡也沒有 → 不會自己好，要人工處理"))
                else:
                    items.append(("放棄待系統遞補", r, None, "熱門款本店沒人排隊，依 SOP 由系統凌晨 3 點計算跨店遞補（不能轉現貨）"))
            else:
                items.append(("放棄未遞補(提醒轉現貨)", r, None, "非熱門、本店沒人排隊 → 依 SOP 用「現貨販售處理」轉現貨"))
        if r["statusName"] in ("已到貨", "保留") and r.get("arrivalEndTime"):
            # 保留的期限門市可以一直延，一律以系統的 reserveEndTime 為準（＝畫面上的「預計取機時間」）；
            # 已到貨沒有這個欄位，用 arrivalEndTime。9/22 實測北區 47 張保留只有 9 張剛好是＋2 天。
            end = r.get("reserveEndTime") or r["arrivalEndTime"]
            due = datetime.strptime(end[:19], "%Y-%m-%d %H:%M:%S")
            if due < now_dt:
                items.append(("逾期未取", r, None, f"{'保留' if r['statusName'] == '保留' else '取貨'}期限 {due:%m/%d %H:%M} 已過"
                              "，系統沒有自動放棄 → 聯絡客人，不取就改放棄讓系統遞補"))

    # 會員核對（二／三）：預約網站總預約數 vs 結帳淨台數
    net_v = defaultdict(int)
    sales_v = defaultdict(list)
    for x in all_sales:
        v = x["VIP"] or (x["LINK"]["vipId"] if x["LINK"] else "")
        if v and not x["STOCK"] and not x["VIP_UNIT"] and not x["SPOT"]:
            net_v[v] += x["QTY"]
            sales_v[v].append(x)
    by_vip = defaultdict(list)
    for r in resv:
        if r.get("vipId"):
            by_vip[r["vipId"]].append(r)
    members = []
    for v, rs in by_vip.items():
        picked_n = sum(1 for r in rs if r["statusName"].startswith(PICKED))
        if picked_n == net_v.get(v, 0):
            continue
        opens = [r for r in rs if r["statusName"] in OPEN_ORDER]
        tag = ("排隊中已買到" if opens else "結帳比已取貨多") if net_v.get(v, 0) > picked_n else "已取貨比結帳多"
        members.append(dict(vip=v, tag=tag, n_resv=len(rs), picked=picked_n, bought=net_v.get(v, 0), opens=len(opens),
                            stores="、".join(sorted({r["shopName"].strip() for r in rs})),
                            resv="；".join(f"{r['shopName'].strip()} {r['stK_ID']} {r['statusName']} …{r['orderSNo'][-5:]}" for r in rs),
                            sales="；".join(f"{BY_CODE.get(x['LOC_ID'], ('', '', x['LOC_ID']))[2]} {x['STK_ID']} {x['QTY']:+d} {fmt_time(x)}"
                                           for x in sorted(sales_v.get(v, []), key=doc_time))))
    unscanned = [s for s in all_sales if s["QTY"] > 0 and s["LINK"] and s["HOW"] != "條碼"]
    sales = [s for s in all_sales if not s["STOCK"] and not s["VIP_UNIT"] and not s["SPOT"] and s["OFF_SRC"] not in ("vip", "spot")]
    stock_sales = [s for s in all_sales if s["STOCK"] or s["OFF_SRC"] == "spot"]
    vip_sold, spot_sold = defaultdict(int), defaultdict(int)
    for s in all_sales:
        if s["VIP_UNIT"] or s["OFF_SRC"] == "vip":
            vip_sold[(s["LOC_ID"], s["STK_ID"])] += s["QTY"]
        if s["SPOT"] or s["OFF_SRC"] == "spot":
            spot_sold[(s["LOC_ID"], s["STK_ID"])] += s["QTY"]
    pool_net = defaultdict(int)
    for s in all_sales:
        if s["OFF_SRC"] == "pool":
            pool_net[s["LINK"]["orderSNo"]] += s["QTY"]
    nobatch_pool = {o for o, n in pool_net.items() if n > 0}
    rows_h = build_hot(resv, sales, stock_sales, is_recent, alias, vip_units, vip_sold, spot_units, spot_sold, nobatch_pool)
    # 對不到預約的結帳，如果那位會員在本店同規格有「放棄(已遞補)／已取消(已遞補)」的預約，
    # 多半就是那張被遞補掉、客人後來還是來買了 → 直接寫出來，不要叫門市去找代取
    hint = defaultdict(list)
    for x in all_sales:
        if x["QTY"] > 0 and not x["LINK"] and not x["STOCK"] and not x["VIP_UNIT"] and not x["SPOT"] and not x["VOID"]:
            for r in by_key.get((x["VIP"], x["STK_ID"]), []):
                if "(已遞補)" in r["statusName"] and code_of(r) == x["LOC_ID"]:
                    hint[(x["LOC_ID"], x["STK_ID"])].append(
                        f"結帳會員 {x['VIP']} 在本店同規格有一張「{r['statusName']}」的預約 {r['orderSNo']}"
                        "：那台機器已經遞補給別人，客人後來還是買了 → 要確認遞補對象拿到的是哪一台")
    for h in rows_h:
        h["hint"] = hint.get((h["loc"], h["sku"]), [])
        h["queue"] = len(queue.get((h["loc"], h["sku"]), []))   # 本店還在排隊（已預約）的人數
        h["queue_all"] = sum(len(v) for k, v in queue.items() if k[1] == h["sku"])   # 全區排隊人數
    # 9/30 起放棄分三種：「放棄(需到店確認)」＝客人在已配貨時放棄，機器到店後門市要按「門市確認」系統才會往後遞補。
    # 這種單沒有 arrivalEndTime，只能用同店同規格「未按到貨但機器已進 EPB」的台數判斷機器到了沒，到了幾台就列幾張
    confirm = defaultdict(list)
    for r in resv:
        if "需到店確認" in r["statusName"] and r["orderSNo"] not in sold_orders:
            confirm[(code_of(r), r["stK_ID"])].append(r)
    for h in rows_h:
        rs = sorted(confirm.get((h["loc"], h["sku"]), []), key=lambda r: r["orderSNo"])
        n = min(len(rs), h["cfg_arrived"])
        for r in rs[:n]:
            q = queue.get((h["loc"], h["sku"]), [])
            nxt = (f"本店還有 {len(q)} 位排隊，按了系統會自動遞補給下一位 {q[0]['orderSNo']}" if q else
                   "本店沒人排隊 → 按完" + ("熱門依 SOP 不能轉現貨，等系統跨店調撥" if h["hot"] else "用「現貨販售處理」轉現貨"))
            items.append(("放棄需到店確認", r, None,
                          f"本店本規格已進 EPB、還沒按到貨的有 {h['cfg_arrived']} 台（未按到貨共 {h['cfg']} 張，"
                          f"其中需到店確認 {len(rs)} 張）；{nxt}"))
    return resv, items, unscanned, rows_h, sku_names, cutoff, members


def build_hot(resv, sales, stock_sales, is_recent, alias, vip_units, vip_sold, spot_units, spot_sold, nobatch_pool=()):
    """門市×規格庫存對帳，不靠會員號（熱門款依 SOP 不能賣現貨；非熱門只看跟現貨無關的確定錯誤）：
       狀態差 ＝ EPB 賣出(滿 SETTLE_MIN 分、不含備註現貨) − 預約已取貨 → >0 賣了沒改、<0 改了沒賣
       進貨 − 已配給預約 → >0 到店還沒配（非異常）、<0 預約顯示到貨但 EPB 沒進貨"""
    hot = fetch_hot_skus()
    skus = {r["stK_ID"] for r in resv}
    code_of = lambda r: BY_RESV_NAME[r["shopName"].strip()][0]
    inbound, onhand, moved_out = fetch_store_moves(skus, alias)
    g = defaultdict(lambda: defaultdict(int))
    for r in resv:
        # 轉現貨(ShopSold)的機器已經離開預約池，不管賣掉沒有都不能再算成「到店還沒配」
        # ⚠️ 9/29 起系統把這些預約的狀態從「放棄(已遞補)」改成「現貨販售」（舊的也一起改），
        #    所以一律看梯次、不要看狀態名稱；漏掉會把轉現貨的機器當成已到店，灌出一堆假的「到貨沒進EPB」
        if (r.get("batchNo") or "").endswith("ShopSold"):
            g[(code_of(r), r["stK_ID"])]["shopsold_q"] += 1
            continue
        # 「(已遞補)」＝這張的機器讓給別人了，不佔本店的機器；但「已取貨(已遞補)」是被遞補到、而且真的來取了，
        # 一樣要算已取貨，不然會平白多出一台「確定賣了沒改」（9/23 微風 07313586、新西門 07313603 就是這樣誤判）。
        picked_up = r["statusName"].startswith(PICKED)
        if not r.get("batchNo") and picked_up and r["orderSNo"] in nobatch_pool:
            g[(code_of(r), r["stK_ID"])]["alloc"] += 1    # 沒配貨、直接拿店裡未配的預約機賣給他
            g[(code_of(r), r["stK_ID"])]["picked"] += 1
            continue
        if not r.get("batchNo") or ("(已遞補)" in r["statusName"] and not picked_up):
            continue
        x = g[(code_of(r), r["stK_ID"])]
        if not r.get("arrivalEndTime") and not picked_up:   # 已取貨(已遞補) 的取貨期限會被系統清掉
            x["cfg"] += 1           # 沒有取貨期限＝門市還沒按到貨（已配貨、待付款、配貨中就放棄…），機器可能還沒到店
            continue
        x["alloc"] += 1
        if r["statusName"].startswith(PICKED):
            x["picked"] += 1
        else:
            x["in_store"] += 1
            x["sheet" if r["statusName"] in ("已到貨", "保留") else "dead"] += 1   # 待取貨庫存清點單＝已到貨＋保留
    for x in sales:
        qty = int(float(x["STK_QTY"] or 0))
        key = "recent" if is_recent(x) else "sold"
        k = (x["LOC_ID"], x["STK_ID"].strip())
        g[k][key] += qty
        # 跨店取機：機器是在這家店賣掉的，但「已取貨」記在預約門市 → 這家店要當成自己的已取貨，
        # 否則會多出一台「確定賣了沒改」（實際上跟預約門市那列的「跨店取機」是同一台）
        lk = x.get("LINK")
        if lk and lk.get("shopName") and BY_RESV_NAME[lk["shopName"].strip()][0] != x["LOC_ID"]:
            g[k]["alloc"] += qty
            g[k]["picked"] += qty
            g[k]["cross"] += qty
    sold_days = defaultdict(set)
    for x in stock_sales:
        k = (x["LOC_ID"], x["STK_ID"].strip())
        q = int(float(x["STK_QTY"] or 0))
        g[k]["stock"] += q
        # 三種現貨要分開講：門市後台轉的（梯次 -ShopSold）跟現貨進貨的機器，POS 備註都是空的，
        # 不能一律寫成「備註寫現貨」（9/23 那 3 台全部是 ShopSold，備註都空的）
        if x.get("SOLD_FILL"):
            g[k]["shopsold"] += q
            sold_days[k].add(doc_time(x)[4:8])
            if hot_at(k[1], doc_time(x)):
                g[k]["shopsold_hot"] += q      # 賣出時已經是熱門
        elif not x.get("SPOT") and x.get("OFF_SRC") != "spot":
            g[k]["remark"] += q
            if hot_at(k[1], doc_time(x)):
                g[k]["remark_hot"] += q        # 熱門清單會變：只有賣出當下是熱門才算違反 SOP
    # 剛結帳（未滿 SETTLE_MIN 分）而且已經改好已取貨的，這張先不算，免得誤判改了沒賣（代取也適用）
    for x in g.values():
        x["picked_recent"] = min(max(x["recent"], 0), max(x["picked"] - x["sold"], 0))
    for k in skus:
        for loc in {s[0] for s in STORES}:
            if inbound.get((loc, k)) or onhand.get((loc, k)):
                v, vs = vip_units.get((loc, k), 0), vip_sold.get((loc, k), 0)
                sp, sps = spot_units.get((loc, k), 0), spot_sold.get((loc, k), 0)
                # 還在店裡的 VIP 機不可能比 EPB 實際庫存多：多出來的就是 VIP 機被調走、或拿去給預約客人了
                # （10/3 高島屋 601／606：VIP 機補給有梯次的預約、調撥出），當成那幾台 VIP 機已離店，進貨也要加回來
                excess = max((v - vs) - max(onhand.get((loc, k), 0), 0), 0)
                g[(loc, k)]["inbound"] = inbound.get((loc, k), 0) - v + excess  # 不含總部 VIP 機
                g[(loc, k)]["onhand"] = onhand.get((loc, k), 0) - (v - vs - excess)   # 扣掉還在店裡的 VIP 機
                g[(loc, k)]["vip"], g[(loc, k)]["vip_sold"] = v, vs
                g[(loc, k)]["spot"], g[(loc, k)]["spot_sold"] = sp, sps
                g[(loc, k)]["moved_out"] = moved_out.get((loc, k), 0)
    rows = []
    for (loc, sku), x in g.items():
        unalloc = x["inbound"] - x["alloc"]
        # 兩個問題要用兩個數字，不能共用一個：
        # ①「進貨夠不夠支應已按到貨的預約」→ 只能扣已經賣掉的現貨（沒賣掉的還在店裡，不構成短少）
        rest_in = unalloc - x["stock"]
        cfg_arrived = min(max(rest_in, 0), x["cfg"])    # 已配貨但機器已進 EPB＝沒按到貨
        free = rest_in - cfg_arrived                    # 進貨但沒配給預約的淨台數（可為負＝預約到貨比進貨多）
        # ②「店裡有沒有沒人認領的預約機」→ 轉現貨(ShopSold)、現貨進貨、備註現貨的機器都已離開預約池，
        #    不管賣掉沒有都要扣；只扣賣掉的，沒賣掉那台會被誤報成「到店還沒配」（9/23 微風 07313590）
        diff = x["sold"] - (x["picked"] - x["picked_recent"])
        u = max(unalloc - x["shopsold_q"] - x["spot"] - x["remark"] - cfg_arrived, 0)
        maybe = min(max(diff, 0), u)                    # 多賣的台數裡，未配的機器夠 cover 的部分：分不出是沒改還是未配機直接賣
        rows.append(dict(loc=loc, sku=sku, hot=sku in hot, inbound=x["inbound"], alloc=x["alloc"], cfg=x["cfg"],
                         cfg_arrived=cfg_arrived, sold=x["sold"], picked=x["picked"] - x["picked_recent"],
                         recent=x["recent"], stock=x["stock"], onhand=x["onhand"], in_store=x["in_store"],
                         sheet=x["sheet"], dead=x["dead"], vip=x["vip"], vip_sold=x["vip_sold"],
                         spot=x["spot"], spot_sold=x["spot_sold"],
                         remark=x["remark"], remark_hot=x["remark_hot"], shopsold=x["shopsold"], shopsold_hot=x["shopsold_hot"],
                         shopsold_q=x["shopsold_q"],
                         sold_days="、".join(f"{d[:2]}/{d[2:]}" for d in sorted(sold_days.get((loc, sku), ()))),
                         moved_out=x["moved_out"], cross=x["cross"],
                         sure=max(diff - u, 0), maybe=maybe, over=max(-diff, 0),
                         no_inbound=max(-free - x["moved_out"], 0), unalloc=u - maybe))
    return rows


def fmt_time(d):
    t = doc_time(d) if d else ""
    return f"{t[4:6]}/{t[6:8]} {t[8:10]}:{t[10:12]}" if len(t) == 14 else (d or {}).get("DOC_DATE", "")[:10]


HEAD_FILL = PatternFill("solid", fgColor="DDEBF7")
BOLD = Font(bold=True)


def store_of(r):
    return BY_RESV_NAME[r["shopName"].strip()]


def record(kind, r, s, note, sku_names):
    st = store_of(r) if r else BY_CODE[s["LOC_ID"]]
    sale_store = BY_CODE.get((s or {}).get("LOC_ID", ""), ("", "", (s or {}).get("LOC_ID", ""), ""))[2] if s else ""
    sku = r["stK_ID"] if r else s["STK_ID"]
    return dict(region=st[3], store=st[2], idx=STORES.index(st), kind=kind, qty=1, todo=TODO[kind],
                urgency=URGENCY.get(kind, "⚪ 參考"), diff=DIFF.get(kind, ""),
                status=r["statusName"] if r else "（無預約）", scene=ACTIVITIES.get(r["reservationActivityId"], "") if r else "",
                order=r.get("orderSNo") if r else "", vip=r["vipId"] if r else (s or {}).get("VIP", ""),
                name=mask_name(r.get("subscriberName")) if r else "", sku=sku,
                product=sku_names.get(sku, "").replace("預約｜", ""), batch=r.get("batchNo") if r else "",
                sale_store=sale_store, sale_time=fmt_time(s) if s else "", doc=(s or {}).get("DOC_ID", ""), note=note)


# 🔴 今天不處理會出事（機器已經不在了、系統還會拿它去遞補，或客人在等）
# 🟡 要查清楚，但不會馬上害到別人　⚪ 不用動作，看一眼就好
URGENCY = {
    "A 結帳未改": "🔴 今天處理", "熱門機在店沒配人": "🔴 今天處理", "放棄逾24小時未遞補": "🔴 今天處理",
    "放棄未遞補(有人排隊)": "🔴 今天處理", "逾期未取": "🔴 今天處理", "放棄需到店確認": "🔴 今天處理",
    "B 改狀態無消費": "🟡 要確認", "疑似代取": "🟡 要確認", "銷退仍已取貨": "🟡 要確認",
    "掃錯規格": "🟡 要確認", "重複掃同一張": "🟡 要確認", "條碼欄打錯": "🟡 要確認",
    "非本人取機": "🟡 要確認", "跨店取機": "🟡 要確認", "熱門 標現貨": "🟡 要確認",
    "調撥機被當現貨賣": "🔴 今天處理", "熱門款轉了現貨": "⚪ 參考", "熱門機轉現貨沒賣掉": "🔴 今天處理",
    "庫存 確定賣了沒改": "🟡 要確認", "庫存 改了沒賣": "🟡 要確認", "庫存 到貨沒進EPB": "🟡 要確認",
    "清點單機器不足": "🟡 要確認", "放棄未遞補(提醒轉現貨)": "🟡 要確認", "未掃條碼": "🟡 要確認",
    "DOA 待客服改狀態": "⚪ 參考", "放棄待系統遞補": "⚪ 參考", "熱門無預約結帳(參考)": "⚪ 參考",
    "到店還沒配": "⚪ 參考",
}
# 「差異在哪」：一句話講兩邊對不起來的地方，不寫動作
DIFF = {
    "A 結帳未改": "EPB 已賣出，預約系統還沒改「已取貨」",
    "B 改狀態無消費": "預約系統寫「已取貨」，EPB 查不到結帳",
    "疑似代取": "預約已取貨但沒有本人結帳，同店同規格有一筆對不到預約的結帳",
    "放棄未遞補(有人排隊)": "機器已放棄，本店還有人在排隊卻沒遞補",
    "放棄未遞補(提醒轉現貨)": "非熱門機已放棄，本店沒人排隊",
    "放棄待系統遞補": "熱門機已放棄、本店沒人排隊，等系統跨店遞補",
    "放棄逾24小時未遞補": "機器放棄超過 24 小時，系統跨店遞補一直沒排到",
    "放棄需到店確認": "客人在配貨中放棄，機器已進 EPB，門市還沒按「門市確認」",
    "逾期未取": "過了預約系統的「預計取機時間」，還是已到貨／保留",
    "清點單機器不足": "清點單（已到貨＋保留）比 EPB 庫存多",
    "熱門機在店沒配人": "EPB 有熱門機在店、系統沒配給任何人，本店卻有人在排隊",
    "銷退仍已取貨": "EPB 已經銷退，預約系統還是「已取貨」",
    "掃錯規格": "賣出的規格，跟掃到的預約條碼不同規格",
    "重複掃同一張": "同一張預約條碼被結帳兩台以上",
    "條碼欄打錯": "F10 第 7 欄填的不是預約單號（序號／電話等）",
    "非本人取機": "結帳會員跟預約會員不同（或沒打會員）",
    "跨店取機": "預約門市跟結帳門市不是同一家",
    "熱門 標現貨": "熱門機用「現貨」備註賣出（SOP 不能賣現貨）",
    "熱門無預約結帳(參考)": "熱門機結帳，買方沒有任何預約（現貨或 VIP）",
    "調撥機被當現貨賣": "別店調撥來要遞補的熱門機，賣給了沒有預約的客人",
    "熱門款轉了現貨": "熱門機被用後台「現貨販售處理」轉成現貨賣掉",
    "熱門機轉現貨沒賣掉": "熱門機已轉現貨、還沒賣掉，機器還在店裡但已不在遞補池",
    "DOA 待客服改狀態": "DOA 已銷退，預約系統還是「已取貨」",
    "到店還沒配": "EPB 有貨，預約系統還沒配給客人",
    "未掃條碼": "狀態是對的，但結帳沒在 F10 第 7 欄掃預約條碼",
}
TODO = {
    "A 結帳未改": "把這張預約改成「已取貨」（不改的話系統會把這台再遞補給別人）",
    "B 改狀態無消費": "客人沒取就改回原狀態；有取就補結帳並補掃條碼",
    "疑似代取": "確認是不是代取；是的話在 POS 備註補「代取 會員號」（非本人只能購買、不能累積會員）",
    "放棄未遞補(有人排隊)": "立即遞補給下一位（備註欄有單號）",
    "放棄未遞補(提醒轉現貨)": "用預約後台的「現貨販售處理」轉現貨",
    "放棄待系統遞補": "不用做，等系統凌晨 3 點跑；隔天還在才處理",
    "放棄逾24小時未遞補": "回報區經理／總部人工處理（不會自己好）",
    "放棄需到店確認": "預約後台按「門市確認」；本店沒人排隊再看非熱門轉現貨、熱門等跨店調撥（備註有寫）",
    "逾期未取": "聯絡客人；不取就改放棄，讓系統遞補",
    "清點單機器不足": "盤點預約區，找出已取機沒改狀態的",
    "熱門機在店沒配人": "按到貨接上遞補；若不是本店要的機器，回報總部重配",
    "銷退仍已取貨": "客人沒取機就改回原狀態（銷退不會自動改；入錯會員／重結最常見，請勿略過）",
    "掃錯規格": "確認賣出的是哪張預約，兩張狀態都改正確",
    "重複掃同一張": "確認另一台是誰的預約，改正確",
    "條碼欄打錯": "確認這筆是哪張預約、手動改狀態；下次要掃預約單號",
    "非本人取機": "確認取機人；SOP 限會員本人領取、須出示本人 LINE 會員卡",
    "跨店取機": "確認兩店預約區庫存；SOP 取貨門市不能改",
    "熱門無預約結帳(參考)": "不用做（現貨或 VIP，不算異常）",
    "調撥機被當現貨賣": "盤點這台機器在不在；不在就回報區經理／總部，那位等遞補的客人要重新安排機器",
    "DOA 待客服改狀態": "不用門市處理，等客服改回「已預約」，追蹤即可",
    "未掃條碼": "下次取貨結帳要在 F10 第 7 欄掃預約單號，系統才會自動改狀態",
}
KIND_ORDER = {k: n for n, k in enumerate([
    "A 結帳未改", "B 改狀態無消費", "疑似代取", "銷退仍已取貨", "掃錯規格", "重複掃同一張",
    "放棄未遞補(有人排隊)", "放棄需到店確認", "放棄逾24小時未遞補", "逾期未取", "清點單機器不足",
    "庫存 確定賣了沒改", "庫存 改了沒賣", "庫存 到貨沒進EPB", "放棄未遞補(提醒轉現貨)",
    "條碼欄打錯", "非本人取機", "跨店取機", "熱門 標現貨", "DOA 待客服改狀態", "放棄待系統遞補",
    "熱門款轉了現貨",
    "熱門無預約結帳(參考)", "到店還沒配", "未掃條碼"])}
KIND_ORDER["熱門機在店沒配人"] = KIND_ORDER["清點單機器不足"] + 0.5
KIND_ORDER["熱門機轉現貨沒賣掉"] = -0.4
KIND_ORDER["調撥機被當現貨賣"] = -0.5          # 最嚴重：機器已經不在，等遞補的人會撲空


def hot_records(rows_h, sku_names):
    recs = []
    for h in rows_h:
        st = BY_CODE[h["loc"]]
        base = dict(region=st[3], store=st[2], idx=STORES.index(st), order="", vip="", name="", sku=h["sku"],
                    product=sku_names.get(h["sku"], "").replace("預約｜", ""), sale_store="", sale_time="",
                    doc="", scene="", batch="",
                    status=f"{'熱門' if h['hot'] else '非熱門'}｜EPB庫存 {h['onhand']}／預約在店 {h['in_store']}"
                           + (f"／本店排隊 {h['queue']} 人" if h.get("queue") else ""))
        extra = f"；另有 {h['recent']} 台 {SETTLE_MIN} 分鐘內剛結帳，暫不判定" if h["recent"] else ""
        head = f"EPB 賣出 {h['sold']} 台、預約已取貨 {h['picked']} 張"
        base["urgency"] = "🟡 要確認"
        if h["sure"]:
            recs.append(dict(base, kind="庫存 確定賣了沒改", qty=h["sure"],
                             diff=f"EPB 賣出 {h['sold']} ＞ 系統已取貨 {h['picked']}，差 {h['sure']} 台（未配的機器也不夠）", todo=f"{head}，未配的機器也不夠 → 確定有 {h['sure']} 張預約"
                             "已賣出卻沒改：找出來改已取貨，否則到期會遞補錯",
                             note=("；".join(h["hint"]) if h.get("hint") else
                                   "下方 A 列可能就是；A 找不到的多半是代取，請對預約區實體機器與名單") + extra))
        if h["over"]:
            recs.append(dict(base, kind="庫存 改了沒賣", qty=h["over"],
                             diff=f"系統已取貨 {h['picked']} ＞ EPB 賣出 {h['sold']}，差 {h['over']} 台", todo=f"{head} → 有 {h['over']} 張沒結帳卻改已取貨："
                             "沒取就改回原狀態，有取就補結帳", note="下方 B／疑似代取 列可能就是" + extra))
        if h["no_inbound"]:
            recs.append(dict(base, kind="庫存 到貨沒進EPB", qty=h["no_inbound"],
                             diff=f"系統算已到店 {h['alloc']} 台 ＞ EPB 進貨 {h['inbound']} 台，差 {h['no_inbound']} 台", todo=f"預約有 {h['alloc']} 台算已到店，EPB 只進貨 {h['inbound']} 台"
                             f" → 有 {h['no_inbound']} 台預約顯示到貨但 EPB 沒進貨：確認實體機器與調撥入庫",
                             note=""))
        short = h["sheet"] - (h["onhand"] + max(h["recent"], 0))
        if short > 0:
            recs.append(dict(base, kind="清點單機器不足", qty=short,
                             diff=f"清點單 {h['sheet']} 張 ＞ EPB 庫存 {h['onhand']} 台，差 {short} 台",
                             todo=f"清點單（已到貨＋保留）{h['sheet']} 張，EPB 庫存只有 {h['onhand']} 台 → 有 {short} 位已到貨客人沒有機器：盤點預約區、找出已取機沒改的",
                             note="待取貨庫存清點單可在預約後台匯出核對"))
        if h["remark_hot"]:
            recs.append(dict(base, kind="熱門 標現貨", qty=h["remark_hot"],
                             diff=f"熱門機用「現貨」備註賣出 {h['remark_hot']} 台（SOP 不能賣現貨）", todo=f"POS 備註寫「現貨」賣出 {h['remark_hot']} 台，賣出當下這個規格已是熱門，依 SOP 不能賣現貨"
                             " → 確認是不是其實是預約機（是的話該張預約要改已取貨）", note=""))
        # 轉現貨會把機器移出遞補池，那張預約結案成「放棄(已遞補)」，系統之後不會再把它排進跨店遞補。
        # 規格後來被列為熱門時，這些機器就等於憑空消失 → 還沒賣掉的要抓回來，已經賣掉的至少要讓總部知道。
        left = min(h["shopsold_q"] - h["shopsold"], max(h["onhand"], 0))   # 還在店裡的不可能比 EPB 庫存多
        hs = hot_since(h["sku"])
        since_txt = f"；這個規格 {hs:%m/%d %H:%M} 才列為熱門" if hs else ""
        why = (f"；全區還有 {h['queue_all']} 人在排這個規格，但轉現貨的那幾張預約已經結案，系統不會再把這些機器排進跨店遞補"
               if h.get("queue_all") else "")
        if h["hot"] and left > 0:
            recs.append(dict(base, kind="熱門機轉現貨沒賣掉", qty=left, urgency="🔴 今天處理",
                             diff=f"{left} 台熱門機已轉現貨、還沒賣掉，機器還在店裡（本店 {h['queue']} 人排隊）",
                             todo=f"店裡有 {left} 台熱門機已經轉成現貨但還沒賣出 → 先不要賣，回報區經理／總部，"
                                  "看要不要收回重新配給排隊的客人",
                             note=f"依據是預約梯次尾端的「-ShopSold」（本店這個規格共轉了 {h['shopsold_q']} 張）" + since_txt + why))
        # 「熱門款轉了現貨」(已賣掉的) 使用者 10/3 決定不列：追不回來、不算異常，列出來只會讓清單更亂
        if h["unalloc"] and h["hot"]:
            if h.get("queue"):
                recs.append(dict(base, kind="熱門機在店沒配人", qty=h["unalloc"], urgency="🔴 今天處理",
                                 diff=f"店內 {h['unalloc']} 台熱門機沒配給任何人，本店還有 {h['queue']} 人在排隊",
                                 todo=f"EPB 有 {h['unalloc']} 台熱門機在店，預約系統沒有配給任何人，本店還有 {h['queue']} 位在排隊"
                                      " → 多半是調撥／遞補進來的機器沒按到貨（按了系統才會把該位客人改成已到貨）；"
                                      "若不是本店要的，請回報總部重配",
                                 note="可對 EPB 收貨單的「我方參考」：空白＝現貨、寫梯次＝總倉配貨、寫調撥單號＝別店調來的"))
                continue
            recs.append(dict(base, kind="到店還沒配", qty=h["unalloc"], urgency="⚪ 參考",
                             diff=f"EPB 有 {h['unalloc']} 台還沒配給預約客人（本店沒人排隊）", todo=f"EPB 有 {h['unalloc']} 台已進貨、還沒配給預約客人（不是異常）",
                             note=""))
    return recs


def write_hot(ws, rows_h, sku_names):
    head = ["區", "門市", "存貨代碼", "品名", "EPB進貨(不含VIP)", "已配給預約(已到店)", "未按到貨(已配貨等,可能在途)",
            f"EPB賣出(滿{SETTLE_MIN}分,不含現貨)", "預約已取貨", "狀態差(賣出−取貨)", "未按到貨但機器已進EPB",
            "現貨賣出(備註/轉現貨/現貨機)", "未配淨額(進貨−已配−已到已配貨−現貨)", "確定賣了沒改", "可能賣了沒改(或未配機直接賣)",
            "改了沒賣", "預約到貨EPB沒進貨", "到店還沒配", f"{SETTLE_MIN}分鐘內剛結帳", "EPB庫存(不含VIP)", "預約在店", "熱門款",
            "清點單數量(已到貨+保留)", "到店後放棄/取消未遞補", "EPB庫存−清點單",
            "總部VIP機(Chester/HQ VIP)", "VIP已結帳", "現貨進貨(我方參考空白)", "現貨已結帳", "調撥出", "跨店取機(本店賣出)",
            "轉現貨張數(ShopSold)", "備註現貨賣出", "可配的未配預約機"]
    ws.append(head)
    for c in ws[1]:
        c.font, c.fill = BOLD, HEAD_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    bad_fill = PatternFill("solid", fgColor="FCE4D6")
    for i, h in enumerate(sorted(rows_h, key=lambda h: (BY_CODE[h["loc"]][3], STORES.index(BY_CODE[h["loc"]]), h["sku"])),
                          start=2):
        st = BY_CODE[h["loc"]]
        ws.append([st[3], st[2], h["sku"], sku_names.get(h["sku"], "").replace("預約｜", ""),
                   h["inbound"], h["alloc"], h["cfg"], h["sold"], h["picked"], f"=H{i}-I{i}",
                   f"=MIN(MAX(E{i}-F{i}-L{i},0),G{i})", h["stock"], f"=E{i}-F{i}-L{i}-K{i}",
                   f"=MAX(J{i}-AH{i},0)", f"=MIN(MAX(J{i},0),AH{i})", f"=MAX(-J{i},0)",
                   f"=MAX(-M{i}-AD{i},0)", f'=IF(V{i}="熱門",AH{i}-O{i},0)', h["recent"], h["onhand"], h["in_store"],
                   "熱門" if h["hot"] else "非熱門", h["sheet"], h["dead"], f"=T{i}-W{i}", h["vip"], h["vip_sold"],
                   h["spot"], h["spot_sold"], h["moved_out"], h["cross"],
                   # 「還有沒有沒人認領的預約機」跟「進貨夠不夠」是兩個問題，分母不一樣：
                   # 這欄扣掉所有已離開預約池的機器（轉現貨的張數、現貨進貨、備註現貨），不管賣掉沒有
                   h["shopsold_q"], h["remark"], f"=MAX(E{i}-F{i}-AF{i}-AB{i}-AG{i}-K{i},0)"])
        if h["sure"] or h["over"] or h["no_inbound"]:
            for c in ws[i]:
                c.fill = bad_fill
    for i, w in enumerate([6, 10, 11, 34] + [10] * 30, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 60
    ws.freeze_panes = "E2"
    ws.auto_filter.ref = ws.dimensions


COLS = [("區", "region", 6), ("門市", "store", 10), ("急", "urgency", 9), ("類型", "kind", 15), ("台數", "qty", 6),
        ("差異在哪", "diff", 32), ("怎麼處理", "todo", 34),
        ("略過(填V)", "skip", 9), ("略過原因", "skip_reason", 20),
        ("預約狀態", "status", 10), ("預約單號", "order", 24), ("會員編號", "vip", 13), ("姓名", "name", 7),
        ("存貨代碼", "sku", 11), ("品名", "product", 34), ("EPB結帳門市", "sale_store", 11),
        ("結帳時間", "sale_time", 11), ("EPB單號", "doc", 32), ("備註", "note", 50), ("場次", "scene", 8),
        ("梯次", "batch", 24)]


def write_list(ws, records, with_groups):
    ws.append([c[0] for c in COLS])
    for c in ws[1]:
        c.font, c.fill = BOLD, HEAD_FILL
        c.alignment = Alignment(horizontal="center")
    grp_fill = PatternFill("solid", fgColor="FFF2CC")
    records = sorted(records, key=lambda x: (x["region"], x["idx"], KIND_ORDER[x["kind"]], x["sku"]))
    cur = None
    for rec in records:
        if with_groups and rec["store"] != cur:
            cur = rec["store"]
            cnt = lambda kind: sum(1 for x in records if x["store"] == cur and x["kind"] == kind)
            parts = [f"{k} {cnt(k)}" for k in KIND_ORDER if cnt(k) and k != "到店還沒配"]
            ws.append([rec["region"], cur, "", "▶ 本店", "", "、".join(parts) or "只有參考項目"])
            ws.merge_cells(start_row=ws.max_row, start_column=6, end_row=ws.max_row, end_column=len(COLS) - 3)
            for c in ws[ws.max_row]:
                c.font, c.fill = BOLD, grp_fill
        ws.append([rec.get(c[1], "") for c in COLS])
    for i, (_, _, w) in enumerate(COLS, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    wrap = {i for i, c in enumerate(COLS, 1) if c[1] in ("todo", "diff", "product", "note", "skip_reason")}
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(vertical="top", wrap_text=c.column in wrap)
    ws.freeze_panes = "G2"
    ws.auto_filter.ref = ws.dimensions


def next_update_dt(now=None):
    now = now or datetime.now()
    h = next((h for h in UPDATE_HOURS if h > now.hour), None)
    day = now if h else now + timedelta(days=1)
    return day.replace(hour=h or UPDATE_HOURS[0], minute=0, second=0, microsecond=0)


def next_update(now=None):
    now = now or datetime.now()
    t = next_update_dt(now)
    return f"{t.hour}:00" if t.date() == now.date() else f"明天 {t.hour}:00"


def ignore_key(kind, store, sku, order):
    return (str(kind or ""), str(store or ""), str(sku or ""), str(order or ""))


IGNORE_HEAD = ["類型", "門市", "存貨代碼", "預約單號", "略過時台數", "略過原因", "加入時間"]


def load_ignores():
    """略過清單：key → (略過時台數, 原因)。要取消略過，直接刪掉那一列。"""
    if not IGNORE_FILE.exists():
        return {}
    ws = load_workbook(IGNORE_FILE).active
    out = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r and r[0]:
            out[ignore_key(*r[:4])] = (int(r[4] or 0), r[5] or "")
    return out


def _harvest_rows(rows, col, ignores):
    added = 0
    for r in rows[1:]:
        marked = str(r[col["略過(填V)"]] or "").strip() in SKIP_MARKS or str(r[col["略過原因"]] or "").strip()
        if not marked or r[col["類型"]] == "▶ 本店":  # 打 V 或只寫原因都算略過
            continue
        key = ignore_key(r[col["類型"]], r[col["門市"]], r[col["存貨代碼"]], r[col["預約單號"]])
        qty = int(r[col["台數"]] or 0)
        if key not in ignores or ignores[key][0] < qty:
            ignores[key] = (qty, r[col["略過原因"]] or "")
            added += 1
    return added


def harvest_skips():
    """把上一份報表「待處理清單」「區經理參考」裡填 V 的列，加進略過清單。"""
    files = sorted(OUT_DIR.glob("iPhone18預約取貨核對_*.xlsx"), key=lambda f: f.stat().st_mtime, reverse=True)
    if not files:
        return 0
    wb = load_workbook(files[0])
    ignores = load_ignores()
    added = 0
    for sheet in ("待處理清單", "區經理參考"):
        if sheet not in wb.sheetnames:
            continue
        rows = list(wb[sheet].iter_rows(values_only=True))
        col = {name: i for i, name in enumerate(rows[0])}
        if "略過(填V)" in col:
            added += _harvest_rows(rows, col, ignores)
    if added:
        save_ignores(ignores)
    return added


def save_ignores(ignores):
    """寫回略過清單（保留每列原本的加入時間）。"""
    wb_out = Workbook()
    ws = wb_out.active
    ws.append(IGNORE_HEAD)
    for c in ws[1]:
        c.font, c.fill = BOLD, HEAD_FILL
    old = {ignore_key(*x[:4]): x[6] for x in (load_workbook(IGNORE_FILE).active.iter_rows(min_row=2, values_only=True)
                                               if IGNORE_FILE.exists() else []) if x and x[0]}
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    for key, (qty, reason) in ignores.items():
        ws.append(list(key) + [qty, reason, old.get(key, now)])
    for i, w in enumerate([16, 10, 11, 24, 10, 30, 16], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    wb_out.save(IGNORE_FILE)


def split_ignored(records, ignores):
    """台數沒有超過略過時的台數才略過；A、B 以預約單號略過。"""
    active, skipped = [], []
    for rec in records:
        hit = ignores.get(ignore_key(rec["kind"], rec["store"], rec["sku"], rec["order"]))
        if hit and rec["qty"] <= hit[0]:
            skipped.append(dict(rec, skip="V", skip_reason=hit[1]))
        else:
            active.append(rec)
    return active, skipped


SUMMARY = [("A 結帳未改", ["A 結帳未改"]), ("B 改狀態無消費", ["B 改狀態無消費"]), ("疑似代取", ["疑似代取"]),
           ("銷退仍已取貨", ["銷退仍已取貨"]), ("掃錯規格／重複掃", ["掃錯規格", "重複掃同一張"]),
           ("放棄未遞補(有人排隊)", ["放棄未遞補(有人排隊)"]),
           ("放棄需到店確認", ["放棄需到店確認"]), ("放棄逾24小時未遞補", ["放棄逾24小時未遞補"]), ("逾期未取", ["逾期未取"]),
           ("清點單機器不足(台)", ["清點單機器不足"]),
           ("庫存 確定賣了沒改(台)", ["庫存 確定賣了沒改"]), ("庫存 改了沒賣(台)", ["庫存 改了沒賣"]),
           ("庫存 到貨沒進EPB(台)", ["庫存 到貨沒進EPB"]), ("提醒轉現貨", ["放棄未遞補(提醒轉現貨)"]),
           ("其他違反SOP(打錯條碼/非本人/跨店)", ["條碼欄打錯", "非本人取機", "跨店取機"]),
           ("熱門機在店沒配人(台)", ["熱門機在店沒配人"]),
           ("調撥機被當現貨賣(台)", ["調撥機被當現貨賣"]),
           ("熱門機轉現貨沒賣掉(台)", ["熱門機轉現貨沒賣掉"]),
           ("參考：熱門款轉了現貨(台)", ["熱門款轉了現貨"]),
           ("參考：現貨/VIP", ["熱門無預約結帳(參考)"]), ("參考：到店還沒配(台)", ["到店還沒配"])]


# 門市只看「要立刻去預約系統改」的（使用者 10/3：像忘記改已取貨這種），其餘移到「區經理參考」分頁
MGR_KINDS = {"庫存 改了沒賣", "庫存 到貨沒進EPB", "清點單機器不足", "非本人取機", "跨店取機", "DOA 待客服改狀態",
             "放棄待系統遞補", "放棄逾24小時未遞補", "熱門 標現貨", "到店還沒配"}


def for_store(rec):
    if rec["kind"] == "條碼欄打錯":            # 狀態已手動改對的只需宣導，不用門市再改
        return "已手動改對" not in (rec.get("note") or "")
    return rec["kind"] not in MGR_KINDS


def write_members(ws, members):
    head = ["會員編號", "標註", "預約門市", "預約總數", "已取貨", "結帳淨台數", "開著的預約", "預約明細", "結帳明細"]
    ws.append(head)
    for c in ws[1]:
        c.font, c.fill = BOLD, HEAD_FILL
    for m in sorted(members, key=lambda m: (m["tag"], m["stores"])):
        ws.append([m["vip"], m["tag"], m["stores"], m["n_resv"], m["picked"], m["bought"], m["opens"], m["resv"], m["sales"]])
    for i, w in enumerate([13, 14, 16, 8, 8, 10, 10, 70, 60], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(vertical="top", wrap_text=c.column in (8, 9))
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions


# 庫存對帳類的差異，若同店同規格已經有單筆列出來（同一件事），就只列解釋不了的台數
# 同一個錯誤會在庫存對帳的好幾個角度同時冒出來（例：一張該改回去的「已取貨」，
# 會讓「改了沒賣」「到貨沒進EPB」「清點單機器不足」各多一台）→ 只列解釋不了的台數
STATUS_ERRORS = ("A 結帳未改", "B 改狀態無消費", "疑似代取", "銷退仍已取貨", "跨店取機", "DOA 待客服改狀態")
EXPLAINS = {"庫存 確定賣了沒改": ("A 結帳未改", "調撥機被當現貨賣"),
            "庫存 改了沒賣": ("B 改狀態無消費", "疑似代取", "銷退仍已取貨", "跨店取機", "DOA 待客服改狀態"),
            "庫存 到貨沒進EPB": STATUS_ERRORS,
            "清點單機器不足": STATUS_ERRORS + ("庫存 確定賣了沒改", "庫存 到貨沒進EPB")}
# 先算完的庫存類可以拿來解釋後面的（清點單不足常常就是「確定賣了沒改」那台）
STOCK_ORDER = ["庫存 確定賣了沒改", "庫存 改了沒賣", "庫存 到貨沒進EPB", "清點單機器不足"]
# 做法一模一樣、門市一次處理的，同店同規格併成一列（台數相加、單號寫進備註）
MERGE_KINDS = {"放棄需到店確認", "放棄待系統遞補", "放棄未遞補(提醒轉現貨)", "放棄逾24小時未遞補", "逾期未取", "熱門無預約結帳(參考)"}


def consolidate(recs):
    """同一個問題只列一次：同一張預約（沒預約的依結帳單）只留最重要的類型，其他類型併進備註；
       庫存對帳類若已被同店同規格的單筆（A／B／銷退）解釋，只列解釋不了的台數。"""
    stock = [r for r in recs if r["kind"] in EXPLAINS]
    groups = defaultdict(list)
    for r in recs:
        if r["kind"] not in EXPLAINS:
            groups[r["order"] or r["doc"] or id(r)].append(r)
    out = []
    for g in groups.values():
        g.sort(key=lambda r: KIND_ORDER[r["kind"]])
        main = dict(g[0])
        extra = [f"{r['kind']}（{r['note']}）" for r in g[1:] if r["kind"] != main["kind"]]
        if extra:
            main["note"] = (main["note"] + "｜" if main["note"] else "") + "另：" + "；".join(extra)
        out.append(main)
    for r in sorted(stock, key=lambda r: STOCK_ORDER.index(r["kind"]) if r["kind"] in STOCK_ORDER else 99):
        done = sum(x["qty"] for x in out if x["store"] == r["store"] and x["sku"] == r["sku"] and x["kind"] in EXPLAINS[r["kind"]])
        left = r["qty"] - done
        if left > 0:
            out.append(dict(r, qty=left, note=(f"其中 {done} 台已列在本店{'／'.join(EXPLAINS[r['kind']])}；" if done else "") + r["note"]))
    return merge_same(out)


def merge_same(recs):
    """同店同規格、處理方式一樣的（等遞補、轉現貨…）併成一列，台數相加、單號列在備註。"""
    groups, out = defaultdict(list), []
    for r in recs:
        if r["kind"] in MERGE_KINDS:
            groups[(r["store"], r["sku"], r["kind"])].append(r)
        else:
            out.append(r)
    for g in groups.values():
        m = dict(g[0], qty=len(g))
        if len(g) > 1:
            orders = "、".join(x["order"] for x in g if x["order"])
            m["order"] = ""          # 併成一列後，略過改用「門市＋規格＋類型」
            m["status"] = g[0]["status"] if len({x["status"] for x in g}) == 1 else "多筆"
            m["name"] = m["vip"] = m["batch"] = ""
            m["note"] = f"{m['note']}｜共 {len(g)} 張：{orders}"
        out.append(m)
    return out


def write_xlsx(resv, items, unscanned, rows_h, sku_names, cutoff, members, out):
    wb = Workbook()
    recs = consolidate([record(k, r, x, note, sku_names) for k, r, x, note in items] + hot_records(rows_h, sku_names))
    # 同店同規格同時有「A 結帳未改」和「B 改狀態無消費」：多半是改錯張（甲買走機器、門市卻把乙改成已取貨），
    # 兩列互相指到對方，門市才知道要一起看（10/3 新莊宏匯 07313604 SA02******／SA00******）
    for b_ in [x for x in recs if x["kind"] == "B 改狀態無消費"]:
        for a_ in [x for x in recs if x["kind"] == "A 結帳未改" and x["store"] == b_["store"] and x["sku"] == b_["sku"]]:
            a_["note"] += (f"；⚠️ 同店同規格 {b_['vip']} 的預約 {b_['order']} 被改成已取貨卻沒有結帳 → 可能改錯張："
                           f"確認實際取機的是 {a_['vip']} 就把他這張改已取貨、{b_['vip']} 那張改回已到貨")
            b_["note"] += (f"；⚠️ 同店同規格 {a_['vip']} 在 {a_['sale_time']} 結帳、他的預約 {a_['order']} 卻沒改 → 可能改錯張"
                           f"（或 {a_['vip']} 幫忙代取）")
    active, skipped = split_ignored(recs, load_ignores())
    print(f"略過 {len(skipped)} 筆（略過清單：{IGNORE_FILE.name}）")
    recs_u = [record("未掃條碼", x["LINK"], x, (f"掃到同會員另一張預約（{x['MIXUP']}）的條碼" if x.get("MIXUP")
                                                  else f"以{x['HOW']}對到"), sku_names) for x in unscanned]

    ws_sum = wb.active
    ws_sum.title = "摘要"
    write_list(wb.create_sheet("待處理清單"), [x for x in active if for_store(x)], with_groups=True)
    write_list(wb.create_sheet("區經理參考"), [x for x in active if not for_store(x)], with_groups=True)
    write_guide(wb.create_sheet("說明"), cutoff)
    write_hot(wb.create_sheet("庫存對帳"), rows_h, sku_names)
    write_list(wb.create_sheet("未掃預約條碼"), recs_u, with_groups=False)
    write_members(wb.create_sheet("會員核對"), members)
    write_list(wb.create_sheet("已略過"), skipped, with_groups=False)

    ws_sum.append([f"iPhone 18 Pro 預約取貨核對（產出 {datetime.now():%Y-%m-%d %H:%M}；"
                   f"下次更新 {next_update()}；結帳未改只判定 {cutoff[8:10]}:{cutoff[10:12]} 以前結帳的）"])
    ws_sum["A1"].font = Font(bold=True, size=13)
    ws_sum.append([])
    ws_sum.append(["區", "門市", "預約已取貨筆數"] + [h for h, _ in SUMMARY]
                  + ["未掃預約條碼(筆)", "未掃率", "已略過(筆)"])
    for c in ws_sum[3]:
        c.font, c.fill = BOLD, HEAD_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws_sum.row_dimensions[3].height = 60
    picked_cnt = defaultdict(int)
    for r in resv:
        if (r.get("statusName") or "").startswith(PICKED):
            picked_cnt[r["shopName"].strip()] += 1
    first = 4
    for st in STORES:
        n = ws_sum.max_row + 1
        q = lambda kinds: "=" + "+".join(f'SUMIFS({sh}!$E:$E,{sh}!$B:$B,B{n},{sh}!$D:$D,"{k}")'
                                         for sh in ("待處理清單", "區經理參考")
                                         for k in kinds)
        u = get_column_letter(3 + len(SUMMARY) + 1)
        ws_sum.append([st[3], st[2], picked_cnt[st[1]]] + [q(k) for _, k in SUMMARY]
                      + [f"=COUNTIFS(未掃預約條碼!$B:$B,B{n})", f'=IF(C{n}=0,"",{u}{n}/C{n})',
                         f"=COUNTIFS(已略過!$B:$B,B{n})"])
    last_row = ws_sum.max_row
    rate_col = get_column_letter(3 + len(SUMMARY) + 2)
    cols = [get_column_letter(i) for i in range(3, 3 + 1 + len(SUMMARY) + 3)]
    for region in ("北一", "北二"):
        n = ws_sum.max_row + 1
        ws_sum.append(["", f"{region}小計"] + [
            f'=SUMIF($A${first}:$A${last_row},"{region}",{c}${first}:{c}${last_row})' for c in cols])
        ws_sum[f"{rate_col}{n}"] = f'=IF(C{n}=0,"",{get_column_letter(3 + len(SUMMARY) + 1)}{n}/C{n})'
        for c in ws_sum[n]:
            c.font = BOLD
    n = ws_sum.max_row + 1
    ws_sum.append(["", "北區合計"] + [f"=SUM({c}{first}:{c}{last_row})" for c in cols])
    ws_sum[f"{rate_col}{n}"] = f'=IF(C{n}=0,"",{get_column_letter(3 + len(SUMMARY) + 1)}{n}/C{n})'
    for c in ws_sum[n]:
        c.font = BOLD
    ws_sum.append([])
    ws_sum.append(["▶ 逐店處理請看「待處理清單」：先看「急」欄（🔴 今天處理 → 🟡 要確認 → ⚪ 參考），"
                   "「差異在哪」講兩邊對不起來的地方、「怎麼處理」就是動作。判斷規則依 9/17 預約系統 SOP，見「說明」頁。"
                   "門市不用動手的（庫存推算、SOP 宣導、等系統／客服）在「區經理參考」。摘要數字兩頁合計。"]) 
    ws_sum.append([f"▶ 特例不想每次出現：在「待處理清單」該列的「略過(填V)」填 V 或寫原因（只寫原因也算），下次更新（每天 {HOURS_TXT} 點）就會移到「已略過」。"
                   "要取消略過，把「已略過」那一列的 V 清掉。"])
    for c in ws_sum[rate_col]:
        c.number_format = "0%"
    for i, w in enumerate([6, 12, 11] + [11] * (len(SUMMARY) + 3), 1):
        ws_sum.column_dimensions[get_column_letter(i)].width = w
    wb.save(out)


def write_guide(ws, cutoff):
    title = Font(bold=True, size=13)
    wrap = Alignment(wrap_text=True, vertical="top")
    for col, w in zip("ABC", [24, 44, 44]):
        ws.column_dimensions[col].width = w
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    def line(*cells, font=None, fill=None, merge=False, height=None):
        ws.append(list(cells))
        for c in ws[ws.max_row]:
            c.alignment = wrap
            if font:
                c.font = font
            if fill and c.value is not None:
                c.fill = fill
        if merge:
            ws.merge_cells(start_row=ws.max_row, start_column=1, end_row=ws.max_row, end_column=3)
        if height:
            ws.row_dimensions[ws.max_row].height = height

    def section(text):
        ws.append([])
        line(text, font=title, merge=True)

    line("iPhone 18 Pro 預約取貨 × EPB 銷售核對　說明", font=Font(bold=True, size=15), merge=True)
    line("目的：確保「結帳了就是已取貨、沒結帳就不是已取貨」，狀態錯了會遞補錯人、晚上預約區庫存也不會對。"
         f"範圍：北區 13 店，iPhone 18 Pro 限定場＋全會員場，EPB POS 結帳自 {SALES_FROM} 起（含銷退）。"
         "規則依「業務_新預約系統操作流程 20260917」SOP。", merge=True, height=48)

    section("一、依 SOP，每筆結帳怎麼對到預約")
    for t in [
        "SOP：取貨結帳時在 POS 按「序號(F10)」，在「7 預約條碼」掃預約單號，完成交易 5 分鐘內系統自動改「已取貨」。",
        "所以程式依序用：① 預約條碼（POS 第 7 欄）② 備註裡的預約單號 ③ 備註裡的會員編號（例：代取 SA00******）"
        "④ 結帳會員＋規格，把每一筆結帳對到一張預約。都對不到、備註寫「現貨」的才算現貨。",
        f"結帳未滿 {SETTLE_MIN} 分鐘的先不判定「結帳未改」（本次判定 {cutoff[8:10]}:{cutoff[10:12]} 以前結帳的）。",
        "同一支手機在 EPB 有三組代碼（07313…、073S13…、905029…），程式依型號一律換成預約系統的代碼再比對。",
        "總部 VIP 機不列入核對：總倉收貨單「我方參考」含 Chester 或 VIP 的，依每店每規格台數，從進貨、庫存扣掉；"
        "同店同規格對不到預約的結帳，在這個台數以內算 VIP 取機，不列任何檢核。",
        "門市轉現貨不列入核對：門市用「現貨販售處理」轉現貨後，該張預約的梯次尾端會加「-ShopSold」。"
        "依每店每規格的 ShopSold 台數，同店同規格對不到預約的結帳在這個台數以內算現貨賣出，不會被判成疑似代取或賣了沒改。",
        "現貨進貨也不列入核對：EPB 收貨單「我方參考」空白的就是現貨（預約配貨會寫梯次「2026 iPhone 18 09xx」、"
        "VIP 寫 Chester／HQ VIP、店間調撥寫調撥單號），比照 VIP 機從進貨、庫存扣掉，對不到預約的結帳在這個台數以內算現貨賣出。"
        "VIP 機與現貨機都只算「已經進到門市庫存異動」的；總倉出了貨、門市還沒收的在途機器不能先扣，否則進貨會憑空少一台。",
    ]:
        line(t, merge=True, height=34)

    section("二、檢核邏輯（七種情況）")
    line("情況", "怎麼判斷", "報表類型", font=BOLD, fill=HEAD_FILL)
    for row in [
        ("一、結帳沒連動改狀態", "沒掃預約條碼。買方完全沒有任何預約＝現貨或 VIP；買方有預約卻沒掃＝狀態要手動改。",
         "A 結帳未改（預約還開著）／未掃預約條碼分頁（已手動改對）／熱門無預約結帳(參考)"),
        ("二、結帳一台連動，但改到錯的預約", "查會員在預約網站的總預約數：同會員另一張→程式自動改對；別人的（入錯會員）→銷退不會改回狀態。",
         "銷退仍已取貨／B 改狀態無消費／會員核對分頁"),
        ("三、結帳兩台都掃同一張", "同會員兩台掃同一張單號→程式自動對回另一張；否則列出。", "重複掃同一張／掃錯規格／會員核對分頁"),
        ("四、有改狀態沒消費（代取）", "同店同規格有一筆對不到預約的結帳（沒打會員或別人會員）。非本人取機只能購買、不能累積會員。", "疑似代取"),
        ("五、有改狀態沒消費", "同店也找不到可能的代取結帳。", "B 改狀態無消費（異常）"),
        ("六、放棄／逾期，沒遞補、沒人排隊", "門市已按到貨後放棄或取消、沒有（已遞補），本店沒有已預約。非熱門提醒轉現貨；熱門等系統跨店遞補。剛到期 12 小時內的先當參考（系統還沒跑遞補）。",
         "放棄未遞補(提醒轉現貨)／放棄待系統遞補"),
        ("放棄(需到店確認)", "9/30 起新狀態：客人在已配貨時放棄，機器到店後門市要按「門市確認」系統才會往後遞補。沒有取貨期限，"
         "用同店同規格「未按到貨但機器已進 EPB」的台數判斷機器到了沒，到了幾台列幾張。", "放棄需到店確認"),
        ("七、放棄／逾期，還沒遞補下一位", "同上，但本店還有人排隊。排序：會員等級→非再購→登記時間，列出下一位。", "放棄未遞補(有人排隊)"),
        ("熱門放棄卡住", "熱門、本店沒人排隊、取貨期限已過 24 小時以上，而且不在預約系統目前的「跨店調撥建議」清單裡 → 系統不會自己補，要人工處理。",
         "放棄逾24小時未遞補"),
        ("逾期未取", "過了期限還是已到貨／保留。期限一律用預約系統的「預計取機時間」（保留門市可以延，不是固定＋2 天）。到期時系統通常會自動改放棄，沒改到的才列。", "逾期未取"),
        ("待取貨庫存清點單", "清點單＝已到貨＋保留。EPB 庫存應 ≥ 清點單；差額＝放棄未遞補＋到店還沒配＋狀態錯誤。", "庫存對帳分頁 W～Y 欄／清點單機器不足"),
        ("熱門機在店沒配人", "EPB 有熱門機、預約系統沒配給任何人，而且本店還有「已預約」在排隊。"
         "「預約在店」算的是按過到貨的張數，跟「有沒有人在排隊」是兩回事，所以這兩個數字要一起看。",
         "熱門機在店沒配人"),
    ]:
        line(*row, height=62)
        ws["A" + str(ws.max_row)].font = BOLD

    section("二之二、各類型的差異在哪、門市要做什麼")
    line("類型", "差異在哪", "門市要做什麼", font=BOLD, fill=HEAD_FILL)
    for kind in KIND_ORDER:
        line(kind, DIFF.get(kind, ""), TODO.get(kind, {"庫存 確定賣了沒改": "找出已取機的預約改已取貨",
                                       "庫存 改了沒賣": "找出沒結帳卻改已取貨的（通常同店會有 B 或疑似代取）",
                                       "庫存 到貨沒進EPB": "確認實體機器與 EPB 調撥入庫",
                                       "到店還沒配": "不用處理，等系統配貨（參考）"}.get(kind, "")), height=34)
        ws["A" + str(ws.max_row)].font = BOLD

    section("三、分頁")
    for t in [
        "• 摘要：各店各類筆數，最右邊有「未掃率」＝沒在 F10 第 7 欄掃預約條碼的比例。"
        "　• 待處理清單：只放門市要立刻去預約系統改的（忘記改已取貨、改錯、要遞補…），先看「急」欄，再看「差異在哪」「怎麼處理」（門市看這頁）。　• 區經理參考：庫存推算、SOP 宣導、等系統／客服處理的，門市不用動手。",
        "• 庫存對帳：每店每規格的 EPB 進貨／賣出 vs 預約配貨／已取貨，是總量的交叉檢查。",
        "• 未掃預約條碼：狀態是對的，但結帳沒掃預約條碼的單（違反 SOP，會讓系統無法自動改狀態），給店長教育用。",
        "• 會員核對：預約網站總預約數／已取貨 vs 結帳台數不一致的會員（含排隊中已買到，只供參考）。",
        "• 已略過：標記為特例的項目。",
    ]:
        line(t, merge=True, height=34)

    section("四、略過特例")
    for t in [
        "• 確認是特例、不用每次出現的：在「待處理清單」那一列的「略過(填V)」填 V，或「略過原因」寫原因（只寫原因也算略過）。",
        f"• 下次更新（每天 {HOURS_TXT} 點）會自動收進略過清單，該列移到「已略過」分頁，不再算進摘要；原因晚點在「已略過」補寫也可以。",
        "• 要取消略過：把「已略過」那一列的 V 清掉，下次更新就會回到待處理清單。",
        "• 有預約單號的依預約單號略過；庫存對帳類依門市＋規格＋類型略過，台數比略過時多就會重新出現。",
    ]:
        line(t, merge=True, height=34)

    section("五、注意")
    for t in [
        "• 資料是執行當下即時抓的，門市改完狀態後重跑，數字就會更新。",
        "• 姓名只留第一個字、不列電話；找人請用會員編號＋預約單號。",
    ]:
        line(t, merge=True, height=24)


def main():
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else OUT_DIR / f"iPhone18預約取貨核對_{datetime.now():%Y%m%d}.xlsx"
    n = harvest_skips()
    if n:
        print(f"從上一份報表收到 {n} 筆新的略過")
    resv, items, unscanned, rows_h, sku_names, cutoff, members = build()
    write_xlsx(resv, items, unscanned, rows_h, sku_names, cutoff, members, out)
    cnt = defaultdict(int)
    for k, *_ in items:
        cnt[k] += 1
    t = lambda k: sum(h[k] for h in rows_h if h["hot"] or k in ("sure", "over", "no_inbound"))
    print("、".join(f"{k} {cnt[k]}" for k in KIND_ORDER if cnt[k]))
    print(f"庫存對帳：確定賣了沒改 {t('sure')} 台、改了沒賣 {t('over')} 台、到貨沒進EPB {t('no_inbound')} 台")
    tags = defaultdict(int)
    for mem in members:
        tags[mem["tag"]] += 1
    print(f"未掃預約條碼 {len(unscanned)} 筆；會員核對不一致 {len(members)} 人（"
          + "、".join(f"{t} {n}" for t, n in tags.items()) + "）")
    print(out)


if __name__ == "__main__":
    main()
