#!/usr/bin/env python3
# TOOL: ARpedia、人流 日報（北二區）
"""
ARpedia 日報 — 北二區（永和、板橋誠品、西門、花蓮、板橋遠百、新莊宏匯、新店裕隆城）
輸出：ARpedia（昨日/本週/月累/年累） + 人流（昨日/月累/上月同期）
人流：有計數器店（西門/板橋遠百/新莊宏匯/新店裕隆城）用 ShopperTrak；
      無計數器店（永和/板橋誠品/花蓮）用 EPB 成交筆數公式。
"""
import subprocess, sys, json
from datetime import date, timedelta
from pathlib import Path
from config import JAVA, EPB_CP, EPB_CWD

STORES     = ['永和', '板橋誠品', '西門', '花蓮', '板橋遠百', '新莊宏匯', '新店裕隆城']
SHOP_CODES = {'永和':'9', '板橋誠品':'25', '西門':'50', '花蓮':'55',
              '板橋遠百':'63', '新莊宏匯':'64', '新店裕隆城':'68'}
SHOP_STR   = {'永和':'009', '板橋誠品':'025', '西門':'050', '花蓮':'055',
              '板橋遠百':'063', '新莊宏匯':'064', '新店裕隆城':'068'}
SHOP_IN    = "'009','025','050','055','063','064','068'"
COUNTER    = ['050', '063', '064', '068']   # 有 ShopperTrak 計數器
NO_COUNTER = {'009', '025', '055'}           # 無計數器 → 公式

# ShopperTrak 模組與帳密借用北二區週報資料夾（在該網頁設定帳密；沒有就略過人流）
ST_DIR    = Path.home() / 'Desktop' / 'north2-weekly-report'
ST_CONFIG = ST_DIR / 'local_config.json'


def epb(sql):
    r = subprocess.run(
        [JAVA, '-Dsun.net.client.defaultReadTimeout=120000',
         '-cp', EPB_CP, 'EPBReportQuery', sql, '5000'],
        capture_output=True, text=True, cwd=EPB_CWD
    )
    lines = [l for l in r.stdout.strip().split('\n') if l.strip()]
    if not lines:
        return []
    hdrs = [h.strip().upper() for h in lines[0].split('\t')]
    return [dict(zip(hdrs, row.split('\t'))) for row in lines[1:]]


def query_units(d_start: date, d_end: date) -> dict:
    ds = f"TO_DATE('{d_start}','yyyy-mm-dd')"
    de = f"TO_DATE('{d_end}','yyyy-mm-dd')"
    rows = epb(
        f"SELECT l.shop_id, "
        f"SUM(CASE WHEN l.trans_type IN ('A','H') THEN l.stk_qty "
        f"WHEN l.trans_type='E' THEN l.stk_qty ELSE 0 END) AS units "
        f"FROM poslinev_bi l "
        f"WHERE l.org_id='01' AND l.shop_id IN ({SHOP_IN}) "
        f"AND l.doc_date>={ds} AND l.doc_date<={de} "
        f"AND l.trans_type IN ('A','E','H') "
        f"AND l.brand_id='496' "
        f"GROUP BY l.shop_id ORDER BY l.shop_id"
    )
    return {str(int(r['SHOP_ID'])): int(float(r.get('UNITS', 0) or 0)) for r in rows}


def query_txn_counts(d_start: date, d_end: date) -> dict:
    """無計數器 3 店：一次查成交筆數（估算人流），回傳 {code: cnt}"""
    ds = f"TO_DATE('{d_start}','yyyy-mm-dd')"
    de = f"TO_DATE('{d_end}','yyyy-mm-dd')"
    rows = epb(
        f"SELECT l.shop_id, COUNT(DISTINCT l.doc_id) AS cnt FROM poslinev_bi l "
        f"WHERE l.org_id='01' AND l.shop_id IN ('009','025','055') "
        f"AND l.doc_date>={ds} AND l.doc_date<={de} AND l.trans_type IN ('A','H') "
        f"GROUP BY l.shop_id"
    )
    return {str(int(r['SHOP_ID'])).zfill(3): int(float(r.get('CNT', 0) or 0))
            for r in rows}


def traffic_formula(txn: int) -> int:
    return round(txn * 0.85 / 0.3)


def fetch_shoppertrak(periods):
    """回傳 {store_name: [各period總和]}；失敗回 {}（略過 ShopperTrak 人流）。"""
    import ssl
    ssl._create_default_https_context = ssl._create_unverified_context
    sys.path.insert(0, str(ST_DIR))
    try:
        import shoppertrak
    except ImportError:
        print("  ⚠️  找不到 shoppertrak 模組（北二區週報未下載？），略過 ShopperTrak 人流")
        return {}
    try:
        cfg = json.loads(ST_CONFIG.read_text())
        username = cfg['shoppertrak']['username']
        password = cfg['shoppertrak']['password']
    except Exception as e:
        print(f"  ⚠️  無法讀取 ShopperTrak 帳密（{e}），略過 ShopperTrak 人流")
        return {}

    fetch_start = min(ds for _, ds, _ in periods)
    fetch_end   = max(de for _, _, de in periods)
    try:
        print("  查詢 ShopperTrak 人流...", flush=True)
        daily_map = shoppertrak.fetch_all(
            COUNTER, fetch_start, fetch_end, username, password,
            log=lambda m: print(f"  {m}", flush=True))
    except Exception as e:
        print(f"  ⚠️  ShopperTrak 失敗：{e}，略過")
        return {}

    result = {}
    for store in STORES:
        code = SHOP_STR[store]
        if code in NO_COUNTER:
            continue
        days = daily_map.get(code, {})
        vals = []
        for _, ds, de in periods:
            total, d = 0, ds
            while d <= de:
                total += days.get(d.isoformat(), 0)
                d += timedelta(days=1)
            vals.append(total)
        result[store] = vals
    return result


def main():
    today     = date.today()
    yesterday = today - timedelta(days=1)

    days_since_sunday = (yesterday.weekday() + 1) % 7
    week_start  = yesterday - timedelta(days=days_since_sunday)
    month_start = yesterday.replace(day=1)
    year_start  = today.replace(month=1, day=1)
    last_month_start = (month_start - timedelta(days=1)).replace(day=1)
    last_month_same  = last_month_start.replace(day=min(yesterday.day,
                       (month_start - timedelta(days=1)).day))

    ar_periods = [
        (f"昨日 {yesterday.strftime('%-m/%-d')}",                                    yesterday,   yesterday),
        (f"本週 {week_start.strftime('%-m/%-d')}~{yesterday.strftime('%-m/%-d')}",   week_start,  yesterday),
        (f"月累 {month_start.strftime('%-m/%-d')}~{yesterday.strftime('%-m/%-d')}", month_start, yesterday),
        (f"年累 {year_start.strftime('%-m/%-d')}~{yesterday.strftime('%-m/%-d')}",  year_start,  yesterday),
    ]
    tr_periods = [
        (f"昨日 {yesterday.strftime('%-m/%-d')}",                                                   yesterday,        yesterday),
        (f"月累 {month_start.strftime('%-m/%-d')}~{yesterday.strftime('%-m/%-d')}",                 month_start,      yesterday),
        (f"上月同期 {last_month_start.strftime('%-m/%-d')}~{last_month_same.strftime('%-m/%-d')}", last_month_start, last_month_same),
    ]

    print(f"\nARpedia、人流 日報 北二區 ({today.strftime('%Y-%m-%d')})")
    print("=" * 80)

    # ── ARpedia 銷售 ──
    print("\n【ARpedia 銷售數量】")
    results = []
    for label, ds, de in ar_periods:
        print(f"  查詢 {label}...", flush=True)
        results.append((label, query_units(ds, de)))

    col_w  = 10
    header = f"{'門市':<7}" + "".join(f"{label:>{col_w}}" for label, _ in results)
    print("\n" + header)
    print("-" * len(header))
    totals = [0] * len(results)
    for store in STORES:
        sid = SHOP_CODES[store]
        row = f"{store:<7}"
        for i, (_, data) in enumerate(results):
            v = data.get(sid, 0)
            totals[i] += v
            row += f"{v:>{col_w},}"
        print(row)
    print("-" * len(header))
    print(f"{'合計':<7}" + "".join(f"{t:>{col_w},}" for t in totals))

    # ── 人流（來客數）──
    print("\n【人流（來客數）】")
    tr_data = fetch_shoppertrak(tr_periods)
    for label, ds, de in tr_periods:           # 無計數器 3 店：每期間一次合併查詢
        print(f"  查詢無計數器門市成交筆數 {label}...", flush=True)
        cnts = query_txn_counts(ds, de)
        for store in STORES:
            code = SHOP_STR[store]
            if code in NO_COUNTER:
                tr_data.setdefault(store, []).append(traffic_formula(cnts.get(code, 0)))

    col_w2  = 16
    tr_labels = [label for label, _, _ in tr_periods]
    header2 = f"{'門市':<7}" + "".join(f"{label:>{col_w2}}" for label in tr_labels)
    print("\n" + header2)
    print("-" * len(header2))
    tr_totals = [0] * len(tr_periods)
    for store in STORES:
        vals = tr_data.get(store)
        row = f"{store:<7}"
        for i, v in enumerate(vals or [None] * len(tr_periods)):
            if v is None:
                row += f"{'--':>{col_w2}}"
            else:
                tr_totals[i] += v
                row += f"{v:>{col_w2},}"
        print(row)
    print("-" * len(header2))
    print(f"{'合計':<7}" + "".join(f"{t:>{col_w2},}" for t in tr_totals))
    print()


if __name__ == '__main__':
    main()
