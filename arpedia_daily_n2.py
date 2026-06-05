#!/usr/bin/env python3
# TOOL: ARpedia 日報（北二區）
"""
ARpedia 日報 — 北二區（永和、板橋誠品、西門、花蓮、板橋遠百、新莊宏匯、新店裕隆城）
輸出：昨日 / 本週 / 月累積 / 年累積
"""
import subprocess
from datetime import date, timedelta
from config import JAVA, EPB_CP, EPB_CWD

STORES     = ['永和', '板橋誠品', '西門', '花蓮', '板橋遠百', '新莊宏匯', '新店裕隆城']
SHOP_CODES = {'永和':'9', '板橋誠品':'25', '西門':'50', '花蓮':'55',
              '板橋遠百':'63', '新莊宏匯':'64', '新店裕隆城':'68'}
SHOP_IN    = "'009','025','050','055','063','064','068'"

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

def main():
    today     = date.today()
    yesterday = today - timedelta(days=1)

    days_since_sunday = (yesterday.weekday() + 1) % 7
    week_start  = yesterday - timedelta(days=days_since_sunday)
    month_start = yesterday.replace(day=1)
    year_start  = today.replace(month=1, day=1)

    periods = [
        (f"昨日 {yesterday.strftime('%-m/%-d')}",                                    yesterday,   yesterday),
        (f"本週 {week_start.strftime('%-m/%-d')}~{yesterday.strftime('%-m/%-d')}",   week_start,  yesterday),
        (f"月累 {month_start.strftime('%-m/%-d')}~{yesterday.strftime('%-m/%-d')}", month_start, yesterday),
        (f"年累 {year_start.strftime('%-m/%-d')}~{yesterday.strftime('%-m/%-d')}",  year_start,  yesterday),
    ]

    print(f"\nARpedia 銷售日報 北二區 ({today.strftime('%Y-%m-%d')})")
    print("=" * 76)

    results = []
    for label, ds, de in periods:
        print(f"查詢 {label}...", flush=True)
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
    print()

if __name__ == '__main__':
    main()
