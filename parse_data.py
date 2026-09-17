#!/usr/bin/env python3
"""
parse_data.py  –  Read the Kuta Selatan sales xlsx and write data.js
                  containing `const D = { … }` for the dashboard.

Usage:  python3 parse_data.py [path/to/file.xlsx]
        (defaults to the first *.xlsx in the current directory)
"""

import glob, json, os, sys
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit("openpyxl not installed – run: pip install openpyxl")

# ── Find xlsx ─────────────────────────────────────────────────────────────────
if len(sys.argv) > 1:
    XLSX = Path(sys.argv[1])
else:
    candidates = glob.glob("*.xlsx")
    if not candidates:
        sys.exit("No .xlsx file found in current directory.")
    import re as _re
    _MON = {'jan':1,'feb':2,'mar':3,'apr':4,'may':5,'jun':6,
            'jul':7,'aug':8,'sep':9,'oct':10,'nov':11,'dec':12}
    def _xlsx_key(p):
        m = _re.search(r'jan-(\w{3})', Path(p).name.lower())
        return _MON.get(m.group(1), 0) if m else 0
    XLSX = Path(max(candidates, key=_xlsx_key))

OUT = Path("data.js")
print(f"Reading: {XLSX}")

MONTH_NAMES = {1:'Jan',2:'Feb',3:'Mar',4:'Apr',5:'May',6:'Jun',
               7:'Jul',8:'Aug',9:'Sep',10:'Oct',11:'Nov',12:'Dec'}

KS_SALES = {'Monica', 'Juni'}

FB_CATEGORIES = {
    'BEVERAGE', 'DAIRY', 'FLOUR MIP', 'FLOUR SRIBOGA',
    'BAKERY', 'FROZEN BAKERY', 'FROZEN FRIES', 'FROZEN MEAT', 'FROZEN SEA FOOD',
    'PIZZA',
}

def is_fb(r):
    cat = (r[7] or '').strip().upper()
    return cat in FB_CATEGORIES or cat == ''

# ── Load ──────────────────────────────────────────────────────────────────────
wb = openpyxl.load_workbook(XLSX, read_only=True, data_only=True)
ws = wb.active
rows = list(ws.iter_rows(values_only=True))
# Columns: Kode_Wil Nama_Wil Kode_Cust Nama_Cust Grup_Cust Kode_Brg Nama_Brg
#          Kat_Brg  Qty DPP PPN Grand_Total Sales Tipe_Order bulan nama_bulan
#          Nama_Div tahun
header = rows[0]
ks = [r for r in rows[1:] if r[12] in KS_SALES]
print(f"  {len(ks):,} rows for Monica + Juni")

# ── Month analysis ──────────────────────────────────────────────────────────────
all_rows     = rows[1:]
month_counts = defaultdict(int)
month_years  = {}
for r in ks:
    if r[14]:
        m = int(r[14])
        month_counts[m] += 1
        if r[17]:
            month_years[m] = int(r[17])

sorted_months = sorted(month_counts)
if not sorted_months:
    sys.exit("No month data found.")

# Detect partial (in-progress) month: the latest month is partial only if it's
# still the current real-world month/year, i.e. it hasn't finished yet.
# (A row-count heuristic was tried before but breaks once the month is more
# than ~1/3 over, since daily transaction volume isn't constant.)
try:
    from zoneinfo import ZoneInfo
    today = datetime.now(ZoneInfo("Asia/Makassar")).date()
except Exception:
    today = date.today()
latest_m   = sorted_months[-1]
latest_y   = month_years.get(latest_m)
partial_m  = latest_m if (latest_y == today.year and latest_m == today.month) else None
ref_m      = sorted_months[-2] if partial_m and len(sorted_months) >= 2 else latest_m
prev_m     = (sorted_months[-3] if partial_m and len(sorted_months) >= 3
              else sorted_months[-2] if len(sorted_months) >= 2 else ref_m)

print(f"  Months in data: {sorted_months}")
print(f"  Partial month:  {partial_m} ({MONTH_NAMES.get(partial_m,'?') if partial_m else 'none'})")
print(f"  Reference month (growth compare): {MONTH_NAMES.get(ref_m,'?')} vs {MONTH_NAMES.get(prev_m,'?')}")

def is_balian(r):
    return 'BALIAN' in (r[7] or '').upper()

def grand(r):
    return float(r[11] or 0)

# ── Monthly revenue ────────────────────────────────────────────────────────────
monica_fb  = defaultdict(float)
monica_bal = defaultdict(float)
juni_fb    = defaultdict(float)
juni_bal   = defaultdict(float)

for r in ks:
    m = int(r[14] or 0)
    if not m:
        continue
    g = grand(r)
    if is_balian(r):
        (monica_bal if r[12] == 'Monica' else juni_bal)[m] += g
    elif is_fb(r):
        (monica_fb  if r[12] == 'Monica' else juni_fb )[m] += g

def month_series(d):
    return [round(d.get(m, 0)) for m in sorted_months]

months_labels = [
    MONTH_NAMES.get(m, str(m)) + ('*' if m == partial_m else '')
    for m in sorted_months
]

# ── Per-customer aggregates ────────────────────────────────────────────────────
cust_total       = defaultdict(float)
cust_fb          = defaultdict(float)
cust_bal         = defaultdict(float)
cust_sp          = {}
cust_by_m        = defaultdict(lambda: defaultdict(float))  # fb
cust_by_m_bal    = defaultdict(lambda: defaultdict(float))  # balian
cust_top_prod    = defaultdict(lambda: defaultdict(float))
cust_skus        = defaultdict(set)

for r in ks:
    m  = int(r[14] or 0)
    cn = (r[3] or '').strip()
    pn = (r[6] or '').strip()
    kb = (r[5] or '').strip()
    g  = grand(r)
    sp = r[12]
    if not cn:
        continue
    cust_total[cn] += g
    cust_sp[cn] = sp
    if is_balian(r):
        cust_bal[cn]            += g
        cust_by_m_bal[cn][m]   += g
    elif is_fb(r):
        cust_fb[cn]             += g
        cust_by_m[cn][m]       += g
    if pn:
        cust_top_prod[cn][pn]  += g
    if kb:
        cust_skus[cn].add(kb)

# ── ABC classification ─────────────────────────────────────────────────────────
sorted_custs = sorted(cust_total.items(), key=lambda x: -x[1])
total_rev    = sum(v for _, v in sorted_custs)
cust_cls     = {}
running      = 0
for cn, rev in sorted_custs:
    running += rev
    frac = running / total_rev if total_rev else 1
    cust_cls[cn] = 'A' if frac <= 0.80 else ('B' if frac <= 0.95 else 'C')

abc_buckets = {c: {'accs': 0, 'rev': 0.0} for c in 'ABC'}
abc_custs = {'A': [], 'B': [], 'C': []}
for cn, rev in sorted_custs:
    c = cust_cls.get(cn, 'C')
    abc_buckets[c]['accs'] += 1
    abc_buckets[c]['rev']  += rev
    abc_custs[c].append({'n': cn, 'rev': round(rev), 'sp': cust_sp.get(cn, '')})

abc_strat = {'A': 'Top 20% — Protect & Grow',
             'B': 'Middle — Upsell',
             'C': 'Long tail — Review'}
abc_list = [
    {'cls': c,
     'accs': abc_buckets[c]['accs'],
     'rev':  round(abc_buckets[c]['rev']),
     'pct':  round(abc_buckets[c]['rev'] / total_rev, 3) if total_rev else 0,
     'strat': abc_strat[c]}
    for c in 'ABC'
]

# ── Top 10 customers (total revenue) ──────────────────────────────────────────
top10 = [
    {'n': cn, 'cls': cust_cls.get(cn, 'C'), 'sp': cust_sp.get(cn, ''), 'rev': round(rev)}
    for cn, rev in sorted_custs[:10]
]

# ── Top 10 by salesperson (FB only) ───────────────────────────────────────────
def sp_fb_sorted(sp):
    return sorted(
        ((cn, v) for cn, v in cust_fb.items() if cust_sp.get(cn) == sp),
        key=lambda x: -x[1]
    )

monica_custs = sp_fb_sorted('Monica')
juni_custs   = sp_fb_sorted('Juni')

def sp_top10_list(custs):
    total = sum(v for _, v in custs) if custs else 1
    return [
        {'n': cn, 'rev': round(v), 'pct': round(v / total * 100, 2)}
        for cn, v in custs[:10]
    ]

# ── Top Products ───────────────────────────────────────────────────────────────
prod_rev   = defaultdict(float)
prod_custs = defaultdict(set)
for r in ks:
    if not is_fb(r):
        continue
    pn = (r[6] or '').strip()
    cn = (r[3] or '').strip()
    if pn:
        prod_rev[pn]   += grand(r)
        if cn:
            prod_custs[pn].add(cn)

products_list = [
    {'n': pn, 'rev': round(rev), 'cust': len(prod_custs[pn])}
    for pn, rev in sorted(prod_rev.items(), key=lambda x: -x[1])[:10]
]

# ── Top 10 Products per Customer ───────────────────────────────────────────────
cust_top10_items = {}
for cn, prods in cust_top_prod.items():
    sorted_prods = sorted(prods.items(), key=lambda x: -x[1])[:10]
    cust_top10_items[cn] = [{'n': p, 'rev': round(r)} for p, r in sorted_prods]

# ── Growth / Decline ───────────────────────────────────────────────────────────
def growth_decline(cust_months_dict, sp):
    grow, decl = [], []
    for cn, md in cust_months_dict.items():
        if cust_sp.get(cn) != sp:
            continue
        lat = md.get(ref_m, 0)
        prv = md.get(prev_m, 0)
        if prv == 0 and lat == 0:
            continue
        if prv == 0:
            g_pct = 999.0
            is_new = True
        else:
            g_pct = round((lat - prv) / prv * 100, 1)
            is_new = False
        cls = cust_cls.get(cn, 'C')
        if g_pct > 5:
            grow.append({'n': cn, 'cls': cls, 'g': g_pct, 'lat': round(lat), 'new': is_new})
        elif g_pct < -5:
            prods = cust_top_prod.get(cn, {})
            top_p = max(prods, key=prods.get) if prods else ''
            decl.append({'n': cn, 'cls': cls, 'd': g_pct, 'act': 'URGENT', 'top_p': top_p})
    grow.sort(key=lambda x: -x['g'])
    decl.sort(key=lambda x:  x['d'])
    return grow[:8], decl[:8]

grow_m,  dec_m  = growth_decline(cust_by_m,     'Monica')
grow_j,  dec_j  = growth_decline(cust_by_m,     'Juni')
grow_bm, dec_bm = growth_decline(cust_by_m_bal, 'Monica')
grow_bj, dec_bj = growth_decline(cust_by_m_bal, 'Juni')

# ── Dormant ────────────────────────────────────────────────────────────────────
def dormant_list(cust_months_dict, sp):
    out = []
    for cn, md in cust_months_dict.items():
        if cust_sp.get(cn) != sp:
            continue
        # Still active in ref month or partial month = not dormant
        if md.get(ref_m, 0) > 0:
            continue
        if partial_m and md.get(partial_m, 0) > 0:
            continue
        # Had orders in prev_m = already captured in decline list, skip here
        if md.get(prev_m, 0) > 0:
            continue
        active_months = [m for m, v in md.items() if v > 0]
        if not active_months:
            continue
        last_m = max(active_months)
        ytd    = sum(md.values())
        prods = cust_top_prod.get(cn, {})
        top_p = max(prods, key=prods.get) if prods else ''
        out.append({
            'n':    cn,
            'cls':  cust_cls.get(cn, 'C'),
            'last': MONTH_NAMES.get(last_m, str(last_m)),
            'ytd':  round(ytd),
            'top_p': top_p
        })
    out.sort(key=lambda x: -x['ytd'])
    return out[:15]

dorm_m  = dormant_list(cust_by_m,     'Monica')
dorm_j  = dormant_list(cust_by_m,     'Juni')
dorm_bm = dormant_list(cust_by_m_bal, 'Monica')
dorm_bj = dormant_list(cust_by_m_bal, 'Juni')

active_m    = sum(1 for cn in cust_total if cust_sp.get(cn) == 'Monica')
active_j    = sum(1 for cn in cust_total if cust_sp.get(cn) == 'Juni')
total_skus  = len(prod_rev)
dormant_cnt = len(dorm_m) + len(dorm_j)

# ── Upsell (top account + their top product) ──────────────────────────────────
def upsell_list(custs):
    out = []
    for cn, rev in custs[:7]:
        prods = cust_top_prod.get(cn, {})
        top_p = max(prods, key=prods.get) if prods else ''
        out.append({'n': cn, 'prod': top_p, 'cls': cust_cls.get(cn, 'C'), 'rev': round(rev)})
    return out

ups_m = upsell_list(monica_custs)
ups_j = upsell_list(juni_custs)

# ── Opportunities (Dynamic Insights) ──────────────────────────────────────────
opp_candidates = []
# Find top accounts with low SKUs or missing cross-sell (FB vs Balian)
for cn, rev in sorted_custs:
    cls  = cust_cls.get(cn, 'C')
    skus = len(cust_skus.get(cn, set()))
    sp   = cust_sp.get(cn, '')
    fb_rev = cust_fb.get(cn, 0)
    bal_rev = cust_bal.get(cn, 0)
    
    if cls in ('A', 'B'):
        # Check cross sell
        if fb_rev > 10_000_000 and bal_rev == 0:
            opp_candidates.append((cn, sp, cls, skus, round(rev), 'Buys F&B but no Balian - cross-sell opportunity!'))
        elif bal_rev > 10_000_000 and fb_rev == 0:
            opp_candidates.append((cn, sp, cls, skus, round(rev), 'Buys Balian but no F&B - pitch food catalog!'))
        # Check low SKUs for A/B class
        elif skus <= 5:
            opp_candidates.append((cn, sp, cls, skus, round(rev), f'Only {skus} SKUs for {cls}-class account. Upsell new categories.'))

# Sort by revenue descending
opp_candidates.sort(key=lambda x: -x[4])

opps = []
seen_opps = set()
for i, (cn, sp, cls, skus, ytd, act) in enumerate(opp_candidates, 1):
    if cn in seen_opps:
        continue
    seen_opps.add(cn)
    opps.append({'rank': len(opps) + 1, 'n': cn, 'sp': sp, 'cls': cls,
                 'skus': skus, 'ytd': ytd, 'act': act})
    if len(opps) >= 5:
        break

# ── Advanced Insights ──────────────────────────────────────────────────────────
combined_md = defaultdict(lambda: defaultdict(float))
for cn, md in cust_by_m.items():
    for m, v in md.items():
        combined_md[cn][m] += v
for cn, md in cust_by_m_bal.items():
    for m, v in md.items():
        combined_md[cn][m] += v

churn_risk = []
recovery = []
consistent = []
aov_drops = []

for cn, md in combined_md.items():
    cls = cust_cls.get(cn, 'C')
    sp = cust_sp.get(cn, '')
    
    # 1. Churn Risk (A/B class, drop >20% in last 2 months vs history)
    if cls in ('A', 'B'):
        active_mons = [m for m, v in md.items() if v > 0]
        recent_2 = sum(md.get(m, 0) for m in (ref_m, prev_m))
        hist_months = [m for m in active_mons if m not in (ref_m, prev_m)]
        if hist_months:
            hist_avg = sum(md[m] for m in hist_months) / len(hist_months)
            recent_avg = recent_2 / 2
            if hist_avg > 0 and (hist_avg - recent_avg) / hist_avg > 0.2:
                churn_risk.append({'n': cn, 'cls': cls, 'sp': sp, 'drop': round((hist_avg - recent_avg) / hist_avg * 100)})

    # 2. Recovery / Win-Backs (gap >= 3 months, but active now)
    if md.get(ref_m, 0) > 0:
        active_before = [m for m, v in md.items() if v > 0 and m < ref_m]
        if active_before:
            gap = ref_m - max(active_before)
            if gap >= 3:
                recovery.append({'n': cn, 'cls': cls, 'sp': sp, 'gap': gap, 'rev': md.get(ref_m, 0)})
    
    # 4. Consistent Buyers (active in all available months)
    active_count = len([m for m, v in md.items() if v > 0])
    if active_count == len(months_labels) or (active_count == len(months_labels)-1 and partial_m):
        consistent.append({'n': cn, 'cls': cls, 'sp': sp, 'rev': sum(md.values())})

    # 5. AOV Drops (>40% drop in recent month vs historical avg)
    if md.get(ref_m, 0) > 0:
        hist_months = [m for m, v in md.items() if v > 0 and m < ref_m]
        if hist_months:
            hist_avg = sum(md[m] for m in hist_months) / len(hist_months)
            cur = md[ref_m]
            if hist_avg > 0 and (hist_avg - cur) / hist_avg > 0.4:
                aov_drops.append({'n': cn, 'cls': cls, 'sp': sp, 'drop': round((hist_avg - cur) / hist_avg * 100)})

# 3. High-Dependence Accounts (>80% rev from single product)
single_prod = []
for cn, prods in cust_top_prod.items():
    tot = sum(prods.values())
    if tot < 5_000_000: continue
    top_p, top_v = max(prods.items(), key=lambda x: x[1])
    if top_v / tot > 0.8:
        single_prod.append({'n': cn, 'cls': cust_cls.get(cn, 'C'), 'sp': cust_sp.get(cn, ''), 'prod': top_p, 'pct': round(top_v/tot*100)})

churn_risk.sort(key=lambda x: -x['drop'])
recovery.sort(key=lambda x: -x['rev'])
single_prod.sort(key=lambda x: -x['pct'])
consistent.sort(key=lambda x: -x['rev'])
aov_drops.sort(key=lambda x: -x['drop'])

# ── Target ─────────────────────────────────────────────────────────────────────
TARGET_H1 = 1_800_000_000
TARGET_H2 = 2_200_000_000

# ── Assemble D ─────────────────────────────────────────────────────────────────
D = {
    'months':      months_labels,
    'monica_fb':   month_series(monica_fb),
    'juni_fb':     month_series(juni_fb),
    'monica_bal':  month_series(monica_bal),
    'juni_bal':    month_series(juni_bal),
    'target_h1':   TARGET_H1,
    'target_h2':   TARGET_H2,
    'top10':       top10,
    'monica_top10': sp_top10_list(monica_custs),
    'juni_top10':   sp_top10_list(juni_custs),
    'abc':         abc_list,
    'abc_custs':   abc_custs,
    'cust_top10_items': cust_top10_items,
    'products':    products_list,
    'grow_m':      grow_m,
    'grow_j':      grow_j,
    'bal_grow_m':  grow_bm,
    'bal_grow_j':  grow_bj,
    'dec_m':       dec_m,
    'dec_j':       dec_j,
    'bal_dec_m':   dec_bm,
    'bal_dec_j':   dec_bj,
    'dorm_m':      dorm_m,
    'dorm_j':      dorm_j,
    'bal_dorm_m':  dorm_bm,
    'bal_dorm_j':  dorm_bj,
    'ups_m':       ups_m,
    'ups_j':       ups_j,
    'opps':        opps,
    'active_m':    active_m,
    'active_j':    active_j,
    'skus':        total_skus,
    'dormant_cnt': dormant_cnt,
    'churn_risk':  churn_risk[:10],
    'recovery':    recovery[:10],
    'single_prod': single_prod[:10],
    'consistent':  consistent[:10],
    'aov_drops':   aov_drops[:10],
}

# ── Write data.js ──────────────────────────────────────────────────────────────
js = 'const D = ' + json.dumps(D, ensure_ascii=False, indent=2) + ';\n'
OUT.write_text(js, encoding='utf-8')
print(f"✓ Written {OUT}  ({len(js):,} bytes)")
print(f"  months:     {months_labels}")
print(f"  monica_fb:  {month_series(monica_fb)}")
print(f"  juni_fb:    {month_series(juni_fb)}")
print(f"  monica_bal: {month_series(monica_bal)}")
print(f"  juni_bal:   {month_series(juni_bal)}")
print(f"  top10:      {[x['n'] for x in top10]}")
print(f"  abc:        A={abc_list[0]['accs']} accs, B={abc_list[1]['accs']} accs, C={abc_list[2]['accs']} accs")
print(f"  grow_m:     {[x['n'] for x in grow_m]}")
print(f"  dec_m:      {[x['n'] for x in dec_m]}")
print(f"  dorm_m:     {[x['n'] for x in dorm_m[:5]]}")
