#!/usr/bin/env python3
"""
parse_data.py  –  Read the Kuta Selatan sales xlsx and write data.js
                  containing `const D = { … }` for the dashboard.

Usage:  python3 parse_data.py [path/to/file.xlsx]
        (defaults to the first *.xlsx in the current directory)
"""

import glob, json, sys
from collections import defaultdict
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
    XLSX = Path(candidates[0])

OUT = Path("data.js")
print(f"Reading: {XLSX}")

MONTH_NAMES = {1:'Jan',2:'Feb',3:'Mar',4:'Apr',5:'May',6:'Jun',
               7:'Jul',8:'Aug',9:'Sep',10:'Oct',11:'Nov',12:'Dec'}

KS_SALES = {'Monica', 'Juni'}

FB_CATEGORIES = {
    'BEVERAGE', 'DAIRY', 'FLOUR MIP', 'FLOUR SRIBOGA',
    'FROZEN BAKERY', 'FROZEN FRIES', 'FROZEN MEAT', 'FROZEN SEA FOOD',
    'PIZZA',
}

def is_fb(r):
    return (r[7] or '').strip().upper() in FB_CATEGORIES

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

# ── Month analysis (use full dataset for reliable partial detection) ───────────
all_rows     = rows[1:]
month_counts_all = defaultdict(int)
month_counts     = defaultdict(int)
for r in all_rows:
    if r[14]:
        month_counts_all[int(r[14])] += 1
for r in ks:
    if r[14]:
        month_counts[int(r[14])] += 1

sorted_months = sorted(month_counts)
if not sorted_months:
    sys.exit("No month data found.")

# Detect partial (in-progress) month using full-dataset counts (< 75 % of mean)
all_months_all = sorted(month_counts_all)
other_all  = [month_counts_all[m] for m in all_months_all[:-1]]
avg_all    = sum(other_all) / len(other_all) if other_all else month_counts_all[all_months_all[-1]]
latest_m   = sorted_months[-1]
partial_m  = latest_m if month_counts_all.get(latest_m, 0) < avg_all * 0.75 else None
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
for cn, rev in cust_total.items():
    b = abc_buckets[cust_cls.get(cn, 'C')]
    b['accs'] += 1
    b['rev']  += rev

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

# ── Top 5 by salesperson (FB only) ────────────────────────────────────────────
def sp_fb_sorted(sp):
    return sorted(
        ((cn, v) for cn, v in cust_fb.items() if cust_sp.get(cn) == sp),
        key=lambda x: -x[1]
    )

monica_custs = sp_fb_sorted('Monica')
juni_custs   = sp_fb_sorted('Juni')

def top5_list(custs):
    total = sum(v for _, v in custs) if custs else 1
    return [
        {'n': cn, 'rev': round(v), 'pct': round(v / total * 100, 2)}
        for cn, v in custs[:5]
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
        else:
            g_pct = round((lat - prv) / prv * 100, 1)
        cls = cust_cls.get(cn, 'C')
        if g_pct > 5:
            grow.append({'n': cn, 'cls': cls, 'g': g_pct, 'lat': round(lat)})
        elif g_pct < -5:
            decl.append({'n': cn, 'cls': cls, 'd': g_pct, 'act': 'URGENT'})
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
        out.append({
            'n':    cn,
            'cls':  cust_cls.get(cn, 'C'),
            'last': MONTH_NAMES.get(last_m, str(last_m)),
            'ytd':  round(ytd)
        })
    out.sort(key=lambda x: -x['ytd'])
    return out[:15]

dorm_m  = dormant_list(cust_by_m,     'Monica')
dorm_j  = dormant_list(cust_by_m,     'Juni')
dorm_bm = dormant_list(cust_by_m_bal, 'Monica')
dorm_bj = dormant_list(cust_by_m_bal, 'Juni')

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

# ── Opportunities (high-rev, low-SKU accounts) ────────────────────────────────
opp_candidates = []
for cn, rev in sorted_custs:
    cls  = cust_cls.get(cn, 'C')
    skus = len(cust_skus.get(cn, set()))
    sp   = cust_sp.get(cn, '')
    if cls in ('A', 'B') and skus <= 8 and rev > 50_000_000:
        opp_candidates.append((cn, sp, cls, skus, round(rev)))

# Sort: A-class first, then by revenue descending
opp_candidates.sort(key=lambda x: (x[2] == 'B', -x[4]))

opps = []
for i, (cn, sp, cls, skus, ytd) in enumerate(opp_candidates[:5], 1):
    act = (f'Expand product range — currently {skus} SKU{"s" if skus != 1 else ""}, '
           'pitch new categories')
    opps.append({'rank': i, 'n': cn, 'sp': sp, 'cls': cls,
                 'skus': skus, 'ytd': ytd, 'act': act})

# ── Target ─────────────────────────────────────────────────────────────────────
TARGET = 2_200_000_000 if min(sorted_months) >= 7 else 1_800_000_000

# ── Assemble D ─────────────────────────────────────────────────────────────────
D = {
    'months':      months_labels,
    'monica_fb':   month_series(monica_fb),
    'juni_fb':     month_series(juni_fb),
    'monica_bal':  month_series(monica_bal),
    'juni_bal':    month_series(juni_bal),
    'target':      TARGET,
    'top10':       top10,
    'monica_top5': top5_list(monica_custs),
    'juni_top5':   top5_list(juni_custs),
    'abc':         abc_list,
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
