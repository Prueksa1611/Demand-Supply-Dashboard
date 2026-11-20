"""Usage: python export.py "Demand and Supply Planning.xlsm" data.json
Writes summarised data only (no VIN numbers). Needs: pip install pandas openpyxl"""
import sys, json, pandas as pd

src, out = sys.argv[1], sys.argv[2]
x = pd.ExcelFile(src)
def ser(v):  # Excel stores some dates as serial numbers (e.g. 46356 = 7 Dec 2026)
    return pd.Timestamp('1899-12-30') + pd.Timedelta(days=float(v)) if isinstance(v, (int, float)) and not pd.isna(v) else v
iso = lambda d: None if pd.isna(ser(d)) else pd.Timestamp(ser(d)).strftime('%Y-%m-%d')
key = lambda m, c: f"{str(m).strip()}|{'' if pd.isna(c) else str(c).strip()}".lower()
dt = lambda s: pd.to_datetime(s.map(ser), errors='coerce')

# Calculated: counts that do not depend on today (cols A,B,C,D,E,G,I,J); row 3 holds source dates
top = x.parse('Calculated', header=None, nrows=3)
cal = x.parse('Calculated', header=None, skiprows=3).iloc[:, [0, 1, 2, 3, 4, 6, 8, 9]]
cal = cal[cal[0].isin(['CBU', 'CKD']) & ~cal[2].isin(['SUB-TTL', 'TOTAL', 'Grand Total'])]
calc = [[r[0], r[1], r[2]] + [int(v) for v in r[3:]] for r in cal.itertuples(index=False)]

# Confirmed Shipment Raw -> daily series (actual received)
sh = x.parse('Confirmed Shipment Raw', usecols='A:L')
sh.columns = ['org', 'sc', 'crm', 'per', 'tq', 'inv', 'model', 'etd', 'eta', 'color', 'qty', 'year']
sh['eta'] = dt(sh['eta'])
g = sh.dropna(subset=['eta']).groupby(['model', 'color', 'per', 'year', 'eta'], dropna=False)['qty'].sum().reset_index()
ship = [[key(r.model, r.color), r.per, r.year, iso(r.eta), int(r.qty)] for r in g.itertuples()]

# VIN Lists -> counts by ETA / tax date, plus pending tax by model (Tax Payment Date = Pending)
v = x.parse('VIN Lists', usecols='A:L')
v.columns = ['org', 'sc', 'per', 'inv', 'vin', 'model', 'color', 'eta', 'tax', 'gr', 'crm', 'year']
pend = v[v['tax'].astype(str).eq('Pending') & v['per'].astype(str).str.match(r'^Y2026-P(0[1-9]|1[0-2])$')].copy()  # only ordering periods Y2026-P01..P12
pend['eta'] = dt(pend['eta'])
tax = [[m, int(len(p)), iso(p.eta.min()), iso(p.eta.max()), int(p.eta.isna().sum())] for m, p in pend.groupby('model')]
v['eta'], v['tax'] = dt(v['eta']), dt(v['tax'])
v['k'] = [key(m, c) for m, c in zip(v.model, v.color)]
g = v.groupby(['k', 'per', 'year', 'eta', 'tax'], dropna=False).size().reset_index(name='n')
vin = [[r.k, r.per, r.year, iso(r.eta), iso(r.tax), int(r.n)] for r in g.itertuples()]

# Summary shipment: CBU shipment plan (ETA), CKD production plan (production date), BYD stock
s = x.parse('Summary shipment', usecols='A:I')
s.columns = ['org', 'model', 'color', 'type', 'etd', 'prod', 'eta', 'qty', 'grid']
ss = [[r.org, key(r.model, r.color), r.type, iso(r.grid), int(r.qty), r.model, r.color] for r in s.itertuples()]

# Sales Contract -> Open LC rows (model, color, qty, CRM PO, period, status, Open LC70%/DLC100% date)
sc = x.parse('Sales Contract')
sc = sc[sc.iloc[:, 6].notna()]
lc = [[r[0], r[1], int(r[2]), r[3], r[4], r[5] if isinstance(r[5], str) else '', iso(r[6])]
      for r in sc.iloc[:, [6, 7, 8, 1, 5, 4, 3]].itertuples(index=False)]

# Source dates
dlr = x.parse('DLRs Stock'); bk = x.parse('RawData Booking')
meta = {'dealer': iso(dt(dlr['ALLOCATION DATE']).max()), 'booking': iso(dt(bk['Booking date']).max()), 'byd': iso(dt(s[s['type'].eq('BYD Stock')]['eta']).min())}  # date of the BYD stock rows
data = {'meta': meta, 'q3': iso(top.iloc[2, 16]), 'r3': iso(top.iloc[2, 17]), 'maxEta': iso(sh['eta'].max()),
        'calc': calc, 'ship': ship, 'vin': vin, 'ss': ss, 'lc': lc, 'tax': tax}
json.dump(data, open(out, 'w'), separators=(',', ':'), default=str)
print(meta, {k: len(v) for k, v in data.items() if isinstance(v, list)})
