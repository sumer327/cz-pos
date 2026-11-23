"""Rebuild retail.json (the POS's in-store prices) from CELLULAR ZONE MASTER BUYBACK SHEET 2.0.
Usage: export the sheet as .xlsx, then  python3 tools/build_retail.py master.xlsx  > retail.json
Reads the RETAIL columns: IPHONES (UNLK  PAID OFF) P, SAMSUNG O, PIXEL O, IPADS / APPLE WATCH / SURFACE PRO / GAME CONSOLE / MACBOOK N
(MACBOOK also its UPGRADE: xxGB RAM retail bonus rows). Row = [family, model key, variant key, storage key, retail, sheet model, sheet variant, sheet storage, Swappa list]."""
import json, re, sys, openpyxl
from datetime import date
wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
def nm(s):
    s = s.lower().replace('+', ' plus ')
    s = re.sub(r'\b(apple|google|samsung|motorola)\b', '', s)
    s = re.sub(r'(\d)\.(\d)', r'\1\2', s)
    s = re.sub(r'[^a-z0-9 ]', ' ', s); return re.sub(r'\s+', ' ', s).strip()
def sto(s):
    m = re.match(r'(\d+)(GB|TB|MM)', str(s or '').upper().replace(' ', '')); return (m.group(1) + m.group(2).lower()) if m else ''
# (tab, header row, RETAIL column, family, competitive Swappa list column: AVG SWAPPA / SWAPPA LIST, refreshed every morning)
TABS = [('IPHONES (UNLK  PAID OFF)', 4, 15, 'phone', 14), ('SAMSUNG', 5, 14, 'phone', 13), ('PIXEL', 5, 14, 'phone', 13), ('IPADS', 4, 13, 'ipad', 11),
        ('APPLE WATCH', 5, 13, 'watch', 11), ('SURFACE PRO', 5, 13, 'other', 11), ('GAME CONSOLE', 5, 13, 'other', 11)]
num = lambda v: float(v) if isinstance(v, (int, float)) and v > 0 else None
rows = []
for name, hdr, col, fam, sw in TABS:
    cur = None
    for r in wb[name].iter_rows(min_row=hdr + 1, values_only=True):
        if r[0]: cur = str(r[0]).strip()
        rt = r[col]
        if not cur or not isinstance(rt, (int, float)) or rt <= 0: continue
        c = str(r[1] or '').strip(); cl = c.lower()
        kv = ('U' if cl.startswith('unl') else 'A' if cl == 'att' else 'O') if fam == 'phone' else ('W' if cl == 'wifi' else 'C' if cl == 'cellular' else 'WC') if fam == 'ipad' else nm(c)
        rows.append([fam, nm(cur), kv, sto(r[2]), float(rt), cur, c, str(r[2] or '').strip(), num(r[sw])])
cur = None
for r in wb['MACBOOK'].iter_rows(min_row=6, values_only=True):
    m = r[0]
    if not m or m == 'MODEL': continue
    if str(m).startswith('UPGRADE'):
        b = re.match(r'\+\$(\d+)', str(r[13] or '')); ram = re.search(r'(\d+)GB', str(m))
        if cur and b and ram and int(b.group(1)) > 0: rows.append(['macup', cur, ram.group(1) + 'gb', '', int(b.group(1)), '', '', ''])
    elif isinstance(r[13], (int, float)):
        cur = nm(str(m)); rows.append(['mac', cur, nm(str(r[1] or '')), sto(r[2]), float(r[13]), str(m).strip(), str(r[1] or '').strip(), str(r[2] or '').strip(), num(r[11])])
print(json.dumps({'asof': 'MASTER BUYBACK SHEET 2.0 · ' + date.today().isoformat(), 'rows': rows}, separators=(',', ':')))
