#!/usr/bin/env python3
"""Generate a formatted Excel report from CSV input to XLSX output."""

import sys, os, re

missing = []
for pkg, imp in [("openpyxl","openpyxl"),("pyyaml","yaml")]:
    try: __import__(imp)
    except ImportError: missing.append(pkg)
if missing:
    print("Missing packages: " + " ".join(missing))
    print("Fix: pip install " + " ".join(missing))
    sys.exit(1)

import csv
import yaml
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import DataBarRule

# ── helpers ──────────────────────────────────────────────────────────────────
def bd():
    s = Side(style='thin', color='BDBDBD')
    return Border(top=s, bottom=s, left=s, right=s)

def fill(c): return PatternFill("solid", fgColor=c)
def hfont(color='FFFFFF', size=10): return Font(name='Arial', bold=True, color=color, size=size)
def cfont(bold=False, size=10): return Font(name='Arial', bold=bold, size=size)
def ctr(): return Alignment(horizontal='center', vertical='center', wrap_text=True)
def lft(): return Alignment(horizontal='left', vertical='center', wrap_text=True)
def cw(ws, col, w): ws.column_dimensions[get_column_letter(col)].width = w

# ── read inputs ───────────────────────────────────────────────────────────────
csv_path = sys.argv[1] if len(sys.argv) >= 2 else 'data/output/result.csv'
out = sys.argv[2] if len(sys.argv) >= 3 else 'data/output/result.xlsx'
people_path = sys.argv[3] if len(sys.argv) >= 4 else 'data/input/people.csv'
cfg_path = 'data/input/config.yaml'

if not os.path.exists(csv_path):
    print(f"Error: {csv_path} not found — run cargo first"); sys.exit(1)
if not os.path.exists(people_path):
    print(f"Error: people file not found: {people_path}"); sys.exit(1)

cfg = yaml.safe_load(open(cfg_path))
dessert_addr = f"{cfg['dessert_address']} {cfg['dessert_postal_code']} {cfg['dessert_city']}"
event_title = cfg.get('event_title', 'Happy agape 2')

rows = []
with open(csv_path) as f:
    for r in csv.DictReader(f):
        rows.append(r)

# ── load distance cache for walk times ───────────────────────────────────────
import json
cache_path = 'data/cache/distance_cache.json'
dist = {}
if os.path.exists(cache_path):
    dist = json.load(open(cache_path)).get('entries', {})

def normalize_address(address):
    s = (address or '').lower()
    s = s.replace('-', ' ').replace('_', ' ')
    s = re.sub(r'[^\w\s]', ' ', s, flags=re.UNICODE)
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def canonical_cache_key(addr_a, addr_b):
    a = normalize_address(addr_a)
    b = normalize_address(addr_b)
    return f"{a}|||{b}" if a <= b else f"{b}|||{a}"

# Build normalization-insensitive index from existing cache entries.
dist_normalized = {}
for k, v in dist.items():
    if '|||' not in k:
        continue
    left, right = k.split('|||', 1)
    ck = canonical_cache_key(left, right)
    if ck not in dist_normalized or v < dist_normalized[ck]:
        dist_normalized[ck] = v

def walk(addr_from, addr_to):
    """Return walk time in minutes from cache, or 0 if not found."""
    if normalize_address(addr_from) == normalize_address(addr_to):
        return 0.0

    key = canonical_cache_key(addr_from, addr_to)
    if key in dist_normalized:
        return round(dist_normalized[key] / 60.0, 1)

    # Walking time is assumed symmetric: A->B == B->A.
    key = f"{addr_from}|||{addr_to}" if addr_from <= addr_to else f"{addr_to}|||{addr_from}"
    if key in dist:
        return round(dist[key] / 60.0, 1)
    # Backward compatibility with older directional cache files.
    legacy = f"{addr_from}|||{addr_to}"
    reverse = f"{addr_to}|||{addr_from}"
    return round(dist.get(legacy, dist.get(reverse, 0)) / 60.0, 1)

# Build address map: name -> address
people_csv = []
with open(people_path) as f:
    for r in csv.DictReader(f):
        addr = f"{r['postal_address']} {r['postal_code']} {r['city']}"
        people_csv.append({'name': r['name'], 'addr': addr})

addr_map = {p['name']: p['addr'] for p in people_csv}

# Compute walk times for each person
for r in rows:
    person_addr      = addr_map.get(r['name'], '')
    drinks_host_addr = addr_map.get(r['drinks_host'], '')
    dinner_host_addr = addr_map.get(r['dinner_host'], '')
    r['w1'] = walk(person_addr, drinks_host_addr)
    r['w2'] = walk(drinks_host_addr, dinner_host_addr)
    r['w3'] = walk(dinner_host_addr, dessert_addr)
    r['total'] = round(r['w1'] + r['w2'] + r['w3'], 1)

# ── build workbook ────────────────────────────────────────────────────────────
wb = Workbook()

# ═══ SHEET 1 — Overview ══════════════════════════════════════════════════════
ws = wb.active
ws.title = "Vue d'ensemble"
ws.sheet_view.showGridLines = False
ws.freeze_panes = "A3"

ws.merge_cells("A1:G1")
ws["A1"] = f"🎉 {event_title.upper()} — SUGGESTION"
ws["A1"].font = Font(name='Arial', bold=True, size=17, color='FFFFFF')
ws["A1"].fill = fill("0B132B"); ws["A1"].alignment = ctr()
ws.row_dimensions[1].height = 40

hdrs   = ["👤 Nom","🍸 Aperitif chez","🍽️ Diner chez","🚶 Maison→Apero (min)","🚶 Apero→Diner (min)","🚶 Diner→Dessert (min)","📊 Total marche (min)"]
colors = ["1D4ED8","1D4ED8","C2410C","059669","059669","7C3AED","B91C1C"]
widths = [34,34,34,18,18,20,17]
ws.row_dimensions[2].height = 26
for ci,(h,c,w) in enumerate(zip(hdrs,colors,widths),1):
    cell = ws.cell(row=2,column=ci,value=h)
    cell.font=hfont(); cell.fill=fill(c); cell.alignment=ctr(); cell.border=bd()
    cw(ws,ci,w)

for ri,r in enumerate(rows,3):
    ws.row_dimensions[ri].height = 19
    same = r['drinks_host']==r['dinner_host']
    bg = fill("FFCCBC") if same else fill("F5F5F5" if ri%2==0 else "FFFFFF")
    vals = [r['name'],r['drinks_host'],r['dinner_host'],
            r['w1'],r['w2'],r['w3'],f"=D{ri}+E{ri}+F{ri}"]
    for ci,v in enumerate(vals,1):
        c = ws.cell(row=ri,column=ci,value=v)
        c.fill=bg; c.font=cfont(bold=(ci==1))
        c.alignment=lft() if ci==1 else ctr(); c.border=bd()
        if ci>=4: c.number_format='0.0'

tr = len(rows)+3
ws.row_dimensions[tr].height = 22
ws.merge_cells(f"A{tr}:C{tr}")
ws[f"A{tr}"]="MAXIMUM"; ws[f"A{tr}"].font=hfont(color='000000')
ws[f"A{tr}"].fill=fill("E3F2FD"); ws[f"A{tr}"].alignment=ctr(); ws[f"A{tr}"].border=bd()
for ci in range(4,8):
    col=get_column_letter(ci)
    c=ws.cell(row=tr,column=ci,value=f"=MAX({col}3:{col}{len(rows)+2})")
    c.font=hfont(color='000000',size=10); c.fill=fill("E3F2FD")
    c.number_format='0.0'; c.alignment=ctr(); c.border=bd()

ws.conditional_formatting.add(f"G3:G{len(rows)+2}",
    DataBarRule(start_type='min',end_type='max',color="0284C7",showValue=True))

# ═══ SHEET 2 — Aperitif ══════════════════════════════════════════════════════
from collections import defaultdict
ws2 = wb.create_sheet("Aperitif")
ws2.sheet_view.showGridLines = False
ws2.merge_cells("A1:C1")
ws2["A1"]="🍸 APERITIF — Repartition par hote"
ws2["A1"].font=Font(name='Arial',bold=True,size=14,color='FFFFFF')
ws2["A1"].fill=fill("1565C0"); ws2["A1"].alignment=ctr()
ws2.row_dimensions[1].height=30
for col,w in zip([1,2,3],[34,20,32]): cw(ws2,col,w)

drinks_groups = defaultdict(list)
for r in rows:
    host = r['drinks_host']
    drinks_groups.setdefault(host, [])
    drinks_groups[host].append(r)

row=2
for host,guests in drinks_groups.items():
    host_addr = addr_map.get(host,'')
    ws2.row_dimensions[row].height=26
    ws2.merge_cells(f"A{row}:C{row}")
    ws2[f"A{row}"]=f"Chez {host}  —  {host_addr}"
    ws2[f"A{row}"].font=Font(name='Arial',bold=True,size=11,color='FFFFFF')
    ws2[f"A{row}"].fill=fill("0D47A1"); ws2[f"A{row}"].alignment=lft(); row+=1
    ws2.row_dimensions[row].height=20
    for ci,h in enumerate(["👥 Participant","🚶 Maison → ici (min)","🍽️ Diner chez"],1):
        c=ws2.cell(row=row,column=ci,value=h)
        c.font=hfont(size=9); c.fill=fill("42A5F5"); c.alignment=ctr(); c.border=bd()
    row+=1
    if not guests:
        ws2.row_dimensions[row].height=18
        for ci,v in enumerate(["(aucun participant)","",""],1):
            c=ws2.cell(row=row,column=ci,value=v)
            c.fill=fill("E3F2FD"); c.font=cfont(bold=(ci==1))
            c.alignment=lft() if ci==1 else ctr(); c.border=bd()
        row+=1
    else:
        for g in guests:
            ws2.row_dimensions[row].height=18
            for ci,v in enumerate([g['name'],g['w1'],g['dinner_host']],1):
                c=ws2.cell(row=row,column=ci,value=v)
                c.fill=fill("E3F2FD"); c.font=cfont(bold=(ci==1))
                c.alignment=lft() if ci in (1,3) else ctr(); c.border=bd()
                if ci==2: c.number_format='0.0'
            row+=1
    row+=1

# ═══ SHEET 3 — Diner ═════════════════════════════════════════════════════════
ws3=wb.create_sheet("Diner")
ws3.sheet_view.showGridLines=False
ws3.merge_cells("A1:C1")
ws3["A1"]="🍽️ DINER — Repartition par hote"
ws3["A1"].font=Font(name='Arial',bold=True,size=14,color='FFFFFF')
ws3["A1"].fill=fill("E65100"); ws3["A1"].alignment=ctr()
ws3.row_dimensions[1].height=30
for col,w in zip([1,2,3],[32,22,30]): cw(ws3,col,w)

dinner_groups = defaultdict(list)
for r in rows:
    host = r['dinner_host']
    dinner_groups.setdefault(host, [])
    dinner_groups[host].append(r)

row=2
for host,guests in dinner_groups.items():
    host_addr=addr_map.get(host,'')
    ws3.row_dimensions[row].height=26
    ws3.merge_cells(f"A{row}:C{row}")
    ws3[f"A{row}"]=f"Chez {host}  —  {host_addr}"
    ws3[f"A{row}"].font=Font(name='Arial',bold=True,size=11,color='FFFFFF')
    ws3[f"A{row}"].fill=fill("BF360C"); ws3[f"A{row}"].alignment=lft(); row+=1
    ws3.row_dimensions[row].height=20
    for ci,h in enumerate(["👥 Participant","🚶 Apero → ici (min)","🍸 Apero chez"],1):
        c=ws3.cell(row=row,column=ci,value=h)
        c.font=hfont(size=9); c.fill=fill("FF7043"); c.alignment=ctr(); c.border=bd()
    row+=1
    if not guests:
        ws3.row_dimensions[row].height=18
        for ci,v in enumerate(["(aucun participant)","",""],1):
            c=ws3.cell(row=row,column=ci,value=v)
            c.fill=fill("FFF3E0"); c.font=cfont(bold=(ci==1))
            c.alignment=lft() if ci==1 else ctr(); c.border=bd()
        row+=1
    else:
        for g in guests:
            ws3.row_dimensions[row].height=18
            for ci,v in enumerate([g['name'],g['w2'],g['drinks_host']],1):
                c=ws3.cell(row=row,column=ci,value=v)
                c.fill=fill("FFF3E0")
                c.font=cfont(bold=(ci==1)); c.alignment=lft() if ci in (1,3) else ctr(); c.border=bd()
                if ci==2: c.number_format='0.0'
            row+=1
    row+=1

# ═══ SHEET 4 — Stats ═════════════════════════════════════════════════════════
ws5=wb.create_sheet("Statistiques")
ws5.sheet_view.showGridLines=False
for col,w in zip([1,2,3],[36,22,26]): cw(ws5,col,w)
ws5.merge_cells("A1:C1")
ws5["A1"]="📈 STATISTIQUES"
ws5["A1"].font=Font(name='Arial',bold=True,size=14,color='FFFFFF')
ws5["A1"].fill=fill("263238"); ws5["A1"].alignment=ctr(); ws5.row_dimensions[1].height=30

total_walk=sum(r['total'] for r in rows)
max_r=max(rows,key=lambda r:r['total']); min_r=min(rows,key=lambda r:r['total'])
row_by_name = {r['name']: r for r in rows}
dinner_host_walks = []
for host_name in sorted(dinner_groups.keys()):
    host_row = row_by_name.get(host_name)
    if host_row:
        dinner_host_walks.append({
            'name': host_name,
            'w2': host_row['w2'],
            'drinks_host': host_row['drinks_host'],
        })

stats=[
    ("","",""),("👥 PARTICIPANTS","",""),
    ("Nombre total",len(rows),"personnes"),
    ("","",""),("🍸 APERITIF","",""),("Hotes",len(drinks_groups),""),
]
for h,g in drinks_groups.items():
    stats.append((f"  Chez {h}",f"{len(g)} participants",""))
stats+=[("","",""),("🍽️ DINER","",""),("Hotes",len(dinner_groups),"")]
for h,g in dinner_groups.items():
    stats.append((f"  Chez {h}",f"{len(g)} participants",""))
stats += [("","",""),("🏠 HOTES DINER — Apero→Diner","","")]
if dinner_host_walks:
    avg_host_walk = sum(r['w2'] for r in dinner_host_walks) / len(dinner_host_walks)
    max_host_walk = max(dinner_host_walks, key=lambda r: r['w2'])
    for r in dinner_host_walks:
        stats.append((f"  {r['name']}",f"{r['w2']:.1f} min",f"depuis chez {r['drinks_host']}"))
    stats.append(("  Moyenne hotes",f"{avg_host_walk:.1f} min",""))
    stats.append(("  Maximum hote",f"{max_host_walk['w2']:.1f} min",f"({max_host_walk['name']})"))
else:
    stats.append(("  (aucun hote diner)","",""))
stats+=[
    ("","",""),("🚶 TEMPS DE MARCHE","",""),
    ("Moyenne / personne",f"{total_walk/len(rows):.1f} min",""),
    ("Maximum",f"{max_r['total']:.1f} min",f"({max_r['name']})"),
    ("Minimum",f"{min_r['total']:.1f} min",f"({min_r['name']})"),
]
sections={"👥 PARTICIPANTS","🍸 APERITIF","🍽️ DINER","🏠 HOTES DINER — Apero→Diner","🚶 TEMPS DE MARCHE"}
for ri,(a,b,c) in enumerate(stats,2):
    ws5.row_dimensions[ri].height=20
    for ci,v in enumerate([a,b,c],1):
        cell=ws5.cell(row=ri,column=ci,value=v); cell.border=bd()
        if str(a) in sections:
            cell.font=Font(name='Arial',bold=True,size=11,color='FFFFFF')
            cell.fill=fill("455A64"); cell.alignment=lft()
        elif str(a).startswith("  "):
            cell.font=cfont(size=10); cell.fill=fill("ECEFF1"); cell.alignment=lft()
        else:
            cell.font=cfont(bold=(ci==1),size=10)
            cell.fill=fill("F5F5F5" if ri%2==0 else "FFFFFF"); cell.alignment=lft() if ci==1 else ctr()

out_dir = os.path.dirname(out)
if out_dir:
    os.makedirs(out_dir, exist_ok=True)
wb.save(out)
print(f"Excel saved: {out}")
