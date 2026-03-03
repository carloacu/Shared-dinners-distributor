#!/usr/bin/env python3
"""Generate a formatted Excel report from CSV input to XLSX output."""

import sys, os

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
cfg_path = 'data/input/config.yaml'

if not os.path.exists(csv_path):
    print(f"Error: {csv_path} not found — run cargo first"); sys.exit(1)

cfg = yaml.safe_load(open(cfg_path))
dessert_addr = f"{cfg['dessert_address']} {cfg['dessert_postal_code']} {cfg['dessert_city']}"

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

def walk(addr_from, addr_to):
    """Return walk time in minutes from cache, or 0 if not found."""
    if addr_from == addr_to: return 0.0
    key = f"{addr_from}|||{addr_to}"
    return round(dist.get(key, 0) / 60.0, 1)

# Build address map: name -> address
people_csv = []
with open('data/input/people.csv') as f:
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
ws.freeze_panes = "A4"

ws.merge_cells("A1:G1")
ws["A1"] = "PROGRESSIVE DINNER — REPARTITION FINALE"
ws["A1"].font = Font(name='Arial', bold=True, size=16, color='FFFFFF')
ws["A1"].fill = fill("1A1A2E"); ws["A1"].alignment = ctr()
ws.row_dimensions[1].height = 36

ws.merge_cells("A2:G2")
ws["A2"] = f"Dessert : {dessert_addr}"
ws["A2"].font = Font(name='Arial', italic=True, size=11, color='546E7A')
ws["A2"].fill = fill("ECEFF1"); ws["A2"].alignment = ctr()
ws.row_dimensions[2].height = 20

hdrs   = ["Nom","Aperitif chez","Diner chez","Maison→Apero","Apero→Diner","Diner→Dessert","Total marche"]
colors = ["1565C0","1565C0","E65100","388E3C","388E3C","7B1FA2","C62828"]
widths = [28,28,28,17,17,19,16]
ws.row_dimensions[3].height = 26
for ci,(h,c,w) in enumerate(zip(hdrs,colors,widths),1):
    cell = ws.cell(row=3,column=ci,value=h)
    cell.font=hfont(); cell.fill=fill(c); cell.alignment=ctr(); cell.border=bd()
    cw(ws,ci,w)

for ri,r in enumerate(rows,4):
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

tr = len(rows)+4
ws.row_dimensions[tr].height = 22
ws.merge_cells(f"A{tr}:C{tr}")
ws[f"A{tr}"]="TOTAL"; ws[f"A{tr}"].font=hfont(color='000000')
ws[f"A{tr}"].fill=fill("E3F2FD"); ws[f"A{tr}"].alignment=ctr(); ws[f"A{tr}"].border=bd()
for ci in range(4,8):
    col=get_column_letter(ci)
    c=ws.cell(row=tr,column=ci,value=f"=SUM({col}4:{col}{len(rows)+3})")
    c.font=hfont(color='000000',size=10); c.fill=fill("E3F2FD")
    c.number_format='0.0'; c.alignment=ctr(); c.border=bd()

ws.conditional_formatting.add(f"G4:G{len(rows)+3}",
    DataBarRule(start_type='min',end_type='max',color="1565C0",showValue=True))

# ═══ SHEET 2 — Aperitif ══════════════════════════════════════════════════════
from collections import defaultdict
ws2 = wb.create_sheet("Aperitif")
ws2.sheet_view.showGridLines = False
ws2.merge_cells("A1:C1")
ws2["A1"]="APERITIF — Repartition par hote"
ws2["A1"].font=Font(name='Arial',bold=True,size=14,color='FFFFFF')
ws2["A1"].fill=fill("1565C0"); ws2["A1"].alignment=ctr()
ws2.row_dimensions[1].height=30
for col,w in zip([1,2,3],[28,18,26]): cw(ws2,col,w)

drinks_groups = defaultdict(list)
for r in rows:
    host = r['drinks_host']
    drinks_groups.setdefault(host, [])
    if r['name'] != host:
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
    for ci,h in enumerate(["Invite","Maison → ici (min)","Diner chez"],1):
        c=ws2.cell(row=row,column=ci,value=h)
        c.font=hfont(size=9); c.fill=fill("42A5F5"); c.alignment=ctr(); c.border=bd()
    row+=1
    if not guests:
        ws2.row_dimensions[row].height=18
        for ci,v in enumerate(["(aucun invite)","",""],1):
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
ws3["A1"]="DINER — Repartition par hote"
ws3["A1"].font=Font(name='Arial',bold=True,size=14,color='FFFFFF')
ws3["A1"].fill=fill("E65100"); ws3["A1"].alignment=ctr()
ws3.row_dimensions[1].height=30
for col,w in zip([1,2,3],[28,18,22]): cw(ws3,col,w)

dinner_groups = defaultdict(list)
for r in rows:
    host = r['dinner_host']
    dinner_groups.setdefault(host, [])
    if r['name'] != host:
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
    for ci,h in enumerate(["Invite","Apero → ici (min)","Meme hote apero ?"],1):
        c=ws3.cell(row=row,column=ci,value=h)
        c.font=hfont(size=9); c.fill=fill("FF7043"); c.alignment=ctr(); c.border=bd()
    row+=1
    if not guests:
        ws3.row_dimensions[row].height=18
        for ci,v in enumerate(["(aucun invite)","",""],1):
            c=ws3.cell(row=row,column=ci,value=v)
            c.fill=fill("FFF3E0"); c.font=cfont(bold=(ci==1))
            c.alignment=lft() if ci==1 else ctr(); c.border=bd()
        row+=1
    else:
        for g in guests:
            same=g['drinks_host']==g['dinner_host']
            ws3.row_dimensions[row].height=18
            for ci,v in enumerate([g['name'],g['w2'],"OUI" if same else "Non"],1):
                c=ws3.cell(row=row,column=ci,value=v)
                c.fill=fill("FFCCBC" if same else "FFF3E0")
                c.font=cfont(bold=(ci==1)); c.alignment=lft() if ci==1 else ctr(); c.border=bd()
                if ci==2: c.number_format='0.0'
            row+=1
    row+=1

# ═══ SHEET 4 — Stats ═════════════════════════════════════════════════════════
ws5=wb.create_sheet("Statistiques")
ws5.sheet_view.showGridLines=False
for col,w in zip([1,2,3],[30,20,22]): cw(ws5,col,w)
ws5.merge_cells("A1:C1")
ws5["A1"]="STATISTIQUES"
ws5["A1"].font=Font(name='Arial',bold=True,size=14,color='FFFFFF')
ws5["A1"].fill=fill("263238"); ws5["A1"].alignment=ctr(); ws5.row_dimensions[1].height=30

total_walk=sum(r['total'] for r in rows)
max_r=max(rows,key=lambda r:r['total']); min_r=min(rows,key=lambda r:r['total'])
same_count=sum(1 for r in rows if r['drinks_host']==r['dinner_host'])

stats=[
    ("","",""),("PARTICIPANTS","",""),
    ("Nombre total",len(rows),"personnes"),
    ("","",""),("APERITIF","",""),("Hotes",len(drinks_groups),""),
]
for h,g in drinks_groups.items():
    stats.append((f"  Chez {h}",f"{len(g)} invites",""))
stats+=[("","",""),("DINER","",""),("Hotes",len(dinner_groups),"")]
for h,g in dinner_groups.items():
    stats.append((f"  Chez {h}",f"{len(g)} invites",""))
stats+=[
    ("","",""),("TEMPS DE MARCHE","",""),
    ("Total cumule",f"{total_walk:.1f} min","tous"),
    ("Moyenne / personne",f"{total_walk/len(rows):.1f} min",""),
    ("Maximum",f"{max_r['total']:.1f} min",f"({max_r['name']})"),
    ("Minimum",f"{min_r['total']:.1f} min",f"({min_r['name']})"),
    ("","",""),("QUALITE","",""),
    ("Meme hote apero+diner",same_count,"ideal = 0"),
]
sections={"PARTICIPANTS","APERITIF","DINER","TEMPS DE MARCHE","QUALITE"}
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
