import urllib.request
import csv
import json
import os
import re

SHEET_ID = "1zLM0_lbs73MMEMgY-QeDOadVYMek8WJXPBxZyI5mvtM"
GID = "237243801"
CSV_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID}"
OUTPUT_PATH = "items.json"

print(f"Fetching live CSV from Google Sheet (GID: {GID})...")
req = urllib.request.Request(CSV_URL, headers={'User-Agent': 'Mozilla/5.0'})
with urllib.request.urlopen(req) as resp:
    raw_bytes = resp.read()

csv_text = raw_bytes.decode('utf-8', errors='ignore')
print(f"Downloaded {len(csv_text):,} characters of CSV data.")

lines = csv_text.splitlines()

# Find row where "Item Code" appears in header
reader = csv.reader(lines)
rows = list(reader)

header_idx = -1
for idx, r in enumerate(rows):
    if len(r) > 1 and r[1].strip().lower() == "item code":
        header_idx = idx
        break

if header_idx == -1:
    print("Error: Could not locate 'Item Code' header row.")
    exit(1)

print(f"Header located at row index {header_idx}")
headers = rows[header_idx]

def parse_num(val):
    if not val:
        return 0.0
    s = str(val).replace(',', '').strip()
    if s == '-' or s == '—' or s.startswith('#'):
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0

def clean_str(val):
    if not val:
        return ""
    s = str(val).strip()
    if s == '-' or s == '—':
        return ""
    return s

items = []

for r_idx in range(header_idx + 1, len(rows)):
    row = rows[r_idx]
    if len(row) <= 1:
        continue
    
    item_code = clean_str(row[1])
    if not item_code or item_code.lower() == "item code":
        continue
    
    no = clean_str(row[0])
    dci_code = clean_str(row[2])
    desc = clean_str(row[3])
    gen_name = clean_str(row[4])
    cat = clean_str(row[5])
    pharm_cat = clean_str(row[6])
    uom = clean_str(row[7])
    acq_cost = parse_num(row[8])
    sell_price = parse_num(row[9])

    # 12-Month Consumption: Most recent 12 months (Oct-25 to Sep-26: Cols 23 to 34)
    # If recent 12 months are all zeros, fallback to 2025 Jan-Dec (Cols 14 to 25)
    recent_12 = [parse_num(row[c]) if c < len(row) else 0.0 for c in range(23, 35)]
    cal_2025 = [parse_num(row[c]) if c < len(row) else 0.0 for c in range(14, 26)]
    
    monthly_series = recent_12 if any(v > 0 for v in recent_12) else cal_2025

    avg_monthly = parse_num(row[36]) if len(row) > 36 else 0.0
    avg_normalized = parse_num(row[37]) if len(row) > 37 else 0.0
    total_qty = parse_num(row[38]) if len(row) > 38 else 0.0
    inv_value = parse_num(row[39]) if len(row) > 39 else 0.0
    level_days = parse_num(row[40]) if len(row) > 40 else 0.0
    impact_date = clean_str(row[41]) if len(row) > 41 else ""
    
    pending_po = parse_num(row[43]) if len(row) > 43 else 0.0
    ending_qty = parse_num(row[44]) if len(row) > 44 else 0.0
    ending_days = parse_num(row[45]) if len(row) > 45 else 0.0
    ending_impact = clean_str(row[46]) if len(row) > 46 else ""
    epa_bal = parse_num(row[47]) if len(row) > 47 else 0.0
    ending_with_epa_qty = parse_num(row[48]) if len(row) > 48 else 0.0
    ending_with_epa_days = parse_num(row[49]) if len(row) > 49 else 0.0

    # Sub-location breakdowns
    # Dispensing: Cols 52-57
    disp_qty = parse_num(row[52]) if len(row) > 52 else 0.0
    disp_val = parse_num(row[53]) if len(row) > 53 else 0.0
    disp_avg_m = parse_num(row[54]) if len(row) > 54 else 0.0
    disp_avg_d = parse_num(row[55]) if len(row) > 55 else 0.0
    disp_days = parse_num(row[56]) if len(row) > 56 else 0.0

    # Storage: Cols 59-64
    stor_qty = parse_num(row[59]) if len(row) > 59 else 0.0
    stor_val = parse_num(row[60]) if len(row) > 60 else 0.0
    stor_avg_m = parse_num(row[61]) if len(row) > 61 else 0.0
    stor_avg_d = parse_num(row[62]) if len(row) > 62 else 0.0
    stor_days = parse_num(row[63]) if len(row) > 63 else 0.0

    # Warehouse: Cols 66-71
    ware_qty = parse_num(row[66]) if len(row) > 66 else 0.0
    ware_val = parse_num(row[67]) if len(row) > 67 else 0.0
    ware_avg_m = parse_num(row[68]) if len(row) > 68 else 0.0
    ware_avg_d = parse_num(row[69]) if len(row) > 69 else 0.0
    ware_days = parse_num(row[70]) if len(row) > 70 else 0.0

    # Consignment: Cols 73-78
    cons_qty = parse_num(row[73]) if len(row) > 73 else 0.0
    cons_val = parse_num(row[74]) if len(row) > 74 else 0.0
    cons_avg_m = parse_num(row[75]) if len(row) > 75 else 0.0
    cons_avg_d = parse_num(row[76]) if len(row) > 76 else 0.0
    cons_days = parse_num(row[77]) if len(row) > 77 else 0.0

    item_obj = {
        "no": no,
        "item_code": item_code,
        "dci_code": dci_code,
        "description": desc,
        "generic_name": gen_name,
        "pharmacy_category": cat,
        "pharmacologic_category": pharm_cat,
        "unit_of_measure": uom,
        "unit_cost": f"{acq_cost:.2f}" if acq_cost > 0 else "0",
        "selling_price": f"{sell_price:.2f}" if sell_price > 0 else "0",
        "acquisition_price": f"{acq_cost:.2f}" if acq_cost > 0 else "0",
        "avg_monthly_consumption": str(avg_monthly),
        "avg_monthly_normalized_demand": str(avg_normalized),
        "total_inventory_qty": str(total_qty),
        "inventory_value_php": f"{inv_value:.2f}" if inv_value > 0 else "0",
        "inventory_level_days": str(level_days),
        "date_of_impact": impact_date,
        "pending_po_co_qty": str(pending_po),
        "ending_inventory_qty": str(ending_qty),
        "ending_inventory_level_days": str(ending_days),
        "ending_impact_date": ending_impact,
        "epa_cy2026_balance": str(epa_bal),
        "ending_inv_with_epa_qty": str(ending_with_epa_qty),
        "ending_inv_with_epa_days": str(ending_with_epa_days),
        "dispensing_inventory_qty": str(disp_qty),
        "dispensing_inventory_value": str(disp_val),
        "dispensing_avg_monthly": str(disp_avg_m),
        "dispensing_avg_daily": str(disp_avg_d),
        "dispensing_level_days": str(disp_days),
        "storage_inventory_qty": str(stor_qty),
        "storage_inventory_value": str(stor_val),
        "storage_avg_monthly": str(stor_avg_m),
        "storage_avg_daily": str(stor_avg_d),
        "storage_level_days": str(stor_days),
        "warehouse_inventory_qty": str(ware_qty),
        "warehouse_inventory_value": str(ware_val),
        "warehouse_avg_monthly": str(ware_avg_m),
        "warehouse_avg_daily": str(ware_avg_d),
        "warehouse_level_days": str(ware_days),
        "consignment_inventory_qty": str(cons_qty),
        "consignment_inventory_value": str(cons_val),
        "consignment_avg_monthly": str(cons_avg_m),
        "consignment_avg_daily": str(cons_avg_d),
        "consignment_level_days": str(cons_days),
        "monthly_consumption": monthly_series
    }
    items.append(item_obj)

print(f"Processed {len(items):,} items from Google Sheet.")

# Write to items.json
with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
    json.dump(items, f, indent=2, ensure_ascii=False)

print(f"Successfully saved {len(items):,} items to {OUTPUT_PATH}!")

# Sample verification
ampi = next((it for it in items if it["dci_code"] == "DMR000072" or "AMPICILLIN SODIUM 250MG" in it["description"]), None)
if ampi:
    print(f"\n--- Verification on Ampicillin ({ampi['dci_code']}) ---")
    print(f"Description: {ampi['description']}")
    print(f"Total Stock: {ampi['total_inventory_qty']}")
    print(f"Dispensing:  {ampi['dispensing_inventory_qty']}")
    print(f"Storage:     {ampi['storage_inventory_qty']}")
    print(f"Warehouse:   {ampi['warehouse_inventory_qty']}")
    print(f"Consignment: {ampi['consignment_inventory_qty']}")
    print(f"Stock Days:  {ampi['inventory_level_days']}")
    print(f"Monthly:     {ampi['monthly_consumption']}")
