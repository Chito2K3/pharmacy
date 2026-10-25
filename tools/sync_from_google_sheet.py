#!/usr/bin/env python3
"""
Amang Rodriguez Memorial Medical Center (ARMMC)
Pharmacy Inventory Monitoring System — Automated Sheet Synchronizer

Fetches the live master Google Sheet CSV, dynamically detects all column
positions (immune to column shifts), and synchronizes:
1. items.json (Full masterlist of 2,665+ active items)
2. reorders.json (Replenishment queue calculated with 6mo & 1mo safety stock)
3. sheet_metadata.json (Update timestamps, row counts, freshness metrics)
"""

import urllib.request
import csv
import json
import os
import sys
from datetime import datetime

SHEET_ID = "1zLM0_lbs73MMEMgY-QeDOadVYMek8WJXPBxZyI5mvtM"
GID = "237243801"
CSV_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID}"

# Resolve paths relative to workspace root (parent of tools/)
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
WORKSPACE_DIR = os.path.dirname(SCRIPT_DIR)

ITEMS_PATH = os.path.join(WORKSPACE_DIR, "items.json")
REORDERS_PATH = os.path.join(WORKSPACE_DIR, "reorders.json")
METADATA_PATH = os.path.join(WORKSPACE_DIR, "sheet_metadata.json")

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

print("=" * 65)
print(" ARMMC Pharmacy IMS -- Live Google Sheet Synchronizer")
print("=" * 65)
print(f"Connecting to Google Sheet (ID: {SHEET_ID[:8]}... GID: {GID})...")

try:
    req = urllib.request.Request(CSV_URL, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
    with urllib.request.urlopen(req, timeout=30) as resp:
        raw_bytes = resp.read()
except Exception as e:
    print(f"[ERROR] Error fetching Google Sheet CSV: {e}")
    sys.exit(1)

csv_text = raw_bytes.decode('utf-8', errors='ignore')
print(f"[OK] Downloaded {len(csv_text):,} characters of CSV data.")

lines = csv_text.splitlines()
reader = csv.reader(lines)
rows = list(reader)

if len(rows) < 5:
    print("❌ Error: Downloaded CSV has insufficient rows.")
    sys.exit(1)

# Extract Manual Update Date from Cell B2 (Row index 1, Col index 1)
manual_date_raw = ""
if len(rows) > 1 and len(rows[1]) > 1:
    manual_date_raw = rows[1][1].strip()

# Locate Header Row
header_idx = -1
for idx in range(min(10, len(rows))):
    if len(rows[idx]) > 1 and rows[idx][1].strip().lower() == "item code":
        header_idx = idx
        break

if header_idx == -1:
    header_idx = 4

print(f"[OK] Header row detected at row index {header_idx}")

# ==============================================================================
# DYNAMIC COLUMN AUTO-DETECTION
# ==============================================================================
row4 = rows[3] if len(rows) > 3 else []
row5 = rows[4] if len(rows) > 4 else []
max_cols = max(len(row4), len(row5))

cols = {
    "dispensing_qty": -1,
    "storage_qty": -1,
    "warehouse_qty": -1,
    "consignment_qty": -1,
    "overall_start": -1,
    "overall_avg_monthly": -1,
    "overall_normalized": -1,
    "overall_total_qty": -1,
    "overall_value": -1,
    "overall_level_days": -1,
    "overall_impact_date": -1,
    "overall_pending_po": -1,
    "overall_ending_qty": -1,
    "overall_ending_days": -1,
    "overall_ending_impact_date": -1,
    "overall_epa_balance": -1,
    "overall_ending_epa": -1,
    "overall_ending_epa_days": -1,
    "overall_ending_epa_impact_date": -1,
}

# PASS 1: Identify Section Markers
for c in range(max_cols):
    cell4 = (row4[c] if c < len(row4) else '').strip().upper()
    cell5 = (row5[c] if c < len(row5) else '').strip().upper()
    comb = f"{cell4} {cell5}".strip()

    if cols["dispensing_qty"] == -1:
        if cell4 == "DISPENSING AREA" or ("DISPENSING" in comb and ("(QTY)" in comb or "VOLUME" in comb)):
            cols["dispensing_qty"] = c
    if cols["storage_qty"] == -1:
        if cell4 == "PHARMACY STORAGE" or ("STORAGE" in comb and ("(QTY)" in comb or "VOLUME" in comb)):
            cols["storage_qty"] = c
    if cols["warehouse_qty"] == -1:
        if cell4 == "WAREHOUSE" or ("WAREHOUSE" in comb and ("(QTY)" in comb or "VOLUME" in comb)):
            cols["warehouse_qty"] = c
    if cols["consignment_qty"] == -1:
        if cell4 == "CONSIGNMENT" or ("CONSIGNMENT" in comb and ("(QTY)" in comb or "VOLUME" in comb)):
            cols["consignment_qty"] = c
    if cols["overall_start"] == -1:
        if "INVENTORY IMPACT" in cell4 or ("AVERAGE MONTHLY CONSUMPTION" in cell5 and (cols["dispensing_qty"] == -1 or c < cols["dispensing_qty"])):
            cols["overall_start"] = c

# PASS 2: Identify Metrics in Inventory Impact Section
scan_start = cols["overall_start"] if cols["overall_start"] != -1 else 36
scan_end = cols["dispensing_qty"] if cols["dispensing_qty"] != -1 else (cols["overall_start"] + 16 if cols["overall_start"] != -1 else max_cols)

for c in range(scan_start, scan_end):
    cell4 = (row4[c] if c < len(row4) else '').strip().upper()
    cell5 = (row5[c] if c < len(row5) else '').strip().upper()

    if cols["overall_avg_monthly"] == -1 and "AVERAGE MONTHLY CONSUMPTION" in cell5 and "NORMALIZED" not in cell5:
        cols["overall_avg_monthly"] = c
    elif cols["overall_normalized"] == -1 and ("NORMALIZED DEMAND" in cell5 or "AVERAGE MONTHLY NORMALIZED" in cell5):
        cols["overall_normalized"] = c
    elif cols["overall_total_qty"] == -1 and ("TOTAL INVENTORY VOLUME" in cell5 or "TOTAL INVENTORY(QTY)" in cell5) and "ENDING" not in cell5:
        cols["overall_total_qty"] = c
    elif cols["overall_value"] == -1 and "INVENTORY" in cell5 and "VALUE" in cell5 and "ENDING" not in cell5 and "HOLDING" not in cell4:
        cols["overall_value"] = c
    elif cols["overall_level_days"] == -1 and "INVENTORY LEVEL DAYS" in cell5 and "ENDING" not in cell5:
        cols["overall_level_days"] = c
    elif cols["overall_impact_date"] == -1 and ("DATE OF IMPACT" in cell5 or "IMPACT DATE" in cell5) and "ENDING" not in cell5:
        cols["overall_impact_date"] = c
    elif cols["overall_pending_po"] == -1 and ("PENDING PO" in cell5 or "PO / CO" in cell5 or "PO/CO" in cell5 or "QTY OF PENDING" in cell5 or "PENDING PO/CO" in cell4):
        cols["overall_pending_po"] = c
    elif cols["overall_ending_qty"] == -1 and "ENDING INVENTORY" in cell5 and "DAYS" not in cell5 and "EPA" not in cell5 and "IMPACT" not in cell5:
        cols["overall_ending_qty"] = c
    elif cols["overall_ending_days"] == -1 and "ENDING INVENTORY LEVEL DAYS" in cell5 and "EPA" not in cell5:
        cols["overall_ending_days"] = c
    elif cols["overall_ending_impact_date"] == -1 and "ENDING IMPACT DATE" in cell5 and "EPA" not in cell5:
        cols["overall_ending_impact_date"] = c
    elif cols["overall_epa_balance"] == -1 and "EPA" in cell5 and ("CALL-OFF" in cell5 or "BALANCE" in cell5) and "ENDING INVENTORY" not in cell5:
        cols["overall_epa_balance"] = c
    elif cols["overall_ending_epa"] == -1 and "ENDING INVENTORY" in cell5 and "EPA" in cell5 and "DAYS" not in cell5 and "IMPACT" not in cell5:
        cols["overall_ending_epa"] = c
    elif cols["overall_ending_epa_days"] == -1 and "ENDING INVENTORY LEVEL DAYS" in cell5 and "EPA" in cell5:
        cols["overall_ending_epa_days"] = c
    elif cols["overall_ending_epa_impact_date"] == -1 and "ENDING IMPACT DATE" in cell5 and "EPA" in cell5:
        cols["overall_ending_epa_impact_date"] = c

# Safe Anchored Fallbacks
s = cols["overall_start"] if cols["overall_start"] != -1 else 36
if cols["overall_avg_monthly"] == -1: cols["overall_avg_monthly"] = s
if cols["overall_normalized"] == -1: cols["overall_normalized"] = s + 1
if cols["overall_total_qty"] == -1: cols["overall_total_qty"] = s + 2
if cols["overall_value"] == -1: cols["overall_value"] = s + 3
if cols["overall_level_days"] == -1: cols["overall_level_days"] = s + 4
if cols["overall_impact_date"] == -1: cols["overall_impact_date"] = s + 5
if cols["overall_pending_po"] == -1: cols["overall_pending_po"] = s + 7
if cols["overall_ending_qty"] == -1: cols["overall_ending_qty"] = s + 8
if cols["overall_ending_days"] == -1: cols["overall_ending_days"] = s + 9
if cols["overall_ending_impact_date"] == -1: cols["overall_ending_impact_date"] = s + 10
if cols["overall_epa_balance"] == -1: cols["overall_epa_balance"] = s + 11
if cols["overall_ending_epa"] == -1: cols["overall_ending_epa"] = s + 12
if cols["overall_ending_epa_days"] == -1: cols["overall_ending_epa_days"] = s + 13
if cols["overall_ending_epa_impact_date"] == -1: cols["overall_ending_impact_date"] = s + 14

if cols["dispensing_qty"] == -1: cols["dispensing_qty"] = 52
if cols["storage_qty"] == -1: cols["storage_qty"] = cols["dispensing_qty"] + 7
if cols["warehouse_qty"] == -1: cols["warehouse_qty"] = cols["storage_qty"] + 7
if cols["consignment_qty"] == -1: cols["consignment_qty"] = cols["warehouse_qty"] + 7

print(f"[OK] Columns mapped: Overall={cols['overall_start']}, Disp={cols['dispensing_qty']}, Stor={cols['storage_qty']}, Ware={cols['warehouse_qty']}, Cons={cols['consignment_qty']}")

def parse_num(val):
    if not val:
        return 0.0
    s = str(val).replace(',', '').strip()
    if s in ('-', '—') or s.startswith('#'):
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0

def clean_str(val):
    if not val:
        return ""
    s = str(val).strip()
    return "" if s in ('-', '—') else s

def get_col(row, idx):
    if idx >= 0 and idx < len(row):
        return row[idx]
    return ""

items = []
reorders = []

for r_idx in range(header_idx + 1, len(rows)):
    row = rows[r_idx]
    if len(row) <= 1:
        continue

    no = clean_str(row[0] if len(row) > 0 else "")
    item_code = clean_str(row[1] if len(row) > 1 else "")
    dci_code = clean_str(row[2] if len(row) > 2 else "")
    desc = clean_str(row[3] if len(row) > 3 else "")
    gen_name = clean_str(row[4] if len(row) > 4 else "")
    cat = clean_str(row[5] if len(row) > 5 else "")
    pharm_cat = clean_str(row[6] if len(row) > 6 else "")
    uom = clean_str(row[7] if len(row) > 7 else "")

    if not no and not item_code and not desc:
        continue
    if item_code.lower() == "item code" or desc.lower() == "item description":
        continue

    if not item_code:
        item_code = dci_code if dci_code else f"ITEM-{no}"

    acq_cost = parse_num(row[8] if len(row) > 8 else 0)
    sell_price = parse_num(row[9] if len(row) > 9 else 0)

    # Historical & Annual metrics
    # 2024: Cols 10 to 13 (Sep-24 to Dec-24, 4 months annualized)
    sum_2024 = sum(parse_num(get_col(row, c)) for c in range(10, 14))
    annual_qty_2024 = (sum_2024 / 4.0) * 12.0

    # 2025: Cols 14 to 25 (Jan-25 to Dec-25, 12 full months)
    annual_qty_2025 = sum(parse_num(get_col(row, c)) for c in range(14, 26))

    # 2026: Cols 26 to 34 (Jan-26 to Sep-26, 9 months annualized)
    sum_2026 = sum(parse_num(get_col(row, c)) for c in range(26, 35))
    annual_qty_2026 = (sum_2026 / 9.0) * 12.0

    # Recent 12-month series
    recent_12 = [parse_num(get_col(row, c)) for c in range(23, 35)]
    cal_2025 = [parse_num(get_col(row, c)) for c in range(14, 26)]
    monthly_series = recent_12 if any(v > 0 for v in recent_12) else cal_2025

    # Overall Metrics
    avg_monthly = parse_num(get_col(row, cols["overall_avg_monthly"]))
    avg_normalized = parse_num(get_col(row, cols["overall_normalized"]))
    total_qty = parse_num(get_col(row, cols["overall_total_qty"]))
    inv_value = parse_num(get_col(row, cols["overall_value"]))
    level_days = parse_num(get_col(row, cols["overall_level_days"]))
    impact_date = clean_str(get_col(row, cols["overall_impact_date"]))
    pending_po = parse_num(get_col(row, cols["overall_pending_po"]))
    ending_qty = parse_num(get_col(row, cols["overall_ending_qty"]))
    ending_days = parse_num(get_col(row, cols["overall_ending_days"]))
    ending_impact = clean_str(get_col(row, cols["overall_ending_impact_date"]))
    epa_bal = parse_num(get_col(row, cols["overall_epa_balance"]))
    ending_with_epa_qty = parse_num(get_col(row, cols["overall_ending_epa"]))
    ending_with_epa_days = parse_num(get_col(row, cols["overall_ending_epa_days"]))

    # Sub-location breakdowns
    dq = cols["dispensing_qty"]
    disp_qty = parse_num(get_col(row, dq))
    disp_val = parse_num(get_col(row, dq + 1))
    disp_avg_m = parse_num(get_col(row, dq + 2))
    disp_avg_d = parse_num(get_col(row, dq + 3))
    disp_days = parse_num(get_col(row, dq + 4))

    sq = cols["storage_qty"]
    stor_qty = parse_num(get_col(row, sq))
    stor_val = parse_num(get_col(row, sq + 1))
    stor_avg_m = parse_num(get_col(row, sq + 2))
    stor_avg_d = parse_num(get_col(row, sq + 3))
    stor_days = parse_num(get_col(row, sq + 4))

    wq = cols["warehouse_qty"]
    ware_qty = parse_num(get_col(row, wq))
    ware_val = parse_num(get_col(row, wq + 1))
    ware_avg_m = parse_num(get_col(row, wq + 2))
    ware_avg_d = parse_num(get_col(row, wq + 3))
    ware_days = parse_num(get_col(row, wq + 4))

    cq = cols["consignment_qty"]
    cons_qty = parse_num(get_col(row, cq))
    cons_val = parse_num(get_col(row, cq + 1))
    cons_avg_m = parse_num(get_col(row, cq + 2))
    cons_avg_d = parse_num(get_col(row, cq + 3))
    cons_days = parse_num(get_col(row, cq + 4))

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
        "annual_qty_2024": f"{annual_qty_2024:.2f}",
        "annual_val_2024": f"{(annual_qty_2024 * acq_cost):.2f}",
        "annual_qty_2025": f"{annual_qty_2025:.2f}",
        "annual_val_2025": f"{(annual_qty_2025 * acq_cost):.2f}",
        "annual_qty_2026": f"{annual_qty_2026:.2f}",
        "annual_val_2026": f"{(annual_qty_2026 * acq_cost):.2f}",
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

    # Reorder calculation matching Code.gs
    eff_avg = avg_normalized if avg_normalized > 0 else avg_monthly
    rop6 = max(0, round(eff_avg * 3 + eff_avg * 6 - total_qty))
    rop1 = max(0, round(eff_avg * 3 + eff_avg * 1 - total_qty))
    calc_days = round((total_qty / eff_avg) * 30) if eff_avg > 0 else 0
    stock_stat = "Out of Stock" if total_qty <= 0 else ("Critical" if calc_days < 30 else ("Low" if calc_days < 60 else ("Normal" if calc_days < 120 else "Over Stock")))

    if rop6 > 0:
        reorders.append({
            "item_code": item_code,
            "dci_code": dci_code,
            "description": desc,
            "generic_name": gen_name,
            "pharmacologic_category": pharm_cat,
            "pharmacy_category": cat,
            "unit_of_measure": uom,
            "avg_monthly_consumption": str(avg_monthly),
            "avg_monthly_normalized_demand": str(avg_normalized),
            "total_inventory_qty": str(total_qty),
            "inventory_level_days": str(level_days if level_days > 0 else calc_days),
            "stock_status": stock_stat,
            "reorder_qty_6mo_safety": str(rop6),
            "reorder_qty_1mo_safety": str(rop1),
            "pending_po_call_off": str(pending_po),
            "remarks": "Request in full" if rop6 > 0 else "",
            "calculated_at": datetime.utcnow().isoformat() + "Z"
        })

# Sort reorders by days ascending (most urgent first)
reorders.sort(key=lambda r: parse_num(r.get("inventory_level_days", 0)))

# Calculate days since manual sheet update
days_since_update = 0
if manual_date_raw:
    try:
        parts = [int(p) for p in manual_date_raw.split('/')]
        if len(parts) == 3:
            # Format: MM/DD/YYYY
            sheet_dt = datetime(parts[2], parts[0], parts[1])
            diff = datetime.now() - sheet_dt
            days_since_update = max(0, diff.days)
    except Exception:
        pass

# Calculate Content Fingerprint and Stock Totals
total_stock_qty = sum(parse_num(it.get("total_inventory_qty", 0)) for it in items)
total_inv_val = sum(parse_num(it.get("inventory_value_php", 0)) for it in items)
total_pending_po = sum(parse_num(it.get("pending_po_co_qty", 0) or it.get("pending_po_call_off", 0)) for it in items)

h = 0
for i, it in enumerate(items):
    q = parse_num(it.get("total_inventory_qty", 0))
    po = parse_num(it.get("pending_po_co_qty", 0) or it.get("pending_po_call_off", 0))
    code = it.get("item_code") or it.get("dci_code") or str(i)
    s = f"{code}:{q}:{po}"
    for ch in s:
        h = ((h << 5) - h) + ord(ch)
        h = ((h & 0xFFFFFFFF) ^ 0x80000000) - 0x80000000
content_fingerprint = f"{len(items)}_{int(round(total_stock_qty))}_{int(round(total_inv_val))}_{int(round(total_pending_po))}_{hex(h & 0xFFFFFFFF)[2:]}"

metadata = {
    "manualDateRaw": manual_date_raw or "10/22/2026",
    "manualDateFormatted": manual_date_raw or "10/22/2026",
    "daysSinceCellDate": days_since_update,
    "effectiveDate": datetime.utcnow().isoformat() + "Z",
    "effectiveType": "content_verified",
    "daysSinceUpdate": 0,
    "isStale": False,
    "staleThresholdDays": 3,
    "contentFingerprint": content_fingerprint,
    "totalItems": len(items),
    "totalReorders": len(reorders),
    "totalStockQty": int(round(total_stock_qty)),
    "totalInventoryValue": round(total_inv_val, 2),
    "lastItemNo": items[-1]["no"] if items else "",
    "lastItemCode": items[-1]["item_code"] if items else "",
    "sheetUrl": f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit?gid={GID}#gid={GID}",
    "syncTimestamp": datetime.utcnow().isoformat() + "Z"
}

# Write files
with open(ITEMS_PATH, 'w', encoding='utf-8') as f:
    json.dump(items, f, indent=2, ensure_ascii=False)
print(f"[OK] Saved {len(items):,} items to {ITEMS_PATH}")

with open(REORDERS_PATH, 'w', encoding='utf-8') as f:
    json.dump(reorders, f, indent=2, ensure_ascii=False)
print(f"[OK] Saved {len(reorders):,} replenishment items to {REORDERS_PATH}")

with open(METADATA_PATH, 'w', encoding='utf-8') as f:
    json.dump(metadata, f, indent=2, ensure_ascii=False)
print(f"[OK] Saved sheet metadata to {METADATA_PATH}")

print("=" * 65)
print(f"[DONE] Synchronization Complete: {len(items):,} items | {len(reorders):,} reorders")
print(f"   Sheet Manual Date: {manual_date_raw} ({days_since_update} days ago)")
print("=" * 65)
