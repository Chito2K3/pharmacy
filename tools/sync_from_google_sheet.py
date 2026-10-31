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
NEAR_EXPIRY_GID = "2052140349"
CSV_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID}"
NEAR_EXPIRY_CSV_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={NEAR_EXPIRY_GID}"

# Resolve paths relative to workspace root (parent of tools/)
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
WORKSPACE_DIR = os.path.dirname(SCRIPT_DIR)

ITEMS_PATH = os.path.join(WORKSPACE_DIR, "items.json")
REORDERS_PATH = os.path.join(WORKSPACE_DIR, "reorders.json")
NEAR_EXPIRY_PATH = os.path.join(WORKSPACE_DIR, "near_expiring.json")
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
if cols["overall_ending_epa_impact_date"] == -1: cols["overall_ending_epa_impact_date"] = s + 14

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

# ==============================================================================
# NEAR EXPIRING DRUGS SYNCHRONIZATION (GID: 2052140349)
# ==============================================================================
print(f"Fetching Near Expiring Drugs tab (GID: {NEAR_EXPIRY_GID})...")
near_expiry_items = []
near_expiry_as_of = ""
near_expiry_tiers = []
near_expiry_fingerprint = "0_0_0"

try:
    ne_req = urllib.request.Request(NEAR_EXPIRY_CSV_URL, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
    with urllib.request.urlopen(ne_req, timeout=30) as ne_resp:
        ne_bytes = ne_resp.read()
    ne_text = ne_bytes.decode('utf-8', errors='ignore')
    ne_rows = list(csv.reader(ne_text.splitlines()))

    # Extract 'as of' date from top rows (typically Row 2, Col B)
    for r_i in range(min(5, len(ne_rows))):
        for c_val in ne_rows[r_i]:
            s_val = str(c_val or "").strip()
            if "as of" in s_val.lower():
                near_expiry_as_of = s_val
                break
        if near_expiry_as_of:
            break

    # Dynamic Header & Column Auto-Detection
    ne_header_idx = 3
    ne_cols = {
        "no": 0, "item_code": 1, "description": 2, "total_qty": 3, "expiration_date": 4,
        "dispensing_qty": 6, "storage_qty": 7, "warehouse_qty": 8, "loc_total": 9,
        "audit_note": 10, "impact_date": 11, "end_users": 12, "status": 13,
        "gl_note": 14, "followup_note": 15, "extra_note": 16
    }

    for r_i in range(min(10, len(ne_rows))):
        row_upper = [str(x or "").strip().upper() for x in ne_rows[r_i]]
        if any("ITEM DESCRIPTION" in x or "EXPIRATION DATE" in x for x in row_upper):
            ne_header_idx = r_i
            for c_i, h_txt in enumerate(row_upper):
                if h_txt in ("NO.", "NO", "#"):
                    ne_cols["no"] = c_i
                elif "ITEM CODE" in h_txt or "DCI CODE" in h_txt:
                    ne_cols["item_code"] = c_i
                elif "DESCRIPTION" in h_txt:
                    ne_cols["description"] = c_i
                elif "TOTAL STOCK" in h_txt or "STOCKS ON HAND" in h_txt:
                    ne_cols["total_qty"] = c_i
                elif "EXPIRATION" in h_txt:
                    ne_cols["expiration_date"] = c_i
                elif "DISPENSING" in h_txt:
                    ne_cols["dispensing_qty"] = c_i
                elif "STORAGE" in h_txt:
                    ne_cols["storage_qty"] = c_i
                elif "WAREHOUSE" in h_txt:
                    ne_cols["warehouse_qty"] = c_i
                elif h_txt == "TOTAL":
                    ne_cols["loc_total"] = c_i
                elif "IMPACT DATE" in h_txt:
                    ne_cols["impact_date"] = c_i
                elif "END-USER" in h_txt or "END USER" in h_txt or "REQUESTING" in h_txt:
                    ne_cols["end_users"] = c_i
                elif h_txt == "STATUS" or "DISPOSITION" in h_txt:
                    ne_cols["status"] = c_i
            break

    # Build lookup by DCI code / Item code for unit cost & masterlist enrichment
    dci_lookup = {}
    for it in items:
        d_code = (it.get("dci_code") or "").strip().upper()
        i_code = (it.get("item_code") or "").strip().upper()
        if d_code and d_code not in dci_lookup:
            dci_lookup[d_code] = it
        if i_code and i_code not in dci_lookup:
            dci_lookup[i_code] = it

    current_tier = "EXPIRING WITHIN 3 MONTHS"
    seq_id = 0

    for r_i in range(ne_header_idx + 1, len(ne_rows)):
        r = ne_rows[r_i]
        if not r or all(not str(x or "").strip() for x in r):
            continue

        # Detect Section Banner row (e.g. "EXPIRING WITHIN 3 MONTHS", "EXPIRING WITHIN 6 MONTHS", etc.)
        joined_upper = " ".join(str(x or "").strip().upper() for x in r if str(x or "").strip())
        c1_upper = clean_str(get_col(r, ne_cols["item_code"])).upper()
        c2_upper = clean_str(get_col(r, ne_cols["description"])).upper()

        if ("EXPIRING WITHIN" in c1_upper or "EXPIRING WITHIN" in c2_upper or "ALREADY EXPIRED" in c1_upper) and not clean_str(get_col(r, ne_cols["expiration_date"])):
            banner = c1_upper if "EXPIR" in c1_upper else c2_upper
            current_tier = banner
            if current_tier not in near_expiry_tiers:
                near_expiry_tiers.append(current_tier)
            continue

        no_val = clean_str(get_col(r, ne_cols["no"]))
        code_val = clean_str(get_col(r, ne_cols["item_code"]))
        desc_val = clean_str(get_col(r, ne_cols["description"]))
        exp_val = clean_str(get_col(r, ne_cols["expiration_date"]))
        tot_qty = parse_num(get_col(r, ne_cols["total_qty"]))

        if no_val.upper() in ("NO.", "NO") and ("EXPIR" in code_val.upper() or not desc_val):
            if "EXPIR" in code_val.upper():
                current_tier = code_val.upper()
                if current_tier not in near_expiry_tiers:
                    near_expiry_tiers.append(current_tier)
            continue

        # Skip trailing notes that have no code, no expiration date, and 0 stock
        if not code_val and not exp_val and tot_qty <= 0:
            continue

        if current_tier not in near_expiry_tiers:
            near_expiry_tiers.append(current_tier)

        disp_q = parse_num(get_col(r, ne_cols["dispensing_qty"]))
        stor_q = parse_num(get_col(r, ne_cols["storage_qty"]))
        ware_q = parse_num(get_col(r, ne_cols["warehouse_qty"]))
        audit_note = clean_str(get_col(r, ne_cols["audit_note"]))
        impact_date = clean_str(get_col(r, ne_cols["impact_date"]))
        end_users = clean_str(get_col(r, ne_cols["end_users"]))
        status_val = clean_str(get_col(r, ne_cols["status"]))
        gl_note = clean_str(get_col(r, ne_cols["gl_note"]))
        followup_note = clean_str(get_col(r, ne_cols["followup_note"]))
        extra_note = clean_str(get_col(r, ne_cols["extra_note"]))

        master_it = dci_lookup.get(code_val.upper())
        unit_cost = parse_num(master_it.get("unit_cost", 0)) if master_it else 0.0
        master_item_code = master_it.get("item_code", "") if master_it else ""
        pharm_cat = master_it.get("pharmacy_category", "") if master_it else ("Donation" if "DON" in code_val.upper() else "")

        seq_id += 1
        near_expiry_items.append({
            "id": seq_id,
            "no": no_val or str(seq_id),
            "item_code": code_val,
            "master_item_code": master_item_code,
            "description": desc_val,
            "pharmacy_category": pharm_cat,
            "total_qty": tot_qty,
            "dispensing_qty": disp_q,
            "storage_qty": stor_q,
            "warehouse_qty": ware_q,
            "expiration_date": exp_val,
            "horizon_tier": current_tier,
            "audit_note": audit_note,
            "impact_date": impact_date,
            "requesting_end_users": end_users,
            "status": status_val,
            "gl_note": gl_note,
            "followup_note": followup_note,
            "extra_note": extra_note,
            "unit_cost": round(unit_cost, 2),
            "estimated_value_php": round(tot_qty * unit_cost, 2)
        })

    # Compute Near Expiry content fingerprint
    ne_total_qty = sum(x["total_qty"] for x in near_expiry_items)
    ne_h = 0
    for x in near_expiry_items:
        s_ne = f"{x['item_code']}:{x['horizon_tier']}:{int(round(x['total_qty']))}:{x['expiration_date']}:{x['impact_date']}:{x['requesting_end_users']}:{x['status']}:{x['gl_note']}"
        for ch in s_ne:
            ne_h = ((ne_h << 5) - ne_h) + ord(ch)
            ne_h = ((ne_h & 0xFFFFFFFF) ^ 0x80000000) - 0x80000000
    near_expiry_fingerprint = f"{len(near_expiry_items)}_{int(round(ne_total_qty))}_{hex(ne_h & 0xFFFFFFFF)[2:]}"
    print(f"[OK] Parsed {len(near_expiry_items)} near-expiring items across {len(near_expiry_tiers)} tiers ({near_expiry_as_of}).")
except Exception as e_ne:
    print(f"[WARN] Could not fetch Near Expiring Drugs CSV: {e_ne}")

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
    "nearExpiryGid": NEAR_EXPIRY_GID,
    "nearExpirySheetUrl": f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit?gid={NEAR_EXPIRY_GID}#gid={NEAR_EXPIRY_GID}",
    "nearExpiryAsOfDate": near_expiry_as_of,
    "nearExpiryFingerprint": near_expiry_fingerprint,
    "totalNearExpiring": len(near_expiry_items),
    "syncTimestamp": datetime.utcnow().isoformat() + "Z"
}

# Write files
with open(ITEMS_PATH, 'w', encoding='utf-8') as f:
    json.dump(items, f, indent=2, ensure_ascii=False)
print(f"[OK] Saved {len(items):,} items to {ITEMS_PATH}")

with open(REORDERS_PATH, 'w', encoding='utf-8') as f:
    json.dump(reorders, f, indent=2, ensure_ascii=False)
print(f"[OK] Saved {len(reorders):,} replenishment items to {REORDERS_PATH}")

if near_expiry_items:
    near_expiry_payload = {
        "asOfDate": near_expiry_as_of,
        "gid": NEAR_EXPIRY_GID,
        "sheetUrl": f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit?gid={NEAR_EXPIRY_GID}#gid={NEAR_EXPIRY_GID}",
        "fingerprint": near_expiry_fingerprint,
        "totalItems": len(near_expiry_items),
        "tiers": near_expiry_tiers,
        "syncedAt": datetime.utcnow().isoformat() + "Z",
        "items": near_expiry_items
    }
    with open(NEAR_EXPIRY_PATH, 'w', encoding='utf-8') as f:
        json.dump(near_expiry_payload, f, indent=2, ensure_ascii=False)
    print(f"[OK] Saved {len(near_expiry_items):,} near-expiring items to {NEAR_EXPIRY_PATH}")

with open(METADATA_PATH, 'w', encoding='utf-8') as f:
    json.dump(metadata, f, indent=2, ensure_ascii=False)
print(f"[OK] Saved sheet metadata to {METADATA_PATH}")

print("=" * 65)
print(f"[DONE] Synchronization Complete: {len(items):,} items | {len(reorders):,} reorders | {len(near_expiry_items):,} near-expiring")
print(f"   Sheet Manual Date: {manual_date_raw} ({days_since_update} days ago)")
print("=" * 65)
