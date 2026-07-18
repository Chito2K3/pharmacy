// =============================================================
//  PharmaDash — Google Apps Script Backend (Code.gs)
//  Paste this entire file into your Google Apps Script editor.
//  Then Deploy > New Deployment > Web App > Anyone can access.
// =============================================================

// ---- SHEET NAMES ----
const SHEET_ITEMS    = 'Inventory Utilization Report';
const SHEET_REORDERS = 'Reorders';
const SHEET_USERS    = 'Users';
const SHEET_LOG      = 'AuditLog';

// =============================================================
//  HTTP ENTRY POINTS
// =============================================================

function doGet(e) {
  const params = e ? e.parameter : {};
  const action = params.action || '';
  const email  = params.email  || '';
  const token  = params.token  || '';

  // CORS-friendly response helper
  const respond = (data) =>
    ContentService
      .createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);

  try {
    switch (action) {
      case 'getItems':     return respond(getItems());
      case 'getReorders':  return respond(getReorders());
      case 'getUser':      return respond(getUser(email));
      case 'getStats':     return respond(getStats());
      case 'getSheetInfo': return respond(getSheetInfo()); // debug endpoint
      case 'ping':         return respond({ status: 'ok', timestamp: new Date().toISOString() });
      default:
        return respond({ error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return respond({ error: err.message, stack: err.stack });
  }
}

function doPost(e) {
  const params = JSON.parse(e.postData.contents || '{}');
  const action = params.action || '';

  const respond = (data) =>
    ContentService
      .createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);

  try {
    switch (action) {
      case 'updateInventory': return respond(updateInventory(params));
      case 'addUser':         return respond(addUser(params));
      case 'updateUser':      return respond(updateUser(params));
      default:
        return respond({ error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return respond({ error: err.message });
  }
}

// =============================================================
//  GETTERS
// =============================================================

/** Helper to parse numbers safely and handle Excel error codes. */
function safeNum(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'string') {
    if (val.trim().indexOf('#') === 0) return 0;
    const parsed = parseFloat(val.replace(/,/g, ''));
    return isNaN(parsed) ? 0 : parsed;
  }
  return isNaN(val) ? 0 : val;
}

/** Helper to format date cells safely. */
function formatDate(val) {
  if (val === null || val === undefined || val === '') return '';
  if (val instanceof Date) {
    return (val.getMonth() + 1) + '/' + val.getDate() + '/' + val.getFullYear();
  }
  const str = String(val).trim();
  if (str.indexOf('#') === 0) return '';
  return str;
}

// =============================================================
//  COLUMN AUTO-DETECTION
// =============================================================

/**
 * Scans the first 6 rows of the Inventory Utilization Report sheet to find
 * the 0-based column index for each section's "INVENTORY Volume (qty)" column.
 * Works by locating the merged section header cells:
 *   "DISPENSING AREA", "PHARMACY STORAGE", "WAREHOUSE", "CONSIGNMENT"
 *
 * In Google Sheets, a merged cell's value appears only in the top-left cell
 * of the merge, so the column index of the header = the column index of the
 * first (inventory volume) column of that section.
 */
function detectInventoryColumns(sheet) {
  const lastCol   = sheet.getLastColumn();
  const scanRows  = Math.min(6, sheet.getLastRow());
  const hData     = sheet.getRange(1, 1, scanRows, lastCol).getValues();

  const cols = { 
    dispensing_qty: -1, storage_qty: -1, warehouse_qty: -1, consignment_qty: -1,
    overall_start: -1,
    overall_total_qty: -1,
    overall_value: -1,
    overall_avg_monthly: -1,
    overall_normalized: -1,
    overall_level_days: -1,
    overall_pending_po: -1,
    overall_ending_qty: -1,
    overall_ending_days: -1,
    overall_epa_balance: -1,
    overall_ending_epa: -1,
    overall_impact_date: -1
  };

  // 1. Detect main sections & find average monthly consumption column to get overall_start (Rows 4 & 5, indices 3 & 4)
  for (let r = 3; r <= 4; r++) {
    if (r >= hData.length) continue;
    const row = hData[r];
    for (let c = 0; c < row.length; c++) {
      const cell = String(row[c] || '').trim().toUpperCase();
      if (cell === '') continue;

      if (cols.dispensing_qty === -1) {
        if (cell === 'DISPENSING AREA' || (cell.includes('DISPENSING INVENTORY') && cell.includes('(QTY)')) || (cell.includes('DISPENSING') && cell.includes('VOLUME'))) {
          cols.dispensing_qty = c;
        }
      }
      if (cols.storage_qty === -1) {
        if (cell === 'PHARMACY STORAGE' || (cell.includes('STORAGE INVENTORY') && cell.includes('(QTY)')) || (cell.includes('STORAGE') && cell.includes('VOLUME'))) {
          cols.storage_qty = c;
        }
      }
      if (cols.warehouse_qty === -1) {
        if (cell === 'WAREHOUSE' || (cell.includes('WAREHOUSE INVENTORY') && cell.includes('(QTY)')) || (cell.includes('WAREHOUSE') && cell.includes('VOLUME'))) {
          cols.warehouse_qty = c;
        }
      }
      if (cols.consignment_qty === -1) {
        if (cell === 'CONSIGNMENT' || (cell.includes('CONSIGNMENT INVENTORY') && cell.includes('(QTY)')) || (cell.includes('CONSIGNMENT') && cell.includes('VOLUME'))) {
          cols.consignment_qty = c;
        }
      }
      if (cols.overall_start === -1) {
        if (cell.includes('AVERAGE MONTHLY CONSUMPTION') && !cell.includes('DISPENSING') && !cell.includes('STORAGE') && !cell.includes('WAREHOUSE') && !cell.includes('CONSIGNMENT')) {
          cols.overall_start = c;
        }
      }
    }
  }

  // Apply layout-specific defaults based on overall_start
  if (cols.overall_start !== -1) {
    const start = cols.overall_start;
    if (start === 59) {
      // Sheet 3 (Excel) layout
      cols.overall_avg_monthly = start;
      cols.overall_normalized  = start + 1;
      cols.overall_total_qty   = start + 2;
      cols.overall_value       = start + 3;
      cols.overall_level_days  = start + 4;
      cols.overall_impact_date = start + 5; // Col BM
      cols.overall_pending_po  = start + 6;
      cols.overall_ending_qty  = start + 7;
      cols.overall_ending_days = start + 8;
      cols.overall_epa_balance = start + 9;
      cols.overall_ending_epa  = start + 11;
    } else {
      // Live Google Sheet / Sheet 2 layout
      cols.overall_avg_monthly = start;
      cols.overall_normalized  = start + 1;
      cols.overall_total_qty   = start + 2;
      cols.overall_value       = start + 3;
      cols.overall_level_days  = start + 4;
      cols.overall_impact_date = start + 5; // Col AO
      cols.overall_pending_po  = start + 8;
      cols.overall_ending_qty  = start + 9;
      cols.overall_ending_days = start + 10;
      cols.overall_epa_balance = start + 12;
      cols.overall_ending_epa  = start + 13;
    }
  } else {
    // Ultimate fallback
    cols.overall_avg_monthly = 35;
    cols.overall_normalized  = 36;
    cols.overall_total_qty   = 37;
    cols.overall_value       = 38;
    cols.overall_level_days  = 39;
    cols.overall_impact_date = 40;
    cols.overall_pending_po  = 43;
    cols.overall_ending_qty  = 44;
    cols.overall_ending_days = 45;
    cols.overall_epa_balance = 47;
    cols.overall_ending_epa  = 48;
  }

  if (cols.dispensing_qty === -1) cols.dispensing_qty = 78;
  if (cols.storage_qty === -1) cols.storage_qty = 85;

  // Refine using scanning if OVERALL start is found
  if (cols.overall_start !== -1) {
    const limitCol = cols.dispensing_qty !== -1 ? cols.dispensing_qty : hData[0].length;
    for (let r = 3; r <= 4; r++) {
      if (r >= hData.length) continue;
      const row = hData[r];
      const endScan = Math.min(limitCol, row.length);
      for (let c = cols.overall_start; c < endScan; c++) {
        const cell = String(row[c] || '').trim().toUpperCase();
        if (!cell) continue;

        if (cell.includes('AVERAGE MONTHLY CONSUMPTION')) cols.overall_avg_monthly = c;
        else if (cell.includes('NORMALIZED DEMAND')) cols.overall_normalized = c;
        else if ((cell.includes('TOTAL INVENTORY VOLUME') || cell.includes('TOTAL INVENTORY(QTY)') || cell.includes('TOTAL INVENTORY VOLUME (QTY)')) && !cell.includes('ENDING')) cols.overall_total_qty = c;
        else if ((cell.includes('INVENTORY (VALUE)') || (cell.includes('INVENTORY') && cell.includes('VALUE'))) && !cell.includes('ENDING')) cols.overall_value = c;
        else if (cell.includes('LEVEL DAYS') && !cell.includes('ENDING')) cols.overall_level_days = c;
        else if ((cell.includes('DATE OF IMPACT') || cell.includes('IMPACT DATE') || cell.includes('DATE OF') || cell.includes('IMPACT')) && !cell.includes('ENDING')) cols.overall_impact_date = c;
        else if ((cell.includes('PENDING PO') || cell.includes('PO/CO') || cell.includes('QTY OF PENDING')) && !cell.includes('ENDING') && !cell.includes('HAND')) cols.overall_pending_po = c;
        else if (cell.includes('ENDING INVENTORY') && !cell.includes('DAYS') && !cell.includes('EPA') && !cell.includes('BALANCES')) cols.overall_ending_qty = c;
        else if (cell.includes('ENDING INVENTORY LEVEL DAYS')) cols.overall_ending_days = c;
        else if (cell.includes('EPA') && cell.includes('BALANCE')) cols.overall_epa_balance = c;
        else if (cell.includes('ENDING IMPACT DATE') || (cell.includes('ENDING') && cell.includes('IMPACT'))) cols.overall_ending_epa = c;
      }
    }
  }

  Logger.log('[PharmaDash] Inventory column map: ' + JSON.stringify(cols));
  return cols;
}

/** Helper to retrieve the items sheet, supporting name fallbacks. */
function getItemSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_ITEMS);
  if (!sheet) {
    sheet = ss.getSheetByName('New Inventory Utilization 2025');
  }
  return sheet;
}

/**
 * Debug endpoint — call ?action=getSheetInfo to inspect detected columns.
 * Returns header cells, detected column indices, and sample data values.
 */
function getSheetInfo() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getItemSheet(ss);
  if (!sheet) return { error: 'Sheet not found: ' + SHEET_ITEMS + ' or "New Inventory Utilization 2025"' };

  function colLetter(idx) {
    let s = '', n = idx + 1;
    while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }

  const lastCol   = sheet.getLastColumn();
  const scanRows  = Math.min(6, sheet.getLastRow());
  const hData     = sheet.getRange(1, 1, scanRows, lastCol).getValues();
  const detected  = detectInventoryColumns(sheet);

  // Collect non-empty, non-numeric header cells
  const headerCells = [];
  for (let r = 0; r < hData.length; r++) {
    for (let c = 0; c < hData[r].length; c++) {
      const v = String(hData[r][c] || '').trim();
      if (v && isNaN(v)) {
        headerCells.push({ row: r + 1, col: colLetter(c), idx: c, value: v.substring(0, 100) });
      }
    }
  }

  // Sample values from first 3 data rows using detected columns
  const dataStart = 5; // row 5 is first potential data row
  const sampleData = [];
  if (sheet.getLastRow() >= dataStart) {
    const rows = sheet.getRange(dataStart, 1, Math.min(5, sheet.getLastRow() - dataStart + 1), lastCol).getValues();
    for (const row of rows) {
      const code = String(row[1] || '').trim();
      if (code && code.toLowerCase() !== 'item code') {
        sampleData.push({
          item_code     : code,
          dispensing_col: colLetter(detected.dispensing_qty),
          dispensing_val: row[detected.dispensing_qty],
          storage_col   : colLetter(detected.storage_qty),
          storage_val   : row[detected.storage_qty],
          warehouse_col : colLetter(detected.warehouse_qty),
          warehouse_val : row[detected.warehouse_qty],
          consignment_col: colLetter(detected.consignment_qty),
          consignment_val: row[detected.consignment_qty],
        });
        if (sampleData.length >= 3) break;
      }
    }
  }

  return {
    sheetName     : SHEET_ITEMS,
    lastRow       : sheet.getLastRow(),
    lastCol       : lastCol,
    detectedCols  : detected,
    headerCells   : headerCells,
    sampleData    : sampleData,
  };
}

/** Returns all items from the Inventory Utilization Report sheet. */
function getItems() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getItemSheet(ss);
  if (!sheet) throw new Error('Sheet "' + SHEET_ITEMS + '" or "New Inventory Utilization 2025" not found.');

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 5) return [];

  // ---- AUTO-DETECT location columns from sheet headers ----
  const lc = detectInventoryColumns(sheet);

  const values = sheet.getRange(5, 1, lastRow - 4, lastCol).getValues();

  return values.map((row, idx) => {
    const itemCode = String(row[1] || '').trim();

    // Location inventory
    const dispensingStock  = lc.dispensing_qty  >= 0 ? safeNum(row[lc.dispensing_qty])  : 0;
    const storageStock     = lc.storage_qty     >= 0 ? safeNum(row[lc.storage_qty])     : 0;
    const warehouseStock   = lc.warehouse_qty   >= 0 ? safeNum(row[lc.warehouse_qty])   : 0;
    const consignmentStock = lc.consignment_qty >= 0 ? safeNum(row[lc.consignment_qty]) : 0;

    // OVERALL stats
    const totalStock = lc.overall_total_qty >= 0 ? safeNum(row[lc.overall_total_qty]) : 0;

    return {
      no                           : String(row[0] || (idx + 1)),
      item_code                    : itemCode,
      status                       : String(row[2] || '').trim(),
      description                  : String(row[3] || '').trim(),
      generic_name                 : String(row[4] || '').trim(),
      pharmacy_category            : String(row[5] || '').trim(),
      pharmacologic_category       : String(row[6] || '').trim(),
      unit_of_measure              : String(row[7] || '').trim(),
      unit_cost                    : safeNum(row[8]),
      
      avg_monthly_consumption      : safeNum(row[lc.overall_avg_monthly]),
      avg_monthly_normalized_demand: safeNum(row[lc.overall_normalized]),
      total_inventory_qty          : totalStock,
      inventory_value_php          : safeNum(row[lc.overall_value]),
      inventory_level_days         : safeNum(row[lc.overall_level_days]),
      date_of_impact               : lc.overall_impact_date >= 0 ? formatDate(row[lc.overall_impact_date]) : '',
      pending_po_co_qty            : safeNum(row[lc.overall_pending_po]),
      ending_inventory_qty         : safeNum(row[lc.overall_ending_qty]),
      ending_inventory_level_days  : safeNum(row[lc.overall_ending_days]),
      epa_balance                  : safeNum(row[lc.overall_epa_balance]),
      ending_with_epa_qty          : safeNum(row[lc.overall_ending_epa]),

      dispensing_inventory_qty     : dispensingStock,
      storage_inventory_qty        : storageStock,
      warehouse_inventory_qty      : warehouseStock,
      consignment_inventory_qty    : consignmentStock,
      rank                         : String(row[0] || '')
    };
  }).filter(item =>
    item.item_code !== '' &&
    item.item_code.toLowerCase() !== 'item code' &&
    !item.item_code.toLowerCase().startsWith('note')
  );
}

/** Returns all reorder records from the Reorders sheet. */
function getReorders() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_REORDERS);
  if (!sheet) {
    try {
      recalculateReorders();
      sheet = ss.getSheetByName(SHEET_REORDERS);
    } catch (e) {
      throw new Error('Sheet "' + SHEET_REORDERS + '" not found and auto-calculation failed: ' + e.message);
    }
  }
  if (!sheet) return [];
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  return sheetToJSON(sheet);
}

/** Looks up a user by email from the Users sheet. Returns role info. */
function getUser(email) {
  if (!email) return { error: 'Email required.' };
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) return { error: 'Users sheet not found.' };

  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim().toLowerCase());
  const emailIdx  = headers.indexOf('email');
  const roleIdx   = headers.indexOf('role');
  const nameIdx   = headers.indexOf('name');
  const activeIdx = headers.indexOf('active');

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[emailIdx]).trim().toLowerCase() === email.toLowerCase()) {
      if (activeIdx >= 0 && !row[activeIdx]) return { error: 'Account inactive.' };
      return {
        email:  String(row[emailIdx]),
        role:   String(row[roleIdx]),
        name:   nameIdx >= 0 ? String(row[nameIdx]) : email,
        active: activeIdx >= 0 ? Boolean(row[activeIdx]) : true,
      };
    }
  }
  return { error: 'User not found.' };
}

/** Returns aggregate statistics for the dashboard. */
function getStats() {
  const items    = getItems();
  const reorders = getReorders();

  const total    = items.length;
  
  // Gated to 'Medicine Regular' for stock status alerts/KPIs
  const regularItems = items.filter(it => it.pharmacy_category === 'Medicine Regular');
  
  const critical  = regularItems.filter(it => stockDays(it) < 30).length;
  const low       = regularItems.filter(it => { const d = stockDays(it); return d >= 30 && d < 60; }).length;
  const normal    = regularItems.filter(it => { const d = stockDays(it); return d >= 60 && d < 120; }).length;
  const overstock = regularItems.filter(it => stockDays(it) >= 120).length;
  
  const pending   = items.filter(it => parseFloat(it.pending_po_co_qty || 0) > 0).length;

  // Category breakdown
  const catMap = {};
  items.forEach(it => {
    const cat = it.pharmacologic_category || 'Other';
    catMap[cat] = (catMap[cat] || 0) + 1;
  });

  return {
    total, critical, low, normal, overstock, pending,
    categories: catMap,
    reorderCount: reorders.length,
    lastUpdated: new Date().toISOString(),
  };
}

// =============================================================
//  UPDATERS
// =============================================================

/** Updates inventory quantity for an item. */
function updateInventory(params) {
  const { item_code, field, value, updated_by } = params;
  if (!item_code || !field || value === undefined) return { error: 'Missing parameters.' };

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getItemSheet(ss);
  if (!sheet) return { error: 'Inventory sheet not found.' };
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  const headers = sheet.getRange(4, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const colIdx  = headers.indexOf(field);

  if (colIdx < 0)  return { error: `Column "${field}" not found.` };
  
  const data = sheet.getRange(5, 1, lastRow - 4, lastCol).getValues();

  for (let i = 0; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(item_code).trim()) {
      const oldVal = data[i][colIdx];
      sheet.getRange(i + 5, colIdx + 1).setValue(value);
      logAction(ss, { action: 'updateInventory', item_code, field, oldVal, newVal: value, updated_by });
      return { success: true, item_code, field, newValue: value };
    }
  }
  return { error: 'Item not found: ' + item_code };
}


/** Adds a new user to the Users sheet. */
function addUser(params) {
  const { email, role, name, added_by } = params;
  if (!email || !role) return { error: 'Email and role are required.' };
  const validRoles = ['Admin', 'Dispensing', 'Storage'];
  if (!validRoles.includes(role)) return { error: `Invalid role. Must be one of: ${validRoles.join(', ')}` };

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_USERS) || ss.insertSheet(SHEET_USERS);

  // Ensure headers exist
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['email', 'role', 'name', 'active', 'created_at']);
  }

  // Check if user already exists
  const existing = getUser(email);
  if (!existing.error) return { error: 'User already exists.' };

  sheet.appendRow([email, role, name || email, true, new Date().toISOString()]);
  logAction(ss, { action: 'addUser', email, role, added_by });
  return { success: true, email, role, name: name || email };
}

/** Updates a user's role or active status. */
function updateUser(params) {
  const { email, role, active, updated_by } = params;
  if (!email) return { error: 'Email required.' };

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) return { error: 'Users sheet not found.' };

  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim().toLowerCase());
  const emailIdx  = headers.indexOf('email');
  const roleIdx   = headers.indexOf('role');
  const activeIdx = headers.indexOf('active');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][emailIdx]).toLowerCase() === email.toLowerCase()) {
      if (role   !== undefined && roleIdx   >= 0) sheet.getRange(i+1, roleIdx+1).setValue(role);
      if (active !== undefined && activeIdx >= 0) sheet.getRange(i+1, activeIdx+1).setValue(active);
      logAction(ss, { action: 'updateUser', email, role, active, updated_by });
      return { success: true };
    }
  }
  return { error: 'User not found.' };
}

// =============================================================
//  DATA IMPORT — Run this ONCE to seed data from JSON
//  Paste items.json and reorders.json content into the variables
//  below, then run importAllData() from the Apps Script editor.
// =============================================================

function importAllData() {
  // INSTRUCTIONS:
  // 1. Open Apps Script editor (Extensions > Apps Script)
  // 2. Paste the contents of items.json into the ITEMS_JSON variable below
  // 3. Paste the contents of reorders.json into the REORDERS_JSON variable below
  // 4. Click Run > importAllData
  // 5. Grant permissions when prompted
  // 6. Check your Google Sheet — data will be populated!

  const ITEMS_JSON    = '[]'; // <-- PASTE items.json content here
  const REORDERS_JSON = '[]'; // <-- PASTE reorders.json content here

  const items    = JSON.parse(ITEMS_JSON);
  const reorders = JSON.parse(REORDERS_JSON);

  if (items.length === 0)    { Logger.log('⚠️ No items to import. Did you paste items.json?'); }
  if (reorders.length === 0) { Logger.log('⚠️ No reorders to import. Did you paste reorders.json?'); }

  importSheet(SHEET_ITEMS,    items);
  importSheet(SHEET_REORDERS, reorders);
  importUsersSheet();

  Logger.log('✅ Import complete!');
  Logger.log(`   Items: ${items.length}`);
  Logger.log(`   Reorders: ${reorders.length}`);
}

function importSheet(sheetName, records) {
  if (!records || records.length === 0) return;
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  else        sheet.clearContents();

  const headers = Object.keys(records[0]);
  const rows    = records.map(r => headers.map(h => r[h] !== null && r[h] !== undefined ? r[h] : ''));

  // Write headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Style headers
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#1a2235');
  headerRange.setFontColor('#06b6d4');
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(10);

  // Write data in chunks (avoids timeout for large datasets)
  const CHUNK = 500;
  for (let i = 0; i < rows.length; i += CHUNK) {
    const chunk = rows.slice(i, i + CHUNK);
    sheet.getRange(i + 2, 1, chunk.length, headers.length).setValues(chunk);
    SpreadsheetApp.flush();
    Utilities.sleep(200); // small pause to avoid rate limits
    Logger.log(`  Wrote rows ${i+1}–${Math.min(i+CHUNK, rows.length)} of ${rows.length} to "${sheetName}"`);
  }

  // Freeze header row and auto-resize
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  Logger.log(`✅ Imported ${rows.length} records to "${sheetName}"`);
}

function importUsersSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(SHEET_USERS);
  if (!sheet) sheet = ss.insertSheet(SHEET_USERS);
  else        sheet.clearContents();

  const headers = ['email', 'role', 'name', 'active', 'created_at'];
  const rows = [
    ['admin@pharma.gov.ph',      'Admin',      'Admin User',       true,  new Date().toISOString()],
    ['dispensing@pharma.gov.ph', 'Dispensing', 'Dispensing Staff', true,  new Date().toISOString()],
    ['storage@pharma.gov.ph',    'Storage',    'Storage Staff',    true,  new Date().toISOString()],
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sheet.setFrozenRows(1);

  // Style it
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#1a2235');
  headerRange.setFontColor('#06b6d4');
  headerRange.setFontWeight('bold');

  Logger.log(`✅ Users sheet created with ${rows.length} demo users.`);
}

// =============================================================
//  UTILITIES
// =============================================================

/** Converts a sheet's data range into an array of plain objects. */
function sheetToJSON(sheet) {
  const data    = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(h => String(h).trim());
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i] !== '' ? row[i] : null; });
    return obj;
  });
}

/** Calculates inventory level days from an item object. */
function stockDays(item) {
  return parseFloat(item.inventory_level_days || item['inventory_level_days'] || 0);
}

/** Writes an entry to the AuditLog sheet. */
function logAction(ss, details) {
  try {
    let log = ss.getSheetByName(SHEET_LOG);
    if (!log) {
      log = ss.insertSheet(SHEET_LOG);
      log.appendRow(['timestamp', 'action', 'details']);
    }
    log.appendRow([new Date().toISOString(), details.action, JSON.stringify(details)]);
  } catch (e) {
    Logger.log('Audit log error: ' + e.message);
  }
}

// =============================================================
//  TRIGGER SETUP — Run once to enable scheduled recalculation
// =============================================================

/**
 * Run this once from the Apps Script editor to set up a daily
 * trigger that recalculates reorder quantities automatically.
 */
function setupTriggers() {
  // Delete existing triggers first
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  // Add daily recalculation at 6:00 AM
  ScriptApp.newTrigger('recalculateReorders')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();
  Logger.log('✅ Daily recalculation trigger set for 6:00 AM.');
}

/**
 * Recalculates reorder quantities for all items and updates the Reorders sheet.
 * Formula: ROP = (Avg Monthly Consumption × 3 months) + (6-month safety stock) − stock on hand
 */
function recalculateReorders() {
  const items = getItems();
  const reorders = items.map(it => {
    const avg   = parseFloat(it.avg_monthly_normalized_demand || it.avg_monthly_consumption || 0);
    const stock = parseFloat(it.total_inventory_qty || 0);
    const rop6  = Math.max(0, Math.round(avg * 3 + avg * 6 - stock));
    const rop1  = Math.max(0, Math.round(avg * 3 + avg * 1 - stock));
    const days  = avg > 0 ? Math.round((stock / avg) * 30) : 0;
    const status= days < 30 ? 'Critical' : days < 60 ? 'Low' : days < 120 ? 'Normal' : 'Over Stock';

    return {
      item_code:                     it.item_code || '',
      description:                   it.description || '',
      pharmacologic_category:        it.pharmacologic_category || '',
      pharmacy_category:             it.pharmacy_category || '',
      unit_of_measure:               it.unit_of_measure || '',
      avg_monthly_consumption:       it.avg_monthly_consumption || 0,
      avg_monthly_normalized_demand: avg,
      total_inventory_qty:           stock,
      inventory_level_days:          days,
      stock_status:                  status,
      reorder_qty_6mo_safety:        rop6,
      reorder_qty_1mo_safety:        rop1,
      pending_po_call_off:           it.pending_po_co_qty || 0,
      remarks:                       rop6 > 0 ? 'Request in full' : '',
      calculated_at:                 new Date().toISOString(),
    };
  }).filter(r => r.reorder_qty_6mo_safety > 0)
    .sort((a,b) => a.inventory_level_days - b.inventory_level_days);

  importSheet(SHEET_REORDERS, reorders);
  Logger.log(`✅ Recalculated reorders: ${reorders.length} items need replenishment.`);
}
