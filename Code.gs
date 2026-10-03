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

  // If no action parameter, serve the dashboard page (supports 'index' or 'dashboard' HTML file names)
  if (!action) {
    try {
      return HtmlService.createHtmlOutputFromFile('index')
        .setTitle('PharmaDash — Pharmacy Inventory Dashboard')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    } catch (e) {
      return HtmlService.createHtmlOutputFromFile('dashboard')
        .setTitle('PharmaDash — Pharmacy Inventory Dashboard')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  }

  // CORS-friendly JSON response helper for API calls
  const respond = (data) =>
    ContentService
      .createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);

  const email  = params.email  || '';
  const token  = params.token  || '';

  try {
    switch (action) {
      case 'getItems':     return respond(getItems());
      case 'getReorders':  return respond(getReorders());
      case 'getLocator':   return respond(getLocatorData());
      case 'getUser':      return respond(getUser(email));
      case 'getStats':     return respond(getStats());
      case 'getSheetInfo': return respond(getSheetInfo());
      case 'getMetadata':  return respond(getSheetMetadata());
      case 'ping':         return respond({ status: 'ok', timestamp: new Date().toISOString() });
      default:
        return respond({ error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return respond({ error: err.message, stack: err.stack });
  }
}

function doPost(e) {
  const respond = (data) =>
    ContentService
      .createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);

  try {
    const rawContent = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    const params = JSON.parse(rawContent || '{}');
    const action = params.action || '';

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

/** Helper to parse a date from header cell values dynamically. */
function parseHeaderDate(val) {
  if (val === null || val === undefined || val === '') return null;
  if (val instanceof Date) return val;
  const num = Number(val);
  if (!isNaN(num) && num > 30000) {
    return new Date((num - 25569) * 86400 * 1000);
  }
  const str = String(val).trim().toLowerCase();
  if (str.indexOf('#') === 0) return null;
  
  const parts = str.split(/[- ]/);
  if (parts.length >= 2) {
    const months = {
      jan:0, feb:1, mar:2, apr:3, may:4, jun:5, jul:6, aug:7, sep:8, oct:9, nov:10, dec:11
    };
    const m = months[parts[0].substring(0,3)];
    let y = parseInt(parts[1]);
    if (m !== undefined && !isNaN(y)) {
      if (y < 100) y += 2000;
      return new Date(y, m, 1);
    }
  }
  
  const standaloneYear = parseInt(str);
  if (!isNaN(standaloneYear) && standaloneYear >= 2020 && standaloneYear <= 2030) {
    return new Date(standaloneYear, 0, 1);
  }
  return null;
}

// =============================================================
//  COLUMN AUTO-DETECTION
// =============================================================

/**
 * Dynamically scans the header rows (Rows 4 & 5) of the Inventory Utilization Report sheet
 * to identify the exact 0-based column indices for all sections and inventory metrics.
 * 
 * Works 100% dynamically via header text matching, making the app immune to column shifts
 * when monthly consumption columns are added, removed, or rolled over.
 */
function detectInventoryColumns(sheet) {
  const lastCol   = sheet.getLastColumn();
  const scanRows  = Math.min(6, sheet.getLastRow());
  const hData     = sheet.getRange(1, 1, scanRows, lastCol).getValues();
  const hRow4     = hData.length >= 4 ? hData[3] : [];
  const hRow5     = hData.length >= 5 ? hData[4] : [];
  const maxCols   = Math.max(hRow4.length, hRow5.length);

  const cols = { 
    dispensing_qty: -1, storage_qty: -1, warehouse_qty: -1, consignment_qty: -1,
    overall_start: -1,
    overall_total_qty: -1,
    overall_value: -1,
    overall_avg_monthly: -1,
    overall_normalized: -1,
    overall_level_days: -1,
    overall_impact_date: -1,
    overall_pending_po: -1,
    overall_ending_qty: -1,
    overall_ending_days: -1,
    overall_ending_impact_date: -1,
    overall_epa_balance: -1,
    overall_ending_epa: -1,
    overall_ending_epa_days: -1,
    overall_ending_epa_impact_date: -1
  };

  // PASS 1: Identify Section Markers in Row 4 & Row 5
  for (let c = 0; c < maxCols; c++) {
    const cell4 = String(hRow4[c] || '').trim().toUpperCase();
    const cell5 = String(hRow5[c] || '').trim().toUpperCase();
    const combined = (cell4 + ' ' + cell5).trim();

    // Section 1: Dispensing Area
    if (cols.dispensing_qty === -1) {
      if (cell4 === 'DISPENSING AREA' || (combined.includes('DISPENSING INVENTORY') && combined.includes('(QTY)')) || (combined.includes('DISPENSING') && combined.includes('VOLUME'))) {
        cols.dispensing_qty = c;
      }
    }
    // Section 2: Pharmacy Storage
    if (cols.storage_qty === -1) {
      if (cell4 === 'PHARMACY STORAGE' || (combined.includes('STORAGE INVENTORY') && combined.includes('(QTY)')) || (combined.includes('STORAGE') && combined.includes('VOLUME'))) {
        cols.storage_qty = c;
      }
    }
    // Section 3: Warehouse
    if (cols.warehouse_qty === -1) {
      if (cell4 === 'WAREHOUSE' || (combined.includes('WAREHOUSE INVENTORY') && combined.includes('(QTY)')) || (combined.includes('WAREHOUSE') && combined.includes('VOLUME'))) {
        cols.warehouse_qty = c;
      }
    }
    // Section 4: Consignment
    if (cols.consignment_qty === -1) {
      if (cell4 === 'CONSIGNMENT' || (combined.includes('CONSIGNMENT INVENTORY') && combined.includes('(QTY)')) || (combined.includes('CONSIGNMENT') && combined.includes('VOLUME'))) {
        cols.consignment_qty = c;
      }
    }

    // Section 0: Overall Start (Inventory Impact)
    if (cols.overall_start === -1) {
      if (cell4.includes('INVENTORY IMPACT') || (cell5.includes('AVERAGE MONTHLY CONSUMPTION') && (cols.dispensing_qty === -1 || c < cols.dispensing_qty))) {
        cols.overall_start = c;
      }
    }
  }

  // PASS 2: Scan columns inside the Overall / Impact / Pending / EPA section
  const scanStart = cols.overall_start !== -1 ? cols.overall_start : 10;
  const scanEnd   = cols.dispensing_qty !== -1 ? cols.dispensing_qty : maxCols;

  for (let c = scanStart; c < scanEnd; c++) {
    const cell4 = String(hRow4[c] || '').trim().toUpperCase();
    const cell5 = String(hRow5[c] || '').trim().toUpperCase();

    // 1. Average Monthly Consumption
    if (cols.overall_avg_monthly === -1 && cell5.includes('AVERAGE MONTHLY CONSUMPTION') && !cell5.includes('NORMALIZED')) {
      cols.overall_avg_monthly = c;
    }
    // 2. Average Monthly Normalized Demand
    else if (cols.overall_normalized === -1 && (cell5.includes('NORMALIZED DEMAND') || cell5.includes('AVERAGE MONTHLY NORMALIZED'))) {
      cols.overall_normalized = c;
    }
    // 3. Total Inventory Volume Qty
    else if (cols.overall_total_qty === -1 && (cell5.includes('TOTAL INVENTORY VOLUME') || cell5.includes('TOTAL INVENTORY(QTY)') || cell5.includes('TOTAL INVENTORY VOLUME (QTY)')) && !cell5.includes('ENDING')) {
      cols.overall_total_qty = c;
    }
    // 4. Inventory Value
    else if (cols.overall_value === -1 && (cell5.includes('INVENTORY') && cell5.includes('VALUE')) && !cell5.includes('ENDING') && !cell5.includes('HOLDING') && !cell4.includes('HOLDING')) {
      cols.overall_value = c;
    }
    // 5. Inventory Level Days
    else if (cols.overall_level_days === -1 && cell5.includes('INVENTORY LEVEL DAYS') && !cell5.includes('ENDING')) {
      cols.overall_level_days = c;
    }
    // 6. Date of Impact
    else if (cols.overall_impact_date === -1 && (cell5.includes('DATE OF IMPACT') || cell5.includes('IMPACT DATE')) && !cell5.includes('ENDING')) {
      cols.overall_impact_date = c;
    }
    // 7. Pending PO / CO Qty
    else if (cols.overall_pending_po === -1 && (cell5.includes('PENDING PO') || cell5.includes('PO / CO') || cell5.includes('PO/CO') || cell5.includes('QTY OF PENDING'))) {
      cols.overall_pending_po = c;
    }
    // 8. Ending Inventory Qty (Quantity on Hand + Qty of pending CO/PO)
    else if (cols.overall_ending_qty === -1 && cell5.includes('ENDING INVENTORY') && !cell5.includes('DAYS') && !cell5.includes('EPA') && !cell5.includes('IMPACT')) {
      cols.overall_ending_qty = c;
    }
    // 9. Ending Inventory Level Days (Without EPA)
    else if (cols.overall_ending_days === -1 && cell5.includes('ENDING INVENTORY LEVEL DAYS') && !cell5.includes('EPA')) {
      cols.overall_ending_days = c;
    }
    // 10. Ending Impact Date (Without EPA)
    else if (cols.overall_ending_impact_date === -1 && cell5.includes('ENDING IMPACT DATE') && !cell5.includes('EPA')) {
      cols.overall_ending_impact_date = c;
    }
    // 11. EPA Balance for Call-Off
    else if (cols.overall_epa_balance === -1 && cell5.includes('EPA') && (cell5.includes('CALL-OFF') || cell5.includes('BALANCE')) && !cell5.includes('ENDING INVENTORY') && !cell5.includes('DAYS') && !cell5.includes('IMPACT')) {
      cols.overall_epa_balance = c;
    }
    // 12. Ending Inventory with EPA
    else if (cols.overall_ending_epa === -1 && cell5.includes('ENDING INVENTORY') && cell5.includes('EPA') && !cell5.includes('DAYS') && !cell5.includes('IMPACT')) {
      cols.overall_ending_epa = c;
    }
    // 13. Ending Inventory Level Days with EPA
    else if (cols.overall_ending_epa_days === -1 && cell5.includes('ENDING INVENTORY LEVEL DAYS') && cell5.includes('EPA')) {
      cols.overall_ending_epa_days = c;
    }
    // 14. Ending Impact Date with EPA
    else if (cols.overall_ending_epa_impact_date === -1 && cell5.includes('ENDING IMPACT DATE') && cell5.includes('EPA')) {
      cols.overall_ending_epa_impact_date = c;
    }
  }

  // PASS 3: Safe relative fallback offsets anchored to overall_start
  // If any individual sub-header was missing or blank, use relative offsets from overall_start
  if (cols.overall_start !== -1) {
    const s = cols.overall_start;
    if (cols.overall_avg_monthly === -1) cols.overall_avg_monthly = s;
    if (cols.overall_normalized  === -1) cols.overall_normalized  = s + 1;
    if (cols.overall_total_qty   === -1) cols.overall_total_qty   = s + 2;
    if (cols.overall_value       === -1) cols.overall_value       = s + 3;
    if (cols.overall_level_days  === -1) cols.overall_level_days  = s + 4;
    if (cols.overall_impact_date === -1) cols.overall_impact_date = s + 5;
    if (cols.overall_pending_po  === -1) cols.overall_pending_po  = s + 7;
    if (cols.overall_ending_qty  === -1) cols.overall_ending_qty  = s + 8;
    if (cols.overall_ending_days === -1) cols.overall_ending_days = s + 9;
    if (cols.overall_ending_impact_date === -1) cols.overall_ending_impact_date = s + 10;
    if (cols.overall_epa_balance === -1) cols.overall_epa_balance = s + 11;
    if (cols.overall_ending_epa  === -1) cols.overall_ending_epa  = s + 12;
  }

  // Fallbacks for breakdown sections if headers are unmerged/unnamed
  if (cols.dispensing_qty  === -1) cols.dispensing_qty  = (cols.overall_start !== -1 ? cols.overall_start + 16 : 52);
  if (cols.storage_qty     === -1) cols.storage_qty     = cols.dispensing_qty + 7;
  if (cols.warehouse_qty   === -1) cols.warehouse_qty   = cols.storage_qty + 7;
  if (cols.consignment_qty === -1) cols.consignment_qty = cols.warehouse_qty + 7;

  Logger.log('[PharmaDash] Dynamic inventory column map: ' + JSON.stringify(cols));
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
 * Returns metadata about the Inventory Utilization Report sheet:
 * - Manual update date from Cell B2 (e.g. '10/01/2026')
 * - Automatic Google Drive file last modified timestamp
 * - Sheet URL and tab GID
 * - Calculated days elapsed and staleness indicator (default: >= 3 days)
 */
function getSheetMetadata() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getItemSheet(ss);
  
  let manualDateRaw = '';
  let manualDateIso = null;
  let manualDateFormatted = '';
  
  if (sheet) {
    try {
      // Cell B2 (Row 2, Column 2)
      const b2 = sheet.getRange(2, 2).getValue();
      if (b2 instanceof Date) {
        manualDateIso = b2.toISOString();
        manualDateFormatted = (b2.getMonth() + 1) + '/' + b2.getDate() + '/' + b2.getFullYear();
        manualDateRaw = manualDateFormatted;
      } else if (b2 !== null && b2 !== undefined && String(b2).trim() !== '') {
        manualDateRaw = String(b2).trim();
        const d = new Date(manualDateRaw);
        if (!isNaN(d.getTime())) {
          manualDateIso = d.toISOString();
          manualDateFormatted = (d.getMonth() + 1) + '/' + d.getDate() + '/' + d.getFullYear();
        } else {
          manualDateFormatted = manualDateRaw;
        }
      }
    } catch (e) {
      Logger.log('Error reading Cell B2: ' + e.message);
    }
  }

  let autoLastModified = null;
  try {
    const file = DriveApp.getFileById(ss.getId());
    if (file) {
      autoLastModified = file.getLastUpdated().toISOString();
    }
  } catch (e) {
    Logger.log('DriveApp getLastUpdated error: ' + e.message);
  }

  let effectiveDate = null;
  let effectiveType = 'none';

  const mTime = manualDateIso ? new Date(manualDateIso).getTime() : 0;
  const aTime = autoLastModified ? new Date(autoLastModified).getTime() : 0;

  if (mTime > 0 && mTime >= aTime) {
    effectiveDate = manualDateIso;
    effectiveType = 'manual';
  } else if (aTime > 0) {
    effectiveDate = autoLastModified;
    effectiveType = 'automatic';
  } else if (mTime > 0) {
    effectiveDate = manualDateIso;
    effectiveType = 'manual';
  }

  let daysSinceUpdate = 0;
  if (effectiveDate) {
    const diffMs = Date.now() - new Date(effectiveDate).getTime();
    daysSinceUpdate = Math.max(0, Math.floor(diffMs / (1000 * 60 * 60 * 24)));
  }

  let sheetUrl = ss.getUrl();
  if (sheet) {
    sheetUrl += '#gid=' + sheet.getSheetId();
  }

  return {
    manualDateRaw       : manualDateRaw,
    manualDateIso       : manualDateIso,
    manualDateFormatted : manualDateFormatted,
    autoLastModified    : autoLastModified,
    effectiveDate       : effectiveDate,
    effectiveType       : effectiveType,
    daysSinceUpdate     : daysSinceUpdate,
    isStale             : daysSinceUpdate >= 3,
    staleThresholdDays  : 3,
    sheetUrl            : sheetUrl,
    sheetName           : sheet ? sheet.getName() : (SHEET_ITEMS || 'Inventory Utilization Report'),
    serverTime          : new Date().toISOString()
  };
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

  // Group all consumption columns dynamically by year
  const hRow4 = sheet.getRange(4, 1, 1, lastCol).getValues()[0];
  const hRow5 = sheet.getRange(5, 1, 1, lastCol).getValues()[0];
  const colsByYear = {};
  
  for (let c = 10; c < lc.overall_start; c++) {
    const val4 = hRow4[c];
    const val5 = hRow5[c];
    const date = parseHeaderDate(val4) || parseHeaderDate(val5);
    if (date) {
      const yr = date.getFullYear();
      if (!colsByYear[yr]) colsByYear[yr] = [];
      colsByYear[yr].push(c);
    }
  }

  const values = sheet.getRange(5, 1, lastRow - 4, lastCol).getValues();

  return values.map((row, idx) => {
    const itemCode = String(row[1] || '').trim();

    // Calculate annualized metrics dynamically for all detected years (2024, 2025, 2026, 2027, etc.)
    const annualMetrics = {};
    for (const yr in colsByYear) {
      const colList = colsByYear[yr];
      let yrSum = 0;
      for (let i = 0; i < colList.length; i++) {
        yrSum += safeNum(row[colList[i]]);
      }
      const count = colList.length;
      const annualQty = count === 12 ? yrSum : (count > 0 ? (yrSum / count) * 12 : 0);
      const annualVal = annualQty * safeNum(row[8]);
      annualMetrics['annual_qty_' + yr] = annualQty;
      annualMetrics['annual_val_' + yr] = annualVal;
    }

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
      dci_code                     : String(row[2] || '').trim(),
      status                       : String(row[2] || '').trim(),
      description                  : String(row[3] || '').trim(),
      generic_name                 : String(row[4] || '').trim(),
      pharmacy_category            : String(row[5] || '').trim(),
      pharmacologic_category       : String(row[6] || '').trim(),
      unit_of_measure              : String(row[7] || '').trim(),
      unit_cost                    : safeNum(row[8]),
      acquisition_price            : safeNum(row[8]),
      selling_price                : safeNum(row[9]),
      avg_monthly_consumption      : safeNum(row[lc.overall_avg_monthly]),
      avg_monthly_normalized_demand: safeNum(row[lc.overall_normalized]),
      total_inventory_qty          : totalStock,
      inventory_value_php          : safeNum(row[lc.overall_value]),
      inventory_level_days         : safeNum(row[lc.overall_level_days]),
      date_of_impact               : lc.overall_impact_date >= 0 ? formatDate(row[lc.overall_impact_date]) : '',
      pending_po_co_qty            : safeNum(row[lc.overall_pending_po]),
      ending_inventory_qty         : safeNum(row[lc.overall_ending_qty]),
      ending_inventory_level_days  : safeNum(row[lc.overall_ending_days]),
      ending_impact_date           : lc.overall_ending_impact_date >= 0 ? formatDate(row[lc.overall_ending_impact_date]) : '',
      epa_balance                  : safeNum(row[lc.overall_epa_balance]),
      epa_cy2026_balance           : safeNum(row[lc.overall_epa_balance]),
      ending_with_epa_qty          : safeNum(row[lc.overall_ending_epa]),
      ending_inv_with_epa_qty      : safeNum(row[lc.overall_ending_epa]),

      // Preserve standard fields and spread dynamic annual metrics
      ...annualMetrics,
      annual_qty_2025              : annualMetrics['annual_qty_2025'] || 0,
      annual_val_2025              : annualMetrics['annual_val_2025'] || 0,
      annual_qty_2026              : annualMetrics['annual_qty_2026'] || 0,
      annual_val_2026              : annualMetrics['annual_val_2026'] || 0,
 
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
      Logger.log('[PharmaDash] getReorders notice: ' + e.message);
      return [];
    }
  }
  if (!sheet) return [];
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  return sheetToJSON(sheet);
}

/** Returns pallet locator records from the Locator sheet or Google Sheet CSV export. */
function getLocatorData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Locator') || ss.getSheetByName('Map Locator');
  if (sheet) {
    return sheetToJSON(sheet);
  }
  try {
    const url = 'https://docs.google.com/spreadsheets/d/1Nyzo7WCS90_t88Pz0wQ35sx4sBN-SSSVFSsLHc9d21w/export?format=csv&gid=0';
    const res = UrlFetchApp.fetch(url);
    const csv = Utilities.parseCsv(res.getContentText());
    if (csv.length < 2) return [];
    const headers = csv[0].map(h => String(h).trim().toLowerCase());
    const itemCodeIdx = headers.indexOf('item code');
    const dciIdx      = headers.indexOf('dci code');
    const descIdx     = headers.indexOf('item description');
    const locIdx      = headers.indexOf('location');

    return csv.slice(1).map(row => ({
      item_code:   itemCodeIdx >= 0 ? String(row[itemCodeIdx]).trim() : '',
      dci_code:    dciIdx      >= 0 ? String(row[dciIdx]).trim()      : '',
      description: descIdx     >= 0 ? String(row[descIdx]).trim()     : '',
      location:    locIdx      >= 0 ? String(row[locIdx]).trim()      : ''
    })).filter(r => r.item_code);
  } catch (e) {
    Logger.log('Locator fetch error: ' + e.message);
    return [];
  }
}

/** Looks up a user by email from the Users sheet. Returns role info. */
function getUser(email) {
  if (!email) return { error: 'Email required.' };
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_USERS);
    sheet.appendRow(['email', 'role', 'name', 'active']);
    sheet.appendRow(['admin@pharma.gov.ph', 'Admin', 'Admin User', true]);
    sheet.appendRow(['dispensing@pharma.gov.ph', 'Dispensing', 'Dispensing Staff', true]);
    sheet.appendRow(['storage@pharma.gov.ph', 'Storage', 'Storage Staff', true]);
  }

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
  
  // Gated to monitored categories (Medicine Regular + Consignment)
  const monitored = items.filter(it => {
    const cat = String(it.pharmacy_category || '').toLowerCase();
    return cat === 'medicine regular' || cat.includes('consig');
  });
  
  const outofstock = monitored.filter(it => parseFloat(it.total_inventory_qty || 0) <= 0).length;
  const critical   = monitored.filter(it => { const q = parseFloat(it.total_inventory_qty || 0); const d = stockDays(it); return q > 0 && d < 30; }).length;
  const low        = monitored.filter(it => { const q = parseFloat(it.total_inventory_qty || 0); const d = stockDays(it); return q > 0 && d >= 30 && d < 60; }).length;
  const normal     = monitored.filter(it => { const q = parseFloat(it.total_inventory_qty || 0); const d = stockDays(it); return q > 0 && d >= 60 && d < 120; }).length;
  const overstock  = monitored.filter(it => { const q = parseFloat(it.total_inventory_qty || 0); const d = stockDays(it); return q > 0 && d >= 120; }).length;
  
  const pending   = items.filter(it => parseFloat(it.pending_po_co_qty || 0) > 0).length;

  // Category breakdown
  const catMap = {};
  items.forEach(it => {
    const cat = it.pharmacologic_category || 'Other';
    catMap[cat] = (catMap[cat] || 0) + 1;
  });

  return {
    total, outofstock, critical, low, normal, overstock, pending,
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
    const status= stock <= 0 ? 'Out of Stock' : (days < 30 ? 'Critical' : (days < 60 ? 'Low' : (days < 120 ? 'Normal' : 'Over Stock')));

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
