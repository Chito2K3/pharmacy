const fs = require('fs');
const path = require('path');

function parseCSVLine(text) {
  const result = [];
  let cur = '';
  let inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (c === '"') {
      if (inQuotes && text[i + 1] === '"') {
        cur += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (c === ',' && !inQuotes) {
      result.push(cur.trim());
      cur = '';
    } else {
      cur += c;
    }
  }
  result.push(cur.trim());
  return result;
}

const csvPath = path.join(__dirname, '..', 'scratch_sheet.csv');
const itemsJsonPath = path.join(__dirname, '..', 'items.json');

if (!fs.existsSync(csvPath)) {
  console.error('scratch_sheet.csv not found');
  process.exit(1);
}

const lines = fs.readFileSync(csvPath, 'utf8').split('\n');
console.log('Total CSV lines:', lines.length);

let headerIdx = -1;
for (let i = 0; i < Math.min(10, lines.length); i++) {
  const parsed = parseCSVLine(lines[i]);
  if (parsed[1] && parsed[1].toLowerCase() === 'item code') {
    headerIdx = i;
    console.log('Header found at line', i, 'Col I header:', parsed[8], 'Col J header:', parsed[9]);
    break;
  }
}

if (headerIdx === -1) {
  console.error('Could not find header row in CSV');
  process.exit(1);
}

const priceMap = new Map();
for (let i = headerIdx + 1; i < lines.length; i++) {
  const row = parseCSVLine(lines[i]);
  const itemCode = row[1];
  if (!itemCode || itemCode.toLowerCase() === 'item code' || itemCode.startsWith('Note:')) continue;
  const acqCost = row[8] ? row[8].replace(/,/g, '').trim() : '0';
  const selling = row[9] ? row[9].replace(/,/g, '').trim() : '0';
  priceMap.set(itemCode, {
    acquisition_price: isNaN(parseFloat(acqCost)) ? '0' : acqCost,
    selling_price: isNaN(parseFloat(selling)) ? '0' : selling
  });
}

console.log('Items parsed from Google Sheet CSV:', priceMap.size);

const items = JSON.parse(fs.readFileSync(itemsJsonPath, 'utf8'));
let matched = 0;
let updatedAcq = 0;
let updatedSell = 0;

for (const it of items) {
  const code = it.item_code;
  if (priceMap.has(code)) {
    matched++;
    const p = priceMap.get(code);
    it.acquisition_price = p.acquisition_price;
    it.selling_price = p.selling_price;
    it.unit_cost = p.acquisition_price; // keep for backward compatibility
    if (parseFloat(p.acquisition_price) > 0) updatedAcq++;
    if (parseFloat(p.selling_price) > 0) updatedSell++;
  } else {
    it.acquisition_price = it.acquisition_price || it.unit_cost || '0';
    it.selling_price = it.selling_price || '0';
  }
}

console.log(`Matched ${matched} / ${items.length} items in items.json`);
console.log(`Items with non-zero acquisition price: ${updatedAcq}`);
console.log(`Items with non-zero selling price: ${updatedSell}`);

fs.writeFileSync(itemsJsonPath, JSON.stringify(items, null, 2), 'utf8');
console.log('Successfully updated items.json with Acquisition Price and Selling Price!');

// Clean up scratch file
if (fs.existsSync(csvPath)) {
  fs.unlinkSync(csvPath);
  console.log('Removed temporary scratch_sheet.csv');
}
