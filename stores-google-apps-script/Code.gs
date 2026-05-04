// ─── Stores Management — Google Apps Script Backend ─────────────────────────
// SETUP: Replace YOUR_SPREADSHEET_ID_HERE with your Google Sheet ID
// (the long string in the sheet's URL between /d/ and /edit)
const SPREADSHEET_ID = '1IGUN4zVLPw6iH9_NCZUlh183It8nZAH3wbbktKwrftc';

const INV_SHEET = 'Inventory';
const TXN_SHEET = 'Transactions';
const SET_SHEET = 'Settings';

const INV_HEADERS = [
  'id','partNumber','description','itemHeading','oem','qty','qtyRequired',
  'cost','reorderPoint','location','threeQuotesRequired',
  'quotation1','quotation2','quotation3',
  'contact1','contact2','contact3',
  'unipressPO','poCopy','cad','link','lastUpdated'
];
const TXN_HEADERS = ['id','partId','type','quantity','user','notes','timestamp'];
const SET_HEADERS = ['key','value'];

// ─── Utilities ───────────────────────────────────────────────────────────────
function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2,7);
}

function jsonResp(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
function ok(data)  { return jsonResp({ ok: true,  data: data }); }
function fail(msg) { return jsonResp({ ok: false, error: msg }); }

function getOrCreate(name, headers) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.appendRow(headers);
    s.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    s.setFrozenRows(1);
  }
  return s;
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1)
    .filter(r => r[0] !== '')
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = r[i]; });
      return obj;
    });
}

// ─── HTTP Handlers ────────────────────────────────────────────────────────────
function doGet(e)  { return route(e.parameter || {}); }
function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (_) {}
  if (e.parameter) Object.keys(e.parameter).forEach(k => { if (!p[k]) p[k] = e.parameter[k]; });
  return route(p);
}

function route(p) {
  try {
    switch (p.action) {
      case 'setup':           return doSetup();
      case 'getInventory':    return doGetInventory();
      case 'getPart':         return doGetPart(p);
      case 'addPart':         return doAddPart(p);
      case 'updatePart':      return doUpdatePart(p);
      case 'deletePart':      return doDeletePart(p);
      case 'adjustStock':     return doAdjustStock(p);
      case 'getLowStock':     return doGetLowStock();
      case 'getTransactions': return doGetTransactions(p);
      case 'sendPO':          return doSendPO(p);
      case 'getSettings':     return doGetSettings();
      case 'updateSettings':  return doUpdateSettings(p);
      case 'importFromSheet': return doImportFromSheet(p);
      default:                return fail('Unknown action: ' + (p.action || 'none'));
    }
  } catch (err) {
    return fail(err.toString());
  }
}

// ─── SETUP ───────────────────────────────────────────────────────────────────
function doSetup() {
  getOrCreate(INV_SHEET, INV_HEADERS);
  getOrCreate(TXN_SHEET, TXN_HEADERS);
  getOrCreate(SET_SHEET, SET_HEADERS);
  return ok({ message: 'Sheets created successfully' });
}

// ─── INVENTORY CRUD ───────────────────────────────────────────────────────────
function doGetInventory() {
  return ok(sheetToObjects(getOrCreate(INV_SHEET, INV_HEADERS)));
}

function doGetPart(p) {
  const items = sheetToObjects(getOrCreate(INV_SHEET, INV_HEADERS));
  const item = items.find(i => i.id === p.id || String(i.partNumber) === String(p.partNumber));
  if (!item) return fail('Part not found');
  return ok(item);
}

function doAddPart(p) {
  const s = getOrCreate(INV_SHEET, INV_HEADERS);
  const part = JSON.parse(p.row || '{}');
  if (!part.id) part.id = uid();
  part.lastUpdated = new Date().toISOString();
  s.appendRow(INV_HEADERS.map(h => part[h] !== undefined ? part[h] : ''));
  return ok(part);
}

function doUpdatePart(p) {
  const s = getOrCreate(INV_SHEET, INV_HEADERS);
  const data = s.getDataRange().getValues();
  const headers = data[0];
  const updates = JSON.parse(p.row || '{}');
  updates.lastUpdated = new Date().toISOString();
  const idCol = headers.indexOf('id');
  for (let r = 1; r < data.length; r++) {
    if (data[r][idCol] === updates.id) {
      headers.forEach((h, c) => {
        if (updates[h] !== undefined) s.getRange(r + 1, c + 1).setValue(updates[h]);
      });
      return ok(updates);
    }
  }
  return fail('Part not found: ' + updates.id);
}

function doDeletePart(p) {
  const s = getOrCreate(INV_SHEET, INV_HEADERS);
  const data = s.getDataRange().getValues();
  const idCol = data[0].indexOf('id');
  for (let r = 1; r < data.length; r++) {
    if (data[r][idCol] === p.id) { s.deleteRow(r + 1); return ok({ deleted: p.id }); }
  }
  return fail('Part not found: ' + p.id);
}

// ─── STOCK ADJUSTMENT ─────────────────────────────────────────────────────────
function doAdjustStock(p) {
  const qty = parseInt(p.quantity, 10);
  if (!qty || qty <= 0) return fail('Invalid quantity');

  const s = getOrCreate(INV_SHEET, INV_HEADERS);
  const data = s.getDataRange().getValues();
  const headers = data[0];
  const idCol      = headers.indexOf('id');
  const qtyCol     = headers.indexOf('qty');
  const reorderCol = headers.indexOf('reorderPoint');
  const lastCol    = headers.indexOf('lastUpdated');

  for (let r = 1; r < data.length; r++) {
    if (data[r][idCol] !== p.partId) continue;

    let current = parseInt(data[r][qtyCol], 10) || 0;
    if (p.type === 'add') {
      current += qty;
    } else {
      if (current < qty) return fail('Insufficient stock: have ' + current + ', need ' + qty);
      current -= qty;
    }

    s.getRange(r + 1, qtyCol + 1).setValue(current);
    s.getRange(r + 1, lastCol + 1).setValue(new Date().toISOString());

    const txn = getOrCreate(TXN_SHEET, TXN_HEADERS);
    txn.appendRow([uid(), p.partId, p.type, qty, p.user || '', p.notes || '', new Date().toISOString()]);

    const reorderPoint = parseInt(data[r][reorderCol], 10) || 0;
    const lowStock = p.type === 'deduct' && reorderPoint > 0 && current <= reorderPoint;

    if (lowStock) {
      try {
        const part = {};
        headers.forEach((h, i) => { part[h] = data[r][i]; });
        part.qty = current;
        tryAutoSendPO(part);
      } catch (e) { Logger.log('Auto PO error: ' + e); }
    }

    return ok({ newQty: current, lowStock: lowStock });
  }
  return fail('Part not found: ' + p.partId);
}

// ─── LOW STOCK ────────────────────────────────────────────────────────────────
function doGetLowStock() {
  const items = sheetToObjects(getOrCreate(INV_SHEET, INV_HEADERS));
  return ok(items.filter(i => {
    const rp = parseInt(i.reorderPoint, 10);
    const q  = parseInt(i.qty, 10);
    return rp > 0 && !isNaN(q) && q <= rp;
  }));
}

// ─── TRANSACTIONS ─────────────────────────────────────────────────────────────
function doGetTransactions(p) {
  const all = sheetToObjects(getOrCreate(TXN_SHEET, TXN_HEADERS));
  if (p.partId) return ok(all.filter(t => t.partId === p.partId).reverse());
  return ok(all.slice(-300).reverse());
}

// ─── PURCHASE ORDER ───────────────────────────────────────────────────────────
function doSendPO(p) {
  const items = sheetToObjects(getOrCreate(INV_SHEET, INV_HEADERS));
  const part  = items.find(i => i.id === p.partId);
  if (!part) return fail('Part not found');

  const toEmail = p.toEmail || part.contact1;
  if (!toEmail) return fail('No supplier email set for this part');

  const settings    = getSettingsObj();
  const companyName = settings.companyName || 'Unipres UK';
  const fromName    = settings.fromName    || 'Stores Department';
  const qtyNeeded   = parseInt(p.qty, 10) || parseInt(part.qtyRequired, 10) || 1;

  MailApp.sendEmail({
    to:       toEmail,
    subject:  'Purchase Order Request — ' + (part.partNumber || part.description),
    name:     fromName + ' — ' + companyName,
    htmlBody: buildPOHtml(part, qtyNeeded, companyName, fromName, p.notes || '', p.poNumber || '')
  });

  // Update PO number on row if provided
  if (p.poNumber) {
    const s = getOrCreate(INV_SHEET, INV_HEADERS);
    const data = s.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('id');
    const poCol = headers.indexOf('unipressPO');
    for (let r = 1; r < data.length; r++) {
      if (data[r][idCol] === p.partId) { s.getRange(r + 1, poCol + 1).setValue(p.poNumber); break; }
    }
  }

  return ok({ sent: true, to: toEmail });
}

function tryAutoSendPO(part) {
  const settings = getSettingsObj();
  if (settings.autoPO !== 'true') return;
  const email = part.contact1 || settings.defaultSupplierEmail;
  if (!email) return;
  doSendPO({ partId: part.id, toEmail: email });
}

function buildPOHtml(part, qty, company, sender, notes, poNumber) {
  const date = new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' });
  const poRef = poNumber ? ('<br><span style="font-size:13px;opacity:0.7;">PO Ref: ' + poNumber + '</span>') : '';
  function tr(label, val) {
    return '<tr style="border-bottom:1px solid #e5e7eb;">'
      + '<td style="padding:10px 12px;font-size:13px;color:#6b7280;width:160px;">' + label + '</td>'
      + '<td style="padding:10px 12px;font-size:13px;color:#111827;font-weight:500;">' + (val || '—') + '</td>'
      + '</tr>';
  }
  return '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;padding:24px;">'
    + '<div style="border:2px solid #1e40af;border-radius:8px;overflow:hidden;">'
    + '<div style="background:#1e40af;color:white;padding:20px 24px;">'
    + '<h2 style="margin:0;font-size:20px;">Purchase Order Request</h2>'
    + '<p style="margin:4px 0 0;opacity:0.8;font-size:13px;">' + company + ' — ' + date + poRef + '</p>'
    + '</div>'
    + '<div style="padding:24px;">'
    + '<table style="width:100%;border-collapse:collapse;margin-bottom:20px;">'
    + tr('Part Number',       part.partNumber)
    + tr('Description',       part.description)
    + tr('OEM / Manufacturer',part.oem)
    + tr('Quantity Required', qty)
    + tr('Current Stock',     part.qty)
    + tr('Unit Cost (Ref)',   part.cost ? '£' + parseFloat(part.cost).toFixed(2) : null)
    + '</table>'
    + (notes ? '<div style="background:#fef3c7;border:1px solid #fde68a;border-radius:6px;padding:12px;margin-bottom:16px;">'
       + '<p style="margin:0;font-size:13px;color:#92400e;"><strong>Notes:</strong> ' + notes + '</p></div>' : '')
    + '<p style="color:#374151;font-size:14px;">Please confirm availability, pricing, and lead time at your earliest convenience.</p>'
    + '<p style="margin-top:16px;color:#374151;font-size:14px;">Kind regards,<br><strong>' + sender + '</strong><br>' + company + '</p>'
    + '</div></div>'
    + '<p style="margin:12px 0 0;font-size:11px;color:#9ca3af;text-align:center;">Unipres Stores Management System</p>'
    + '</div>';
}

// ─── SETTINGS ─────────────────────────────────────────────────────────────────
function doGetSettings() { return ok(getSettingsObj()); }

function getSettingsObj() {
  try {
    const s    = getOrCreate(SET_SHEET, SET_HEADERS);
    const data = s.getDataRange().getValues();
    const obj  = {};
    for (let r = 1; r < data.length; r++) { if (data[r][0]) obj[data[r][0]] = data[r][1]; }
    return obj;
  } catch (e) { return {}; }
}

function doUpdateSettings(p) {
  const settings = JSON.parse(p.settings || '{}');
  const s    = getOrCreate(SET_SHEET, SET_HEADERS);
  const data = s.getDataRange().getValues();
  const idx  = {};
  for (let r = 1; r < data.length; r++) { if (data[r][0]) idx[data[r][0]] = r + 1; }
  Object.keys(settings).forEach(key => {
    if (idx[key]) s.getRange(idx[key], 2).setValue(settings[key]);
    else          s.appendRow([key, settings[key]]);
  });
  return ok({ updated: true });
}

// ─── IMPORT FROM EXISTING SHEET ───────────────────────────────────────────────
// Maps your existing column names to the Inventory fields.
// Default map covers the exact column names from your stores spreadsheet.
function doImportFromSheet(p) {
  if (!p.sheetId) return fail('sheetId required');

  const ss       = SpreadsheetApp.openById(p.sheetId);
  const srcSheet = p.sheetName ? ss.getSheetByName(p.sheetName) : ss.getSheets()[0];
  if (!srcSheet) return fail('Source sheet not found');

  const data = srcSheet.getDataRange().getValues();
  if (data.length < 2) return fail('No data rows found');

  const srcHeaders = data[0];
  // colMap: source column name → destination field name
  const colMap = p.columnMap ? JSON.parse(p.columnMap) : buildDefaultColMap(srcHeaders);

  const dest    = getOrCreate(INV_SHEET, INV_HEADERS);
  const existing = sheetToObjects(dest).map(i => String(i.partNumber));

  let imported = 0, skipped = 0;
  for (let r = 1; r < data.length; r++) {
    const srcRow = data[r];
    if (!srcRow.some(c => c !== '')) { skipped++; continue; }

    const part = {};
    srcHeaders.forEach((h, i) => {
      const dest = colMap[h];
      if (dest) part[dest] = srcRow[i];
    });

    if (!part.description && !part.partNumber) { skipped++; continue; }
    if (part.partNumber && existing.includes(String(part.partNumber))) { skipped++; continue; }

    part.id          = uid();
    part.lastUpdated = new Date().toISOString();
    dest.appendRow(INV_HEADERS.map(h => part[h] !== undefined ? part[h] : ''));
    imported++;
  }
  return ok({ imported: imported, skipped: skipped });
}

// Builds a best-effort column map from detected headers
function buildDefaultColMap(headers) {
  // Maps your spreadsheet's exact header names → Inventory field names
  const knownMap = {
    'Item':                 'item',
    'Qty Required':         'qtyRequired',
    'Item Heading':         'itemHeading',
    'Description':          'description',
    'OEM':                  'oem',
    '3 Quotes Required?':   'threeQuotesRequired',
    'Part Number':          'partNumber',
    'Link':                 'link',
    'Qty':                  'qty',
    'Cost':                 'cost',
    'Quotation 1':          'quotation1',
    'Quotation 2':          'quotation2',
    'Quotation 3':          'quotation3',
    'Contact 1':            'contact1',
    'Contact 2':            'contact2',
    'Contact 3':            'contact3',
    'Unipres PO':           'unipressPO',
    'PO Copy':              'poCopy',
    'CAD':                  'cad'
  };
  const map = {};
  headers.forEach(h => { if (knownMap[h]) map[h] = knownMap[h]; });
  return map;
}
