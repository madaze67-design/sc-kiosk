// ═══════════════════════════════════════════════════════════════════
// SC FABRICATION KIOSK — Google Apps Script Backend
// ═══════════════════════════════════════════════════════════════════
//
// SETUP INSTRUCTIONS:
// 1. Open your Google Sheet → Extensions → Apps Script
// 2. Delete all existing code, paste this entire file
// 3. Click Save (Ctrl+S)
// 4. Click Deploy → New Deployment
// 5. Type: Web App
// 6. Execute as: Me
// 7. Who has access: Anyone
// 8. Click Deploy → Copy the Web App URL
// 9. Paste that URL into the program's "Connect to Google Sheets" panel
//
// The script will automatically create a sheet named "SC_Users"
// in your spreadsheet to store all user data.
// ═══════════════════════════════════════════════════════════════════

const SHEET_NAME = 'SC_Users';
const HEADERS    = ['ID', 'Rank', 'Name', 'Password', 'Blueprints (JSON)', 'Materials (JSON)', 'Created', 'Updated', 'Role', 'Permissions'];

// ── Helpers ──────────────────────────────────────────────────────────────────

function getOrCreateSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    // Style header row
    const hdr = sheet.getRange(1, 1, 1, HEADERS.length);
    hdr.setBackground('#1a3530');
    hdr.setFontColor('#4dc8b0');
    hdr.setFontWeight('bold');
    sheet.setColumnWidth(5, 400); // Blueprints JSON
    sheet.setColumnWidth(6, 600); // Materials JSON
  }
  return sheet;
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function findUserRow(sheet, userId) {
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(userId)) return i + 1; // 1-based row
  }
  return -1;
}

function rowToUser(row) {
  return {
    id:         row[0],
    rank:       row[1],
    name:       row[2],
    password:   String(row[3] || ''),
    blueprints: tryParse(row[4], []),
    materials:  tryParse(row[5], []),
    created:    row[6],
    updated:    row[7],
    role:       String(row[8] || ''),
    permissions: String(row.length > 9 ? (row[9] || '') : ''),
  };
}

function normMat(name) {
  return (name||'').replace(/ \(Ore\)/g, '').trim();
}

function tryParse(val, fallback) {
  try { return JSON.parse(val || 'null') || fallback; }
  catch(e) { return fallback; }
}

// ── GET handler — all reads ───────────────────────────────────────────────────

function doGet(e) {
  const action = (e.parameter.action || '').toLowerCase();

  try {
    if (action === 'getusers' || action === 'getall') {
      const sheet  = getOrCreateSheet();
      const data   = sheet.getDataRange().getValues();
      const users  = data.length <= 1 ? [] : data.slice(1).map(rowToUser);
      return jsonResponse({ success: true, users });
    }

    if (action === 'saveuser') {
      const payload = tryParse(e.parameter.payload, null);
      if (!payload) return jsonResponse({ success: false, error: 'No payload' });
      return handleSaveUser(payload);
    }

    if (action === 'deleteuser') {
      const payload = tryParse(e.parameter.payload, null);
      if (!payload || !payload.id) return jsonResponse({ success: false, error: 'No user id' });
      return handleDeleteUser(payload.id);
    }

    if (action === 'getorders') {
      const sheet = getOrCreateOrdersSheet();
      const data  = sheet.getDataRange().getValues();
      const orders = data.length<=1 ? [] : data.slice(1).map(rowToOrder);
      return jsonResponse({success:true, orders});
    }
    if (action === 'saveorder') {
      const payload = tryParse(e.parameter.payload, null);
      if(!payload) return jsonResponse({success:false,error:'No payload'});
      return handleSaveOrder(payload);
    }
    if (action === 'getconfig') {
      const key = e.parameter.key || '';
      if(!key) return jsonResponse({success:false, error:'No key'});
      const value = getConfigValue(key);
      return jsonResponse({success: value !== null, value: value || ''});
    }
    if (action === 'setconfig') {
      const payload = tryParse(e.parameter.payload, null);
      if(!payload || !payload.key) return jsonResponse({success:false, error:'No key'});
      setConfigValue(payload.key, String(payload.value || ''));
      return jsonResponse({success:true});
    }
    if (action === 'ensureadmin') {
      // Check if an admin user exists; if not, create one with default password
      const sheet = getOrCreateSheet();
      const data  = sheet.getDataRange().getValues();
      const users = data.length <= 1 ? [] : data.slice(1).map(rowToUser);
      const existing = users.find(u => u.role === 'admin');
      if(existing) {
        return jsonResponse({ success: true, exists: true, id: existing.id });
      }
      // Create the admin user
      const now   = new Date().toISOString();
      const newId = Date.now();
      sheet.appendRow([newId,'Gnome Admin','Admin','rQjcpWA8KD','[]','[]',now,now,'admin']);
      return jsonResponse({ success: true, exists: false, id: newId });
    }
    if (action === 'deleteorder') {
      const payload2 = tryParse(e.parameter.payload, null);
      if(!payload2 || !payload2.id) return jsonResponse({success:false,error:'No id'});
      const os = getOrCreateOrdersSheet();
      const ov = os.getDataRange().getValues();
      for(let i=1;i<ov.length;i++) {
        if(String(ov[i][0])===String(payload2.id)) { os.deleteRow(i+1); return jsonResponse({success:true}); }
      }
      return jsonResponse({success:false,error:'Order not found'});
    }
    if (action === 'ping') {
      return jsonResponse({ success: true, message: 'SC Kiosk API online' });
    }

    return jsonResponse({ success: false, error: 'Unknown action: ' + action });

  } catch(err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

// ── POST handler — also handles writes (fallback) ────────────────────────────

function doPost(e) {
  try {
    // Support both form-urlencoded POST (preferred, no CORS preflight)
    // and JSON POST (fallback)
    let action, inner;
    if(e.parameter && e.parameter.action) {
      // Form-urlencoded: params available via e.parameter
      action = (e.parameter.action || '').toLowerCase();
      inner  = tryParse(e.parameter.payload, {});
    } else {
      // JSON body fallback
      const body = tryParse(e.postData ? e.postData.contents : null, null);
      if(!body) return jsonResponse({ success: false, error: 'Invalid payload' });
      action = (body.action || '').toLowerCase();
      inner  = body.payload !== undefined ? body.payload : body;
    }

    if (action === 'saveuser')     return handleSaveUser(inner);
    if (action === 'deleteuser')   return handleDeleteUser(inner.id);
    if (action === 'saveorder')    return handleSaveOrder(inner);
    if (action === 'deleteorder') {
      if(!inner.id) return jsonResponse({success:false,error:'No id'});
      const os2 = getOrCreateOrdersSheet();
      const ov2 = os2.getDataRange().getValues();
      for(let i=1;i<ov2.length;i++) { if(String(ov2[i][0])===String(inner.id)) { os2.deleteRow(i+1); return jsonResponse({success:true}); } }
      return jsonResponse({success:false,error:'Order not found'});
    }
    if (action === 'setconfig') {
      if(!inner.key) return jsonResponse({success:false, error:'No key'});
      setConfigValue(inner.key, String(inner.value || ''));
      return jsonResponse({success:true});
    }
    return jsonResponse({ success: false, error: 'Unknown action: ' + action });
  } catch(err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

// ── Write operations ──────────────────────────────────────────────────────────

function handleSaveUser(data) {
  const sheet   = getOrCreateSheet();
  const now     = new Date().toISOString();
  const rowIdx  = data.id ? findUserRow(sheet, data.id) : -1;

  if (rowIdx === -1) {
    // New user — insert
    const newId = data.id || Date.now();
    sheet.appendRow([
      newId,
      data.rank       || '',
      data.name       || '',
      String(data.password   || ''),
      JSON.stringify(data.blueprints || []),
      JSON.stringify(data.materials  || []),
      now,
      now,
      data.role         || '',
      data.permissions  || '',
    ]);
    return jsonResponse({ success: true, action: 'created', id: newId });
  } else {
    // Existing user — update only provided fields
    const range = sheet.getRange(rowIdx, 1, 1, 10);
    const row   = range.getValues()[0];
    const updated = [
      row[0],
      data.hasOwnProperty('rank')       ? data.rank       : row[1],
      data.hasOwnProperty('name')       ? data.name       : row[2],
      data.hasOwnProperty('password')   ? String(data.password || '') : String(row[3] || ''),
      data.hasOwnProperty('blueprints') ? JSON.stringify(data.blueprints) : row[4],
      data.hasOwnProperty('materials')  ? JSON.stringify(data.materials)  : row[5],
      row[6],
      now,
      data.hasOwnProperty('role')       ? data.role       : (row[8] || ''),
      data.hasOwnProperty('permissions') ? data.permissions : (row[9] || ''),
    ];
    range.setValues([updated]);
    return jsonResponse({ success: true, action: 'updated', id: row[0] });
  }
}

function handleDeleteUser(userId) {
  const sheet  = getOrCreateSheet();
  const rowIdx = findUserRow(sheet, userId);
  if (rowIdx === -1) return jsonResponse({ success: false, error: 'User not found' });
  sheet.deleteRow(rowIdx);
  return jsonResponse({ success: true, action: 'deleted', id: userId });
}


// ── Additional sheet names ────────────────────────────────────────────────────
const ORDERS_SHEET   = 'SC_Orders';

const ORDER_HEADERS  = ['ID','Placer Name','Callsign','RSI Handle','User ID','Items (JSON)','Materials (JSON)','Status','Assigned To','Created','Updated','Crafted Items (JSON)','Priority','Notes'];

function getOrCreateOrdersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(ORDERS_SHEET);
  if(!sheet) {
    sheet = ss.insertSheet(ORDERS_SHEET);
    sheet.appendRow(ORDER_HEADERS);
    sheet.setFrozenRows(1);
    const hdr = sheet.getRange(1,1,1,ORDER_HEADERS.length);
    hdr.setBackground('#1a1a2e'); hdr.setFontColor('#c8941a'); hdr.setFontWeight('bold');
    sheet.setColumnWidth(6,300); sheet.setColumnWidth(7,400); sheet.setColumnWidth(13,120); sheet.setColumnWidth(14,300);
  }
  return sheet;
}

function rowToOrder(row) {
  return {
    id:          row[0],
    placerName:  row[1],
    callsign:    row[2],
    rsiHandle:   row[3],
    userId:      row[4],
    items:       tryParse(row[5], []),
    materials:   tryParse(row[6], []),
    status:      row[7]  || 'New Order',
    assignedTo:  row[8]  || '',
    created:     row[9],
    updated:     row[10],
    crafted:     tryParse(row[11], {}),
    priority:    String(row[12] || ''),
    notes:       String(row[13] || ''),
  };
}

function handleSaveOrder(data) {
  const sheet  = getOrCreateOrdersSheet();
  const now    = new Date().toISOString();
  const values = sheet.getDataRange().getValues();
  let rowIdx = -1;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(data.id)) { rowIdx = i + 1; break; }
  }

  if (rowIdx === -1) {
    // New order — append all 14 columns
    sheet.appendRow([
      data.id,
      data.placerName  || '',
      data.callsign    || '',
      data.rsiHandle   || '',
      data.userId      || '',
      JSON.stringify(data.items     || []),
      JSON.stringify(data.materials || []),
      data.status      || 'New Order',
      data.assignedTo  || '',
      data.created     || now,
      now,
      JSON.stringify(data.crafted   || {}),
      data.priority    || '',
      data.notes       || '',
    ]);
    return jsonResponse({ success: true, action: 'created', id: data.id });
  }

  // Existing order — read current 14 cols, update only what was sent
  const existingRange = sheet.getRange(rowIdx, 1, 1, 14);
  const row = existingRange.getValues()[0];
  existingRange.setValues([[
    row[0],
    data.placerName || row[1],
    data.callsign   || row[2],
    data.rsiHandle  || row[3],
    data.userId     || row[4],
    data.items        ? JSON.stringify(data.items)      : row[5],
    data.materials    ? JSON.stringify(data.materials)  : row[6],
    data.status       || row[7],
    data.hasOwnProperty('assignedTo') ? data.assignedTo             : row[8],
    row[9],
    now,
    data.hasOwnProperty('crafted')    ? JSON.stringify(data.crafted) : (row[11] || '{}'),
    data.hasOwnProperty('priority')   ? (data.priority || '')        : (row[12] || ''),
    data.hasOwnProperty('notes')      ? (data.notes    || '')        : (row[13] || ''),
  ]]);
  return jsonResponse({ success: true, action: 'updated', id: data.id });
}

// ── SC_Config sheet — stores admin password and other server-side settings ───
const CONFIG_SHEET = 'SC_Config';

function getOrCreateConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG_SHEET);
  if(!sheet) {
    sheet = ss.insertSheet(CONFIG_SHEET);
    sheet.appendRow(['Key', 'Value', 'Updated']);
    sheet.setFrozenRows(1);
    const hdr = sheet.getRange(1,1,1,3);
    hdr.setBackground('#2a1a3e');
    hdr.setFontColor('#aa44ff');
    hdr.setFontWeight('bold');
    sheet.setColumnWidth(1, 180);
    sheet.setColumnWidth(2, 260);
    sheet.appendRow(['admin_password', 'rQjcpWA8KD', new Date().toISOString()]);
  }
  return sheet;
}

function getConfigValue(key) {
  const sheet = getOrCreateConfigSheet();
  const values = sheet.getDataRange().getValues();
  for(let i=1; i<values.length; i++) {
    if(String(values[i][0]) === key) return String(values[i][1] || '');
  }
  return null;
}

function setConfigValue(key, value) {
  const sheet = getOrCreateConfigSheet();
  const values = sheet.getDataRange().getValues();
  const now = new Date().toISOString();
  for(let i=1; i<values.length; i++) {
    if(String(values[i][0]) === key) {
      sheet.getRange(i+1, 2, 1, 2).setValues([[String(value), now]]);
      return;
    }
  }
  sheet.appendRow([key, String(value), now]);
}
