/**
 * Mantis - write endpoint for Meter Reading (Running Hour) & Maintenance activity entries.
 *
 * Deploy steps:
 * 1. Buka spreadsheet Mantis (ID: 11-sqiV3BYqx738Rhr3Y9SF4Q7FnAGKNqMn-Gvoq6rFs) di Google Sheets.
 * 2. Menu Extensions > Apps Script.
 * 3. Hapus isi default, tempel seluruh isi file ini.
 * 4. Menu Project Settings (ikon gerigi) > Script Properties > Add script property:
 *      Property: MANTIS_WRITE_TOKEN
 *      Value:    (buat string acak yang rahasia, contoh: hasil dari `openssl rand -hex 16`)
 * 5. Menu Deploy > New deployment > pilih tipe "Web app".
 *      Execute as: Me
 *      Who has access: Anyone
 *    Klik Deploy, izinkan akses saat diminta.
 * 6. Salin "Web app URL" yang muncul.
 * 7. Di index.html, isi:
 *      const MANTIS_METER_WRITE_URL = '<Web app URL dari langkah 6>';
 *      const MANTIS_METER_WRITE_TOKEN = '<value yang sama dengan MANTIS_WRITE_TOKEN>';
 *
 * Setiap kali kode ini diubah, buat "New deployment" baru (bukan cukup Save) agar
 * perubahan aktif di Web app URL.
 *
 * Daftar teknisi pada popup "Teknisi" di Activities diambil dari sheet "User"
 * (kolom username & role), hanya baris dengan role = "Teknisi" yang muncul.
 * Kolom password pada sheet itu tidak pernah dikirim ke browser.
 *
 * Login (index.html) memvalidasi username/password terhadap sheet "User"
 * (kolom username/nama, password, role) lewat handleLogin() di bawah - juga
 * server-side saja, responsnya cuma success/username/role, tidak pernah
 * mengembalikan password. Tidak butuh MANTIS_WRITE_TOKEN karena ini justru
 * langkah untuk mendapatkan akses.
 */

var SHEET_ID = '11-sqiV3BYqx738Rhr3Y9SF4Q7FnAGKNqMn-Gvoq6rFs';
var MACHINE_COLUMN_LABEL = 'Nama Mesin';
var USER_SHEET_NAME = 'Users';

function doGet(e) {
  try {
    var action = (e.parameter && e.parameter.action || '').toString();

    if (action === 'teknisiList') {
      return handleTeknisiList(e);
    }

    if (action === 'bufferMeta') {
      return handleBufferMeta(e);
    }

    return jsonResponse({ success: false, error: 'Unknown action' });
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

/**
 * Returns when a "User_<username>" buffer tab (see handleSyncBuffer) was last
 * pushed, so a client can decide whether ITS local queue or the cloud buffer
 * is more current before doing a full gviz read - cheap enough to call on
 * every Entries page visit without re-fetching the whole buffer table each
 * time. Stored in Script Properties (not a sheet cell) since it has nothing
 * to do with the buffer tab's own row/column layout, and needs to read back
 * even when the tab currently holds zero rows (an empty queue is a valid,
 * meaningfully-timestamped state, not "no info").
 */
function handleBufferMeta(e) {
  var expectedToken = PropertiesService.getScriptProperties().getProperty('MANTIS_WRITE_TOKEN');
  var token = e.parameter && e.parameter.token;
  if (!expectedToken || token !== expectedToken) {
    return jsonResponse({ success: false, error: 'Unauthorized' });
  }

  var sheetName = (e.parameter && e.parameter.sheet || '').toString().trim();
  if (!/^User_/.test(sheetName)) {
    return jsonResponse({ success: false, error: 'Buffer sheet name must start with "User_"' });
  }

  var raw = PropertiesService.getScriptProperties().getProperty('bufferUpdatedAt_' + sheetName);
  return jsonResponse({ success: true, updatedAt: raw ? Number(raw) : 0 });
}

/**
 * Returns technician names (role = Teknisi) from the "User" sheet, for the
 * Activities "Problem & Action" Teknisi picker. Only the username/role
 * columns are read - the password column is never included in the response,
 * since this runs server-side and the sheet's gviz endpoint (used elsewhere
 * in index.html) is public and would otherwise leak it to the browser.
 */
function handleTeknisiList(e) {
  var expectedToken = PropertiesService.getScriptProperties().getProperty('MANTIS_WRITE_TOKEN');
  var token = e.parameter && e.parameter.token;
  if (!expectedToken || token !== expectedToken) {
    return jsonResponse({ success: false, error: 'Unauthorized' });
  }

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(USER_SHEET_NAME);
  if (!sheet) {
    return jsonResponse({ success: false, error: 'Sheet not found: ' + USER_SHEET_NAME });
  }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  var nameCol = -1, roleCol = -1;
  for (var i = 0; i < headers.length; i++) {
    var label = String(headers[i]).trim().toLowerCase();
    if (label === 'username' || label === 'nama') nameCol = i;
    if (label === 'role') roleCol = i;
  }
  if (nameCol === -1 || roleCol === -1) {
    return jsonResponse({ success: false, error: 'Kolom username/role tidak ditemukan di sheet ' + USER_SHEET_NAME });
  }

  var data = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  var names = [];
  data.forEach(function (row) {
    var role = String(row[roleCol] || '').trim().toLowerCase();
    var name = String(row[nameCol] || '').trim();
    if (name && role === 'teknisi') names.push(name);
  });

  return jsonResponse({ success: true, names: names });
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);

    if (body.action === 'login') {
      return handleLogin(body);
    }

    var expectedToken = PropertiesService.getScriptProperties().getProperty('MANTIS_WRITE_TOKEN');
    if (!expectedToken || body.token !== expectedToken) {
      return jsonResponse({ success: false, error: 'Unauthorized' });
    }

    if (body.action === 'appendRow') {
      return handleAppendRow(body);
    }

    if (body.action === 'syncBuffer') {
      return handleSyncBuffer(body);
    }

    return handleSetMeterReading(body);
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

/**
 * Checks username/password against the "User" sheet, server-side only - the
 * password column never leaves this function. Returns just success/username/
 * role so the browser has something to show without ever seeing the password
 * itself (the sheet's public gviz read, used elsewhere in index.html, would
 * otherwise leak it).
 */
function handleLogin(body) {
  var username = (body.username || '').toString().trim();
  var password = (body.password || '').toString();

  if (!username || !password) {
    return jsonResponse({ success: false, error: 'Username dan password wajib diisi' });
  }

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(USER_SHEET_NAME);
  if (!sheet) {
    return jsonResponse({ success: false, error: 'Sheet not found: ' + USER_SHEET_NAME });
  }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  var nameCol = -1, roleCol = -1, passCol = -1;
  for (var i = 0; i < headers.length; i++) {
    var label = String(headers[i]).trim().toLowerCase();
    if (label === 'username' || label === 'nama') nameCol = i;
    if (label === 'role') roleCol = i;
    if (label === 'password') passCol = i;
  }
  if (nameCol === -1 || passCol === -1) {
    return jsonResponse({ success: false, error: 'Kolom username/password tidak ditemukan di sheet ' + USER_SHEET_NAME });
  }

  var data = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  for (var r = 0; r < data.length; r++) {
    var rowName = String(data[r][nameCol] || '').trim();
    if (rowName.toLowerCase() !== username.toLowerCase()) continue;

    var rowPass = String(data[r][passCol] || '');
    if (rowPass !== password) {
      // Same generic error as "user not found" below - don't reveal which
      // part (username vs password) was wrong.
      return jsonResponse({ success: false, error: 'Username atau password salah' });
    }

    var role = roleCol === -1 ? '' : String(data[r][roleCol] || '').trim();
    return jsonResponse({ success: true, username: rowName, role: role });
  }

  return jsonResponse({ success: false, error: 'Username atau password salah' });
}

var APPEND_ROW_ID_COLUMN_LABEL = 'EntryID';

/**
 * Appends one or more rows to the given sheet, mapping body.values (an object,
 * or an array of objects, keyed by header label) onto the matching columns.
 * Used by the Activities "Problem & Action" queue to publish the actions
 * collected during a shift after the existing reports, in one call.
 *
 * Each incoming row carries an EntryID (the queue item's client-side id) so a
 * duplicated publish call - double click, a keepalive retry, or re-publishing
 * a still-pending queue in a later session - upserts (overwrites the row with
 * that EntryID) instead of appending a second copy. The EntryID column is
 * auto-created (like a missing week column in handleSetMeterReading) if the
 * sheet doesn't have it yet.
 */
function handleAppendRow(body) {
  var sheetName = (body.sheet || '').toString().trim();
  var values = body.values;

  if (!sheetName || !values) {
    return jsonResponse({ success: false, error: 'Missing sheet or values' });
  }

  var rowsToAppend = Array.isArray(values) ? values : [values];
  if (!rowsToAppend.length) {
    return jsonResponse({ success: false, error: 'No values to append' });
  }

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return jsonResponse({ success: false, error: 'Sheet not found: ' + sheetName });
  }

  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  var idCol = -1;
  for (var h = 0; h < headers.length; h++) {
    if (String(headers[h]).trim() === APPEND_ROW_ID_COLUMN_LABEL) { idCol = h; break; }
  }
  if (idCol === -1) {
    idCol = headers.length;
    headers.push(APPEND_ROW_ID_COLUMN_LABEL);
    sheet.getRange(1, idCol + 1).setValue(APPEND_ROW_ID_COLUMN_LABEL);
    lastCol = headers.length;
  }

  var lastRow = sheet.getLastRow();
  var existingIds = lastRow > 1 ? sheet.getRange(2, idCol + 1, lastRow - 1, 1).getValues() : [];
  var idToRow = {};
  existingIds.forEach(function (row, i) {
    var id = String(row[0] || '').trim();
    if (id) idToRow[id] = i + 2;
  });

  var rowToMatrix = function (entryValues) {
    return headers.map(function (header) {
      var key = String(header).trim();
      if (!key || !entryValues || typeof entryValues !== 'object' || !(key in entryValues)) return '';
      return parseAppendValue(entryValues[key]);
    });
  };

  var toAppend = [];
  var updated = 0;
  rowsToAppend.forEach(function (entryValues) {
    var id = entryValues && typeof entryValues === 'object' ? String(entryValues[APPEND_ROW_ID_COLUMN_LABEL] || '').trim() : '';
    var existingRow = id ? idToRow[id] : undefined;
    if (existingRow) {
      sheet.getRange(existingRow, 1, 1, lastCol).setValues([rowToMatrix(entryValues)]);
      updated++;
    } else {
      toAppend.push(rowToMatrix(entryValues));
    }
  });

  var startRow = sheet.getLastRow() + 1;
  if (toAppend.length) {
    sheet.getRange(startRow, 1, toAppend.length, lastCol).setValues(toAppend);
  }

  return jsonResponse({ success: true, row: startRow, appended: toAppend.length, updated: updated, count: rowsToAppend.length });
}

function parseAppendValue(value) {
  if (typeof value === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(value.trim())) {
    return new Date(value.trim() + 'T00:00:00');
  }
  if (typeof value === 'string' && /^-?\d+(\.\d+)?$/.test(value.trim())) {
    return Number(value);
  }
  return value;
}

var BUFFER_SHEET_HEADERS = ['EntryID', 'Area', 'Mesin', 'Issue', 'Action', 'Link', 'Link 2', 'Link 3', 'Link 4'];

/**
 * Mirrors a technician's local "Problem & Action" queue (index.html's
 * localStorage mantisPendingActions) into their own per-account sheet tab -
 * e.g. "User_Hariyanto" - as a cross-device backup. This is a buffer, not a
 * log: every call replaces that tab's data rows wholesale with body.values
 * (whatever the client's current local queue is right now), rather than
 * appending/upserting like handleAppendRow does for Records.
 *
 * sheet must start with "User_" - this endpoint wipes and rewrites whatever
 * sheet it's given, so without this check a client bug (or anyone else with
 * the write token) could pass "Records" or "Downtime" and blank out real
 * data. The buffer tabs are expected to already exist (created ahead of
 * time per account) but are auto-created with headers if missing, so a new
 * account's first sync doesn't fail.
 */
function handleSyncBuffer(body) {
  var sheetName = (body.sheet || '').toString().trim();
  if (!/^User_/.test(sheetName)) {
    return jsonResponse({ success: false, error: 'Buffer sheet name must start with "User_"' });
  }

  // Client sends the timestamp of the local change that triggered this push
  // (not "now" - a slow/retried request can arrive late). If an older-timestamped
  // push shows up after a newer one already landed (out-of-order delivery),
  // skip it rather than regressing the buffer back to stale data - the device
  // that sent it still has its own locally-queued edit and will catch up
  // (and re-push) once bufferMeta tells it the cloud has moved ahead.
  var propKey = 'bufferUpdatedAt_' + sheetName;
  var props = PropertiesService.getScriptProperties();
  var incomingUpdatedAt = Number(body.updatedAt) || Date.now();
  var storedUpdatedAt = Number(props.getProperty(propKey)) || 0;
  if (storedUpdatedAt && incomingUpdatedAt < storedUpdatedAt) {
    return jsonResponse({ success: true, skipped: true, reason: 'stale', updatedAt: storedUpdatedAt });
  }

  var rows = Array.isArray(body.values) ? body.values : [];

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  var headers = BUFFER_SHEET_HEADERS;
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    var lastCol = Math.max(sheet.getLastColumn(), 1);
    var existingHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    if (existingHeaders[0]) {
      headers = existingHeaders.map(function (h) { return String(h); });
    } else {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }

  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }

  if (rows.length) {
    var matrix = rows.map(function (entryValues) {
      return headers.map(function (header) {
        var key = String(header).trim();
        if (!key || !entryValues || typeof entryValues !== 'object' || !(key in entryValues)) return '';
        return entryValues[key];
      });
    });
    sheet.getRange(2, 1, matrix.length, headers.length).setValues(matrix);
  }

  props.setProperty(propKey, String(incomingUpdatedAt));

  return jsonResponse({ success: true, count: rows.length, updatedAt: incomingUpdatedAt });
}

function handleSetMeterReading(body) {
  try {
    var sheetName = (body.sheet || 'Meters').toString();
    var machine = (body.machine || '').toString().trim();
    var week = (body.week || '').toString().trim();
    var value = body.value;

    if (!machine || !/^W\d+$/i.test(week) || value === undefined || value === null || value === '') {
      return jsonResponse({ success: false, error: 'Missing or invalid machine, week, or value' });
    }
    week = week.toUpperCase();

    // value bisa berupa angka (reading hour/count) atau tanggal 'YYYY-MM-DD' (reading tanggal).
    var cellValue;
    if (typeof value === 'number') {
      cellValue = value;
    } else if (typeof value === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(value.trim())) {
      cellValue = new Date(value.trim() + 'T00:00:00');
    } else if (typeof value === 'string' && value.trim() !== '' && !isNaN(Number(value))) {
      cellValue = Number(value);
    } else {
      return jsonResponse({ success: false, error: 'Invalid value format' });
    }

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return jsonResponse({ success: false, error: 'Sheet not found: ' + sheetName });
    }

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

    var machineCol = -1;
    for (var h = 0; h < headers.length; h++) {
      if (String(headers[h]).trim() === MACHINE_COLUMN_LABEL) { machineCol = h; break; }
    }
    if (machineCol === -1) {
      return jsonResponse({ success: false, error: 'Column "' + MACHINE_COLUMN_LABEL + '" not found' });
    }

    var machineValues = sheet.getRange(2, machineCol + 1, Math.max(lastRow - 1, 0), 1).getValues();
    var rowIndex = -1;
    for (var i = 0; i < machineValues.length; i++) {
      if (String(machineValues[i][0]).trim().toLowerCase() === machine.toLowerCase()) {
        rowIndex = i + 2;
        break;
      }
    }
    if (rowIndex === -1) {
      return jsonResponse({ success: false, error: 'Machine not found: ' + machine });
    }

    var weekCol = -1;
    for (var w = 0; w < headers.length; w++) {
      if (String(headers[w]).trim().toUpperCase() === week) { weekCol = w; break; }
    }

    var targetCol;
    if (weekCol === -1) {
      targetCol = lastCol + 1;
      sheet.getRange(1, targetCol).setValue(week);
    } else {
      targetCol = weekCol + 1;
    }

    sheet.getRange(rowIndex, targetCol).setValue(cellValue);

    return jsonResponse({ success: true, row: rowIndex, col: targetCol, week: week });
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
