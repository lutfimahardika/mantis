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
 *
 * Foto pada "New Action"/"Edit Action" (Activities > Problem & Action) diupload
 * lewat handleUploadImage() ke folder Drive MANTIS_PHOTOS_FOLDER_NAME milik akun
 * yang men-deploy Web app ini (karena "Execute as: Me" = identitas deployer).
 * Supaya foto masuk ke Drive akun "bottomfolder@hcloud.my.id", deploy WAJIB
 * dilakukan sambil login sebagai akun tsb:
 * 1. Share spreadsheet Mantis ke bottomfolder@hcloud.my.id sebagai Editor
 *    (kalau belum).
 * 2. Login Google sebagai bottomfolder@hcloud.my.id, buka spreadsheet Mantis,
 *    lalu ikuti "Deploy steps" di atas seperti biasa (Execute as: Me, Who has
 *    access: Anyone) dari akun tsb.
 * 3. Salin Web app URL hasil deploy itu ke MANTIS_METER_WRITE_URL di index.html.
 * Link hasil upload disimpan lewat handleAppendRow() ke kolom "Link Gambar" di
 * sheet "Maintenance" - kolom ini otomatis dibuat kalau belum ada, tidak perlu
 * ditambah manual.
 */

var SHEET_ID = '11-sqiV3BYqx738Rhr3Y9SF4Q7FnAGKNqMn-Gvoq6rFs';
var MACHINE_COLUMN_LABEL = 'Nama Mesin';
var USER_SHEET_NAME = 'User';
var MANTIS_PHOTOS_FOLDER_NAME = 'Mantis - Problem & Action Photos';

function doGet(e) {
  try {
    var action = (e.parameter && e.parameter.action || '').toString();

    if (action === 'teknisiList') {
      return handleTeknisiList(e);
    }

    return jsonResponse({ success: false, error: 'Unknown action' });
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
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

    if (body.action === 'uploadImage') {
      return handleUploadImage(body);
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

/**
 * Appends one or more rows at the end of the given sheet, mapping body.values
 * (an object, or an array of objects, keyed by header label) onto the matching
 * columns. Used by the Activities "Problem & Action" queue to publish the
 * actions collected during a shift after the existing reports, in one call.
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
  var headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];

  // Any key present in the incoming rows that doesn't match an existing
  // header gets its own new column appended at the end - e.g. "Link Gambar"
  // for the New/Edit Action photo feature - so the sheet doesn't need to be
  // hand-edited first.
  var headerSet = {};
  headers.forEach(function (h) { headerSet[String(h).trim()] = true; });
  var newHeaders = [];
  rowsToAppend.forEach(function (entryValues) {
    if (!entryValues || typeof entryValues !== 'object') return;
    Object.keys(entryValues).forEach(function (key) {
      var trimmed = String(key).trim();
      if (trimmed && !headerSet[trimmed]) {
        headerSet[trimmed] = true;
        newHeaders.push(trimmed);
      }
    });
  });
  if (newHeaders.length) {
    sheet.getRange(1, headers.length + 1, 1, newHeaders.length).setValues([newHeaders]);
    headers = headers.concat(newHeaders);
  }

  var startRow = sheet.getLastRow() + 1;

  var matrix = rowsToAppend.map(function (entryValues) {
    return headers.map(function (header) {
      var key = String(header).trim();
      if (!key || !entryValues || typeof entryValues !== 'object' || !(key in entryValues)) return '';
      return parseAppendValue(entryValues[key]);
    });
  });

  sheet.getRange(startRow, 1, matrix.length, headers.length).setValues(matrix);

  return jsonResponse({ success: true, row: startRow, count: matrix.length });
}

/**
 * Uploads one photo (base64, sent from the New/Edit Action modal) to the
 * MANTIS_PHOTOS_FOLDER_NAME folder in the Drive of whichever account this Web
 * app is deployed as, and shares it view-only "anyone with link" so the URL
 * saved in the sheet's "Link Gambar" column can be opened by anyone with
 * access to the sheet.
 */
function handleUploadImage(body) {
  var base64Data = (body.imageBase64 || '').toString();
  if (!base64Data) {
    return jsonResponse({ success: false, error: 'Missing imageBase64' });
  }

  var mimeType = (body.mimeType || 'image/jpeg').toString();
  var commaIdx = base64Data.indexOf(',');
  if (base64Data.slice(0, 5) === 'data:' && commaIdx !== -1) {
    var mimeMatch = base64Data.slice(5, commaIdx).split(';')[0];
    if (mimeMatch) mimeType = mimeMatch;
    base64Data = base64Data.slice(commaIdx + 1);
  }

  var filename = (body.filename || ('mantis_' + new Date().getTime())).toString().trim();
  if (filename.indexOf('.') === -1) {
    filename += '.' + (mimeType.split('/')[1] || 'jpg');
  }

  var bytes;
  try {
    bytes = Utilities.base64Decode(base64Data);
  } catch (err) {
    return jsonResponse({ success: false, error: 'Invalid image data' });
  }

  var blob = Utilities.newBlob(bytes, mimeType, filename);
  var file = getMantisPhotosFolder().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return jsonResponse({ success: true, url: file.getUrl(), fileId: file.getId() });
}

function getMantisPhotosFolder() {
  var folders = DriveApp.getFoldersByName(MANTIS_PHOTOS_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(MANTIS_PHOTOS_FOLDER_NAME);
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

function handleSetMeterReading(body) {
  try {
    var sheetName = (body.sheet || 'Running Hour').toString();
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
