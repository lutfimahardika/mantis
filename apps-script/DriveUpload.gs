/**
 * Mantis - Drive upload endpoint for report photos (Activities > New Action).
 *
 * Ini SENGAJA dipisah dari Code.gs: Code.gs berjalan di akun pemilik Sheet
 * (lutfimahardika@gmail.com), sedangkan script ini harus di-deploy dari akun
 * bottomfolder@hcloud.my.id supaya foto laporan tersimpan di Drive akun itu.
 *
 * Deploy steps (login dulu sebagai bottomfolder@hcloud.my.id):
 * 1. Buka https://script.google.com/, klik "New project".
 * 2. Hapus isi default, tempel seluruh isi file ini.
 * 3. Menu Project Settings (ikon gerigi) > Script Properties > Add script property:
 *      Property: MANTIS_DRIVE_TOKEN
 *      Value:    (buat string acak yang rahasia, contoh: hasil dari `openssl rand -hex 16`;
 *                 boleh sama atau beda dengan MANTIS_WRITE_TOKEN punya Code.gs)
 * 4. (Opsional) Ganti DRIVE_FOLDER_NAME di bawah kalau mau nama folder lain -
 *    foldernya otomatis dibuat di My Drive akun ini kalau belum ada.
 * 5. Menu Deploy > New deployment > pilih tipe "Web app".
 *      Execute as: Me
 *      Who has access: Anyone
 *    Klik Deploy, izinkan akses saat diminta.
 * 6. Salin "Web app URL" yang muncul.
 * 7. Di index.html, isi:
 *      const MANTIS_DRIVE_UPLOAD_URL = '<Web app URL dari langkah 6>';
 *      const MANTIS_DRIVE_UPLOAD_TOKEN = '<value yang sama dengan MANTIS_DRIVE_TOKEN>';
 *
 * Setiap kali kode ini diubah, buat "New deployment" baru (bukan cukup Save) agar
 * perubahan aktif di Web app URL.
 *
 * Foto yang diupload disimpan di folder DRIVE_FOLDER_NAME pada Drive akun yang
 * deploy script ini, lalu dibagikan sebagai "Anyone with the link - Viewer"
 * supaya link-nya (disimpan di kolom "Link" sheet Records) bisa dibuka
 * langsung tanpa perlu login ke akun manapun.
 */

var DRIVE_FOLDER_NAME = 'Mantis - Foto Laporan';

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);

    var expectedToken = PropertiesService.getScriptProperties().getProperty('MANTIS_DRIVE_TOKEN');
    if (!expectedToken || body.token !== expectedToken) {
      return jsonResponse({ success: false, error: 'Unauthorized' });
    }

    return handleUpload(body);
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

function handleUpload(body) {
  var base64 = (body.data || '').toString();
  var mimeType = (body.mimeType || 'image/jpeg').toString();
  var filename = (body.filename || ('laporan_' + Date.now() + '.jpg')).toString();

  if (!base64) {
    return jsonResponse({ success: false, error: 'Missing image data' });
  }

  var bytes;
  try {
    bytes = Utilities.base64Decode(base64);
  } catch (err) {
    return jsonResponse({ success: false, error: 'Invalid base64 image data' });
  }

  var blob = Utilities.newBlob(bytes, mimeType, filename);
  var folder = getOrCreateFolder(DRIVE_FOLDER_NAME);
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return jsonResponse({
    success: true,
    url: file.getUrl(),
    fileId: file.getId()
  });
}

function getOrCreateFolder(name) {
  var folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
