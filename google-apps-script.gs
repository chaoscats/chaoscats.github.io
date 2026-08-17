/**
 * ChaosCats Whitelist — Google Apps Script Web App
 * ------------------------------------------------
 * Menerima submit dari form Whitelist di website dan menuliskannya
 * sebagai baris baru ke Google Sheet ini.
 *
 * CARA PASANG:
 * 1. Buka Google Sheet kosong (atau buat baru) → Extensions > Apps Script.
 * 2. Hapus isi default, lalu paste seluruh isi file ini.
 * 3. Klik "Deploy" > "New deployment".
 *    - Type: "Web app"
 *    - Description: bebas, mis. "ChaosCats WL v1"
 *    - Execute as: "Me"
 *    - Who has access: "Anyone"
 * 4. Klik "Deploy", izinkan akses (Authorize access) saat diminta.
 * 5. Copy URL Web App yang muncul (diakhiri /exec).
 * 6. Paste URL tersebut ke variabel WL_CONFIG.SHEET_WEBHOOK_URL
 *    di index.html (cari teks PASTE_YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE).
 *
 * Setiap kali script diedit & di-deploy ulang, pilih "Manage deployments"
 * > edit > New version, supaya URL /exec tetap sama dan website tidak perlu diubah lagi.
 */

const SHEET_NAME = 'Whitelist'; // nama tab sheet tempat data akan ditulis

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    const sheet = getOrCreateSheet_();

    sheet.appendRow([
      new Date(),                 // waktu server menerima submit
      data.name || '',
      data.wallet || '',
      data.followed ? 'YES' : 'NO',
      data.reposted ? 'YES' : 'NO',
      data.timestamp || '',       // waktu dari browser user
      data.source || ''
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getOrCreateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      'Timestamp (Server)',
      'Name / X Handle',
      'Wallet Address',
      'Followed on X',
      'Liked & Reposted',
      'Timestamp (Browser)',
      'Source'
    ]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}
