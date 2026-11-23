/**
 * PAPAN MARKAH — MISI ARKIB SEJARAH (Escape Room Sejarah Tingkatan 4: Bab 1, 3, 8)
 * Cikgu Rabbani, SMKDUMA · CG R 2026
 *
 * CARA PASANG (guna akaun Gmail peribadi):
 * 1. Buka Google Sheet baharu > Extensions > Apps Script.
 * 2. Padam semua kod sedia ada, tampal SEMUA kod ini, tekan Save.
 * 3. Deploy > New deployment > pilih jenis "Web app".
 *      Execute as : Me
 *      Who has access : Anyone
 * 4. Tekan Deploy, benarkan akses (Authorize), salin "Web app URL" (berakhir /exec).
 * 5. Dalam index.html, tampal URL itu pada  SHEET_URL: "..."  kemudian upload ke GitHub.
 *
 * Sheet "Markah_T4" akan dicipta secara automatik pada markah pertama.
 * Jika kod ini diubah, buat Deploy > Manage deployments > Edit > Version: New version.
 */

const SHEET_NAME = 'Markah_T4';
const HEADERS = ['Tarikh', 'Permainan', 'Topik', 'Nama', 'Kelas', 'Kod Bilik', 'Bilik', 'Markah', 'Masa (saat)', 'Kunci Dibuka', 'Lolos', 'Bintang'];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#2b2318').setFontColor('#ecd27a');
  }
  return sh;
}

// Elak suntikan formula & potong teks terlalu panjang
function clean_(v, max) {
  let s = String(v == null ? '' : v).trim().slice(0, max || 60);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (!d.nama || !d.kelas) return out_({ ok: false, error: 'Nama/kelas kosong' });
    getSheet_().appendRow([
      new Date(),
      clean_(d.game, 30),
      clean_(d.topik, 10),
      clean_(d.nama, 50),
      clean_(String(d.kelas).toUpperCase(), 20),
      clean_(d.bilik, 10),
      clean_(d.namaBilik, 40),
      Math.max(0, Math.min(5000, Number(d.markah) || 0)),
      Math.max(0, Number(d.masa) || 0),
      clean_(d.kunci, 10),
      clean_(d.lolos, 6),
      Number(d.bintang) || 0
    ]);
    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  const game = p.game || '';
  const values = getSheet_().getDataRange().getValues();
  values.shift();
  const rows = values
    .filter(r => r[3] && (!game || r[1] === game))
    .map(r => ({
      t: r[0], game: r[1], topik: r[2], nama: r[3], kelas: r[4], bilik: r[5], namaBilik: r[6],
      markah: r[7], masa: r[8], kunci: r[9], lolos: r[10], bintang: r[11]
    }));
  return out_({ ok: true, rows: rows }, p.callback);
}

function out_(obj, callback) {
  const json = JSON.stringify(obj);
  if (callback && /^[A-Za-z0-9_]+$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
