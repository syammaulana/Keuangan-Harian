/**
 * ═══════════════════════════════════════════════════════════
 *  BUKU KAS KELUARGA — Backend Sinkronisasi v3
 *  (Sinkron dua arah ke sheet TRANSAKSI, format kolom terbaru
 *   dengan Timestamp di kolom A)
 * ═══════════════════════════════════════════════════════════
 *
 * ⚠️ SEBELUM MULAI: File → Make a copy di Google Sheets ini dulu,
 *    simpan sebagai cadangan.
 *
 * STRUKTUR KOLOM YANG DIHARAPKAN DI SHEET "TRANSAKSI":
 *   A = Timestamp        (diisi otomatis oleh Form, atau kosong)
 *   B = Tanggal
 *   C = Jenis Transaksi
 *   D = Rekening
 *   E = Kategori
 *   F = Nominal
 *   G = Lokasi
 *   H = Keterangan
 *   I = ID   (kolom baru, dibuat & dikelola otomatis oleh script ini —
 *             JANGAN diisi/diedit manual)
 *
 * ⚠️ PENTING SOAL NAMA TAB: script ini mencari tab bernama persis
 *    "TRANSAKSI" (lihat SHEET_NAME di bawah). Kalau nama tab Anda
 *    berbeda, ganti nilai SHEET_NAME sebelum deploy.
 *
 * KENAPA ADA 2 TRIGGER (bukan cuma 1):
 * - onEdit  → mendeteksi saat Anda mengetik/mengedit manual langsung
 *   di sheet.
 * - onFormSubmit → mendeteksi saat ada submission baru dari Google
 *   Form (submission form TIDAK memicu onEdit, jadi butuh trigger
 *   terpisah supaya baris baru dari Form juga langsung kebaca app).
 *   Kalau Anda tidak pakai Google Form sama sekali, trigger ini tetap
 *   aman terpasang, hanya tidak akan pernah terpicu.
 *
 * CARA PASANG:
 * 1. Di spreadsheet ini: Extensions → Apps Script.
 * 2. Hapus kode default, tempel SELURUH isi file ini.
 * 3. Ganti FAMILY_CODE di bawah dengan kode rahasia bebas.
 * 4. Pilih fungsi "setupTrigger" di dropdown atas → klik ▶ Run →
 *    izinkan akses saat diminta. WAJIB dijalankan sekali.
 * 5. Deploy → New deployment → Web app → Execute as: Me,
 *    Who has access: Anyone → Deploy → salin URL .../exec.
 * 6. Tempel URL + FAMILY_CODE ke Pengaturan app Buku Kas Keluarga.
 *
 * BATASAN:
 * - Near-realtime (polling tiap 5 detik), bukan instan.
 * - Edit baris yang sama di Sheets & app dalam waktu sangat
 *   berdekatan (sebelum sempat sync): yang tersimpan terakhir yang
 *   menang, tidak digabung otomatis.
 * - Kolom I (ID) jangan diisi manual — biarkan kosong untuk baris
 *   baru, script otomatis mengisinya saat sinkron berikutnya.
 */

const FAMILY_CODE = 'GANTI-KODE-INI'; // <-- WAJIB DIGANTI sebelum deploy!
const SHEET_NAME = 'TRANSAKSI';       // <-- Ganti kalau nama tab Anda berbeda
const SETTINGS_SHEET = 'AppSettings';
const VERSION_KEY = 'dataVersion';

// Urutan kolom di sheet TRANSAKSI (1-indexed)
const COL = {
  TIMESTAMP: 1, TANGGAL: 2, JENIS: 3, REKENING: 4,
  KATEGORI: 5, NOMINAL: 6, LOKASI: 7, KETERANGAN: 8, ID: 9
};
const NUM_COLS = 9;

function _jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function _checkCode_(code) {
  return code && FAMILY_CODE !== 'GANTI-KODE-INI' && code === FAMILY_CODE;
}
function _getTxSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" tidak ditemukan. Cek nama tab persis di SHEET_NAME.');
  if (String(sheet.getRange(1, COL.ID).getValue()).trim() !== 'ID') {
    sheet.getRange(1, COL.ID).setValue('ID');
  }
  return sheet;
}
function _getSettingsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SETTINGS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS_SHEET);
    sheet.hideSheet();
    sheet.getRange(1, 1, 1, 2).setValues([['key', 'json']]);
    sheet.getRange(2, 1, 1, 2).setValues([['settings', '{}']]);
  }
  return sheet;
}
function _bumpVersion_() {
  const props = PropertiesService.getScriptProperties();
  const v = Number(props.getProperty(VERSION_KEY) || '0') + 1;
  props.setProperty(VERSION_KEY, String(v));
  return v;
}
function _getVersion_() {
  return Number(PropertiesService.getScriptProperties().getProperty(VERSION_KEY) || '0');
}

/** Trigger INSTALLABLE — dipasang lewat setupTrigger(), bukan simple trigger,
 *  supaya bisa akses PropertiesService dengan otorisasi penuh. */
function onEditInstallable(e) {
  try {
    const name = e.range.getSheet().getName();
    if (name === SHEET_NAME || name === SETTINGS_SHEET) _bumpVersion_();
  } catch (err) { /* abaikan supaya tidak mengganggu editing normal */ }
}
function onFormSubmitInstallable(e) {
  try { _bumpVersion_(); } catch (err) { /* abaikan */ }
}

/** Jalankan fungsi ini SEKALI secara manual dari editor Apps Script. */
function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === 'onEditInstallable' || fn === 'onFormSubmitInstallable') ScriptApp.deleteTrigger(t);
  });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('onEditInstallable').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('onFormSubmitInstallable').forSpreadsheet(ss).onFormSubmit().create();
  Logger.log('Trigger onEdit & onFormSubmit terpasang. Sinkronisasi 2 arah siap dipakai.');
}

// ── Konversi tanggal: sheet (Date object ATAU teks dd/mm/yyyy) <-> app (yyyy-mm-dd) ──
function _sheetDateToApp_(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone() || 'GMT+7', 'yyyy-MM-dd');
  }
  const s = String(val || '').trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); // dd/mm/yyyy
  if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
  const m2 = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/); // sudah yyyy-mm-dd
  if (m2) return s;
  return s;
}
function _appDateToSheet_(val) {
  const m = String(val || '').match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (!m) return val;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}
function _typeToApp_(v) { return String(v || '').trim().toLowerCase() === 'pemasukan' ? 'pemasukan' : 'pengeluaran'; }
function _typeToSheet_(v) { return v === 'pemasukan' ? 'Pemasukan' : 'Pengeluaran'; }

/**
 * GET ?code=XXXX&action=load
 */
function doGet(e) {
  const params = e.parameter || {};
  if (!_checkCode_(params.code)) return _jsonResponse_({ ok: false, error: 'Kode keluarga salah' });

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = _getTxSheet_();
    const lastRow = sheet.getLastRow();
    const data = [];
    if (lastRow >= 2) {
      const range = sheet.getRange(2, 1, lastRow - 1, NUM_COLS);
      const values = range.getValues();
      let idWasMissing = false;
      for (let i = 0; i < values.length; i++) {
        const row = values[i];
        if (!row[COL.TANGGAL - 1] && !row[COL.KATEGORI - 1] && !row[COL.NOMINAL - 1]) continue; // baris kosong
        if (!row[COL.ID - 1]) { row[COL.ID - 1] = Utilities.getUuid(); idWasMissing = true; }
        data.push({
          id: row[COL.ID - 1],
          date: _sheetDateToApp_(row[COL.TANGGAL - 1]),
          type: _typeToApp_(row[COL.JENIS - 1]),
          account: row[COL.REKENING - 1] || '',
          category: row[COL.KATEGORI - 1] || '',
          amount: Number(row[COL.NOMINAL - 1]) || 0,
          location: row[COL.LOKASI - 1] || '',
          note: row[COL.KETERANGAN - 1] || ''
        });
      }
      if (idWasMissing) range.setValues(values); // tulis-balik ID yang baru dibuat, sekali batch
    }
    const settingsSheet = _getSettingsSheet_();
    let settings = {};
    try { settings = JSON.parse(settingsSheet.getRange(2, 2).getValue() || '{}'); } catch (err) { settings = {}; }

    return _jsonResponse_({ ok: true, version: _getVersion_(), data: data, settings: settings });
  } finally {
    lock.releaseLock();
  }
}

/**
 * POST body JSON: { code, data: [...transaksi dari app], settings: {categories,budgets,accounts,customIcons} }
 */
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return _jsonResponse_({ ok: false, error: 'Payload tidak valid' }); }
  if (!_checkCode_(body.code)) return _jsonResponse_({ ok: false, error: 'Kode keluarga salah' });

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = _getTxSheet_();
    const lastRow = sheet.getLastRow();
    const existing = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, NUM_COLS).getValues() : [];
    const idToRow = {};
    existing.forEach((row, idx) => { if (row[COL.ID - 1]) idToRow[row[COL.ID - 1]] = idx; });

    const incoming = Array.isArray(body.data) ? body.data : [];
    const usedIds = {};
    const output = existing.map(r => r.slice()); // salinan, akan ditimpa/ditandai

    incoming.forEach(tx => {
      const id = tx.id || Utilities.getUuid();
      usedIds[id] = true;
      const isExisting = id in idToRow;
      const timestamp = isExisting ? output[idToRow[id]][COL.TIMESTAMP - 1] : new Date(); // pertahankan timestamp asli kalau update

      const rowArr = [];
      rowArr[COL.TIMESTAMP - 1] = timestamp;
      rowArr[COL.TANGGAL - 1] = _appDateToSheet_(tx.date);
      rowArr[COL.JENIS - 1] = _typeToSheet_(tx.type);
      rowArr[COL.REKENING - 1] = tx.account || '';
      rowArr[COL.KATEGORI - 1] = tx.category || '';
      rowArr[COL.NOMINAL - 1] = Number(tx.amount) || 0;
      rowArr[COL.LOKASI - 1] = tx.location || '';
      rowArr[COL.KETERANGAN - 1] = tx.note || '';
      rowArr[COL.ID - 1] = id;

      if (isExisting) output[idToRow[id]] = rowArr; // update baris yang sudah ada
      else output.push(rowArr); // transaksi baru dari app
    });

    // Baris lama yang ID-nya tidak ada lagi di data dari app = dihapus di app -> buang dari sheet.
    // (Baris yang belum punya ID sama sekali dilewati/dipertahankan supaya input manual/Form yang
    //  belum sempat di-load app tidak ikut kehapus.)
    const finalRows = output.filter(row => !row[COL.ID - 1] || usedIds[row[COL.ID - 1]] || !idToRow.hasOwnProperty(row[COL.ID - 1]));

    // Tulis ulang seluruh area data dalam satu batch (efisien untuk ribuan baris).
    const neededRows = finalRows.length;
    if (neededRows > 0) {
      sheet.getRange(2, 1, neededRows, NUM_COLS).setValues(finalRows);
    }
    const currentMaxRow = sheet.getLastRow();
    if (currentMaxRow > neededRows + 1) {
      sheet.getRange(neededRows + 2, 1, currentMaxRow - neededRows - 1, NUM_COLS).clearContent();
    }

    if (body.settings) {
      const settingsSheet = _getSettingsSheet_();
      settingsSheet.getRange(2, 2).setValue(JSON.stringify(body.settings));
    }

    const newVersion = _bumpVersion_();
    return _jsonResponse_({ ok: true, version: newVersion });
  } finally {
    lock.releaseLock();
  }
}
