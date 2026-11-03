/**
 * Rekap Supervisi Konstruksi
 * ------------------------------------------------------------------
 * Membaca semua file Excel/Google Sheet di satu folder Drive, mengambil
 * sheet "List_*" (tabel datar) dari tiap file, lalu merekapnya ke sheet
 * "Rekap" beserta keterangan KAPAN data diambil.
 *
 * Cara pakai: lihat README.md. Jalankan `updateRekap` (atau menu "Rekap").
 */

const CONFIG = {
  FOLDER_ID: '1IyXIvkUK5uMIuO7bx4lrb9aGeBTgYBhN',
  SHEET_REKAP: 'Rekap',
  SHEET_LOG: 'Log',
  LIST_SHEET_PREFIX: 'List_',          // sheet sumber: List_SKTT, List_SUTT, ...
  TZ: 'Asia/Jakarta',
  TIME_BUDGET_MS: 4.5 * 60 * 1000,     // batas Apps Script 6 menit; sisakan margin
  CONTINUE_AFTER_MS: 60 * 1000,        // jeda sebelum lanjut otomatis
  DAILY_HOUR: 19,                      // jam jadwal harian (zona waktu script)
};

// Kolom rekap sesuai permintaan (urutan dipertahankan).
const COLUMNS = [
  'Kode Lokasi', 'Status', 'Periode', 'No Surat Penugasan', 'Jangka Waktu Penugasan',
  'No Amandemen Penugasan', 'UPMK', 'PSO', 'Pemberi Tugas', 'Direksi Pekerjaan',
  'Nama Proyek', 'No Kontrak', 'Kontraktor', 'Nilai Kontrak Original', 'Nilai Kontrak Akhir',
  'No. Amandemen Nilai Kontrak Akhir', 'Sumber Dana', 'Tanggal Kontrak', 'Efektif Kontrak',
  'COD Kontrak', 'COD Amandemen', 'No. Amd Waktu Kontrak', 'Kapasitas', 'Satuan Kapasitas',
  'Detail', 'Masa Berlaku Bank Garansi', 'Lokasi', 'Progres Aktual (%)', 'Progres Plan(%)',
  'Aktivitas', 'RLB', 'Energize', 'SLO', 'BAST 1', 'BAST 2',
  'Kategori1', 'Detail1', 'Tindaklanjut1', 'Kategori2', 'Detail2', 'Tindaklanjut2',
  'Kategori3', 'Detail3', 'Tindaklanjut3', 'Kategori4', 'Detail4', 'Tindaklanjut4',
  'Kategori5', 'Detail5', 'Tindaklanjut5',
  'Jenis', 'Milestone', 'to decimal actual', 'to decimal plan', 'Deviasi', 'Bullet',
];

// Kolom keterangan pengambilan data (ditambahkan di kanan).
const META = ['Diambil Pada', 'Nama File Sumber', 'File Diubah Terakhir', 'ID File', 'Catatan'];
const ALL_COLUMNS = COLUMNS.concat(META);

// Nama header di file sumber yang berbeda dengan nama kolom rekap.
const ALIASES = { 'pso': ['pst'] };

/* ============================ ENTRY POINTS ============================ */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Rekap')
    .addItem('Ambil data sekarang (hanya file yang berubah)', 'updateRekap')
    .addItem('Ambil ulang SEMUA file', 'updateRekapSemua')
    .addSeparator()
    .addItem('Pasang jadwal harian', 'pasangJadwalHarian')
    .addItem('Hapus jadwal', 'hapusJadwal')
    .addToUi();
}

function updateRekap() { run_(false); }
function updateRekapSemua() { run_(true); }

/* ================================ MAIN ================================ */

function run_(forceAll) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10 * 1000)) { toast_('Proses lain sedang berjalan.'); return; }
  const started = Date.now();
  const stats = { processed: 0, skipped: 0, errors: 0, pending: 0 };
  try {
    clearContinuation_();
    const ss = SpreadsheetApp.getActive();
    const rekap = ensureRekapSheet_(ss);
    const log = ensureLogSheet_(ss);

    const files = listSourceFiles_();
    const existing = readRekap_(rekap);               // {fileId: {modified, rows[]}}
    const takenAt = new Date();
    const byFile = {};
    Object.keys(existing).forEach(id => { byFile[id] = existing[id]; });

    // Hapus data file yang sudah tidak ada di folder.
    const liveIds = new Set(files.map(f => f.getId()));
    Object.keys(byFile).forEach(id => { if (!liveIds.has(id)) delete byFile[id]; });

    for (const file of files) {
      const id = file.getId();
      const modified = file.getLastUpdated();
      const prev = byFile[id];
      if (!forceAll && prev && prev.modified && Math.abs(prev.modified - modified.getTime()) < 1000) {
        stats.skipped++;
        continue;
      }
      if (Date.now() - started > CONFIG.TIME_BUDGET_MS) { stats.pending++; continue; }
      try {
        const rows = extractRows_(file, takenAt);
        byFile[id] = { modified: modified.getTime(), rows: rows };
        stats.processed++;
      } catch (e) {
        stats.errors++;
        appendLog_(log, takenAt, file.getName(), 'ERROR', String(e && e.message || e));
        // Pertahankan data lama bila ada; kalau tidak, catat baris kosong ber-catatan.
        if (!byFile[id]) byFile[id] = { modified: 0, rows: [blankRow_(file, takenAt, 'GAGAL DIBACA: ' + (e && e.message || e))] };
      }
    }

    writeRekap_(rekap, byFile);
    appendLog_(log, new Date(), '(ringkasan)', stats.pending ? 'SEBAGIAN' : 'OK',
      'diproses=' + stats.processed + ', dilewati(tidak berubah)=' + stats.skipped +
      ', error=' + stats.errors + ', tersisa=' + stats.pending);

    if (stats.pending > 0) scheduleContinuation_();
    toast_('Selesai. Diproses ' + stats.processed + ', dilewati ' + stats.skipped +
      ', error ' + stats.errors + (stats.pending ? ', tersisa ' + stats.pending + ' (lanjut otomatis)' : ''));
  } finally {
    lock.releaseLock();
  }
}

/* ============================ SOURCE FILES ============================ */

function listSourceFiles_() {
  const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
  const out = [];
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    const mt = f.getMimeType();
    if (mt === MimeType.GOOGLE_SHEETS || mt === MimeType.MICROSOFT_EXCEL) out.push(f);
  }
  out.sort((a, b) => a.getName().localeCompare(b.getName()));
  return out;
}

/** Buka file (konversi sementara jika .xlsx) dan ambil baris dari sheet List_*. */
function extractRows_(file, takenAt) {
  let tmpId = null, ss;
  try {
    if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
      ss = SpreadsheetApp.openById(file.getId());
    } else {
      const created = Drive.Files.create(
        { name: 'tmp_rekap_' + file.getId(), mimeType: MimeType.GOOGLE_SHEETS },
        file.getBlob());
      tmpId = created.id;
      ss = SpreadsheetApp.openById(tmpId);
    }
    const sheets = ss.getSheets().filter(s => s.getName().toLowerCase().indexOf(CONFIG.LIST_SHEET_PREFIX.toLowerCase()) === 0);
    if (!sheets.length) throw new Error('Sheet "' + CONFIG.LIST_SHEET_PREFIX + '*" tidak ditemukan');
    const sh = sheets[0];
    const jenis = sh.getName().substring(CONFIG.LIST_SHEET_PREFIX.length).replace(/^[\\_\s]+/, '');
    const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
    if (lastRow < 2) throw new Error('Sheet ' + sh.getName() + ' tidak punya baris data');

    const range = sh.getRange(1, 1, lastRow, lastCol);
    const values = range.getValues();
    const formats = range.getNumberFormats();
    const header = values[0].map(norm_);
    const index = {};
    COLUMNS.forEach(col => {
      const keys = [norm_(col)].concat(ALIASES[norm_(col)] || []);
      for (const k of keys) { const i = header.indexOf(k); if (i >= 0) { index[col] = i; break; } }
    });

    const rows = [];
    for (let r = 1; r < values.length; r++) {
      const src = values[r];
      const keyCols = ['Kode Lokasi', 'Nama Proyek', 'No Kontrak'];
      if (!keyCols.some(c => index[c] !== undefined && String(src[index[c]]).trim() !== '')) continue;

      const notes = [];
      let errCells = 0;
      const row = COLUMNS.map(col => {
        const i = index[col];
        if (i === undefined) return '';
        const v = src[i];
        if (typeof v === 'string' && /^#(NAME|REF|VALUE|N\/A|DIV\/0|NUM|NULL)/.test(v)) { errCells++; return ''; }
        return v;
      });
      if (errCells) notes.push(errCells + ' sel berisi error rumus (dikosongkan)');

      // Kolom turunan
      const set = (col, val) => { row[COLUMNS.indexOf(col)] = val; };
      if (index['Jenis'] === undefined || row[COLUMNS.indexOf('Jenis')] === '') set('Jenis', jenis);
      const act = toDecimal_(src[index['Progres Aktual (%)']], (formats[r] || [])[index['Progres Aktual (%)']]);
      const plan = toDecimal_(src[index['Progres Plan(%)']], (formats[r] || [])[index['Progres Plan(%)']]);
      set('to decimal actual', act);
      set('to decimal plan', plan);
      set('Deviasi', (act === '' || plan === '') ? '' : act - plan);
      if (index['Milestone'] === undefined) notes.push('Milestone tidak ada di sumber');
      if (index['Bullet'] === undefined) notes.push('Bullet tidak ada di sumber');
      const missing = COLUMNS.filter(c => index[c] === undefined &&
        ['Jenis', 'Milestone', 'Bullet', 'to decimal actual', 'to decimal plan', 'Deviasi'].indexOf(c) < 0);
      if (missing.length) notes.push('Header tidak ditemukan: ' + missing.join(', '));

      rows.push(row.concat([takenAt, file.getName(), file.getLastUpdated(), file.getId(), notes.join('; ')]));
    }
    if (!rows.length) throw new Error('Tidak ada baris data terisi di ' + sh.getName());
    return rows;
  } finally {
    if (tmpId) { try { Drive.Files.remove(tmpId); } catch (e) { /* abaikan */ } }
  }
}

function blankRow_(file, takenAt, note) {
  const row = new Array(COLUMNS.length).fill('');
  return row.concat([takenAt, file.getName(), file.getLastUpdated(), file.getId(), note]);
}

/* ============================== HELPERS =============================== */

function norm_(s) { return String(s).toLowerCase().replace(/[^a-z0-9]/g, ''); }

/** Ubah nilai progres ke desimal (0-1). Sel berformat % sudah berupa pecahan. */
function toDecimal_(v, fmt) {
  if (v === '' || v === null || v === undefined) return '';
  if (typeof v === 'string') {
    const s = v.trim();
    if (/^#/.test(s) || s === '-' || s === '') return '';
    const n = parseFloat(s.replace('%', '').replace(',', '.'));
    if (isNaN(n)) return '';
    return n / 100;   // teks "0,732%" atau "0.732" dianggap persen
  }
  if (typeof v !== 'number') return '';
  return (fmt && String(fmt).indexOf('%') >= 0) ? v : v / 100;
}

/* ============================ REKAP SHEET ============================= */

function ensureRekapSheet_(ss) {
  let sh = ss.getSheetByName(CONFIG.SHEET_REKAP);
  if (!sh) {
    const sheets = ss.getSheets();
    if (sheets.length === 1 && sheets[0].getLastRow() <= 1) { sh = sheets[0]; sh.setName(CONFIG.SHEET_REKAP); }
    else sh = ss.insertSheet(CONFIG.SHEET_REKAP);
  }
  return sh;
}

function ensureLogSheet_(ss) {
  let sh = ss.getSheetByName(CONFIG.SHEET_LOG);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.SHEET_LOG);
    sh.appendRow(['Waktu', 'File', 'Status', 'Keterangan']);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 4).setFontWeight('bold');
  }
  return sh;
}

function appendLog_(log, when, file, status, msg) {
  log.appendRow([when, file, status, msg]);
  log.getRange(log.getLastRow(), 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
}

/** Baca Rekap saat ini, dikelompokkan per ID File (untuk melewati file yang tidak berubah). */
function readRekap_(sh) {
  const out = {};
  const lastRow = sh.getLastRow();
  if (lastRow < 2 || sh.getLastColumn() < ALL_COLUMNS.length) return out;
  const vals = sh.getRange(2, 1, lastRow - 1, ALL_COLUMNS.length).getValues();
  const iId = ALL_COLUMNS.indexOf('ID File'), iMod = ALL_COLUMNS.indexOf('File Diubah Terakhir');
  vals.forEach(r => {
    const id = r[iId];
    if (!id) return;
    const t = r[iMod] instanceof Date ? r[iMod].getTime() : 0;
    if (!out[id]) out[id] = { modified: t, rows: [] };
    out[id].rows.push(r);
  });
  return out;
}

function writeRekap_(sh, byFile) {
  let rows = [];
  Object.keys(byFile).forEach(id => { rows = rows.concat(byFile[id].rows); });
  const iNama = ALL_COLUMNS.indexOf('Nama File Sumber');
  rows.sort((a, b) => String(a[iNama]).localeCompare(String(b[iNama])));

  sh.clearContents();
  sh.getRange(1, 1, 1, ALL_COLUMNS.length).setValues([ALL_COLUMNS])
    .setFontWeight('bold').setBackground('#1f3b63').setFontColor('#ffffff').setWrap(true);
  sh.setFrozenRows(1);
  if (rows.length) {
    sh.getRange(2, 1, rows.length, ALL_COLUMNS.length).setValues(rows);
    const ts = ['Diambil Pada', 'File Diubah Terakhir'].map(c => ALL_COLUMNS.indexOf(c) + 1);
    ts.forEach(c => sh.getRange(2, c, rows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss'));
    ['to decimal actual', 'to decimal plan', 'Deviasi'].forEach(c =>
      sh.getRange(2, ALL_COLUMNS.indexOf(c) + 1, rows.length, 1).setNumberFormat('0.000%'));
  }
  // Keterangan global: kapan rekap terakhir dijalankan.
  sh.getRange(1, ALL_COLUMNS.length + 2).setValue('Rekap terakhir dijalankan:')
    .setFontWeight('bold');
  sh.getRange(2, ALL_COLUMNS.length + 2).setValue(new Date()).setNumberFormat('yyyy-mm-dd hh:mm:ss');
}

/* ============================== TRIGGERS ============================== */

function pasangJadwalHarian() {
  hapusJadwal();
  ScriptApp.newTrigger('updateRekap').timeBased().everyDays(1).atHour(CONFIG.DAILY_HOUR).create();
  toast_('Jadwal harian dipasang (sekitar jam ' + CONFIG.DAILY_HOUR + ').');
}

function hapusJadwal() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'updateRekap') ScriptApp.deleteTrigger(t);
  });
}

function scheduleContinuation_() {
  const t = ScriptApp.newTrigger('updateRekap').timeBased().after(CONFIG.CONTINUE_AFTER_MS).create();
  PropertiesService.getScriptProperties().setProperty('CONTINUATION_ID', t.getUniqueId());
}

function clearContinuation_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('CONTINUATION_ID');
  if (!id) return;
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getUniqueId() === id) ScriptApp.deleteTrigger(t); });
  props.deleteProperty('CONTINUATION_ID');
}

function toast_(msg) {
  try { SpreadsheetApp.getActive().toast(msg, 'Rekap', 10); } catch (e) { /* tanpa UI */ }
  console.log(msg);
}
