/**
 * SiteGuide CKD - penerima data (Google Apps Script)
 * Tempel di Extensions > Apps Script pada Google Sheet, ganti TOKEN, lalu Deploy sebagai Web app.
 * Tab Log, NCR, dan Rekap dibuat otomatis saat data pertama masuk.
 */
const TOKEN = 'GANTI-DENGAN-KODE-RAHASIA';

const SHEETS = {
  Log:   ['Waktu', 'ID Event', 'Jenis', 'Proyek', 'Kontraktor', 'Pengguna', 'Peran', 'Area', 'Detail'],
  NCR:   ['ID NCR', 'No NCR', 'Diperbarui', 'Proyek', 'Pelapor', 'Peran', 'Lokasi', 'Pekerjaan', 'Tingkat', 'Uraian', 'Tindakan korektif', 'Batas waktu', 'Status'],
  Rekap: ['Waktu', 'Proyek', 'Kontraktor', 'Pengguna', 'Peran', 'Area', 'Laporan']
};

function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(SHEETS[name]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, SHEETS[name].length).setFontWeight('bold');
  }
  return sh;
}

// Cegah teks yang diawali = + - @ dibaca Sheets sebagai rumus.
function safe_(v) {
  if (typeof v === 'string' && /^[=+\-@]/.test(v)) return "'" + v;
  return v === undefined || v === null ? '' : v;
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return out_({ ok: true, service: 'SiteGuide CKD', time: new Date().toISOString() });
}

function detail_(ev) {
  if (ev.type === 'check') return safe_(ev.jobName + ' | ' + ev.step + ' | ' + (ev.checked ? 'SELESAI' : 'BATAL'));
  if (ev.type === 'ncr') return safe_((ev.no || '') + ' | ' + (ev.status || ''));
  return safe_(ev.type);
}

function upsertNcr_(ev) {
  const sh = sheet_('NCR');
  const row = [ev.ncrId, ev.no, new Date(ev.t), ev.project, ev.by || ev.user, ev.role, ev.lokasi,
               ev.jobName, ev.level, ev.desc, ev.act, ev.due, ev.status].map(safe_);
  const rows = Math.max(1, sh.getLastRow() - 1);
  const hit = sh.getRange(2, 1, rows, 1).createTextFinder(String(ev.ncrId)).matchEntireCell(true).findNext();
  if (hit) sh.getRange(hit.getRow(), 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return out_({ ok: false, error: 'Format data tidak valid' }); }
  if (body.token !== TOKEN) return out_({ ok: false, error: 'Kode rahasia salah' });

  const events = Array.isArray(body.events) ? body.events.slice(0, 100) : [];
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const log = sheet_('Log');
    const last = log.getLastRow();
    const start = Math.max(2, last - 999);
    const seen = new Set(last >= 2 ? log.getRange(start, 2, last - start + 1, 1).getValues().flat() : []);
    let added = 0;
    events.forEach(function (ev) {
      if (!ev || !ev.id || seen.has(ev.id)) return;   // kiriman ulang diabaikan
      seen.add(ev.id);
      log.appendRow([new Date(ev.t), ev.id, ev.type, ev.project, ev.contractor, ev.user, ev.role, ev.area, detail_(ev)].map(safe_));
      if (ev.type === 'ncr' && ev.ncrId) upsertNcr_(ev);
      if (ev.type === 'rekap') sheet_('Rekap').appendRow([new Date(ev.t), ev.project, ev.contractor, ev.user, ev.role, ev.area, ev.text].map(safe_));
      added++;
    });
    return out_({ ok: true, n: added });
  } finally {
    lock.releaseLock();
  }
}
