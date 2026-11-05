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

// ===================== SINKRON METODE KERJA DARI DRIVE =====================
// Membaca Method Statement (.docx atau Google Docs) di folder CSA dan MEP, lalu mengirim ringkasannya ke aplikasi.
// Akun yang men-deploy script ini harus punya akses ke kedua folder. File PDF dilewati.
const MS_FOLDERS = { csa: '1MDh1TG8OtYdd-lqL7u9H5uq7ksjrV2Rr', mep: '1XbIGOO56k4eMpbtXphE0isRyztxczZq9' };

// ---- bagian murni (tidak memakai layanan Google), diuji di Node ----
function halfTitle_(t) { const i = String(t).indexOf(' / '); return (i > 0 ? t.slice(0, i) : t).replace(/\s+/g, ' ').trim(); }

// Paragraf dwibahasa "Indonesia / English": ambil bagian Indonesia.
function halfId_(t) {
  t = String(t).replace(/\s+/g, ' ').trim();
  const idx = []; let i = -1;
  while ((i = t.indexOf(' / ', i + 1)) >= 0) idx.push(i);
  if (!idx.length) return t;
  const mid = t.length / 2; let best = idx[0];
  idx.forEach(function (x) { if (Math.abs(x - mid) < Math.abs(best - mid)) best = x; });
  const a = t.slice(0, best), b = t.slice(best + 3);
  const r = a.length / Math.max(1, b.length);
  return (a.length >= 15 && b.length >= 15 && r > 0.4 && r < 2.5) ? a.trim() : t;
}

const MS_SECTIONS_ = [
  [/^TUJUAN/, 'purpose'], [/^LINGKUP/, 'scope'], [/^DOKUMEN ACUAN/, 'refs'], [/^MATERIAL/, 'mat'], [/^PERALATAN/, 'equip'],
  [/^SURVEY/, 'survey'], [/^PELAKSANAAN/, 'steps'], [/^PENGUJIAN/, 'tests'], [/^KRITERIA/, 'accept'],
  [/^PERSYARATAN KESELAMATAN/, 'safety'], [/^PENGENDALIAN LINGKUNGAN/, 'env'], [/^RINGKASAN/, 'seq']
];

function msFileInfo_(name) {
  const m = /^(S[A-Z]P-CKD-(CSA|MEP)-MS(?:-[A-Z]+)?-\d{2})(?:-R(\d+))?[\s_-]*(.*?)\.(docx|pdf)$/i.exec(name) || /^(S[A-Z]P-CKD-(CSA|MEP)-MS(?:-[A-Z]+)?-\d{2})(?:-R(\d+))?[\s_-]*(.*)$/i.exec(name);
  if (!m) return null;
  return { no: m[1].toUpperCase(), role: m[2].toLowerCase(), rev: m[3] || '0', title: (m[4] || '').replace(/_/g, ' ').trim() };
}

// paras: [{style:'Heading1'|'Heading2'|'', text}], rows: [[cell,...],...] dari tabel sampul
function parseMs_(paras, rows, file) {
  const info = msFileInfo_(file.name) || { no: file.name, role: file.role || '', rev: '0', title: file.name };
  const d = { id: info.no.toLowerCase(), no: info.no, role: info.role, title: info.title, rev: 'R' + info.rev, status: '', mod: file.updated || 0,
    purpose: '', scope: '', refs: [], equip: [], mat: [], steps: [], tests: [], accept: [], safety: [], env: [], seq: '', jsa: [] };
  (rows || []).forEach(function (r) {
    for (let i = 0; i < r.length - 1; i++) if (/STATUS DOKUMEN/i.test(r[i])) d.status = String(r[i + 1]).trim();
  });
  let sec = '', grp = null, survey = [];
  paras.forEach(function (p) {
    const raw = String(p.text || '').replace(/\s+/g, ' ').trim();
    if (!raw || raw.length < 3) return;
    const st = String(p.style || '');
    if (/^Heading\s?1$/i.test(st)) {
      const up = raw.toUpperCase(); sec = ''; grp = null;
      MS_SECTIONS_.forEach(function (s) { if (s[0].test(up)) sec = s[1]; });
      return;
    }
    if (/^Heading\s?2$/i.test(st)) { if (sec === 'steps') { grp = { t: halfTitle_(raw), d: [] }; d.steps.push(grp); } return; }
    if (/^(TOC|Contents|Daftar)/i.test(st) || !sec) return;
    const t = halfId_(raw);
    if (sec === 'purpose') d.purpose = d.purpose ? d.purpose + ' ' + t : t;
    else if (sec === 'scope') d.scope = d.scope ? d.scope + ' ' + t : t;
    else if (sec === 'seq') d.seq = d.seq ? d.seq + ' ' + raw : raw;
    else if (sec === 'survey') survey.push(t);
    else if (sec === 'steps') { if (!grp) { grp = { t: 'Pelaksanaan', d: [] }; d.steps.push(grp); } grp.d.push(t); }
    else if (d[sec]) d[sec].push(t);
  });
  if (survey.length) d.steps.unshift({ t: 'Survey pra-konstruksi', d: survey });
  return d;
}

// ---- pembaca file (layanan Google, tidak bisa diuji di luar Apps Script) ----
function wText_(el) {
  let s = '';
  el.getChildren().forEach(function (c) {
    const n = c.getName();
    if (n === 't') s += c.getText();
    else if (n === 'tab' || n === 'br') s += ' ';
    else if (n === 'delText' || n === 'instrText' || n === 'drawing' || n === 'pict' || n === 'AlternateContent') return;
    else s += wText_(c);
  });
  return s;
}

function docxRead_(file) {
  const blob = file.getBlob().setContentType('application/zip');
  const part = Utilities.unzip(blob).filter(function (p) { return p.getName() === 'word/document.xml'; })[0];
  const W = XmlService.getNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main');
  const body = XmlService.parse(part.getDataAsString('UTF-8')).getRootElement().getChild('body', W);
  const paras = [], rows = [];
  body.getChildren().forEach(function (el) {
    const n = el.getName();
    if (n === 'p') {
      const ppr = el.getChild('pPr', W), ps = ppr && ppr.getChild('pStyle', W);
      const a = ps && ps.getAttribute('val', W);
      paras.push({ style: a ? a.getValue() : '', text: wText_(el) });
    } else if (n === 'tbl' && rows.length < 40) {
      el.getChildren('tr', W).forEach(function (tr) { rows.push(tr.getChildren('tc', W).map(wText_)); });
    }
  });
  return { paras: paras, rows: rows };
}

function gdocRead_(id) {
  const body = DocumentApp.openById(id).getBody(), H = DocumentApp.ParagraphHeading, paras = [], rows = [];
  body.getParagraphs().forEach(function (p) {
    const h = p.getHeading();
    paras.push({ style: h === H.HEADING1 ? 'Heading1' : h === H.HEADING2 ? 'Heading2' : '', text: p.getText() });
  });
  body.getTables().slice(0, 3).forEach(function (t) {
    for (let r = 0; r < t.getNumRows(); r++) { const row = t.getRow(r), c = []; for (let k = 0; k < row.getNumCells(); k++) c.push(row.getCell(k).getText()); rows.push(c); }
  });
  return { paras: paras, rows: rows };
}

function msList_(folderId, out) {
  const f = DriveApp.getFolderById(folderId), fs = f.getFiles();
  while (fs.hasNext()) { const x = fs.next(); out.push({ id: x.getId(), name: x.getName(), mime: x.getMimeType(), updated: x.getLastUpdated().getTime() }); }
  const ds = f.getFolders();
  while (ds.hasNext()) msList_(ds.next().getId(), out);
}

function syncMs_() {
  const cache = CacheService.getScriptCache(), best = {}, skipped = [];
  const rank = { docx: 3, gdoc: 2, pdf: 1 };
  Object.keys(MS_FOLDERS).forEach(function (role) {
    const files = []; msList_(MS_FOLDERS[role], files);
    files.forEach(function (f) {
      const kind = f.mime === 'application/vnd.google-apps.document' ? 'gdoc' : /wordprocessingml/.test(f.mime) ? 'docx' : f.mime === 'application/pdf' ? 'pdf' : '';
      if (!kind) return;
      const info = msFileInfo_(f.name);
      if (!info) { skipped.push({ name: f.name, reason: 'Nama file tidak sesuai pola nomor MS' }); return; }
      const k = info.no, cur = best[k];
      if (!cur || rank[kind] > rank[cur.kind] || (rank[kind] === rank[cur.kind] && f.updated > cur.f.updated)) best[k] = { f: f, kind: kind };
    });
  });
  const docs = [];
  Object.keys(best).sort().forEach(function (k) {
    const f = best[k].f, kind = best[k].kind;
    if (kind === 'pdf') { skipped.push({ name: f.name, reason: 'Hanya tersedia PDF, belum bisa dibaca otomatis' }); return; }
    const ck = 'ms3:' + f.id + ':' + f.updated;
    let d = null; try { const c = cache.get(ck); if (c) d = JSON.parse(c); } catch (e) {}
    if (!d) {
      try {
        const r = kind === 'docx' ? docxRead_(DriveApp.getFileById(f.id)) : gdocRead_(f.id);
        d = parseMs_(r.paras, r.rows, f);
        const s = JSON.stringify(d); if (s.length < 90000) cache.put(ck, s, 21600);
      } catch (e) { skipped.push({ name: f.name, reason: 'Gagal dibaca: ' + e.message }); return; }
    }
    if (d.steps.length) docs.push(d); else skipped.push({ name: f.name, reason: 'Struktur heading tidak dikenali (pakai gaya Heading 1 dan Heading 2)' });
  });
  return { ok: true, docs: docs, skipped: skipped, at: new Date().toISOString() };
}
// =================== AKHIR SINKRON METODE KERJA ===================

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return out_({ ok: false, error: 'Format data tidak valid' }); }
  if (body.token !== TOKEN) return out_({ ok: false, error: 'Kode rahasia salah' });
  if (body.action === 'ms') {
    try { return out_(syncMs_()); }
    catch (err) { return out_({ ok: false, error: 'Sinkron metode gagal: ' + err.message }); }
  }

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
