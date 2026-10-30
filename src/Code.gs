const SS_ID = '1itdPdb6NATZawFNjsNPPOxBD-HdTpA8heT8hmP2hOc4';
const COMPANY = 'PT Aksara Perdana Multimedia';
const TTL = 21600; // sesi login 6 jam
const SC = 120;    // cache data sheet; dipanaskan ulang tiap 1 menit oleh trigger warm()
let _ss; const ss_ = () => _ss || (_ss = SpreadsheetApp.openById(SS_ID));

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Slip Gaji - ' + COMPANY)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// API JSON untuk GitHub Pages (POST text/plain agar tanpa preflight CORS)
const API = { pub, login, getData, getImages, getSigs, getKpi, getRef, getJadwal, getAttend, changePassword, regDevice, loginDevice, unregDevice };
function doPost(e) {
  let out;
  try {
    const q = JSON.parse(e.postData.contents);
    if (!API.hasOwnProperty(q.fn)) throw new Error('Fungsi tidak dikenal');
    out = { data: API[q.fn].apply(null, q.args || []) };
  } catch (err) { out = { error: String(err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_(name) {
  return ss_().getSheets().find(s => s.getName().toLowerCase() === name.toLowerCase());
}

// Baca sheet apa adanya (header dicari otomatis) + cache singkat agar cepat
function tbl_(name, key, force) {
  const c = CacheService.getScriptCache(), ck = 't_' + name;
  const hit = !force && c.get(ck);
  if (hit) return JSON.parse(hit);
  const v = sheet_(name).getDataRange().getDisplayValues();
  const h = v.findIndex(r => r.some(x => x.trim() === key));
  const head = v[h].map(x => x.trim());
  const rows = v.slice(h + 1).map((r, i) => ({ r, row: h + 2 + i }))
    .filter(x => x.r.some(y => y !== ''))
    .map(x => { const o = { _row: x.row }; head.forEach((k, j) => o[k] = x.r[j]); return o; });
  const out = { head, rows };
  try { c.put(ck, JSON.stringify(out), SC); } catch (e) {}
  return out;
}

// Gambar Drive -> data URI kecil (di-cache, tetap jalan walau file Drive tidak publik)
function img_(url, sz) {
  const m = String(url || '').match(/[-\w]{25,}/);
  if (!m) return '';
  const c = CacheService.getScriptCache(), k = 'i_' + sz + m[0];
  let d = c.get(k);
  if (d) return d;
  try {
    const r = UrlFetchApp.fetch('https://drive.google.com/thumbnail?id=' + m[0] + '&sz=w' + sz,
      { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) return '';
    const b = r.getBlob();
    d = 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes());
    if (d.length < 95000) c.put(k, d, TTL);
    return d;
  } catch (e) { return ''; }
}

function logo_() {
  const r = tbl_('Slip_Gaji', 'Periode').rows.find(x => x['Logo Perusahaan']);
  return r ? img_(r['Logo Perusahaan'], 300) : '';
}

function pub() { return { company: COMPANY, logo: logo_() }; }

function who_(t) {
  const a = CacheService.getScriptCache().get('s_' + t);
  const u = a && tbl_('Users', 'Akun').rows.find(r => r['Akun'] === a);
  if (!u) throw new Error('SESI');
  return u;
}

const BLN = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
// Slip periode "Oktober 2026" baru dirilis tanggal 3 bulan berikutnya (3 November 2026), 00:00 WIB
function rel_(p) {
  const m = String(p).trim().toLowerCase().match(/^([a-z]+)\s+(\d{4})$/);
  const i = m ? BLN.findIndex(b => b.toLowerCase() === m[1]) : -1;
  return i < 0 ? null : new Date(+m[2], i + 1, 3);
}

// Data ringan (tanpa gambar): semua slip milik user (SPV: semua karyawan); filter periode dilakukan di layar
function build_(u) {
  const spv = u['Jabatan'].toUpperCase() === 'SPV';
  const s = tbl_('Slip_Gaji', 'Periode').rows, emp = tbl_('Data_Karyawan', 'NIK').rows;
  const now = new Date();
  // Hanya SPV yang melihat semua; karyawan tidak menerima data periode yang belum dirilis (dicegah di server)
  const periods = [...new Set(s.map(r => r['Periode']))].map(p => {
    const d = rel_(p);
    return { p, locked: !spv && !!d && now < d, rel: d ? Utilities.formatDate(d, 'Asia/Jakarta', 'yyyy-MM-dd') : '' };
  });
  const lockedSet = new Set(periods.filter(x => x.locked).map(x => x.p));
  const slips = s.filter(r => (spv || r['NIP'] === u['NIK']) && !lockedSet.has(r['Periode'])).map(r => {
    const e = emp.find(x => x['NIK'] === r['NIP']) || {};
    const o = Object.assign({}, r, { Nama: e['Nama Karyawan'] || r['NIP'], Jabatan: e['Jabatan'] || '', Departemen: e['Departemen'] || '', Bank: e['Bank'] || '', Rek: e['No Rekening'] || '' });
    delete o['Logo Perusahaan']; delete o['TTD']; delete o._row; return o;
  });
  const me = emp.find(x => x['NIK'] === u['NIK']) || {};
  return {
    spv, slips, periods,
    me: { nik: u['NIK'], nama: u['Nama Karyawan'], jabatan: u['Jabatan'], dept: u['Departemen'],
      hasFoto: !!me['Foto Karyawan'], namaRek: me['Nama Rekening'] || u['Nama Karyawan'], bank: me['Bank'] || '', rek: me['No Rekening'] || '',
      pwLog: JSON.parse(PropertiesService.getScriptProperties().getProperty('P_' + u['Akun']) || '[]') }
  };
}

function session_(u) {
  const t = Utilities.getUuid();
  CacheService.getScriptCache().put('s_' + t, u['Akun'], TTL);
  return { ok: true, token: t, data: build_(u) };
}

function login(akun, pw) {
  const u = tbl_('Users', 'Akun').rows.find(r => r['Akun'].trim() === String(akun).trim());
  if (!u || u['Password'] !== String(pw)) { Utilities.sleep(400); return { ok: false, code: 'bad' }; }
  return session_(u);
}

// Login sidik jari: HP memverifikasi sidik jari, server memverifikasi token perangkat (maks. 3 HP per akun)
const hash_ = s => Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s));
function regDevice(t) {
  const u = who_(t), P = PropertiesService.getScriptProperties(), k = 'D_' + u['Akun'];
  const dt = u['Akun'] + '.' + Utilities.getUuid() + Utilities.getUuid();
  const a = JSON.parse(P.getProperty(k) || '[]'); a.push(hash_(dt));
  P.setProperty(k, JSON.stringify(a.slice(-3)));
  return dt;
}
function loginDevice(dt) {
  dt = String(dt); const ak = dt.slice(0, dt.lastIndexOf('.'));
  const a = JSON.parse(PropertiesService.getScriptProperties().getProperty('D_' + ak) || '[]');
  const u = tbl_('Users', 'Akun').rows.find(r => r['Akun'] === ak);
  if (!u || !a.includes(hash_(dt))) { Utilities.sleep(400); return { ok: false, code: 'dev' }; }
  return session_(u);
}
function unregDevice(dt) {
  dt = String(dt); const ak = dt.slice(0, dt.lastIndexOf('.')), P = PropertiesService.getScriptProperties();
  const a = JSON.parse(P.getProperty('D_' + ak) || '[]').filter(x => x !== hash_(dt));
  a.length ? P.setProperty('D_' + ak, JSON.stringify(a)) : P.deleteProperty('D_' + ak);
  return true;
}

// Pemanasan cache tiap 1 menit: login selalu cepat & data selalu segar. Jalankan setupTrigger() SEKALI dari editor.
function warm() { tbl_('Users', 'Akun', true); tbl_('Slip_Gaji', 'Periode', true); tbl_('Data_Karyawan', 'NIK', true); logo_(); try { kpiData_(true); } catch (e) {} try { attMap_(true); } catch (e) {} try { CacheService.getScriptCache().remove('jdw'); jadwalData_(); } catch (e) {} }
function setupTrigger() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'warm').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('warm').timeBased().everyMinutes(1).create();
}

function getData(t) { return build_(who_(t)); }

// Tanda tangan manager untuk PDF: link di sheet Slip_Gaji sel M2 (kiri) dan M3 (kanan); file asli diambil tanpa kompresi
function cellUrl_(cell) {
  const rt = cell.getRichTextValue();
  let u = rt && rt.getLinkUrl();
  if (!u && rt) for (const r of rt.getRuns()) if (r.getLinkUrl()) { u = r.getLinkUrl(); break; }
  if (!u) { const f = cell.getFormula(), m = f && f.match(/"(https?:[^"]+)"/); u = m ? m[1] : cell.getDisplayValue(); }
  return u;
}
function sigImg_(url) {
  const m = String(url || '').match(/[-\w]{25,}/);
  if (!m) return '';
  const c = CacheService.getScriptCache(), k = 'g_' + m[0];
  let d = c.get(k);
  if (d) return d;
  try {
    const r = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + m[0] + '?alt=media',
      { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    if (r.getResponseCode() === 200) {
      const b = r.getBlob();
      d = 'data:' + (b.getContentType() || 'image/png').split(';')[0] + ';base64,' + Utilities.base64Encode(b.getBytes());
      if (d.length < 95000) c.put(k, d, TTL);
      return d;
    }
  } catch (e) {}
  return img_(url, 900);
}
function getSigs(t) {
  who_(t);
  const sh = sheet_('Slip_Gaji');
  return { a: sigImg_(cellUrl_(sh.getRange('M2'))), b: sigImg_(cellUrl_(sh.getRange('M3'))) };
}

// ===== KPI & Referensi Nilai Poin (spreadsheet "Format KPI Teknisi APM Depok") =====
const KPI_ID = '1czzXRndFHBYehpti8TvlsPdtRizGE77LmFEs0L6rNOg';
function ksheet_(n) { return SpreadsheetApp.openById(KPI_ID).getSheets().find(s => s.getName().toLowerCase() === n.toLowerCase()); }
function num_(v) {
  if (typeof v === 'number') return v;
  let s = String(v).replace(/[^\d.,-]/g, '');
  if (!s) return 0;
  if (/^-?\d{1,3}(\.\d{3})+(,\d+)?$/.test(s)) s = s.replace(/\./g, '').replace(',', '.'); else s = s.replace(/,/g, '');
  return Number(s) || 0;
}
// Total Pemasangan Individu per teknisi per bulan (sheet "Laporan Harian", kolom "Total Pemasangan Individu")
function kpiData_(force) {
  const c = CacheService.getScriptCache(), hit = !force && c.get('kpi');
  if (hit) return JSON.parse(hit);
  const v = ksheet_('Laporan Harian').getDataRange().getDisplayValues();
  let year = new Date().getFullYear(), target = 66, month = '', col = -1;
  const months = [], rows = {};
  v.forEach(r => {
    const a = String(r[0]).trim();
    let m = a.match(/TAHUN\s+(\d{4})/i); if (m) year = +m[1];
    m = a.match(/=\s*(\d+)\s*Pemasangan/i); if (m) target = +m[1];
    m = a.match(/^BULAN\s+(\S+)/i);
    if (m) { month = m[1].charAt(0).toUpperCase() + m[1].slice(1).toLowerCase() + ' ' + year; if (!months.includes(month)) months.push(month); col = -1; return; }
    if (a === 'ID Karyawan') { col = r.findIndex(x => String(x).trim() === 'Total Pemasangan Individu'); return; }
    if (month && col >= 0 && /^[A-Za-z]+\d+$/.test(a)) (rows[a] = rows[a] || { id: a, name: r[1], m: {} }).m[month] = num_(r[col]);
  });
  const out = { target, months, rows: Object.values(rows) };
  try { c.put('kpi', JSON.stringify(out), 60); } catch (e) {}
  return out;
}
function getKpi(t) {
  const u = who_(t), d = kpiData_(), spv = u['Jabatan'].toUpperCase() === 'SPV';
  const mine = d.rows.find(r => r.id === u['NIK']);
  return { target: d.target, months: d.months, hasMe: !!mine, me: mine ? mine.m : {}, all: spv ? d.rows : null };
}
// Tabel referensi: Target Poin -> Total Rp (sheet "Reference", kolom A dan E)
function getRef(t) {
  who_(t);
  const c = CacheService.getScriptCache(), hit = c.get('ref');
  if (hit) return JSON.parse(hit);
  const sh = ksheet_('Reference'), v = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 5).getValues();
  const out = { rows: v.filter(r => r[0] !== '' && num_(r[0]) > 0).map(r => [num_(r[0]), num_(r[4])]) };
  try { c.put('ref', JSON.stringify(out), 600); } catch (e) {}
  return out;
}

// ===== Jadwal kerja (spreadsheet "Jadwal Masuk Periode ...": L = libur, P1 = masuk) =====
const JDW_ID = '1fTM6NeNs6zZ50IeeTTGH6jhE9fFMc0P1wRqAogz47cI';
const BLN_ALL = { januari: 0, februari: 1, maret: 2, april: 3, mei: 4, juni: 5, juli: 6, agustus: 7, september: 8, oktober: 9, november: 10, desember: 11 };
function parseJadwal_(sh) {
  const v = sh.getDataRange().getDisplayValues();
  let m = null;
  for (let i = 0; i < Math.min(v.length, 6) && !m; i++) for (const x of v[i]) { m = String(x).match(/PERIODE\s+(\d{1,2})\s+([A-Za-z]+)\s+S\s*\/?\s*D\s+(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})/i); if (m) break; }
  if (!m) return null;
  const sm = BLN_ALL[m[2].toLowerCase()], em = BLN_ALL[m[4].toLowerCase()], ey = +m[5];
  if (sm === undefined || em === undefined) return null;
  const hi = v.findIndex(r => String(r[0]).trim().toUpperCase() === 'NIK');
  if (hi < 0 || !v[hi + 1]) return null;
  let n = 0; while (/^\d{1,2}$/.test(String(v[hi + 1][2 + n]).trim())) n++;
  if (!n) return null;
  const start = Date.UTC(sm > em ? ey - 1 : ey, sm, +m[1]), days = [];
  for (let i = 0; i < n; i++) days.push(new Date(start + i * 864e5).toISOString().slice(0, 10));
  const rows = []; let shift = null;
  for (let r = hi + 2; r < v.length && /^[A-Za-z]+\d+$/.test(String(v[r][0]).trim()); r++)
    rows.push({ id: String(v[r][0]).trim(), name: v[r][1], codes: v[r].slice(2, 2 + n).map(x => String(x).trim()) });
  v.forEach(r => r.forEach(x => { const t = String(x).match(/(\d{1,2})[.:](\d{2})\s*-\s*(\d{1,2})[.:](\d{2})/); if (t && !shift) shift = { in: t[1].padStart(2, '0') + ':' + t[2], out: t[3].padStart(2, '0') + ':' + t[4] }; }));
  return { days, shift, rows };
}
function jadwalData_() {
  const c = CacheService.getScriptCache(), hit = c.get('jdw');
  if (hit) return JSON.parse(hit);
  const out = [];
  SpreadsheetApp.openById(JDW_ID).getSheets().forEach(sh => { const p = parseJadwal_(sh); if (p) out.push(p); });
  try { c.put('jdw', JSON.stringify(out), 120); } catch (e) {}
  return out;
}
function getJadwal(t) {
  const u = who_(t), spv = u['Jabatan'].toUpperCase() === 'SPV';
  const today = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd'), ps = jadwalData_();
  if (spv) { // Supervisor: hanya karyawan yang masuk hari ini
    const p = ps.find(x => x.days.includes(today)), i = p ? p.days.indexOf(today) : -1;
    return { spv: true, today, shift: p ? p.shift : null,
      list: i < 0 ? [] : p.rows.filter(r => r.codes[i] && r.codes[i].toUpperCase() !== 'L').map(r => ({ name: r.name, code: r.codes[i] })) };
  }
  return { spv: false, today, periods: ps.map(p => { const r = p.rows.find(x => x.id === u['NIK']); return r ? { days: p.days, shift: p.shift, codes: r.codes } : null; }).filter(Boolean) };
}

// ===== Daftar Kehadiran: jadwal (P1 = masuk, L = libur) + clock in/out dari sheet "Detail Absensi Harian" =====
const ATT_ID = '1tHZzQ4JEfgImeBiDH_5DM3sP-Ld7ovI7Yx7PJ3bf3aw';
function hm_(s) { const m = String(s || '').match(/(\d{1,2}):(\d{2})/); return m ? ('0' + m[1]).slice(-2) + ':' + m[2] : ''; }
function dkey_(s) {
  s = String(s || '').trim();
  let m = s.match(/(\d{4})-(\d{2})-(\d{2})/); if (m) return m[0];
  m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
  return m ? m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2) : '';
}
function attMap_(force) {
  const cc = CacheService.getScriptCache(), hit = !force && cc.get('att');
  if (hit) return JSON.parse(hit);
  const att = {}, sh = SpreadsheetApp.openById(ATT_ID).getSheets().find(s => s.getName().toLowerCase() === 'detail absensi harian');
  if (sh) {
    const v = sh.getDataRange().getDisplayValues();
    const hr = v.findIndex(r => String(r[0]).trim() === 'Hari' && String(r[1]).trim() === 'Tanggal');
    for (let r = hr + 1; hr >= 0 && r < v.length; r++) {
      const d = dkey_(v[r][1]), nik = String(v[r][2]).trim();
      if (d && nik) (att[nik] = att[nik] || {})[d] = { i: hm_(v[r][4]), o: hm_(v[r][5]) };
    }
  }
  try { cc.put('att', JSON.stringify(att), 60); } catch (e) {}
  return att;
}
function getAttend(t) {
  const u = who_(t), spv = u['Jabatan'].toUpperCase() === 'SPV', ps = jadwalData_(), att = attMap_();
  const today = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd'), tm = Utilities.formatDate(new Date(Date.now() + 864e5), 'Asia/Jakarta', 'yyyy-MM-dd');
  const cell = (r, d) => { const p = ps.find(x => x.days.includes(d)); if (!p) return null; const code = (r.codes[p.days.indexOf(d)] || '').toUpperCase(), a = (att[r.id] || {})[d] || {}; return { d, code, i: a.i || '', o: a.o || '' }; };
  if (spv) { // Supervisor: hanya karyawan yang masuk hari ini dan besok
    const g = d => { const p = ps.find(x => x.days.includes(d)); return { d, rows: !p ? [] : p.rows.map(r => Object.assign({ name: r.name }, cell(r, d))).filter(x => x.code && x.code !== 'L') }; };
    return { today, spv: true, groups: [g(today), g(tm)] };
  }
  const p = ps.find(x => x.days.includes(today) && x.rows.some(r => r.id === u['NIK'])) || ps.slice().reverse().find(x => x.rows.some(r => r.id === u['NIK']));
  const r = p && p.rows.find(x => x.id === u['NIK']);
  if (!r) return { today, none: true, days: [] };
  return { today, shift: p.shift, days: p.days.map(d => cell(r, d)) };
}

function getImages(t) {
  const u = who_(t);
  const e = tbl_('Data_Karyawan', 'NIK').rows.find(x => x['NIK'] === u['NIK']);
  return { logo: logo_(), foto: e ? img_(e['Foto Karyawan'], 240) : '' };
}

// Ganti password -> langsung ditulis ke kolom "Password" di sheet Users
function changePassword(t, oldPw, newPw) {
  const u = who_(t);
  if (u['Password'] !== String(oldPw)) return { ok: false, code: 'old' };
  newPw = String(newPw);
  if (newPw.length < 6) return { ok: false, code: 'short' };
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const col = tbl_('Users', 'Akun').head.indexOf('Password') + 1;
    sheet_('Users').getRange(u._row, col).setNumberFormat('@').setValue(newPw);
    CacheService.getScriptCache().remove('t_Users');
    const P = PropertiesService.getScriptProperties();
    P.deleteProperty('D_' + u['Akun']); // cabut semua sidik jari terdaftar
    const log = JSON.parse(P.getProperty('P_' + u['Akun']) || '[]'); log.push(Date.now());
    P.setProperty('P_' + u['Akun'], JSON.stringify(log.slice(-10))); // riwayat waktu ganti password (untuk Beranda)
  } finally { lock.releaseLock(); }
  return { ok: true };
}
