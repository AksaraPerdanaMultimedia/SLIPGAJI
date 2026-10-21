const SS_ID = '1itdPdb6NATZawFNjsNPPOxBD-HdTpA8heT8hmP2hOc4';
const COMPANY = 'PT Aksara Perdana Multimedia';
const TTL = 21600; // sesi login 6 jam
const SC = 60;     // cache data sheet 60 detik (supaya cepat, tetap hampir real-time)

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Slip Gaji - ' + COMPANY)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// API JSON untuk GitHub Pages (POST text/plain agar tanpa preflight CORS)
const API = { pub, login, getData, getImages, changePassword };
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
  return SpreadsheetApp.openById(SS_ID).getSheets().find(s => s.getName().toLowerCase() === name.toLowerCase());
}

// Baca sheet apa adanya (header dicari otomatis) + cache singkat agar cepat
function tbl_(name, key) {
  const c = CacheService.getScriptCache(), ck = 't_' + name;
  const hit = c.get(ck);
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
    return { p, locked: !spv && !!d && now < d, rel: d ? d.getDate() + ' ' + BLN[d.getMonth()] + ' ' + d.getFullYear() : '' };
  });
  const lockedSet = new Set(periods.filter(x => x.locked).map(x => x.p));
  const slips = s.filter(r => (spv || r['NIP'] === u['NIK']) && !lockedSet.has(r['Periode'])).map(r => {
    const e = emp.find(x => x['NIK'] === r['NIP']) || {};
    const o = Object.assign({}, r, { Nama: e['Nama Karyawan'] || r['NIP'], Jabatan: e['Jabatan'] || '', Departemen: e['Departemen'] || '', Bank: e['Bank'] || '', Rek: e['No Rekening'] || '' });
    delete o['Logo Perusahaan']; delete o._row; return o;
  });
  return {
    spv, slips, periods,
    me: { nik: u['NIK'], nama: u['Nama Karyawan'], jabatan: u['Jabatan'], dept: u['Departemen'] }
  };
}

function login(akun, pw) {
  const u = tbl_('Users', 'Akun').rows.find(r => r['Akun'].trim() === String(akun).trim());
  if (!u || u['Password'] !== String(pw)) { Utilities.sleep(500); return { ok: false, msg: 'Akun atau password salah' }; }
  const t = Utilities.getUuid();
  CacheService.getScriptCache().put('s_' + t, u['Akun'], TTL);
  return { ok: true, token: t, data: build_(u) };
}

function getData(t) { return build_(who_(t)); }

function getImages(t) {
  const u = who_(t);
  const e = tbl_('Data_Karyawan', 'NIK').rows.find(x => x['NIK'] === u['NIK']);
  return { logo: logo_(), foto: e ? img_(e['Foto Karyawan'], 240) : '' };
}

// Ganti password -> langsung ditulis ke kolom "Password" di sheet Users
function changePassword(t, oldPw, newPw) {
  const u = who_(t);
  if (u['Password'] !== String(oldPw)) return { ok: false, msg: 'Password lama salah' };
  newPw = String(newPw);
  if (newPw.length < 6) return { ok: false, msg: 'Password baru minimal 6 karakter' };
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const col = tbl_('Users', 'Akun').head.indexOf('Password') + 1;
    sheet_('Users').getRange(u._row, col).setNumberFormat('@').setValue(newPw);
    CacheService.getScriptCache().remove('t_Users');
  } finally { lock.releaseLock(); }
  return { ok: true };
}
