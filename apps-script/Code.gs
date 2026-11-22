/**
 * BACKEND — Form Approval TF Internasional Komputer
 * Tempel di: Spreadsheet "DATABASE APPROVAL TF" > Ekstensi > Apps Script
 * Script properties (Setelan project > Properti skrip):
 *   PIN          : PIN akses aplikasi (wajib)
 *   GEMINI_KEY   : API key Google AI Studio (untuk baca foto)
 *   GEMINI_MODEL : opsional, kosongkan agar model Flash terbaru dipilih otomatis
 * Laporan otomatis: jalankan pasangJadwalLaporan() sekali (kirim tiap tanggal 1 jam 07.00 WIB).
 * Deploy: Terapkan > Deployment baru > Aplikasi web
 *   Jalankan sebagai: Saya | Yang memiliki akses: Siapa saja
 */
var VERSION = '1.2';
var TZ = 'Asia/Jakarta';
var DEF_EMAIL = 'internasionalkomp.karanganyar@gmail.com';
var SHEETS = {
  requests: { name: 'REQUESTS', head: ['id','no','tanggal','jenis','supplier / rekanan','nominal','status','diajukan oleh','rekening','pesan_wa','updatedAt','data_json'] },
  rekening: { name: 'REKENING', head: ['id','nama','bank','nomor','atas_nama','kategori','updatedAt'] },
  settings: { name: 'SETTINGS', head: ['id','data_json'] }
};
var STATUS_LABEL = { wait:'Menunggu ACC', acc:'ACC Manager', tf:'Sudah TF', no:'Ditolak' };
var STATUS_KEY = { 'MENUNGGU ACC':'wait', 'ACC MANAGER':'acc', 'ACC':'acc', 'SUDAH TF':'tf', 'DITOLAK':'no' };

function doGet() {
  return out_({ ok: true, app: 'Approval TF', version: VERSION, info: 'Gunakan POST dengan PIN.' });
}

function doPost(e) {
  var req = {};
  try { req = JSON.parse(e.postData.contents || '{}'); } catch (x) { return out_({ ok: false, error: 'Format permintaan salah' }); }
  var pin = PropertiesService.getScriptProperties().getProperty('PIN');
  if (!pin) return out_({ ok: false, error: 'PIN belum diatur di Script properties', code: 'no_pin' });
  if (String(req.pin || '') !== String(pin)) return out_({ ok: false, error: 'PIN salah', code: 'bad_pin' });
  try {
    switch (req.action) {
      case 'ping': return out_({ ok: true, version: VERSION, ai: !!PropertiesService.getScriptProperties().getProperty('GEMINI_KEY') });
      case 'load': return out_({ ok: true, data: loadAll_(), ts: Date.now() });
      case 'set': case 'update': case 'delete': return out_(write_(req));
      case 'ai': return out_({ ok: true, text: ai_(req.prompt, req.images || []) });
      case 'sendReport': return out_(kirimLaporan_(req.month || bulanIni_()));
      default: return out_({ ok: false, error: 'Aksi tidak dikenal' });
    }
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  }
}

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function sheet_(key) {
  var cfg = SHEETS[key], ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(cfg.name);
  if (!sh) {
    sh = ss.insertSheet(cfg.name);
    sh.getRange(1, 1, 1, cfg.head.length).setValues([cfg.head]).setFontWeight('bold').setBackground('#0d47a1').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function rows_(key) {
  var sh = sheet_(key), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, SHEETS[key].head.length).getValues();
}

function loadAll_() {
  var req = {}, rek = {}, set = {};
  rows_('requests').forEach(function (r) {
    if (!r[0]) return;
    var d = {}; try { d = JSON.parse(r[11] || '{}'); } catch (x) {}
    var st = STATUS_KEY[String(r[6] || '').toUpperCase().trim()];
    if (st) d.status = st;               // status boleh diubah langsung di sheet
    req[r[0]] = d;
  });
  rows_('rekening').forEach(function (r) {
    if (!r[0] && !r[1]) return;
    var id = r[0] || slug_(r[1]);
    rek[id] = { nama: up_(r[1]), bank: up_(r[2]), nomor: String(r[3] || '').replace(/[^\d]/g, ''), an: up_(r[4]), kategori: String(r[5] || 'Supplier'), updatedAt: Number(r[6]) || 0 };
  });
  rows_('settings').forEach(function (r) { if (r[0]) { try { set[r[0]] = JSON.parse(r[1] || '{}'); } catch (x) {} } });
  return { requests: req, rekening: rek, settings: set };
}

function toRow_(key, id, d) {
  if (key === 'requests') {
    return [id, d.no || '', d.createdAt ? new Date(d.createdAt) : '', d.typeLabel || '', d.tag || '', Number(d.nominal) || 0,
      STATUS_LABEL[d.status] || d.status || '', d.pengaju || '', d.rek || '', d.wa || '', d.updatedAt || Date.now(), JSON.stringify(d)];
  }
  if (key === 'rekening') {
    return [id, d.nama || '', d.bank || '', d.nomor ? "'" + d.nomor : '', d.an || '', d.kategori || '', d.updatedAt || Date.now()];
  }
  return [id, JSON.stringify(d)];
}

function write_(req) {
  var key = req.col;
  if (!SHEETS[key]) return { ok: false, error: 'Koleksi tidak dikenal' };
  var id = String(req.id || '');
  if (!id) return { ok: false, error: 'ID kosong' };
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = sheet_(key), n = sh.getLastRow(), row = -1;
    if (n >= 2) {
      var ids = sh.getRange(2, 1, n - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) { row = i + 2; break; }
    }
    if (req.action === 'delete') {
      if (row > 0) sh.deleteRow(row);
      return { ok: true };
    }
    var data = req.data || {};
    if (req.action === 'update') {
      if (row < 0) return { ok: false, error: 'Data tidak ditemukan' };
      var cur = loadAll_()[key][id] || {};
      for (var k in data) cur[k] = data[k];
      data = cur;
    }
    var vals = toRow_(key, id, data);
    if (row > 0) sh.getRange(row, 1, 1, vals.length).setValues([vals]);
    else sh.appendRow(vals);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/* ---------------- Gemini ---------------- */
function ai_(prompt, images) {
  var p = PropertiesService.getScriptProperties(), key = p.getProperty('GEMINI_KEY');
  if (!key) throw new Error('GEMINI_KEY belum diatur di Script properties');
  var parts = [{ text: String(prompt || '') }];
  (images || []).slice(0, 4).forEach(function (im) { if (im && im.b64) parts.push({ inline_data: { mime_type: im.mime || 'image/jpeg', data: im.b64 } }); });
  var body = { contents: [{ role: 'user', parts: parts }], generationConfig: { temperature: 0.1, responseMimeType: 'application/json' } };
  var model = p.getProperty('GEMINI_MODEL') || pickModel_(key, false);
  var res = callGemini_(key, model, body);
  if (res.code === 404 || res.code === 400 && /model/i.test(res.text)) {   // model sudah diganti Google -> pilih ulang
    model = pickModel_(key, true); res = callGemini_(key, model, body);
  }
  if (res.code !== 200) throw new Error('Gemini ' + res.code + ': ' + res.text.slice(0, 200));
  var j = JSON.parse(res.text), c = j.candidates && j.candidates[0];
  var txt = c && c.content && c.content.parts ? c.content.parts.map(function (x) { return x.text || ''; }).join('') : '';
  if (!txt) throw new Error('Gemini tidak memberi jawaban');
  return txt;
}

function callGemini_(key, model, body) {
  var r = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent?key=' + encodeURIComponent(key),
    { method: 'post', contentType: 'application/json', payload: JSON.stringify(body), muteHttpExceptions: true });
  return { code: r.getResponseCode(), text: r.getContentText() };
}

/** Pilih model Flash stabil terbaru yang mendukung generateContent (disimpan 7 hari). */
function pickModel_(key, force) {
  var p = PropertiesService.getScriptProperties(), cached = p.getProperty('GEMINI_MODEL_AUTO'), at = Number(p.getProperty('GEMINI_MODEL_AT') || 0);
  if (!force && cached && Date.now() - at < 7 * 864e5) return cached;
  var r = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=200&key=' + encodeURIComponent(key), { muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('Gagal membaca daftar model Gemini (' + r.getResponseCode() + '). Cek GEMINI_KEY.');
  var list = (JSON.parse(r.getContentText()).models || []).filter(function (m) {
    var n = m.name.replace('models/', '');
    return /flash/i.test(n) && (m.supportedGenerationMethods || []).indexOf('generateContent') > -1 &&
      !/lite|tts|image|audio|live|preview|exp|embed|thinking|latest|\d{3}$/i.test(n);
  }).map(function (m) { return m.name.replace('models/', ''); });
  if (!list.length) throw new Error('Tidak ada model Gemini Flash yang tersedia untuk API key ini');
  list.sort(function (a, b) { return ver_(b) - ver_(a); });
  p.setProperty('GEMINI_MODEL_AUTO', list[0]); p.setProperty('GEMINI_MODEL_AT', String(Date.now()));
  return list[0];
}
function ver_(n) { var m = n.match(/(\d+)(?:\.(\d+))?/); return m ? Number(m[1]) * 100 + Number(m[2] || 0) : 0; }

function up_(s) { return String(s == null ? '' : s).replace(/\s+/g, ' ').trim().toUpperCase(); }
function slug_(s) { return up_(s).replace(/[^A-Z0-9 ]/g, '').replace(/ /g, '-').toLowerCase() || 'x'; }

/* ================= LAPORAN BULANAN OTOMATIS ================= */
var BLN = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
var LOGNM = { buat:'Dibuat', wait:'Menunggu ACC', acc:'ACC', tf:'TF', no:'Ditolak' };
function bulanIni_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM'); }
function bulanLalu_() { var d = new Date(); d.setDate(1); d.setMonth(d.getMonth() - 1); return Utilities.formatDate(d, TZ, 'yyyy-MM'); }
function label_(k) { var p = k.split('-'); return BLN[Number(p[1]) - 1] + ' ' + p[0]; }
function rp_(n) { return 'Rp ' + String(Math.round(n || 0)).replace(/\B(?=(\d{3})+(?!\d))/g, '.'); }
function tagOf_(d) { if (d.tag) return up_(d.tag); var f = d.fields || {}; return up_(f.suplier || f.rekanan || f.platform || f.karyawan || ''); }
function names_(s) { var o = []; String(s || '').toUpperCase().split(/\s*(?:,|&|\/|\bDAN\b)\s*/).forEach(function (n) { n = up_(n).replace(/^(T|M|MKT|TEK|K)\s*[.:]\s*/, ''); if (n && o.indexOf(n) < 0) o.push(n); }); return o.join(', '); }

function dataBulan_(k) {
  var all = loadAll_(), list = [];
  Object.keys(all.requests).forEach(function (id) {
    var d = all.requests[id]; if (!d || !d.createdAt) return;
    if (Utilities.formatDate(new Date(d.createdAt), TZ, 'yyyy-MM') !== k) return;
    list.push(d);
  });
  list.sort(function (a, b) { return a.createdAt - b.createdAt; });
  var act = list.filter(function (d) { return d.status !== 'no'; });
  var sum = function (a) { return a.reduce(function (s, d) { return s + (Number(d.nominal) || 0); }, 0); };
  function grp(fn) { var m = {}; act.forEach(function (d) { var g = fn(d) || '(tanpa nama)'; if (!m[g]) m[g] = { nama: g, n: 0, tot: 0, tf: 0, blm: 0 };
      var v = Number(d.nominal) || 0; m[g].n++; m[g].tot += v; if (d.status === 'tf') m[g].tf += v; else m[g].blm += v; });
    return Object.keys(m).map(function (g) { return m[g]; }).sort(function (a, b) { return b.tot - a.tot; }); }
  var set = (all.settings.pic2) || {};
  return { key: k, label: label_(k), list: list, set: set,
    kpi: { n: list.length, total: sum(act), tf: sum(act.filter(function (d) { return d.status === 'tf'; })),
      blm: sum(act.filter(function (d) { return d.status !== 'tf'; })), tolak: sum(list.filter(function (d) { return d.status === 'no'; })) },
    bySup: grp(tagOf_), byJenis: grp(function (d) { return String(d.typeLabel || '').replace('Request TF ', ''); }) };
}

function buatFileLaporan_(D) {
  var ss = SpreadsheetApp.create('Laporan TF ' + D.label), id = ss.getId();
  ss.setSpreadsheetTimeZone(TZ);
  var HB = '#1565c0', HD = '#0d47a1', ZB = '#f5f9ff', RP = '"Rp"#,##0;-"Rp"#,##0;"-"';
  var STL = { wait: 'Menunggu ACC', acc: 'ACC Manager', tf: 'Sudah TF', no: 'Ditolak' }, STC = { wait: '#a16207', acc: '#15803d', tf: '#1e6fd9', no: '#dc2626' };
  /* RINGKASAN */
  var s = ss.getSheets()[0]; s.setName('RINGKASAN'); s.setHiddenGridlines(true);
  [4, 300, 70, 130, 130, 130, 70].forEach(function (w, i) { s.setColumnWidth(i + 1, w); });
  s.getRange('B1').setValue('LAPORAN BULANAN REQUEST TF').setFontSize(16).setFontWeight('bold').setFontColor(HD);
  s.getRange('B2').setValue('Internasional Komputer Karanganyar').setFontWeight('bold');
  s.getRange('B3').setValue('Periode ' + D.label + '  ·  Dikirim otomatis ' + Utilities.formatDate(new Date(), TZ, 'd MMM yyyy, HH:mm')).setFontColor('#5b6c86').setFontSize(9);
  var r = 5;
  function sec(t) { s.getRange(r, 2, 1, 6).merge().setValue(t).setFontWeight('bold').setFontColor(HD).setBorder(null, null, true, null, null, null, HB, SpreadsheetApp.BorderStyle.SOLID_MEDIUM); r++; }
  sec('RINGKASAN ' + D.label.toUpperCase());
  [['Jumlah request (semua status)', D.kpi.n, '0', HD], ['Total pengajuan (tanpa ditolak)', D.kpi.total, RP, HD], ['Sudah ditransfer', D.kpi.tf, RP, '#1e6fd9'],
   ['Belum ditransfer (Menunggu ACC + ACC Manager)', D.kpi.blm, RP, '#a16207'], ['Ditolak', D.kpi.tolak, RP, '#dc2626']].forEach(function (k, i) {
    s.getRange(r, 2, 1, 3).merge().setValue(k[0]);
    s.getRange(r, 5, 1, 2).merge().setValue(k[1]).setNumberFormat(k[2]).setFontWeight('bold').setFontColor(k[3]).setHorizontalAlignment('right');
    s.getRange(r, 2, 1, 5).setBackground(i % 2 ? ZB : '#ffffff').setBorder(true, true, true, true, true, true, '#c9d8ee', SpreadsheetApp.BorderStyle.SOLID); r++; });
  r++;
  function tabel(t, items) {
    sec(t);
    s.getRange(r, 2, 1, 6).setValues([['Nama', 'Req', 'Total', 'Sudah TF', 'Belum TF', 'Porsi']]).setBackground(HB).setFontColor('#ffffff').setFontWeight('bold'); r++;
    if (!items.length) { s.getRange(r, 2).setValue('Belum ada data').setFontStyle('italic'); r += 2; return; }
    var rows = items.map(function (x) { return [x.nama, x.n, x.tot, x.tf, x.blm, D.kpi.total ? x.tot / D.kpi.total : 0]; });
    var tot = items.reduce(function (a, x) { return [a[0] + x.n, a[1] + x.tot, a[2] + x.tf, a[3] + x.blm]; }, [0, 0, 0, 0]);
    rows.push(['TOTAL', tot[0], tot[1], tot[2], tot[3], 1]);
    var rg = s.getRange(r, 2, rows.length, 6); rg.setValues(rows).setBorder(true, true, true, true, true, true, '#c9d8ee', SpreadsheetApp.BorderStyle.SOLID);
    s.getRange(r, 4, rows.length, 3).setNumberFormat(RP); s.getRange(r, 7, rows.length, 1).setNumberFormat('0.0%'); s.getRange(r, 3, rows.length, 1).setHorizontalAlignment('center');
    s.getRange(r, 5, rows.length - 1, 1).setFontColor('#1e6fd9'); s.getRange(r, 6, rows.length - 1, 1).setFontColor('#a16207');
    s.getRange(r + rows.length - 1, 2, 1, 6).setBackground(HD).setFontColor('#ffffff').setFontWeight('bold');
    r += rows.length + 1;
  }
  tabel('REKAP PER SUPPLIER / REKANAN / PLATFORM', D.bySup);
  tabel('REKAP PER JENIS REQUEST', D.byJenis);
  sec('PENGESAHAN');
  var nm = function (x) { return String(x || '').replace(/^@/, '').replace(/\s*\(.*\)$/, '').trim().toUpperCase(); };
  s.getRange(r, 2, 1, 5).setValues([['Dibuat oleh', '', 'Diperiksa oleh', '', 'Disetujui oleh']]).setFontColor('#5b6c86').setHorizontalAlignment('center');
  s.getRange(r + 4, 2, 1, 5).setValues([[nm(D.set.req || 'Nisa'), '', [nm(D.set.tf1 || 'Fachri'), nm(D.set.tf2)].filter(String).join(' / '), '', nm(D.set.acc || 'Azzam Muhammad')]])
    .setFontWeight('bold').setFontLine('underline').setHorizontalAlignment('center');
  /* DETAIL */
  var w = ss.insertSheet('DETAIL');
  var head = ['No', 'Tanggal', 'No Request', 'Jenis', 'Supplier / Rekanan', 'Nama Barang', 'Keterangan', 'No Nota', 'Karyawan Order', 'Diajukan', 'Status', 'Nominal', 'Jejak Status'];
  [36, 110, 120, 120, 170, 170, 260, 150, 130, 80, 95, 110, 260].forEach(function (x, i) { w.setColumnWidth(i + 1, x); });
  w.getRange(1, 1).setValue('DETAIL REQUEST TF — ' + D.label.toUpperCase()).setFontSize(14).setFontWeight('bold').setFontColor(HD);
  w.getRange(3, 1, 1, head.length).setValues([head]).setBackground(HB).setFontColor('#ffffff').setFontWeight('bold').setWrap(true);
  w.setFrozenRows(3);
  var body = D.list.map(function (d, i) {
    var f = d.fields || {};
    return [i + 1, Utilities.formatDate(new Date(d.createdAt), TZ, 'dd/MM/yyyy HH:mm'), d.no || '', String(d.typeLabel || '').replace('Request TF ', ''), tagOf_(d),
      up_(f.barang), String(f.ket || '').replace(/[*_]/g, '').replace(/\s*\n\s*/g, ' | '), up_(f.nota), names_(f.order || f.karyawan), up_(d.pengaju),
      STL[d.status] || d.status || '', Number(d.nominal) || 0,
      (d.log || []).map(function (e) { return (LOGNM[e.a] || e.a) + (e.by ? ' ' + e.by : '') + ' ' + Utilities.formatDate(new Date(e.at), TZ, 'dd/MM HH:mm'); }).join(' → ')];
  });
  if (body.length) {
    var br = w.getRange(4, 1, body.length, head.length);
    br.setValues(body).setVerticalAlignment('top').setWrap(true).setBorder(true, true, true, true, true, true, '#c9d8ee', SpreadsheetApp.BorderStyle.SOLID);
    w.getRange(4, 12, body.length, 1).setNumberFormat(RP);
    body.forEach(function (b, i) { var d = D.list[i]; w.getRange(4 + i, 11).setFontColor(STC[d.status] || '#a16207').setFontWeight('bold'); if (i % 2) w.getRange(4 + i, 1, 1, head.length).setBackground(ZB); });
    w.getRange(4 + body.length, 11, 1, 2).setValues([['TOTAL', D.kpi.total]]).setBackground(HD).setFontColor('#ffffff').setFontWeight('bold');
    w.getRange(4 + body.length, 12).setNumberFormat(RP);
  } else w.getRange(4, 1).setValue('Belum ada request bulan ini.');
  SpreadsheetApp.flush();
  return ss;
}

function export_(id, q) {
  var url = 'https://docs.google.com/spreadsheets/d/' + id + '/export?' + q;
  var r = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('Gagal ekspor laporan (' + r.getResponseCode() + ')');
  return r.getBlob();
}

function penerima_(D) {
  var e = String(D.set.emails || DEF_EMAIL).split(/[,;\s]+/).map(function (x) { return x.trim(); }).filter(function (x) { return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(x); });
  return e.length ? e : [DEF_EMAIL];
}

function kirimLaporan_(k) {
  var D = dataBulan_(k), ss = buatFileLaporan_(D), id = ss.getId(), nama = 'Laporan TF ' + D.label;
  try {
    var pdf = export_(id, 'format=pdf&size=A4&portrait=false&fitw=true&gridlines=false&printtitle=false&sheetnames=false&pagenum=CENTER&top_margin=0.4&bottom_margin=0.4&left_margin=0.4&right_margin=0.4').setName(nama + '.pdf');
    var xls = export_(id, 'format=xlsx').setName(nama + '.xlsx');
    var top = D.bySup.slice(0, 5).map(function (x) { return '<tr><td style="padding:4px 10px 4px 0">' + x.nama + '</td><td style="padding:4px 10px;text-align:right">' + x.n + 'x</td><td style="padding:4px 0;text-align:right;font-weight:bold">' + rp_(x.tot) + '</td></tr>'; }).join('');
    var html = '<div style="font-family:Segoe UI,Arial,sans-serif;color:#13223a;max-width:560px">' +
      '<div style="background:#0d47a1;color:#fff;padding:16px 18px;border-radius:12px 12px 0 0"><b style="font-size:17px">Laporan Bulanan Request TF</b><br><span style="opacity:.85">Internasional Komputer Karanganyar · ' + D.label + '</span></div>' +
      '<div style="border:1px solid #dbe6f5;border-top:0;padding:16px 18px;border-radius:0 0 12px 12px">' +
      '<table style="width:100%;font-size:14px;border-collapse:collapse">' +
      '<tr><td>Jumlah request</td><td style="text-align:right;font-weight:bold">' + D.kpi.n + '</td></tr>' +
      '<tr><td>Total pengajuan</td><td style="text-align:right;font-weight:bold">' + rp_(D.kpi.total) + '</td></tr>' +
      '<tr><td>Sudah ditransfer</td><td style="text-align:right;font-weight:bold;color:#1e6fd9">' + rp_(D.kpi.tf) + '</td></tr>' +
      '<tr><td>Belum ditransfer</td><td style="text-align:right;font-weight:bold;color:#a16207">' + rp_(D.kpi.blm) + '</td></tr>' +
      '<tr><td>Ditolak</td><td style="text-align:right;font-weight:bold;color:#dc2626">' + rp_(D.kpi.tolak) + '</td></tr></table>' +
      (top ? '<p style="margin:16px 0 6px;font-weight:bold;color:#0d47a1">Top supplier / rekanan</p><table style="font-size:13.5px;border-collapse:collapse">' + top + '</table>' : '') +
      '<p style="font-size:12px;color:#5b6c86;margin-top:16px">Lampiran: PDF &amp; Excel lengkap. Email ini dikirim otomatis oleh Dashboard Form Approval TF · SG.</p></div></div>';
    var to = penerima_(D);
    MailApp.sendEmail({ to: to.join(','), subject: 'Laporan TF ' + D.label + ' — Internasional Komputer Karanganyar', htmlBody: html, attachments: [pdf, xls], name: 'Dashboard Form Approval TF' });
    return { ok: true, label: D.label, to: to.join(', '), n: D.kpi.n };
  } finally {
    DriveApp.getFileById(id).setTrashed(true);
  }
}

/** Dipanggil pemicu tiap tanggal 1: kirim laporan bulan sebelumnya. */
function kirimLaporanOtomatis() { kirimLaporan_(bulanLalu_()); }

/** Jalankan SEKALI dari editor untuk memasang jadwal tanggal 1 pukul 07.00. */
function pasangJadwalLaporan() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'kirimLaporanOtomatis') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('kirimLaporanOtomatis').timeBased().onMonthDay(1).atHour(7).inTimezone(TZ).create();
  Logger.log('Jadwal laporan terpasang: tiap tanggal 1 sekitar pukul 07.00 WIB');
}

/** Uji kirim laporan bulan ini ke penerima di Pengaturan. */
function tesKirimLaporan() { Logger.log(JSON.stringify(kirimLaporan_(bulanIni_()))); }

/** Jalankan sekali dari editor untuk mengecek setelan (lihat Log eksekusi). */
function tesSetelan() {
  ['requests', 'rekening', 'settings'].forEach(sheet_);
  var p = PropertiesService.getScriptProperties();
  Logger.log('PIN: ' + (p.getProperty('PIN') ? 'sudah diatur' : 'BELUM diatur'));
  Logger.log('GEMINI_KEY: ' + (p.getProperty('GEMINI_KEY') ? 'sudah diatur' : 'BELUM diatur'));
  if (p.getProperty('GEMINI_KEY')) Logger.log('Model Gemini terpilih: ' + pickModel_(p.getProperty('GEMINI_KEY'), true));
  Logger.log('Data: ' + JSON.stringify(Object.keys(loadAll_().requests).length) + ' request');
}
