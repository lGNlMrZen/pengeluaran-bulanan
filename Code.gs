// Backend sinkronisasi Pengeluaran Bulanan (Google Apps Script)
// Ganti PIN di bawah dengan kata sandi buatanmu sendiri (minimal 8 karakter).
const PIN = 'GANTI-PIN-INI';

const WARNA = '#0f6b57'; // warna header spreadsheet (hijau, sama dengan tampilan halaman)
const RP = '"Rp "#,##0;[Red]-"Rp "#,##0';
const KEPALA_P = ['ID', 'Tanggal', 'Kategori', 'Keterangan', 'Jumlah'];
const KEPALA_I = ['Bulan', 'Pendapatan'];

function doGet() {
  return ContentService.createTextOutput('Backend pengeluaran aktif.');
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const req = JSON.parse(e.postData.contents);
    if (PIN === 'GANTI-PIN-INI') return keluar_({ ok: false, error: 'PIN di Code.gs belum diganti' });
    if (req.pin !== PIN) return keluar_({ ok: false, error: 'pin' });

    const st = ambil_();
    const ops = req.ops || [];
    ops.forEach(op => {
      if (op.t === 'add') {
        const id = Number(op.e.id);
        if (!st.data.some(d => d.id === id)) {
          st.data.push({ id: id, tgl: String(op.e.tgl), kat: String(op.e.kat), ket: String(op.e.ket), jml: Number(op.e.jml) });
        }
      } else if (op.t === 'del') {
        st.data = st.data.filter(d => d.id !== Number(op.id));
      } else if (op.t === 'delbulan') {
        st.data = st.data.filter(d => !d.tgl.startsWith(String(op.bulan)));
      } else if (op.t === 'inc') {
        const j = Number(op.jml);
        if (j > 0) st.pemasukan[String(op.bulan)] = j; else delete st.pemasukan[String(op.bulan)];
      } else if (op.t === 'incbila') { // isi hanya jika bulan itu belum punya pendapatan
        if (st.pemasukan[String(op.bulan)] === undefined && Number(op.jml) > 0) st.pemasukan[String(op.bulan)] = Number(op.jml);
      }
    });
    if (ops.length) tulis_(st);
    return keluar_({ ok: true, data: st.data, pemasukan: st.pemasukan });
  } catch (err) {
    return keluar_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Jalankan SEKALI dari editor (pilih fungsi ini di menu samping tombol Jalankan)
// untuk merapikan spreadsheet yang sudah ada tanpa menunggu ada catatan baru.
function rapikanSekarang() {
  PropertiesService.getScriptProperties().deleteProperty('rapi');
  tulis_(ambil_());
}

function keluar_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_(nama, kepala) {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(nama);
  if (!s) { s = ss.insertSheet(nama); s.appendRow(kepala); }
  return s;
}

function ambil_() {
  const sp = sheet_('Pengeluaran', KEPALA_P);
  const si = sheet_('Pendapatan', KEPALA_I);
  const data = sp.getDataRange().getValues().slice(1)
    .filter(r => r[0] !== '')
    .map(r => ({ id: Number(r[0]), tgl: String(r[1]), kat: String(r[2]), ket: String(r[3]), jml: Number(r[4]) }));
  const pemasukan = {};
  si.getDataRange().getValues().slice(1).forEach(r => {
    if (r[0] !== '') pemasukan[String(r[0])] = Number(r[1]);
  });
  return { data: data, pemasukan: pemasukan };
}

function tulis_(st) {
  // Urutkan berdasarkan tanggal agar di spreadsheet rapi dari awal sampai akhir bulan
  st.data.sort((a, b) => (a.tgl < b.tgl ? -1 : a.tgl > b.tgl ? 1 : a.id - b.id));
  isi_(sheet_('Pengeluaran', KEPALA_P),
    st.data.map(d => [d.id, d.tgl, d.kat, d.ket, d.jml]), 5, [2, 3, 4]);
  isi_(sheet_('Pendapatan', KEPALA_I),
    Object.keys(st.pemasukan).sort().map(b => [b, st.pemasukan[b]]), 2, [1]);
  ringkasan_(st);

  // Gaya (warna header, lebar kolom, filter) cukup diterapkan sekali agar sinkron tetap cepat
  const p = PropertiesService.getScriptProperties();
  if (p.getProperty('rapi') !== 'v1') { rapikan_(); p.setProperty('rapi', 'v1'); }
}

// Kolom teks diformat "teks biasa" agar tanggal/bulan tidak diubah Sheets jadi tanggal otomatis.
function isi_(s, rows, jumlahKolom, kolomTeks) {
  const akhir = s.getLastRow();
  if (akhir > 1) s.getRange(2, 1, akhir - 1, jumlahKolom).clearContent();
  if (!rows.length) return;
  kolomTeks.forEach(c => s.getRange(2, c, rows.length, 1).setNumberFormat('@'));
  s.getRange(2, 1, rows.length, jumlahKolom).setValues(rows);
}

// Sheet "Ringkasan": total per bulan dan per kategori.
// Nilainya dihitung ulang di skrip setiap ada perubahan (bukan rumus), supaya tidak
// bergantung pada pengaturan bahasa/wilayah spreadsheet.
function ringkasan_(st) {
  const rk = sheet_('Ringkasan', ['Bulan']);
  const bulan = {};
  st.data.forEach(d => { bulan[d.tgl.slice(0, 7)] = 1; });
  Object.keys(st.pemasukan).forEach(b => { bulan[b] = 1; });
  const bl = Object.keys(bulan).sort().reverse(); // bulan terbaru di atas
  const kat = Array.from(new Set(st.data.map(d => d.kat))).sort();
  const kepala = ['Bulan', 'Pendapatan', 'Pengeluaran', 'Sisa', '% terpakai'].concat(kat);
  const n = kepala.length;

  rk.getBandings().forEach(b => b.remove());
  rk.clear();
  rk.getRange(1, 1, 1, n).setValues([kepala])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground(WARNA)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  rk.setRowHeight(1, 32);
  rk.setFrozenRows(1);
  rk.setColumnWidth(1, 100);
  for (let c = 2; c <= n; c++) rk.setColumnWidth(c, 130);
  if (!bl.length) return;

  const jumlah = arr => arr.reduce((t, d) => t + d.jml, 0);
  const rows = bl.map(b => {
    const dat = st.data.filter(d => d.tgl.startsWith(b));
    const keluar = jumlah(dat);
    const masuk = st.pemasukan[b] || 0;
    return [b, masuk, keluar, masuk - keluar, masuk > 0 ? keluar / masuk : '']
      .concat(kat.map(k => jumlah(dat.filter(d => d.kat === k))));
  });
  rk.getRange(2, 1, bl.length, 1).setNumberFormat('@');
  rk.getRange(2, 1, bl.length, n).setValues(rows);
  rk.getRange(2, 2, bl.length, 3).setNumberFormat(RP).setHorizontalAlignment('right');
  rk.getRange(2, 5, bl.length, 1).setNumberFormat('0%').setHorizontalAlignment('right');
  if (kat.length) rk.getRange(2, 6, bl.length, kat.length).setNumberFormat(RP).setHorizontalAlignment('right');
  rk.getRange(1, 1, bl.length + 1, n)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
    .setHeaderRowColor(WARNA).setFirstRowColor('#ffffff').setSecondRowColor('#eef4f1');
}

// Setiap langkah gaya dicoba sendiri-sendiri: kalau satu gagal (misalnya karena fitur
// "Tabel" bawaan Google Sheets), langkah lain tetap berjalan.
function coba_(f) {
  try { f(); } catch (e) { console.log('Dilewati: ' + e); }
}

function gaya_(s, kepala, lebar) {
  const n = kepala.length;
  coba_(() => s.getRange(1, 1, 1, n).setValues([kepala])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground(WARNA)
    .setHorizontalAlignment('center').setVerticalAlignment('middle'));
  coba_(() => s.setRowHeight(1, 32));
  coba_(() => s.setFrozenRows(1));
  coba_(() => lebar.forEach((w, i) => s.setColumnWidth(i + 1, w)));
  coba_(() => {
    s.getBandings().forEach(b => b.remove());
    s.getRange(1, 1, s.getMaxRows(), n)
      .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
      .setHeaderRowColor(WARNA).setFirstRowColor('#ffffff').setSecondRowColor('#eef4f1');
  });
}

function rapikan_() {
  const ss = SpreadsheetApp.getActive();
  const sp = sheet_('Pengeluaran', KEPALA_P);
  const si = sheet_('Pendapatan', KEPALA_I);

  gaya_(sp, KEPALA_P, [60, 110, 140, 300, 130]);
  coba_(() => sp.hideColumns(1)); // kolom ID disembunyikan agar tidak terubah tanpa sengaja
  coba_(() => sp.getRange(2, 5, sp.getMaxRows() - 1, 1).setNumberFormat(RP).setHorizontalAlignment('right'));
  coba_(() => { // bisa saring per kategori/tanggal di Sheets
    if (sp.getFilter()) sp.getFilter().remove();
    sp.getRange(1, 1, sp.getMaxRows(), 5).createFilter();
  });

  gaya_(si, KEPALA_I, [110, 150]);
  coba_(() => si.getRange(2, 2, si.getMaxRows() - 1, 1).setNumberFormat(RP).setHorizontalAlignment('right'));

  // Ringkasan jadi tab pertama
  coba_(() => { ss.setActiveSheet(sheet_('Ringkasan', ['Bulan'])); ss.moveActiveSheet(1); });
}
