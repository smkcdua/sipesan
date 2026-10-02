/**
 * Penerima formulir "Jalur Orang Dalam" — SMK Cokroaminoto 2 Banjarnegara
 * Menulis data ke Google Spreadsheet dan menyimpan foto rapor ke Google Drive.
 *
 * CARA PASANG
 * 1. Buka script.google.com > Proyek baru > tempel seluruh kode ini.
 * 2. Pastikan akun yang menjalankan script punya akses EDIT ke spreadsheet di SHEET_ID.
 * 3. Deploy > Deployment baru > Jenis: Aplikasi web
 *      Jalankan sebagai: Saya
 *      Yang memiliki akses: Siapa saja
 * 4. Izinkan akses (Spreadsheet + Drive), salin URL /exec,
 *    lalu tempel ke APPS_SCRIPT_URL di jalurorangdalam.html.
 * 5. Setiap kode diubah, buat Deployment baru (versi baru) agar perubahan aktif.
 */
const SHEET_ID    = '1TkQxKO_F7_zL-lJq6PdcS6NMpmvubO1mHMwCM3s2xhU';
const SHEET_NAME  = 'Jalur Orang Dalam';
const FOLDER_NAME = 'Rapor Jalur Orang Dalam';
const MAX_BASE64  = 9 * 1024 * 1024;   // batas aman foto (±6,5 MB biner)

const HEADERS = [
  'No Pendaftaran', 'Waktu Kirim', 'Referal',
  'Nama', 'NISN', 'NIK', 'Jenis Kelamin', 'Asal SMP/MTs', 'Tahun Lulus',
  'Tempat Lahir', 'Tanggal Lahir', 'Desa', 'Kecamatan', 'Kabupaten/Kota',
  'WA Anak', 'Nama Ayah', 'WA Ayah', 'Nama Ibu', 'WA Ibu', 'Jarak (km)',
  'Jurusan Dipilih', 'Jurusan Rekomendasi', 'Skor Rekomendasi (%)', 'Peringkat Jurusan',
  'Nilai Mapel', 'Bakat', 'Minat', 'Tujuan', 'Gaya Belajar',
  'Link Foto Rapor', 'Status'
];

const WAJIB = ['nama', 'nisn', 'nik', 'jenis_kelamin', 'asal_sekolah', 'tahun_lulus', 'tempat_lahir',
  'tanggal_lahir', 'desa', 'kecamatan', 'kabupaten', 'wa_siswa', 'nama_ayah', 'wa_ayah',
  'nama_ibu', 'wa_ibu', 'jurusan_pilihan', 'referal'];

function doGet() {
  return json_({ ok: true, info: 'Endpoint Jalur Orang Dalam aktif.' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    const body = JSON.parse(e.postData.contents);
    const d = body.data || {};
    const r = body.rapor || {};

    // ---- validasi sisi server ----
    WAJIB.forEach(function (k) {
      if (!clean_(d[k])) throw new Error('Kolom "' + k + '" belum diisi.');
    });
    if (!/^\d{10}$/.test(clean_(d.nisn))) throw new Error('NISN harus 10 digit.');
    if (!/^\d{16}$/.test(clean_(d.nik))) throw new Error('NIK harus 16 digit.');
    if (!r.base64) throw new Error('Foto rapor wajib dilampirkan.');
    if (r.base64.length > MAX_BASE64) throw new Error('Ukuran foto terlalu besar.');
    if (['image/jpeg', 'image/png'].indexOf(r.mime) < 0) throw new Error('Format foto harus JPG atau PNG.');

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#ffe9d6');
    }

    // ---- cegah pendaftaran ganda (NISN sama) ----
    const kolomNisn = HEADERS.indexOf('NISN') + 1;
    if (sh.getLastRow() > 1) {
      const nisnAda = sh.getRange(2, kolomNisn, sh.getLastRow() - 1, 1).getValues().map(function (x) { return String(x[0]); });
      const idx = nisnAda.indexOf(clean_(d.nisn));
      if (idx >= 0) {
        const noLama = sh.getRange(idx + 2, 1).getValue();
        throw new Error('NISN ini sudah terdaftar dengan nomor ' + noLama + '. Hubungi panitia bila perlu mengubah data.');
      }
    }

    // ---- nomor pendaftaran & simpan foto ----
    const nomor = 'JOD-' + Utilities.formatString('%04d', sh.getLastRow());   // baris header dihitung 1
    const folder = folder_();
    const namaFile = nomor + '_' + clean_(d.nama).replace(/[^\w\- ]/g, '').trim().replace(/\s+/g, '_') + (r.mime === 'image/png' ? '.png' : '.jpg');
    const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(r.base64), r.mime, namaFile));

    const row = [
      nomor, new Date(), clean_(d.referal),
      clean_(d.nama), clean_(d.nisn), clean_(d.nik), clean_(d.jenis_kelamin), clean_(d.asal_sekolah), clean_(d.tahun_lulus),
      clean_(d.tempat_lahir), clean_(d.tanggal_lahir), clean_(d.desa), clean_(d.kecamatan), clean_(d.kabupaten),
      clean_(d.wa_siswa), clean_(d.nama_ayah), clean_(d.wa_ayah), clean_(d.nama_ibu), clean_(d.wa_ibu), clean_(d.jarak_km),
      clean_(d.jurusan_pilihan), clean_(d.jurusan_rekomendasi), clean_(d.skor_rekomendasi), clean_(d.ranking),
      clean_(d.nilai), clean_(d.bakat), clean_(d.minat), clean_(d.tujuan), clean_(d.gaya),
      file.getUrl(), 'Baru'
    ];
    const baris = sh.getLastRow() + 1;
    const rng = sh.getRange(baris, 1, 1, HEADERS.length);
    rng.setNumberFormat('@');                                   // semua teks: NISN/NIK/WA tidak berubah format
    sh.getRange(baris, 2).setNumberFormat('dd/MM/yyyy HH:mm:ss');
    rng.setValues([row]);

    return json_({ ok: true, no: nomor });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function folder_() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function clean_(v) {
  return (v === undefined || v === null) ? '' : String(v).trim().slice(0, 500);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
