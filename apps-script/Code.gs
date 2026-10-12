/**
 * BUKU TEMPAHAN — HOMESTAY DESA BARAKAH
 * ------------------------------------------------------------
 * Kod ini menyambungkan app tempahan dengan Google Sheet ini.
 *
 *  - Setiap tempahan baru dari app masuk ke tab "Tempahan" dengan status "Menunggu".
 *  - Tukar status kepada "Disahkan" bila pelanggan dah bayar deposit.
 *    Tarikh yang "Disahkan" akan jadi kelabu (penuh) dalam app.
 *  - Tab "Tarikh Blok" untuk blok tarikh sendiri (contoh: tempahan melalui telefon, cuti keluarga).
 *
 * Langkah pertama: pilih fungsi "setup" di atas, kemudian tekan "Run".
 */

const HARGA_SEMALAM = 200; // RM, tukar di sini jika harga berubah
const TAB_TEMPAHAN = 'Tempahan';
const TAB_BLOK = 'Tarikh Blok';
const STATUS = ['Menunggu', 'Disahkan', 'Batal'];
const HEADERS = ['Kod', 'Masa Tempah', 'Nama', 'Telefon', 'Tarikh Datang', 'Tarikh Balik',
  'Malam', 'Tetamu', 'Jumlah (RM)', 'Permintaan Khas', 'Status'];
const COL = { kod: 1, masuk: 5, keluar: 6, status: 11 };

/** Jalankan SEKALI sahaja untuk sediakan Sheet. */
function setup() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone('Asia/Kuala_Lumpur');

  let t = ss.getSheetByName(TAB_TEMPAHAN);
  if (!t) {
    t = ss.getSheets()[0];
    if (t.getLastRow() > 0) t = ss.insertSheet();
    t.setName(TAB_TEMPAHAN);
  }
  t.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setBackground('#2b774a').setFontColor('#ffffff');
  t.setFrozenRows(1);
  t.getRange('B2:B').setNumberFormat('d/m/yyyy h:mm am/pm');
  t.getRange('E2:F').setNumberFormat('d/m/yyyy');
  t.getRange('I2:I').setNumberFormat('#,##0');
  const widths = [90, 140, 180, 120, 110, 110, 60, 70, 100, 220, 110];
  widths.forEach((w, i) => t.setColumnWidth(i + 1, w));
  t.getRange(2, COL.status, 999, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(STATUS, true).setAllowInvalid(false).build());
  t.clearConditionalFormatRules();
  const statusRange = t.getRange(2, COL.status, 999, 1);
  t.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Disahkan')
      .setBackground('#d4edda').setFontColor('#1e5e34').setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Menunggu')
      .setBackground('#fff3cd').setFontColor('#7a5a00').setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Batal')
      .setBackground('#eeeeee').setFontColor('#777777').setRanges([statusRange]).build()
  ]);

  let b = ss.getSheetByName(TAB_BLOK);
  if (!b) b = ss.insertSheet(TAB_BLOK);
  b.getRange(1, 1, 1, 3).setValues([['Tarikh Mula', 'Tarikh Akhir', 'Sebab']])
    .setFontWeight('bold').setBackground('#2b774a').setFontColor('#ffffff');
  b.setFrozenRows(1);
  b.getRange('A2:B').setNumberFormat('d/m/yyyy');
  b.setColumnWidth(1, 120); b.setColumnWidth(2, 120); b.setColumnWidth(3, 260);
  b.getRange(2, 1, 999, 2).setDataValidation(
    SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false)
      .setHelpText('Masukkan tarikh, contoh 25/12/2026').build());
  b.getRange('E1').setValue('Cara guna: isi Tarikh Mula dan Tarikh Akhir. Semua malam dari tarikh mula hingga tarikh akhir akan jadi penuh dalam app.')
    .setFontColor('#56635c').setFontStyle('italic');

  ss.setActiveSheet(t);
  SpreadsheetApp.getUi().alert('Siap! Buku tempahan sudah disediakan.');
}

/** App bertanya: tarikh mana yang sudah penuh? (Tiada maklumat pelanggan dikongsi.) */
function doGet() {
  const tz = 'Asia/Kuala_Lumpur';
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const booked = {};
  const addRange = (start, endExclusive) => {
    if (!(start instanceof Date) || !(endExclusive instanceof Date)) return;
    const d = new Date(start.getTime());
    let guard = 0;
    while (d < endExclusive && guard++ < 400) {
      const s = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
      if (s >= today) booked[s] = true;
      d.setDate(d.getDate() + 1);
    }
  };

  const ss = SpreadsheetApp.getActive();
  const t = ss.getSheetByName(TAB_TEMPAHAN);
  if (t && t.getLastRow() > 1) {
    t.getRange(2, 1, t.getLastRow() - 1, HEADERS.length).getValues().forEach(r => {
      if (String(r[COL.status - 1]).trim() === 'Disahkan') addRange(r[COL.masuk - 1], r[COL.keluar - 1]);
    });
  }
  const b = ss.getSheetByName(TAB_BLOK);
  if (b && b.getLastRow() > 1) {
    b.getRange(2, 1, b.getLastRow() - 1, 2).getValues().forEach(r => {
      if (r[0] instanceof Date) {
        const last = r[1] instanceof Date ? r[1] : r[0];
        const end = new Date(last.getTime()); end.setDate(end.getDate() + 1);
        addRange(r[0], end);
      }
    });
  }
  return json({ ok: true, booked: Object.keys(booked).sort() });
}

/** App menghantar tempahan baru. */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (d.action) return adminAction(d); // permintaan dari app Haji Berkuri
    const kod = clean(d.kod, 12), nama = safe(clean(d.nama, 80)), tel = safe(clean(d.telefon, 20));
    const nota = safe(clean(d.nota, 300));
    const isoOk = s => /^\d{4}-\d{2}-\d{2}$/.test(s);
    if (!/^DB-[A-Z0-9]{4,8}$/.test(kod) || !nama || !tel || !isoOk(d.masuk) || !isoOk(d.keluar) || d.keluar <= d.masuk) {
      return json({ ok: false, error: 'Maklumat tidak lengkap' });
    }
    const masuk = toDate(d.masuk), keluar = toDate(d.keluar);
    const malam = Math.round((keluar - masuk) / 864e5);
    if (malam < 1 || malam > 60) return json({ ok: false, error: 'Tarikh tidak sah' });

    const t = SpreadsheetApp.getActive().getSheetByName(TAB_TEMPAHAN);
    if (t.getLastRow() > 1) {
      const kods = t.getRange(2, COL.kod, t.getLastRow() - 1, 1).getValues().map(r => r[0]);
      if (kods.indexOf(kod) !== -1) return json({ ok: true, kod: kod, duplicate: true });
    }
    const tetamu = Math.max(1, Math.min(50, parseInt(d.tetamu, 10) || 1));
    t.appendRow([kod, new Date(), nama, tel, masuk, keluar, malam, tetamu, HARGA_SEMALAM * malam, nota, 'Menunggu']);
    return json({ ok: true, kod: kod });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/* ============================================================
 * APP PENGURUSAN HAJI BERKURI (dilindungi PIN)
 * ============================================================ */
const TZ = 'Asia/Kuala_Lumpur';
const clean = (v, n) => String(v == null ? '' : v).replace(/[\r\n\t]+/g, ' ').trim().slice(0, n);
const safe = v => /^[=+\-@]/.test(v) ? "'" + v : v; // elak formula pelik dalam Sheet
const isoOk = s => /^\d{4}-\d{2}-\d{2}$/.test(String(s));
const toDate = s => Utilities.parseDate(s, TZ, 'yyyy-MM-dd');
const toIso = v => v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : '';
const nextDay = s => { const d = toDate(s); d.setDate(d.getDate() + 1); return toIso(d); };
const overlap = (a1, a2, b1, b2) => a1 < b2 && b1 < a2; // julat [mula, tamat)

/** Menu dalam Google Sheet. */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Homestay Desa Barakah')
    .addItem('Tetapkan PIN app Haji Berkuri', 'tetapkanPin')
    .addToUi();
}

/** Tetapkan / tukar PIN untuk app pengurusan. */
function tetapkanPin() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('PIN app Haji Berkuri',
    'Masukkan PIN baru (6 hingga 12 nombor). PIN ini digunakan untuk masuk ke app pengurusan.',
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const pin = r.getResponseText().trim();
  if (!/^\d{6,12}$/.test(pin)) { ui.alert('PIN mesti 6 hingga 12 nombor sahaja. Sila cuba lagi.'); return; }
  PropertiesService.getScriptProperties().setProperty('ADMIN_PIN', pin);
  CacheService.getScriptCache().remove('pinfail');
  ui.alert('PIN berjaya disimpan. Gunakan PIN ini untuk masuk ke app Haji Berkuri.');
}

function checkPin(pin) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('pinfail') || 0);
  if (fails >= 5) return 'TERKUNCI';
  const stored = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN');
  if (!stored) return 'TIADA_PIN';
  if (String(pin || '') !== stored) { cache.put('pinfail', String(fails + 1), 900); return 'PIN_SALAH'; }
  cache.remove('pinfail');
  return '';
}

function readBookings() {
  const t = SpreadsheetApp.getActive().getSheetByName(TAB_TEMPAHAN);
  if (!t || t.getLastRow() < 2) return [];
  return t.getRange(2, 1, t.getLastRow() - 1, HEADERS.length).getValues().map((r, i) => ({
    row: i + 2, kod: String(r[0]), masa: r[1] instanceof Date ? Utilities.formatDate(r[1], TZ, 'yyyy-MM-dd HH:mm') : '',
    nama: String(r[2]), telefon: String(r[3]), masuk: toIso(r[4]), keluar: toIso(r[5]),
    malam: Number(r[6]) || 0, tetamu: Number(r[7]) || 0, jumlah: Number(r[8]) || 0,
    nota: String(r[9]), status: String(r[10]).trim()
  })).filter(b => b.kod);
}

function readBlocks() {
  const b = SpreadsheetApp.getActive().getSheetByName(TAB_BLOK);
  if (!b || b.getLastRow() < 2) return [];
  return b.getRange(2, 1, b.getLastRow() - 1, 3).getValues().map((r, i) => ({
    row: i + 2, mula: toIso(r[0]), akhir: toIso(r[1]) || toIso(r[0]), sebab: String(r[2])
  })).filter(x => x.mula);
}

function adminAction(d) {
  const err = checkPin(d.pin);
  if (err) return json({ ok: false, error: err });

  if (d.action === 'list') {
    return json({ ok: true, bookings: readBookings(), blocks: readBlocks(), today: toIso(new Date()), harga: HARGA_SEMALAM });
  }

  if (d.action === 'status') {
    const status = String(d.status);
    if (STATUS.indexOf(status) === -1) return json({ ok: false, error: 'Status tidak sah' });
    const all = readBookings();
    const b = all.filter(x => x.kod === String(d.kod))[0];
    if (!b) return json({ ok: false, error: 'Tempahan tidak dijumpai' });
    if (status === 'Disahkan') {
      const clash = all.filter(x => x.kod !== b.kod && x.status === 'Disahkan' && overlap(b.masuk, b.keluar, x.masuk, x.keluar))[0];
      if (clash) return json({ ok: false, error: 'Tarikh ini bertindih dengan tempahan ' + clash.kod + ' (' + clash.nama + ') yang sudah disahkan.' });
      const blk = readBlocks().filter(x => overlap(b.masuk, b.keluar, x.mula, nextDay(x.akhir)))[0];
      if (blk) return json({ ok: false, error: 'Tarikh ini bertindih dengan tarikh yang diblok' + (blk.sebab ? ' (' + blk.sebab + ')' : '') + '.' });
    }
    SpreadsheetApp.getActive().getSheetByName(TAB_TEMPAHAN).getRange(b.row, COL.status).setValue(status);
    return json({ ok: true });
  }

  if (d.action === 'block') {
    if (!isoOk(d.mula) || !isoOk(d.akhir) || d.akhir < d.mula) return json({ ok: false, error: 'Tarikh tidak sah. Tarikh akhir mesti sama atau selepas tarikh mula.' });
    SpreadsheetApp.getActive().getSheetByName(TAB_BLOK).appendRow([toDate(d.mula), toDate(d.akhir), safe(clean(d.sebab, 120))]);
    const clash = readBookings().filter(x => x.status === 'Disahkan' && overlap(d.mula, nextDay(d.akhir), x.masuk, x.keluar))[0];
    return json({ ok: true, warning: clash ? 'Perhatian: ada tempahan disahkan (' + clash.kod + ', ' + clash.nama + ') dalam tempoh ini.' : '' });
  }

  if (d.action === 'unblock') {
    const blk = readBlocks().filter(x => x.row === Number(d.row) && x.mula === d.mula)[0];
    if (!blk) return json({ ok: false, error: 'Tarikh blok tidak dijumpai. Sila muat semula.' });
    SpreadsheetApp.getActive().getSheetByName(TAB_BLOK).deleteRow(blk.row);
    return json({ ok: true });
  }

  return json({ ok: false, error: 'Arahan tidak dikenali' });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
