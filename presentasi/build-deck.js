const pptxgen = require('pptxgenjs');
const path = require('path');
const JADE='0B3B2E', JADE_2='134C3B', CREAM='F6F1E7', GOLD='C8A24A',
      CLAY='A65A3A', TEAL='2E6E7E', INK='10221A', MUTE='6B7A72', LINE='DDD3BE';
const HEAD='Cambria', BODY='Calibri';
const EMBLEM = path.join(__dirname, '..', 'assets', 'emblem.png');
const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
pres.title = 'Spill de Tea — BMC & Robust Design';

/* 1 — title */
const s1 = pres.addSlide(); s1.background = { color: JADE };
s1.addShape(pres.ShapeType.ellipse, { x: 0.85, y: 0.8, w: 1.5, h: 1.5, fill: { color: CREAM }, line: { color: GOLD, width: 1.25 } });
s1.addImage({ path: EMBLEM, x: 0.96, y: 0.91, w: 1.28, h: 1.28 });
s1.addText('VENTURE CREATION PROJECT  ·  BINUS UNIVERSITY', { x: 0.85, y: 2.75, w: 11.6, h: 0.3, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12, bold: true, charSpacing: 2, color: GOLD });
s1.addText('Spill de Tea', { x: 0.85, y: 3.15, w: 11.6, h: 1.9, isTextBox: true, margin: 0, fontFace: HEAD, fontSize: 60, bold: true, color: CREAM });
s1.addText('Business Model Canvas  ·  Robust Design Identification Factors', { x: 0.85, y: 5.15, w: 11.6, h: 0.4, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 16, color: 'CFE0D4' });
s1.addText('Produksi batch H-1  ·  Pickup jalur khusus  ·  Pilot 500 botol di Binus Festival', { x: 0.85, y: 6.35, w: 11.6, h: 0.35, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12, color: '8FA89A' });
s1.addNotes('Slide pembuka. Deck ini dua bagian: model bisnisnya dulu (BMC), lalu pemetaan robust design untuk sistem produksi dan pickup.');

/* 2 — Business Model Canvas */
{
const TINT='F1F5F2';
const s = pres.addSlide();
s.background = { color: 'FFFFFF' };

s.addImage({ path: path.join(__dirname, '..', 'assets', 'emblem.png'), x: 0.45, y: 0.28, w: 0.62, h: 0.62 });
s.addText('Business Model Canvas', {
  x: 1.2, y: 0.3, w: 7, h: 0.42, isTextBox: true, margin: 0,
  fontFace: HEAD, fontSize: 24, bold: true, color: JADE
});
s.addText('Spill de Tea — milk tea botolan pre-order untuk event kampus', {
  x: 1.2, y: 0.7, w: 7, h: 0.26, isTextBox: true, margin: 0,
  fontFace: BODY, fontSize: 11.5, color: MUTE
});

const X0 = 0.45, GAP = 0.12, W = 13.333 - 2 * X0;
const CW = (W - 4 * GAP) / 5;
const TOP = 1.1, TOPH = 4.3, BOTY = TOP + TOPH + GAP, BOTH = 7.08 - BOTY;
const colX = i => X0 + i * (CW + GAP);
const halfH = (TOPH - GAP) / 2;
const FS = 9.5;
const fits = [];

function block(title, items, x, y, w, h, opt = {}) {
  const dark = !!opt.dark, fs = opt.fs || FS;
  s.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.06,
    fill: { color: dark ? JADE : TINT }, line: { color: dark ? JADE : TINT }
  });
  s.addShape(pres.ShapeType.rect, {
    x: x + 0.14, y: y + 0.17, w: 0.1, h: 0.1, fill: { color: GOLD }, line: { color: GOLD }
  });
  s.addText(title, {
    x: x + 0.32, y: y + 0.09, w: w - 0.44, h: 0.26, isTextBox: true, margin: 0,
    fontFace: BODY, fontSize: 11, bold: true, color: dark ? CREAM : JADE
  });
  const bw = w - 0.3, bh = h - 0.5;
  s.addText(items.map((t, i) => ({
    text: t, options: { bullet: { indent: 10 }, breakLine: i < items.length - 1, paraSpaceAfter: 2.5 }
  })), {
    x: x + 0.12, y: y + 0.42, w: bw, h: bh, isTextBox: true, margin: 0, valign: 'top',
    fontFace: BODY, fontSize: fs, color: dark ? 'E6EFE9' : INK
  });
  const cpl = Math.floor(((bw - 0.15) * 72) / (fs * 0.48));
  const lines = items.reduce((n, t) => n + Math.ceil(t.length / cpl), 0);
  const need = (lines * fs * 1.2 + items.length * 2.5) / 72;
  fits.push({ title, need: need.toFixed(2), have: bh.toFixed(2), ok: need <= bh * 0.95 });
}

block('Key Partners', [
  'Supplier grosir teh bubuk, susu bubuk & gula aren cair',
  'Pemasok bahan bean curd',
  'Vendor botol, label custom & sealer',
  'Panitia / EO Binus Festival & himpunan',
  'Penyedia pembayaran QRIS',
  'Jotform — platform form pre-order'
], colX(0), TOP, CW, TOPH, { fs: 10.5 });

block('Key Activities', [
  'Rekap pre-order & perencanaan batch',
  'Produksi batch H-1',
  'QC rasa, suhu & kemasan',
  'Operasi booth pickup jalur khusus',
  'Konten IG/TikTok & evaluasi tiap event'
], colX(1), TOP, CW, halfH);

block('Key Resources', [
  'Resep & SOP batch',
  'Alat produksi & cooler rantai dingin',
  'Tim mahasiswa TI — Service Systems Eng.',
  'Brand: emblem, desain cup, website'
], colX(1), TOP + halfH + GAP, CW, halfH);

block('Value Propositions', [
  'Rasa premium harga mahasiswa: Rp15.000–25.000 vs kompetitor Rp25.000–38.000',
  'Tanpa antre — pickup lewat jalur khusus pre-order',
  'Selalu fresh: diproduksi H-1, dijaga dingin sampai pickup',
  'Gula aren asli & bean curd sutra',
  'Cup bercetak penuh yang Instagrammable',
  'Zero waste — produksi sesuai pesanan pasti'
], colX(2), TOP, CW, TOPH, { dark: true, fs: 11 });

block('Customer Relationships', [
  'Komunitas lewat Instagram',
  'Paket bareng temen: 5 botol hemat Rp5.000, 10 botol gratis 1',
  'Konfirmasi personal via WhatsApp',
  'Kode order untuk pickup cepat'
], colX(3), TOP, CW, halfH);

block('Channels', [
  'Website + form pre-order Jotform',
  'Instagram & TikTok',
  'WhatsApp untuk konfirmasi & bayar',
  'Booth pickup di lokasi event'
], colX(3), TOP + halfH + GAP, CW, halfH);

block('Customer Segments', [
  'Mahasiswa Binus',
  'Pengunjung Binus Festival & event kampus',
  'Panitia & himpunan — pesanan kolektif',
  'Event kampus lain (fase 2)',
  'Pelanggan pre-order mingguan di area kampus (fase 3)'
], colX(4), TOP, CW, TOPH, { fs: 10.5 });

const BW = (W - GAP) / 2;
block('Cost Structure', [
  'Bahan baku ± Rp2.600/botol (teh, susu, gula aren, bean curd)',
  'Kemasan Rp1.500/botol (botol, label, sealer) → HPP ± Rp4.100',
  'Konten & promosi skala kecil · biaya booth event bila ada',
  'Tanpa sewa toko tetap & tanpa pegawai harian'
], X0, BOTY, BW, BOTH);

block('Revenue Streams', [
  'Penjualan per botol: Aren Rp15.000 · Jasmine Bean Curd Rp20.000 · Matcha Cloud Rp25.000',
  'Paket bundling 5 & 10 botol',
  'Deposit pre-order — kas masuk sebelum produksi',
  'Margin kotor ± 73% per botol (Aren)'
], X0 + BW + GAP, BOTY, BW, BOTH);

s.addNotes('Bacanya dari kanan: segmen dan value proposition dulu, lalu cara menjangkau, lalu yang di belakang layar. Tekankan bahwa deposit pre-order membuat kas masuk sebelum biaya produksi keluar.');

}

/* 3 — P-diagram */
const s2 = pres.addSlide(); s2.background = { color: 'FFFFFF' };
s2.addText('P-Diagram', { x: 0.65, y: 0.45, w: 8, h: 0.6, isTextBox: true, margin: 0, fontFace: HEAD, fontSize: 36, bold: true, color: JADE });
s2.addText('Sinyal masuk, respons keluar, noise menekan dari atas, kendali menopang dari bawah', { x: 0.65, y: 1.08, w: 11.5, h: 0.35, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 14, color: MUTE });
s2.addShape(pres.ShapeType.roundRect, { x: 4.75, y: 3.25, w: 3.8, h: 1.25, rectRadius: 0.12, fill: { color: 'F4F8F5' }, line: { color: JADE, width: 2 } });
s2.addText('Produksi batch\n& pickup Spill de Tea', { x: 4.75, y: 3.25, w: 3.8, h: 1.25, isTextBox: true, align: 'center', valign: 'middle', fontFace: BODY, fontSize: 16, bold: true, color: INK });
[5.45, 6.65, 7.85].forEach(x => s2.addShape(pres.ShapeType.line, { x, y: 2.15, w: 0, h: 1.0, line: { color: CLAY, width: 2.25, endArrowType: 'triangle' } }));
s2.addShape(pres.ShapeType.rect, { x: 6.42, y: 1.62, w: 0.13, h: 0.13, fill: { color: CLAY }, line: { color: CLAY } });
s2.addText('NOISE FACTOR', { x: 6.65, y: 1.52, w: 3, h: 0.3, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 14, bold: true, color: CLAY });
s2.addText('tidak bisa dikendalikan', { x: 4.75, y: 1.82, w: 3.8, h: 0.28, isTextBox: true, margin: 0, align: 'center', fontFace: BODY, fontSize: 11.5, color: MUTE });
[5.45, 6.65, 7.85].forEach(x => s2.addShape(pres.ShapeType.line, { x, y: 4.55, w: 0, h: 1.0, line: { color: JADE, width: 2.25, beginArrowType: 'triangle' } }));
s2.addShape(pres.ShapeType.rect, { x: 6.28, y: 5.72, w: 0.13, h: 0.13, fill: { color: JADE }, line: { color: JADE } });
s2.addText('CONTROL FACTOR', { x: 6.51, y: 5.62, w: 3.2, h: 0.3, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 14, bold: true, color: JADE });
s2.addText('kami kunci di SOP batch', { x: 4.75, y: 5.92, w: 3.8, h: 0.28, isTextBox: true, margin: 0, align: 'center', fontFace: BODY, fontSize: 11.5, color: MUTE });
s2.addShape(pres.ShapeType.line, { x: 1.95, y: 3.88, w: 2.7, h: 0, line: { color: GOLD, width: 2.25, endArrowType: 'triangle' } });
s2.addShape(pres.ShapeType.rect, { x: 0.68, y: 3.17, w: 0.13, h: 0.13, fill: { color: GOLD }, line: { color: GOLD } });
s2.addText('SIGNAL FACTOR', { x: 0.91, y: 3.07, w: 3, h: 0.3, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 14, bold: true, color: 'A8842F' });
s2.addText('input dari pelanggan', { x: 0.68, y: 3.38, w: 3.2, h: 0.28, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 11.5, color: MUTE });
s2.addShape(pres.ShapeType.line, { x: 8.65, y: 3.88, w: 2.7, h: 0, line: { color: TEAL, width: 2.25, endArrowType: 'triangle' } });
s2.addShape(pres.ShapeType.rect, { x: 9.42, y: 3.17, w: 0.13, h: 0.13, fill: { color: TEAL }, line: { color: TEAL } });
s2.addText('OUTPUT RESPONSE', { x: 9.65, y: 3.07, w: 2.95, h: 0.3, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 14, bold: true, color: TEAL });
s2.addText('harus tetap stabil', { x: 9.42, y: 3.38, w: 3.2, h: 0.28, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 11.5, color: MUTE });
s2.addNotes('Inti robust design: kita tidak melawan noise, tapi mengatur control factor supaya output tetap stabil meski noise berubah.');

/* 4 — table */
const s3 = pres.addSlide(); s3.background = { color: 'FFFFFF' };
s3.addText('Identifikasi Faktor', { x: 0.65, y: 0.46, w: 8, h: 0.6, isTextBox: true, margin: 0, fontFace: HEAD, fontSize: 32, bold: true, color: JADE });
s3.addText('Product / Service: Spill de Tea — milk tea botolan 250 ml, pre-order event kampus', { x: 0.65, y: 1.06, w: 11.9, h: 0.3, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12.5, color: MUTE });
const bullets = items => items.map((t, i) => ({ text: t, options: { bullet: true, fontSize: 10, color: INK, breakLine: i < items.length - 1 } }));
const typeCell = (n, sub, c) => ({ text: [ { text: n, options: { bold: true, fontSize: 12, color: c, breakLine: true } }, { text: sub, options: { fontSize: 9.5, color: MUTE } } ] });
const desc = t => ({ text: t, options: { fontSize: 10, color: '43534A' } });
const hdr = t => ({ text: t, options: { bold: true, fontSize: 10.5, color: MUTE, charSpacing: 1 } });
s3.addTable([
  [hdr('Factor Type'), hdr('Indikator'), hdr('Deskripsi / Dampak')],
  [typeCell('Control Factor', 'Kami tetapkan sendiri', JADE),
   { text: bullets(['Dosis teh & gula aren cair per liter','Suhu dan lama seduh teh','Rasio teh : susu : gula aren','Takaran bean curd per botol','Ukuran batch produksi','Produksi dikunci di H-1 sebelum pickup','Suhu rantai dingin dapur sampai booth','Jumlah petugas & jalur di booth pickup']) },
   desc('Penentu langsung konsistensi rasa antar botol dan kecepatan layani. Ditetapkan di SOP batch sebelum produksi, dan inilah yang kami setel ketika evaluasi event menunjukkan variasi.')],
  [typeCell('Noise Factor', 'Di luar kendali kami', CLAY),
   { text: bullets(['Suhu & kelembapan venue saat event','Variasi mutu bahan baku antar kiriman','Lama botol dibawa sebelum diminum','Lonjakan pengambilan di jam istirahat','Ketersediaan listrik & es di booth','Pembeli telat ambil atau tidak datang','Cuaca — hujan menurunkan kunjungan']) },
   desc('Sumber variasi rasa, suhu, dan antrean yang tidak bisa dihilangkan. Desain tidak melawan noise ini, tapi harus tetap memberi hasil sama meski semuanya berubah.')],
  [typeCell('Signal Factor', 'Datang dari pelanggan', 'A8842F'),
   { text: bullets(['Jumlah botol per pesanan (1 sampai paket 10)','Varian: Aren, Jasmine Bean Curd, Matcha Cloud','Kustomisasi: less sugar, tanpa bean curd','Slot waktu pickup yang dipilih','Total volume pre-order per event']) },
   desc('Input yang menggerakkan besar dan komposisi batch. Terkumpul lewat form pre-order sebelum hari-H, jadi sinyalnya sudah pasti saat produksi dimulai — bukan tebakan seperti booth walk-in.')],
  [typeCell('Output Response', 'Kami ukur tiap event', TEAL),
   { text: bullets(['Konsistensi rasa antar botol (QC sensori)','Suhu minuman saat diterima pembeli','Waktu tunggu di jalur pickup','Ketepatan pemenuhan — tepat waktu & tepat isi','HPP aktual per botol vs target Rp4.100','Botol terbuang terhadap jumlah produksi']) },
   desc('Hasil yang harus tetap stabil walau noise berubah. Dicatat tiap event lalu dipakai menyetel ulang control factor untuk siklus berikutnya — ini yang menutup loop evaluasi.')]
], { x: 0.65, y: 1.48, w: 12.0, colW: [2.5, 5.2, 4.3], border: { type: 'solid', color: LINE, pt: 0.75 },
     fill: { color: 'FFFFFF' }, fontFace: BODY, valign: 'top', margin: [5, 9, 5, 9] });
s3.addNotes('Tabel utama sesuai template. Tekankan bahwa signal factor sudah pasti karena pre-order, itu keunggulan sistem kami dibanding booth walk-in.');

/* 5 — targets */
const s4 = pres.addSlide(); s4.background = { color: JADE };
s4.addText('Target Output Response', { x: 0.65, y: 0.6, w: 9, h: 0.65, isTextBox: true, margin: 0, fontFace: HEAD, fontSize: 34, bold: true, color: CREAM });
s4.addText('Batas yang kami pakai untuk menilai tiap event', { x: 0.65, y: 1.26, w: 9, h: 0.35, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 14, color: '9FBBAB' });
[['≥ 95%','Lolos QC sensori'],['≤ 10 °C','Suhu saat diterima'],['≤ 2 menit','Waktu tunggu pickup'],
 ['≥ 98%','Pesanan tepat waktu'],['≤ Rp4.100','HPP aktual per botol'],['≤ 2%','Botol terbuang']].forEach((t, i) => {
  const x = 0.65 + (i % 3) * 4.12, y = 2.1 + ((i / 3) | 0) * 2.05;
  s4.addShape(pres.ShapeType.roundRect, { x, y, w: 3.72, h: 1.72, rectRadius: 0.1, fill: { color: JADE_2 }, line: { color: '1E6650', width: 1 } });
  s4.addText(t[0], { x: x + 0.3, y: y + 0.3, w: 3.12, h: 0.7, isTextBox: true, margin: 0, fontFace: HEAD, fontSize: 30, bold: true, color: GOLD });
  s4.addText(t[1], { x: x + 0.3, y: y + 1.02, w: 3.12, h: 0.4, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12.5, color: 'B9CFC2' });
});
s4.addText('Angka di atas adalah target yang kami tetapkan sebagai tolok ukur, belum hasil pengukuran. Pengukuran sebenarnya diambil pada pilot 500 botol di Binus Festival.', { x: 0.65, y: 6.45, w: 11.9, h: 0.55, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 11.5, italic: true, color: '8FA89A' });
s4.addNotes('Penting: sebut bahwa ini target, bukan data hasil pengukuran. Pengukuran nyata baru ada setelah pilot.');

pres.writeFile({ fileName: path.join(__dirname, 'Spill_de_Tea_BMC_Robust_Design.pptx') }).then(f => console.log('tersimpan:', f));
