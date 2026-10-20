/* Spill de Tea — Creativity & Innovation deck.
   Structure follows the course example; every figure here comes from the
   project itself (pitch deck, website, BMC). Survey slides are deliberately
   absent until a survey has actually been run. */
const pptxgen = require('pptxgenjs');
const path = require('path');

const JADE='0B3B2E', JADE_2='134C3B', CREAM='F6F1E7', GOLD='C8A24A', GOLD_D='9A7A2C',
      INK='10221A', MUTE='5E6E65', TINT='F1F5F2', CLAY='A65A3A', CLAY_T='F7EDE8',
      TEAL='2E6E7E', TEAL_T='EAF2F4', LINE='D9D2C1';
const HEAD='Cambria', BODY='Calibri';
const SW=13.333, SH=7.5, MX=0.6;
const img = f => path.join(__dirname, 'img', f);
const EMBLEM = path.join(__dirname, '..', 'assets', 'emblem.png');

const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
pres.title = 'Spill de Tea — Creativity & Innovation';

/* ── helpers ─────────────────────────────────────────────── */
function txt(s, text, o) { s.addText(text, Object.assign({ isTextBox: true, margin: 0, fontFace: BODY, color: INK }, o)); }

function heading(s, title, sub) {
  txt(s, title, { x: MX, y: 0.42, w: 10.5, h: 0.62, fontFace: HEAD, fontSize: 30, bold: true, color: JADE });
  if (sub) txt(s, sub, { x: MX, y: 1.04, w: 11.8, h: 0.3, fontSize: 13, color: MUTE });
  s.addImage({ path: EMBLEM, x: SW - MX - 0.62, y: 0.38, w: 0.62, h: 0.62 });
}

function bullets(items, fs, color) {
  return items.map((t, i) => ({ text: t, options: {
    bullet: { indent: 11 }, breakLine: i < items.length - 1, paraSpaceAfter: 3, fontSize: fs, color } }));
}

function card(s, x, y, w, h, head, body, o = {}) {
  const fill = o.fill || TINT, hc = o.headColor || JADE, fs = o.fs || 11.5, tc = o.textColor || INK;
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.07, fill: { color: fill }, line: { color: fill } });
  s.addShape(pres.ShapeType.rect, { x: x + 0.18, y: y + 0.24, w: 0.11, h: 0.11, fill: { color: o.mark || GOLD }, line: { color: o.mark || GOLD } });
  txt(s, head, { x: x + 0.38, y: y + 0.14, w: w - 0.55, h: 0.3, fontSize: 13, bold: true, color: hc });
  const content = Array.isArray(body) ? bullets(body, fs, tc) : body;
  txt(s, content, { x: x + 0.18, y: y + 0.55, w: w - 0.36, h: h - 0.7, fontSize: fs, color: tc, valign: 'top',
    lineSpacingMultiple: Array.isArray(body) ? 1 : 1.12 });
}

/* ── 1. cover ────────────────────────────────────────────── */
{
  const s = pres.addSlide(); s.background = { color: JADE };
  s.addShape(pres.ShapeType.ellipse, { x: MX, y: 0.75, w: 1.35, h: 1.35, fill: { color: CREAM }, line: { color: GOLD, width: 1.25 } });
  s.addImage({ path: EMBLEM, x: MX + 0.1, y: 0.85, w: 1.15, h: 1.15 });
  txt(s, 'CREATIVITY & INNOVATION  ·  VENTURE CREATION PROJECT', { x: MX, y: 2.55, w: 7.2, h: 0.3, fontSize: 12, bold: true, charSpacing: 2, color: GOLD });
  txt(s, 'Spill de Tea', { x: MX, y: 2.95, w: 7.2, h: 1.25, fontFace: HEAD, fontSize: 66, bold: true, color: CREAM });
  txt(s, 'Premium taste, student price.', { x: MX, y: 4.25, w: 7.2, h: 0.5, fontFace: HEAD, fontSize: 24, italic: true, color: GOLD });
  txt(s, 'Milk tea botolan pre-order untuk event kampus — diproduksi segar per batch, diambil tanpa antre.', {
    x: MX, y: 5.05, w: 6.6, h: 0.75, fontSize: 15, color: 'CFE0D4' });
  txt(s, 'Binus University', { x: MX, y: 6.6, w: 6, h: 0.3, fontSize: 12, color: '8FA89A' });
  s.addShape(pres.ShapeType.rect, { x: 8.55, y: 0.55, w: 4.2, h: 6.4, fill: { color: JADE_2 }, line: { color: GOLD, width: 1 } });
  s.addImage({ path: img('cup3d.jpg'), x: 8.65, y: 0.65, w: 4.0, h: 6.2, sizing: { type: 'cover', w: 4.0, h: 6.2 } });
  s.addNotes('Pembuka. Satu kalimat posisi kami: rasa premium, harga mahasiswa, tanpa antre.');
}

/* ── 2. business overview ────────────────────────────────── */
{
  const s = pres.addSlide(); s.background = { color: 'FFFFFF' };
  heading(s, 'Business Overview', 'Apa yang kami jual, dan kenapa caranya berbeda');
  s.addShape(pres.ShapeType.roundRect, { x: MX, y: 1.6, w: 7.45, h: 5.2, rectRadius: 0.08, fill: { color: TINT }, line: { color: TINT } });
  txt(s, [
    { text: 'Spill de Tea adalah brand milk tea botolan yang berjualan lewat sistem pre-order untuk event kampus.', options: { bold: true, breakLine: true } },
    { text: ' ', options: { fontSize: 8, breakLine: true } },
    { text: 'Kami berangkat dari keresahan sederhana: minuman kekinian itu enak, tapi antre di booth makan waktu istirahat, harganya berat untuk kantong mahasiswa, dan esnya keburu cair sebelum sempat diminum.', options: { breakLine: true } },
    { text: ' ', options: { fontSize: 8, breakLine: true } },
    { text: 'Pelanggan memesan lewat form sebelum hari-H. Kami memproduksi satu batch sesuai jumlah pesanan pasti pada H-1, lalu pesanan diambil di jalur pickup khusus yang terpisah dari antrean walk-in. Hasilnya minuman dengan gula aren asli dan bean curd, dalam cup bercetak penuh, mulai Rp13.000 — sekitar setengah harga brand teh premium.' }
  ], { x: MX + 0.35, y: 1.9, w: 6.75, h: 4.7, fontSize: 16.5, valign: 'top', lineSpacingMultiple: 1.15 });
  [['Rp13.000–15.000', 'per botol 250 ml, vs Rp25.000–38.000 di brand premium'],
   ['Produksi H-1', 'satu batch sesuai pesanan pasti — tidak ada stok yang dibuang'],
   ['Jalur khusus', 'pickup pre-order terpisah dari antrean walk-in']].forEach((f, i) => {
    const y = 1.6 + i * 1.78;
    s.addShape(pres.ShapeType.roundRect, { x: 8.35, y, w: 4.38, h: 1.6, rectRadius: 0.08, fill: { color: JADE }, line: { color: JADE } });
    txt(s, f[0], { x: 8.65, y: y + 0.25, w: 3.9, h: 0.55, fontFace: HEAD, fontSize: 24, bold: true, color: GOLD });
    txt(s, f[1], { x: 8.65, y: y + 0.85, w: 3.9, h: 0.6, fontSize: 12, color: 'D6E4DB' });
  });
  s.addNotes('Tiga hal yang membedakan: harga, produksi sesuai pesanan, dan jalur pickup khusus.');
}

/* ── 3–5. personas ───────────────────────────────────────── */
const personas = [
  { n: 'Dimas', ini: 'D', age: '20 tahun', role: 'Mahasiswa Teknik Industri, semester 5',
    arch: 'Si mahasiswa dengan jadwal padat',
    bg: 'Kuliah dari pagi sampai sore dengan jeda istirahat yang pendek. Sering datang ke event kampus di sela kelas dan biasanya beli minum di booth yang ada.',
    goal: ['Minuman enak tanpa menghabiskan waktu istirahat', 'Harga yang aman untuk dibeli rutin'],
    mot: ['Hemat waktu', 'Hemat uang', 'Tetap dapat rasa yang enak'],
    fr: ['Antre panjang di booth saat event', 'Minuman premium Rp25.000+ terlalu mahal untuk tiap hari', 'Es cepat cair, rasa berubah sebelum habis'],
    q: 'Istirahat cuma sebentar, masa habis buat antre minum.' },
  { n: 'Nadia', ini: 'N', age: '21 tahun', role: 'Bendahara himpunan, panitia event kampus',
    arch: 'Si pengurus konsumsi panitia',
    bg: 'Mengurus konsumsi panitia dan tamu di setiap acara himpunan. Harus memesan banyak sekaligus dengan anggaran yang sudah dipatok sejak awal.',
    goal: ['Pesan puluhan botol sekali jalan', 'Pesanan pasti ada tepat waktu', 'Harga sesuai anggaran'],
    mot: ['Satu kontak untuk pesanan kolektif', 'Harga paket', 'Kepastian sebelum hari-H'],
    fr: ['Membeli satu per satu di booth makan waktu', 'Harga premium membuat anggaran jebol', 'Tidak ada kepastian stok saat acara'],
    q: 'Yang penting pesanan pasti ada dan sesuai budget.' },
  { n: 'Kayla', ini: 'K', age: '19 tahun', role: 'Mahasiswi Komunikasi, aktif di Instagram & TikTok',
    arch: 'Si pemburu yang layak di-post',
    bg: 'Suka mencoba minuman baru dan hampir selalu memotretnya sebelum diminum. Produk yang tampil beda lebih mudah menarik perhatiannya.',
    goal: ['Mencoba minuman yang unik', 'Konten ringan untuk media sosial'],
    mot: ['Kemasan yang estetik', 'Brand yang punya cerita', 'Harga tetap ramah kantong'],
    fr: ['Cup polos yang terlihat biasa saja', 'Minuman yang cantik biasanya mahal', 'Tidak ada pembeda dari brand lain'],
    q: 'Kalau cup-nya cantik, pasti aku post.' }
];
personas.forEach((p, i) => {
  const s = pres.addSlide(); s.background = { color: 'FFFFFF' };
  heading(s, `User Persona ${i + 1}`, p.arch);
  s.addShape(pres.ShapeType.roundRect, { x: MX, y: 1.6, w: 3.75, h: 5.2, rectRadius: 0.08, fill: { color: JADE }, line: { color: JADE } });
  s.addShape(pres.ShapeType.ellipse, { x: MX + 1.15, y: 1.95, w: 1.45, h: 1.45, fill: { color: CREAM }, line: { color: GOLD, width: 1.5 } });
  txt(s, p.ini, { x: MX + 1.15, y: 1.95, w: 1.45, h: 1.45, align: 'center', valign: 'middle', fontFace: HEAD, fontSize: 44, bold: true, color: JADE });
  txt(s, p.n, { x: MX + 0.3, y: 3.55, w: 3.15, h: 0.5, align: 'center', fontFace: HEAD, fontSize: 24, bold: true, color: CREAM });
  txt(s, `${p.age}\n${p.role}`, { x: MX + 0.3, y: 4.05, w: 3.15, h: 0.85, align: 'center', fontSize: 12.5, color: 'D6E4DB' });
  txt(s, `“${p.q}”`, { x: MX + 0.3, y: 5.15, w: 3.15, h: 1.4, align: 'center', valign: 'middle', fontFace: HEAD, fontSize: 16, italic: true, color: GOLD });
  const X = MX + 4.0, CWd = (SW - MX - X - 0.2) / 2, CH = 2.5;
  card(s, X, 1.6, CWd, CH, 'Latar Belakang', p.bg, { fs: 14 });
  card(s, X + CWd + 0.2, 1.6, CWd, CH, 'Goal', p.goal, { fs: 14 });
  card(s, X, 1.6 + CH + 0.2, CWd, CH, 'Motivasi', p.mot, { fs: 14 });
  card(s, X + CWd + 0.2, 1.6 + CH + 0.2, CWd, CH, 'Frustrasi', p.fr, { fs: 14, fill: CLAY_T, headColor: CLAY, mark: CLAY });
  txt(s, 'Proto-persona: disusun dari masalah yang kami identifikasi, belum divalidasi survei.', {
    x: MX, y: 6.95, w: 9, h: 0.25, fontSize: 9.5, italic: true, color: MUTE });
  s.addNotes(`Persona ${i + 1}: ${p.arch}. Hubungkan ke segmen di BMC.`);
});

/* ── 6. empathy map ──────────────────────────────────────── */
{
  const s = pres.addSlide(); s.background = { color: 'FFFFFF' };
  heading(s, 'Empathy Map', 'Pelanggan utama: mahasiswa yang membeli minuman di event kampus');
  const cells = [
    ['Think & Feel', ['Ingin minuman kekinian, tapi sayang uangnya', 'Waktu istirahat terlalu berharga untuk antre'], {}],
    ['See', ['Antrean panjang di booth saat event', 'Brand teh premium dengan harga Rp25.000 ke atas', 'Teman memposting minuman dengan cup estetik'], {}],
    ['Hear', ['Ajakan teman untuk beli bareng biar hemat', 'Info event dan promo dari Instagram himpunan'], {}],
    ['Say & Do', ['Beli minuman di sela jadwal kuliah', 'Patungan memesan bareng satu kelas', 'Memotret minuman sebelum diminum'], {}],
    ['Pain', ['Antre lama', 'Harga berat untuk dibeli rutin', 'Es cair, rasa berubah'], { fill: CLAY_T, headColor: CLAY, mark: CLAY }],
    ['Gain', ['Minuman premium dengan harga mahasiswa', 'Tinggal ambil tanpa antre', 'Cup yang layak jadi konten'], { fill: TEAL_T, headColor: TEAL, mark: TEAL }]
  ];
  const cw = (SW - 2 * MX - 0.4) / 3, ch = 2.5;
  cells.forEach((c, i) => card(s, MX + (i % 3) * (cw + 0.2), 1.6 + ((i / 3) | 0) * (ch + 0.2), cw, ch, c[0], c[1], Object.assign({ fs: 15 }, c[2])));
  s.addNotes('Baris atas apa yang mereka alami; baris bawah apa yang mereka lakukan, sakitnya, dan yang mereka cari.');
}

/* ── 7. about + products ─────────────────────────────────── */
{
  const s = pres.addSlide(); s.background = { color: 'FFFFFF' };
  heading(s, 'About Spill de Tea', 'Tiga racikan inti — semua 250 ml, botolan, fresh batch');
  txt(s, 'Spill de Tea meracik milk tea dengan teh hitam dan melati yang diseduh pekat, gula aren cair asli, dan bean curd yang lembut — lalu menjualnya di harga yang masuk akal untuk dibeli rutin. Setiap botol dibuat pada H-1 sesuai pesanan yang sudah masuk, dijaga dingin sampai diambil, dan dikemas dalam cup bercetak penuh yang layak dipotret.', {
    x: MX, y: 1.6, w: 12.1, h: 1.15, fontSize: 14.5, lineSpacingMultiple: 1.12, color: '2A3A31' });
  const menu = [
    ['menu-1.jpg', 'Aren Milk Tea', 'Rp13.000', 'Teh hitam pekat, susu creamy, gula aren asli.'],
    ['menu-2.jpg', 'Jasmine Bean Curd Latte', 'Rp15.000', 'Teh melati dan susu lembut, ditutup bean curd.'],
    ['menu-3.jpg', 'Matcha Cloud', 'Rp15.000', 'Matcha diseduh dingin dengan lapisan cream cloud.']];
  const w = (SW - 2 * MX - 0.5) / 3;
  menu.forEach((m, i) => {
    const x = MX + i * (w + 0.25), y = 3.0;
    s.addImage({ path: img(m[0]), x, y, w, h: w * 0.786 });
    txt(s, m[1], { x, y: y + w * 0.786 + 0.12, w: w - 1.3, h: 0.4, fontFace: HEAD, fontSize: 17, bold: true, color: JADE });
    txt(s, m[2], { x: x + w - 1.3, y: y + w * 0.786 + 0.12, w: 1.3, h: 0.4, align: 'right', fontFace: HEAD, fontSize: 17, bold: true, color: GOLD_D });
    txt(s, m[3], { x, y: y + w * 0.786 + 0.55, w, h: 0.4, fontSize: 12, color: MUTE });
  });
  s.addNotes('Thai Tea sudah tidak dijual; menu sekarang tiga ini.');
}

/* ── 8. SDGs ─────────────────────────────────────────────── */
{
  const s = pres.addSlide(); s.background = { color: 'FFFFFF' };
  heading(s, 'SDGs', 'Kontribusi kami: SDG 12 — Responsible Consumption and Production');
  s.addShape(pres.ShapeType.rect, { x: MX, y: 1.6, w: 3.3, h: 3.3, fill: { color: 'BF8B2E' }, line: { color: 'BF8B2E' } });
  txt(s, '12', { x: MX + 0.3, y: 1.8, w: 1.2, h: 0.95, fontSize: 54, bold: true, color: 'FFFFFF' });
  txt(s, 'RESPONSIBLE\nCONSUMPTION\nAND PRODUCTION', { x: MX + 1.45, y: 1.86, w: 1.8, h: 0.95, fontSize: 12, bold: true, color: 'FFFFFF' });
  txt(s, '∞', { x: MX, y: 3.0, w: 3.3, h: 1.6, align: 'center', fontSize: 80, color: 'FFFFFF' });
  txt(s, 'Kontribusi kami lahir langsung dari cara kami berproduksi, bukan dari klaim kemasan.', {
    x: 4.3, y: 1.6, w: 8.4, h: 0.7, fontFace: HEAD, fontSize: 18, bold: true, color: JADE });
  const pts = [
    ['Produksi sesuai pesanan pasti', 'Batch dibuat pada H-1 dari jumlah pre-order yang sudah masuk, jadi tidak ada stok berlebih yang berakhir dibuang.'],
    ['Belanja bahan seperlunya', 'Bahan baku dibeli grosir per batch sesuai kebutuhan, bukan ditimbun.'],
    ['Waste diukur, bukan diasumsikan', 'Target botol terbuang ≤ 2% per event, dicatat setiap event dan dipakai memperbaiki batch berikutnya.']];
  pts.forEach((p, i) => {
    const y = 2.55 + i * 1.12;
    s.addShape(pres.ShapeType.ellipse, { x: 4.3, y: y + 0.05, w: 0.42, h: 0.42, fill: { color: JADE }, line: { color: JADE } });
    txt(s, String(i + 1), { x: 4.3, y: y + 0.05, w: 0.42, h: 0.42, align: 'center', valign: 'middle', fontSize: 13, bold: true, color: GOLD });
    txt(s, p[0], { x: 4.95, y, w: 7.7, h: 0.35, fontSize: 14.5, bold: true, color: INK });
    txt(s, p[1], { x: 4.95, y: y + 0.37, w: 7.7, h: 0.6, fontSize: 12.5, color: MUTE });
  });
  s.addShape(pres.ShapeType.roundRect, { x: MX, y: 5.95, w: SW - 2 * MX, h: 0.9, rectRadius: 0.06, fill: { color: TINT }, line: { color: TINT } });
  txt(s, [{ text: 'Langkah berikutnya: ', options: { bold: true, color: JADE } },
           { text: 'botol dan cup saat ini masih kemasan sekali pakai. Program tukar botol atau kemasan yang bisa didaur ulang adalah perbaikan yang sedang kami pertimbangkan.' }],
      { x: MX + 0.3, y: 6.08, w: SW - 2 * MX - 0.6, h: 0.65, fontSize: 12.5, valign: 'middle', color: '2A3A31' });
  s.addNotes('Jujur soal kemasan: kontribusi kami di sisi produksi sesuai pesanan, kemasannya belum.');
}

/* ── 9. value proposition canvas ─────────────────────────── */
{
  const s = pres.addSlide(); s.background = { color: 'FFFFFF' };
  heading(s, 'Value Proposition Canvas', 'Sisi kiri yang kami tawarkan, sisi kanan yang pelanggan butuhkan');
  const SX = MX, SY = 1.55, SQ = 5.35, split = 2.15;
  s.addShape(pres.ShapeType.rect, { x: SX, y: SY, w: SQ, h: SQ, fill: { color: TINT }, line: { color: JADE, width: 1.5 } });
  s.addShape(pres.ShapeType.line, { x: SX + split, y: SY, w: 0, h: SQ, line: { color: JADE, width: 1.25 } });
  s.addShape(pres.ShapeType.line, { x: SX + split, y: SY + SQ / 2, w: SQ - split, h: 0, line: { color: JADE, width: 1.25 } });
  const sect = (x, y, w, h, head, items, c) => {
    txt(s, head, { x: x + 0.15, y: y + 0.12, w: w - 0.3, h: 0.3, fontSize: 12.5, bold: true, color: c });
    txt(s, bullets(items, 13.5, INK), { x: x + 0.15, y: y + 0.48, w: w - 0.3, h: h - 0.6, fontSize: 13.5, valign: 'top' });
  };
  sect(SX, SY, split, SQ, 'Products & Services', ['Aren Milk Tea', 'Jasmine Bean Curd Latte', 'Matcha Cloud', 'Pre-order lewat form', 'Jalur pickup khusus', 'Paket 5 & 10 botol'], JADE);
  sect(SX + split, SY, SQ - split, SQ / 2, 'Gain Creators', ['Gula aren asli & bean curd', 'Cup bercetak penuh', 'Bundling hemat untuk kelompok', 'Konfirmasi personal via WhatsApp'], TEAL);
  sect(SX + split, SY + SQ / 2, SQ - split, SQ / 2, 'Pain Relievers', ['Pickup tanpa antre walk-in', 'Rp13.000–15.000 per botol', 'Produksi H-1, dijaga dingin', 'Pesanan pasti tersedia'], CLAY);

  const R = 2.68, CX = 10.05, CY = SY + SQ / 2;
  s.addShape(pres.ShapeType.ellipse, { x: CX - R, y: CY - R, w: 2 * R, h: 2 * R, fill: { color: TINT }, line: { color: JADE, width: 1.5 } });
  const dx = R * Math.cos(Math.PI / 3), dy = R * Math.sin(Math.PI / 3);
  s.addShape(pres.ShapeType.line, { x: CX, y: CY - dy, w: dx, h: dy, flipV: true, line: { color: JADE, width: 1.25 } });
  s.addShape(pres.ShapeType.line, { x: CX, y: CY, w: dx, h: dy, line: { color: JADE, width: 1.25 } });
  s.addShape(pres.ShapeType.line, { x: CX - R, y: CY, w: R, h: 0, line: { color: JADE, width: 1.25 } });
  sect(CX - 2.05, CY - 2.05, 2.55, 2.0, 'Gains', ['Harga ramah kantong', 'Rasa premium', 'Cepat didapat', 'Kemasan layak konten'], TEAL);
  sect(CX - 2.05, CY + 0.12, 2.55, 2.0, 'Pains', ['Antre lama di booth', 'Harga Rp25.000+', 'Es cair, rasa berubah', 'Stok habis saat event'], CLAY);
  sect(CX + 0.45, CY - 0.95, 1.95, 2.1, 'Customer Jobs', ['Beli minum di sela kuliah', 'Pesan untuk satu kelompok', 'Punya minuman untuk difoto'], JADE);

  s.addShape(pres.ShapeType.line, { x: SX + SQ + 0.12, y: CY, w: CX - R - SX - SQ - 0.24, h: 0,
    line: { color: GOLD_D, width: 2.5, beginArrowType: 'triangle', endArrowType: 'triangle' } });
  txt(s, 'FIT', { x: SX + SQ, y: CY - 0.42, w: CX - R - SX - SQ, h: 0.3, align: 'center', fontSize: 11, bold: true, color: GOLD_D });
  s.addNotes('Setiap pain di kanan dijawab satu pain reliever di kiri: antre → jalur khusus, mahal → Rp13–15 ribu, es cair → produksi H-1 dijaga dingin.');
}

/* ── 10. STP ─────────────────────────────────────────────── */
{
  const s = pres.addSlide(); s.background = { color: 'FFFFFF' };
  heading(s, 'Segmentation · Targeting · Positioning');
  const segW = 3.0, segH = 2.45;
  [['Demografis', ['18–24 tahun', 'Mahasiswa, pria & wanita', 'Uang saku terbatas']],
   ['Geografis', ['Kampus Binus & sekitarnya', 'Venue event kampus']],
   ['Psikografis', ['Suka minuman kekinian', 'Aktif di media sosial', 'Mau tampil beda']],
   ['Behavioral', ['Beli minum saat event', 'Pesan berkelompok', 'Sensitif harga & waktu']]].forEach((c, i) =>
    card(s, MX + (i % 2) * (segW + 0.2), 1.55 + ((i / 2) | 0) * (segH + 0.2), segW, segH, c[0], c[1], { fs: 12 }));
  txt(s, 'SEGMENTATION', { x: MX, y: 6.75, w: 6.2, h: 0.3, align: 'center', fontSize: 11, bold: true, charSpacing: 2, color: MUTE });

  card(s, 7.05, 1.55, 5.68, 2.3, 'Targeting', [
    'Utama: mahasiswa Binus & pengunjung event kampus',
    'Kedua: panitia & himpunan yang memesan kolektif',
    'Berikutnya: event kampus lain, lalu pre-order mingguan'], { fs: 12.5 });
  s.addShape(pres.ShapeType.roundRect, { x: 7.05, y: 4.05, w: 5.68, h: 2.6, rectRadius: 0.07, fill: { color: JADE }, line: { color: JADE } });
  txt(s, 'Positioning', { x: 7.35, y: 4.2, w: 5, h: 0.3, fontSize: 13, bold: true, color: GOLD });
  txt(s, 'Untuk mahasiswa yang ingin minuman kekinian berasa premium tanpa antre dan tanpa mahal, Spill de Tea adalah milk tea botolan pre-order yang dibuat segar per batch dan diambil di jalur khusus — dengan harga sekitar setengah brand teh premium.', {
    x: 7.35, y: 4.6, w: 5.1, h: 1.9, fontFace: HEAD, fontSize: 15, italic: true, color: CREAM, valign: 'top', lineSpacingMultiple: 1.1 });
  s.addNotes('Positioning statement adalah ringkasan seluruh deck: siapa, apa, dan kenapa kami.');
}

/* ── 11. positioning map ─────────────────────────────────── */
{
  const s = pres.addSlide(); s.background = { color: 'FFFFFF' };
  heading(s, 'Brand Positioning Map', 'Harga per cup terhadap kesan brand');
  const GX = 1.6, GY = 1.6, GW = 9.6, GH = 5.0, mx = GX + GW / 2, my = GY + GH / 2;
  s.addShape(pres.ShapeType.rect, { x: GX, y: GY, w: GW, h: GH, fill: { color: 'FAF8F3' }, line: { color: LINE, width: 1 } });
  for (let i = 1; i < 8; i++) s.addShape(pres.ShapeType.line, { x: GX + GW * i / 8, y: GY, w: 0, h: GH, line: { color: 'ECE6DA', width: 0.5 } });
  for (let i = 1; i < 5; i++) s.addShape(pres.ShapeType.line, { x: GX, y: GY + GH * i / 5, w: GW, h: 0, line: { color: 'ECE6DA', width: 0.5 } });
  s.addShape(pres.ShapeType.line, { x: GX, y: my, w: GW, h: 0, line: { color: INK, width: 1.5, endArrowType: 'triangle' } });
  s.addShape(pres.ShapeType.line, { x: mx, y: GY, w: 0, h: GH, line: { color: INK, width: 1.5, beginArrowType: 'triangle' } });
  txt(s, 'Premium', { x: mx - 1, y: GY - 0.32, w: 2, h: 0.28, align: 'center', fontSize: 12, bold: true, color: INK });
  txt(s, 'Kasual', { x: mx - 1, y: GY + GH + 0.05, w: 2, h: 0.28, align: 'center', fontSize: 12, bold: true, color: INK });
  txt(s, 'Terjangkau', { x: GX - 1.05, y: my - 0.15, w: 1.0, h: 0.3, align: 'right', fontSize: 12, bold: true, color: INK });
  txt(s, 'Mahal', { x: GX + GW + 0.08, y: my - 0.15, w: 1.0, h: 0.3, fontSize: 12, bold: true, color: INK });
  const P = (px, py) => [GX + GW * px, GY + GH * (1 - py)];   // px: price 0→1, py: premium 0→1
  [['Chagee', 0.86, 0.86], ['Chatime', 0.72, 0.6], ['Kopi Kenangan', 0.58, 0.45], ['Haus!', 0.28, 0.22], ['Mixue', 0.12, 0.12]].forEach(b => {
    const [x, y] = P(b[1], b[2]);
    s.addShape(pres.ShapeType.ellipse, { x: x - 0.11, y: y - 0.11, w: 0.22, h: 0.22, fill: { color: '8E9A93' }, line: { color: 'FFFFFF', width: 1 } });
    txt(s, b[0], { x: x + 0.16, y: y - 0.16, w: 1.9, h: 0.32, fontSize: 12.5, color: '3B4A42' });
  });
  const [sx, sy] = P(0.24, 0.82);
  s.addShape(pres.ShapeType.ellipse, { x: sx - 0.48, y: sy - 0.48, w: 0.96, h: 0.96, fill: { color: CREAM }, line: { color: GOLD, width: 2 } });
  s.addImage({ path: EMBLEM, x: sx - 0.4, y: sy - 0.4, w: 0.8, h: 0.8 });
  txt(s, 'Spill de Tea', { x: sx + 0.58, y: sy - 0.2, w: 2.2, h: 0.4, fontFace: HEAD, fontSize: 17, bold: true, color: JADE });
  txt(s, 'Kuadran kosong yang kami isi: kesan premium di harga terjangkau.', {
    x: GX, y: SH - 0.62, w: GW, h: 0.3, align: 'center', fontSize: 11.5, italic: true, color: MUTE });
  txt(s, 'Posisi berdasarkan penilaian tim, belum survei.', { x: GX + GW - 4, y: GY + GH - 0.32, w: 3.9, h: 0.26, align: 'right', fontSize: 9.5, italic: true, color: MUTE });
  s.addNotes('Brand premium ada di kanan atas, brand murah di kiri bawah. Kiri atas — premium tapi terjangkau — masih kosong, di situ kami.');
}

/* ── 12. mind map ────────────────────────────────────────── */
{
  const s = pres.addSlide(); s.background = { color: 'FFFFFF' };
  heading(s, 'About Us — Mind Map');
  const CX = SW / 2, CY = 4.3, R0 = 0.95;
  const branches = [
    ['MENU', -1, -1, ['Aren Milk Tea', 'Jasmine Bean Curd Latte', 'Matcha Cloud'], JADE],
    ['SISTEM', 1, -1, ['Pre-order via form', 'Produksi batch H-1', 'Pickup jalur khusus'], TEAL],
    ['TARGET', -1, 1, ['Mahasiswa Binus', 'Panitia & himpunan', 'Pengunjung event'], GOLD_D],
    ['BRAND', 1, 1, ['Cup bercetak penuh', 'Emblem penjaga teh', 'Website & Instagram'], CLAY]];
  branches.forEach(b => {
    const nx = CX + b[1] * 2.15, ny = CY + b[2] * 1.45;
    s.addShape(pres.ShapeType.line, { x: Math.min(CX, nx), y: Math.min(CY, ny), w: Math.abs(nx - CX), h: Math.abs(ny - CY),
      flipV: b[1] * b[2] < 0, line: { color: b[4], width: 2.5 } });
    s.addShape(pres.ShapeType.roundRect, { x: nx - 0.9, y: ny - 0.32, w: 1.8, h: 0.64, rectRadius: 0.3, fill: { color: b[4] }, line: { color: b[4] } });
    txt(s, b[0], { x: nx - 0.9, y: ny - 0.32, w: 1.8, h: 0.64, align: 'center', valign: 'middle', fontSize: 14, bold: true, color: 'FFFFFF' });
    const lx = b[1] < 0 ? MX : SW - MX - 2.85;
    txt(s, bullets(b[3], 13, INK), { x: lx, y: ny - 0.55, w: 2.85, h: 1.15, valign: 'middle', align: 'left' });
  });
  s.addShape(pres.ShapeType.ellipse, { x: CX - R0, y: CY - R0, w: 2 * R0, h: 2 * R0, fill: { color: CREAM }, line: { color: GOLD, width: 2 } });
  s.addImage({ path: EMBLEM, x: CX - R0 + 0.12, y: CY - R0 + 0.12, w: 2 * R0 - 0.24, h: 2 * R0 - 0.24 });
  s.addNotes('Empat cabang ide bisnis: apa yang dijual, bagaimana sistemnya, untuk siapa, dan bagaimana brand-nya dikenali.');
}

/* ── 13. digital prototype ───────────────────────────────── */
{
  const s = pres.addSlide(); s.background = { color: 'FFFFFF' };
  heading(s, 'Prototype Digital', 'Website pre-order dan desain cup — sudah live');
  s.addShape(pres.ShapeType.rect, { x: MX, y: 1.6, w: 8.0, h: 5.0, fill: { color: JADE }, line: { color: LINE, width: 1 } });
  s.addImage({ path: img('site-hero.jpg'), x: MX, y: 1.6, w: 8.0, h: 5.0 });
  s.addShape(pres.ShapeType.rect, { x: 8.95, y: 1.6, w: 3.78, h: 5.0, fill: { color: JADE }, line: { color: JADE } });
  s.addImage({ path: img('cup3d.jpg'), x: 8.95, y: 1.6, w: 3.78, h: 5.0, sizing: { type: 'cover', w: 3.78, h: 5.0 } });
  txt(s, [{ text: 'Website: ', options: { bold: true, color: JADE } }, { text: 'beixuantianjun.github.io/spill-de-tea', options: { breakLine: true } },
          { text: 'Isi: cup 3D interaktif, menu, alur pre-order, dan form pre-order Jotform yang sudah menerima pesanan.' }],
      { x: MX, y: 6.72, w: 12.1, h: 0.6, fontSize: 12, color: '2A3A31' });
  s.addNotes('Prototype digital: situs live dengan form pre-order. Foto prototype fisik ditambahkan di slide berikutnya.');
}

/* ── 14. business model canvas ───────────────────────────── */
{
  const s = pres.addSlide(); s.background = { color: 'FFFFFF' };
  heading(s, 'Business Model Canvas');
  const X0 = MX, GAP = 0.12, W = SW - 2 * X0, CW = (W - 4 * GAP) / 5;
  const TOP = 1.3, TOPH = 4.0, BOTY = TOP + TOPH + GAP, BOTH = 7.1 - BOTY, halfH = (TOPH - GAP) / 2;
  const cx = i => X0 + i * (CW + GAP);
  const blk = (head, items, x, y, w, h, o = {}) => {
    const dark = !!o.dark, fs = o.fs || 9.5;
    s.addShape(pres.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.06, fill: { color: dark ? JADE : TINT }, line: { color: dark ? JADE : TINT } });
    s.addShape(pres.ShapeType.rect, { x: x + 0.14, y: y + 0.17, w: 0.1, h: 0.1, fill: { color: GOLD }, line: { color: GOLD } });
    txt(s, head, { x: x + 0.32, y: y + 0.09, w: w - 0.44, h: 0.26, fontSize: 11, bold: true, color: dark ? CREAM : JADE });
    txt(s, bullets(items, fs, dark ? 'E6EFE9' : INK), { x: x + 0.12, y: y + 0.42, w: w - 0.3, h: h - 0.5, valign: 'top' });
  };
  blk('Key Partners', ['Supplier grosir teh, susu bubuk & gula aren', 'Pemasok bahan bean curd', 'Vendor botol, label & sealer', 'Panitia / EO event & himpunan', 'Pembayaran QRIS', 'Jotform — form pre-order'], cx(0), TOP, CW, TOPH, { fs: 10 });
  blk('Key Activities', ['Rekap pre-order & rencana batch', 'Produksi batch H-1', 'QC rasa, suhu & kemasan', 'Operasi booth pickup', 'Konten & evaluasi event'], cx(1), TOP, CW, halfH);
  blk('Key Resources', ['Resep & SOP batch', 'Alat produksi & cooler', 'Tim mahasiswa TI', 'Brand: emblem, cup, website'], cx(1), TOP + halfH + GAP, CW, halfH);
  blk('Value Propositions', ['Rasa premium harga mahasiswa: Rp13.000–15.000 vs Rp25.000–38.000', 'Tanpa antre lewat jalur pickup khusus', 'Selalu fresh: produksi H-1', 'Gula aren asli & bean curd', 'Cup bercetak penuh', 'Zero waste — produksi sesuai pesanan'], cx(2), TOP, CW, TOPH, { dark: true, fs: 10.5 });
  blk('Customer Relationships', ['Komunitas Instagram', 'Paket 5 botol hemat Rp5.000, 10 botol gratis 1', 'Konfirmasi via WhatsApp', 'Kode order pickup'], cx(3), TOP, CW, halfH);
  blk('Channels', ['Website + form Jotform', 'Instagram & TikTok', 'WhatsApp', 'Booth pickup di event'], cx(3), TOP + halfH + GAP, CW, halfH);
  blk('Customer Segments', ['Mahasiswa Binus', 'Pengunjung event kampus', 'Panitia & himpunan', 'Event kampus lain (fase 2)', 'Pre-order mingguan (fase 3)'], cx(4), TOP, CW, TOPH, { fs: 10 });
  const BW = (W - GAP) / 2;
  blk('Cost Structure', ['Bahan baku ± Rp2.600/botol', 'Kemasan Rp1.500/botol → HPP ± Rp4.100', 'Promosi skala kecil · biaya booth bila ada', 'Tanpa sewa toko & pegawai harian'], X0, BOTY, BW, BOTH, { fs: 10 });
  blk('Revenue Streams', ['Aren Rp13.000 · Jasmine Bean Curd & Matcha Cloud Rp15.000', 'Paket bundling 5 & 10 botol', 'Deposit pre-order — kas masuk sebelum produksi', 'Margin kotor ± 68% per botol (Aren)'], X0 + BW + GAP, BOTY, BW, BOTH, { fs: 10 });
  s.addNotes('Baca dari segmen dan value proposition, lalu cara menjangkau, lalu yang di belakang layar.');
}

/* ── 15. revenue & cost simulation ───────────────────────── */
{
  const s = pres.addSlide(); s.background = { color: 'FFFFFF' };
  heading(s, 'Revenue Streams & Cost Structure', 'Simulasi laba kotor per fase, dari harga dan HPP di pitch deck');
  const fmt = n => 'Rp' + n.toLocaleString('id-ID');
  const PRICE = 13000, HPP = 4100;
  const phases = [['Pilot — Binus Festival', 500, 'sekali event'], ['Fase 2 — multi-event kampus', 2000, 'per bulan'], ['Fase 3 — pre-order reguler', 6000, 'per bulan']];
  const H = (t, al) => ({ text: t, options: { bold: true, color: 'FFFFFF', fill: { color: JADE }, align: al || 'left' } });
  const rows = [[H('Fase'), H('Botol', 'right'), H('Pendapatan', 'right'), H('HPP', 'right'), H('Laba kotor', 'right')]];
  phases.forEach(p => rows.push([
    { text: `${p[0]}\n`, options: { bold: true } },
    { text: `${p[1].toLocaleString('id-ID')}\n${p[2]}`, options: { align: 'right' } },
    { text: fmt(p[1] * PRICE), options: { align: 'right' } },
    { text: fmt(p[1] * HPP), options: { align: 'right', color: CLAY } },
    { text: fmt(p[1] * (PRICE - HPP)), options: { align: 'right', bold: true, color: JADE } }]));
  s.addTable(rows, { x: MX, y: 1.6, w: 7.55, colW: [2.35, 1.1, 1.4, 1.3, 1.4], fontFace: BODY, fontSize: 11.5, color: INK,
    border: { type: 'solid', color: LINE, pt: 0.75 }, margin: [6, 8, 6, 8], valign: 'middle', rowH: [0.42, 0.62, 0.62, 0.62] });

  const hpp = [['Teh bubuk premium', 800], ['Susu bubuk / creamer', 700], ['Gula aren cair', 500], ['Topping', 600], ['Botol + label + sealer', 1500]];
  const hrows = [[H('HPP per botol (Aren)'), H('Biaya', 'right')]];
  hpp.forEach(h => hrows.push([h[0], { text: fmt(h[1]), options: { align: 'right' } }]));
  hrows.push([{ text: 'Total HPP', options: { bold: true } }, { text: fmt(4100), options: { bold: true, align: 'right', color: JADE } }]);
  s.addTable(hrows, { x: MX, y: 4.25, w: 4.4, colW: [3.0, 1.4], fontFace: BODY, fontSize: 11, color: INK,
    border: { type: 'solid', color: LINE, pt: 0.75 }, margin: [4, 8, 4, 8], rowH: 0.3 });

  s.addShape(pres.ShapeType.roundRect, { x: MX + 4.6, y: 4.25, w: 2.95, h: 2.1, rectRadius: 0.07, fill: { color: JADE }, line: { color: JADE } });
  txt(s, '68%', { x: MX + 4.85, y: 4.45, w: 2.5, h: 0.8, fontFace: HEAD, fontSize: 40, bold: true, color: GOLD });
  txt(s, 'margin kotor per botol Aren\n(Rp8.900 dari Rp13.000)', { x: MX + 4.85, y: 5.3, w: 2.6, h: 0.8, fontSize: 11.5, color: 'D6E4DB' });

  s.addChart(pres.ChartType.bar, [
    { name: 'Pendapatan', labels: ['Pilot', 'Fase 2', 'Fase 3'], values: phases.map(p => p[1] * PRICE / 1e6) },
    { name: 'Laba kotor', labels: ['Pilot', 'Fase 2', 'Fase 3'], values: phases.map(p => p[1] * (PRICE - HPP) / 1e6) }
  ], { x: 8.45, y: 1.5, w: 4.3, h: 4.9, barDir: 'col', barGrouping: 'clustered', chartColors: ['B9C9C0', JADE],
       showValue: true, dataLabelPosition: 'outEnd', dataLabelFontSize: 10, dataLabelColor: INK, dataLabelFormatCode: '0.0',
       showTitle: true, title: 'Juta rupiah', titleFontSize: 11, titleColor: MUTE,
       showLegend: true, legendPos: 'b', legendFontSize: 10, catAxisLabelColor: INK, catAxisLabelFontSize: 11,
       valAxisHidden: true, valGridLine: { style: 'none' }, catGridLine: { style: 'none' } });

  txt(s, 'Dihitung dengan harga dan HPP Aren. Biaya tetap (promosi, booth) belum dikurangi, dan HPP menu Rp15.000 belum dihitung — angka ini laba kotor, bukan laba bersih.', {
    x: MX, y: 6.6, w: SW - 2 * MX, h: 0.5, fontSize: 11, italic: true, color: MUTE });
  s.addNotes('Semua angka dari pitch deck: Rp13.000, HPP Rp4.100, volume tiap fase. Tegaskan ini laba kotor.');
}

/* ── 16. contact ─────────────────────────────────────────── */
{
  const s = pres.addSlide(); s.background = { color: JADE };
  txt(s, 'Connect', { x: MX, y: 1.9, w: 6, h: 0.45, fontFace: HEAD, fontSize: 20, italic: true, color: GOLD });
  txt(s, 'Pesan Spill de Tea\nuntuk event\nberikutnya.', { x: MX, y: 2.45, w: 6.4, h: 2.9, fontFace: HEAD, fontSize: 44, bold: true, color: CREAM, lineSpacingMultiple: 1.0 });
  [['WEBSITE', 'beixuantianjun.github.io/spill-de-tea'], ['FORM PRE-ORDER', 'form.jotform.com/262514421222040']].forEach((c, i) => {
    const y = 2.55 + i * 1.35;
    txt(s, c[0], { x: 7.4, y, w: 5.3, h: 0.3, fontSize: 11, bold: true, charSpacing: 2, color: '8FA89A' });
    txt(s, c[1], { x: 7.4, y: y + 0.35, w: 5.3, h: 0.45, fontSize: 18, color: GOLD });
    s.addShape(pres.ShapeType.line, { x: 7.4, y: y + 0.95, w: 5.3, h: 0, line: { color: '2E5A4C', width: 1 } });
  });
  s.addShape(pres.ShapeType.ellipse, { x: SW - MX - 1.2, y: SH - 1.75, w: 1.2, h: 1.2, fill: { color: CREAM }, line: { color: GOLD, width: 1.25 } });
  s.addImage({ path: EMBLEM, x: SW - MX - 1.12, y: SH - 1.67, w: 1.04, h: 1.04 });
  s.addNotes('Penutup. Kontak Instagram dan WhatsApp ditambahkan setelah datanya ada.');
}

pres.writeFile({ fileName: path.join(__dirname, 'Spill_de_Tea_Creativity_Innovation.pptx') }).then(f => console.log('tersimpan:', f));
