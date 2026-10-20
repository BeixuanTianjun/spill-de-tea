# Presentasi

## Deck Creativity & Innovation

`Spill_de_Tea_Creativity_Innovation.pptx` — 16 slide mengikuti struktur contoh
kelas: overview, 3 user persona, empathy map, about & menu, SDGs, value
proposition canvas, STP, brand positioning map, mind map, prototype digital,
BMC, simulasi revenue & cost, dan kontak. Dibangun oleh `build-cni.js`; visual
produk di `img/` diambil dari website ini sendiri.

Belum ada di deck ini, menunggu data asli: hasil survei, slide CEO & tim, foto
prototype fisik, dan feedback grid.

## Deck BMC & Robust Design

`Spill_de_Tea_BMC_Robust_Design.pptx` — deck kelas Venture Creation Project, 5 slide:
cover, Business Model Canvas, P-diagram, identifikasi faktor robust design, dan
target output response. Tiap slide punya speaker notes.

Deck ini dibangun oleh `build-deck.js`. Untuk mengubah isinya, edit skrip itu lalu
bangun ulang:

```bash
cd presentasi
npm install pptxgenjs
node build-deck.js   # deck BMC & robust design
node build-cni.js    # deck Creativity & Innovation
```

Emblem diambil dari `../assets/emblem.png`, jadi skrip harus dijalankan dari dalam repo ini.
