# SiteGuide CKD

Panduan standar lapangan untuk SPV CSA, SPV MEP, dan QC (PWA, bisa dipasang di layar utama HP dan jalan offline).

Proyek: Batik Project - Safran. Kontraktor: PT Citra Karsa Dinamika.

## Isi
- `index.html`: aplikasi (checklist per pekerjaan, NCR, alat uji, rekap, panduan)
- `manifest.json`, `sw.js`, `icon-*.png`: komponen PWA

## Menayangkan lewat GitHub Pages
Settings > Pages > Build and deployment > Source: Deploy from a branch > Branch: `main`, folder `/ (root)` > Save.
Pages dari repo privat butuh paket GitHub berbayar. Pada paket gratis, ubah repo ke Public lebih dulu.

## Catatan
- Data tersimpan di browser tiap perangkat (localStorage). Belum ada sinkronisasi terpusat.
- Nilai ambang uji ditulis "spek proyek": isi dari spesifikasi kontrak.
- Nomor dan edisi standar perlu dicocokkan dengan dokumen resmi.
