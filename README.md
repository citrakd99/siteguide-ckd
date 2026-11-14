# SiteGuide CKD

Panduan standar lapangan untuk SPV CSA, SPV MEP, dan QC (PWA, bisa dipasang di layar utama HP dan jalan offline).

Kontraktor: PT Citra Karsa Dinamika.

## Isi
- `index.html`: aplikasi (checklist per pekerjaan, NCR, alat uji, rekap, panduan)
- `manifest.json`, `sw.js`, `icon-*.png`: komponen PWA

## Menayangkan lewat GitHub Pages
Settings > Pages > Build and deployment > Source: Deploy from a branch > Branch: `main`, folder `/ (root)` > Save.
Pages dari repo privat butuh paket GitHub berbayar. Pada paket gratis, ubah repo ke Public lebih dulu.

## Penyimpanan terpusat (Google Sheets)
1. Buat Google Sheet baru. Buka Extensions > Apps Script, tempel isi `apps-script/Code.gs`, ganti `TOKEN`.
2. Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone. Salin alamat `/exec`.
3. Di aplikasi: Rekap > Sinkronisasi Google Sheets. Tempel alamat dan kode rahasia, ketuk Simpan dan tes.
4. Data masuk ke tab Log (semua kejadian), NCR (satu baris per NCR, status diperbarui), dan Rekap (laporan harian).
Setiap perubahan kode Apps Script perlu deployment baru (Manage deployments > Edit > New version).

## Catatan
- Data tetap disimpan di browser tiap perangkat. Sinkronisasi ke Sheet bersifat opsional dan hanya menambah baris.
- Nilai ambang uji ditulis "spek proyek": isi dari spesifikasi kontrak.
- Nomor dan edisi standar perlu dicocokkan dengan dokumen resmi.
