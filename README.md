# Dashboard IT

Dashboard IT adalah sistem manajemen arsip dan link berbasis Google Apps Script dengan penyimpanan file terstruktur berdasarkan kategori Google Drive.

## 🔗 Akses Deploy
Gunakan link berikut untuk mengakses aplikasi:

https://script.google.com/macros/s/AKfycbxvttNyVrtp-JsJYPVZDXwvGnPR5j4dTCVDSOYE0uCdP6a2liwFq2y0oSQvou__1bDq/exec

## ✨ Fitur Utama
- Upload file ke Google Drive berdasarkan kategori
- Preview dan download file dengan nama asli lengkap dan ekstensi
- Manajemen kategori arsip dan kategori dashboard
- Hapus paksa kategori beserta semua isinya
- Loading states dan notifikasi success/error untuk semua proses
- Diagnostic tools untuk memeriksa akses Drive dan memperbaiki record hilang

## 🚀 Setup Singkat
1. Buka Google Sheets dan Apps Script
2. Atur `MAIN_FOLDER_ID` di `kode.gs`
3. Jalankan fungsi `setupDatabase()`
4. Jalankan `initializeCategoryFolders()`
5. Gunakan menu admin untuk tambah/hapus/ubah kategori

## 📝 Dokumen Tambahan
Lihat `SETUP_GUIDE.md` untuk panduan lengkap instalasi dan troubleshooting.

## 📦 Versi
`v2.4`

## 📌 Catatan
- Pastikan Apps Script memiliki izin Google Drive penuh
- Upload menggunakan nama file asli di Drive; kolom input hanya berfungsi sebagai deskripsi tambahan
- Gunakan deploy link di atas jika ingin akses langsung dari URL eksternal
