# 📋 Setup Guide - Dashboard IT dengan Folder Kategori

> **Akses langsung:** https://script.google.com/macros/s/AKfycbwYSxxKqQUfCmDtNzk97XxG4KAaIxZSaiTINgSoJxZrOqbNPSHnq1QgptCC6QEmJS_v/exec

## ⚙️ Konfigurasi Awal

### 1. **Dapatkan Folder ID dari Google Drive**
- Buka [Google Drive](https://drive.google.com)
- Buat folder untuk dashboard (misal: "Dashboard IT Archive")
- Klik folder tersebut → Copas **ID dari URL** (bagian setelah `/folders/`)
  - Contoh URL: `https://drive.google.com/drive/folders/1TImiBi0slgDECtrLY6vIqnczOCaOxjrO`
  - **Folder ID**: `1TImiBi0slgDECtrLY6vIqnczOCaOxjrO`

### 2. **Set Folder ID di kode.gs**
```javascript
const MAIN_FOLDER_ID = 'GANTI_DENGAN_ID_FOLDER_ANDA';
```

### 3. **Jalankan Setup Database**
- Buka **Apps Script** (dari spreadsheet Google Sheets)
- Jalankan fungsi: `setupDatabase()`
- Tunggu selesai

### 4. **Inisialisasi Folder Kategori**
- Jalankan fungsi: `initializeCategoryFolders()`
- Fungsi ini akan membuat folder untuk semua kategori yang sudah ada

---

## 📁 Struktur Folder Otomatis

Ketika setup berhasil, struktur folder akan seperti ini:

```
📦 MAIN_FOLDER (Dashboard IT Archive)
├── 📂 IM
│   ├── 📄 file1.pdf
│   └── 📄 file2.docx
├── 📂 IK
│   └── 📄 report.xlsx
├── 📂 Form
│   └── 📄 form_template.pdf
└── 📂 Memo
    └── 📄 memo.txt
```

---

## ➕ Cara Menambah Kategori Baru

### Di Admin Panel:
1. Buka "Manajemen Kategori"
2. Klik "Tambah Kategori"
3. Input nama kategori (misal: "Surat")
4. Klik "Simpan"

### Sistem otomatis akan:
✅ Mencatat kategori di sheet `Categories`
✅ **Membuat folder baru** di Google Drive dengan nama sama
✅ Siap untuk upload file ke kategori tersebut

---

## 📤 Upload File

1. Buka "Upload Arsip"
2. Pilih file yang ingin diupload
3. **Pilih kategori** dari dropdown (misal: "IM")
4. Isi deskripsi (optional)
5. Klik "Upload ke Drive"

### Yang terjadi otomatis:
- File disimpan ke folder **IM** di Google Drive
- Catatan disimpan di sheet `Archives`
- File dapat diakses via link

---

## 🔧 Troubleshooting

### Error: "Gagal upload: Exception: Akses ditolak: DriveApp"
**Gejala:** File terupload ke Drive tapi tidak tercatat di spreadsheet
**Penyebab:** Permission issue atau error saat setSharing/record

**Solusi:**
1. **Jalankan Diagnosa Sistem** (tombol 🔧 Diagnosa di halaman upload)
2. Cek **Execution logs** di Apps Script untuk detail error
3. Pastikan folder ID benar dan dapat diakses
4. Jika masih error, coba:
   - Reload halaman dashboard
   - Login ulang
   - Jalankan `testDriveAccess()` di Apps Script

### File Upload tapi Tidak Tercatat
**Solusi Otomatis:**
- Klik tombol **🔧 Diagnosa** di halaman upload
- Sistem akan otomatis mencari file yang belum tercatat dan memperbaikinya

**Solusi Manual:**
```javascript
// Jalankan di Apps Script
repairMissingArchives(); // Perbaiki record yang hilang
```

### Error: "Akses ditolak" Saat Membuat Folder
**Solusi:**
- Pastikan Apps Script memiliki permission penuh ke folder
- Buka folder di Google Drive → Bagikan → Pastikan email Apps Script ada
- Atau gunakan folder yang dibuat oleh akun yang sama dengan Apps Script

### Permission Google Drive API
**Jika masih bermasalah:**
1. Buka [Google Cloud Console](https://console.cloud.google.com)
2. Pilih project Apps Script
3. Enable **Google Drive API**
4. Pastikan OAuth consent screen sudah setup

---

## 🔄 Update Kategori

Jika ingin mengganti nama kategori:
1. Edit di "Manajemen Kategori"
2. Sistem akan:
   - Update nama di sheet `Categories`
   - Update nama di sheet `Links` dan `Archives`
   - ⚠️ Folder lama tetap ada, buat folder baru dengan nama baru (manual jika diperlukan)

---

## � Diagnostic Tools

### Tombol Diagnosa (Admin Only)
- **Lokasi:** Halaman Upload Arsip → Tombol orange 🔧 Diagnosa
- **Fungsi:**
  - Test akses Google Drive
  - Perbaiki record arsip yang hilang
  - **Loading indicator** saat proses berlangsung
  - Refresh daftar arsip otomatis setelah selesai

### Manual Testing di Apps Script
```javascript
// Test Drive access
testDriveAccess();

// Perbaiki record hilang
repairMissingArchives();

// Setup ulang folder kategori
initializeCategoryFolders();
```

### Kelola Kategori Arsip (Admin Only)
- **Tombol:** "Kelola Kategori" di halaman Upload Arsip
- **Fitur:**
  - ✅ **Tambah kategori baru** dengan input field
  - ✅ **Hapus kategori** yang tidak sedang digunakan
  - ✅ **Validasi otomatis** - kategori yang masih digunakan file arsip tidak bisa dihapus
  - ✅ **Folder management** - folder kosong otomatis dihapus saat kategori dihapus

---

## 📊 Monitoring & Logs

### Melihat Execution Logs
1. Buka Apps Script editor
2. Klik **Executions** (ikon clock)
3. Klik execution terbaru untuk melihat logs
4. Cari error messages untuk debugging

### Common Log Messages
- `File uploaded: [nama] ke kategori: [kategori]` → Upload berhasil
- `Error upload: [error]` → Upload gagal
- `Category folder ready: [nama]` → Folder kategori OK
- `Warning: Failed to set sharing` → Sharing gagal tapi file OK

---

## 🚨 Emergency Recovery

Jika sistem rusak parah:

1. **Backup spreadsheet** (Download as Excel)
2. **Jalankan setup ulang:**
   ```javascript
   setupDatabase();
   initializeCategoryFolders();
   repairMissingArchives();
   ```
3. **Test upload** dengan file kecil
4. **Periksa folder Drive** untuk file yang terupload tapi tidak tercatat

---

## �💡 Tips

- **Selalu test upload** setelah perubahan kode
- **Monitor Execution logs** secara berkala
- **Backup spreadsheet** sebelum upgrade besar
- **Gunakan folder Drive dedicated** untuk dashboard
- **Pastikan permission consistent** antara Apps Script dan Drive
- **Jangan hapus folder kategori manual** di Google Drive, gunakan sistem admin
- **Test upload file** setelah setup untuk memastikan semua berfungsi

---

## 📞 Support

Jika ada masalah:
1. Cek file `kode.gs` di Apps Script
2. Buka **Execution logs** untuk melihat error detail
3. Jalankan `setupDatabase()` dan `initializeCategoryFolders()` kembali

---

**Last Updated:** April 2026
**Version:** 2.3 (Bug fixes, preview/download, force delete, loading states, deploy link)
