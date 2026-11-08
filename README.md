# HASHIRASAI Community Portal

Project ini memisahkan pengelolaan source code dari layanan backend: GitHub untuk version control/source hosting, Google Apps Script untuk menjalankan aplikasi dan mengakses Google Sheets/Drive.

## Struktur

- `src/index.html` — frontend portal (sidebar responsif sudah diperbaiki).
- `src/Code.gs` — backend utama Apps Script dari file utama yang diberikan, termasuk fungsi profil dan ubah password.
- `src/appsscript.json` — konfigurasi project Apps Script.
- `patches/ProfilePatch.gs` — salinan patch lama sebagai arsip/referensi. **Jangan tambahkan file ini lagi ke project Apps Script**, karena `updateMemberProfile`, `changeMemberPassword`, dan helper profil sudah ada di `src/Code.gs`.

## Menjalankan tetap melalui Google Apps Script

1. Buka project Apps Script yang terikat pada spreadsheet HASHIRASAI.
2. Cadangkan project dan spreadsheet terlebih dahulu.
3. Salin isi `src/Code.gs` ke file backend utama `Code.gs` di Apps Script (atau bandingkan lalu terapkan perubahan). Jangan membuat salinan kedua dari fungsi-fungsi yang sama.
4. Pastikan `src/index.html` menggantikan/menyelaraskan file HTML bernama `index` di Apps Script.
5. Jalankan `setupDatabase()` satu kali dari editor Apps Script, lalu tinjau izin akses yang diminta. Fungsi ini menyiapkan/migrasikan header sheet; tetap buat backup sebelum menjalankannya.
6. Deploy sebagai Web App dari Apps Script. Gunakan URL deployment Apps Script untuk pengguna.

**Penting:** Jangan membuka `index.html` langsung dari GitHub Pages dengan harapan `google.script.run` berfungsi. API `google.script.run`, `SpreadsheetApp`, dan `DriveApp` berjalan dalam lingkungan Google Apps Script HTML Service. GitHub Pages tidak menyediakan runtime tersebut.

## Catatan integrasi

- Kolom `USERS.FOTO_URL` didukung oleh backend. Migrasi header diperbaiki agar kolom `FOTO_URL` tetap ditambahkan walaupun `STATUS_AKUN` sudah ada.
- Foto profil disimpan ke Google Drive dan backend membagikan file dengan akses `ANYONE_WITH_LINK` agar thumbnail dapat ditampilkan. Periksa kebijakan privasi komunitas sebelum mengaktifkan fitur foto profil.
- Perubahan username dicek agar tidak duplikat; perubahan password memerlukan password lama.
- Backend menggunakan `SpreadsheetApp.getActiveSpreadsheet()`. Pastikan project tetap bound ke spreadsheet yang benar.

## Sidebar

Perbaikan sidebar berada di `src/index.html`: drawer mobile, backdrop, state expanded, lebar desktop/collapsed yang konsisten, fokus keyboard, dan shortcut pencarian menu Ctrl/⌘+K.

## GitHub

Commit dan push folder project ini ke repository GitHub privat bila source mengandung konfigurasi atau detail internal. Jangan pernah memasukkan password, token deployment, atau data anggota ke repository.
