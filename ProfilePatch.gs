/* =====================================================================
   HASHIRASAI — PATCH Code.gs  (Pengaturan Akun: profil, foto, password)
   ---------------------------------------------------------------------
   A) Ubah 4 tempat kecil di kode lama (lihat di bawah).
   B) Tempel semua fungsi di bagian "FUNGSI BARU" ke bawah file Code.gs.
   C) Deploy > Manage deployments > Edit > New version.
   ===================================================================== */


/* ---------- A1. HEADERS.USERS: tambah kolom FOTO_URL di akhir ----------
   Ubah:
     USERS: ["USER_ID", ..., "JABATAN", "STATUS_AKUN"],
   Menjadi:
     USERS: ["USER_ID", "NAMA_IC", "DISCORD_USERNAME", "DISCORD_ID", "PASSWORD_HASH",
             "SALT", "ROLE", "STATUS", "TANGGAL_DAFTAR", "JABATAN", "STATUS_AKUN", "FOTO_URL"],
*/

/* ---------- A2. setupDatabase(): di dalam  else if (key === "USERS") { ... }
   TEMPEL tepat sebelum kurung tutup blok tersebut (setelah blok STATUS_AKUN):

      const headersNow = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1))
        .getDisplayValues()[0].map(v => String(v || "").trim().toUpperCase());
      if (!headersNow.includes("FOTO_URL")) {
        sheet.getRange(1, sheet.getLastColumn() + 1).setValue("FOTO_URL");
      }
   (Opsional: fungsi baru di bawah juga membuat kolom ini otomatis bila belum ada.)
*/

/* ---------- A3. loginMember(): pada objek  member: { ... }  tambahkan baris:
        fotoUrl: user[11] || "",
*/

/* ---------- A4. validateMemberSession(): pada objek  member: { ... }  tambahkan baris:
        fotoUrl: user[11] || "",
*/


/* =====================================================================
   FUNGSI BARU  (tempel di bagian bawah Code.gs)
   ===================================================================== */

// Pastikan kolom ada di sheet USERS; kembalikan nomor kolom (1-based).
function _ensureUserColumn(sheet, name) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0]
    .map(v => String(v || "").trim().toUpperCase());
  const idx = headers.indexOf(name);
  if (idx >= 0) return idx + 1;
  const col = lastCol + 1;
  sheet.getRange(1, col).setValue(name);
  return col;
}

// Cegah teks yang diawali = + - @ dibaca sebagai formula oleh Sheets.
function _safeCellText(value) {
  const text = String(value || "");
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

// Simpan foto profil (JPEG) ke Drive, bagikan via link, hapus foto lama.
function _saveProfilePhoto(photoData, userID, oldUrl) {
  const bytes = Utilities.base64Decode(String(photoData));
  if (bytes.length < 4 || bytes[0] !== -1 || bytes[1] !== -40) {
    throw new Error("Foto harus berformat JPEG/gambar yang valid.");
  }
  if (bytes.length > 1024 * 1024) throw new Error("Ukuran foto profil maksimal 1 MB.");

  const folderName = "HASHIRASAI_FOTO_PROFIL";
  const folders = DriveApp.getFoldersByName(folderName);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
  const blob = Utilities.newBlob(bytes, "image/jpeg", "PROFIL_" + userID + "_" + Date.now() + ".jpg");
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const old = String(oldUrl || "").match(/[?&]id=([^&]+)/);
  if (old) {
    try { DriveApp.getFileById(old[1]).setTrashed(true); } catch (e) { /* abaikan */ }
  }
  return "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w256";
}

/**
 * Ubah username, nama IC, dan foto profil milik anggota yang sedang login.
 * data = { token, username, namaIC, photoData (base64 JPEG, opsional) }
 */
function updateMemberProfile(data) {
  const lock = LockService.getScriptLock();
  try {
    if (!data) throw new Error("Data profil tidak ditemukan.");
    const validation = validateMemberSession(String(data.token || ""));
    if (!validation.success) throw new Error(validation.message || "Sesi tidak valid.");
    const member = validation.member;

    const username = String(data.username || "").trim();
    const namaIC = String(data.namaIC || "").trim();
    if (namaIC.length < 3) throw new Error("Nama IC minimal 3 karakter.");
    if (username.length < 2 || username.length > 64) {
      throw new Error("Username harus terdiri dari 2 sampai 64 karakter.");
    }
    if (!lock.tryLock(10000)) throw new Error("Sistem sedang sibuk. Coba lagi.");

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.USERS);
    if (!sheet) throw new Error("Sheet USERS tidak ditemukan.");
    const fotoCol = _ensureUserColumn(sheet, "FOTO_URL");

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error("Data anggota tidak ditemukan.");
    const rows = sheet.getRange(2, 1, lastRow - 1, Math.max(11, fotoCol, sheet.getLastColumn())).getValues();
    const idx = rows.findIndex(r => String(r[0]) === String(member.userID));
    if (idx < 0) throw new Error("Anggota tidak ditemukan.");

    const taken = rows.some((r, i) => i !== idx &&
      String(r[2] || "").trim().toLowerCase() === username.toLowerCase());
    if (taken) throw new Error("Username sudah dipakai anggota lain.");

    let fotoUrl = String(rows[idx][fotoCol - 1] || "");
    if (data.photoData) fotoUrl = _saveProfilePhoto(data.photoData, member.userID, fotoUrl);

    const rowNumber = idx + 2;
    sheet.getRange(rowNumber, 2).setValue(_safeCellText(namaIC));
    sheet.getRange(rowNumber, 3).setValue(_safeCellText(username));
    if (data.photoData) sheet.getRange(rowNumber, fotoCol).setValue(fotoUrl);

    _writeMemberLog(member.userID, "UBAH PROFIL",
      "Memperbarui profil" + (data.photoData ? " dan foto." : "."));

    return {
      success: true,
      message: "Profil berhasil diperbarui.",
      member: Object.assign({}, member, {
        namaIC: namaIC,
        discordUsername: username,
        fotoUrl: fotoUrl
      })
    };
  } catch (error) {
    return { success: false, message: error.message };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

/**
 * Ubah password anggota yang sedang login (wajib password lama).
 */
function changeMemberPassword(token, oldPassword, newPassword) {
  const lock = LockService.getScriptLock();
  try {
    const validation = validateMemberSession(String(token || ""));
    if (!validation.success) throw new Error(validation.message || "Sesi tidak valid.");
    const member = validation.member;

    const oldPass = String(oldPassword || "");
    const newPass = String(newPassword || "");
    if (!oldPass) throw new Error("Password saat ini wajib diisi.");
    if (newPass.length < 8) throw new Error("Password baru minimal 8 karakter.");
    if (newPass === oldPass) throw new Error("Password baru harus berbeda dari password lama.");
    if (!lock.tryLock(10000)) throw new Error("Sistem sedang sibuk. Coba lagi.");

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.USERS);
    if (!sheet) throw new Error("Sheet USERS tidak ditemukan.");
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error("Data anggota tidak ditemukan.");
    const rows = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    const idx = rows.findIndex(r => String(r[0]) === String(member.userID));
    if (idx < 0) throw new Error("Anggota tidak ditemukan.");

    const row = rows[idx];
    if (_hashMemberPassword(oldPass, String(row[5])) !== String(row[4])) {
      throw new Error("Password saat ini salah.");
    }

    const salt = Utilities.getUuid().replace(/-/g, "");
    sheet.getRange(idx + 2, 5, 1, 2).setValues([[_hashMemberPassword(newPass, salt), salt]]);
    _writeMemberLog(member.userID, "UBAH PASSWORD", "Anggota mengubah password akun.");

    return { success: true, message: "Password berhasil diubah." };
  } catch (error) {
    return { success: false, message: error.message };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}
