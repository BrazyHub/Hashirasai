const CONFIG = {
  SPREADSHEET_ID: SpreadsheetApp.getActiveSpreadsheet().getId(),

  SHEETS: {
    USERS: "USERS",
    ABSENSI: "ABSENSI",
    LOG: "LOG_AKTIVITAS",
    MISSION: "MISSION",
    MISSION_PROGRESS: "MISSION_PROGRESS",
    EVENT: "EVENT",
    EVENT_PESERTA: "EVENT_PESERTA",
    TARGET_SETORAN: "TARGET_SETORAN",
    SETORAN: "SETORAN"
  }
};

const HEADERS = {
  USERS: ["USER_ID", "NAMA_IC", "DISCORD_USERNAME", "DISCORD_ID", "PASSWORD_HASH",
             "SALT", "ROLE", "STATUS", "TANGGAL_DAFTAR", "JABATAN", "STATUS_AKUN", "FOTO_URL"],
  ABSENSI: [
    "ABSENSI_ID", "USER_ID", "KEGIATAN_ID", "CHECK_IN", "CHECK_OUT", "DURASI", "STATUS", "FOTO_URL", "KETERANGAN_IZIN"
  ],
  LOG: ["LOG_ID", "USER_ID", "AKTIVITAS", "WAKTU", "KETERANGAN"],
  MISSION: ["MISSION_ID", "JUDUL", "DESKRIPSI", "TIPE", "REWARD", "STATUS", "DIBUAT_OLEH", "TANGGAL_DIBUAT"],
  MISSION_PROGRESS: ["PROGRESS_ID", "MISSION_ID", "USER_ID", "PERIODE", "WAKTU"],
  EVENT: ["EVENT_ID", "NAMA_EVENT", "JENIS", "TANGGAL", "JAM_MULAI", "JAM_SELESAI", "LOKASI", "DESKRIPSI", "STATUS", "DIBUAT_OLEH", "TANGGAL_DIBUAT"],
  EVENT_PESERTA: ["PESERTA_ID", "EVENT_ID", "USER_ID", "WAKTU"],
  TARGET_SETORAN: ["TARGET_ID", "JUDUL", "NOMINAL_TARGET", "PERIODE_MULAI", "PERIODE_SELESAI", "KETERANGAN", "STATUS", "DIBUAT_OLEH", "TANGGAL_DIBUAT"],
  SETORAN: ["SETORAN_ID", "TARGET_ID", "USER_ID", "NOMINAL", "CATATAN", "DICATAT_OLEH", "WAKTU"]
};

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Object.keys(CONFIG.SHEETS).forEach(key => {
    const sheetName = CONFIG.SHEETS[key];
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }

    const headers = HEADERS[key];

    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    } else if (key === "ABSENSI") {
      const current = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getDisplayValues()[0];
      headers.forEach((header, index) => {
        if (!current[index]) sheet.getRange(1, index + 1).setValue(header);
      });
    } else if (key === "USERS") {
      const current = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getDisplayValues()[0];
      // Migrasi aman: ubah nama kolom DIVISI menjadi JABATAN pada kolom yang sama.
      const oldIndex = current.findIndex(value => String(value).trim().toUpperCase() === "DIVISI");
      const newIndex = current.findIndex(value => String(value).trim().toUpperCase() === "JABATAN");
      if (oldIndex >= 0 && newIndex < 0) {
        sheet.getRange(1, oldIndex + 1).setValue("JABATAN");
      } else if (newIndex < 0) {
        sheet.getRange(1, sheet.getLastColumn() + 1).setValue("JABATAN");
      }

      let headersNow = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1))
        .getDisplayValues()[0].map(value => String(value || "").trim().toUpperCase());

      // Tambahkan STATUS_AKUN secara independen agar migrasi FOTO_URL tetap berjalan.
      if (!headersNow.includes("STATUS_AKUN")) {
        const statusAccountCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, statusAccountCol).setValue("STATUS_AKUN");
        if (sheet.getLastRow() > 1) {
          const defaults = Array.from({ length: sheet.getLastRow() - 1 }, () => ["ACTIVE"]);
          sheet.getRange(2, statusAccountCol, defaults.length, 1).setValues(defaults);
        }
      }

      headersNow = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1))
        .getDisplayValues()[0].map(value => String(value || "").trim().toUpperCase());
      if (!headersNow.includes("FOTO_URL")) {
        sheet.getRange(1, sheet.getLastColumn() + 1).setValue("FOTO_URL");
      }
    }

    sheet.setFrozenRows(1);

    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#202020")
      .setFontColor("#D4AF37");

    sheet.autoResizeColumns(1, headers.length);
  });

  return "Database HASHIRASAI berhasil diperiksa.";
}

function getDatabaseStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  return Object.keys(CONFIG.SHEETS).map(key => {
    const sheet = ss.getSheetByName(CONFIG.SHEETS[key]);

    return {
      sheet: CONFIG.SHEETS[key],
      exists: !!sheet,
      rows: sheet ? sheet.getLastRow() : 0
    };
  });
}


/**
 * PENDAFTARAN ANGGOTA HASHIRASAI
 * Status awal: PENDING
 * Role awal: MEMBER
 */

function registerMember(data) {
  const lock = LockService.getScriptLock();

  try {
    if (!data) throw new Error("Data pendaftaran tidak ditemukan.");

    const namaIC = String(data.namaIC || "").trim();
    const discordUsername = String(data.discordUsername || "").trim();
    const password = String(data.password || "");

    if (!namaIC || !discordUsername || !password) {
      throw new Error("Nama IC, username Discord, dan password wajib diisi.");
    }
    if (namaIC.length < 3) throw new Error("Nama IC minimal 3 karakter.");
    if (discordUsername.length < 2 || discordUsername.length > 64) {
      throw new Error("Username Discord harus terdiri dari 2 sampai 64 karakter.");
    }
    if (password.length < 8) throw new Error("Password minimal 8 karakter.");

    if (!lock.tryLock(10000)) throw new Error("Sistem sedang sibuk. Silakan coba lagi.");

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.USERS);
    if (!sheet) throw new Error("Sheet USERS tidak ditemukan.");

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const existingUsernames = sheet.getRange(2, 3, lastRow - 1, 1)
        .getDisplayValues().flat().map(value => String(value).trim().toLowerCase());
      if (existingUsernames.includes(discordUsername.toLowerCase())) {
        throw new Error("Username Discord sudah terdaftar. Gunakan username lain atau hubungi admin.");
      }
    }

    const userID = "USR-" + Utilities.getUuid();
    const salt = Utilities.getUuid().replace(/-/g, "");
    const passwordHash = _hashMemberPassword(password, salt);

    // Kolom DISCORD_ID dipertahankan untuk kompatibilitas data lama, tetapi kosong untuk pendaftar baru.
    sheet.appendRow([
      userID, namaIC, discordUsername, "", passwordHash, salt,
      "MEMBER", "PENDING", new Date(), "", "ACTIVE"
    ]);

    _writeMemberLog(userID, "PENDAFTARAN", "Pendaftaran anggota baru menunggu persetujuan admin.");
    return { success: true, message: "Pendaftaran berhasil. Silakan tunggu persetujuan admin." };

  } catch (error) {
    return { success: false, message: error.message };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

/**
 * Hash password dengan salt.
 */
function _hashMemberPassword(password, salt) {
  let result = salt + ":" + password;

  for (let i = 0; i < 1000; i++) {
    const digest = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      result,
      Utilities.Charset.UTF_8
    );

    result = Utilities.base64Encode(digest);
  }

  return result;
}


/**
 * Mencatat aktivitas anggota.
 */
function _writeMemberLog(userID, aktivitas, keterangan) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEETS.LOG);

  if (!sheet) {
    throw new Error("Sheet LOG_AKTIVITAS tidak ditemukan.");
  }

  sheet.appendRow([
    "LOG-" + Utilities.getUuid(),
    userID,
    aktivitas,
    new Date(),
    keterangan
  ]);
}


/**
 * LOGIN ANGGOTA HASHIRASAI
 */
function loginMember(data) {
  try {
    const discordUsername = String(data?.discordUsername || "").trim();
    const password = String(data?.password || "");

    if (!discordUsername || !password) {
      throw new Error("Username Discord dan password wajib diisi.");
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.USERS);
    if (!sheet) throw new Error("Sheet USERS tidak ditemukan.");
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error("Belum ada anggota terdaftar.");

    const rows = sheet.getRange(2, 1, lastRow - 1, Math.max(11, sheet.getLastColumn())).getValues();
    const normalizedUsername = discordUsername.toLowerCase();
    const user = rows.find(row => String(row[2] || "").trim().toLowerCase() === normalizedUsername);

    if (!user) throw new Error("Username Discord atau password salah.");

    const userID = user[0];
    const passwordHash = user[4];
    const salt = user[5];
    const role = user[6];
    const status = String(user[7]).toUpperCase();
    const accountStatus = String(user[10] || "ACTIVE").toUpperCase();
    const inputHash = _hashMemberPassword(password, salt);

    if (inputHash !== passwordHash) throw new Error("Username Discord atau password salah.");
    if (status === "PENDING") throw new Error("Pendaftaran kamu masih menunggu persetujuan admin.");
    if (status !== "APPROVED") throw new Error("Akun kamu belum diizinkan mengakses portal.");
    if (accountStatus !== "ACTIVE") throw new Error("Akun kamu sedang TIDAK AKTIF. Hubungi ADMIN untuk mengaktifkannya kembali.");

    const token = Utilities.getUuid() + Utilities.getUuid();
    const rememberMe = data?.rememberMe === true;
    const sessionDuration = rememberMe
      ? 30 * 24 * 60 * 60 * 1000   // Ingat Saya: 30 hari
      : 12 * 60 * 60 * 1000;      // Sesi biasa: 12 jam
    const sessionRecord = {
      userID: userID,
      expiresAt: Date.now() + sessionDuration
    };

    // ScriptProperties lebih tahan terhadap cache yang dibersihkan Google.
    PropertiesService.getScriptProperties().setProperty(
      "HASHIRASAI_SESSION_" + token,
      JSON.stringify(sessionRecord)
    );
    _writeMemberLog(userID, "LOGIN", rememberMe
      ? "Anggota berhasil login (Ingat Saya 30 hari)."
      : "Anggota berhasil login (sesi 12 jam).");

    return {
      success: true,
      message: "Login berhasil.",
      token: token,
      member: {
        userID: user[0], namaIC: user[1], discordUsername: user[2],
        discordID: user[3] || "", role: role, status: status, accountStatus: accountStatus, jabatan: user[9],fotoUrl: user[11] || "",
      }
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}


/**
 * VALIDASI SESSION ANGGOTA
 */
function validateMemberSession(token) {
  try {
    if (!token) {
      throw new Error("Sesi tidak ditemukan. Silakan login kembali.");
    }

    const properties = PropertiesService.getScriptProperties();
    const sessionKey = "HASHIRASAI_SESSION_" + token;
    const rawSession = properties.getProperty(sessionKey);

    if (!rawSession) {
      throw new Error("Sesi tidak ditemukan atau telah berakhir. Silakan login kembali.");
    }

    let sessionRecord;
    try {
      sessionRecord = JSON.parse(rawSession);
    } catch (parseError) {
      properties.deleteProperty(sessionKey);
      throw new Error("Data sesi tidak valid. Silakan login kembali.");
    }

    if (!sessionRecord.userID || !sessionRecord.expiresAt ||
        Date.now() > Number(sessionRecord.expiresAt)) {
      properties.deleteProperty(sessionKey);
      throw new Error("Sesi telah berakhir. Silakan login kembali.");
    }

    const userID = sessionRecord.userID;

    const sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(CONFIG.SHEETS.USERS);

    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      throw new Error("Data anggota tidak ditemukan.");
    }

    const rows = sheet.getRange(2, 1, lastRow - 1, Math.max(11, sheet.getLastColumn())).getValues();

    const user = rows.find(row => row[0] === userID);

    if (!user || String(user[7]).toUpperCase() !== "APPROVED") {
      properties.deleteProperty(sessionKey);
      throw new Error("Akun tidak memiliki akses ke portal.");
    }
    if (String(user[10] || "ACTIVE").toUpperCase() !== "ACTIVE") {
      properties.deleteProperty(sessionKey);
      throw new Error("Akun kamu sedang TIDAK AKTIF. Hubungi ADMIN.");
    }

    return {
      success: true,
      member: {
        userID: user[0],
        namaIC: user[1],
        discordUsername: user[2],
        discordID: user[3],
        role: user[6],
        status: user[7],
        accountStatus: String(user[10] || "ACTIVE").toUpperCase(),
        jabatan: user[9],
        fotoUrl: user[11] || "",
      }
    };

  } catch (error) {
    return {
      success: false,
      message: error.message
    };
  }
}


/**
 * LOGOUT ANGGOTA
 */
function logoutMember(token) {
  try {
    if (token) {
      PropertiesService.getScriptProperties()
        .deleteProperty("HASHIRASAI_SESSION_" + token);
    }

    return {
      success: true,
      message: "Berhasil logout."
    };

  } catch (error) {
    return {
      success: false,
      message: error.message
    };
  }
}


function doGet() {
  return HtmlService
    .createHtmlOutputFromFile("index")
    .setTitle("HASHIRASAI | Community Portal")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}


/**
 * DATA DASHBOARD ANGGOTA
 */
function getMemberDashboard(token) {
  try {
    const validation = validateMemberSession(token);
    if (!validation || validation.success !== true || !validation.member) {
      throw new Error((validation && validation.message) || 'Sesi tidak valid. Silakan login kembali.');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('Spreadsheet database tidak dapat diakses.');

    const userSheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
    const attendanceSheet = ss.getSheetByName(CONFIG.SHEETS.ABSENSI);
    if (!userSheet) throw new Error('Sheet USERS tidak ditemukan.');
    if (!attendanceSheet) throw new Error('Sheet ABSENSI tidak ditemukan.');

    const member = validation.member;
    const readRows = (sheet) => {
      const lastRow = sheet.getLastRow();
      const lastCol = Math.max(sheet.getLastColumn(), 1);
      if (lastRow < 2) return { headers: [], rows: [] };
      const values = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
      return {
        headers: values[0].map(v => String(v || '').trim().toUpperCase()),
        rows: values.slice(1)
      };
    };
    const col = (headers, name, fallback) => {
      const found = headers.indexOf(name);
      return found >= 0 ? found : fallback;
    };

    const userData = readRows(userSheet);
    const u = {
      id: col(userData.headers, 'USER_ID', 0),
      nama: col(userData.headers, 'NAMA_IC', 1),
      discord: col(userData.headers, 'DISCORD_USERNAME', 2),
      role: col(userData.headers, 'ROLE', 6),
      status: col(userData.headers, 'STATUS', 7),
      jabatan: col(userData.headers, 'JABATAN', 9),
      accountStatus: col(userData.headers, 'STATUS_AKUN', 10)
    };
    const directory = userData.rows
      .filter(row => String(row[u.status] || '').trim().toUpperCase() === 'APPROVED')
      .map(row => ({
        userID: row[u.id] || '',
        namaIC: row[u.nama] || '',
        discordUsername: row[u.discord] || '',
        role: String(row[u.role] || 'MEMBER').toUpperCase(),
        jabatan: row[u.jabatan] || ''
      }));

    const attendanceData = readRows(attendanceSheet);
    const at = {
      id: col(attendanceData.headers, 'ABSENSI_ID', 0),
      user: col(attendanceData.headers, 'USER_ID', 1),
      kegiatan: col(attendanceData.headers, 'KEGIATAN_ID', 2),
      checkIn: col(attendanceData.headers, 'CHECK_IN', 3),
      checkOut: col(attendanceData.headers, 'CHECK_OUT', 4),
      durasi: col(attendanceData.headers, 'DURASI', 5),
      status: col(attendanceData.headers, 'STATUS', 6)
    };
    const myAttendance = attendanceData.rows
      .filter(row => String(row[at.user] || '') === String(member.userID || ''))
      .map(row => ({
        absensiID: row[at.id] || '', kegiatanID: row[at.kegiatan] || '',
        namaIC: member.namaIC || '', checkIn: row[at.checkIn] || '',
        checkOut: row[at.checkOut] || '', durasi: row[at.durasi] || '',
        status: row[at.status] || ''
      })).reverse();

    const counts = {};
    attendanceData.rows.forEach(row => {
      const status = String(row[at.status] || '').trim().toUpperCase();
      const userID = String(row[at.user] || '');
      if (userID && ['HADIR', 'SELESAI'].includes(status)) counts[userID] = (counts[userID] || 0) + 1;
    });
    const leaderboard = directory.map(person => ({
      namaIC: person.namaIC, jabatan: person.jabatan || '-',
      totalHadir: counts[String(person.userID)] || 0
    })).sort((x, y) => y.totalHadir - x.totalHadir || String(x.namaIC).localeCompare(String(y.namaIC)));

        const openEvents = _getOpenEventsForDashboard();

    return {
      success: true,
      member: member,
      leaderboard: leaderboard,
      activities: openEvents,
      statistics: {
        totalAnggota: directory.length,
        totalKegiatan: openEvents.length,
        totalKehadiran: myAttendance.filter(a => ["HADIR", "SELESAI"].includes(String(a.status).toUpperCase())).length
      },
      directory: directory,
      attendance: myAttendance
    };
  } catch (error) {
    console.error('getMemberDashboard error:', error && error.stack ? error.stack : error);
    return { success: false, message: error && error.message ? error.message : String(error) };
  }
}


/**
 * ADMIN: daftar pendaftar dan anggota untuk pengelolaan.
 * OWNER dan ADMIN boleh melihat; hanya ADMIN yang dapat mengubah role.
 */
function getAdminMemberManagement(token) {
  try {
    const validation = validateMemberSession(token);
    if (!validation.success) throw new Error(validation.message);
    const actor = validation.member;
    const actorRole = String(actor.role || "").toUpperCase();
    if (!["OWNER", "ADMIN"].includes(actorRole)) {
      throw new Error("Akses ditolak. Fitur ini khusus Admin dan Owner.");
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.USERS);
    if (!sheet) throw new Error("Sheet USERS tidak ditemukan.");
    const lastRow = sheet.getLastRow();
    const rows = lastRow > 1
      ? sheet.getRange(2, 1, lastRow - 1, Math.max(11, sheet.getLastColumn())).getDisplayValues()
      : [];

    const members = rows.map(row => ({
      userID: row[0],
      namaIC: row[1],
      discordUsername: row[2],
      role: String(row[6] || "MEMBER").toUpperCase(),
      status: String(row[7] || "").toUpperCase(),
      tanggalDaftar: row[8],
      jabatan: row[9] || "",
      accountStatus: String(row[10] || "ACTIVE").toUpperCase()
    }));

    return {
      success: true,
      actorRole: actorRole,
      canManageRoles: actorRole === "ADMIN",
      canManageJabatan: ["OWNER", "ADMIN"].includes(actorRole),
      canManageAccountStatus: ["OWNER", "ADMIN"].includes(actorRole),
      pending: members.filter(m => m.status === "PENDING"),
      approved: members.filter(m => m.status === "APPROVED"),
      rejected: members.filter(m => m.status === "REJECTED")
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * ADMIN dan OWNER: mengubah status akun ACTIVE/INACTIVE.
 * OWNER tidak dapat mengubah status akun ADMIN; ADMIN memiliki wewenang lebih tinggi.
 */
function setMemberAccountStatus(token, targetUserID, newStatus) {
  const lock = LockService.getScriptLock();
  try {
    const validation = validateMemberSession(token);
    if (!validation.success) throw new Error(validation.message);
    const actor = validation.member;
    const actorRole = String(actor.role || "").toUpperCase();
    if (!["ADMIN", "OWNER"].includes(actorRole)) {
      throw new Error("Akses ditolak. Hanya ADMIN dan OWNER yang dapat mengubah status akun.");
    }
    const accountStatus = String(newStatus || "").trim().toUpperCase();
    if (!["ACTIVE", "INACTIVE"].includes(accountStatus)) {
      throw new Error("Status akun hanya boleh ACTIVE atau INACTIVE.");
    }
    if (!targetUserID) throw new Error("ID anggota tidak ditemukan.");
    if (String(targetUserID) === String(actor.userID)) {
      throw new Error("Kamu tidak dapat mengubah status akun sendiri.");
    }
    if (!lock.tryLock(10000)) throw new Error("Sistem sedang sibuk. Coba lagi.");

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.USERS);
    if (!sheet) throw new Error("Sheet USERS tidak ditemukan.");
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0]
      .map(value => String(value || "").trim().toUpperCase());
    const statusCol = headers.indexOf("STATUS_AKUN") + 1;
    if (!statusCol) throw new Error("Kolom STATUS_AKUN belum tersedia. Jalankan setupDatabase().");

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error("Belum ada data anggota.");
    const rows = sheet.getRange(2, 1, lastRow - 1, Math.max(11, sheet.getLastColumn())).getValues();
    const index = rows.findIndex(row => String(row[0]) === String(targetUserID));
    if (index < 0) throw new Error("Anggota tidak ditemukan.");

    const target = rows[index];
    if (String(target[7] || "").toUpperCase() !== "APPROVED") {
      throw new Error("Status akun hanya dapat diubah untuk anggota yang sudah disetujui.");
    }
    const targetRole = String(target[6] || "MEMBER").toUpperCase();
    if (actorRole === "OWNER" && targetRole === "ADMIN") {
      throw new Error("OWNER tidak dapat mengubah status akun ADMIN.");
    }

    const oldStatus = String(target[statusCol - 1] || "ACTIVE").toUpperCase();
    sheet.getRange(index + 2, statusCol).setValue(accountStatus);
    _writeMemberLog(actor.userID, "MENGUBAH STATUS AKUN",
      "Mengubah status akun " + target[2] + " (" + targetUserID + ") dari " + oldStatus + " menjadi " + accountStatus + ".");
    return { success: true, message: "Status akun berhasil diubah menjadi " + (accountStatus === "ACTIVE" ? "AKTIF" : "TIDAK AKTIF") + "." };
  } catch (error) {
    return { success: false, message: error.message };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

/**
 * ADMIN/OWNER: setujui atau tolak pendaftar.
 */
function setApplicantStatus(token, targetUserID, newStatus) {
  const lock = LockService.getScriptLock();
  try {
    const validation = validateMemberSession(token);
    if (!validation.success) throw new Error(validation.message);
    const actor = validation.member;
    const actorRole = String(actor.role || "").toUpperCase();
    if (actorRole !== "ADMIN") {
      throw new Error("Akses ditolak. Hanya ADMIN yang dapat memproses pendaftar.");
    }

    const status = String(newStatus || "").toUpperCase();
    if (!["APPROVED", "REJECTED"].includes(status)) {
      throw new Error("Status hanya boleh APPROVED atau REJECTED.");
    }
    if (!targetUserID) throw new Error("ID anggota tidak ditemukan.");
    if (!lock.tryLock(10000)) throw new Error("Sistem sedang sibuk. Coba lagi.");

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.USERS);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error("Belum ada data anggota.");
    const rows = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    const index = rows.findIndex(row => String(row[0]) === String(targetUserID));
    if (index < 0) throw new Error("Pendaftar tidak ditemukan.");

    const target = rows[index];
    if (String(target[7]).toUpperCase() !== "PENDING") {
      throw new Error("Pendaftar ini sudah diproses sebelumnya.");
    }

    const rowNumber = index + 2;
    sheet.getRange(rowNumber, 8).setValue(status);
    _writeMemberLog(actor.userID, status === "APPROVED" ? "MENYETUJUI PENDAFTAR" : "MENOLAK PENDAFTAR",
      "Memproses " + target[2] + " (" + targetUserID + ") menjadi " + status + ".");
    return { success: true, message: status === "APPROVED" ? "Pendaftar berhasil disetujui." : "Pendaftar berhasil ditolak." };
  } catch (error) {
    return { success: false, message: error.message };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

/**
 * ADMIN: ubah role anggota yang sudah disetujui.
 * Akun tidak dapat mengubah role sendiri atau mengubah role OWNER utama.
 */
function setMemberRole(token, targetUserID, newRole) {
  const lock = LockService.getScriptLock();
  try {
    const validation = validateMemberSession(token);
    if (!validation.success) throw new Error(validation.message);
    const actor = validation.member;
    if (String(actor.role || "").toUpperCase() !== "ADMIN") {
      throw new Error("Akses ditolak. Hanya ADMIN yang dapat mengelola role anggota.");
    }

    const role = String(newRole || "").toUpperCase();
    if (!["MEMBER", "ADMIN"].includes(role)) {
      throw new Error("Role yang dapat diberikan adalah MEMBER atau ADMIN.");
    }
    if (!targetUserID) throw new Error("ID anggota tidak ditemukan.");
    if (String(targetUserID) === String(actor.userID)) {
      throw new Error("Role akun yang sedang digunakan tidak dapat diubah dari menu ini.");
    }
    if (!lock.tryLock(10000)) throw new Error("Sistem sedang sibuk. Coba lagi.");

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.USERS);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error("Belum ada data anggota.");
    const rows = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    const index = rows.findIndex(row => String(row[0]) === String(targetUserID));
    if (index < 0) throw new Error("Anggota tidak ditemukan.");

    const target = rows[index];
    if (String(target[7]).toUpperCase() !== "APPROVED") {
      throw new Error("Role hanya dapat diubah untuk anggota yang sudah disetujui.");
    }
    if (String(target[6]).toUpperCase() === "OWNER") {
      throw new Error("Role OWNER utama tidak dapat diubah melalui menu ini.");
    }

    sheet.getRange(index + 2, 7).setValue(role);
    _writeMemberLog(actor.userID, "MENGUBAH ROLE",
      "Mengubah role " + target[2] + " (" + targetUserID + ") menjadi " + role + ".");
    return { success: true, message: "Role anggota berhasil diperbarui." };
  } catch (error) {
    return { success: false, message: error.message };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}


/**
 * OWNER dan ADMIN: mengubah jabatan anggota. Role sistem tidak ikut berubah.
 */
function setMemberJabatan(token, targetUserID, newJabatan) {
  const lock = LockService.getScriptLock();
  try {
    const validation = validateMemberSession(token);
    if (!validation.success) throw new Error(validation.message);
    const actor = validation.member;
    const actorRole = String(actor.role || "").toUpperCase();
    if (!["OWNER", "ADMIN"].includes(actorRole)) {
      throw new Error("Akses ditolak. Hanya ADMIN dan OWNER yang dapat mengelola jabatan.");
    }
    const jabatan = String(newJabatan || "").trim().toUpperCase();
    const allowed = ["OYABUN", "WAKAGASHIRA", "KYODAI", "SHATEGASHIRA", "SHATEI"];
    if (!allowed.includes(jabatan)) throw new Error("Jabatan tidak valid.");
    if (!targetUserID) throw new Error("ID anggota tidak ditemukan.");
    if (!lock.tryLock(10000)) throw new Error("Sistem sedang sibuk. Coba lagi.");

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.USERS);
    if (!sheet) throw new Error("Sheet USERS tidak ditemukan.");
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0]
      .map(value => String(value || "").trim().toUpperCase());
    const jabatanCol = headers.indexOf("JABATAN") + 1;
    if (!jabatanCol) throw new Error("Kolom JABATAN tidak ditemukan. Jalankan setupDatabase terlebih dahulu.");
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error("Belum ada data anggota.");
    const rows = sheet.getRange(2, 1, lastRow - 1, Math.max(10, sheet.getLastColumn())).getValues();
    const index = rows.findIndex(row => String(row[0]) === String(targetUserID));
    if (index < 0) throw new Error("Anggota tidak ditemukan.");
    if (String(rows[index][7] || "").toUpperCase() !== "APPROVED") {
      throw new Error("Jabatan hanya dapat diubah untuk anggota yang sudah disetujui.");
    }
    const oldJabatan = String(rows[index][jabatanCol - 1] || "-" );
    sheet.getRange(index + 2, jabatanCol).setValue(jabatan);
    _writeMemberLog(actor.userID, "MENGUBAH JABATAN",
      "Mengubah jabatan " + rows[index][2] + " (" + targetUserID + ") dari " + oldJabatan + " menjadi " + jabatan + ".");
    return { success: true, message: "Jabatan anggota berhasil diperbarui." };
  } catch (error) {
    return { success: false, message: error.message };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}


/**
 * Menyimpan check-in, check-out, atau izin anggota yang sedang login.
 * Foto disimpan di Google Drive; status yang didukung hanya HADIR, SELESAI, IZIN.
 */
function _saveAbsenPhoto(photoData, userID) {
  if (!photoData) return "";
  const bytes = Utilities.base64Decode(String(photoData));
  if (bytes.length > 2 * 1024 * 1024) throw new Error("Ukuran foto maksimal 2 MB.");
  const folderName = "HASHIRASAI_BUKTI_ABSENSI";
  const folders = DriveApp.getFoldersByName(folderName);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
  const blob = Utilities.newBlob(bytes, "image/jpeg", "ABSEN_" + userID + "_" + Date.now() + ".jpg");
  return folder.createFile(blob).getUrl();
}

function submitMemberAttendance(data) {
  const lock = LockService.getScriptLock();
  try {
    if (!data) throw new Error("Data absensi tidak ditemukan.");
    if (!lock.tryLock(10000)) throw new Error("Sistem sedang sibuk. Coba lagi.");
    const validation = validateMemberSession(String(data.token || ""));
    if (!validation.success) throw new Error(validation.message || "Sesi tidak valid.");
    const member = validation.member;
    const action = String(data.action || "").toUpperCase();
    const reason = String(data.reason || "").trim();
    if (!["HADIR", "SELESAI", "IZIN", "ALPA"].includes(action)) throw new Error("Status absensi tidak valid.");
    if (action === "IZIN" && !reason) throw new Error("Alasan izin wajib diisi.");

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.ABSENSI);
    if (!sheet) throw new Error("Sheet absensi tidak ditemukan.");

    const kegiatanID = "ABSENSI_UMUM";
    const tz = Session.getScriptTimeZone();
    const now = new Date();
    const todayKey = Utilities.formatDate(now, tz, "yyyy-MM-dd");
    const last = sheet.getLastRow();
    const rows = last > 1
      ? sheet.getRange(2, 1, last - 1, Math.max(9, sheet.getLastColumn())).getValues()
      : [];
    const toDate = v => v instanceof Date ? v : (v ? new Date(v) : null);

    const todayRows = rows.map((row, i) => ({ row: row, i: i })).filter(o => {
      if (String(o.row[1]) !== String(member.userID)) return false;
      if (String(o.row[2]) !== kegiatanID) return false;
      const d = toDate(o.row[3]);
      return d && !isNaN(d.getTime()) && Utilities.formatDate(d, tz, "yyyy-MM-dd") === todayKey;
    });
    const presence = todayRows.find(o => ["HADIR", "SELESAI"].includes(String(o.row[6]).toUpperCase()));
    const absent = todayRows.find(o => ["IZIN", "ALPA"].includes(String(o.row[6]).toUpperCase()));

    if (action === "IZIN" || action === "ALPA") {
      if (presence || absent) throw new Error("Kamu sudah memiliki catatan absensi hari ini.");
    } else if (action === "HADIR") {
      if (absent) throw new Error("Hari ini kamu sudah tercatat " + String(absent.row[6]).toUpperCase() + ".");
      if (presence) throw new Error("Kamu sudah melakukan check-in hari ini.");
    } else {
      if (!presence) throw new Error("Lakukan check-in terlebih dahulu sebelum check-out.");
      if (presence.row[4]) throw new Error("Kamu sudah melakukan check-out hari ini.");
    }

    const photoUrl = _saveAbsenPhoto(data.photoData, member.userID);

    if (action === "IZIN" || action === "ALPA") {
      sheet.appendRow(["ABS-" + Utilities.getUuid(), member.userID, kegiatanID, now, "", "", action, photoUrl, reason]);
    } else if (action === "HADIR") {
      sheet.appendRow(["ABS-" + Utilities.getUuid(), member.userID, kegiatanID, now, "", "", "HADIR", photoUrl, ""]);
    } else {
      const rowNumber = presence.i + 2;
      const checkIn = toDate(presence.row[3]);
      const minutes = Math.max(0, Math.round((now.getTime() - checkIn.getTime()) / 60000));
      const durasi = Math.floor(minutes / 60) + " jam " + (minutes % 60) + " menit";
      sheet.getRange(rowNumber, 5, 1, 3).setValues([[now, durasi, "SELESAI"]]);
      if (photoUrl) sheet.getRange(rowNumber, 8).setValue(photoUrl);
    }

    _writeMemberLog(member.userID, "ABSENSI_" + action, "Absensi umum: " + action);
    const msgs = {
      HADIR: "Check-in berhasil disimpan.",
      SELESAI: "Check-out berhasil disimpan.",
      IZIN: "Pengajuan izin berhasil dikirim.",
      ALPA: "Status Alpa berhasil dicatat."
    };
    return { success: true, message: msgs[action] };
  } catch (error) {
    return { success: false, message: error.message };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

/* =====================================================================
   UTILITAS UMUM (Mission, Event, Target Setoran)
   ===================================================================== */
function _eq(a, b) { return String(a) === String(b); }

function _newId(prefix) {
  return prefix + "-" + Utilities.getUuid().replace(/-/g, "").slice(0, 12).toUpperCase();
}

function _now() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
}

function _todayKey() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function _weekKey(d) {
  const tz = Session.getScriptTimeZone();
  const date = new Date(Date.UTC(
    Number(Utilities.formatDate(d, tz, "yyyy")),
    Number(Utilities.formatDate(d, tz, "M")) - 1,
    Number(Utilities.formatDate(d, tz, "d"))
  ));
  const day = date.getUTCDay() || 7;
  date.setUTCDate(date.getUTCDate() + 4 - day);
  const yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
  const week = Math.ceil(((date - yearStart) / 86400000 + 1) / 7);
  return date.getUTCFullYear() + "-W" + ("0" + week).slice(-2);
}

function _sheetByKey(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = CONFIG.SHEETS[key];
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    const h = HEADERS[key];
    sheet.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function _readTable(key) {
  const sheet = _sheetByKey(key);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  const tz = Session.getScriptTimeZone();
  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0].map(h => String(h || "").trim());
  return values.slice(1)
    .filter(r => r.some(c => c !== ""))
    .map(row => {
      const o = {};
      headers.forEach((h, i) => {
        if (!h) return;
        let v = row[i];
        if (v instanceof Date) v = Utilities.formatDate(v, tz, "yyyy-MM-dd HH:mm:ss");
        o[h] = (v === null || v === undefined) ? "" : v;
      });
      return o;
    });
}

// Semua nilai disimpan sebagai teks agar tanggal/jam tidak diubah otomatis oleh Sheets.
function _appendRow(sheet, values) {
  const row = sheet.getLastRow() + 1;
  if (row > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), 1);
  const range = sheet.getRange(row, 1, 1, values.length);
  range.setNumberFormat("@");
  range.setValues([values.map(v => String(v === null || v === undefined ? "" : v))]);
}

function _rowOf(sheet, headerName, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const col = headers.indexOf(headerName) + 1;
  if (!col) return -1;
  const ids = sheet.getRange(2, col, lastRow - 1, 1).getValues().map(r => String(r[0]));
  const i = ids.indexOf(String(id));
  return i < 0 ? -1 : i + 2;
}

function _setCell(sheet, row, headerName, value) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const col = headers.indexOf(headerName) + 1;
  if (!col) throw new Error("Kolom " + headerName + " tidak ditemukan.");
  sheet.getRange(row, col).setNumberFormat("@").setValue(String(value));
}

function _authMember(token) {
  const v = validateMemberSession(token);
  if (!v || !v.success) throw new Error((v && v.message) || "Sesi tidak valid. Silakan login kembali.");
  return v.member;
}

// Yang boleh mengisi/mengelola: ROLE ADMIN atau JABATAN OYABUN.
// Mau OWNER ikut boleh? Tambahkan: || String(member.role||"").toUpperCase()==="OWNER"
function _canManageContent(member) {
  return String(member.role || "").toUpperCase() === "ADMIN" ||
         String(member.jabatan || "").trim().toUpperCase() === "OYABUN";
}

function _readAction(token, work) {
  try {
    return work(_authMember(token));
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function _writeAction(token, needManager, work) {
  const lock = LockService.getScriptLock();
  try {
    const member = _authMember(token);
    if (needManager && !_canManageContent(member)) {
      throw new Error("Akses ditolak. Hanya OYABUN dan ADMIN yang dapat mengisi/mengelola data ini.");
    }
    if (!lock.tryLock(10000)) throw new Error("Sistem sedang sibuk. Coba lagi.");
    return work(member);
  } catch (error) {
    return { success: false, message: error.message };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

/* =====================================================================
   MISSION HARIAN & MINGGUAN
   ===================================================================== */
function getMissions(token) {
  return _readAction(token, function (member) {
    const canManage = _canManageContent(member);
    const progress = _readTable("MISSION_PROGRESS");
    const today = _todayKey();
    const week = _weekKey(new Date());

    const missions = _readTable("MISSION")
      .filter(m => canManage || String(m.STATUS).toUpperCase() === "AKTIF")
      .reverse()
      .map(m => {
        const tipe = String(m.TIPE).toUpperCase();
        const periode = tipe === "MINGGUAN" ? week : today;
        const done = progress.filter(p => _eq(p.MISSION_ID, m.MISSION_ID) && _eq(p.PERIODE, periode));
        return {
          missionID: m.MISSION_ID,
          judul: m.JUDUL,
          deskripsi: m.DESKRIPSI,
          tipe: tipe,
          reward: m.REWARD,
          status: String(m.STATUS).toUpperCase(),
          selesaiSaya: done.some(p => _eq(p.USER_ID, member.userID)),
          totalSelesai: done.length
        };
      });
    return { success: true, canManage: canManage, missions: missions };
  });
}

function saveMission(token, data) {
  return _writeAction(token, true, function (member) {
    const judul = String((data && data.judul) || "").trim();
    const tipe = String((data && data.tipe) || "").toUpperCase();
    if (judul.length < 3) throw new Error("Judul mission minimal 3 karakter.");
    if (!["HARIAN", "MINGGUAN"].includes(tipe)) throw new Error("Tipe mission tidak valid.");
    _appendRow(_sheetByKey("MISSION"), [
      _newId("MIS"), judul, String(data.deskripsi || "").trim(), tipe,
      String(data.reward || "").trim(), "AKTIF", member.userID, _now()
    ]);
    _writeMemberLog(member.userID, "MEMBUAT MISSION", "Mission: " + judul);
    return { success: true, message: "Mission berhasil ditambahkan." };
  });
}

function setMissionStatus(token, missionID, status) {
  return _writeAction(token, true, function (member) {
    const st = String(status || "").toUpperCase();
    if (!["AKTIF", "NONAKTIF"].includes(st)) throw new Error("Status tidak valid.");
    const sheet = _sheetByKey("MISSION");
    const row = _rowOf(sheet, "MISSION_ID", missionID);
    if (row < 0) throw new Error("Mission tidak ditemukan.");
    _setCell(sheet, row, "STATUS", st);
    _writeMemberLog(member.userID, "UBAH STATUS MISSION", missionID + " -> " + st);
    return { success: true, message: "Status mission diperbarui." };
  });
}

function deleteMission(token, missionID) {
  return _writeAction(token, true, function (member) {
    const sheet = _sheetByKey("MISSION");
    const row = _rowOf(sheet, "MISSION_ID", missionID);
    if (row < 0) throw new Error("Mission tidak ditemukan.");
    sheet.deleteRow(row);
    _writeMemberLog(member.userID, "HAPUS MISSION", missionID);
    return { success: true, message: "Mission dihapus." };
  });
}

function completeMission(token, missionID) {
  return _writeAction(token, false, function (member) {
    const m = _readTable("MISSION").find(x => _eq(x.MISSION_ID, missionID));
    if (!m) throw new Error("Mission tidak ditemukan.");
    if (String(m.STATUS).toUpperCase() !== "AKTIF") throw new Error("Mission ini sedang tidak aktif.");
    const tipe = String(m.TIPE).toUpperCase();
    const periode = tipe === "MINGGUAN" ? _weekKey(new Date()) : _todayKey();
    const already = _readTable("MISSION_PROGRESS").some(p =>
      _eq(p.MISSION_ID, missionID) && _eq(p.USER_ID, member.userID) && _eq(p.PERIODE, periode));
    if (already) throw new Error("Mission ini sudah kamu selesaikan untuk periode ini.");
    _appendRow(_sheetByKey("MISSION_PROGRESS"), [_newId("PRG"), missionID, member.userID, periode, _now()]);
    _writeMemberLog(member.userID, "SELESAI MISSION", m.JUDUL + " (" + periode + ")");
    return { success: true, message: "Mission ditandai selesai. Kerja bagus!" };
  });
}

/* =====================================================================
   EVENT
   ===================================================================== */
function _getOpenEventsForDashboard() {
  try {
    return _readTable("EVENT")
      .filter(e => String(e.STATUS).toUpperCase() === "OPEN")
      .map(e => ({
        namaKegiatan: e.NAMA_EVENT, jenis: e.JENIS, tanggal: e.TANGGAL,
        jamMulai: e.JAM_MULAI, jamSelesai: e.JAM_SELESAI, status: String(e.STATUS).toUpperCase()
      }));
  } catch (error) {
    return [];
  }
}

function getEvents(token) {
  return _readAction(token, function (member) {
    const canManage = _canManageContent(member);
    const peserta = _readTable("EVENT_PESERTA");
    const nameMap = {};
    if (canManage) _readTable("USERS").forEach(u => { nameMap[u.USER_ID] = u.NAMA_IC; });

    const list = _readTable("EVENT").map(e => {
      const ps = peserta.filter(p => _eq(p.EVENT_ID, e.EVENT_ID));
      const item = {
        eventID: e.EVENT_ID, nama: e.NAMA_EVENT, jenis: e.JENIS, tanggal: e.TANGGAL,
        jamMulai: e.JAM_MULAI, jamSelesai: e.JAM_SELESAI, lokasi: e.LOKASI,
        deskripsi: e.DESKRIPSI, status: String(e.STATUS).toUpperCase(),
        jumlahPeserta: ps.length,
        ikutSaya: ps.some(p => _eq(p.USER_ID, member.userID))
      };
      if (canManage) item.daftarPeserta = ps.map(p => nameMap[p.USER_ID] || p.USER_ID);
      return item;
    });

    list.sort((a, b) => {
      if (a.status !== b.status) return a.status === "OPEN" ? -1 : 1;
      return a.status === "OPEN"
        ? String(a.tanggal).localeCompare(String(b.tanggal))
        : String(b.tanggal).localeCompare(String(a.tanggal));
    });
    return { success: true, canManage: canManage, events: list };
  });
}

function saveEvent(token, data) {
  return _writeAction(token, true, function (member) {
    const nama = String((data && data.nama) || "").trim();
    const tanggal = String((data && data.tanggal) || "").trim();
    const jamMulai = String((data && data.jamMulai) || "").trim();
    const jamSelesai = String((data && data.jamSelesai) || "").trim();
    if (nama.length < 3) throw new Error("Nama event minimal 3 karakter.");
    if (!/^\d{4}-\d{2}-\d{2}$/.test(tanggal)) throw new Error("Tanggal event wajib diisi.");
    if (jamMulai && !/^\d{2}:\d{2}$/.test(jamMulai)) throw new Error("Format jam mulai tidak valid.");
    if (jamSelesai && !/^\d{2}:\d{2}$/.test(jamSelesai)) throw new Error("Format jam selesai tidak valid.");
    _appendRow(_sheetByKey("EVENT"), [
      _newId("EVT"), nama, String(data.jenis || "").trim(), tanggal, jamMulai, jamSelesai,
      String(data.lokasi || "").trim(), String(data.deskripsi || "").trim(),
      "OPEN", member.userID, _now()
    ]);
    _writeMemberLog(member.userID, "MEMBUAT EVENT", "Event: " + nama);
    return { success: true, message: "Event berhasil ditambahkan." };
  });
}

function setEventStatus(token, eventID, status) {
  return _writeAction(token, true, function (member) {
    const st = String(status || "").toUpperCase();
    if (!["OPEN", "CLOSED"].includes(st)) throw new Error("Status tidak valid.");
    const sheet = _sheetByKey("EVENT");
    const row = _rowOf(sheet, "EVENT_ID", eventID);
    if (row < 0) throw new Error("Event tidak ditemukan.");
    _setCell(sheet, row, "STATUS", st);
    _writeMemberLog(member.userID, "UBAH STATUS EVENT", eventID + " -> " + st);
    return { success: true, message: "Status event diperbarui." };
  });
}

function deleteEvent(token, eventID) {
  return _writeAction(token, true, function (member) {
    const sheet = _sheetByKey("EVENT");
    const row = _rowOf(sheet, "EVENT_ID", eventID);
    if (row < 0) throw new Error("Event tidak ditemukan.");
    sheet.deleteRow(row);
    const ps = _sheetByKey("EVENT_PESERTA");
    const vals = ps.getDataRange().getValues();
    for (let i = vals.length - 1; i >= 1; i--) {
      if (_eq(vals[i][1], eventID)) ps.deleteRow(i + 1);
    }
    _writeMemberLog(member.userID, "HAPUS EVENT", eventID);
    return { success: true, message: "Event dihapus." };
  });
}

function toggleEventJoin(token, eventID) {
  return _writeAction(token, false, function (member) {
    const ev = _readTable("EVENT").find(e => _eq(e.EVENT_ID, eventID));
    if (!ev) throw new Error("Event tidak ditemukan.");
    if (String(ev.STATUS).toUpperCase() !== "OPEN") throw new Error("Event ini sudah ditutup.");
    const sheet = _sheetByKey("EVENT_PESERTA");
    const vals = sheet.getDataRange().getValues();
    let rowIdx = -1;
    for (let i = 1; i < vals.length; i++) {
      if (_eq(vals[i][1], eventID) && _eq(vals[i][2], member.userID)) { rowIdx = i + 1; break; }
    }
    if (rowIdx > 0) {
      sheet.deleteRow(rowIdx);
      return { success: true, message: "Kamu batal mengikuti event ini." };
    }
    _appendRow(sheet, [_newId("PST"), eventID, member.userID, _now()]);
    _writeMemberLog(member.userID, "IKUT EVENT", ev.NAMA_EVENT);
    return { success: true, message: "Kamu terdaftar di event ini." };
  });
}

/* =====================================================================
   TARGET SETORAN
   ===================================================================== */
function getTargets(token) {
  return _readAction(token, function (member) {
    const canManage = _canManageContent(member);
    const targets = _readTable("TARGET_SETORAN");
    const setoran = _readTable("SETORAN");
    const users = _readTable("USERS");
    const nameMap = {};
    users.forEach(u => { nameMap[u.USER_ID] = u.NAMA_IC; });
    const members = users.filter(u =>
      String(u.STATUS).toUpperCase() === "APPROVED" &&
      String(u.STATUS_AKUN || "ACTIVE").toUpperCase() === "ACTIVE");

    const sum = arr => arr.reduce((a, s) => a + (Number(s.NOMINAL) || 0), 0);
    const pct = (v, t) => t > 0 ? Math.min(100, Math.round(v / t * 100)) : 0;

    const out = targets.slice().reverse().map(t => {
      const nominal = Number(t.NOMINAL_TARGET) || 0;
      const rows = setoran.filter(s => _eq(s.TARGET_ID, t.TARGET_ID));
      const mine = rows.filter(s => _eq(s.USER_ID, member.userID));
      const totalSaya = sum(mine);
      const item = {
        targetID: t.TARGET_ID, judul: t.JUDUL, nominalTarget: nominal,
        mulai: t.PERIODE_MULAI, selesai: t.PERIODE_SELESAI, keterangan: t.KETERANGAN,
        status: String(t.STATUS).toUpperCase(),
        totalSaya: totalSaya, persenSaya: pct(totalSaya, nominal),
        setoranSaya: mine.slice(-5).reverse().map(s => ({
          nominal: Number(s.NOMINAL) || 0, catatan: s.CATATAN, waktu: s.WAKTU
        }))
      };
      if (canManage) {
        item.totalTerkumpul = sum(rows);
        item.rekap = members.map(m => {
          const total = sum(rows.filter(s => _eq(s.USER_ID, m.USER_ID)));
          return { userID: m.USER_ID, namaIC: m.NAMA_IC, total: total, persen: pct(total, nominal) };
        }).sort((a, b) => b.total - a.total);
        item.setoranTerbaru = rows.slice(-8).reverse().map(s => ({
          setoranID: s.SETORAN_ID, namaIC: nameMap[s.USER_ID] || s.USER_ID,
          nominal: Number(s.NOMINAL) || 0, catatan: s.CATATAN, waktu: s.WAKTU
        }));
      }
      return item;
    });

    return {
      success: true,
      canManage: canManage,
      targets: out,
      members: canManage ? members.map(m => ({ userID: m.USER_ID, namaIC: m.NAMA_IC })) : []
    };
  });
}

function saveTarget(token, data) {
  return _writeAction(token, true, function (member) {
    const judul = String((data && data.judul) || "").trim();
    const nominal = Number(data && data.nominalTarget);
    const mulai = String((data && data.mulai) || "").trim();
    const selesai = String((data && data.selesai) || "").trim();
    if (judul.length < 3) throw new Error("Judul target minimal 3 karakter.");
    if (!isFinite(nominal) || nominal <= 0) throw new Error("Nominal target harus lebih dari 0.");
    if (!/^\d{4}-\d{2}-\d{2}$/.test(mulai) || !/^\d{4}-\d{2}-\d{2}$/.test(selesai)) {
      throw new Error("Periode mulai dan selesai wajib diisi.");
    }
    if (selesai < mulai) throw new Error("Tanggal selesai tidak boleh sebelum tanggal mulai.");
    _appendRow(_sheetByKey("TARGET_SETORAN"), [
      _newId("TGT"), judul, nominal, mulai, selesai,
      String(data.keterangan || "").trim(), "AKTIF", member.userID, _now()
    ]);
    _writeMemberLog(member.userID, "MEMBUAT TARGET SETORAN", judul);
    return { success: true, message: "Target setoran berhasil dibuat." };
  });
}

function setTargetStatus(token, targetID, status) {
  return _writeAction(token, true, function (member) {
    const st = String(status || "").toUpperCase();
    if (!["AKTIF", "SELESAI"].includes(st)) throw new Error("Status tidak valid.");
    const sheet = _sheetByKey("TARGET_SETORAN");
    const row = _rowOf(sheet, "TARGET_ID", targetID);
    if (row < 0) throw new Error("Target tidak ditemukan.");
    _setCell(sheet, row, "STATUS", st);
    _writeMemberLog(member.userID, "UBAH STATUS TARGET", targetID + " -> " + st);
    return { success: true, message: "Status target diperbarui." };
  });
}

function deleteTarget(token, targetID) {
  return _writeAction(token, true, function (member) {
    const sheet = _sheetByKey("TARGET_SETORAN");
    const row = _rowOf(sheet, "TARGET_ID", targetID);
    if (row < 0) throw new Error("Target tidak ditemukan.");
    sheet.deleteRow(row);
    const sh = _sheetByKey("SETORAN");
    const vals = sh.getDataRange().getValues();
    for (let i = vals.length - 1; i >= 1; i--) {
      if (_eq(vals[i][1], targetID)) sh.deleteRow(i + 1);
    }
    _writeMemberLog(member.userID, "HAPUS TARGET", targetID);
    return { success: true, message: "Target beserta data setorannya dihapus." };
  });
}

function addSetoran(token, data) {
  return _writeAction(token, true, function (member) {
    const targetID = String((data && data.targetID) || "");
    const userID = String((data && data.userID) || "");
    const nominal = Number(data && data.nominal);
    if (!isFinite(nominal) || nominal <= 0) throw new Error("Nominal setoran harus lebih dari 0.");
    const t = _readTable("TARGET_SETORAN").find(x => _eq(x.TARGET_ID, targetID));
    if (!t) throw new Error("Target tidak ditemukan.");
    if (String(t.STATUS).toUpperCase() !== "AKTIF") throw new Error("Target ini sudah selesai/ditutup.");
    const u = _readTable("USERS").find(x => _eq(x.USER_ID, userID) && String(x.STATUS).toUpperCase() === "APPROVED");
    if (!u) throw new Error("Anggota tidak ditemukan.");
    _appendRow(_sheetByKey("SETORAN"), [
      _newId("STR"), targetID, userID, nominal, String(data.catatan || "").trim(), member.userID, _now()
    ]);
    _writeMemberLog(member.userID, "CATAT SETORAN", u.NAMA_IC + " : " + nominal + " (" + t.JUDUL + ")");
    return { success: true, message: "Setoran " + u.NAMA_IC + " berhasil dicatat." };
  });
}

function deleteSetoran(token, setoranID) {
  return _writeAction(token, true, function (member) {
    const sheet = _sheetByKey("SETORAN");
    const row = _rowOf(sheet, "SETORAN_ID", setoranID);
    if (row < 0) throw new Error("Data setoran tidak ditemukan.");
    sheet.deleteRow(row);
    _writeMemberLog(member.userID, "HAPUS SETORAN", setoranID);
    return { success: true, message: "Data setoran dihapus." };
  });
}

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