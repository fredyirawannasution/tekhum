const CONFIG = {
  APP_NAME: 'Sistem Upload Surat KPU',
  DEFAULT_ADMIN_NAME: 'Administrator',
  DEFAULT_ADMIN_EMAIL: 'admin@example.com',
  DEFAULT_ADMIN_PASSWORD: 'Admin@12345',
  SESSION_TTL_SECONDS: 300,
  RESET_TTL_SECONDS: 1800,
  SHEETS: {
    USERS: 'Users',
    UPLOADS: 'Uploads',
    FOLDERS: 'Folders'
  }
};

function doGet(e) {
  initializeApp_();
  const template = HtmlService.createTemplateFromFile('index');
  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'app';
  const token = (e && e.parameter && e.parameter.token) ? e.parameter.token : '';
  template.bootstrap = JSON.stringify({
    appName: CONFIG.APP_NAME,
    page: page,
    token: token
  });
  return template.evaluate()
    .setTitle(CONFIG.APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function initializeApp_() {
  const spreadsheet = getOrCreateDatabase_();
  const usersSheet = getOrCreateSheet_(spreadsheet, CONFIG.SHEETS.USERS, [
    'ID', 'Nama', 'Email', 'Username', 'PasswordHash', 'Role', 'Active', 'CreatedAt', 'UpdatedAt'
  ]);
  const uploadsSheet = getOrCreateSheet_(spreadsheet, CONFIG.SHEETS.UPLOADS, [
    'ID', 'Tanggal', 'NomorSurat', 'NamaSurat', 'JenisSurat', 'Kategori', 'SubFolder', 'DriveFolderId', 'DriveFolderLink', 'FileId', 'FileName', 'FileUrl', 'UploadedBy', 'UploadedByEmail', 'CreatedAt'
  ]);
  const foldersSheet = getOrCreateSheet_(spreadsheet, CONFIG.SHEETS.FOLDERS, [
    'ID', 'Kategori', 'NamaFolder', 'ParentDriveLink', 'ParentFolderId', 'DriveLink', 'FolderId', 'Active', 'CreatedAt', 'UpdatedAt'
  ]);

  migrateUsersSheet_(usersSheet);
  migrateFoldersSheet_(foldersSheet);

  if (usersSheet.getLastRow() === 1) {
    usersSheet.appendRow([
      Utilities.getUuid(),
      CONFIG.DEFAULT_ADMIN_NAME,
      CONFIG.DEFAULT_ADMIN_EMAIL.toLowerCase(),
      'admin',
      hashPassword_(CONFIG.DEFAULT_ADMIN_PASSWORD),
      'admin',
      'TRUE',
      nowIso_(),
      nowIso_()
    ]);
  }

  // Folder tidak lagi dibuat dari daftar hardcoded di kode.
  // Isi sheet Folders secara manual, buat dari menu Admin Folder, atau tekan tombol "Buat Folder dari Sheet".

  // Hindari auto-format setiap request agar login lebih cepat.
  // Format sheet dilakukan saat ada perubahan data besar atau manual bila perlu.
}

function getOrCreateDatabase_() {
  return SpreadsheetApp.openById('11bZJMKTRj90X4im_phoIpo9-Z8fgJMOoxO3nBQr8EcE');
}

function getOrCreateSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function migrateUsersSheet_(sheet) {
  const newHeaders = ['ID', 'Nama', 'Email', 'Username', 'PasswordHash', 'Role', 'Active', 'CreatedAt', 'UpdatedAt'];
  const oldHeaders = ['ID', 'Nama', 'Email', 'PasswordHash', 'Role', 'Active', 'CreatedAt', 'UpdatedAt'];
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow === 0) {
    sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
    sheet.setFrozenRows(1);
    return;
  }
  const headers = sheet.getRange(1, 1, 1, Math.max(lastCol, newHeaders.length)).getValues()[0].map(String);
  const isNew = newHeaders.every(function(header, index) { return headers[index] === header; });
  const isOld = oldHeaders.every(function(header, index) { return headers[index] === header; });
  if (isNew) return;
  if (isOld) {
    const existing = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, oldHeaders.length).getValues() : [];
    sheet.clear();
    sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
    if (existing.length) {
      const used = {};
      const migrated = existing.map(function(row) {
        let username = normalizeUsername_(String(row[2] || '').split('@')[0] || row[1] || 'user');
        const baseName = username || 'user';
        let counter = 2;
        while (used[username]) {
          username = baseName + counter;
          counter++;
        }
        used[username] = true;
        return [row[0], row[1], row[2], username, row[3], row[4], row[5], row[6], row[7]];
      });
      sheet.getRange(2, 1, migrated.length, newHeaders.length).setValues(migrated);
    }
    sheet.setFrozenRows(1);
    return;
  }
  sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
  sheet.setFrozenRows(1);
}

function migrateFoldersSheet_(sheet) {
  const newHeaders = ['ID', 'Kategori', 'NamaFolder', 'ParentDriveLink', 'ParentFolderId', 'DriveLink', 'FolderId', 'Active', 'CreatedAt', 'UpdatedAt'];
  const oldHeaders = ['ID', 'Kategori', 'NamaFolder', 'DriveLink', 'FolderId', 'Active', 'CreatedAt', 'UpdatedAt'];
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow === 0) {
    sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
    sheet.setFrozenRows(1);
    return;
  }

  const headers = sheet.getRange(1, 1, 1, Math.max(lastCol, oldHeaders.length)).getValues()[0].map(String);
  const isOld = oldHeaders.every(function(header, index) { return headers[index] === header; });
  const isNew = newHeaders.every(function(header, index) { return headers[index] === header; });

  if (isNew) return;

  if (isOld) {
    const existing = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, oldHeaders.length).getValues() : [];
    sheet.clear();
    sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
    if (existing.length) {
      const migrated = existing.map(function(row) {
        return [
          row[0], row[1], row[2], '', '', row[3], row[4], row[5], row[6], row[7]
        ];
      });
      sheet.getRange(2, 1, migrated.length, newHeaders.length).setValues(migrated);
    }
    sheet.setFrozenRows(1);
    return;
  }

  sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
  sheet.setFrozenRows(1);
}

function fitSheet_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastCol > 0) sheet.autoResizeColumns(1, lastCol);
  if (lastRow < 1 || lastCol < 1) return;

  sheet.getRange(1, 1, lastRow, lastCol)
    .setBorder(true, true, true, true, true, true)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.getRange(1, 1, 1, lastCol)
    .setFontWeight('bold')
    .setBackground('#f8fafc');
}

function formatDataRow_(sheet, rowNumber) {
  const lastCol = sheet.getLastColumn();
  if (!rowNumber || rowNumber < 1 || lastCol < 1) return;
  sheet.getRange(rowNumber, 1, 1, lastCol)
    .setBorder(true, true, true, true, true, true)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.getRange(1, 1, 1, lastCol)
    .setFontWeight('bold')
    .setBackground('#f8fafc')
    .setBorder(true, true, true, true, true, true);
}

function getPublicAppInfo() {
  initializeApp_();
  return {
    ok: true,
    appName: CONFIG.APP_NAME,
    hasAdmin: true
  };
}

function login(payload) {
  initializeApp_();
  const identifier = String(payload.identifier || payload.email || payload.username || '').trim().toLowerCase();
  const password = String(payload.password || '');
  if (!identifier || !password) throw new Error('Email/username dan password wajib diisi.');
  const user = findUserByLogin_(identifier);
  if (!user || user.active !== true) throw new Error('Email/username atau password salah.');
  if (user.passwordHash !== hashPassword_(password)) throw new Error('Email/username atau password salah.');
  const sessionToken = Utilities.getUuid() + '.' + Utilities.getUuid();
  const cache = CacheService.getScriptCache();
  cache.put('session:' + sessionToken, JSON.stringify({
    userId: user.id,
    email: user.email,
    username: user.username,
    role: user.role,
    name: user.name,
    createdAt: nowIso_(),
    lastActiveAt: nowIso_(),
    loginMode: 'password'
  }), CONFIG.SESSION_TTL_SECONDS);
  return {
    ok: true,
    sessionToken: sessionToken,
    user: sanitizeUser_(user),
    dashboard: buildDashboardResponse_(user),
    loginMode: 'password'
  };
}
function logout(sessionToken) {
  if (sessionToken) {
    CacheService.getScriptCache().remove('session:' + sessionToken);
  }
  return { ok: true };
}

function getSessionUser(sessionToken) {
  return { ok: true, user: sanitizeUser_(requireSession_(sessionToken)) };
}

function requestPasswordReset(email) {
  initializeApp_();
  email = String(email || '').trim().toLowerCase();
  if (!email) throw new Error('Email wajib diisi.');

  const user = findUserByEmail_(email);
  if (!user || user.active !== true) {
    return { ok: true, message: 'Jika email terdaftar, tautan reset password sudah dikirim.' };
  }

  const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
  CacheService.getScriptCache().put('reset:' + token, JSON.stringify({
    userId: user.id,
    email: user.email
  }), CONFIG.RESET_TTL_SECONDS);

  const appUrl = ScriptApp.getService().getUrl();
  const resetUrl = appUrl ? (appUrl + '?page=reset&token=' + encodeURIComponent(token)) : '(deploy web app terlebih dahulu untuk mendapatkan link reset)';
  const body = [
    'Halo ' + user.name + ',',
    '',
    'Anda meminta reset password untuk aplikasi ' + CONFIG.APP_NAME + '.',
    'Klik tautan berikut untuk membuat password baru:',
    resetUrl,
    '',
    'Tautan ini berlaku selama 30 menit.',
    'Jika Anda tidak merasa meminta reset password, abaikan email ini.'
  ].join('\n');

  MailApp.sendEmail({
    to: user.email,
    subject: '[' + CONFIG.APP_NAME + '] Reset Password',
    body: body
  });

  return { ok: true, message: 'Jika email terdaftar, tautan reset password sudah dikirim.' };
}

function validateResetToken(token) {
  const data = CacheService.getScriptCache().get('reset:' + token);
  return { ok: !!data };
}

function resetPassword(token, newPassword) {
  const raw = CacheService.getScriptCache().get('reset:' + token);
  if (!raw) throw new Error('Token reset tidak valid atau sudah kedaluwarsa.');
  const payload = JSON.parse(raw);
  validatePasswordStrength_(newPassword);
  const ss = getOrCreateDatabase_();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
  migrateUsersSheet_(sheet);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === payload.userId) {
      sheet.getRange(i + 1, 5).setValue(hashPassword_(newPassword));
      sheet.getRange(i + 1, 9).setValue(nowIso_());
      CacheService.getScriptCache().remove('reset:' + token);
      fitSheet_(sheet);
      SpreadsheetApp.flush();
      return { ok: true, message: 'Password berhasil diubah. Silakan login.' };
    }
  }
  throw new Error('Pengguna untuk token reset tidak ditemukan.');
}

function buildDashboardResponse_(user) {
  return {
    ok: true,
    user: sanitizeUser_(user),
    uploads: listUploads_(user),
    folders: listFolders_(user),
    stats: getStats_(user),
    dbInfo: user.role === 'admin' ? {
      spreadsheetId: getOrCreateDatabase_().getId(),
      spreadsheetUrl: getOrCreateDatabase_().getUrl()
    } : null
  };
}

function getDashboardData(sessionToken) {
  const user = requireSession_(sessionToken);
  return buildDashboardResponse_(user);
}

function uploadDocument(sessionToken, payload) {
  const user = requireSession_(sessionToken);
  const tanggal = String(payload.tanggal || '').trim();
  const nomorSurat = String(payload.nomorSurat || '').trim();
  const namaSurat = String(payload.namaSurat || '').trim();
  const jenisSurat = String(payload.jenisSurat || '').trim();
  const folderId = String(payload.folderId || '').trim();
  const fileObject = payload.fileObject || null;

  if (!tanggal || !nomorSurat || !namaSurat || !jenisSurat || !folderId || !fileObject) {
    throw new Error('Tanggal, nomor surat, nama surat, jenis surat, folder, dan file wajib diisi.');
  }

  const folder = getFolderRecordById_(folderId);
  if (!folder || !folder.active) throw new Error('Folder tujuan tidak ditemukan atau tidak aktif.');
  if (!folder.folderId) throw new Error('Folder Google Drive belum tersedia. Admin dapat mengisi link folder atau membuat folder otomatis dari aplikasi/sheet.');

  let driveFolder;
  try {
    driveFolder = DriveApp.getFolderById(folder.folderId);
  } catch (err) {
    throw new Error('Folder Google Drive tidak dapat diakses. Pastikan link/folder ID benar dan izin Drive tersedia.');
  }

  const contentType = String(fileObject.mimeType || 'application/octet-stream');
  const fileName = String(fileObject.fileName || 'file');
  const base64Data = String(fileObject.base64 || '');
  const bytes = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(bytes, contentType, fileName);
  const createdFile = driveFolder.createFile(blob);
  createdFile.setDescription([
    'Nomor Surat: ' + nomorSurat,
    'Nama Surat: ' + namaSurat,
    'Jenis Surat: ' + jenisSurat,
    'Tanggal: ' + tanggal,
    'Pengupload: ' + user.name + ' (' + user.email + ')'
  ].join('\n'));

  const ss = getOrCreateDatabase_();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.UPLOADS);
  sheet.appendRow([
    Utilities.getUuid(),
    tanggal,
    nomorSurat,
    namaSurat,
    jenisSurat,
    folder.category,
    folder.folderName,
    folder.folderId,
    folder.driveLink,
    createdFile.getId(),
    createdFile.getName(),
    createdFile.getUrl(),
    user.name,
    user.email,
    nowIso_()
  ]);
  formatDataRow_(sheet, sheet.getLastRow());
  SpreadsheetApp.flush();

  return {
    ok: true,
    message: 'File berhasil diupload.',
    upload: listUploads_(user)[0]
  };
}

function adminUpdateUpload(sessionToken, payload) {
  requireAdmin_(sessionToken);
  payload = payload || {};
  const id = String(payload.id || '').trim();
  if (!id) throw new Error('ID upload tidak valid.');

  const ss = getOrCreateDatabase_();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.UPLOADS);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0] || '') === id) {
      const tanggal = String(payload.tanggal || '').trim();
      const nomorSurat = String(payload.nomorSurat || '').trim();
      const namaSurat = String(payload.namaSurat || '').trim();
      const jenisSurat = String(payload.jenisSurat || '').trim();
      const kategori = String(payload.kategori || '').trim();
      const subFolder = String(payload.subFolder || '').trim();
      if (!tanggal || !nomorSurat || !namaSurat || !jenisSurat || !kategori || !subFolder) {
        throw new Error('Tanggal, nomor surat, nama surat, jenis surat, kategori, dan folder wajib diisi.');
      }
      sheet.getRange(i + 1, 2, 1, 6).setValues([[tanggal, nomorSurat, namaSurat, jenisSurat, kategori, subFolder]]);
      formatDataRow_(sheet, i + 1);
      SpreadsheetApp.flush();
      return { ok: true, message: 'Data tabel berhasil diperbarui.', uploads: listUploads_(requireSession_(sessionToken)) };
    }
  }
  throw new Error('Data upload tidak ditemukan.');
}

function adminDeleteUpload(sessionToken, uploadId) {
  requireAdmin_(sessionToken);
  uploadId = String(uploadId || '').trim();
  if (!uploadId) throw new Error('ID upload tidak valid.');

  const ss = getOrCreateDatabase_();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.UPLOADS);
  const values = sheet.getDataRange().getValues();

  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][0] || '') === uploadId) {
      const fileId = String(values[i][9] || '').trim();
      let driveMessage = 'File Drive tidak ditemukan pada data.';
      if (fileId) {
        try {
          DriveApp.getFileById(fileId).setTrashed(true);
          driveMessage = 'File di Google Drive juga sudah dipindahkan ke Sampah.';
        } catch (err) {
          driveMessage = 'Baris dihapus, tetapi file Drive tidak dapat dihapus: ' + (err && err.message ? err.message : err);
        }
      }

      sheet.deleteRow(i + 1);
      if (sheet.getLastRow() >= 1) fitSheet_(sheet);
      SpreadsheetApp.flush();
      return {
        ok: true,
        message: 'Data upload berhasil dihapus dari sheet. ' + driveMessage
      };
    }
  }

  throw new Error('Data upload tidak ditemukan.');
}

function adminSyncDriveFilesToUploads(sessionToken) {
  const user = requireAdmin_(sessionToken);
  const ss = getOrCreateDatabase_();
  const uploadSheet = ss.getSheetByName(CONFIG.SHEETS.UPLOADS);
  const folders = listFolders_(user).filter(function(folder) {
    return folder.active && folder.folderId;
  });

  const existingValues = uploadSheet.getDataRange().getValues();
  const existingFileIds = {};
  for (let i = 1; i < existingValues.length; i++) {
    const fileId = String(existingValues[i][9] || '').trim();
    if (fileId) existingFileIds[fileId] = true;
  }

  let count = 0;
  folders.forEach(function(folder) {
    let driveFolder;
    try {
      driveFolder = DriveApp.getFolderById(folder.folderId);
    } catch (err) {
      return;
    }

    const files = driveFolder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      const fileId = file.getId();
      if (existingFileIds[fileId]) continue;

      const createdAt = cellText_(file.getDateCreated());
      uploadSheet.appendRow([
        Utilities.getUuid(),
        createdAt ? createdAt.substring(0, 10) : nowIso_().substring(0, 10),
        '-',
        file.getName(),
        'File Drive',
        folder.category,
        folder.folderName,
        folder.folderId,
        folder.driveLink,
        fileId,
        file.getName(),
        file.getUrl(),
        'Sinkronisasi Drive',
        user.email,
        nowIso_()
      ]);
      formatDataRow_(uploadSheet, uploadSheet.getLastRow());
      existingFileIds[fileId] = true;
      count++;
    }
  });

  SpreadsheetApp.flush();

  return {
    ok: true,
    message: count + ' file dari folder Google Drive berhasil dibaca ke tabel.',
    uploads: listUploads_(user),
    stats: getStats_(user)
  };
}

function getAdminData(sessionToken) {
  const user = requireAdmin_(sessionToken);
  return {
    ok: true,
    user: sanitizeUser_(user),
    users: listUsers_(),
    folders: listFolders_(user)
  };
}
function adminCreateUser(sessionToken, payload) {
  requireAdmin_(sessionToken);
  const nama = String(payload.nama || '').trim();
  const email = String(payload.email || '').trim().toLowerCase();
  const username = normalizeUsername_(payload.username || (email ? email.split('@')[0] : ''));
  const password = String(payload.password || '');
  const role = normalizeRole_(payload.role);
  if (!nama || !email || !username || !password) throw new Error('Nama, email, username, dan password wajib diisi.');
  assertUniqueUserFields_(email, username, '');
  validatePasswordStrength_(password);
  const ss = getOrCreateDatabase_();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
  migrateUsersSheet_(sheet);
  sheet.appendRow([Utilities.getUuid(), nama, email, username, hashPassword_(password), role, 'TRUE', nowIso_(), nowIso_()]);
  fitSheet_(sheet);
  return { ok: true, users: listUsers_() };
}

function adminUpdateUser(sessionToken, payload) {
  const admin = requireAdmin_(sessionToken);
  const id = String(payload.id || '').trim();
  if (!id) throw new Error('ID user tidak valid.');
  const ss = getOrCreateDatabase_();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
  migrateUsersSheet_(sheet);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === id) {
      const newNama = String(payload.nama || values[i][1]).trim();
      const newEmail = String(payload.email || values[i][2]).trim().toLowerCase();
      const newUsername = normalizeUsername_(payload.username || values[i][3]);
      const newRole = normalizeRole_(payload.role || values[i][5]);
      const active = String(payload.active !== undefined ? payload.active : values[i][6]).toLowerCase() === 'true';
      const newPassword = String(payload.password || '');
      if (!newNama || !newEmail || !newUsername) throw new Error('Nama, email, dan username wajib diisi.');
      assertUniqueUserFields_(newEmail, newUsername, id);
      if (newPassword) validatePasswordStrength_(newPassword);
      if (admin.id === id && newRole !== 'admin') throw new Error('Admin yang sedang login tidak boleh menurunkan rolenya sendiri.');
      sheet.getRange(i + 1, 2).setValue(newNama);
      sheet.getRange(i + 1, 3).setValue(newEmail);
      sheet.getRange(i + 1, 4).setValue(newUsername);
      sheet.getRange(i + 1, 6).setValue(newRole);
      sheet.getRange(i + 1, 7).setValue(active ? 'TRUE' : 'FALSE');
      sheet.getRange(i + 1, 9).setValue(nowIso_());
      if (newPassword) sheet.getRange(i + 1, 5).setValue(hashPassword_(newPassword));
      fitSheet_(sheet);
      return { ok: true, users: listUsers_() };
    }
  }
  throw new Error('User tidak ditemukan.');
}

function adminDeleteUser(sessionToken, userId) {
  const admin = requireAdmin_(sessionToken);
  userId = String(userId || '').trim();
  if (!userId) throw new Error('ID user tidak valid.');

  const ss = getOrCreateDatabase_();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
  const values = sheet.getDataRange().getValues();

  for (let i = values.length - 1; i >= 1; i--) {
    if (values[i][0] === userId) {
      if (values[i][0] === admin.id) throw new Error('Admin yang sedang login tidak dapat menghapus akunnya sendiri.');
      sheet.deleteRow(i + 1);
      fitSheet_(sheet);
      return { ok: true, users: listUsers_() };
    }
  }
  throw new Error('User tidak ditemukan.');
}

function adminSaveFolder(sessionToken, payload) {
  requireAdmin_(sessionToken);
  const id = String(payload.id || '').trim();
  const kategori = String(payload.kategori || '').trim();
  const namaFolder = String(payload.namaFolder || '').trim();
  let parentDriveLink = String(payload.parentDriveLink || '').trim();
  let parentFolderId = extractDriveFolderId_(parentDriveLink);
  let driveLink = String(payload.driveLink || '').trim();
  let folderId = extractDriveFolderId_(driveLink);
  const active = String(payload.active !== undefined ? payload.active : true).toLowerCase() === 'true';
  const createInDrive = String(payload.createInDrive || 'false').toLowerCase() === 'true';

  if (!kategori || !namaFolder) throw new Error('Kategori dan nama folder wajib diisi.');
  if (parentDriveLink && !parentFolderId) throw new Error('Link Google Drive parent tidak valid.');
  if (driveLink && !folderId) throw new Error('Link Google Drive folder tidak valid.');

  if (createInDrive && !folderId) {
    const created = createDriveFolderPath_(kategori, namaFolder, parentFolderId);
    parentDriveLink = created.parentFolderUrl;
    parentFolderId = created.parentFolderId;
    driveLink = created.subFolderUrl;
    folderId = created.subFolderId;
  }

  const ss = getOrCreateDatabase_();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.FOLDERS);

  if (!id) {
    sheet.appendRow([
      Utilities.getUuid(),
      kategori,
      namaFolder,
      parentDriveLink,
      parentFolderId || '',
      driveLink,
      folderId || '',
      active ? 'TRUE' : 'FALSE',
      nowIso_(),
      nowIso_()
    ]);
  } else {
    const values = sheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === id) {
        sheet.getRange(i + 1, 2, 1, 9).setValues([[
          kategori,
          namaFolder,
          parentDriveLink,
          parentFolderId || '',
          driveLink,
          folderId || '',
          active ? 'TRUE' : 'FALSE',
          values[i][8] || nowIso_(),
          nowIso_()
        ]]);
        found = true;
        break;
      }
    }
    if (!found) throw new Error('Folder tidak ditemukan.');
  }

  SpreadsheetApp.flush();
  return { ok: true, folders: listFolders_(requireSession_(sessionToken)) };
}

function adminCreateMissingDriveFoldersFromSheet(sessionToken) {
  requireAdmin_(sessionToken);

  const ss = getOrCreateDatabase_();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.FOLDERS);
  migrateFoldersSheet_(sheet);
  const values = sheet.getDataRange().getValues();
  let count = 0;

  for (let i = 1; i < values.length; i++) {
    const kategori = String(values[i][1] || '').trim();
    const namaFolder = String(values[i][2] || '').trim();
    let parentDriveLink = String(values[i][3] || '').trim();
    let parentFolderId = String(values[i][4] || '').trim() || extractDriveFolderId_(parentDriveLink);
    const existingFolderId = String(values[i][6] || '').trim();

    if (!kategori || !namaFolder || existingFolderId) continue;

    const created = createDriveFolderPath_(kategori, namaFolder, parentFolderId);
    parentDriveLink = created.parentFolderUrl;
    parentFolderId = created.parentFolderId;

    sheet.getRange(i + 1, 4).setValue(parentDriveLink);
    sheet.getRange(i + 1, 5).setValue(parentFolderId);
    sheet.getRange(i + 1, 6).setValue(created.subFolderUrl);
    sheet.getRange(i + 1, 7).setValue(created.subFolderId);
    sheet.getRange(i + 1, 10).setValue(nowIso_());
    count++;
  }

  SpreadsheetApp.flush();

  return {
    ok: true,
    message: count + ' folder/sub folder berhasil dibuat dari sheet.',
    folders: listFolders_(requireSession_(sessionToken))
  };
}

function adminDeleteFolder(sessionToken, id) {
  requireAdmin_(sessionToken);
  id = String(id || '').trim();
  if (!id) throw new Error('ID folder tidak valid.');

  const ss = getOrCreateDatabase_();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.FOLDERS);
  const values = sheet.getDataRange().getValues();
  for (let i = values.length - 1; i >= 1; i--) {
    if (values[i][0] === id) {
      sheet.deleteRow(i + 1);
      fitSheet_(sheet);
      return { ok: true, folders: listFolders_(requireSession_(sessionToken)) };
    }
  }
  throw new Error('Folder tidak ditemukan.');
}

function createDriveFolderPath_(kategori, namaFolder, parentFolderId) {
  let base;
  if (parentFolderId) {
    try {
      base = DriveApp.getFolderById(parentFolderId);
    } catch (err) {
      throw new Error('Folder induk Google Drive tidak dapat diakses. Pastikan link/folder ID parent benar dan izin Drive tersedia.');
    }
  } else {
    base = DriveApp.getRootFolder();
  }

  const parent = getOrCreateChildFolder_(base, kategori);
  const sub = getOrCreateChildFolder_(parent, namaFolder);

  return {
    parentFolderId: parent.getId(),
    parentFolderUrl: parent.getUrl(),
    subFolderId: sub.getId(),
    subFolderUrl: sub.getUrl()
  };
}

function getOrCreateChildFolder_(parent, name) {
  const safeName = String(name || '').trim();
  if (!safeName) throw new Error('Nama folder tidak boleh kosong.');
  const folders = parent.getFoldersByName(safeName);
  return folders.hasNext() ? folders.next() : parent.createFolder(safeName);
}

function cellText_(value, pattern) {
  if (value === null || value === undefined) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, 'Asia/Jakarta', pattern || 'yyyy-MM-dd HH:mm:ss');
  }
  return String(value);
}

function cellBool_(value) {
  return String(value || '').trim().toLowerCase() === 'true';
}

function listUploads_(user) {
  const ss = getOrCreateDatabase_();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.UPLOADS);
  const values = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    rows.push({
      id: values[i][0],
      tanggal: cellText_(values[i][1], 'yyyy-MM-dd'),
      nomorSurat: values[i][2],
      namaSurat: values[i][3],
      jenisSurat: values[i][4],
      kategori: values[i][5],
      subFolder: values[i][6],
      driveFolderId: values[i][7],
      driveFolderLink: values[i][8],
      fileId: values[i][9],
      fileName: values[i][10],
      fileUrl: values[i][11],
      uploadedBy: values[i][12],
      uploadedByEmail: values[i][13],
      createdAt: cellText_(values[i][14])
    });
  }
  rows.sort(function(a, b) {
    return String(b.createdAt).localeCompare(String(a.createdAt));
  });
  return rows;
}

function listFolders_(user) {
  const ss = getOrCreateDatabase_();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.FOLDERS);
  migrateFoldersSheet_(sheet);
  const values = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const active = cellBool_(values[i][7]);
    if (user.role !== 'admin' && !active) continue;
    rows.push({
      id: values[i][0],
      category: values[i][1],
      folderName: values[i][2],
      parentDriveLink: values[i][3],
      parentFolderId: values[i][4],
      driveLink: values[i][5],
      folderId: values[i][6],
      active: active,
      createdAt: cellText_(values[i][8]),
      updatedAt: cellText_(values[i][9])
    });
  }
  rows.sort(function(a, b) {
    return String(a.category + a.folderName).localeCompare(String(b.category + b.folderName));
  });
  return rows;
}

function listUsers_() {
  const ss = getOrCreateDatabase_();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
  migrateUsersSheet_(sheet);
  const values = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    rows.push({
      id: values[i][0], nama: values[i][1], email: values[i][2], username: values[i][3],
      role: values[i][5], active: cellBool_(values[i][6]),
      createdAt: cellText_(values[i][7]), updatedAt: cellText_(values[i][8])
    });
  }
  rows.sort(function(a, b) { return String(a.nama).localeCompare(String(b.nama)); });
  return rows;
}

function getStats_(user) {
  const uploads = listUploads_(user);
  const folders = listFolders_(user);
  return {
    totalUploads: uploads.length,
    totalFolders: folders.length,
    totalCategories: unique_(folders.map(function(f) { return f.category; })).length
  };
}

function getFolderRecordById_(id) {
  const ss = getOrCreateDatabase_();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.FOLDERS);
  migrateFoldersSheet_(sheet);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === id) {
      return {
        id: values[i][0],
        category: values[i][1],
        folderName: values[i][2],
        parentDriveLink: values[i][3],
        parentFolderId: values[i][4],
        driveLink: values[i][5],
        folderId: values[i][6],
        active: cellBool_(values[i][7])
      };
    }
  }
  return null;
}
function findUserByEmail_(email) {
  const ss = getOrCreateDatabase_();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
  migrateUsersSheet_(sheet);
  const values = sheet.getDataRange().getValues();
  email = String(email || '').toLowerCase();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][2] || '').toLowerCase() === email) return buildUserFromRow_(values[i]);
  }
  return null;
}

function findUserByUsername_(username) {
  const ss = getOrCreateDatabase_();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
  migrateUsersSheet_(sheet);
  const values = sheet.getDataRange().getValues();
  username = normalizeUsername_(username);
  for (let i = 1; i < values.length; i++) {
    if (normalizeUsername_(values[i][3]) === username) return buildUserFromRow_(values[i]);
  }
  return null;
}

function findUserByLogin_(identifier) {
  identifier = String(identifier || '').trim().toLowerCase();
  if (!identifier) return null;
  return identifier.indexOf('@') > -1 ? findUserByEmail_(identifier) : findUserByUsername_(identifier);
}

function buildUserFromRow_(row) {
  return { id: row[0], name: row[1], email: row[2], username: row[3], passwordHash: row[4], role: row[5], active: cellBool_(row[6]), createdAt: cellText_(row[7]), updatedAt: cellText_(row[8]) };
}

function assertUniqueUserFields_(email, username, ignoreUserId) {
  const ss = getOrCreateDatabase_();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
  migrateUsersSheet_(sheet);
  const values = sheet.getDataRange().getValues();
  email = String(email || '').trim().toLowerCase();
  username = normalizeUsername_(username);
  for (let i = 1; i < values.length; i++) {
    const rowId = String(values[i][0] || '');
    if (ignoreUserId && rowId === ignoreUserId) continue;
    if (String(values[i][2] || '').trim().toLowerCase() === email) throw new Error('Email sudah terdaftar.');
    if (normalizeUsername_(values[i][3]) === username) throw new Error('Username sudah terdaftar.');
  }
}

function normalizeUsername_(username) {
  return String(username || '').trim().toLowerCase().replace(/[^a-z0-9._-]/g, '').replace(/^[._-]+|[._-]+$/g, '');
}
function requireSession_(sessionToken) {
  const token = String(sessionToken || '').trim();
  if (!token) throw new Error('Sesi tidak ditemukan. Silakan login kembali.');
  const cache = CacheService.getScriptCache();
  const raw = cache.get('session:' + token);
  if (!raw) throw new Error('Sesi berakhir karena tidak aktif lebih dari 5 menit. Silakan login kembali.');
  const session = JSON.parse(raw);
  const user = findUserByEmail_(session.email);
  if (!user || !user.active) throw new Error('User tidak aktif atau tidak ditemukan.');
  session.lastActiveAt = nowIso_();
  cache.put('session:' + token, JSON.stringify(session), CONFIG.SESSION_TTL_SECONDS);
  return user;
}

function requireAdmin_(sessionToken) {
  const user = requireSession_(sessionToken);
  if (user.role !== 'admin') throw new Error('Akses admin diperlukan.');
  return user;
}

function sanitizeUser_(user) {
  return { id: user.id, name: user.name || user.nama, email: user.email, username: user.username, role: user.role, active: user.active, createdAt: user.createdAt, updatedAt: user.updatedAt };
}

function normalizeRole_(role) {
  role = String(role || 'user').trim().toLowerCase();
  return role === 'admin' ? 'admin' : 'user';
}

function validatePasswordStrength_(password) {
  password = String(password || '');
  if (password.length < 8) {
    throw new Error('Password minimal 8 karakter.');
  }
}

function hashPassword_(password) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(password), Utilities.Charset.UTF_8);
  return bytes.map(function(b) {
    const v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

function extractDriveFolderId_(urlOrId) {
  const input = String(urlOrId || '').trim();
  if (!input) return '';
  const patterns = [
    /\/folders\/([a-zA-Z0-9_-]+)/,
    /id=([a-zA-Z0-9_-]+)/,
    /^([a-zA-Z0-9_-]{10,})$/
  ];
  for (let i = 0; i < patterns.length; i++) {
    const match = input.match(patterns[i]);
    if (match && match[1]) return match[1];
  }
  return '';
}

function unique_(arr) {
  const map = {};
  const out = [];
  arr.forEach(function(item) {
    const key = String(item || '');
    if (!map[key]) {
      map[key] = true;
      out.push(item);
    }
  });
  return out;
}

function nowIso_() {
  return Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");
}