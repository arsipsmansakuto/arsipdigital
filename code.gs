/**
 * ============================================================================
 * BACKEND PRODUCTION ENGINE: Google Apps Script (Code.gs)
 * ArsipCloud Enterprise v4.1 - Production Cloud Storage & Sheets Database
 * ============================================================================
 */

// Default Root Google Drive ID jika belum dikonfigurasi pada sheet Pengaturan
const DEFAULT_ROOT_DRIVE_FOLDER_ID = "1rxWfplF9QTj4-j0TMPrYTfvRB0_5by7r";

// Definisi Nama Lembar Kerja Database Spreadsheet
const SHEET_NAMES = {
  ARCHIVE: 'Data_Arsip',
  CATEGORY: 'Kategori_Drive',
  USERS: 'Users',
  LOGS: 'Audit_Log',
  SETTINGS: 'Pengaturan'
};

/**
 * Handler HTTP POST utama untuk seluruh operasi CRUD & Sinkronisasi
 * Menggunakan LockService untuk menjamin integritas transaksi data multi-pengguna
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    // Tunggu antrean eksekusi maksimal 15 detik
    lock.waitLock(15000);
    setupDatabase();

    if (!e || !e.postData || !e.postData.contents) {
      return createJsonResponse({
        status: 'ERROR',
        message: 'Payload POST kosong atau tidak valid'
      });
    }

    let contents;
    try {
      contents = JSON.parse(e.postData.contents);
    } catch (parseErr) {
      return createJsonResponse({
        status: 'ERROR',
        message: 'Format data JSON tidak valid: ' + parseErr.toString()
      });
    }

    const action = contents.action;
    let result = { status: 'SUCCESS', message: 'Operasi berhasil dieksekusi' };

    switch (action) {
      case 'GET_SETTINGS':
        result = getSettingsData();
        break;

      case 'GET_ALL_DATA':
        result = getAllData();
        break;

      case 'SAVE_ARCHIVE':
        const saveRes = saveArchive(contents.data || {});
        result.file_url = saveRes.file_url;
        result.drive_file_id = saveRes.drive_file_id;
        break;

      case 'DELETE_ARCHIVE':
        deleteArchive(contents.id, contents.drive_file_id);
        break;

      case 'SAVE_CATEGORY':
        result = saveCategory(contents.data || {});
        break;

      case 'DELETE_CATEGORY':
        result = deleteCategory(contents.id, contents.folder_id);
        break;

      case 'SAVE_USER':
        saveUser(contents.data || {});
        break;

      case 'DELETE_USER':
        deleteUser(contents.username);
        break;

      case 'SAVE_SETTINGS':
        result = saveSettings(contents.data || {});
        break;

      default:
        result = { status: 'ERROR', message: 'Aksi permintaan tidak dikenali: ' + action };
    }

    return createJsonResponse(result);

  } catch (error) {
    Logger.log("doPost Error: " + error.toString());
    return createJsonResponse({ 
      status: 'ERROR', 
      message: error.toString() 
    });
  } finally {
    try {
      lock.releaseLock();
    } catch (lockErr) {
      Logger.log("Lock Release Warning: " + lockErr.toString());
    }
  }
}

/**
 * Handler HTTP GET untuk pengujian status Web App dan pemuatan konfigurasi awal
 */
function doGet(e) {
  setupDatabase();
  const settingsObj = getSettingsData().settings || {};
  const currentRootId = settingsObj.rootDriveId || DEFAULT_ROOT_DRIVE_FOLDER_ID;

  return createJsonResponse({ 
    status: 'ACTIVE', 
    version: 'ArsipCloud Enterprise v4.1 Production Engine',
    root_folder_id: currentRootId,
    settings: settingsObj
  });
}

/**
 * Memformat objek output menjadi MIME Type JSON murni
 */
function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Mengambil direktori folder induk Google Drive berdasarkan konfigurasi Pengaturan
 */
function getRootDriveFolder() {
  let targetFolderId = DEFAULT_ROOT_DRIVE_FOLDER_ID;
  try {
    const settings = getSettingsData().settings;
    if (settings && settings.rootDriveId && String(settings.rootDriveId).trim() !== "") {
      targetFolderId = String(settings.rootDriveId).trim();
    }
  } catch (e) {
    Logger.log("Gagal membaca rootDriveId dari pengaturan: " + e.toString());
  }

  try {
    return DriveApp.getFolderById(targetFolderId);
  } catch (e) {
    Logger.log("Folder Google Drive (" + targetFolderId + ") tidak ditemukan, mengalihkan ke folder lokal: " + e.toString());
    return getOrCreateDriveFolder("ArsipCloud_Enterprise_Drive");
  }
}

/**
 * Mencari atau membuat folder Google Drive di bawah folder induk
 */
function getOrCreateDriveFolder(folderName, parentFolder = null) {
  const targetParent = parentFolder || getRootDriveFolder();
  const folders = targetParent.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  const newFld = targetParent.createFolder(folderName);
  newFld.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return newFld;
}

/**
 * Zero-Config Database Setup: Memeriksa dan membuat 5 sheet lengkap dengan header kolom
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Tab Sheet: Data_Arsip
  let sheetArchive = ss.getSheetByName(SHEET_NAMES.ARCHIVE);
  if (!sheetArchive) {
    sheetArchive = ss.insertSheet(SHEET_NAMES.ARCHIVE);
    sheetArchive.appendRow([
      'ID', 'No. Dokumen', 'Nama Arsip', 'Keterangan', 'Kategori', 
      'Tanggal', 'Ukuran (KB)', 'Nama Berkas', 'URL Berkas', 'Pengunggah', 'Drive File ID'
    ]);
    sheetArchive.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#f1f5f9');
  }

  // 2. Tab Sheet: Kategori_Drive
  let sheetCategory = ss.getSheetByName(SHEET_NAMES.CATEGORY);
  if (!sheetCategory) {
    sheetCategory = ss.insertSheet(SHEET_NAMES.CATEGORY);
    sheetCategory.appendRow(['ID Kategori', 'Nama Kategori', 'Folder ID Google Drive', 'Keterangan', 'Status']);
    sheetCategory.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#f1f5f9');
    
    sheetCategory.appendRow(['CAT-01', 'Keuangan & Perpajakan', 'FLD_KEU_8372', 'Laporan neraca, pajak, dan audit tahunan', 'ACTIVE']);
    sheetCategory.appendRow(['CAT-02', 'SDM & Kepegawaian', 'FLD_HRD_1928', 'SK pegawai, kontrak kerja, dan absensi', 'ACTIVE']);
    sheetCategory.appendRow(['CAT-03', 'Legalitas & Notaris', 'FLD_LGL_4412', 'Akta pendirian, NIB, dan perjanjian MOU', 'ACTIVE']);
    sheetCategory.appendRow(['CAT-04', 'Operasional Proyek', 'FLD_OPS_9921', 'RAB, dokumen tender, dan berita acara', 'ACTIVE']);
  }

  // 3. Tab Sheet: Users (dengan kolom Password & Hak Akses)
  let sheetUsers = ss.getSheetByName(SHEET_NAMES.USERS);
  if (!sheetUsers) {
    sheetUsers = ss.insertSheet(SHEET_NAMES.USERS);
    sheetUsers.appendRow(['Username', 'Password', 'Nama Lengkap', 'Role Hak Akses', 'Status Akun']);
    sheetUsers.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#f1f5f9');
    sheetUsers.appendRow(['admin', 'admin123', 'Administrator Utama', 'Admin', 'Aktif']);
    sheetUsers.appendRow(['operator', 'operator123', 'Budi Perkasa', 'Operator', 'Aktif']);
    sheetUsers.appendRow(['viewer', 'viewer123', 'Tamu Peninjau', 'Viewer', 'Aktif']);
  } else {
    // Migrasi cerdas jika sheet Users versi lama belum memiliki kolom Password
    const headerRow = sheetUsers.getRange(1, 1, 1, sheetUsers.getLastColumn()).getValues()[0];
    const headerStr = headerRow.map(h => String(h).toLowerCase()).join(' ');
    if (!headerStr.includes('password')) {
      sheetUsers.insertColumnAfter(1);
      sheetUsers.getRange(1, 2).setValue('Password').setFontWeight('bold').setBackground('#f1f5f9');
      const numRows = sheetUsers.getLastRow();
      if (numRows > 1) {
        for (let r = 2; r <= numRows; r++) {
          const uVal = String(sheetUsers.getRange(r, 1).getValue()).toLowerCase();
          sheetUsers.getRange(r, 2).setValue(uVal ? uVal + '123' : 'admin123');
        }
      }
    }
  }

  // 4. Tab Sheet: Audit_Log
  let sheetLogs = ss.getSheetByName(SHEET_NAMES.LOGS);
  if (!sheetLogs) {
    sheetLogs = ss.insertSheet(SHEET_NAMES.LOGS);
    sheetLogs.appendRow(['Waktu', 'Username', 'Aktivitas', 'Rincian']);
    sheetLogs.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#f1f5f9');
  }

  // 5. Tab Sheet: Pengaturan
  let sheetSettings = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!sheetSettings) {
    sheetSettings = ss.insertSheet(SHEET_NAMES.SETTINGS);
    sheetSettings.appendRow(['Key', 'Value']);
    sheetSettings.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#f1f5f9');
  }
}

/**
 * Membaca seluruh pasangan kunci & nilai dari sheet Pengaturan
 */
function getSettingsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  let settingsObj = {
    appName: "ArsipCloud Enterprise",
    appDesc: "Sistem Manajemen Arsip Digital Terintegrasi Google Drive",
    appLogo: "https://lh3.googleusercontent.com/d/1rxWfplF9QTj4-j0TMPrYTfvRB0_5by7r",
    bgUrl: "",
    bgOpacity: 100,
    bgBlur: 0,
    sidebarBg: "#0f172a",
    sidebarText: "#f8fafc",
    sidebarActive: "#38bdf8",
    rootDriveId: DEFAULT_ROOT_DRIVE_FOLDER_ID,
    maxUploadSizeMB: 20,
    allowedExtensions: "jpg,jpeg,png,gif,webp,svg,pdf,doc,docx,xls,xlsx,csv,ppt,pptx,txt,rtf,odt,ods,odp,zip,rar"
  };

  if (settingsSheet && settingsSheet.getLastRow() > 1) {
    const sRows = settingsSheet.getRange(2, 1, settingsSheet.getLastRow() - 1, 2).getValues();
    for (let i = 0; i < sRows.length; i++) {
      try {
        settingsObj[sRows[i][0]] = JSON.parse(sRows[i][1]);
      } catch (e) {
        settingsObj[sRows[i][0]] = sRows[i][1];
      }
    }
  }
  return {
    status: 'SUCCESS',
    settings: settingsObj
  };
}

/**
 * Mengambil seluruh data dari 5 sheet untuk inisialisasi frontend
 */
function getAllData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Data Arsip
  const archiveSheet = ss.getSheetByName(SHEET_NAMES.ARCHIVE);
  const archives = archiveSheet && archiveSheet.getLastRow() > 1 
    ? archiveSheet.getDataRange().getValues().slice(1).map(r => ({
        id: String(r[0]),
        no: String(r[1]),
        title: String(r[2]),
        desc: String(r[3] || ''),
        category: String(r[4]),
        uploadedAt: String(r[5]),
        size: parseInt(r[6], 10) || 0,
        filename: String(r[7] || ''),
        driveUrl: String(r[8] || ''),
        uploader: String(r[9] || 'admin'),
        driveFileId: String(r[10] || '')
      })) 
    : [];

  // 2. Data Kategori
  const categorySheet = ss.getSheetByName(SHEET_NAMES.CATEGORY);
  const categories = categorySheet && categorySheet.getLastRow() > 1 
    ? categorySheet.getDataRange().getValues().slice(1).map(r => ({
        id: String(r[0]),
        name: String(r[1]),
        driveId: String(r[2]),
        desc: String(r[3] || ''),
        status: String(r[4] || 'ACTIVE')
      })) 
    : [];

  // 3. Data Pengguna
  const usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);
  let users = [];
  if (usersSheet && usersSheet.getLastRow() > 1) {
    const data = usersSheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).toLowerCase().trim());
    const uIdx = headers.indexOf('username');
    const pIdx = headers.indexOf('password');
    const nIdx = headers.findIndex(h => h.includes('nama'));
    const rIdx = headers.findIndex(h => h.includes('role'));
    const sIdx = headers.findIndex(h => h.includes('status'));

    users = data.slice(1).map(r => {
      const uname = String(uIdx !== -1 ? r[uIdx] : r[0]).trim();
      let pwd = uname ? uname.toLowerCase() + '123' : 'admin123';
      if (pIdx !== -1 && r[pIdx] && String(r[pIdx]).trim() !== '') {
        pwd = String(r[pIdx]).trim();
      }
      return {
        username: uname,
        password: pwd,
        name: String(nIdx !== -1 ? r[nIdx] : (r[2] || uname)),
        role: String(rIdx !== -1 ? r[rIdx] : (r[3] || 'Operator')),
        status: String(sIdx !== -1 ? r[sIdx] : (r[4] || 'Aktif'))
      };
    });
  }

  // 4. Data Log Aktivitas
  const logsSheet = ss.getSheetByName(SHEET_NAMES.LOGS);
  const logs = logsSheet && logsSheet.getLastRow() > 1 
    ? logsSheet.getDataRange().getValues().slice(1).map(r => ({
        timestamp: String(r[0]),
        user: String(r[1]),
        action: String(r[2]),
        detail: String(r[3])
      })).reverse() 
    : [];

  const settingsObj = getSettingsData().settings;

  return {
    status: 'SUCCESS',
    archives,
    categories,
    users,
    logs,
    settings: settingsObj
  };
}

/**
 * Menyimpan atau memperbarui data dokumen arsip beserta pengunggahan fisik ke Google Drive
 */
function saveArchive(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.ARCHIVE);
  let fileDirectUrl = data.driveUrl || data.file_url || "";
  let driveFileId = data.driveFileId || data.drive_file_id || "";

  let targetFolderId = getCategoryFolderId(data.category || data.kategori);
  const rawBase64 = data.contentData || data.file_data || (fileDirectUrl.startsWith("data:") ? fileDirectUrl : "");

  // Jika payload berisi file Base64 murni, konversi dan simpan ke folder Google Drive
  if (rawBase64 && String(rawBase64).startsWith("data:")) {
    const uploadObj = uploadBase64ToDrive(
      rawBase64, 
      data.filename || data.file_name || (data.no + "." + (data.ext || "pdf").toLowerCase()), 
      data.category || data.kategori, 
      targetFolderId
    );
    if (uploadObj.url) fileDirectUrl = uploadObj.url;
    if (uploadObj.file_id) driveFileId = uploadObj.file_id;
  }

  const rows = sheet.getDataRange().getValues();
  let foundIndex = -1;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).toLowerCase() === String(data.id).toLowerCase()) {
      foundIndex = i + 1;
      break;
    }
  }

  const rowData = [
    data.id,
    data.no,
    data.title || data.nama,
    data.desc || data.deskripsi || "",
    data.category || data.kategori,
    data.uploadedAt || data.tanggal,
    data.size || 0,
    data.filename || data.file_name || "",
    fileDirectUrl,
    data.uploader || "admin",
    driveFileId
  ];

  if (foundIndex > 0) {
    sheet.getRange(foundIndex, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }

  logAction(data.uploader || 'admin', foundIndex > 0 ? 'EDIT_ARCHIVE' : 'ADD_ARCHIVE', `Dokumen ${data.no} - ${data.title || data.nama} disimpan`);
  return { file_url: fileDirectUrl, drive_file_id: driveFileId };
}

/**
 * Mengunggah data Base64 ke dalam direktori Google Drive dan memulangkan URL pratinjau langsung
 */
function uploadBase64ToDrive(base64Data, fileName, categoryName, targetFolderId) {
  try {
    let parentFolder;
    
    if (targetFolderId && String(targetFolderId).trim() !== "" && !String(targetFolderId).startsWith("FLD_")) {
      try {
        parentFolder = DriveApp.getFolderById(String(targetFolderId).trim());
      } catch (fErr) {
        Logger.log("Folder ID kategori kustom tidak ditemukan: " + fErr.toString());
      }
    }

    if (!parentFolder) {
      parentFolder = getRootDriveFolder();
    }

    const splitData = base64Data.split(',');
    if (splitData.length < 2) return { url: "", file_id: "" };

    const mimeMatch = splitData[0].match(/:(.*?);/);
    const contentType = mimeMatch ? mimeMatch[1] : "application/octet-stream";
    const decoded = Utilities.base64Decode(splitData[1]);
    const blob = Utilities.newBlob(decoded, contentType, fileName || "Dokumen_Arsip");

    const file = parentFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return {
      url: "https://lh3.googleusercontent.com/d/" + file.getId(),
      file_id: file.getId()
    };
  } catch (err) {
    Logger.log("Drive Upload Error: " + err.toString());
    return { url: "", file_id: "" };
  }
}

/**
 * Mengambil Folder ID Google Drive yang terikat pada nama kategori
 */
function getCategoryFolderId(categoryName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.CATEGORY);
  if (!sheet || sheet.getLastRow() <= 1) return "";
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).toLowerCase() === String(categoryName).toLowerCase()) {
      return String(rows[i][2]);
    }
  }
  return "";
}

/**
 * Menyimpan atau memperbarui folder kategori pada Google Drive
 */
function saveCategory(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.CATEGORY);
  const rows = sheet.getDataRange().getValues();
  let foundIndex = -1;

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      foundIndex = i + 1;
      break;
    }
  }

  let folderId = data.driveId || data.folder_id ? String(data.driveId || data.folder_id).trim() : "";

  // Otomatis buat folder Google Drive baru di bawah Root Drive jika belum ditentukan
  if (!folderId || folderId === "" || folderId.startsWith("FLD_") || folderId.toUpperCase() === "AUTO") {
    try {
      const rootFolder = getRootDriveFolder();
      const newFolder = rootFolder.createFolder(data.name || data.nama);
      newFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      folderId = newFolder.getId();
    } catch (e) {
      Logger.log("Gagal membuat folder di root Google Drive: " + e.toString());
      folderId = `FLD_${Date.now().toString().slice(-6)}`;
    }
  }

  const rowData = [
    data.id, 
    data.name || data.nama, 
    folderId, 
    data.desc || data.keterangan || "", 
    data.status || "ACTIVE"
  ];

  if (foundIndex > 0) {
    sheet.getRange(foundIndex, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }

  logAction(data.uploader || 'admin', foundIndex > 0 ? 'EDIT_CATEGORY' : 'ADD_CATEGORY', `Kategori "${data.name || data.nama}" terhubung ke Drive Folder ID: ${folderId}`);
  return { 
    status: 'SUCCESS', 
    message: `Kategori "${data.name || data.nama}" berhasil disimpan!`,
    category: {
      id: data.id,
      name: data.name || data.nama,
      driveId: folderId,
      desc: data.desc || data.keterangan || "",
      status: data.status || 'ACTIVE'
    }
  };
}

/**
 * Menghapus kategori dan memindahkan folder Google Drive ke tempat sampah
 */
function deleteCategory(id, folderId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.CATEGORY);
  if (!sheet || sheet.getLastRow() <= 1) return { status: 'SUCCESS' };
  const rows = sheet.getDataRange().getValues();
  let targetFolderId = folderId || "";

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).toLowerCase() === String(id).toLowerCase()) {
      if (!targetFolderId && rows[i][2]) {
        targetFolderId = String(rows[i][2]);
      }
      sheet.deleteRow(i + 1);
      break;
    }
  }

  // Pindahkan folder fisik di Google Drive ke tempat sampah (Trash)
  if (targetFolderId && String(targetFolderId).trim() !== "" && !targetFolderId.startsWith("FLD_")) {
    try {
      const folder = DriveApp.getFolderById(String(targetFolderId).trim());
      folder.setTrashed(true);
      Logger.log("Folder Google Drive berhasil dipindahkan ke sampah: " + targetFolderId);
    } catch (fErr) {
      Logger.log("Gagal memindahkan folder ke sampah: " + fErr.toString());
    }
  }

  logAction('admin', 'DELETE_CATEGORY', `Kategori ID ${id} (${targetFolderId}) dihapus`);
  return { status: 'SUCCESS', message: 'Kategori dan folder Drive berhasil dihapus' };
}

/**
 * Menghapus data arsip dari lembar kerja dan memindahkan berkas fisiknya ke Trash Google Drive
 */
function deleteArchive(id, driveFileId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.ARCHIVE);
  if (!sheet || sheet.getLastRow() <= 1) return;
  const rows = sheet.getDataRange().getValues();

  let targetDriveId = driveFileId || "";

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).toLowerCase() === String(id).toLowerCase()) {
      if (!targetDriveId && rows[i][10]) {
        targetDriveId = String(rows[i][10]);
      }
      sheet.deleteRow(i + 1);
      break;
    }
  }

  if (targetDriveId && String(targetDriveId).trim() !== "") {
    try {
      const file = DriveApp.getFileById(String(targetDriveId).trim());
      file.setTrashed(true);
      Logger.log("Berkas Google Drive berhasil dipindahkan ke sampah: " + targetDriveId);
    } catch (err) {
      Logger.log("Gagal memindahkan berkas Google Drive ke sampah: " + err.toString());
    }
  }
}

/**
 * Menyimpan atau memperbarui data akun pengguna dan kata sandi
 */
function saveUser(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.USERS);
  const rows = sheet.getDataRange().getValues();
  let foundIndex = -1;

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).toLowerCase() === String(data.username).toLowerCase()) {
      foundIndex = i + 1;
      break;
    }
  }

  const rowData = [
    data.username, 
    data.password || 'admin123', 
    data.name, 
    data.role, 
    data.status
  ];

  if (foundIndex > 0) {
    sheet.getRange(foundIndex, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
  logAction('admin', foundIndex > 0 ? 'EDIT_USER' : 'ADD_USER', `Pengguna @${data.username} (${data.role}) disimpan`);
}

/**
 * Menghapus akun pengguna dari lembar kerja
 */
function deleteUser(username) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.USERS);
  if (!sheet || sheet.getLastRow() <= 1) return;
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).toLowerCase() === String(username).toLowerCase()) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  logAction('admin', 'DELETE_USER', `Pengguna @${username} dihapus dari sistem`);
}

/**
 * Menyimpan konfigurasi sistem, wallpaper, dan batas upload ke sheet Pengaturan
 */
function saveSettings(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.SETTINGS);
  }
  sheet.clearContents();
  sheet.appendRow(['Key', 'Value']);
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#f1f5f9');

  let sanitizedSettings = {};

  if (data && typeof data === 'object') {
    const assetFolder = getOrCreateDriveFolder("ArsipCloud_System_Assets");
    
    Object.keys(data).forEach(k => {
      let val = data[k];

      // Jika wallpaper diunggah dalam format Base64 besar, simpan ke Drive agar spreadsheet tetap ringan
      if (typeof val === 'string' && val.startsWith('data:image')) {
        try {
          const splitData = val.split(',');
          if (splitData.length >= 2) {
            const mimeMatch = splitData[0].match(/:(.*?);/);
            const contentType = mimeMatch ? mimeMatch[1] : "image/png";
            const decoded = Utilities.base64Decode(splitData[1]);
            const ext = contentType.includes('jpeg') || contentType.includes('jpg') ? '.jpg' : '.png';
            const blob = Utilities.newBlob(decoded, contentType, "Asset_" + k + "_" + Date.now() + ext);
            
            const file = assetFolder.createFile(blob);
            file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            val = "https://lh3.googleusercontent.com/d/" + file.getId();
          }
        } catch (assetErr) {
          Logger.log("Gagal mengonversi aset gambar ke Drive: " + assetErr.toString());
        }
      }

      sanitizedSettings[k] = val;
      sheet.appendRow([k, JSON.stringify(val)]);
    });
  }

  logAction('system', 'SAVE_SETTINGS', 'Konfigurasi sistem, batas unggah, dan tema visual diperbarui');
  return { status: 'SUCCESS', settings: sanitizedSettings };
}

/**
 * Mencatat rekam jejak aktivitas audit sistem ke sheet Audit_Log
 */
function logAction(user, action, detail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.LOGS);
  if (!sheet) return;
  const timeStr = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss");
  sheet.appendRow([timeStr, user || 'system', action, detail]);
}
