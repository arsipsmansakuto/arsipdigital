/**
 * ============================================================================
 * BACKEND PRODUCTION ENGINE: Google Apps Script (Code.gs)
 * ArsipCloud Enterprise v3.8 - Centralized Drive Folder ID & Synchronized Trash
 * ============================================================================
 */

const ROOT_DRIVE_FOLDER_ID = "1rxWfplF9QTj4-j0TMPrYTfvRB0_5by7r";

const SHEET_NAMES = {
  ARCHIVE: 'Data_Arsip',
  CATEGORY: 'Kategori_Drive',
  USERS: 'Users',
  LOGS: 'Audit_Log',
  SETTINGS: 'Pengaturan'
};

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
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

function doGet(e) {
  setupDatabase();
  const settingsObj = getSettingsData().settings || {};
  return createJsonResponse({ 
    status: 'ACTIVE', 
    version: 'ArsipCloud Enterprise v3.8 Production Engine',
    root_folder_id: ROOT_DRIVE_FOLDER_ID,
    settings: settingsObj
  });
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getRootDriveFolder() {
  try {
    return DriveApp.getFolderById(ROOT_DRIVE_FOLDER_ID);
  } catch (e) {
    Logger.log("Gagal mengakses ROOT_DRIVE_FOLDER_ID (" + ROOT_DRIVE_FOLDER_ID + "): " + e.toString());
    return getOrCreateDriveFolder("ArsipCloud_Enterprise_Drive");
  }
}

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
    sheetCategory.appendRow(['ID Kategori', 'Nama Kategori', 'Folder ID Google Drive', 'Status']);
    sheetCategory.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#f1f5f9');
    
    const suratFolder = getOrCreateDriveFolder("Surat Masuk & Keluar");
    const invoiceFolder = getOrCreateDriveFolder("Invoice & Perpajakan");
    const hrdFolder = getOrCreateDriveFolder("Dokumen Kepegawaian");

    sheetCategory.appendRow(['CAT-101', 'Surat Masuk / Keluar', suratFolder.getId(), 'ACTIVE']);
    sheetCategory.appendRow(['CAT-102', 'Invoice & Tagihan', invoiceFolder.getId(), 'ACTIVE']);
    sheetCategory.appendRow(['CAT-103', 'Dokumen Kepegawaian', hrdFolder.getId(), 'ACTIVE']);
  }

  // 3. Tab Sheet: Users
  let sheetUsers = ss.getSheetByName(SHEET_NAMES.USERS);
  if (!sheetUsers) {
    sheetUsers = ss.insertSheet(SHEET_NAMES.USERS);
    sheetUsers.appendRow(['Username', 'Password', 'Nama Lengkap', 'Role Hak Akses', 'Status Akun']);
    sheetUsers.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#f1f5f9');
    sheetUsers.appendRow(['admin', 'admin123', 'Administrator System', 'ADMINISTRATOR', 'ACTIVE']);
    sheetUsers.appendRow(['operator', 'operator123', 'Staf Operator Arsip', 'OPERATOR', 'ACTIVE']);
    sheetUsers.appendRow(['viewer', 'viewer123', 'Auditor External', 'VIEWER', 'ACTIVE']);
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

function getSettingsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  let settingsObj = {};
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

function getAllData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const archiveSheet = ss.getSheetByName(SHEET_NAMES.ARCHIVE);
  const archives = archiveSheet && archiveSheet.getLastRow() > 1 
    ? archiveSheet.getDataRange().getValues().slice(1).map(r => ({
        id: String(r[0]),
        no: String(r[1]),
        nama: String(r[2]),
        deskripsi: String(r[3]),
        kategori: String(r[4]),
        tanggal: String(r[5]),
        size: parseInt(r[6], 10) || 0,
        file_name: String(r[7] || ''),
        file_url: String(r[8] || ''),
        uploader: String(r[9] || 'admin'),
        drive_file_id: String(r[10] || '')
      })) 
    : [];

  const categorySheet = ss.getSheetByName(SHEET_NAMES.CATEGORY);
  const categories = categorySheet && categorySheet.getLastRow() > 1 
    ? categorySheet.getDataRange().getValues().slice(1).map(r => ({
        id: String(r[0]),
        nama: String(r[1]),
        folder_id: String(r[2]),
        status: String(r[3])
      })) 
    : [];

  const usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);
  const users = usersSheet && usersSheet.getLastRow() > 1 
    ? usersSheet.getDataRange().getValues().slice(1).map(r => ({
        username: String(r[0]),
        password: String(r[1] || 'admin123'),
        name: String(r[2]),
        role: String(r[3]),
        status: String(r[4])
      })) 
    : [];

  const logsSheet = ss.getSheetByName(SHEET_NAMES.LOGS);
  const logs = logsSheet && logsSheet.getLastRow() > 1 
    ? logsSheet.getDataRange().getValues().slice(1).map(r => ({
        time: String(r[0]),
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

function saveArchive(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.ARCHIVE);
  let fileDirectUrl = data.file_url || "";
  let driveFileId = data.drive_file_id || "";

  let targetFolderId = getCategoryFolderId(data.kategori);
  const rawBase64 = data.file_data || (data.file_url && String(data.file_url).startsWith("data:") ? data.file_url : "");

  if (rawBase64 && String(rawBase64).startsWith("data:")) {
    const uploadObj = uploadBase64ToDrive(rawBase64, data.file_name || (data.no + "." + (data.file_type || "pdf").toLowerCase()), data.kategori, targetFolderId);
    if (uploadObj.url) fileDirectUrl = uploadObj.url;
    if (uploadObj.file_id) driveFileId = uploadObj.file_id;
  }

  const rows = sheet.getDataRange().getValues();
  let foundIndex = -1;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      foundIndex = i + 1;
      break;
    }
  }

  const rowData = [
    data.id,
    data.no,
    data.nama,
    data.deskripsi || "",
    data.kategori,
    data.tanggal,
    data.size || 0,
    data.file_name || "",
    fileDirectUrl,
    data.uploader || "admin",
    driveFileId
  ];

  if (foundIndex > 0) {
    sheet.getRange(foundIndex, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }

  logAction(data.uploader, foundIndex > 0 ? 'EDIT_ARCHIVE' : 'ADD_ARCHIVE', `Dokumen ${data.no} - ${data.nama} disimpan`);
  return { file_url: fileDirectUrl, drive_file_id: driveFileId };
}

function uploadBase64ToDrive(base64Data, fileName, categoryName, targetFolderId) {
  try {
    let parentFolder;
    
    if (targetFolderId && String(targetFolderId).trim() !== "") {
      try {
        parentFolder = DriveApp.getFolderById(String(targetFolderId).trim());
      } catch (fErr) {
        Logger.log("Folder ID Kategori tidak ditemukan, mengalihkan ke root parent: " + fErr.toString());
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

  let folderId = data.folder_id ? String(data.folder_id).trim() : "";

  // Buat folder baru di dalam folder Google Drive root utama jika ID kosong/otomatis
  if (!folderId || folderId === "" || folderId.startsWith("FLD_") || folderId.toUpperCase() === "AUTO" || folderId.startsWith("AUTO_")) {
    try {
      const rootFolder = getRootDriveFolder();
      const newFolder = rootFolder.createFolder(data.nama);
      newFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      folderId = newFolder.getId();
      Logger.log("Folder Google Drive baru berhasil dibuat di root folder (" + ROOT_DRIVE_FOLDER_ID + "): " + data.nama + " (ID: " + folderId + ")");
    } catch (e) {
      Logger.log("Gagal membuat folder di root Google Drive: " + e.toString());
      folderId = getOrCreateDriveFolder(data.nama).getId();
    }
  } else {
    try {
      const existingFolder = DriveApp.getFolderById(folderId);
    } catch (err) {
      Logger.log("Folder ID kustom tidak ditemukan, membuat folder baru di root: " + err.toString());
      const rootFolder = getRootDriveFolder();
      const fallbackFolder = rootFolder.createFolder(data.nama);
      fallbackFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      folderId = fallbackFolder.getId();
    }
  }

  const rowData = [data.id, data.nama, folderId, data.status || "ACTIVE"];
  if (foundIndex > 0) {
    sheet.getRange(foundIndex, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }

  logAction(data.uploader || 'admin', foundIndex > 0 ? 'EDIT_CATEGORY' : 'ADD_CATEGORY', `Kategori "${data.nama}" tersambung ke Google Drive Folder ID: ${folderId}`);
  return { 
    status: 'SUCCESS', 
    message: `Folder Google Drive "${data.nama}" berhasil dihubungkan!`,
    category: {
      id: data.id,
      nama: data.nama,
      folder_id: folderId,
      status: data.status || 'ACTIVE'
    }
  };
}

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

  if (targetFolderId && String(targetFolderId).trim() !== "" && !targetFolderId.startsWith("FLD_") && !targetFolderId.startsWith("AUTO_")) {
    try {
      const folder = DriveApp.getFolderById(String(targetFolderId).trim());
      folder.setTrashed(true);
      Logger.log("Folder Google Drive kategori berhasil dipindahkan ke sampah: " + targetFolderId);
    } catch (fErr) {
      Logger.log("Gagal memindahkan folder Google Drive ke sampah: " + fErr.toString());
    }
  }

  logAction('admin', 'DELETE_CATEGORY', `Kategori ID ${id} dan folder Google Drive (${targetFolderId}) berhasil dihapus`);
  return { status: 'SUCCESS', message: 'Kategori dan folder Drive berhasil dihapus' };
}

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
      Logger.log("Gagal menghapus berkas di Google Drive: " + err.toString());
    }
  }
}

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

  const rowData = [data.username, data.password || 'admin123', data.name, data.role, data.status];
  if (foundIndex > 0) {
    sheet.getRange(foundIndex, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
}

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
}

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
          Logger.log("Gagal mengonversi asset gambar ke Drive: " + assetErr.toString());
        }
      }

      sanitizedSettings[k] = val;
      sheet.appendRow([k, JSON.stringify(val)]);
    });
  }

  logAction('system', 'SAVE_SETTINGS', 'Konfigurasi sistem & tema visual diperbarui');
  return { status: 'SUCCESS', settings: sanitizedSettings };
}

function logAction(user, action, detail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.LOGS);
  if (!sheet) return;
  const timeStr = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm");
  sheet.appendRow([timeStr, user || 'system', action, detail]);
}
