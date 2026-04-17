/**
 * KONFIGURASI DASHBOARD IT
 * Ganti MAIN_FOLDER_ID dengan ID folder Google Drive Anda
 */
const MAIN_FOLDER_ID = '1TImiBi0slgDECtrLY6vIqnczOCaOxjrO';

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Dashboard IT')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ----------------------------------------------------
// DIAGNOSTIC & TROUBLESHOOTING
// ----------------------------------------------------

/**
 * Test akses ke Google Drive dan folder utama
 * Jalankan ini untuk memeriksa permission
 */
function testDriveAccess() {
  try {
    Logger.log('Testing Drive access...');
    
    // Test akses ke DriveApp
    const driveTest = DriveApp.getRootFolder();
    Logger.log('DriveApp access OK');
    
    // Test akses ke folder utama
    const mainFolder = DriveApp.getFolderById(MAIN_FOLDER_ID);
    Logger.log('Main folder access OK: ' + mainFolder.getName());
    
    // Test create folder
    const testFolderName = 'TEST_FOLDER_' + new Date().getTime();
    const testFolder = mainFolder.createFolder(testFolderName);
    Logger.log('Folder creation OK: ' + testFolder.getName());
    
    // Cleanup test folder
    testFolder.setTrashed(true);
    Logger.log('Test folder cleanup OK');
    
    return { 
      success: true, 
      message: 'Akses Drive berhasil! Folder utama: ' + mainFolder.getName()
    };
  } catch (e) {
    Logger.log('Drive access test failed: ' + e.toString());
    return { 
      success: false, 
      message: 'Akses Drive gagal: ' + e.toString() + '\n\nPastikan:\n1. Folder ID benar\n2. Apps Script memiliki permission Drive\n3. Folder dapat diakses'
    };
  }
}

/**
 * Perbaiki record arsip yang hilang
 * Cari file di Drive yang belum tercatat di sheet
 */
function repairMissingArchives() {
  try {
    Logger.log('Starting archive repair...');
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Archives');
    const existingRecords = sheet.getDataRange().getValues();
    const existingFileIds = existingRecords.slice(1).map(row => row[6]); // FileID column
    
    const mainFolder = DriveApp.getFolderById(MAIN_FOLDER_ID);
    const allFiles = getAllFilesInFolder(mainFolder);
    
    let repaired = 0;
    let skipped = 0;
    
    for (const file of allFiles) {
      if (!existingFileIds.includes(file.getId())) {
        // File belum tercatat, tambahkan record
        const id = 'ARK-' + new Date().getTime();
        const timestamp = file.getDateCreated();
        const fileName = file.getName();
        
        // Coba tentukan kategori dari folder
        let kategori = 'General';
        try {
          const parents = file.getParents();
          if (parents.hasNext()) {
            const parentFolder = parents.next();
            if (parentFolder.getId() !== MAIN_FOLDER_ID) {
              kategori = parentFolder.getName();
            }
          }
        } catch (e) {
          Logger.log('Could not determine category for file: ' + fileName);
        }
        
        sheet.appendRow([id, timestamp, fileName, kategori, 'Record diperbaiki otomatis', file.getUrl(), file.getId(), file.getMimeType()]);
        repaired++;
        Logger.log('Repaired record for: ' + fileName);
      } else {
        skipped++;
      }
    }
    
    return { 
      success: true, 
      message: 'Perbaikan selesai. File diperbaiki: ' + repaired + ', dilewati: ' + skipped
    };
  } catch (e) {
    Logger.log('Archive repair failed: ' + e.toString());
    return { 
      success: false, 
      message: 'Perbaikan gagal: ' + e.toString()
    };
  }
}

/**
 * Helper: Get all files recursively in folder
 */
function getAllFilesInFolder(folder) {
  const files = [];
  const subFolders = folder.getFolders();
  
  // Get files in current folder
  const folderFiles = folder.getFiles();
  while (folderFiles.hasNext()) {
    files.push(folderFiles.next());
  }
  
  // Recursively get files in subfolders
  while (subFolders.hasNext()) {
    const subFolder = subFolders.next();
    files.push(...getAllFilesInFolder(subFolder));
  }
  
  return files;
}
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const sheets = {
    'Users': ['Username', 'Password', 'Role'],
    'Categories': ['Nama Kategori'],
    'Archives': ['ID', 'Timestamp', 'File Name', 'Category', 'Description', 'URL', 'FileID', 'MIME Type'],
    'Links': ['ID', 'Kategori', 'Judul', 'Deskripsi', 'URL', 'Jenis'],
    'DashboardCategories': ['Kategori Dashboard', 'Order']
  };

  for (let name in sheets) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(sheets[name]);
      
      if (name === 'Users') {
        sheet.appendRow(['admin', 'admin123', 'admin']);
        sheet.appendRow(['user', 'user123', 'user']);
      }
      if (name === 'Categories') {
        ['IM', 'IK', 'Form', 'Memo'].forEach(c => sheet.appendRow([c]));
      }
      if (name === 'DashboardCategories') {
        ['Google Sheets', 'Maintenance'].forEach((c, index) => sheet.appendRow([c, index + 1]));
      }
    } else if (name === 'Links') {
      let headers = sheet.getRange(1, 1, 1, 6).getValues()[0];
      if (headers[5] !== 'Jenis') {
        sheet.getRange(1, 6).setValue('Jenis');
      }
    }
  }
  
  // Buat folder kategori untuk kategori yang sudah ada
  initializeCategoryFolders();
  
  return "Database berhasil disiapkan!";
}

/**
 * Membuat folder kategori untuk semua kategori yang ada di sheet
 * Jalankan ini jika sudah punya kategori tapi belum ada folder di Drive
 */
function initializeCategoryFolders() {
  try {
    const categories = getCategories();
    let createdCount = 0;
    
    for (const cat of categories) {
      try {
        getOrCreateCategoryFolder(cat);
        createdCount++;
        Logger.log('Folder initialized: ' + cat);
      } catch (e) {
        Logger.log('Error initing folder ' + cat + ': ' + e.toString());
      }
    }
    
    Logger.log('Total folder terinialisasi: ' + createdCount);
    return { success: true, message: 'Total folder terinialisasi: ' + createdCount };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

// ----------------------------------------------------
// FUNGSI AUTHENTICATION
// ----------------------------------------------------
function checkLogin(username, password) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toLowerCase() === username.toLowerCase() && data[i][1].toLowerCase() === password.toLowerCase()) {
      return { success: true, role: data[i][2] };
    }
  }
  return { success: false, message: 'Username atau Password salah!' };
}

// ----------------------------------------------------
// MANAJEMEN LINKS (DASHBOARD)
// ----------------------------------------------------
function getLinks() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Links');
  if (!sheet) return [];
  let data = sheet.getDataRange().getValues();
  let result = [];
  
  for (let i = 1; i < data.length; i++) {
    result.push({
      id: data[i][0],
      kategori: data[i][1],
      judul: data[i][2],
      deskripsi: data[i][3],
      url: data[i][4],
      jenis: data[i][5] || 'Menu'
    });
  }
  return result;
}

function getDashboardCategories() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DashboardCategories');
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const categories = values.map(row => ({ name: row[0], order: row[1] || 999 })).sort((a, b) => a.order - b.order);
  return categories.map(cat => cat.name);
}

function addDashboardCategory(name) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DashboardCategories');
  if (!sheet) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    sheet = ss.insertSheet('DashboardCategories');
    sheet.appendRow(['Kategori Dashboard', 'Order']);
  }

  const lastRow = sheet.getLastRow();
  const data = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 2).getValues() : [];
  if (data.some(row => row[0] === name)) {
    return { success: false, message: 'Kategori sudah ada!' };
  }

  const maxOrder = data.length ? Math.max(...data.map(row => row[1] || 0)) : 0;
  sheet.appendRow([name, maxOrder + 1]);
  return { success: true, message: 'Kategori berhasil ditambahkan!' };
}

function updateDashboardCategory(oldName, newName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DashboardCategories');
  if (!sheet) return { success: false, message: 'Sheet kategori tidak ditemukan.' };

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
  if (values.includes(newName)) {
    return { success: false, message: 'Kategori baru sudah ada.' };
  }

  const data = sheet.getDataRange().getValues();
  let updated = false;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === oldName) {
      sheet.getRange(i + 1, 1).setValue(newName);
      updated = true;
      break;
    }
  }

  if (!updated) return { success: false, message: 'Kategori tidak ditemukan.' };

  // Update kategori di Links sheet
  const linksSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Links');
  if (linksSheet) {
    const linksData = linksSheet.getDataRange().getValues();
    for (let i = 1; i < linksData.length; i++) {
      if (linksData[i][1] === oldName) {
        linksSheet.getRange(i + 1, 2).setValue(newName);
      }
    }
  }

  return { success: true, message: 'Kategori berhasil diperbarui!' };
}

function deleteDashboardCategory(name) {
  try {
    // Cek berapa link yang menggunakan kategori ini
    const linksSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Links');
    if (!linksSheet) {
      return deleteDashboardCategoryConfirmed(name);
    }

    const linksData = linksSheet.getDataRange().getValues();
    let linkCount = 0;

    // Cek dari baris kedua (skip header)
    for (let i = 1; i < linksData.length; i++) {
      if (linksData[i][1] === name) { // Kolom kategori (index 1)
        linkCount++;
      }
    }

    // Jika ada link, return info untuk konfirmasi
    if (linkCount > 0) {
      return {
        success: false,
        needsConfirmation: true,
        message: 'Kategori "' + name + '" berisi ' + linkCount + ' link dashboard.',
        linkCount: linkCount,
        categoryName: name
      };
    }

    // Jika tidak ada link, hapus langsung
    return deleteDashboardCategoryConfirmed(name);
  } catch (e) {
    return { success: false, message: 'Gagal menghapus kategori: ' + e.toString() };
  }
}

function deleteDashboardCategoryConfirmed(name) {
  try {
    // Hapus kategori dari sheet DashboardCategories
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DashboardCategories');
    if (!sheet) return { success: false, message: 'Sheet kategori tidak ditemukan.' };

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === name) {
        sheet.deleteRow(i + 1);

        // Hapus semua link di kategori ini
        const linksSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Links');
        if (linksSheet) {
          const linksData = linksSheet.getDataRange().getValues();

          // Cari dan hapus dari bawah ke atas untuk menghindari index shifting
          for (let j = linksData.length - 1; j >= 1; j--) {
            if (linksData[j][1] === name) { // Kolom kategori (index 1)
              linksSheet.deleteRow(j + 1);
            }
          }
        }

        return { success: true, message: 'Kategori "' + name + '" dan semua link di dalamnya berhasil dihapus!' };
      }
    }

    return { success: false, message: 'Kategori tidak ditemukan.' };
  } catch (e) {
    return { success: false, message: 'Gagal menghapus kategori: ' + e.toString() };
  }
}

function addLink(kategori, judul, deskripsi, url, jenis) {
  try {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Links');
    if (!sheet) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Links');
      sheet.appendRow(['ID', 'Kategori', 'Judul', 'Deskripsi', 'URL', 'Jenis']);
    }
    
    let id = 'L-' + new Date().getTime();
    sheet.appendRow([id, kategori, judul, deskripsi, url, jenis]);
    
    return { success: true, message: 'Menu berhasil ditambahkan!' };
  } catch (e) {
    return { success: false, message: 'Gagal: ' + e.toString() };
  }
}

function deleteLink(id) {
  try {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Links');
    let data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Menu berhasil dihapus!' };
      }
    }
    return { success: false, message: 'ID tidak ditemukan.' };
  } catch (e) {
    return { success: false, message: 'Gagal hapus: ' + e.toString() };
  }
}

// ----------------------------------------------------
// MANAJEMEN ARSIP & KATEGORI
// ----------------------------------------------------
function getCategories() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Categories');
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  return sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(r => r[0]);
}

/**
 * Membuat atau mengambil folder kategori
 * Jika folder belum ada, akan otomatis dibuat
 */
function getOrCreateCategoryFolder(categoryName) {
  try {
    const mainFolder = DriveApp.getFolderById(MAIN_FOLDER_ID);
    const folders = mainFolder.getFoldersByName(categoryName);
    
    if (folders.hasNext()) {
      return folders.next();
    } else {
      // Jika folder tidak ada, buat folder baru
      const newFolder = mainFolder.createFolder(categoryName);
      Logger.log('Folder kategori dibuat: ' + categoryName);
      return newFolder;
    }
  } catch (e) {
    Logger.log('Error di getOrCreateCategoryFolder: ' + e.toString());
    throw new Error('Gagal mengakses/membuat folder kategori: ' + e.toString());
  }
}

function addCategory(name) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Categories');

    const data = sheet.getRange(2, 1, sheet.getLastRow(), 1)
      .getValues()
      .flat();

    if (data.includes(name)) {
      return { success: false, message: 'Kategori sudah ada!' };
    }

    // Tambahkan kategori ke sheet
    sheet.appendRow([name]);

    // Buat folder kategori di Drive secara otomatis
    try {
      getOrCreateCategoryFolder(name);
    } catch (folderError) {
      Logger.log('Warning: Folder tidak bisa dibuat tapi kategori disimpan: ' + folderError.toString());
    }

    return { success: true, message: 'Kategori berhasil ditambahkan dan folder dibuat!' };
  } catch (e) {
    return { success: false, message: 'Gagal: ' + e.toString() };
  }
}

function deleteCategory(name) {
  try {
    // Cek berapa file yang menggunakan kategori ini
    const archivesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Archives');
    const archivesData = archivesSheet.getDataRange().getValues();

    let fileCount = 0;
    // Cek dari baris kedua (skip header)
    for (let i = 1; i < archivesData.length; i++) {
      if (archivesData[i][3] === name) { // Kolom kategori (index 3)
        fileCount++;
      }
    }

    // Jika ada file, return info untuk konfirmasi
    if (fileCount > 0) {
      return {
        success: false,
        needsConfirmation: true,
        message: 'Kategori "' + name + '" berisi ' + fileCount + ' file arsip.',
        fileCount: fileCount,
        categoryName: name
      };
    }

    // Jika tidak ada file, hapus langsung
    return deleteCategoryConfirmed(name);
  } catch (e) {
    return { success: false, message: 'Gagal menghapus kategori: ' + e.toString() };
  }
}

function deleteCategoryConfirmed(name) {
  try {
    // Hapus semua file di kategori ini dari Drive
    const mainFolder = DriveApp.getFolderById(MAIN_FOLDER_ID);
    const subFolders = mainFolder.getFoldersByName(name);

    if (subFolders.hasNext()) {
      const categoryFolder = subFolders.next();

      // Hapus semua file di folder kategori
      const files = categoryFolder.getFiles();
      let deletedFiles = 0;
      while (files.hasNext()) {
        const file = files.next();
        file.setTrashed(true);
        deletedFiles++;
      }

      // Hapus folder kategori
      categoryFolder.setTrashed(true);
      Logger.log('Folder kategori dan ' + deletedFiles + ' file dihapus: ' + name);
    }

    // Hapus dari sheet Categories
    const categoriesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Categories');
    const categoriesData = categoriesSheet.getDataRange().getValues();

    for (let i = 1; i < categoriesData.length; i++) {
      if (categoriesData[i][0] === name) {
        categoriesSheet.deleteRow(i + 1);
        break;
      }
    }

    // Hapus semua record arsip di kategori ini
    const archivesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Archives');
    const archivesData = archivesSheet.getDataRange().getValues();

    // Cari dan hapus dari bawah ke atas untuk menghindari index shifting
    for (let i = archivesData.length - 1; i >= 1; i--) {
      if (archivesData[i][3] === name) { // Kolom kategori (index 3)
        archivesSheet.deleteRow(i + 1);
      }
    }

    return { success: true, message: 'Kategori "' + name + '" dan semua file di dalamnya berhasil dihapus!' };
  } catch (e) {
    return { success: false, message: 'Gagal menghapus kategori: ' + e.toString() };
  }
}

function updateCategory(oldName, newName) {
  try {
    // Cek apakah nama baru sudah ada
    const categoriesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Categories');
    const categoriesData = categoriesSheet.getDataRange().getValues();

    for (let i = 1; i < categoriesData.length; i++) {
      if (categoriesData[i][0] === newName && categoriesData[i][0] !== oldName) {
        return { success: false, message: 'Kategori "' + newName + '" sudah ada!' };
      }
    }

    // Update nama kategori di sheet Categories
    for (let i = 1; i < categoriesData.length; i++) {
      if (categoriesData[i][0] === oldName) {
        categoriesSheet.getRange(i + 1, 1).setValue(newName);
        break;
      }
    }

    // Update nama kategori di sheet Archives
    const archivesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Archives');
    const archivesData = archivesSheet.getDataRange().getValues();

    for (let i = 1; i < archivesData.length; i++) {
      if (archivesData[i][3] === oldName) { // Kolom kategori (index 3)
        archivesSheet.getRange(i + 1, 4).setValue(newName);
      }
    }

    // Rename folder di Drive jika ada
    try {
      const mainFolder = DriveApp.getFolderById(MAIN_FOLDER_ID);
      const subFolders = mainFolder.getFoldersByName(oldName);

      if (subFolders.hasNext()) {
        const oldFolder = subFolders.next();
        oldFolder.setName(newName);
        Logger.log('Folder kategori di-rename: ' + oldName + ' → ' + newName);
      } else {
        // Jika folder belum ada, buat dengan nama baru
        getOrCreateCategoryFolder(newName);
      }
    } catch (folderError) {
      Logger.log('Warning: Gagal rename folder kategori: ' + folderError.toString());
    }

    return { success: true, message: 'Kategori berhasil diubah dari "' + oldName + '" menjadi "' + newName + '"!' };
  } catch (e) {
    return { success: false, message: 'Gagal mengubah kategori: ' + e.toString() };
  }
}

// FORCE DELETE - Hapus kategori beserta semua isinya
function forceDeleteCategory(name) {
  try {
    Logger.log('Memulai force delete kategori: ' + name);

    // 1. Hapus semua file di kategori ini dari Drive
    const mainFolder = DriveApp.getFolderById(MAIN_FOLDER_ID);
    const subFolders = mainFolder.getFoldersByName(name);

    let deletedFiles = 0;
    if (subFolders.hasNext()) {
      const categoryFolder = subFolders.next();

      // Hapus semua file di folder kategori
      const files = categoryFolder.getFiles();
      while (files.hasNext()) {
        const file = files.next();
        file.setTrashed(true);
        deletedFiles++;
      }

      // Hapus folder kategori
      categoryFolder.setTrashed(true);
      Logger.log('Folder kategori dan ' + deletedFiles + ' file dihapus: ' + name);
    }

    // 2. Hapus dari sheet Categories
    const categoriesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Categories');
    const categoriesData = categoriesSheet.getDataRange().getValues();

    for (let i = 1; i < categoriesData.length; i++) {
      if (categoriesData[i][0] === name) {
        categoriesSheet.deleteRow(i + 1);
        break;
      }
    }

    // 3. Hapus semua record arsip di kategori ini
    const archivesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Archives');
    const archivesData = archivesSheet.getDataRange().getValues();

    let deletedRecords = 0;
    // Cari dan hapus dari bawah ke atas untuk menghindari index shifting
    for (let i = archivesData.length - 1; i >= 1; i--) {
      if (archivesData[i][3] === name) { // Kolom kategori (index 3)
        archivesSheet.deleteRow(i + 1);
        deletedRecords++;
      }
    }

    return {
      success: true,
      message: 'Kategori "' + name + '" berhasil dihapus paksa!\n\nDihapus:\n• ' + deletedFiles + ' file dari Drive\n• ' + deletedRecords + ' record dari database\n• 1 folder kategori'
    };
  } catch (e) {
    Logger.log('Error force delete category: ' + e.toString());
    return { success: false, message: 'Gagal menghapus paksa kategori: ' + e.toString() };
  }
}

function forceDeleteDashboardCategory(name) {
  try {
    Logger.log('Memulai force delete kategori dashboard: ' + name);

    // 1. Hapus semua link yang menggunakan kategori ini
    const linksSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Links');
    let deletedLinks = 0;

    if (linksSheet) {
      const linksData = linksSheet.getDataRange().getValues();

      // Cari dan hapus dari bawah ke atas untuk menghindari index shifting
      for (let i = linksData.length - 1; i >= 1; i--) {
        if (linksData[i][1] === name) { // Kolom kategori (index 1)
          linksSheet.deleteRow(i + 1);
          deletedLinks++;
        }
      }
    }

    // 2. Hapus kategori dari sheet DashboardCategories
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DashboardCategories');
    if (!sheet) return { success: false, message: 'Sheet kategori tidak ditemukan.' };

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === name) {
        sheet.deleteRow(i + 1);
        break;
      }
    }

    return {
      success: true,
      message: 'Kategori dashboard "' + name + '" berhasil dihapus paksa!\n\nDihapus:\n• ' + deletedLinks + ' link dashboard\n• 1 kategori dashboard'
    };
  } catch (e) {
    Logger.log('Error force delete dashboard category: ' + e.toString());
    return { success: false, message: 'Gagal menghapus paksa kategori dashboard: ' + e.toString() };
  }
}

function uploadFileToDrive(base64Data, fileName, kategori, deskripsi) {
  try {
    Logger.log('Starting upload for file: ' + fileName + ', category: ' + kategori);
    
    // Parsing base64
    const splitData = base64Data.split(',');
    if (splitData.length < 2) {
      throw new Error('Format base64 tidak valid');
    }
    
    const contentType = splitData[0].match(/:(.*?);/)[1] || 'application/octet-stream';
    const byteCharacters = Utilities.base64Decode(splitData[1]);
    const blob = Utilities.newBlob(byteCharacters, contentType, fileName);
    
    Logger.log('Blob created, size: ' + blob.getBytes().length + ', MIME: ' + contentType);
    
    // Pastikan kategori folder ada, jika tidak buat
    let categoryFolder;
    try {
      categoryFolder = getOrCreateCategoryFolder(kategori);
      Logger.log('Category folder ready: ' + categoryFolder.getName());
    } catch (folderError) {
      Logger.log('Error getting category folder: ' + folderError.toString());
      // Fallback ke main folder jika gagal
      categoryFolder = DriveApp.getFolderById(MAIN_FOLDER_ID);
      Logger.log('Using main folder as fallback');
    }
    
    // Upload ke folder kategori
    const file = categoryFolder.createFile(blob);
    Logger.log('File created in Drive: ' + file.getName() + ', ID: ' + file.getId() + ', MIME: ' + file.getMimeType());
    
    // Set sharing - ini bisa gagal tapi file tetap ada
    let sharingSuccess = true;
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      Logger.log('Sharing set successfully');
    } catch (sharingError) {
      Logger.log('Warning: Failed to set sharing: ' + sharingError.toString());
      sharingSuccess = false;
    }
    
    // Catat di Archives sheet - tambahkan MIME type
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Archives');
    if (!sheet) {
      throw new Error('Sheet Archives tidak ditemukan');
    }
    
    const id = 'ARK-' + new Date().getTime();
    const timestamp = new Date();
    const fileUrl = file.getUrl();
    const fileId = file.getId();
    const mimeType = file.getMimeType();
    
    sheet.appendRow([id, timestamp, fileName, kategori, deskripsi, fileUrl, fileId, mimeType]);
    Logger.log('Record added to Archives sheet: ' + id + ', MIME: ' + mimeType);
    
    const message = sharingSuccess 
      ? 'File berhasil diupload ke folder ' + kategori + '!'
      : 'File berhasil diupload ke folder ' + kategori + ' (perhatian: sharing mungkin perlu diset manual)!';
    
    return { 
      success: true, 
      message: message,
      fileId: fileId,
      fileUrl: fileUrl,
      mimeType: mimeType
    };
  } catch (e) {
    Logger.log('Error upload: ' + e.toString());
    // Jika file sudah terupload tapi gagal catat, coba cari file dan catat
    try {
      const mainFolder = DriveApp.getFolderById(MAIN_FOLDER_ID);
      const files = mainFolder.getFilesByName(fileName);
      if (files.hasNext()) {
        const uploadedFile = files.next();
        Logger.log('Found uploaded file, attempting to record: ' + uploadedFile.getId());
        
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Archives');
        const id = 'ARK-' + new Date().getTime();
        const timestamp = new Date();
        const mimeType = uploadedFile.getMimeType();
        sheet.appendRow([id, timestamp, fileName, kategori, deskripsi, uploadedFile.getUrl(), uploadedFile.getId(), mimeType]);
        
        return { 
          success: true, 
          message: 'File berhasil diupload (record diperbaiki)!',
          fileId: uploadedFile.getId(),
          fileUrl: uploadedFile.getUrl(),
          mimeType: mimeType
        };
      }
    } catch (recoveryError) {
      Logger.log('Recovery failed: ' + recoveryError.toString());
    }
    
    return { success: false, message: 'Gagal upload: ' + e.toString() };
  }
}

function getArchives() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Archives');
  const data = sheet.getDataRange().getValues();
  const result = [];
  for (let i = data.length - 1; i >= 1; i--) { 
     result.push({
       id: data[i][0],
       timestamp: Utilities.formatDate(new Date(data[i][1]), "GMT+7", "dd/MM/yy HH:mm"),
       fileName: data[i][2],
       kategori: data[i][3],
       keterangan: data[i][4],
       url: data[i][5],
       mimeType: data[i][7] || 'application/octet-stream' // MIME type dari kolom 8
     });
  }
  return result;
}

function deleteArchive(id) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Archives');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        try { DriveApp.getFileById(data[i][6]).setTrashed(true); } catch(e) {}
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Arsip dihapus!' };
      }
    }
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}
