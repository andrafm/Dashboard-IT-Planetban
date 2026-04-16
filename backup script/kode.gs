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
// SETUP DATABASE (Jalankan ini jika sheet belum ada)
// ----------------------------------------------------
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const sheets = {
    'Users': ['Username', 'Password', 'Role'],
    'Categories': ['Nama Kategori'],
    'Archives': ['ID', 'Timestamp', 'File Name', 'Category', 'Description', 'URL', 'FileID'],
    'Links': ['ID', 'Kategori', 'Judul', 'Deskripsi', 'URL']
  };

  for (let name in sheets) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(sheets[name]);
      
      // Data Awal untuk Users
      if (name === 'Users') {
        sheet.appendRow(['admin', 'admin123', 'admin']);
        sheet.appendRow(['user', 'user123', 'user']);
      }
      // Data Awal untuk Categories
      if (name === 'Categories') {
        ['IM', 'IK', 'Form', 'Memo'].forEach(c => sheet.appendRow([c]));
      }
      // Setup Sheet Links jika belum ada
      let sheetLinks = ss.getSheetByName('Links');
      if (!sheetLinks) {
      sheetLinks = ss.insertSheet('Links');
      sheetLinks.appendRow(['ID', 'Kategori Group', 'Judul Menu', 'Deskripsi', 'URL', 'Jenis']);
      } else {
      // Pastikan header kolom ke-6 adalah 'Jenis'
      let headers = sheetLinks.getRange(1, 1, 1, 6).getValues()[0];
      if (headers[5] !== 'Jenis') {
      sheetLinks.getRange(1, 6).setValue('Jenis');
    }
  }
    }
  }
  return "Database berhasil disiapkan!";
}

// ----------------------------------------------------
// FUNGSI AUTHENTICATION
// ----------------------------------------------------
function checkLogin(username, password) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username && data[i][1] === password) {
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
      jenis: data[i][5] || 'Menu' // Default ke 'Menu' jika kosong
    });
  }
  return result;
}

function addLink(kategori, judul, deskripsi, url, jenis) {
  try {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Links');
    if (!sheet) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Links');
      sheet.appendRow(['ID', 'Kategori Group', 'Judul Menu', 'Deskripsi', 'URL', 'Jenis']);
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
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().map(r => r[0]);
}

function addCategory(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Categories');

  const data = sheet.getRange(2, 1, sheet.getLastRow(), 1)
    .getValues()
    .flat();

  if (data.includes(name)) {
    return { success: false, message: 'Kategori sudah ada!' };
  }

  sheet.appendRow([name]);

  return { success: true, message: 'Kategori berhasil ditambahkan!' };
}

function uploadFileToDrive(base64Data, fileName, kategori, deskripsi) {
  try {
    const splitData = base64Data.split(',');
    const contentType = splitData[0].match(/:(.*?);/)[1];
    const byteCharacters = Utilities.base64Decode(splitData[1]);
    const blob = Utilities.newBlob(byteCharacters, contentType, fileName);
    
    const folder = DriveApp.getFolderById(MAIN_FOLDER_ID);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Archives');
    const id = 'ARK-' + new Date().getTime();
    sheet.appendRow([id, new Date(), fileName, kategori, deskripsi, file.getUrl(), file.getId()]);
    
    return { success: true, message: 'File berhasil diupload!' };
  } catch (e) {
    return { success: false, message: 'Gagal: ' + e.toString() };
  }
}

function getArchives() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Archives');
  const data = sheet.getDataRange().getValues();
  const result = [];
  for (let i = data.length - 1; i >= 1; i--) { 
     result.push({
       id: data[i][0],
       timestamp: Utilities.formatDate(new Date(data[i][1]), "GMT+7", "dd/MM/yyyy HH:mm"),
       fileName: data[i][2],
       kategori: data[i][3],
       keterangan: data[i][4],
       url: data[i][5]
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
