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
  return "Database berhasil disiapkan!";
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
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DashboardCategories');
  if (!sheet) return { success: false, message: 'Sheet kategori tidak ditemukan.' };

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === name) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'Kategori berhasil dihapus!' };
    }
  }

  return { success: false, message: 'Kategori tidak ditemukan.' };
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
       timestamp: Utilities.formatDate(new Date(data[i][1]), "GMT+7", "dd/MM/yy HH:mm"),
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
