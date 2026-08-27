const SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1kJMH5vweXd4YU8j6urk2pk6iuGcLaUeSCzCWT6szmLM/edit';

function doGet(e) {
  initSheets();
  
  const action = e.parameter.action;
  let result = {};

  try {
    if (action === 'getMasterData') {
      result = getMasterData();
    } else if (action === 'getDashboardData') {
      result = getDashboardData();
    } else if (action === 'getOdcHistory') {
      result = getOdcHistory(e.parameter.odc);
    } else if (action === 'getTechnicianHistory') {
      result = getTechnicianHistory(e.parameter.query);
    } else {
      result = { error: 'Invalid action' };
    }
  } catch (err) {
    result = { error: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let result = {};
  
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    if (action === 'submitBorrow') {
      result = submitBorrow(data.payload);
    } else if (action === 'submitReturn') {
      result = submitReturn(data.payload);
    } else if (action === 'submitEvidence') {
      result = submitEvidence(data.payload);
    } else {
      result = { success: false, message: 'Invalid action' };
    }
  } catch (err) {
    result = { success: false, message: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function initSheets() {
  const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  if (!ss.getSheetByName('ActiveBorrowings')) {
    const sheet = ss.insertSheet('ActiveBorrowings');
    sheet.appendRow(['ID', 'STO', 'ODC', 'User', 'Kegiatan', 'Estimasi', 'Waktu Pinjam', 'Selfie Pinjam URL']);
  }
  if (!ss.getSheetByName('History')) {
    const sheet = ss.insertSheet('History');
    sheet.appendRow(['ID', 'STO', 'ODC', 'User', 'Kegiatan', 'Estimasi', 'Waktu Pinjam', 'Selfie Pinjam URL', 'Waktu Kembali', 'Selfie Kembali URL']);
  }
  if (!ss.getSheetByName('EvidenceKegiatan')) {
    const sheet = ss.insertSheet('EvidenceKegiatan');
    sheet.appendRow(['ID', 'STO', 'ODC', 'Teknisi', 'Kegiatan', 'Waktu', 'Foto Evidence URL']);
  }
}

function submitEvidence(data) {
  const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  const sheet = ss.getSheetByName('EvidenceKegiatan');
  
  let urls = [];
  if (data.photos && data.photos.length > 0) {
    for (let i = 0; i < data.photos.length; i++) {
      const fileName = 'EVIDENCE_' + data.odc + '_' + new Date().getTime() + '_' + i + '.png';
      const url = uploadImageToDrive(data.photos[i], fileName);
      if (url) urls.push(url);
    }
  }
  
  const photoUrls = urls.join(', \n');
  const id = generateTicketId(data.sto, true);
  const time = new Date().toLocaleString();
  
  sheet.appendRow([id, data.sto, data.odc, data.user, data.kegiatan, time, photoUrls, data.extend || '-', data.reason || '-']);
  return { success: true, id: id };
}

function generateTicketId(sto, isEvidence) {
  const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  let maxCount = 0;
  
  if (!isEvidence) {
    const activeSheet = ss.getSheetByName('ActiveBorrowings');
    const historySheet = ss.getSheetByName('History');
    
    if (activeSheet) {
      const activeData = activeSheet.getDataRange().getValues();
      for (let i = 1; i < activeData.length; i++) {
        if (activeData[i][1] === sto) {
          const idStr = activeData[i][0].toString();
          const num = parseInt(idStr.replace(sto, ''), 10);
          if (!isNaN(num) && num > maxCount) maxCount = num;
        }
      }
    }
    
    if (historySheet) {
      const historyData = historySheet.getDataRange().getValues();
      for (let i = 1; i < historyData.length; i++) {
        if (historyData[i][1] === sto) {
          const idStr = historyData[i][0].toString();
          const num = parseInt(idStr.replace(sto, ''), 10);
          if (!isNaN(num) && num > maxCount) maxCount = num;
        }
      }
    }
    
    return sto + (maxCount + 1);
  } else {
    const evSheet = ss.getSheetByName('EvidenceKegiatan');
    if (evSheet) {
      const evData = evSheet.getDataRange().getValues();
      for (let i = 1; i < evData.length; i++) {
        if (evData[i][1] === sto) {
          const idStr = evData[i][0].toString();
          const num = parseInt(idStr.replace('EVD-' + sto, ''), 10);
          if (!isNaN(num) && num > maxCount) maxCount = num;
        }
      }
    }
    return 'EVD-' + sto + (maxCount + 1);
  }
}

function getMasterData() {
  const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  const sheet = ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();
  
  const stas = {};
  for (let i = 1; i < data.length; i++) {
    const sto = data[i][5];
    const odc = data[i][7];
    
    if (sto && odc) {
      if (!stas[sto]) stas[sto] = [];
      if (!stas[sto].includes(odc)) stas[sto].push(odc);
    }
  }
  return stas;
}

function getActiveBorrowings() {
  const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  const sheet = ss.getSheetByName('ActiveBorrowings');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  
  const active = [];
  for (let i = 1; i < data.length; i++) {
    active.push({
      id: data[i][0],
      sto: data[i][1],
      odc: data[i][2],
      user: data[i][3],
      kegiatan: data[i][4],
      estimasi: data[i][5],
      waktuPinjam: data[i][6],
      selfiePinjam: data[i][7],
      dasarKegiatan: data[i][8] || '',
      dasarEvidence: data[i][9] || ''
    });
  }
  return active;
}

function getDashboardData() {
  const stas = getMasterData();
  const active = getActiveBorrowings();
  
  const dashboard = {};
  const activeMap = {};
  active.forEach(b => activeMap[b.odc] = b);
  
  for (const sto in stas) {
    dashboard[sto] = { name: sto, status: 'green', odcs: [] };
    
    stas[sto].forEach(odc => {
      const isBorrowed = !!activeMap[odc];
      if (isBorrowed) dashboard[sto].status = 'red';
      
      dashboard[sto].odcs.push({
        name: odc,
        status: isBorrowed ? 'red' : 'green',
        borrowDetails: isBorrowed ? activeMap[odc] : null
      });
    });
  }
  return dashboard;
}

function uploadImageToDrive(base64Str, fileName) {
  try {
    const splitBase = base64Str.split(',');
    const type = splitBase[0].split(';')[0].replace('data:', '');
    const byteCharacters = Utilities.base64Decode(splitBase[1]);
    const blob = Utilities.newBlob(byteCharacters, type, fileName);
    
    let folder;
    const folders = DriveApp.getFoldersByName('ODC_Selfie_Evidence');
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder('ODC_Selfie_Evidence');
    }
    
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) {
    return 'Error saving image: ' + e.message;
  }
}

function submitBorrow(data) {
  const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  const sheet = ss.getSheetByName('ActiveBorrowings');
  
  let urls = [];
  if (data.selfie && data.selfie.length > 0) {
    for (let i = 0; i < data.selfie.length; i++) {
      const fileName = 'PINJAM_' + data.odc + '_' + new Date().getTime() + '_' + i + '.png';
      const url = uploadImageToDrive(data.selfie[i], fileName);
      if (url) urls.push(url);
    }
  }
  const selfieUrls = urls.join(', \n');
  
  let dasarUrl = '';
  if (data.dasarEvidence) {
    const dasarFileName = 'DASAR_' + data.odc + '_' + new Date().getTime() + '.png';
    dasarUrl = uploadImageToDrive(data.dasarEvidence, dasarFileName);
  }

  const id = generateTicketId(data.sto, false);
  const time = new Date().toLocaleString();
  
  sheet.appendRow([id, data.sto, data.odc, data.user, data.kegiatan, data.estimasi, time, selfieUrls, data.dasarKegiatan || '', dasarUrl]);
  return { success: true, id: id };
}

function submitReturn(data) {
  const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  const activeSheet = ss.getSheetByName('ActiveBorrowings');
  const historySheet = ss.getSheetByName('History');
  
  const activeData = activeSheet.getDataRange().getValues();
  let rowIndex = -1;
  let record = null;
  
  for (let i = 1; i < activeData.length; i++) {
    if (activeData[i][2] && data.odc && activeData[i][2].toString().trim() === data.odc.toString().trim()) {
      rowIndex = i + 1;
      record = activeData[i];
      break;
    }
  }
  
  if (rowIndex === -1) return { success: false, message: 'Kunci ODC ini belum dipinjam / tidak ada di daftar aktif' };
  
  const fileName = 'KEMBALI_' + data.odc + '_' + new Date().getTime() + '.png';
  const selfieUrl = uploadImageToDrive(data.selfie, fileName);
  const time = new Date().toLocaleString();
  
  historySheet.appendRow([record[0], record[1], record[2], record[3], record[4], record[5], record[6], record[7], time, selfieUrl, record[8] || '', record[9] || '']);
  activeSheet.deleteRow(rowIndex);
  
  return { success: true };
}

function getOdcHistory(odc) {
  const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  const sheet = ss.getSheetByName('History');
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const history = [];
  
  for (let i = data.length - 1; i > 0; i--) {
    if (data[i][2] === odc) {
      history.push({
        id: data[i][0],
        user: data[i][3],
        kegiatan: data[i][4],
        waktuPinjam: data[i][6],
        waktuKembali: data[i][8],
        selfiePinjam: data[i][7],
        selfieKembali: data[i][9],
        dasarKegiatan: data[i][10] || '',
        dasarEvidence: data[i][11] || ''
      });
    }
  }
  return history;
}

function getTechnicianHistory(query) {
  if (!query) return { active: [], history: [] };
  query = query.toString().toLowerCase().trim();
  const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  
  const activeSheet = ss.getSheetByName('ActiveBorrowings');
  let activeList = [];
  if (activeSheet) {
    const activeData = activeSheet.getDataRange().getValues();
    for (let i = 1; i < activeData.length; i++) {
      const id = (activeData[i][0] || '').toString().toLowerCase();
      const user = (activeData[i][3] || '').toString().toLowerCase();
      if (id.includes(query) || user.includes(query)) {
        activeList.push({
          id: activeData[i][0],
          sto: activeData[i][1],
          odc: activeData[i][2],
          user: activeData[i][3],
          kegiatan: activeData[i][4],
          waktuPinjam: activeData[i][6],
          status: 'Sedang Dipinjam'
        });
      }
    }
  }

  const historySheet = ss.getSheetByName('History');
  let historyList = [];
  if (historySheet) {
    const historyData = historySheet.getDataRange().getValues();
    for (let i = historyData.length - 1; i > 0; i--) {
      const id = (historyData[i][0] || '').toString().toLowerCase();
      const user = (historyData[i][3] || '').toString().toLowerCase();
      if (id.includes(query) || user.includes(query)) {
        historyList.push({
          id: historyData[i][0],
          sto: historyData[i][1],
          odc: historyData[i][2],
          user: historyData[i][3],
          kegiatan: historyData[i][4],
          waktuPinjam: historyData[i][6],
          waktuKembali: historyData[i][8],
          status: 'Sudah Kembali'
        });
      }
      if (historyList.length >= 20) break; // Limit to 20 recent records
    }
  }
  
  return { active: activeList, history: historyList };
}
