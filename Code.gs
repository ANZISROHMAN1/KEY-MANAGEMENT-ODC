const SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1kJMH5vweXd4YU8j6urk2pk6iuGcLaUeSCzCWT6szmLM/edit';

// ==========================================
// PENGATURAN TELEGRAM BOT
// ==========================================
// Ganti dengan Token Bot dari @BotFather
const TELEGRAM_BOT_TOKEN = '8944833527:AAEnIP7ZCii1Q2GF-RKC1s_sKtIX9DlTE1A'; 
// Ganti dengan Chat ID Grup atau Pribadi Anda (bisa didapatkan dari @userinfobot atau sejenisnya)
const TELEGRAM_CHAT_ID = '-5572220998'; 

function sendTelegramMessage(text) {
  if (TELEGRAM_BOT_TOKEN === 'ISI_BOT_TOKEN_ANDA_DISINI' || TELEGRAM_CHAT_ID === 'ISI_CHAT_ID_ANDA_DISINI') {
    return; // Abaikan jika token/chat id belum diisi
  }
  
  const url = `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`;
  const payload = {
    chat_id: TELEGRAM_CHAT_ID,
    text: text,
    parse_mode: 'HTML'
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    UrlFetchApp.fetch(url, options);
  } catch (e) {
    console.error('Error sending Telegram message: ' + e.message);
  }
}

function doGet(e) {
  initSheets();
  
  const action = e.parameter.action;
  let result = {};

  try {
    if (action === 'getInitialData') {
      result = {
        dashboard: getDashboardData(),
        users: getUsersList(),
        oltData: getOltData(),
        tifVerifications: getTifVerificationsData()
      };
    } else if (action === 'getTifVerifications') {
      result = getTifVerificationsData();
    } else if (action === 'getMasterData') {
      result = getMasterData();
    } else if (action === 'getDashboardData') {
      result = getDashboardData();
    } else if (action === 'getOltData') {
      result = getOltData();
    } else if (action === 'getUsersList') {
      result = getUsersList();
    } else if (action === 'getOdcHistory') {
      result = getOdcHistory(e.parameter.odc);
    } else if (action === 'getAllHistory') {
      result = getAllHistory();
    } else if (action === 'getTechnicianHistory') {
      result = getTechnicianHistory(e.parameter.query);
    } else if (action === 'getPendingApprovals') {
      result = getPendingApprovals(e.parameter.status);
    } else if (action === 'getApprovedBa') {
      result = getApprovedBa();
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
    } else if (action === 'submitPergantianGembok') {
      result = submitPergantianGembok(data.payload);
    } else if (action === 'login') {
      result = handleLogin(data.payload);
    } else if (action === 'register') {
      result = handleRegister(data.payload);
    } else if (action === 'loginHsa') {
      result = loginHsa(data.payload);
    } else if (action === 'loginTif') {
      result = loginTif(data.payload);
    } else if (action === 'registerAdmin') {
      result = registerAdmin(data.payload);
    } else if (action === 'actionApproval') {
      result = actionApproval(data.payload);
    } else if (action === 'updateOlt') {
      result = updateOlt(data.payload);
    } else if (action === 'verifyEvidenceTif') {
      result = verifyEvidenceTif(data.payload);
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
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('SHEETS_INITIALIZED') === 'true') return;

  const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  if (!ss.getSheetByName('ActiveBorrowings')) {
    const sheet = ss.insertSheet('ActiveBorrowings');
    sheet.appendRow(['ID', 'STO', 'ODC', 'User', 'Kegiatan', 'Estimasi', 'Waktu Pinjam', 'Selfie Pinjam URL', 'Status', 'HSA Approver', 'Dasar Kegiatan URL', 'Dasar Evidence URL', 'TIF Approver']);
  }
  if (!ss.getSheetByName('History')) {
    const sheet = ss.insertSheet('History');
    sheet.appendRow(['ID', 'STO', 'ODC', 'User', 'Kegiatan', 'Estimasi', 'Waktu Pinjam', 'Selfie Pinjam URL', 'Waktu Kembali', 'Selfie Kembali URL']);
  }
  if (!ss.getSheetByName('EvidenceKegiatan')) {
    const sheet = ss.insertSheet('EvidenceKegiatan');
    sheet.appendRow(['ID', 'STO', 'ODC', 'Teknisi', 'Kegiatan', 'Waktu', 'Foto Evidence URL']);
  }
  if (!ss.getSheetByName('HistoryGembok')) {
    const sheet = ss.insertSheet('HistoryGembok');
    sheet.appendRow(['ID Tiket', 'STO', 'ODC', 'Petugas', 'Tanggal', 'Keterangan', 'Foto Sebelum URL', 'Foto Sesudah URL', 'Foto Tampak Samping URL', 'Foto Tampak Dalam URL', 'Foto Sebelum Overpal URL', 'Foto Sesudah Overpal URL']);
  }
  if (!ss.getSheetByName('Hsa')) {
    const sheet = ss.insertSheet('Hsa');
    sheet.appendRow(['Username', 'Password', 'STO List (Comma Separated)', 'Telegram ID']);
    sheet.appendRow(['admin_hsa', '123456', 'BOO, CBD', '@contoh_hsa']); // Default dummy HSA
  }
  if (!ss.getSheetByName('Tif')) {
    const sheet = ss.insertSheet('Tif');
    sheet.appendRow(['Username', 'Password', 'Telegram ID']);
    sheet.appendRow(['admin_tif', '123456', '@contoh_tif']); // Default dummy TIF
  }
  if (!ss.getSheetByName('DB OLT')) {
    const sheet = ss.insertSheet('DB OLT');
    sheet.appendRow(['No', 'Hostname', 'Lokasi', 'Keterangan', 'Foto Perangkat']);
  }
  if (!ss.getSheetByName('VerifikasiEvidenceTIF')) {
    const sheet = ss.insertSheet('VerifikasiEvidenceTIF');
    sheet.appendRow(['ODC', 'Status', 'Catatan', 'Reviewer', 'Waktu']);
  }
  
  props.setProperty('SHEETS_INITIALIZED', 'true');
}

function getAdminTelegramTags(sto) {
  const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  let hsaTag = '';
  let tifTag = '';

  const hsaSheet = ss.getSheetByName('Hsa');
  if (hsaSheet) {
    const data = hsaSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const stos = data[i][2] ? data[i][2].toString().toUpperCase() : '';
      if (stos.split(',').map(s => s.trim()).includes(sto.toUpperCase())) {
        if (data[i][3]) {
           hsaTag += data[i][3].toString().trim() + ' ';
        }
      }
    }
  }

  const tifSheet = ss.getSheetByName('Tif');
  if (tifSheet) {
    const data = tifSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][2]) {
        tifTag += data[i][2].toString().trim() + ' ';
      }
    }
  }

  return { hsa: hsaTag.trim(), tif: tifTag.trim() };
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
  const id = generateTicketId(data.sto, data.odc, true);
  const time = new Date().toLocaleString();
  
  sheet.appendRow([id, data.sto, data.odc, data.user, data.kegiatan, time, photoUrls, data.extend || '-', data.reason || '-']);
  
  // Update estimasi di ActiveBorrowings jika ada extend
  if (data.extend && parseInt(data.extend) > 0) {
    const activeSheet = ss.getSheetByName('ActiveBorrowings');
    if (activeSheet) {
      const activeData = activeSheet.getDataRange().getValues();
      for (let i = 1; i < activeData.length; i++) {
        if (activeData[i][2] && activeData[i][2].toString().trim() === data.odc.toString().trim()) {
          const currentEstimasi = parseInt(activeData[i][5]) || 0;
          const newEstimasi = currentEstimasi + parseInt(data.extend);
          activeSheet.getRange(i + 1, 6).setValue(newEstimasi); // Kolom 6 adalah 'Estimasi'
          break;
        }
      }
    }
  }
  
  return { success: true, id: id };
}

function submitPergantianGembok(data) {
  const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  const sheet = ss.getSheetByName('DB ODC') || ss.getSheets()[0];
  const sheetData = sheet.getDataRange().getValues();
  
  let rowIndex = -1;
  const targetOdc = (data.odc || '').toString().toUpperCase().trim();
  const targetSto = (data.sto || '').toString().toUpperCase().trim();
  
  for (let i = 1; i < sheetData.length; i++) {
    const sto = sheetData[i][5] ? sheetData[i][5].toString().toUpperCase().trim() : '';
    const odc = sheetData[i][7] ? sheetData[i][7].toString().toUpperCase().trim() : '';
    
    if (odc === targetOdc) {
      if (!targetSto || sto === targetSto) {
        rowIndex = i + 1; // 1-based row index in spreadsheet
        break;
      }
    }
  }
  
  if (rowIndex === -1) {
    return { success: false, message: 'ODC ' + data.odc + ' tidak ditemukan di database.' };
  }
  
  // Upload foto ke Google Drive jika ada Base64
  let fotoSebelumUrl = '';
  if (data.photoBefore) {
    const fileName = 'SEBELUM_' + targetOdc + '_' + new Date().getTime() + '.png';
    fotoSebelumUrl = uploadImageToDrive(data.photoBefore, fileName);
  } else if (data.photoBeforeUrl) {
    fotoSebelumUrl = data.photoBeforeUrl;
  }

  let fotoSesudahUrl = '';
  if (data.photoAfter) {
    const fileName = 'SESUDAH_' + targetOdc + '_' + new Date().getTime() + '.png';
    fotoSesudahUrl = uploadImageToDrive(data.photoAfter, fileName);
  } else if (data.photo) {
    const fileName = 'SESUDAH_' + targetOdc + '_' + new Date().getTime() + '.png';
    fotoSesudahUrl = uploadImageToDrive(data.photo, fileName);
  } else if (data.photoAfterUrl) {
    fotoSesudahUrl = data.photoAfterUrl;
  }

  let fotoSampingUrl = '';
  if (data.photoSide) {
    const fileName = 'SAMPING_' + targetOdc + '_' + new Date().getTime() + '.png';
    fotoSampingUrl = uploadImageToDrive(data.photoSide, fileName);
  } else if (data.photoSideUrl) {
    fotoSampingUrl = data.photoSideUrl;
  }

  let fotoDalamUrl = '';
  if (data.photoInside) {
    const fileName = 'DALAM_' + targetOdc + '_' + new Date().getTime() + '.png';
    fotoDalamUrl = uploadImageToDrive(data.photoInside, fileName);
  } else if (data.photoInsideUrl) {
    fotoDalamUrl = data.photoInsideUrl;
  }

  let fotoOverpalSebelumUrl = '';
  if (data.photoOverpalBefore) {
    const fileName = 'OVERPAL_SEBELUM_' + targetOdc + '_' + new Date().getTime() + '.png';
    fotoOverpalSebelumUrl = uploadImageToDrive(data.photoOverpalBefore, fileName);
  } else if (data.photoOverpalBeforeUrl) {
    fotoOverpalSebelumUrl = data.photoOverpalBeforeUrl;
  }

  let fotoOverpalSesudahUrl = '';
  if (data.photoOverpalAfter) {
    const fileName = 'OVERPAL_SESUDAH_' + targetOdc + '_' + new Date().getTime() + '.png';
    fotoOverpalSesudahUrl = uploadImageToDrive(data.photoOverpalAfter, fileName);
  } else if (data.photoOverpalAfterUrl) {
    fotoOverpalSesudahUrl = data.photoOverpalAfterUrl;
  }
  
  const tanggal = data.tanggal ? data.tanggal : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
  const petugas = data.petugas || data.user || '-';
  const keterangan = data.keterangan || 'Gembok Baru Terpasang';
  
  // Pastikan header di baris 1 tersedia (mulai dari kolom R / 18 agar aman dari data A-Q)
  if (sheetData.length > 0) {
    if (!sheetData[0][17]) sheet.getRange(1, 18).setValue('GEMBOK DIGANTI');
    if (!sheetData[0][18]) sheet.getRange(1, 19).setValue('FOTO SEBELUM DIGANTI');
    if (!sheetData[0][19]) sheet.getRange(1, 20).setValue('FOTO SESUDAH DIGANTI');
    if (!sheetData[0][20]) sheet.getRange(1, 21).setValue('FOTO TAMPAK SAMPING');
    if (!sheetData[0][21]) sheet.getRange(1, 22).setValue('FOTO TAMPAK DALAM');
    if (!sheetData[0][22]) sheet.getRange(1, 23).setValue('PETUGAS GEMBOK');
    if (!sheetData[0][23]) sheet.getRange(1, 24).setValue('KETERANGAN GEMBOK');
    if (!sheetData[0][24]) sheet.getRange(1, 25).setValue('FOTO SEBELUM OVERPAL');
    if (!sheetData[0][25]) sheet.getRange(1, 26).setValue('FOTO SESUDAH OVERPAL');
  }

  // Kolom R (18): Tanggal Gembok Diganti
  sheet.getRange(rowIndex, 18).setValue(tanggal);
  
  // Kolom S (19): Foto Sebelum URL
  if (fotoSebelumUrl) {
    sheet.getRange(rowIndex, 19).setValue(fotoSebelumUrl);
  }

  // Kolom T (20): Foto Sesudah URL
  if (fotoSesudahUrl) {
    sheet.getRange(rowIndex, 20).setValue(fotoSesudahUrl);
  }
  
  // Kolom U (21): Foto Tampak Samping URL
  if (fotoSampingUrl) {
    sheet.getRange(rowIndex, 21).setValue(fotoSampingUrl);
  }

  // Kolom V (22): Foto Tampak Dalam URL
  if (fotoDalamUrl) {
    sheet.getRange(rowIndex, 22).setValue(fotoDalamUrl);
  }

  // Kolom W (23): Petugas / Teknisi
  sheet.getRange(rowIndex, 23).setValue(petugas);

  // Kolom X (24): Keterangan
  if (keterangan) {
    sheet.getRange(rowIndex, 24).setValue(keterangan);
  }

  // Kolom Y (25): Foto Sebelum Overpal URL
  if (fotoOverpalSebelumUrl) {
    sheet.getRange(rowIndex, 25).setValue(fotoOverpalSebelumUrl);
  }

  // Kolom Z (26): Foto Sesudah Overpal URL
  if (fotoOverpalSesudahUrl) {
    sheet.getRange(rowIndex, 26).setValue(fotoOverpalSesudahUrl);
  }

  // Update status gembok (Ada Gembok: YA, Kondisi: BAGUS, Secure: YA jika dipilih)
  sheet.getRange(rowIndex, 10).setValue('YA'); // Kolom J: Ada Gembok
  sheet.getRange(rowIndex, 12).setValue('BAGUS'); // Kolom L: Kondisi Gembok
  if (data.isSecure) {
    sheet.getRange(rowIndex, 13).setValue('YA'); // Kolom M: Gembok Secure
  }

  // Catat ke sheet HistoryGembok untuk riwayat lengkap
  let histSheet = ss.getSheetByName('HistoryGembok');
  if (!histSheet) {
    histSheet = ss.insertSheet('HistoryGembok');
    histSheet.appendRow(['ID Tiket', 'STO', 'ODC', 'Petugas', 'Tanggal', 'Keterangan', 'Foto Sebelum URL', 'Foto Sesudah URL', 'Foto Tampak Samping URL', 'Foto Tampak Dalam URL', 'Foto Sebelum Overpal URL', 'Foto Sesudah Overpal URL']);
  } else {
    const histHeaders = histSheet.getRange(1, 1, 1, Math.max(histSheet.getLastColumn(), 12)).getValues()[0];
    if (!histHeaders[8]) histSheet.getRange(1, 9).setValue('Foto Tampak Samping URL');
    if (!histHeaders[9]) histSheet.getRange(1, 10).setValue('Foto Tampak Dalam URL');
    if (!histHeaders[10]) histSheet.getRange(1, 11).setValue('Foto Sebelum Overpal URL');
    if (!histHeaders[11]) histSheet.getRange(1, 12).setValue('Foto Sesudah Overpal URL');
  }
  const ticketId = generateTicketId(targetSto || 'STO', targetOdc, true).replace('#EVD-', '#GM-');
  histSheet.appendRow([ticketId, targetSto || '-', targetOdc, petugas, tanggal, keterangan, fotoSebelumUrl, fotoSesudahUrl, fotoSampingUrl, fotoDalamUrl, fotoOverpalSebelumUrl, fotoOverpalSesudahUrl]);

  // Clear cache
  const cache = CacheService.getScriptCache();
  try { cache.remove("MASTER_DATA"); } catch(e){}
  invalidateHistoryCache();

  // Telegram Notification
  const msg = `🔐 <b>UPDATE PERGANTIAN GEMBOK & OVERPAL</b>\n\n` +
              `<b>ODC:</b> ${targetOdc}\n` +
              `<b>Petugas:</b> ${petugas}\n` +
              `<b>Tanggal:</b> ${tanggal}\n` +
              `<b>Keterangan:</b> ${keterangan}\n` +
              (fotoSebelumUrl ? `<b>Evidence Sebelum:</b> <a href="${fotoSebelumUrl}">Lihat Foto</a>\n` : '') +
              (fotoSesudahUrl ? `<b>Evidence Sesudah:</b> <a href="${fotoSesudahUrl}">Lihat Foto</a>\n` : '') +
              (fotoSampingUrl ? `<b>Tampak Samping:</b> <a href="${fotoSampingUrl}">Lihat Foto</a>\n` : '') +
              (fotoDalamUrl ? `<b>Tampak Dalam:</b> <a href="${fotoDalamUrl}">Lihat Foto</a>\n` : '') +
              (fotoOverpalSebelumUrl ? `<b>Sebelum Ganti Overpal:</b> <a href="${fotoOverpalSebelumUrl}">Lihat Foto</a>\n` : '') +
              (fotoOverpalSesudahUrl ? `<b>Sesudah Ganti Overpal:</b> <a href="${fotoOverpalSesudahUrl}">Lihat Foto</a>\n` : '');
  sendTelegramMessage(msg);

  return { 
    success: true, 
    message: 'Pergantian gembok berhasil dicatat!', 
    fotoSebelumUrl: fotoSebelumUrl,
    fotoSesudahUrl: fotoSesudahUrl,
    fotoSampingUrl: fotoSampingUrl,
    fotoDalamUrl: fotoDalamUrl,
    fotoOverpalSebelumUrl: fotoOverpalSebelumUrl,
    fotoOverpalSesudahUrl: fotoOverpalSesudahUrl,
    fotoUrl: fotoSesudahUrl || fotoSebelumUrl || fotoSampingUrl || fotoDalamUrl || fotoOverpalSesudahUrl || fotoOverpalSebelumUrl, 
    tanggal: tanggal,
    petugas: petugas 
  };
}

function generateTicketId(sto, odc, isEvidence) {
  const dateStr = Utilities.formatDate(new Date(), "Asia/Jakarta", "ddMMyy");
  
  // Mapping STO to abbreviation
  const stoMapping = {
    'BOGOR': 'BOO',
    'CIBINONG': 'CBN',
    'CIBADAK': 'CBD',
    'CILEUNGSI': 'CLS',
    'DRAMAGA': 'DRM',
    'CIAWI': 'CAW',
    'SUKABUMI': 'SKI',
    'PAGELARAN': 'PGR',
    'SEMPLAK': 'SPK',
    'GUNUNG PUTRI': 'GNP'
  };
  
  let stoAbbr = stoMapping[sto.toUpperCase()];
  if (!stoAbbr) {
    stoAbbr = sto.substring(0, 3).toUpperCase();
  }
  
  const odcStr = odc ? odc.toUpperCase() : 'ODC';
  
  const prefix = isEvidence ? "#EVD-" : "#";
  return prefix + dateStr + stoAbbr + "-" + odcStr;
}

function getMasterData() {
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get("MASTER_DATA");
  if (cachedData) {
    try {
      return JSON.parse(cachedData);
    } catch(e) {}
  }

  const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  const sheet = ss.getSheetByName('DB ODC') || ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) return {};

  const headers = data[0].map(h => (h || '').toString().toUpperCase().trim());
  const stoIdx = headers.indexOf('STO') !== -1 ? headers.indexOf('STO') : 5;
  const odcIdx = headers.indexOf('ODC NAME') !== -1 ? headers.indexOf('ODC NAME') : 7;
  const saIdx = headers.indexOf('SA') !== -1 ? headers.indexOf('SA') : 8;
  const adaGembokIdx = headers.indexOf('ADA GEMBOK') !== -1 ? headers.indexOf('ADA GEMBOK') : 9;
  const gembokWasakaIdx = headers.indexOf('GEMBOK WASAKA') !== -1 ? headers.indexOf('GEMBOK WASAKA') : 10;
  const kondisiGembokIdx = headers.indexOf('KONDISI GEMBOK') !== -1 ? headers.indexOf('KONDISI GEMBOK') : 11;
  const krangkengIdx = headers.indexOf('KRANGKENG') !== -1 ? headers.indexOf('KRANGKENG') : 12;
  const gembokKrankengIdx = headers.indexOf('GEMBOK KRANKENG') !== -1 ? headers.indexOf('GEMBOK KRANKENG') : 13;
  const tikorIdx = headers.indexOf('TIKOR') !== -1 ? headers.indexOf('TIKOR') : 14;

  // Extra columns in DB ODC (must be >= 17 to avoid standard columns A-Q)
  const gembokDigantiIdx = headers.indexOf('GEMBOK DIGANTI');
  const fotoSebelumIdx = headers.indexOf('FOTO SEBELUM DIGANTI') !== -1 ? headers.indexOf('FOTO SEBELUM DIGANTI') : headers.indexOf('FOTO SEBELUM');
  const fotoSesudahIdx = headers.indexOf('FOTO SESUDAH DIGANTI') !== -1 ? headers.indexOf('FOTO SESUDAH DIGANTI') : headers.indexOf('FOTO SESUDAH');
  const fotoSampingIdx = headers.indexOf('FOTO TAMPAK SAMPING') !== -1 ? headers.indexOf('FOTO TAMPAK SAMPING') : headers.indexOf('TAMPAK SAMPING');
  const fotoDalamIdx = headers.indexOf('FOTO TAMPAK DALAM') !== -1 ? headers.indexOf('FOTO TAMPAK DALAM') : headers.indexOf('TAMPAK DALAM');
  const fotoOverpalSebelumIdx = headers.indexOf('FOTO SEBELUM OVERPAL') !== -1 ? headers.indexOf('FOTO SEBELUM OVERPAL') : headers.indexOf('SEBELUM OVERPAL');
  const fotoOverpalSesudahIdx = headers.indexOf('FOTO SESUDAH OVERPAL') !== -1 ? headers.indexOf('FOTO SESUDAH OVERPAL') : headers.indexOf('SESUDAH OVERPAL');
  const petugasGembokIdx = headers.indexOf('PETUGAS GEMBOK');
  const ketGembokIdx = headers.indexOf('KETERANGAN GEMBOK');

  // Read HistoryGembok to populate latest replaced padlock info
  const gembokHistoryMap = {};
  const histSheet = ss.getSheetByName('HistoryGembok');
  if (histSheet) {
    const histData = histSheet.getDataRange().getValues();
    for (let h = 1; h < histData.length; h++) {
      const odcName = histData[h][2] ? histData[h][2].toString().toUpperCase().trim() : '';
      if (odcName) {
        gembokHistoryMap[odcName] = {
          petugas: histData[h][3] ? histData[h][3].toString().trim() : '',
          tanggal: histData[h][4] ? (histData[h][4] instanceof Date ? Utilities.formatDate(histData[h][4], Session.getScriptTimeZone(), "dd/MM/yyyy") : histData[h][4].toString().trim()) : '',
          keterangan: histData[h][5] ? histData[h][5].toString().trim() : '',
          fotoSebelum: histData[h][6] ? histData[h][6].toString().trim() : '',
          fotoSesudah: histData[h][7] ? histData[h][7].toString().trim() : '',
          fotoSamping: histData[h][8] ? histData[h][8].toString().trim() : '',
          fotoDalam: histData[h][9] ? histData[h][9].toString().trim() : '',
          fotoOverpalSebelum: histData[h][10] ? histData[h][10].toString().trim() : '',
          fotoOverpalSesudah: histData[h][11] ? histData[h][11].toString().trim() : ''
        };
      }
    }
  }

  const sas = {};
  for (let i = 1; i < data.length; i++) {
    const sto = data[i][stoIdx];
    const odc = data[i][odcIdx];
    const sa = data[i][saIdx] || 'OTHER';
    
    const adaGembok = data[i][adaGembokIdx] ? data[i][adaGembokIdx].toString().toUpperCase().trim() : '';
    const gembokWasaka = data[i][gembokWasakaIdx] ? data[i][gembokWasakaIdx].toString().toUpperCase().trim() : '';
    const kondisiGembok = data[i][kondisiGembokIdx] ? data[i][kondisiGembokIdx].toString().toUpperCase().trim() : '';
    const krangkeng = krangkengIdx !== -1 && data[i][krangkengIdx] ? data[i][krangkengIdx].toString().toUpperCase().trim() : '';
    const gembokKrankeng = gembokKrankengIdx !== -1 && data[i][gembokKrankengIdx] ? data[i][gembokKrankengIdx].toString().toUpperCase().trim() : '';
    
    let tikor = '';
    if (tikorIdx !== -1 && data[i][tikorIdx]) {
      const rawT = data[i][tikorIdx].toString().trim();
      if (rawT && rawT !== '#N/A') {
        tikor = rawT;
      }
    }

    const odcKey = (odc || '').toString().toUpperCase().trim();
    const histRecord = gembokHistoryMap[odcKey];

    let gembokDigantiStr = histRecord ? histRecord.tanggal : '';
    let fotoSebelum = histRecord ? histRecord.fotoSebelum : '';
    let fotoSesudah = histRecord ? histRecord.fotoSesudah : '';
    let fotoSamping = histRecord ? histRecord.fotoSamping : '';
    let fotoDalam = histRecord ? histRecord.fotoDalam : '';
    let fotoOverpalSebelum = histRecord ? histRecord.fotoOverpalSebelum : '';
    let fotoOverpalSesudah = histRecord ? histRecord.fotoOverpalSesudah : '';
    let petugasGembok = histRecord ? histRecord.petugas : '';
    let keteranganGembok = histRecord ? histRecord.keterangan : '';

    if (!gembokDigantiStr && gembokDigantiIdx >= 17 && data[i][gembokDigantiIdx]) {
      const rawGembokDiganti = data[i][gembokDigantiIdx];
      if (rawGembokDiganti instanceof Date) {
        gembokDigantiStr = Utilities.formatDate(rawGembokDiganti, Session.getScriptTimeZone(), "dd/MM/yyyy");
      } else {
        gembokDigantiStr = rawGembokDiganti.toString().trim();
      }
    }
    if (!fotoSebelum && fotoSebelumIdx >= 17 && data[i][fotoSebelumIdx]) {
      fotoSebelum = data[i][fotoSebelumIdx].toString().trim();
    }
    if (!fotoSesudah && fotoSesudahIdx >= 17 && data[i][fotoSesudahIdx]) {
      fotoSesudah = data[i][fotoSesudahIdx].toString().trim();
    }
    if (!fotoSamping && fotoSampingIdx >= 17 && data[i][fotoSampingIdx]) {
      fotoSamping = data[i][fotoSampingIdx].toString().trim();
    }
    if (!fotoDalam && fotoDalamIdx >= 17 && data[i][fotoDalamIdx]) {
      fotoDalam = data[i][fotoDalamIdx].toString().trim();
    }
    if (!fotoOverpalSebelum && fotoOverpalSebelumIdx >= 17 && data[i][fotoOverpalSebelumIdx]) {
      fotoOverpalSebelum = data[i][fotoOverpalSebelumIdx].toString().trim();
    }
    if (!fotoOverpalSesudah && fotoOverpalSesudahIdx >= 17 && data[i][fotoOverpalSesudahIdx]) {
      fotoOverpalSesudah = data[i][fotoOverpalSesudahIdx].toString().trim();
    }
    if (!petugasGembok && petugasGembokIdx >= 17 && data[i][petugasGembokIdx]) {
      petugasGembok = data[i][petugasGembokIdx].toString().trim();
    }
    if (!keteranganGembok && ketGembokIdx >= 17 && data[i][ketGembokIdx]) {
      keteranganGembok = data[i][ketGembokIdx].toString().trim();
    }

    const fotoGembok = fotoSesudah || fotoSebelum || fotoSamping || fotoDalam;

    if (sto && odc) {
      if (!sas[sa]) sas[sa] = {};
      if (!sas[sa][sto]) sas[sa][sto] = [];
      const exists = sas[sa][sto].find(item => (typeof item === 'object' ? item.name === odc : item === odc));
      if (!exists) {
        sas[sa][sto].push({
          name: odc,
          tikor: tikor,
          adaGembok: adaGembok,
          gembokWasaka: gembokWasaka,
          kondisiGembok: kondisiGembok,
          krangkeng: krangkeng,
          gembokKrankeng: gembokKrankeng,
          gembokSecure: (gembokWasaka === 'YA' || gembokKrankeng === 'YA') ? 'YA' : '',
          gembokDiganti: gembokDigantiStr,
          fotoSebelum: fotoSebelum,
          fotoSesudah: fotoSesudah,
          fotoSamping: fotoSamping,
          fotoDalam: fotoDalam,
          fotoOverpalSebelum: fotoOverpalSebelum,
          fotoOverpalSesudah: fotoOverpalSesudah,
          fotoGembok: fotoGembok,
          petugasGembok: petugasGembok,
          keteranganGembok: keteranganGembok
        });
      }
    }
  }

  try {
    cache.put("MASTER_DATA", JSON.stringify(sas), 60); // Cache for 1 minute
  } catch(e) {}

  return sas;
}

function getActiveBorrowings() {
  const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  
  // Build NIK Map
  const techSheet = ss.getSheetByName('Teknisi');
  const nikMap = {};
  if (techSheet) {
    const techData = techSheet.getDataRange().getValues();
    for (let i = 1; i < techData.length; i++) {
      if (techData[i][0]) {
        nikMap[techData[i][0].toString().trim()] = techData[i][2] || '-';
      }
    }
  }

  const sheet = ss.getSheetByName('ActiveBorrowings');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  
  const active = [];
  for (let i = 1; i < data.length; i++) {
    const userName = data[i][3] ? data[i][3].toString().trim() : '';
    active.push({
      id: data[i][0],
      sto: data[i][1],
      odc: data[i][2],
      user: userName,
      nik: nikMap[userName] || '-',
      kegiatan: data[i][4],
      estimasi: data[i][5],
      waktuPinjam: data[i][6],
      selfiePinjam: data[i][7],
      status: data[i][8] || '',
      hsaApprover: data[i][9] || '',
      dasarKegiatan: data[i][10] || '',
      dasarEvidence: data[i][11] || '',
      tifApprover: data[i][12] || ''
    });
  }
  return active;
}

function getDashboardData() {
  const sas = getMasterData();
  const active = getActiveBorrowings();
  
  const dashboard = {};
  const activeMap = {};
  active.forEach(b => activeMap[b.odc] = b);
  
  for (const sa in sas) {
    dashboard[sa] = { name: sa, stos: {} };
    
    for (const sto in sas[sa]) {
      dashboard[sa].stos[sto] = { name: sto, status: 'green', odcs: [] };
      
      sas[sa][sto].forEach(item => {
        const odc = typeof item === 'object' ? item.name : item;
        const activeRecord = activeMap[odc];
        const isBorrowed = activeRecord && activeRecord.status === 'Approved';
        
        if (isBorrowed) dashboard[sa].stos[sto].status = 'red';
        if (activeRecord) {
          activeRecord.tikor = item.tikor || '';
        }
        
        dashboard[sa].stos[sto].odcs.push({
          name: odc,
          tikor: item.tikor || '',
          status: isBorrowed ? 'red' : 'green',
          borrowDetails: isBorrowed ? activeRecord : null,
          adaGembok: item.adaGembok || '',
          gembokWasaka: item.gembokWasaka || '',
          kondisiGembok: item.kondisiGembok || '',
          krangkeng: item.krangkeng || '',
          gembokKrankeng: item.gembokKrankeng || '',
          gembokSecure: item.gembokSecure || '',
          gembokDiganti: item.gembokDiganti || '',
          fotoSebelum: item.fotoSebelum || '',
          fotoSesudah: item.fotoSesudah || '',
          fotoSamping: item.fotoSamping || '',
          fotoDalam: item.fotoDalam || '',
          fotoOverpalSebelum: item.fotoOverpalSebelum || '',
          fotoOverpalSesudah: item.fotoOverpalSesudah || '',
          fotoGembok: item.fotoGembok || item.fotoSesudah || item.fotoSebelum || item.fotoSamping || item.fotoDalam || item.fotoOverpalSesudah || item.fotoOverpalSebelum || '',
          petugasGembok: item.petugasGembok || '',
          keteranganGembok: item.keteranganGembok || ''
        });
      });
    }
  }
  
  return dashboard;
}

function resolveLightshotImage(url) {
  if (!url || !url.toString().includes('prnt.sc/')) return url;
  try {
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' } });
    const html = res.getContentText();
    const match = html.match(/<meta\s+property=["']og:image["']\s+content=["']([^"']+)["']/i) ||
                  html.match(/<meta\s+name=["']twitter:image:src["']\s+content=["']([^"']+)["']/i);
    if (match && match[1]) return match[1];
  } catch (e) {}
  return url;
}

function getOltData() {
  const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  
  // Build STO to SA mapping from DB ODC
  const odcSheet = ss.getSheetByName('DB ODC') || ss.getSheets()[0];
  const odcData = odcSheet.getDataRange().getValues();
  const stoToSa = {};
  for (let i = 1; i < odcData.length; i++) {
    const sto = odcData[i][5] ? odcData[i][5].toString().toUpperCase().trim() : '';
    const sa = odcData[i][8] ? odcData[i][8].toString().toUpperCase().trim() : '';
    if (sto && sa && !stoToSa[sto]) {
      stoToSa[sto] = sa;
    }
  }

  const oltSheet = ss.getSheetByName('DB OLT');
  if (!oltSheet) {
    return { olts: [], dashboard: {}, stats: { total: 0, butuhGembok: 0, aman: 0, totalSto: 0, withPhoto: 0 } };
  }
  
  const oltData = oltSheet.getDataRange().getValues();
  const olts = [];
  const dashboard = {};
  let totalButuhGembok = 0;
  let totalAman = 0;
  let totalWithPhoto = 0;
  const stoSet = {};

  for (let i = 1; i < oltData.length; i++) {
    const row = oltData[i];
    const no = row[0];
    const hostname = row[1] ? row[1].toString().trim() : '';
    if (!hostname) continue;
    
    const lokasi = row[2] ? row[2].toString().trim() : '-';
    const keterangan = row[3] ? row[3].toString().trim() : '-';
    const foto = row[4] ? row[4].toString().trim() : '';
    const directFoto = resolveLightshotImage(foto);

    // Extract STO from Hostname: GPON00-D2-JGL-3#AG -> JGL
    let sto = '';
    const parts = hostname.split('-');
    if (parts.length >= 3) {
      sto = parts[2].toUpperCase().trim();
    } else {
      for (const knownSto in stoToSa) {
        if (hostname.toUpperCase().includes(knownSto)) {
          sto = knownSto;
          break;
        }
      }
    }
    if (!sto) sto = 'OTHER';
    
    const sa = stoToSa[sto] || 'AREA LAIN';
    stoSet[sto] = true;

    const isButuhGembok = keterangan.toLowerCase().includes('butuh') || 
                          keterangan.toLowerCase().includes('gembok') || 
                          keterangan.toLowerCase().includes('kerangkeng') ||
                          keterangan.toLowerCase().includes('rusak');
    
    if (isButuhGembok) {
      totalButuhGembok++;
    } else {
      totalAman++;
    }

    if (foto) totalWithPhoto++;

    const oltItem = {
      no: no || i,
      hostname: hostname,
      lokasi: lokasi,
      keterangan: keterangan,
      foto: foto,
      directFoto: directFoto,
      sto: sto,
      sa: sa,
      status: isButuhGembok ? 'warning' : 'green',
      isButuhGembok: isButuhGembok
    };

    olts.push(oltItem);

    // Grouping by SA and STO for Dashboard View
    if (!dashboard[sa]) {
      dashboard[sa] = { name: sa, stos: {} };
    }
    if (!dashboard[sa].stos[sto]) {
      dashboard[sa].stos[sto] = { name: sto, status: 'green', olts: [] };
    }
    if (isButuhGembok) {
      dashboard[sa].stos[sto].status = 'warning';
    }
    dashboard[sa].stos[sto].olts.push(oltItem);
  }

  return {
    olts: olts,
    dashboard: dashboard,
    stats: {
      total: olts.length,
      butuhGembok: totalButuhGembok,
      aman: totalAman,
      totalSto: Object.keys(stoSet).length,
      withPhoto: totalWithPhoto
    }
  };
}

function updateOlt(data) {
  const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  const sheet = ss.getSheetByName('DB OLT');
  if (!sheet) return { success: false, message: 'Sheet DB OLT tidak ditemukan' };
  
  const targetHost = (data.hostname || '').toString().trim().toUpperCase();
  const sheetData = sheet.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < sheetData.length; i++) {
    if (sheetData[i][1] && sheetData[i][1].toString().trim().toUpperCase() === targetHost) {
      rowIndex = i + 1;
      break;
    }
  }
  if (rowIndex === -1) {
    return { success: false, message: 'OLT ' + data.hostname + ' tidak ditemukan' };
  }
  
  let fotoUrl = data.fotoUrl || '';
  if (data.photoBase64) {
    const fileName = 'OLT_' + targetHost.replace(/[^a-zA-Z0-9]/g, '_') + '_' + new Date().getTime() + '.png';
    fotoUrl = uploadImageToDrive(data.photoBase64, fileName);
  }
  
  if (data.lokasi) sheet.getRange(rowIndex, 3).setValue(data.lokasi);
  if (data.keterangan) sheet.getRange(rowIndex, 4).setValue(data.keterangan);
  if (fotoUrl) sheet.getRange(rowIndex, 5).setValue(fotoUrl);
  
  return { success: true, fotoUrl: fotoUrl };
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
  
  const fileName = 'PINJAM_' + data.odc + '_' + new Date().getTime() + '.png';
  const selfieUrl = uploadImageToDrive(data.selfie, fileName);
  
  let dasarUrl = '';
  if (data.dasarEvidence) {
    const dasarFileName = 'DASAR_' + data.odc + '_' + new Date().getTime() + '.png';
    dasarUrl = uploadImageToDrive(data.dasarEvidence, dasarFileName);
  }

  const id = generateTicketId(data.sto, data.odc, false);
  const time = data.waktuPinjam ? new Date(data.waktuPinjam).toLocaleString('id-ID', {timeZone: 'Asia/Jakarta'}) : new Date().toLocaleString('id-ID', {timeZone: 'Asia/Jakarta'});
  
  sheet.appendRow([id, data.sto, data.odc, data.user, data.kegiatan, data.estimasi, time, selfieUrl, 'Pending HSA', '', dasarUrl, '', '']);
  
  // Send Telegram Notification
  const contact = getTechnicianContact(data.user);
  const adminTags = getAdminTelegramTags(data.sto);
  
  let tagText = '';
  if (adminTags.hsa || adminTags.tif) {
    tagText = '\n<b>Tags:</b> ';
    if (adminTags.hsa) tagText += adminTags.hsa + ' (HSA) ';
    if (adminTags.tif) tagText += adminTags.tif + ' (TIF)';
  }

  const msg = `🔔 <b>REQUEST PINJAM KUNCI</b>\n\n` +
              `<b>ID:</b> ${id}\n` +
              `<b>STO:</b> ${data.sto}\n` +
              `<b>ODC:</b> ${data.odc}\n` +
              `<b>Teknisi:</b> ${data.user}\n` +
              `<b>ID Telegram:</b> ${contact.telegram}\n` +
              `<b>No WA:</b> ${contact.wa}\n` +
              `<b>Kegiatan:</b> ${data.kegiatan}\n` +
              `<b>Estimasi Waktu:</b> ${data.estimasi} Jam\n\n` +
              `<i>Status: Menunggu Approval</i>${tagText}`;
  sendTelegramMessage(msg);

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
  
  historySheet.appendRow([record[0], record[1], record[2], record[3], record[4], record[5], record[6], record[7], time, selfieUrl]);
  activeSheet.deleteRow(rowIndex);
  invalidateHistoryCache();
  
  return { success: true };
}

function handleLogin(data) {
  const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  const sheet = ss.getSheetByName('Teknisi');
  
  if (!sheet) {
    return { success: false, message: 'Sheet "Teknisi" tidak ditemukan di Spreadsheet. Harap hubungi Admin.' };
  }
  
  const techData = sheet.getDataRange().getValues();
  for (let i = 1; i < techData.length; i++) {
    const user = techData[i][0] ? techData[i][0].toString().trim() : '';
    const pass = techData[i][1] ? techData[i][1].toString().trim() : '';
    
    if (user === data.username && pass === data.password) {
      let isHsa = false;
      let hsaStos = [];
      const hsaSheet = ss.getSheetByName('Hsa');
      if (hsaSheet) {
        const hsaData = hsaSheet.getDataRange().getValues();
        for (let j = 1; j < hsaData.length; j++) {
          if (hsaData[j][0] && hsaData[j][0].toString() === user) {
            isHsa = true;
            hsaStos = hsaData[j][2] ? hsaData[j][2].toString().split(',').map(s => s.trim()) : [];
            break;
          }
        }
      }
      return { success: true, username: user, isHsa: isHsa, hsaStos: hsaStos };
    }
  }
  
  return { success: false, message: 'Username atau Password salah!' };
}

function handleRegister(data) {
  const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  let sheet = ss.getSheetByName('Teknisi');
  
  if (!sheet) {
    sheet = ss.insertSheet('Teknisi');
    sheet.appendRow(['Username', 'Password', 'NIK', 'STO', 'ID Telegram', 'No WA', 'Foto URL', 'Waktu Register']);
  }
  
  const username = data.username ? data.username.toString().trim() : '';
  const password = data.password ? data.password.toString().trim() : '';
  
  if (!username || !password) {
    return { success: false, message: 'Username dan Password tidak boleh kosong!' };
  }
  
  const techData = sheet.getDataRange().getValues();
  for (let i = 1; i < techData.length; i++) {
    const user = techData[i][0] ? techData[i][0].toString().trim() : '';
    if (user.toLowerCase() === username.toLowerCase()) {
      return { success: false, message: 'Username sudah terdaftar! Silakan gunakan username lain.' };
    }
  }
  
  let fotoUrl = '';
  if (data.foto) {
    const fileName = 'REG_' + username + '_' + new Date().getTime() + '.png';
    fotoUrl = uploadImageToDrive(data.foto, fileName);
  }
  
  const waktuRegister = new Date().toLocaleString('id-ID');
  sheet.appendRow([username, password, data.nik || '', data.sto || '', data.telegram || '', data.wa || '', fotoUrl, new Date().toLocaleString()]);
  
  return { success: true, message: 'Registrasi berhasil! Silakan login.' };
}

function registerAdmin(data) {
  const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  const role = data.role === 'HSA' ? 'Hsa' : 'Tif';
  const sheetName = role;
  
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return { success: false, message: `Database ${role} tidak ditemukan` };
  }
  
  const username = data.username ? data.username.toString().trim() : '';
  const password = data.password ? data.password.toString().trim() : '';
  const stos = data.stos ? data.stos.toString().trim() : '';
  const telegram = data.telegram ? data.telegram.toString().trim() : '';
  
  if (!username || !password) {
    return { success: false, message: 'Username dan Password tidak boleh kosong!' };
  }
  
  const adminData = sheet.getDataRange().getValues();
  for (let i = 1; i < adminData.length; i++) {
    const user = adminData[i][0] ? adminData[i][0].toString().trim() : '';
    if (user.toLowerCase() === username.toLowerCase()) {
      return { success: false, message: 'Username sudah terdaftar! Silakan gunakan username lain.' };
    }
  }
  
  if (role === 'Hsa') {
    sheet.appendRow([username, password, stos, telegram]);
  } else {
    sheet.appendRow([username, password, telegram]);
  }
  
  return { success: true, message: 'Registrasi berhasil! Silakan login.' };
}

function getHistoryCacheVersion() {
  const cache = CacheService.getScriptCache();
  let v = cache.get('HIST_VER');
  if (!v) {
    v = new Date().getTime().toString();
    try { cache.put('HIST_VER', v, 21600); } catch(e){}
  }
  return v;
}

function invalidateHistoryCache() {
  const cache = CacheService.getScriptCache();
  try { cache.put('HIST_VER', new Date().getTime().toString(), 21600); } catch(e){}
}

function getOdcHistory(odc) {
  const cache = CacheService.getScriptCache();
  const v = getHistoryCacheVersion();
  const cacheKey = 'HIST_ODC_' + odc + '_' + v;
  const cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch(e){}
  }

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
        selfieKembali: data[i][9]
      });
    }
  }
  
  try { cache.put(cacheKey, JSON.stringify(history), 1800); } catch(e){}
  return history;
}

function getTechnicianHistory(query) {
  if (!query) return { active: [], history: [] };
  query = query.toString().toLowerCase().trim();
  const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  
  const activeSheet = ss.getSheetByName('ActiveBorrowings');
  let activeList = [];
  if (activeSheet) {
    const techSheet = ss.getSheetByName('Teknisi');
    const nikMap = {};
    if (techSheet) {
      const techData = techSheet.getDataRange().getValues();
      for (let i = 1; i < techData.length; i++) {
        if (techData[i][0]) {
          nikMap[techData[i][0].toString().trim()] = techData[i][2] || '-';
        }
      }
    }

    const activeData = activeSheet.getDataRange().getValues();
    for (let i = 1; i < activeData.length; i++) {
      const id = (activeData[i][0] || '').toString().toLowerCase();
      const userName = (activeData[i][3] || '').toString().trim();
      const userLower = userName.toLowerCase();
      if (id.includes(query) || userLower.includes(query)) {
        activeList.push({
          id: activeData[i][0],
          sto: activeData[i][1],
          odc: activeData[i][2],
          user: userName,
          nik: nikMap[userName] || '-',
          kegiatan: activeData[i][4],
          estimasi: activeData[i][5],
          waktuPinjam: activeData[i][6],
          status: activeData[i][8] || 'Sedang Dipinjam',
          hsaApprover: activeData[i][9] || '-',
          tifApprover: activeData[i][12] || '-'
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
      if (historyList.length >= 20) break;
    }
  }
  
  return { active: activeList, history: historyList };
}

function getUsersList() {
  const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  const sheet = ss.getSheetByName('Teknisi');
  if (!sheet) return [];
  
  const techData = sheet.getDataRange().getValues();
  const users = [];
  
  for (let i = 1; i < techData.length; i++) {
    const user = techData[i][0] ? techData[i][0].toString().trim() : '';
    if (user) {
      users.push({
        username: user,
        nik: techData[i][2] || '-',
        sto: techData[i][3] || '-',
        telegram: techData[i][4] || '-',
        wa: techData[i][5] || '-',
        fotoUrl: techData[i][6] || '',
        waktuRegister: techData[i][7] || '-'
      });
    }
  }
  return users;
}

function getAllHistory() {
  const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  const history = [];
  
  const activeSheet = ss.getSheetByName('ActiveBorrowings');
  if (activeSheet) {
    const activeData = activeSheet.getDataRange().getValues();
    for (let i = activeData.length - 1; i > 0; i--) {
      history.push({
        id: activeData[i][0],
        sto: activeData[i][1],
        odc: activeData[i][2],
        user: activeData[i][3],
        kegiatan: activeData[i][4],
        estimasi: activeData[i][5],
        waktuPinjam: activeData[i][6],
        waktuKembali: 'Belum Kembali'
      });
    }
  }

  const historySheet = ss.getSheetByName('History');
  if (historySheet) {
    const historyData = historySheet.getDataRange().getValues();
    for (let i = historyData.length - 1; i > 0; i--) {
      history.push({
        id: historyData[i][0],
        sto: historyData[i][1],
        odc: historyData[i][2],
        user: historyData[i][3],
        kegiatan: historyData[i][4],
        estimasi: historyData[i][5],
        waktuPinjam: historyData[i][6],
        waktuKembali: historyData[i][8] || 'Sudah Kembali'
      });
      if (history.length >= 100) break;
    }
  }
  
  try { cache.put(cacheKey, JSON.stringify(history), 1800); } catch(e){}
  return history;
}

function loginHsa(payload) {
  const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  const sheet = ss.getSheetByName('Hsa');
  if (!sheet) return { success: false, message: 'Database HSA tidak ditemukan' };
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString() === payload.user && data[i][1] && data[i][1].toString() === payload.pass) {
      const stos = data[i][2] ? data[i][2].toString().split(',').map(s => s.trim()) : [];
      return { success: true, username: data[i][0], stos: stos, role: 'HSA' };
    }
  }
  return { success: false, message: 'Username atau Password salah!' };
}

function loginTif(payload) {
  const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  const sheet = ss.getSheetByName('Tif');
  if (!sheet) return { success: false, message: 'Database TIF tidak ditemukan' };
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString() === payload.user && data[i][1] && data[i][1].toString() === payload.pass) {
      return { success: true, username: data[i][0], role: 'TIF' };
    }
  }
  return { success: false, message: 'Username atau Password salah!' };
}

function getPendingApprovals(statusFilter) {
  const active = getActiveBorrowings();
  return active.filter(item => item.status === statusFilter);
}

function actionApproval(payload) {
  const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  const activeSheet = ss.getSheetByName('ActiveBorrowings');
  
  if (!activeSheet) return { success: false, message: 'ActiveBorrowings sheet missing' };
  
  const data = activeSheet.getDataRange().getValues();
  let rowIndex = -1;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString() === payload.id) {
      rowIndex = i + 1;
      break;
    }
  }
  
  if (rowIndex === -1) return { success: false, message: 'Tiket tidak ditemukan' };
  
  let newStatus = '';
  if (payload.actionType === 'Approve') {
    if (payload.role === 'HSA') {
      newStatus = 'Pending TIF';
      activeSheet.getRange(rowIndex, 10).setValue(payload.name);
    } else if (payload.role === 'TIF') {
      newStatus = 'Approved';
      activeSheet.getRange(rowIndex, 13).setValue(payload.name);
    }
  } else {
    newStatus = 'Rejected';
    if (payload.role === 'HSA') activeSheet.getRange(rowIndex, 10).setValue(payload.name);
    if (payload.role === 'TIF') activeSheet.getRange(rowIndex, 13).setValue(payload.name);
  }
  
  activeSheet.getRange(rowIndex, 9).setValue(newStatus);
  
  // Kirim Notifikasi Telegram jika status menjadi Approved
  if (newStatus === 'Approved') {
    const rowData = activeSheet.getRange(rowIndex, 1, 1, 13).getValues()[0];
    const ticketId = rowData[0];
    const odc = rowData[2];
    const user = rowData[3];
    const kegiatan = rowData[4];
    const estimasi = rowData[5];
    
    const contact = getTechnicianContact(user);
    
    const msg = `🟢 <b>KUNCI DIPINJAM</b>\n\n` +
                `<b>ID:</b> ${ticketId}\n` +
                `<b>ODC:</b> ${odc}\n` +
                `<b>Teknisi:</b> ${user}\n` +
                `<b>ID Telegram:</b> ${contact.telegram}\n` +
                `<b>No WA:</b> ${contact.wa}\n` +
                `<b>Kegiatan:</b> ${kegiatan}\n` +
                `<b>Estimasi Waktu:</b> ${estimasi} Jam\n\n` +
                `<i>Timer sedang berjalan. Peringatan akan dikirim 10 menit sebelum waktu habis.</i>`;
    sendTelegramMessage(msg);
  }
  
  return { success: true };
}

function getTechnicianContact(username) {
  const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  const sheet = ss.getSheetByName('Teknisi');
  if (!sheet) return { telegram: '-', wa: '-' };
  if (!username) return { telegram: '-', wa: '-' };
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() === username.toString().trim()) {
      return { 
        telegram: data[i][4] || '-', 
        wa: data[i][5] || '-' 
      };
    }
  }
  return { telegram: '-', wa: '-' };
}

// ==========================================
// TELEGRAM TRIGGER & SCHEDULER
// ==========================================

function checkExpiringTimers() {
  const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  const activeSheet = ss.getSheetByName('ActiveBorrowings');
  if (!activeSheet) return;
  
  const data = activeSheet.getDataRange().getValues();
  const nowMs = new Date().getTime();
  
  for (let i = 1; i < data.length; i++) {
    const status = data[i][8];
    if (status === 'Approved' || status === 'Sedang Dipinjam') {
      const warned = data[i][13]; // Kolom N (index 13)
      const lastOverdueWarningMs = parseInt(data[i][14]) || 0; // Kolom O (index 14)
      const waktuPinjam = data[i][6];
      const estimasi = parseInt(data[i][5]) || 0;
      
      if (waktuPinjam && estimasi > 0) {
        const borrowTimeMs = parseIdDateForServer(waktuPinjam).getTime();
        const alarmTimeMs = borrowTimeMs + (estimasi * 60 * 60 * 1000);
        
        const diffMs = alarmTimeMs - nowMs;
        const diffMinutes = Math.floor(diffMs / (1000 * 60));
        
        const ticketId = data[i][0];
        const user = data[i][3];
        const odc = data[i][2];
        const contact = getTechnicianContact(user);
        
        // 1. Peringatan 10 Menit sebelum habis (Hanya sekali)
        if (warned !== 'Yes' && diffMinutes <= 10 && diffMinutes >= 0) {
          const msg = `⚠️ <b>PERINGATAN WAKTU HABIS</b>\n\n` +
                      `<b>ID:</b> ${ticketId}\n` +
                      `<b>ODC:</b> ${odc}\n` +
                      `<b>Teknisi:</b> ${user}\n` +
                      `<b>ID Telegram:</b> ${contact.telegram}\n` +
                      `<b>No WA:</b> ${contact.wa}\n\n` +
                      `Waktu peminjaman tersisa kurang dari 10 menit! Segera selesaikan pekerjaan dan kembalikan kunci.`;
          sendTelegramMessage(msg);
          activeSheet.getRange(i + 1, 14).setValue('Yes');
        }
        
        // 2. Peringatan Overdue Setiap 1 Jam setelah waktu habis
        if (diffMinutes < 0) {
          const timeSinceLastWarning = nowMs - lastOverdueWarningMs;
          if (timeSinceLastWarning >= 3600000) { // 3600000 ms = 1 Jam
            const overdueHours = Math.abs(Math.floor(diffMinutes / 60));
            let keterlambatanText = overdueHours > 0 ? `(Terlambat ${overdueHours} Jam)` : '(Waktu Habis)';
            const msg = `🚨 <b>ALARM: KUNCI BELUM DIKEMBALIKAN!</b>\n\n` +
                        `<b>ID:</b> ${ticketId}\n` +
                        `<b>ODC:</b> ${odc}\n` +
                        `<b>Teknisi:</b> ${user}\n` +
                        `<b>ID Telegram:</b> ${contact.telegram}\n` +
                        `<b>No WA:</b> ${contact.wa}\n\n` +
                        `Waktu peminjaman sudah habis ${keterlambatanText}. Segera kembalikan kunci ODC!`;
            sendTelegramMessage(msg);
            activeSheet.getRange(i + 1, 15).setValue(nowMs.toString());
          }
        }
      }
    }
  }
}

function parseIdDateForServer(dateStr) {
  if (!dateStr) return new Date();
  let cleanStr = dateStr.toString().replace(/\./g, ':');
  let parts = cleanStr.split(/[\s,]+/);
  if (parts.length < 2) return new Date(cleanStr);
  
  let dParts = parts[0].split('/');
  if (dParts.length === 3) {
    let day = parseInt(dParts[0]), month = parseInt(dParts[1]) - 1, year = parseInt(dParts[2]);
    let tParts = parts[1].split(':');
    let hours = parseInt(tParts[0]) || 0, minutes = parseInt(tParts[1]) || 0, seconds = parseInt(tParts[2]) || 0;
    return new Date(year, month, day, hours, minutes, seconds);
  }
  return new Date(dateStr);
}

function setupTrigger() {
  // Hapus trigger lama jika ada
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'checkExpiringTimers') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Buat trigger baru setiap 5 menit
  ScriptApp.newTrigger('checkExpiringTimers')
    .timeBased()
    .everyMinutes(5)
    .create();
}

function getApprovedBa() {
  const active = getActiveBorrowings();
  return active.filter(item => item.status === 'Approved');
}

// ==========================================
// VERIFIKASI EVIDENCE TIF (MASTER ADMIN)
// ==========================================
function getTifVerificationsData() {
  const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  const sheet = ss.getSheetByName('VerifikasiEvidenceTIF');
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return {};
  
  const map = {};
  for (let i = 1; i < data.length; i++) {
    const odc = (data[i][0] || '').toString().trim();
    if (odc) {
      map[odc] = {
        status: (data[i][1] || '').toString().trim(),
        note: (data[i][2] || '').toString().trim(),
        updatedBy: (data[i][3] || '').toString().trim(),
        updatedAt: (data[i][4] || '').toString().trim()
      };
    }
  }
  return map;
}

function verifyEvidenceTif(payload) {
  if (!payload || !payload.odc) {
    return { success: false, message: 'ODC tidak valid' };
  }
  const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  let sheet = ss.getSheetByName('VerifikasiEvidenceTIF');
  if (!sheet) {
    sheet = ss.insertSheet('VerifikasiEvidenceTIF');
    sheet.appendRow(['ODC', 'Status', 'Catatan', 'Reviewer', 'Waktu']);
  }

  const odcTarget = payload.odc.toString().trim();
  const status = (payload.status !== undefined ? payload.status : '').toString().trim();
  const note = (payload.note !== undefined ? payload.note : '').toString().trim();
  const reviewer = (payload.reviewer || 'TIF').toString().trim();
  const time = payload.time || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');

  const data = sheet.getDataRange().getValues();
  let foundRow = -1;
  for (let i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toString().trim().toUpperCase() === odcTarget.toUpperCase()) {
      foundRow = i + 1; // 1-based index
      break;
    }
  }

  if (foundRow !== -1) {
    const currentStatus = data[foundRow - 1][1] || '';
    const currentNote = data[foundRow - 1][2] || '';
    const finalStatus = (payload.status !== undefined) ? status : currentStatus;
    const finalNote = (payload.note !== undefined) ? note : currentNote;

    sheet.getRange(foundRow, 1, 1, 5).setValues([[odcTarget, finalStatus, finalNote, reviewer, time]]);
  } else {
    sheet.appendRow([odcTarget, status, note, reviewer, time]);
  }

  return { success: true, odc: odcTarget, status: status, note: note };
}
