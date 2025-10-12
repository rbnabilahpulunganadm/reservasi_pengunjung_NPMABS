// =================================================================
//      BACKEND - KLINIK NABILAH PULUNGAN (Google Apps Script) - VERSI REVISI
// =================================================================
// Deskripsi Revisi:
// - Menambahkan manajemen Terapis (sheet, CRUD functions).
// - Menambahkan fitur notifikasi untuk reservasi baru (kolom 'Telah_Dilihat', fungsi get & mark).
// - Menambahkan fitur pengecekan duplikasi data pasien sebelum membuat reservasi baru.
// =================================================================

// --- 1. KONSTANTA GLOBAL ---
const PATIENT_SHEET_NAME = 'Pasien';
const RESERVATION_SHEET_NAME = 'Reservasi';
const TREATMENT_SHEET_NAME = 'Treatments';
const PRODUCT_SHEET_NAME = 'Products';
const THERAPIST_SHEET_NAME = 'Terapis'; // [FITUR BARU] Sheet untuk terapis
const LOG_SHEET_NAME = 'Log'; 

const TEMPLATE_ID = '1a9EMmne_y3pDUu5yoq7a9L0zVcgV2h-pfoyGRbooWak'; 

// --- 2. FUNGSI UTAMA WEB APP (ENTRY POINT) ---

function doGet(e) {
  try {
    const action = e.parameter.action;
    const payload = e.parameter.payload ? JSON.parse(e.parameter.payload) : {};
    let response;

    switch(action) {
      case 'getReservationsAndNotifications': response = getReservationsAndNotifications(); break; // [FITUR BARU] Mengambil reservasi + notif
      case 'getPatients':       response = getPatients(payload.query); break;
      case 'getItems':          response = getItems(); break;
      case 'getRekapData':      response = getRekapData(); break;
      case 'getPatientHistory': response = getPatientHistory(payload); break; 
      case 'generatePdf':       response = generatePatientStatusPdf(payload); break;
      case 'getTherapists':     response = getTherapists(); break; // [FITUR BARU]
      case 'checkExistingPatient': response = checkExistingPatient(payload); break; // [FITUR BARU]
      default:                  response = { status: 'error', message: 'Invalid GET action' };
    }
    return createJsonResponse(response);
  } catch (error) {
    logError('doGet', error);
    return createJsonResponse({ status: 'error', message: error.message });
  }
}

function doPost(e) {
  try {
    const requestData = JSON.parse(e.postData.contents);
    const action = requestData.action;
    const payload = requestData.payload;
    let response;

    if (!action) return createJsonResponse({ status: 'error', message: 'Action not specified' });

    switch (action) {
      case 'newReservation':      response = handleNewReservation(payload); break;
      case 'completeReservation': response = handleCompleteReservation(payload); break;
      case 'addOrUpdateItem':     response = handleAddOrUpdateItem(payload); break;
      case 'updatePatient':       response = handleUpdatePatient(payload); break;
      case 'addOrUpdateTherapist': response = handleAddOrUpdateTherapist(payload); break; // [FITUR BARU]
      case 'markReservationsAsSeen': response = markReservationsAsSeen(payload); break; // [FITUR BARU]
      default:                    response = { status: 'error', message: 'Invalid POST action' };
    }
    return createJsonResponse(response);
  } catch (error) {
    logError('doPost', error);
    return createJsonResponse({ status: 'error', message: error.message, stack: error.stack });
  }
}


// --- 3. FUNGSI-FUNGSI HANDLER (LOGIKA BISNIS) ---

/**
 * [FITUR BARU]
 * Memeriksa apakah ada pasien yang sudah ada dengan nama yang mirip.
 * @param {object} payload - Berisi `patientName` dan `requesterName`.
 * @returns {object} - Daftar pasien yang cocok.
 */
function checkExistingPatient(payload) {
    const { patientName, requesterName } = payload;
    if (!patientName && !requesterName) {
        return { status: 'success', data: [] };
    }

    const patientSheet = getSheet(PATIENT_SHEET_NAME);
    const patients = sheetToJSON(patientSheet);

    const searchWords = `${patientName.toLowerCase()} ${requesterName.toLowerCase()}`.split(/\s+/).filter(w => w.length > 2);

    const matches = patients.filter(p => {
        const patientWords = `${p.Nama_Pasien.toLowerCase()} ${p.Nama_Pemesan.toLowerCase()}`.split(/\s+/);
        // Cek jika ada kata dari input yang sama dengan kata dari data pasien
        return searchWords.some(searchWord => patientWords.includes(searchWord));
    });

    return { status: 'success', data: matches };
}

/**
 * [FUNGSI YANG DIPERBARUI]
 * Menangani pembuatan reservasi baru. Sekarang bisa menggunakan RME yang sudah ada.
 * @param {object} data - Data lengkap dari form reservasi, mungkin termasuk `existingRME`.
 * @returns {object} - Objek status dan RME pasien.
 */
function handleNewReservation(data) {
    let rme;
    // Jika ada 'existingRME', berarti pengguna memilih pasien lama.
    if (data.existingRME) {
        rme = data.existingRME;
    } else {
        // Jika tidak, cari atau buat pasien baru seperti biasa.
        const patientRecord = findOrCreatePatient(data);
        rme = patientRecord.rme;
    }

    const reservationSheet = getSheet(RESERVATION_SHEET_NAME);

    if (data.visitTime !== '24_jam') {
      const reservations = sheetToJSON(reservationSheet);
      const visitDateTime = new Date(`${data.visitDate}T${data.visitTime}`);
      const conflict = reservations.filter(r => {
          if (!r.Tanggal_Datang || r.Status === 'Selesai') return false;
          try { return new Date(r.Tanggal_Datang).getTime() === visitDateTime.getTime(); } catch (err) { return false; }
      });
      if (conflict.length >= 2) {
        return { status: 'error', message: 'Slot pada jam dan tanggal tersebut sudah penuh (Maks 2 reservasi).' };
      }
    }

    const reservationId = 'RES-' + Date.now();
    const visitDateTime = new Date(`${data.visitDate}T${data.visitTime}`);
    const newReservation = [
      reservationId, new Date(), 'Menunggu', rme, data.patientName, data.requesterName,
      data.phone, data.address, 
      !isNaN(visitDateTime.getTime()) ? visitDateTime.toISOString() : `${data.visitDate}T${data.visitTime}`,
      data.visitTime, JSON.stringify(data.selectedItems.map(item => item.name)),
      data.complaint || '', data.notes || '', '', '', '', false // Tambah kolom 'Telah_Dilihat' = false
    ];
    reservationSheet.appendRow(newReservation);
    return { status: 'success', message: 'Reservasi berhasil dibuat.', data: { rme: rme } };
}

/**
 * [FITUR BARU]
 * Menangani penambahan atau pembaruan data terapis.
 * @param {object} payload - Data terapis (id, name, status).
 * @returns {object} - Objek status.
 */
function handleAddOrUpdateTherapist(payload) {
    const { id, name, status } = payload;
    const sheet = getSheet(THERAPIST_SHEET_NAME);
    
    if (id) { // Update
        const data = sheet.getDataRange().getValues();
        const rowIndex = data.findIndex(row => row[0].toString() === id.toString());
        if (rowIndex !== -1) {
            const realRowIndex = rowIndex + 1;
            sheet.getRange(realRowIndex, 1, 1, 3).setValues([[id, name, status]]);
            return { status: 'success', message: 'Data terapis berhasil diperbarui.' };
        } else {
            return { status: 'error', message: 'Terapis tidak ditemukan.' };
        }
    } else { // Tambah baru
        const newId = 'TRP-' + Date.now();
        sheet.appendRow([newId, name, status || 'Aktif']);
        return { status: 'success', message: 'Terapis berhasil ditambahkan.' };
    }
}

/**
 * [FITUR BARU]
 * Menandai reservasi sebagai 'telah dilihat' untuk mematikan notifikasi.
 * @param {object} payload - Berisi array `reservationIds`.
 * @returns {object} - Objek status.
 */
function markReservationsAsSeen(payload) {
    const { reservationIds } = payload;
    if (!reservationIds || reservationIds.length === 0) {
        return { status: 'error', message: 'Tidak ada ID reservasi yang diberikan.' };
    }

    const sheet = getSheet(RESERVATION_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idColIndex = headers.indexOf('ID_Reservasi');
    const seenColIndex = headers.indexOf('Telah_Dilihat');

    if (idColIndex === -1 || seenColIndex === -1) {
        return { status: 'error', message: 'Kolom ID_Reservasi atau Telah_Dilihat tidak ditemukan.' };
    }

    data.forEach((row, index) => {
        if (index > 0 && reservationIds.includes(row[idColIndex])) {
            // Cek untuk menghindari penulisan yang tidak perlu
            if (row[seenColIndex] !== true) {
                sheet.getRange(index + 1, seenColIndex + 1).setValue(true);
            }
        }
    });

    return { status: 'success', message: 'Reservasi telah ditandai.' };
}

// --- FUNGSI-FUNGSI LAINNYA (TIDAK BERUBAH SECARA SIGNIFIKAN, KECUALI YANG DISEBUTKAN) ---

function findOrCreatePatient(data) {
    const patientSheet = getSheet(PATIENT_SHEET_NAME);
    const patients = sheetToJSON(patientSheet);
    let existingPatient = patients.find(p => 
        p.Nama_Pasien.toLowerCase() === data.patientName.toLowerCase() && 
        String(p.No_HP).trim() === String(data.phone).trim()
    );

    if (existingPatient) {
        return { ...existingPatient, rme: existingPatient.RME, isNew: false };
    } else {
        const newRME = generateRME(patientSheet);
        const dobDate = new Date(data.dob);
        const newPatientRow = [ newRME, data.patientName, data.requesterName, data.phone, data.instagram || '', data.address, !isNaN(dobDate.getTime()) ? dobDate.toISOString() : data.dob, new Date(), data.gender ];
        patientSheet.appendRow(newPatientRow);
        return { rme: newRME, isNew: true };
    }
}

function handleCompleteReservation(payload) {
    const { reservationId, therapist, updatedItems, updatedComplaint } = payload;
    const reservationSheet = getSheet(RESERVATION_SHEET_NAME);
    const data = reservationSheet.getDataRange().getValues();
    const headers = data[0];
    const idColIndex = headers.indexOf('ID_Reservasi');
    if (idColIndex === -1) return { status: 'error', message: 'Kolom ID_Reservasi tidak ditemukan.'};
    
    let rowIndex = data.findIndex(row => row[idColIndex] === reservationId);
    if (rowIndex === -1) return { status: 'error', message: 'Reservasi tidak ditemukan.'};
    
    const realRowIndex = rowIndex + 1; 
    const examData = { suhu: payload.temp, berat: payload.weight, tinggi: payload.height, lila: payload.lila, catatan: payload.examNotes };
    
    reservationSheet.getRange(realRowIndex, headers.indexOf('Status') + 1).setValue('Selesai');
    reservationSheet.getRange(realRowIndex, headers.indexOf('Terapis') + 1).setValue(therapist);
    reservationSheet.getRange(realRowIndex, headers.indexOf('Data_Pemeriksaan') + 1).setValue(JSON.stringify(examData));
    reservationSheet.getRange(realRowIndex, headers.indexOf('Items') + 1).setValue(JSON.stringify(updatedItems || []));
    reservationSheet.getRange(realRowIndex, headers.indexOf('Keluhan') + 1).setValue(updatedComplaint || '');
    
    return { status: 'success', message: 'Reservasi telah diselesaikan.'};
}

// (Fungsi lain seperti generatePdf, getPatientHistory, handleUpdatePatient, handleAddOrUpdateItem tidak berubah)
// ... (salin fungsi-fungsi tersebut dari kode asli Anda)


// --- 4. FUNGSI-FUNGSI PENGAMBILAN DATA (GETTERS) ---

/**
 * [FUNGSI YANG DIPERBARUI]
 * Mengambil semua reservasi dan menghitung jumlah notifikasi baru.
 */
function getReservationsAndNotifications() {
    const reservations = sheetToJSON(getSheet(RESERVATION_SHEET_NAME));
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    // Notifikasi dihitung jika: Status 'Menunggu', dibuat hari ini, dan belum dilihat.
    const newReservationCount = reservations.filter(r => {
        const reservationDate = new Date(r.Timestamp);
        return r.Status === 'Menunggu' && reservationDate >= today && r.Telah_Dilihat !== true;
    }).length;

    return { 
        status: 'success', 
        data: {
            reservations: reservations,
            newReservationCount: newReservationCount
        } 
    };
}

/**
 * [FITUR BARU]
 * Mengambil daftar semua terapis yang aktif.
 */
function getTherapists() {
    const therapists = sheetToJSON(getSheet(THERAPIST_SHEET_NAME));
    const activeTherapists = therapists.filter(t => t.Status === 'Aktif');
    return { status: 'success', data: activeTherapists };
}

function getItems() { return { status: 'success', data: { treatments: sheetToJSON(getSheet(TREATMENT_SHEET_NAME)), products: sheetToJSON(getSheet(PRODUCT_SHEET_NAME)) }}; }
function getPatients(query) {
    const patients = sheetToJSON(getSheet(PATIENT_SHEET_NAME));
    if (!query) return { status: 'success', data: patients };
    const lowerCaseQuery = query.toLowerCase();
    const filtered = patients.filter(p => (p.RME && p.RME.toLowerCase().includes(lowerCaseQuery)) || (p.Nama_Pasien && p.Nama_Pasien.toLowerCase().includes(lowerCaseQuery)));
    return { status: 'success', data: filtered };
}
// (Fungsi getRekapData tidak berubah)
// ... (salin fungsi getRekapData dari kode asli Anda)

// --- 5. FUNGSI-FUNGSI UTILITAS (HELPERS) ---

function createJsonResponse(data) { return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON); }
function getSheet(sheetName) { const ss = SpreadsheetApp.getActiveSpreadsheet(); let sheet = ss.getSheetByName(sheetName); if (!sheet) { sheet = ss.insertSheet(sheetName); const headers = getHeadersForSheet(sheetName); if (headers) { sheet.appendRow(headers); sheet.setFrozenRows(1); sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold'); } } return sheet; }
function getHeadersForSheet(sheetName) { 
    const headerMap = { 
        [PATIENT_SHEET_NAME]: ['RME', 'Nama_Pasien', 'Nama_Pemesan', 'No_HP', 'Instagram', 'Alamat', 'Tanggal_Lahir', 'Tanggal_Registrasi', 'Jenis_Kelamin'], 
        // [FITUR BARU] Tambah kolom 'Telah_Dilihat'
        [RESERVATION_SHEET_NAME]: ['ID_Reservasi', 'Timestamp', 'Status', 'RME', 'Nama_Pasien', 'Nama_Pemesan', 'No_HP', 'Alamat', 'Tanggal_Datang', 'Jam_Datang', 'Items', 'Keluhan', 'Catatan', 'Terapis', 'Data_Pemeriksaan', 'Telah_Dilihat'], 
        [TREATMENT_SHEET_NAME]: ['ID_Treatment', 'Kategori', 'Nama', 'Deskripsi'], 
        [PRODUCT_SHEET_NAME]: ['ID_Produk', 'Nama', 'Deskripsi'],
        [THERAPIST_SHEET_NAME]: ['ID_Terapis', 'Nama_Terapis', 'Status'], // [FITUR BARU]
        [LOG_SHEET_NAME]: ['Timestamp', 'Function', 'Message', 'ErrorStack'] 
    }; 
    return headerMap[sheetName]; 
}
function sheetToJSON(sheet) { const data = sheet.getDataRange().getValues(); if (data.length < 2) return []; const headers = data.shift(); return data.map(row => { let obj = {}; headers.forEach((col, index) => { obj[col] = row[index]; }); return obj; }); }
function generateRME(sheet) { const lastRow = sheet.getLastRow(); if (lastRow < 2) return 'NBLH-001'; try { const lastRME = sheet.getRange(lastRow, 1).getValue(); const lastNumber = parseInt(lastRME.split('-')[1]); const newNumber = (lastNumber + 1).toString().padStart(3, '0'); return `NBLH-${newNumber}`; } catch(e) { return `NBLH-${(lastRow).toString().padStart(3, '0')}`; } }
function logError(functionName, error) { try { getSheet(LOG_SHEET_NAME).appendRow([new Date(), functionName, error.message, error.stack]); } catch(e) { Logger.log("Gagal menulis log: " + e.message); } }

/**
 * [FUNGSI YANG DIPERBARUI]
 * Menjalankan setup awal, termasuk sheet Terapis.
 */
function setupInitialSheet() { 
    getSheet(PATIENT_SHEET_NAME); 
    getSheet(RESERVATION_SHEET_NAME); 
    getSheet(TREATMENT_SHEET_NAME); 
    getSheet(PRODUCT_SHEET_NAME);
    getSheet(THERAPIST_SHEET_NAME); // [FITUR BARU]
    getSheet(LOG_SHEET_NAME); 
    SpreadsheetApp.flush(); 
    Logger.log('Setup completed.'); 
}

// =================================================================
// SALIN DAN TEMPEL SEMUA FUNGSI LAIN DARI KODE ASLI ANDA KE SINI
// UNTUK MEMASTIKAN TIDAK ADA YANG TERLEWAT
// CONTOH: generatePatientStatusPdf, getPatientHistory, handleUpdatePatient, 
//         handleAddOrUpdateItem, getRekapData
// =================================================================
function generatePatientStatusPdf(payload) {
  if (TEMPLATE_ID === 'GANTI_DENGAN_ID_GOOGLE_SLIDE_ANDA' || !TEMPLATE_ID) {
    return { status: 'error', message: 'ID Template Google Slide belum diatur di skrip backend (Code.gs).' };
  }

  const { reservationId } = payload;
  if (!reservationId) return { status: 'error', message: 'ID Reservasi tidak valid.' };

  const reservations = sheetToJSON(getSheet(RESERVATION_SHEET_NAME));
  const patients = sheetToJSON(getSheet(PATIENT_SHEET_NAME));

  const reservation = reservations.find(r => r.ID_Reservasi === reservationId);
  if (!reservation) return { status: 'error', message: 'Data reservasi tidak ditemukan.' };

  const patient = patients.find(p => p.RME === reservation.RME);
  if (!patient) return { status: 'error', message: 'Data pasien tidak ditemukan.' };
  
  try {
    const copyFile = DriveApp.getFileById(TEMPLATE_ID).makeCopy(`Status Pasien - ${patient.Nama_Pasien} - ${new Date().getTime()}`);
    const presentation = SlidesApp.openById(copyFile.getId());
    const slide = presentation.getSlides()[0]; 

    const dob = new Date(patient.Tanggal_Lahir);
    const visitDate = new Date(reservation.Tanggal_Datang);
    
    let examData = {};
    try {
        if(reservation.Data_Pemeriksaan) {
            examData = JSON.parse(reservation.Data_Pemeriksaan);
        }
    } catch(e) {
        logError('generatePdf.parseExamData', e);
    }

    let ageString = 'Tanggal lahir invalid';
    if (!isNaN(dob.getTime())) {
      let years = visitDate.getFullYear() - dob.getFullYear();
      let months = visitDate.getMonth() - dob.getMonth();
      let days = visitDate.getDate() - dob.getDate();
      if (days < 0) { months--; days += new Date(visitDate.getFullYear(), visitDate.getMonth(), 0).getDate(); }
      if (months < 0) { years--; months += 12; }
      ageString = `${years} thn, ${months} bln, ${days} hr`;
    }
    
    const replacements = {
      '<<NAMABAYI>>': patient.Nama_Pasien || '',
      '<<TTL>>': !isNaN(dob.getTime()) ? dob.toLocaleDateString('id-ID', { day: '2-digit', month: 'long', year: 'numeric' }) : 'N/A',
      '<<UMUR>>': ageString,
      '<<JENISKELAMIN>>': patient.Jenis_Kelamin || '',
      '<<ALAMAT>>': patient.Alamat || '',
      '<<NAMAPEMESAN>>': reservation.Nama_Pemesan || '',
      '<<NOHP>>': reservation.No_HP || '',
      '<<INSTAGRAM>>': patient.Instagram || '',
      '<<TGL>>': !isNaN(visitDate.getTime()) ? visitDate.toLocaleDateString('id-ID', { weekday: 'long', day: '2-digit', month: 'long', year: 'numeric' }) : 'N/A',
      '<<KELUHAN>>': reservation.Keluhan || '',
      '<<TREATMENT>>': JSON.parse(reservation.Items || '[]').join(', '),
      '<<RME>>': patient.RME || '',
      '<<TIMESTAMP>>': new Date(reservation.Timestamp).toLocaleString('id-ID', { dateStyle: 'long', timeStyle: 'short' }) + ' WIB',
      '<<suhu>>': examData.suhu ? `${examData.suhu} °C` : 'N/A',
      '<<berat_badan>>': examData.berat ? `${examData.berat} kg` : 'N/A',
      '<<tinggi_badan>>': examData.tinggi ? `${examData.tinggi} cm` : 'N/A',
      '<<lila>>': examData.lila ? `${examData.lila} cm` : 'N/A',
      '<<TERAPIS>>': reservation.Terapis || 'N/A'
    };

    for (const placeholder in replacements) {
      slide.replaceAllText(placeholder, replacements[placeholder]);
    }

    presentation.saveAndClose();
    
    const pdfBlob = copyFile.getAs('application/pdf');
    const base64Pdf = Utilities.base64Encode(pdfBlob.getBytes());
    const fileName = `StatusReservasi-${patient.Nama_Pasien.replace(/ /g, '_')}-${patient.RME}.pdf`;

    copyFile.setTrashed(true);

    return { status: 'success', data: { base64: base64Pdf, fileName: fileName } };
  } catch (error) {
    logError('generatePatientStatusPdf', error);
    return { status: 'error', message: 'Gagal membuat PDF: ' + error.message };
  }
}

function getPatientHistory(payload) {
  if (!payload || !payload.rme) return { status: 'error', message: 'RME tidak valid.' };
  const reservations = sheetToJSON(getSheet(RESERVATION_SHEET_NAME));
  const patientHistory = reservations
    .filter(res => res.RME === payload.rme)
    .sort((a, b) => new Date(b.Tanggal_Datang) - new Date(a.Tanggal_Datang));
  return { status: 'success', data: patientHistory };
}

function handleUpdatePatient(payload) {
    const { rme, patientName, requesterName, phone, instagram, address, dob, gender } = payload;
    if (!rme) return { status: 'error', message: 'RME tidak ditemukan untuk pembaruan.' };

    const patientSheet = getSheet(PATIENT_SHEET_NAME);
    const data = patientSheet.getDataRange().getValues();
    const headers = data[0];
    const rmeColIndex = headers.indexOf('RME');
    if (rmeColIndex === -1) return { status: 'error', message: 'Kolom RME tidak ditemukan.'};
    
    const rowIndex = data.findIndex(row => row[rmeColIndex] === rme);
    if (rowIndex === -1) return { status: 'error', message: 'Pasien tidak ditemukan.'};
    
    const realRowIndex = rowIndex + 1;
    const dobDate = new Date(dob);
    
    patientSheet.getRange(realRowIndex, headers.indexOf('Nama_Pasien') + 1).setValue(patientName);
    patientSheet.getRange(realRowIndex, headers.indexOf('Nama_Pemesan') + 1).setValue(requesterName);
    patientSheet.getRange(realRowIndex, headers.indexOf('No_HP') + 1).setValue(phone);
    patientSheet.getRange(realRowIndex, headers.indexOf('Instagram') + 1).setValue(instagram);
    patientSheet.getRange(realRowIndex, headers.indexOf('Alamat') + 1).setValue(address);
    patientSheet.getRange(realRowIndex, headers.indexOf('Tanggal_Lahir') + 1).setValue(!isNaN(dobDate.getTime()) ? dobDate.toISOString() : dob);
    patientSheet.getRange(realRowIndex, headers.indexOf('Jenis_Kelamin') + 1).setValue(gender);
    
    const reservationSheet = getSheet(RESERVATION_SHEET_NAME);
    const resData = reservationSheet.getDataRange().getValues();
    const resHeaders = resData[0];
    resData.forEach((row, index) => {
        if (index > 0 && row[resHeaders.indexOf('RME')] === rme && row[resHeaders.indexOf('Status')] === 'Menunggu') {
            reservationSheet.getRange(index + 1, resHeaders.indexOf('Nama_Pasien') + 1).setValue(patientName);
        }
    });

    return { status: 'success', message: 'Data pasien berhasil diperbarui.' };
}

function handleAddOrUpdateItem(payload) {
    const { type, id, name, category, description } = payload;
    const sheetName = type === 'treatment' ? TREATMENT_SHEET_NAME : PRODUCT_SHEET_NAME;
    const sheet = getSheet(sheetName);
    
    if (id) { 
        const data = sheet.getDataRange().getValues();
        const rowIndex = data.findIndex(row => row[0].toString() === id.toString());
        if (rowIndex !== -1) {
            const realRowIndex = rowIndex + 1;
            const rowData = type === 'treatment' ? [id, category, name, description] : [id, name, description];
            sheet.getRange(realRowIndex, 1, 1, rowData.length).setValues([rowData]);
            return { status: 'success', message: 'Data berhasil diperbarui.' };
        } else {
            return { status: 'error', message: 'Item tidak ditemukan.' };
        }
    } else { 
        const newId = (type.charAt(0).toUpperCase()) + '-' + Date.now();
        const newRow = type === 'treatment' ? [newId, category, name, description] : [newId, name, description];
        sheet.appendRow(newRow);
        return { status: 'success', message: 'Data berhasil ditambahkan.' };
    }
}

function getRekapData() {
    const reservations = sheetToJSON(getSheet(RESERVATION_SHEET_NAME));
    const patients = sheetToJSON(getSheet(PATIENT_SHEET_NAME));
    const treatments = sheetToJSON(getSheet(TREATMENT_SHEET_NAME));
    
    const treatmentCategoryMap = treatments.reduce((map, item) => { map[item.Nama] = item.Kategori; return map; }, {});
    
    const stats = { categoryCounts: {}, treatmentNameCounts: {}, genderCounts: {}, dayCounts: {}, peakHourCounts: {}, monthCounts: {}, therapistCounts: {}, dailyTrend: {}, calendarData: {}, ageDemographics: { 'Bayi (0-1)': 0, 'Balita (2-5)': 0, 'Anak (6-12)': 0, 'Remaja (13-18)': 0, 'Dewasa (19-40)': 0, 'Lansia (41+)': 0 }, addressDemographics: {} };
    
    const days = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
    for(let i=0; i<24; i++) { stats.peakHourCounts[i.toString().padStart(2,'0') + ':00'] = 0; }
    
    reservations.forEach(res => {
        if (!res.Tanggal_Datang) return;
        const visitDate = new Date(res.Tanggal_Datang);
        if (isNaN(visitDate.getTime())) return;
        
        const dateKey = visitDate.toISOString().split('T')[0];
        stats.dailyTrend[dateKey] = (stats.dailyTrend[dateKey] || 0) + 1;
        stats.calendarData[dateKey] = (stats.calendarData[dateKey] || 0) + 1;
        
        try { JSON.parse(res.Items || '[]').forEach(itemName => { stats.treatmentNameCounts[itemName] = (stats.treatmentNameCounts[itemName] || 0) + 1; const category = treatmentCategoryMap[itemName]; if (category) stats.categoryCounts[category] = (stats.categoryCounts[category] || 0) + 1; }); } catch(e) {}
        
        const patient = patients.find(p => p.RME === res.RME);
        if (patient && patient.Jenis_Kelamin) { stats.genderCounts[patient.Jenis_Kelamin] = (stats.genderCounts[patient.Jenis_Kelamin] || 0) + 1; }
        
        stats.dayCounts[days[visitDate.getDay()]] = (stats.dayCounts[days[visitDate.getDay()]] || 0) + 1;
        const hourKey = visitDate.getHours().toString().padStart(2, '0') + ':00';
        if (stats.peakHourCounts.hasOwnProperty(hourKey)) stats.peakHourCounts[hourKey]++;
        
        const monthYearKey = `${months[visitDate.getMonth()]} ${visitDate.getFullYear()}`;
        stats.monthCounts[monthYearKey] = (stats.monthCounts[monthYearKey] || 0) + 1;
        
        if (res.Terapis) { stats.therapistCounts[res.Terapis] = (stats.therapistCounts[res.Terapis] || 0) + 1; }
    });
    
    const addressCounts = {};
    patients.forEach(p => {
        if (p.Tanggal_Lahir) {
            const dob = new Date(p.Tanggal_Lahir);
            if (!isNaN(dob.getTime())) {
                const age = (new Date() - dob) / (1000 * 60 * 60 * 24 * 365.25);
                if (age <= 1) stats.ageDemographics['Bayi (0-1)']++; else if (age <= 5) stats.ageDemographics['Balita (2-5)']++; else if (age <= 12) stats.ageDemographics['Anak (6-12)']++; else if (age <= 18) stats.ageDemographics['Remaja (13-18)']++; else if (age <= 40) stats.ageDemographics['Dewasa (19-40)']++; else stats.ageDemographics['Lansia (41+)']++;
            }
        }
        if (p.Alamat && p.Alamat.trim() !== '') { const address = p.Alamat.trim().toLowerCase(); addressCounts[address] = (addressCounts[address] || 0) + 1; }
    });
    stats.addressDemographics = Object.entries(addressCounts).sort(([,a],[,b]) => b-a).slice(0, 10).reduce((r, [k, v]) => { const capitalizedKey = k.replace(/\b\w/g, l => l.toUpperCase()); r[capitalizedKey] = v; return r; }, {});

    return { status: 'success', data: {stats, rawReservations: reservations} };
}
