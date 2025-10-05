// =================================================================
//      BACKEND - KLINIK NABILAH PULUNGAN (Google Apps Script)
// =================================================================
// File: googleappscript.js
// Author: Gemini
// Instructions:
// 1. Ganti seluruh kode di Apps Script Anda dengan kode ini.
// 2. Tidak perlu menjalankan `setupInitialSheet` lagi jika sheet sudah ada.
// 3. Deploy ulang sebagai Web App (PENTING: Pilih versi BARU saat deploy ulang).
// =================================================================

const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const PATIENT_SHEET_NAME = 'Pasien';
const RESERVATION_SHEET_NAME = 'Reservasi';
const TREATMENT_SHEET_NAME = 'Treatments';
const PRODUCT_SHEET_NAME = 'Products';
const LOG_SHEET_NAME = 'Log';

function doGet(e) {
  try {
    const action = e.parameter.action;
    const payload = e.parameter.payload ? JSON.parse(e.parameter.payload) : {};
    let response;

    switch(action) {
      case 'getReservations': response = getReservations(); break;
      case 'getPatients': response = getPatients(payload.query); break;
      case 'getItems': response = getItems(); break;
      case 'getRekapData': response = getRekapData(); break;
      case 'getPatientHistory': response = getPatientHistory(payload); break; 
      default: response = { status: 'error', message: 'Invalid GET action' };
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
      case 'newReservation': response = handleNewReservation(payload); break;
      case 'completeReservation': response = handleCompleteReservation(payload); break;
      case 'addOrUpdateItem': response = handleAddOrUpdateItem(payload); break;
      default: response = { status: 'error', message: 'Invalid POST action' };
    }
    return createJsonResponse(response);
  } catch (error) {
    logError('doPost', error);
    return createJsonResponse({ status: 'error', message: error.message, stack: error.stack });
  }
}

function getPatientHistory(payload) {
  if (!payload || !payload.rme) {
    return { status: 'error', message: 'RME tidak valid.' };
  }
  const reservations = sheetToJSON(getSheet(RESERVATION_SHEET_NAME));
  const patientHistory = reservations
    .filter(res => res.RME === payload.rme)
    .sort((a, b) => new Date(b.Tanggal_Datang) - new Date(a.Tanggal_Datang));
  
  return { status: 'success', data: patientHistory };
}

function handleNewReservation(data) {
  const patientRecord = findOrCreatePatient(data);
  const rme = patientRecord.rme;
  const reservationSheet = getSheet(RESERVATION_SHEET_NAME);
  
  if (data.visitTime !== '24_jam') {
    const reservations = sheetToJSON(reservationSheet);
    const visitDateTime = new Date(`${data.visitDate}T${data.visitTime}`);
    const conflict = reservations.filter(r => {
        if (!r.Tanggal_Datang || r.Status === 'Selesai') return false;
        try {
            const existingDate = new Date(r.Tanggal_Datang);
            return existingDate.getTime() === visitDateTime.getTime();
        } catch (err) { return false; }
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
    data.complaint || '', data.notes || '', '', ''
  ];
  reservationSheet.appendRow(newReservation);
  return { status: 'success', message: 'Reservasi berhasil dibuat.', data: { rme: rme } };
}

function findOrCreatePatient(data) {
    const patientSheet = getSheet(PATIENT_SHEET_NAME);
    const patients = sheetToJSON(patientSheet);

    // --- PERBAIKAN DI SINI ---
    // Memastikan perbandingan Nama (case-insensitive) dan No HP (sebagai string)
    let existingPatient = patients.find(p => 
        p.Nama_Pasien.toLowerCase() === data.patientName.toLowerCase() && 
        String(p.No_HP).trim() === String(data.phone).trim()
    );
    // -------------------------

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
    const { reservationId, therapist } = payload;
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
    return { status: 'success', message: 'Reservasi telah diselesaikan.'};
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
        } else return { status: 'error', message: 'Item tidak ditemukan.' };
    } else {
        const newId = (type.charAt(0).toUpperCase()) + '-' + Date.now();
        const newRow = type === 'treatment' ? [newId, category, name, description] : [newId, name, description];
        sheet.appendRow(newRow);
        return { status: 'success', message: 'Data berhasil ditambahkan.' };
    }
}

function getReservations() { return { status: 'success', data: sheetToJSON(getSheet(RESERVATION_SHEET_NAME)) }; }
function getItems() { return { status: 'success', data: { treatments: sheetToJSON(getSheet(TREATMENT_SHEET_NAME)), products: sheetToJSON(getSheet(PRODUCT_SHEET_NAME)) } }; }
function getPatients(query) {
    const patients = sheetToJSON(getSheet(PATIENT_SHEET_NAME));
    if (!query) return { status: 'success', data: patients };
    const lowerCaseQuery = query.toLowerCase();
    const filtered = patients.filter(p => (p.RME && p.RME.toLowerCase().includes(lowerCaseQuery)) || (p.Nama_Pasien && p.Nama_Pasien.toLowerCase().includes(lowerCaseQuery)));
    return { status: 'success', data: filtered };
}

function getRekapData() {
    const reservations = sheetToJSON(getSheet(RESERVATION_SHEET_NAME));
    const patients = sheetToJSON(getSheet(PATIENT_SHEET_NAME));
    const treatments = sheetToJSON(getSheet(TREATMENT_SHEET_NAME));
    const treatmentCategoryMap = treatments.reduce((map, item) => { map[item.Nama] = item.Kategori; return map; }, {});
    const patientGenderMap = patients.reduce((map, p) => { map[p.RME] = p.Jenis_Kelamin; return map; }, {});
    const stats = {
        categoryCounts: {}, treatmentNameCounts: {}, genderCounts: {'Wanita': 0, 'Pria': 0},
        dayCounts: {'Minggu': 0, 'Senin': 0, 'Selasa': 0, 'Rabu': 0, 'Kamis': 0, 'Jumat': 0, 'Sabtu': 0},
        peakHourCounts: {}, monthCounts: {}, therapistCounts: {}, dailyTrend: {}, calendarData: {},
        ageDemographics: { 'Bayi (0-1)': 0, 'Balita (2-5)': 0, 'Anak (6-12)': 0, 'Remaja (13-18)': 0, 'Dewasa (19-40)': 0, 'Lansia (41+)': 0 },
        addressDemographics: {}
    };
    const days = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
    for(let i=0; i<24; i++) { stats.peakHourCounts[i.toString().padStart(2,'0') + ':00'] = 0; }
    
    reservations.forEach(res => {
        if (!res.Tanggal_Datang) return;
        const visitDate = new Date(res.Tanggal_Datang);
        if (isNaN(visitDate.getTime())) return;

        const dateKey = visitDate.toISOString().split('T')[0];
        stats.calendarData[dateKey] = (stats.calendarData[dateKey] || 0) + 1;
        try {
            JSON.parse(res.Items || '[]').forEach(itemName => {
                stats.treatmentNameCounts[itemName] = (stats.treatmentNameCounts[itemName] || 0) + 1;
                const category = treatmentCategoryMap[itemName];
                if (category) stats.categoryCounts[category] = (stats.categoryCounts[category] || 0) + 1;
            });
        } catch(e) {}
        const gender = patientGenderMap[res.RME];
        if (gender && stats.genderCounts.hasOwnProperty(gender)) stats.genderCounts[gender]++;
        stats.dayCounts[days[visitDate.getDay()]]++;
        const hourKey = visitDate.getHours().toString().padStart(2, '0') + ':00';
        if (stats.peakHourCounts.hasOwnProperty(hourKey)) stats.peakHourCounts[hourKey]++;
        const monthYearKey = `${months[visitDate.getMonth()]} ${visitDate.getFullYear()}`;
        stats.monthCounts[monthYearKey] = (stats.monthCounts[monthYearKey] || 0) + 1;
        if (res.Status === 'Selesai' && res.Terapis) { stats.therapistCounts[res.Terapis] = (stats.therapistCounts[res.Terapis] || 0) + 1; }
        stats.dailyTrend[dateKey] = (stats.dailyTrend[dateKey] || 0) + 1;
    });

    // Process patient demographics
    const addressCounts = {};
    patients.forEach(p => {
        // Age
        if (p.Tanggal_Lahir) {
            const dob = new Date(p.Tanggal_Lahir);
            if (!isNaN(dob.getTime())) {
                const age = (new Date() - dob) / (1000 * 60 * 60 * 24 * 365.25);
                if (age <= 1) stats.ageDemographics['Bayi (0-1)']++;
                else if (age <= 5) stats.ageDemographics['Balita (2-5)']++;
                else if (age <= 12) stats.ageDemographics['Anak (6-12)']++;
                else if (age <= 18) stats.ageDemographics['Remaja (13-18)']++;
                else if (age <= 40) stats.ageDemographics['Dewasa (19-40)']++;
                else stats.ageDemographics['Lansia (41+)']++;
            }
        }
        // Address
        if (p.Alamat && p.Alamat.trim() !== '') {
            const address = p.Alamat.trim().toLowerCase();
            addressCounts[address] = (addressCounts[address] || 0) + 1;
        }
    });

    stats.addressDemographics = Object.entries(addressCounts)
        .sort(([,a],[,b]) => b-a)
        .slice(0, 10)
        .reduce((r, [k, v]) => {
            const capitalizedKey = k.replace(/\b\w/g, l => l.toUpperCase());
            r[capitalizedKey] = v;
            return r;
        }, {});


    const sortedDailyTrend = Object.keys(stats.dailyTrend).sort().reduce((obj, key) => { obj[key] = stats.dailyTrend[key]; return obj; }, {});
    stats.dailyTrend = sortedDailyTrend;
    return { status: 'success', data: {stats, rawReservations: reservations} };
}

function createJsonResponse(data) { return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON); }
function getSheet(sheetName) { const ss = SpreadsheetApp.getActiveSpreadsheet(); let sheet = ss.getSheetByName(sheetName); if (!sheet) { sheet = ss.insertSheet(sheetName); const headers = getHeadersForSheet(sheetName); if (headers) { sheet.appendRow(headers); sheet.setFrozenRows(1); sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold'); } } return sheet; }
function getHeadersForSheet(sheetName) { const headerMap = { [PATIENT_SHEET_NAME]: ['RME', 'Nama_Pasien', 'Nama_Pemesan', 'No_HP', 'Instagram', 'Alamat', 'Tanggal_Lahir', 'Tanggal_Registrasi', 'Jenis_Kelamin'], [RESERVATION_SHEET_NAME]: ['ID_Reservasi', 'Timestamp', 'Status', 'RME', 'Nama_Pasien', 'Nama_Pemesan', 'No_HP', 'Alamat', 'Tanggal_Datang', 'Jam_Datang', 'Items', 'Keluhan', 'Catatan', 'Terapis', 'Data_Pemeriksaan'], [TREATMENT_SHEET_NAME]: ['ID_Treatment', 'Kategori', 'Nama', 'Deskripsi'], [PRODUCT_SHEET_NAME]: ['ID_Produk', 'Nama', 'Deskripsi'], [LOG_SHEET_NAME]: ['Timestamp', 'Function', 'Message', 'ErrorStack'] }; return headerMap[sheetName]; }
function sheetToJSON(sheet) { const data = sheet.getDataRange().getValues(); if (data.length < 2) return []; const headers = data.shift(); return data.map(row => { let obj = {}; headers.forEach((col, index) => { obj[col] = row[index]; }); return obj; }); }
function generateRME(sheet) { const lastRow = sheet.getLastRow(); if (lastRow < 2) return 'NBLH-001'; try { const lastRME = sheet.getRange(lastRow, 1).getValue(); const lastNumber = parseInt(lastRME.split('-')[1]); const newNumber = (lastNumber + 1).toString().padStart(3, '0'); return `NBLH-${newNumber}`; } catch(e) { return `NBLH-${(lastRow).toString().padStart(3, '0')}`; } }
function logError(functionName, error) { try { getSheet(LOG_SHEET_NAME).appendRow([new Date(), functionName, error.message, error.stack]); } catch(e) { Logger.log("Gagal menulis log: " + e.message); } }
function setupInitialSheet() { getSheet(PATIENT_SHEET_NAME); getSheet(RESERVATION_SHEET_NAME); getSheet(TREATMENT_SHEET_NAME); getSheet(PRODUCT_SHEET_NAME); getSheet(LOG_SHEET_NAME); SpreadsheetApp.flush(); Logger.log('Setup completed.'); }

