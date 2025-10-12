// =================================================================
//      BACKEND - KLINIK NABILAH PULUNGAN (Google Apps Script)
// =================================================================
// File: Code.gs (atau nama file skrip Anda di Google Apps Script)
// Deskripsi:
// Skrip ini berfungsi sebagai backend (server) untuk aplikasi web reservasi.
// Ia menangani semua logika bisnis seperti menyimpan, mengambil, dan
// memanipulasi data yang tersimpan di Google Sheets.
//
// Petunjuk Penggunaan:
// 1. Salin seluruh kode ini dan tempelkan ke editor skrip proyek Google Apps Script Anda.
// 2. Jika Anda baru memulai, jalankan fungsi `setupInitialSheet` sekali dari editor untuk membuat semua sheet yang diperlukan.
// 3. Deploy skrip sebagai "Web App". Pastikan untuk memilih versi BARU setiap kali Anda melakukan perubahan pada kode.
// =================================================================

// --- 1. KONSTANTA GLOBAL ---
// Mendefinisikan nama-nama sheet sebagai konstanta.
// Ini adalah praktik yang baik untuk menghindari kesalahan pengetikan (typo) dan memudahkan jika suatu saat nama sheet perlu diubah.
const PATIENT_SHEET_NAME = 'Pasien';
const RESERVATION_SHEET_NAME = 'Reservasi';
const TREATMENT_SHEET_NAME = 'Treatments';
const PRODUCT_SHEET_NAME = 'Products';
const LOG_SHEET_NAME = 'Log'; // Sheet untuk mencatat error

// !!! PENTING: Ganti ID di bawah ini dengan ID file Google Slide template Anda !!!
// Petunjuk: Buka file Google Slide, salin ID dari URL. Contoh: .../d/INI_ADALAH_ID_NYA/edit
const TEMPLATE_ID = '1a9EMmne_y3pDUu5yoq7a9L0zVcgV2h-pfoyGRbooWak';

// --- 2. FUNGSI UTAMA WEB APP (ENTRY POINT) ---

/**
 * Fungsi `doGet(e)` adalah fungsi khusus Google Apps Script yang akan
 * dieksekusi setiap kali ada request HTTP GET ke URL web app.
 * Fungsi ini biasanya digunakan untuk MENGAMBIL (get) data.
 * @param {object} e - Objek event yang berisi parameter dari URL request.
 * @returns {ContentService.TextOutput} - Respon dalam format JSON.
 */
function doGet(e) {
  try { // Blok `try...catch` digunakan untuk menangani potensi error selama eksekusi.
    // Mengambil parameter 'action' dari URL. Contoh: ?action=getReservations
    const action = e.parameter.action;
    // Mengambil parameter 'payload' (jika ada) dan mengubahnya dari string JSON menjadi objek JavaScript.
    const payload = e.parameter.payload ? JSON.parse(e.parameter.payload) : {};
    let response;

    // `switch` digunakan untuk memilih blok kode yang akan dijalankan berdasarkan nilai `action`.
    switch(action) {
      case 'getReservations':   response = getReservations(); break;
      case 'getPatients':       response = getPatients(payload.query); break;
      case 'getItems':          response = getItems(); break;
      case 'getRekapData':      response = getRekapData(); break;
      case 'getPatientHistory': response = getPatientHistory(payload); break; 
      case 'generatePdf':       response = generatePatientStatusPdf(payload); break; // <-- PENAMBAHAN FUNGSI BARU
      default:                  response = { status: 'error', message: 'Invalid GET action' };
    }
    // Mengembalikan hasil sebagai response JSON.
    return createJsonResponse(response);
  } catch (error) {
    // Jika terjadi error di dalam blok `try`, catat error tersebut dan kirim response error.
    logError('doGet', error);
    return createJsonResponse({ status: 'error', message: error.message });
  }
}

/**
 * Fungsi `doPost(e)` adalah fungsi khusus yang dieksekusi setiap kali
 * ada request HTTP POST ke URL web app.
 * Fungsi ini biasanya digunakan untuk MENGIRIM atau MENGUBAH data.
 * @param {object} e - Objek event yang berisi data yang dikirim di body request.
 * @returns {ContentService.TextOutput} - Respon dalam format JSON.
 */
function doPost(e) {
  try {
    // Mengambil data dari body POST request dan mengubahnya dari string JSON menjadi objek JavaScript.
    const requestData = JSON.parse(e.postData.contents);
    const action = requestData.action;
    const payload = requestData.payload;
    let response;

    // Validasi sederhana untuk memastikan `action` ada.
    if (!action) return createJsonResponse({ status: 'error', message: 'Action not specified' });

    switch (action) {
      case 'newReservation':      response = handleNewReservation(payload); break;
      case 'completeReservation': response = handleCompleteReservation(payload); break;
      case 'addOrUpdateItem':     response = handleAddOrUpdateItem(payload); break;
      case 'updatePatient':       response = handleUpdatePatient(payload); break;
      default:                    response = { status: 'error', message: 'Invalid POST action' };
    }
    return createJsonResponse(response);
  } catch (error) {
    logError('doPost', error);
    return createJsonResponse({ status: 'error', message: error.message, stack: error.stack });
  }
}


// --- 3. FUNGSI-FUNGSI HANDLER (LOGIKA BISNIS) ---
// Fungsi-fungsi ini berisi logika utama dari aplikasi.

/**
 * **[FUNGSI BARU]**
 * Membuat PDF status pasien berdasarkan template Google Slide.
 * @param {object} payload - Objek yang berisi `reservationId`.
 * @returns {object} - Objek status berisi data PDF dalam format base64.
 */
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
    const slide = presentation.getSlides()[0]; // Mengasumsikan template hanya 1 slide

    const dob = new Date(patient.Tanggal_Lahir);
    const visitDate = new Date(reservation.Tanggal_Datang);

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
      '<<TIMESTAMP>>': new Date(reservation.Timestamp).toLocaleString('id-ID', { dateStyle: 'long', timeStyle: 'short' }) + ' WIB'
    };

    for (const placeholder in replacements) {
      slide.replaceAllText(placeholder, replacements[placeholder]);
    }

    presentation.saveAndClose();
    
    const pdfBlob = copyFile.getAs('application/pdf');
    const base64Pdf = Utilities.base64Encode(pdfBlob.getBytes());
    const fileName = `StatusReservasi-${patient.Nama_Pasien.replace(/ /g, '_')}-${patient.RME}.pdf`;

    copyFile.setTrashed(true); // Menghapus file salinan sementara

    return { status: 'success', data: { base64: base64Pdf, fileName: fileName } };
  } catch (error) {
    logError('generatePatientStatusPdf', error);
    return { status: 'error', message: 'Gagal membuat PDF: ' + error.message };
  }
}

/**
 * Mengambil riwayat reservasi untuk seorang pasien berdasarkan RME.
 * @param {object} payload - Objek yang berisi `rme`.
 * @returns {object} - Objek status dan data riwayat.
 */
function getPatientHistory(payload) {
  if (!payload || !payload.rme) return { status: 'error', message: 'RME tidak valid.' };
  const reservations = sheetToJSON(getSheet(RESERVATION_SHEET_NAME));
  // Memfilter semua reservasi untuk mendapatkan yang cocok dengan RME, lalu mengurutkannya.
  const patientHistory = reservations
    .filter(res => res.RME === payload.rme)
    .sort((a, b) => new Date(b.Tanggal_Datang) - new Date(a.Tanggal_Datang));
  return { status: 'success', data: patientHistory };
}

/**
 * Menangani pembuatan reservasi baru.
 * Ini termasuk mencari atau membuat data pasien dan menambahkan data reservasi ke sheet.
 * @param {object} data - Data lengkap dari form reservasi di frontend.
 * @returns {object} - Objek status dan RME pasien.
 */
function handleNewReservation(data) {
  // Mencari pasien yang ada atau membuat yang baru jika tidak ditemukan.
  const patientRecord = findOrCreatePatient(data);
  const rme = patientRecord.rme;
  const reservationSheet = getSheet(RESERVATION_SHEET_NAME);
  
  // Cek konflik jadwal: jika ada 2 atau lebih reservasi pada jam yang sama, tolak.
  if (data.visitTime !== '24_jam') {
    const reservations = sheetToJSON(reservationSheet);
    const visitDateTime = new Date(`${data.visitDate}T${data.visitTime}`);
    const conflict = reservations.filter(r => {
        if (!r.Tanggal_Datang || r.Status === 'Selesai') return false;
        try {
            return new Date(r.Tanggal_Datang).getTime() === visitDateTime.getTime();
        } catch (err) { return false; }
    });
    if (conflict.length >= 2) {
      return { status: 'error', message: 'Slot pada jam dan tanggal tersebut sudah penuh (Maks 2 reservasi).' };
    }
  }

  // Membuat ID unik untuk reservasi.
  const reservationId = 'RES-' + Date.now();
  const visitDateTime = new Date(`${data.visitDate}T${data.visitTime}`);
  // Menyiapkan data baris baru untuk dimasukkan ke sheet Reservasi.
  const newReservation = [
    reservationId, new Date(), 'Menunggu', rme, data.patientName, data.requesterName,
    data.phone, data.address, 
    !isNaN(visitDateTime.getTime()) ? visitDateTime.toISOString() : `${data.visitDate}T${data.visitTime}`,
    data.visitTime, JSON.stringify(data.selectedItems.map(item => item.name)),
    data.complaint || '', data.notes || '', '', ''
  ];
  // Menambahkan baris baru ke akhir sheet.
  reservationSheet.appendRow(newReservation);
  return { status: 'success', message: 'Reservasi berhasil dibuat.', data: { rme: rme } };
}

/**
 * Mencari pasien berdasarkan nama dan nomor HP. Jika tidak ada, buat data pasien baru.
 * @param {object} data - Data dari form reservasi.
 * @returns {object} - Objek berisi data pasien dan status (baru/lama).
 */
function findOrCreatePatient(data) {
    const patientSheet = getSheet(PATIENT_SHEET_NAME);
    const patients = sheetToJSON(patientSheet);
    // Mencari pasien yang cocok. Mengubah ke huruf kecil untuk pencocokan case-insensitive.
    let existingPatient = patients.find(p => 
        p.Nama_Pasien.toLowerCase() === data.patientName.toLowerCase() && 
        String(p.No_HP).trim() === String(data.phone).trim()
    );

    if (existingPatient) {
        // Jika pasien ditemukan, kembalikan datanya.
        return { ...existingPatient, rme: existingPatient.RME, isNew: false };
    } else {
        // Jika tidak ditemukan, buat RME baru dan tambahkan pasien baru ke sheet.
        const newRME = generateRME(patientSheet);
        const dobDate = new Date(data.dob);
        const newPatientRow = [ newRME, data.patientName, data.requesterName, data.phone, data.instagram || '', data.address, !isNaN(dobDate.getTime()) ? dobDate.toISOString() : data.dob, new Date(), data.gender ];
        patientSheet.appendRow(newPatientRow);
        return { rme: newRME, isNew: true };
    }
}

/**
 * Menangani pembaruan data pasien dan menyinkronkan perubahan ke reservasi aktif.
 * @param {object} payload - Data baru pasien.
 * @returns {object} - Objek status.
 */
function handleUpdatePatient(payload) {
    const { rme, patientName, requesterName, phone, instagram, address, dob, gender } = payload;
    if (!rme) return { status: 'error', message: 'RME tidak ditemukan untuk pembaruan.' };

    const patientSheet = getSheet(PATIENT_SHEET_NAME);
    const data = patientSheet.getDataRange().getValues();
    const headers = data[0];
    const rmeColIndex = headers.indexOf('RME');
    if (rmeColIndex === -1) return { status: 'error', message: 'Kolom RME tidak ditemukan.'};
    
    // Mencari baris yang sesuai dengan RME.
    const rowIndex = data.findIndex(row => row[rmeColIndex] === rme);
    if (rowIndex === -1) return { status: 'error', message: 'Pasien tidak ditemukan.'};
    
    const realRowIndex = rowIndex + 1; // Index di array + 1 = nomor baris di sheet.
    const dobDate = new Date(dob);
    
    // Memperbarui setiap sel di baris tersebut.
    patientSheet.getRange(realRowIndex, headers.indexOf('Nama_Pasien') + 1).setValue(patientName);
    patientSheet.getRange(realRowIndex, headers.indexOf('Nama_Pemesan') + 1).setValue(requesterName);
    patientSheet.getRange(realRowIndex, headers.indexOf('No_HP') + 1).setValue(phone);
    patientSheet.getRange(realRowIndex, headers.indexOf('Instagram') + 1).setValue(instagram);
    patientSheet.getRange(realRowIndex, headers.indexOf('Alamat') + 1).setValue(address);
    patientSheet.getRange(realRowIndex, headers.indexOf('Tanggal_Lahir') + 1).setValue(!isNaN(dobDate.getTime()) ? dobDate.toISOString() : dob);
    patientSheet.getRange(realRowIndex, headers.indexOf('Jenis_Kelamin') + 1).setValue(gender);
    
    // Sinkronisasi nama ke reservasi yang statusnya masih 'Menunggu'.
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

/**
 * Menangani penyelesaian reservasi: mengubah status, menambahkan data pemeriksaan, dll.
 * @param {object} payload - Data dari form penyelesaian.
 * @returns {object} - Objek status.
 */
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
    
    // Memperbarui sel-sel yang relevan di baris reservasi.
    reservationSheet.getRange(realRowIndex, headers.indexOf('Status') + 1).setValue('Selesai');
    reservationSheet.getRange(realRowIndex, headers.indexOf('Terapis') + 1).setValue(therapist);
    reservationSheet.getRange(realRowIndex, headers.indexOf('Data_Pemeriksaan') + 1).setValue(JSON.stringify(examData));
    reservationSheet.getRange(realRowIndex, headers.indexOf('Items') + 1).setValue(JSON.stringify(updatedItems || []));
    reservationSheet.getRange(realRowIndex, headers.indexOf('Keluhan') + 1).setValue(updatedComplaint || '');
    
    return { status: 'success', message: 'Reservasi telah diselesaikan.'};
}

/**
 * Menangani penambahan atau pembaruan data treatment/produk.
 * @param {object} payload - Data item (nama, kategori, dll).
 * @returns {object} - Objek status.
 */
function handleAddOrUpdateItem(payload) {
    const { type, id, name, category, description } = payload;
    const sheetName = type === 'treatment' ? TREATMENT_SHEET_NAME : PRODUCT_SHEET_NAME;
    const sheet = getSheet(sheetName);
    
    if (id) { // Jika ada ID, berarti ini adalah operasi update.
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
    } else { // Jika tidak ada ID, ini adalah operasi tambah data baru.
        const newId = (type.charAt(0).toUpperCase()) + '-' + Date.now();
        const newRow = type === 'treatment' ? [newId, category, name, description] : [newId, name, description];
        sheet.appendRow(newRow);
        return { status: 'success', message: 'Data berhasil ditambahkan.' };
    }
}


// --- 4. FUNGSI-FUNGSI PENGAMBILAN DATA (GETTERS) ---
// Fungsi-fungsi ini hanya bertugas mengambil data dari sheet dan mengembalikannya.

function getReservations() { return { status: 'success', data: sheetToJSON(getSheet(RESERVATION_SHEET_NAME)) }; }
function getItems() { return { status: 'success', data: { treatments: sheetToJSON(getSheet(TREATMENT_SHEET_NAME)), products: sheetToJSON(getSheet(PRODUCT_SHEET_NAME)) }}; }
function getPatients(query) {
    const patients = sheetToJSON(getSheet(PATIENT_SHEET_NAME));
    if (!query) return { status: 'success', data: patients }; // Jika tidak ada query, kembalikan semua.
    const lowerCaseQuery = query.toLowerCase();
    // Jika ada query, filter berdasarkan RME atau Nama Pasien.
    const filtered = patients.filter(p => (p.RME && p.RME.toLowerCase().includes(lowerCaseQuery)) || (p.Nama_Pasien && p.Nama_Pasien.toLowerCase().includes(lowerCaseQuery)));
    return { status: 'success', data: filtered };
}

/**
 * Mengambil dan mengolah semua data yang diperlukan untuk halaman laporan/rekap.
 * @returns {object} - Objek status dan data yang sudah diolah untuk statistik.
 */
function getRekapData() {
    const reservations = sheetToJSON(getSheet(RESERVATION_SHEET_NAME));
    const patients = sheetToJSON(getSheet(PATIENT_SHEET_NAME));
    const treatments = sheetToJSON(getSheet(TREATMENT_SHEET_NAME));
    
    // Membuat "map" untuk pencarian cepat (lebih efisien daripada iterasi berulang).
    const treatmentCategoryMap = treatments.reduce((map, item) => { map[item.Nama] = item.Kategori; return map; }, {});
    
    // Objek untuk menampung semua hasil kalkulasi statistik.
    const stats = { categoryCounts: {}, treatmentNameCounts: {}, genderCounts: {}, dayCounts: {}, peakHourCounts: {}, monthCounts: {}, therapistCounts: {}, dailyTrend: {}, calendarData: {}, ageDemographics: { 'Bayi (0-1)': 0, 'Balita (2-5)': 0, 'Anak (6-12)': 0, 'Remaja (13-18)': 0, 'Dewasa (19-40)': 0, 'Lansia (41+)': 0 }, addressDemographics: {} };
    
    // Inisialisasi data.
    const days = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
    for(let i=0; i<24; i++) { stats.peakHourCounts[i.toString().padStart(2,'0') + ':00'] = 0; }
    
    // Proses setiap reservasi untuk mengumpulkan data statistik.
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
    
    // Proses data pasien untuk demografi usia dan alamat.
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


// --- 5. FUNGSI-FUNGSI UTILITAS (HELPERS) ---
// Fungsi-fungsi pembantu untuk tugas-tugas yang berulang.

function createJsonResponse(data) { return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON); }
function getSheet(sheetName) { const ss = SpreadsheetApp.getActiveSpreadsheet(); let sheet = ss.getSheetByName(sheetName); if (!sheet) { sheet = ss.insertSheet(sheetName); const headers = getHeadersForSheet(sheetName); if (headers) { sheet.appendRow(headers); sheet.setFrozenRows(1); sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold'); } } return sheet; }
function getHeadersForSheet(sheetName) { const headerMap = { [PATIENT_SHEET_NAME]: ['RME', 'Nama_Pasien', 'Nama_Pemesan', 'No_HP', 'Instagram', 'Alamat', 'Tanggal_Lahir', 'Tanggal_Registrasi', 'Jenis_Kelamin'], [RESERVATION_SHEET_NAME]: ['ID_Reservasi', 'Timestamp', 'Status', 'RME', 'Nama_Pasien', 'Nama_Pemesan', 'No_HP', 'Alamat', 'Tanggal_Datang', 'Jam_Datang', 'Items', 'Keluhan', 'Catatan', 'Terapis', 'Data_Pemeriksaan'], [TREATMENT_SHEET_NAME]: ['ID_Treatment', 'Kategori', 'Nama', 'Deskripsi'], [PRODUCT_SHEET_NAME]: ['ID_Produk', 'Nama', 'Deskripsi'], [LOG_SHEET_NAME]: ['Timestamp', 'Function', 'Message', 'ErrorStack'] }; return headerMap[sheetName]; }
function sheetToJSON(sheet) { const data = sheet.getDataRange().getValues(); if (data.length < 2) return []; const headers = data.shift(); return data.map(row => { let obj = {}; headers.forEach((col, index) => { obj[col] = row[index]; }); return obj; }); }
function generateRME(sheet) { const lastRow = sheet.getLastRow(); if (lastRow < 2) return 'NBLH-001'; try { const lastRME = sheet.getRange(lastRow, 1).getValue(); const lastNumber = parseInt(lastRME.split('-')[1]); const newNumber = (lastNumber + 1).toString().padStart(3, '0'); return `NBLH-${newNumber}`; } catch(e) { return `NBLH-${(lastRow).toString().padStart(3, '0')}`; } }
function logError(functionName, error) { try { getSheet(LOG_SHEET_NAME).appendRow([new Date(), functionName, error.message, error.stack]); } catch(e) { Logger.log("Gagal menulis log: " + e.message); } }
function setupInitialSheet() { getSheet(PATIENT_SHEET_NAME); getSheet(RESERVATION_SHEET_NAME); getSheet(TREATMENT_SHEET_NAME); getSheet(PRODUCT_SHEET_NAME); getSheet(LOG_SHEET_NAME); SpreadsheetApp.flush(); Logger.log('Setup completed.'); }
