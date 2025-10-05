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

// --- KONSTANTA GLOBAL ---
// Mendefinisikan konstanta untuk nama-nama sheet agar mudah dikelola dan tidak ada salah ketik.
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId(); // (Tidak digunakan secara eksplisit, tapi baik untuk ada)
const PATIENT_SHEET_NAME = 'Pasien';
const RESERVATION_SHEET_NAME = 'Reservasi';
const TREATMENT_SHEET_NAME = 'Treatments';
const PRODUCT_SHEET_NAME = 'Products';
const LOG_SHEET_NAME = 'Log'; // Sheet untuk mencatat error

// --- FUNGSI UTAMA WEB APP (ENTRY POINT) ---

/**
 * Fungsi ini dijalankan ketika ada request HTTP GET ke URL Web App.
 * Berfungsi sebagai router untuk mengambil data (read operations).
 * @param {Object} e - Objek event yang berisi parameter dari request.
 * @returns {ContentService.TextOutput} - Respon dalam format JSON.
 */
function doGet(e) {
  try {
    // Mengambil parameter 'action' dari URL.
    const action = e.parameter.action;
    // Mengambil dan mem-parsing 'payload' dari URL (jika ada).
    const payload = e.parameter.payload ? JSON.parse(e.parameter.payload) : {};
    let response;

    // Memilih fungsi yang akan dijalankan berdasarkan 'action'.
    switch(action) {
      case 'getReservations':   response = getReservations(); break;
      case 'getPatients':       response = getPatients(payload.query); break;
      case 'getItems':          response = getItems(); break;
      case 'getRekapData':      response = getRekapData(); break;
      case 'getPatientHistory': response = getPatientHistory(payload); break; 
      default:                  response = { status: 'error', message: 'Invalid GET action' };
    }
    // Mengembalikan hasil sebagai JSON.
    return createJsonResponse(response);
  } catch (error) {
    // Jika terjadi error, catat error dan kembalikan pesan error.
    logError('doGet', error);
    return createJsonResponse({ status: 'error', message: error.message });
  }
}

/**
 * Fungsi ini dijalankan ketika ada request HTTP POST ke URL Web App.
 * Berfungsi sebagai router untuk mengubah data (create, update, delete operations).
 * @param {Object} e - Objek event yang berisi data POST.
 * @returns {ContentService.TextOutput} - Respon dalam format JSON.
 */
function doPost(e) {
  try {
    // Mem-parsing data JSON dari body request.
    const requestData = JSON.parse(e.postData.contents);
    const action = requestData.action;
    const payload = requestData.payload;
    let response;

    // Validasi: pastikan 'action' ada.
    if (!action) return createJsonResponse({ status: 'error', message: 'Action not specified' });

    // Memilih fungsi yang akan dijalankan berdasarkan 'action'.
    switch (action) {
      case 'newReservation':      response = handleNewReservation(payload); break;
      case 'completeReservation': response = handleCompleteReservation(payload); break;
      case 'addOrUpdateItem':     response = handleAddOrUpdateItem(payload); break;
      default:                    response = { status: 'error', message: 'Invalid POST action' };
    }
    // Mengembalikan hasil sebagai JSON.
    return createJsonResponse(response);
  } catch (error) {
    // Jika terjadi error, catat error dan kembalikan pesan error.
    logError('doPost', error);
    return createJsonResponse({ status: 'error', message: error.message, stack: error.stack });
  }
}


// --- FUNGSI-FUNGSI HANDLER (LOGIKA BISNIS) ---

/**
 * Mengambil riwayat semua kunjungan seorang pasien berdasarkan RME.
 * @param {Object} payload - Objek yang berisi { rme: 'NBLH-XXX' }.
 * @returns {Object} - Objek status dan data riwayat pasien.
 */
function getPatientHistory(payload) {
  // Validasi input.
  if (!payload || !payload.rme) {
    return { status: 'error', message: 'RME tidak valid.' };
  }
  // Ambil semua data reservasi.
  const reservations = sheetToJSON(getSheet(RESERVATION_SHEET_NAME));
  // Filter reservasi berdasarkan RME yang cocok.
  const patientHistory = reservations
    .filter(res => res.RME === payload.rme)
    .sort((a, b) => new Date(b.Tanggal_Datang) - new Date(a.Tanggal_Datang)); // Urutkan dari yang terbaru.
  
  return { status: 'success', data: patientHistory };
}

/**
 * Menangani pembuatan reservasi baru.
 * Mencari/membuat data pasien, lalu menambahkan data reservasi ke sheet.
 * @param {Object} data - Data dari form reservasi di frontend.
 * @returns {Object} - Objek status dan pesan hasil operasi.
 */
function handleNewReservation(data) {
  // Cari atau buat pasien baru dan dapatkan RME-nya.
  const patientRecord = findOrCreatePatient(data);
  const rme = patientRecord.rme;
  const reservationSheet = getSheet(RESERVATION_SHEET_NAME);
  
  // Cek konflik jadwal (jika bukan janji partus 24 jam).
  if (data.visitTime !== '24_jam') {
    const reservations = sheetToJSON(reservationSheet);
    const visitDateTime = new Date(`${data.visitDate}T${data.visitTime}`);
    // Cari reservasi lain pada jam dan tanggal yang sama.
    const conflict = reservations.filter(r => {
        if (!r.Tanggal_Datang || r.Status === 'Selesai') return false; // Abaikan yang sudah selesai.
        try {
            const existingDate = new Date(r.Tanggal_Datang);
            return existingDate.getTime() === visitDateTime.getTime();
        } catch (err) { return false; }
    });
    // Jika sudah ada 2 atau lebih, tolak reservasi baru.
    if (conflict.length >= 2) {
      return { status: 'error', message: 'Slot pada jam dan tanggal tersebut sudah penuh (Maks 2 reservasi).' };
    }
  }

  // Buat ID reservasi unik.
  const reservationId = 'RES-' + Date.now();
  // Gabungkan tanggal dan waktu kunjungan.
  const visitDateTime = new Date(`${data.visitDate}T${data.visitTime}`);

  // Siapkan baris baru untuk dimasukkan ke sheet Reservasi.
  const newReservation = [
    reservationId, new Date(), 'Menunggu', rme, data.patientName, data.requesterName,
    data.phone, data.address, 
    // Pastikan tanggal valid sebelum diubah ke ISO string.
    !isNaN(visitDateTime.getTime()) ? visitDateTime.toISOString() : `${data.visitDate}T${data.visitTime}`,
    data.visitTime, JSON.stringify(data.selectedItems.map(item => item.name)), // Simpan item sebagai string JSON
    data.complaint || '', data.notes || '', '', '' // Kolom Terapis & Data Pemeriksaan dikosongkan dulu
  ];
  // Tambahkan baris baru ke sheet.
  reservationSheet.appendRow(newReservation);
  return { status: 'success', message: 'Reservasi berhasil dibuat.', data: { rme: rme } };
}

/**
 * Mencari pasien berdasarkan nama dan No HP. Jika tidak ada, buat pasien baru.
 * @param {Object} data - Data dari form reservasi.
 * @returns {Object} - Objek yang berisi data pasien dan status (apakah pasien baru atau tidak).
 */
function findOrCreatePatient(data) {
    const patientSheet = getSheet(PATIENT_SHEET_NAME);
    const patients = sheetToJSON(patientSheet);

    // Cari pasien yang ada dengan mencocokkan nama (case-insensitive) DAN nomor HP.
    let existingPatient = patients.find(p => 
        p.Nama_Pasien.toLowerCase() === data.patientName.toLowerCase() && 
        String(p.No_HP).trim() === String(data.phone).trim()
    );

    if (existingPatient) {
        // Jika pasien ditemukan, kembalikan datanya.
        return { ...existingPatient, rme: existingPatient.RME, isNew: false };
    } else {
        // Jika tidak ditemukan, buat RME baru.
        const newRME = generateRME(patientSheet);
        const dobDate = new Date(data.dob);
        // Siapkan baris baru untuk dimasukkan ke sheet Pasien.
        const newPatientRow = [ newRME, data.patientName, data.requesterName, data.phone, data.instagram || '', data.address, !isNaN(dobDate.getTime()) ? dobDate.toISOString() : data.dob, new Date(), data.gender ];
        patientSheet.appendRow(newPatientRow);
        // Kembalikan RME yang baru dibuat.
        return { rme: newRME, isNew: true };
    }
}


/**
 * Menangani penyelesaian reservasi.
 * Mengubah status menjadi 'Selesai' dan menambahkan nama terapis serta data pemeriksaan.
 * @param {Object} payload - Data dari form penyelesaian reservasi.
 * @returns {Object} - Objek status dan pesan hasil operasi.
 */
function handleCompleteReservation(payload) {
    const { reservationId, therapist } = payload;
    const reservationSheet = getSheet(RESERVATION_SHEET_NAME);
    const data = reservationSheet.getDataRange().getValues();
    const headers = data[0];
    // Cari indeks kolom dan baris berdasarkan ID Reservasi.
    const idColIndex = headers.indexOf('ID_Reservasi');
    if (idColIndex === -1) return { status: 'error', message: 'Kolom ID_Reservasi tidak ditemukan.'};
    let rowIndex = data.findIndex(row => row[idColIndex] === reservationId);
    if (rowIndex === -1) return { status: 'error', message: 'Reservasi tidak ditemukan.'};
    
    // Nomor baris di sheet adalah index + 1.
    const realRowIndex = rowIndex + 1; 
    // Gabungkan data pemeriksaan menjadi satu objek JSON.
    const examData = { suhu: payload.temp, berat: payload.weight, tinggi: payload.height, lila: payload.lila, catatan: payload.examNotes };
    
    // Update nilai sel di sheet.
    reservationSheet.getRange(realRowIndex, headers.indexOf('Status') + 1).setValue('Selesai');
    reservationSheet.getRange(realRowIndex, headers.indexOf('Terapis') + 1).setValue(therapist);
    reservationSheet.getRange(realRowIndex, headers.indexOf('Data_Pemeriksaan') + 1).setValue(JSON.stringify(examData));
    
    return { status: 'success', message: 'Reservasi telah diselesaikan.'};
}

/**
 * Menangani penambahan atau pembaruan data item (treatment/produk).
 * @param {Object} payload - Data dari form item di halaman manajemen.
 * @returns {Object} - Objek status dan pesan hasil operasi.
 */
function handleAddOrUpdateItem(payload) {
    const { type, id, name, category, description } = payload;
    const sheetName = type === 'treatment' ? TREATMENT_SHEET_NAME : PRODUCT_SHEET_NAME;
    const sheet = getSheet(sheetName);
    
    // Jika ada ID, berarti ini adalah operasi update.
    if (id) {
        const data = sheet.getDataRange().getValues();
        const rowIndex = data.findIndex(row => row[0].toString() === id.toString());
        if (rowIndex !== -1) {
            const realRowIndex = rowIndex + 1;
            const rowData = type === 'treatment' ? [id, category, name, description] : [id, name, description];
            // Update satu baris penuh.
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


// --- FUNGSI-FUNGSI PENGAMBILAN DATA (GETTERS) ---

// Mengambil semua data dari sheet Reservasi.
function getReservations() { 
  return { status: 'success', data: sheetToJSON(getSheet(RESERVATION_SHEET_NAME)) }; 
}

// Mengambil semua data treatment dan produk.
function getItems() { 
  return { status: 'success', data: { 
    treatments: sheetToJSON(getSheet(TREATMENT_SHEET_NAME)), 
    products: sheetToJSON(getSheet(PRODUCT_SHEET_NAME)) 
  }}; 
}

// Mengambil data pasien, bisa semua atau difilter berdasarkan query.
function getPatients(query) {
    const patients = sheetToJSON(getSheet(PATIENT_SHEET_NAME));
    if (!query) return { status: 'success', data: patients };
    const lowerCaseQuery = query.toLowerCase();
    const filtered = patients.filter(p => 
      (p.RME && p.RME.toLowerCase().includes(lowerCaseQuery)) || 
      (p.Nama_Pasien && p.Nama_Pasien.toLowerCase().includes(lowerCaseQuery))
    );
    return { status: 'success', data: filtered };
}

/**
 * Mengambil dan mengolah semua data yang diperlukan untuk halaman Laporan/Rekapitulasi.
 * @returns {Object} - Objek status dan data statistik yang sudah diolah.
 */
function getRekapData() {
    // Ambil semua data mentah yang diperlukan.
    const reservations = sheetToJSON(getSheet(RESERVATION_SHEET_NAME));
    const patients = sheetToJSON(getSheet(PATIENT_SHEET_NAME));
    const treatments = sheetToJSON(getSheet(TREATMENT_SHEET_NAME));
    
    // Buat "peta" untuk mempermudah pencarian data relasional.
    const treatmentCategoryMap = treatments.reduce((map, item) => { map[item.Nama] = item.Kategori; return map; }, {});
    const patientGenderMap = patients.reduce((map, p) => { map[p.RME] = p.Jenis_Kelamin; return map; }, {});
    
    // Siapkan objek 'stats' untuk menampung hasil perhitungan.
    const stats = {
        categoryCounts: {}, treatmentNameCounts: {}, genderCounts: {'Wanita': 0, 'Pria': 0},
        dayCounts: {'Minggu': 0, 'Senin': 0, 'Selasa': 0, 'Rabu': 0, 'Kamis': 0, 'Jumat': 0, 'Sabtu': 0},
        peakHourCounts: {}, monthCounts: {}, therapistCounts: {}, dailyTrend: {}, calendarData: {},
        ageDemographics: { 'Bayi (0-1)': 0, 'Balita (2-5)': 0, 'Anak (6-12)': 0, 'Remaja (13-18)': 0, 'Dewasa (19-40)': 0, 'Lansia (41+)': 0 },
        addressDemographics: {}
    };
    
    // Inisialisasi data.
    const days = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
    for(let i=0; i<24; i++) { stats.peakHourCounts[i.toString().padStart(2,'0') + ':00'] = 0; }
    
    // Iterasi melalui setiap reservasi untuk mengumpulkan statistik.
    reservations.forEach(res => {
        if (!res.Tanggal_Datang) return;
        const visitDate = new Date(res.Tanggal_Datang);
        if (isNaN(visitDate.getTime())) return;

        const dateKey = visitDate.toISOString().split('T')[0];
        stats.calendarData[dateKey] = (stats.calendarData[dateKey] || 0) + 1;
        
        // Hitung item dan kategori.
        try {
            JSON.parse(res.Items || '[]').forEach(itemName => {
                stats.treatmentNameCounts[itemName] = (stats.treatmentNameCounts[itemName] || 0) + 1;
                const category = treatmentCategoryMap[itemName];
                if (category) stats.categoryCounts[category] = (stats.categoryCounts[category] || 0) + 1;
            });
        } catch(e) {}
        
        // Hitung statistik lainnya.
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

    // Proses demografi pasien (usia dan alamat).
    const addressCounts = {};
    patients.forEach(p => {
        // Usia
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
        // Alamat
        if (p.Alamat && p.Alamat.trim() !== '') {
            const address = p.Alamat.trim().toLowerCase();
            addressCounts[address] = (addressCounts[address] || 0) + 1;
        }
    });

    // Ambil 10 alamat teratas.
    stats.addressDemographics = Object.entries(addressCounts)
        .sort(([,a],[,b]) => b-a)
        .slice(0, 10)
        .reduce((r, [k, v]) => {
            const capitalizedKey = k.replace(/\b\w/g, l => l.toUpperCase());
            r[capitalizedKey] = v;
            return r;
        }, {});


    // Urutkan data tren harian berdasarkan tanggal.
    const sortedDailyTrend = Object.keys(stats.dailyTrend).sort().reduce((obj, key) => { obj[key] = stats.dailyTrend[key]; return obj; }, {});
    stats.dailyTrend = sortedDailyTrend;
    
    // Kembalikan data statistik dan data reservasi mentah.
    return { status: 'success', data: {stats, rawReservations: reservations} };
}


// --- FUNGSI-FUNGSI UTILITAS (HELPERS) ---

// Mengubah objek JavaScript menjadi format JSON yang dapat dikirim sebagai respon.
function createJsonResponse(data) { 
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON); 
}

// Mendapatkan objek Sheet berdasarkan nama. Jika tidak ada, buat sheet baru beserta headernya.
function getSheet(sheetName) { 
  const ss = SpreadsheetApp.getActiveSpreadsheet(); 
  let sheet = ss.getSheetByName(sheetName); 
  if (!sheet) { 
    sheet = ss.insertSheet(sheetName); 
    const headers = getHeadersForSheet(sheetName); 
    if (headers) { 
      sheet.appendRow(headers); 
      sheet.setFrozenRows(1); 
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold'); 
    } 
  } 
  return sheet; 
}

// Mendefinisikan header untuk setiap sheet.
function getHeadersForSheet(sheetName) { 
  const headerMap = { 
    [PATIENT_SHEET_NAME]: ['RME', 'Nama_Pasien', 'Nama_Pemesan', 'No_HP', 'Instagram', 'Alamat', 'Tanggal_Lahir', 'Tanggal_Registrasi', 'Jenis_Kelamin'], 
    [RESERVATION_SHEET_NAME]: ['ID_Reservasi', 'Timestamp', 'Status', 'RME', 'Nama_Pasien', 'Nama_Pemesan', 'No_HP', 'Alamat', 'Tanggal_Datang', 'Jam_Datang', 'Items', 'Keluhan', 'Catatan', 'Terapis', 'Data_Pemeriksaan'], 
    [TREATMENT_SHEET_NAME]: ['ID_Treatment', 'Kategori', 'Nama', 'Deskripsi'], 
    [PRODUCT_SHEET_NAME]: ['ID_Produk', 'Nama', 'Deskripsi'], 
    [LOG_SHEET_NAME]: ['Timestamp', 'Function', 'Message', 'ErrorStack'] 
  }; 
  return headerMap[sheetName]; 
}

// Mengubah data dari sheet (array 2D) menjadi array of objects (JSON).
function sheetToJSON(sheet) { 
  const data = sheet.getDataRange().getValues(); 
  if (data.length < 2) return []; 
  const headers = data.shift(); 
  return data.map(row => { 
    let obj = {}; 
    headers.forEach((col, index) => { 
      obj[col] = row[index]; 
    }); 
    return obj; 
  }); 
}

// Membuat nomor Rekam Medis Elektronik (RME) baru secara otomatis.
function generateRME(sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return 'NBLH-001'; // Jika belum ada data sama sekali.
    try {
        const lastRME = sheet.getRange(lastRow, 1).getValue();
        const lastNumber = parseInt(lastRME.split('-')[1]);
        const newNumber = (lastNumber + 1).toString().padStart(3, '0'); // Tambah 1 dan format jadi 3 digit.
        return `NBLH-${newNumber}`;
    } catch(e) {
        // Fallback jika format RME terakhir tidak sesuai.
        return `NBLH-${(lastRow).toString().padStart(3, '0')}`;
    }
}

// Mencatat detail error ke dalam sheet Log.
function logError(functionName, error) {
    try {
        getSheet(LOG_SHEET_NAME).appendRow([new Date(), functionName, error.message, error.stack]);
    } catch(e) {
        // Jika logging ke sheet gagal, log ke logger bawaan Apps Script.
        Logger.log("Gagal menulis log: " + e.message);
    }
}

/**
 * Fungsi setup awal. Dijalankan manual sekali saja dari editor Apps Script.
 * Fungsinya untuk membuat semua sheet yang diperlukan jika belum ada.
 */
function setupInitialSheet() {
    getSheet(PATIENT_SHEET_NAME);
    getSheet(RESERVATION_SHEET_NAME);
    getSheet(TREATMENT_SHEET_NAME);
    getSheet(PRODUCT_SHEET_NAME);
    getSheet(LOG_SHEET_NAME);
    SpreadsheetApp.flush(); // Memastikan semua perubahan disimpan.
    Logger.log('Setup completed.');
}
