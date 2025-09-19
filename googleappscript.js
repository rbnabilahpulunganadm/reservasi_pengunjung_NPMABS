// Hak cipta (c) 2025 A. Ro'pat Tanjung - Dibuat untuk Klinik Nabilah Pulungan
// Script ini berfungsi sebagai backend untuk menerima data dari aplikasi web reservasi.

// ID Spreadsheet bisa didapatkan dari URL Google Sheet Anda
// Contoh: docs.google.com/spreadsheets/d/THIS_IS_THE_ID/edit
const SPREADSHEET_ID = "1NuDpvUSqY7WpxSHz1CDTWEhfnUYbZRO3HIJxvThj45w"; 
const RESERVATION_SHEET_NAME = "Reservasi";
const CUSTOMER_SHEET_NAME = "Pelanggan";

function doGet(e) {
  try {
    const action = e.parameter.action;
    
    if (action === 'getData') {
      return getDataFromSheets();
    } else if (action === 'searchCustomers') {
      const searchTerm = e.parameter.term;
      return searchCustomers(searchTerm);
    } else if (action === 'getQueueData') {
      return getQueueData();
    } else if (action === 'getReservationById') {
      const reservationId = e.parameter.id;
      return getReservationById(reservationId);
    } else if (action === 'checkTriggers') {
      return ContentService
        .createTextOutput(JSON.stringify(checkActiveTriggers()))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService
      .createTextOutput(JSON.stringify({ "status": "error", "message": "Action tidak valid" }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    Logger.log(error.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ "status": "error", "message": error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const action = e.parameter.action;
    const data = JSON.parse(e.parameter.data);
    
    if (action === 'newReservation') {
      return handleNewReservation(data);
    } else if (action === 'updateStatus') {
      return handleUpdateStatus(data);
    } else if (action === 'completeTreatment') {
      return handleCompleteTreatment(data);
    } else if (action === 'createReservation') {
      return createReservation(data);
    } else if (action === 'runAutocrat') {
      autoRunAutocrat();
      return ContentService
        .createTextOutput(JSON.stringify({ "status": "success", "message": "Autocrat triggered" }))
        .setMimeType(ContentService.MimeType.JSON);
    } else if (action === 'createOnChangeTrigger') {
      createOnChangeTrigger();
      return ContentService
        .createTextOutput(JSON.stringify({ "status": "success", "message": "OnChange trigger created" }))
        .setMimeType(ContentService.MimeType.JSON);
    } else if (action === 'deleteAllTriggers') {
      deleteAllTriggers();
      return ContentService
        .createTextOutput(JSON.stringify({ "status": "success", "message": "All triggers deleted" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ "status": "error", "message": "Action tidak valid" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log(error.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ "status": "error", "message": error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getDataFromSheets() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const reservationSheet = ss.getSheetByName(RESERVATION_SHEET_NAME);
    const customerSheet = ss.getSheetByName(CUSTOMER_SHEET_NAME);
    
    // Jika sheet belum ada, buat sheet baru
    if (!reservationSheet || !customerSheet) {
      setupSheets();
      return ContentService
        .createTextOutput(JSON.stringify({ reservations: [], customers: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // Ambil data dari sheet
    const reservationData = reservationSheet.getDataRange().getValues();
    const customerData = customerSheet.getDataRange().getValues();
    
    // Konversi ke format JSON
    const headersReservation = reservationData.shift();
    const headersCustomer = customerData.shift();
    
    const reservations = reservationData.map(row => {
      let obj = {};
      headersReservation.forEach((header, i) => {
        obj[header] = row[i];
      });
      return obj;
    });
    
    const customers = customerData.map(row => {
      let obj = {};
      headersCustomer.forEach((header, i) => {
        obj[header] = row[i];
      });
      return obj;
    });
    
    return ContentService
      .createTextOutput(JSON.stringify({ reservations, customers }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    Logger.log(error.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ "status": "error", "message": error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function searchCustomers(searchTerm) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const customerSheet = ss.getSheetByName(CUSTOMER_SHEET_NAME);
    
    if (!customerSheet) {
      return ContentService
        .createTextOutput(JSON.stringify([]))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    const customerData = customerSheet.getDataRange().getValues();
    const headers = customerData.shift();
    
    const results = customerData
      .filter(row => {
        // Cari di berbagai kolom: nama pasien, nama pemesan, telepon, Instagram, RME
        return (
          (row[headers.indexOf('Nama Pasien')] && row[headers.indexOf('Nama Pasien')].toString().toLowerCase().includes(searchTerm.toLowerCase())) ||
          (row[headers.indexOf('Nama Pemesan')] && row[headers.indexOf('Nama Pemesan')].toString().toLowerCase().includes(searchTerm.toLowerCase())) ||
          (row[headers.indexOf('Telepon')] && row[headers.indexOf('Telepon')].toString().includes(searchTerm)) ||
          (row[headers.indexOf('Instagram')] && row[headers.indexOf('Instagram')].toString().toLowerCase().includes(searchTerm.toLowerCase())) ||
          (row[headers.indexOf('RME')] && row[headers.indexOf('RME')].toString().toLowerCase().includes(searchTerm.toLowerCase()))
        );
      })
      .map(row => {
        return {
          rme: row[headers.indexOf('RME')],
          patientName: row[headers.indexOf('Nama Pasien')],
          bookerName: row[headers.indexOf('Nama Pemesan')],
          phone: row[headers.indexOf('Telepon')],
          instagram: row[headers.indexOf('Instagram')],
          address: row[headers.indexOf('Alamat')],
          dob: row[headers.indexOf('Tgl Lahir')],
          gender: row[headers.indexOf('Gender')]
        };
      });
    
    return ContentService
      .createTextOutput(JSON.stringify(results))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    Logger.log(error.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ "status": "error", "message": error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getQueueData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const reservationSheet = ss.getSheetByName(RESERVATION_SHEET_NAME);
    
    if (!reservationSheet) {
      return ContentService
        .createTextOutput(JSON.stringify([]))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    const reservationData = reservationSheet.getDataRange().getValues();
    const headers = reservationData.shift();
    
    // Ambil hanya reservasi yang belum selesai atau hari ini
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    const queueData = reservationData
      .filter(row => {
        const status = row[headers.indexOf('Status')];
        const bookingDate = new Date(row[headers.indexOf('Tgl Booking')]);
        bookingDate.setHours(0, 0, 0, 0);
        
        // Tampilkan reservasi yang masih dalam antrian atau sedang treatment, 
        // atau reservasi yang sudah selesai hari ini
        return (status !== 'Selesai') || (bookingDate.getTime() === today.getTime());
      })
      .map(row => {
        return {
          id: row[headers.indexOf('ID')],
          rme: row[headers.indexOf('RME')],
          patientName: row[headers.indexOf('Nama Pasien')],
          treatment: row[headers.indexOf('Treatments')],
          date: formatDateForDisplay(row[headers.indexOf('Tgl Booking')]),
          time: row[headers.indexOf('Jam Booking')],
          status: row[headers.indexOf('Status')] || 'Menunggu',
          bookerName: row[headers.indexOf('Nama Pemesan')],
          phone: row[headers.indexOf('Telepon')]
        };
      });
    
    return ContentService
      .createTextOutput(JSON.stringify(queueData))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    Logger.log(error.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ "status": "error", "message": error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function createReservation(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let reservationSheet = ss.getSheetByName(RESERVATION_SHEET_NAME);
    
    // Jika sheet belum ada, buat sheet baru
    if (!reservationSheet) {
      setupSheets();
      reservationSheet = ss.getSheetByName(RESERVATION_SHEET_NAME);
    }
    
    // Generate ID unik untuk reservasi
    const reservationId = 'RES-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000);
    
    // Menambahkan data reservasi ke sheet "Reservasi"
    const reservationHeaders = ['ID', 'Timestamp', 'RME', 'Nama Pemesan', 'Telepon', 'Instagram', 'Alamat', 'Nama Pasien', 'Tgl Lahir', 'Usia', 'Gender', 'Kategori', 'Treatments', 'Ada Keluhan', 'Detail Keluhan', 'Tgl Booking', 'Jam Booking', 'Status', 'Terapis', 'TreatmentTambahan', 'Tarif', 'Suhu', 'Berat', 'Tinggi', 'BMI', 'Alergi', 'Tekanan Darah', 'Catatan Medis'];
    
    // Cek jika header sudah ada
    if (reservationSheet.getLastRow() === 0) {
      reservationSheet.appendRow(reservationHeaders);
    }
    
    const reservationValues = [
      reservationId, 
      new Date(), 
      data.rme, 
      data.bookerName, 
      data.phone, 
      data.instagram || '', 
      data.address || '',
      data.patientName, 
      data.dob, 
      data.age, 
      data.gender, 
      data.category, 
      Array.isArray(data.treatments) ? data.treatments.join(', ') : data.treatments,
      data.hasComplaint ? "Ya" : "Tidak", 
      data.complaintText || '', 
      data.bookingDate, 
      data.bookingTime, 
      data.status || 'Dalam Antrian', 
      '', // Terapis (kosong dulu)
      '', // TreatmentTambahan (kosong dulu)
      '', // Tarif (kosong dulu)
      '', // Suhu (kosong dulu)
      '', // Berat (kosong dulu)
      '', // Tinggi (kosong dulu)
      '', // BMI (kosong dulu)
      '', // Alergi (kosong dulu)
      '', // Tekanan Darah (kosong dulu)
      ''  // Catatan Medis (kosong dulu)
    ];
    
    reservationSheet.appendRow(reservationValues);

    // Cek dan tambahkan data pelanggan baru ke sheet "Pelanggan"
    let customerSheet = ss.getSheetByName(CUSTOMER_SHEET_NAME);
    if (!customerSheet) {
      setupSheets();
      customerSheet = ss.getSheetByName(CUSTOMER_SHEET_NAME);
    }
    
    const customerHeaders = ['RME', 'Nama Pasien', 'Nama Pemesan', 'Telepon', 'Instagram', 'Alamat', 'Tgl Lahir', 'Gender'];
    
    // Cek jika header sudah ada di sheet pelanggan
    if (customerSheet.getLastRow() === 0) {
      customerSheet.appendRow(customerHeaders);
    }

    // Cari apakah RME sudah ada
    const rmeColumn = customerSheet.getRange('A:A').getValues();
    let rmeExists = false;
    for (let i = 0; i < rmeColumn.length; i++) {
      if (rmeColumn[i][0] == data.rme) {
        rmeExists = true;
        break;
      }
    }

    if (!rmeExists) {
      const customerValues = [
        data.rme, 
        data.patientName, 
        data.bookerName, 
        data.phone, 
        data.instagram || '', 
        data.address || '', 
        data.dob, 
        data.gender
      ];
      customerSheet.appendRow(customerValues);
    }
    
    // Kembalikan data reservasi yang berhasil disimpan
    return ContentService
      .createTextOutput(JSON.stringify({ 
        success: true, 
        reservationId: reservationId,
        message: "Reservasi berhasil dibuat"
      }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    Logger.log(error.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ 
        success: false, 
        message: error.toString() 
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleNewReservation(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let reservationSheet = ss.getSheetByName(RESERVATION_SHEET_NAME);
    
    // Jika sheet belum ada, buat sheet baru
    if (!reservationSheet) {
      setupSheets();
      reservationSheet = ss.getSheetByName(RESERVATION_SHEET_NAME);
    }
    
    // Menambahkan data reservasi ke sheet "Reservasi"
    const reservationHeaders = ['ID', 'Timestamp', 'RME', 'Nama Pemesan', 'Telepon', 'Instagram', 'Alamat', 'Nama Pasien', 'Tgl Lahir', 'Usia', 'Gender', 'Kategori', 'Treatments', 'Ada Keluhan', 'Detail Keluhan', 'Tgl Booking', 'Jam Booking', 'Status', 'Terapis', 'TreatmentTambahan', 'Tarif', 'Suhu', 'Berat', 'Tinggi', 'BMI', 'Alergi', 'Tekanan Darah', 'Catatan Medis'];
    
    // Cek jika header sudah ada
    if (reservationSheet.getLastRow() === 0) {
      reservationSheet.appendRow(reservationHeaders);
    }
    
    const reservationValues = [
      data.id, 
      new Date(), 
      data.rme, 
      data.bookerName, 
      data.phone, 
      data.instagram || '', 
      data.address || '',
      data.patientName, 
      data.dob, 
      data.age, 
      data.gender, 
      data.category, 
      data.treatments,
      data.hasComplaint ? "Ya" : "Tidak", 
      data.complaintText || '', 
      data.bookingDate, 
      data.bookingTime, 
      data.status || 'Dalam Antrian', 
      data.therapist || '',
      data.additionalTreatment || '',
      data.fee || '',
      // Data medis baru
      data.temperature || '',
      data.weight || '',
      data.height || '',
      data.bmi || '',
      data.allergy || '',
      data.bloodPressure || '',
      data.medicalNotes || ''
    ];
    
    reservationSheet.appendRow(reservationValues);

    // Cek dan tambahkan data pelanggan baru ke sheet "Pelanggan"
    let customerSheet = ss.getSheetByName(CUSTOMER_SHEET_NAME);
    if (!customerSheet) {
      setupSheets();
      customerSheet = ss.getSheetByName(CUSTOMER_SHEET_NAME);
    }
    
    const customerHeaders = ['RME', 'Nama Pasien', 'Nama Pemesan', 'Telepon', 'Instagram', 'Alamat', 'Tgl Lahir', 'Gender'];
    
    // Cek jika header sudah ada di sheet pelanggan
    if (customerSheet.getLastRow() === 0) {
      customerSheet.appendRow(customerHeaders);
    }

    // Cari apakah RME sudah ada
    const rmeColumn = customerSheet.getRange('A:A').getValues();
    let rmeExists = false;
    for (let i = 0; i < rmeColumn.length; i++) {
      if (rmeColumn[i][0] == data.rme) {
        rmeExists = true;
        break;
      }
    }

    if (!rmeExists) {
      const customerValues = [
        data.rme, 
        data.patientName, 
        data.bookerName, 
        data.phone, 
        data.instagram || '', 
        data.address || '', 
        data.dob, 
        data.gender
      ];
      customerSheet.appendRow(customerValues);
    }
    
    // Kembalikan data reservasi yang berhasil disimpan
    return ContentService
      .createTextOutput(JSON.stringify({ 
        success: true, 
        data: {
          id: data.id,
          rme: data.rme,
          patientName: data.patientName,
          treatment: data.treatments,
          date: data.bookingDate,
          time: data.bookingTime,
          queueNumber: generateQueueNumber(data.bookingTime)
        }
      }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    Logger.log(error.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ "status": "error", "message": error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleUpdateStatus(data) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(RESERVATION_SHEET_NAME);
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[0];

    // Cari baris berdasarkan ID reservasi
    for (let i = 1; i < values.length; i++) {
      if (values[i][headers.indexOf('ID')] == data.id) {
        // Update status
        sheet.getRange(i + 1, headers.indexOf('Status') + 1).setValue(data.status);
        
        // Update terapis jika ada
        if (data.therapist) {
          sheet.getRange(i + 1, headers.indexOf('Terapis') + 1).setValue(data.therapist);
        }
        
        // Update treatment tambahan jika ada
        if (data.additionalTreatment !== undefined) {
          sheet.getRange(i + 1, headers.indexOf('TreatmentTambahan') + 1).setValue(data.additionalTreatment);
        }
        
        // Update tarif jika ada
        if (data.fee !== undefined) {
          sheet.getRange(i + 1, headers.indexOf('Tarif') + 1).setValue(data.fee);
        }
        
        // Update data medis jika ada
        if (data.temperature !== undefined) {
          sheet.getRange(i + 1, headers.indexOf('Suhu') + 1).setValue(data.temperature);
        }
        if (data.weight !== undefined) {
          sheet.getRange(i + 1, headers.indexOf('Berat') + 1).setValue(data.weight);
        }
        if (data.height !== undefined) {
          sheet.getRange(i + 1, headers.indexOf('Tinggi') + 1).setValue(data.height);
        }
        if (data.bmi !== undefined) {
          sheet.getRange(i + 1, headers.indexOf('BMI') + 1).setValue(data.bmi);
        }
        if (data.allergy !== undefined) {
          sheet.getRange(i + 1, headers.indexOf('Alergi') + 1).setValue(data.allergy);
        }
        if (data.bloodPressure !== undefined) {
          sheet.getRange(i + 1, headers.indexOf('Tekanan Darah') + 1).setValue(data.bloodPressure);
        }
        if (data.medicalNotes !== undefined) {
          sheet.getRange(i + 1, headers.indexOf('Catatan Medis') + 1).setValue(data.medicalNotes);
        }
        
        break;
      }
    }
    
    return ContentService
      .createTextOutput(JSON.stringify({ "status": "success" }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    Logger.log(error.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ "status": "error", "message": error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleCompleteTreatment(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const reservationSheet = ss.getSheetByName(RESERVATION_SHEET_NAME);
    const dataRange = reservationSheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[0];

    // Cari baris berdasarkan ID reservasi
    for (let i = 1; i < values.length; i++) {
      if (values[i][headers.indexOf('ID')] == data.reservationId) {
        // Update status menjadi Selesai
        reservationSheet.getRange(i + 1, headers.indexOf('Status') + 1).setValue('Selesai');
        
        // Update nama terapis
        reservationSheet.getRange(i + 1, headers.indexOf('Terapis') + 1).setValue(data.therapistName);
        
        // Update treatment tambahan
        if (data.additionalTreatment) {
          reservationSheet.getRange(i + 1, headers.indexOf('TreatmentTambahan') + 1).setValue(data.additionalTreatment);
        }
        
        // Update tarif
        if (data.fee) {
          reservationSheet.getRange(i + 1, headers.indexOf('Tarif') + 1).setValue(data.fee);
        }
        
        // Update data medis
        if (data.temperature) {
          reservationSheet.getRange(i + 1, headers.indexOf('Suhu') + 1).setValue(data.temperature);
        }
        if (data.weight) {
          reservationSheet.getRange(i + 1, headers.indexOf('Berat') + 1).setValue(data.weight);
        }
        if (data.height) {
          reservationSheet.getRange(i + 1, headers.indexOf('Tinggi') + 1).setValue(data.height);
        }
        if (data.bmi) {
          reservationSheet.getRange(i + 1, headers.indexOf('BMI') + 1).setValue(data.bmi);
        }
        if (data.allergy) {
          reservationSheet.getRange(i + 1, headers.indexOf('Alergi') + 1).setValue(data.allergy);
        }
        if (data.bloodPressure) {
          reservationSheet.getRange(i + 1, headers.indexOf('Tekanan Darah') + 1).setValue(data.bloodPressure);
        }
        if (data.medicalNotes) {
          reservationSheet.getRange(i + 1, headers.indexOf('Catatan Medis') + 1).setValue(data.medicalNotes);
        }
        
        break;
      }
    }
    
    return ContentService
      .createTextOutput(JSON.stringify({ "status": "success" }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    Logger.log(error.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ "status": "error", "message": error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getReservationById(reservationId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const reservationSheet = ss.getSheetByName(RESERVATION_SHEET_NAME);
    
    if (!reservationSheet) {
      return ContentService
        .createTextOutput(JSON.stringify({}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    const reservationData = reservationSheet.getDataRange().getValues();
    const headers = reservationData.shift();
    
    // Cari reservasi berdasarkan ID
    const reservation = reservationData.find(row => row[headers.indexOf('ID')] === reservationId);
    
    if (!reservation) {
      return ContentService
        .createTextOutput(JSON.stringify({}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    const reservationObj = {};
    headers.forEach((header, i) => {
      reservationObj[header] = reservation[i];
    });
    
    return ContentService
      .createTextOutput(JSON.stringify(reservationObj))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    Logger.log(error.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ "status": "error", "message": error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Fungsi untuk generate nomor antrian
function generateQueueNumber(bookingTime) {
  const timePart = bookingTime.split(':')[0];
  const randomNum = Math.floor(Math.random() * 99) + 1;
  return `${timePart}-${String(randomNum).padStart(2, '0')}`;
}

// Fungsi untuk format tanggal untuk display
function formatDateForDisplay(dateValue) {
  if (!dateValue) return '';
  
  const date = new Date(dateValue);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

// Fungsi ini bisa dijalankan manual untuk setup awal sheet
function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // Buat sheet Reservasi jika belum ada
  if (!ss.getSheetByName(RESERVATION_SHEET_NAME)) {
    ss.insertSheet(RESERVATION_SHEET_NAME);
    const reservationSheet = ss.getSheetByName(RESERVATION_SHEET_NAME);
    const reservationHeaders = ['ID', 'Timestamp', 'RME', 'Nama Pemesan', 'Telepon', 'Instagram', 'Alamat', 'Nama Pasien', 'Tgl Lahir', 'Usia', 'Gender', 'Kategori', 'Treatments', 'Ada Keluhan', 'Detail Keluhan', 'Tgl Booking', 'Jam Booking', 'Status', 'Terapis', 'TreatmentTambahan', 'Tarif', 'Suhu', 'Berat', 'Tinggi', 'BMI', 'Alergi', 'Tekanan Darah', 'Catatan Medis'];
    reservationSheet.appendRow(reservationHeaders);
    reservationSheet.setFrozenRows(1);
    reservationSheet.getRange("A1:AA1").setFontWeight("bold");
  }

  // Buat sheet Pelanggan jika belum ada
  if (!ss.getSheetByName(CUSTOMER_SHEET_NAME)) {
    ss.insertSheet(CUSTOMER_SHEET_NAME);
    const customerSheet = ss.getSheetByName(CUSTOMER_SHEET_NAME);
    const customerHeaders = ['RME', 'Nama Pasien', 'Nama Pemesan', 'Telepon', 'Instagram', 'Alamat', 'Tgl Lahir', 'Gender'];
    customerSheet.appendRow(customerHeaders);
    customerSheet.setFrozenRows(1);
    customerSheet.getRange("A1:H1").setFontWeight("bold");
  }
}

// ==================== FUNGSI AUTOCRAT YANG DIOPTIMALKAN ====================

// Fungsi utama untuk menjalankan Autocrat
function autoRunAutocrat(e) {
  try {
    console.log("🔍 Memeriksa perubahan data...");
    
    // Mencari job Autocrat berdasarkan nama
    const jobs = AutocratApp.getJobs();
    const job = jobs.find(j => j.name === AUTOCRAT_JOB_NAME);

    if (job) {
      // Menjalankan job yang ditemukan
      job.start();
      console.log("✅ Autocrat job '" + AUTOCRAT_JOB_NAME + "' berhasil dijalankan pada: " + new Date());
    } else {
      console.error("❌ Job dengan nama '" + AUTOCRAT_JOB_NAME + "' tidak ditemukan.");
    }
  } catch (error) {
    console.error("❌ Terjadi kesalahan saat menjalankan job: ", error.toString());
  }
}

// Fungsi untuk MEMBUAT trigger onChange (jalankan sekali saja)
function createOnChangeTrigger() {
  // Hapus semua trigger yang mungkin sudah ada
  deleteAllTriggers();
  
  // Dapatkan spreadsheet
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // Buat trigger baru yang berjalan ketika ada perubahan di spreadsheet
  ScriptApp.newTrigger('autoRunAutocrat')
    .forSpreadsheet(ss)
    .onChange()
    .create();
  
  console.log("✅ Trigger onChange berhasil dibuat! Autocrat akan berjalan otomatis saat ada perubahan data.");
}

// Fungsi untuk MEMBUAT trigger onFormSubmit (jika menggunakan Google Form)
function createOnFormSubmitTrigger() {
  // Hapus semua trigger yang mungkin sudah ada
  deleteAllTriggers();
  
  // Dapatkan spreadsheet
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // Buat trigger baru yang berjalan ketika form disubmit
  ScriptApp.newTrigger('autoRunAutocrat')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();
  
  console.log("✅ Trigger onFormSubmit berhasil dibuat! Autocrat akan berjalan otomatis saat form disubmit.");
}

// Fungsi untuk MENGHAPUS semua trigger
function deleteAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === "autoRunAutocrat") {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  console.log("✅ Semua trigger autocrat dihapus.");
}

// Tambahkan fungsi ini untuk memeriksa trigger yang aktif
function checkActiveTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  console.log("Trigger yang aktif:");
  const result = [];
  triggers.forEach(trigger => {
    const triggerInfo = {
      function: trigger.getHandlerFunction(),
      type: trigger.getEventType().toString()
    };
    console.log(`- Fungsi: ${triggerInfo.function}, Tipe: ${triggerInfo.type}`);
    result.push(triggerInfo);
  });
  return result;
}

// ==================== FUNGSI TRIGGER LAMA (UNTUK KOMPATIBILITAS) ====================

// Fungsi untuk MEMBUAT trigger time-based (jalankan sekali saja)
function createTimeBasedTrigger() {
  // Hapus semua trigger yang mungkin sudah ada
  deleteAllTriggers();
  
  // Buat trigger baru yang berjalan setiap menit
  ScriptApp.newTrigger('autoRunAutocrat')
    .timeBased()
    .everyMinutes(1)
    .create();
  
  console.log("✅ Trigger time-based berhasil dibuat! Autocrat akan berjalan otomatis setiap menit.");
}

// Fungsi untuk MENGHAPUS trigger time-based
function deleteTimeBasedTrigger() {
  deleteAllTriggers();
  console.log("✅ Time-based trigger dihapus.");
}