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
      handleNewReservation(data);
    } else if (action === 'updateStatus') {
      handleUpdateStatus(data);
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

function handleNewReservation(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let reservationSheet = ss.getSheetByName(RESERVATION_SHEET_NAME);
  
  // Jika sheet belum ada, buat sheet baru
  if (!reservationSheet) {
    setupSheets();
    reservationSheet = ss.getSheetByName(RESERVATION_SHEET_NAME);
  }
  
  // Menambahkan data reservasi ke sheet "Reservasi"
  const reservationHeaders = ['ID', 'Timestamp', 'RME', 'Nama Pemesan', 'Telepon', 'Instagram', 'Alamat', 'Nama Pasien', 'Tgl Lahir', 'Usia', 'Gender', 'Kategori', 'Treatments', 'Ada Keluhan', 'Detail Keluhan', 'Tgl Booking', 'Jam Booking', 'Status', 'Terapis'];
  const reservationValues = [
    data.id, data.timestamp, data.rme, data.bookerName, data.phone, data.instagram, data.address,
    data.patientName, data.dob, data.age, data.gender, data.category, data.treatments,
    data.hasComplaint ? "Ya" : "Tidak", data.complaintText, data.bookingDate, data.bookingTime, data.status, data.therapist
  ];
  
  // Cek jika header sudah ada
  if (reservationSheet.getLastRow() === 0) {
    reservationSheet.appendRow(reservationHeaders);
  }
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
      data.rme, data.patientName, data.bookerName, data.phone, data.instagram, data.address, data.dob, data.gender
    ];
    customerSheet.appendRow(customerValues);
  }
}

function handleUpdateStatus(data) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(RESERVATION_SHEET_NAME);
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  // Cari baris berdasarkan ID reservasi (kolom pertama)
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] == data.id) {
      // Kolom Status ada di index 17 (R), Terapis di index 18 (S)
      sheet.getRange(i + 1, 18).setValue(data.status); // Kolom R
      sheet.getRange(i + 1, 19).setValue(data.therapist); // Kolom S
      break;
    }
  }
}

// Fungsi ini bisa dijalankan manual untuk setup awal sheet
function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // Buat sheet Reservasi jika belum ada
  if (!ss.getSheetByName(RESERVATION_SHEET_NAME)) {
    ss.insertSheet(RESERVATION_SHEET_NAME);
    const reservationSheet = ss.getSheetByName(RESERVATION_SHEET_NAME);
    const reservationHeaders = ['ID', 'Timestamp', 'RME', 'Nama Pemesan', 'Telepon', 'Instagram', 'Alamat', 'Nama Pasien', 'Tgl Lahir', 'Usia', 'Gender', 'Kategori', 'Treatments', 'Ada Keluhan', 'Detail Keluhan', 'Tgl Booking', 'Jam Booking', 'Status', 'Terapis'];
    reservationSheet.appendRow(reservationHeaders);
    reservationSheet.setFrozenRows(1);
    reservationSheet.getRange("A1:S1").setFontWeight("bold");
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