// =================================================================
//      BACKEND - KLINIK NABILAH PULUNGAN (Google Apps Script)
// =================================================================
// Deskripsi:
// Skrip ini berfungsi sebagai backend untuk aplikasi web reservasi.
// Fitur termasuk manajemen pasien, reservasi, treatment, produk, dan terapis,
// serta fungsionalitas notifikasi, pengecekan data duplikat, dan pembuatan laporan PDF.
// =================================================================

// --- 1. KONSTANTA GLOBAL ---
// Konstanta digunakan untuk menyimpan nilai yang tidak akan berubah,
// sehingga kode lebih mudah dibaca dan dikelola.

// Mendefinisikan nama sheet untuk data Pasien.
const PATIENT_SHEET_NAME = 'Pasien';
// Mendefinisikan nama sheet untuk data Reservasi.
const RESERVATION_SHEET_NAME = 'Reservasi';
// Mendefinisikan nama sheet untuk data Treatments.
const TREATMENT_SHEET_NAME = 'Treatments';
// Mendefinisikan nama sheet untuk data Produk.
const PRODUCT_SHEET_NAME = 'Products';
// Mendefinisikan nama sheet untuk data Terapis.
const THERAPIST_SHEET_NAME = 'Terapis';
// Mendefinisikan nama sheet untuk mencatat log error.
const LOG_SHEET_NAME = 'Log';
// Mendefinisikan ID template Google Slide yang akan digunakan untuk membuat PDF status pasien.
const TEMPLATE_ID = '1a9EMmne_y3pDUu5yoq7a9L0zVcgV2h-pfoyGRbooWak';

// --- 2. FUNGSI UTAMA WEB APP (ENTRY POINT) ---
// Fungsi ini adalah titik masuk utama ketika aplikasi web diakses.

/**
 * Menangani permintaan HTTP GET dari frontend.
 * @param {object} e - Objek event yang berisi parameter dari URL.
 * @returns {ContentService.TextOutput} - Respon dalam format JSON.
 */
function doGet(e) {
  // Blok try-catch digunakan untuk menangani potensi error yang terjadi selama eksekusi.
  try {
    // Mengambil parameter 'action' dari URL untuk menentukan operasi yang diminta.
    const action = e.parameter.action;
    // Mengambil 'payload' dari URL (jika ada), lalu mengubahnya dari string JSON menjadi objek JavaScript.
    const payload = e.parameter.payload ? JSON.parse(e.parameter.payload) : {};
    // Mendeklarasikan variabel untuk menyimpan hasil respon.
    let response;

    // 'switch' digunakan untuk menjalankan kode yang berbeda berdasarkan nilai dari 'action'.
    switch (action) {
      // Jika action adalah 'getReservationsAndNotifications', panggil fungsi terkait.
      case 'getReservationsAndNotifications':
        response = getReservationsAndNotifications();
        break;
      // Jika action adalah 'getPatients', panggil fungsi terkait dengan payload query.
      case 'getPatients':
        response = getPatients(payload.query);
        break;
      // Jika action adalah 'getItems', panggil fungsi untuk mengambil treatment dan produk.
      case 'getItems':
        response = getItems();
        break;
      // Jika action adalah 'getRekapData', panggil fungsi untuk mengambil data rekapitulasi.
      case 'getRekapData':
        response = getRekapData();
        break;
      // Jika action adalah 'getPatientHistory', panggil fungsi riwayat pasien.
      case 'getPatientHistory':
        response = getPatientHistory(payload);
        break;
      // Jika action adalah 'generatePdf', panggil fungsi untuk membuat PDF.
      case 'generatePdf':
        response = generatePatientStatusPdf(payload);
        break;
      // Jika action adalah 'getTherapists', panggil fungsi untuk mengambil data terapis.
      case 'getTherapists':
        response = getTherapists();
        break;
      // Jika action adalah 'checkExistingPatient', panggil fungsi untuk memeriksa duplikasi pasien.
      case 'checkExistingPatient':
        response = checkExistingPatient(payload);
        break;
      // 'default' akan dijalankan jika tidak ada 'action' yang cocok.
      default:
        response = {
          status: 'error',
          message: 'Invalid GET action'
        };
    }
    // Mengembalikan respon dalam format JSON ke frontend.
    return createJsonResponse(response);
  } catch (error) {
    // Jika terjadi error, catat error tersebut.
    logError('doGet', error);
    // Kirim respon error ke frontend.
    return createJsonResponse({
      status: 'error',
      message: error.message
    });
  }
}

/**
 * Menangani permintaan HTTP POST dari frontend.
 * @param {object} e - Objek event yang berisi data POST.
 * @returns {ContentService.TextOutput} - Respon dalam format JSON.
 */
function doPost(e) {
  // Blok try-catch untuk menangani potensi error.
  try {
    // Mengambil data yang dikirim dari frontend dan mengubahnya menjadi objek JavaScript.
    const requestData = JSON.parse(e.postData.contents);
    // Mengambil 'action' dari data yang dikirim.
    const action = requestData.action;
    // Mengambil 'payload' (data utama) dari data yang dikirim.
    const payload = requestData.payload;
    // Mendeklarasikan variabel untuk menyimpan respon.
    let response;

    // Memeriksa apakah 'action' telah ditentukan. Jika tidak, kirim error.
    if (!action) return createJsonResponse({
      status: 'error',
      message: 'Action not specified'
    });

    // 'switch' untuk memilih operasi berdasarkan 'action'.
    switch (action) {
      // Jika action adalah 'newReservation', panggil fungsi untuk menangani reservasi baru.
      case 'newReservation':
        response = handleNewReservation(payload);
        break;
      // Jika action adalah 'completeReservation', panggil fungsi untuk menyelesaikan reservasi.
      case 'completeReservation':
        response = handleCompleteReservation(payload);
        break;
      // Jika action adalah 'addOrUpdateItem', panggil fungsi untuk menambah/memperbarui item.
      case 'addOrUpdateItem':
        response = handleAddOrUpdateItem(payload);
        break;
      // Jika action adalah 'updatePatient', panggil fungsi untuk memperbarui data pasien.
      case 'updatePatient':
        response = handleUpdatePatient(payload);
        break;
      // Jika action adalah 'addOrUpdateTherapist', panggil fungsi untuk menambah/memperbarui terapis.
      case 'addOrUpdateTherapist':
        response = handleAddOrUpdateTherapist(payload);
        break;
      // Jika action adalah 'markReservationsAsSeen', panggil fungsi untuk menandai reservasi telah dilihat.
      case 'markReservationsAsSeen':
        response = markReservationsAsSeen(payload);
        break;
      // 'default' untuk 'action' yang tidak valid.
      default:
        response = {
          status: 'error',
          message: 'Invalid POST action'
        };
    }
    // Mengembalikan respon dalam format JSON.
    return createJsonResponse(response);
  } catch (error) {
    // Mencatat error yang terjadi.
    logError('doPost', error);
    // Mengirim respon error ke frontend, termasuk stack trace untuk debugging.
    return createJsonResponse({
      status: 'error',
      message: error.message,
      stack: error.stack
    });
  }
}


// --- 3. FUNGSI-FUNGSI HANDLER (LOGIKA BISNIS) ---

/**
 * Memeriksa apakah ada pasien yang sudah ada dengan nama yang mirip.
 * @param {object} payload - Berisi `patientName` dan `requesterName`.
 * @returns {object} - Objek status dan daftar pasien yang cocok.
 */
function checkExistingPatient(payload) {
  // Mengekstrak nama pasien dan nama pemesan dari payload.
  const {
    patientName,
    requesterName
  } = payload;
  // Jika kedua nama kosong, kembalikan array kosong karena tidak ada yang bisa dicari.
  if (!patientName && !requesterName) {
    return {
      status: 'success',
      data: []
    };
  }

  // Mengambil sheet data pasien.
  const patientSheet = getSheet(PATIENT_SHEET_NAME);
  // Mengubah data dari sheet menjadi format JSON (array of objects).
  const patients = sheetToJSON(patientSheet);

  // Menggabungkan nama pasien dan pemesan, mengubah ke huruf kecil, lalu memecahnya menjadi kata-kata.
  // Menyaring kata-kata yang terlalu pendek (kurang dari 3 huruf) untuk menghindari pencocokan yang tidak relevan.
  const searchWords = `${patientName.toLowerCase()} ${requesterName.toLowerCase()}`.split(/\s+/).filter(w => w.length > 2);

  // Menyaring daftar pasien untuk menemukan yang cocok.
  const matches = patients.filter(p => {
    // Menggabungkan dan memecah nama pasien dan pemesan dari data sheet menjadi kata-kata.
    const patientWords = `${p.Nama_Pasien.toLowerCase()} ${p.Nama_Pemesan.toLowerCase()}`.split(/\s+/);
    // Memeriksa apakah setidaknya ada satu kata dari input pencarian yang ada di dalam data pasien.
    return searchWords.some(searchWord => patientWords.includes(searchWord));
  });

  // Mengembalikan hasil pencocokan.
  return {
    status: 'success',
    data: matches
  };
}

/**
 * Menangani pembuatan reservasi baru, termasuk penggunaan RME yang sudah ada.
 * @param {object} data - Data lengkap dari form reservasi.
 * @returns {object} - Objek status dan RME pasien.
 */
function handleNewReservation(data) {
  // Mendeklarasikan variabel untuk menyimpan nomor Rekam Medis Elektronik (RME).
  let rme;
  // Memeriksa apakah frontend mengirim 'existingRME'. Ini berarti pengguna memilih pasien lama.
  if (data.existingRME) {
    // Jika ya, gunakan RME yang sudah ada.
    rme = data.existingRME;
  } else {
    // Jika tidak, cari pasien berdasarkan data yang diberikan atau buat pasien baru jika tidak ditemukan.
    const patientRecord = findOrCreatePatient(data);
    // Ambil RME dari hasil pencarian atau pembuatan pasien baru.
    rme = patientRecord.rme;
  }

  // Mengambil sheet untuk data reservasi.
  const reservationSheet = getSheet(RESERVATION_SHEET_NAME);

  // Memeriksa konflik jadwal, kecuali jika waktu kunjungan adalah '24_jam' (untuk partus).
  if (data.visitTime !== '24_jam') {
    // Mengambil semua data reservasi yang ada.
    const reservations = sheetToJSON(reservationSheet);
    // Membuat objek Date dari tanggal dan waktu kunjungan yang dipilih.
    const visitDateTime = new Date(`${data.visitDate}T${data.visitTime}`);
    // Menyaring reservasi untuk menemukan yang berkonflik.
    const conflict = reservations.filter(r => {
      // Lewati reservasi yang tidak punya tanggal datang atau yang sudah selesai.
      if (!r.Tanggal_Datang || r.Status === 'Selesai') return false;
      // Bandingkan waktu kunjungan. try-catch untuk menangani format tanggal yang mungkin tidak valid.
      try {
        return new Date(r.Tanggal_Datang).getTime() === visitDateTime.getTime();
      } catch (err) {
        return false;
      }
    });
    // Jika sudah ada 2 atau lebih reservasi pada slot waktu yang sama, kembalikan pesan error.
    if (conflict.length >= 2) {
      return {
        status: 'error',
        message: 'Slot pada jam dan tanggal tersebut sudah penuh (Maks 2 reservasi).'
      };
    }
  }

  // Membuat ID unik untuk reservasi baru menggunakan timestamp.
  const reservationId = 'RES-' + Date.now();
  // Membuat objek Date untuk waktu kunjungan.
  const visitDateTime = new Date(`${data.visitDate}T${data.visitTime}`);
  // Menyiapkan data untuk baris baru di sheet reservasi.
  const newReservation = [
    reservationId, // ID_Reservasi
    new Date(), // Timestamp (waktu pembuatan)
    'Menunggu', // Status awal
    rme, // RME pasien
    data.patientName, // Nama_Pasien
    data.requesterName, // Nama_Pemesan
    data.phone, // No_HP
    data.address, // Alamat
    // Jika tanggal valid, simpan dalam format ISO. Jika tidak, simpan sebagai string asli.
    !isNaN(visitDateTime.getTime()) ? visitDateTime.toISOString() : `${data.visitDate}T${data.visitTime}`, // Tanggal_Datang
    data.visitTime, // Jam_Datang
    // Mengubah array item yang dipilih menjadi string JSON.
    JSON.stringify(data.selectedItems.map(item => item.name)), // Items
    data.complaint || '', // Keluhan (atau string kosong jika tidak ada)
    data.notes || '', // Catatan (atau string kosong jika tidak ada)
    '', // Terapis (kosong saat awal)
    '', // Data_Pemeriksaan (kosong saat awal)
    '', // Kosongkan kolom yang mungkin sudah tidak relevan
    false // Menambahkan nilai 'false' untuk kolom 'Telah_Dilihat'
  ];
  // Menambahkan baris baru ke sheet reservasi.
  reservationSheet.appendRow(newReservation);
  // Mengembalikan pesan sukses beserta RME pasien.
  return {
    status: 'success',
    message: 'Reservasi berhasil dibuat.',
    data: {
      rme: rme
    }
  };
}

/**
 * Menangani penambahan atau pembaruan data terapis.
 * @param {object} payload - Data terapis (id, name, status).
 * @returns {object} - Objek status (sukses atau error).
 */
function handleAddOrUpdateTherapist(payload) {
  // Mengekstrak id, nama, dan status dari payload.
  const {
    id,
    name,
    status
  } = payload;
  // Mengambil sheet data terapis.
  const sheet = getSheet(THERAPIST_SHEET_NAME);

  // Jika 'id' ada, berarti ini adalah operasi update.
  if (id) {
    // Mengambil semua data dari sheet.
    const data = sheet.getDataRange().getValues();
    // Mencari indeks baris dimana ID terapis cocok.
    const rowIndex = data.findIndex(row => row[0].toString() === id.toString());
    // Jika baris ditemukan (indeks bukan -1).
    if (rowIndex !== -1) {
      // Menghitung indeks baris yang sebenarnya di sheet (karena array dimulai dari 0, sheet dari 1).
      const realRowIndex = rowIndex + 1;
      // Memperbarui data pada baris tersebut (ID, Nama, Status).
      sheet.getRange(realRowIndex, 1, 1, 3).setValues([
        [id, name, status]
      ]);
      // Mengembalikan pesan sukses.
      return {
        status: 'success',
        message: 'Data terapis berhasil diperbarui.'
      };
    } else {
      // Jika ID tidak ditemukan, kembalikan pesan error.
      return {
        status: 'error',
        message: 'Terapis tidak ditemukan.'
      };
    }
  } else { // Jika 'id' tidak ada, ini adalah operasi tambah data baru.
    // Membuat ID baru untuk terapis.
    const newId = 'TRP-' + Date.now();
    // Menambahkan baris baru ke sheet dengan data terapis baru.
    sheet.appendRow([newId, name, status || 'Aktif']);
    // Mengembalikan pesan sukses.
    return {
      status: 'success',
      message: 'Terapis berhasil ditambahkan.'
    };
  }
}

/**
 * Menandai reservasi sebagai 'telah dilihat' untuk mematikan notifikasi.
 * @param {object} payload - Berisi array `reservationIds`.
 * @returns {object} - Objek status.
 */
function markReservationsAsSeen(payload) {
  // Mengekstrak array reservationIds dari payload.
  const {
    reservationIds
  } = payload;
  // Jika tidak ada ID yang diberikan, kembalikan error.
  if (!reservationIds || reservationIds.length === 0) {
    return {
      status: 'error',
      message: 'Tidak ada ID reservasi yang diberikan.'
    };
  }

  // Mengambil sheet reservasi.
  const sheet = getSheet(RESERVATION_SHEET_NAME);
  // Mengambil semua data dari sheet.
  const data = sheet.getDataRange().getValues();
  // Mengambil baris header.
  const headers = data[0];
  // Mencari indeks kolom 'ID_Reservasi'.
  const idColIndex = headers.indexOf('ID_Reservasi');
  // Mencari indeks kolom 'Telah_Dilihat'.
  const seenColIndex = headers.indexOf('Telah_Dilihat');

  // Jika salah satu kolom tidak ditemukan, kembalikan error.
  if (idColIndex === -1 || seenColIndex === -1) {
    return {
      status: 'error',
      message: 'Kolom ID_Reservasi atau Telah_Dilihat tidak ditemukan.'
    };
  }

  // Melakukan iterasi pada setiap baris data.
  data.forEach((row, index) => {
    // Lewati baris header (index 0) dan hanya proses baris yang ID-nya ada di dalam array `reservationIds`.
    if (index > 0 && reservationIds.includes(row[idColIndex])) {
      // Memeriksa apakah nilai saat ini bukan 'true' untuk menghindari penulisan yang tidak perlu.
      if (row[seenColIndex] !== true) {
        // Mengatur nilai sel di kolom 'Telah_Dilihat' menjadi 'true'.
        sheet.getRange(index + 1, seenColIndex + 1).setValue(true);
      }
    }
  });

  // Mengembalikan pesan sukses.
  return {
    status: 'success',
    message: 'Reservasi telah ditandai.'
  };
}


/**
 * Mencari pasien berdasarkan nama dan nomor HP, atau membuat data baru jika tidak ditemukan.
 * @param {object} data - Data dari form reservasi.
 * @returns {object} - Objek berisi data pasien, RME, dan status (baru/lama).
 */
function findOrCreatePatient(data) {
  // Mengambil sheet pasien.
  const patientSheet = getSheet(PATIENT_SHEET_NAME);
  // Mengubah data sheet menjadi JSON.
  const patients = sheetToJSON(patientSheet);
  // Mencari pasien yang cocok berdasarkan nama (case-insensitive) dan nomor HP.
  let existingPatient = patients.find(p =>
    p.Nama_Pasien.toLowerCase() === data.patientName.toLowerCase() &&
    String(p.No_HP).trim() === String(data.phone).trim()
  );

  // Jika pasien ditemukan.
  if (existingPatient) {
    // Kembalikan data pasien yang ada, tambahkan properti 'isNew: false'.
    return { ...existingPatient,
      rme: existingPatient.RME,
      isNew: false
    };
  } else { // Jika pasien tidak ditemukan.
    // Buat nomor RME baru.
    const newRME = generateRME(patientSheet);
    // Ubah tanggal lahir menjadi objek Date.
    const dobDate = new Date(data.dob);
    // Siapkan data untuk baris pasien baru.
    const newPatientRow = [
      newRME, // RME
      data.patientName, // Nama_Pasien
      data.requesterName, // Nama_Pemesan
      data.phone, // No_HP
      data.instagram || '', // Instagram
      data.address, // Alamat
      // Simpan tanggal lahir dalam format ISO jika valid, jika tidak, simpan sebagai string.
      !isNaN(dobDate.getTime()) ? dobDate.toISOString() : data.dob, // Tanggal_Lahir
      new Date(), // Tanggal_Registrasi
      data.gender // Jenis_Kelamin
    ];
    // Tambahkan baris baru ke sheet pasien.
    patientSheet.appendRow(newPatientRow);
    // Kembalikan RME baru dengan status 'isNew: true'.
    return {
      rme: newRME,
      isNew: true
    };
  }
}

/**
 * Menangani penyelesaian reservasi, memperbarui status dan menambahkan data pemeriksaan.
 * @param {object} payload - Data yang diperlukan untuk menyelesaikan reservasi.
 * @returns {object} - Objek status.
 */
function handleCompleteReservation(payload) {
  // Mengekstrak data yang relevan dari payload.
  const {
    reservationId,
    therapist,
    updatedItems,
    updatedComplaint
  } = payload;
  // Mengambil sheet reservasi.
  const reservationSheet = getSheet(RESERVATION_SHEET_NAME);
  // Mengambil semua data dari sheet.
  const data = reservationSheet.getDataRange().getValues();
  // Mengambil baris header.
  const headers = data[0];
  // Mencari indeks kolom 'ID_Reservasi'.
  const idColIndex = headers.indexOf('ID_Reservasi');
  // Jika kolom tidak ditemukan, kembalikan error.
  if (idColIndex === -1) return {
    status: 'error',
    message: 'Kolom ID_Reservasi tidak ditemukan.'
  };

  // Mencari indeks baris dari reservasi yang cocok.
  let rowIndex = data.findIndex(row => row[idColIndex] === reservationId);
  // Jika tidak ditemukan, kembalikan error.
  if (rowIndex === -1) return {
    status: 'error',
    message: 'Reservasi tidak ditemukan.'
  };

  // Menghitung indeks baris yang sebenarnya di sheet.
  const realRowIndex = rowIndex + 1;
  // Menyiapkan objek untuk data pemeriksaan.
  const examData = {
    suhu: payload.temp,
    berat: payload.weight,
    tinggi: payload.height,
    lila: payload.lila,
    catatan: payload.examNotes
  };

  // Memperbarui sel 'Status' menjadi 'Selesai'.
  reservationSheet.getRange(realRowIndex, headers.indexOf('Status') + 1).setValue('Selesai');
  // Memperbarui sel 'Terapis'.
  reservationSheet.getRange(realRowIndex, headers.indexOf('Terapis') + 1).setValue(therapist);
  // Memperbarui sel 'Data_Pemeriksaan' dengan data yang sudah di-stringify.
  reservationSheet.getRange(realRowIndex, headers.indexOf('Data_Pemeriksaan') + 1).setValue(JSON.stringify(examData));
  // Memperbarui sel 'Items' dengan daftar item yang diperbarui.
  reservationSheet.getRange(realRowIndex, headers.indexOf('Items') + 1).setValue(JSON.stringify(updatedItems || []));
  // Memperbarui sel 'Keluhan'.
  reservationSheet.getRange(realRowIndex, headers.indexOf('Keluhan') + 1).setValue(updatedComplaint || '');

  // Mengembalikan pesan sukses.
  return {
    status: 'success',
    message: 'Reservasi telah diselesaikan.'
  };
}

/**
 * Memperbarui data pasien di sheet Pasien dan reservasi yang masih menunggu.
 * @param {object} payload - Data pasien yang akan diperbarui.
 * @returns {object} - Objek status.
 */
function handleUpdatePatient(payload) {
  // Ekstrak data pasien dari payload.
  const {
    rme,
    patientName,
    requesterName,
    phone,
    instagram,
    address,
    dob,
    gender
  } = payload;
  // Jika RME tidak ada, operasi update tidak bisa dilakukan.
  if (!rme) return {
    status: 'error',
    message: 'RME tidak ditemukan untuk pembaruan.'
  };

  // Ambil sheet pasien.
  const patientSheet = getSheet(PATIENT_SHEET_NAME);
  // Ambil semua data dari sheet pasien.
  const data = patientSheet.getDataRange().getValues();
  // Ambil header.
  const headers = data[0];
  // Cari indeks kolom RME.
  const rmeColIndex = headers.indexOf('RME');
  if (rmeColIndex === -1) return {
    status: 'error',
    message: 'Kolom RME tidak ditemukan.'
  };

  // Cari baris pasien berdasarkan RME.
  const rowIndex = data.findIndex(row => row[rmeColIndex] === rme);
  if (rowIndex === -1) return {
    status: 'error',
    message: 'Pasien tidak ditemukan.'
  };

  // Hitung indeks baris sebenarnya di sheet.
  const realRowIndex = rowIndex + 1;
  // Konversi tanggal lahir ke objek Date.
  const dobDate = new Date(dob);

  // Update setiap sel pada baris pasien.
  patientSheet.getRange(realRowIndex, headers.indexOf('Nama_Pasien') + 1).setValue(patientName);
  patientSheet.getRange(realRowIndex, headers.indexOf('Nama_Pemesan') + 1).setValue(requesterName);
  patientSheet.getRange(realRowIndex, headers.indexOf('No_HP') + 1).setValue(phone);
  patientSheet.getRange(realRowIndex, headers.indexOf('Instagram') + 1).setValue(instagram);
  patientSheet.getRange(realRowIndex, headers.indexOf('Alamat') + 1).setValue(address);
  patientSheet.getRange(realRowIndex, headers.indexOf('Tanggal_Lahir') + 1).setValue(!isNaN(dobDate.getTime()) ? dobDate.toISOString() : dob);
  patientSheet.getRange(realRowIndex, headers.indexOf('Jenis_Kelamin') + 1).setValue(gender);

  // Ambil sheet reservasi untuk memperbarui nama pasien pada reservasi yang masih 'Menunggu'.
  const reservationSheet = getSheet(RESERVATION_SHEET_NAME);
  const resData = reservationSheet.getDataRange().getValues();
  const resHeaders = resData[0];
  // Iterasi melalui setiap baris reservasi.
  resData.forEach((row, index) => {
    // Jika RME cocok dan status 'Menunggu', perbarui nama pasien.
    if (index > 0 && row[resHeaders.indexOf('RME')] === rme && row[resHeaders.indexOf('Status')] === 'Menunggu') {
      reservationSheet.getRange(index + 1, resHeaders.indexOf('Nama_Pasien') + 1).setValue(patientName);
    }
  });

  // Kembalikan pesan sukses.
  return {
    status: 'success',
    message: 'Data pasien berhasil diperbarui.'
  };
}

/**
 * Menambahkan atau memperbarui item (treatment atau produk).
 * @param {object} payload - Data item yang akan ditambahkan/diperbarui.
 * @returns {object} - Objek status.
 */
function handleAddOrUpdateItem(payload) {
  // Ekstrak data item dari payload.
  const {
    type,
    id,
    name,
    category,
    description
  } = payload;
  // Tentukan nama sheet berdasarkan tipe item.
  const sheetName = type === 'treatment' ? TREATMENT_SHEET_NAME : PRODUCT_SHEET_NAME;
  // Ambil sheet yang sesuai.
  const sheet = getSheet(sheetName);

  // Jika ID ada, ini adalah operasi update.
  if (id) {
    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex(row => row[0].toString() === id.toString());
    if (rowIndex !== -1) {
      const realRowIndex = rowIndex + 1;
      // Siapkan data baris berdasarkan tipe item.
      const rowData = type === 'treatment' ? [id, category, name, description] : [id, name, description];
      // Update data pada baris tersebut.
      sheet.getRange(realRowIndex, 1, 1, rowData.length).setValues([rowData]);
      return {
        status: 'success',
        message: 'Data berhasil diperbarui.'
      };
    } else {
      return {
        status: 'error',
        message: 'Item tidak ditemukan.'
      };
    }
  } else { // Jika ID tidak ada, ini adalah operasi tambah baru.
    // Buat ID baru.
    const newId = (type.charAt(0).toUpperCase()) + '-' + Date.now();
    // Siapkan data baris baru.
    const newRow = type === 'treatment' ? [newId, category, name, description] : [newId, name, description];
    // Tambahkan baris baru ke sheet.
    sheet.appendRow(newRow);
    return {
      status: 'success',
      message: 'Data berhasil ditambahkan.'
    };
  }
}

/**
 * Membuat dan mengembalikan file PDF status pasien dalam format base64.
 * @param {object} payload - Berisi `reservationId`.
 * @returns {object} - Objek status berisi data PDF base64.
 */
function generatePatientStatusPdf(payload) {
  // Validasi ID template.
  if (TEMPLATE_ID === 'GANTI_DENGAN_ID_GOOGLE_SLIDE_ANDA' || !TEMPLATE_ID) {
    return {
      status: 'error',
      message: 'ID Template Google Slide belum diatur di skrip backend (Code.gs).'
    };
  }

  // Ekstrak ID reservasi.
  const {
    reservationId
  } = payload;
  if (!reservationId) return {
    status: 'error',
    message: 'ID Reservasi tidak valid.'
  };

  // Ambil semua data reservasi dan pasien.
  const reservations = sheetToJSON(getSheet(RESERVATION_SHEET_NAME));
  const patients = sheetToJSON(getSheet(PATIENT_SHEET_NAME));

  // Cari data reservasi dan pasien yang sesuai.
  const reservation = reservations.find(r => r.ID_Reservasi === reservationId);
  if (!reservation) return {
    status: 'error',
    message: 'Data reservasi tidak ditemukan.'
  };

  const patient = patients.find(p => p.RME === reservation.RME);
  if (!patient) return {
    status: 'error',
    message: 'Data pasien tidak ditemukan.'
  };

  try {
    // Buat salinan dari file template Google Slide.
    const copyFile = DriveApp.getFileById(TEMPLATE_ID).makeCopy(`Status Pasien - ${patient.Nama_Pasien} - ${new Date().getTime()}`);
    // Buka salinan file sebagai presentasi.
    const presentation = SlidesApp.openById(copyFile.getId());
    // Ambil slide pertama.
    const slide = presentation.getSlides()[0];

    // Persiapkan data tanggal.
    const dob = new Date(patient.Tanggal_Lahir);
    const visitDate = new Date(reservation.Tanggal_Datang);

    // Parsing data pemeriksaan.
    let examData = {};
    try {
      if (reservation.Data_Pemeriksaan) {
        examData = JSON.parse(reservation.Data_Pemeriksaan);
      }
    } catch (e) {
      logError('generatePdf.parseExamData', e);
    }

    // Hitung umur pasien.
    let ageString = 'Tanggal lahir invalid';
    if (!isNaN(dob.getTime())) {
      let years = visitDate.getFullYear() - dob.getFullYear();
      let months = visitDate.getMonth() - dob.getMonth();
      let days = visitDate.getDate() - dob.getDate();
      if (days < 0) {
        months--;
        days += new Date(visitDate.getFullYear(), visitDate.getMonth(), 0).getDate();
      }
      if (months < 0) {
        years--;
        months += 12;
      }
      ageString = `${years} thn, ${months} bln, ${days} hr`;
    }

    // Siapkan data placeholder untuk diganti di slide.
    const replacements = {
      '<<NAMABAYI>>': patient.Nama_Pasien || '',
      '<<TTL>>': !isNaN(dob.getTime()) ? dob.toLocaleDateString('id-ID', {
        day: '2-digit',
        month: 'long',
        year: 'numeric'
      }) : 'N/A',
      '<<UMUR>>': ageString,
      '<<JENISKELAMIN>>': patient.Jenis_Kelamin || '',
      '<<ALAMAT>>': patient.Alamat || '',
      '<<NAMAPEMESAN>>': reservation.Nama_Pemesan || '',
      '<<NOHP>>': reservation.No_HP || '',
      '<<INSTAGRAM>>': patient.Instagram || '',
      '<<TGL>>': !isNaN(visitDate.getTime()) ? visitDate.toLocaleDateString('id-ID', {
        weekday: 'long',
        day: '2-digit',
        month: 'long',
        year: 'numeric'
      }) : 'N/A',
      '<<KELUHAN>>': reservation.Keluhan || '',
      '<<TREATMENT>>': JSON.parse(reservation.Items || '[]').join(', '),
      '<<RME>>': patient.RME || '',
      '<<TIMESTAMP>>': new Date(reservation.Timestamp).toLocaleString('id-ID', {
        dateStyle: 'long',
        timeStyle: 'short'
      }) + ' WIB',
      '<<suhu>>': examData.suhu ? `${examData.suhu} °C` : 'N/A',
      '<<berat_badan>>': examData.berat ? `${examData.berat} kg` : 'N/A',
      '<<tinggi_badan>>': examData.tinggi ? `${examData.tinggi} cm` : 'N/A',
      '<<lila>>': examData.lila ? `${examData.lila} cm` : 'N/A',
      '<<TERAPIS>>': reservation.Terapis || 'N/A'
    };

    // Lakukan penggantian teks di slide.
    for (const placeholder in replacements) {
      slide.replaceAllText(placeholder, replacements[placeholder]);
    }

    // Simpan dan tutup presentasi.
    presentation.saveAndClose();

    // Dapatkan file sebagai PDF blob.
    const pdfBlob = copyFile.getAs('application/pdf');
    // Encode PDF ke format base64.
    const base64Pdf = Utilities.base64Encode(pdfBlob.getBytes());
    // Buat nama file yang unik.
    const fileName = `StatusReservasi-${patient.Nama_Pasien.replace(/ /g, '_')}-${patient.RME}.pdf`;

    // Hapus file salinan Google Slide.
    copyFile.setTrashed(true);

    // Kembalikan data PDF base64.
    return {
      status: 'success',
      data: {
        base64: base64Pdf,
        fileName: fileName
      }
    };
  } catch (error) {
    logError('generatePatientStatusPdf', error);
    return {
      status: 'error',
      message: 'Gagal membuat PDF: ' + error.message
    };
  }
}

// --- 4. FUNGSI-FUNGSI PENGAMBILAN DATA (GETTERS) ---

/**
 * Mengambil semua data reservasi dan menghitung jumlah notifikasi baru.
 * @returns {object} - Objek berisi data reservasi dan jumlah notifikasi.
 */
function getReservationsAndNotifications() {
  // Mengambil semua data reservasi.
  const reservations = sheetToJSON(getSheet(RESERVATION_SHEET_NAME));
  // Mendapatkan tanggal hari ini, diatur ke awal hari (jam 00:00).
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // Menghitung jumlah reservasi baru.
  const newReservationCount = reservations.filter(r => {
    // Ubah timestamp reservasi menjadi objek Date.
    const reservationDate = new Date(r.Timestamp);
    // Kriteria notifikasi: Status 'Menunggu', dibuat pada atau setelah hari ini, dan belum dilihat.
    return r.Status === 'Menunggu' && reservationDate >= today && r.Telah_Dilihat !== true;
  }).length; // .length untuk mendapatkan jumlahnya.

  // Mengembalikan data reservasi dan jumlah notifikasi baru.
  return {
    status: 'success',
    data: {
      reservations: reservations,
      newReservationCount: newReservationCount
    }
  };
}

/**
 * Mengambil daftar semua terapis yang aktif.
 * @returns {object} - Objek status dan daftar terapis aktif.
 */
function getTherapists() {
  // Mengambil semua data terapis.
  const therapists = sheetToJSON(getSheet(THERAPIST_SHEET_NAME));
  // Menyaring untuk mendapatkan terapis yang statusnya 'Aktif'.
  const activeTherapists = therapists.filter(t => t.Status === 'Aktif');
  // Mengembalikan data terapis yang aktif.
  return {
    status: 'success',
    data: activeTherapists
  };
}

/**
 * Mengambil semua data treatment dan produk.
 * @returns {object} - Objek status dan data item.
 */
function getItems() {
  // Mengembalikan objek yang berisi data treatments dan products.
  return {
    status: 'success',
    data: {
      treatments: sheetToJSON(getSheet(TREATMENT_SHEET_NAME)),
      products: sheetToJSON(getSheet(PRODUCT_SHEET_NAME))
    }
  };
}

/**
 * Mengambil data pasien, bisa dengan filter pencarian.
 * @param {string} query - Kata kunci pencarian (RME atau nama).
 * @returns {object} - Objek status dan daftar pasien.
 */
function getPatients(query) {
  // Mengambil semua data pasien.
  const patients = sheetToJSON(getSheet(PATIENT_SHEET_NAME));
  // Jika tidak ada query, kembalikan semua data pasien.
  if (!query) return {
    status: 'success',
    data: patients
  };
  // Ubah query ke huruf kecil untuk pencarian case-insensitive.
  const lowerCaseQuery = query.toLowerCase();
  // Saring pasien berdasarkan RME atau Nama Pasien.
  const filtered = patients.filter(p =>
    (p.RME && p.RME.toLowerCase().includes(lowerCaseQuery)) ||
    (p.Nama_Pasien && p.Nama_Pasien.toLowerCase().includes(lowerCaseQuery))
  );
  // Kembalikan data pasien yang telah disaring.
  return {
    status: 'success',
    data: filtered
  };
}


/**
 * Mengambil riwayat kunjungan seorang pasien berdasarkan RME.
 * @param {object} payload - Berisi `rme`.
 * @returns {object} - Objek status dan riwayat reservasi.
 */
function getPatientHistory(payload) {
  // Validasi payload.
  if (!payload || !payload.rme) return {
    status: 'error',
    message: 'RME tidak valid.'
  };
  // Ambil semua data reservasi.
  const reservations = sheetToJSON(getSheet(RESERVATION_SHEET_NAME));
  // Saring reservasi berdasarkan RME, lalu urutkan dari yang terbaru.
  const patientHistory = reservations
    .filter(res => res.RME === payload.rme)
    .sort((a, b) => new Date(b.Tanggal_Datang) - new Date(a.Tanggal_Datang));
  // Kembalikan data riwayat.
  return {
    status: 'success',
    data: patientHistory
  };
}

/**
 * Mengambil dan memproses data untuk halaman laporan/rekapitulasi.
 * @returns {object} - Objek status dan data statistik.
 */
function getRekapData() {
  // Ambil semua data yang relevan.
  const reservations = sheetToJSON(getSheet(RESERVATION_SHEET_NAME));
  const patients = sheetToJSON(getSheet(PATIENT_SHEET_NAME));
  const treatments = sheetToJSON(getSheet(TREATMENT_SHEET_NAME));

  // Buat map untuk memetakan nama treatment ke kategorinya untuk efisiensi.
  const treatmentCategoryMap = treatments.reduce((map, item) => {
    map[item.Nama] = item.Kategori;
    return map;
  }, {});

  // Inisialisasi objek untuk menyimpan statistik.
  const stats = {
    categoryCounts: {},
    treatmentNameCounts: {},
    genderCounts: {},
    dayCounts: {},
    peakHourCounts: {},
    monthCounts: {},
    therapistCounts: {},
    dailyTrend: {},
    calendarData: {},
    ageDemographics: {
      'Bayi (0-1)': 0,
      'Balita (2-5)': 0,
      'Anak (6-12)': 0,
      'Remaja (13-18)': 0,
      'Dewasa (19-40)': 0,
      'Lansia (41+)': 0
    },
    addressDemographics: {}
  };

  // Persiapan label untuk grafik.
  const days = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
  for (let i = 0; i < 24; i++) {
    stats.peakHourCounts[i.toString().padStart(2, '0') + ':00'] = 0;
  }

  // Iterasi melalui setiap reservasi untuk mengumpulkan statistik.
  reservations.forEach(res => {
    if (!res.Tanggal_Datang) return; // Lewati jika tanggal tidak ada.
    const visitDate = new Date(res.Tanggal_Datang);
    if (isNaN(visitDate.getTime())) return; // Lewati jika tanggal tidak valid.

    // Statistik tren harian dan kalender.
    const dateKey = visitDate.toISOString().split('T')[0];
    stats.dailyTrend[dateKey] = (stats.dailyTrend[dateKey] || 0) + 1;
    stats.calendarData[dateKey] = (stats.calendarData[dateKey] || 0) + 1;

    // Statistik treatment dan kategori.
    try {
      JSON.parse(res.Items || '[]').forEach(itemName => {
        stats.treatmentNameCounts[itemName] = (stats.treatmentNameCounts[itemName] || 0) + 1;
        const category = treatmentCategoryMap[itemName];
        if (category) stats.categoryCounts[category] = (stats.categoryCounts[category] || 0) + 1;
      });
    } catch (e) {}

    // Statistik gender.
    const patient = patients.find(p => p.RME === res.RME);
    if (patient && patient.Jenis_Kelamin) {
      stats.genderCounts[patient.Jenis_Kelamin] = (stats.genderCounts[patient.Jenis_Kelamin] || 0) + 1;
    }

    // Statistik hari, jam, bulan, dan terapis.
    stats.dayCounts[days[visitDate.getDay()]] = (stats.dayCounts[days[visitDate.getDay()]] || 0) + 1;
    const hourKey = visitDate.getHours().toString().padStart(2, '0') + ':00';
    if (stats.peakHourCounts.hasOwnProperty(hourKey)) stats.peakHourCounts[hourKey]++;

    const monthYearKey = `${months[visitDate.getMonth()]} ${visitDate.getFullYear()}`;
    stats.monthCounts[monthYearKey] = (stats.monthCounts[monthYearKey] || 0) + 1;

    if (res.Terapis) {
      stats.therapistCounts[res.Terapis] = (stats.therapistCounts[res.Terapis] || 0) + 1;
    }
  });

  // Iterasi melalui pasien untuk demografi usia dan alamat.
  const addressCounts = {};
  patients.forEach(p => {
    // Demografi Usia
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
    // Demografi Alamat
    if (p.Alamat && p.Alamat.trim() !== '') {
      const address = p.Alamat.trim().toLowerCase();
      addressCounts[address] = (addressCounts[address] || 0) + 1;
    }
  });
  // Ambil 10 alamat teratas.
  stats.addressDemographics = Object.entries(addressCounts).sort(([, a], [, b]) => b - a).slice(0, 10).reduce((r, [k, v]) => {
    const capitalizedKey = k.replace(/\b\w/g, l => l.toUpperCase());
    r[capitalizedKey] = v;
    return r;
  }, {});

  // Kembalikan semua statistik.
  return {
    status: 'success',
    data: {
      stats,
      rawReservations: reservations
    }
  };
}


// --- 5. FUNGSI-FUNGSI UTILITAS (HELPERS) ---

/**
 * Membuat objek respon JSON untuk dikirim ke frontend.
 * @param {object} data - Objek JavaScript yang akan diubah menjadi JSON.
 * @returns {ContentService.TextOutput} - Objek output teks dengan tipe MIME JSON.
 */
function createJsonResponse(data) {
  // Mengubah objek JavaScript menjadi string JSON.
  return ContentService.createTextOutput(JSON.stringify(data))
    // Mengatur tipe konten menjadi JSON agar browser tahu cara menanganinya.
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Mendapatkan objek sheet berdasarkan namanya. Jika sheet tidak ada, maka akan dibuat.
 * @param {string} sheetName - Nama sheet yang ingin diakses atau dibuat.
 * @returns {Sheet} - Objek sheet dari Google Sheets.
 */
function getSheet(sheetName) {
  // Mendapatkan spreadsheet yang aktif saat ini.
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Mencoba mendapatkan sheet berdasarkan nama.
  let sheet = ss.getSheetByName(sheetName);
  // Jika sheet tidak ditemukan (hasilnya null).
  if (!sheet) {
    // Buat sheet baru dengan nama yang diberikan.
    sheet = ss.insertSheet(sheetName);
    // Dapatkan header yang sesuai untuk sheet baru ini.
    const headers = getHeadersForSheet(sheetName);
    // Jika ada definisi header.
    if (headers) {
      // Tambahkan header sebagai baris pertama.
      sheet.appendRow(headers);
      // Bekukan baris pertama agar header selalu terlihat saat scrolling.
      sheet.setFrozenRows(1);
      // Jadikan teks header tebal.
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
  }
  // Kembalikan objek sheet.
  return sheet;
}

/**
 * Menyediakan daftar header kolom untuk setiap jenis sheet.
 * @param {string} sheetName - Nama sheet.
 * @returns {Array<string>|undefined} - Array berisi nama-nama kolom header.
 */
function getHeadersForSheet(sheetName) {
  // Objek ini memetakan nama sheet ke daftar headernya.
  const headerMap = {
    [PATIENT_SHEET_NAME]: ['RME', 'Nama_Pasien', 'Nama_Pemesan', 'No_HP', 'Instagram', 'Alamat', 'Tanggal_Lahir', 'Tanggal_Registrasi', 'Jenis_Kelamin'],
    [RESERVATION_SHEET_NAME]: ['ID_Reservasi', 'Timestamp', 'Status', 'RME', 'Nama_Pasien', 'Nama_Pemesan', 'No_HP', 'Alamat', 'Tanggal_Datang', 'Jam_Datang', 'Items', 'Keluhan', 'Catatan', 'Terapis', 'Data_Pemeriksaan', 'Telah_Dilihat'],
    [TREATMENT_SHEET_NAME]: ['ID_Treatment', 'Kategori', 'Nama', 'Deskripsi'],
    [PRODUCT_SHEET_NAME]: ['ID_Produk', 'Nama', 'Deskripsi'],
    [THERAPIST_SHEET_NAME]: ['ID_Terapis', 'Nama_Terapis', 'Status'],
    [LOG_SHEET_NAME]: ['Timestamp', 'Function', 'Message', 'ErrorStack']
  };
  // Mengembalikan array header yang sesuai dengan nama sheet yang diberikan.
  return headerMap[sheetName];
}

/**
 * Mengubah data dari sheet menjadi array objek JSON.
 * @param {Sheet} sheet - Objek sheet yang akan diubah.
 * @returns {Array<object>} - Array objek, di mana setiap objek merepresentasikan satu baris.
 */
function sheetToJSON(sheet) {
  // Mendapatkan semua data dari sheet sebagai array 2D.
  const data = sheet.getDataRange().getValues();
  // Jika sheet kosong atau hanya berisi header, kembalikan array kosong.
  if (data.length < 2) return [];
  // Mengambil baris pertama sebagai header, dan menghapusnya dari data utama.
  const headers = data.shift();
  // Menggunakan .map() untuk mengubah setiap baris (array) menjadi objek.
  return data.map(row => {
    // Membuat objek kosong untuk baris saat ini.
    let obj = {};
    // Melakukan iterasi pada setiap header.
    headers.forEach((col, index) => {
      // Menetapkan nilai sel (row[index]) ke properti objek (obj[col]).
      obj[col] = row[index];
    });
    // Mengembalikan objek yang sudah jadi.
    return obj;
  });
}

/**
 * Membuat nomor Rekam Medis Elektronik (RME) baru secara berurutan.
 * @param {Sheet} sheet - Sheet pasien untuk mendapatkan RME terakhir.
 * @returns {string} - Nomor RME baru, contoh: 'NBLH-002'.
 */
function generateRME(sheet) {
  // Mendapatkan nomor baris terakhir yang berisi data.
  const lastRow = sheet.getLastRow();
  // Jika sheet masih kosong (hanya header), mulai dari 001.
  if (lastRow < 2) return 'NBLH-001';
  // Blok try-catch untuk menangani jika RME terakhir formatnya tidak sesuai.
  try {
    // Mendapatkan nilai RME dari baris terakhir, kolom pertama.
    const lastRME = sheet.getRange(lastRow, 1).getValue();
    // Memecah string RME (misal: 'NBLH-001') dan mengambil bagian nomornya.
    const lastNumber = parseInt(lastRME.split('-')[1]);
    // Menambahkan 1 ke nomor terakhir, lalu memformatnya menjadi 3 digit dengan nol di depan.
    const newNumber = (lastNumber + 1).toString().padStart(3, '0');
    // Menggabungkan kembali menjadi format RME lengkap.
    return `NBLH-${newNumber}`;
  } catch (e) {
    // Jika terjadi error (misal format RME salah), buat RME baru berdasarkan jumlah baris.
    return `NBLH-${(lastRow).toString().padStart(3, '0')}`;
  }
}

/**
 * Mencatat informasi error ke dalam sheet 'Log'.
 * @param {string} functionName - Nama fungsi tempat error terjadi.
 * @param {Error} error - Objek error yang ditangkap.
 */
function logError(functionName, error) {
  // Blok try-catch untuk mencegah error saat proses logging itu sendiri.
  try {
    // Menambahkan baris baru ke sheet 'Log' dengan informasi waktu, nama fungsi, pesan error, dan stack trace.
    getSheet(LOG_SHEET_NAME).appendRow([new Date(), functionName, error.message, error.stack]);
  } catch (e) {
    // Jika logging ke sheet gagal, catat ke log standar Apps Script.
    Logger.log("Gagal menulis log: " + e.message);
  }
}

/**
 * Menjalankan setup awal untuk memastikan semua sheet yang dibutuhkan sudah ada.
 */
function setupInitialSheet() {
  // Memanggil getSheet untuk setiap sheet yang dibutuhkan.
  // Ini akan membuat sheet jika belum ada.
  getSheet(PATIENT_SHEET_NAME);
  getSheet(RESERVATION_SHEET_NAME);
  getSheet(TREATMENT_SHEET_NAME);
  getSheet(PRODUCT_SHEET_NAME);
  getSheet(THERAPIST_SHEET_NAME);
  getSheet(LOG_SHEET_NAME);
  // Memastikan semua perubahan yang tertunda ditulis ke spreadsheet.
  SpreadsheetApp.flush();
  // Mencatat pesan bahwa setup telah selesai.
  Logger.log('Setup completed.');
}