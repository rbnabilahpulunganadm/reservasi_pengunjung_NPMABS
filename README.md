# 🏥 Aplikasi Reservasi & Manajemen Klinik Nabilah Pulungan

## Deskripsi Proyek

Proyek ini adalah sistem aplikasi web yang dikembangkan menggunakan **Google Apps Script (GAS)** sebagai *backend* dan **Google Sheets** sebagai *database* utama. Tujuannya adalah untuk memfasilitasi dan mengotomatisasi proses reservasi, manajemen data pasien, *treatment*, produk, dan terapis di **Klinik Nabilah Pulungan**.

Antarmuka pengguna (UI) dirancang menggunakan teknologi web modern untuk pengalaman pengguna yang responsif dan interaktif.

## Fitur Utama

Aplikasi ini mencakup fungsionalitas menyeluruh, antara lain:

### ⚙️ Backend (Google Apps Script)

  * **Manajemen Data Master**: Mengelola data Pasien, Reservasi, Treatments, Products, dan Terapis, yang semuanya disimpan di Google Sheets.
  * **Pembuatan ID Otomatis**: Menghasilkan ID Rekam Medis Elektronik (RME) baru dengan format unik `NBLH-xxx`.
  * **Pengecekan Duplikat**: Logika untuk mencegah entri data ganda (misalnya, data pasien).
  * **Pembuatan Laporan PDF**: Otomatisasi pembuatan dokumen status atau laporan menggunakan template **Google Slides** tertentu (`TEMPLATE_ID`).
  * **Sistem Logging**: Mencatat semua *error* yang terjadi di backend ke dalam sheet khusus (`Log`) untuk memudahkan *debugging*.
  * **API Service**: Menyediakan *endpoint* melalui fungsi `doGet` dan `doPost` untuk komunikasi dengan *frontend*.

### 💻 Frontend (index.html)

  * **Desain Responsif**: Dibangun dengan **Tailwind CSS** agar dapat diakses dengan baik di berbagai ukuran perangkat.
  * **Antarmuka Interaktif**: Menggunakan **JavaScript Vanilla** untuk mengelola navigasi tab, animasi pembuka (*opening animation*), dan interaksi pengguna.
  * **Visualisasi Data**: Mengintegrasikan **Chart.js** untuk menampilkan grafik dan *dashboard* data secara visual.
  * **Tema**: Mendukung *dark mode* dan *light mode* yang dapat diatur oleh pengguna.
  * **Notifikasi**: Menyediakan notifikasi *toast* (pop-up) untuk umpan balik operasi (sukses/gagal).
  * **Download PDF Klien**: Fungsi untuk mengunduh file PDF yang dikirimkan dari backend (*base64 string*).

-----

## Teknologi yang Digunakan

| Kategori | Teknologi | Deskripsi |
| :--- | :--- | :--- |
| **Backend & Hosting** | **Google Apps Script (GAS)** | Logika bisnis dan layanan web. |
| **Database** | **Google Sheets** | Penyimpanan data relasional dan non-relasional. |
| **Frontend** | **HTML5, JavaScript** | Struktur dan interaktivitas. |
| **Styling** | **Tailwind CSS** | Kerangka kerja CSS *utility-first*. |
| **Grafik** | **Chart.js** | Pustaka untuk visualisasi data. |
| **Ikon** | **Font Awesome** | Koleksi ikon. |

-----

## Struktur File

```
.
├── googleappscript.js  # Backend: Logika utama, fungsi CRUD, logging, dan Web App Service.
└── index.html          # Frontend: Antarmuka pengguna (UI), HTML, styling (Tailwind), dan JavaScript sisi klien.
```

-----

## Cuplikan

<img width="521" height="650" alt="image" src="https://github.com/user-attachments/assets/f5eedb44-0f0e-4e89-8ca5-119919bf5445" />
<img width="513" height="648" alt="image" src="https://github.com/user-attachments/assets/db3be0ae-681c-484b-9a25-6101c5ad1136" />
<img width="519" height="651" alt="image" src="https://github.com/user-attachments/assets/5bfd70cd-3bed-4dad-8f0c-946c05b86c52" />
<img width="517" height="640" alt="image" src="https://github.com/user-attachments/assets/ef6cc54d-18be-4900-9eea-17da2e38d516" />

