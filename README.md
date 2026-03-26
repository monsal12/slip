# Website Slip Gaji + Auto Email (MongoDB)

Aplikasi web sederhana untuk:
- Input data slip gaji karyawan
- Upload batch slip dari Excel
- Simpan data ke MongoDB
- Generate PDF slip gaji
- QR code otomatis di setiap PDF slip
- Kirim otomatis ke email karyawan yang terdaftar

## 1) Persiapan

Pastikan sudah terpasang:
- Node.js 18+
- MongoDB (lokal atau Atlas)

## 2) Setup

```bash
npm install
```

Salin file environment:

```bash
copy .env.example .env
```

Lalu isi `.env` minimal:

```env
PORT=3000
MONGODB_URI=mongodb://127.0.0.1:27017/slip_gaji
SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_SECURE=false
SMTP_USER=your_email@gmail.com
SMTP_PASS=app_password
SMTP_FROM_NAME=HRD RS Mulia Raya
SMTP_FROM_EMAIL=your_email@gmail.com
```

## 3) Menjalankan

Mode normal:

```bash
npm start
```

Mode development:

```bash
npm run dev
```

Buka: `http://localhost:3000`

### Login Default

- Username: `putri`
- Password: `putri`

Kredensial bisa diubah lewat `.env`:
- `APP_LOGIN_USER`
- `APP_LOGIN_PASS`

## 4) Cara Pakai

1. Isi data karyawan + komponen gaji.
2. Centang `Kirim otomatis ke email terdaftar`.
3. Klik `Buat Slip dan Proses Email`.
4. Sistem akan:
   - Simpan data ke MongoDB
   - Buat PDF ke folder `generated_slips/`
   - Kirim PDF ke email karyawan

## 5) Catatan SMTP Gmail

Jika memakai Gmail, gunakan App Password (bukan password akun utama).

## 6) Upload Batch Excel

1. Buka halaman utama aplikasi.
2. Klik `Download Template Excel` pada bagian `Upload Batch Excel`.
3. Isi file template dengan satu baris per karyawan.
: Gunakan sheet sesuai format:
   - `template_karyawan`
   - `template_dokter_umum`
   - `template_dokter_spesialis`
   - `template_spesialis_radiologi`
4. Upload file Excel, lalu klik `Upload Batch dan Proses`.
5. Jika centang `Kirim email ke semua baris pada file` aktif, sistem langsung kirim ke semua data valid.

Kolom template:

- `Institusi`
- `Bulan`
- `Tahun`
- `Nama Karyawan`
- `Posisi`
- `Tipe Slip` (`otomatis` / `Dokter Umum` / `Karyawan`)
- `No Rekening`
- `Email Karyawan`
- `Gaji Pokok`
- `Gaji Jasa`
- `Gaji Jaga`
- `BPJS Ketenagakerjaan (Pendapatan)`
- `BPJS Kesehatan (Pendapatan)`
- `Bonus`
- `BPJS Ketenagakerjaan (Potongan)`
- `BPJS Kesehatan (Potongan)`
- `Potongan Lain`
- `Tampilkan Gaji Jasa` (`yes/no`)
- `Tampilkan Gaji Jaga` (`yes/no`)
- `Tampilkan BPJS Pendapatan` (`yes/no`)
- `Tampilkan Bonus` (`yes/no`)
- `Tampilkan Potongan Lain` (`yes/no`)
- `Kirim Email` (`yes`/`true`/`1` untuk kirim email per baris)

Nilai posisi yang diterima:
- `Dokter Umum`
- `Dokter Spesialis`
- `Karyawan`

Catatan:
- Saat upload batch, sistem akan membaca semua sheet template (kecuali `referensi_posisi` dan `panduan`).
