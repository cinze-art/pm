# Pemuda Pemudi Mawe V10 — UI Dark Neon

Portal warga Pemuda Pemudi Mawe dengan Vite + Supabase.

## Perubahan utama

### UI V10
- Halaman utama mengikuti referensi dark/black + red neon Pemuda Mawe.
- Admin memakai dashboard kartu/statistik dan quick actions.
- Agenda memakai layout daftar + kalender + agenda mendatang.
- Organisasi memakai bagan kepengurusan bertingkat dengan kartu minimalis.
- Jimpitan memakai tabel rekap bulanan dengan baris PENARIK di bawah TOTAL, filter RT/program, dan jenis jimpitan.
- Alur Supabase, login, program, KK, penarikan, koreksi, laporan, Excel/PDF, dan WhatsApp tetap dipertahankan.


Sistem jimpitan sekarang tidak lagi mengunci 2× per bulan atau jenis tertentu.

### Admin mengatur
- Nama program
- Jenis jimpitan bebas
- Target nominal kumulatif per KK
- Tanggal mulai
- Tanggal selesai
- Frekuensi: mingguan, 2 mingguan, bulanan, atau custom
- Wilayah/RT yang ikut program

Saat program disimpan, sistem membuat jadwal otomatis sesuai frekuensi. Contoh 2 minggu sekali selama 10 bulan menghasilkan semua tanggal penarikan yang berada di rentang tersebut.

**RT 07 sekarang bernama HUNTAP.**

### Penarik
Admin tidak menentukan nama penarik.

Orang yang berada di lapangan memilih jadwal lalu memasukkan:
- nama penarik
- identitas tambahan opsional
- waktu mulai otomatis

Setiap transaksi KK terhubung dengan sesi penarikan tersebut.

### KK
- Selalu tersimpan walaupun nominal Rp0.
- Urutan nama otomatis A–Z.
- Nonaktifkan KK tidak menghapus riwayat transaksi.
- Import CSV/Excel tersedia.
- Import PDF sederhana dicoba melalui ekstraksi teks; untuk PDF tabel yang kompleks lebih aman gunakan Excel/CSV.
- Tambah manual tersedia.

### Transaksi
Nominal bebas pada setiap penarikan.

Contoh target Rp200.000:
- Rp50.000
- Rp40.000
- Rp5.000
- Rp50.000

Total Rp145.000, sehingga progress 72,5%.

Tidak ada kewajiban nominal tetap setiap penarikan.

### Rp0
Jika KK tidak memberikan uang, penarik memasukkan `0` atau membiarkannya kosong lalu mengonfirmasi.

Database menyimpan `0`, sedangkan tampilan menggunakan `—`.

### Koreksi
Admin dapat membuka riwayat KK dan mengoreksi nominal. Sebelum perubahan, sistem menyimpan:
- nominal lama
- nominal baru
- alasan
- Admin yang mengoreksi
- waktu koreksi

Riwayat lama tidak dihapus.

### Laporan
Filter:
- tahun
- bulan
- RT/HUNTAP
- program

Export:
- Excel transaksi detail
- PDF melalui dialog Print browser, sehingga dapat dipilih **Save as PDF**

## Instalasi

```bash
npm install
npm run dev
```

## Supabase

1. Buka Supabase SQL Editor.
2. Jalankan seluruh `supabase.sql`.
3. Buka Authentication → Users.
4. Buat akun email/password untuk Admin.
5. Isi environment:

```env
VITE_SUPABASE_URL=https://PROJECT.supabase.co
VITE_SUPABASE_ANON_KEY=YOUR_ANON_KEY
```

Jangan pernah memasukkan `service_role` key ke frontend.

## Catatan import PDF

Browser tidak mengirim file PDF ke server. Versi ini mencoba membaca teks sederhana dari file PDF yang dipilih. PDF hasil scan/foto atau tabel dengan layout kompleks sebaiknya dikonversi ke Excel/CSV terlebih dahulu agar nama KK tidak salah terbaca.

## Build

```bash
npm run build
npm run preview
```

## Vercel

- Build command: `npm run build`
- Output directory: `dist`
- Environment Variables:
  - `VITE_SUPABASE_URL`
  - `VITE_SUPABASE_ANON_KEY`
