# Kalkulasi Nilai Karyawan

Proyek ini akan melakukan kalkulasi pada file excel untuk mengetahui kinerja karyawan pada setiap bulan.

## Tahapan Setup

### 1. Clone Repositori

Untuk memulai, pertama-tama Anda perlu meng-clone repositori ini ke dalam direktori lokal Anda. Jalankan perintah berikut:

```bash
git clone https://github.com/haikalirhamna/excel-calculation.git
cd excel-calculation
```

### 2. Install dependencies

Setelah repositori berhasil di-clone, Anda perlu menginstall dependensi yang diperlukan. Gunakan pip untuk menginstall package yang ada di requirements.txt

```bash
pip install -r requirements.txt
```

### 3. Jalankan Proyek

```bash
python main.py
```

## Rumus Perhitungan Nilai Akhir

Nilai akhir dihitung dengan menggunakan bobot dari **nilai bekerja** dan **nilai absensi** berdasarkan total skor bulanan yang telah dihitung.

### 1. Nilai Bekerja
Nilai bekerja dihitung berdasarkan aktivitas dalam seminggu, dengan nilai maksimum dalam satu bulan adalah 120. Nilai ini dihitung berdasarkan 6 hari kerja setiap minggu dan maksimum nilai per hari adalah 5. 

Rumus perhitungannya adalah:
- Nilai per hari kerja bisa mencapai **5** (nilai tertinggi jika bekerja pada level "critical").
- Dalam seminggu, terdapat **6 hari kerja** sehingga maksimum nilai dalam seminggu adalah:
  \[
  5 * 6 = 30
  \]
- Dalam satu bulan (4 minggu), maksimum nilai bekerja adalah:
  \[
  30 * 4 = 120
  \]

### 2. Nilai Absensi
Nilai absensi dihitung berdasarkan kehadiran dalam seminggu, dengan nilai maksimum dalam satu bulan adalah 72. Nilai absensi dihitung berdasarkan jumlah hari hadir, izin, sakit, dan tidak hadir.

Rumus perhitungannya adalah:
- Nilai untuk hadir adalah **3**, izin **2**, sakit **1**, dan tidak hadir **0**.
- Dalam seminggu, terdapat **6 hari kerja**, sehingga nilai maksimum absensi dalam seminggu adalah:
  \[
  3 * 6 = 18
  \]
- Dalam satu bulan (4 minggu), maksimum nilai absensi adalah:
  \[
  18 * 4 = 72
  \]

### 3. Rumus Total Nilai Akhir
Untuk menghitung nilai akhir, rumus yang digunakan adalah:


**Keterangan:**
- **Nilai Bekerja** adalah jumlah nilai bekerja dalam satu bulan (maksimal 120).
- **Nilai Absensi** adalah jumlah nilai absensi dalam satu bulan (maksimal 72).
- **192** adalah jumlah maksimum gabungan nilai antara bekerja dan absensi, yaitu \( 120 + 72 = 192 \).
- Hasil perhitungan akan dikali 100 untuk mendapatkan persentase.

### Contoh Perhitungan:

Misalnya, **Nilai Bekerja** = 100 dan **Nilai Absensi** = 60.

\[
Nilai Akhir = ((100 + 60)/(192) * 100) = ((160/192) * 100) = 83.33\%
\]
