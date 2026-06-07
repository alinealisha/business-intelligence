# Tugas Individu — Business Intelligence (IBIS)

> **Inteligensi Bisnis (IBIS) | Fakultas Ilmu Komputer, Universitas Indonesia**  
> Alisha Aline Athiyyah — 2306207921

Repositori ini berisi empat tugas individu end-to-end dalam mata kuliah Business Intelligence, mencakup persiapan data, pembuatan kalkulasi, visualisasi dashboard, dan analisis keputusan bisnis.

---

## Tugas 1 — Microsoft Excel (Power Query & Pivot Table)

**Dataset:** Data penjualan elektronik (Data_Pelanggan, Data_Produk, Data_Penjualan)  
**Tool:** Microsoft Excel — Power Query, PivotTable, PivotChart

### Data Cleaning & Transformation (Power Query)

Proses persiapan data dilakukan di Power Query sebelum data masuk ke model. Perbedaan utama antara keduanya:

- **Data Cleaning** → memperbaiki data yang salah, tidak konsisten, atau tidak lengkap (hapus baris kosong, standarisasi format tanggal, bersihkan prefix "Rp"/"RP", seragamkan kapitalisasi nama Sales Rep, replace nilai null)
- **Data Transformation** → mengubah bentuk/struktur data untuk keperluan analisis (tambah kolom Notes Gudang, hitung Revenue/Profit/Margin, gabungkan tiga dataset menjadi satu laporan)

| Tabel | Langkah Cleaning & Transformation |
|---|---|
| Data Pelanggan | Hapus baris kosong |
| Data Produk | Hapus baris kosong · Standarisasi Harga Jual (replace "Rp"/"RP", trim, ubah ke integer) · Split column status · Tambah kolom Notes Gudang (logika stok kondisional) |
| Data Penjualan | Hapus baris kosong · Standarisasi format tanggal (multi-format → Date via Custom Column) · Replace null Jumlah → 10 · Gabungkan data duplikat (Group By) · Seragamkan kapitalisasi Sales Rep · Merge dengan produk & pelanggan · Hitung Revenue, Profit, Margin |

**Formula Notes Gudang:**
```
if [Status] = "Aktif" then
    if [Stok] < [Min_Stok] then "Perlu Restock"
    else if [Stok] <= [Min_Stok] * 1.5 then "Stok Hampir Habis"
    else "Stok Aman"
else
    if [Stok] = 0 then "Habis"
    else "Barang Hilang"
```

### Pivot Table yang Dibuat

| Label | Deskripsi |
|---|---|
| Pivot Table A | Revenue per Kota × Kategori produk |
| Pivot Table B.1 | Jumlah transaksi per Sales Rep per Bulan (grouped by month) |
| Pivot Table B.2 | Total Profit per Sales Rep |
| Pivot Table C | Stok vs Min_Stok per Produk + filter Notes Gudang + Clustered Bar Chart |
| Pivot Table D | Total transaksi & rata-rata margin per Pelanggan (dengan Tipe pelanggan) |
| Summary COO | Laporan ringkasan untuk Chief Operating Officer |

### Perbandingan VLOOKUP vs Power Query

VLOOKUP lebih mudah untuk kebutuhan cepat dengan satu-dua kolom dari tabel kecil, namun punya keterbatasan: harus tahu nomor urut kolom, rawan error jika kolom bergeser, dan harus menulis formula terpisah untuk tiap kolom yang ingin diambil. Untuk dataset kompleks seperti tugas ini (3 tabel, ratusan baris, proses yang perlu diulang), Power Query lebih unggul karena merge dilakukan sekali untuk semua kolom, setiap langkah tercatat di Applied Steps, dan seluruh proses bisa di-refresh otomatis saat data sumber diperbarui.

---

## Tugas 2 — Microsoft Power BI

**Dataset:** `spotify_data_clean.xlsx` + `track_data_final.xlsx` (implementasi DAX) · `TCustomers`, `TProducts`, `TSales` dari Excel (tutorial Power BI)  
**Tool:** Microsoft Power BI Desktop

### Ringkasan Materi Tutorial

Tugas ini mencakup pemahaman menyeluruh terhadap ekosistem Power BI:

- **Power BI Interface** — Report View, Table View, Model View, DAX Query View, Ribbon, Fields/Visualizations/Filters Pane
- **Connecting to Data Source** — import dari Excel, auto-detect relasi antar tabel (one-to-many Customer ID dan Product ID)
- **Power Query** — transformasi data (Group By, Transpose, Split Column, Custom Column, Remove Duplicates, tipe data)
- **Data Modelling** — relasi one-to-many antar tabel, tabel kalender DAX (`CALENDARAUTO()`), Mark as Date Table
- **Data Visualization** — Bar Chart, Line Graph, Pie Chart, Slicer, Timeline Filter, Bookmarks
- **DAX Functions** — SUM, SUMX, AVERAGE, CALCULATE, FILTER, RELATED, IF, DATEADD, RANKX

### DAX Functions yang Diimplementasikan

| Fungsi | Kegunaan | Contoh pada Dataset |
|---|---|---|
| `IF()` | Kategorisasi kondisional bertingkat | Kategorikan lagu: "Populer" / "Cukup Populer" / "Kurang Populer" berdasarkan `track_popularity` |
| `FILTER()` | Menghasilkan tabel subset berdasarkan kondisi | Saring lagu album dengan `track_popularity ≥ 80` dan `album_type = "album"` |
| `DATEADD()` | Pergeseran konteks waktu (time intelligence) | Bandingkan rata-rata popularitas tahun ini vs tahun sebelumnya (year-over-year) |
| `RANKX()` | Peringkat nilai dalam tabel | Ranking artis berdasarkan jumlah followers dan rata-rata popularitas (DESC, DENSE) |
| `SUMX()` | Kalkulasi per baris sebelum dijumlahkan | Konversi durasi lagu dari ms ke menit per baris (`/ 60000`), lalu jumlahkan per artis |
| `CALCULATE()` | Modifikasi konteks filter kalkulasi | Hitung total penjualan untuk produk tertentu atau bulan sebelumnya |
| `RELATED()` | Ambil nilai dari tabel lain via relasi | Tarik nama produk, kategori, harga dari TProducts ke TSales sebagai calculated column |

### Komponen Power BI

| Komponen | Deskripsi |
|---|---|
| Data | Data mentah yang diimpor dari berbagai sumber |
| Data Model | Struktur yang mengorganisir tabel beserta hubungannya |
| Relationships | Koneksi antar tabel berdasarkan kolom key yang sama |
| Measures | Kalkulasi DAX yang dihitung berdasarkan konteks filter aktif |
| Calculated Columns | Kolom baru berbasis DAX yang dihitung saat data di-load |
| Hierarchies | Susunan bertingkat kolom untuk mendukung drill-down |
| Filters | Pembatasan data pada level visual, halaman, atau seluruh laporan |
| Dashboard | Tampilan satu halaman berisi kumpulan visualisasi penting |

### Tipe Data Power BI

`Decimal Number` · `Fixed Decimal Number (Currency)` · `Whole Number` · `Date/Time` · `Text` · `True/False (Boolean)`

### Drill Down, Drill Up & Drill Through

- **Drill Down / Drill Up** — navigasi hierarkis dalam satu visual yang sama (contoh: release_year → album_name → track_name pada `album_release_date`)
- **Drill Through** — pindah ke halaman baru yang berisi detail spesifik dari data yang diklik (contoh: klik lagu "Bye Bye Bye" → halaman detail menampilkan popularitas 79.00, durasi 3.34 menit, album "No Strings Attached", tipe "album")

### Dashboard yang Dibuat

**Dashboard 1 — Music Overview**

| Chart | Deskripsi |
|---|---|
| 5a | Top 10 artis berdasarkan jumlah followers (bar chart) |
| 5b | Top 10 lagu terpopuler sepanjang masa (bar chart) |
| 5c | Rata-rata durasi lagu per tahun (line chart + drill down/up) |
| 5d | Proporsi tipe album: single vs album vs compilation (pie chart) |

**Dashboard 2 — Artist Comparison** (filter: Jan 2010 – Jan 2025)

| Chart | Deskripsi |
|---|---|
| 5f | Tren popularitas dari waktu ke waktu — Eminem vs Sabrina Carpenter vs Taylor Swift (line chart) |
| 5g | Rata-rata jumlah lagu per album per tahun — ketiga artis yang sama (line/bar chart) |

---

## Tugas 3 — Lab Tableau

**Dataset:** `DatasetPerformaSiswa.csv`  
**Tool:** Python (Jupyter Notebook) · Tableau Prep · Tableau Public

### Data Cleaning dengan Python

Proses pembersihan data dilakukan secara penuh di Jupyter Notebook sebelum divisualisasikan di Tableau:

| Langkah | Temuan & Penanganan |
|---|---|
| Missing Values | 3 kolom: `Teacher_Quality` (78), `Parental_Education_Level` (90), `Distance_from_Home` (67) → diimputasi dengan modus karena kolom kategorikal |
| Duplicate Data | Tidak ditemukan duplikat |
| Tipe Data & Distribusi Kategorikal | Semua nilai valid dan konsisten, tidak ada typo atau inkonsistensi |
| Outliers (IQR method) | `Hours_Studied` (43 outlier) → capping · `Exam_Score` (104 outlier) → capping · `Tutoring_Sessions` (430 outlier) → dibiarkan karena nilai 4–8 sesi masih logis secara domain |

### Komponen Tableau yang Dipelajari

**Data Pane:**

| Komponen | Deskripsi |
|---|---|
| Dimensions | Field kategorikal (teks, tanggal, boolean) untuk mengelompokkan data, tidak diagregasi secara default |
| Measures | Field numerik yang diagregasi (SUM, AVG, COUNT, dll) |
| Parameters | Variabel dinamis yang bisa diubah user saat runtime |
| Sets | Subset data custom berdasarkan kondisi tertentu atau pilihan manual |
| Calculated Fields | Field baru hasil komputasi menggunakan formula Tableau |
| Bins | Pengelompokan nilai numerik ke dalam rentang berukuran sama |
| Group | Penggabungan beberapa member dalam satu dimensi menjadi satu kategori baru secara manual |

**Perbedaan Bins, Group, dan Convert to Dimension:**

- **Bins** — untuk kolom angka yang ingin dilihat distribusinya dalam rentang tertentu, cocok untuk histogram (contoh: `Hours_Studied` dipecah jadi 0–5, 5–10, dst)
- **Group** — untuk kolom kategori yang ingin digabungkan secara manual (contoh: Low + Medium pada `Motivation_Level` → "Non-High")
- **Convert to Dimension** — untuk kolom angka yang merepresentasikan label, bukan nilai yang perlu diagregasi (contoh: `Tutoring_Sessions` sebagai sumbu kategori)

### Visualisasi yang Dibuat (DatasetPerformaSiswa)

| Chart | Pertanyaan | Insight |
|---|---|---|
| 5a | Perbandingan jam belajar berdasarkan gender | Female 19.99 jam vs Male 19.95 jam — hampir identik, gender tidak memengaruhi durasi belajar |
| 5b | Komposisi siswa berdasarkan gender | 57.73% Female, 42.27% Male |
| 5c | Gambaran tingkat motivasi siswa | Medium 3.351 · Low 1.937 · High 1.319 |
| 5d | Kehadiran: siswa disabilitas vs non-disabilitas | Non-disabilitas 80.07 vs disabilitas 79.23 |
| 5e | Rata-rata nilai ujian berdasarkan tingkat kehadiran | Tren positif: kehadiran rendah (~60) → nilai 64.15, kehadiran tinggi (~100) → nilai 70.73 |
| 5f | Distribusi performa akademik berdasarkan nilai ujian | Terpusat di rentang 65–70, bin 65 terbanyak (3.530 siswa), menyerupai kurva normal |
| 5g | Durasi belajar vs nilai ujian | Hubungan positif, tren naik untuk kedua gender |
| 5h | Kehadiran vs performa akademik | Scatter plot menunjukkan hubungan positif yang konsisten |
| 5i | Previous Score vs Exam Score berdasarkan kualitas guru | Pola positif kuat; kualitas guru tidak terlalu membedakan pola |
| 5j | Pendidikan orangtua vs nilai ujian | High School 66.78 → College 67.21 → Postgraduate 67.87 |
| 5k | Keterlibatan orangtua vs nilai ujian | Low 66.24 → Medium 67.02 → High 67.94 |
| 5l | Durasi tidur vs nilai ujian | Relatif stabil di ~67 untuk semua durasi tidur, tidak signifikan |
| 5m | Rata-rata jam belajar per kategori performa | High Performer 26.10 jam vs Low Performer 19.86 jam |
| 5n | Keseimbangan kebiasaan belajar vs nilai ujian | Balanced 67.13 vs Unbalanced 67.12 — perbedaan tidak signifikan |

**Dashboard Story (4 grafik berkorelasi):** Kebiasaan belajar sebagai penentu utama performa akademik — dimulai dari kesetaraan gender dalam durasi belajar, lalu korelasi positif kehadiran terhadap nilai, scatter plot jam belajar vs nilai, hingga perbedaan jam belajar High vs Low Performer (26.1 vs 19.86 jam).

---

## Tugas 4 — Analisis NPV (Net Present Value)

**Dataset:** Soal keputusan investasi multi-skenario (dua opsi)  
**Tool:** Microsoft Excel (perhitungan manual) · ChatGPT GPT-5.2 (pembanding Generative AI)

### Ringkasan Analisis

Tugas ini membandingkan hasil perhitungan Expected NPV secara manual menggunakan spreadsheet dengan hasil yang dihasilkan oleh Generative AI (ChatGPT GPT-5.2 versi gratis). Kedua opsi dievaluasi berdasarkan nilai Expected NPV untuk menentukan pilihan investasi terbaik.

### Hasil Perhitungan

| Metode | Expected NPV Opsi 1 | Expected NPV Opsi 2 | Rekomendasi |
|---|---|---|---|
| Manual (Spreadsheet) | $401,657,851.24 | $396,872,727.27 | Opsi 1 |
| Generative AI (ChatGPT) | ≈ $124 juta | ≈ $123 juta | Opsi 1 |

Kesimpulan akhir keduanya sama: **Opsi 1 lebih baik dipilih** dibandingkan Opsi 2.

### Sumber Perbedaan Nilai

| Aspek | Perhitungan Manual | ChatGPT |
|---|---|---|
| Periode yang dihitung | P0, P1, P2 | Hanya P2 |
| Perlakuan diskonto | P1 didiskon 1× dengan (1.1)¹, P2 didiskon 2× dengan (1.1)² | EV dihitung dulu, lalu didiskon sekaligus dengan (1.1)² |
| Komponen NPV | P0 net cash flow dimasukkan sebagai komponen NPV | P0 tidak dimasukkan |
| Asumsi alokasi produksi | Eksplisit per skenario | Berbeda dari perhitungan manual |

### Analisis Kepercayaan Hasil

Perhitungan manual dengan spreadsheet **lebih dapat dipercaya** karena:

- Menghitung semua periode (P0, P1, P2) sesuai permintaan soal — ChatGPT melewatkan P0 dan P1 sehingga dua sumber cash flow yang signifikan tidak diperhitungkan
- Perlakuan diskonto dilakukan per periode secara tepat — ChatGPT mendiskon seluruh EV sekaligus tanpa membedakan periode
- Setiap langkah kalkulasi transparan dan dapat diverifikasi di spreadsheet

Meski demikian, keduanya menghasilkan **rekomendasi akhir yang sama (Opsi 1)**, menunjukkan bahwa ChatGPT masih bisa memberikan arah keputusan yang benar meskipun nilai absolutnya tidak akurat.

---

## Tools & Teknologi

| Kategori | Tools |
|---|---|
| Spreadsheet & ETL | Microsoft Excel, Power Query (M Language) |
| BI & Visualisasi | Microsoft Power BI Desktop |
| Data Viz Lanjutan | Tableau Public |
| Data Preparation | Python (Jupyter Notebook), pandas |
| Formula Language | DAX (Data Analysis Expressions) |
| Analisis Keputusan | Net Present Value (NPV), Expected Value |

---

*Tugas Individu — Inteligensi Bisnis (IBIS), Genap 2025/2026, Fakultas Ilmu Komputer, Universitas Indonesia*
