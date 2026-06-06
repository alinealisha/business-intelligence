# Tugas Individu — Business Intelligence (IBIS)

> **Inteligensi Bisnis (IBIS) | Fakultas Ilmu Komputer, Universitas Indonesia**  
> Alisha Aline Athiyyah — 2306207921

Repositori ini berisi dua tugas individu end-to-end dalam mata kuliah Business Intelligence, mencakup persiapan data, pembuatan kalkulasi, dan penyusunan dashboard untuk analisis bisnis.

---

## Tugas 1 — Microsoft Excel (Power Query & Pivot Table)

**Dataset:** Data penjualan elektronik (Data_Pelanggan, Data_Produk, Data_Penjualan)  
**Tool:** Microsoft Excel — Power Query, PivotTable, PivotChart

### Data Cleaning & Transformation (Power Query)

Proses persiapan data dilakukan di Power Query sebelum data masuk ke model. Perbedaan utama antara keduanya:

- **Data Cleaning** → memperbaiki data yang salah, tidak konsisten, atau tidak lengkap (hapus baris kosong, standarisasi format tanggal, bersihkan prefix "Rp", seragamkan kapitalisasi nama Sales Rep, replace nilai null)
- **Data Transformation** → mengubah bentuk/struktur data untuk keperluan analisis (tambah kolom Notes Gudang, hitung Revenue/Profit/Margin, gabungkan tiga dataset menjadi satu laporan)

| Tabel | Langkah Cleaning & Transformation |
|---|---|
| Data Pelanggan | Hapus baris kosong |
| Data Produk | Hapus baris kosong · Standarisasi Harga Jual (replace "Rp"/"RP", trim, ubah ke integer) · Tambah kolom Notes Gudang (logika stok kondisional) |
| Data Penjualan | Hapus baris kosong · Standarisasi format tanggal (multi-format → Date) · Replace null Jumlah → 10 · Gabungkan data duplikat · Seragamkan kapitalisasi Sales Rep · Merge dengan produk & pelanggan · Hitung Revenue, Profit, Margin |

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

VLOOKUP lebih mudah untuk kebutuhan cepat dengan satu-dua kolom dari tabel kecil. Namun untuk dataset kompleks seperti tugas ini (3 tabel, ratusan baris, proses yang perlu diulang), Power Query lebih unggul karena: merge dilakukan sekali untuk semua kolom, setiap langkah tercatat di Applied Steps, dan seluruh proses bisa di-refresh otomatis saat data sumber diperbarui.

---

## Tugas 2 — Microsoft Power BI

**Dataset:** `spotify_data_clean.xlsx` + `track_data_final.xlsx` (data musik Spotify)  
**Tool:** Microsoft Power BI Desktop

### Ringkasan Materi Tutorial

Tugas ini mencakup pemahaman menyeluruh terhadap ekosistem Power BI:

- **Power BI Interface** — Report View, Table View, Model View, DAX Query View, Ribbon, Fields/Visualizations/Filters Pane
- **Connecting to Data Source** — import dari Excel, auto-detect relasi antar tabel
- **Power Query** — transformasi data (Group By, Transpose, Split Column, Custom Column, Remove Duplicates, tipe data)
- **Data Modelling** — relasi one-to-many antar tabel, tabel kalender DAX (`CALENDARAUTO()`), Mark as Date Table
- **Data Visualization** — Bar Chart, Line Graph, Pie Chart, Slicer, Timeline Filter, Bookmarks
- **DAX Functions** — SUM, SUMX, AVERAGE, CALCULATE, FILTER, RELATED, IF, DATEADD, RANKX

### DAX Functions yang Diimplementasikan

| Fungsi | Kegunaan | Contoh pada Dataset |
|---|---|---|
| `IF()` | Kategorisasi kondisional bertingkat | Kategorikan lagu: "Populer" / "Cukup Populer" / "Kurang Populer" berdasarkan `track_popularity` |
| `FILTER()` | Menghasilkan tabel subset berdasarkan kondisi | Saring lagu album dengan `track_popularity ≥ 80` |
| `DATEADD()` | Pergeseran konteks waktu (time intelligence) | Bandingkan rata-rata popularitas tahun ini vs tahun sebelumnya (year-over-year) |
| `RANKX()` | Peringkat nilai dalam tabel | Ranking artis berdasarkan jumlah followers dan rata-rata popularitas (DESC, DENSE) |
| `SUMX()` | Kalkulasi per baris sebelum dijumlahkan | Konversi durasi lagu dari ms ke menit per baris, lalu jumlahkan per artis |

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

- **Drill Down / Drill Up** — navigasi hierarkis dalam satu visual yang sama (contoh: Year → Quarter → Month pada `album_release_date`)
- **Drill Through** — pindah ke halaman baru yang berisi detail spesifik dari data yang diklik (contoh: klik lagu "Bye Bye Bye" → halaman detail menampilkan popularitas, durasi, nama album, tipe album)

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

## Tools & Teknologi

| Kategori | Tools |
|---|---|
| Spreadsheet & ETL | Microsoft Excel, Power Query (M Language) |
| BI & Visualisasi | Microsoft Power BI Desktop |
| Formula Language | DAX (Data Analysis Expressions) |

---


*Tugas Individu — Inteligensi Bisnis (IBIS), Genap 2025/2026, Fakultas Ilmu Komputer, Universitas Indonesia*
