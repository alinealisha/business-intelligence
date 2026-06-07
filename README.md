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

**Dataset:** `DatasetPerformaSiswa.csv` — 6.607 siswa, 20 variabel akademik & demografis  
**Tool:** Python · pandas · Tableau Public

### Business Problem

Institusi pendidikan sering salah alokasi resource karena tidak tahu faktor mana yang benar-benar menggerakkan performa akademik. Analisis ini menjawab: **apa yang harus diprioritaskan — program kehadiran, jam belajar, kualitas tidur, atau keterlibatan orangtua?**

### Data Preparation (Python — pandas)

Notebook menjalankan 5 langkah pembersihan secara berurutan sebelum data diekspor ke Tableau:

| # | Langkah | Temuan | Penanganan | Reasoning |
|---|---|---|---|---|
| 1 | Missing values | `Teacher_Quality` (78), `Parental_Education_Level` (90), `Distance_from_Home` (67) | Imputasi modus | Kolom kategorikal — tidak ada urutan numerik, modus adalah representasi terbaik |
| 2 | Duplicate check | 0 duplikat | Tidak perlu penanganan | — |
| 3 | Konsistensi kategorikal | Semua nilai valid, tidak ada typo | Tidak perlu penanganan | — |
| 4 | Outlier (IQR method) | `Hours_Studied` (43), `Exam_Score` (104) outlier; `Tutoring_Sessions` (430) flagged | `Hours_Studied` & `Exam_Score` di-cap ke batas IQR; `Tutoring_Sessions` dibiarkan | Nilai 4–8 sesi masih domain-valid — upper IQR bound hanya 3.5 tapi nilai 8 sesi logis secara nyata |
| 5 | Valid range check | `Attendance` (0–100), `Sleep_Hours` (0–24), dll. | Semua dalam range valid | Verifikasi tambahan pasca-capping |

> Keputusan *tidak* men-cap `Tutoring_Sessions` adalah domain judgment yang disengaja — aturan statistik tidak selalu menang atas konteks realitas data.

### Visualisasi & Technical Implementation (Tableau Public)

14 chart dibangun, memanfaatkan fitur Tableau berikut:

| Fitur Tableau | Digunakan Untuk |
|---|---|
| **Bins** | `Hours_Studied` di-bin per 5 jam → histogram distribusi durasi belajar |
| **Calculated Field** | `Performance Score` (High/Low Performer dari `Exam_Score`) dan `Study Habit Balance` (`Hours_Studied ≥ 5` AND `Sleep_Hours ≥ 7`) |
| **Hierarchy + Drill Down/Up** | `Parental Education Level → Gender` untuk breakdown nilai ujian per level |
| **Scatter Plot + Trend Line** | Korelasi `Hours_Studied` vs `Exam_Score`, warna per Gender |
| **Reference Line** | Rata-rata keseluruhan sebagai baseline di line chart kehadiran vs nilai |
| **Dashboard (Tiled Layout)** | 4 chart berkorelasi disusun dalam satu story dashboard dengan narasi terhubung |

### Key Findings

| Variabel | Data | Signifikansi | Implikasi Bisnis |
|---|---|---|---|
| Kehadiran | Nilai: 64.15 (kehadiran ~60%) → 70.73 (kehadiran ~100%) | ✅ Tinggi | Program pemantauan kehadiran = ROI tertinggi |
| Durasi belajar | High Performer: 26.1 jam/minggu vs Low Performer: 19.86 jam (+32%) | ✅ Tinggi | Structured study program terbukti worth it |
| Keterlibatan orangtua | Low 66.24 → High 67.94 (selisih ~1.7 poin) | ⚠️ Rendah | Efek ada tapi kecil; program pendukung saja |
| Gender | Female 19.99 jam vs Male 19.95 jam (hampir identik) | ❌ Tidak signifikan | Tidak perlu intervensi berbasis gender |
| Durasi tidur | Flat ~67 untuk semua kelompok (3–11 jam) | ❌ Tidak signifikan | Program sleep hygiene tidak akan menggerakkan nilai |

### Business Recommendation

Dari 6.600+ siswa, **kehadiran dan durasi belajar adalah dua lever utama yang actionable**. Institusi sebaiknya fokus di sana — bukan di faktor demografis (gender) atau gaya hidup (tidur) yang datanya sendiri menunjukkan tidak membuat perbedaan.

---

## Tugas 4 — Analisis NPV (Net Present Value)

**Dataset:** Keputusan investasi ekspansi kapasitas — 2 opsi, 4 skenario permintaan, 3 periode  
**Tool:** Microsoft Excel

### Business Problem

Perusahaan perlu memilih antara dua opsi ekspansi kapasitas produksi Asia untuk melayani pasar Asia dan Eropa, di tengah ketidakpastian permintaan masa depan. Keputusan ini tidak bisa hanya dilihat dari biaya investasi — harus memperhitungkan **seluruh arus kas yang bisa di-capture di setiap kemungkinan skenario**.

### Metodologi (Excel)

Model dibangun end-to-end di Excel, dengan struktur kalkulasi sebagai berikut:

**1. Margin kontribusi per rute produksi** (dihitung dari parameter input)

| Rute | Harga Jual | Biaya Produksi | Biaya Kirim | Margin/Unit |
|---|---|---|---|---|
| Asia → Asia | $40 | $15 | $0 | **$25** |
| Asia → Eropa | $40 | $15 | $2 | **$23** |
| Eropa → Eropa | $40 | $20 | $0 | **$20** |
| Eropa → Asia | $40 | $20 | $2 | **$18** |

**2. Kapasitas setelah ekspansi**

| | Kapasitas Asia | Kapasitas Eropa | Berlaku |
|---|---|---|---|
| Existing (P0) | 2.400.000 | 4.500.000 | P0 |
| Opsi 1 (+2 juta unit) | 4.400.000 | 4.500.000 | P1 & P2 |
| Opsi 2 (+1.5 juta unit) | 3.900.000 | 4.500.000 | P1 & P2 |

**3. Skenario permintaan** — 4 kombinasi dari probabilitas independen Asia × Eropa

| Skenario | Pertumbuhan Asia | Pertumbuhan Eropa | Probabilitas |
|---|---|---|---|
| HH | +40% | +10% | 0.5 × 0.6 = **0.3** |
| HL | +40% | −20% | 0.5 × 0.4 = **0.2** |
| LH | +20% | +10% | 0.5 × 0.6 = **0.3** |
| LL | +20% | −20% | 0.5 × 0.4 = **0.2** |

**4. Formula Expected NPV**

Alokasi produksi dioptimalkan per skenario (Asia diprioritaskan karena margin lebih tinggi). NPV dihitung per skenario lalu di-weighted dengan probabilitas:

```
NPV(skenario) = (Profit_P0 − Investasi) + Profit_P1/(1.1)¹ + Profit_P2/(1.1)²
Expected NPV  = Σ NPV(skenario) × Probabilitas(skenario)
```

### Hasil

| | Opsi 1 | Opsi 2 |
|---|---|---|
| Investasi | $17.000.000 | $15.000.000 |
| P0 Net Cash Flow | $118.600.000 | $120.600.000 |
| **Expected NPV** | **$401.657.851** | **$396.872.727** |
| Selisih | +$4.784.851 vs Opsi 2 | — |

### Business Recommendation

**Pilih Opsi 1.** Tambahan investasi $2 juta menghasilkan selisih NPV ~$4,8 juta — incremental ROI positif. Kapasitas lebih besar memberi kemampuan menyerap upside demand pada skenario pertumbuhan tinggi (HH + LH), yang secara agregat memiliki bobot probabilitas 60%. Dengan kata lain, memilih Opsi 2 menghemat $2 juta di muka tapi meninggalkan ~$4,8 juta di atas meja.

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
