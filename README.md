# Portfolio Excel — Dashboard Penjualan Toko Kelontong
**Nadya Theresia Tangkere| Data Analyst Portfolio Project**

---

## Tujuan Proyek

Proyek ini mensimulasikan sistem analisis penjualan untuk toko kelontong/sembako skala UMKM menggunakan Microsoft Excel tingkat lanjut. Proyek dirancang sebagai **portfolio data analyst**.

---

## Struktur File

```
toko_makmur_sejahtera_dashboard.xlsx
│
├── 🏠 DASHBOARD          ← Halaman utama, ringkasan semua analisis
├── DATA_TRANSAKSI        ← Raw data (3.335 baris, Jul–Des 2024)
├── REKAP_BULANAN         ← Agregasi penjualan per bulan
├── TARGET_VS_AKTUAL      ← Monitoring pencapaian target
├── ANALISIS_PRODUK       ← Ranking produk & margin per SKU
└── ANALISIS_MARGIN       ← Profitabilitas per kategori
```

---

## Penjelasan Tiap Sheet

### 1. 🏠 DASHBOARD
Halaman utama yang menampilkan ringkasan seluruh performa toko dalam satu layar.

| Komponen | Keterangan |
|---|---|
| KPI Cards | Total Penjualan, Laba Kotor, Rata-rata Margin, Total Transaksi |
| Tren Bulanan | Tabel dinamis Target vs Aktual lengkap dengan KPI bar visual |
| Top 5 Best Seller | Produk dengan penjualan tertinggi beserta margin-nya |
| Margin per Kategori | Profitabilitas 7 kategori produk dengan color scale |

> Semua data di Dashboard **otomatis terupdate** saat data di sheet lain diubah.

---

### 2. DATA_TRANSAKSI
Sheet ini berfungsi sebagai **"database" utama** toko — simulasi data transaksi harian.

| Kolom | Keterangan |
|---|---|
| ID Transaksi | Kode unik per transaksi (TRX-00001 dst.) |
| Tanggal | Tanggal transaksi (format DD/MM/YYYY) |
| Nama Produk | 20 SKU produk kelontong & sembako |
| Kategori | Sembako, Mie & Snack, Minuman, Kebersihan, dll. |
| Qty | Jumlah unit terjual |
| Harga Jual / Harga Beli | Harga per unit |
| Total Penjualan | `=Qty × Harga Jual` |
| Total HPP | `=Qty × Harga Beli` |
| Laba Kotor | `=Total Penjualan − Total HPP` |
| Bulan / Hari | Diekstrak otomatis menggunakan fungsi `TEXT()` |

**Volume data:** 3.335 baris transaksi selama 6 bulan (Jul–Des 2025), termasuk variasi weekend vs weekday.

---

### 3. REKAP_BULANAN
Agregasi otomatis dari DATA_TRANSAKSI menggunakan `SUMPRODUCT` berbasis kondisi bulan dan tahun.

**Formula utama:**
```excel
=SUMPRODUCT(
  (MONTH(DATA_TRANSAKSI!B$2:B$3336)=7) *
  (YEAR(DATA_TRANSAKSI!B$2:B$3336)=2024) *
  DATA_TRANSAKSI!H$2:H$3336
)
```

**Fitur:**
- Perhitungan otomatis Total Penjualan, HPP, Laba, Margin, dan % Pencapaian Target
- **Conditional Formatting RAG** (🟢 Hijau / 🟡 Kuning / 🔴 Merah) pada kolom % Pencapaian
- Warna otomatis pada kolom Selisih: hijau jika surplus, merah jika defisit

---

### 4. TARGET_VS_AKTUAL
Sheet monitoring performa bulanan dibandingkan target yang telah ditetapkan.

**Fitur unggulan:**
- **Kolom input biru** (`Target Penjualan`) — dapat diubah langsung oleh user
- **Status emoji otomatis** menggunakan nested `IF`:
  ```excel
  =IF(E6>=1,"✅ TERCAPAI",
    IF(E6>=0.85,"⚠️ HAMPIR",
      IF(E6>=0.7,"🔸 PERLU PERHATIAN","❌ TIDAK TERCAPAI")))
  ```
- **KPI Bar Visual** menggunakan fungsi `REPT`:
  ```excel
  =REPT("█", MIN(ROUND(E6*10,0),10)) &
   REPT("░", MAX(10-ROUND(E6*10,0),0)) &
   " " & TEXT(E6,"0%")
  ```
  Contoh output: `███████░░░ 70%`

---

### 5. ANALISIS_PRODUK
Ranking 20 SKU produk berdasarkan total penjualan, menggunakan `SUMIF` untuk agregasi per nama produk.

```excel
=SUMIF(DATA_TRANSAKSI!C:C, B5, DATA_TRANSAKSI!H:H)
```

**Fitur:**
- Ranking produk berdasarkan total penjualan & laba kotor
- **Color Scale** pada kolom Margin (merah = rendah → hijau = tinggi)
- Grade otomatis: ⭐ Premium / ✅ Baik / 🔸 Rata-rata / ⚠️ Rendah

---

### 6. ANALISIS_MARGIN
Breakdown profitabilitas per kategori produk menggunakan `SUMIF` pada kolom Kategori.

**Fitur:**
- Margin (%) dengan **Color Scale** kondisional
- Kontribusi penjualan tiap kategori terhadap total dengan **Data Bar**
- Grade margin otomatis berdasarkan threshold persentase

---

## Formula & Teknik Excel yang Digunakan

| Kategori | Formula / Fitur |
|---|---|
| **Agregasi Dinamis** | `SUMPRODUCT`, `SUMIF`, `SUMIFS`, `COUNTIFS` |
| **Logika & Teks** | `IF`, `IFS`, `IFERROR`, `TEXT`, `REPT` |
| **Tanggal** | `MONTH()`, `YEAR()`, `EOMONTH()`, `DATE()` |
| **Referensi** | Cross-sheet reference (`Sheet!Range`), Named Range |
| **Visualisasi Data** | Conditional Formatting (Color Scale, Data Bar, Icon Set, Formula-based) |
| **UX/Proteksi** | Input cell berwarna biru, footer dokumentasi, freeze panes |

---

## Konvensi Warna (Industry Standard)

| Warna | Arti |
|---|---|
| 🔵 Teks Biru | Input manual — angka yang bisa diubah user |
| ⚫ Teks Hitam | Formula — hasil kalkulasi otomatis |
| 🟢 Latar Hijau | Nilai di atas target / margin baik |
| 🟡 Latar Kuning | Mendekati target / perlu perhatian |
| 🔴 Latar Merah | Di bawah target / margin rendah |

---

## Cara Menggunakan

1. **Buka sheet `DATA_TRANSAKSI`** — ini adalah sumber data utama
2. **Tambah/edit transaksi baru** di baris berikutnya mengikuti format yang ada
3. **Ubah target penjualan** di sheet `TARGET_VS_AKTUAL` kolom B (sel berwarna kuning)
4. **Lihat hasilnya otomatis** di sheet `🏠 DASHBOARD`

> ⚠️ Jangan mengubah formula di kolom selain kolom input biru. Semua kalkulasi sudah terhubung secara dinamis.

---

## Insight Bisnis dari Data Simulasi

- **Desember** adalah bulan terbaik — penjualan tertinggi karena efek libur akhir tahun
- **Rokok & Gas LPG** memiliki margin rendah (<10%) namun volume penjualan tinggi
- **Kategori Kebersihan** konsisten memberi margin tertinggi (>25%)
- **Indomie Goreng** secara konsisten menjadi produk terlaris berdasarkan volume unit

---

## Tentang Proyek Ini

| | |
|---|---|
| **Jenis** | Portfolio Data Analyst — Excel Advanced |
| **Tools** | Microsoft Excel (kompatibel Office 2016+) |
| **Data** | Simulasi realistis, bukan data nyata |
| **Level** | Advanced (Power Query ready, dynamic formula) |
| **Industri** | Retail / UMKM — relevan untuk pasar Indonesia |

---
=======
# dashboard_toko_kelontong
>>>>>>> 6d7e8b56b1d95602929e32490d9320b045409410
