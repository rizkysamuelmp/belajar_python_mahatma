# 🐍 SCRIPT PYTHON UNTUK ANALISIS DOKUMEN WORD

## 📄 Deskripsi

Kumpulan script Python untuk menganalisis dokumen Word:
**Laporan_30_Artikel_Meditasi_Yogyakarta_2025_REVISI.docx**

## 📋 Daftar Script

### 1. **analisis_dokumen.py**
Menganalisis dokumen Word dan mengekstrak informasi penting.

**Fitur:**
- Membaca dokumen Word
- Menghitung statistik (paragraf, kata, artikel)
- Mengekstrak kutipan verbatim
- Analisis kata kunci yang sering muncul
- Analisis tema berdasarkan keyword

**Output:**
- Statistik dokumen di terminal
- Top 10 kata kunci
- Contoh kutipan verbatim
- Analisis tema

**Cara Menjalankan:**
```bash
python3 analisis_dokumen.py
```

---

### 2. **ekstrak_data.py**
Mengekstrak data artikel dari dokumen Word ke format CSV dan JSON.

**Fitur:**
- Ekstrak 30 artikel lengkap
- Parsing judul, tanggal, media, link, ringkasan, kutipan
- Export ke CSV (untuk Excel/spreadsheet)
- Export ke JSON (untuk programming)

**Output:**
- `data_artikel_meditasi_yogyakarta.csv`
- `data_artikel_meditasi_yogyakarta.json`

**Cara Menjalankan:**
```bash
python3 ekstrak_data.py
```

---

### 3. **visualisasi_data.py**
Membuat visualisasi dari data artikel.

**Fitur:**
- Visualisasi distribusi artikel per media
- Timeline publikasi artikel
- Analisis kata kunci dalam kutipan
- Laporan statistik lengkap

**Output:**
- `distribusi_media_python.png`
- `timeline_publikasi_python.png`
- `kata_kunci_python.png`
- `laporan_statistik.txt`

**Cara Menjalankan:**
```bash
python3 visualisasi_data.py
```

---

### 4. **run_all_analysis.py** ⭐
Script master yang menjalankan semua analisis sekaligus.

**Fitur:**
- Menjalankan ketiga script di atas secara berurutan
- Menampilkan ringkasan eksekusi
- Daftar file yang dihasilkan

**Cara Menjalankan:**
```bash
python3 run_all_analysis.py
```

---

## 🚀 Quick Start

### Instalasi Dependencies
```bash
pip install python-docx matplotlib pandas
```

### Menjalankan Semua Analisis
```bash
python3 run_all_analysis.py
```

### Menjalankan Analisis Individual
```bash
# Analisis dokumen saja
python3 analisis_dokumen.py

# Ekstrak data saja
python3 ekstrak_data.py

# Visualisasi saja
python3 visualisasi_data.py
```

---

## 📊 Output yang Dihasilkan

### File Data:
1. **data_artikel_meditasi_yogyakarta.csv** - Data dalam format CSV
2. **data_artikel_meditasi_yogyakarta.json** - Data dalam format JSON

### File Visualisasi:
3. **distribusi_media_python.png** - Bar chart distribusi media
4. **timeline_publikasi_python.png** - Bar chart timeline publikasi
5. **kata_kunci_python.png** - Bar chart kata kunci

### File Laporan:
6. **laporan_statistik.txt** - Laporan statistik lengkap

---

## 📈 Contoh Output

### Analisis Dokumen:
```
📊 STATISTIK DOKUMEN:
   Total Paragraf: 300
   Total Kata: 1,563
   Jumlah Artikel: 30

🔑 KATA KUNCI PALING SERING MUNCUL:
    1. meditation           :  12x
    2. healing              :   6x
    3. yoga                 :   6x
```

### Ekstrak Data:
```
✓ Berhasil mengekstrak 30 artikel
✓ File CSV berhasil dibuat
✓ File JSON berhasil dibuat
```

### Visualisasi:
```
✓ Visualisasi distribusi media berhasil dibuat
✓ Visualisasi timeline publikasi berhasil dibuat
✓ Visualisasi kata kunci berhasil dibuat
```

---

## 🔧 Struktur Data JSON

```json
[
  {
    "no": 1,
    "judul": "Wonderful Indonesia Wellness Festival 2025 Awaits",
    "tanggal": "17 Oktober 2025",
    "media": "PR Newswire APAC",
    "link": "https://www.prnewswire.com/...",
    "ringkasan": "Festival wellness perdana...",
    "kutipan": "\"spiritual spa treatments, meditation at historical sites\""
  }
]
```

---

## 📝 Catatan

- Semua script menggunakan Python 3
- Dependencies: python-docx, matplotlib, pandas
- File input harus ada di direktori yang sama
- Output akan disimpan di direktori yang sama

---

## 🎯 Use Cases

1. **Analisis Konten** - Menganalisis tema dan kata kunci
2. **Data Processing** - Mengolah data untuk penelitian
3. **Visualisasi** - Membuat grafik untuk presentasi
4. **Export Data** - Mengekspor ke format yang berbeda

---

## 📞 Troubleshooting

### Error: ModuleNotFoundError
```bash
pip install python-docx matplotlib pandas
```

### Error: File not found
Pastikan file `Laporan_30_Artikel_Meditasi_Yogyakarta_2025_REVISI.docx` ada di direktori yang sama.

### Error: Permission denied
```bash
chmod +x *.py
```

---

**Dibuat**: 3 Desember 2025  
**Lokasi**: /home/mahatma/belajar_python_mahatma/Diseratasi/
