#!/usr/bin/env python3
"""
SCRIPT MASTER - Analisis Lengkap Dokumen Word
Laporan_30_Artikel_Meditasi_Yogyakarta_2025_REVISI.docx

Menjalankan semua analisis:
1. Analisis dokumen
2. Ekstrak data ke CSV dan JSON
3. Visualisasi data
"""

import subprocess
import os

def run_script(script_name, description):
    """Menjalankan script Python"""
    print(f"\n{'='*70}")
    print(f"🚀 {description}")
    print(f"{'='*70}")
    
    result = subprocess.run(['python3', script_name], capture_output=True, text=True)
    print(result.stdout)
    if result.stderr:
        print("⚠️  Warnings/Errors:")
        print(result.stderr)
    
    return result.returncode == 0

def main():
    print("""
╔══════════════════════════════════════════════════════════════════╗
║           ANALISIS LENGKAP DOKUMEN WORD - MASTER SCRIPT          ║
╚══════════════════════════════════════════════════════════════════╝

File Input: Laporan_30_Artikel_Meditasi_Yogyakarta_2025_REVISI.docx

Proses yang akan dijalankan:
  1. Analisis Dokumen (statistik, kutipan, kata kunci)
  2. Ekstrak Data (CSV dan JSON)
  3. Visualisasi Data (grafik dan laporan)
""")
    
    input("Tekan ENTER untuk memulai...")
    
    scripts = [
        ('analisis_dokumen.py', 'Analisis Dokumen Word'),
        ('ekstrak_data.py', 'Ekstrak Data ke CSV dan JSON'),
        ('visualisasi_data.py', 'Visualisasi Data')
    ]
    
    hasil = []
    for script, desc in scripts:
        success = run_script(script, desc)
        hasil.append((desc, success))
    
    # Ringkasan
    print(f"\n{'='*70}")
    print("📋 RINGKASAN EKSEKUSI")
    print(f"{'='*70}")
    
    for desc, success in hasil:
        status = "✅ BERHASIL" if success else "❌ GAGAL"
        print(f"  {status} - {desc}")
    
    # Daftar file yang dihasilkan
    print(f"\n{'='*70}")
    print("📁 FILE YANG DIHASILKAN")
    print(f"{'='*70}")
    
    files = [
        'data_artikel_meditasi_yogyakarta.csv',
        'data_artikel_meditasi_yogyakarta.json',
        'distribusi_media_python.png',
        'timeline_publikasi_python.png',
        'kata_kunci_python.png',
        'laporan_statistik.txt'
    ]
    
    for i, file in enumerate(files, 1):
        if os.path.exists(file):
            size = os.path.getsize(file)
            print(f"  {i}. {file:45s} ({size:,} bytes)")
        else:
            print(f"  {i}. {file:45s} (tidak ditemukan)")
    
    print(f"\n{'='*70}")
    print("✅ SEMUA PROSES SELESAI!")
    print(f"{'='*70}\n")

if __name__ == "__main__":
    main()
