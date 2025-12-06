#!/usr/bin/env python3
"""
Script untuk membuat visualisasi dari data artikel
"""

import json
import matplotlib.pyplot as plt
from collections import Counter
from datetime import datetime
import pandas as pd

def load_data():
    """Load data dari file JSON"""
    with open('data_artikel_meditasi_yogyakarta.json', 'r', encoding='utf-8') as f:
        return json.load(f)

def visualisasi_distribusi_media(data):
    """Visualisasi distribusi artikel per media"""
    media_count = Counter([artikel['media'] for artikel in data])
    
    plt.figure(figsize=(12, 6))
    media = list(media_count.keys())
    counts = list(media_count.values())
    
    plt.barh(media, counts, color='skyblue', edgecolor='black')
    plt.xlabel('Jumlah Artikel', fontweight='bold')
    plt.title('Distribusi Artikel Berdasarkan Media', fontweight='bold', fontsize=14)
    plt.tight_layout()
    plt.savefig('distribusi_media_python.png', dpi=300, bbox_inches='tight')
    print("✓ Visualisasi distribusi media berhasil dibuat")
    plt.close()

def visualisasi_timeline(data):
    """Visualisasi timeline publikasi artikel"""
    # Parse tanggal
    tanggal_list = []
    for artikel in data:
        try:
            tgl = artikel['tanggal']
            # Parse format "DD Bulan YYYY"
            if tgl:
                tanggal_list.append(tgl)
        except:
            pass
    
    # Hitung per bulan
    bulan_count = Counter([tgl.split()[1] if len(tgl.split()) > 1 else 'Unknown' for tgl in tanggal_list])
    
    plt.figure(figsize=(10, 6))
    bulan = list(bulan_count.keys())
    counts = list(bulan_count.values())
    
    plt.bar(bulan, counts, color='coral', edgecolor='black')
    plt.xlabel('Bulan', fontweight='bold')
    plt.ylabel('Jumlah Artikel', fontweight='bold')
    plt.title('Timeline Publikasi Artikel per Bulan', fontweight='bold', fontsize=14)
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig('timeline_publikasi_python.png', dpi=300, bbox_inches='tight')
    print("✓ Visualisasi timeline publikasi berhasil dibuat")
    plt.close()

def analisis_kata_kunci(data):
    """Analisis kata kunci dari kutipan"""
    semua_kutipan = ' '.join([artikel['kutipan'] for artikel in data])
    
    # Kata kunci yang dicari
    keywords = ['meditation', 'wellness', 'spiritual', 'healing', 'yoga', 
                'mindfulness', 'retreat', 'balance', 'clarity', 'resilience']
    
    keyword_count = {}
    for keyword in keywords:
        count = semua_kutipan.lower().count(keyword)
        if count > 0:
            keyword_count[keyword] = count
    
    # Visualisasi
    plt.figure(figsize=(10, 6))
    keywords_sorted = sorted(keyword_count.items(), key=lambda x: x[1], reverse=True)
    words = [k[0] for k in keywords_sorted]
    counts = [k[1] for k in keywords_sorted]
    
    plt.barh(words, counts, color='lightgreen', edgecolor='black')
    plt.xlabel('Frekuensi', fontweight='bold')
    plt.title('Kata Kunci yang Sering Muncul dalam Kutipan', fontweight='bold', fontsize=14)
    plt.tight_layout()
    plt.savefig('kata_kunci_python.png', dpi=300, bbox_inches='tight')
    print("✓ Visualisasi kata kunci berhasil dibuat")
    plt.close()
    
    return keyword_count

def buat_laporan_statistik(data):
    """Buat laporan statistik dalam bentuk teks"""
    total_artikel = len(data)
    media_unik = len(set([artikel['media'] for artikel in data]))
    
    # Hitung panjang rata-rata kutipan
    panjang_kutipan = [len(artikel['kutipan']) for artikel in data if artikel['kutipan']]
    rata_rata_kutipan = sum(panjang_kutipan) / len(panjang_kutipan) if panjang_kutipan else 0
    
    laporan = f"""
╔══════════════════════════════════════════════════════════════════╗
║                    LAPORAN STATISTIK ARTIKEL                     ║
╚══════════════════════════════════════════════════════════════════╝

📊 STATISTIK UMUM:
   Total Artikel        : {total_artikel}
   Media Unik           : {media_unik}
   Rata-rata Panjang    : {rata_rata_kutipan:.0f} karakter
   Kutipan

📅 PERIODE PUBLIKASI:
   Januari - November 2025

📰 MEDIA TERBANYAK:
"""
    
    media_count = Counter([artikel['media'] for artikel in data])
    for i, (media, count) in enumerate(media_count.most_common(5), 1):
        laporan += f"   {i}. {media:30s} : {count:2d} artikel\n"
    
    return laporan

def main():
    print("="*70)
    print("VISUALISASI DATA ARTIKEL")
    print("="*70)
    
    # Load data
    print("\n📖 Memuat data dari JSON...")
    data = load_data()
    print(f"   ✓ {len(data)} artikel berhasil dimuat")
    
    # Buat visualisasi
    print("\n📊 Membuat visualisasi...")
    visualisasi_distribusi_media(data)
    visualisasi_timeline(data)
    keyword_count = analisis_kata_kunci(data)
    
    # Buat laporan
    print("\n📝 Membuat laporan statistik...")
    laporan = buat_laporan_statistik(data)
    print(laporan)
    
    # Simpan laporan
    with open('laporan_statistik.txt', 'w', encoding='utf-8') as f:
        f.write(laporan)
    print("   ✓ Laporan disimpan ke: laporan_statistik.txt")
    
    print("\n" + "="*70)
    print("✅ VISUALISASI SELESAI!")
    print("="*70)
    print("\nFile yang dihasilkan:")
    print("  1. distribusi_media_python.png")
    print("  2. timeline_publikasi_python.png")
    print("  3. kata_kunci_python.png")
    print("  4. laporan_statistik.txt")

if __name__ == "__main__":
    main()
