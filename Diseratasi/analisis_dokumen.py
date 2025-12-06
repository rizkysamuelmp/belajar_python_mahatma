#!/usr/bin/env python3
"""
Script untuk menganalisis dokumen Word:
Laporan_30_Artikel_Meditasi_Yogyakarta_2025_REVISI.docx
"""

from docx import Document
from collections import Counter
import re

def baca_dokumen(file_path):
    """Membaca dokumen Word dan mengekstrak teks"""
    doc = Document(file_path)
    teks_lengkap = []
    
    for para in doc.paragraphs:
        if para.text.strip():
            teks_lengkap.append(para.text)
    
    return teks_lengkap

def ekstrak_kutipan(teks_list):
    """Mengekstrak semua kutipan verbatim dari dokumen"""
    kutipan = []
    for teks in teks_list:
        # Cari teks dalam tanda petik
        matches = re.findall(r'"([^"]+)"', teks)
        kutipan.extend(matches)
    return kutipan

def analisis_kata_kunci(kutipan_list):
    """Menganalisis kata kunci yang sering muncul dalam kutipan"""
    semua_kata = []
    stopwords = {'the', 'a', 'an', 'and', 'or', 'but', 'in', 'on', 'at', 'to', 'for', 'of', 'with', 'by', 'from', 'up', 'is', 'are'}
    
    for kutipan in kutipan_list:
        kata = kutipan.lower().split()
        kata_bersih = [k.strip('.,!?;:') for k in kata if k.strip('.,!?;:') not in stopwords and len(k) > 2]
        semua_kata.extend(kata_bersih)
    
    return Counter(semua_kata)

def hitung_statistik(teks_list):
    """Menghitung statistik dokumen"""
    total_paragraf = len(teks_list)
    total_kata = sum(len(teks.split()) for teks in teks_list)
    
    # Hitung jumlah artikel
    jumlah_artikel = sum(1 for teks in teks_list if teks.startswith('Artikel '))
    
    return {
        'total_paragraf': total_paragraf,
        'total_kata': total_kata,
        'jumlah_artikel': jumlah_artikel
    }

def main():
    file_path = 'Laporan_30_Artikel_Meditasi_Yogyakarta_2025_REVISI.docx'
    
    print("="*70)
    print("ANALISIS DOKUMEN WORD")
    print("="*70)
    print(f"\nFile: {file_path}\n")
    
    # Baca dokumen
    print("📖 Membaca dokumen...")
    teks = baca_dokumen(file_path)
    
    # Statistik dasar
    print("\n📊 STATISTIK DOKUMEN:")
    stats = hitung_statistik(teks)
    print(f"   Total Paragraf: {stats['total_paragraf']}")
    print(f"   Total Kata: {stats['total_kata']:,}")
    print(f"   Jumlah Artikel: {stats['jumlah_artikel']}")
    
    # Ekstrak kutipan
    print("\n💬 MENGEKSTRAK KUTIPAN VERBATIM...")
    kutipan = ekstrak_kutipan(teks)
    print(f"   Total Kutipan Ditemukan: {len(kutipan)}")
    
    # Analisis kata kunci
    print("\n🔑 KATA KUNCI PALING SERING MUNCUL:")
    kata_freq = analisis_kata_kunci(kutipan)
    top_10 = kata_freq.most_common(10)
    
    for i, (kata, freq) in enumerate(top_10, 1):
        print(f"   {i:2d}. {kata:20s} : {freq:3d}x")
    
    # Tampilkan beberapa kutipan
    print("\n📝 CONTOH KUTIPAN VERBATIM:")
    for i, kutipan_text in enumerate(kutipan[:5], 1):
        print(f"   {i}. \"{kutipan_text}\"")
    
    # Analisis tema
    print("\n🎯 ANALISIS TEMA:")
    tema_keywords = {
        'Meditasi': ['meditation', 'meditative', 'mindfulness'],
        'Kesehatan Mental': ['mental', 'clarity', 'balance', 'resilience'],
        'Spiritual': ['spiritual', 'healing', 'wellness'],
        'Retreat': ['retreat', 'sessions', 'program']
    }
    
    for tema, keywords in tema_keywords.items():
        count = sum(kata_freq.get(kw, 0) for kw in keywords)
        print(f"   {tema:20s} : {count:3d} kemunculan")
    
    print("\n" + "="*70)
    print("✅ ANALISIS SELESAI!")
    print("="*70)

if __name__ == "__main__":
    main()
