#!/usr/bin/env python3
"""
Script untuk mengekstrak data dari dokumen Word ke format CSV dan JSON
"""

from docx import Document
import csv
import json
import re

def ekstrak_artikel_dari_dokumen(file_path):
    """Ekstrak data artikel dari dokumen Word"""
    doc = Document(file_path)
    artikel_list = []
    
    current_artikel = {}
    mode = None
    
    for para in doc.paragraphs:
        text = para.text.strip()
        
        if text.startswith('Artikel '):
            # Simpan artikel sebelumnya jika ada
            if current_artikel:
                artikel_list.append(current_artikel)
            
            # Mulai artikel baru
            match = re.match(r'Artikel (\d+): (.+)', text)
            if match:
                current_artikel = {
                    'no': int(match.group(1)),
                    'judul': match.group(2),
                    'tanggal': '',
                    'media': '',
                    'link': '',
                    'ringkasan': '',
                    'kutipan': ''
                }
                mode = 'header'
        
        elif text.startswith('Tanggal:'):
            current_artikel['tanggal'] = text.replace('Tanggal:', '').strip()
        
        elif text.startswith('Media:'):
            current_artikel['media'] = text.replace('Media:', '').strip()
        
        elif text.startswith('Link:'):
            current_artikel['link'] = text.replace('Link:', '').strip()
        
        elif text == 'RINGKASAN ARTIKEL:':
            mode = 'ringkasan'
        
        elif text == 'KUTIPAN VERBATIM:':
            mode = 'kutipan'
        
        elif mode == 'ringkasan' and text and not text.startswith('_'):
            current_artikel['ringkasan'] = text
        
        elif mode == 'kutipan' and text and not text.startswith('_'):
            current_artikel['kutipan'] = text
    
    # Simpan artikel terakhir
    if current_artikel:
        artikel_list.append(current_artikel)
    
    return artikel_list

def simpan_ke_csv(artikel_list, output_file):
    """Simpan data artikel ke file CSV"""
    with open(output_file, 'w', newline='', encoding='utf-8') as f:
        fieldnames = ['no', 'judul', 'tanggal', 'media', 'link', 'ringkasan', 'kutipan']
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        
        writer.writeheader()
        for artikel in artikel_list:
            writer.writerow(artikel)

def simpan_ke_json(artikel_list, output_file):
    """Simpan data artikel ke file JSON"""
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(artikel_list, f, ensure_ascii=False, indent=2)

def main():
    file_path = 'Laporan_30_Artikel_Meditasi_Yogyakarta_2025_REVISI.docx'
    
    print("="*70)
    print("EKSTRAK DATA DARI DOKUMEN WORD")
    print("="*70)
    print(f"\nFile Input: {file_path}\n")
    
    # Ekstrak data
    print("📖 Mengekstrak data artikel...")
    artikel_list = ekstrak_artikel_dari_dokumen(file_path)
    print(f"   ✓ Berhasil mengekstrak {len(artikel_list)} artikel")
    
    # Simpan ke CSV
    csv_file = 'data_artikel_meditasi_yogyakarta.csv'
    print(f"\n💾 Menyimpan ke CSV: {csv_file}")
    simpan_ke_csv(artikel_list, csv_file)
    print(f"   ✓ File CSV berhasil dibuat")
    
    # Simpan ke JSON
    json_file = 'data_artikel_meditasi_yogyakarta.json'
    print(f"\n💾 Menyimpan ke JSON: {json_file}")
    simpan_ke_json(artikel_list, json_file)
    print(f"   ✓ File JSON berhasil dibuat")
    
    # Tampilkan contoh data
    print("\n📋 CONTOH DATA (Artikel 1):")
    if artikel_list:
        artikel = artikel_list[0]
        print(f"   No      : {artikel['no']}")
        print(f"   Judul   : {artikel['judul'][:50]}...")
        print(f"   Tanggal : {artikel['tanggal']}")
        print(f"   Media   : {artikel['media']}")
        print(f"   Kutipan : {artikel['kutipan'][:60]}...")
    
    print("\n" + "="*70)
    print("✅ EKSTRAK DATA SELESAI!")
    print("="*70)
    print("\nFile yang dihasilkan:")
    print(f"  1. {csv_file}")
    print(f"  2. {json_file}")

if __name__ == "__main__":
    main()
