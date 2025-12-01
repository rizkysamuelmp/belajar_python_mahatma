import matplotlib.pyplot as plt
from analisis_wacana import AnalisisWacanaKritis

# Load data
analyzer = AnalisisWacanaKritis('data_mahatma.csv')
analyzer.analisis_sentimen()

# 1. Diagram batang kategori sentimen
sentimen_counts = analyzer.df['kategori_sentimen'].value_counts()
plt.figure(figsize=(10, 6))
sentimen_counts.plot(kind='bar', color=['green', 'gray', 'red'])
plt.title('Distribusi Kategori Sentimen', fontsize=14)
plt.xlabel('Kategori Sentimen')
plt.ylabel('Jumlah Artikel')
plt.xticks(rotation=0)
plt.tight_layout()
plt.savefig('diagram_kategori_sentimen.png', dpi=300, bbox_inches='tight')
plt.close()
print("✓ Diagram kategori sentimen disimpan ke: diagram_kategori_sentimen.png")

# 2. Diagram batang jumlah artikel per media
artikel_per_media = analyzer.df['Media Name'].value_counts().head(10)
plt.figure(figsize=(12, 6))
artikel_per_media.plot(kind='bar', color='steelblue')
plt.title('10 Media dengan Artikel Terbanyak', fontsize=14)
plt.xlabel('Media')
plt.ylabel('Jumlah Artikel')
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
plt.savefig('diagram_artikel_per_media.png', dpi=300, bbox_inches='tight')
plt.close()
print("✓ Diagram artikel per media disimpan ke: diagram_artikel_per_media.png")

# 3. Diagram batang frekuensi kata kunci
keyword_freq = analyzer.frekuensi_kata_kunci()
plt.figure(figsize=(10, 6))
plt.bar(keyword_freq.keys(), keyword_freq.values(), color='coral')
plt.title('Frekuensi Kata Kunci', fontsize=14)
plt.xlabel('Kata Kunci')
plt.ylabel('Frekuensi')
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
plt.savefig('diagram_kata_kunci.png', dpi=300, bbox_inches='tight')
plt.close()
print("✓ Diagram kata kunci disimpan ke: diagram_kata_kunci.png")

print("\nSemua diagram berhasil dibuat!")
