"""
Program Analisis Wacana Nostalgia & Identitas Generasi Z
Menganalisis artikel media massa tentang Olivia Rodrigo dan tren Y2K

Library yang digunakan:
- pandas: Manipulasi data tabular
- textblob: Sentiment analysis
- wordcloud: Visualisasi word cloud
- matplotlib: Plotting
- re: Regular expression untuk text cleaning
"""

# [EKSEKUSI-1] Import semua library yang dibutuhkan
import pandas as pd
from textblob import TextBlob
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.feature_extraction.text import TfidfVectorizer
from collections import Counter
import re

# [EKSEKUSI-2] Python membaca definisi class (belum dieksekusi)
class AnalisisWacanaNostalgia:
    # [EKSEKUSI-4] Method __init__ dipanggil saat instantiasi
    def __init__(self, file_path):
        """
        Inisialisasi analisis dengan membaca dan membersihkan data CSV
        
        Input: file_path (str) - Path ke file CSV
        Output: DataFrame yang sudah dibersihkan dengan kolom full_content
        
        Perubahan data:
        1. Baca CSV dengan separator ';' dan encoding latin-1/cp1252
        2. Hapus baris dengan Media Name atau Article Tittle kosong
        3. Gabungkan Article Tittle + Article Media + Precise Quote menjadi full_content
        4. Definisikan 3 set keywords untuk analisis
        
        Sintaks: pd.read_csv(), try/except, dropna(), str.strip(), fillna()
        """
        # [EKSEKUSI-4.1] Baca CSV dengan encoding yang sesuai
        try:
            self.df = pd.read_csv(file_path, sep=';', encoding='latin-1')
        except:
            self.df = pd.read_csv(file_path, sep=';', encoding='cp1252')
        
        # [EKSEKUSI-4.2] Hapus baris dengan nilai kosong pada kolom penting
        # Sintaks: dropna(subset=[]) untuk drop baris dengan null di kolom tertentu
        self.df = self.df.dropna(subset=['Media Name', 'Article Tittle'])
        # [EKSEKUSI-4.3] Boolean indexing untuk filter baris dengan Media Name tidak kosong
        self.df = self.df[self.df['Media Name'].str.strip() != '']
        
        # [EKSEKUSI-4.4] Gabungkan 3 kolom menjadi satu kolom full_content
        # Sintaks: fillna('') untuk ganti null dengan string kosong
        self.df['full_content'] = (
            self.df['Article Tittle'].fillna('') + ' ' + 
            self.df['Article Media'].fillna('') + ' ' + 
            self.df['Precise Quote'].fillna('')
        )
        
        # [EKSEKUSI-4.5] Definisikan keywords untuk 3 kategori analisis
        self.nostalgia_keywords = [
            'nostalgia', 'nostalgic', 'throwback', 'retro', 'vintage', 
            'early 2000s', '2000s', 'y2k', 'memories', 'childhood',
            'past', 'remember', 'reminiscent', 'flashback'
        ]
        
        self.genz_identity_keywords = [
            'generation z', 'gen z', 'genz', 'young', 'youth', 'teen',
            'millennial', 'generation', 'identity', 'culture', 'trend',
            'influence', 'social media', 'tiktok', 'instagram'
        ]
        
        self.fashion_memory_keywords = [
            'fashion', 'style', 'outfit', 'look', 'aesthetic', 'vibe',
            'chainmail', 'leather', 'boots', 'jeans', 'accessories'
        ]
    
    # [EKSEKUSI-13] Method ini dipanggil dari visualisasi_nostalgia_wordcloud()
    def preprocessing_text(self, text):
        """
        Membersihkan teks untuk analisis
        
        Input: text (str) - Teks mentah
        Output: String yang sudah dibersihkan (lowercase, tanpa karakter spesial)
        
        Perubahan data:
        1. Konversi ke lowercase
        2. Hapus karakter non-alphanumeric (kecuali spasi)
        3. Hapus spasi berlebih
        4. Hapus spasi di awal/akhir
        
        Sintaks: str(), .lower(), re.sub(), r'pattern', .strip()
        """
        text = str(text).lower()  # [EKSEKUSI-13.1] Konversi ke string dan lowercase
        # [EKSEKUSI-13.2] re.sub() untuk replace pattern dengan spasi
        # r'[^\w\s]' = regex untuk karakter selain word dan whitespace
        text = re.sub(r'[^\w\s]', ' ', text)
        # [EKSEKUSI-13.3] r'\s+' = regex untuk multiple whitespace
        text = re.sub(r'\s+', ' ', text)
        return text.strip()  # [EKSEKUSI-13.4] Hapus spasi di awal/akhir
    
    # [EKSEKUSI-6] Method ini dipanggil pertama dari laporan_analisis_wacana()
    def analisis_wacana_nostalgia(self):
        """
        Menghitung frekuensi kemunculan kata kunci nostalgia
        
        Input: DataFrame dengan kolom full_content
        Output: List tuple (keyword, count) diurutkan descending
        
        Perubahan data:
        1. Gabungkan semua teks menjadi satu string
        2. Konversi ke lowercase
        3. Hitung kemunculan setiap keyword
        4. Urutkan berdasarkan frekuensi tertinggi
        
        Sintaks: .join(), .astype(str), .count(), dict, sorted(), lambda, .items()
        """
        # [EKSEKUSI-6.1] Gabungkan semua konten artikel menjadi satu string
        # Sintaks: ' '.join() untuk gabung list string dengan separator spasi
        all_text = ' '.join(self.df['full_content'].astype(str)).lower()
        
        # [EKSEKUSI-6.2] Hitung frekuensi setiap keyword nostalgia
        nostalgia_analysis = {}
        for keyword in self.nostalgia_keywords:
            count = all_text.count(keyword)  # .count() menghitung substring
            nostalgia_analysis[keyword] = count
        
        # [EKSEKUSI-6.3] Urutkan berdasarkan count (descending)
        # Sintaks: sorted() dengan key=lambda untuk custom sorting
        # x[1] mengambil nilai count (index 1 dari tuple)
        return sorted(nostalgia_analysis.items(), key=lambda x: x[1], reverse=True)
    
    # [EKSEKUSI-7] Method ini dipanggil kedua dari laporan_analisis_wacana()
    def analisis_identitas_genz(self):
        """
        Menghitung frekuensi kata kunci identitas Generasi Z
        
        Input: DataFrame dengan kolom full_content
        Output: List tuple (keyword, count) diurutkan descending
        
        Proses sama dengan analisis_wacana_nostalgia() tapi untuk Gen Z keywords
        Sintaks: .join(), .count(), sorted(), lambda
        """
        all_text = ' '.join(self.df['full_content'].astype(str)).lower()
        
        genz_analysis = {}
        for keyword in self.genz_identity_keywords:
            count = all_text.count(keyword)
            genz_analysis[keyword] = count
        
        return sorted(genz_analysis.items(), key=lambda x: x[1], reverse=True)
    
    # [EKSEKUSI-8] Method ini dipanggil ketiga dari laporan_analisis_wacana()
    def analisis_memori_fashion(self):
        """
        Menghitung frekuensi elemen fashion sebagai penanda nostalgia
        
        Input: DataFrame dengan kolom full_content
        Output: List tuple (keyword, count) diurutkan descending
        
        Proses sama dengan analisis sebelumnya tapi untuk fashion keywords
        Sintaks: .join(), .count(), sorted(), lambda
        """
        all_text = ' '.join(self.df['full_content'].astype(str)).lower()
        
        fashion_analysis = {}
        for keyword in self.fashion_memory_keywords:
            count = all_text.count(keyword)
            fashion_analysis[keyword] = count
        
        return sorted(fashion_analysis.items(), key=lambda x: x[1], reverse=True)
    
    # [EKSEKUSI-9] Method ini dipanggil keempat dari laporan_analisis_wacana()
    def ekstrak_frasa_nostalgia(self):
        """
        Mengekstrak kalimat yang mengandung kata kunci nostalgia
        
        Input: DataFrame dengan kolom full_content
        Output: List 10 frasa unik yang mengandung keyword nostalgia
        
        Perubahan data:
        1. Iterasi setiap artikel
        2. Split artikel menjadi kalimat (delimiter: '.')
        3. Filter kalimat yang mengandung keyword nostalgia
        4. Batasi panjang frasa max 200 karakter, min 20 karakter
        5. Hapus duplikat dengan set()
        6. Ambil 10 frasa teratas
        
        Sintaks: nested for loop, .split(), if in, string slicing [:200], 
                 len(), list(set()), list slicing [:10]
        """
        nostalgia_phrases = []
        
        # [EKSEKUSI-9.1] Iterasi setiap konten artikel
        for content in self.df['full_content']:
            text = str(content).lower()
            # [EKSEKUSI-9.2] Split teks menjadi kalimat dengan delimiter '.'
            sentences = text.split('.')
            
            # [EKSEKUSI-9.3] Nested loop untuk cek setiap kalimat
            for sentence in sentences:
                # [EKSEKUSI-9.4] Cek apakah kalimat mengandung keyword nostalgia tertentu
                for keyword in ['nostalgia', 'nostalgic', 'throwback', 'retro', 'y2k']:
                    if keyword in sentence:
                        # [EKSEKUSI-9.5] String slicing [:200] untuk batasi panjang
                        clean_sentence = sentence.strip()[:200]
                        # Filter frasa minimal 20 karakter
                        if clean_sentence and len(clean_sentence) > 20:
                            nostalgia_phrases.append(clean_sentence)
        
        # [EKSEKUSI-9.6] list(set()) untuk hapus duplikat, [:10] ambil 10 teratas
        return list(set(nostalgia_phrases))[:10]
    
    # [EKSEKUSI-10] Method ini dipanggil kelima dari laporan_analisis_wacana()
    def analisis_sentimen_nostalgia(self):
        """
        Analisis sentimen artikel yang membahas nostalgia menggunakan TextBlob
        
        Input: DataFrame dengan kolom full_content
        Output: List dictionary berisi info sentimen per artikel
        
        Perubahan data:
        1. Filter artikel yang mengandung 5 keyword nostalgia utama
        2. Untuk setiap artikel relevan, hitung:
           - Polarity: -1 (negatif) hingga 1 (positif)
           - Subjectivity: 0 (objektif) hingga 1 (subjektif)
        3. Simpan metadata: index, media, title, sentiment, subjectivity
        
        Sintaks: enumerate(), any(), TextBlob(), .sentiment.polarity, 
                 .sentiment.subjectivity, .iloc[], string slicing [:100]
        """
        nostalgia_articles = []
        
        # [EKSEKUSI-10.1] enumerate() untuk iterasi dengan index
        for idx, content in enumerate(self.df['full_content']):
            text = str(content).lower()
            # [EKSEKUSI-10.2] any() cek apakah ada keyword yang cocok
            if any(keyword in text for keyword in self.nostalgia_keywords[:5]):
                # [EKSEKUSI-10.3] TextBlob untuk sentiment analysis
                blob = TextBlob(str(content))
                nostalgia_articles.append({
                    'index': idx,
                    # [EKSEKUSI-10.4] .iloc[] untuk akses baris by index
                    'media': self.df.iloc[idx]['Media Name'],
                    # [:100] batasi panjang title
                    'title': self.df.iloc[idx]['Article Tittle'][:100],
                    # [EKSEKUSI-10.5] .sentiment.polarity: nilai sentimen -1 hingga 1
                    'sentiment': blob.sentiment.polarity,
                    # .sentiment.subjectivity: nilai subjektivitas 0 hingga 1
                    'subjectivity': blob.sentiment.subjectivity
                })
        
        return nostalgia_articles
    
    # [EKSEKUSI-11] Method ini dipanggil keenam dari laporan_analisis_wacana()
    def pola_representasi_media(self):
        """
        Analisis bagaimana setiap media merepresentasikan nostalgia
        
        Input: DataFrame dengan kolom Media Name dan full_content
        Output: Dictionary dengan statistik per media
        
        Perubahan data:
        1. Untuk setiap media unik:
           - Filter artikel dari media tersebut
           - Gabungkan semua teks artikel
           - Hitung total sebutan nostalgia keywords
           - Hitung total sebutan Gen Z keywords
           - Hitung densitas nostalgia (sebutan per artikel)
        2. Return dictionary nested dengan struktur:
           {media_name: {nostalgia_mentions, genz_mentions, total_articles, nostalgia_density}}
        
        Sintaks: .unique(), boolean indexing df[df['col']==value], 
                 sum() dengan generator expression, conditional expression if/else inline
        """
        media_representation = {}
        
        # [EKSEKUSI-11.1] .unique() untuk mendapatkan nilai unik dari kolom
        for media in self.df['Media Name'].unique():
            # [EKSEKUSI-11.2] Boolean indexing untuk filter artikel dari media tertentu
            media_articles = self.df[self.df['Media Name'] == media]['full_content']
            all_media_text = ' '.join(media_articles.astype(str)).lower()
            
            # [EKSEKUSI-11.3] sum() dengan generator expression untuk hitung total kemunculan
            nostalgia_count = sum(all_media_text.count(kw) for kw in self.nostalgia_keywords)
            genz_count = sum(all_media_text.count(kw) for kw in self.genz_identity_keywords)
            
            # [EKSEKUSI-11.4] Nested dictionary untuk simpan statistik per media
            media_representation[media] = {
                'nostalgia_mentions': nostalgia_count,
                'genz_mentions': genz_count,
                'total_articles': len(media_articles),
                # Conditional expression inline: value_if_true if condition else value_if_false
                'nostalgia_density': nostalgia_count / len(media_articles) if len(media_articles) > 0 else 0
            }
        
        return media_representation
    
    # [EKSEKUSI-12] Method ini dipanggil ketujuh dari laporan_analisis_wacana()
    def visualisasi_nostalgia_wordcloud(self):
        """
        Membuat word cloud dari artikel yang membahas nostalgia
        
        Input: DataFrame dengan kolom full_content
        Output: Visualisasi matplotlib (tidak ada return value)
        
        Perubahan data:
        1. Filter artikel yang mengandung keyword nostalgia
        2. Gabungkan semua teks artikel relevan
        3. Bersihkan teks dengan preprocessing_text()
        4. Generate word cloud dengan parameter:
           - Ukuran: 1000x500 px
           - Background: putih
           - Colormap: plasma
           - Max words: 100
        5. Tampilkan dengan matplotlib
        
        Sintaks: list comprehension dengan conditional, WordCloud(), 
                 plt.figure(figsize=()), plt.imshow(), plt.axis(), 
                 plt.title(), plt.tight_layout(), plt.show()
        """
        # [EKSEKUSI-12.1] List comprehension dengan conditional untuk filter artikel
        nostalgia_texts = []
        for content in self.df['full_content']:
            text = str(content).lower()
            # any() untuk cek apakah ada keyword yang cocok
            if any(keyword in text for keyword in self.nostalgia_keywords):
                nostalgia_texts.append(text)
        
        if nostalgia_texts:
            # [EKSEKUSI-12.2] Gabungkan semua teks
            all_nostalgia_text = ' '.join(nostalgia_texts)
            # [EKSEKUSI-12.3] Panggil preprocessing_text() → LONCAT KE EKSEKUSI-13
            clean_text = self.preprocessing_text(all_nostalgia_text)
            
            # [EKSEKUSI-12.4] WordCloud() dengan named parameters
            wordcloud = WordCloud(
                width=1000, height=500,
                background_color='white',
                colormap='plasma',  # Skema warna
                max_words=100  # Maksimal 100 kata
            ).generate(clean_text)
            
            # [EKSEKUSI-12.5] plt.figure(figsize=()) untuk set ukuran plot
            plt.figure(figsize=(15, 8))
            # [EKSEKUSI-12.6] plt.imshow() untuk tampilkan image
            plt.imshow(wordcloud, interpolation='bilinear')
            # [EKSEKUSI-12.7] plt.axis('off') untuk sembunyikan axis
            plt.axis('off')
            # [EKSEKUSI-12.8] plt.title() untuk set judul dengan fontsize dan padding
            plt.title('Word Cloud: Wacana Nostalgia dalam Media', fontsize=18, pad=20)
            # [EKSEKUSI-12.9] plt.tight_layout() untuk optimasi layout
            plt.tight_layout()
            # [EKSEKUSI-12.10] plt.show() untuk tampilkan plot
            plt.show()
    
    # [EKSEKUSI-5] Method ini dipanggil dari main block
    def laporan_analisis_wacana(self):
        """
        Menghasilkan laporan lengkap analisis wacana
        
        Input: Semua data dari DataFrame
        Output: Dictionary dengan semua hasil analisis + print laporan ke console
        
        Perubahan data:
        1. Memanggil semua fungsi analisis
        2. Menghitung statistik agregat (total, rata-rata, rasio)
        3. Mencetak laporan terstruktur dengan 8 bagian
        4. Return dictionary dengan semua hasil
        
        Sintaks: print(), f-string, .nunique(), sum() dengan generator,
                 sorted() dengan lambda, conditional expression, string multiplication "="*70
        """
        # [EKSEKUSI-5.1] String multiplication untuk membuat garis pembatas
        print("=" * 70)
        print("ANALISIS WACANA: MEMORI NOSTALGIA & IDENTITAS GENERASI Z")
        print("Olivia Rodrigo & Tren Y2K dalam Media Massa")
        print("=" * 70)
        
        # [EKSEKUSI-5.2] f-string untuk interpolasi variabel dalam string
        # len() untuk hitung jumlah baris, .nunique() untuk hitung nilai unik
        print(f"📊 Total artikel dianalisis: {len(self.df)}")
        print(f"📰 Jumlah media: {self.df['Media Name'].nunique()}")
        
        print("\n" + "=" * 50)  # \n untuk newline
        print("1. ANALISIS WACANA NOSTALGIA")
        print("=" * 50)
        
        # [EKSEKUSI-5.3] Panggil fungsi analisis dan simpan hasilnya → LONCAT KE EKSEKUSI-6
        nostalgia_results = self.analisis_wacana_nostalgia()
        print("Frekuensi kata kunci nostalgia:")
        for keyword, count in nostalgia_results:
            if count > 0:
                # f-string dengan format untuk print
                print(f"  • {keyword}: {count} kali")
        
        print("\n" + "=" * 50)
        print("2. ANALISIS IDENTITAS GENERASI Z")
        print("=" * 50)
        
        # [EKSEKUSI-5.4] Panggil analisis Gen Z → LONCAT KE EKSEKUSI-7
        genz_results = self.analisis_identitas_genz()
        print("Frekuensi kata kunci identitas Gen Z:")
        for keyword, count in genz_results:
            if count > 0:
                print(f"  • {keyword}: {count} kali")
        
        print("\n" + "=" * 50)
        print("3. MEMORI FASHION SEBAGAI PENANDA NOSTALGIA")
        print("=" * 50)
        
        # [EKSEKUSI-5.5] Panggil analisis fashion → LONCAT KE EKSEKUSI-8
        fashion_results = self.analisis_memori_fashion()
        print("Elemen fashion yang menjadi penanda memori:")
        for keyword, count in fashion_results:
            if count > 0:
                print(f"  • {keyword}: {count} kali")
        
        print("\n" + "=" * 50)
        print("4. FRASA KUNCI NOSTALGIA")
        print("=" * 50)
        
        # [EKSEKUSI-5.6] Panggil ekstrak frasa → LONCAT KE EKSEKUSI-9
        phrases = self.ekstrak_frasa_nostalgia()
        print("Frasa yang mengandung wacana nostalgia:")
        # enumerate(start=1) untuk iterasi dengan index mulai dari 1
        for i, phrase in enumerate(phrases, 1):
            print(f"  {i}. \"{phrase}\"")
        
        print("\n" + "=" * 50)
        print("5. SENTIMEN ARTIKEL NOSTALGIA")
        print("=" * 50)
        
        # [EKSEKUSI-5.7] Panggil analisis sentimen → LONCAT KE EKSEKUSI-10
        nostalgia_sentiment = self.analisis_sentimen_nostalgia()
        if nostalgia_sentiment:
            # sum() dengan generator expression untuk hitung total
            avg_sentiment = sum(art['sentiment'] for art in nostalgia_sentiment) / len(nostalgia_sentiment)
            # .3f untuk format float dengan 3 desimal
            print(f"Rata-rata sentimen artikel nostalgia: {avg_sentiment:.3f}")
            print(f"Jumlah artikel dengan tema nostalgia: {len(nostalgia_sentiment)}")
            
            print("\nArtikel dengan sentimen tertinggi:")
            # sorted() dengan lambda untuk sort by sentiment descending
            sorted_sentiment = sorted(nostalgia_sentiment, key=lambda x: x['sentiment'], reverse=True)
            # List slicing [:3] untuk ambil 3 teratas
            for art in sorted_sentiment[:3]:
                print(f"  • {art['media']}: {art['title']} (sentimen: {art['sentiment']:.3f})")
        
        print("\n" + "=" * 50)
        print("6. POLA REPRESENTASI MEDIA")
        print("=" * 50)
        
        # [EKSEKUSI-5.8] Panggil pola representasi media → LONCAT KE EKSEKUSI-11
        media_patterns = self.pola_representasi_media()
        print("Bagaimana setiap media merepresentasikan nostalgia:")
        # sorted() dengan lambda untuk sort dictionary by value
        for media, data in sorted(media_patterns.items(), key=lambda x: x[1]['nostalgia_density'], reverse=True):
            print(f"  • {media}:")
            # .2f untuk format float dengan 2 desimal
            print(f"    - Densitas nostalgia: {data['nostalgia_density']:.2f}")
            print(f"    - Total sebutan nostalgia: {data['nostalgia_mentions']}")
            print(f"    - Total sebutan Gen Z: {data['genz_mentions']}")
        
        print("\n" + "=" * 50)
        print("7. TEMUAN KRITIS")
        print("=" * 50)
        
        # [EKSEKUSI-5.9] sum() dengan generator expression untuk hitung total dari tuple
        total_nostalgia = sum(count for _, count in nostalgia_results)
        total_genz = sum(count for _, count in genz_results)
        
        print(f"🔍 TEMUAN UTAMA:")
        print(f"  • Intensitas wacana nostalgia: {total_nostalgia} sebutan")
        print(f"  • Intensitas wacana identitas Gen Z: {total_genz} sebutan")
        # Conditional expression inline untuk handle division by zero
        print(f"  • Rasio nostalgia vs identitas: {total_nostalgia/total_genz:.2f}" if total_genz > 0 else "  • Rasio: Tidak dapat dihitung")
        
        # Akses tuple dengan indexing [0][1] untuk ambil count tertinggi
        if nostalgia_results[0][1] > 0:
            print(f"  • Kata nostalgia dominan: '{nostalgia_results[0][0]}' ({nostalgia_results[0][1]} kali)")
        
        print(f"\n💡 INTERPRETASI WACANA:")
        # Conditional untuk interpretasi berdasarkan perbandingan
        if total_nostalgia > total_genz:
            print("  • Media lebih fokus pada aspek nostalgia daripada identitas generasi")
            print("  • Olivia Rodrigo diposisikan sebagai simbol memori kolektif")
        else:
            print("  • Media lebih menekankan aspek identitas generasional")
            print("  • Olivia Rodrigo diposisikan sebagai representasi Gen Z")
        
        print("\n" + "=" * 50)
        print("8. VISUALISASI")
        print("=" * 50)
        # [EKSEKUSI-5.10] Panggil fungsi visualisasi → LONCAT KE EKSEKUSI-12
        self.visualisasi_nostalgia_wordcloud()
        
        # [EKSEKUSI-5.11] Return dictionary dengan semua hasil analisis
        return {
            'nostalgia_analysis': nostalgia_results,
            'genz_analysis': genz_results,
            'fashion_analysis': fashion_results,
            'sentiment_analysis': nostalgia_sentiment,
            'media_patterns': media_patterns
        }

# [EKSEKUSI-3] Main execution block
# if __name__ == "__main__": memastikan code hanya jalan saat file dieksekusi langsung
# (tidak jalan saat di-import sebagai module)
if __name__ == "__main__":
    # [EKSEKUSI-3.1] Instantiasi class dengan parameter file path → LONCAT KE EKSEKUSI-4
    analyzer = AnalisisWacanaNostalgia('data_mahatma.csv')
    # [EKSEKUSI-3.2] Panggil method untuk generate laporan lengkap → LONCAT KE EKSEKUSI-5
    # hasil berisi dictionary dengan semua hasil analisis
    hasil = analyzer.laporan_analisis_wacana()
