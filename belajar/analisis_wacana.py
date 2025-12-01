import pandas as pd
from textblob import TextBlob
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.feature_extraction.text import TfidfVectorizer
from collections import Counter
import re

class AnalisisWacanaKritis:
    def __init__(self, file_path):
        """
        Analisis wacana kritis untuk data Olivia Rodrigo & Y2K di media massa
        """
        # Read CSV with semicolon separator and different encoding
        try:
            self.df = pd.read_csv(file_path, sep=';', encoding='latin-1')
        except:
            self.df = pd.read_csv(file_path, sep=';', encoding='cp1252')
        
        # Clean and prepare data
        self.df = self.df.dropna(subset=['Media Name', 'Article Tittle'])
        self.df = self.df[self.df['Media Name'].str.strip() != '']
        
        # Combine article content
        self.df['full_content'] = (
            self.df['Article Tittle'].fillna('') + ' ' + 
            self.df['Article Media'].fillna('') + ' ' + 
            self.df['Precise Quote'].fillna('')
        )
        
        self.keywords = ['olivia rodrigo', 'y2k', 'media massa', 'fashion', 'style', 'trend']
        
    def preprocessing_text(self, text):
        """Preprocessing teks"""
        text = str(text).lower()
        text = re.sub(r'[^\w\s]', '', text)
        return text
    
    def analisis_sentimen(self):
        """Analisis sentimen artikel"""
        sentiments = []
        for text in self.df['full_content']:
            blob = TextBlob(str(text))
            sentiments.append(blob.sentiment.polarity)
        
        self.df['sentimen'] = sentiments
        self.df['kategori_sentimen'] = pd.cut(self.df['sentimen'], 
                                            bins=[-1, -0.1, 0.1, 1], 
                                            labels=['Negatif', 'Netral', 'Positif'])
        return self.df
    
    def analisis_media(self):
        """Analisis berdasarkan media"""
        media_analysis = self.df.groupby('Media Name').agg({
            'sentimen': ['mean', 'count'],
            'Article Tittle': 'count'
        }).round(3)
        
        return media_analysis
    
    def frekuensi_kata_kunci(self):
        """Analisis frekuensi kata kunci"""
        all_text = ' '.join(self.df['full_content'].astype(str))
        all_text = self.preprocessing_text(all_text)
        
        keyword_count = {}
        for keyword in self.keywords:
            count = all_text.count(keyword.lower())
            keyword_count[keyword] = count
        
        return keyword_count
    
    def analisis_y2k_discourse(self):
        """Analisis khusus wacana Y2K"""
        y2k_terms = ['y2k', '2000s', 'nostalgic', 'throwback', 'retro', 'vintage', 'early 2000s']
        
        y2k_mentions = {}
        all_text = ' '.join(self.df['full_content'].astype(str)).lower()
        
        for term in y2k_terms:
            count = all_text.count(term)
            y2k_mentions[term] = count
            
        return y2k_mentions
    
    def tfidf_analysis(self):
        """Analisis TF-IDF untuk kata penting"""
        corpus = [self.preprocessing_text(text) for text in self.df['full_content']]
        
        vectorizer = TfidfVectorizer(max_features=20, stop_words='english')
        tfidf_matrix = vectorizer.fit_transform(corpus)
        
        feature_names = vectorizer.get_feature_names_out()
        tfidf_scores = tfidf_matrix.sum(axis=0).A1
        
        word_scores = dict(zip(feature_names, tfidf_scores))
        return sorted(word_scores.items(), key=lambda x: x[1], reverse=True)
    
    def visualisasi_wordcloud(self):
        """Membuat word cloud"""
        all_text = ' '.join(self.df['full_content'].astype(str))
        all_text = self.preprocessing_text(all_text)
        
        wordcloud = WordCloud(width=800, height=400, 
                            background_color='white',
                            colormap='viridis').generate(all_text)
        
        plt.figure(figsize=(12, 6))
        plt.imshow(wordcloud, interpolation='bilinear')
        plt.axis('off')
        plt.title('Word Cloud: Olivia Rodrigo & Y2K dalam Media Massa', fontsize=16)
        plt.tight_layout()
        plt.savefig('wordcloud.png', dpi=300, bbox_inches='tight')
        plt.close()
        print("✓ Word cloud disimpan ke: wordcloud.png")
    
    def visualisasi_sentimen_media(self):
        """Visualisasi sentimen per media"""
        media_sentiment = self.df.groupby('Media Name')['sentimen'].mean().sort_values(ascending=False)
        
        plt.figure(figsize=(12, 6))
        media_sentiment.plot(kind='bar', color='skyblue')
        plt.title('Rata-rata Sentimen per Media', fontsize=14)
        plt.xlabel('Media')
        plt.ylabel('Skor Sentimen')
        plt.xticks(rotation=45, ha='right')
        plt.tight_layout()
        plt.savefig('sentimen_media.png', dpi=300, bbox_inches='tight')
        plt.close()
        print("✓ Grafik sentimen disimpan ke: sentimen_media.png")
        
        return media_sentiment
    
    def laporan_lengkap(self):
        """Generate laporan lengkap analisis wacana kritis"""
        print("=" * 60)
        print("ANALISIS WACANA KRITIS: OLIVIA RODRIGO & Y2K DI MEDIA MASSA")
        print("=" * 60)
        print(f"Total artikel: {len(self.df)}")
        print(f"Jumlah media: {self.df['Media Name'].nunique()}")
        print(f"Keywords: {', '.join(self.keywords)}")
        
        print("\n" + "=" * 40)
        print("1. ANALISIS SENTIMEN KESELURUHAN")
        print("=" * 40)
        
        # Sentimen
        sentimen_df = self.analisis_sentimen()
        sentimen_counts = sentimen_df['kategori_sentimen'].value_counts()
        print(sentimen_counts)
        print(f"\nRata-rata sentimen: {sentimen_df['sentimen'].mean():.3f}")
        
        print("\n" + "=" * 40)
        print("2. ANALISIS PER MEDIA")
        print("=" * 40)
        media_analysis = self.analisis_media()
        print(media_analysis.head(10))
        
        print("\n" + "=" * 40)
        print("3. FREKUENSI KATA KUNCI")
        print("=" * 40)
        keyword_freq = self.frekuensi_kata_kunci()
        for keyword, count in keyword_freq.items():
            print(f"- {keyword}: {count} kali")
        
        print("\n" + "=" * 40)
        print("4. ANALISIS WACANA Y2K")
        print("=" * 40)
        y2k_analysis = self.analisis_y2k_discourse()
        for term, count in sorted(y2k_analysis.items(), key=lambda x: x[1], reverse=True):
            if count > 0:
                print(f"- {term}: {count} kali")
        
        print("\n" + "=" * 40)
        print("5. KATA PENTING (TF-IDF)")
        print("=" * 40)
        tfidf_results = self.tfidf_analysis()
        for word, score in tfidf_results[:15]:
            print(f"- {word}: {score:.3f}")
        
        print("\n" + "=" * 40)
        print("6. TEMUAN KRITIS")
        print("=" * 40)
        
        # Critical findings
        total_positive = len(sentimen_df[sentimen_df['kategori_sentimen'] == 'Positif'])
        total_negative = len(sentimen_df[sentimen_df['kategori_sentimen'] == 'Negatif'])
        
        print(f"• Dominasi sentimen: {'Positif' if total_positive > total_negative else 'Negatif'}")
        print(f"• Media dengan sentimen tertinggi: {media_analysis.iloc[0].name if len(media_analysis) > 0 else 'N/A'}")
        print(f"• Fokus wacana Y2K: {max(y2k_analysis.items(), key=lambda x: x[1])[0] if any(y2k_analysis.values()) else 'Tidak ditemukan'}")
        
        # Visualizations
        print("\n" + "=" * 40)
        print("7. VISUALISASI")
        print("=" * 40)
        self.visualisasi_wordcloud()
        self.visualisasi_sentimen_media()
        
        return sentimen_df

# Jalankan analisis
if __name__ == "__main__":
    analyzer = AnalisisWacanaKritis('data_mahatma.csv')
    hasil = analyzer.laporan_lengkap()
