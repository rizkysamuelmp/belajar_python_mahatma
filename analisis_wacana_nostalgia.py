import pandas as pd
from textblob import TextBlob
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.feature_extraction.text import TfidfVectorizer
from collections import Counter
import re

class AnalisisWacanaNostalgia:
    def __init__(self, file_path):
        """Analisis Wacana: Memori Nostalgia & Identitas Generasi Z"""
        try:
            self.df = pd.read_csv(file_path, sep=';', encoding='latin-1')
        except:
            self.df = pd.read_csv(file_path, sep=';', encoding='cp1252')
        
        # Clean data
        self.df = self.df.dropna(subset=['Media Name', 'Article Tittle'])
        self.df = self.df[self.df['Media Name'].str.strip() != '']
        
        # Combine content
        self.df['full_content'] = (
            self.df['Article Tittle'].fillna('') + ' ' + 
            self.df['Article Media'].fillna('') + ' ' + 
            self.df['Precise Quote'].fillna('')
        )
        
        # Keywords untuk analisis wacana nostalgia & identitas Gen Z
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
    
    def preprocessing_text(self, text):
        """Clean text for analysis"""
        text = str(text).lower()
        text = re.sub(r'[^\w\s]', ' ', text)
        text = re.sub(r'\s+', ' ', text)
        return text.strip()
    
    def analisis_wacana_nostalgia(self):
        """Analisis khusus wacana nostalgia"""
        all_text = ' '.join(self.df['full_content'].astype(str)).lower()
        
        nostalgia_analysis = {}
        for keyword in self.nostalgia_keywords:
            count = all_text.count(keyword)
            nostalgia_analysis[keyword] = count
        
        return sorted(nostalgia_analysis.items(), key=lambda x: x[1], reverse=True)
    
    def analisis_identitas_genz(self):
        """Analisis wacana identitas Generasi Z"""
        all_text = ' '.join(self.df['full_content'].astype(str)).lower()
        
        genz_analysis = {}
        for keyword in self.genz_identity_keywords:
            count = all_text.count(keyword)
            genz_analysis[keyword] = count
        
        return sorted(genz_analysis.items(), key=lambda x: x[1], reverse=True)
    
    def analisis_memori_fashion(self):
        """Analisis memori fashion sebagai penanda nostalgia"""
        all_text = ' '.join(self.df['full_content'].astype(str)).lower()
        
        fashion_analysis = {}
        for keyword in self.fashion_memory_keywords:
            count = all_text.count(keyword)
            fashion_analysis[keyword] = count
        
        return sorted(fashion_analysis.items(), key=lambda x: x[1], reverse=True)
    
    def ekstrak_frasa_nostalgia(self):
        """Ekstrak frasa yang mengandung kata nostalgia"""
        nostalgia_phrases = []
        
        for content in self.df['full_content']:
            text = str(content).lower()
            sentences = text.split('.')
            
            for sentence in sentences:
                for keyword in ['nostalgia', 'nostalgic', 'throwback', 'retro', 'y2k']:
                    if keyword in sentence:
                        clean_sentence = sentence.strip()[:200]  # Limit length
                        if clean_sentence and len(clean_sentence) > 20:
                            nostalgia_phrases.append(clean_sentence)
        
        return list(set(nostalgia_phrases))[:10]  # Top 10 unique phrases
    
    def analisis_sentimen_nostalgia(self):
        """Analisis sentimen khusus artikel yang membahas nostalgia"""
        nostalgia_articles = []
        
        for idx, content in enumerate(self.df['full_content']):
            text = str(content).lower()
            if any(keyword in text for keyword in self.nostalgia_keywords[:5]):
                blob = TextBlob(str(content))
                nostalgia_articles.append({
                    'index': idx,
                    'media': self.df.iloc[idx]['Media Name'],
                    'title': self.df.iloc[idx]['Article Tittle'][:100],
                    'sentiment': blob.sentiment.polarity,
                    'subjectivity': blob.sentiment.subjectivity
                })
        
        return nostalgia_articles
    
    def pola_representasi_media(self):
        """Analisis bagaimana media merepresentasikan nostalgia"""
        media_representation = {}
        
        for media in self.df['Media Name'].unique():
            media_articles = self.df[self.df['Media Name'] == media]['full_content']
            all_media_text = ' '.join(media_articles.astype(str)).lower()
            
            nostalgia_count = sum(all_media_text.count(kw) for kw in self.nostalgia_keywords)
            genz_count = sum(all_media_text.count(kw) for kw in self.genz_identity_keywords)
            
            media_representation[media] = {
                'nostalgia_mentions': nostalgia_count,
                'genz_mentions': genz_count,
                'total_articles': len(media_articles),
                'nostalgia_density': nostalgia_count / len(media_articles) if len(media_articles) > 0 else 0
            }
        
        return media_representation
    
    def visualisasi_nostalgia_wordcloud(self):
        """Word cloud khusus untuk kata-kata nostalgia"""
        # Filter text yang mengandung kata nostalgia
        nostalgia_texts = []
        for content in self.df['full_content']:
            text = str(content).lower()
            if any(keyword in text for keyword in self.nostalgia_keywords):
                nostalgia_texts.append(text)
        
        if nostalgia_texts:
            all_nostalgia_text = ' '.join(nostalgia_texts)
            clean_text = self.preprocessing_text(all_nostalgia_text)
            
            wordcloud = WordCloud(
                width=1000, height=500,
                background_color='white',
                colormap='plasma',
                max_words=100
            ).generate(clean_text)
            
            plt.figure(figsize=(15, 8))
            plt.imshow(wordcloud, interpolation='bilinear')
            plt.axis('off')
            plt.title('Word Cloud: Wacana Nostalgia dalam Media', fontsize=18, pad=20)
            plt.tight_layout()
            plt.show()
    
    def laporan_analisis_wacana(self):
        """Laporan lengkap analisis wacana nostalgia & identitas Gen Z"""
        print("=" * 70)
        print("ANALISIS WACANA: MEMORI NOSTALGIA & IDENTITAS GENERASI Z")
        print("Olivia Rodrigo & Tren Y2K dalam Media Massa")
        print("=" * 70)
        
        print(f"📊 Total artikel dianalisis: {len(self.df)}")
        print(f"📰 Jumlah media: {self.df['Media Name'].nunique()}")
        
        print("\n" + "=" * 50)
        print("1. ANALISIS WACANA NOSTALGIA")
        print("=" * 50)
        
        nostalgia_results = self.analisis_wacana_nostalgia()
        print("Frekuensi kata kunci nostalgia:")
        for keyword, count in nostalgia_results:
            if count > 0:
                print(f"  • {keyword}: {count} kali")
        
        print("\n" + "=" * 50)
        print("2. ANALISIS IDENTITAS GENERASI Z")
        print("=" * 50)
        
        genz_results = self.analisis_identitas_genz()
        print("Frekuensi kata kunci identitas Gen Z:")
        for keyword, count in genz_results:
            if count > 0:
                print(f"  • {keyword}: {count} kali")
        
        print("\n" + "=" * 50)
        print("3. MEMORI FASHION SEBAGAI PENANDA NOSTALGIA")
        print("=" * 50)
        
        fashion_results = self.analisis_memori_fashion()
        print("Elemen fashion yang menjadi penanda memori:")
        for keyword, count in fashion_results:
            if count > 0:
                print(f"  • {keyword}: {count} kali")
        
        print("\n" + "=" * 50)
        print("4. FRASA KUNCI NOSTALGIA")
        print("=" * 50)
        
        phrases = self.ekstrak_frasa_nostalgia()
        print("Frasa yang mengandung wacana nostalgia:")
        for i, phrase in enumerate(phrases, 1):
            print(f"  {i}. \"{phrase}\"")
        
        print("\n" + "=" * 50)
        print("5. SENTIMEN ARTIKEL NOSTALGIA")
        print("=" * 50)
        
        nostalgia_sentiment = self.analisis_sentimen_nostalgia()
        if nostalgia_sentiment:
            avg_sentiment = sum(art['sentiment'] for art in nostalgia_sentiment) / len(nostalgia_sentiment)
            print(f"Rata-rata sentimen artikel nostalgia: {avg_sentiment:.3f}")
            print(f"Jumlah artikel dengan tema nostalgia: {len(nostalgia_sentiment)}")
            
            print("\nArtikel dengan sentimen tertinggi:")
            sorted_sentiment = sorted(nostalgia_sentiment, key=lambda x: x['sentiment'], reverse=True)
            for art in sorted_sentiment[:3]:
                print(f"  • {art['media']}: {art['title']} (sentimen: {art['sentiment']:.3f})")
        
        print("\n" + "=" * 50)
        print("6. POLA REPRESENTASI MEDIA")
        print("=" * 50)
        
        media_patterns = self.pola_representasi_media()
        print("Bagaimana setiap media merepresentasikan nostalgia:")
        for media, data in sorted(media_patterns.items(), key=lambda x: x[1]['nostalgia_density'], reverse=True):
            print(f"  • {media}:")
            print(f"    - Densitas nostalgia: {data['nostalgia_density']:.2f}")
            print(f"    - Total sebutan nostalgia: {data['nostalgia_mentions']}")
            print(f"    - Total sebutan Gen Z: {data['genz_mentions']}")
        
        print("\n" + "=" * 50)
        print("7. TEMUAN KRITIS")
        print("=" * 50)
        
        total_nostalgia = sum(count for _, count in nostalgia_results)
        total_genz = sum(count for _, count in genz_results)
        
        print(f"🔍 TEMUAN UTAMA:")
        print(f"  • Intensitas wacana nostalgia: {total_nostalgia} sebutan")
        print(f"  • Intensitas wacana identitas Gen Z: {total_genz} sebutan")
        print(f"  • Rasio nostalgia vs identitas: {total_nostalgia/total_genz:.2f}" if total_genz > 0 else "  • Rasio: Tidak dapat dihitung")
        
        if nostalgia_results[0][1] > 0:
            print(f"  • Kata nostalgia dominan: '{nostalgia_results[0][0]}' ({nostalgia_results[0][1]} kali)")
        
        print(f"\n💡 INTERPRETASI WACANA:")
        if total_nostalgia > total_genz:
            print("  • Media lebih fokus pada aspek nostalgia daripada identitas generasi")
            print("  • Olivia Rodrigo diposisikan sebagai simbol memori kolektif")
        else:
            print("  • Media lebih menekankan aspek identitas generasional")
            print("  • Olivia Rodrigo diposisikan sebagai representasi Gen Z")
        
        print("\n" + "=" * 50)
        print("8. VISUALISASI")
        print("=" * 50)
        self.visualisasi_nostalgia_wordcloud()
        
        return {
            'nostalgia_analysis': nostalgia_results,
            'genz_analysis': genz_results,
            'fashion_analysis': fashion_results,
            'sentiment_analysis': nostalgia_sentiment,
            'media_patterns': media_patterns
        }

# Jalankan analisis
if __name__ == "__main__":
    analyzer = AnalisisWacanaNostalgia('data_mahatma.csv')
    hasil = analyzer.laporan_analisis_wacana()
