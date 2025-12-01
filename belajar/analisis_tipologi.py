import pandas as pd
from textblob import TextBlob
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.cluster import KMeans
from sklearn.decomposition import PCA
import matplotlib.pyplot as plt
import seaborn as sns
from collections import Counter
import re
import numpy as np

class AnalisisTipologi:
    def __init__(self, file_path):
        """Analisis Tipologi berdasarkan kata dominan dan tema"""
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
        
        # Define themes
        self.themes = {
            'Fashion & Style': ['fashion', 'style', 'outfit', 'look', 'wear', 'dress', 'clothes', 'aesthetic', 'chainmail', 'leather', 'boots', 'jeans'],
            'Nostalgia & Memory': ['nostalgia', 'nostalgic', 'throwback', 'retro', 'vintage', 'y2k', '2000s', 'memories', 'past', 'remember'],
            'Identity & Generation': ['generation', 'gen z', 'youth', 'young', 'identity', 'culture', 'millennial', 'teen', 'influence'],
            'Media & Performance': ['performance', 'stage', 'concert', 'festival', 'music', 'singer', 'artist', 'show', 'event'],
            'Social & Cultural': ['social', 'cultural', 'trend', 'popular', 'iconic', 'symbol', 'represent', 'influence', 'impact']
        }
    
    def preprocessing_text(self, text):
        """Clean text"""
        text = str(text).lower()
        text = re.sub(r'[^\w\s]', ' ', text)
        text = re.sub(r'\s+', ' ', text)
        return text.strip()
    
    def ekstrak_kata_dominan(self, n_words=50):
        """Ekstrak kata-kata dominan menggunakan TF-IDF"""
        corpus = [self.preprocessing_text(text) for text in self.df['full_content']]
        
        vectorizer = TfidfVectorizer(
            max_features=n_words,
            stop_words='english',
            min_df=2,
            ngram_range=(1, 2)
        )
        
        tfidf_matrix = vectorizer.fit_transform(corpus)
        feature_names = vectorizer.get_feature_names_out()
        tfidf_scores = tfidf_matrix.sum(axis=0).A1
        
        word_scores = list(zip(feature_names, tfidf_scores))
        word_scores.sort(key=lambda x: x[1], reverse=True)
        
        return word_scores
    
    def klasifikasi_tema_artikel(self):
        """Klasifikasi artikel berdasarkan tema dominan"""
        artikel_tema = []
        
        for idx, content in enumerate(self.df['full_content']):
            text = self.preprocessing_text(content)
            tema_scores = {}
            
            for tema, keywords in self.themes.items():
                score = sum(text.count(keyword) for keyword in keywords)
                tema_scores[tema] = score
            
            # Tentukan tema dominan
            tema_dominan = max(tema_scores.items(), key=lambda x: x[1])
            
            artikel_tema.append({
                'index': idx,
                'media': self.df.iloc[idx]['Media Name'],
                'title': self.df.iloc[idx]['Article Tittle'][:80],
                'tema_dominan': tema_dominan[0],
                'skor_tema': tema_dominan[1],
                'semua_skor': tema_scores
            })
        
        return artikel_tema
    
    def tipologi_media(self):
        """Buat tipologi media berdasarkan fokus tema"""
        media_tipologi = {}
        
        for media in self.df['Media Name'].unique():
            media_articles = self.df[self.df['Media Name'] == media]['full_content']
            all_text = ' '.join(media_articles.astype(str))
            clean_text = self.preprocessing_text(all_text)
            
            tema_scores = {}
            for tema, keywords in self.themes.items():
                score = sum(clean_text.count(keyword) for keyword in keywords)
                tema_scores[tema] = score
            
            # Normalisasi berdasarkan jumlah artikel
            total_articles = len(media_articles)
            normalized_scores = {tema: score/total_articles for tema, score in tema_scores.items()}
            
            # Tentukan tipologi
            dominant_theme = max(normalized_scores.items(), key=lambda x: x[1])
            
            media_tipologi[media] = {
                'tipologi': dominant_theme[0],
                'skor_dominan': dominant_theme[1],
                'distribusi_tema': normalized_scores,
                'total_artikel': total_articles
            }
        
        return media_tipologi
    
    def clustering_artikel(self, n_clusters=4):
        """Clustering artikel berdasarkan kesamaan konten"""
        corpus = [self.preprocessing_text(text) for text in self.df['full_content']]
        
        vectorizer = TfidfVectorizer(max_features=100, stop_words='english', min_df=2)
        tfidf_matrix = vectorizer.fit_transform(corpus)
        
        # K-means clustering
        kmeans = KMeans(n_clusters=n_clusters, random_state=42)
        clusters = kmeans.fit_predict(tfidf_matrix)
        
        # PCA untuk visualisasi
        pca = PCA(n_components=2)
        pca_result = pca.fit_transform(tfidf_matrix.toarray())
        
        # Analisis cluster
        cluster_analysis = {}
        for i in range(n_clusters):
            cluster_indices = np.where(clusters == i)[0]
            cluster_texts = [corpus[idx] for idx in cluster_indices]
            
            # Kata dominan per cluster
            cluster_vectorizer = TfidfVectorizer(max_features=10, stop_words='english')
            cluster_tfidf = cluster_vectorizer.fit_transform(cluster_texts)
            cluster_words = cluster_vectorizer.get_feature_names_out()
            
            cluster_analysis[f'Cluster {i+1}'] = {
                'jumlah_artikel': len(cluster_indices),
                'kata_dominan': list(cluster_words),
                'artikel_sample': [self.df.iloc[idx]['Article Tittle'][:60] for idx in cluster_indices[:3]]
            }
        
        return clusters, pca_result, cluster_analysis
    
    def analisis_evolusi_tema(self):
        """Analisis evolusi tema berdasarkan waktu (jika ada data tanggal)"""
        if 'Date' in self.df.columns:
            try:
                self.df['Date_parsed'] = pd.to_datetime(self.df['Date'], errors='coerce')
                self.df = self.df.dropna(subset=['Date_parsed'])
                
                # Group by month
                monthly_themes = {}
                for month in self.df['Date_parsed'].dt.to_period('M').unique():
                    month_articles = self.df[self.df['Date_parsed'].dt.to_period('M') == month]
                    month_text = ' '.join(month_articles['full_content'].astype(str))
                    clean_text = self.preprocessing_text(month_text)
                    
                    tema_scores = {}
                    for tema, keywords in self.themes.items():
                        score = sum(clean_text.count(keyword) for keyword in keywords)
                        tema_scores[tema] = score
                    
                    monthly_themes[str(month)] = tema_scores
                
                return monthly_themes
            except:
                return None
        return None
    
    def visualisasi_tipologi(self, clusters, pca_result):
        """Visualisasi hasil clustering dan tipologi"""
        fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(15, 12))
        
        # 1. Scatter plot clustering
        scatter = ax1.scatter(pca_result[:, 0], pca_result[:, 1], c=clusters, cmap='viridis', alpha=0.7)
        ax1.set_title('Clustering Artikel berdasarkan Konten')
        ax1.set_xlabel('PCA Component 1')
        ax1.set_ylabel('PCA Component 2')
        plt.colorbar(scatter, ax=ax1)
        
        # 2. Distribusi tema per media
        media_tipologi = self.tipologi_media()
        media_names = list(media_tipologi.keys())[:10]  # Top 10 media
        tema_counts = {}
        
        for tema in self.themes.keys():
            tema_counts[tema] = sum(1 for media in media_names 
                                  if media_tipologi[media]['tipologi'] == tema)
        
        ax2.bar(tema_counts.keys(), tema_counts.values(), color='skyblue')
        ax2.set_title('Distribusi Tipologi Media')
        ax2.set_xlabel('Tema Dominan')
        ax2.set_ylabel('Jumlah Media')
        ax2.tick_params(axis='x', rotation=45)
        
        # 3. Kata dominan
        kata_dominan = self.ekstrak_kata_dominan(15)
        words, scores = zip(*kata_dominan)
        
        ax3.barh(words[:10], scores[:10], color='lightcoral')
        ax3.set_title('Top 10 Kata Dominan (TF-IDF)')
        ax3.set_xlabel('Skor TF-IDF')
        
        # 4. Distribusi cluster
        cluster_counts = Counter(clusters)
        ax4.pie(cluster_counts.values(), labels=[f'Cluster {i+1}' for i in cluster_counts.keys()], 
                autopct='%1.1f%%', startangle=90)
        ax4.set_title('Distribusi Artikel per Cluster')
        
        plt.tight_layout()
        plt.show()
    
    def laporan_tipologi_lengkap(self):
        """Laporan lengkap analisis tipologi"""
        print("=" * 70)
        print("ANALISIS TIPOLOGI: KATA DOMINAN & TEMA")
        print("Olivia Rodrigo & Y2K dalam Media Massa")
        print("=" * 70)
        
        print(f"📊 Total artikel: {len(self.df)}")
        print(f"📰 Total media: {self.df['Media Name'].nunique()}")
        
        print("\n" + "=" * 50)
        print("1. KATA-KATA DOMINAN (TF-IDF)")
        print("=" * 50)
        
        kata_dominan = self.ekstrak_kata_dominan(20)
        for i, (word, score) in enumerate(kata_dominan, 1):
            print(f"{i:2d}. {word:<20} | Skor: {score:.3f}")
        
        print("\n" + "=" * 50)
        print("2. KLASIFIKASI TEMA ARTIKEL")
        print("=" * 50)
        
        artikel_tema = self.klasifikasi_tema_artikel()
        tema_distribusi = Counter(art['tema_dominan'] for art in artikel_tema)
        
        print("Distribusi artikel per tema:")
        for tema, count in tema_distribusi.most_common():
            percentage = (count / len(artikel_tema)) * 100
            print(f"  • {tema:<25} | {count:3d} artikel ({percentage:.1f}%)")
        
        print("\n" + "=" * 50)
        print("3. TIPOLOGI MEDIA")
        print("=" * 50)
        
        media_tipologi = self.tipologi_media()
        print("Tipologi media berdasarkan fokus tema:")
        
        for media, data in sorted(media_tipologi.items(), key=lambda x: x[1]['skor_dominan'], reverse=True):
            print(f"\n📰 {media}")
            print(f"   Tipologi: {data['tipologi']}")
            print(f"   Skor dominan: {data['skor_dominan']:.2f}")
            print(f"   Total artikel: {data['total_artikel']}")
        
        print("\n" + "=" * 50)
        print("4. CLUSTERING ARTIKEL")
        print("=" * 50)
        
        clusters, pca_result, cluster_analysis = self.clustering_artikel()
        
        print("Hasil clustering artikel:")
        for cluster_name, data in cluster_analysis.items():
            print(f"\n🔍 {cluster_name}")
            print(f"   Jumlah artikel: {data['jumlah_artikel']}")
            print(f"   Kata dominan: {', '.join(data['kata_dominan'])}")
            print(f"   Contoh artikel:")
            for sample in data['artikel_sample']:
                print(f"     - {sample}")
        
        print("\n" + "=" * 50)
        print("5. EVOLUSI TEMA")
        print("=" * 50)
        
        evolusi = self.analisis_evolusi_tema()
        if evolusi:
            print("Evolusi tema per periode:")
            for periode, tema_scores in evolusi.items():
                dominant_tema = max(tema_scores.items(), key=lambda x: x[1])
                print(f"  📅 {periode}: {dominant_tema[0]} (skor: {dominant_tema[1]})")
        else:
            print("Data tanggal tidak tersedia untuk analisis evolusi")
        
        print("\n" + "=" * 50)
        print("6. TEMUAN TIPOLOGI")
        print("=" * 50)
        
        # Temuan utama
        tema_terpopuler = tema_distribusi.most_common(1)[0]
        media_terfokus = max(media_tipologi.items(), key=lambda x: x[1]['skor_dominan'])
        
        print("🔍 TEMUAN UTAMA:")
        print(f"  • Tema dominan: {tema_terpopuler[0]} ({tema_terpopuler[1]} artikel)")
        print(f"  • Media paling terfokus: {media_terfokus[0]} ({media_terfokus[1]['tipologi']})")
        print(f"  • Jumlah cluster optimal: {len(cluster_analysis)}")
        
        print(f"\n💡 INTERPRETASI TIPOLOGI:")
        if tema_terpopuler[0] == 'Fashion & Style':
            print("  • Media massa memposisikan Olivia Rodrigo sebagai ikon fashion")
            print("  • Fokus pada aspek visual dan estetika Y2K")
        elif tema_terpopuler[0] == 'Nostalgia & Memory':
            print("  • Media menekankan aspek nostalgia dan memori kolektif")
            print("  • Y2K dipandang sebagai fenomena retrospektif")
        
        print("\n" + "=" * 50)
        print("7. VISUALISASI TIPOLOGI")
        print("=" * 50)
        self.visualisasi_tipologi(clusters, pca_result)
        
        return {
            'kata_dominan': kata_dominan,
            'artikel_tema': artikel_tema,
            'media_tipologi': media_tipologi,
            'cluster_analysis': cluster_analysis,
            'evolusi_tema': evolusi
        }

# Jalankan analisis
if __name__ == "__main__":
    analyzer = AnalisisTipologi('data_mahatma.csv')
    hasil = analyzer.laporan_tipologi_lengkap()
