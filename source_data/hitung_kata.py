from docx import Document
from collections import Counter
import re

doc = Document('source_data/Olivia Rodrigo/Analisis_Tema_Olivia_Rodrigo_Revisi.docx')

tema1_text = ""
tema2_text = ""

for table in doc.tables:
    if len(table.rows) > 0:
        header = [cell.text.strip() for cell in table.rows[0].cells]
        if any('Kutipan' in h or 'Tema' in h for h in header):
            for row in table.rows[1:]:
                cells = [cell.text.strip() for cell in row.cells]
                if len(cells) >= 2:
                    tema1_text += " " + cells[-2]
                    tema2_text += " " + cells[-1]

def count_words(text, keywords):
    words = re.findall(r'\b[a-zA-Z]{3,}\b', text.lower())
    filtered = [w for w in words if w in keywords]
    return Counter(filtered)

# Kata kunci Tema 1: Revival Memory Y2K
tema1_keywords = {'y2k', 'nostalgia', 'nostalgic', 'vintage', 'retro', 'revival', 'throwback', '2000s', 'nineties', 'early', 'memory', 'memories', 'past', 'era', 'trends', 'fashion', 'style', 'aesthetic', 'inspired'}

# Kata kunci Tema 2: Identitas Generasional
tema2_keywords = {'generation', 'generational', 'gen z', 'identity', 'millennial', 'millennials', 'youth', 'young', 'teen', 'teenage', 'cultural', 'culture', 'icon', 'influence', 'celebrity', 'star', 'figure', 'role', 'model', 'representation'}

tema1_counts = count_words(tema1_text, tema1_keywords)
tema2_counts = count_words(tema2_text, tema2_keywords)

print("=== TEMA 1: Revival Memory Y2K ===")
print(f"Total kata relevan: {sum(tema1_counts.values())}")
for word, count in tema1_counts.most_common():
    print(f"{word}: {count}")

print("\n=== TEMA 2: Identitas Generasional ===")
print(f"Total kata relevan: {sum(tema2_counts.values())}")
for word, count in tema2_counts.most_common():
    print(f"{word}: {count}")
