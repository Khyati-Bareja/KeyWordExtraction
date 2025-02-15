
import yake
from docx import Document
from sklearn.feature_extraction.text import TfidfVectorizer

# Step 1: Read text from the .docx file
file_path = "/Users/khyati/Downloads/example.docx"
doc = Document(file_path)
text = " ".join([para.text for para in doc.paragraphs if para.text.strip()])  # Extracted text

# Step 2: Extract keywords using YAKE
yake_extractor = yake.KeywordExtractor(lan="en", n=2, dedupLim=0.9, top=3)
yake_keywords = [kw[0] for kw in yake_extractor.extract_keywords(text)]

# Step 3: Extract keywords using TF-IDF
vectorizer = TfidfVectorizer(stop_words="english", ngram_range=(1, 2))
X = vectorizer.fit_transform([text])
tfidf_scores = dict(zip(vectorizer.get_feature_names_out(), X.toarray()[0]))
tfidf_keywords = sorted(tfidf_scores, key=tfidf_scores.get, reverse=True)[:3]  # Top 5 keywords

# Step 4: Combine Both Keyword Lists
final_keywords = list(set(yake_keywords + tfidf_keywords))  # Remove duplicates

# Step 5: Print Final Keywords
print("Optimized approach extracted Keywords:", final_keywords)
