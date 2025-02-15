
from docx import Document
from keybert import KeyBERT

# Step 1: Read text from the .docx file
file_path = "/Users/khyati/Downloads/example.docx"
doc = Document(file_path)
text = " ".join([para.text for para in doc.paragraphs if para.text.strip()])  # Extracted text

# Step 2: Load BERT-based keyword extractor
kw_model = KeyBERT()

# Step 3: Extract keywords (top 5)
bert_keywords = kw_model.extract_keywords(text, keyphrase_ngram_range=(1, 2), stop_words="english", top_n=3)
bert_keywords = [kw[0] for kw in bert_keywords]

# Step 4: Print results
print("BERT Keywords:", bert_keywords)

