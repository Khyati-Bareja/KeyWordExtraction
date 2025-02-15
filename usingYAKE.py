import yake
from docx import Document

file_path = "/Users/khyati/Downloads/example.docx"
doc = Document(file_path)
text = " ".join([para.text for para in doc.paragraphs if para.text.strip()])  # Extracted text
# Define YAKE keyword extractor
yake_extractor = yake.KeywordExtractor(lan="en", n=2, dedupLim=0.9, top=3)

# Extract keywords
yake_keywords = [kw[0] for kw in yake_extractor.extract_keywords(text)]
print("YAKE Keywords:", yake_keywords)
