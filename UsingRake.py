from rake_nltk import Rake
from docx import Document

def extract_text_from_docx(docx_path):
    """Extract text from a DOCX file."""
    doc = Document(docx_path)
    return "\n".join([para.text for para in doc.paragraphs])

def extract_keywords_rake(docx_path):
    text = extract_text_from_docx(docx_path)

    rake = Rake()
    rake.extract_keywords_from_text(text)
    return rake.get_ranked_phrases()[:10]  # Get top 10 ranked phrases

# Example usage
docx_path = "input.docx"  # Replace with your actual file path
rake_keywords = extract_keywords_rake(docx_path)
print("RAKE Extracted Keywords:", rake_keywords)
