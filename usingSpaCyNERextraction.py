# Using SpaCy NER for Keyword Extraction
# This script extracts named entities like product names, brand names, and proper nouns.

import spacy
from docx import Document
import re
from keybert import KeyBERT
import pandas as pd
from collections import Counter
from spacy.matcher import PhraseMatcher


# Load English NLP model
nlp = spacy.load("en_core_web_sm")


def extract_text_from_docx(docx_path):
    """Extract text from a DOCX file."""
    doc = Document(docx_path)
    return "\n".join([para.text for para in doc.paragraphs])

# def extract_keywords_ner(docx_path):
#     text = extract_text_from_docx(docx_path)
#     doc = nlp(text)
#     keywords = set()

#     # Extract Named Entities (NER)
#     for ent in doc.ents:
#         if ent.label_ in ["ORG", "PRODUCT", "GPE", "PERSON", "BAKERY"]: 
#             keywords.add(ent.text)

#      custom_food_terms = ["Bagels", "Sashimi", "Croissant", "Brioche", "Tuna Tartare"]
#     return list(keywords)

def extract_keywords_ner_with_custom_rules(text):
    doc = nlp(text)
    keywords = set()

    # Extract Named Entities
    for ent in doc.ents:
        if ent.label_ in ["ORG", "PRODUCT", "GPE", "PERSON"]:
            keywords.add(ent.text)

    # Custom food keyword list
    custom_food_terms = ["Bagels", "Sashimi", "Croissant", "Brioche", "Tuna Tartare"]
    matcher = PhraseMatcher(nlp.vocab)
    patterns = [nlp.make_doc(food) for food in custom_food_terms]
    matcher.add("FOOD_TERMS", patterns)

    matches = matcher(doc)
    for match_id, start, end in matches:
        keywords.add(doc[start:end].text)

    return list(keywords)

file_path = "example_changed_test.docx"  
ner_keywords = extract_keywords_ner_with_custom_rules(file_path)
print("NER Extracted Keywords from the example file:", ner_keywords)


exception_file = "exception_modified.xlsx"
exception_df = pd.read_excel(exception_file, usecols=[2, 4])  # Read only 3rd and 5th columns

# Clean the exception list by removing numbered prefixes (e.g., '1. Keyword') and converting to lowercase
exception_list = set(
    exception_df.stack()
    .dropna()
    .astype(str)
    .apply(lambda x: re.sub(r'^\d+\.\s*', '', x).strip().lower())  # Remove numbered prefixes like '1. '
)

print("Cleaned Exception List:", exception_list)

valid_keywords = [
    kw for kw in ner_keywords if kw not in exception_list and kw not in exception_list
]
print("Valid Keywords after Filtering:", valid_keywords)