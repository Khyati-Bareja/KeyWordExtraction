import spacy
import pandas as pd
import re
from collections import Counter

# Load spaCy's English model
nlp = spacy.load("en_core_web_sm")

# Step 1: Read text from the .docx file
from docx import Document
file_path = "example_changed_test.docx"
doc = Document(file_path)
text = " ".join([para.text for para in doc.paragraphs if para.text.strip()])  # Extracted text

# Step 2: Load the exception list from Excel (columns 3 and 5)
# exception_file = "Exception list Food and Beverages New.xlsx"
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
# Step 3: Process the text with spaCy NLP
doc_spacy = nlp(text)

# Step 4: Extract named entities (focus on food and product names)
entities = [ent.text.lower() for ent in doc_spacy.ents if ent.label_ in ['ORG', 'PRODUCT', 'GPE', 'MONEY', 'PERSON']]

# Step 5: Count the occurrences of each entity (this helps rank them)
entity_counts = Counter(entities)

# Step 6: Remove any entities found in the exception list
filtered_entities = [ent for ent in entity_counts if ent not in exception_list]

# Step 7: Sort entities by frequency to get the top ones
sorted_filtered_entities = sorted(filtered_entities, key=lambda x: entity_counts[x], reverse=True)
print("Entities:", sorted_filtered_entities)
# Step 8: Show top valid keywords
top_keywords = sorted_filtered_entities[:3]  # Get top 3 valid keywords

# Display the results
print("Top Valid Keywords Extracted Using spaCy:")
print(top_keywords)
