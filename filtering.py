import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LogisticRegression
from sklearn.pipeline import make_pipeline
from sklearn.metrics import classification_report

# Sample data
data = {
    'Email': ['example1@example.com', 'example2@example.com','cs@4care.co.th','custcare@ameen.com.my','customercare@jivo.in','customercare@rasoitatva.com','customerser@unionbankofindia.co.in','donko@donko.info','export@tolido-group.com','info@amul.coop'],
    'First_Name': ['John', 'Jane','Gaglan','','T','A'],
    'Last_Name': ['Doe', 'Smith','Bansal','Kumar','R','B',''],
    'Products_Name': [
        'Cola, Sour Spray - Cola, Orange, Kacha Aam,  Cola Sour Spray, Bursters Candy, Gel Candy, Icy Juicy Ice Pops, Icy Juicy All Natural Ice Pops, Dare Drop Kacha Aam Sour Spray',
        'Breads and Confectionary, Fruit based Ingredients, Canned and Frozen Seafood, Canned Juices (Preserve Beverages), Dried Fruits and Nuts, Fresh produces (Fruits and Veggies), Organic, Halal and Kosher, Ice Cream and Cakes, Nestle / Maaggi and DOLE Products, Canned and Frozen Meat, Canned and Frozen Fruit, Coconut Products, Packed Ethnic Foods,Mixes and Condiments',
        'DAITY ICE CREAM, DAITY DAIRY, DAITY FALOODEH, BIG N BEN ICE CREAM, DAITY EGG',
        'Gelatin, Fish Oils, Amino Acids, Essential Oils, Nutraceuticals, Plant Extracts & Ayurvedic',
        'HOLEY ROLLS Chicken Crunch, HOLEY CREPES Banoffee, HOLEY CREPES Thai Tea, HOLEY Grainy Bar, Crispy Coconut Rolls',
        'Tablete de Chocolate Branco com Morango, Tablete de Chocolate Branco com Cookies, Tablete de Chocolate Cacau ao Leite, Tablete de Chocolate Branco com Cranberry, Kibbles Chocolate Cocoa Origin Line, Cocoa Chocolate Bar with Baru, Cocoa Chocolate Bar with Coconut Milk, White Chocolate Bar with Coconut Milk, Chocolate Bar Kit Cocoa Zero Sugar, Cocoa Chocolate Bar with Cranberry, Cocoa Chocolate Bar, Cocoa Chocolate Tablet w/ Fleur de Sal, Cocoa Chocolate Bar with Zero Sugar, Kibbles Chocolate Cocoa Origin Line, Cocoa Chocolate Bar with Nibs, Cocoa Chocolate Bar with Roasted Coffee, Gift Box of Flavors Chocolates, White Chocolate Bar with Strawberry, White Chocolate Tablet with Cookies, Cocoa Milk Chocolate Bar, White Chocolate Bar with Cranberry, Tablete de Chocolate Cacau ao Leite de Coco, Tablete de Chocolate Branco Cacau ao Leite de Coco',
        'Biryani Masala, Butter Chicken Masala, Chat Masala, Chhole Masala, Chicken Masala, Kitchen King Masala, Meat Masala, Misal Masala, Paneer Butter Masala, Pani Puri Masala, Pav Bhaji Masala, Sabji Masala, Sambar Masala, Tandoori Masala, Hing - Yellow, Kasoori Methi, Saffron, Tea Masala, Cassia (Taj), Coriander Powder, Coriander Cumin Powder, Cloves, Chilli Powder Hot, Black Pepper, Chilli Powder Kashmiri, Coriander Seed, Cumin Seed, Fennel Seed, Garam Masala, Green Cardamom, Methi Seed, Mustard, Turmeric Powder',
        'Curd, Ghee, Desserts, Food Enhancers, Fruit Yogurt, Premium Yogurt, UHT, Ready To Drink, Condensed Milk, Ice Cream, Capella Chocolate, Asal'
    ]
}

# Create DataFrame
df = pd.DataFrame(data)

# Predefined valid and invalid product names
invalid_products = [
    "Fresh Tomatoes", "Fresh Fruits", "Frozen Pizza", 
    "Fresh Vegetables", "Yogurt", "Fresh Dairy Products"
]

valid_products = [
    "Cola", "Sour Spray", "Orange", "Canned Fruits"
]

# Create labels for the training set
labels = ['valid' if product in valid_products else 'invalid' for product in invalid_products + valid_products]

# Create a list of product names for training
products = invalid_products + valid_products

# Split the data
X_train, X_test, y_train, y_test = train_test_split(products, labels, test_size=0.2, random_state=42)

# Build a classification pipeline
pipeline = make_pipeline(TfidfVectorizer(), LogisticRegression())
pipeline.fit(X_train, y_train)

# Evaluate the model
y_pred = pipeline.predict(X_test)
print(classification_report(y_test, y_pred))

# Classify products from the original DataFrame
def classify_product(product_name):
    prediction = pipeline.predict([product_name])[0]
    return prediction

# Filter valid products
df['Validity'] = df['Products_Name'].apply(lambda x: classify_product(x))
valid_records = df[df['Validity'] == 'valid']

# Save valid records to a new Excel file
valid_records.to_excel('valid_products.xlsx', index=False)
print("Valid products saved to 'valid_products.xlsx' file.")