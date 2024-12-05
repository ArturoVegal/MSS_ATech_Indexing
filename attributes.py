import pandas as pd
from nltk.tokenize import sent_tokenize, word_tokenize
from nltk import pos_tag
from nltk.tag.perceptron import PerceptronTagger
import nltk

# Download necessary NLTK resources
nltk.download('punkt')
nltk.download('averaged_perceptron_tagger')

# Explicitly initialize the PerceptronTagger
tagger = PerceptronTagger()

# Function to classify text into Type, Purpose, and Attributes
def extract_features(description):
    # Tokenize the description into sentences
    sentences = sent_tokenize(description)
    types, purposes, attributes = [], [], []

    # Iterate over each sentence
    for sentence in sentences:
        # Tokenize words and POS tag them
        words = word_tokenize(sentence.lower())
        tagged_words = pos_tag(words, tagset=None, lang='eng')  # Correct language code
        
        # Extract Types (Nouns)
        types += [word for word, tag in tagged_words if tag.startswith('NN')]
        
        # Extract Purposes (Verbs)
        purposes += [word for word, tag in tagged_words if tag.startswith('VB')]
        
        # Extract Attributes (Adjectives)
        attributes += [word for word, tag in tagged_words if tag.startswith('JJ')]
    
    # Join the lists into strings for output
    return {
        "Type": ", ".join(set(types)),
        "Purpose": ", ".join(set(purposes)),
        "Attributes": ", ".join(set(attributes))
    }

# Load the dataset
file_path = '/Users/jarturovega/Downloads/MSS Defence Products_Keychain.xlsx'
df = pd.read_excel(file_path)

# Ensure there is a 'Description' column
if 'Description' not in df.columns:
    raise ValueError("The dataset must have a 'Description' column!")

# Apply the feature extraction to each product
for index, row in df.iterrows():
    if pd.notna(row['Description']):  # Ensure the description is not empty
        features = extract_features(row['Description'])
        df.at[index, 'TYPE'] = features["Type"]  # Populate 'TYPE' column
        df.at[index, 'PURPOSE'] = features["Purpose"]  # Populate 'PURPOSE' column
        df.at[index, 'ATTRIBUTES'] = features["Attributes"]  # Populate 'ATTRIBUTES' column

# Save the updated file
output_file = '/Users/jarturovega/Downloads/updated_MSS_Defence_Products.xlsx'
df.to_excel(output_file, index=False)

print(f"Feature extraction complete. Results saved to {output_file}.")