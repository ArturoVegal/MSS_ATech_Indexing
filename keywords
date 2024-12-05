import pandas as pd
from collections import Counter
from nltk.tokenize import word_tokenize
import string

# Load your dataset
file_path = '/Users/jarturovega/Downloads/MSS Defence Products_Keychain.xlsx'
df = pd.read_excel(file_path)

# Combine product titles and descriptions for analysis
if 'Product' in df.columns and 'Description' in df.columns:
    text_data = df['Product'].astype(str) + " " + df['Description'].astype(str)
else:
    text_data = df['Description'].astype(str)  # Fallback to Description if Product is missing

# Tokenize words and remove punctuation
tokens = [word.lower() for text in text_data for word in word_tokenize(str(text)) if word.isalpha()]

# Count most common words
word_counts = Counter(tokens)
common_words = word_counts.most_common(100)  # Adjust the number to get more or fewer words

# Convert results to a DataFrame for saving and reviewing
keywords_df = pd.DataFrame(common_words, columns=['Keyword', 'Count'])

# Save the keywords to a file
output_path = output_path = '/Users/jarturovega/MSS Advanced Tech/extracted_keywords.xlsx'
keywords_df.to_excel(output_path, index=False)

print(f"Most common words saved to {output_path}")
