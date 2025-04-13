import pandas as pd

# Load your Excel file
df = pd.read_excel("Product-List.xlsx")  # Replace with your file name

# Print actual column names to verify
print("Columns in Excel:", df.columns.tolist())

# Strip column names in case there are extra spaces
df.columns = df.columns.str.strip().str.lower()

# Function to capitalize each word, skipping numbers
def capitalize_words(text):
    if isinstance(text, str):
        return ' '.join([word.capitalize() if not word.isdigit() else word for word in text.split()])
    return text

# Apply to the correct column
if 'particular' in df.columns:
    df['particular'] = df['particular'].apply(capitalize_words)
    df.to_excel("capitalized_output.xlsx", index=False)
    print("✅ Done! Saved as 'capitalized_output.xlsx'")
else:
    print("❌ Column 'particular' not found.")

