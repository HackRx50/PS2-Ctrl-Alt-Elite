import pandas as pd
import re

# Load ICD-10 Codes
def load_icd10_codes(icd10_file):
    """
    Load ICD-10 dataset containing Level-3 codes and descriptions.
    The data should have columns for Level-3 Code and Level-3 Desc.
    Returns a dictionary mapping Level-3 Description (lowercased) to Level-3 Code and Desc.
    """
    icd10_df = pd.read_csv(icd10_file)

    # Handle missing values (fill NaNs with empty strings)
    icd10_df['Level-3 Desc'] = icd10_df['Level-3 Desc'].fillna('')

    # Create a dictionary for faster lookups (mapping Level-3 Description to Level-3 Code and Desc)
    icd10_dict = {
        row["Level-3 Desc"].lower(): (row["Level-3 Code"], row["Level-3 Desc"])
        for _, row in icd10_df.iterrows() if row["Level-3 Desc"]
    }

    return icd10_dict

# Clean and tokenize text, ignoring punctuation and focusing only on words
def clean_and_tokenize(text):
    """
    Clean the input text by removing punctuation and splitting into words.
    """
    if not text or pd.isna(text):
        return []
    
    # Remove multiple spaces and punctuations but keep the original structure of words
    cleaned_text = re.sub(r'[^\w\s]', '', str(text))  # Remove any punctuation
    return re.findall(r'\b\w+\b', cleaned_text.lower())  # Tokenized as lowercase words

# Function to compare corrected field with ICD-10 codes
def extract_icd_code(corrected_text, icd10_dict):
    """
    Extract relevant ICD-10 Code and Level-3 Description from corrected text 
    by matching against the ICD-10 dataset's Level-3 descriptions.
    Returns the Level-3 Code and its Level-3 description.
    """
    words = clean_and_tokenize(corrected_text)  # Tokenize corrected sentence

    # Compare each word against the ICD-10 Level-3 description
    for word in words:
        for desc in icd10_dict.keys():
            # Match whether the word exists in any Level-3 description
            if word in desc:
                return icd10_dict[desc]  # Return the (Code, Description)
    
    # If no ICD-10 code is found
    return None, None  # No match found

def compare_corrected_with_icd10(corrected_excel, icd10_file, output_excel):
    """
    Compare corrected diagnoses in the 'Corrected Output' column
    of a corrected Excel file, extract relevant ICD-10 codes,
    and save to a new Excel file.
    """
    # Load the ICD-10 data (Level-3 Description and Level-3 Code)
    icd10_dict = load_icd10_codes(icd10_file)
    
    # Load the corrected diagnosis data from the provided Excel
    df = pd.read_excel(corrected_excel)
    
    # Ensure necessary columns are present
    if 'File Name' not in df.columns or 'Corrected Output' not in df.columns:
        raise ValueError("Input Excel must contain 'File Name' and 'Corrected Output' columns.")
    
    # Function to process each row and extract ICD-10 Code and Description
    def process_row(row):
        corrected_text = str(row['Corrected Output']).strip()  # Corrected diagnosis from the row
        icd_code, icd_desc = extract_icd_code(corrected_text, icd10_dict)  # Compare and get ICD-10 code
        return pd.Series({
            'ICD-10 Code': icd_code,
            'Level 3 Description': icd_desc
        })
    
    # Apply comparison and ICD-10 extraction to each row
    df[['ICD-10 Code', 'Level 3 Description']] = df.apply(process_row, axis=1)

    # Save the final result back into Excel
    df.to_excel(output_excel, index=False)

    print(f"ICD-10 comparison completed. Saved to {output_excel}")

# Path to the corrected Excel file (previously processed with spelling corrections)
corrected_excel = r"/Users/apple/Desktop/HackRx/HackRx_Finals/datasets/output_corrected_l2.xlsx"

# Path to the ICD-10 dataset file (contains Level 3 codes and descriptions)
icd10_file = r"/Users/apple/Desktop/HackRx/HackRx_Finals/datasets/ICD.xlsx"

# Path to save the final output with ICD-10 comparison
output_excel = r"/Users/apple/Desktop/HackRx/HackRx_Finals/Images/output_comparison_with_icd10.xlsx"

# Run the ICD-10 code comparison process
compare_corrected_with_icd10(corrected_excel, icd10_file, output_excel)

