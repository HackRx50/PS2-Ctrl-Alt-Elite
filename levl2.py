import pandas as pd
import re
from spellchecker import SpellChecker
import xlsxwriter

# Load medical terms from the text file into a set for fast lookups (case-insensitive)
def load_medical_terms(medical_terms_file):
    """
    Load a list of valid medical terms from medical_terms.txt file.
    Returns a set of medical terms in lowercase for case-insensitive matching.
    """
    with open(medical_terms_file, 'r', encoding='utf-8') as file:
        medical_terms = {line.strip().lower() for line in file if line.strip()}
    return medical_terms

# Clean and tokenize text, ignoring punctuation and focusing only on words
def clean_and_tokenize(text):
    """
    Clean and tokenize input text, ignoring punctuation.
    Returns a list of words without punctuation, while keeping original case for structure.
    """
    if not text or pd.isna(text):
        return []
    
    # Remove multiple spaces and punctuations but keep the original structure of words
    cleaned_text = re.sub(r'[^\w\s]', '', str(text))  # Remove punctuation
    return re.findall(r'\w+', cleaned_text)

# Spell check with PySpellChecker (case-insensitive)
def spell_check_strict(word, medical_terms=None):
    """
    Perform case-insensitive spell checking for words not in medical terms.
    It returns the corrected word if misspelled.
    If the word is found in the medical terms set, it skips spell-checking.
    """
    # Skip spell-check if the word is a known medical term.
    if medical_terms and word.lower() in medical_terms:
        return word  # No correction needed; valid medical term

    # Perform spell-check only for words not in medical terms
    spell = SpellChecker()
    corrected = spell.correction(word.lower())  # Lowercase for spell checking
    return corrected if corrected else word  # Return corrected or original word

def correct_text(text, medical_terms):
    """
    Process the extracted text for spell correction based on medical terms (case-insensitive).
    Highlights corrected words by marking their cells during Excel writing.
    """
    words = clean_and_tokenize(text)  # Cleaned version of the text (for spell-checking)
    original_words = re.findall(r'\S+', str(text))  # Capture all original words/punctuation/whitespace
    
    corrected_words = []  # Store corrected words for final output
    has_corrections = False  # Track if any corrections were made
    
    original_idx = 0
    for word in words:
        # Ensure we match words from the original, ignoring punctuation
        while original_idx < len(original_words) and re.sub(r'[^\w]', '', original_words[original_idx]).lower() != word.lower():
            corrected_words.append(original_words[original_idx])
            original_idx += 1

        # Check if the word needs correction (not in the medical terms)
        corrected_word = spell_check_strict(word, medical_terms)

        if corrected_word.lower() != word.lower():
            has_corrections = True  # Note that a correction was made
            corrected_words.append(corrected_word)  # Add the corrected word
        else:
            corrected_words.append(original_words[original_idx])  # No change, use the original word
            
        original_idx += 1

    # Append any remaining unprocessed tokens
    while original_idx < len(original_words):
        corrected_words.append(original_words[original_idx])
        original_idx += 1

    corrected_text = ' '.join(corrected_words)
    return corrected_text, has_corrections  # Return corrected text and a flag for corrections

# Main function to process spell-checking in Excel
def correct_diagnoses_in_excel(input_excel, medical_terms_file, output_excel):
    """
    Correct extracted diagnoses in the 'provisional_diagnosis' column.
    Check against a list of medical terms, spell-correct missing terms, and highlight corrected fields.
    """
    # Load medical terms for spell-check
    medical_terms = load_medical_terms(medical_terms_file)

    # Load the OCR-extracted data from the input Excel file
    df = pd.read_excel(input_excel)

    # Ensure necessary columns are present
    if 'file_name' not in df.columns or 'provisional_diagnosis' not in df.columns:
        raise ValueError("Input Excel must contain 'file_name' and 'provisional_diagnosis' columns.")

    corrections_made = []  # Track rows where corrections were made
    
    # Process each row for corrections
    def process_row(row):
        corrected_text, has_corrections = correct_text(str(row['provisional_diagnosis']), medical_terms)
        corrections_made.append(has_corrections)
        return corrected_text

    # Apply corrections to all rows
    df['corrected_output'] = df.apply(process_row, axis=1)

    # Add column tracking whether corrections were made
    df['Corrections Applied'] = corrections_made

    # Rename columns for clarity
    df.rename(columns={
        'file_name': 'File Name',
        'provisional_diagnosis': 'Extracted Output',
        'corrected_output': 'Corrected Output'
    }, inplace=True)

    # Save to Excel with yellow highlighting for corrected rows
    save_to_excel_with_highlight(df, output_excel)

    print(f"Corrected data saved successfully to {output_excel}")

def save_to_excel_with_highlight(df, output_excel):
    """Save DataFrame to Excel and highlight cells where corrections have been made."""
    
    # Create an Excel writer
    with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        
        # Define a yellow highlight format for corrected cells
        highlight_format = workbook.add_format({'bg_color': 'yellow'})
        
        for idx, row in df.iterrows():
            if row['Corrections Applied']:  # If corrections were applied, highlight cells
                worksheet.write(idx + 1, 3, row['Corrected Output'], highlight_format)  # Highlight "Corrected Output" column (corresponds to column index 3)

### Example Usage ###

# Path to the medical terms file (custom medical terms for spell-checking)
medical_terms_file = r"/Users/apple/Desktop/HackRx/HackRx_Finals/datasets/medical_terms.txt"

# Path to the input Excel file with extracted diagnoses
input_excel = r"/Users/apple/Desktop/HackRx/HackRx_Finals/datasets/output_diagnoses.xlsx"

# Path to save the corrected Excel output
output_excel = r"/Users/apple/Desktop/HackRx/HackRx_Finals/Images/output_corrected.xlsx"

# Run the correction process with yellow-highlighted corrections in Excel
correct_diagnoses_in_excel(input_excel, medical_terms_file, output_excel)