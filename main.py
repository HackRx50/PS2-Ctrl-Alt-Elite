import os
import re
import sys
import pandas as pd
from pathlib import Path
import pytesseract
from PIL import Image
import cv2
from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential
from concurrent.futures import ThreadPoolExecutor, as_completed
from gliner import GLiNER

# Initialize pytesseract for OCR
pytesseract.pytesseract.tesseract_cmd = r'/opt/homebrew/bin/tesseract'

# Azure Credentials for Form Recognizer
os.environ['AZURE_FORM_RECOGNIZER_ENDPOINT'] = ''
os.environ['AZURE_FORM_RECOGNIZER_KEY'] = ''

subscription_key = os.environ['AZURE_FORM_RECOGNIZER_KEY']
endpoint = os.environ['AZURE_FORM_RECOGNIZER_ENDPOINT']

# Azure Form Recognizer client
document_analysis_client = DocumentAnalysisClient(
    endpoint=endpoint, credential=AzureKeyCredential(subscription_key)
)

# Initialize GLiNER with pre-trained model for advanced entity recognition in medical texts
gliner_model = GLiNER.from_pretrained("urchade/gliner_mediumv2.1")

# Define labels for medical entity extraction using GLiNER - focus on Diagnosis and Condition
medical_labels = ["Diagnosis", "Condition"]

def clean_text(text):
    """Clean up the OCR-extracted text."""
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'[^\w\s]', ' ', text) 
    text = re.sub(r'\b(i)\b', ' ', text, flags=re.IGNORECASE)
    return text.strip()

def preprocess_image(image_path):
    """Preprocess the image for optimal OCR clarity."""
    image = cv2.imread(str(image_path))
    gray_image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    _, binary_image = cv2.threshold(gray_image, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return binary_image

def extract_text_tesseract(image_path):
    """Extract text using Tesseract OCR."""
    image = Image.open(image_path).convert("RGB")
    return pytesseract.image_to_string(image)

def read_image_with_document_intelligence(image_path):
    """Use Azure Form Recognizer to extract textual data from images."""
    with open(image_path, "rb") as image_file:
        image_bytes = image_file.read()

    poller = document_analysis_client.begin_analyze_document("prebuilt-layout", document=image_bytes)
    result = poller.result()
    return result

def extract_diagnosis_keyword_based(result):
    """Extract diagnosis text based on predefined target keywords."""
    target_keywords_set = {
        "provisional diagnosis", "diagnosis", "dx", "clinical impression", 
        "diagnostic impression", "suspected condition", "differential diagnosis",
        "impression", "initial diagnosis", "medical assessment",
        "discharge diagnosis", "primary diagnosis", "secondary diagnosis"
    }
    
    diagnosis_text = None
    capture_next_line = False

    for page in result.pages:
        for line in page.lines:
            text = line.content.strip().lower()

            # Search for presence of provisional diagnosis-related keywords
            if any(keyword in text for keyword in target_keywords_set):
                capture_next_line = True  # Mark for capture in next line
                continue

            if capture_next_line:  # Extract diagnosis in the next line
                diagnosis_text = clean_text(line.content)
                break

    return diagnosis_text if diagnosis_text else None

def extract_diagnosis_ner_based(ocr_text):
    """Use GLiNER NER to extract diagnosis-related entities."""
    ocr_text = clean_text(ocr_text)  # Clean the extracted text
    # Use GLiNER to predict medical entities
    extracted_entities = gliner_model.predict_entities(ocr_text, labels=medical_labels, threshold=0.5)

    # Extract any diagnoses/conditions
    diagnoses = [
        entity['text'] for entity in extracted_entities 
        if entity["label"] == "Diagnosis" or entity["label"] == "Condition"
    ]
    
    return ', '.join(diagnoses).strip() if diagnoses else "No Provisional Diagnosis Found"

def extract_all_text_from_layout(result):
    """Extract all text from Azure Layout model."""
    return ' '.join([line.content.strip() for page in result.pages for line in page.lines])

def process_image(image_path):
    """Main logic to process a single image and extract diagnosis."""
    print(f"Processing file: {image_path}...")

    try:
        # Use Azure Form Recognizer for text extraction
        result = read_image_with_document_intelligence(image_path)
        
        # Step 1: Try keyword-based extraction
        diagnosis = extract_diagnosis_keyword_based(result)
        
        if not diagnosis:
            # If no diagnosis found by keyword, switch to NER
            ocr_text = extract_all_text_from_layout(result)
            diagnosis = extract_diagnosis_ner_based(ocr_text)

            # Check if NER extraction returned "No Provisional Diagnosis Found" or contains "Provisional Diagnosis"
            if diagnosis == "No Provisional Diagnosis Found" or "Provisional Diagnosis" in diagnosis:
                # Directly use Azure Form Recognizer again
                result = read_image_with_document_intelligence(image_path)
                ocr_text = extract_all_text_from_layout(result)
                diagnosis = extract_diagnosis_ner_based(ocr_text)

    except Exception as e:
        print(f"Azure Form Recognizer failed for {image_path}, falling back to Tesseract OCR: {e}")
        ocr_text = extract_text_tesseract(image_path)
        diagnosis = extract_diagnosis_ner_based(ocr_text)


    return image_path.name, diagnosis

def save_to_excel(data, output_excel_path):
    """Save extracted diagnosis data into Excel."""
    df = pd.DataFrame(data, columns=["file_name", "provisional_diagnosis"])
    with pd.ExcelWriter(output_excel_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Diagnoses')
        workbook = writer.book
        worksheet = writer.sheets['Diagnoses']
        header_format = workbook.add_format({'bold': True})
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)

def process_folder(folder_path):
    """Process all images in a folder and extract diagnosis information."""
    folder_path = Path(folder_path)
    image_extensions = {'.jpg', '.jpeg', '.png'}
    image_files = [f for f in folder_path.iterdir() if f.suffix.lower() in image_extensions]

    if not image_files:
        print(f"No image files found in {folder_path}.")
        return

    results = []

    with ThreadPoolExecutor() as executor:
        futures = {executor.submit(process_image, img_file): img_file for img_file in image_files}
        
        for future in as_completed(futures):
            img_file = futures[future]
            try:
                result = future.result()  # Retrieve processed result
                results.append(result)  # Append the result
            except Exception as e:
                print(f"Error processing {img_file}: {e}")

    output_excel_path = folder_path / "output_diagnoses.xlsx"
    save_to_excel(results, output_excel_path)
    print(f"Results saved to {output_excel_path}.")

# Example folder path where images are located
folder_path = r"/Users/apple/Desktop/HackRx/HackRx_Finals/Images"

 
def main():
    if len(sys.argv) == 2:  # Check if there is a command-line argument provided.
        folder_path = sys.argv[1]
        process_folder(folder_path)  

if __name__ == "__main__":
    main()
