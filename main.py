import os
import re
import pandas as pd
from pathlib import Path
import pytesseract
import sys
from PIL import Image
import cv2
from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential
from concurrent.futures import ThreadPoolExecutor, as_completed  
from gliner import GLiNER

# Initialize pytesseract for OCR
pytesseract.pytesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Azure Credentials for Form Recognizer
os.environ['AZURE_FORM_RECOGNIZER_ENDPOINT'] = 'https://hackerx47.cognitiveservices.azure.com/'
os.environ['AZURE_FORM_RECOGNIZER_KEY'] = '69db9919b6834f8e8b8ab528798b42e8'

subscription_key = os.environ['AZURE_FORM_RECOGNIZER_KEY']
endpoint = os.environ['AZURE_FORM_RECOGNIZER_ENDPOINT']


document_analysis_client = DocumentAnalysisClient(
    endpoint=endpoint, credential=AzureKeyCredential(subscription_key)
)


gliner_model = GLiNER.from_pretrained("urchade/gliner_mediumv2.1")


target_labels = {
    "provisional diagnosis", "diagnosis", "dx", "impression", "working diagnosis"
}


medical_labels = ["Diagnosis", "Condition"]


def clean_text(text):
    text = re.sub(r'\s+', ' ', text)  
    text = re.sub(r'[^\w\s]', ' ', text)  
    text = re.sub(r'\b(proposed treatment|treatment plan|surgery|surgical management|icd 10 code|next steps)\b', '', text, flags=re.IGNORECASE)
    text = re.sub(r'\b(i\.|ii\.|iii\.|iv\.|v\.|1\.|2\.|3\.|4\.|5\.)\b', '', text)  
    text = re.sub(r'\b(proposed treatment|treatment plan|surgery|surgical management|icd 10 code|next steps|G|i|1|L|Proposed line of treatment)\b', '', text, flags=re.IGNORECASE)
    return text.strip()  



def preprocess_image(image_path):
    """
    Preprocesses the image to improve clarity for OCR - useful for handwritten and printed text.
    """
    image = cv2.imread(str(image_path))
    gray_image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    _, binary_image = cv2.threshold(gray_image, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    blurred_image = cv2.GaussianBlur(binary_image, (5, 5), 0)
    return binary_image


def read_image_with_document_intelligence(image_path):
    with open(image_path, "rb") as image_file:
        image_bytes = image_file.read()

    poller = document_analysis_client.begin_analyze_document(
        "prebuilt-layout", document=image_bytes
    )
    result = poller.result()
    return result


def extract_text_from_key_section(result):
    diagnosis_text = None
    capture_next_line = False

    for page in result.pages:
        for line in page.lines:
            text = line.content.strip().lower()

            
            print(f"Detected Line: {text}")

            
            if any(label in text for label in target_labels):
                capture_next_line = True
                continue

           
            if capture_next_line:
                diagnosis_text = clean_text(line.content)
                print(f"Caught Provisional Diagnosis: {diagnosis_text}")  # Debugging log
                capture_next_line = False
                break  # Only one diagnosis expected

    # Feedback diagnostics if nothing was found
    diagnosis_text = diagnosis_text if diagnosis_text else "No Provisional Diagnosis Extracted"
    print(f"Final Diagnosis Extracted: {diagnosis_text}")
    
    return diagnosis_text  


def process_image(image_path):
    
    print(f"Processing file: {image_path}...")

   
    result = read_image_with_document_intelligence(image_path)

    
    provisional_diagnosis_text = extract_text_from_key_section(result)

    
    return image_path.name, provisional_diagnosis_text   


def save_to_excel(data, output_excel_path):
    df = pd.DataFrame(data, columns=["file_name", "provisional_diagnosis"])
    with pd.ExcelWriter(output_excel_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Diagnosis Results')
        workbook = writer.book
        worksheet = writer.sheets['Diagnosis Results']
        header_format = workbook.add_format({'bold': True})
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)

# Process a folder containing multiple medical forms/images
def process_folder(folder_path):
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
                result = future.result()  
                results.append(result)
            except Exception as e:
                print(f"Error processing {img_file}: {e}")

    # Save final results into an Excel file
    output_excel_path = folder_path / "output_diagnoses.xlsx"
    save_to_excel(results, output_excel_path)
    print(f"Results have been successfully saved to {output_excel_path}.")

def main():
    """Main function to run either via command line or Streamlit."""
    
    if len(sys.argv) == 2:  # Check if there is a command-line argument provided.
        folder_path = sys.argv[1]
        process_folder(folder_path)  

if __name__ == "__main__":
    main()
