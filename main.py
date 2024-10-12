import os
import sys
import time
import pandas as pd
from azure.cognitiveservices.vision.computervision import ComputerVisionClient
from azure.cognitiveservices.vision.computervision.models import OperationStatusCodes
from msrest.authentication import CognitiveServicesCredentials
from pathlib import Path

# Set Azure credentials
os.environ['AZURE_COMPUTER_VISION_SUBSCRIPTION_KEY'] = '58f17158504d499f88750389e3646849'
os.environ['AZURE_COMPUTER_VISION_ENDPOINT'] = 'https://testingmine.cognitiveservices.azure.com/'

# Retrieve Azure credentials from environment variables
subscription_key = os.environ['AZURE_COMPUTER_VISION_SUBSCRIPTION_KEY']
endpoint = os.environ['AZURE_COMPUTER_VISION_ENDPOINT']

# Create the Computer Vision client
computervision_client = ComputerVisionClient(endpoint, CognitiveServicesCredentials(subscription_key))

def read_image_from_path(image_path):
    with open(image_path, "rb") as image:
        read_response = computervision_client.read_in_stream(image, raw=True)
        operation_location = read_response.headers["Operation-Location"]
        operation_id = operation_location.split("/")[-1]

        while True:
            result = computervision_client.get_read_result(operation_id)
            if result.status not in ['notStarted', 'running']:
                break
            time.sleep(1)
        if result.status == OperationStatusCodes.succeeded:
            return result.analyze_result.read_results
    return None


def extract_diagnosis(read_results):
    target_keywords = {"provisional diagnosis", "diagnosis", "provisional diagnosis:"}
    section_indicators = {
        "proposed line of treatment", "treatment:", "investigation:", "surgical:",
        "management:", "care:", "drug administration", "surgery:", "icd 10 code", "next steps", "plan of care",
        "i.", "ii.", "iii.", "iv.", "v."
    }
    diagnosis_lines = []
    capture = False
    for page in read_results:
        for line in page.lines:
            text = line.text.strip().lower()

            if not capture and any(keyword in text for keyword in target_keywords):
                capture = True
                parts = text.split(max(target_keywords, key=len))
                diagnosis_part = parts[1].strip() if len(parts) > 1 else ""
                if diagnosis_part:  # Only start capture if there's text after the keyword
                    diagnosis_lines.append(diagnosis_part)
            elif capture:
                if any(indicator in text for indicator in section_indicators):
                    capture = False  # Stop capturing if a section indicator line is detected
                else:
                    # Continue capturing all text until we hit a stopping condition
                    diagnosis_lines.append(line.text.strip())
    diagnosis_text = ' '.join(diagnosis_lines).strip()
    # Post-processing to remove any trailing "i.", "ii."
    diagnosis_text = diagnosis_text.rstrip("i. ").rstrip("ii. ").rstrip("iii. ").rstrip("iv. ").rstrip("v. ").strip()
    return diagnosis_text if diagnosis_text else "No provisional diagnosis found"


def save_to_excel(data, output_excel_path):
    df = pd.DataFrame(data, columns=["file_name", "provisional_diagnosis"])
    with pd.ExcelWriter(output_excel_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        header_format = workbook.add_format({'bold': True})
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)


def process_folder(folder_path):
    folder_path = Path(folder_path)
    data = []
    if not folder_path.exists():
        print(f"The folder path {folder_path} does not exist.")
        return
    if not os.access(folder_path, os.R_OK):
        print(f"The folder path {folder_path} is not readable.")
        return
    image_extensions = {'.jpg', '.jpeg', '.png'}
    processed_file_count = 0
    for filename in folder_path.iterdir():
        if filename.suffix.lower() in image_extensions:
            print(f"Processing file: {filename.name}")
            read_results = read_image_from_path(filename)
            if read_results:
                diagnosis = extract_diagnosis(read_results)
                data.append([filename.name, diagnosis])
            else:
                print(f"No results for {filename}.")
            processed_file_count += 1

    output_excel_path = folder_path / "output_diagnoses.xlsx"
    save_to_excel(data, output_excel_path)
    print(f"Results saved to {output_excel_path}. Processed {processed_file_count} files.")

# def main():
#     # Set the folder path directly
#     folder_path = '/Users/apple/Downloads/HackRx/Sample_Imgs' 
#     process_folder(folder_path)
    
def main():
    if len(sys.argv) != 2:
        print("Usage: python main.py <folder_path>")
        sys.exit(1)

    # Get the folder path from the command-line argument
    folder_path = sys.argv[1]
    process_folder(folder_path)  


if __name__ == "__main__":
    main()