# Medical-Diagnosis-Extraction

## About the Solution
- Uses Azure Computer Vision for Optical Character Recognition (OCR).
- Extracts and processes the diagnosis text specifically from handwritten medical forms (70% Accuracy).
- Saves the extracted data into an Excel file (output_diagnoses) with high accuracy and precision.

## How to run
Create and activate a new virtual environment (recommended) by running
the following:

```bash
python3 -m venv myvenv
myvenv\Scripts\activate
```

Install the dependencies:
```bash
pip install -r requirements.txt
pip install azure-cognitiveservices-vision-computervision
pip install msrest
pip install pandas
pip install pathlib
pip install pytesseract
pip install opencv-python
pip install Pillow
pip install azure-core
```
Run the script:
```bash
python main.py <folder_path>
```

## Output
![Output Image](Output.png)
