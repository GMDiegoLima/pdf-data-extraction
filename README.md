# PDF to Text & Document Converter
This Python tool allows you to extract text from images embedded in PDF files and save the output as either a new PDF or a DOCX document. Additionally, it supports converting entire PDF documents to DOCX format. The tool is especially useful for digitizing documents where the text is present in images rather than as selectable text.

### Features
PDF Image to Text (PDF output): Extracts text from images in a PDF file and saves it as a new PDF.
PDF Image to Text (DOCX output): Extracts text from images in a PDF file and saves it as a DOCX document.
PDF to DOCX Conversion: Converts the entire PDF, including all text, into a DOCX document.


### Installation
To get started, clone this repository and install the required dependencies:
1. `git clone https://github.com/GMDiegoLima/pdf-data-extraction.git`<br>
2. `cd pdf-data-extraction`<br>
3. `pip install -r requirements.txt`

Ensure you have Tesseract-OCR installed on your system, as it is required for OCR operations.

### Usage
Place the PDF files you want to process in the PDFs to analyse folder.

Run the script:
<br>`python main.py`<br>
Follow the on-screen prompts:

Option 1: Extract text from images in a PDF and save it as a new PDF.
Option 2: Extract text from images in a PDF and save it as a DOCX file.
Option 3: Convert the entire PDF document to a DOCX file.
The processed files will be saved in the Final results folder.

### Example
If you want to extract text from images in a PDF and save it as a DOCX file:
<br>`python main.py`<br>
Type 2 and press enter to select the option 2
Enter a name for the output DOCX file.
The tool will process each PDF in the PDFs to analyse folder and save the results in the Final results folder.

### Requirements
- Python 3.x<br>
- Tesseract-OCR installed and added to your PATH
- **Libraries**: pytesseract, Pillow, docx, reportlab, PyMuPDF

### License
This project is licensed under the Apache-2.0 license.

Acknowledgments
Tesseract-OCR for the OCR functionality.
PyMuPDF for PDF processing.
python-docx for DOCX creation.
