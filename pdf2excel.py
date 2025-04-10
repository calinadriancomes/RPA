import pytesseract
from pdf2image import convert_from_path
import pandas as pd
import os

# Optional: Set the path to the Tesseract executable manually if not in PATH
# Example for Windows:
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Step 1: Convert each page of the scanned PDF to a high-resolution image
pdf_path = "scanned_application.pdf"
images = convert_from_path(pdf_path, dpi=300)  # Higher DPI gives better OCR results

# Step 2: Initialize an Excel writer for saving the results
output_excel = "scanned_output_ocr.xlsx"
excel_writer = pd.ExcelWriter(output_excel, engine='openpyxl')

# Step 3: Process each image with OCR and store the results in Excel
for i, img in enumerate(images):
    # Extract raw text using Tesseract OCR
    text = pytesseract.image_to_string(img, lang='eng')  # Use 'ron' for Romanian text

    # Convert the extracted text into a list of lines and columns (basic parsing)
    lines = text.strip().split('\n')
    data = [line.split() for line in lines if line.strip()]  # Remove empty lines

    # Create a DataFrame from the parsed data
    df = pd.DataFrame(data)

    # Save each page's DataFrame as a separate sheet in the Excel file
    df.to_excel(excel_writer, sheet_name=f"Page_{i+1}", index=False)

# Step 4: Finalize the Excel file
excel_writer.close()

print(f"OCR process completed. Excel file saved as: {output_excel}")
