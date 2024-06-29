import os
import re
import cv2
import pytesseract
import pandas as pd

# --- Configuration ---
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
input_folder = 'input_images'
output_folder = 'output_data'
confidence_threshold = 60
excel_file = 'your_excel_file.xlsx'


# --- Helper Functions ---
def extract_data_from_image(image_path):
    """Extracts text from an image using Tesseract OCR.

    Args:
        image_path (str): Path to the image file.

    Returns:
        str: Extracted text from the image.
    """

    img = cv2.imread(image_path)
    text = pytesseract.image_to_string(img)
    return text


def parse_extracted_text(text):
    """Parses the extracted text to find relevant data.

    Args:
        text (str): Text extracted from the OCR process.

    Returns:
        dict: A dictionary containing the extracted data
              (name, date_seen, total, check_number).
              Returns None if parsing fails.
    """

    try:
        name_match = re.search(r"Name:\s*(.+)", text)
        date_seen_match = re.search(r"Date Seen:\s*(\d{2}/\d{2}/\d{4})", text)
        total_match = re.search(r"Total:\s*\$([\d.]+)", text)
        check_number_match = re.search(r"Check Number:\s*(\d+)", text)

        if all([name_match, date_seen_match, total_match, check_number_match]):
            return {
                'name': name_match.group(1).strip(),
                'date_seen': date_seen_match.group(1),
                'total': float(total_match.group(1)),
                'check_number': check_number_match.group(1)
            }
        else:
            return None
    finally:
        return None


# --- Main Script Logic ---

# 1. Load Excel Workbook
excel_workbook = pd.ExcelFile(excel_file)
sheet_names = excel_workbook.sheet_names

# 2. Process Images and Match Data for Each Sheet
for sheet_name in sheet_names:
    df = excel_workbook.parse(sheet_name)

    for filename in os.listdir(input_folder):
        if filename.endswith(('.jpg', '.png', '.jpeg')):
            image_path = os.path.join(input_folder, filename)
            extracted_text = extract_data_from_image(image_path)
            data = parse_extracted_text(extracted_text)

            if data:
                matching_rows = df[
                    (df['name'] == data['name']) &
                    (df['date seen'] == data['date_seen'])
                    ]

                if not matching_rows.empty:
                    row_index = matching_rows.index[0]

                    if data['total'] == df.loc[row_index, 'total sum paid']:
                        df.loc[row_index, 'check number'] = data['check_number']
                    else:
                        df.loc[row_index, 'check number'] = f"disc__{data['check_number']}"
                else:
                    print(f"No match found in sheet '{sheet_name}' for data from {filename}: {data}")
            else:
                print(f"Unable to extract data from {filename}")

    # 3. Update Sheet in Excel Workbook (in-memory)
    excel_workbook.book.remove(excel_workbook.book[sheet_name])  # Remove old sheet
    excel_workbook.book.sheets.append(df)  # Add updated sheet

# 4. Save Updated Excel File
excel_workbook.save(os.path.join(output_folder, 'updated_data.xlsx'))
print("Processing complete. Updated data saved to updated_data.xlsx")
