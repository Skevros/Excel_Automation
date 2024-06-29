Imports: Include necessary libraries for image processing (OpenCV), OCR (Tesseract), data manipulation (Pandas), regular expressions (re), and file handling (os).
Configuration:
Set the path to your Tesseract OCR installation.
Specify the input folder for the images and the output folder for the updated CSV.
Adjust the confidence_threshold if needed based on image quality.
extract_data_from_image(image_path) Function:
Reads an image from the given path using OpenCV (cv2.imread).
Performs OCR using Tesseract (pytesseract.image_to_string) to extract text from the image.
parse_extracted_text(text) Function:
Uses regular expressions to search for patterns within the extracted text to find:
Name (e.g., "Name: John Doe")
Date Seen (e.g., "Date Seen: 08/15/2023")
Total amount (e.g., "Total: $150.00")
Check Number (e.g., "Check Number: 12345")
If all fields are found, returns a dictionary with the extracted data; otherwise, returns None to indicate parsing failure.
Main Script:
Load Excel Data: Loads the Excel spreadsheet using Pandas (pd.read_excel).
Process Images and Match Data:
Loops through each file in the input folder.
Checks if the file is an image (supported formats: .jpg, .png, .jpeg).
Extracts text from the image using extract_data_from_image().
Parses the extracted text using parse_extracted_text().
If parsing is successful:
Attempts to find a matching row in the DataFrame based on the extracted name and date.
If a match is found:
Compares the extracted total to the 'total sum paid' in the Excel sheet.
If totals match: Updates the 'check number' column in the DataFrame with the extracted check number.
If totals don't match: Adds a "disc__" prefix along with the check number to the 'check number' column.
If no match is found: Prints an error message indicating the missing data.
If parsing fails: Prints an error message.
Save Updated Excel: Saves the updated DataFrame back to a new Excel file (updated_data.xlsx) in the specified output folder.
