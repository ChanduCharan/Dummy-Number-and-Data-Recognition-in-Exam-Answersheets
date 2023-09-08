import cv2
import os
import pytesseract
import pandas as pd
import openpyxl
from PIL import Image


# Path to the Tesseract executable
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\Darcodux\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'

# Path to the folder containing the exam answer sheets
folder_path = r'C:\Users\Darcodux\Documents\FYPP\Marks_sheet'

# Get a list of all the image files in the folder
image_files = [f for f in os.listdir(folder_path) if f.endswith('.jpg')]

# Iterate over each image file
for image_file in image_files:
    # Construct the full image path
    image_path = os.path.join(folder_path, image_file)

    # Load the image
    Image = cv2.imread(image_path)

    # Convert the image to grayscale for better OCR accuracy
    gray_img = cv2.cvtColor(Image, cv2.COLOR_BGR2GRAY)

    # Perform OCR on the image
    extracted_text = pytesseract.image_to_string(gray_img)

    # Split the extracted text into lines
    lines = extracted_text.split('\n')

    # Initialize variables to store data
    student_data = {}
    answers = {}

    # Iterate through each line of the extracted text
    for line in lines:
        # Check if the line contains "Name" (case-insensitive) and extract the name
        if "Name" in line:
            student_data['Student Name'] = line.split("Name")[1].strip()
        # Check if the line contains "Roll Number" (case-insensitive) and extract the roll number
        if "Roll Number" in line:
            student_data['Roll Number'] = line.split("Roll Number")[1].strip()

        # Check if the line starts with "Q" and ends with a question number (1-20)
        if line.startswith("Q") and line[1:].isdigit() and int(line[1:]) in range(1, 21):
            question_number = int(line[1:])
            answer = line.split(":")[1].strip()
            answers[question_number] = answer

    # Create a DataFrame to store the extracted data
    df = pd.DataFrame(student_data, index=[0])

    for question_number, answer in answers.items():
        df[f'Q{question_number}'] = [answer]

    # Create a new Excel workbook and select the default sheet
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    # Split the extracted text into lines (assuming each line corresponds to a response)
    response_lines = extracted_text.split('\n')

    # Write each response to the Excel sheet, one response per row
    for i, response in enumerate(response_lines, start=1):
        sheet.cell(row=i, column=1, value=response)

    # Save the Excel file with the extracted responses
    workbook.save(f'{image_file}_responses.xlsx')

    print(f"Data has been written to {image_file}_responses.xlsx")