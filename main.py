import csv
import os
from barcode import Code128
from barcode.writer import ImageWriter
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib import colors
import json

# Extarcting Setting file

file_path = 'settings.json'

# Write the settings to the specified JSON file
with open(file_path, 'r') as f:
    settings = json.load(f)

output_type = settings.get("output_type")
input_csv = settings.get("input_csv")
output_file = settings.get("output_file")
barcode_folder = settings.get("barcode_folder")
imageformat = settings.get("imageformat")
ImageWidth = settings.get("ImageWidth")
ImageHeight = settings.get("ImageHeight")


def output_file_pdf(imageformat, ImageWidth, ImageHeight, barcode_folder, output_file):
    with open(input_csv, "r") as file:
        reader = csv.reader(file)
        y_position = height - 50  # Start a bit from the top to leave space for the header

        for idx, row in enumerate(reader):
            if not row or not row[0].strip():  # Skip empty rows or blank data
                continue

            data = row[0].strip()  # Clean and get the first column value

            # Set the barcode image file name without the .png extension being duplicated
            barcode_path = os.path.abspath(os.path.join(barcode_folder, f"{data}"))

            #barcode_path = os.path.join(barcode_folder, f"{data}")  # Do not add .png here

            if not os.path.exists(f"{barcode_path}{imageformat}"):  # Check if the file already exists
                try:
                    barcode = Code128(data, writer=ImageWriter())
                    # Save barcode image with the .png extension correctly
                    barcode.save(barcode_path)  # Only barcode_path without .png extension here
                except Exception as e:
                    print(f"Failed to generate barcode for '{data}': {e}")
                    continue

            # Add barcode data (text) to the PDF
            c.setFont("Helvetica", 10)
            c.drawString(50, y_position, f"Barcode Number: {data}")

            # Add barcode image to the PDF
            try:
                c.drawImage(f"{barcode_path}.png", 200, y_position - 10, width=ImageWidth, height=ImageHeight)
            except Exception as e:
                print(f"Failed to add image for data '{data}': {e}")

            y_position -= 70  # Move down for the next barcode (adjust spacing as needed)

            # Check if the content is near the bottom of the page, and if so, create a new page
            if y_position < 100:
                c.showPage()  # Start a new page in the PDF
                y_position = height - 50  # Reset the y_position to start from the top of the new page
    try:
        c.save()
        print(f"Barcodes saved in '{output_file}' and images in the '{barcode_folder}' folder.")
    except Exception as e:
        print(f"Failed to save the PDF file: {e}")        

def output_file_excel(imageformat, ImageWidth, ImageHeight, barcode_folder, output_file):
    with open(input_csv, "r") as file:
        reader = csv.reader(file)
        for idx, row in enumerate(reader, start=2):  # Start from row 2 (after header)
            if not row or not row[0].strip():  # Skip empty rows or blank data
                continue

            data = row[0].strip()  # Clean and get the first column value

            # Set the barcode image file name without the .png extension being duplicated
            #barcode_path = os.path.join(barcode_folder, f"{data}")  # Do not add .png here
            barcode_path = os.path.abspath(os.path.join(barcode_folder, f"{data}"))

            if not os.path.exists(f"{barcode_path}{imageformat}"):  # Check if the file already exists
                try:
                    barcode = Code128(data, writer=ImageWriter())
                    # Save barcode image with the .png extension correctly
                    barcode.save(barcode_path)  # Only barcode_path without .png extension here
                except Exception as e:
                    print(f"Failed to generate barcode for '{data}': {e}")
                    continue

            # Add the barcode number to column 1 (A)
            sheet[f"A{idx}"] = data

            # Add the barcode image to column 2 (B)
            try:
                img = ExcelImage(f"{barcode_path}{imageformat}")  # Load barcode image with .png extension
                img.height = ImageHeight  # Resize image (optional)
                img.width = ImageWidth
                cell_position = f"B{idx}"  # Position of the barcode image
                sheet.add_image(img, cell_position)
            except Exception as e:
                print(f"Failed to add image for data '{data}': {e}")

    # Save the Excel file
    try:
        workbook.save(output_file)
        print(f"Barcodes saved in '{output_file}' and images in the '{barcode_folder}' folder.")
    except Exception as e:
        print(f"Failed to save the Excel file: {e}")


try:
    os.makedirs(barcode_folder, exist_ok=True)
    if output_type == "PDF":
        # Initialize the PDF
        c = canvas.Canvas(output_file, pagesize=letter)
        width, height = letter  # Page size (8.5 x 11 inches)
        #call pdf function
        output_file_pdf(imageformat, ImageWidth, ImageHeight, barcode_folder, output_file)
    elif output_type == "Excel":
        # Initialize an Excel workbook
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Barcodes"
        sheet.append(["Barcode Number", "Barcode Image"])  # Header row
        # call excel function
        output_file_excel(imageformat, ImageWidth, ImageHeight, barcode_folder, output_file)
except Exception as e:
    print(f"Failed to initialize the output file. {e}")
    exit()

