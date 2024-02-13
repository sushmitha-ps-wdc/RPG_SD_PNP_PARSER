from PIL import Image
from reportlab.pdfgen import canvas
import os

def convert_jpg_to_pdf(folder_path, output_pdf_path):
    # Get a list of all JPG files in the folder
    jpg_files = [f for f in os.listdir(folder_path) if f.lower().endswith(".jpg")]

    if not jpg_files:
        print("No JPG files found in the folder.")
        return

    # Sort the list of files to maintain order
    jpg_files.sort()

    for jpg_file in jpg_files:
        jpg_path = os.path.join(folder_path, jpg_file)

        # Create a PDF file for each JPG file
        pdf_path = os.path.join(output_pdf_path, f"{os.path.splitext(jpg_file)[0]}.pdf")
        pdf_canvas = canvas.Canvas(pdf_path)

        img = Image.open(jpg_path)

        # Assuming A4 size (you can adjust the dimensions accordingly)
        pdf_canvas.setPageSize((img.width, img.height))
        pdf_canvas.drawImage(jpg_path, 0, 0, width=img.width, height=img.height)

        pdf_canvas.showPage()

        # Save the PDF file
        pdf_canvas.save()
        print(f"PDF file created at {pdf_path}")

if __name__ == "__main__":
    # Replace 'input_folder' and 'output_folder' with your actual paths
    input_folder = r'C:\Users\42395\Downloads\marks_card_jpg'
    output_folder = r'C:\Users\42395\Downloads\marks_card_pdf'

    convert_jpg_to_pdf(input_folder, output_folder)
