import PyPDF2
import getpass
import os

def unlock_pdfs(input_folder, output_folder):
    # Get the password from the user without echoing it to the console
    password = getpass.getpass(prompt="Enter the common password for the PDFs: ")

    # Create the output folder if it doesn't exist
    os.makedirs(output_folder, exist_ok=True)

    # Get a list of all PDF files in the input folder
    pdf_files = [f for f in os.listdir(input_folder) if f.lower().endswith(".pdf")]

    if not pdf_files:
        print("No PDF files found in the folder.")
        return

    # Sort the list of files to maintain order
    pdf_files.sort()

    for pdf_file in pdf_files:
        input_pdf_path = os.path.join(input_folder, pdf_file)
        output_pdf_path = os.path.join(output_folder, pdf_file)

        with open(input_pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfFileReader(file)

            # Check if the PDF is encrypted
            if pdf_reader.isEncrypted:
                # Try to decrypt the PDF using the provided password
                if pdf_reader.decrypt(password):
                    # Create a new PDF writer
                    pdf_writer = PyPDF2.PdfFileWriter()

                    # Add all pages to the new writer
                    for page_num in range(pdf_reader.numPages):
                        pdf_writer.addPage(pdf_reader.getPage(page_num))

                    # Write the new PDF to the output file without encryption
                    with open(output_pdf_path, 'wb') as output_file:
                        pdf_writer.write(output_file)
                    print(f"Document '{pdf_file}' unlocked and saved at {output_pdf_path}")
                else:
                    print(f"Incorrect password for document '{pdf_file}'. Unable to unlock.")
            else:
                print(f"Document '{pdf_file}' is not encrypted. No password required.")

if __name__ == "__main__":
    # Replace 'input_folder' and 'output_folder' with your actual paths
    input_folder = r'C:\Users\42395\Downloads\documnets_with_password'
    output_folder = r'C:\Users\42395\Downloads\documnets_without_password'

    unlock_pdfs(input_folder, output_folder)
