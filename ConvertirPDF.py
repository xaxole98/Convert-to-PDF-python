import PyPDF2  # Import PyPDF2 library for working with PDF files
import os  # Import os module for file and path operations
import subprocess  # Import subprocess module for executing external commands
# Import filedialog from tkinter for file selection dialog
import tkinter.filedialog as fd
import tkinter as tk  # Import tkinter module for creating GUI


def convert_to_pdf(input_path):
    """
    Function to convert a document to PDF.

    Args:
        input_path (str): Path of the input document file.

    Returns:
        None
    """
    input_extension = os.path.splitext(
        input_path)[1][1:].lower()  # Get the extension of the input file

    if input_extension == 'pdf':
        # If the file is already a PDF, print a warning message
        print("The selected file is already a PDF.")
    else:
        output_pdf = PyPDF2.PdfFileWriter()  # Create an output PDF object

        if input_extension in ['doc', 'docx']:
            # Get the directory where the script is located
            output_path = os.path.dirname(os.path.abspath(__file__))
            subprocess.run(['libreoffice', '--headless', '--convert-to',
                            'pdf', '--outdir', output_path, input_path])
            # Use LibreOffice to convert Word documents to PDF
        elif input_extension == 'txt':
            # Get the directory where the script is located
            output_path = os.path.dirname(os.path.abspath(__file__))
            output_file_path = os.path.join(output_path, 'output.pdf')
            with open(input_path, 'rb') as txt_file:
                text = txt_file.read()
                output_pdf.addPage(
                    PyPDF2.pdf.PageObject.createTextString(text))
                # Convert text files to PDF using PyPDF2
                with open(output_file_path, 'wb') as output_file:
                    output_pdf.write(output_file)
        else:
            print(
                f"Cannot convert the file with extension '{input_extension}' to PDF.")
            # If the file extension is not recognized, print an error message


root = tk.Tk()  # Create a tkinter window object
root.withdraw()  # Hide the main tkinter window

input_file_path = fd.askopenfilename(
    title="Select input file")  # Open a file selection dialog

if input_file_path:
    convert_to_pdf(input_file_path)  # Convert the selected file to PDF
    print("Conversion completed.")  # Print a completion message

root.mainloop()  # Run the tkinter event loop

# Enjoy :p