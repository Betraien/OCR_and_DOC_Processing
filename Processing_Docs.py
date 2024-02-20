# Requires Python 3.6 or higher due to f-strings

# Import libraries
import os
import platform
from tempfile import TemporaryDirectory
from pathlib import Path

from spire.doc import *
from spire.doc.common import *

import pytesseract
from pdf2image import convert_from_path
from PIL import Image

import pandas

if platform.system() == "Windows":
    # We may need to do some additional downloading and setup...
    # Windows needs a PyTesseract Download
    # https://github.com/UB-Mannheim/tesseract/wiki/Downloading-Tesseract-OCR-Engine

    pytesseract.pytesseract.tesseract_cmd = (
        r"C:\Users\User\Desktop\E&pm OCR - Shabib\Scripting\Libraries\Tesseract-OCR\tesseract.exe"
    )

    # Windows also needs poppler_exe
    path_to_poppler_exe = Path(r"C:\Users\User\Desktop\E&pm OCR - Shabib\Scripting\Libraries\poppler-23.11.0\Library\bin")
    
    # Put our output files in a sane place...
    out_directory = Path(r"C:\Users\User\Desktop\E&pm OCR - Shabib\Scripting")
 

# Path of the Input file
file = r"C:\Users\User\Desktop\E&pm OCR - Shabib\Scripting\file-sample_100kB.doc"
root_ext = os.path.splitext(file)[1].lower()


# Store all the pages of the PDF in a variable
image_file_list = []

text_file = out_directory / Path("out_text.txt")

def word():

    global file

    # Create a Document object
    document = Document()

    # Load a Word file from disk
    document.LoadFromFile(file)

    # Save the Word file in txt format
    document.SaveToFile("out_text.txt", FileFormat.Txt)
    document.Close()


#////////////////////////////////////////////////////////////////////////////////////////////////

def excel():
    
    global file
    
    if root_ext == '.xlsx':
        df = pandas.read_excel(file, engine='openpyxl')
    elif root_ext == '.xls':
        df = pandas.read_excel(file)
    elif root_ext == '.csv':
        df = pandas.read_csv(file)

    # Convert the dataframe to text (string)
    text_data = df.to_csv(sep='\t', index=False)

    # Write to a text file
    txt_file = r'out_text.txt'  # Replace with your text file name
    with open(txt_file, 'w') as file:
        file.write(text_data)
        
#//////////////////////////////////////////////////////////////////////////////////////////////////

def pdf():
    ''' Main execution point of the program'''
    with TemporaryDirectory() as tempdir:
        # Create a temporary directory to hold our temporary images.

        """
        Part #1 : Converting PDF to images
        """

        if platform.system() == "Windows":
            pdf_pages = convert_from_path(
                file, 500, poppler_path=path_to_poppler_exe
            )
        else:
            pdf_pages = convert_from_path(file, 500)
        # Read in the PDF file at 500 DPI

        # Iterate through all the pages stored above
        for page_enumeration, page in enumerate(pdf_pages, start=1):
            # enumerate() "counts" the pages for us.

            # Create a file name to store the image
            filename = f"{tempdir}\page_{page_enumeration:03}.jpg"

            # Declaring filename for each page of PDF as JPG
            # For each page, filename will be:
            # PDF page 1 -> page_001.jpg
            # PDF page 2 -> page_002.jpg
            # PDF page 3 -> page_003.jpg
            # ....
            # PDF page n -> page_00n.jpg

            # Save the image of the page in system
            page.save(filename, "JPEG")
            image_file_list.append(filename)

        """
        Part #2 - Recognizing text from the images using OCR
        """

        with open(text_file, "a") as output_file:
            # Open the file in append mode so that
            # All contents of all images are added to the same file

            # Iterate from 1 to total number of pages
            for image_file in image_file_list:

                # Set filename to recognize text from
                # Again, these files will be:
                # page_1.jpg
                # page_2.jpg
                # ....
                # page_n.jpg

                # Recognize the text as string in image using pytesserct
                text = str(((pytesseract.image_to_string(Image.open(image_file)))))

                # The recognized text is stored in variable text
                # Any string processing may be applied on text
                # Here, basic formatting has been done:
                # In many PDFs, at line ending, if a word can't
                # be written fully, a 'hyphen' is added.
                # The rest of the word is written in the next line
                # Eg: This is a sample text this word here GeeksF-
                # orGeeks is half on first line, remaining on next.
                # To remove this, we replace every '-\n' to ''.
                text = text.replace("-\n", "")

                # Finally, write the processed text to the file.
                output_file.write(text)

            # At the end of the with .. output_file block
            # the file is closed after writing all the text.
        # At the end of the with .. tempdir block, the 
        # TemporaryDirectory() we're using gets removed!	 
    # End of main function!
    
if __name__ == "__main__":
    # We only want to run this if it's directly executed!
    
    if(root_ext == ".pdf"):
        print("im running pdf")
        pdf()
    elif(root_ext == ".xlsx" or root_ext == ".xls" or root_ext == ".csv"):
        print("im running excel")
        excel()
    elif(root_ext == ".doc" or root_ext == ".docx"):
        print("im running word")
        word()
    else:		
        print("Unrecognised file type")
    

    

    