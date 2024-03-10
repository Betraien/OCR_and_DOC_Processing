import easyocr
import cv2

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
        r"C:\Users\Shabib\Desktop\Shabib\E&Pm\OCR\libraries\tesserect\tesseract.exe"
    )

    # Windows also needs poppler_exe
    path_to_poppler_exe = Path(r"C:\Users\Shabib\Desktop\Shabib\E&Pm\OCR\libraries\poppler-23.11.0\Library\bin")
    
    # Put our output files in a sane place...
 


# Store all the pages of the PDF in a variable
image_file_list = []


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
                file, 1000, poppler_path=path_to_poppler_exe
            )
        else:
            pdf_pages = convert_from_path(file, 1000)
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
        reader = easyocr.Reader(['en'])  # Replace 'fr' with your desired language code
        

        with open(text_file, "a") as output_file:
            # Open the file in append mode so that
            # All contents of all images are added to the same file

            # Iterate from 1 to total number of pages
            counter = 0
            for image_file in image_file_list:
                counter = counter + 1

                # Set filename to recognize text from
                # Again, these files will be:
                # page_1.jpg
                # page_2.jpg
                # ....
                # page_n.jpg

                # Recognize the text as string in image using pytesserct


                image_path = image_file
                img = cv2.imread(image_path)

                detect = reader.readtext(img)



                ##text = str(((pytesseract.image_to_string(Image.open(image_file)))))

                # The recognized text is stored in variable text
                # Any string processing may be applied on text
                # Here, basic formatting has been done:
                # In many PDFs, at line ending, if a word can't
                # be written fully, a 'hyphen' is added.
                # The rest of the word is written in the next line
                # Eg: This is a sample text this word here GeeksF-
                # orGeeks is half on first line, remaining on next.
                # To remove this, we replace every '-\n' to ''.

                ##text = text.replace("-\n", "")

                for result in detect:

                    text = result[1]
                    confidence = result[2]  # Access other elements based on their index
                    bounding_box = result[0]

                    # Print the extracted text, confidence score, and bounding box coordinates
                    print(f"Text: {text}, Confidence: {confidence:.2f}")

                    # Optionally, draw the bounding box on the image
                    cv2.rectangle(img, (int(bounding_box[0][0]), int(bounding_box[0][1])), 
                                    (int(bounding_box[2][0]), int(bounding_box[2][1])), (0, 255, 0), 2)

                # Display the image with bounding boxes (if drawn)
                    
                cv2.imwrite(out_directory + "\output_image"+str(counter)+".jpg", img)


                # Finally, write the processed text to the file.
                #string_to_write += text[0][1]  # Join elements with newline characters


            #output_file.write(string_to_write)

            # At the end of the with .. output_file block
            # the file is closed after writing all the text.
        # At the end of the with .. tempdir block, the 
        # TemporaryDirectory() we're using gets removed!	 
    # End of main function!
    
if __name__ == "__main__":

    out_directory = Path(r"C:\Users\Shabib\Desktop\Shabib\E&Pm\OCR\outputs")
    text_file = out_directory / Path("PDT-030.txt")



    # Path of the Input file
    file = r"C:\Users\Shabib\Desktop\Shabib\E&Pm\OCR\inputs\PDT-030.pdf"
    root_ext = os.path.splitext(file)[1].lower()
    
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
    

    

    
