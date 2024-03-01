import sys
import os

import comtypes.client

def convert_to_pdf(in_file, out_file):
    # Check if the input file exists
    if not os.path.exists(PathIn):
        print("File not found: " + PathIn)
        return
    
    # Check if the output file exists, remove it if it does
    if os.path.exists(PathOut):
        os.remove(PathOut)
    
    # Create a new instance of Word
    wdFormatPDF = 17

    # Open the input file
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(in_file)
    
    # Save the file as a PDF
    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
    
    # Close the file and the Word instance
    doc.Close()
    word.Quit()

if __name__ == "__main__":
    in_file = os.path.abspath(sys.argv[1])
    out_file = os.path.abspath(sys.argv[2])
    convert_to_pdf(in_file, out_file)
