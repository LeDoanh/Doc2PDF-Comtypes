import sys
import os

import comtypes.client

def convert_to_pdf(in_file, out_file):
    wdFormatPDF = 17

    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(in_file)
    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()

if __name__ == "__main__":
    in_file = os.path.abspath(sys.argv[1])
    out_file = os.path.abspath(sys.argv[2])
    convert_to_pdf(in_file, out_file)
