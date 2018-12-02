import sys
import os
import comtypes.client

wdFormatPDF = 17

in_file = r"E:\TUTORIALS_AND_CODINGS\Python_Coding_practice\CONSEN\Text.docx"
out_file = r"E:\TUTORIALS_AND_CODINGS\Python_Coding_practice\CONSEN\PDF_Files\Text.pdf"

word = comtypes.client.CreateObject('Word.Application')
doc = word.Documents.Open(in_file)
doc.SaveAs(out_file, FileFormat=wdFormatPDF)
doc.Close()
word.Quit()